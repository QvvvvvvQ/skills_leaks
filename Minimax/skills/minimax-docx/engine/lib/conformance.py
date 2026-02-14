"""
conformance.py - OOXML child-element ordering rectifier.

Uses the ECMA376Schema parser to dynamically retrieve canonical
element sequences per ISO/IEC 29500-1:2016.

Reference: ECMA-376 5th Edition, ISO/IEC 29500-1:2016
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.etree.ElementTree import Element

from .xmlns import NS

# Import schema from new location
_SKILL_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(_SKILL_ROOT))
from schema.ecma376_parser import ECMA376Schema
sys.path.pop(0)


# ═════════════════════════════════════════════════════════════════════════
# Schema loading
# ═════════════════════════════════════════════════════════════════════════

_SCHEMA = ECMA376Schema()


# ═════════════════════════════════════════════════════════════════════════
# Low-level helpers
# ═════════════════════════════════════════════════════════════════════════

def _local_name(el: Element) -> str:
    """Extract the local name from a Clark-notation tag."""
    t = el.tag
    pos = t.rfind("}")
    return t[pos + 1:] if pos >= 0 else t


def _rank_vector(seq: list[str]) -> dict[str, int]:
    """Map element local-names to their ordinal positions."""
    return {name: idx for idx, name in enumerate(seq)}


# ═════════════════════════════════════════════════════════════════════════
# Core reordering
# ═════════════════════════════════════════════════════════════════════════

def enforce_child_order(parent: Element, seq: list[str]) -> bool:
    """Sort *parent*'s children per *seq*, preserving unknown elements.

    Unknown elements keep their mutual order and are appended after
    the last recognised element.  Returns True when the order changed.
    """
    children = list(parent)
    if len(children) < 2:
        return False

    ranks = _rank_vector(seq)
    ceiling = len(seq)

    original_ids = {id(c): i for i, c in enumerate(children)}

    def _sort_key(child: Element) -> tuple[int, int]:
        r = ranks.get(_local_name(child), ceiling)
        return (r, original_ids[id(child)])

    reordered = sorted(children, key=_sort_key)
    if all(a is b for a, b in zip(children, reordered)):
        return False

    for c in children:
        parent.remove(c)
    parent.extend(reordered)
    return True


# ═════════════════════════════════════════════════════════════════════════
# Structural patches
# ═════════════════════════════════════════════════════════════════════════

_PARA_BORDER_LOCALS = frozenset(
    _SCHEMA.get_child_order("pBdr") or ["top", "left", "bottom", "right", "between", "bar"]
)


def _anchor_body_sectpr(body: Element) -> bool:
    """Ensure the body-level ``sectPr`` is the very last child."""
    children = list(body)
    for idx, child in enumerate(children):
        if _local_name(child) == "sectPr" and idx < len(children) - 1:
            body.remove(child)
            body.append(child)
            return True
    return False


def wrap_orphan_borders(ppr: Element) -> int:
    """Move stray border elements into a ``pBdr`` wrapper."""
    orphans = [c for c in ppr if _local_name(c) in _PARA_BORDER_LOCALS]
    if not orphans:
        return 0

    wrapper = ppr.find(NS.wtag("pBdr"))
    if wrapper is None:
        wrapper = Element(NS.wtag("pBdr"))
        # Insert at a reasonable position (after style/keep/widow/num)
        _prior = {"pStyle", "keepNext", "keepLines", "pageBreakBefore",
                   "framePr", "widowControl", "numPr", "suppressLineNumbers"}
        slot = 0
        for i, c in enumerate(ppr):
            if _local_name(c) not in _prior:
                slot = i
                break
        else:
            slot = len(list(ppr))
        ppr.insert(slot, wrapper)

    for orphan in orphans:
        ppr.remove(orphan)
        wrapper.append(orphan)

    pbdr_seq = _SCHEMA.get_child_order("pBdr")
    if pbdr_seq:
        enforce_child_order(wrapper, pbdr_seq)
    return len(orphans)


# ═════════════════════════════════════════════════════════════════════════
# Top-level rectifier - schema-driven dispatch
# ═════════════════════════════════════════════════════════════════════════

def rectify_document_tree(root: Element) -> int:
    """Walk *root* and reorder children of all known property containers.

    Returns the total number of containers whose child order was corrected.
    """
    corrections = 0

    for local_tag in _SCHEMA.get_all_containers():
        seq = _SCHEMA.get_child_order(local_tag)
        if not seq:
            continue
        qualified = NS.wtag(local_tag)
        for elem in root.iter(qualified):
            if local_tag == "pPr":
                corrections += wrap_orphan_borders(elem)
            if enforce_child_order(elem, seq):
                corrections += 1

    for body in root.iter(NS.wtag("body")):
        if _anchor_body_sectpr(body):
            corrections += 1

    return corrections


def rectify_settings_root(root: Element) -> int:
    """Reorder children of the ``<w:settings>`` root element."""
    settings_seq = _SCHEMA.get_child_order("settings")
    if not settings_seq:
        return 0
    target = root if _local_name(root) == "settings" else root.find(NS.wtag("settings"))
    if target is not None and enforce_child_order(target, settings_seq):
        return 1
    return 0


# ═════════════════════════════════════════════════════════════════════════
# Table cell-width alignment
# ═════════════════════════════════════════════════════════════════════════

_W_ATTR = f"{{{NS.MAIN}}}w"
_TYPE_ATTR = f"{{{NS.MAIN}}}type"
_VAL_ATTR = f"{{{NS.MAIN}}}val"


def _is_nested_table(tbl: Element, all_tables: list[Element]) -> bool:
    for candidate in all_tables:
        if candidate is tbl:
            continue
        if any(inner is tbl for inner in candidate.iter(NS.wtag("tbl"))):
            return True
    return False


def _read_grid_widths(grid: Element) -> list[int] | None:
    cols = grid.findall(NS.wtag("gridCol"))
    if not cols:
        return None
    widths: list[int] = []
    for col in cols:
        raw = col.get(_W_ATTR)
        if raw is None:
            return None
        try:
            widths.append(int(raw))
        except (ValueError, TypeError):
            return None
    return widths


def align_cell_widths(root: Element) -> int:
    """Sync ``tcW`` values with ``tblGrid/gridCol`` definitions.

    Only absolute (dxa) widths are touched.  Percentage/auto widths
    and cells in nested tables are skipped.

    Returns the number of corrections applied.
    """
    all_tables = root.findall(f".//{{{NS.MAIN}}}tbl")
    patched = 0

    for tbl in all_tables:
        if _is_nested_table(tbl, all_tables):
            continue

        grid = tbl.find(NS.wtag("tblGrid"))
        if grid is None:
            continue
        col_widths = _read_grid_widths(grid)
        if not col_widths:
            continue

        for row in tbl.findall(NS.wtag("tr")):
            ci = 0
            for cell in row.findall(NS.wtag("tc")):
                if ci >= len(col_widths):
                    break
                tc_pr = cell.find(NS.wtag("tcPr"))
                if tc_pr is None:
                    ci += 1
                    continue
                tc_w = tc_pr.find(NS.wtag("tcW"))
                if tc_w is None:
                    ci += 1
                    continue

                w_type = tc_w.get(_TYPE_ATTR)
                if w_type not in (None, "", "dxa"):
                    ci += 1
                    continue

                span_el = tc_pr.find(NS.wtag("gridSpan"))
                span = 1
                if span_el is not None:
                    try:
                        span = int(span_el.get(_VAL_ATTR, "1"))
                    except (ValueError, TypeError):
                        pass

                end = ci + span
                if end > len(col_widths):
                    ci = end
                    continue

                target_w = sum(col_widths[ci:end])
                raw = tc_w.get(_W_ATTR)
                if raw is None:
                    ci = end
                    continue
                try:
                    current_w = int(raw)
                except (ValueError, TypeError):
                    ci = end
                    continue

                if target_w > 0 and abs(current_w - target_w) / target_w > 0.04:
                    tc_w.set(_W_ATTR, str(target_w))
                    patched += 1

                ci = end

    return patched
