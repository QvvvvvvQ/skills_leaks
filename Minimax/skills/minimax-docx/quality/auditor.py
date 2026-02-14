#!/usr/bin/env python3
"""
Document Auditor - Rule-based conformance and validation pipeline.

Combines structural normalization (element reordering) with
rule-engine-based verification into a unified pass.

Usage::

    python auditor.py <document.docx>
"""

import sys
import zipfile
import tempfile
import shutil
from pathlib import Path
from xml.etree import ElementTree as ET

# Add parent to path for imports
_SKILL_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_ROOT))

from schema.ecma376_parser import ECMA376Schema
from quality.rules import RuleEngine, DocumentContext
from quality.findings import FindingLevel
from packaging.opc import OPCPackager

sys.path.pop(0)

# Import engine lib modules
sys.path.insert(0, str(_SKILL_ROOT / "engine" / "lib"))
from xmlns import NS
sys.path.pop(0)


class StructuralNormalizer:
    """Normalize OOXML element ordering per ECMA-376 schema.

    Uses the ECMA376Schema parser to dynamically retrieve canonical
    element sequences rather than relying on hardcoded mappings.
    """

    def __init__(self):
        self._schema = ECMA376Schema()
        self._para_border_locals = frozenset(
            self._schema.get_child_order("pBdr") or
            ["top", "left", "bottom", "right", "between", "bar"]
        )

    def rectify_tree(self, root: ET.Element) -> int:
        """Walk tree and reorder children of all known property containers.

        Returns the total number of containers whose child order was corrected.
        """
        corrections = 0

        for local_tag in self._schema.get_all_containers():
            seq = self._schema.get_child_order(local_tag)
            if not seq:
                continue
            qualified = NS.wtag(local_tag)
            for elem in root.iter(qualified):
                if local_tag == "pPr":
                    corrections += self._wrap_orphan_borders(elem)
                if self._enforce_child_order(elem, seq):
                    corrections += 1

        # Ensure body sectPr is last
        for body in root.iter(NS.wtag("body")):
            if self._anchor_body_sectpr(body):
                corrections += 1

        return corrections

    def rectify_settings(self, root: ET.Element) -> int:
        """Reorder children of the settings root element."""
        settings_seq = self._schema.get_child_order("settings")
        if not settings_seq:
            return 0

        local = self._local_name(root)
        target = root if local == "settings" else root.find(NS.wtag("settings"))
        if target is not None and self._enforce_child_order(target, settings_seq):
            return 1
        return 0

    def _enforce_child_order(self, parent: ET.Element, seq: list[str]) -> bool:
        """Sort parent's children per seq, preserving unknown elements."""
        children = list(parent)
        if len(children) < 2:
            return False

        ranks = {name: idx for idx, name in enumerate(seq)}
        ceiling = len(seq)
        original_ids = {id(c): i for i, c in enumerate(children)}

        def _sort_key(child: ET.Element) -> tuple[int, int]:
            r = ranks.get(self._local_name(child), ceiling)
            return (r, original_ids[id(child)])

        reordered = sorted(children, key=_sort_key)
        if all(a is b for a, b in zip(children, reordered)):
            return False

        for c in children:
            parent.remove(c)
        parent.extend(reordered)
        return True

    def _wrap_orphan_borders(self, ppr: ET.Element) -> int:
        """Move stray border elements into a pBdr wrapper."""
        orphans = [c for c in ppr if self._local_name(c) in self._para_border_locals]
        if not orphans:
            return 0

        wrapper = ppr.find(NS.wtag("pBdr"))
        if wrapper is None:
            wrapper = ET.Element(NS.wtag("pBdr"))
            _prior = {"pStyle", "keepNext", "keepLines", "pageBreakBefore",
                      "framePr", "widowControl", "numPr", "suppressLineNumbers"}
            slot = 0
            for i, c in enumerate(ppr):
                if self._local_name(c) not in _prior:
                    slot = i
                    break
            else:
                slot = len(list(ppr))
            ppr.insert(slot, wrapper)

        for orphan in orphans:
            ppr.remove(orphan)
            wrapper.append(orphan)

        pbdr_seq = self._schema.get_child_order("pBdr")
        if pbdr_seq:
            self._enforce_child_order(wrapper, pbdr_seq)
        return len(orphans)

    def _anchor_body_sectpr(self, body: ET.Element) -> bool:
        """Ensure the body-level sectPr is the very last child."""
        children = list(body)
        for idx, child in enumerate(children):
            if self._local_name(child) == "sectPr" and idx < len(children) - 1:
                body.remove(child)
                body.append(child)
                return True
        return False

    @staticmethod
    def _local_name(el: ET.Element) -> str:
        """Extract local name from Clark notation tag."""
        t = el.tag
        pos = t.rfind("}")
        return t[pos + 1:] if pos >= 0 else t


class CellWidthAligner:
    """Sync tcW values with tblGrid/gridCol definitions."""

    _W_ATTR = f"{{{NS.MAIN}}}w"
    _TYPE_ATTR = f"{{{NS.MAIN}}}type"
    _VAL_ATTR = f"{{{NS.MAIN}}}val"
    _TOLERANCE = 0.04  # 4% width mismatch threshold (empirical cross-platform testing)

    def align(self, root: ET.Element) -> int:
        """Align cell widths. Returns correction count."""
        all_tables = root.findall(f".//{{{NS.MAIN}}}tbl")
        patched = 0

        for tbl in all_tables:
            if self._is_nested(tbl, all_tables):
                continue

            grid = tbl.find(NS.wtag("tblGrid"))
            if grid is None:
                continue
            col_widths = self._read_grid_widths(grid)
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

                    w_type = tc_w.get(self._TYPE_ATTR)
                    if w_type not in (None, "", "dxa"):
                        ci += 1
                        continue

                    span_el = tc_pr.find(NS.wtag("gridSpan"))
                    span = 1
                    if span_el is not None:
                        try:
                            span = int(span_el.get(self._VAL_ATTR, "1"))
                        except (ValueError, TypeError):
                            pass

                    end = ci + span
                    if end > len(col_widths):
                        ci = end
                        continue

                    target_w = sum(col_widths[ci:end])
                    raw = tc_w.get(self._W_ATTR)
                    if raw is None:
                        ci = end
                        continue
                    try:
                        current_w = int(raw)
                    except (ValueError, TypeError):
                        ci = end
                        continue

                    if target_w > 0 and abs(current_w - target_w) / target_w > self._TOLERANCE:
                        tc_w.set(self._W_ATTR, str(target_w))
                        patched += 1

                    ci = end

        return patched

    def _is_nested(self, tbl: ET.Element, all_tables: list[ET.Element]) -> bool:
        for candidate in all_tables:
            if candidate is tbl:
                continue
            if any(inner is tbl for inner in candidate.iter(NS.wtag("tbl"))):
                return True
        return False

    def _read_grid_widths(self, grid: ET.Element) -> list[int] | None:
        cols = grid.findall(NS.wtag("gridCol"))
        if not cols:
            return None
        widths: list[int] = []
        for col in cols:
            raw = col.get(self._W_ATTR)
            if raw is None:
                return None
            try:
                widths.append(int(raw))
            except (ValueError, TypeError):
                return None
        return widths


class DocumentAuditor:
    """Extract, normalize, validate, and repackage a .docx file."""

    _CORE_PARTS = ("document.xml", "styles.xml", "numbering.xml")

    def __init__(self, docx_path: str | Path):
        self.docx_path = Path(docx_path)
        self.corrections = 0
        self.issues: list[str] = []
        self.advisories: list[str] = []
        self._normalizer = StructuralNormalizer()
        self._aligner = CellWidthAligner()
        self._rule_engine = RuleEngine.default_engine()

    def run(self) -> tuple[int, list[str], list[str]]:
        """Execute the full pipeline. Returns (corrections, issues, advisories)."""
        if not self.docx_path.exists():
            return 0, [f"LOCATION: file not found: {self.docx_path}"], []

        with tempfile.TemporaryDirectory() as staging:
            pkg = Path(staging) / "content"

            try:
                with zipfile.ZipFile(self.docx_path, "r") as arc:
                    arc.extractall(pkg)
            except zipfile.BadZipFile:
                return 0, ["INTEGRITY: archive corrupted or unreadable"], []

            self._normalize_xml(pkg)
            self._validate_rules(pkg)
            self._repackage_if_needed(pkg)

        return self.corrections, self.issues, self.advisories

    def _normalize_xml(self, pkg: Path):
        word_dir = pkg / "word"

        for part_name in self._CORE_PARTS:
            part = word_dir / part_name
            if part.exists():
                tree = ET.parse(part)
                root = tree.getroot()
                delta = self._normalizer.rectify_tree(root)
                if part_name == "document.xml":
                    delta += self._aligner.align(root)
                if delta > 0:
                    tree.write(part, encoding="UTF-8", xml_declaration=True)
                    self.corrections += delta

        settings = word_dir / "settings.xml"
        if settings.exists():
            tree = ET.parse(settings)
            root = tree.getroot()
            delta = self._normalizer.rectify_settings(root)
            delta += self._normalizer.rectify_tree(root)
            if delta > 0:
                tree.write(settings, encoding="UTF-8", xml_declaration=True)
                self.corrections += delta

        for pattern in ("header*.xml", "footer*.xml"):
            for part in word_dir.glob(pattern):
                tree = ET.parse(part)
                delta = self._normalizer.rectify_tree(tree.getroot())
                if delta > 0:
                    tree.write(part, encoding="UTF-8", xml_declaration=True)
                    self.corrections += delta

    def _validate_rules(self, pkg: Path):
        doc_xml = pkg / "word" / "document.xml"
        if doc_xml.exists():
            tree = ET.parse(doc_xml)
            root = tree.getroot()
            context = DocumentContext(pkg, root)
            findings = self._rule_engine.run(context)

            for finding in findings:
                if finding.level == FindingLevel.ERROR:
                    self.issues.append(finding.to_legacy_format())
                elif finding.level == FindingLevel.WARNING:
                    self.issues.append(finding.to_legacy_format())
                else:
                    self.advisories.append(finding.to_legacy_format())

    def _repackage_if_needed(self, pkg: Path):
        if self.corrections == 0:
            return

        backup = self.docx_path.with_suffix(".docx.bak")
        shutil.copy2(self.docx_path, backup)

        packager = OPCPackager()
        packager.repackage(pkg, self.docx_path)

        backup.unlink(missing_ok=True)


def main():
    if len(sys.argv) < 2:
        print("Usage: python auditor.py <document.docx>")
        sys.exit(1)

    auditor = DocumentAuditor(sys.argv[1])
    corrections, issues, advisories = auditor.run()

    if corrections > 0:
        print(f"Corrected {corrections} element ordering issues")

    for advisory in advisories:
        print(f"Advisory: {advisory}")

    if issues:
        for issue in issues:
            print(f"Issue: {issue}")
        sys.exit(1)
    else:
        print("Validation passed")
        sys.exit(0)


if __name__ == "__main__":
    main()
