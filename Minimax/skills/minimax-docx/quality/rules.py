"""
rules.py - Extensible validation rule engine.

Defines a protocol-based validation system where each rule is a
self-contained checker with configurable severity.

Architecture:
- ValidationRule protocol for custom rules
- RuleEngine for coordinated execution
- Built-in rules for common document quality issues
"""

from __future__ import annotations

import io
from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol, runtime_checkable
from xml.etree import ElementTree as ET

from .findings import DocumentFinding, FindingLevel


# Import namespace helpers
import sys
_SKILL_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_ROOT / "engine" / "lib"))
from xmlns import NS
sys.path.pop(0)


@dataclass
class Finding:
    """Simple finding for backward compatibility."""
    message: str
    severity: str = "error"


class DocumentContext:
    """Context object passed to validation rules.

    Provides access to document structure and content without
    requiring each rule to parse the document independently.
    """

    def __init__(self, pkg_dir: Path, doc_root: ET.Element | None = None):
        self.pkg_dir = pkg_dir
        self.word_dir = pkg_dir / "word"
        self._doc_root = doc_root
        self._rel_index: dict[str, Path] | None = None

    @property
    def doc_root(self) -> ET.Element | None:
        """Lazily load document.xml root."""
        if self._doc_root is None:
            doc_path = self.word_dir / "document.xml"
            if doc_path.exists():
                self._doc_root = ET.parse(doc_path).getroot()
        return self._doc_root

    @property
    def rel_index(self) -> dict[str, Path]:
        """Build relationship index from document.xml.rels."""
        if self._rel_index is None:
            self._rel_index = self._build_rel_index()
        return self._rel_index

    def _build_rel_index(self) -> dict[str, Path]:
        rels_path = self.word_dir / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return {}

        pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        tree = ET.parse(rels_path).getroot()
        index: dict[str, Path] = {}
        for rel in tree.findall(f"{{{pkg_ns}}}Relationship"):
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid and target:
                resolved = self.word_dir / target if not target.startswith("/") else Path(target[1:])
                index[rid] = resolved
        return index


@runtime_checkable
class ValidationRule(Protocol):
    """Protocol for validation rules.

    Each rule has a name, severity level, and check method
    that returns a list of findings.
    """

    name: str
    severity: str  # "error", "warning", "info"

    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        """Execute the rule against the document context."""
        ...


class BaseRule(ABC):
    """Base class for validation rules with common functionality."""

    name: str
    severity: str = "error"

    @abstractmethod
    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        """Execute the rule."""
        pass

    def _finding(
        self,
        message: str,
        location: str | None = None,
        level: FindingLevel | None = None,
    ) -> DocumentFinding:
        """Helper to create a finding with this rule's defaults."""
        if level is None:
            level = FindingLevel(self.severity)
        return DocumentFinding(
            rule_name=self.name,
            level=level,
            message=message,
            location=location,
        )


# ═══════════════════════════════════════════════════════════════════════════
# Built-in Rules
# ═══════════════════════════════════════════════════════════════════════════

class TableIntegrityRule(BaseRule):
    """Verify table grid/cell width consistency.

    Checks that tblGrid/gridCol widths match first-row tcW values.
    Uses a 4% tolerance based on empirical testing with
    cross-platform rendering (Word, LibreOffice, WPS).
    """

    name = "table-integrity"
    severity = "error"

    _TOLERANCE = 0.04  # 4% width mismatch allowed (empirical threshold)

    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        findings: list[DocumentFinding] = []
        root = context.doc_root
        if root is None:
            return findings

        for t_idx, tbl in enumerate(root.iter(NS.wtag("tbl")), 1):
            grid = tbl.find(NS.wtag("tblGrid"))
            if grid is None:
                findings.append(self._finding(
                    "tblGrid absent - layout may collapse",
                    location=f"TABLE[{t_idx}]",
                ))
                continue

            col_widths: list[int | None] = []
            for gc in grid.findall(NS.wtag("gridCol")):
                raw = gc.get(f"{{{NS.MAIN}}}w")
                col_widths.append(int(raw) if raw else None)

            first_row = tbl.find(NS.wtag("tr"))
            if first_row is None:
                continue

            for ci, cell in enumerate(first_row.findall(NS.wtag("tc"))):
                if ci >= len(col_widths):
                    break
                expected = col_widths[ci]
                if expected is None:
                    continue

                tc_pr = cell.find(NS.wtag("tcPr"))
                if tc_pr is None:
                    continue
                tc_w = tc_pr.find(NS.wtag("tcW"))
                if tc_w is None:
                    continue

                raw_w = tc_w.get(f"{{{NS.MAIN}}}w")
                if not raw_w:
                    continue
                actual = int(raw_w)
                if expected > 0 and abs(actual - expected) > expected * self._TOLERANCE:
                    findings.append(self._finding(
                        f"col[{ci}] gridCol={expected} tcW={actual} - width mismatch",
                        location=f"TABLE[{t_idx}]",
                    ))

        return findings


class MediaProportionRule(BaseRule):
    """Check that images maintain their source aspect ratio.

    Flags images where display extent distorts the source aspect ratio
    beyond acceptable tolerance. The 2.5% threshold catches visible
    distortion while allowing minor rounding in EMU conversions.
    """

    name = "media-proportion"
    severity = "warning"

    _TOLERANCE = 0.025  # 2.5% aspect ratio drift (EMU rounding tolerance)

    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        findings: list[DocumentFinding] = []
        root = context.doc_root
        if root is None:
            return findings

        for idx, drawing in enumerate(root.iter(NS.wtag("drawing")), 1):
            extent = drawing.find(f".//{{{NS.DML_WP}}}extent")
            if extent is None:
                continue
            cx_raw, cy_raw = extent.get("cx"), extent.get("cy")
            if not cx_raw or not cy_raw:
                continue
            cx, cy = int(cx_raw), int(cy_raw)
            if cy == 0:
                continue
            display_ratio = cx / cy

            blip = drawing.find(f".//{{{NS.DML_MAIN}}}blip")
            if blip is None:
                continue
            rid = blip.get(f"{{{NS.REL}}}embed")
            if not rid or rid not in context.rel_index:
                continue

            img_path = context.rel_index[rid]
            if not img_path.exists():
                continue

            src_w, src_h = self._image_dimensions(img_path.read_bytes())
            if src_w is None or src_h is None or src_h == 0:
                continue

            source_ratio = src_w / src_h
            if abs(display_ratio - source_ratio) / source_ratio > self._TOLERANCE:
                findings.append(self._finding(
                    f"display={display_ratio:.2f} source={source_ratio:.2f} - aspect ratio distorted",
                    location=f"IMAGE[{idx}] {img_path.name}",
                    level=FindingLevel.WARNING,
                ))

        return findings

    @staticmethod
    def _image_dimensions(blob: bytes) -> tuple[int | None, int | None]:
        """Return (width, height) using Pillow library.

        Pillow provides reliable format detection and dimension extraction
        for PNG, JPEG, GIF, BMP, TIFF and other common formats.
        Falls back gracefully if the library is unavailable.
        """
        try:
            from PIL import Image
            with Image.open(io.BytesIO(blob)) as img:
                return img.width, img.height
        except ImportError:
            # Pillow not installed - skip dimension check
            return None, None
        except Exception:
            # Corrupted or unsupported image format
            return None, None


class AnnotationConsistencyRule(BaseRule):
    """Verify comment companion files exist when needed.

    When threaded comments are present (with paraId attributes),
    commentsExtended.xml and commentsIds.xml must also exist.
    """

    name = "annotation-consistency"
    severity = "error"

    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        findings: list[DocumentFinding] = []
        comments_xml = context.word_dir / "comments.xml"
        if not comments_xml.exists():
            return findings

        root = ET.parse(comments_xml).getroot()

        # Check for threaded comments (have paraId attribute)
        threaded = False
        for comment in root.iter(NS.wtag("comment")):
            for para in comment.iter(NS.wtag("p")):
                if para.get(f"{{{NS.EXT_2010}}}paraId"):
                    threaded = True
                    break
            if threaded:
                break

        if not threaded:
            return findings

        for companion in ("commentsExtended.xml", "commentsIds.xml"):
            if not (context.word_dir / companion).exists():
                findings.append(self._finding(
                    f"{companion} needed by threaded comments but absent",
                    location="ANNOTATIONS",
                ))

        return findings


class LayoutBoundaryRule(BaseRule):
    """Check for dangerously tight margins.

    Warns when all four margins in any section are below a safe minimum.
    """

    name = "layout-boundary"
    severity = "info"

    _MIN_MARGIN_TWIPS = 288  # ~0.2 in

    def check(self, context: DocumentContext) -> list[DocumentFinding]:
        findings: list[DocumentFinding] = []
        root = context.doc_root
        if root is None:
            return findings

        body = root.find(f".//{{{NS.MAIN}}}body")
        if body is None:
            return findings

        sections: list[tuple[str, ET.Element]] = []
        for p in body.iter(NS.wtag("p")):
            ppr = p.find(NS.wtag("pPr"))
            if ppr is not None:
                sp = ppr.find(NS.wtag("sectPr"))
                if sp is not None:
                    sections.append(("mid", sp))

        tail_sp = body.find(NS.wtag("sectPr"))
        if tail_sp is not None:
            sections.append(("tail", tail_sp))

        if len(sections) < 2:
            return findings

        for pos, (label, sp) in enumerate(sections):
            vals = self._margins(sp)
            if vals is None:
                continue
            if all(v < self._MIN_MARGIN_TWIPS for v in vals.values()):
                if label == "tail" and len(sections) > 2:
                    findings.append(self._finding(
                        "terminal section margins near zero - content may bleed",
                        location="LAYOUT",
                        level=FindingLevel.INFO,
                    ))
                elif pos > 0:
                    findings.append(self._finding(
                        f"section[{pos + 1}] margins too tight (top={vals['top']} left={vals['left']}) - clipping risk",
                        location="LAYOUT",
                        level=FindingLevel.INFO,
                    ))

        return findings

    def _margins(self, sp: ET.Element) -> dict[str, int] | None:
        m = sp.find(NS.wtag("pgMar"))
        if m is None:
            return None
        try:
            return {
                side: abs(int(m.get(f"{{{NS.MAIN}}}{side}", "1440")))
                for side in ("top", "bottom", "left", "right")
            }
        except ValueError:
            return None


# ═══════════════════════════════════════════════════════════════════════════
# Rule Engine
# ═══════════════════════════════════════════════════════════════════════════

class RuleEngine:
    """Coordinate execution of multiple validation rules.

    Supports rule registration, selective execution, and
    aggregated result collection.
    """

    def __init__(self):
        self._rules: list[ValidationRule] = []
        self._disabled: set[str] = set()

    def register(self, rule: ValidationRule) -> "RuleEngine":
        """Register a rule with the engine."""
        self._rules.append(rule)
        return self

    def disable(self, rule_name: str) -> "RuleEngine":
        """Disable a rule by name."""
        self._disabled.add(rule_name)
        return self

    def enable(self, rule_name: str) -> "RuleEngine":
        """Re-enable a previously disabled rule."""
        self._disabled.discard(rule_name)
        return self

    def run(self, context: DocumentContext) -> list[DocumentFinding]:
        """Execute all enabled rules and collect findings."""
        findings: list[DocumentFinding] = []
        for rule in self._rules:
            if rule.name in self._disabled:
                continue
            findings.extend(rule.check(context))
        return findings

    @classmethod
    def default_engine(cls) -> "RuleEngine":
        """Create an engine with all built-in rules registered."""
        engine = cls()
        engine.register(TableIntegrityRule())
        engine.register(MediaProportionRule())
        engine.register(AnnotationConsistencyRule())
        engine.register(LayoutBoundaryRule())
        return engine
