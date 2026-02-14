"""
opc.py - Content-Type driven OPC packaging.

Implements Open Packaging Conventions (ECMA-376 Part 2) by reading
the actual [Content_Types].xml manifest to determine part ordering.

Design Principle:
  Instead of path-prefix heuristics, we parse the content type manifest
  and order parts based on their semantic role in the package structure.
  This follows the OPC specification's intent of content-type-first design.

Reference: ECMA-376-2:2016, Section 10 (Physical Package)
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from typing import Callable


# OPC Content Types namespace
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


class ContentTypeManifest:
    """Parser for [Content_Types].xml manifest."""

    def __init__(self, content_types_xml: str | bytes | None = None):
        self._defaults: dict[str, str] = {}  # extension -> content-type
        self._overrides: dict[str, str] = {}  # part-name -> content-type

        if content_types_xml:
            self._parse(content_types_xml)

    def _parse(self, xml_content: str | bytes) -> None:
        """Parse content types from XML."""
        try:
            root = ET.fromstring(xml_content if isinstance(xml_content, bytes)
                                 else xml_content.encode("utf-8"))
            for child in root:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if tag == "Default":
                    ext = child.get("Extension", "").lower()
                    ct = child.get("ContentType", "")
                    if ext:
                        self._defaults[ext] = ct
                elif tag == "Override":
                    part = child.get("PartName", "")
                    ct = child.get("ContentType", "")
                    if part:
                        # Normalize: remove leading slash
                        self._overrides[part.lstrip("/")] = ct
        except ET.ParseError:
            pass

    def get_content_type(self, part_name: str) -> str:
        """Get content type for a part, checking overrides then defaults."""
        normalized = part_name.lstrip("/")
        if normalized in self._overrides:
            return self._overrides[normalized]
        ext = Path(normalized).suffix.lstrip(".").lower()
        return self._defaults.get(ext, "application/octet-stream")


class OPCPartClassifier:
    """Semantic classifier for OPC package parts.

    Categorizes parts by their role in the document structure,
    derived from content-type analysis rather than path patterns.
    """

    # Content-type to semantic category mapping
    _CONTENT_TYPE_CATEGORIES = {
        # Structural metadata
        "application/vnd.openxmlformats-package.relationships+xml": "relationships",
        "application/vnd.openxmlformats-package.core-properties+xml": "metadata",
        "application/vnd.openxmlformats-officedocument.extended-properties+xml": "metadata",
        "application/vnd.openxmlformats-officedocument.custom-properties+xml": "metadata",

        # Main document parts
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "main_document",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml": "definitions",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml": "definitions",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml": "definitions",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml": "definitions",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml": "definitions",

        # Themes
        "application/vnd.openxmlformats-officedocument.theme+xml": "theming",

        # Headers/Footers
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml": "page_layout",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml": "page_layout",

        # Annotations
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml": "annotations",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml": "annotations",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml": "annotations",

        # Media
        "image/png": "media",
        "image/jpeg": "media",
        "image/gif": "media",
        "image/tiff": "media",
        "image/x-emf": "media",
        "image/x-wmf": "media",
    }

    # Semantic category ordering (lower = earlier in archive)
    _CATEGORY_ORDER = {
        "manifest": 0,       # [Content_Types].xml
        "relationships": 1,  # .rels files
        "metadata": 2,       # docProps/*
        "main_document": 3,  # Primary content
        "definitions": 4,    # styles, numbering, settings
        "theming": 5,        # theme
        "page_layout": 6,    # headers, footers
        "annotations": 7,    # comments, footnotes
        "media": 8,          # images, embedded objects
        "unknown": 9,        # Fallback
    }

    def __init__(self, manifest: ContentTypeManifest | None = None):
        self._manifest = manifest or ContentTypeManifest()

    def classify(self, part_name: str) -> str:
        """Classify a part into a semantic category."""
        # Special cases that don't depend on content-type
        if part_name == "[Content_Types].xml":
            return "manifest"
        if part_name.endswith(".rels"):
            return "relationships"

        ct = self._manifest.get_content_type(part_name)
        return self._CONTENT_TYPE_CATEGORIES.get(ct, "unknown")

    def sort_key(self, part_name: str) -> tuple[int, int, str]:
        """Generate sort key: (category_order, relationship_depth, path).

        Parts are ordered by:
        1. Semantic category (manifest first, media last)
        2. Relationship depth (root rels before nested rels)
        3. Alphabetical path (stable ordering)
        """
        category = self.classify(part_name)
        order = self._CATEGORY_ORDER.get(category, 99)

        # For relationships, sort by nesting depth
        # _rels/.rels < word/_rels/document.xml.rels < word/_rels/header1.xml.rels
        rel_depth = part_name.count("/") if category == "relationships" else 0

        return (order, rel_depth, part_name)


class OPCPackager:
    """OPC-compliant package assembler using content-type semantics."""

    def repackage(self, source_dir: Path, output_path: Path) -> None:
        """Repackage a directory into a DOCX with semantic ordering.

        Reads [Content_Types].xml to understand part roles,
        then orders parts appropriately.
        """
        ct_path = source_dir / "[Content_Types].xml"
        manifest = None
        if ct_path.exists():
            manifest = ContentTypeManifest(ct_path.read_bytes())

        classifier = OPCPartClassifier(manifest)

        all_files = [f for f in source_dir.rglob("*") if f.is_file()]
        sorted_files = sorted(
            all_files,
            key=lambda f: classifier.sort_key(str(f.relative_to(source_dir)))
        )

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as arc:
            for fpath in sorted_files:
                arc.write(fpath, fpath.relative_to(source_dir))

    def create_package(
        self,
        output_path: Path,
        content_provider: Callable[[str], bytes | None],
        manifest: list[str],
    ) -> None:
        """Create a new DOCX package from content callbacks.

        For new packages, uses heuristic ordering since we don't have
        a pre-existing [Content_Types].xml to analyze.
        """
        classifier = OPCPartClassifier()
        sorted_manifest = sorted(manifest, key=classifier.sort_key)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as arc:
            for path in sorted_manifest:
                content = content_provider(path)
                if content is not None:
                    arc.writestr(path, content)
