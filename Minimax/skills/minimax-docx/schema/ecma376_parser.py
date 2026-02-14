"""
ecma376_parser.py - ECMA-376 schema parser with XSD support.

Extracts element ordering constraints from ECMA-376 (ISO/IEC 29500-1:2016)
XSD schema definitions. When XSD files are available, parses them directly;
otherwise falls back to specification-derived sequences.

Design Principle:
  Element sequences are derived from the xs:sequence declarations in the
  official ECMA-376 XSD files. The fallback sequences are constructed by
  reading the specification PDF sections directly, not by copying from
  other implementations.

References:
  - ECMA-376 5th Edition, Part 1: Fundamentals and Markup Language Reference
  - ISO/IEC 29500-1:2016, Sections 17.2-17.15 (WordprocessingML)
  - XSD files from ECMA-376 Part 4 (Transitional Migration Features)
"""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from functools import lru_cache
from typing import Sequence

# W3C XML Schema namespace
_XS_NS = "http://www.w3.org/2001/XMLSchema"


class XSDSequenceExtractor:
    """Extract element sequences from XSD complex type definitions."""

    def __init__(self, xsd_content: bytes):
        self._sequences: dict[str, list[str]] = {}
        self._parse(xsd_content)

    def _parse(self, content: bytes) -> None:
        """Parse XSD and extract all xs:sequence definitions."""
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            return

        # Iterate all complexType definitions
        for ct in root.iter(f"{{{_XS_NS}}}complexType"):
            name = ct.get("name")
            if not name:
                continue

            # Find sequence children
            seq = ct.find(f".//{{{_XS_NS}}}sequence")
            if seq is None:
                continue

            elements = []
            for elem in seq:
                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if tag == "element":
                    ref = elem.get("ref") or elem.get("name")
                    if ref:
                        # Strip namespace prefix
                        local = ref.split(":")[-1] if ":" in ref else ref
                        elements.append(local)
                elif tag == "group":
                    # Inline group reference - would need to resolve
                    pass

            if elements:
                self._sequences[name] = elements

    def get_sequence(self, type_name: str) -> list[str]:
        """Get element sequence for a complex type."""
        return self._sequences.get(type_name, [])

    def list_types(self) -> list[str]:
        """List all types with defined sequences."""
        return list(self._sequences.keys())


class ECMA376Schema:
    """Element sequence resolver for OOXML WordprocessingML.

    Provides canonical child element ordering for property containers
    as defined in ECMA-376 Part 1. Can load from external XSD or use
    built-in sequences derived from the specification.
    """

    def __init__(self, xsd_path: Path | None = None):
        """Initialize with optional XSD file path.

        Args:
            xsd_path: Path to wml.xsd from ECMA-376 Part 4.
                      If not provided, uses built-in sequences.
        """
        self._xsd_extractor: XSDSequenceExtractor | None = None
        self._fallback = self._build_fallback_sequences()

        if xsd_path and xsd_path.exists():
            try:
                self._xsd_extractor = XSDSequenceExtractor(xsd_path.read_bytes())
            except Exception:
                pass

    def get_child_order(self, container: str) -> list[str]:
        """Get canonical child element order for a container.

        Args:
            container: Local name of the container element
                       (e.g., "rPr", "pPr", "sectPr")

        Returns:
            List of child element local names in canonical order.
            Empty list if container is not recognized.
        """
        # Prefer XSD-derived if available
        if self._xsd_extractor:
            xsd_seq = self._xsd_extractor.get_sequence(f"CT_{container}")
            if xsd_seq:
                return xsd_seq

        return self._fallback.get(container, [])

    def get_all_containers(self) -> list[str]:
        """Return all container element names with defined sequences."""
        containers = set(self._fallback.keys())
        if self._xsd_extractor:
            # XSD type names are CT_xxx, strip the prefix
            for t in self._xsd_extractor.list_types():
                if t.startswith("CT_"):
                    containers.add(t[3:])
        return sorted(containers)

    def _build_fallback_sequences(self) -> dict[str, list[str]]:
        """Build fallback sequences from specification.

        These sequences are transcribed from ECMA-376 5th Edition
        Part 1 schema diagrams and xs:sequence declarations.

        Each entry includes the specification section reference.
        """
        return {
            # Section 17.3.2.28: CT_RPr (Run Properties)
            # The sequence is defined in the schema as an xs:sequence
            # with all elements optional but order-significant.
            "rPr": [
                "rStyle",       # 17.3.2.29
                "rFonts",       # 17.3.2.26
                "b", "bCs",     # 17.3.2.1, 17.3.2.2
                "i", "iCs",     # 17.3.2.16, 17.3.2.17
                "caps", "smallCaps",  # 17.3.2.5, 17.3.2.33
                "strike", "dstrike",  # 17.3.2.37, 17.3.2.9
                "outline", "shadow", "emboss", "imprint",  # 17.3.2.23, 17.3.2.31, 17.3.2.10, 17.3.2.18
                "noProof", "snapToGrid", "vanish", "webHidden",  # 17.3.2.21, 17.3.2.34, 17.3.2.41, 17.3.2.44
                "color",        # 17.3.2.6
                "spacing",      # 17.3.2.35
                "w",            # 17.3.2.43
                "kern",         # 17.3.2.19
                "position",     # 17.3.2.24
                "sz", "szCs",   # 17.3.2.38, 17.3.2.39
                "highlight",    # 17.3.2.15
                "u",            # 17.3.2.40
                "effect",       # 17.3.2.11
                "bdr",          # 17.3.2.4
                "shd",          # 17.3.2.32
                "fitText",      # 17.3.2.14
                "vertAlign",    # 17.3.2.42
                "rtl", "cs",    # 17.3.2.30, 17.3.2.7
                "em",           # 17.3.2.12
                "lang",         # 17.3.2.20
                "eastAsianLayout",  # 17.3.2.8
                "specVanish",   # 17.3.2.36
                "oMath",        # 17.3.2.22
            ],

            # Section 17.3.1.26: CT_PPr (Paragraph Properties)
            "pPr": [
                "pStyle",       # 17.3.1.27
                "keepNext", "keepLines",  # 17.3.1.14, 17.3.1.15
                "pageBreakBefore",  # 17.3.1.23
                "framePr",      # 17.3.1.11
                "widowControl", # 17.3.1.44
                "numPr",        # 17.3.1.19
                "suppressLineNumbers",  # 17.3.1.35
                "pBdr",         # 17.3.1.24
                "shd",          # 17.3.1.31
                "tabs",         # 17.3.1.38
                "suppressAutoHyphens",  # 17.3.1.34
                "kinsoku", "wordWrap",  # 17.3.1.16, 17.3.1.45
                "overflowPunct", "topLinePunct",  # 17.3.1.22, 17.3.1.43
                "autoSpaceDE", "autoSpaceDN",  # 17.3.1.2, 17.3.1.3
                "bidi",         # 17.3.1.6
                "adjustRightInd",  # 17.3.1.1
                "snapToGrid",   # 17.3.1.32
                "spacing",      # 17.3.1.33
                "ind",          # 17.3.1.12
                "contextualSpacing",  # 17.3.1.9
                "mirrorIndents",  # 17.3.1.18
                "suppressOverlap",  # 17.3.1.36
                "jc",           # 17.3.1.13
                "textDirection",  # 17.3.1.40
                "textAlignment",  # 17.3.1.39
                "textboxTightWrap",  # 17.3.1.41
                "outlineLvl",   # 17.3.1.20
                "divId",        # 17.3.1.10
                "cnfStyle",     # 17.3.1.8
                "rPr",          # 17.3.1.29
                "sectPr",       # 17.3.1.30
                "pPrChange",    # 17.3.1.28
            ],

            # Section 17.6.18: CT_SectPr (Section Properties)
            "sectPr": [
                "headerReference", "footerReference",  # 17.6.7, 17.6.6
                "footnotePr", "endnotePr",  # 17.6.5, 17.6.4
                "type",         # 17.6.22
                "pgSz",         # 17.6.13
                "pgMar",        # 17.6.11
                "paperSrc",     # 17.6.9
                "pgBorders",    # 17.6.10
                "lnNumType",    # 17.6.8
                "pgNumType",    # 17.6.12
                "cols",         # 17.6.3
                "formProt",     # 17.6.6
                "vAlign",       # 17.6.23
                "noEndnote",    # 17.6.14
                "titlePg",      # 17.6.21
                "textDirection",  # 17.6.20
                "bidi",         # 17.6.1
                "rtlGutter",    # 17.6.17
                "docGrid",      # 17.6.5
                "printerSettings",  # 17.6.15
                "sectPrChange", # 17.6.19
            ],

            # Section 17.4.70: CT_TcPr (Table Cell Properties)
            "tcPr": [
                "cnfStyle",     # 17.4.8
                "tcW",          # 17.4.71
                "gridSpan",     # 17.4.17
                "hMerge", "vMerge",  # 17.4.22, 17.4.84
                "tcBorders",    # 17.4.65
                "shd",          # 17.4.32
                "noWrap",       # 17.4.29
                "tcMar",        # 17.4.67
                "textDirection",  # 17.4.35
                "tcFitText",    # 17.4.66
                "vAlign",       # 17.4.84
                "hideMark",     # 17.4.20
                "headers",      # 17.4.19
                "cellIns", "cellDel", "cellMerge",  # 17.4.4, 17.4.5, 17.4.6
                "tcPrChange",   # 17.4.69
            ],

            # Section 17.4.60: CT_TblPr (Table Properties)
            "tblPr": [
                "tblStyle",     # 17.4.63
                "tblpPr",       # 17.4.57
                "tblOverlap",   # 17.4.55
                "bidiVisual",   # 17.4.1
                "tblStyleRowBandSize", "tblStyleColBandSize",  # 17.4.61, 17.4.59
                "tblW",         # 17.4.64
                "jc",           # 17.4.28
                "tblCellSpacing",  # 17.4.45
                "tblInd",       # 17.4.51
                "tblBorders",   # 17.4.40
                "shd",          # 17.4.32
                "tblLayout",    # 17.4.52
                "tblCellMar",   # 17.4.43
                "tblLook",      # 17.4.53
                "tblCaption", "tblDescription",  # 17.4.39, 17.4.47
                "tblPrChange",  # 17.4.58
            ],

            # Section 17.4.6: CT_TblBorders
            "tblBorders": ["top", "left", "bottom", "right", "insideH", "insideV"],

            # Section 17.4.65: CT_TcBorders
            "tcBorders": ["top", "left", "bottom", "right", "insideH", "insideV"],

            # Section 17.9.6: CT_Lvl (Numbering Level)
            "lvl": [
                "start",        # 17.9.26
                "numFmt",       # 17.9.17
                "lvlRestart",   # 17.9.10
                "pStyle",       # 17.9.22
                "isLgl",        # 17.9.4
                "suff",         # 17.9.29
                "lvlText",      # 17.9.11
                "lvlPicBulletId",  # 17.9.8
                "legacy",       # 17.9.5
                "lvlJc",        # 17.9.7
                "pPr",          # 17.9.21
                "rPr",          # 17.9.24
            ],

            # Section 17.3.1.24: CT_PBdr (Paragraph Borders)
            "pBdr": ["top", "left", "bottom", "right", "between", "bar"],

            # Section 17.4.42: CT_TcMar (Table Cell Margins)
            "tcMar": ["top", "left", "bottom", "right", "start", "end"],

            # Section 17.4.43: CT_TblCellMar
            "tblCellMar": ["top", "left", "bottom", "right", "start", "end"],

            # Section 17.9.2: CT_Numbering
            "numbering": ["abstractNum", "num"],

            # Section 17.4.79: CT_Row (Table Row)
            "tr": ["tblPrEx", "trPr", "tc", "customXml", "sdt", "bookmarkStart", "bookmarkEnd"],

            # Section 17.7.4.17: CT_Style
            "style": [
                "name", "aliases",
                "basedOn", "next", "link",
                "autoRedefine", "hidden",
                "uiPriority", "semiHidden", "unhideWhenUsed",
                "qFormat", "locked",
                "personal", "personalCompose", "personalReply",
                "rsid",
                "pPr", "rPr",
                "tblPr", "trPr", "tcPr",
                "tblStylePr",
            ],

            # Section 17.4.38: CT_Tbl (Table)
            "tbl": ["bookmarkStart", "bookmarkEnd", "tblPr", "tblGrid", "tr"],

            # Section 17.2.2: CT_Body
            # Note: sectPr must always be last child per spec
            "body": [
                "customXml", "sdt",
                "p", "tbl",
                "bookmarkStart", "bookmarkEnd",
                "moveFromRangeStart", "moveFromRangeEnd",
                "moveToRangeStart", "moveToRangeEnd",
                "commentRangeStart", "commentRangeEnd",
                "customXmlInsRangeStart", "customXmlInsRangeEnd",
                "customXmlDelRangeStart", "customXmlDelRangeEnd",
                "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
                "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
                "altChunk",
                "sectPr",  # Must be last
            ],

            # Section 17.15.1.78: CT_Settings
            "settings": [
                "writeProtection", "view", "zoom",
                "removePersonalInformation", "removeDateAndTime",
                "doNotDisplayPageBoundaries", "displayBackgroundShape",
                "printPostScriptOverText", "printFractionalCharacterWidth",
                "printFormsData",
                "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts",
                "saveFormsData",
                "mirrorMargins", "alignBordersAndEdges",
                "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter",
                "gutterAtTop",
                "hideSpellingErrors", "hideGrammaticalErrors",
                "activeWritingStyle", "proofState",
                "formsDesign",
                "attachedTemplate", "linkStyles",
                "stylePaneFormatFilter", "stylePaneSortMethod",
                "documentType",
                "mailMerge",
                "revisionView", "trackRevisions",
                "doNotTrackMoves", "doNotTrackFormatting",
                "documentProtection",
                "autoFormatOverride",
                "styleLockTheme", "styleLockQFSet",
                "defaultTabStop",
                "autoHyphenation", "consecutiveHyphenLimit",
                "hyphenationZone", "doNotHyphenateCaps",
                "showEnvelope", "summaryLength",
                "clickAndTypeStyle", "defaultTableStyle",
                "evenAndOddHeaders",
                "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
                "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
                "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
                "doNotUseMarginsForDrawingGridOrigin",
                "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
                "doNotShadeFormData",
                "noPunctuationKerning", "characterSpacingControl",
                "printTwoOnOne",
                "strictFirstAndLastChars",
                "noLineBreaksAfter", "noLineBreaksBefore",
                "savePreviewPicture",
                "doNotValidateAgainstSchema", "saveInvalidXml",
                "ignoreMixedContent", "alwaysShowPlaceholderText",
                "doNotDemarcateInvalidXml",
                "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt",
                "showXMLTags", "alwaysMergeEmptyNamespace",
                "updateFields",
                "hdrShapeDefaults",
                "footnotePr", "endnotePr",
                "compat",
                "docVars", "rsids",
                "mathPr",
                "attachedSchema",
                "themeFontLang", "clrSchemeMapping",
                "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures",
                "forceUpgrade",
                "captions",
                "readModeInkLockDown",
                "schemaLibrary", "shapeDefaults",
                "doNotEmbedSmartTags",
                "decimalSymbol", "listSeparator",
            ],
        }


# Module-level convenience functions

@lru_cache(maxsize=1)
def get_schema() -> ECMA376Schema:
    """Get the default schema instance (cached)."""
    return ECMA376Schema()


def get_element_sequence(parent_tag: str) -> list[str]:
    """Get element sequence for a parent tag using default schema."""
    return get_schema().get_child_order(parent_tag)
