"""
xmlns.py - OOXML namespace definitions.

Provides namespace URIs from the ECMA-376 / ISO 29500 standard via
a lightweight accessor object.  Every OpenXML-compliant tool shares
these same URIs; they are part of the public specification.

Reference: ECMA-376 5th Edition, Part 1 Annex A
"""

from xml.etree import ElementTree as _ET


class OoxmlNs:
    """Central namespace registry with Clark-notation tag builder."""

    MAIN       = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    EXT_2010   = "http://schemas.microsoft.com/office/word/2010/wordml"
    EXT_2012   = "http://schemas.microsoft.com/office/word/2012/wordml"
    REL        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    DML_WP     = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    DML_MAIN   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    PKG_REL    = "http://schemas.openxmlformats.org/package/2006/relationships"

    _PREFIXES = {
        "w":      MAIN,
        "w14":    EXT_2010,
        "w15":    EXT_2012,
        "r":      REL,
        "wp":     DML_WP,
        "a":      DML_MAIN,
        "mc":     "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "pic":    "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "m":      "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "v":      "urn:schemas-microsoft-com:vml",
        "o":      "urn:schemas-microsoft-com:office:office",
        "wpc":    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "wpg":    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "wps":    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "w16se":  "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    }

    def __init__(self):
        for pfx, uri in self._PREFIXES.items():
            _ET.register_namespace(pfx, uri)

    # ── tag helpers ──────────────────────────────────────────────────────

    def tag(self, prefix: str, local: str) -> str:
        """Build a Clark-notation tag: ``{uri}local``."""
        uri = self._PREFIXES.get(prefix)
        if uri is None:
            raise KeyError(f"unknown namespace prefix: {prefix!r}")
        return f"{{{uri}}}{local}"

    def wtag(self, local: str) -> str:
        """Shorthand for ``tag('w', local)``."""
        return f"{{{self.MAIN}}}{local}"

    @property
    def xpath_map(self) -> dict[str, str]:
        """Prefix → URI mapping suitable for ElementTree XPath."""
        return dict(self._PREFIXES)


NS = OoxmlNs()
