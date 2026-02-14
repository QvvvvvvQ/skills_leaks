#!/usr/bin/env python3
"""
inspector.py - Business-rule validation for .docx packages.

Covers checks that the OpenXML SDK schema validator cannot perform:

* Table grid / cell-width consistency
* Image aspect-ratio preservation
* Document-settings hygiene (e.g. TOC updateFields)

Results are printed as structured single-line messages prefixed with the
severity level (ERROR / WARN / OK).

Usage::

    python inspector.py report.docx
"""

import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from lib import (
    NS,
    check_column_widths,
    check_image_scaling,
    check_comment_files,
)


def _verify_toc_settings(word_dir: Path) -> list[str]:
    """Warn if a TOC exists but updateFields is absent in settings.xml."""
    warnings: list[str] = []

    doc_file = word_dir / "document.xml"
    settings_file = word_dir / "settings.xml"
    if not doc_file.exists() or not settings_file.exists():
        return warnings

    doc_root = ET.parse(doc_file).getroot()
    instrs = doc_root.findall(f".//{{{NS.MAIN}}}instrText")
    has_toc = any("TOC" in (el.text or "") for el in instrs)
    if not has_toc:
        return warnings

    settings_root = ET.parse(settings_file).getroot()
    if settings_root.find(f".//{{{NS.MAIN}}}updateFields") is None:
        warnings.append(
            "WARN  settings.xml: TOC present but <w:updateFields> is missing; "
            "table of contents will not auto-refresh on open"
        )
    return warnings


def inspect(docx_path: str | Path) -> dict:
    """Run all business-rule checks.  Returns ``{"errors": [...], "warnings": [...]}``."""
    docx_path = Path(docx_path)
    result: dict[str, list[str]] = {"errors": [], "warnings": []}

    if not docx_path.is_file():
        result["errors"].append(f"file not found: {docx_path}")
        return result

    try:
        with tempfile.TemporaryDirectory(prefix="docx_inspect_") as work:
            pkg = Path(work) / "pkg"
            with zipfile.ZipFile(docx_path, "r") as arc:
                arc.extractall(pkg)

            doc_xml = pkg / "word" / "document.xml"
            if not doc_xml.exists():
                result["errors"].append("word/document.xml is missing from the package")
                return result

            root = ET.parse(doc_xml).getroot()

            result["errors"].extend(check_column_widths(root))
            result["errors"].extend(check_image_scaling(root, pkg))
            result["errors"].extend(check_comment_files(pkg))
            result["warnings"].extend(_verify_toc_settings(pkg / "word"))

    except zipfile.BadZipFile:
        result["errors"].append("archive is corrupt or not a valid .docx")
    except Exception as exc:
        result["errors"].append(f"unexpected parse failure: {exc}")

    return result


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python inspector.py <file.docx>", file=sys.stderr)
        raise SystemExit(1)

    report = inspect(sys.argv[1])

    for w in report["warnings"]:
        print(w)
    if report["errors"]:
        for e in report["errors"]:
            print(f"ERROR {e}")
        raise SystemExit(1)
    print("OK    all business-rule checks passed")


if __name__ == "__main__":
    main()
