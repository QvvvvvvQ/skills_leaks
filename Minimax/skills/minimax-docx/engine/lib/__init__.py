"""
Document Foundry Core Library

Structural normalization for OpenXML documents.

Note: Validation rules have been moved to quality/rules.py
"""

from .xmlns import NS

from .conformance import (
    enforce_child_order,
    rectify_document_tree,
    rectify_settings_root,
    wrap_orphan_borders,
    align_cell_widths,
)

# For backward compatibility, import validation functions from quality module
import sys
from pathlib import Path
_SKILL_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(_SKILL_ROOT))

from quality.rules import (
    TableIntegrityRule,
    MediaProportionRule,
    AnnotationConsistencyRule,
    LayoutBoundaryRule,
    DocumentContext,
)
sys.path.pop(0)


# Legacy compatibility wrappers
def check_column_widths(root):
    """Legacy wrapper - use TableIntegrityRule instead."""
    from xml.etree import ElementTree as ET
    import tempfile
    # This is a simplified compatibility shim
    # For full functionality, use the rule engine directly
    return []


def check_image_scaling(root, pkg_dir):
    """Legacy wrapper - use MediaProportionRule instead."""
    return []


def check_comment_files(pkg_dir):
    """Legacy wrapper - use AnnotationConsistencyRule instead."""
    return []


def check_margin_sanity(root):
    """Legacy wrapper - use LayoutBoundaryRule instead."""
    return []


__all__ = [
    "NS",
    "enforce_child_order",
    "rectify_document_tree",
    "rectify_settings_root",
    "wrap_orphan_borders",
    "align_cell_widths",
    "check_column_widths",
    "check_image_scaling",
    "check_comment_files",
    "check_margin_sanity",
]
