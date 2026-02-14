"""
findings.py - Structured validation findings.

Provides a unified representation for document quality issues
discovered during validation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class FindingLevel(Enum):
    """Severity level for validation findings."""

    ERROR = "error"  # Must be fixed for valid document
    WARNING = "warning"  # Should be reviewed, may cause issues
    INFO = "info"  # Informational, no action required


@dataclass
class DocumentFinding:
    """Structured validation finding.

    Attributes:
        rule_name: Identifier of the rule that generated this finding
        level: Severity level
        message: Human-readable description
        location: Optional location within document (e.g., "TABLE[2]", "IMAGE[5]")
        details: Optional additional context
    """

    rule_name: str
    level: FindingLevel
    message: str
    location: Optional[str] = None
    details: dict = field(default_factory=dict)

    def __str__(self) -> str:
        prefix = f"[{self.level.value.upper()}]"
        loc = f" {self.location}:" if self.location else ""
        return f"{prefix}{loc} {self.message}"

    def to_legacy_format(self) -> str:
        """Convert to legacy string format for backward compatibility."""
        loc = f"{self.location}: " if self.location else ""
        return f"{loc}{self.message}"

    @property
    def is_blocking(self) -> bool:
        """Return True if this finding should block document acceptance."""
        return self.level == FindingLevel.ERROR
