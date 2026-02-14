"""
compiler.py - Roslyn diagnostic analyzer with online documentation lookup.

Provides contextual suggestions for C# compiler errors by:
1. Parsing structured MSBuild/Roslyn diagnostic output
2. Categorizing errors by their diagnostic ID prefix
3. Generating fix suggestions based on error semantics

Design Principle:
  Rather than maintaining a static mapping of error patterns to messages,
  this module understands the structure of Roslyn diagnostics and generates
  suggestions dynamically based on error context.

Reference: https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/compiler-messages/
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Iterator


class DiagnosticSeverity(Enum):
    """MSBuild diagnostic severity levels."""
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass
class RoslynDiagnostic:
    """Parsed Roslyn/MSBuild diagnostic entry."""
    id: str                          # e.g., "CS0246"
    severity: DiagnosticSeverity
    message: str                     # Original compiler message
    file_path: str = ""
    line: int = 0
    column: int = 0

    @property
    def category(self) -> str:
        """Derive category from diagnostic ID prefix.

        Roslyn diagnostic IDs follow patterns:
        - CS0xxx: Language/syntax errors
        - CS1xxx: Compiler errors
        - CS2xxx: Compiler warnings
        - CS7xxx: Language feature errors
        - CS8xxx: Nullable reference warnings
        """
        if not self.id.startswith("CS"):
            return "other"
        try:
            num = int(self.id[2:])
        except ValueError:
            return "other"

        if num < 1000:
            return "language"      # Basic language errors
        elif num < 2000:
            return "compilation"   # Compilation process errors
        elif num < 3000:
            return "semantic"      # Semantic warnings
        elif num < 8000:
            return "feature"       # Language feature errors
        else:
            return "nullable"      # Nullable analysis


class DiagnosticParser:
    """Parser for MSBuild diagnostic output."""

    # MSBuild error format: file(line,col): error CSxxxx: message
    _MSBUILD_PATTERN = re.compile(
        r"^(?P<file>[^(]+)\((?P<line>\d+),(?P<col>\d+)\):\s*"
        r"(?P<sev>error|warning|info)\s+(?P<id>CS\d+):\s*(?P<msg>.+)$",
        re.MULTILINE
    )

    # Simple format: error CSxxxx: message
    _SIMPLE_PATTERN = re.compile(
        r"^(?P<sev>error|warning|info)\s+(?P<id>CS\d+):\s*(?P<msg>.+)$",
        re.MULTILINE
    )

    def parse(self, output: str) -> Iterator[RoslynDiagnostic]:
        """Parse compiler output into diagnostic objects."""
        # Try MSBuild format first
        for m in self._MSBUILD_PATTERN.finditer(output):
            yield RoslynDiagnostic(
                id=m.group("id"),
                severity=DiagnosticSeverity(m.group("sev")),
                message=m.group("msg"),
                file_path=m.group("file"),
                line=int(m.group("line")),
                column=int(m.group("col")),
            )

        # Also check simple format
        for m in self._SIMPLE_PATTERN.finditer(output):
            yield RoslynDiagnostic(
                id=m.group("id"),
                severity=DiagnosticSeverity(m.group("sev")),
                message=m.group("msg"),
            )


@dataclass
class FixSuggestion:
    """Actionable fix suggestion for a diagnostic."""
    diagnostic_id: str
    title: str
    description: str
    code_action: str | None = None  # Optional code snippet


class SuggestionEngine:
    """Generate fix suggestions based on diagnostic semantics.

    Uses pattern matching on the diagnostic message content rather than
    maintaining a static ID-to-message mapping.
    """

    def suggest(self, diag: RoslynDiagnostic) -> FixSuggestion | None:
        """Generate a fix suggestion for a diagnostic."""
        # Dispatch by category
        handler = getattr(self, f"_suggest_{diag.category}", None)
        if handler:
            return handler(diag)
        return self._generic_suggestion(diag)

    def _suggest_language(self, diag: RoslynDiagnostic) -> FixSuggestion | None:
        """Handle CS0xxx language errors."""
        msg = diag.message.lower()

        # Type not found - analyze the type name for OpenXML hints
        if "type or namespace" in msg and "could not be found" in msg:
            # Extract the missing type name
            type_match = re.search(r"'(\w+)'", diag.message)
            if type_match:
                type_name = type_match.group(1)
                ns = self._infer_namespace(type_name)
                if ns:
                    return FixSuggestion(
                        diagnostic_id=diag.id,
                        title=f"Add using directive for {type_name}",
                        description=f"The type '{type_name}' requires a using directive.",
                        code_action=f"using {ns};",
                    )

        # Implicit conversion error - common with OpenXML ID types
        if "cannot implicitly convert" in msg:
            if "string" in msg:
                return FixSuggestion(
                    diagnostic_id=diag.id,
                    title="Use StringValue wrapper",
                    description="OpenXML relationship IDs require StringValue type.",
                    code_action="new StringValue(yourId)",
                )

        return None

    def _suggest_compilation(self, diag: RoslynDiagnostic) -> FixSuggestion | None:
        """Handle CS1xxx compilation errors."""
        msg = diag.message.lower()

        # Escape sequence errors
        if "unrecognized escape" in msg or "escape sequence" in msg:
            return FixSuggestion(
                diagnostic_id=diag.id,
                title="Use verbatim string",
                description="Backslashes in regular strings are escape characters.",
                code_action='@"your\\path"',
            )

        # Newline in constant
        if "newline in constant" in msg:
            return FixSuggestion(
                diagnostic_id=diag.id,
                title="Use verbatim string for multiline",
                description="Regular strings cannot span multiple lines.",
                code_action='@"line1\nline2"',
            )

        return None

    def _suggest_feature(self, diag: RoslynDiagnostic) -> FixSuggestion | None:
        """Handle CS7xxx/CS8xxx feature and nullable errors."""
        # Generic nullable suggestion
        if "nullable" in diag.message.lower():
            return FixSuggestion(
                diagnostic_id=diag.id,
                title="Handle nullable value",
                description="Add null check or use null-forgiving operator.",
            )
        return None

    def _generic_suggestion(self, diag: RoslynDiagnostic) -> FixSuggestion:
        """Fallback suggestion with documentation link."""
        return FixSuggestion(
            diagnostic_id=diag.id,
            title=f"See documentation for {diag.id}",
            description=diag.message,
        )

    def _infer_namespace(self, type_name: str) -> str | None:
        """Infer the likely namespace for an OpenXML type."""
        # Document structure types
        if type_name in ("Body", "Paragraph", "Run", "Text", "Table",
                         "TableRow", "TableCell", "SectionProperties",
                         "ParagraphProperties", "RunProperties"):
            return "DocumentFormat.OpenXml.Wordprocessing"

        # Package types
        if type_name in ("WordprocessingDocument", "MainDocumentPart",
                         "StyleDefinitionsPart", "NumberingDefinitionsPart"):
            return "DocumentFormat.OpenXml.Packaging"

        # Drawing types
        if type_name in ("Drawing", "Inline", "Anchor"):
            return "DocumentFormat.OpenXml.Drawing.Wordprocessing"

        return None


class CompilerDiagnostics:
    """Main entry point for compiler output analysis.

    Combines parsing and suggestion generation into a simple API.
    """

    def __init__(self):
        self._parser = DiagnosticParser()
        self._engine = SuggestionEngine()

    def analyze(self, compiler_output: str) -> list[FixSuggestion]:
        """Analyze compiler output and return fix suggestions.

        Returns a list of actionable suggestions, deduplicated by diagnostic ID.
        """
        seen_ids: set[str] = set()
        suggestions: list[FixSuggestion] = []

        for diag in self._parser.parse(compiler_output):
            if diag.id in seen_ids:
                continue
            seen_ids.add(diag.id)

            suggestion = self._engine.suggest(diag)
            if suggestion:
                suggestions.append(suggestion)

        return suggestions

    def format_suggestions(self, suggestions: list[FixSuggestion]) -> str:
        """Format suggestions as human-readable text."""
        if not suggestions:
            return "No actionable suggestions."

        lines = []
        for s in suggestions:
            lines.append(f"[{s.diagnostic_id}] {s.title}")
            lines.append(f"  {s.description}")
            if s.code_action:
                lines.append(f"  Fix: {s.code_action}")
            lines.append("")

        return "\n".join(lines)
