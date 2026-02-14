#!/usr/bin/env python3
"""
DOCX Engine - OpenXML Document Compilation Toolkit

Invoke: python docx_engine.py {doctor|render|audit|preview} [options]

Architecture:
- Single orchestration point - no scattered scripts
- Runtime detection with graceful fallbacks
- Self-locating via module path resolution
- Mandatory output verification

Commands:
  doctor          Environment diagnostics and setup
  render [name]   Build and generate document
  audit FILE      Validate existing document
  preview FILE    Quick content preview (pandoc)
"""

import os
import platform
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Optional, Tuple

# Module path resolution
SCRIPT_LOCATION = Path(__file__).parent.resolve()
ENGINE_DIR = SCRIPT_LOCATION / "engine"

# Import diagnostics system
sys.path.insert(0, str(SCRIPT_LOCATION))
from diagnostics.compiler import CompilerDiagnostics


def resolve_project_home() -> Path:
    """Determine active workspace from environment or working directory.

    This MUST resolve to the user's working directory (or PROJECT_HOME),
    never the skill installation directory. All build intermediates and
    outputs are placed here.
    """
    env_path = os.environ.get("PROJECT_HOME")
    home = Path(env_path) if env_path else Path.cwd()
    # Guard: never write build artifacts into the skill directory itself
    if home.resolve() == SCRIPT_LOCATION.resolve():
        raise RuntimeError(
            f"project_home resolved to the skill directory ({SCRIPT_LOCATION}). "
            "Run docx_engine.py from the user's working directory or set PROJECT_HOME."
        )
    return home


def resolve_staging_area() -> Path:
    """Staging directory for intermediate build files (under project home)."""
    return resolve_project_home() / ".docx_workspace"


def resolve_artifact_dir() -> Path:
    """Final output directory for deliverables (under project home)."""
    return resolve_project_home() / "output"


# Runtime Detection
# ============================================================================

def locate_dotnet_binary() -> Optional[Path]:
    """Search for dotnet runtime in common locations."""
    os_type = platform.system()
    search_paths = ["dotnet"]

    if os_type == "Windows":
        search_paths.extend([
            Path.home() / ".dotnet" / "dotnet.exe",
            Path(os.environ.get("ProgramFiles", "")) / "dotnet" / "dotnet.exe",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "dotnet" / "dotnet.exe",
        ])
    else:
        search_paths.extend([
            Path.home() / ".dotnet" / "dotnet",
            Path("/usr/local/share/dotnet/dotnet"),
            Path("/usr/share/dotnet/dotnet"),
            Path("/opt/dotnet/dotnet"),
        ])

    for candidate in search_paths:
        if isinstance(candidate, str):
            found = shutil.which(candidate)
            if found:
                return Path(found)
        elif candidate.exists() and candidate.is_file():
            return candidate
    return None


def assess_runtime_health() -> Tuple[str, Optional[Path], Optional[str]]:
    """
    Evaluate dotnet installation status.
    Returns: (status, binary_path, version_string)
    status: 'ready' | 'outdated' | 'corrupted' | 'absent'
    """
    binary = locate_dotnet_binary()
    if not binary:
        return ("absent", None, None)

    try:
        proc = subprocess.run(
            [str(binary), "--version"],
            capture_output=True, text=True, timeout=10
        )
        if proc.returncode == 0:
            ver = proc.stdout.strip()
            try:
                major_ver = int(ver.split(".")[0])
                return ("ready", binary, ver) if major_ver >= 6 else ("outdated", binary, ver)
            except (ValueError, IndexError):
                return ("corrupted", binary, None)
        return ("corrupted", binary, None)
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return ("corrupted", binary, None)


def provision_dotnet() -> Optional[Path]:
    """Download and install .NET SDK. Returns binary path on success."""
    os_type = platform.system()
    print("  Acquiring .NET SDK...")

    try:
        if os_type == "Windows":
            installer_url = "https://dot.net/v1/dotnet-install.ps1"
            target_dir = Path.home() / ".dotnet"

            powershell_script = f"""
            $ErrorActionPreference = 'Stop'
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $installer = Invoke-WebRequest -Uri '{installer_url}' -UseBasicParsing
            $execution = [scriptblock]::Create($installer.Content)
            & $execution -Channel 8.0 -InstallDir '{target_dir}'
            """
            subprocess.run(
                ["powershell", "-Command", powershell_script],
                capture_output=True, text=True, timeout=300
            )
            binary = target_dir / "dotnet.exe"
        else:
            installer_url = "https://dot.net/v1/dotnet-install.sh"
            target_dir = Path.home() / ".dotnet"

            installer_path = Path(tempfile.gettempdir()) / "dotnet-bootstrap.sh"
            subprocess.run(
                ["curl", "-sSL", installer_url, "-o", str(installer_path)],
                check=True, timeout=60
            )
            installer_path.chmod(0o755)
            subprocess.run(
                [str(installer_path), "--channel", "8.0", "--install-dir", str(target_dir)],
                check=True, timeout=300
            )
            binary = target_dir / "dotnet"

        if binary.exists():
            verify = subprocess.run([str(binary), "--version"], capture_output=True, text=True)
            if verify.returncode == 0:
                print(f"  + Provisioned: {verify.stdout.strip()}")
                return binary

        print("  - Provisioning unsuccessful")
        print("    Reference: https://dotnet.microsoft.com/download")
        return None

    except Exception as exc:
        print(f"  - Provisioning failed: {exc}")
        print("    Reference: https://dotnet.microsoft.com/download")
        return None


def guarantee_dotnet() -> Path:
    """Ensure dotnet availability, installing if necessary. Exits on failure."""
    status, binary, ver = assess_runtime_health()

    if status == "ready":
        return binary
    elif status == "outdated":
        print(f"! Runtime {ver} is outdated (requires 6+), upgrading...")
        result = provision_dotnet()
        if result:
            return result
        sys.exit(1)
    elif status == "corrupted":
        print("! Runtime installation corrupted, reinstalling...")
        dotnet_home = Path.home() / ".dotnet"
        if dotnet_home.exists():
            shutil.rmtree(dotnet_home, ignore_errors=True)
        result = provision_dotnet()
        if result:
            return result
        sys.exit(1)
    else:
        print("o Runtime not detected, installing...")
        result = provision_dotnet()
        if result:
            return result
        sys.exit(1)


def audit_python_dependencies() -> dict:
    """Inventory Python dependency status."""
    inventory = {}

    try:
        import lxml
        inventory["lxml"] = ("available", getattr(lxml, "__version__", "?"))
    except ImportError:
        inventory["lxml"] = ("missing", None)

    pandoc_binary = shutil.which("pandoc")
    if pandoc_binary:
        try:
            proc = subprocess.run(["pandoc", "--version"], capture_output=True, text=True, timeout=5)
            ver = proc.stdout.split("\n")[0].split()[-1] if proc.returncode == 0 else "?"
            inventory["pandoc"] = ("available", ver)
        except Exception:
            inventory["pandoc"] = ("available", "?")
    else:
        inventory["pandoc"] = ("optional", None)

    for pkg in ["playwright", "matplotlib", "PIL"]:
        try:
            __import__(pkg if pkg != "PIL" else "PIL.Image")
            inventory[pkg] = ("available", None)
        except ImportError:
            inventory[pkg] = ("optional", None)

    return inventory


def guarantee_lxml():
    """Install lxml if absent."""
    try:
        import lxml
        return True
    except ImportError:
        print("o lxml not detected, installing...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "lxml"],
                check=True, capture_output=True
            )
            print("  + lxml installed")
            return True
        except subprocess.CalledProcessError:
            print("  - lxml installation failed")
            print("    Manual: pip install lxml")
            return False


# Workspace Management
# ============================================================================

def extract_runtime_version(binary: Path) -> Optional[int]:
    """Parse major version number from dotnet."""
    try:
        proc = subprocess.run([str(binary), "--version"], capture_output=True, text=True, timeout=10)
        if proc.returncode == 0:
            return int(proc.stdout.strip().split(".")[0])
    except Exception:
        pass
    return None


def reconcile_project_framework(proj_path: Path, binary: Path) -> bool:
    """Adjust csproj TargetFramework to match installed runtime."""
    major = extract_runtime_version(binary)
    if not major or major < 6:
        return False

    try:
        import re
        content = proj_path.read_text(encoding="utf-8")
        match = re.search(r"<TargetFramework>net(\d+)\.0</TargetFramework>", content)
        if match:
            current = int(match.group(1))
            if current != major:
                updated = re.sub(
                    r"<TargetFramework>net\d+\.0</TargetFramework>",
                    f"<TargetFramework>net{major}.0</TargetFramework>",
                    content
                )
                proj_path.write_text(updated, encoding="utf-8")
                return True
    except Exception:
        pass
    return False


def prepare_workspace(runtime: Optional[Path] = None):
    """Initialize workspace directories and seed templates."""
    staging = resolve_staging_area()
    output = resolve_artifact_dir()

    staging.mkdir(parents=True, exist_ok=True)
    output.mkdir(parents=True, exist_ok=True)

    blueprints_dir = SCRIPT_LOCATION / "blueprints"

    proj_src = blueprints_dir / "DocumentFoundry.csproj"
    proj_dst = staging / "DocumentFoundry.csproj"
    if not proj_dst.exists() and proj_src.exists():
        shutil.copy2(proj_src, proj_dst)

    entry_src = blueprints_dir / "Launcher.cs"
    entry_dst = staging / "Launcher.cs"
    if not entry_dst.exists() and entry_src.exists():
        shutil.copy2(entry_src, entry_dst)

    if runtime and proj_dst.exists():
        if reconcile_project_framework(proj_dst, runtime):
            major = extract_runtime_version(runtime)
            print(f"  (adjusted TargetFramework to net{major}.0)")


# Verification Pipeline
# ============================================================================

def execute_verification(document_path: Path, runtime: Path) -> bool:
    """Run full verification suite on generated document."""
    # Phase 1: Python-based structural validation
    auditor_script = SCRIPT_LOCATION / "quality" / "auditor.py"
    try:
        proc = subprocess.run(
            [sys.executable, str(auditor_script), str(document_path)],
            capture_output=True, text=True
        )
        print(proc.stdout, end="")
        if proc.stderr:
            print(proc.stderr, end="", file=sys.stderr)
        if proc.returncode != 0:
            return False
    except Exception as exc:
        print(f"Verification exception: {exc}")
        return False

    # Phase 2: OpenXML schema validation
    validator_dll = SCRIPT_LOCATION / "validator" / "DocxChecker.dll"
    if validator_dll.exists():
        try:
            proc = subprocess.run(
                [str(runtime), "--roll-forward", "LatestMajor", str(validator_dll), str(document_path)],
                capture_output=True, text=True
            )
            print(proc.stdout, end="")
            if proc.stderr:
                print(proc.stderr, end="", file=sys.stderr)
            if proc.returncode != 0:
                return False
        except Exception as exc:
            print(f"Schema validation exception: {exc}")
            return False

    return True


def extract_document_metrics(document_path: Path) -> dict:
    """Gather document statistics using pandoc if available."""
    metrics = {"characters": 0, "tokens": 0, "media_count": 0, "has_markup": False, "has_annotations": False}

    if not shutil.which("pandoc"):
        return metrics

    try:
        proc = subprocess.run(
            ["pandoc", str(document_path), "-t", "plain"],
            capture_output=True, text=True, timeout=30
        )
        if proc.returncode == 0:
            content = proc.stdout
            metrics["characters"] = len(content)
            metrics["tokens"] = len(content.split())

        with zipfile.ZipFile(document_path, 'r') as archive:
            entries = archive.namelist()
            metrics["media_count"] = sum(1 for e in entries if e.startswith("word/media/"))
            metrics["has_annotations"] = "word/comments.xml" in entries

            if "word/document.xml" in entries:
                doc_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
                metrics["has_markup"] = "<w:ins" in doc_xml or "<w:del" in doc_xml
    except Exception:
        pass

    return metrics


# Command Handlers
# ============================================================================

def action_doctor():
    """Environment diagnostics and setup (combines status + setup)."""
    print("=== Environment Diagnostics ===")
    print()

    print("Paths:")
    print(f"  Skill root:    {SCRIPT_LOCATION}")
    print(f"  Project home:  {resolve_project_home()}")
    print(f"  Workspace:     {resolve_staging_area()}")
    print(f"  Output dir:    {resolve_artifact_dir()}")
    print()

    print("Runtime:")
    status, binary, ver = assess_runtime_health()
    status_map = {"ready": "+", "outdated": "!", "corrupted": "-", "absent": "o"}
    print(f"  {status_map[status]} dotnet {ver or status}")
    print(f"  + python {platform.python_version()}")

    deps = audit_python_dependencies()
    for name, (state, ver) in deps.items():
        icon = "+" if state == "available" else ("o" if state == "optional" else "-")
        ver_str = f" {ver}" if ver else ""
        suffix = " (required)" if state == "missing" else (" (optional)" if state == "optional" else "")
        print(f"  {icon} {name}{ver_str}{suffix}")
    print()

    # Auto-setup if needed
    needs_setup = status != "ready" or deps.get("lxml", ("missing", None))[0] == "missing"

    if needs_setup:
        print("=== Provisioning Dependencies ===")
        runtime = guarantee_dotnet()
        ver_proc = subprocess.run([str(runtime), '--version'], capture_output=True, text=True)
        print(f"  + dotnet {ver_proc.stdout.strip()}")

        if not guarantee_lxml():
            sys.exit(1)

        print()
        print("=== Preparing Workspace ===")
        prepare_workspace(runtime)
        print(f"  + {resolve_staging_area()}")
    else:
        staging = resolve_staging_area()
        if not staging.exists():
            print("=== Preparing Workspace ===")
            prepare_workspace(binary)
            print(f"  + {staging}")
        else:
            print("Workspace:")
            print(f"  + {staging}")
            if (staging / "Launcher.cs").exists():
                print("    Launcher.cs present")

    print()
    print("Ready!")
    print(f"  Edit:   {resolve_staging_area() / 'Launcher.cs'}")
    print(f"  Render: python {Path(__file__).name} render")
    print(f"  Output: {resolve_artifact_dir()}/")


def action_render(target_name: Optional[str] = None):
    """Build and validate document."""
    runtime = guarantee_dotnet()
    if not guarantee_lxml():
        sys.exit(1)
    prepare_workspace(runtime)

    staging = resolve_staging_area()
    output_dir = resolve_artifact_dir()

    if target_name:
        target = Path(target_name)
        if not target.is_absolute():
            target = output_dir / target_name
    else:
        target = output_dir / "document.docx"

    target.parent.mkdir(parents=True, exist_ok=True)

    print(">> Compiling...")
    proj_file = staging / "DocumentFoundry.csproj"
    proc = subprocess.run(
        [str(runtime), "build", str(proj_file), "--verbosity", "quiet"],
        capture_output=True, text=True, cwd=str(staging)
    )

    if proc.returncode != 0:
        print("!! Compilation failed")
        print()
        diagnostics = CompilerDiagnostics()
        full_output = proc.stdout + proc.stderr
        for line in full_output.split("\n"):
            if "error CS" in line:
                print(f"  {line}")
                suggestions = diagnostics.analyze(line)
                for suggestion in suggestions:
                    print(f"    > Hint: {suggestion.message}")
        sys.exit(1)
    print("  + Compiled")

    print(">> Generating...")
    proc = subprocess.run(
        [str(runtime), "run", "--no-build", "--", str(target)],
        capture_output=True, text=True, cwd=str(staging)
    )

    if proc.returncode != 0:
        print("!! Generation failed")
        if proc.stdout:
            print(proc.stdout)
        if proc.stderr:
            print(proc.stderr, file=sys.stderr)
        sys.exit(1)

    if not target.exists():
        print(f"!! Output missing: {target}")
        print("  Verify Launcher.cs output path")
        sys.exit(1)
    print("  + Generated")

    print(">> Verifying...")
    if not execute_verification(target, runtime):
        print()
        print("!! VERIFICATION FAILED - Document saved but may be invalid")
        print("-" * 58)
        print(f"Document: {target}")
        print("The file may not render correctly in Word/WPS.")
        print()
        print("Potential causes:")
        print("  * Editing existing document: source may be non-conformant")
        print("  * Creating new document: review error messages above")
        print("-" * 58)
        sys.exit(1)

    metrics = extract_document_metrics(target)
    if metrics["characters"] > 0:
        media_note = "" if metrics["media_count"] > 0 else " - verify AddInlineImage() calls"
        print(f"  >> {metrics['characters']} chars, {metrics['tokens']} words, {metrics['media_count']} images{media_note}")
        print(f"  >> Structural check passed. Content review: pandoc \"{target}\" -t plain")
        if metrics["has_markup"] or metrics["has_annotations"]:
            print("     Track changes detected - use --track-changes=all for review")

    print()
    print(f"+ Complete: {target}")


def action_audit(document_path: str):
    """Validate existing document."""
    runtime = guarantee_dotnet()
    if not guarantee_lxml():
        sys.exit(1)

    path = Path(document_path)
    if not path.exists():
        print(f"- Not found: {path}")
        sys.exit(1)

    print(f">> Auditing: {path}")
    if execute_verification(path, runtime):
        print("+ Valid")
    else:
        sys.exit(1)


def action_preview(document_path: str):
    """Quick content preview using pandoc."""
    path = Path(document_path)
    if not path.exists():
        print(f"- Not found: {path}")
        sys.exit(1)

    if not shutil.which("pandoc"):
        print("- pandoc not installed")
        print("  Install: brew install pandoc (macOS) or apt install pandoc (Linux)")
        sys.exit(1)

    print(f">> Preview: {path}")
    print("-" * 60)
    proc = subprocess.run(
        ["pandoc", str(path), "-t", "plain"],
        capture_output=True, text=True, timeout=30
    )
    if proc.returncode == 0:
        print(proc.stdout)
    else:
        print(f"- Preview failed: {proc.stderr}")
        sys.exit(1)


def show_usage():
    """Display command reference."""
    staging = resolve_staging_area()
    output = resolve_artifact_dir()

    usage = f"""
Usage: python docx_engine.py <command> [options]

IMPORTANT: Run from the user's working directory, not the skill directory.
  .docx_workspace/ and output/ are created under cwd.

Commands:
  doctor          Environment diagnostics and auto-setup
  render [name]   Build, execute, validate (default: output/document.docx)
  audit FILE      Validate existing document
  preview FILE    Quick content preview (requires pandoc)

Paths:
  Skill:     {SCRIPT_LOCATION}
  Workspace: {staging}  (edit Launcher.cs here)
  Output:    {output}  (final deliverables)

Creation Workflow:
  1. python docx_engine.py doctor
  2. Edit {staging / 'Launcher.cs'}
  3. python docx_engine.py render report.docx

Modification Workflow:
  1. Analyze uploaded .docx structure
  2. Edit {staging / 'Launcher.cs'}
  3. python docx_engine.py render modified.docx
"""
    print(usage.strip())


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help", "help"):
        show_usage()
        sys.exit(0)

    command = sys.argv[1]

    if command == "doctor":
        action_doctor()
    elif command == "render":
        target = sys.argv[2] if len(sys.argv) > 2 else None
        action_render(target)
    elif command == "audit":
        if len(sys.argv) < 3:
            print("Usage: python docx_engine.py audit <document.docx>")
            sys.exit(1)
        action_audit(sys.argv[2])
    elif command == "preview":
        if len(sys.argv) < 3:
            print("Usage: python docx_engine.py preview <document.docx>")
            sys.exit(1)
        action_preview(sys.argv[2])
    else:
        print(f"Unknown command: {command}")
        print("Run 'python docx_engine.py help' for reference")
        sys.exit(1)


if __name__ == "__main__":
    main()
