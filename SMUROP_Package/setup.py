#!/usr/bin/env python3
"""
AUTOMATED SETUP SCRIPT
═════════════════════════════════════════════════════════════════════════════════

This script sets up the complete environment for the JSON-to-DOCX generation workflow:
  1. Checks Python version (requires Python 3.8+, recommends 3.10+)
  2. Creates a virtual environment (.venv) if not present
  3. Upgrades pip, setuptools, and wheel
  4. Installs all required packages:
     - python-docx (for DOCX file manipulation)
     - requests (for optional image downloading)
     - pywin32 (for Windows COM integration, optional)

Run this script once on a new computer, then use the workflow normally.

Usage:
  python setup.py              (interactive)
  python setup.py --no-prompt  (non-interactive, defaults to 'yes')
  python setup.py --uninstall  (remove virtual environment)
"""

import sys
import os
import subprocess
import shutil
from pathlib import Path
import io

# Ensure stdout/stderr use UTF-8 for Unicode characters
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
else:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass


REQUIRED_PACKAGES = [
    "python-docx>=0.8.11",
    "requests>=2.28.0",
]

OPTIONAL_PACKAGES = [
    "pywin32>=305",  # Optional but useful on Windows
]

MIN_PYTHON_VERSION = (3, 8)
RECOMMENDED_PYTHON_VERSION = (3, 10)


def print_banner(title):
    """Print a formatted banner."""
    width = 80
    print("\n" + "=" * width)
    print(f"  {title}")
    print("=" * width + "\n")


def print_step(step_num, description):
    """Print a step indicator."""
    print(f"\n[{step_num}] {description}")
    print("-" * 60)


def check_python_version():
    """Check Python version and warn if below recommended."""
    print_step(1, "Checking Python version")
    
    version = sys.version_info[:2]
    py_str = f"{version[0]}.{version[1]}"
    
    print(f"Current Python: {py_str} ({sys.executable})")
    
    if version < MIN_PYTHON_VERSION:
        print(f"❌ ERROR: Python {MIN_PYTHON_VERSION[0]}.{MIN_PYTHON_VERSION[1]}+ required!")
        print(f"   Current version is too old. Please install Python 3.8 or newer.")
        return False
    
    if version < RECOMMENDED_PYTHON_VERSION:
        rec = RECOMMENDED_PYTHON_VERSION
        print(f"⚠️  WARNING: Python {rec[0]}.{rec[1]}+ recommended (you have {py_str})")
        print(f"   The workflow may have issues with very old Python versions.")
        return True
    
    print(f"✓ Python {py_str} OK (recommended: 3.10+)")
    return True


def get_project_root():
    """Get the project root directory (where setup.py is located)."""
    return Path(__file__).resolve().parent


def check_and_create_venv():
    """Check if .venv exists, create if needed."""
    print_step(2, "Setting up virtual environment")
    
    root = get_project_root()
    venv_path = root / ".venv"
    
    if venv_path.exists():
        print(f"✓ Virtual environment already exists: {venv_path}")
        return venv_path
    
    print(f"Creating virtual environment at: {venv_path}")
    try:
        subprocess.run([sys.executable, "-m", "venv", str(venv_path)], 
                       check=True, capture_output=True)
        print(f"✓ Virtual environment created successfully")
        return venv_path
    except Exception as e:
        print(f"❌ ERROR: Failed to create virtual environment: {e}")
        return None


def get_venv_python(venv_path):
    """Get the Python executable path in the virtual environment."""
    if sys.platform == "win32":
        return venv_path / "Scripts" / "python.exe"
    else:
        return venv_path / "bin" / "python"


def get_venv_pip(venv_path):
    """Get the pip executable path in the virtual environment."""
    if sys.platform == "win32":
        return venv_path / "Scripts" / "pip.exe"
    else:
        return venv_path / "bin" / "pip"


def upgrade_pip(venv_path):
    """Upgrade pip, setuptools, and wheel."""
    print_step(3, "Upgrading pip, setuptools, and wheel")
    
    pip_exe = get_venv_pip(venv_path)
    
    packages_to_upgrade = ["pip", "setuptools", "wheel"]
    
    for pkg in packages_to_upgrade:
        print(f"  Upgrading {pkg}...", end=" ", flush=True)
        try:
            subprocess.run([str(pip_exe), "install", "--upgrade", pkg],
                          check=True, capture_output=True, text=True)
            print("✓")
        except subprocess.CalledProcessError as e:
            print(f"⚠️  (warning: {e})")
    
    print("✓ pip and dependencies upgraded")


def install_packages(venv_path, packages, section_name="Required"):
    """Install a list of packages."""
    print_step(4, f"Installing {section_name} packages")
    
    pip_exe = get_venv_pip(venv_path)
    
    for pkg in packages:
        print(f"  Installing {pkg}...", end=" ", flush=True)
        try:
            subprocess.run([str(pip_exe), "install", pkg],
                          check=True, capture_output=True, text=True)
            print("✓")
        except subprocess.CalledProcessError as e:
            print(f"❌ FAILED")
            print(f"     Error: {e.stderr}")
            return False
    
    print(f"✓ All {section_name.lower()} packages installed")
    return True


def verify_imports(venv_path):
    """Verify that all required packages can be imported."""
    print_step(5, "Verifying package imports")
    
    py_exe = get_venv_python(venv_path)
    
    # Test imports
    test_code = """
import sys
try:
    import docx
    print("✓ python-docx imported successfully")
except ImportError as e:
    print(f"❌ Failed to import python-docx: {e}")
    sys.exit(1)

try:
    import requests
    print("✓ requests imported successfully")
except ImportError as e:
    print(f"❌ Failed to import requests: {e}")
    sys.exit(1)

try:
    import pywin32
    print("✓ pywin32 imported successfully")
except ImportError:
    print("⚠️  pywin32 not available (optional, skipping)")

print("All required packages verified!")
"""
    
    try:
        result = subprocess.run([str(py_exe), "-c", test_code],
                              capture_output=True, text=True, check=False)
        print(result.stdout)
        if result.returncode != 0 and "python-docx" in result.stderr:
            print("❌ Import verification failed")
            return False
    except Exception as e:
        print(f"⚠️  Could not verify imports: {e}")
    
    return True


def print_next_steps(venv_path):
    """Print instructions for next steps."""
    print_banner("SETUP COMPLETE ✓")
    
    root = get_project_root()
    py_exe = get_venv_python(venv_path)
    rel_path = py_exe.relative_to(root) if root in py_exe.parents else py_exe
    
    print("You can now use the workflow:")
    print()
    print("Option 1 - Using the virtual environment:")
    print(f"  {rel_path} compact_factory.py --help")
    print(f"  {rel_path} compact_factory.py list-students --unified e_dnevnik_unified_state_v7.json")
    print(f"  {rel_path} compact_factory.py gen-template --unified e_dnevnik_unified_state_v7.json --template template.docx --out OUT")
    print()
    print("Option 2 - On Windows, double-click RUN_SMUROP_v7.bat to start the workflow")
    print()
    print("Option 3 - Manually activate venv (cross-platform):")
    if sys.platform == "win32":
        print(f"  .venv\\Scripts\\activate.bat")
    else:
        print(f"  source .venv/bin/activate")
    print(f"  python compact_factory.py --help")
    print()


def cleanup_and_exit(venv_path):
    """Option to remove virtual environment."""
    if input("\nRemove virtual environment? (y/N): ").lower() == "y":
        print(f"Removing {venv_path}...")
        try:
            shutil.rmtree(venv_path)
            print("✓ Virtual environment removed")
        except Exception as e:
            print(f"⚠️  Could not remove: {e}")


def main():
    """Main setup routine."""
    print_banner("JSON-to-DOCX WORKFLOW SETUP")
    print("This script will set up Python environment and install dependencies.")
    print()
    
    # Parse arguments
    no_prompt = "--no-prompt" in sys.argv
    uninstall = "--uninstall" in sys.argv
    
    if uninstall:
        venv_path = get_project_root() / ".venv"
        if venv_path.exists():
            cleanup_and_exit(venv_path)
        else:
            print("Virtual environment not found.")
        return
    
    # Step 1: Check Python version
    if not check_python_version():
        print("\n❌ Setup aborted: Python version too old")
        sys.exit(1)
    
    # Step 2: Create virtual environment
    venv_path = check_and_create_venv()
    if not venv_path:
        print("\n❌ Setup aborted: Could not create virtual environment")
        sys.exit(1)
    
    # Step 3: Upgrade pip
    upgrade_pip(venv_path)
    
    # Step 4: Install required packages
    if not install_packages(venv_path, REQUIRED_PACKAGES, "Required"):
        print("\n❌ Setup aborted: Failed to install required packages")
        sys.exit(1)
    
    # Step 4b: Install optional packages (don't fail if these fail)
    print_step(4, "Installing optional packages")
    install_packages(venv_path, OPTIONAL_PACKAGES, "Optional")
    
    # Step 5: Verify imports
    if not verify_imports(venv_path):
        print("\n⚠️  WARNING: Some packages could not be verified")
    
    # Print next steps
    print_next_steps(venv_path)
    
    print("=" * 80)
    print("  Setup finished! You're ready to use the workflow.")
    print("=" * 80)
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nSetup cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
