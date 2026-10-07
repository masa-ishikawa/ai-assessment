"""Portable entry point for the AI assessment Python tools.

Creates a repository-local virtual environment on first use, installs the
declared dependencies, then forwards arguments to the selected tool.
"""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import venv
from pathlib import Path


PROJECT = Path(__file__).resolve().parent.parent
VENV = PROJECT / ".venv"
REQUIREMENTS = PROJECT / "requirements.txt"
MARKER = VENV / ".ai_assess_requirements.sha256"
SKILL = PROJECT / ".agents" / "skills" / "ai-assessment"
TOOLS = {
    "generate": PROJECT / "generate_assessment.py",
    "fill": PROJECT / "fill_assessment_input.py",
    "restyle": SKILL / "scripts" / "restyle_pptx.py",
    "revise": SKILL / "scripts" / "revise_pptx.py",
}


def venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def dependency_fingerprint() -> str:
    payload = REQUIREMENTS.read_bytes() + f"\n{sys.version_info.major}.{sys.version_info.minor}".encode()
    return hashlib.sha256(payload).hexdigest()


def setup() -> Path:
    if sys.version_info < (3, 12):
        raise RuntimeError("Python 3.12 or newer is required")
    python = venv_python()
    if not python.is_file():
        print(f"Creating virtual environment: {VENV}", flush=True)
        venv.EnvBuilder(with_pip=True).create(VENV)
        MARKER.unlink(missing_ok=True)
    fingerprint = dependency_fingerprint()
    if not MARKER.is_file() or MARKER.read_text(encoding="utf-8").strip() != fingerprint:
        print("Installing Python dependencies from requirements.txt", flush=True)
        subprocess.run(
            [str(python), "-m", "pip", "install", "-r", str(REQUIREMENTS)],
            cwd=PROJECT,
            check=True,
        )
        MARKER.write_text(fingerprint + "\n", encoding="utf-8")
    return python


def main() -> int:
    usage = "Usage: python scripts/run_assessment.py {setup|doctor|generate|fill|restyle|revise} [arguments...]"
    if len(sys.argv) < 2 or sys.argv[1] in {"-h", "--help"}:
        print(usage)
        return 0 if len(sys.argv) > 1 else 2
    command, arguments = sys.argv[1], sys.argv[2:]
    if command == "doctor":
        print(f"Project: {PROJECT}")
        print(f"Bootstrap Python: {sys.executable}")
        ready = venv_python().is_file() and MARKER.is_file() and MARKER.read_text(encoding="utf-8").strip() == dependency_fingerprint()
        print(f"Virtual environment: {venv_python()} ({'ready' if ready else 'setup needed'})")
        print(f"Repository skill: {SKILL / 'SKILL.md'} ({'found' if (SKILL / 'SKILL.md').is_file() else 'missing'})")
        return 0
    if command not in TOOLS and command != "setup":
        print(usage, file=sys.stderr)
        return 2
    try:
        python = setup()
        if command == "setup":
            print(f"Ready: {python}")
            return 0
        tool = TOOLS[command]
        if not tool.is_file():
            raise FileNotFoundError(f"Missing tool: {tool}")
        if command == "restyle":
            if len(arguments) < 2:
                raise ValueError("restyle requires SOURCE_PPTX OUTPUT_PPTX --plan PLAN_JSON")
            arguments = [arguments[0], str(SKILL / "assets" / "base.pptx"), *arguments[1:]]
        return subprocess.run([str(python), str(tool), *arguments], cwd=PROJECT).returncode
    except (OSError, subprocess.CalledProcessError, RuntimeError, ValueError) as error:
        print(f"AI assessment setup failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
