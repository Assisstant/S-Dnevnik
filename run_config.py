#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_config.py - Reads gen_config.json and runs the document generator.

WORKFLOW:
  1. S-Dnevnik_v7.html Generator tab -> "Save Config" -> gen_config.json
  2. Put gen_config.json in this folder (same as LAUNCHER.bat)
  3. Double-click LAUNCHER.bat
  4. This script:
     a) Auto-archives JSON data to ARCHIVE/ with timestamp
     b) Runs document_factory_v7.py or compact_factory.py
     c) Renames gen_config.json -> gen_config_done.json (won't re-run)
  5. DOCX files appear in OUT/

YOUR FOLDER after a few runs:
  SMUROP_Package/
    e_dnevnik_unified_state_v7.json   <- always the LATEST
    gen_config_done.json              <- last used config
    ARCHIVE/
      e_dnevnik_unified_state_v7_2026-02-17_1430.json
      e_dnevnik_unified_state_v7_2026-02-18_0900.json
      e_dnevnik_unified_state_v7_2026-02-20_1100.json
    OUT/
      DOSIE_Student1.docx
      DOSIE_Student2.docx
"""

import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def archive_json(base, json_filename):
    """Copy JSON to ARCHIVE/ with timestamp before generating."""
    json_path = base / json_filename
    if not json_path.exists():
        return None

    archive_dir = base / "ARCHIVE"
    archive_dir.mkdir(exist_ok=True)

    stem = json_path.stem
    ext = json_path.suffix
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    archive_name = f"{stem}_{timestamp}{ext}"
    archive_path = archive_dir / archive_name

    # Skip if already archived this minute (re-run same session)
    if archive_path.exists():
        if archive_path.stat().st_size == json_path.stat().st_size:
            return archive_path

    shutil.copy2(json_path, archive_path)
    return archive_path


def main():
    base = Path(__file__).resolve().parent

    config_path = base / "gen_config.json"
    if "--config" in sys.argv:
        idx = sys.argv.index("--config")
        if idx + 1 < len(sys.argv):
            config_path = Path(sys.argv[idx + 1])

    if not config_path.exists():
        print("[ERROR] Config file not found:", config_path)
        print("        Export it from the Generator tab in S-Dnevnik_v7.html")
        return 1

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    json_filename = cfg.get("jsonPath", "e_dnevnik_unified_state_v7.json")

    print("=" * 60)
    print("  SMUROP Document Generator")
    print("=" * 60)
    print(f"  Config:    {config_path.name}")
    print(f"  Generator: {cfg.get('generator', 'document_factory_v7')}")
    print(f"  Data:      {json_filename}")
    print(f"  Doc types: {', '.join(cfg.get('docTypes', ['dosie']))}")
    print(f"  Output:    {cfg.get('outDir', 'OUT')}/")
    print("=" * 60)

    # Check JSON data file exists
    json_path = base / json_filename
    if not json_path.exists():
        print(f"\n[ERROR] Data file not found: {json_filename}")
        print(f"        Expected at: {json_path}")
        print()
        print("  To fix this:")
        print("  1. Open S-Dnevnik_v7.html in Chrome")
        print("  2. Go to Generator tab")
        print(f"  3. Click 'Export JSON' and save as: {json_filename}")
        print(f"  4. Put it in: {base}")
        return 1

    # Auto-archive JSON before generation
    archive_path = archive_json(base, json_filename)
    if archive_path:
        print(f"\n  [ARCHIVE] {json_filename}")
        print(f"         -> ARCHIVE/{archive_path.name}")

    # Find Python executable
    venv_py = base / ".venv" / "Scripts" / "python.exe"
    if not venv_py.exists():
        venv_py = base / ".venv" / "bin" / "python"
    if not venv_py.exists():
        venv_py = Path(sys.executable)
    py = str(venv_py)

    generator = cfg.get("generator", "document_factory_v7")

    if generator == "compact_factory":
        script = base / "compact_factory.py"
        if not script.exists():
            print(f"\n[ERROR] compact_factory.py not found in {base}")
            return 1

        cmd = [
            py, str(script), "gen-template",
            "--unified", json_filename,
            "--template", cfg.get("templatePath", "template.docx"),
            "--out", cfg.get("outDir", "OUT"),
        ]
        if cfg.get("selectedStudent"):
            cmd += ["--student", cfg["selectedStudent"]]
    else:
        script = base / "document_factory_v7.py"
        if not script.exists():
            print(f"\n[ERROR] document_factory_v7.py not found in {base}")
            return 1

        cmd = [
            py, str(script),
            "--unified", json_filename,
            "--out", cfg.get("outDir", "OUT"),
            "--doc-types", ",".join(cfg.get("docTypes", ["dosie"])),
            "--school-year", cfg.get("schoolYear", "2025-2026"),
            "--output-mode", cfg.get("outputMode", "separate"),
            "--font-name", cfg.get("fontName", "Times New Roman"),
            "--font-size", str(cfg.get("fontSize", 12)),
        ]
        if cfg.get("therapistName"):
            cmd += ["--therapist-name", cfg["therapistName"]]
        if cfg.get("therapistTitle"):
            cmd += ["--therapist-title", cfg["therapistTitle"]]
        if cfg.get("institution"):
            cmd += ["--institution", cfg["institution"]]
        if cfg.get("periodLabel"):
            cmd += ["--period-label", cfg["periodLabel"]]
        if cfg.get("selectedStudent"):
            cmd += ["--student", cfg["selectedStudent"]]
        if cfg.get("downloadImages"):
            cmd += ["--download-images"]

    print(f"\n  [RUN] {Path(cmd[1]).name}")
    for i in range(2, len(cmd), 2):
        flag = cmd[i] if i < len(cmd) else ""
        val = cmd[i + 1] if i + 1 < len(cmd) else ""
        print(f"        {flag} {val}")
    print()

    result = subprocess.run(cmd, cwd=str(base))

    # Rename config so it doesn't re-run next time
    done_path = config_path.with_name("gen_config_done.json")
    try:
        if done_path.exists():
            done_path.unlink()
        config_path.rename(done_path)
    except Exception:
        pass

    out_dir = base / cfg.get("outDir", "OUT")
    if result.returncode == 0:
        count = len(list(out_dir.glob("*.docx"))) if out_dir.exists() else 0
        print(f"\n  [OK] Done! {count} DOCX file(s) in: {out_dir.name}/")
        print(f"  [OK] Config saved as: gen_config_done.json")
        print(f"  [OK] Data archived in: ARCHIVE/")
    else:
        print(f"\n  [FAIL] Generator exited with code {result.returncode}")

    return result.returncode


if __name__ == "__main__":
    code = main()
    print()
    input("Press Enter to close...")
    sys.exit(code)
