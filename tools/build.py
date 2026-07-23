#!/usr/bin/env python3
"""
Liest TSV-Dateien aus data/ und bettet sie als JSON in hallenbelegung.gs ein.

Die Arrays ersetzen den Bereich zwischen den Markern:
    // --- SEED DATA BEGIN ---
    // --- SEED DATA END ---

Usage: python3 tools/build.py
"""

import csv
import json
import os
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
GS_FILE = ROOT / "hallenbelegung.gs"

BEGIN_MARKER = "// --- SEED DATA BEGIN ---"
END_MARKER = "// --- SEED DATA END ---"


def read_tsv(path: Path) -> list[list[str]]:
    """Liest TSV-Datei und gibt Liste von Zeilen (ohne Header) zurück."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f, delimiter="\t")
        header = next(reader, None)
        if header is None:
            return rows
        for row in reader:
            if any(cell.strip() for cell in row):
                rows.append([cell.strip() for cell in row])
    return rows


def format_json_array(name: str, data: list[list[str]]) -> str:
    """Erzeugt: var NAME = JSON.parse('...');"""
    json_str = json.dumps(data, ensure_ascii=False)
    return f"var {name} = JSON.parse('{json_str}');"


def main():
    setup = read_tsv(DATA_DIR / "setup_teams.tsv")
    eingabe = read_tsv(DATA_DIR / "eingabe.tsv")
    sperrungen = read_tsv(DATA_DIR / "sperrungen.tsv")

    seed_lines = [
        format_json_array("SEED_SETUP", setup),
        format_json_array("SEED_EINGABE", eingabe),
        format_json_array("SEED_SPERRUNGEN", sperrungen),
    ]

    content = GS_FILE.read_text(encoding="utf-8")

    if BEGIN_MARKER not in content or END_MARKER not in content:
        print(f"Error: Marker {BEGIN_MARKER} / {END_MARKER} not found in {GS_FILE}")
        return

    new_block = "\n".join(seed_lines)
    pattern = re.compile(
        re.escape(BEGIN_MARKER) + r"\n.*?\n" + re.escape(END_MARKER),
        re.DOTALL,
    )
    replacement = BEGIN_MARKER + "\n" + new_block + "\n" + END_MARKER
    new_content = pattern.sub(replacement, content)

    GS_FILE.write_text(new_content, encoding="utf-8")
    print(f"Wrote {GS_FILE}")
    print(f"  SEED_SETUP: {len(setup)} rows")
    print(f"  SEED_EINGABE: {len(eingabe)} rows")
    print(f"  SEED_SPERRUNGEN: {len(sperrungen)} rows")


if __name__ == "__main__":
    main()
