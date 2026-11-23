#!/usr/bin/env python3
"""
Extract variables (name + type) from C source files.
Outputs a table with 4 columns:
  | Variable Name | Type | Kind | File:Line |
"""

import re
import sys
import argparse
from pathlib import Path
from dataclasses import dataclass
from typing import List


@dataclass
class Cvar:
    name: str
    type: str
    kind: str      # "local", "global", "static", "param", "typedef"
    location: str  # "file.c:42"


# ---- Regexes (deliberately conservative; not a full C parser) -------------
# Matches declarations like:
#   int x;   unsigned long *p;   const char *s = "x";
#   struct Foo foo;   int arr[10];   static int counter;
DECL_RE = re.compile(
    r"""
    ^\s*
    (?P<storage>(?:static|extern|register|auto|volatile|const)\s+)*
    (?P<type>
        (?:unsigned\s+|signed\s+)?
        (?:int|char|short|long|float|double|void|_Bool|size_t|
           u?int(?:8|16|32|64)_t|
           [A-Za-z_]\w*(?:\s*\*)?)      # user-defined types / pointers
        (?:\s+[A-Za-z_]\w*)*             # multi-word types e.g. unsigned long int
    )
    \s+
    (?P<name>[A-Za-z_]\w*)
    \s*
    (?P<suffix>\[[^\]]*\])?              # optional array dims
    \s*(?:=[^;]*)?;                      # optional initializer
    """,
    re.VERBOSE,
)

FUNC_RE = re.compile(
    r"""
    (?P<ret>[A-Za-z_][\w\s\*]*?)         # return type
    \s+
    (?P<name>[A-Za-z_]\w*)               # function name
    \s*\((?P<params>[^)]*)\)\s*(?:;|\{)
    """,
    re.VERBOSE,
)

TYPEDEF_RE = re.compile(
    r"typedef\s+(?P<type>[\w\s\*]+?)\s+(?P<name>[A-Za-z_]\w*)\s*;"
)


def strip_comments_and_strings(src: str) -> str:
    """Remove // comments, /* */ comments, and string/char literals."""
    src = re.sub(r"/\*.*?\*/", "", src, flags=re.DOTALL)
    src = re.sub(r"//[^\n]*", "", src)
    src = re.sub(r'"(?:\\.|[^"\\])*"', '""', src)
    src = re.sub(r"'(?:\\.|[^'\\])*'", "''", src)
    return src


def split_statements(src: str):
    """Yield (line_number, statement) for top-level ; separated chunks."""
    for match in re.finditer(r"[^;{}]+;", src):
        line = src.count("\n", 0, match.start()) + 1
        yield line, match.group(0)


def extract_from_file(path: Path) -> List[Cvar]:
    raw = path.read_text(encoding="utf-8", errors="replace")
    src = strip_comments_and_strings(raw)
    found: List[Cvar] = []

    # Skip function bodies when hunting declarations? Simpler: process every
    # statement. Parameters are captured separately from function headers.
    for line, stmt in split_statements(src):
        s = stmt.strip()
        if not s or s.startswith("#"):
            continue

        # --- typedefs ---
        m = TYPEDEF_RE.search(s)
        if m:
            found.append(Cvar(m.group("name"), m.group("type").strip(),
                              "typedef", f"{path.name}:{line}"))
            continue

        # --- function declarations / definitions -> extract parameters ---
        fm = FUNC_RE.fullmatch(s + ("{" if s.endswith(")") else ""))
        if fm:
            params = fm.group("params").strip()
            if params and params != "void":
                for p in params.split(","):
                    p = p.strip().lstrip("*")
                    pm = re.match(r"([\w\s\*]+?)\s+([A-Za-z_]\w*)$", p)
                    if pm:
                        found.append(Cvar(pm.group(2).strip(),
                                          pm.group(1).strip(),
                                          "param",
                                          f"{path.name}:{line}"))
            continue

        # --- ordinary variable declarations ---
        dm = DECL_RE.match(s)
        if dm:
            storage = (dm.group("storage") or "").strip()
            kind = "static" if "static" in storage else "local"
            ctype = (dm.group("type") or "").strip()
            if dm.group("suffix"):
                ctype += " " + dm.group("suffix").strip()
            found.append(Cvar(dm.group("name"), ctype, kind,
                              f"{path.name}:{line}"))
    return found


def to_markdown(vars: List[Cvar], extra_header="Notes") -> str:
    lines = ["| Variable | Type | Kind | Location | " + extra_header + " |",
             "|---|---|---|---|---|"]
    for v in vars:
        lines.append(f"| `{v.name}` | `{v.type}` | {v.kind} | {v.location} |  |")
    return "\n".join(lines)


def to_docx(vars: List[Cvar], out: Path, extra_header="Notes"):
    """Generate a .docx directly (pip install python-docx)."""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    doc.add_heading("C Variable Inventory", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(["Variable", "Type", "Kind", "Location", extra_header]):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for v in vars:
        row = table.add_row().cells
        row[0].text = v.name
        row[1].text = v.type
        row[2].text = v.kind
        row[3].text = v.location
        row[4].text = ""

    for p in doc.paragraphs:
        for r in p.runs:
            r.font.size = Pt(10)
    doc.save(out)


def main():
    ap = argparse.ArgumentParser(description="Extract C variables into a table.")
    ap.add_argument("files", nargs="+", help="C source/header files")
    ap.add_argument("--format", choices=["markdown", "csv", "docx"],
                    default="markdown")
    ap.add_argument("--out", help="Output file (for csv/docx)")
    args = ap.parse_args()

    all_vars: List[Cvar] = []
    for f in args.files:
        all_vars.extend(extract_from_file(Path(f)))

    if args.format == "markdown":
        print(to_markdown(all_vars))
    elif args.format == "csv":
        import csv
        out = args.out or "variables.csv"
        with open(out, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["Variable", "Type", "Kind", "Location", "Notes"])
            for v in all_vars:
                w.writerow([v.name, v.type, v.kind, v.location, ""])
        print(f"Wrote {out} ({len(all_vars)} variables)")
    else:
        out = Path(args.out or "variables.docx")
        to_docx(all_vars, out)
        print(f"Wrote {out} ({len(all_vars)} variables)")


if __name__ == "__main__":
    main()