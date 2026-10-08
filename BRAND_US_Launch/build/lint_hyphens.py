#!/usr/bin/env python3
"""Hyphen and dash lint for every text bearing output.
Usage: python3 lint_hyphens.py <file or dir> [...]   (exit 1 if any violation)
Rule: no hyphens, en dashes or em dashes in prose. Official form and registration numbers keep their format.
"""
import re, sys, subprocess, zipfile, io, json
from pathlib import Path

ALLOW = [
    r"SS-4", r"W-8BEN(-E)?", r"W-9", r"599-64-00496", r"C-TPAT", r"ISO-\d+", r"UL-?\d*", r"IEC-?\d*",
    r"\b\d{3}-\d{2}-\d{5}\b",          # Korean business registration numbers
    r"\b\d{2}-\d{7}\b",                # EIN format 12-3456789
    r"\b[A-Z]\d-[A-Z0-9]+\b",          # odd official codes
    r"Form\s+\d+-[A-Z0-9]+",           # Form 1120-F style
    r"\b1120-F\b", r"\b1040-NR\b", r"\b8804-C\b", r"\b1042-S\b",
    r"https?://\S+",                   # URLs
    r"[A-Za-z0-9_./-]+\.(?:py|js|json|md|mmd|xlsx|pptx|docx|pdf|png|jpg|csv|sh|ttf|otf)\b",  # file names
    r"--?[a-z][a-z-]*",                # CLI flags in code blocks
    r"-->|---|-\.->|==>|-\.-|<--|\|--", # mermaid arrows
    r"\b[A-Z]{2,}-\d+\b",              # ticket style codes like HTS-1
    r"\bCL3-[A-Z]*\b", r"\bRG-?\d+\b",
    r"\b8544\.\d{2}\.\d{2}\b",
    r"\bNoto Sans KR\b",
    r"(?<=\d)-(?=\d{4}\b)",            # date ranges inside official citations only
]
ALLOW_RE = re.compile("|".join(ALLOW))
DASH_RE = re.compile(r"[-–—−]")

def text_of(p: Path) -> str:
    s = p.suffix.lower()
    if s in {".md", ".txt", ".mmd", ".csv", ".json", ".html"}:
        return p.read_text(errors="ignore")
    if s == ".docx":
        return subprocess.run(["pandoc", "-t", "plain", str(p)], capture_output=True, text=True).stdout
    if s == ".pptx":
        return subprocess.run(["markitdown", str(p)], capture_output=True, text=True).stdout
    if s == ".xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(p, data_only=True, read_only=True)
        out = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str): out.append(v)
        return "\n".join(out)
    if s == ".pdf":
        return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout
    return ""

def violations(text: str):
    out = []
    for i, line in enumerate(text.splitlines(), 1):
        if "```" in line: continue
        if re.match(r"^\s*[-*+]\s", line): line = re.sub(r"^\s*[-*+]\s", "", line, count=1)   # markdown list marker
        if re.match(r"^\s*\|?\s*:?-{2,}", line): continue                                    # markdown table rule
        cleaned = ALLOW_RE.sub("", line)
        if DASH_RE.search(cleaned):
            out.append((i, line.strip()[:160]))
    return out

def main(paths):
    bad = 0
    files = []
    for a in paths:
        p = Path(a)
        if p.is_dir(): files += [f for f in p.rglob("*") if f.is_file() and f.suffix.lower() in {".md",".txt",".mmd",".csv",".docx",".pptx",".xlsx",".pdf"}]
        else: files.append(p)
    for f in sorted(set(files)):
        if "/build/" in str(f) and f.suffix in {".py", ".js"}: continue
        v = violations(text_of(f))
        if v:
            bad += len(v)
            print(f"== {f} ({len(v)} lines)")
            for i, l in v[:40]: print(f"  L{i}: {l}")
            if len(v) > 40: print(f"  ... {len(v)-40} more")
    print(json.dumps({"files_checked": len(files), "violating_lines": bad}))
    sys.exit(1 if bad else 0)

if __name__ == "__main__":
    main(sys.argv[1:] or ["."])
