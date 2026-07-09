#!/usr/bin/env python3
"""Parse Australian Curriculum v9 F-6 content descriptions from ACARA's
official DOCX downloads into a flat JSON dataset.

Reproducibility script for the `acara-v9-f6-json` dataset. Given the seven
official curriculum-content DOCX files (URLs below), it extracts every
content description with its AC9 code, strand, year level(s), subject and
learning area.

Method
------
Each DOCX lays out content descriptions in tables. A table whose header row
begins "Strand: <name>" starts a new strand; the same header row carries the
year label ("Foundation", "Year 3", "Years 1-2"). Tables headed
"Sub-strand: <name>" continue the current strand and year. Within a content
table, the first cell of each body row holds the description text followed by
its AC9 code on its own line; codes are matched with the regex
\\bAC9[A-Z0-9]+\\b. Year levels are taken from the table's year label - the
document's own placement of each code - not inferred from the code string.

Australian Curriculum v9 (c) ACARA (Australian Curriculum, Assessment and
Reporting Authority), licensed CC BY 4.0. This script is MIT licensed.

Usage
-----
    python parse_acara_docx.py <docx-dir> -o out.json
    python parse_acara_docx.py <docx-dir> --verify reference.json

<docx-dir> must contain the seven files named as in DOCX_SOURCES keys
(e.g. mathematics.docx). Download them from the URLs in DOCX_SOURCES.
Requires: python-docx  (pip install python-docx)
"""

import argparse
import json
import re
import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")

# Official ACARA download URLs, as published on v9.australiancurriculum.edu.au
# (the v9. host 301-redirects to www.; both work). Verified 2026-07-09.
_BASE = ("https://www.australiancurriculum.edu.au/content/dam/en/curriculum/"
         "ac-version-9/downloads")
DOCX_SOURCES = {
    "english": {
        "url": f"{_BASE}/english/english-curriculum-content-f-6-v9.docx",
        "subject": "English",
        "learningArea": "English",
    },
    "mathematics": {
        "url": f"{_BASE}/mathematics/mathematics-curriculum-content-f-6-v9.docx",
        "subject": "Mathematics",
        "learningArea": "Mathematics",
    },
    "science": {
        "url": f"{_BASE}/science/science-curriculum-content-f-6-v9.docx",
        "subject": "Science",
        "learningArea": "Science",
    },
    "health-and-physical-education": {
        "url": (f"{_BASE}/health-and-physical-education/"
                "health-and-physical-education-curriculum-content-f-6-v9.docx"),
        "subject": "Health and Physical Education",
        "learningArea": "Health and Physical Education",
    },
    "hass": {
        "url": (f"{_BASE}/humanities-and-social-sciences/hass-f-6/"
                "humanities-and-social-sciences-hass-f-6-curriculum-content-v9.docx"),
        "subject": "HASS F-6",
        "learningArea": "Humanities and Social Sciences",
    },
    "digital-technologies": {
        "url": (f"{_BASE}/technologies/digital-technologies/"
                "technologies-digital-technologies-curriculum-content-f-6-v9.docx"),
        "subject": "Digital Technologies",
        "learningArea": "Technologies",
    },
    "design-and-technologies": {
        "url": (f"{_BASE}/technologies/design-and-technologies/"
                "technologies-design-and-technologies-curriculum-content-f-6-v9.docx"),
        "subject": "Design and Technologies",
        "learningArea": "Technologies",
    },
}

CODE_RE = re.compile(r"\bAC9[A-Z0-9]+\b")
F6_SET = ["F", "1", "2", "3", "4", "5", "6"]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_SPACE_RE = re.compile("[\u00a0\u2007\u2009\u202f]")


def _walk_text(el, seg):
    """Depth-first text extraction handling OMML fractions and hyphens."""
    if el.tag == _M + "f":  # OMML fraction -> 'numerator/denominator'
        num, den = el.find(_M + "num"), el.find(_M + "den")
        if num is not None and den is not None:
            n = "".join(t.text or "" for t in num.iter(_M + "t"))
            d = "".join(t.text or "" for t in den.iter(_M + "t"))
            seg.append(f"{n}/{d}")
        return  # do not descend: fraction already rendered
    if el.tag in (_W + "t", _M + "t"):
        if el.text:
            seg.append(el.text)
        return
    if el.tag == _W + "noBreakHyphen":
        seg.append("-")
        return
    if el.tag in (_W + "br", _W + "cr"):
        seg.append("\n")
        return
    for child in el:
        _walk_text(child, seg)


def cell_text(cell):
    """Cell text including OMML equation fractions (rendered as 'n/d').

    python-docx's cell.text silently drops math runs, which ACARA uses for
    fractions (e.g. in AC9M3N02). This walks the raw XML instead.
    """
    paragraphs = []
    for p in cell._tc.iter(_W + "p"):
        seg = []
        _walk_text(p, seg)
        paragraphs.append("".join(seg))
    return "\n".join(paragraphs)


def clean_text(text):
    """Normalise unicode spaces to ASCII and collapse runs of spaces."""
    text = _SPACE_RE.sub(" ", text)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def parse_year_label(label):
    """'Foundation' -> ['F']; 'Year 3' -> ['3']; 'Years 1-2' -> ['1','2'].

    Handles hyphen, en dash and em dash; 'Foundation to Year 2' -> F,1,2.
    Returns None if the label is not a year label.
    """
    label = label.strip()
    if not label:
        return None
    if re.fullmatch(r"Foundation( [Yy]ear)?", label):
        return ["F"]
    m = re.fullmatch(r"Foundation to Year (\d)", label)
    if m:
        end = F6_SET.index(m.group(1))
        return F6_SET[: end + 1]
    m = re.fullmatch(r"Year (\d)", label)
    if m:
        return [m.group(1)]
    m = re.fullmatch(r"Years (\d)\s*[-–—]\s*(\d)", label)
    if m:
        i, j = F6_SET.index(m.group(1)), F6_SET.index(m.group(2))
        return F6_SET[i : j + 1]
    return None


def parse_docx(path, subject, learning_area):
    """Extract content-description records from one curriculum DOCX."""
    doc = Document(str(path))
    records = {}
    strand = None
    years = None
    for table in doc.tables:
        header_cells = [c.text.strip() for c in table.rows[0].cells]
        # A 'Strand:' header starts a new strand and carries the year label.
        strand_cells = [c for c in header_cells if c.startswith("Strand:")]
        if strand_cells:
            strand = strand_cells[0].split(":", 1)[1].strip()
            year_labels = [y for y in (parse_year_label(c) for c in header_cells) if y]
            years = year_labels[0] if year_labels else years
        elif any(c.startswith("Sub-strand:") for c in header_cells):
            pass  # continues the current strand and year band
        else:
            continue  # year-level descriptions, achievement standards, etc.
        if strand is None or years is None:
            continue
        for row in table.rows[1:]:
            cell = cell_text(row.cells[0])
            codes = CODE_RE.findall(cell)
            if not codes:
                continue
            assert len(set(codes)) == 1, f"multiple codes in one cell: {codes}"
            code = codes[0]
            description = clean_text(cell.split(code)[0])
            if not description:
                continue
            if code in records:
                # A code repeated across tables keeps its first description
                # and accumulates any additional year levels.
                for y in years:
                    if y not in records[code]["years"]:
                        records[code]["years"].append(y)
                continue
            records[code] = {
                "code": code,
                "learningArea": learning_area,
                "subject": subject,
                "strand": strand,
                "year": None,  # set below
                "description": description,
                "years": list(years),
            }
    for rec in records.values():
        rec["years"].sort(key=F6_SET.index)
        rec["year"] = rec["years"][0] if len(rec["years"]) == 1 else None
    return list(records.values())


def parse_all(docx_dir):
    docx_dir = Path(docx_dir)
    out = []
    for stem, cfg in DOCX_SOURCES.items():
        path = docx_dir / f"{stem}.docx"
        if not path.exists():
            sys.exit(f"missing {path} - download it from {cfg['url']}")
        out.extend(parse_docx(path, cfg["subject"], cfg["learningArea"]))
    return out


def verify(records, reference_path):
    """Assert parsed records exactly reproduce the published dataset."""
    with open(reference_path, encoding="utf-8") as f:
        ref = json.load(f)
    assert len(records) == 614, f"expected 614 records, got {len(records)}"
    assert len(records) == len(ref), f"count mismatch: {len(records)} vs {len(ref)}"
    got_full = {r["code"]: r for r in records}
    assert set(got_full) == {r["code"] for r in ref}, "code sets differ"
    diffs = []
    for r in ref:
        g = got_full[r["code"]]
        for k in ("learningArea", "subject", "strand", "year", "years",
                  "description"):
            if g[k] != r[k]:
                diffs.append((r["code"], k, r[k], g[k]))
    for code, k, want, got in diffs:
        print(f"  field diff: {code}.{k}: reference={want!r} parsed={got!r}")
    assert not diffs, f"{len(diffs)} field difference(s)"
    print("PASS: 614 records parsed from the official DOCX files exactly "
          "reproduce the published dataset (all fields, order-independent).")
    return diffs


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("docx_dir", help="directory containing the seven DOCX files")
    ap.add_argument("-o", "--output", help="write records to this JSON file")
    ap.add_argument("--verify", metavar="REFERENCE_JSON",
                    help="assert output reproduces the published dataset")
    args = ap.parse_args()

    recs = parse_all(args.docx_dir)
    order = {cfg["subject"]: i for i, cfg in enumerate(DOCX_SOURCES.values())}
    recs.sort(key=lambda r: (order[r["subject"]], r["code"]))
    print(f"parsed {len(recs)} content descriptions "
          f"across {len({r['subject'] for r in recs})} subjects")
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(recs, f, indent=1, ensure_ascii=False)
            f.write("\n")
        print(f"wrote {args.output}")
    if args.verify:
        verify(recs, args.verify)
