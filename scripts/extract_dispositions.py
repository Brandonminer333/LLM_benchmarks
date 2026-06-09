"""
Extract tabular case data from "Final Disposition Report 2015 - 2021_Redacted.pdf".

Uses PyMuPDF to read the PDF text layer and parse each page's fixed-position form
layout. No API or OCR required.

Each page contains several case records. A record is:
  Row A : File# Type Court# NewFile# Issued Disposed Div Defendant Race Sex DOB Height Weight SocSec#
  Row B : Prosecuting-Attorney Defense-Attorney Arrest# Report# Police-Dept Officer SSN
  then a "Cnt / Charge / Disp / Sentenced" sub-table with one row per charge.

Columns are left-aligned at x-positions that are constant for every page, so each
word is assigned to the column whose anchor is the greatest anchor <= word.x0.

Outputs two CSVs in outputs/:
  - charges.csv : one row per charge, with the case/defendant fields repeated (denormalized)
  - cases.csv   : one row per case
"""

import csv
import os
import re
from pathlib import Path

import fitz  # PyMuPDF

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")

BASE_DIR = Path(__file__).resolve().parents[1]
PDF_PATH = BASE_DIR / "Final Disposition Report 2015 - 2021_Redacted.pdf"
OUT_DIR = BASE_DIR / "outputs"

# Left-aligned column anchors (verified constant across all pages).
#
# height/weight/soc_sec (Row A) and ssn (Row B) are redacted PII. We keep their
# x-anchors so adjacent columns get correct right-hand boundaries, but we do NOT
# emit them: the redactor deleted the values and only ever leaves stray leftover
# characters (e.g. a lone "6" or "0"). The officer column has no right neighbor
# in our output, so it absorbs the (redacted) ssn region and we strip any trailing
# digit leak from it.
ROW_A = [
    ("file", 20.0), ("type", 89.8), ("court", 122.8), ("new_file", 189.8),
    ("issued", 249.8), ("disposed", 303.7), ("div", 357.7), ("defendant", 393.0),
    ("race", 479.3), ("sex", 544.1), ("dob", 577.9), ("_height", 632.7),
    ("_weight", 676.6), ("_soc_sec", 718.3),
]
ROW_B = [
    ("pros_attorney", 20.0), ("def_attorney", 249.8), ("arrest", 393.0),
    ("report", 479.3), ("police_dept", 544.1), ("officer", 676.6),
]
CHARGE = [("cnt", 374.0), ("charge", 396.9), ("disp", 632.7), ("sentenced", 718.3)]

# Columns that are redacted PII -- consumed for boundaries but never written out.
DROP_FIELDS = {"_height", "_weight", "_soc_sec"}

# A Row A (record start) is the only line with a Disposed *date* in this x-window.
# (Defense-attorney names occasionally land here too, so we require a date to
#  avoid mistaking a Row B continuation line for a new record.)
DISPOSED_COL = (300.0, 316.0)
Y_TOL = 3.0
ROW_A_HALF = 7.0  # vertical window around a Row A baseline

CASE_FIELDS = [
    "case_id", "page", "file", "type", "court", "new_file", "issued", "disposed",
    "div", "defendant", "race", "sex", "dob",
    "pros_attorney", "def_attorney", "arrest", "report", "police_dept", "officer",
    "num_charges",
]
CHARGE_FIELDS = CASE_FIELDS[:-1] + ["cnt", "charge", "disp", "sentenced"]


def strip_trailing_digits(text):
    """Officer names never end in a number; a trailing all-digit token is an
    SSN redaction leak that overflowed into the officer column."""
    parts = text.split()
    while parts and parts[-1].isdigit():
        parts.pop()
    return " ".join(parts)


def assign_column(x0, anchors):
    """Return the field name whose left-aligned anchor best fits x0."""
    chosen = anchors[0][0]
    for name, ax in anchors:
        if x0 >= ax - 2.0:
            chosen = name
        else:
            break
    return chosen


def bucket(words, anchors):
    """Group words into {field: 'joined text'} using left-aligned anchors."""
    cols = {name: [] for name, _ in anchors}
    for w in sorted(words, key=lambda t: (t[0], t[1])):
        cols[assign_column(w[0], anchors)].append((w[1], w[0], w[4]))
    out = {}
    for name, items in cols.items():
        items.sort(key=lambda t: (round(t[0]), t[1]))  # by line, then x
        text = " ".join(t[2] for t in items).strip()
        if text:
            out[name] = text
    return out


def cluster_lines(words, tol=Y_TOL):
    lines = []
    for w in sorted(words, key=lambda t: (t[1], t[0])):
        if lines and abs(w[1] - lines[-1][0]) <= tol:
            lines[-1][1].append(w)
        else:
            lines.append([w[1], [w]])
    return [(y, ws) for y, ws in lines]


def parse_page(page):
    """Return list of case dicts (each with a 'charges' list) for one page."""
    words = [tuple(w[:5]) for w in page.get_text("words")]
    lines = cluster_lines(words)

    # Locate structural anchor lines by y-coordinate.
    rowA_ys, charge_hdr_ys = [], []
    ceiling = page.rect.height  # exclude the "Page X of Y" footer and anything below
    for y, ws in lines:
        texts = {w[4] for w in ws}
        if "Page" in texts and "of" in texts and any(w[4].isdigit() for w in ws):
            ceiling = min(ceiling, y - 2.0)
            continue
        if {"Cnt", "Charge", "Disp"}.issubset(texts):
            charge_hdr_ys.append(y)
            continue
        if "File" in texts and "Defendant" in texts:  # the static label row
            continue
        if "Prosecuting" in texts and "Defense" in texts:  # the static label row
            continue
        if any(DISPOSED_COL[0] <= w[0] <= DISPOSED_COL[1] and DATE_RE.match(w[4])
               for w in ws):
            rowA_ys.append(y)

    rowA_ys.sort()
    charge_hdr_ys.sort()
    cases = []

    for i, ay in enumerate(rowA_ys):
        # The next record's Row A (with its sex overflow band) bounds this record.
        next_ay = rowA_ys[i + 1] if i + 1 < len(rowA_ys) else ceiling
        record_end = min(next_ay - ROW_A_HALF, ceiling)
        # charge header that belongs to this record
        chdr = next((c for c in charge_hdr_ys if c > ay and c < next_ay), None)
        header_end = chdr if chdr is not None else record_end

        rowA_words = [w for w in words if ay - ROW_A_HALF <= w[1] <= ay + ROW_A_HALF]
        rowB_words = [w for w in words if ay + ROW_A_HALF < w[1] < header_end]

        rec = bucket(rowA_words, ROW_A)
        rec.update(bucket(rowB_words, ROW_B))
        for k in DROP_FIELDS:
            rec.pop(k, None)
        if "sex" in rec:
            rec["sex"] = rec["sex"].replace(" ", "")
        if "officer" in rec:  # drop trailing redacted-SSN digit leaks
            rec["officer"] = strip_trailing_digits(rec["officer"])
            if not rec["officer"]:
                rec.pop("officer")

        charges = []
        if chdr is not None:
            charge_words = [w for w in words if chdr < w[1] < record_end and w[0] >= 368]
            for _, ws in cluster_lines(charge_words):
                col = bucket(ws, CHARGE)
                if not col:
                    continue
                cnt = col.get("cnt", "")
                if cnt.isdigit():
                    charges.append({
                        "cnt": cnt,
                        "charge": col.get("charge", ""),
                        "disp": col.get("disp", ""),
                        "sentenced": col.get("sentenced", ""),
                    })
                elif charges:  # wrapped continuation line
                    if col.get("charge"):
                        charges[-1]["charge"] = (charges[-1]["charge"] + " " + col["charge"]).strip()
                    for k in ("disp", "sentenced"):
                        if col.get(k) and not charges[-1][k]:
                            charges[-1][k] = col[k]
        rec["charges"] = charges
        cases.append(rec)

    return cases


def main():
    OUT_DIR.mkdir(exist_ok=True)
    doc = fitz.open(PDF_PATH)
    n_pages = doc.page_count

    cases_path = OUT_DIR / "cases.csv"
    charges_path = OUT_DIR / "charges.csv"

    case_id = 0
    n_charges = 0
    with open(cases_path, "w", newline="") as cf, open(charges_path, "w", newline="") as chf:
        cw = csv.DictWriter(cf, fieldnames=CASE_FIELDS)
        chw = csv.DictWriter(chf, fieldnames=CHARGE_FIELDS)
        cw.writeheader()
        chw.writeheader()

        for pi in range(n_pages):
            for rec in parse_page(doc[pi]):
                case_id += 1
                charges = rec.pop("charges")
                base = {k: rec.get(k, "") for k in CASE_FIELDS}
                base["case_id"] = case_id
                base["page"] = pi + 1
                base["num_charges"] = len(charges)
                cw.writerow(base)
                if charges:
                    for c in charges:
                        row = dict(base)
                        row.pop("num_charges", None)
                        row.update(c)
                        chw.writerow(row)
                        n_charges += 1
                else:
                    row = dict(base)
                    row.pop("num_charges", None)
                    row.update({"cnt": "", "charge": "", "disp": "", "sentenced": ""})
                    chw.writerow(row)
                    n_charges += 1

            if (pi + 1) % 500 == 0:
                print(f"  ...processed {pi + 1}/{n_pages} pages, {case_id} cases")

    print(f"Done. {case_id} cases, {n_charges} charge rows from {n_pages} pages.")
    print(f"  {cases_path}")
    print(f"  {charges_path}")


if __name__ == "__main__":
    main()
