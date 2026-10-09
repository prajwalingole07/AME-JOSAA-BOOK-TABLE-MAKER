"""
MS Word automation for the updated JoSAA book.

Two jobs, both driven through the Word COM automation model (pywin32) because
only Word itself can lay the pages out:

  1. fit_all_tables()      - every cut-off table must sit on ONE landscape A4
                             page.  Tables that break are shrunk (font size)
                             and re-checked, up to MAX_PASSES times.
  2. fix_index_pages()     - the "Index / Page No." table in the front of the
                             book is rewritten with the real page numbers of
                             the finished document.  Entries 6-9 (IIT / NIT /
                             IIIT / GFTI sections) are taken from the first
                             cut-off table of that section, every other entry
                             is located by searching for its heading text.

Both functions work in place on the given .docx file.
"""

import os
import re

import win32com.client

WD_COLLAPSE_START = 1
WD_ACTIVE_END_PAGE_NUMBER = 3      # WdInformation.wdActiveEndPageNumber
WD_FIND_STOP = 0                   # WdReplace.wdFindStop
WD_DO_NOT_SAVE_CHANGES = 0
WD_SAVE_CHANGES = -1

MIN_FONT_PT = 5.5                  # never shrink a table below this size
MAX_PASSES = 6
SHRINK_STEP = 0.5                  # points taken off per pass

SECTION_KINDS = ("IIT", "NIT", "IIIT", "GFTI")


# --------------------------------------------------------------------------- #
#  helpers
# --------------------------------------------------------------------------- #

def _clean(txt):
    return (txt or "").replace("\x07", "").replace("\r", "").replace("\x0b", " ").strip()


def _kind_of(name):
    m = re.search(r"\((IIT|NIT|IIIT|IIEST)\)\s*$", name)
    if not m:
        return "GFTI"
    tag = m.group(1)
    return "NIT" if tag in ("NIT", "IIEST") else tag


def _is_cutoff_table(table):
    try:
        return _clean(table.Cell(1, 2).Range.Text) == "College Name & Type"
    except Exception:
        return False


def _table_name(table):
    try:
        return _clean(table.Cell(1, 3).Range.Text)
    except Exception:
        return ""


def _page_of(rng):
    return int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))


def _start_page_of(table):
    """Page the table starts on - taken from its first cell, never from the
    range start (that sits in the page-break paragraph before the table)."""
    return _page_of(table.Cell(1, 1).Range)


def _end_page_of(table):
    return _page_of(table.Range)


def _shrink_table(table):
    """Take SHRINK_STEP points off every paragraph of the table.

    Returns False when the table is already at MIN_FONT_PT.
    """
    sizes = []
    for para in table.Range.Paragraphs:
        try:
            size = para.Range.Font.Size
        except Exception:
            size = None
        if size is None:            # mixed runs inside one paragraph
            size = MIN_FONT_PT + SHRINK_STEP
        sizes.append(size)
    if not sizes or min(sizes) - SHRINK_STEP < MIN_FONT_PT:
        return False

    for para in table.Range.Paragraphs:
        try:
            size = para.Range.Font.Size
            if size is None:
                size = MIN_FONT_PT + SHRINK_STEP
            para.Range.Font.Size = max(MIN_FONT_PT, size - SHRINK_STEP)
        except Exception:
            pass
        try:
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        except Exception:
            pass
    return True


def open_word():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


# --------------------------------------------------------------------------- #
#  1. one page per cut-off table
# --------------------------------------------------------------------------- #

def fit_all_tables(path, progress_cb=None):
    """Shrink any cut-off table that spills onto a second page."""
    cb = progress_cb or (lambda pct, msg: None)
    word = open_word()
    report = {"checked": 0, "shrunk": [], "still_split": []}
    try:
        doc = word.Documents.Open(os.path.abspath(path))
        doc.Repaginate()

        tables = [doc.Tables.Item(i + 1) for i in range(doc.Tables.Count)]
        targets = [(i + 1, t) for i, t in enumerate(tables) if _is_cutoff_table(t)]
        report["checked"] = len(targets)

        for idx, table in targets:
            name = _table_name(table) or f"table {idx}"
            passes = 0
            while passes < MAX_PASSES:
                start = _start_page_of(table)
                end = _end_page_of(table)
                if start == end:
                    cb(100, f"OK (1 page): {name[:52]}")
                    break
                if not _shrink_table(table):
                    report["still_split"].append(f"{name} (pages {start}-{end})")
                    cb(100, f"STILL SPLIT: {name[:52]}")
                    break
                doc.Repaginate()
                passes += 1
                report["shrunk"].append(f"{name}: pass {passes}")
            else:
                report["still_split"].append(f"{name} (max passes)")

        doc.Save()
        doc.Close(WD_SAVE_CHANGES)
    finally:
        word.Quit()
    cb(100, "Page-fit audit finished")
    return report


# --------------------------------------------------------------------------- #
#  2. index page numbers
# --------------------------------------------------------------------------- #

# Word's physical page number sits 2 pages above the printed folio (two
# unnumbered cover pages at the front).  Verified against the original
# template: 'Reasons...' is physical page 3 but printed page 1.
PAGE_OFFSET = 2

# Index label -> a distinctive phrase that actually appears as a heading in
# the body.  Written by hand where the index wording differs from the heading
# (Word's Find needs the exact body text).
HEADINGS = {
    "What is JOSAA and CSAB in Brief.": "What is JOSAA?",
    "List of Required Documents for JOSAA/CSAB":
        "List of documents to be produced by the candidate",
    "Important Websites of Central Process.":
        "Central Admission Process-Important Websites.",
    "List of Required Documents for Maharashtra State":
        "Documents for Admission-All Type-A and Type-B Students",
    "Details of Tuition Fee Waiver Scheme (TFWS).":
        "Tuition Fee Waiver Scheme-TFWS",
    "Quota wise Seat Distribution.": "Type of Institution",
    "Important Websites of State Process.":
        "State Admission Process-Important Websites.",
    "Government/ Government Aided Engineering Colleges":
        "Government / Govt. Aided / University Departments.",
    "Details: Non-Creamy Layer Certificate (NCL).":
        "Important Information about Non-Creamy Layer Certificate (NCL)",
    "Top Private Engineering Colleges Nearby Pune.":
        "Top Private Engineering Colleges Nearby Pune.",
    "Top Private Engineering Colleges Nearby Mumbai.":
        "Top Private Engineering Colleges Nearby Mumbai.",
    "Top Private Engineering Colleges in Maharashtra. (Other than Pune & Mumbai)":
        "Top Private Engineering Colleges in Maharashtra. (Other than Pune & Mumbai)",
}

SKIP_LABELS = {"index", "sr."}

# The three "Top Private Engineering Colleges ..." sections have no heading
# table in the body - they start straight with a college table.  Anchor each
# on the five digit college code of its first college.
CODE_ANCHORS = {
    "Government/ Government Aided Engineering Colleges in Maharashtra.": "02020",
    "Top Private Engineering Colleges Nearby Pune.": "06141",
    "Top Private Engineering Colleges Nearby Mumbai.": "03176",
    "Top Private Engineering Colleges in Maharashtra. (Other than Pune & Mumbai)":
        "06267",
}


def _page_of_code(doc, code):
    """Physical page of the first cut-off-style table whose info rows mention
    the five digit college code."""
    for i in range(1, doc.Tables.Count + 1):
        table = doc.Tables.Item(i)
        for r in (1, 2):
            try:
                text = _clean(table.Cell(r, 2).Range.Text) + " " + \
                    _clean(table.Cell(r, 3).Range.Text)
            except Exception:
                continue
            if code in text:
                return _start_page_of(table)
    return None


def _heading_phrase(label):
    """Exact body-heading text for an index label, if we have one mapped."""
    for key, phrase in HEADINGS.items():
        if label.startswith(key[:40]):
            return phrase
    return None


def _first_table_pages(doc):
    """Word physical page of the first IIT / NIT / IIIT / GFTI cut-off table."""
    found = {}
    for i in range(1, doc.Tables.Count + 1):
        table = doc.Tables.Item(i)
        if not _is_cutoff_table(table):
            continue
        kind = _kind_of(_table_name(table))
        if kind not in found:
            found[kind] = _start_page_of(table)
    return found


def _find_page(doc, search_from, phrase):
    """Physical page of the first occurrence of `phrase` at/after search_from."""
    if not phrase or len(phrase) < 6:
        return None
    rng = doc.Range(search_from, doc.Content.End)
    finder = rng.Find
    finder.ClearFormatting()
    finder.Text = phrase
    finder.Forward = True
    finder.Wrap = WD_FIND_STOP
    finder.MatchCase = False
    finder.MatchWildcards = False
    if finder.Execute():
        return _page_of(rng)
    return None


def _norm_key(text):
    s = re.sub(r"\s+", " ", (text or "")).strip().lower().rstrip(" .")
    return s


def _body_headings(doc):
    """Every non-cut-off table title -> its physical page.

    Section headings in this book live in the first row of a table, so a
    table scan is far more reliable than Word's Find (which happily matches
    the Index entry itself).
    """
    found = []
    for i in range(1, doc.Tables.Count + 1):
        table = doc.Tables.Item(i)
        if _is_cutoff_table(table):
            continue
        try:
            text = _clean(table.Cell(1, 1).Range.Text)
        except Exception:
            continue
        key = _norm_key(text)
        if len(key) < 8:
            continue
        found.append((key, _start_page_of(table)))
    return found


def _page_from_headings(headings, label):
    """Best physical page for an index label from the table-title scan."""
    want = _norm_key(label)
    if len(want) < 8:
        return None
    for key, page in headings:          # exact / prefix on the heading side
        if key == want or key.startswith(want) or want.startswith(key):
            return page
    return None


def _search_variants(label, phrase):
    """Phrases to try, most specific first.  Body headings often drop the
    trailing full stop the index uses, so try both."""
    out = []
    if phrase:
        out.append(phrase)
    out.append(label)
    out.append(label.rstrip(" ."))
    out.append(label[:52].rstrip(" ."))
    seen, uniq = set(), []
    for s in out:
        s = re.sub(r"\s+", " ", s).strip()
        if s and s not in seen:
            seen.add(s)
            uniq.append(s)
    return uniq


def fix_index_pages(path, progress_cb=None):
    """Rewrite the 'Page No.' column of the front Index table.

    Every index entry is located in the body and its printed page number is
    written back.  Section entries (IIT/NIT/IIIT/GFTI) are taken from the
    first cut-off table of that section instead of a text search.
    """
    cb = progress_cb or (lambda pct, msg: None)
    word = open_word()
    updated, missing = [], []
    try:
        doc = word.Documents.Open(os.path.abspath(path))
        doc.Repaginate()

        index = doc.Tables.Item(1)
        if _clean(index.Cell(1, 2).Range.Text).lower() != "index":
            for i in range(1, doc.Tables.Count + 1):
                cand = doc.Tables.Item(i)
                try:
                    if _clean(cand.Cell(1, 2).Range.Text).lower() == "index":
                        index = cand
                        break
                except Exception:
                    continue

        section_pages = _first_table_pages(doc)
        search_from = index.Range.End
        headings = _body_headings(doc)

        for r in range(1, index.Rows.Count + 1):
            row = index.Rows.Item(r)
            if row.Cells.Count < 3:
                continue
            label = _clean(row.Cells.Item(2).Range.Text)
            if not label or label.lower() in SKIP_LABELS:
                continue

            phys = None
            m = re.search(r"Cut-off:\s*(IITs|NITs|IIITs|Some other GFTIs)", label)
            if m:
                key = {"IITs": "IIT", "NITs": "NIT", "IIITs": "IIIT",
                       "Some other GFTIs": "GFTI"}[m.group(1)]
                phys = section_pages.get(key)

            if phys is None:
                code = CODE_ANCHORS.get(label)
                if code:
                    phys = _page_of_code(doc, code)

            if phys is None:
                phys = _page_from_headings(headings, label)

            if phys is None:
                for phrase in _search_variants(label, _heading_phrase(label)):
                    phys = _find_page(doc, search_from, phrase)
                    if phys is not None:
                        break

            if phys is None:
                missing.append(label)
                continue

            page = max(1, phys - PAGE_OFFSET)
            old = _clean(row.Cells.Item(3).Range.Text)
            row.Cells.Item(3).Range.Text = str(page)
            updated.append(f"{label[:48]}: {old} -> {page}")
            cb(0, f"Index {r}: {label[:44]} = {page}")

        doc.Save()
        doc.Close(WD_SAVE_CHANGES)
    finally:
        word.Quit()
    cb(100, "Index page numbers written")
    return {"updated": updated, "missing": missing}


# --------------------------------------------------------------------------- #
#  CLI
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys

    target = (sys.argv[1] if len(sys.argv) > 1 else
              r"C:\Users\prajw\Downloads\ENGINNERING FINAL BOOK 2027 - "
              r"UPDATED IIT NIT IIIT.docx")

    def _cb(_pct, msg):
        if msg:
            print("  *", msg, flush=True)

    print("== page-fit audit ==")
    fit = fit_all_tables(target, progress_cb=_cb)
    print("   tables checked :", fit["checked"])
    print("   shrunk         :", len(fit["shrunk"]))
    for s in fit["shrunk"]:
        print("      ", s)
    print("   still split    :", fit["still_split"])

    print("== index ==")
    idx = fix_index_pages(target, progress_cb=_cb)
    for u in idx["updated"]:
        print("   ", u)
    print("    not found:", idx["missing"])
