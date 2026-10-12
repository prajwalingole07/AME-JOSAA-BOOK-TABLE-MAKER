"""
Book Updater for the ENGINEERING FINAL BOOK (JoSAA IIT / NIT / IIIT sections).

Reads the user supplied Word template and the latest JoSAA Excel workbook, then
rebuilds every IIT / NIT / IIIT / GFTI cut-off table in place so that:

  * every institute gets its OWN table,
  * every table starts on a fresh page and fits inside ONE landscape A4 page,
  * NIT Nagpur, NIT Jaipur and NIT Warangal get the combined
    "50% Home State Quota" + "50% Other State Quota" blocks,
  * every other NIT gets the "50% Other State Quota" block only,
  * IIT / IIIT keep their single All-India block,
  * rank data is taken from the Excel workbook (PwD / Defence / Foreign dropped),
  * header info box, serial number and branch seat counts are preserved from
    the template (they do not exist in the Excel).

Everything outside tables 7..67 of the template is left byte-identical.
"""

import copy
import math
import os
import re
from difflib import SequenceMatcher

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from parser import JosaaExcelParser

# --------------------------------------------------------------------------- #
#  Constants
# --------------------------------------------------------------------------- #

CAT_ORDER = [
    "OPEN-N", "OPEN-W",
    "OBC-N", "OBC-W",
    "EWS-N", "EWS-W",
    "SC-N", "SC-W",
    "ST-N", "ST-W",
]

# NITs that must show BOTH the Home State and the Other State block.
COMBINED_QUOTA_KEYWORDS = ("jaipur", "nagpur", "warangal")

MISSING = "-----"

# Height budget for one landscape A4 page (inches) - margins taken from template.
PAGE_HEIGHT_IN = 8.27
MARGIN_TOP_IN = 0.4
MARGIN_BOTTOM_IN = 0.4
# leave a little slack so the page-break paragraph never creates a blank page
FIT_SAFETY = 0.96

LINE_FACTOR = 1.20          # conservative Word single line-height factor
MIN_SCALE = 0.62            # never shrink below 62 % of the template font size


# --------------------------------------------------------------------------- #
#  Small helpers
# --------------------------------------------------------------------------- #

def _norm_text(s):
    """Aggressive normaliser used to match institutes / branches."""
    s = str(s or "").replace("\xa0", " ")
    s = s.lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\((4|5|6|8)\s*years.*?\)", " ", s)
    s = re.sub(r"\[[^\]]*\]", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


_ABBR = {
    "eng": "engineering",
    "comm": "communication",
    "sci": "science",
    "tech": "technology",
    "info": "information",
    "math": "mathematics",
    "maths": "mathematics",
    "metall": "metallurgical",
    "electron": "electronics",
    "comp": "computer",
    "biotech": "biotechnology",
    "biochem": "biochemical",
    "chem": "chemical",
    "mech": "mechanical",
    "int": "intelligence",
    "mgmt": "management",
    "me": "mechanical",
    "ce": "civil engineering",
    "cse": "computer science and engineering",
    "ece": "electronics and communication engineering",
    "eee": "electrical and electronics engineering",
    "ai": "artificial intelligence",
    "ml": "machine learning",
    "ds": "data science",
    "it": "information technology",
    "aids": "artificial intelligence and data science",
    "aiml": "artificial intelligence and machine learning",
    "bioengineering": "bio engineering",
    "lntelligence": "intelligence",     # JoSAA sheet typo (lower-case L)
}

# programmes the book and the workbook name differently (same course)
_RENAME = {
    ("applied geology", "geological engineering"),
    ("applied geophysics", "geophysical technology"),
    ("bioengineering", "biomedical engineering"),
}

# degree / boilerplate words that carry no meaning for branch matching
_NOISE = {
    "and", "of", "the", "in", "with", "for", "to", "a", "an", "on", "by", "at",
    "from", "as", "is", "are", "b", "m", "ms", "dd", "degree", "dual", "years",
    "year", "bachelor", "master", "mtech", "iim", "plus", "bs", "phd", "s", "e",
}


def _expand_abbr(s):
    out = []
    for w in s.split():
        out.append(_ABBR.get(w, w))
    return " ".join(out)


def norm_branch(s):
    """Normalised branch key used for template <-> Excel matching."""
    return _expand_abbr(_norm_text(s))


def norm_institute(s):
    s = _norm_text(s)
    s = re.sub(r"\b(nit|iit|iiit|iiest)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _similar(a, b, threshold=0.72):
    if not a or not b:
        return False
    if a == b:
        return True
    if a in b or b in a:
        return True
    return SequenceMatcher(None, a, b).ratio() >= threshold


# Generic words that carry no institute identity - everything else (city /
# state / founder name) is a distinguishing token used for matching.
_INST_NOISE = frozenset({
    "indian", "institute", "institutes", "information", "technology",
    "national", "of", "the", "and", "or", "in", "for", "at", "on",
    "iit", "nit", "iiit", "iiest", "iiitm", "iist",
    "international", "campus", "design", "manufacturing", "manufac",
    "engineering", "university", "college", "school", "dept", "department",
})


def _inst_tokens(s):
    """Distinctive identity tokens of an institute name (city/state/…)."""
    return {t for t in re.findall(r"[a-z0-9]+", str(s or "").lower())
            if len(t) > 1 and t not in _INST_NOISE}


# --------------------------------------------------------------------------- #
#  Branch matching (template row  <->  Excel programme)
# --------------------------------------------------------------------------- #

def _match_tokens(name):
    """Token list used only for scoring - boilerplate stripped, abbreviations
    expanded, duplicates preserved (they are matched one by one)."""
    toks = []
    for w in _norm_text(name).split():
        if w in _NOISE:
            continue
        for part in _ABBR.get(w, w).split():
            if part in _NOISE:
                continue
            toks.append(part)
    return toks


def _tok_eq(a, b):
    """Fuzzy single-token equality (materials ~ material, metallurgy ~
    metallurgical, engineer ~ engineering)."""
    if a == b:
        return True
    if min(len(a), len(b)) < 5:
        return False
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    if n >= 5 and n >= 0.8 * min(len(a), len(b)):
        return True
    return SequenceMatcher(None, a, b).ratio() >= 0.85


def branch_score(tpl_name, tpl_period, xl_name, xl_period, ignore_period=False):
    """Similarity in [0, 1]; 0 means 'these are different programmes'.

    Course period is a hard constraint: a 5-year row may never take the ranks
    of the 4-year programme of the same name (and vice versa) - unless
    `ignore_period` is set, which is only used for leftover rows whose names
    are identical anyway.
    """
    if (_norm_text(tpl_name), _norm_text(xl_name)) in _RENAME:
        return 0.95
    if not ignore_period and tpl_period and xl_period and \
            str(tpl_period) != str(xl_period):
        return 0.0
    A = _match_tokens(tpl_name)
    B = _match_tokens(xl_name)
    if not A or not B:
        return 0.0

    used, hit = set(), 0
    for a in A:
        for j, b in enumerate(B):
            if j in used:
                continue
            if _tok_eq(a, b):
                used.add(j)
                hit += 1
                break
    if hit == 0:
        return 0.0
    if hit == len(A) == len(B):
        return 1.0
    union = len(A) + len(B) - hit
    jaccard = hit / union
    containment = hit / min(len(A), len(B))
    if containment >= 0.999 and hit >= 2:
        # smaller name fully contained in the longer one
        return max(jaccard, 0.60)
    if containment >= 0.999 and min(len(A), len(B)) == 1:
        # 'Physics' inside 'Engineering Physics', 'Mechanical' inside ...
        return max(jaccard, 0.60)
    if containment >= 0.70 and hit >= 3:
        # 70 % of the book's wording found in the workbook name
        # ('AI & Data Sci. with MBA in Hospital' vs '... MBA in Digital ...')
        return 0.50
    if jaccard >= 0.40 and hit >= 2:
        return jaccard
    return 0.0


def assign_branches(tpl_rows, xl_entries, threshold=0.40):
    """Greedy global assignment: strongest pairs first, one-to-one.

    Returns ({xl_index: tpl_index}, [unmatched tpl index], [leftover xl index]).
    A second pass lets an unmatched row take an identical name whose course
    period differs (the workbook re-classified the programme).
    """
    pairs, unmatched, leftover = _assign(tpl_rows, xl_entries, threshold, False)
    if unmatched and leftover:
        pairs2, unmatched2, leftover2 = _assign(
            [tpl_rows[i] for i in unmatched],
            [xl_entries[j] for j in leftover],
            0.99, True)
        remap = {}
        for j2, i2 in pairs2.items():
            remap[leftover[j2]] = unmatched[i2]
        pairs.update(remap)
        unmatched = [unmatched[i] for i in unmatched2]
        leftover = [leftover[j] for j in leftover2]
    return pairs, unmatched, leftover


def _assign(tpl_rows, xl_entries, threshold, ignore_period):
    edges = []
    for i, tr in enumerate(tpl_rows):
        for j, e in enumerate(xl_entries):
            sc = branch_score(tr["name"], tr.get("period") or "",
                              e["short"], e.get("period") or "",
                              ignore_period=ignore_period)
            if sc >= threshold:
                edges.append((sc, i, j))
    edges.sort(key=lambda t: -t[0])

    taken_t, taken_x, pairs = set(), set(), {}
    for sc, i, j in edges:
        if i in taken_t or j in taken_x:
            continue
        taken_t.add(i)
        taken_x.add(j)
        pairs[j] = i
    unmatched = [i for i in range(len(tpl_rows)) if i not in taken_t]
    leftover = [j for j in range(len(xl_entries)) if j not in taken_x]
    return pairs, unmatched, leftover


def _branch_matches(a, b):
    return branch_score(a, "", b, "") >= 0.40



def _short_branch_name(program_name):
    """'Civil Engineering (4 Years, Bachelor of Technology)' -> 'Civil Engineering'."""
    s = str(program_name or "").strip()
    s = re.sub(r"\s*\((4|5|6|8)\s*Years.*$", "", s)
    s = re.sub(r"\s*\([^)]*Years.*$", "", s)
    s = s.replace(" and ", " & ")
    return s.strip(" ,")


def _course_period(program_name):
    m = re.search(r"\((\d+)\s*Years", str(program_name or ""))
    if m:
        return m.group(1)
    m = re.search(r"\[(\d+)/(\d+)\]", str(program_name or ""))
    return m.group(2) if m else ""


def parse_rank(val):
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    m = re.search(r"\d+", s)
    return int(m.group(0)) if m else None


# --------------------------------------------------------------------------- #
#  Template scanning
# --------------------------------------------------------------------------- #

def _unique_cells(row):
    """Row cells without duplicates introduced by horizontal merges."""
    out, seen = [], set()
    for c in row.cells:
        if c._tc in seen:
            continue
        seen.add(c._tc)
        out.append(c)
    return out


def _row_texts(row):
    return [c.text.strip() for c in _unique_cells(row)]


def _is_info_row(row):
    t = _row_texts(row)
    return len(t) >= 3 and t[1] == "College Name & Type"


def _is_cat_row(row):
    t = _row_texts(row)
    return bool(t) and t[0] in ("Category-Gender", "Category - Gender") and "OPEN-N" in t


def _is_open_close_row(row):
    t = _row_texts(row)
    return "Open" in t and "Close" in t


def _is_quota_row(row):
    t = _row_texts(row)
    return len(t) <= 3 and any("Quota" in x for x in t)


def _is_data_row(row):
    t = _row_texts(row)
    if len(t) < 21:
        return False
    if t[0].startswith("Category-Gender"):
        return False
    if "Open" in t and "Close" in t:
        return False
    if any("Quota" in x for x in t):
        return False
    if len(t) > 1 and t[1] == "College Name & Type":
        return False
    return bool(re.search(r"\d{2,}", " ".join(t[1:])) or "-" in " ".join(t[1:]))


def _classify_table(tbl_el):
    """Is this a cut-off table (info box + rank grid)?"""
    rows = tbl_el.findall(qn("w:tr"))
    if len(rows) < 5:
        return False
    first = rows[0]
    texts = ["".join(t.text or "" for t in tc.iter(qn("w:t"))) for tc in first.findall(qn("w:tc"))]
    return any("College Name & Type" in x for x in texts)


def _parse_label(label):
    """'Computer Science & Engineering [68/4]' -> (name, seats, period)."""
    s = str(label or "").replace("\xa0", " ").strip()
    m = re.search(r"\[([^\]]*)\]\s*$", s)
    seats = period = ""
    if m:
        spec = m.group(1).strip()
        s = s[: m.start()].strip()
        if "/" in spec:
            a, b = spec.split("/", 1)
            seats, period = a.strip(), b.strip()
        else:
            seats = spec
    return s.strip(), seats, period


def _dd_label(label, period="", name="", program=""):
    """Put '(DD)' after a branch only where '(Dual Degree)' is written."""
    lab = str(label or "")
    has_marker = bool(re.search(r"\(DD\)", lab, re.I))
    is_dd = (has_marker
             or bool(re.search(r"Dual\s+Degree",
                               f"{lab} {name or ''} {program or ''}", re.I)))
    if not lab.strip() or not is_dd:
        return lab
    lab = re.sub(r"\s*\(DD\)\s*", " ", lab, flags=re.I)   # drop old markers
    lab = re.sub(r"\s+", " ", lab).strip()
    return lab + " (DD)"


def _block_kind(name):
    """The book itself labels every institute - trust that label.

    '(IIT)' / '(NIT)' / '(IIIT)' / '(IIEST)' at the end of the printed name
    decides which Excel sheet feeds the block; anything without such a suffix
    is a GFTI and keeps the template's own data.
    """
    m = re.search(r"\((IIT|NIT|IIIT|IIEST)\)\s*$", str(name or ""))
    if not m:
        return "GFTI"
    tag = m.group(1)
    if tag in ("NIT", "IIEST"):
        return "NIT"
    return tag


def scan_template(doc):
    """
    Splits every cut-off table of the template into per-institute blocks.

    Returns a list of dicts:
        {
          'kind'      : 'IIT' | 'NIT' | 'IIIT' | 'GFTI',
          'name'      : full college name as printed in the book,
          'info_rows' : [row element x3] (verbatim header box),
          'cat_row'   : category token header row element,
          'oc_row'    : Open / Close header row element,
          'div_rows'  : {label_lower: row element} quota divider templates,
          'data_tpl'  : data row element used as a template,
          'sections'  : [{'label': str, 'branches': [{'label','ranks'}]}],
        }
    """
    blocks = []
    for tbl in doc.tables:
        tbl_el = tbl._tbl
        if not _classify_table(tbl_el):
            continue

        rows = list(tbl.rows)
        # locate info rows of every institute inside this table
        starts = [i for i, r in enumerate(rows) if _is_info_row(r)]
        for si, start in enumerate(starts):
            end = starts[si + 1] if si + 1 < len(starts) else len(rows)
            chunk = rows[start:end]

            info_rows = [r._tr for r in chunk[:3]]
            cat_row = oc_row = None
            div_rows = {}
            data_tpl = None
            sections = []
            current = None

            for r in chunk[3:]:
                if _is_cat_row(r):
                    cat_row = r._tr
                    continue
                if _is_open_close_row(r) and not _is_data_row(r):
                    oc_row = r._tr
                    continue
                if _is_quota_row(r):
                    label = [x for x in _row_texts(r) if "Quota" in x]
                    txt = label[0] if label else "Quota"
                    div_rows[txt.lower()] = r._tr
                    current = {"label": txt, "branches": []}
                    sections.append(current)
                    continue
                if _is_data_row(r):
                    if current is None:
                        current = {"label": "", "branches": []}
                        sections.append(current)
                    texts = _row_texts(r)
                    name, seats, period = _parse_label(texts[0])
                    ranks = {}
                    for i, tok in enumerate(CAT_ORDER):
                        op = parse_rank(texts[1 + i * 2])
                        cl = parse_rank(texts[2 + i * 2])
                        ranks[tok] = (op, cl)
                    current["branches"].append(
                        {"label": texts[0], "name": name, "seats": seats,
                         "period": period, "ranks": ranks}
                    )
                    if data_tpl is None:
                        data_tpl = r._tr

            name = _row_texts(rows[start])[2] if len(_row_texts(rows[start])) > 2 else ""
            blocks.append({
                "kind": _block_kind(name),
                "name": name,
                "info_rows": info_rows,
                "cat_row": cat_row,
                "oc_row": oc_row,
                "div_rows": div_rows,
                "data_tpl": data_tpl,
                "sections": sections,
                "tbl_el": tbl_el,
                "grid_w": tbl_el.tblGrid,
            })
    return blocks


# --------------------------------------------------------------------------- #
#  Excel side
# --------------------------------------------------------------------------- #

def build_excel_index(excel_path, progress_cb=None):
    """processed[ kind ] -> { norm_institute : [ {program, short, period, quota, ranks} ] }"""
    parser = JosaaExcelParser(excel_path)
    if progress_cb:
        progress_cb(5, "Reading Excel workbook ...")
    parser.inspect()
    if progress_cb:
        progress_cb(18, "Consolidating cut-off ranks ...")
    processed = parser.process_all(category_order_mode="IMAGE1")

    index = {}
    for kind, payload in processed.items():
        table = {}
        for inst in payload["institutes"]:
            entries = []
            for br in inst["branches"]:
                entries.append({
                    "program": br["program_name"],
                    "short": _short_branch_name(br["program_name"]),
                    "period": _course_period(br["program_name"]),
                    "quota": str(br.get("quota", "AI")).strip().upper(),
                    "ranks": br.get("ranks", {}),
                })
            table[norm_institute(inst["institute_name"])] = {
                "raw_name": inst["institute_name"],
                "entries": entries,
            }
        index[kind] = table
    if progress_cb:
        progress_cb(30, "Excel index ready")
    return index, processed


def _match_one(index_for_kind, block_name):
    key = norm_institute(block_name)
    if key in index_for_kind:
        return index_for_kind[key]
    # SequenceMatcher ratios inside one institute family are all ~0.9
    # (they share the 40-char "indian institute of information …" prefix),
    # so the discriminating signal is the distinctive token(s) - city, state
    # or founder name.  Prefer the candidate sharing the most of them and
    # only fall back to plain similarity when no token overlaps at all.
    btoks = _inst_tokens(key)
    best, best_score = None, 0.0
    tok_best, tok_ov, tok_score = None, 0, 0.0
    for k, v in index_for_kind.items():
        if not k or not key:
            continue
        score = SequenceMatcher(None, key, k).ratio()
        if key in k or k in key:
            score = max(score, 0.9)
        if score > best_score:
            best, best_score = v, score
        if btoks:
            ov = len(btoks & _inst_tokens(k))
            if ov > tok_ov or (ov == tok_ov and ov and score > tok_score):
                tok_best, tok_ov, tok_score = v, ov, score
    if tok_best is not None:
        return tok_best
    return best if best_score >= 0.70 else None


def _find_institute(excel_index, block_name, preferred_kind):
    """Only the sheet implied by the printed label is consulted - a fuzzy hit
    in a foreign sheet can never be trusted (e.g. Birla Institute of
    Technology, Patna vs IIT Patna)."""
    return _match_one(excel_index.get(preferred_kind, {}), block_name)


# --------------------------------------------------------------------------- #
#  Row construction helpers
# --------------------------------------------------------------------------- #

def _set_tc_text(tc, text):
    """Replace the text of a table cell keeping paragraph + run formatting."""
    ps = tc.findall(qn("w:p"))
    if not ps:
        p = OxmlElement("w:p")
        tc.append(p)
        ps = [p]
    for p in ps[1:]:
        tc.remove(p)
    p = ps[0]
    rpr = None
    runs = p.findall(qn("w:r"))
    if runs:
        rp = runs[0].find(qn("w:rPr"))
        if rp is not None:
            rpr = copy.deepcopy(rp)
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)


def _clone_row(tpl_tr, texts):
    """Deep-copy a template row and pour `texts` into its cells."""
    tr = copy.deepcopy(tpl_tr)
    tcs = tr.findall(qn("w:tc"))
    for i, tc in enumerate(tcs):
        _set_tc_text(tc, texts[i] if i < len(texts) else "")
    return tr


def _make_quota_row(data_tpl, quota_label):
    """Build a divider row (branch label + one merged quota cell) from a data row."""
    tr = copy.deepcopy(data_tpl)
    tcs = tr.findall(qn("w:tc"))
    if len(tcs) < 2:
        return tr
    _set_tc_text(tcs[0], "Branch [Seats/Course Period]")
    span = 0
    for tc in tcs[1:]:
        g = 1
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is not None:
            gs = tcPr.find(qn("w:gridSpan"))
            if gs is not None:
                g = int(gs.get(qn("w:val")) or 1)
        span += g
    keep = tcs[1]
    tcPr = keep.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        keep.insert(0, tcPr)
    for old in tcPr.findall(qn("w:gridSpan")):
        tcPr.remove(old)
    gs = OxmlElement("w:gridSpan")
    gs.set(qn("w:val"), str(max(1, span)))
    tcPr.append(gs)
    _set_tc_text(keep, quota_label)
    for tc in tcs[2:]:
        tr.remove(tc)
    # centre the quota label
    p = keep.findall(qn("w:p"))[0]
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    for old in pPr.findall(qn("w:jc")):
        pPr.remove(old)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pPr.append(jc)
    return tr


def _force_cant_split(tr):
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def _scale_row_font(tr, scale):
    if abs(scale - 1.0) < 0.005:
        return
    for r in tr.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        for tag in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            cur = el.get(qn("w:val"))
            try:
                cur_v = float(cur)
            except (TypeError, ValueError):
                cur_v = 15.0
            el.set(qn("w:val"), str(max(8, int(round(cur_v * scale)))))


def _run_font_pt(tr, default=7.5):
    for r in tr.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is not None:
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                try:
                    return float(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
    return default


def _cell_pad_pt(tr, default=2.0):
    for tc in tr.findall(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            continue
        mar = tcPr.find(qn("w:tcMar"))
        if mar is None:
            continue
        tot = 0.0
        for side in ("top", "bottom"):
            el = mar.find(qn("w:" + side))
            if el is not None:
                try:
                    tot += float(el.get(qn("w:w")) or 0) / 20.0
                except (TypeError, ValueError):
                    pass
        if tot:
            return tot
    return default


def _para_space_pt(tr, default=2.0):
    for p in tr.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        sp = pPr.find(qn("w:spacing"))
        if sp is None:
            continue
        b = a = 0.0
        try:
            b = float(sp.get(qn("w:before")) or 0) / 20.0
            a = float(sp.get(qn("w:after")) or 0) / 20.0
        except (TypeError, ValueError):
            pass
        return b + a
    return default


def _row_height_pt(tr, font_pt, pad_pt, space_pt, chars_per_line):
    lines = 1
    for tc in tr.findall(qn("w:tc")):
        txt = "".join(t.text or "" for t in tc.iter(qn("w:t")))
        if txt:
            n = math.ceil(len(txt) / max(1, chars_per_line))
            lines = max(lines, n)
    return lines * font_pt * LINE_FACTOR + space_pt + pad_pt


def _branch_col_chars(font_pt, col_w_in):
    usable_pt = col_w_in * 72 - 6.0
    return max(6, int(usable_pt / (font_pt * 0.50)))


# --------------------------------------------------------------------------- #
#  Main updater
# --------------------------------------------------------------------------- #

class BookUpdater:
    def __init__(self, excel_path, template_path, output_path, progress_cb=None):
        self.excel_path = excel_path
        self.template_path = template_path
        self.output_path = output_path
        self.cb = progress_cb or (lambda pct, msg: None)
        self.report = {
            "institutes": 0, "matched": 0, "unmatched_institutes": [],
            "rows": 0, "rank_cells": 0, "new_branches": [],
            "shrink_log": [],
        }

    # ---------------------------------------------------------------- Excel --
    def _load_excel(self):
        self.excel_index, self.processed = build_excel_index(
            self.excel_path, progress_cb=lambda p, m: self.cb(p, m))

    # -------------------------------------------------------- section spec --
    def _sections_for(self, block):
        """Decide which quota blocks this institute must print, and with what data."""
        kind = block["kind"]
        name_l = block["name"].lower()
        excel = _find_institute(self.excel_index, block["name"], kind)

        # ---- which quotas does the workbook actually carry? -----------------
        quotas = set()
        if excel is not None:
            quotas = {e["quota"] for e in excel["entries"]}
        has_state_quotas = bool(quotas & {"HS", "OS"})

        # ---- template branch rows grouped by quota (source of seat counts) --
        tpl_by_quota = {}
        for sec in block["sections"]:
            lab = sec["label"] or ""
            quota = "HS" if "home" in lab.lower() else ("OS" if "other" in lab.lower() else "AI")
            tpl_by_quota.setdefault(quota, []).extend(sec["branches"])

        # ---- which quotas do we print? -------------------------------------
        if has_state_quotas:
            combined = any(k in name_l for k in COMBINED_QUOTA_KEYWORDS)
            want = ["HS", "OS"] if combined else ["OS"]
        else:
            want = ["AI"]

        if excel is None:
            # No Excel data (GFTIs, or an institute missing from the workbook):
            # keep whatever the template already has, one block per template quota.
            sections = []
            for sec in block["sections"]:
                label = sec["label"]
                lab_l = (label or "").lower()
                sec_quota = "HS" if "home" in lab_l else ("OS" if "other" in lab_l else "AI")
                if sec_quota not in want:
                    continue
                rows = []
                for b in sec["branches"]:
                    rows.append({"label": _dd_label(b["label"], b.get("period"),
                                                    b.get("name")),
                                 "ranks": b["ranks"],
                                 "name": b["name"], "from_tpl": True})
                if rows:
                    sections.append({"label": label, "rows": rows})
            return sections, None

        # ---- group Excel entries by quota ----------------------------------
        by_quota = {}
        for e in excel["entries"]:
            by_quota.setdefault(e["quota"], []).append(e)

        sections = []
        for quota in want:
            entries = by_quota.get(quota, [])
            if has_state_quotas:
                label = ("50% Home State Quota" if quota == "HS"
                         else "50% Other State Quota")
            else:
                label = ""          # IIT / IIIT: no divider row (same as template)

            tpl_rows = (tpl_by_quota.get(quota) or tpl_by_quota.get("AI")
                        or next(iter(tpl_by_quota.values()), []))

            rows = []
            if not entries:
                # workbook carries no row for this quota -> template data only
                for b in tpl_rows:
                    rows.append({"label": _dd_label(b["label"], b.get("period"),
                                                    b.get("name")),
                                 "ranks": b["ranks"],
                                 "name": b["name"], "from_tpl": True})
                if rows:
                    sections.append({"label": label, "rows": rows})
                continue

            pairs, unmatched, leftover = assign_branches(tpl_rows, entries)
            matched_to_tpl = {i: j for j, i in pairs.items()}

            # 1. template order first -> keeps the book's branch sequence + seats
            for i, b in enumerate(tpl_rows):
                j = matched_to_tpl.get(i)
                if j is None:
                    self.report.setdefault("dropped_branches", []).append(
                        f"{block['name']} :: {b['name'].replace(chr(10), ' ')}")
                    continue
                hit = entries[j]
                label_txt = b["label"]          # preserve "Name [seats/period]"
                if hit["period"] and "/" not in label_txt:
                    label_txt = f"{b['name']} [{hit['period']}]"
                label_txt = _dd_label(label_txt,
                                      hit["period"] or b.get("period"),
                                      b["name"], hit.get("program"))
                rows.append({"label": label_txt, "ranks": hit["ranks"],
                             "name": b["name"], "from_tpl": True})

            # 2. brand new Excel programmes (seat count unknown -> no bracket)
            for j in leftover:
                e = entries[j]
                nm = e["short"] or e["program"]
                rows.append({"label": _dd_label(nm, e.get("period"),
                                                e.get("program")),
                             "ranks": e["ranks"],
                             "name": nm, "from_tpl": False})
                self.report["new_branches"].append(
                    f"{block['name']} :: {nm}")

            if rows:
                sections.append({"label": label, "rows": rows})

        return sections, excel

    # ------------------------------------------------------------ assemble --
    def _build_table(self, doc, block, sections, scale):
        """Create one fresh <w:tbl> for a single institute and return its element."""
        if block["data_tpl"] is None:
            return None

        # header rows -------------------------------------------------------
        head_rows = []
        if block["cat_row"] is not None:
            head_rows.append(copy.deepcopy(block["cat_row"]))
        if block["oc_row"] is not None:
            head_rows.append(copy.deepcopy(block["oc_row"]))

        def divider_for(label):
            if not label:
                return None
            for key, tr in block["div_rows"].items():
                if key.strip() == label.strip().lower():
                    return copy.deepcopy(tr)
            return _make_quota_row(block["data_tpl"], label)

        rows_xml = [copy.deepcopy(r) for r in block["info_rows"]]

        # when the template block only has one divider but we need two (or the
        # other way round) the divider helper keeps everything consistent.
        div_template_label = next(iter(block["div_rows"]), None)

        for sec in sections:
            dv = divider_for(sec["label"])
            if dv is not None:
                rows_xml.append(dv)
            elif sec["label"]:
                rows_xml.append(_make_quota_row(block["data_tpl"], sec["label"]))
            for row in sec["rows"]:
                ranks = row["ranks"]
                texts = [row["label"]]
                for tok in CAT_ORDER:
                    pair = ranks.get(tok)
                    op = cl = None
                    if pair:
                        op, cl = pair
                    texts.append(str(int(op)) if op is not None else MISSING)
                    texts.append(str(int(cl)) if cl is not None else MISSING)
                rows_xml.append(_clone_row(block["data_tpl"], texts))
                self.report["rows"] += 1

        if not rows_xml:
            return None

        # rebuild head rows (they were created before we knew we had data) ---
        final = rows_xml[:3] + head_rows + rows_xml[3:]

        new_tbl = copy.deepcopy(block["tbl_el"])
        for tr in list(new_tbl.findall(qn("w:tr"))):
            new_tbl.remove(tr)
        for tr in final:
            _force_cant_split(tr)
            _scale_row_font(tr, scale)
            new_tbl.append(tr)
        return new_tbl

    # ---------------------------------------------------------- page break --
    @staticmethod
    def _page_break_paragraph(doc):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = 0
        p.paragraph_format.space_after = 0
        pPr = p._p.get_or_add_pPr()
        for tag, val in (("w:spacing", {"w:before": "0", "w:after": "0"}),):
            el = OxmlElement(tag)
            for k, v in val.items():
                el.set(qn(k), v)
            pPr.append(el)
        r = p.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        r._r.append(br)
        return p._p

    # ----------------------------------------------------------------- run --
    def run(self):
        self.cb(35, "Loading Word template ...")
        doc = docx.Document(self.template_path)

        self.cb(42, "Scanning cut-off tables ...")
        blocks = scan_template(doc)

        self._load_excel()

        body = doc.element.body
        # elements we must remove at the end
        original_tables = []
        stray_paragraphs = []
        for tbl in doc.tables:
            if _classify_table(tbl._tbl):
                original_tables.append(tbl._tbl)
                sib = tbl._tbl.getprevious()
                while sib is not None and sib.tag == qn("w:p"):
                    prev = sib.getprevious()
                    text = "".join(t.text or "" for t in sib.iter(qn("w:t")))
                    if text.strip():
                        break
                    stray_paragraphs.append(sib)
                    sib = prev

        usable_in = PAGE_HEIGHT_IN - MARGIN_TOP_IN - MARGIN_BOTTOM_IN
        target_pt = usable_in * 72 * FIT_SAFETY

        total = len(blocks)
        for i, block in enumerate(blocks):
            pct = 45 + int(45 * (i + 1) / max(1, total))
            self.cb(pct, f"Rebuilding {i + 1}/{total}: {block['name'][:58]}")

            sections, excel = self._sections_for(block)
            if excel is not None:
                self.report["matched"] += 1
            else:
                if block["kind"] != "GFTI":
                    self.report["unmatched_institutes"].append(block["name"])
            if not sections:
                continue

            scale = 1.0
            while True:
                tbl = self._build_table(doc, block, sections, scale)
                if tbl is None:
                    break
                if self._fits(tbl, target_pt) or scale <= MIN_SCALE:
                    break
                scale = round(max(MIN_SCALE, scale - 0.08), 3)
                self.report["shrink_log"].append(
                    f"{block['name'][:40]} -> scale {scale}")
            if tbl is None:
                continue

            brk = self._page_break_paragraph(doc)
            # move the freshly appended elements in front of the original table
            anchor = block["tbl_el"]
            for el in (brk, tbl):
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                anchor.addprevious(el)
            self.report["institutes"] += 1

        self.cb(92, "Removing old tables ...")
        for tr in stray_paragraphs:
            parent = tr.getparent()
            if parent is not None:
                parent.remove(tr)
        for tbl in original_tables:
            parent = tbl.getparent()
            if parent is not None:
                parent.remove(tbl)

        self.cb(96, "Saving document ...")
        out_dir = os.path.dirname(os.path.abspath(self.output_path))
        if out_dir and not os.path.isdir(out_dir):
            raise RuntimeError("Output folder must not be created: " + out_dir)
        doc.save(self.output_path)
        self.cb(100, "Word file written")
        return self.report

    # ----------------------------------------------------------- fit check --
    @staticmethod
    def _fits(tbl, target_pt):
        rows = tbl.findall(qn("w:tr"))
        if not rows:
            return True
        font = _run_font_pt(rows[min(3, len(rows) - 1)])
        pad = _cell_pad_pt(rows[min(3, len(rows) - 1)])
        space = _para_space_pt(rows[min(3, len(rows) - 1)])
        branch_w = 2.49
        cpl = _branch_col_chars(font, branch_w)
        total = 0.0
        for tr in rows:
            total += _row_height_pt(tr, font, pad, space, cpl)
        return total <= target_pt


# --------------------------------------------------------------------------- #
#  Convenience entry point
# --------------------------------------------------------------------------- #

def update_book(excel_path, template_path, output_path, progress_cb=None):
    upd = BookUpdater(excel_path, template_path, output_path, progress_cb)
    return upd.run()
