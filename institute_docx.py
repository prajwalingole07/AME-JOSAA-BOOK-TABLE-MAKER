"""
Standalone one-institute cut-off Word builder.

Produces an A4 landscape .docx for a single IIT / NIT / IIIT using ONLY the
parsed Excel ranks (never the old Word file):

    - tables only (no logo / institute info page above the tables)
    - one rank table per quota block, each block on its own page:
        Category-Gender | OPEN-N | OPEN-W | OBC-N | OBC-W | EWS-N | EWS-W |
                          SC-N    | SC-W   | ST-N   | ST-W
        Branch          | Open Close (x10)
        one row per branch; all headers share the same background
    - a table too tall for one page continues on the next page (headers repeat)
    - quota divider rows (50% Home State / Other State) for NITs
    - missing ranks render as -----
"""

import io
import os
import re

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "static", "logo.jpg")

CAT_TOKENS = ["OPEN-N", "OPEN-W", "OBC-N", "OBC-W", "EWS-N", "EWS-W",
              "SC-N", "SC-W", "ST-N", "ST-W"]

MISSING = "-----"
HEAD_BG = "D9D9D9"
DIV_BG = "FFF2CC"
BRANCH_BG = "EDEDED"

_TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
               "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
               "hideMark"]


# --------------------------------------------------------------------------- #
#  tcPr helpers (insert elements in schema order)
# --------------------------------------------------------------------------- #
def _tcpr_add(cell, el):
    tcPr = cell._tc.get_or_add_tcPr()
    tag = el.tag.split("}")[-1]
    pos = _TCPR_ORDER.index(tag) if tag in _TCPR_ORDER else len(_TCPR_ORDER)
    for child in tcPr:
        ctag = child.tag.split("}")[-1]
        cpos = _TCPR_ORDER.index(ctag) if ctag in _TCPR_ORDER else len(_TCPR_ORDER)
        if cpos > pos:
            child.addprevious(el)
            return
    tcPr.append(el)


def _shade(cell, color):
    _tcpr_add(cell, parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))


def _borders(cell, sz="4", color="808080"):
    _tcpr_add(cell, parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'))


def _margins(cell, top=20, bottom=20, left=40, right=40):
    _tcpr_add(cell, parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'))


def _vmerge(cell, restart=False):
    el = OxmlElement("w:vMerge")
    if restart:
        el.set(qn("w:val"), "restart")
    _tcpr_add(cell, el)


def _run(p, text, bold=False, size=8, color=None, font="Arial"):
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = font
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor(*color)
    return r


def _cell_text(cell, text, bold=False, size=8, align=WD_ALIGN_PARAGRAPH.CENTER,
               color=None, bg=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    _run(p, text, bold=bold, size=size, color=color)
    if bg:
        _shade(cell, bg)
    _borders(cell)
    _margins(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _cant_split(row):
    row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))


def _repeat_header(row):
    row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))


def safe_filename(name):
    s = re.sub(r'[\\/:*?"<>|]+', "", str(name or "")).strip()
    s = re.sub(r"\s+", " ", s)
    return (s.rstrip(". ") or "Institute") + ".docx"


def _rank_str(pair, idx):
    if not pair:
        return MISSING
    val = pair[idx]
    return MISSING if val is None else str(int(val))


_YEARS_RE = re.compile(r"\(\s*(4|5|6|8)\s*Years\b", re.I)
_FIVE_RE = re.compile(r"\(\s*5\s*Years\b", re.I)
_DD_RE = re.compile(r"Dual\s+Degree", re.I)


def _clean_name(name):
    s = re.sub(r"\s+", " ", str(name or "").replace("\xa0", " ")).strip()
    return re.sub(r"\[\s*(HS|OS|AI)\s*\]\s*$", "", s, flags=re.I).strip()


def _strip_years_group(s):
    """Drop the '(N Years, ...)' tail, honouring nested parentheses."""
    m = _YEARS_RE.search(s)
    if not m:
        return re.sub(r"\s*\((Bachelor of Technology|B\.?\s*Tech\.?|Degree)"
                      r"[^()]*\)\s*$", "", s, flags=re.I)
    depth = 0
    for i in range(m.start(), len(s)):
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
            if depth == 0:
                return (s[: m.start()] + " " + s[i + 1:]).strip()
    return s[: m.start()].strip()


def _branch_label(name, is_new=False):
    """(label, is_dual) - '(DD)' only where '(Dual Degree)' is written.

    The course period from '(4 Years, ...)' is printed in the same bracket
    style the book template uses, e.g. 'Chemical Engineering [4 Years]'."""
    s = _clean_name(name)
    prefix = "(NEW) " if is_new else ""
    base = re.sub(r"\s+", " ", _strip_years_group(s)).strip()
    m = _YEARS_RE.search(s)
    period = m.group(1) if m else ""
    is_dual = bool(_DD_RE.search(s))
    label = (prefix + base + (" (DD)" if is_dual else "")).strip()
    if period:
        label = f"{label} [{period} Years]"
    return label, is_dual


# --------------------------------------------------------------------------- #
#  document parts
# --------------------------------------------------------------------------- #
def _init_document():
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Inches(11.69)
    sec.page_height = Inches(8.27)
    sec.top_margin = Inches(0.4)
    sec.bottom_margin = Inches(0.4)
    sec.left_margin = Inches(0.4)
    sec.right_margin = Inches(0.4)
    return doc


_BOTH_QUOTA_NITS = ("nagpur", "jaipur")


def _quota_branches(inst, kind=None):
    """All NITs keep only the Other State quota; Nagpur & Jaipur keep both
    Home State and Other State."""
    branches = inst.get("branches") or []
    k = (kind or "").strip().upper()
    if not k:
        k = ("NIT" if "national institute of technology"
             in str(inst.get("institute_name", "")).lower() else "")
    if k != "NIT":
        return branches
    name = str(inst.get("institute_name", "")).lower()
    if any(t in name for t in _BOTH_QUOTA_NITS):
        return [b for b in branches if b.get("quota") in ("HS", "OS")]
    return [b for b in branches if b.get("quota") == "OS"]


def _quota_sections(branches):
    order, groups = [], {}
    for b in branches:
        q = b.get("quota") or "AI"
        if q not in groups:
            groups[q] = []
            order.append(q)
        groups[q].append(b)
    return [(q, groups[q]) for q in order]


def _has_ranks(ranks, tokens):
    return any(ranks.get(t) for t in tokens)


def _add_page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def _render_rank_table(doc, body, inst_name="", kind="IIT"):
    """One cut-off table for a single quota block, starting on this page,
    with the 3-row institute info box on top matching the official template."""
    from metadata_db import get_institute_metadata
    meta = get_institute_metadata(inst_name)

    ncols = 1 + 2 * len(CAT_TOKENS)
    nrows = 3 + 2 + len(body)
    tbl = doc.add_table(rows=nrows, cols=ncols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    widths = [Inches(2.3)] + [Inches(0.43)] * 20

    # ---- 3-row info box ----------------------------------------------------
    # Row 0: Sr No | 'College Name & Type' | college name | Established In
    _cell_text(tbl.cell(0, 0), "1.", bold=True, size=8.5, bg="FFFFFF")
    _vmerge(tbl.cell(0, 0), restart=True)
    _cell_text(tbl.cell(0, 1).merge(tbl.cell(0, 3)), "College Name & Type",
               bold=True, size=8.5, align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c1 = tbl.cell(0, 4).merge(tbl.cell(0, 16))
    tag = f"({kind})" if kind else ""
    disp = f"{inst_name}. {tag}".strip()
    _cell_text(c1, disp, bold=True, size=8.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c2 = tbl.cell(0, 17).merge(tbl.cell(0, 20))
    estd = meta.get("established") or "—"
    _cell_text(c2, f"Established In {estd}", bold=True, size=8, bg="FFFFFF")

    # Row 1: State, Address, Phone
    _vmerge(tbl.cell(1, 0))
    _cell_text(tbl.cell(1, 1).merge(tbl.cell(1, 3)),
               f"State: {meta.get('state') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c_addr = tbl.cell(1, 4).merge(tbl.cell(1, 16))
    _cell_text(c_addr, f"Address: {meta.get('address') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c_ph = tbl.cell(1, 17).merge(tbl.cell(1, 20))
    _cell_text(c_ph, f"Phone: {meta.get('phone') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")

    # Row 2: Intake, Web, Email, Fees
    _vmerge(tbl.cell(2, 0))
    _cell_text(tbl.cell(2, 1).merge(tbl.cell(2, 3)),
               f"Total Intake: {meta.get('total_intake') or '—'}",
               size=7.5, align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c_web = tbl.cell(2, 4).merge(tbl.cell(2, 9))
    _cell_text(c_web, f"Web: {meta.get('web') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c_em = tbl.cell(2, 10).merge(tbl.cell(2, 16))
    _cell_text(c_em, f"Email: {meta.get('email') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")
    c_fee = tbl.cell(2, 17).merge(tbl.cell(2, 20))
    _cell_text(c_fee, f"Fees per Year: {meta.get('fees') or '—'}", size=7.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg="FFFFFF")

    # ---- header (same grid & colours for every category-gender) ------------
    _cell_text(tbl.cell(3, 0), "Category-Gender", bold=True, size=8.5,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg=HEAD_BG, color=(0, 0, 0))
    _cell_text(tbl.cell(4, 0), "Branch [Seats/Course Period]", bold=True, size=8,
               align=WD_ALIGN_PARAGRAPH.LEFT, bg=HEAD_BG, color=(0, 0, 0))
    for i, tok in enumerate(CAT_TOKENS):
        c1 = tbl.cell(3, 1 + i * 2).merge(tbl.cell(3, 2 + i * 2))
        _cell_text(c1, tok, bold=True, size=8, bg=HEAD_BG, color=(0, 0, 0))
        _cell_text(tbl.cell(4, 1 + i * 2), "Open", bold=True, size=7,
                   bg=HEAD_BG, color=(0, 0, 0))
        _cell_text(tbl.cell(4, 2 + i * 2), "Close", bold=True, size=7,
                   bg=HEAD_BG, color=(0, 0, 0))

    _repeat_header(tbl.rows[0])
    _repeat_header(tbl.rows[1])
    _repeat_header(tbl.rows[2])
    _repeat_header(tbl.rows[3])
    _repeat_header(tbl.rows[4])

    # ---- body: one row per branch -----------------------------------------
    r = 5
    for item in body:
        row = tbl.rows[r]
        _cant_split(row)
        if item[0] == "divider":
            merged = row.cells[0].merge(row.cells[ncols - 1])
            _cell_text(merged, item[1], bold=True, size=9, bg=DIV_BG,
                       color=(128, 90, 0))
            r += 1
            continue

        branch = item[1]
        ranks = branch.get("ranks") or {}
        label, _ = _branch_label(branch.get("display_name")
                                 or branch.get("program_name", ""))
        _cell_text(row.cells[0], label, bold=True, size=8,
                   align=WD_ALIGN_PARAGRAPH.LEFT, bg=BRANCH_BG)
        for i, tok in enumerate(CAT_TOKENS):
            pair = ranks.get(tok)
            _cell_text(row.cells[1 + i * 2], _rank_str(pair, 0), size=7.5)
            _cell_text(row.cells[2 + i * 2], _rank_str(pair, 1), size=7.5)
        r += 1

    for row in tbl.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths):
                cell.width = widths[i]


def _add_rank_table(doc, branches, inst_name="", kind="IIT"):
    """Tables only. Every quota block gets its own table on its own page; a
    table too tall for one page flows onto the next page (headers repeat)."""
    sections = []
    for quota, rows in _quota_sections(branches):
        keep = [b for b in rows
                if _has_ranks(b.get("ranks") or {}, CAT_TOKENS)]
        if keep:
            sections.append((quota, keep))
    if not sections:
        return

    show_divider = any(q in ("HS", "OS") for q, _ in sections) or len(sections) > 1

    for idx, (quota, rows) in enumerate(sections):
        if idx > 0:
            _add_page_break(doc)
        body = []      # ("divider", label) | ("data", branch)
        if show_divider:
            body.append(("divider", {"HS": "50% Home State Quota",
                                     "OS": "50% Other State Quota",
                                     "AI": "All India Quota"}.get(quota, f"{quota} Quota")))
        for b in rows:
            body.append(("data", b))
        _render_rank_table(doc, body, inst_name=inst_name, kind=kind)


# --------------------------------------------------------------------------- #
#  public API
# --------------------------------------------------------------------------- #
def build_institute_docx(inst, kind=None):
    """Return (bytes, filename) for one parser institute entry.

    Produces landscape A4 Word document matching Image 2 with official 3-row
    info box on top, grey header rows, and all quota cut-offs.
    """
    doc = _init_document()
    inst_name = inst.get("institute_name", "Institute")
    k = kind or inst.get("kind", "GFTI")
    _add_rank_table(doc, _quota_branches(inst, k), inst_name=inst_name, kind=k)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), safe_filename(inst_name)
