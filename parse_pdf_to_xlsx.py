"""
Parse NRC ML25051A096.pdf → ML25051A096.xlsx
One Excel column per data field. Columns detected dynamically per page.

Usage:
    python parse_pdf_to_xlsx.py [PDF_PATH [XLSX_PATH]]
Defaults to ML25051A096.pdf / ML25051A096.xlsx in the same directory as
this script when no arguments are supplied.
"""
import re
import sys
from pathlib import Path
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

_HERE = Path(__file__).parent
PDF_PATH  = Path(sys.argv[1]) if len(sys.argv) > 1 else _HERE / "ML25051A096.pdf"
XLSX_PATH = Path(sys.argv[2]) if len(sys.argv) > 2 else _HERE / "ML25051A096.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# Row / word utilities
# ─────────────────────────────────────────────────────────────────────────────

def cluster_rows(words, gap=4.0):
    """Group word dicts into horizontal rows (words within `gap` px of top)."""
    if not words:
        return []
    ws = sorted(words, key=lambda w: w["top"])
    groups, cur = [], [ws[0]]
    for w in ws[1:]:
        if w["top"] - cur[-1]["top"] <= gap:
            cur.append(w)
        else:
            groups.append(cur)
            cur = [w]
    groups.append(cur)
    return [(sum(w["top"] for w in g) / len(g),
             sorted(g, key=lambda w: w["x0"])) for g in groups]


def words_in(row_words, x0, x1):
    """Text of words whose x0 ∈ [x0, x1)."""
    return " ".join(w["text"] for w in row_words if x0 <= w["x0"] < x1).strip()


def clean(text):
    """Remove stray isolated periods and normalise whitespace."""
    # remove isolated period tokens (PDF artifact e.g. "Corp. .")
    text = re.sub(r"\s+\.\s*$", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def url_left_of(row_words, x_cut):
    """Join all words left of x_cut without spaces (for URL fragments)."""
    parts = [w["text"] for w in row_words if w["x0"] < x_cut]
    return "".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# Skip-row detection
# ─────────────────────────────────────────────────────────────────────────────

_SKIP = {
    "APPENDIX","A","Commercial","Nuclear","Power","Reactors",
    "Operating","(continued)","Under","Active","Construction",
    "or","Deferred","Policy",
    "CP","Issued","OL","LR","SR","Comm.","Op","Op.",
    "Plant","Name,","Unit","Number",
    "Licensee","NSSS","MWt","Location","NRC","Web","page","Address",
    "Docket","Region","Constructor","License",
    "Architect","Engineer","MWe","Exp.","Date",
    "Con","Type","Licensed","10","CFR","52.103(g)","Finding",
    "COL","APPENDICES",
}


def is_skip(rw):
    if not rw:
        return True
    texts = [w["text"] for w in rw]
    if re.fullmatch(r"[\d\s|]+", " ".join(texts)):
        return True
    if set(texts) <= _SKIP:
        return True
    if texts[0] in ("F:","A:","*","**","***","Note:","Source:"):
        return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Dynamic column detection from each page's header rows
# ─────────────────────────────────────────────────────────────────────────────
# Column layout (for all pages):
#   left     |  region  |  mid (ConType/NSSS/ArchEng/Constructor)  |  numeric (MWt/MWe/License#)  |  date
#
# The x boundaries are detected from header keywords on each page.
# PDF data text often sits 0–3 px to the LEFT of its column header, so we
# subtract COL_TOLERANCE from every detected boundary to capture it correctly.
COL_TOLERANCE = 3   # px: data can sit this far left of the column-header marker

# How far left of the detected "Region" header the region column starts
# (accommodates unit numbers that sit slightly to the right of xl).
REGION_LEFT_BUFFER = 2    # px: xl = Region_header_x - REGION_LEFT_BUFFER

# Extra tolerance applied when searching for a region token to the left of
# the detected xl boundary (guards against sub-pixel shifts across pages).
REGION_SEARCH_BUFFER = 10  # px: accept region tokens in [xl - REGION_SEARCH_BUFFER, xm)

# Approximate pixel width of the widest region token ("III" ~ 12 px).
# Used to skip past the region token when extracting Con Type.
REGION_TOKEN_WIDTH = 15    # px

# Maximum vertical gap (px) between rows that belong to the same plant entry.
MAX_ROW_GAP = 32           # px for Operating Reactors
MAX_ROW_GAP_UC = 35        # px for Under-Construction entries (slightly taller blocks)


def detect_columns(rows):
    """
    Scan header rows of a page to determine x-column boundaries.
    Returns dict: xl (left/region split), xm (region/mid split),
                  xn (mid/numeric split), xd (numeric/date split).
    """
    # Sensible defaults (work for pages whose header can't be fully scanned)
    xl, xm, xn, xd = 222, 270, 365, 420

    for y, rw in rows:
        for w in rw:
            t = w["text"]
            x = w["x0"]
            if t == "Region":
                xl = x - REGION_LEFT_BUFFER    # left column ends just before "Region"
            if t in ("Architect", "Constructor") and x > 200:
                xm = x - COL_TOLERANCE  # mid column starts just before "Architect"
            if t == "Licensed":
                xn = x - COL_TOLERANCE  # numeric column starts just before "Licensed"
            if t in ("CP", "COL"):
                xd = x - COL_TOLERANCE  # date column starts just before "CP"
        if not is_skip(rw):
            break                   # stop at first data row

    return dict(xl=xl, xm=xm, xn=xn, xd=xd)


# ─────────────────────────────────────────────────────────────────────────────
# Detect the NRC-Region word on a plant-name row
# Returns (region_x, region_text) or (None, None)
# ─────────────────────────────────────────────────────────────────────────────

def find_region(row_words, xl, xm):
    """
    Return the (x0, text) of the I/II/III/IV token that falls in [xl, xm).
    This excludes Roman numerals that are part of the plant name (e.g. "Lee III").
    """
    for w in row_words:
        if w["text"] in ("I","II","III","IV") and xl - REGION_SEARCH_BUFFER <= w["x0"] < xm:
            return w["x0"], w["text"]
    return None, None


def is_plant_row(row_words, xl, xm):
    rx, _ = find_region(row_words, xl, xm)
    return rx is not None


# ─────────────────────────────────────────────────────────────────────────────
# Operating Reactor parser  (pages 1–15)
# ─────────────────────────────────────────────────────────────────────────────
#
# Standard 6-row layout:
#   0  Plant Name | Region | Con Type | MWt | CP Issued
#   1  Licensee   | NSSS   | MWe      | OL Issued
#   2  Location   | ArchEng|          | Comm. Op
#   3  Address    |Constructor|License#| LR Issued
#   4  Docket     |           |        | SR Issued
#   5  URL        |           |        | Exp. Date
#
# 7-row layout (2-line licensee):
#   0  Plant Name | Region | Con Type | MWt | CP Issued
#   1  Licensee1  | NSSS   | MWe      | OL Issued
#   2  Licensee2  | ArchEng|          | Comm. Op
#   3  Location   |Constructor|License#| LR Issued
#   4  Address    |           |        | SR Issued
#   5  Docket     |           |        | Exp. Date
#   6  URL

def parse_op_entry(block, cols):
    xl, xm, xn, xd = cols["xl"], cols["xm"], cols["xn"], cols["xd"]
    n = len(block)
    if n < 3:
        return None

    r = {}

    # Row 0 ─ plant name, region, con type, MWt, CP Issued
    b0 = block[0]
    rx, region = find_region(b0, xl, xm)
    if rx is None:
        return None
    # Plant name = all words to the left of the region word
    r["Plant Name"]   = " ".join(w["text"] for w in b0 if w["x0"] < rx).strip()
    r["NRC Region"]   = region
    # Con Type = tokens between end-of-region and xn
    r["Con Type"]     = words_in(b0, rx + REGION_TOKEN_WIDTH, xn)   # skip past region token
    r["Licensed MWt"] = words_in(b0, xn, xd)
    r["CP Issued"]    = words_in(b0, xd, 9999)

    # Row 1 ─ licensee, NSSS, MWe, OL Issued
    b1 = block[1]
    lic1   = words_in(b1, 0,  xm)
    nsss   = words_in(b1, xm, xn)
    mwe    = words_in(b1, xn, xd)
    ol_iss = words_in(b1, xd, 9999)

    # Detect 2-line licensee
    two_line = lic1.rstrip().endswith("/")
    if not two_line and n >= 7:
        b2l = words_in(block[2], 0, xm)
        if b2l and not b2l.startswith("(") \
                and not re.match(r".+,\s+[A-Z]{2}$", b2l.strip()):
            two_line = True

    if two_line:
        b2 = block[2]
        lic2 = words_in(b2, 0,  xm)
        ae   = words_in(b2, xm, xn)
        comm = words_in(b2, xd, 9999)
        r["Licensee"]           = clean((lic1.rstrip("/").strip() + " / " + lic2).strip(" /"))
        r["NSSS"]               = nsss
        r["MWe"]                = mwe
        r["OL Issued"]          = ol_iss
        r["Architect Engineer"] = ae
        r["Comm. Op"]           = comm
        off = 1
    else:
        r["Licensee"]           = clean(lic1)
        r["NSSS"]               = nsss
        r["MWe"]                = mwe
        r["OL Issued"]          = ol_iss
        r["Architect Engineer"] = ""
        r["Comm. Op"]           = ""
        off = 0

    # Location row
    ri = 2 + off
    if ri < n:
        bl = block[ri]
        r["Location"] = words_in(bl, 0, xm)
        if not two_line:
            r["Architect Engineer"] = words_in(bl, xm, xn)
            r["Comm. Op"]           = words_in(bl, xd, 9999)
        else:
            r["Constructor"]    = words_in(bl, xm, xn)
            r["License Number"] = words_in(bl, xn, xd)
            r["LR Issued"]      = words_in(bl, xd, 9999)

    # Address row
    ri = 3 + off
    if ri < n:
        ba = block[ri]
        r["Address"] = words_in(ba, 0, xm)
        if not two_line:
            r["Constructor"]    = words_in(ba, xm, xn)
            r["License Number"] = words_in(ba, xn, xd)
            r["LR Issued"]      = words_in(ba, xd, 9999)
        else:
            r["SR Issued"] = words_in(ba, xd, 9999)

    # Docket row
    ri = 4 + off
    if ri < n:
        bd = block[ri]
        r["Docket Number"] = words_in(bd, 0, xm)
        if not two_line:
            r["SR Issued"]  = words_in(bd, xd, 9999)
        else:
            r["Exp. Date"]  = words_in(bd, xd, 9999)

    # URL row
    ri = 5 + off
    if ri < n:
        bu = block[ri]
        r["NRC Web Page"] = url_left_of(bu, xd)
        if not two_line:
            r["Exp. Date"] = words_in(bu, xd, 9999)

    return r


def parse_op_pages(pages):
    records = []
    for page in pages:
        words = page.extract_words()
        rows  = cluster_rows(words)
        cols  = detect_columns(rows)

        i = 0
        while i < len(rows):
            y, rw = rows[i]
            if is_skip(rw) or not is_plant_row(rw, cols["xl"], cols["xm"]):
                i += 1
                continue

            block = [rw]
            j = i + 1
            while j < len(rows):
                ny, nrw = rows[j]
                if is_skip(nrw):
                    j += 1
                    continue
                if is_plant_row(nrw, cols["xl"], cols["xm"]):
                    break
                if ny - rows[j - 1][0] > MAX_ROW_GAP:
                    break
                block.append(nrw)
                j += 1

            rec = parse_op_entry(block, cols)
            if rec and rec.get("Plant Name"):
                records.append(rec)
            i = j

    return records


# ─────────────────────────────────────────────────────────────────────────────
# Under-Construction / Deferred parser  (pages 16–17)
# ─────────────────────────────────────────────────────────────────────────────
#
# 4–5-row layout:
#   0  Plant Name | Region | Con Type | MWt | CP Issued
#   1  Licensee   | NSSS   | License# | COL Issued
#   2  (address or location) | Arch/Constructor
#   3  Docket | Constructor
#   4  URL

def parse_uc_entry(block, cols):
    xl, xm, xn, xd = cols["xl"], cols["xm"], cols["xn"], cols["xd"]
    n = len(block)
    r = {}
    if n < 2:
        return None

    b0 = block[0]
    rx, region = find_region(b0, xl, xm)
    if rx is None:
        return None

    r["Plant Name"]   = " ".join(w["text"] for w in b0 if w["x0"] < rx).strip()
    r["NRC Region"]   = region
    r["Con Type"]     = words_in(b0, rx + REGION_TOKEN_WIDTH, xn)
    r["Licensed MWt"] = words_in(b0, xn, xd)
    r["CP Issued"]    = words_in(b0, xd, 9999)

    if n > 1:
        b1 = block[1]
        r["Licensee"]       = words_in(b1, 0,  xm)
        r["NSSS"]           = words_in(b1, xm, xn)
        r["License Number"] = words_in(b1, xn, xd)
        r["COL Issued"]     = words_in(b1, xd, 9999)

    for i in range(2, min(n, 7)):
        bi = block[i]
        left = words_in(bi, 0, xm)
        if left.startswith("http"):
            r.setdefault("NRC Web Page", url_left_of(bi, xd))
        elif re.match(r"^0[57]", left):
            r.setdefault("Docket Number", left)
            r.setdefault("Constructor", words_in(bi, xm, xn))
        elif left.startswith("("):
            r.setdefault("Address", left)
            r.setdefault("Constructor", words_in(bi, xm, xn))
        else:
            r.setdefault("Location", left)
            r.setdefault("Architect Engineer", words_in(bi, xm, xn))
            r.setdefault("MWe", words_in(bi, xn, xd))

    return r


def parse_uc_pages(pages):
    records = []
    for page in pages:
        words = page.extract_words()
        rows  = cluster_rows(words)
        cols  = detect_columns(rows)

        i = 0
        while i < len(rows):
            y, rw = rows[i]
            if is_skip(rw) or not is_plant_row(rw, cols["xl"], cols["xm"]):
                i += 1
                continue

            block = [rw]
            j = i + 1
            while j < len(rows):
                ny, nrw = rows[j]
                if is_skip(nrw):
                    j += 1
                    continue
                if is_plant_row(nrw, cols["xl"], cols["xm"]):
                    break
                if ny - rows[j - 1][0] > MAX_ROW_GAP_UC:
                    break
                block.append(nrw)
                j += 1

            rec = parse_uc_entry(block, cols)
            if rec and rec.get("Plant Name"):
                records.append(rec)
            i = j

    return records


# ─────────────────────────────────────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────────────────────────────────────

OP_COLUMNS = [
    "Plant Name", "NRC Region", "Con Type", "Licensed MWt",
    "Licensee", "NSSS", "MWe",
    "Location", "Address",
    "Architect Engineer", "Constructor", "License Number",
    "Docket Number",
    "CP Issued", "OL Issued", "Comm. Op",
    "LR Issued", "SR Issued", "Exp. Date",
    "NRC Web Page",
]

UC_COLUMNS = [
    "Plant Name", "NRC Region", "Con Type", "Licensed MWt",
    "Licensee", "NSSS", "MWe",
    "Location", "Address",
    "Architect Engineer", "Constructor", "License Number",
    "Docket Number",
    "CP Issued", "COL Issued",
    "NRC Web Page",
]

_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
_DAT_FONT = Font(name="Calibri", size=10)
_ALT_FILL = PatternFill("solid", fgColor="DCE6F1")
_A_CTR    = Alignment(horizontal="center", vertical="top", wrap_text=True)
_A_LEFT   = Alignment(horizontal="left",   vertical="top", wrap_text=True)
_thin     = Side(style="thin", color="BFBFBF")
_BORDER   = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

_CTR_COLS = {
    "NRC Region","Con Type","Licensed MWt","MWe",
    "CP Issued","OL Issued","Comm. Op","LR Issued","SR Issued","Exp. Date",
    "COL Issued","Docket Number","License Number",
}
_WIDTHS = {
    "Plant Name":40,"Licensee":40,"Location":22,"Address":34,
    "NRC Web Page":46,"NSSS":12,"Architect Engineer":18,
    "Constructor":14,"License Number":14,"Docket Number":14,
    "NRC Region":10,"Con Type":14,"Licensed MWt":13,"MWe":8,
    "CP Issued":12,"OL Issued":12,"Comm. Op":12,"COL Issued":12,
    "LR Issued":12,"SR Issued":12,"Exp. Date":12,
}


def write_sheet(ws, columns, records, title):
    ws.title = title

    for ci, col in enumerate(columns, 1):
        c = ws.cell(1, ci, col)
        c.fill = _HDR_FILL; c.font = _HDR_FONT
        c.alignment = _A_CTR; c.border = _BORDER
    ws.row_dimensions[1].height = 28

    for ri, rec in enumerate(records, 2):
        bg = _ALT_FILL if ri % 2 == 0 else None
        for ci, col in enumerate(columns, 1):
            c = ws.cell(ri, ci, rec.get(col, ""))
            c.font = _DAT_FONT; c.border = _BORDER
            c.alignment = _A_CTR if col in _CTR_COLS else _A_LEFT
            if bg:
                c.fill = bg

    for ci, col in enumerate(columns, 1):
        ws.column_dimensions[ws.cell(1,ci).column_letter].width = _WIDTHS.get(col, 14)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Parsing PDF …")
    with pdfplumber.open(PDF_PATH) as pdf:
        op_records = parse_op_pages(pdf.pages[:15])
        uc_records = parse_uc_pages(pdf.pages[15:17])

    print(f"  Operating Reactors : {len(op_records)}")
    print(f"  Under Construction : {len(uc_records)}")

    wb = Workbook()
    write_sheet(wb.active, OP_COLUMNS, op_records, "Operating Reactors")
    write_sheet(wb.create_sheet(), UC_COLUMNS, uc_records, "Under Construction")
    wb.save(XLSX_PATH)
    print(f"Saved → {XLSX_PATH}\n")

    # ── Validation ──────────────────────────────────────────────────────────
    issues = 0
    for r in op_records:
        reg = r.get("NRC Region","")
        if reg not in ("I","II","III","IV"):
            print(f"  BAD REGION: {r['Plant Name']!r} → {reg!r}")
            issues += 1
        if not r.get("CP Issued"):
            print(f"  NO CP:   {r['Plant Name']!r}")
            issues += 1

    print(f"\nValidation issues: {issues}")

    print("\nSample (first 10):")
    for r in op_records[:10]:
        print(f"  {r['Plant Name']!r:50s} | {r['NRC Region']:4s} | {r['CP Issued']:12s} | OL={r.get('OL Issued',''):12s} | Exp={r.get('Exp. Date','')}")

    print("\nUnder Construction:")
    for r in uc_records:
        print(f"  {r['Plant Name']!r:52s} | {r.get('NRC Region',''):4s}")
