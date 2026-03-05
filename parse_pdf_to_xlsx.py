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
import urllib.request
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
# County & FIPS enrichment
# ─────────────────────────────────────────────────────────────────────────────

FCC_FIPS_URL = "https://transition.fcc.gov/oet/info/maps/census/fips/fips.txt"

# Minimum number of entries expected after successfully parsing the FCC FIPS
# file (~3 200 county entries in the real file; 200 guards against truncation).
_MIN_EXPECTED_FIPS_ENTRIES = 200

# Maps the 2-digit state FIPS prefix to the state abbreviation.
# Used to determine state from a 5-digit county FIPS code while parsing the
# FCC file without tracking section headers.
_STATE_FIPS_TO_ABBREV = {
    "01": "AL", "02": "AK", "04": "AZ", "05": "AR", "06": "CA",
    "08": "CO", "09": "CT", "10": "DE", "11": "DC", "12": "FL",
    "13": "GA", "15": "HI", "16": "ID", "17": "IL", "18": "IN",
    "19": "IA", "20": "KS", "21": "KY", "22": "LA", "23": "ME",
    "24": "MD", "25": "MA", "26": "MI", "27": "MN", "28": "MS",
    "29": "MO", "30": "MT", "31": "NE", "32": "NV", "33": "NH",
    "34": "NJ", "35": "NM", "36": "NY", "37": "NC", "38": "ND",
    "39": "OH", "40": "OK", "41": "OR", "42": "PA", "44": "RI",
    "45": "SC", "46": "SD", "47": "TN", "48": "TX", "49": "UT",
    "50": "VT", "51": "VA", "53": "WA", "54": "WV", "55": "WI",
    "56": "WY",
}

# County lookup keyed by (city_lower, STATE_ABBREV).
# Sourced from public web data for all NRC nuclear plant locations, using
# Plant Name + Location as identifiers.
_CITY_COUNTY_MAP = {
    ("london",           "AR"): "Pope County",
    ("shippingport",     "PA"): "Beaver County",
    ("braceville",       "IL"): "Will County",
    ("southport",        "NC"): "Brunswick County",
    ("byron",            "IL"): "Ogle County",
    ("fulton",           "MO"): "Callaway County",
    ("lusby",            "MD"): "Calvert County",
    ("york",             "SC"): "York County",
    ("clinton",          "IL"): "DeWitt County",
    ("glen rose",        "TX"): "Somervell County",
    ("brownville",       "NE"): "Nemaha County",
    ("oak harbor",       "OH"): "Ottawa County",
    ("avila beach",      "CA"): "San Luis Obispo County",
    ("bridgman",         "MI"): "Berrien County",
    ("baxley",           "GA"): "Appling County",
    ("newport",          "MI"): "Monroe County",
    ("port gibson",      "MS"): "Claiborne County",
    ("hartsville",       "SC"): "Darlington County",
    ("hancocks bridge",  "NJ"): "Salem County",
    ("scriba",           "NY"): "Oswego County",
    ("columbia",         "AL"): "Houston County",
    ("marseilles",       "IL"): "LaSalle County",
    ("limerick",         "PA"): "Montgomery County",
    ("huntersville",     "NC"): "Mecklenburg County",
    ("waterford",        "CT"): "New London County",
    ("monticello",       "MN"): "Wright County",
    ("seneca",           "SC"): "Oconee County",
    ("wintersburg",      "AZ"): "Maricopa County",
    ("delta",            "PA"): "York County",
    ("perry",            "OH"): "Lake County",
    ("two rivers",       "WI"): "Manitowoc County",
    ("welch",            "MN"): "Goodhue County",
    ("cordova",          "IL"): "Rock Island County",
    ("ontario",          "NY"): "Wayne County",
    ("st. francisville", "LA"): "West Feliciana Parish",
    ("jensen beach",     "FL"): "St. Lucie County",
    ("seabrook",         "NH"): "Rockingham County",
    ("soddy-daisy",      "TN"): "Hamilton County",
    ("new hill",         "NC"): "Wake County",
    ("bay city",         "TX"): "Matagorda County",
    ("surry",            "VA"): "Surry County",
    ("homestead",        "FL"): "Miami-Dade County",
    ("jenkinsville",     "SC"): "Fairfield County",
    ("waynesboro",       "GA"): "Burke County",
    ("killona",          "LA"): "St. Charles Parish",
    ("spring city",      "TN"): "Rhea County",
    # Reference-city fallback used from the Address field when Location is None:
    ("scottsboro",       "AL"): "Jackson County",
}

# Embedded FIPS data used as fallback when the FCC URL is inaccessible.
# Source: US Census Bureau / FCC county FIPS reference (public domain).
# Format: (state_abbrev_lower, county_name_lower) → 5-digit FIPS code.
_EMBEDDED_FIPS = {
    ("al", "limestone county"):        "01083",
    ("al", "houston county"):          "01069",
    ("al", "jackson county"):          "01071",
    ("ar", "pope county"):             "05115",
    ("az", "maricopa county"):         "04013",
    ("ca", "san luis obispo county"):  "06079",
    ("ct", "new london county"):       "09011",
    ("fl", "miami-dade county"):       "12086",
    ("fl", "st. lucie county"):        "12111",
    ("ga", "appling county"):          "13001",
    ("ga", "burke county"):            "13033",
    ("il", "dewitt county"):           "17039",
    ("il", "grundy county"):           "17063",
    ("il", "lasalle county"):          "17099",
    ("il", "ogle county"):             "17141",
    ("il", "rock island county"):      "17161",
    ("il", "will county"):             "17197",
    ("ks", "coffey county"):           "20031",
    ("la", "st. charles parish"):      "22089",
    ("la", "west feliciana parish"):   "22125",
    ("md", "calvert county"):          "24009",
    ("mi", "berrien county"):          "26021",
    ("mi", "monroe county"):           "26115",
    ("mn", "goodhue county"):          "27049",
    ("mn", "wright county"):           "27171",
    ("mo", "callaway county"):         "29027",
    ("ms", "claiborne county"):        "28021",
    ("nc", "brunswick county"):        "37019",
    ("nc", "mecklenburg county"):      "37119",
    ("nc", "wake county"):             "37183",
    ("ne", "nemaha county"):           "31117",
    ("nh", "rockingham county"):       "33015",
    ("nj", "salem county"):            "34033",
    ("ny", "oswego county"):           "36075",
    ("ny", "wayne county"):            "36117",
    ("oh", "lake county"):             "39085",
    ("oh", "ottawa county"):           "39123",
    ("pa", "beaver county"):           "42007",
    ("pa", "luzerne county"):          "42079",
    ("pa", "montgomery county"):       "42091",
    ("pa", "york county"):             "42133",
    ("sc", "cherokee county"):         "45021",
    ("sc", "darlington county"):       "45031",
    ("sc", "fairfield county"):        "45039",
    ("sc", "oconee county"):           "45073",
    ("sc", "york county"):             "45091",
    ("tn", "hamilton county"):         "47065",
    ("tn", "rhea county"):             "47143",
    ("tx", "matagorda county"):        "48321",
    ("tx", "somervell county"):        "48423",
    ("va", "louisa county"):           "51109",
    ("va", "surry county"):            "51181",
    ("wa", "benton county"):           "53005",
    ("wi", "manitowoc county"):        "55071",
}


def _parse_fips_txt(text):
    """Parse FCC FIPS text file into {(state_abbrev_lower, county_lower): fips}.

    Strategy: find any line of the form  "County Name   XXXXX"  (the 5-digit
    FIPS code appears at the end, separated by ≥2 spaces or a tab). The state
    abbreviation is derived from the first two digits of the FIPS code via
    _STATE_FIPS_TO_ABBREV, so no state-header tracking is needed.
    Both the full name ("Pope County") and the bare name ("Pope") are stored.
    County names in the FCC file use ASCII characters (A-Z, space, period,
    hyphen, apostrophe) – Unicode names do not appear in this dataset.
    """
    lookup = {}
    # Match "County Name   XXXXX" with 2+ spaces or tabs before the FIPS code.
    pattern = re.compile(r'^([A-Za-z][A-Za-z\s.\-\']+?)(?:\s{2,}|\t+)(\d{5})\s*$')
    for line in text.splitlines():
        m = pattern.match(line.rstrip())
        if not m:
            continue
        county_raw = m.group(1).strip()
        fips_code  = m.group(2)
        # Skip state-level entries (county portion "000")
        if fips_code[2:] == "000":
            continue
        state_abbrev = _STATE_FIPS_TO_ABBREV.get(fips_code[:2])
        if not state_abbrev:
            continue
        sa  = state_abbrev.lower()
        key = county_raw.lower()
        lookup[(sa, key)] = fips_code
        # Also store without trailing geographic-type suffix for fuzzy matching.
        # "census area" is matched as a complete phrase to avoid false matches
        # on county names that merely end in "area".
        stripped = re.sub(
            r'\s+(county|parish|borough|municipality|census area)\s*$',
            '', key, flags=re.IGNORECASE).strip()
        if stripped != key:
            lookup.setdefault((sa, stripped), fips_code)
    return lookup


def _load_fips_lookup():
    """Download and parse the FCC FIPS file; fall back to embedded data."""
    try:
        req = urllib.request.Request(
            FCC_FIPS_URL,
            headers={"User-Agent": "NRC-PDF-parser/1.0 (python urllib)"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read()
        # Try UTF-8 first; fall back to Latin-1 (the file is plain ASCII but
        # some mirrors may add a BOM or use Windows-1252 for special characters).
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("latin-1")
        table = _parse_fips_txt(text)
        if len(table) >= _MIN_EXPECTED_FIPS_ENTRIES:
            print(f"  FIPS table loaded from FCC ({len(table)} entries).")
            return table
        print("  FCC FIPS file parsed but too few entries; using embedded data.")
    except Exception as exc:
        print(f"  FCC FIPS URL unavailable ({exc}); using embedded data.")
    return dict(_EMBEDDED_FIPS)


def _county_from_location(location):
    """Extract (county_name, state_abbrev) when the county is embedded in the
    Location string.  Returns (None, state) when only city+state is present.

    Handles patterns such as:
      "Limestone County, AL"                      → county at start
      "Hanford Reservation in Benton County, WA"  → "in X County"
      "Morris (Grundy County), IL"                → county in parens
      "Burlington (Coffey County), KS"            → county in parens
    """
    if not location:
        return None, None
    state_m = re.search(r',\s*([A-Z]{2})\s*$', location)
    state   = state_m.group(1) if state_m else None

    # "(Name County)" or "(Name Parish)" anywhere in string
    m = re.search(r'\(([^)]+(?:County|Parish))\)', location, re.IGNORECASE)
    if m and state:
        return m.group(1).strip(), state

    # "in X County, ST" or "in X Parish, ST"
    m = re.search(
        r'\bin\s+([A-Za-z][A-Za-z\s]+(?:County|Parish))\s*,',
        location, re.IGNORECASE)
    if m and state:
        return m.group(1).strip(), state

    # Starts with "X County, ST" or "X Parish, ST"
    m = re.match(r'^([A-Za-z][A-Za-z\s]+(?:County|Parish))\s*,',
                 location, re.IGNORECASE)
    if m and state:
        return m.group(1).strip(), state

    return None, state


def _lookup_city_county(location):
    """Return (county_name, state_abbrev) by looking up city+state in
    _CITY_COUNTY_MAP.  Returns (None, None) when not found."""
    if not location:
        return None, None
    m = re.match(r'^([^,(]+),\s*([A-Z]{2})', location)
    if not m:
        return None, None
    city  = m.group(1).strip().lower()
    state = m.group(2)
    county = _CITY_COUNTY_MAP.get((city, state))
    return (county, state) if county else (None, None)


def _fips_for(state_abbrev, county_name, fips_table):
    """Return the 5-digit FIPS code string, or '' if not found."""
    sa   = state_abbrev.lower()
    base = county_name.lower().strip()
    for key in (base,
                re.sub(r'\s+(county|parish|borough|municipality|census area)\s*$',
                       '', base, flags=re.IGNORECASE).strip()):
        code = fips_table.get((sa, key))
        if code:
            return code
    return ""


def get_county_fips(location, addr, fips_table):
    """Return (county_name, fips_code) for a plant record.

    1. Try to extract county directly from Location (embedded patterns).
    2. Fall back to city-based lookup in _CITY_COUNTY_MAP.
    3. Last resort: extract the reference city from Address (e.g. Bellefonte
       which has no Location but whose Address names a nearby town).
    """
    county, state = _county_from_location(location)

    if not county:
        county, state = _lookup_city_county(location)

    if not county and addr:
        # Address format: "(N miles DIR of City, ST)"
        am = re.search(r'of\s+([A-Za-z][A-Za-z.\s]+),\s*([A-Z]{2})', addr)
        if am:
            ref_city  = am.group(1).strip().lower()
            ref_state = am.group(2)
            county    = _CITY_COUNTY_MAP.get((ref_city, ref_state))
            state     = ref_state

    if not county or not state:
        return "", ""

    return county, _fips_for(state, county, fips_table)


def enrich_with_county_fips(records, fips_table):
    """Add 'County' and 'FIPS Code' fields to every record dict in-place."""
    for rec in records:
        county, fips = get_county_fips(
            rec.get("Location"), rec.get("Address"), fips_table)
        rec["County"]    = county
        rec["FIPS Code"] = fips


# ─────────────────────────────────────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────────────────────────────────────

OP_COLUMNS = [
    "Plant Name", "NRC Region", "Con Type", "Licensed MWt",
    "Licensee", "NSSS", "MWe",
    "Location", "Address", "County", "FIPS Code",
    "Architect Engineer", "Constructor", "License Number",
    "Docket Number",
    "CP Issued", "OL Issued", "Comm. Op",
    "LR Issued", "SR Issued", "Exp. Date",
    "NRC Web Page",
]

UC_COLUMNS = [
    "Plant Name", "NRC Region", "Con Type", "Licensed MWt",
    "Licensee", "NSSS", "MWe",
    "Location", "Address", "County", "FIPS Code",
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
    "COL Issued","Docket Number","License Number","FIPS Code",
}
_WIDTHS = {
    "Plant Name":40,"Licensee":40,"Location":22,"Address":34,
    "County":22,"FIPS Code":10,
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

    print("Loading FIPS lookup …")
    fips_table = _load_fips_lookup()
    enrich_with_county_fips(op_records, fips_table)
    enrich_with_county_fips(uc_records, fips_table)

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
