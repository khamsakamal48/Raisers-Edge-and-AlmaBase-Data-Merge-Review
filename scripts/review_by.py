r"""
Data Review Tool  v6.0
======================
100% LOCAL. Zero network calls. Zero data exposure.

ADDRESS STRATEGY (v5.1+)
-----------------------
Actual file has address split across component columns:
    primary street, primary city, primary state, primary country, primary zip
    permanent street, permanent city, permanent state, permanent country, permanent zip

DB1 has up to 10 address slots, each with components:
    Address 1 Street / City / State / Country / Zip
    ...
    Address 10 Street / City / State / Country / Zip

DB2 has:
    Home Address Line 1, Home Address Line 2, Home Address City,
    Home Address State, Home Address Country, Home Address Zip Code
    Permanent Address, Permanent Address Line 2

Strategy:
  For each person, COMBINE the actual address components into one string:
      "primary street + primary city + primary state + primary country + primary zip"
  Then for DB1, combine EACH slot (1-10) into one string and compare
  actual vs EACH slot. TRUE if ANY slot fuzzy-matches above threshold.
  For DB2, combine the Home/Permanent components and compare.

COLOUR LOGIC (v6.0)
-------------------
  Actual column cell:
    GREEN       — value matched via EXACT match in DB1 OR DB2,
                  OR value is boolean/categorical (true/false/yes/no/alumni/gold/…),
                  OR column has no DB mapping (unmapped → treated as green)
    FUZZY_GREEN — value matched only via FUZZY match (no exact match found)
                  Uses a different, slightly darker shade of green so reviewers
                  can distinguish fuzzy-matched cells at a glance
    RED         — value did not match in either DB
  YELLOW never appears on the Actual column.

  Individual DB lookup column cell:
    TRUE    (GREEN)      — exact match in this DB
    ~FUZZY  (TEAL-GREEN) — fuzzy-only match in this DB
    MISSING (YELLOW)     — this DB has NO value for this person
    FALSE   (RED)        — this DB has a value but it does not match

MATCH LOGIC (v6.0)
------------------
  For ALL field types (text, phone, email, date, number, url, address):
    1. Exact match  — normalised strings are equal
    2. Fuzzy match  — only tried if exact match FAILED and both sides non-blank.
       Phone: suffix match (last 10 / 7 digits) + general ratio
       Email: same-domain local-part comparison + general ratio
       Others: rapidfuzz.fuzz.token_sort_ratio (or SequenceMatcher fallback)
       If score >= FUZZY_THRESHOLD the cell gets FUZZY_GREEN / ~FUZZY colouring.

  Thresholds (lowered to maximise match frequency):
    FUZZY_THRESHOLD        = 0.90  (was 0.80)
    ADDRESS_HIGH_THRESHOLD = 0.90  (was 0.70) → GREEN
    ADDRESS_LOW_THRESHOLD  = 0.70  (was 0.45) → YELLOW

Security:
  - No import of: requests, urllib, http.client, socket, paramiko, boto3, etc.
  - All operations: disk read -> RAM -> disk write only

Usage:
    python review.py ^
      --actual .../input/actual_data.xlsx ^
      --db1    .../input/db1_dump.xlsx ^
      --db2    .../input/db2_dump.xlsx ^
      --actual-db1-id DB1_ID --actual-db2-id DB2_ID ^
      --db1-id DB1-ID --db2-id DB2-ID ^
      --config .../config/fields_config.json ^
      --output .../output/review.xlsx ^
      --addr-high 0.60 --addr-low 0.35
"""

import argparse
import gc
import json
import os
import pickle
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# Module-level helper — strips Excel-illegal control characters from strings.
# Defined here (not inside a function) so every workbook-writing function
# can use it without redefining it locally.
_illegal_sub = ILLEGAL_CHARACTERS_RE.sub
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from rapidfuzz import fuzz as _rf_fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ── VS Code / Windows RAM hint ────────────────────────────────────────────────
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.SetProcessWorkingSetSize(
            -1, 200_000_000, 2_000_000_000)
    except Exception:
        pass

# ── Optional: phonenumbers ────────────────────────────────────────────────────
try:
    import phonenumbers
    HAS_PHONENUMBERS = True
except ImportError:
    HAS_PHONENUMBERS = False
    print("[WARN] phonenumbers not found — pip install phonenumbers")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECURITY GUARD
# ═══════════════════════════════════════════════════════════════════════════════
_BANNED = {
    "requests","urllib3","httpx","aiohttp","http.client",
    "ftplib","smtplib","paramiko","boto3",
    "google.cloud","azure","dropbox","onedrive",
}

def _security_guard():
    hits = set(sys.modules.keys()) & _BANNED
    if hits:
        raise RuntimeError(
            f"SECURITY VIOLATION: Network-capable module(s) detected: {hits}\n"
            "This script must run offline. Aborting."
        )

_security_guard()


# ═══════════════════════════════════════════════════════════════════════════════
#  GLOBAL CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

FUZZY_THRESHOLD        = 0.90   # increased for higher match frequency
MAX_VARIANT_COLS       = 10
ADDRESS_HIGH_THRESHOLD = 0.90   # increased: score >= this -> GREEN
ADDRESS_LOW_THRESHOLD  = 0.70   # increased: score >= this -> YELLOW, else RED

# ── DB1 address slot definitions (Address 1..10) ──────────────────────────────
DB1_ADDR_SLOTS = list(range(1, 11))   # 1 to 10

def db1_slot_components(slot: int) -> dict:
    """Return the DB1 column names for a given address slot (1-10)."""
    return {
        "street":  f"Address {slot} Street",
        "city":    f"Address {slot} City",
        "state":   f"Address {slot} State",
        "country": f"Address {slot} Country",
        "zip":     f"Address {slot} Zip",
    }

# ── DB2 address definitions ───────────────────────────────────────────────────
DB2_HOME_COMPONENTS = [
    "Home Address Line 1",
    "Home Address Line 2",
    "Home Address City",
    "Home Address State",
    "Home Address Country",
    "Home Address Zip Code",
]
DB2_PERMANENT_COMPONENTS = [
    "Permanent Address",
    "Permanent Address Line 2",
]

# ── Actual address component column names ─────────────────────────────────────
ACTUAL_PRIMARY_COMPONENTS   = [
    "Primary_Street", "Primary_City", "Primary_State",
    "Primary_Country", "Primary_Zip",
]
ACTUAL_PERMANENT_COMPONENTS = [
    "Permanent_Street", "Permanent_City", "Permanent_State",
    "Permanent_Country", "Permanent_Zip",
]

# ── Fields that are PART of the address section but NOT fuzzy-matched ─────────
# These are reviewed with plain text / exact matching
ADDR_META_FIELDS = {
    "primary verified", "primary verification source", "primary mail delivery",
    "permanent verified", "permanent verification source", "permanent mail delivery",
    "primary verification", "permanent verification",
}


# ── Output column ordering ────────────────────────────────────────────────────
# Output files must present columns in the SAME order as the Salesforce actual
# input file (not the alphabetical order of the field config). These helpers map
# every output column back to its position in actual_df.columns.

def _input_col_positions(actual_df) -> dict:
    """Map each actual-file column name -> its index in the input file order."""
    return {c: i for i, c in enumerate(actual_df.columns)}

def _addr_group_position(ag, col_pos: dict) -> float:
    """
    Sort position for a combined address group: the index of its FIRST present
    component column in the input file (e.g. primary address sits where
    Primary_Street is). Falls back to end-of-sheet if no component is present.
    """
    comps = _addr_group_components(ag, col_pos)
    idxs = [col_pos[c] for c in comps]
    return min(idxs) if idxs else float("inf")

def _addr_group_components(ag, col_pos: dict) -> list:
    """Component column names for an address group that exist in the actual file."""
    comps = (ACTUAL_PERMANENT_COMPONENTS
             if str(ag.get("label", "")).lower().startswith("permanent")
             else ACTUAL_PRIMARY_COMPONENTS)
    return [c for c in comps if c in col_pos]


# ═══════════════════════════════════════════════════════════════════════════════
#  COLOURS & FONTS
# ═══════════════════════════════════════════════════════════════════════════════

GREEN_FILL         = PatternFill("solid", start_color="C6EFCE")
FUZZY_ACTUAL_FILL  = PatternFill("solid", start_color="A9D18E")
FUZZY_DB_FILL      = PatternFill("solid", start_color="D5E8D4")
YELLOW_FILL        = PatternFill("solid", start_color="FFEB9C")
RED_FILL           = PatternFill("solid", start_color="FFC7CE")
LIGHT_GREEN_FILL   = PatternFill("solid", start_color="E2EFDA")
LIGHT_ORANGE_FILL  = PatternFill("solid", start_color="FCE4D6")
DIFF_FILL          = PatternFill("solid", start_color="FFF2CC")
DB1_HDR_FILL       = PatternFill("solid", start_color="1F4E79")
DB2_HDR_FILL       = PatternFill("solid", start_color="375623")
REV_HDR_FILL       = PatternFill("solid", start_color="7B3F00")
ADDR_HDR_FILL      = PatternFill("solid", start_color="833C00")
DIFF_HDR_FILL      = PatternFill("solid", start_color="4C4C4C")
AUDIT_FILL         = PatternFill("solid", start_color="4A235A")
HEADER_FILL        = PatternFill("solid", start_color="2E4057")
FILLED_ACTUAL_FILL = PatternFill("solid", start_color="BDD7EE")
FILLED_PCT_FILL    = PatternFill("solid", start_color="D9E1F2")

GREEN_FONT   = Font(size=9, bold=True, color="375623")
FUZZY_FONT   = Font(size=9, bold=True, color="2E7D32")
YELLOW_FONT  = Font(size=9, bold=True, color="7D6608")
RED_FONT      = Font(size=9, bold=True, color="9C0006")
NORMAL_FONT   = Font(size=9)
HEADER_FONT   = Font(size=9, bold=True, color="FFFFFF")
DIFF_FONT     = Font(size=9, italic=True, color="595959")
SCORE_FONT    = Font(size=9, bold=True, color="203864")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
WRAP_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
WRAP_C = Alignment(horizontal="center", vertical="center", wrap_text=True)

_ADDR_ACTUAL_STYLE = {
    "HIGH": (GREEN_FILL,       GREEN_FONT),
    "MID":  (YELLOW_FILL,      YELLOW_FONT),
    "LOW":  (RED_FILL,         RED_FONT),
}
_ADDR_CELL_STYLE = {
    "HIGH": (LIGHT_GREEN_FILL, GREEN_FONT),
    "MID":  (LIGHT_ORANGE_FILL,YELLOW_FONT),
    "LOW":  (RED_FILL,         RED_FONT),
}
_STD_ACTUAL_STYLE = {
    "GREEN":       (GREEN_FILL,       GREEN_FONT),
    "FUZZY_GREEN": (FUZZY_ACTUAL_FILL, FUZZY_FONT),
    "YELLOW":      (YELLOW_FILL,      YELLOW_FONT),
    "RED":         (RED_FILL,         RED_FONT),
}
_REVIEW_STYLE = {
    "TRUE":    (GREEN_FILL,    GREEN_FONT),
    "FUZZY":   (FUZZY_DB_FILL, FUZZY_FONT),
    "FALSE":   (RED_FILL,      RED_FONT),
    "MISSING": (YELLOW_FILL,   YELLOW_FONT),
}

_EMPTY_TOKENS = {"nan", "", "none", "-", "n/a", "nil"}

# ── Boolean / categorical values that are always shown GREEN ─────────────────
# These fields carry no DB-lookup meaning (they're flags/labels); force GREEN.
_BOOL_LIKE_VALUES = {
    "true", "false", "yes", "no", "y", "n", "t", "f",
    "alumni", "gold", "silver", "bronze",
    "1", "0", "active", "inactive", "verified", "unverified",
}

def _is_bool_like_series(series: pd.Series) -> pd.Series:
    """Return boolean mask: True where cell value looks like a flag/categorical."""
    return series.astype(str).str.strip().str.lower().isin(_BOOL_LIKE_VALUES)


# ═══════════════════════════════════════════════════════════════════════════════
#  TYPE DETECTION
# ═══════════════════════════════════════════════════════════════════════════════

_PHONE_KEYS  = {"phone","mobile","cell","contact","tel","telephone","ph","mob","landline"}
_EMAIL_KEYS  = {"email","e-mail","mail","emailid","emailaddress"}
_URL_KEYS    = {"url","website","web","site","link","http","www","homepage","portal"}
_DATE_KEYS   = {"date","dob","birth","doj","joining","expiry","expiration","anniversary",
               "from","to","since","until","start","end"}
_NUMBER_KEYS = {"amount","salary","income","balance","score","count","number","no","num",
               "code","pin","pincode","zip","age","year","rate","percent","pct"}
_ADDR_COMPONENT_KEYS = {
    "street","road","lane","line","locality","area","society","building",
    "flat","house","plot","sector","nagar","colony","village","city","town",
    "district","tehsil","taluka","state","province","country","nation",
    "zip","pincode","postal","postcode","addr","address",
}
_ADDR_META_KEYS  = {"verified","verification","source","delivery","mail","status","flag"}
_ADDR_PREFIX_KEYS= {"primary","permanent","correspondence","home","office","comm","perm","curr"}


def detect_norm_type(col_name: str, sample_values: pd.Series = None) -> str:
    """
    Classify a column into a normalisation type.

    Key rules for address sub-fields:
      - Meta words (verified / source / delivery) always → "text"
      - Component words (street / city / state / zip / country) → "address"
      - Primary/permanent ALONE (no component) → "text"
    """
    col_lower = col_name.lower().strip()

    # Check exact match against known meta fields first
    if col_lower in ADDR_META_FIELDS:
        return "text"

    words = set(re.findall(r'[a-z]+', col_lower))

    if words & _ADDR_META_KEYS:   return "text"
    if words & _EMAIL_KEYS:        return "email"
    if words & _URL_KEYS:          return "url"
    if words & _PHONE_KEYS:        return "phone"
    if words & _DATE_KEYS:         return "date"

    # Address component detection (strip prefix words before checking)
    meaningful = words - _ADDR_PREFIX_KEYS
    if meaningful & _ADDR_COMPONENT_KEYS:
        return "address"

    if words & _NUMBER_KEYS:       return "number"

    if sample_values is not None:
        sample = sample_values.dropna().astype(str).head(20)
        if sample.str.match(r'^[\+\d\s\-\(\)]{7,20}$').mean() > 0.6:  return "phone"
        if sample.str.contains(r'@.+\..+').mean() > 0.6:               return "email"
        avg_len = sample.str.len().mean()
        has_dig = sample.str.contains(r'\d').mean()
        has_com = sample.str.contains(r',').mean()
        if avg_len > 20 and has_dig > 0.2 and has_com > 0.2:           return "address"

    return "text"


# ═══════════════════════════════════════════════════════════════════════════════
#  PHONE NORMALIZATION v2
# ═══════════════════════════════════════════════════════════════════════════════

_CC_SORTED = sorted({
    "91","1","44","61","971","966","65","60","49","33",
    "86","81","55","27","92","94","880","977","64","31",
    "32","39","34","7","46","41","45","47","48","420",
}, key=len, reverse=True)


def _strip_cc(digits: str) -> str:
    if not digits: return digits
    if digits.startswith("00"):    digits = digits[2:]
    elif digits.startswith("011"): digits = digits[3:]
    for cc in _CC_SORTED:
        if digits.startswith(cc):
            nat = digits[len(cc):]
            if len(nat) >= 6: return nat
            break
    if digits.startswith("0") and len(digits) > 7:
        return digits[1:]
    return digits


def normalize_phone(val: str) -> str:
    val = str(val).strip()
    if not val or val.lower() in _EMPTY_TOKENS: return ""
    if HAS_PHONENUMBERS:
        try:
            p = phonenumbers.parse(val, "IN")
            if phonenumbers.is_valid_number(p):
                return str(p.national_number)
        except Exception: pass
    return _strip_cc(re.sub(r'[^\d]', '', val))


# ═══════════════════════════════════════════════════════════════════════════════
#  ADDRESS NORMALIZATION & FUZZY MATCHING
# ═══════════════════════════════════════════════════════════════════════════════

_ADDR_SUBS = [
    (r'\bs/o\b',''), (r'\bw/o\b',''), (r'\bd/o\b',''), (r'\bc/o\b',''),
    (r'\bnear\b',''), (r'\bopp(?:osite)?\.?\b',''), (r'\bbehind\b',''),
    (r'\bapartment\b','apt'), (r'\bflat\b','f'), (r'\bfl\b','f'),
    (r'\bhouse\b','h'), (r'\bh\.?\s*no\.?\b','h'),
    (r'\bplot\s*no\.?\b','plot'), (r'\bshop\s*no\.?\b','shop'),
    (r'\bbuilding\b','bldg'), (r'\bbldg\b','bldg'), (r'\bblock\b','blk'),
    (r'\btower\b','twr'), (r'\bwing\b','wng'), (r'\bunit\b','u'),
    (r'\bsociety\b','soc'), (r'\bnagar\b','ngr'), (r'\bcolony\b','col'),
    (r'\bsector\b','sec'), (r'\bphase\b','ph'), (r'\bextension\b','ext'),
    (r'\benclave\b','enc'), (r'\bchowk\b','chk'), (r'\bmarket\b','mkt'),
    (r'\bcomplex\b','cplx'), (r'\bpark\b','prk'), (r'\bheights\b','hts'),
    (r'\bstreet\b','st'), (r'\broad\b','rd'), (r'\bmarg\b','rd'),
    (r'\blane\b','ln'), (r'\bnational\s+highway\b','nh'),
    (r'\bwest\b','w'), (r'\beast\b','e'), (r'\bnorth\b','n'), (r'\bsouth\b','s'),
    (r'\bno\.?\s*',''), (r'\bnum\.?\s*',''), (r'\bnumber\b',''),
    (r'[,\.\-\(\)\[\]\/\\:#@&]',' '), (r'\s{2,}',' '),
]
_ADDR_PATTERNS = [(re.compile(p, re.IGNORECASE), r) for p, r in _ADDR_SUBS]


def normalize_address(val: str) -> str:
    """Normalize a single address string for comparison."""
    val = str(val).strip().lower()
    if not val or val in _EMPTY_TOKENS: return ""
    for pat, repl in _ADDR_PATTERNS:
        val = pat.sub(repl, val)
    return val.strip()


def combine_and_normalize(components: list) -> str:
    """
    Join a list of address component strings into one normalized string.
    Blank components are skipped.

    Example:
        ["Flat 4 Shree Nagar", "Andheri West", "Mumbai", "Maharashtra", "400058"]
        → "f 4 shree ngr andheri w mumbai maharashtra 400058"
    """
    parts = []
    for c in components:
        c = str(c).strip()
        if c and c.lower() not in _EMPTY_TOKENS:
            parts.append(normalize_address(c))
    return " ".join(p for p in parts if p).strip()


def _tok(s: str) -> set:
    return {t for t in s.split() if len(t) >= 2}


# Minimum token count on the smaller address before its full containment in the
# larger one earns containment credit. Stops tiny stubs ("India India") from
# scoring GREEN against any address that happens to mention the same word.
ADDR_CONTAINMENT_MIN_TOKENS = 4

def address_similarity(na: str, nb: str) -> float:
    """
    Token-overlap similarity + 0.25 boost if zip/pincode matches.
    Returns 0.0–1.0. Both blank → 1.0.

    Score is the max of:
      - Jaccard            = |A∩B| / |A∪B|         (symmetric overlap)
      - Containment        = |A∩B| / min(|A|,|B|)  (subset coverage)
    Containment only counts when the smaller address has at least
    ADDR_CONTAINMENT_MIN_TOKENS tokens. This rescues the common case where one
    DB stores only a street line while the actual blob also carries
    city/state/country/zip: the DB address is fully contained in the actual one,
    so it should score high instead of being penalised by the longer union.
    """
    if not na and not nb: return 1.0
    if not na or not nb:  return 0.0
    ta, tb = _tok(na), _tok(nb)
    if not ta and not tb: return 1.0
    if not ta or not tb:  return 0.0
    inter = len(ta & tb)
    jaccard = inter / len(ta | tb)
    score = jaccard
    smaller = min(len(ta), len(tb))
    if smaller >= ADDR_CONTAINMENT_MIN_TOKENS:
        containment = inter / smaller
        score = max(score, containment)
    pa = set(re.findall(r'\b\d{5,6}\b', na))
    pb = set(re.findall(r'\b\d{5,6}\b', nb))
    if pa and pb and (pa & pb):
        score = min(1.0, score + 0.25)
    return round(score, 4)


def address_diff(na: str, nb: str) -> str:
    """
    Show which tokens are present in actual but missing from DB (ACTUAL ONLY►)
    and which are in DB but not actual (◄DB ONLY).
    """
    if not na and not nb: return "both blank"
    if not na: return f"actual blank | DB: {nb[:80]}"
    if not nb: return f"actual: {na[:80]} | DB blank/missing"
    ta, tb = _tok(na), _tok(nb)
    only_a = sorted(ta - tb)
    only_b = sorted(tb - ta)
    parts = []
    if only_a: parts.append(f"ACTUAL ONLY► {' '.join(only_a)}")
    if only_b: parts.append(f"◄DB ONLY: {' '.join(only_b)}")
    return " | ".join(parts) if parts else "tokens match"


def _addr_colour(score: float) -> str:
    if score >= ADDRESS_HIGH_THRESHOLD: return "HIGH"
    if score >= ADDRESS_LOW_THRESHOLD:  return "MID"
    return "LOW"


# ═══════════════════════════════════════════════════════════════════════════════
#  OTHER NORMALIZERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_email(val: str) -> str:
    return re.sub(r'^mailto:', '', str(val).strip().lower())

def normalize_url(val: str) -> str:
    val = str(val).strip().lower()
    val = re.sub(r'^https?://', '', val)
    val = re.sub(r'^www\.', '', val)
    return val.rstrip('/')

def normalize_date(val: str) -> str:
    """
    Normalise any date-like value to dd-mmm-yyyy (e.g. 15-Jan-2024).

    Handles three common problems that arise when reading Excel files with
    dtype=str:

      1. Timestamp suffix  — "2021-03-21 00:00:00" → strip " HH:MM:SS" first.
      2. Excel serial numbers — Excel stores dates as integers (days since
         1899-12-30). Reading with dtype=str produces strings like "44276"
         or "44276.0" that must be converted back to real dates.
      3. yyyy-mm-dd bare strings — convert to target format dd-mmm-yyyy.
    """
    val = str(val).strip()
    if not val or val.lower() in _EMPTY_TOKENS:
        return ""

    # ── Step 1: strip trailing timestamp "HH:MM:SS" or "HH:MM:SS.ffffff" ──────
    val = re.sub(r'\s+\d{1,2}:\d{2}:\d{2}(\.\d+)?\s*$', '', val).strip()

    # ── Step 2: Excel serial number (e.g. "44276" or "44276.0") ──────────────
    # Range 20000–60000 covers roughly year 1955 to 2064 — avoids collisions
    # with plain 5-digit numeric IDs that are too small or too large.
    m_serial = re.match(r'^(\d{4,5})(?:\.0+)?$', val)
    if m_serial:
        serial = int(m_serial.group(1))
        if 20_000 <= serial <= 60_000:          # plausible Excel date range
            try:
                dt = pd.Timestamp('1899-12-30') + pd.Timedelta(days=serial)
                return dt.strftime('%d-%b-%Y')
            except Exception:
                pass

    # ── Step 3: yyyy-mm-dd ────────────────────────────────────────────────────
    if re.match(r'^\d{4}-\d{2}-\d{2}$', val):
        try:
            return pd.to_datetime(val, format='%Y-%m-%d').strftime('%d-%b-%Y')
        except Exception:
            pass

    # ── Step 4: dd-mmm-yyyy (target format — already correct) ────────────────
    try:
        return pd.to_datetime(val, format='%d-%b-%Y', errors='raise').strftime('%d-%b-%Y')
    except Exception:
        pass

    # ── Step 5: generic fallback — let pandas try any format ─────────────────
    try:
        return pd.to_datetime(val, dayfirst=False, errors='raise').strftime('%d-%b-%Y')
    except Exception:
        pass

    return re.sub(r'\s+', ' ', val).lower()

def normalize_number(val: str) -> str:
    val = str(val).strip()
    if not val or val.lower() in _EMPTY_TOKENS: return ""
    val = re.sub(r'[₹$£€,\s]', '', val)
    val = re.sub(r'^[a-zA-Z\.\s]+', '', val)
    val = re.sub(r'\.0+$', '', val)
    return val.strip()

def normalize_text(val: str) -> str:
    """
    Normalise a free-text value for comparison.

    Steps
    -----
    1. Strip leading/trailing whitespace, lowercase.
    2. Collapse all internal whitespace runs to a single space.
    3. Strip trailing punctuation characters  (. , ; : ! ?)
       so that values like "G." and "G", or "Dr." and "Dr",
       or "Inc." and "Inc" compare as equal at Step-1 exact match
       without needing to fall through to fuzzy or combined matching.

    Note: only TRAILING punctuation is removed.  Internal punctuation
    (e.g. "U.S.A.", "Ph.D.", "e.g.") keeps its dots after the collapse
    step — only the final dot is stripped, which is the desired behaviour
    for abbreviations written with or without a closing period.
    """
    val = re.sub(r'\s+', ' ', str(val).strip().lower())
    val = val.rstrip('.,;:!?')
    # Collapse a pure integer-valued float to its integer form so that values
    # read by pandas as floats (e.g. "1976.0") compare equal to the same value
    # stored as an int/string elsewhere (e.g. "1976"). Only applies when the
    # ENTIRE value is digits + a trailing ".0…" — never touches real text.
    val = re.sub(r'^(\d+)\.0+$', r'\1', val)
    return val.strip()

def normalize_value(val: str, norm_type: str) -> str:
    return {
        "phone":   normalize_phone,
        "email":   normalize_email,
        "url":     normalize_url,
        "date":    normalize_date,
        "number":  normalize_number,
        "address": normalize_address,
        "text":    normalize_text,
    }.get(norm_type, normalize_text)(val)

def _norm_via_unique(series: pd.Series, fn) -> pd.Series:
    s = series.astype(str)
    m = {u: fn(u) for u in s.unique()}
    return s.map(m)

def normalize_series(series: pd.Series, norm_type: str) -> pd.Series:
    if norm_type == "email":
        return (series.astype(str).str.strip().str.lower()
                .str.replace(r'^mailto:', '', regex=True))
    if norm_type == "url":
        s = series.astype(str).str.strip().str.lower()
        s = s.str.replace(r'^https?://', '', regex=True)
        s = s.str.replace(r'^www\.', '', regex=True)
        return s.str.rstrip('/')
    if norm_type == "text":
        s = (series.astype(str).str.strip().str.lower()
             .str.replace(r'\s+', ' ', regex=True))
        # Collapse pure integer-valued floats ("1976.0" → "1976") so numeric
        # values stored as floats compare equal to int/string equivalents.
        return s.str.replace(r'^(\d+)\.0+$', r'\1', regex=True)
    if norm_type == "number":
        s = series.astype(str).str.strip()
        empty = s.str.lower().isin(_EMPTY_TOKENS)
        s = (s.str.replace(r'[₹$£€,\s]', '', regex=True)
              .str.replace(r'^[a-zA-Z\.\s]+', '', regex=True)
              .str.replace(r'\.0+$', '', regex=True).str.strip())
        return s.mask(empty, "")
    return _norm_via_unique(series, lambda v: normalize_value(v, norm_type))


# ═══════════════════════════════════════════════════════════════════════════════
#  PHONE-SPECIFIC FUZZY MATCHING
# ═══════════════════════════════════════════════════════════════════════════════

def _phone_fuzzy(a: str, b: str, threshold: float) -> bool:
    """
    Fuzzy match for already-normalised digit strings.
      1. Suffix match on last 10 digits (handles different country-code forms)
      2. Suffix match on last 7 digits  (regional number core)
      3. General ratio fallback
    """
    if not a or not b:
        return False
    # Strip any remaining non-digits just in case
    da, db_ = re.sub(r'\D', '', a), re.sub(r'\D', '', b)
    if not da or not db_:
        return False
    # Suffix matches (most reliable for phone numbers)
    for n in (10, 7):
        if len(da) >= n and len(db_) >= n and da[-n:] == db_[-n:]:
            return True
    # General ratio fallback
    if HAS_RAPIDFUZZ:
        return _rf_fuzz.ratio(da, db_) / 100.0 >= threshold
    return SequenceMatcher(None, da, db_).ratio() >= threshold


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL-SPECIFIC FUZZY MATCHING
# ═══════════════════════════════════════════════════════════════════════════════

def _email_fuzzy(a: str, b: str, threshold: float) -> bool:
    """
    Fuzzy match for already-normalised email strings.
      1. Exact (already handled upstream, but kept as safety)
      2. Same domain → fuzzy compare local parts only
      3. Full string ratio fallback
    """
    if not a or not b:
        return False
    if a == b:
        return True
    if '@' in a and '@' in b:
        al, ad = a.rsplit('@', 1)
        bl, bd = b.rsplit('@', 1)
        if ad == bd:   # same domain: compare only local parts
            ratio = SequenceMatcher(None, al, bl).ratio()
            return ratio >= threshold
    if HAS_RAPIDFUZZ:
        return _rf_fuzz.token_sort_ratio(a, b) / 100.0 >= threshold
    return SequenceMatcher(None, a, b).ratio() >= threshold


# ═══════════════════════════════════════════════════════════════════════════════
#  FUZZY VALUE MATCHING  (field-level, applied after exact match fails)
# ═══════════════════════════════════════════════════════════════════════════════

def _fuzzy_match_series(
        s_actual: pd.Series,
        s_db: pd.Series,
        mask: pd.Series,
        threshold: float = None,
        norm_type: str = "text",
) -> pd.Series:
    """
    Element-wise fuzzy similarity check, evaluated ONLY for rows where mask=True.

    Dispatches to specialised matchers for phone and email, then falls back to
    the general token_sort_ratio / SequenceMatcher for all other types.

    Returns a boolean Series (True = fuzzy match at or above threshold).
    Only called on rows that already FAILED exact match and have non-blank
    values on both sides.
    """
    if threshold is None:
        threshold = FUZZY_THRESHOLD
    result = pd.Series(False, index=s_actual.index)
    mask_idx = s_actual.index[mask]
    if len(mask_idx) == 0:
        return result
    a_vals = s_actual[mask].tolist()
    b_vals = s_db[mask].tolist()
    matched = []
    for a, b in zip(a_vals, b_vals):
        if not a or not b:
            matched.append(False)
            continue
        if norm_type == "phone":
            matched.append(_phone_fuzzy(a, b, threshold))
        elif norm_type == "email":
            matched.append(_email_fuzzy(a, b, threshold))
        else:
            # General: token_sort_ratio handles word-order variations gracefully
            if HAS_RAPIDFUZZ:
                score = _rf_fuzz.token_sort_ratio(a, b) / 100.0
            else:
                score = SequenceMatcher(None, a, b).ratio()
            matched.append(score >= threshold)
    result[mask_idx] = matched
    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  FUZZY COLUMN MATCHING
# ═══════════════════════════════════════════════════════════════════════════════

@lru_cache(maxsize=None)
def normalize_col_name(name: str) -> str:
    n = str(name).lower().strip()
    n = re.sub(r'[\s_\-\.]+', ' ', n)
    n = re.sub(r'(\s*\d+(\.\d+)*)+$', '', n).strip()
    return n

@lru_cache(maxsize=None)
def _seq_ratio(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()

def clear_all_caches():
    """
    Wipe every lru_cache in this module.

    Must be called at the start of main() so that corrected config files
    and updated column mappings are always picked up fresh — stale cached
    column-name similarity scores from a previous run in the same Python
    process (e.g. Jupyter / interactive shell) will NOT carry over.
    """
    normalize_col_name.cache_clear()
    _seq_ratio.cache_clear()
    print("  [cache] All lru_caches cleared — fresh column matching guaranteed.")

def col_similarity(a: str, b: str) -> float:
    return _seq_ratio(normalize_col_name(a), normalize_col_name(b))

def find_matching_columns(target: str, columns: list, threshold: float) -> list:
    t_norm = normalize_col_name(target)
    results = [(c, round(_seq_ratio(t_norm, normalize_col_name(c)), 3)) for c in columns]
    results = [(c, s) for c, s in results if s >= threshold]
    results.sort(key=lambda x: -x[1])
    return [c for c, _ in results]

def group_variants(columns: list, threshold: float) -> dict:
    used, groups = set(), {}
    norm_map = {c: normalize_col_name(c) for c in columns}
    for col in columns:
        if col in used: continue
        norm = norm_map[col]
        groups.setdefault(norm, []).append(col)
        used.add(col)
        cn = norm_map[col]
        for other in columns:
            if other not in used and _seq_ratio(cn, norm_map[other]) >= threshold:
                groups[norm].append(other)
                used.add(other)
    return groups


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════════════════════════════════════

def load_excel(path: str) -> pd.DataFrame:
    print(f"  Loading: {os.path.basename(path)} ...")
    t0 = time.time()
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv", ".txt"):
        sep = "\t" if ext == ".tsv" else ","
        try:
            df = pd.read_csv(path, dtype=str, sep=sep, encoding="utf-8",
                             keep_default_na=False, na_values=[""], low_memory=False)
            eng = f"csv:utf-8"
        except UnicodeDecodeError:
            df = pd.read_csv(path, dtype=str, sep=sep, encoding="latin-1",
                             keep_default_na=False, na_values=[""], low_memory=False)
            eng = f"csv:latin-1"
    else:
        try:
            df = pd.read_excel(path, dtype=str, engine="calamine"); eng = "calamine"
        except Exception:
            df = pd.read_excel(path, dtype=str, engine="openpyxl"); eng = "openpyxl"
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    print(f"  -> {len(df):,} rows x {len(df.columns):,} cols ({time.time()-t0:.1f}s via {eng})")
    return df

def detect_id_column(df: pd.DataFrame, hint: str) -> str:
    matches = find_matching_columns(hint, df.columns.tolist(), threshold=0.65)
    if not matches:
        raise ValueError(f"Cannot find column '{hint}'. First 15: {list(df.columns[:15])}")
    best = matches[0]
    others = matches[1:3]
    print(f"  [ID] '{hint}' -> '{best}'" + (f"  (also: {others})" if others else ""))
    return best


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _is_blank(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin(_EMPTY_TOKENS)

def xlookup(lookup_vals, key_col, val_col):
    m = dict(zip(key_col.astype(str).str.strip(), val_col.astype(str).str.strip()))
    return lookup_vals.astype(str).str.strip().map(m).fillna("")


def _multi_key_lookup(key: str, mapping: dict) -> str:
    """
    Look up a value from mapping for a key that may contain multiple IDs
    in a single cell (e.g. "123,456" or "123|456" or "123;456").

    Strategy:
      1. Try the full key as-is (fast path for normal single-ID cells).
      2. If no non-blank value found, split on common delimiters and try
         each individual ID in order — returning the first non-blank hit.
      3. Return "" if nothing matched.

    This ensures that cells like "ID001,ID002" will check ID001 in the dump,
    then ID002, and return the first column value that is actually populated.
    """
    key = str(key).strip()

    # Fast path: direct match (handles normal single-ID cells with zero overhead)
    direct = mapping.get(key, "")
    if direct and direct.lower() not in _EMPTY_TOKENS:
        return direct

    # Split on comma, semicolon, pipe, or forward-slash (trim whitespace around each)
    parts = [p.strip() for p in re.split(r'[,;|/]+', key) if p.strip()]
    if len(parts) > 1:
        for part in parts:
            val = mapping.get(part, "")
            if val and val.lower() not in _EMPTY_TOKENS:
                return val

    return direct  # return whatever direct lookup gave (may be "")


def _multi_key_lookup_match(key: str, mapping: dict, target_norm: str,
                            norm_type: str) -> str:
    """
    Like _multi_key_lookup, but for cells holding multiple IDs (e.g. "123,456")
    it prefers the candidate whose normalized value EQUALS target_norm.

    A multi-ID actual row corresponds to several merged DB records.  The plain
    "first non-blank" rule can return a value from a different record than the
    one the actual value came from (e.g. batch 1978 vs the first record's 1984),
    producing a false mismatch.  When the actual value is present in ANY linked
    record we return that record's value so the exact match succeeds; otherwise
    we fall back to first-non-blank (unchanged behaviour).
    """
    key = str(key).strip()

    candidates = []
    direct = mapping.get(key, "")
    if direct:
        candidates.append(direct)
    parts = [p.strip() for p in re.split(r'[,;|/]+', key) if p.strip()]
    if len(parts) > 1:
        for p in parts:
            v = mapping.get(p, "")
            if v:
                candidates.append(v)

    if target_norm != "":
        for v in candidates:
            if v.lower() in _EMPTY_TOKENS:
                continue
            if normalize_value(v, norm_type) == target_norm:
                return v

    for v in candidates:
        if v and v.lower() not in _EMPTY_TOKENS:
            return v
    return direct


# ── Stop-words filtered during token-overlap matching ─────────────────────────
# These are so common that matching on them alone produces false positives.
_STOP_WORDS = {
    "the", "and", "of", "in", "at", "to", "for", "on", "by", "with",
    "a", "an", "is", "are", "was", "were", "be", "as",
}

# Regex that catches common encoding artifacts produced when UTF-8 text is
# mis-decoded as Latin-1 (e.g. non-breaking space U+00A0 → "Â ").
# Also removes zero-width characters and other invisible Unicode noise.
_ENCODING_ARTIFACT_RE = re.compile(
    r'[\u00c2\u00c3\u00a0\u00b0\u200b\u200c\u200d\ufeff\ufffd\xa0]+'
)


def _clean_encoding(text: str) -> str:
    """
    Strip common encoding artefacts (Â, non-breaking spaces, mojibake, etc.)
    then collapse any resulting double-spaces.

    Example
    -------
    "Â project engineering"   →  "project engineering"
    "PetroleumÂ "             →  "Petroleum"
    """
    text = _ENCODING_ARTIFACT_RE.sub(' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def _token_overlap_match(db_val: str, actual_val: str,
                         min_len: int = 4,
                         threshold: float = 0.85) -> bool:
    """
    Robust combined/contains matching — three passes in order of cost.

    Designed to handle the full range of real-world problems:
      • Actual field contains DB1 value + semicolon + DB2 value (concatenated)
      • Comma/semicolon-separated skill lists with different ordering
      • Encoding artefacts: "Â project engineering" / "PetroleumÂ "
      • Word-order differences: "New York, USA" vs "USA, New York"
      • Punctuation differences: "Oil & Gas" vs "Oil &Gas" vs "Oil/Gas"
      • Case differences (handled upstream by normalize_text)

    Pass 1 — Exact substring  (O(n), zero overhead)
    ------------------------------------------------
    Checks whether the DB value appears verbatim inside the actual string.
    Catches the common case where actual = "DB1_value; DB2_value" and one of
    them is a literal prefix/suffix.

        db  = "Civil Engineering, Engineering Design"
        act = "Autocad, Civil Engineering, Engineering Design, Epc; ..."
        → "civil engineering, engineering design" in actual → HIT

    Pass 2 — Encoding-cleaned substring  (cheap string clean + O(n) check)
    -----------------------------------------------------------------------
    Strips encoding artefacts (Â, U+00A0, mojibake…) from both sides then
    repeats the substring check.  Catches cases where the DB dump was saved
    with an encoding mismatch and has stray non-ASCII bytes.

        db  = "Â project engineering, PetroleumÂ"
        cleaned_db  = "project engineering, Petroleum"
        cleaned_act = "... project engineering, Petroleum ..."
        → hit after clean

    Pass 3 — Token overlap  (tokenise + set intersection)
    -------------------------------------------------------
    Extracts individual alphanumeric words from both cleaned strings, filters
    out stop-words and tokens shorter than 3 chars, then checks what fraction
    of the DB's meaningful tokens appear anywhere inside the actual value.
    Default threshold = 85 % → allows 1 word to be absent in a 7-word phrase.

        db_tokens  = {autocad, civil, engineering, epc, feed, management,
                      oil, gas, petrochemical, petroleum, project, planning,
                      refinery, operations, structural, analysis}   (16 tokens)
        act_tokens = (superset — actual contains DB1 + DB2 content)
        matched    = 16 / 16 = 100 % ≥ 85 % → HIT

    Why threshold < 100 %?
    ~~~~~~~~~~~~~~~~~~~~~~
    Comma-separated lists often have minor trailing/leading space differences,
    and very long DB values may have one token that the actual abbreviates or
    omits.  85 % gives a one-token grace for lists up to 7 items while still
    being strict enough to prevent random false positives.

    Parameters
    ----------
    db_val     : normalised DB value string (after normalize_series)
    actual_val : normalised actual value string
    min_len    : DB value must be at least this long before we try (prevents
                 trivially short values like "Mr", "01" from matching anywhere)
    threshold  : fraction of DB tokens that must appear in actual (0–1)

    Returns True if any pass finds a match, False otherwise.
    """
    if len(db_val) < min_len:
        return False

    # ── Pass 1: exact substring ───────────────────────────────────────────────
    if db_val in actual_val:
        return True

    # ── Strip encoding artefacts for Passes 2 & 3 ────────────────────────────
    db_clean     = _clean_encoding(db_val)
    actual_clean = _clean_encoding(actual_val)

    # ── Pass 2: encoding-cleaned substring ────────────────────────────────────
    if len(db_clean) >= min_len and db_clean in actual_clean:
        return True

    # ── Pass 3: token overlap ─────────────────────────────────────────────────
    # Use the cleaned strings so artefact characters don't end up as spurious tokens.
    db_tokens = [
        t for t in re.findall(r'[a-z0-9]+', db_clean)
        if len(t) >= 3 and t not in _STOP_WORDS
    ]
    if not db_tokens:
        return False
    actual_tokens = set(re.findall(r'[a-z0-9]+', actual_clean))
    matched = sum(1 for t in db_tokens if t in actual_tokens)
    return (matched / len(db_tokens)) >= threshold

def build_field_plan(actual_df, db1_df, db2_df, fields=None) -> list:
    print("\n[3] Building field plan ...")
    actual_groups = group_variants(actual_df.columns.tolist(), FUZZY_THRESHOLD)
    target_canonicals = (
        [normalize_col_name(f) for f in fields] if fields
        else list(actual_groups.keys())
    )
    plan = []
    for canonical in target_canonicals:
        actual_cols = []
        for norm_key, cols in actual_groups.items():
            if col_similarity(canonical, norm_key) >= FUZZY_THRESHOLD:
                actual_cols.extend(cols)
        if not actual_cols:
            print(f"  [WARN] No actual column for '{canonical}' — skipping")
            continue
        db1_cols = find_matching_columns(
            canonical, db1_df.columns.tolist(), FUZZY_THRESHOLD)[:MAX_VARIANT_COLS]
        db2_cols = find_matching_columns(
            canonical, db2_df.columns.tolist(), FUZZY_THRESHOLD)[:MAX_VARIANT_COLS]
        for actual_col in actual_cols:
            norm_type = detect_norm_type(actual_col, actual_df[actual_col])
            plan.append({
                "canonical":    canonical,
                "actual_col":   actual_col,
                "db1_cols":     db1_cols,
                "db2_cols":     db2_cols,
                "norm_type":    norm_type,
                "is_addr_group":False,
            })
    print(f"  -> {len(plan)} actual columns to review")
    for p in plan:
        print(f"     {p['actual_col']:<42} type={p['norm_type']:<8} "
              f"DB1:{len(p['db1_cols'])}  DB2:{len(p['db2_cols'])}")
    return plan


# ═══════════════════════════════════════════════════════════════════════════════
#  ADDRESS GROUP COMPUTATION
#  Handles the special case where actual has split address components
#  (primary street / city / state / country / zip) that map to
#  multiple DB slots (Address 1..10 in DB1, Home/Permanent in DB2).
# ═══════════════════════════════════════════════════════════════════════════════

def compute_address_group(
        actual_df, db1_df, db2_df,
        actual_component_cols,   # e.g. ["primary street","primary city",...]
        db1_slot_col_lists,      # list of lists: [[slot1_street,slot1_city,...], ...]
        db2_component_cols,      # e.g. ["Home Address Line 1","Home Address City",...]
        actual_lookup, db1_key, db2_key,
        label: str,              # "primary" or "permanent"
        db2_extra_component_cols=None,  # optional 2nd DB2 set, e.g. Permanent
    ) -> dict:
    """
    Build a combined address block for one address group (primary or permanent).

    Returns a dict suitable for appending to blocks[]:
    {
        "label":           "primary address",
        "actual_combined": Series of raw combined actual address strings,
        "actual_norm":     Series of normalized combined strings,
        "actual_blank":    boolean Series,
        "filled_in_actual":int,
        "db1_slots": [
            {"slot": 1, "combined": Series, "norm": Series,
             "score": Series, "colour": Series, "diff": Series},
            ...
        ],
        "db2": {
            "combined": Series, "norm": Series,
            "score": Series, "colour": Series, "diff": Series,
        },
        "best_colour":  Series ("HIGH"/"MID"/"LOW"),
        "review_db1":   Series ("TRUE"/"FALSE"/"MISSING"),
        "review_db2":   Series ("TRUE"/"FALSE"/"MISSING"),
        "stats_db1": { matched, missing, mismatch, total, match_pct, filled_match_pct },
        "stats_db2": { ... },
    }
    """
    n = len(actual_df)
    idx = actual_df.index

    # ── Build actual combined address ─────────────────────────────────────────
    existing_actual = [c for c in actual_component_cols if c in actual_df.columns]
    if existing_actual:
        raw_parts = [actual_df[c].astype(str).str.strip() for c in existing_actual]
        actual_raw = pd.Series(
            [" ".join(p for p in parts if p and p.lower() not in _EMPTY_TOKENS)
             for parts in zip(*[s.tolist() for s in raw_parts])],
            index=idx
        )
    else:
        actual_raw = pd.Series([""] * n, index=idx)

    actual_blank = _is_blank(actual_raw)
    filled_in_actual = int((~actual_blank).sum())
    actual_norm_list = [combine_and_normalize(
        [actual_df[c].iloc[i] for c in existing_actual if c in actual_df.columns]
    ) for i in range(n)]
    actual_norm = pd.Series(actual_norm_list, index=idx)

    def _score_and_diff(norm_a_list, norm_b_list, blank_a, blank_b):
        scores  = []
        colours = []
        diffs   = []
        for na, nb, ba, bb in zip(norm_a_list, norm_b_list, blank_a, blank_b):
            if ba and bb:
                sc = 1.0; col = "HIGH"; di = "both blank"
            elif ba:
                sc = 1.0; col = "HIGH"; di = "actual blank"
            elif bb:
                sc = 0.0; col = "LOW";  di = f"actual: {na[:60]} | DB blank/missing"
            else:
                sc  = address_similarity(na, nb)
                col = _addr_colour(sc)
                di  = address_diff(na, nb)
            scores.append(sc); colours.append(col); diffs.append(di)
        return (pd.Series(scores, index=idx),
                pd.Series(colours, index=idx),
                pd.Series(diffs, index=idx))

    # ── DB1: compare against each slot (1-10) ─────────────────────────────────
    db1_slots_out = []
    best_score_db1 = pd.Series([0.0] * n, index=idx)

    for slot_cols in db1_slot_col_lists:
        existing_slot = [c for c in slot_cols if c in db1_df.columns]
        if not existing_slot:
            continue

        # Lookup each component for this slot.
        # Use _multi_key_lookup so cells containing multiple IDs (e.g. "123,456")
        # are split and each ID is tried until a non-blank value is found.
        slot_parts = []
        for sc in existing_slot:
            _m = dict(zip(db1_key, db1_df[sc].astype(str).str.strip()))
            lv = actual_lookup["db1"].apply(lambda k: _multi_key_lookup(k, _m))
            slot_parts.append(lv)

        # Combine into one string per row
        slot_raw = pd.Series(
            [" ".join(p for p in parts if p and p.lower() not in _EMPTY_TOKENS)
             for parts in zip(*[s.tolist() for s in slot_parts])],
            index=idx
        )
        slot_blank = _is_blank(slot_raw)
        slot_norm  = pd.Series(
            [combine_and_normalize([p.iloc[i] for p in slot_parts]) for i in range(n)],
            index=idx
        )

        score_s, colour_s, diff_s = _score_and_diff(
            actual_norm.tolist(), slot_norm.tolist(),
            actual_blank.tolist(), slot_blank.tolist()
        )

        # Track best score
        improve = score_s > best_score_db1
        best_score_db1 = best_score_db1.where(~improve, score_s)

        slot_label = existing_slot[0].split()[0] + " " + existing_slot[0].split()[1]  # e.g. "Address 1"
        db1_slots_out.append({
            "slot_label": slot_label,
            "combined":   slot_raw,
            "norm":       slot_norm,
            "score":      score_s,
            "colour":     colour_s,
            "diff":       diff_s,
        })

    # Consolidated DB1 result
    best_colour_db1 = best_score_db1.apply(_addr_colour)
    # blank actual + no DB data → HIGH (green)
    all_db1_blank = pd.Series(True, index=idx)
    for sl in db1_slots_out:
        all_db1_blank &= _is_blank(sl["combined"])
    best_colour_db1[actual_blank & all_db1_blank] = "HIGH"
    best_colour_db1[actual_blank & ~all_db1_blank] = "HIGH"  # blank actual always green

    review_db1 = pd.Series("FALSE", index=idx)
    review_db1[all_db1_blank & ~actual_blank] = "MISSING"
    review_db1[best_score_db1 >= ADDRESS_HIGH_THRESHOLD] = "TRUE"
    review_db1[actual_blank] = "TRUE"  # blank actual always green

    matched_db1  = int((review_db1 == "TRUE").sum())
    missing_db1  = int((review_db1 == "MISSING").sum())
    mismatch_db1 = int((review_db1 == "FALSE").sum())
    pct_db1      = round(100 * matched_db1 / n, 1) if n else 0.0
    nb_true_db1  = int(((review_db1 == "TRUE") & ~actual_blank).sum())
    fpct_db1     = round(100 * nb_true_db1 / filled_in_actual, 1) if filled_in_actual else 0.0

    # ── DB2 ───────────────────────────────────────────────────────────────────
    # DB2 may store an address in more than one place (e.g. Almabase keeps a
    # clean split under "Home Address *" but ALSO a free-text copy under
    # "Permanent Address"). The actual primary address sometimes matches one but
    # not the other, so we score the actual value against EACH available DB2
    # component set and keep, per row, the best-scoring set. This rescues cases
    # where Home is a clean split (low token overlap) yet Permanent holds a
    # near-verbatim copy of the actual blob.
    db2_component_sets = [db2_component_cols]
    if db2_extra_component_cols:
        db2_component_sets.append(db2_extra_component_cols)

    cand_raws, cand_norms, cand_scores, cand_colours, cand_diffs = [], [], [], [], []
    for comp_set in db2_component_sets:
        existing_db2 = [c for c in comp_set if c in db2_df.columns]
        if not existing_db2:
            continue
        db2_parts = []
        for dc in existing_db2:
            _m = dict(zip(db2_key, db2_df[dc].astype(str).str.strip()))
            lv = actual_lookup["db2"].apply(lambda k: _multi_key_lookup(k, _m))
            db2_parts.append(lv)

        c_raw = pd.Series(
            [" ".join(p for p in parts if p and p.lower() not in _EMPTY_TOKENS)
             for parts in zip(*[s.tolist() for s in db2_parts])],
            index=idx
        )
        c_blank = _is_blank(c_raw)
        c_norm  = pd.Series(
            [combine_and_normalize([p.iloc[i] for p in db2_parts]) for i in range(n)],
            index=idx
        )
        c_score, c_colour, c_diff = _score_and_diff(
            actual_norm.tolist(), c_norm.tolist(),
            actual_blank.tolist(), c_blank.tolist()
        )
        cand_raws.append(c_raw.tolist());     cand_norms.append(c_norm.tolist())
        cand_scores.append(c_score.tolist()); cand_colours.append(c_colour.tolist())
        cand_diffs.append(c_diff.tolist())

    if cand_scores:
        # Per row, keep the DB2 component set with the highest similarity score.
        k = len(cand_scores)
        raw_l, norm_l, score_l, colour_l, diff_l = [], [], [], [], []
        for i in range(n):
            best_j = max(range(k), key=lambda j: cand_scores[j][i])
            raw_l.append(cand_raws[best_j][i]);     norm_l.append(cand_norms[best_j][i])
            score_l.append(cand_scores[best_j][i]);  colour_l.append(cand_colours[best_j][i])
            diff_l.append(cand_diffs[best_j][i])
        db2_raw    = pd.Series(raw_l, index=idx)
        db2_blank  = _is_blank(db2_raw)
        db2_norm   = pd.Series(norm_l, index=idx)
        score_db2  = pd.Series(score_l, index=idx)
        colour_db2 = pd.Series(colour_l, index=idx)
        diff_db2   = pd.Series(diff_l, index=idx)
    else:
        db2_raw    = pd.Series([""] * n, index=idx)
        db2_blank  = pd.Series([True] * n, index=idx)
        db2_norm   = pd.Series([""] * n, index=idx)
        score_db2  = pd.Series([0.0] * n, index=idx)
        colour_db2 = pd.Series(["LOW"] * n, index=idx)
        diff_db2   = pd.Series(["DB2 columns not found"] * n, index=idx)

    colour_db2[actual_blank] = "HIGH"
    review_db2 = pd.Series("FALSE", index=idx)
    review_db2[db2_blank & ~actual_blank]          = "MISSING"
    review_db2[score_db2 >= ADDRESS_HIGH_THRESHOLD]= "TRUE"
    review_db2[actual_blank]                       = "TRUE"

    matched_db2  = int((review_db2 == "TRUE").sum())
    missing_db2  = int((review_db2 == "MISSING").sum())
    mismatch_db2 = int((review_db2 == "FALSE").sum())
    pct_db2      = round(100 * matched_db2 / n, 1) if n else 0.0
    nb_true_db2  = int(((review_db2 == "TRUE") & ~actual_blank).sum())
    fpct_db2     = round(100 * nb_true_db2 / filled_in_actual, 1) if filled_in_actual else 0.0

    # ── Overall cell colour (best across DB1 + DB2) ───────────────────────────
    priority = {"HIGH": 2, "MID": 1, "LOW": 0}
    overall_colour = best_colour_db1.copy()
    upgrade = colour_db2.map(priority) > overall_colour.map(priority)
    overall_colour[upgrade] = colour_db2[upgrade]

    return {
        "label":           f"{label} address",
        "actual_combined": actual_raw,
        "actual_norm":     actual_norm,
        "actual_blank":    actual_blank,
        "filled_in_actual": filled_in_actual,
        "db1_slots":       db1_slots_out,
        "db2": {
            "combined": db2_raw,
            "norm":     db2_norm,
            "score":    score_db2,
            "colour":   colour_db2,
            "diff":     diff_db2,
        },
        "best_colour":  overall_colour,
        "review_db1":   review_db1,
        "review_db2":   review_db2,
        "stats_db1": {
            "matched": matched_db1, "missing": missing_db1,
            "mismatch": mismatch_db1, "total": n,
            "match_pct": pct_db1, "filled_match_pct": fpct_db1,
            "filled_in_actual": filled_in_actual,
        },
        "stats_db2": {
            "matched": matched_db2, "missing": missing_db2,
            "mismatch": mismatch_db2, "total": n,
            "match_pct": pct_db2, "filled_match_pct": fpct_db2,
            "filled_in_actual": filled_in_actual,
        },
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  REVIEW COMPUTATION  (standard fields)
# ═══════════════════════════════════════════════════════════════════════════════

def compute_review(actual_df, db1_df, db2_df,
                   actual_db1_id, actual_db2_id,
                   db1_id_col, db2_id_col,
                   plan: list, workers: int = 1,
                   always_green_cols: set = None) -> list:
    always_green_cols = always_green_cols or set()

    print(f"\n[4] Running XLOOKUP for {len(plan)} fields x {len(actual_df):,} rows ...")

    actual_db1_lookup = actual_df[actual_db1_id].astype(str).str.strip()
    actual_db2_lookup = actual_df[actual_db2_id].astype(str).str.strip()
    db1_key = db1_df[db1_id_col].astype(str).str.strip()
    db2_key = db2_df[db2_id_col].astype(str).str.strip()

    db1_val_cache, db2_val_cache = {}, {}
    cache_lock = threading.Lock()

    def safe_xlookup(actual_lookup, db_df, key_stripped, col_name, cache,
                     target_norm=None, norm_type="text"):
        with cache_lock:
            m = cache.get(col_name)
        if m is None:
            m_local = dict(zip(key_stripped, db_df[col_name].astype(str).str.strip()))
            with cache_lock:
                m = cache.setdefault(col_name, m_local)
        # Use _multi_key_lookup so cells with multiple IDs (e.g. "123,456")
        # are split and each ID is tried against the dump until a value is found.
        # When the actual (normalized) value is provided, prefer the linked
        # record that actually matches it, so multi-ID merges don't false-RED.
        if target_norm is not None:
            return pd.Series(
                [_multi_key_lookup_match(k, m, t, norm_type)
                 for k, t in zip(actual_lookup, target_norm)],
                index=actual_lookup.index)
        return actual_lookup.apply(lambda k: _multi_key_lookup(k, m))

    progress = {"done": 0}
    prog_lock = threading.Lock()
    total_fields = len(plan)

    def process_field(field_idx, entry):
        ac        = entry["actual_col"]
        norm_type = entry["norm_type"]
        actual_vals = actual_df[ac].astype(str).str.strip()
        actual_blank = _is_blank(actual_vals)
        filled_in_actual = int((~actual_blank).sum())

        # ── Always-green short-circuit: skip all validation, force GREEN. ──────
        if ac in always_green_cols:
            color = pd.Series("GREEN", index=actual_df.index)
            with prog_lock:
                progress["done"] += 1
                print(f"  [{progress['done']}/{total_fields}] {ac} "
                      f"(always-green, filled={filled_in_actual:,})")
            return field_idx, {
                "actual_col": ac, "norm_type": norm_type,
                "actual_vals": actual_vals, "actual_blank": actual_blank,
                "filled_in_actual": filled_in_actual,
                "db1_sub": [], "db2_sub": [],
                "actual_color": color, "overall_pct": 100.0,
                "is_address": False, "is_addr_group": False,
            }

        actual_norm = normalize_series(actual_vals, norm_type)

        def make_sub(db_df, actual_lookup, key_stripped, cache, db_cols_list):
            valid = [c for c in db_cols_list if c in db_df.columns]
            if not valid: return []
            total = len(actual_df)
            lu_all, norm_all = {}, {}
            for dc in valid:
                lv = safe_xlookup(actual_lookup, db_df, key_stripped, dc, cache,
                                  target_norm=actual_norm, norm_type=norm_type)
                lu_all[dc]   = lv
                norm_all[dc] = normalize_series(lv, norm_type)

            # ── Minimum DB value length for combined/contains matching ────────────
            # Prevents trivially short values (e.g. "Mr", "01") from producing
            # false-positive combined hits.  Only applied for text-like field types;
            # dates and numbers are excluded because substring matching there would
            # almost always be a false positive (e.g. "2024" inside "2024-01-01").
            _COMBINED_MIN_LEN  = 4
            _COMBINED_ELIGIBLE = norm_type not in ("date", "number")

            match_any       = pd.Series(False, index=actual_df.index)
            fuzzy_match_any = pd.Series(False, index=actual_df.index)
            any_data        = pd.Series(False, index=actual_df.index)
            mslot           = pd.Series("",    index=actual_df.index)
            fuzzy_results    = {}
            exact_results    = {}
            combined_results = {}

            for dc in valid:
                db_blank  = _is_blank(lu_all[dc])
                has_data  = lu_all[dc] != ""
                both_blank= actual_blank & db_blank

                # ── Step 1: exact match ───────────────────────────────────────
                exact_match = both_blank | ((norm_all[dc] == actual_norm) & has_data)
                exact_results[dc] = exact_match

                # ── Step 2: fuzzy match ───────────────────────────────────────
                needs_fuzzy = ~exact_match & ~actual_blank & has_data & ~db_blank
                if needs_fuzzy.any():
                    fuzzy_m = _fuzzy_match_series(
                        actual_norm, norm_all[dc], needs_fuzzy,
                        norm_type=norm_type)
                else:
                    fuzzy_m = pd.Series(False, index=actual_df.index)
                fuzzy_results[dc] = fuzzy_m

                # ── Step 3: combined/contains match ───────────────────────────
                # Triggered when actual may be a concatenation of values from
                # multiple DB sources.  Uses _token_overlap_match which runs
                # two passes:
                #   Pass 1 — exact substring (fast, handles verbatim cases)
                #   Pass 2 — token overlap   (handles different word order,
                #             punctuation differences, and long multi-word values)
                combined_m = pd.Series(False, index=actual_df.index)
                if _COMBINED_ELIGIBLE:
                    needs_combined = (
                        ~exact_match & ~fuzzy_m
                        & ~actual_blank & has_data & ~db_blank
                    )
                    if needs_combined.any():
                        db_vals  = norm_all[dc]
                        act_vals = actual_norm
                        idx_list = actual_df.index
                        combined_arr = [
                            _token_overlap_match(
                                db_vals.iloc[i],
                                act_vals.iloc[i],
                                min_len=_COMBINED_MIN_LEN,
                            )
                            if needs_combined.iloc[i] else False
                            for i in range(len(idx_list))
                        ]
                        combined_m = pd.Series(combined_arr, index=idx_list)
                combined_results[dc] = combined_m

                is_match = exact_match | fuzzy_m | combined_m
                any_data |= has_data | both_blank
                mslot     = mslot.where(~is_match, other=dc)
                match_any       |= is_match
                fuzzy_match_any |= (fuzzy_m & ~exact_match)

            no_db_blank = actual_blank & ~any_data
            match_any  |= no_db_blank
            any_data   |= no_db_blank

            # ── Consolidated review across all DB columns for this DB source ──
            # Priority (highest wins): TRUE > FUZZY > MISSING > FALSE
            # Combined/contains match is treated identically to an exact match
            # → "TRUE" and GREEN. Only RED if nothing matched at all.
            exact_any    = pd.Series(False, index=actual_df.index)
            fuzzy_any    = pd.Series(False, index=actual_df.index)
            combined_any = pd.Series(False, index=actual_df.index)
            for dc in valid:
                exact_any    |= exact_results[dc]
                fuzzy_any    |= fuzzy_results.get(dc,    pd.Series(False, index=actual_df.index))
                combined_any |= combined_results.get(dc, pd.Series(False, index=actual_df.index))

            cons = pd.Series("FALSE", index=actual_df.index)
            cons[~any_data]  = "MISSING"
            cons[match_any]  = "TRUE"
            # Downgrade to FUZZY only when: fuzzy hit exists, no exact hit, no combined hit
            pure_fuzzy = match_any & ~exact_any & ~combined_any & fuzzy_any & ~no_db_blank
            cons[pure_fuzzy] = "FUZZY"
            # Combined match stays "TRUE" — same as exact match for review and colour

            matched  = int(((cons == "TRUE") | (cons == "FUZZY")).sum())
            missing  = int((cons == "MISSING").sum())
            mismatch = int((cons == "FALSE").sum())
            pct      = round(100 * matched / total, 1) if total else 0.0
            nb_true  = int((((cons == "TRUE") | (cons == "FUZZY")) & ~actual_blank).sum())
            fpct     = round(100 * nb_true / filled_in_actual, 1) if filled_in_actual else 0.0

            sub = []
            for dc in valid:
                lv         = lu_all[dc]
                db_blnk    = _is_blank(lv)
                no_data    = lv == ""
                exact_m    = exact_results[dc]
                fuzzy_m    = fuzzy_results.get(dc,    pd.Series(False, index=actual_df.index))
                combined_m = combined_results.get(dc, pd.Series(False, index=actual_df.index))
                indiv      = pd.Series("FALSE", index=actual_df.index)
                indiv[actual_blank & db_blnk]              = "TRUE"
                indiv[actual_blank & no_data]              = "TRUE"
                indiv[no_data & ~actual_blank]             = "MISSING"
                indiv[exact_m & ~no_data & ~actual_blank]  = "TRUE"
                if fuzzy_m.any():
                    indiv[fuzzy_m & (indiv == "FALSE")]    = "FUZZY"
                # Combined match → TRUE (green), same as exact
                if combined_m.any():
                    indiv[combined_m & (indiv == "FALSE")] = "TRUE"
                sub.append({
                    "db_col": dc,
                    "values": norm_all[dc] if norm_type == "date" else lv,
                    "review": indiv, "consolidated_review": cons,
                    "match_pct": pct, "filled_match_pct": fpct,
                    "matched": matched, "missing": missing, "mismatch": mismatch,
                    "total": total, "filled_in_actual": filled_in_actual,
                    "is_match_any": True,
                })
            del lu_all, norm_all, match_any, any_data, mslot
            return sub

        db1_sub = make_sub(db1_df, actual_db1_lookup, db1_key,
                           db1_val_cache, entry["db1_cols"])
        db2_sub = make_sub(db2_df, actual_db2_lookup, db2_key,
                           db2_val_cache, entry["db2_cols"])
        all_sub = db1_sub + db2_sub

        # ── Detect boolean/categorical values → always GREEN ──────────────────
        bool_mask = _is_bool_like_series(actual_vals)

        if all_sub:
            reviews = []
            if db1_sub: reviews.append(db1_sub[0]["consolidated_review"])
            if db2_sub: reviews.append(db2_sub[0]["consolidated_review"])
            # Actual column colour: GREEN (exact or combined match), FUZZY_GREEN (fuzzy only), RED (no match)
            exact_true = sum(((r == "TRUE").astype(int))  for r in reviews)
            fuzzy_true = sum(((r == "FUZZY").astype(int)) for r in reviews)
            color = pd.Series("RED", index=actual_df.index)
            color[actual_blank]       = "GREEN"
            color[fuzzy_true  >= 1]   = "FUZZY_GREEN"
            color[exact_true  >= 1]   = "GREEN"   # exact AND combined both map to TRUE → GREEN
            color[bool_mask]          = "GREEN"
        else:
            # No DB mapping found for this column → show GREEN (unmapped = not reviewable)
            color = pd.Series("GREEN", index=actual_df.index)
            color[actual_blank] = "GREEN"

        all_pcts    = [s["match_pct"] for s in all_sub]
        overall_pct = round(sum(all_pcts) / len(all_pcts), 1) if all_pcts else 0.0

        del actual_norm
        with prog_lock:
            progress["done"] += 1
            print(f"  [{progress['done']}/{total_fields}] {ac} "
                  f"(type={norm_type}, filled={filled_in_actual:,})")

        # For date fields, display the normalized dd-mmm-yyyy format in the
        # output Excel instead of the raw value (which may carry a "00:00:00"
        # timestamp from pandas Excel parsing).
        display_actual_vals = (normalize_series(actual_vals, norm_type)
                               if norm_type == "date" else actual_vals)

        return field_idx, {
            "actual_col": ac, "norm_type": norm_type,
            "actual_vals": display_actual_vals, "actual_blank": actual_blank,
            "filled_in_actual": filled_in_actual,
            "db1_sub": db1_sub, "db2_sub": db2_sub,
            "actual_color": color, "overall_pct": overall_pct,
            "is_address": False, "is_addr_group": False,
        }

    items = list(enumerate(plan, 1))
    if workers <= 1:
        results = [process_field(i, e) for i, e in items]
    else:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            results = list(ex.map(lambda ie: process_field(*ie), items))

    results.sort(key=lambda x: x[0])
    blocks = [b for _, b in results]
    gc.collect()
    print("  -> Computation complete.")
    return blocks


# ═══════════════════════════════════════════════════════════════════════════════
#  WORKBOOK WRITER HELPER
# ═══════════════════════════════════════════════════════════════════════════════

def _make_actuals_only_workbook(actual_df, blocks, addr_groups, out_path, workers):
    """
    Compact review file: IDs + each standard actual column (coloured GREEN /
    FUZZY_GREEN / RED / YELLOW) + combined address columns + per-row Red count
    + Yellow count.

    Columns are ordered to match the Salesforce input file. Each combined
    address column is slotted where its first component column sits.

    Counting rule per row:
      RED state (or address LOW)    -> red_n
      YELLOW state (or address MID) -> yellow_n
      GREEN / FUZZY_GREEN / HIGH    -> neither
    """
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Actuals Review")
    ws.freeze_panes = "A2"

    # Build the data columns (standard blocks + address groups) in input order.
    col_pos  = _input_col_positions(actual_df)
    unsorted = []
    for b in blocks:
        unsorted.append((col_pos.get(b["actual_col"], float("inf")), {
            "header":  b["actual_col"],
            "vals":    b["actual_vals"].tolist(),
            "codes":   b["actual_color"].tolist(),
            "is_addr": False,
        }))
    for ag in addr_groups:
        comp_cols   = _addr_group_components(ag, col_pos)
        addr_codes  = ag["best_colour"].tolist()
        for c in comp_cols:
            unsorted.append((col_pos[c], {
                "header":  c,
                "vals":    actual_df[c].astype(str).tolist(),
                "codes":   addr_codes,
                "is_addr": True,
            }))
        combo_pos = (max(col_pos[c] for c in comp_cols) + 0.5
                     if comp_cols else _addr_group_position(ag, col_pos))
        unsorted.append((combo_pos, {
            "header":  f"{ag['label']} (address)",
            "vals":    ag["actual_combined"].tolist(),
            "codes":   addr_codes,
            "is_addr": True,
        }))
    unsorted.sort(key=lambda x: x[0])
    data_cols = [d for _, d in unsorted]

    id_cols = list(actual_df.columns[:2])
    actual_cols = [d["header"] for d in data_cols]
    headers = id_cols + actual_cols + ["Red cells", "Yellow cells"]
    n_id    = len(id_cols)
    n_act   = len(actual_cols)
    n_cols  = len(headers)
    n_rows  = len(actual_df)

    # Column widths + header row
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(14, min(len(h) + 4, 38))
    ws.row_dimensions[1].height = 32

    hdr_row = []
    for i, h in enumerate(headers):
        c = WriteOnlyCell(ws, value=h)
        c.font = HEADER_FONT
        if i < n_id:               c.fill = HEADER_FILL
        elif i < n_id + n_act:     c.fill = HEADER_FILL
        elif h == "Red cells":     c.fill = RED_FILL; c.font = HEADER_FONT
        else:                      c.fill = YELLOW_FILL; c.font = HEADER_FONT
        c.alignment = CENTER
        hdr_row.append(c)
    ws.append(hdr_row)

    def _clean(v):
        return _illegal_sub("", v) if isinstance(v, str) and v else v

    # Pre-materialise per-column value + colour lists
    id_vals   = [actual_df[c].astype(str).tolist() for c in id_cols]
    act_vals  = [d["vals"]    for d in data_cols]
    act_codes = [d["codes"]   for d in data_cols]
    act_addr  = [d["is_addr"] for d in data_cols]

    def build_row(r_idx):
        row = [None] * n_cols
        # ID cells
        for i in range(n_id):
            row[i] = _clean(id_vals[i][r_idx])
        # Actual cells
        red_n = 0
        yellow_n = 0
        for i in range(n_act):
            code = act_codes[i][r_idx]
            v    = _clean(act_vals[i][r_idx])
            cell = WriteOnlyCell(ws, value=v)
            cell.font = NORMAL_FONT
            if act_addr[i]:
                style = _ADDR_MERGE_STYLE.get(code)
                if code == "LOW":   red_n += 1
                elif code == "MID": yellow_n += 1
            else:
                style = _STD_ACTUAL_STYLE.get(code)
                if code == "RED":      red_n += 1
                elif code == "YELLOW": yellow_n += 1
            if style: cell.fill, cell.font = style
            row[n_id + i] = cell
        # Red count
        c_red = WriteOnlyCell(ws, value=red_n)
        c_red.alignment = CENTER
        c_red.fill, c_red.font = (RED_FILL, RED_FONT) if red_n > 0 else (GREEN_FILL, GREEN_FONT)
        row[n_id + n_act] = c_red
        # Yellow count
        c_yel = WriteOnlyCell(ws, value=yellow_n)
        c_yel.alignment = CENTER
        c_yel.fill, c_yel.font = (YELLOW_FILL, YELLOW_FONT) if yellow_n > 0 else (GREEN_FILL, GREEN_FONT)
        row[n_id + n_act + 1] = c_yel
        return row

    print(f"  Writing actuals-only sheet: {n_cols} cols x {n_rows:,} rows ...")
    t0 = time.time()
    if workers <= 1 or n_rows <= 2000:
        for r in range(n_rows):
            ws.append(build_row(r))
            if (r + 1) % 10000 == 0:
                print(f"    ... {r+1:,}/{n_rows:,}")
    else:
        CHUNK = max(500, n_rows // (workers * 8))
        chunks = [(i, min(i+CHUNK, n_rows)) for i in range(0, n_rows, CHUNK)]
        appended = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            for chunk_rows in ex.map(lambda b: [build_row(r) for r in range(*b)], chunks):
                for row in chunk_rows:
                    ws.append(row)
                appended += len(chunk_rows)
                if appended % 10000 < CHUNK:
                    print(f"    ... {appended:,}/{n_rows:,}")
    ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}1"
    print(f"  Actuals sheet done in {time.time()-t0:.1f}s")

    print(f"  Saving -> {out_path}")
    wb.save(out_path)
    return out_path


def _make_workbook_and_review_sheet(actual_df, blocks, addr_groups,
                                    out_path, workers):
    """
    Build the full workbook (Review + Summary + Field Audit + Address Legend).
    Returns the saved workbook path.
    """
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Review")
    ws.freeze_panes = "A2"

    col_entries = []
    audit_rows  = []
    # strip Excel-illegal characters from cell values (module-level _illegal_sub)

    def add(header, series, col_type, colour_series=None,
            hdr_fill=HEADER_FILL, extra=None):
        col_entries.append({
            "header": header, "series": series, "type": col_type,
            "colour_series": colour_series, "hdr_fill": hdr_fill,
            "extra": extra or {},
        })
        return len(col_entries)

    # ── ID columns ─────────────────────────────────────────────────────────────
    for col in list(actual_df.columns[:2]):
        add(col, actual_df[col], "value")

    # ── Standard field blocks ──────────────────────────────────────────────────
    for block in blocks:
        ac    = block["actual_col"]
        db1_s = block["db1_sub"]
        db2_s = block["db2_sub"]
        n1, n2 = len(db1_s), len(db2_s)
        nt    = block["norm_type"]

        act_idx = add(ac, block["actual_vals"], "actual",
                      colour_series=block["actual_color"])

        db1_val_idxs, db1_rev_idxs = [], []
        for i, sub in enumerate(db1_s):
            sfx = f" {i+1}" if n1 > 1 else ""
            db1_val_idxs.append(add(f"{ac} in db1{sfx}", sub["values"],
                                    "value", hdr_fill=DB1_HDR_FILL))
        db2_val_idxs, db2_rev_idxs = [], []
        for i, sub in enumerate(db2_s):
            sfx = f" {i+1}" if n2 > 1 else ""
            db2_val_idxs.append(add(f"{ac} in db2{sfx}", sub["values"],
                                    "value", hdr_fill=DB2_HDR_FILL))

        is_any_db1 = db1_s and db1_s[0].get("is_match_any", False)
        if is_any_db1:
            db1_rev_idxs.append(add(f"{ac} db1 review (any match)",
                db1_s[0]["consolidated_review"], "review", hdr_fill=REV_HDR_FILL))
        else:
            for i, sub in enumerate(db1_s):
                sfx = f" {i+1}" if n1 > 1 else ""
                db1_rev_idxs.append(add(f"{ac} db1{sfx} review",
                    sub["review"], "review", hdr_fill=REV_HDR_FILL))

        is_any_db2 = db2_s and db2_s[0].get("is_match_any", False)
        if is_any_db2:
            db2_rev_idxs.append(add(f"{ac} db2 review (any match)",
                db2_s[0]["consolidated_review"], "review", hdr_fill=REV_HDR_FILL))
        else:
            for i, sub in enumerate(db2_s):
                sfx = f" {i+1}" if n2 > 1 else ""
                db2_rev_idxs.append(add(f"{ac} db2{sfx} review",
                    sub["review"], "review", hdr_fill=REV_HDR_FILL))

        def _rl(rev_idxs, is_any, i):
            if not rev_idxs:       return "—"
            if is_any:             return get_column_letter(rev_idxs[0])
            if i < len(rev_idxs): return get_column_letter(rev_idxs[i])
            return "—"

        for i, sub in enumerate(db1_s):
            audit_rows.append({
                "Actual field": ac, "Norm type": nt, "DB": "DB1",
                "DB col": sub["db_col"],
                "Val col": get_column_letter(db1_val_idxs[i]) if i < len(db1_val_idxs) else "—",
                "Rev col": _rl(db1_rev_idxs, is_any_db1, i),
                "Total rows": sub["total"], "Filled in actual": sub["filled_in_actual"],
                "Matched": sub["matched"], "Mismatch": sub["mismatch"],
                "Missing": sub["missing"],
                "Exact match %": f"{sub['match_pct']}%",
                "Filled match %": f"{sub['filled_match_pct']}%",
            })
        for i, sub in enumerate(db2_s):
            audit_rows.append({
                "Actual field": ac, "Norm type": nt, "DB": "DB2",
                "DB col": sub["db_col"],
                "Val col": get_column_letter(db2_val_idxs[i]) if i < len(db2_val_idxs) else "—",
                "Rev col": _rl(db2_rev_idxs, is_any_db2, i),
                "Total rows": sub["total"], "Filled in actual": sub["filled_in_actual"],
                "Matched": sub["matched"], "Mismatch": sub["mismatch"],
                "Missing": sub["missing"],
                "Exact match %": f"{sub['match_pct']}%",
                "Filled match %": f"{sub['filled_match_pct']}%",
            })
        if not db1_s and not db2_s:
            audit_rows.append({
                "Actual field": ac, "Norm type": nt, "DB": "—",
                "DB col": "NOT FOUND", "Val col": get_column_letter(act_idx),
                "Rev col": "—", "Total rows": len(actual_df),
                "Filled in actual": block["filled_in_actual"],
                "Matched": 0, "Mismatch": 0, "Missing": len(actual_df),
                "Exact match %": "0.0%", "Filled match %": "0.0%",
            })

    # ── Address group blocks ───────────────────────────────────────────────────
    for ag in addr_groups:
        lbl = ag["label"]
        n   = len(actual_df)

        # Actual combined column
        add(f"{lbl} (combined)",
            ag["actual_combined"], "actual",
            colour_series=ag["best_colour"],
            hdr_fill=ADDR_HDR_FILL,
            extra={"is_address": True})

        # DB1 slots
        for sl in ag["db1_slots"]:
            add(f"{lbl} | {sl['slot_label']} (combined)",
                sl["combined"], "value", hdr_fill=DB1_HDR_FILL)
            score_str = sl["score"].apply(lambda s: f"{s:.2f}")
            add(f"{lbl} | {sl['slot_label']} score",
                score_str, "addr_score",
                colour_series=sl["colour"], hdr_fill=DIFF_HDR_FILL)
            add(f"{lbl} | {sl['slot_label']} diff ◄►",
                sl["diff"], "addr_diff", hdr_fill=DIFF_HDR_FILL)

        add(f"{lbl} DB1 review (any slot)",
            ag["review_db1"], "addr_review",
            colour_series=ag["review_db1"].map(
                {"TRUE": "HIGH", "FALSE": "LOW", "MISSING": "MID"}),
            hdr_fill=REV_HDR_FILL)

        # DB2
        db2 = ag["db2"]
        add(f"{lbl} | DB2 home/permanent (combined)",
            db2["combined"], "value", hdr_fill=DB2_HDR_FILL)
        score_str2 = db2["score"].apply(lambda s: f"{s:.2f}")
        add(f"{lbl} | DB2 score",
            score_str2, "addr_score",
            colour_series=db2["colour"], hdr_fill=DIFF_HDR_FILL)
        add(f"{lbl} | DB2 diff ◄►",
            db2["diff"], "addr_diff", hdr_fill=DIFF_HDR_FILL)
        add(f"{lbl} DB2 review",
            ag["review_db2"], "addr_review",
            colour_series=ag["review_db2"].map(
                {"TRUE": "HIGH", "FALSE": "LOW", "MISSING": "MID"}),
            hdr_fill=REV_HDR_FILL)

        # Audit rows for address group
        s1 = ag["stats_db1"]
        audit_rows.append({
            "Actual field": lbl, "Norm type": "address", "DB": "DB1 (any slot 1-10)",
            "DB col": "Address 1-10 Street/City/State/Country/Zip",
            "Val col": "—", "Rev col": "—",
            "Total rows": s1["total"], "Filled in actual": s1["filled_in_actual"],
            "Matched": s1["matched"], "Mismatch": s1["mismatch"], "Missing": s1["missing"],
            "Exact match %": f"{s1['match_pct']}%",
            "Filled match %": f"{s1['filled_match_pct']}%",
        })
        s2 = ag["stats_db2"]
        audit_rows.append({
            "Actual field": lbl, "Norm type": "address", "DB": "DB2",
            "DB col": "Home/Permanent Address combined",
            "Val col": "—", "Rev col": "—",
            "Total rows": s2["total"], "Filled in actual": s2["filled_in_actual"],
            "Matched": s2["matched"], "Mismatch": s2["mismatch"], "Missing": s2["missing"],
            "Exact match %": f"{s2['match_pct']}%",
            "Filled match %": f"{s2['filled_match_pct']}%",
        })

    # ── Write header row ───────────────────────────────────────────────────────
    for c_idx, entry in enumerate(col_entries, start=1):
        w = 55 if entry["type"] in ("addr_diff",) else max(14, min(len(entry["header"]) + 4, 38))
        ws.column_dimensions[get_column_letter(c_idx)].width = w
    ws.row_dimensions[1].height = 42

    hdr_row = []
    for entry in col_entries:
        c = WriteOnlyCell(ws, value=entry["header"])
        c.font = HEADER_FONT; c.fill = entry["hdr_fill"]; c.alignment = CENTER
        hdr_row.append(c)
    ws.append(hdr_row)

    # ── Pre-materialise ────────────────────────────────────────────────────────
    n_rows = len(actual_df)
    n_cols = len(col_entries)
    val_lists   = [e["series"].tolist() if e["series"] is not None else [""]*n_rows
                   for e in col_entries]
    col_lists   = [e["colour_series"].tolist() if e.get("colour_series") is not None else None
                   for e in col_entries]
    types  = [e["type"]           for e in col_entries]
    extras = [e.get("extra", {})  for e in col_entries]

    def _clean(v):
        return _illegal_sub("", v) if isinstance(v, str) and v else v

    def build_row(r_idx):
        row = [None] * n_cols
        for ci in range(n_cols):
            t  = types[ci]
            v  = _clean(val_lists[ci][r_idx])
            ex = extras[ci]
            if t == "value":
                row[ci] = v; continue
            cell = WriteOnlyCell(ws, value=v)
            cell.font = NORMAL_FONT
            if t == "actual":
                code = col_lists[ci][r_idx] if col_lists[ci] else None
                style = (_ADDR_ACTUAL_STYLE if ex.get("is_address") else _STD_ACTUAL_STYLE
                         ).get(code)
                if style: cell.fill, cell.font = style
            elif t == "review":
                cell.alignment = CENTER
                code = v  # "TRUE", "FUZZY", "FALSE", "MISSING"
                if code == "FUZZY":
                    cell.value = "~FUZZY"
                style = _REVIEW_STYLE.get(code)
                if style: cell.fill, cell.font = style
            elif t == "addr_review":
                cell.alignment = CENTER
                code  = col_lists[ci][r_idx] if col_lists[ci] else "LOW"
                label = {"HIGH": "✅ MATCH", "MID": "⚠ PARTIAL", "LOW": "❌ MISMATCH"}.get(code, v)
                cell.value = label
                style = _ADDR_CELL_STYLE.get(code)
                if style: cell.fill, cell.font = style
            elif t == "addr_score":
                cell.alignment = CENTER; cell.font = SCORE_FONT
                code  = col_lists[ci][r_idx] if col_lists[ci] else "LOW"
                style = _ADDR_CELL_STYLE.get(code)
                if style: cell.fill = style[0]
            elif t == "addr_diff":
                cell.alignment = WRAP_L; cell.font = DIFF_FONT; cell.fill = DIFF_FILL
                if v == "tokens match":
                    cell.fill = LIGHT_GREEN_FILL; cell.font = GREEN_FONT
            row[ci] = cell
        return row

    print(f"  Writing {n_cols} columns x {n_rows:,} rows ...")
    t_w = time.time()
    if workers <= 1 or n_rows <= 2000:
        for r_idx in range(n_rows):
            ws.append(build_row(r_idx))
            if (r_idx + 1) % 10000 == 0:
                print(f"    ... {r_idx+1:,}/{n_rows:,}")
    else:
        CHUNK = max(500, n_rows // (workers * 8))
        chunks = [(i, min(i+CHUNK, n_rows)) for i in range(0, n_rows, CHUNK)]
        appended = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            for chunk_rows in ex.map(lambda b: [build_row(r) for r in range(*b)], chunks):
                for row in chunk_rows:
                    ws.append(row)
                appended += len(chunk_rows)
                if appended % 10000 < CHUNK:
                    print(f"    ... {appended:,}/{n_rows:,}")
    print(f"  Review sheet done in {time.time()-t_w:.1f}s")
    ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}1"

    # ── SHEET 2: Summary ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    sum_hdrs = ["Field","Norm type","DB1 mapped","DB2 mapped","Filled in actual",
                "GREEN / HIGH","YELLOW / MID","RED / LOW","Overall match %"]
    sum_fills = [HEADER_FILL,HEADER_FILL,DB1_HDR_FILL,DB2_HDR_FILL,FILLED_ACTUAL_FILL,
                 PatternFill("solid",start_color="375623"),
                 PatternFill("solid",start_color="7D6608"),
                 PatternFill("solid",start_color="9C0006"), HEADER_FILL]
    for c in range(1, len(sum_hdrs)+1):
        ws2.column_dimensions[get_column_letter(c)].width = 28
    ws2.row_dimensions[1].height = 42
    ws2.freeze_panes = "A2"
    h2 = []
    for h, f in zip(sum_hdrs, sum_fills):
        c = WriteOnlyCell(ws2, value=h)
        c.font = HEADER_FONT; c.fill = f; c.alignment = WRAP_C
        h2.append(c)
    ws2.append(h2)

    def _sc(ws_, v, fill, font, aln=None):
        c = WriteOnlyCell(ws_, value=v)
        c.fill=fill; c.font=font
        if aln: c.alignment=aln
        return c

    # Standard blocks
    for block in blocks:
        ac_col = block["actual_color"]
        green  = int(((ac_col == "GREEN") | (ac_col == "FUZZY_GREEN")).sum())
        yellow = int((ac_col == "YELLOW").sum())
        red    = int((ac_col == "RED").sum())
        pct    = block["overall_pct"]
        filled = block["filled_in_actual"]
        d1n    = ", ".join(s["db_col"] for s in block["db1_sub"]) or "not found"
        d2n    = ", ".join(s["db_col"] for s in block["db2_sub"]) or "not found"
        c1=WriteOnlyCell(ws2,value=block["actual_col"]); c1.font=NORMAL_FONT
        c2=WriteOnlyCell(ws2,value=block["norm_type"]);  c2.font=NORMAL_FONT
        c3=WriteOnlyCell(ws2,value=d1n);                 c3.font=NORMAL_FONT
        c4=WriteOnlyCell(ws2,value=d2n);                 c4.font=NORMAL_FONT
        c5=WriteOnlyCell(ws2,value=filled); c5.font=NORMAL_FONT; c5.alignment=CENTER; c5.fill=FILLED_ACTUAL_FILL
        pf,pft = (GREEN_FILL,GREEN_FONT) if pct>=95 else (YELLOW_FILL,YELLOW_FONT) if pct>=80 else (RED_FILL,RED_FONT)
        ws2.append([c1,c2,c3,c4,c5,
                    _sc(ws2,green,GREEN_FILL,GREEN_FONT,CENTER),
                    _sc(ws2,yellow,YELLOW_FILL,YELLOW_FONT,CENTER),
                    _sc(ws2,red,RED_FILL,RED_FONT,CENTER),
                    _sc(ws2,f"{pct}%",pf,pft,CENTER)])

    # Address group rows
    for ag in addr_groups:
        ac_col = ag["best_colour"]
        high   = int((ac_col=="HIGH").sum())
        mid    = int((ac_col=="MID").sum())
        low    = int((ac_col=="LOW").sum())
        pct    = ag["stats_db1"]["match_pct"]
        filled = ag["filled_in_actual"]
        c1=WriteOnlyCell(ws2,value=ag["label"]); c1.font=NORMAL_FONT
        c2=WriteOnlyCell(ws2,value="address");   c2.font=NORMAL_FONT
        c3=WriteOnlyCell(ws2,value="Address 1-10 (any slot)"); c3.font=NORMAL_FONT
        c4=WriteOnlyCell(ws2,value="Home/Permanent combined"); c4.font=NORMAL_FONT
        c5=WriteOnlyCell(ws2,value=filled); c5.font=NORMAL_FONT; c5.alignment=CENTER; c5.fill=FILLED_ACTUAL_FILL
        pf,pft=(GREEN_FILL,GREEN_FONT) if pct>=95 else (YELLOW_FILL,YELLOW_FONT) if pct>=80 else (RED_FILL,RED_FONT)
        ws2.append([c1,c2,c3,c4,c5,
                    _sc(ws2,high,GREEN_FILL,GREEN_FONT,CENTER),
                    _sc(ws2,mid,YELLOW_FILL,YELLOW_FONT,CENTER),
                    _sc(ws2,low,RED_FILL,RED_FONT,CENTER),
                    _sc(ws2,f"{pct}%",pf,pft,CENTER)])

    ws2.auto_filter.ref = f"A1:{get_column_letter(len(sum_hdrs))}{1+len(blocks)+len(addr_groups)}"

    # ── SHEET 3: Field Audit ───────────────────────────────────────────────────
    ws3 = wb.create_sheet("Field Audit")
    ws3.sheet_tab_color = "4A235A"
    aud_hdrs = ["Actual field","Norm type","DB","DB col used",
                "Val col (letter)","Rev col (letter)",
                "Total rows","Filled in actual",
                "Matched (TRUE)","Mismatch (FALSE)","Missing",
                "Exact match %","Filled match %"]
    aud_fills= [AUDIT_FILL,AUDIT_FILL,AUDIT_FILL,AUDIT_FILL,
                DB1_HDR_FILL,REV_HDR_FILL,
                HEADER_FILL,FILLED_ACTUAL_FILL,
                PatternFill("solid",start_color="375623"),
                PatternFill("solid",start_color="9C0006"),
                PatternFill("solid",start_color="7D6608"),
                HEADER_FILL,FILLED_PCT_FILL]
    aud_widths=[38,14,16,42,16,16,12,16,16,16,12,14,14]
    for c,w in enumerate(aud_widths,1):
        ws3.column_dimensions[get_column_letter(c)].width=w
    ws3.row_dimensions[1].height=42; ws3.freeze_panes="A2"
    h3=[]
    for h,f in zip(aud_hdrs,aud_fills):
        c=WriteOnlyCell(ws3,value=h)
        c.font=HEADER_FONT; c.fill=f; c.alignment=WRAP_C
        h3.append(c)
    ws3.append(h3)
    key_map={"Actual field":"Actual field","Norm type":"Norm type","DB":"DB",
             "DB col used":"DB col","Val col (letter)":"Val col",
             "Rev col (letter)":"Rev col","Total rows":"Total rows",
             "Filled in actual":"Filled in actual","Matched (TRUE)":"Matched",
             "Mismatch (FALSE)":"Mismatch","Missing":"Missing",
             "Exact match %":"Exact match %","Filled match %":"Filled match %"}
    se=PatternFill("solid",start_color="F2EEF7")
    so=PatternFill("solid",start_color="FFFFFF")
    for ri,row in enumerate(audit_rows,2):
        stripe=se if ri%2==0 else so
        out=[]
        for ci,hdr in enumerate(aud_hdrs,1):
            key=key_map[hdr]
            val=row.get(key,"")
            cell=WriteOnlyCell(ws3,value=val)
            cell.font=NORMAL_FONT
            cell.alignment=CENTER if ci>=5 else LEFT
            if hdr in ("Exact match %","Filled match %"):
                try:
                    pv=float(str(val).replace("%",""))
                    cell.fill,cell.font=(GREEN_FILL,GREEN_FONT) if pv>=95 else (YELLOW_FILL,YELLOW_FONT) if pv>=80 else (RED_FILL,RED_FONT)
                except: pass
            elif hdr=="Matched (TRUE)":   cell.fill=GREEN_FILL; cell.font=GREEN_FONT
            elif hdr=="Mismatch (FALSE)": cell.fill=RED_FILL;   cell.font=RED_FONT
            elif hdr=="Missing":          cell.fill=YELLOW_FILL;cell.font=YELLOW_FONT
            elif hdr=="Filled in actual": cell.fill=FILLED_ACTUAL_FILL
            elif ci<=7:                   cell.fill=stripe
            out.append(cell)
        ws3.append(out)
    ws3.auto_filter.ref=f"A1:{get_column_letter(len(aud_hdrs))}1"

    # ── SHEET 4: Address Legend ────────────────────────────────────────────────
    ws4 = wb.create_sheet("Address Legend")
    ws4.sheet_tab_color = "833C00"
    legend = [
        ("How address matching works","",""),
        ("Step","What happens","Example"),
        ("1. Combine components",
         "primary street + primary city + primary state + primary country + primary zip",
         "Flat 4 Shree Nagar + Andheri West + Mumbai + Maharashtra + 400058"),
        ("2. Normalize","Abbreviations collapsed, punctuation removed, lowercase",
         "f 4 shree ngr andheri w mumbai maharashtra 400058"),
        ("3. Compare vs DB1 slots","Each of Address 1–10 combined and compared",
         "Best matching slot score used"),
        ("4. Compare vs DB2","Home/Permanent combined and compared",""),
        ("5. Score formula","Jaccard(shared tokens / all tokens) + 0.25 if ZIP matches",""),
        ("","",""),
        ("Colour thresholds","",""),
        ("GREEN (✅ MATCH)",f"Score >= {ADDRESS_HIGH_THRESHOLD}","Strong match"),
        ("YELLOW (⚠ PARTIAL)",f"Score >= {ADDRESS_LOW_THRESHOLD}","Partial match"),
        ("RED (❌ MISMATCH)",f"Score < {ADDRESS_LOW_THRESHOLD}","Poor/no match"),
        ("BLANK vs BLANK","Both sides blank = GREEN regardless of score",""),
        ("","",""),
        ("Standard field colour legend (v6.0)","",""),
        ("GREEN (actual col)","Exact match in DB1 or DB2, OR boolean/categorical value, OR unmapped column",""),
        ("FUZZY_GREEN (actual col)","Matched only via fuzzy logic — different shade of green so reviewers can spot fuzzy hits",""),
        ("RED (actual col)","No match found in either DB",""),
        ("TRUE (DB col)","Exact match in this specific DB column",""),
        ("~FUZZY (DB col)","Fuzzy-only match in this specific DB column (teal-green shade)",""),
        ("MISSING (DB col)","This DB has no value for this person",""),
        ("FALSE (DB col)","This DB has a value but it does not match",""),
        ("","",""),
        ("Boolean / categorical → always GREEN","",""),
        ("Values treated as boolean/categorical:",
         "true, false, yes, no, y, n, t, f, alumni, gold, silver, bronze, active, inactive, verified, unverified, 1, 0",
         "These are flag fields with no meaningful DB comparison"),
        ("","",""),
        ("CLI flags to tune thresholds","",""),
        ("--fuzzy-threshold","General fuzzy threshold (default 0.90)","python review.py ... --fuzzy-threshold 0.90"),
        ("--addr-high","GREEN threshold (default 0.90)","python review.py ... --addr-high 0.90"),
        ("--addr-low", "YELLOW threshold (default 0.70)","python review.py ... --addr-low 0.70"),
    ]
    for c,w in zip("ABC",[35,65,55]):
        ws4.column_dimensions[c].width=w
    for ri,row in enumerate(legend,1):
        out=[]
        for ci,val in enumerate(row,1):
            cell=WriteOnlyCell(ws4,value=val)
            cell.font=HEADER_FONT if ri<=2 or (val and not row[1]) else NORMAL_FONT
            cell.fill=ADDR_HDR_FILL if ri<=2 else PatternFill()
            cell.alignment=LEFT
            out.append(cell)
        ws4.append(out)

    # ── Save ──────────────────────────────────────────────────────────────────
    print(f"  Saving -> {out_path}")
    wb.save(out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
#  FINAL MERGED OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

# Colour mapping for address group cells in the merged output
_ADDR_MERGE_STYLE = {
    "HIGH": (GREEN_FILL,  GREEN_FONT),
    "MID":  (YELLOW_FILL, YELLOW_FONT),
    "LOW":  (RED_FILL,    RED_FONT),
}

# Header fill colours for the two trailing summary columns
_RED_CNT_HDR_FILL  = PatternFill("solid", start_color="9C0006")  # dark red
_RED_FLD_HDR_FILL  = PatternFill("solid", start_color="C00000")  # lighter red


def write_final_merged_output(actual_df, blocks, addr_groups, out_path,
                              id_cols=None):
    """
    Write a clean "final_merged_output" Excel file that contains:

    • The actual ID columns (e.g. RE_IDs / AB_IDs) first, uncoloured, so each
      row can be traced back to the source record.

    • The mapped actual data columns (raw actual values with their match-based
      cell colours), PLUS the raw address component columns (Primary_Street …
      Permanent_Zip) exactly as they appear in the actual file, followed by a
      combined "<group> (address)" string column.

    • Two summary columns appended at the END of every row:
        – "Red Cell Count"       — integer: how many mapped columns are RED
                                   (mismatch) for this person/row.
        – "Red Fields (Mismatch)"— comma-separated list of the column names
                                   that are RED, so reviewers can see at a
                                   glance which fields need attention.

    Colour rules (same as the main review sheet):
        GREEN / FUZZY_GREEN → green cell  (exact or fuzzy match)
        YELLOW              → yellow cell (blank in actual)
        RED                 → red cell    (no match found)
        Address HIGH        → green, MID → yellow, LOW → red

    Parameters
    ----------
    actual_df   : original actual DataFrame (used only for row count)
    blocks      : list of standard field result blocks from compute_review()
    addr_groups : list of address group result dicts
    out_path    : full path for the output .xlsx file
    """
    print(f"\n  Building final_merged_output ...")

    n_rows = len(actual_df)

    # ── Collect column data ────────────────────────────────────────────────────
    # Each entry: header name, per-row values, per-row colour codes, is_addr flag.
    # Values are stripped of Excel-illegal control characters here so the .pkl
    # companion file always contains clean data — combine_all_batches() never
    # needs to know about illegal character handling.
    # Order all output columns by their position in the Salesforce input file.
    # Address groups slot in where their first component column sits (so e.g.
    # "primary address" lands at the Primary_Street position, not the far end).
    col_pos  = _input_col_positions(actual_df)

    def _clean_list(values):
        return [_illegal_sub("", v) if isinstance(v, str) and v else v
                for v in values]

    # ID columns lead the sheet (uncoloured). Position before everything else.
    id_cols  = [c for c in (id_cols or []) if c in actual_df.columns]
    id_data  = []
    for k, c in enumerate(id_cols):
        id_data.append({
            "header":  c,
            "values":  _clean_list(actual_df[c].astype(str).tolist()),
            "colors":  [None] * n_rows,
            "is_addr": False,
            "is_id":   True,
        })

    unsorted = []

    for block in blocks:
        unsorted.append((col_pos.get(block["actual_col"], float("inf")), {
            "header":  block["actual_col"],
            "values":  _clean_list(block["actual_vals"].tolist()),
            "colors":  block["actual_color"].tolist(),
            "is_addr": False,
            "is_id":   False,
        }))

    for ag in addr_groups:
        comp_cols   = _addr_group_components(ag, col_pos)
        addr_colors = ag["best_colour"].tolist()
        # Raw component columns, exactly as in the actual file, coloured by the
        # group's address-match result.
        for c in comp_cols:
            unsorted.append((col_pos[c], {
                "header":  c,
                "values":  _clean_list(actual_df[c].astype(str).tolist()),
                "colors":  addr_colors,
                "is_addr": True,
                "is_id":   False,
            }))
        # Combined string column, placed right after the last component.
        combo_pos = (max(col_pos[c] for c in comp_cols) + 0.5
                     if comp_cols else _addr_group_position(ag, col_pos))
        unsorted.append((combo_pos, {
            "header":  f"{ag['label']} (address)",
            "values":  _clean_list(ag["actual_combined"].tolist()),
            "colors":  addr_colors,
            "is_addr": True,
            "is_id":   False,
        }))

    unsorted.sort(key=lambda x: x[0])
    col_data = id_data + [d for _, d in unsorted]

    n_data_cols = len(col_data)

    # ── Per-row red count and red field names ──────────────────────────────────
    red_counts = [0] * n_rows
    red_fields = [[] for _ in range(n_rows)]

    for col_entry in col_data:
        hdr     = col_entry["header"]
        colors  = col_entry["colors"]
        is_addr = col_entry["is_addr"]
        for r in range(n_rows):
            # "LOW" = address mismatch, "RED" = standard field mismatch
            is_red = (colors[r] == "LOW") if is_addr else (colors[r] == "RED")
            if is_red:
                red_counts[r] += 1
                red_fields[r].append(hdr)

    # ── Build workbook ─────────────────────────────────────────────────────────
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Final Merged Output")
    ws.freeze_panes = "A2"

    # Column widths: data columns 22, Red Count 16, Red Fields 50
    total_cols = n_data_cols + 2
    for ci in range(1, total_cols + 1):
        letter = get_column_letter(ci)
        if ci == total_cols:          # Red Fields (last column) — wider
            ws.column_dimensions[letter].width = 55
        elif ci == total_cols - 1:    # Red Cell Count
            ws.column_dimensions[letter].width = 16
        else:                          # Data columns
            ws.column_dimensions[letter].width = 22
    ws.row_dimensions[1].height = 42

    # ── Header row ─────────────────────────────────────────────────────────────
    hdr_row = []

    for col_entry in col_data:
        c = WriteOnlyCell(ws, value=col_entry["header"])
        c.font      = HEADER_FONT
        c.fill      = HEADER_FILL
        c.alignment = WRAP_C
        hdr_row.append(c)

    # "Red Cell Count" header
    c_cnt = WriteOnlyCell(ws, value="Red Cell Count")
    c_cnt.font      = HEADER_FONT
    c_cnt.fill      = _RED_CNT_HDR_FILL
    c_cnt.alignment = CENTER
    hdr_row.append(c_cnt)

    # "Red Fields (Mismatch)" header
    c_fld = WriteOnlyCell(ws, value="Red Fields (Mismatch)")
    c_fld.font      = HEADER_FONT
    c_fld.fill      = _RED_FLD_HDR_FILL
    c_fld.alignment = CENTER
    hdr_row.append(c_fld)

    ws.append(hdr_row)

    # ── Data rows ──────────────────────────────────────────────────────────────
    for r_idx in range(n_rows):
        row = []

        for col_entry in col_data:
            v       = col_entry["values"][r_idx]   # already stripped of illegal chars
            c_code  = col_entry["colors"][r_idx]
            is_addr = col_entry["is_addr"]

            cell = WriteOnlyCell(ws, value=v)
            cell.font = NORMAL_FONT

            if is_addr:
                style = _ADDR_MERGE_STYLE.get(c_code)
            else:
                style = _STD_ACTUAL_STYLE.get(c_code)

            if style:
                cell.fill, cell.font = style

            row.append(cell)

        # ── Red Cell Count cell ────────────────────────────────────────────────
        cnt       = red_counts[r_idx]
        cell_cnt  = WriteOnlyCell(ws, value=cnt)
        cell_cnt.alignment = CENTER
        if cnt == 0:
            cell_cnt.fill = GREEN_FILL
            cell_cnt.font = GREEN_FONT
        elif cnt <= 2:
            cell_cnt.fill = YELLOW_FILL
            cell_cnt.font = YELLOW_FONT
        else:
            cell_cnt.fill = RED_FILL
            cell_cnt.font = RED_FONT
        row.append(cell_cnt)

        # ── Red Fields (Mismatch) cell ─────────────────────────────────────────
        fld_str  = ", ".join(red_fields[r_idx]) if red_fields[r_idx] else "—"
        cell_fld = WriteOnlyCell(ws, value=fld_str)
        cell_fld.font      = RED_FONT  if red_fields[r_idx] else GREEN_FONT
        cell_fld.fill      = LIGHT_ORANGE_FILL if red_fields[r_idx] else LIGHT_GREEN_FILL
        cell_fld.alignment = WRAP_L
        row.append(cell_fld)

        ws.append(row)

    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}1"

    print(f"  Saving -> {out_path}")
    wb.save(out_path)
    print(f"  final_merged_output: {n_data_cols} columns × {n_rows:,} rows written.")

    # ── Save companion data file (.pkl) ────────────────────────────────────────
    # This lightweight binary file stores the raw values + colour codes for
    # every mapped column.  run_all_batches.py reads these files after all
    # batches complete and stitches them into one FINAL_MERGED_ALL_BATCHES.xlsx
    # by merging columns horizontally (all batches share the same rows).
    pkl_path = out_path.replace(".xlsx", ".pkl")
    with open(pkl_path, "wb") as _f:
        pickle.dump({"col_data": col_data, "n_rows": n_rows}, _f,
                    protocol=pickle.HIGHEST_PROTOCOL)
    print(f"  Companion data  -> {pkl_path}")
    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    global FUZZY_THRESHOLD, ADDRESS_HIGH_THRESHOLD, ADDRESS_LOW_THRESHOLD

    # ── Clear all caches first ─────────────────────────────────────────────────
    # Ensures that any corrections to the config or DB column names since the
    # last run are fully picked up. Without this, lru_cache results from a
    # previous run in the same Python process (Jupyter / interactive shell)
    # silently return stale column-match scores, causing the tool to continue
    # mapping against old/wrong DB columns even after you fix the config.
    clear_all_caches()

    parser = argparse.ArgumentParser(
        description="Data Review Tool v6.0 — fuzzy-green distinction + phone/email fuzzy + bool/unmapped green"
    )
    parser.add_argument("--actual",         required=True)
    parser.add_argument("--db1",            required=True)
    parser.add_argument("--db2",            required=True)
    parser.add_argument("--actual-db1-id",  default="db1_id")
    parser.add_argument("--actual-db2-id",  default="db2_id")
    parser.add_argument("--db1-id",         default="id")
    parser.add_argument("--db2-id",         default="id")
    parser.add_argument("--fields",  nargs="*")
    parser.add_argument("--config")
    parser.add_argument("--always-green-config",
                        default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                             "..", "config", "always_green_columns.json"),
                        help="JSON file listing columns to always mark GREEN (no validation).")
    parser.add_argument("--output",  default=None)
    parser.add_argument("--fuzzy-threshold", type=float, default=FUZZY_THRESHOLD)
    parser.add_argument("--addr-high", type=float, default=ADDRESS_HIGH_THRESHOLD,
                        help="Address similarity threshold for GREEN (default 0.90)")
    parser.add_argument("--addr-low",  type=float, default=ADDRESS_LOW_THRESHOLD,
                        help="Address similarity threshold for YELLOW (default 0.70)")
    parser.add_argument("--workers", type=int,
                        default=max(1, (os.cpu_count() or 2) // 2))
    parser.add_argument("--detailed-review", action="store_true",
                        help="Also build the detailed review workbook (the slow file "
                             "that appends every field's db1/db2 + review + diff "
                             "columns). OFF by default because it is the slowest output "
                             "on very large datasets. The lean final_merged + actuals "
                             "files are always produced.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Testing only: evaluate just the first N records of the "
                             "actual file. DB1/DB2 are kept full so lookups still work. "
                             "Use to smoke-test the script fast before a full run.")
    args = parser.parse_args()

    FUZZY_THRESHOLD        = args.fuzzy_threshold
    ADDRESS_HIGH_THRESHOLD = args.addr_high
    ADDRESS_LOW_THRESHOLD  = args.addr_low
    workers                = max(1, args.workers)

    print("=" * 65)
    print("  Data Review Tool v6.0 — 100% Local, Zero Network")
    print(f"  {datetime.now().strftime('%d-%b-%Y')}")
    print(f"  Threads : {workers}")
    print(f"  Fuzzy threshold : {FUZZY_THRESHOLD}")
    print(f"  Addr thresholds: HIGH={ADDRESS_HIGH_THRESHOLD}  LOW={ADDRESS_LOW_THRESHOLD}")
    print("=" * 65)

    _security_guard()

    # ── Load files ─────────────────────────────────────────────────────────────
    print(f"\n[1] Loading files ...")
    with ThreadPoolExecutor(max_workers=min(3, workers)) as ex:
        fa = ex.submit(load_excel, args.actual)
        f1 = ex.submit(load_excel, args.db1)
        f2 = ex.submit(load_excel, args.db2)
        actual_df = fa.result()
        db1_df    = f1.result()
        db2_df    = f2.result()

    # ── Testing limit: keep only first N actual records (DBs stay full) ─────────
    if args.limit is not None and args.limit > 0 and len(actual_df) > args.limit:
        print(f"  [LIMIT] Testing mode — evaluating first {args.limit:,} of "
              f"{len(actual_df):,} actual records.")
        actual_df = actual_df.head(args.limit).reset_index(drop=True)
    gc.collect()

    # ── Detect ID columns ──────────────────────────────────────────────────────
    print("\n[2] Detecting ID columns ...")
    actual_db1_id = detect_id_column(actual_df, args.actual_db1_id)
    actual_db2_id = detect_id_column(actual_df, args.actual_db2_id)
    db1_id_col    = detect_id_column(db1_df,    args.db1_id)
    db2_id_col    = detect_id_column(db2_df,    args.db2_id)

    actual_db1_lookup = actual_df[actual_db1_id].astype(str).str.strip()
    actual_db2_lookup = actual_df[actual_db2_id].astype(str).str.strip()
    db1_key = db1_df[db1_id_col].astype(str).str.strip()
    db2_key = db2_df[db2_id_col].astype(str).str.strip()

    # ── Build plan ─────────────────────────────────────────────────────────────
    fields      = args.fields
    manual_plan = None

    if args.config:
        config_abs = os.path.abspath(args.config)
        print(f"\n[3] Loading config: {config_abs}")
        with open(config_abs, encoding="utf-8-sig") as f:   # utf-8-sig handles BOM
            cfg = json.load(f)
        raw = cfg.get("fields", [])

        # ── CONFIG IS THE SOLE AUTHORITY ───────────────────────────────────────
        # When a config file is supplied, ALL field→DB column mappings come
        # exclusively from it.  Auto-discovery is NEVER used for standard fields.
        #
        # Rules per entry:
        #   "db1_cols": ["X"]  → map against column X in DB1
        #   "db1_cols": []     → no DB1 mapping for this field
        #   "db1_cols" absent  → treated the same as [] (no mapping)

        # Build case-insensitive lookup maps so a column typed as
        # "ABC xyz" in the config matches "ABC XYZ" in the Excel file,
        # and so invisible whitespace differences never cause silent misses.
        db1_col_map = {c.strip().lower(): c for c in db1_df.columns}
        db2_col_map = {c.strip().lower(): c for c in db2_df.columns}
        # Exact column-name sets, checked BEFORE the case-insensitive map so a
        # config column that exists verbatim in the DB always resolves to itself
        # (avoids case-collision picking the wrong duplicate-named column).
        db1_set = set(db1_df.columns)
        db2_set = set(db2_df.columns)

        manual_plan  = []
        skipped_cols: list[str] = []

        for entry in raw:
            if not isinstance(entry, dict):
                continue
            canonical     = entry.get("canonical", "")
            actual_cols   = entry.get("actual_cols", [])
            # Strip whitespace from every configured column name so a
            # trailing space typed in the JSON never causes a silent miss.
            db1_cols_cfg  = [c.strip() for c in entry.get("db1_cols", [])]
            db2_cols_cfg  = [c.strip() for c in entry.get("db2_cols", [])]
            norm_override = entry.get("norm_type", None)

            for actual_col in actual_cols:
                if actual_col not in actual_df.columns:
                    skipped_cols.append(actual_col)
                    continue

                # Resolve each configured column name against the DB dataframe
                # using case-insensitive, whitespace-stripped matching.
                # This means "ABC xyz" in the config will correctly resolve
                # to "ABC XYZ" as it appears in the DB Excel file.
                db1_cols, db2_cols = [], []

                for c in db1_cols_cfg:
                    # Exact-case match wins over case-insensitive fallback.
                    # Some DB files contain columns that differ only by case
                    # (e.g. "Class year (1)" vs "Class Year (1)"); the lower()
                    # map would otherwise collide and silently resolve to the
                    # wrong (often near-empty) column.
                    resolved = c if c in db1_set else db1_col_map.get(c.lower())
                    if resolved:
                        db1_cols.append(resolved)
                    else:
                        print(f"  ⚠  DB1 column NOT FOUND: '{c}' "
                              f"(field '{canonical}') — check the column name "
                              f"in your config matches exactly what is in DB1.")

                for c in db2_cols_cfg:
                    resolved = c if c in db2_set else db2_col_map.get(c.lower())
                    if resolved:
                        db2_cols.append(resolved)
                    else:
                        print(f"  ⚠  DB2 column NOT FOUND: '{c}' "
                              f"(field '{canonical}') — check the column name "
                              f"in your config matches exactly what is in DB2.")

                norm_type = (norm_override if norm_override
                             else detect_norm_type(actual_col, actual_df[actual_col]))

                manual_plan.append({
                    "canonical":     canonical,
                    "actual_col":    actual_col,
                    "db1_cols":      db1_cols,
                    "db2_cols":      db2_cols,
                    "norm_type":     norm_type,
                    "is_addr_group": False,
                })

        if skipped_cols:
            print(f"  ⚠  Skipped {len(skipped_cols)} actual column(s) not found "
                  f"in actual file: {skipped_cols}")

        # ── Print mapping plan table ───────────────────────────────────────────
        print(f"\n  {'Actual Column':<35} {'DB1 Columns':<35} {'DB2 Columns':<35}")
        print("  " + "─" * 105)
        for mp in manual_plan:
            db1_str = ", ".join(mp["db1_cols"]) if mp["db1_cols"] else "(none)"
            db2_str = ", ".join(mp["db2_cols"]) if mp["db2_cols"] else "(none)"
            print(f"  {mp['actual_col']:<35} {db1_str:<35} {db2_str:<35}")
        print("  " + "─" * 105)
        print(f"  Total: {len(manual_plan)} mapping(s) | "
              f"config-only mode: auto-discovery disabled")
        print("  ↑ If any field shows wrong DB column, update THAT batch config file.\n")

    plan = (manual_plan if manual_plan is not None
            else build_field_plan(actual_df, db1_df, db2_df, fields))

    # ── Remove individual address component fields from plan ──────────────────
    # They will be handled by the address group computation instead.
    all_addr_components = set(ACTUAL_PRIMARY_COMPONENTS + ACTUAL_PERMANENT_COMPONENTS)
    plan_filtered = [p for p in plan
                     if p["actual_col"].lower() not in
                     {c.lower() for c in all_addr_components}]

    skipped = len(plan) - len(plan_filtered)
    if skipped:
        print(f"\n  [INFO] {skipped} individual address component columns removed from "
              f"standard plan — handled by address group computation instead.")

    # ── Always-green columns: load + resolve canonicals → actual_cols ─────────
    always_green_cols: set = set()
    agc_path = args.always_green_config
    if agc_path and os.path.exists(agc_path):
        try:
            with open(agc_path) as f:
                agc = json.load(f)
            raw_list = agc.get("columns", []) if isinstance(agc, dict) else agc
            # Build canonical → actual_cols map from --config (if any)
            canonical_map = {}
            if args.config:
                try:
                    with open(args.config) as f:
                        _cfg = json.load(f)
                    for entry in _cfg.get("fields", []):
                        if isinstance(entry, dict):
                            canonical_map[entry.get("canonical", "").strip().lower()] = \
                                entry.get("actual_cols", [])
                except Exception:
                    pass
            for item in raw_list:
                if not isinstance(item, str) or not item.strip():
                    continue
                key = item.strip()
                if key in actual_df.columns:
                    always_green_cols.add(key)
                elif key.lower() in canonical_map:
                    for ac in canonical_map[key.lower()]:
                        if ac in actual_df.columns:
                            always_green_cols.add(ac)
                else:
                    print(f"  [WARN] always-green entry not found in actual file: {key!r}")
            if always_green_cols:
                print(f"\n[3b] Always-green: {len(always_green_cols)} columns will skip validation.")
        except Exception as e:
            print(f"  [WARN] Failed to load always-green config {agc_path}: {e}")

    # ── Standard field computation ─────────────────────────────────────────────
    blocks = compute_review(
        actual_df, db1_df, db2_df,
        actual_db1_id, actual_db2_id,
        db1_id_col, db2_id_col,
        plan_filtered, workers=workers,
        always_green_cols=always_green_cols,
    )

    # ── Address group computation ──────────────────────────────────────────────
    print("\n[4b] Computing address groups ...")

    db1_slot_col_lists = [
        list(db1_slot_components(s).values()) for s in DB1_ADDR_SLOTS
    ]

    addr_groups = []

    # Primary address group
    primary_cols_present = [c for c in ACTUAL_PRIMARY_COMPONENTS if c in actual_df.columns]
    if primary_cols_present:
        print("  Computing: primary address group ...")
        ag_primary = compute_address_group(
            actual_df, db1_df, db2_df,
            ACTUAL_PRIMARY_COMPONENTS,
            db1_slot_col_lists,
            DB2_HOME_COMPONENTS,
            {"db1": actual_db1_lookup, "db2": actual_db2_lookup},
            db1_key, db2_key,
            label="primary",
            db2_extra_component_cols=DB2_PERMANENT_COMPONENTS,
        )
        addr_groups.append(ag_primary)
        print(f"    DB1 match%={ag_primary['stats_db1']['match_pct']}%  "
              f"DB2 match%={ag_primary['stats_db2']['match_pct']}%  "
              f"filled={ag_primary['filled_in_actual']:,}")

    # Permanent address group
    permanent_cols_present = [c for c in ACTUAL_PERMANENT_COMPONENTS if c in actual_df.columns]
    if permanent_cols_present:
        print("  Computing: permanent address group ...")
        ag_permanent = compute_address_group(
            actual_df, db1_df, db2_df,
            ACTUAL_PERMANENT_COMPONENTS,
            db1_slot_col_lists,
            DB2_PERMANENT_COMPONENTS,
            {"db1": actual_db1_lookup, "db2": actual_db2_lookup},
            db1_key, db2_key,
            label="permanent"
        )
        addr_groups.append(ag_permanent)
        print(f"    DB1 match%={ag_permanent['stats_db1']['match_pct']}%  "
              f"DB2 match%={ag_permanent['stats_db2']['match_pct']}%  "
              f"filled={ag_permanent['filled_in_actual']:,}")

    gc.collect()

    # ── Output path ────────────────────────────────────────────────────────────
    if not args.output:
        ts = datetime.now().strftime("%d-%b-%Y") 
        args.output = f"review_output_{ts}.xlsx"

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)

    # ── Write main (detailed) review workbook ──────────────────────────────────
    # This is the slowest output: per field it appends db1/db2 value columns,
    # review columns and diff columns. Skipped unless --detailed-review is given.
    if args.detailed_review:
        _make_workbook_and_review_sheet(
            actual_df, blocks, addr_groups,
            args.output, workers
        )
    else:
        print("\n[5] Detailed review workbook SKIPPED "
              "(pass --detailed-review to build it).")

    # ── Write final merged output ──────────────────────────────────────────────
    # Name is derived from the main output filename so run_all_batches.py can
    # locate each batch's merged file reliably (no timestamp guessing).
    # e.g.  review_batch_001.xlsx  →  final_merged_batch_001.xlsx
    out_dir     = os.path.dirname(os.path.abspath(args.output))
    main_stem   = os.path.splitext(os.path.basename(args.output))[0]
    # Strip leading "review_" prefix if present so the merged name is clean
    merged_stem = main_stem[len("review_"):] if main_stem.startswith("review_") else main_stem
    merged_path = os.path.join(out_dir, f"final_merged_{merged_stem}.xlsx")
    write_final_merged_output(actual_df, blocks, addr_groups, merged_path,
                              id_cols=[actual_db1_id, actual_db2_id])

    # ── Write compact "actuals only" output ───────────────────────────────────
    base, ext = os.path.splitext(args.output)
    actuals_path = f"{base}_actuals{ext or '.xlsx'}"
    _make_actuals_only_workbook(actual_df, blocks, addr_groups, actuals_path, workers)

    # ── Final summary ──────────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  FINAL SUMMARY")
    print("=" * 65)
    print(f"  {'Field':<40} {'Type':<9} {'Filled':>7} {'Match%':>8}")
    print("  " + "-" * 65)
    for block in blocks:
        print(f"  {block['actual_col']:<40} {block['norm_type']:<9} "
              f"{block['filled_in_actual']:>7,} {block['overall_pct']:>7.1f}%")
    for ag in addr_groups:
        pct = ag["stats_db1"]["match_pct"]
        print(f"  {ag['label']:<40} {'address':<9} "
              f"{ag['filled_in_actual']:>7,} {pct:>7.1f}%  (DB1, any slot)")

    if args.detailed_review:
        print(f"\n  Review output  -> {args.output}")
    else:
        print(f"\n  Review output  -> (skipped — use --detailed-review)")
    print(f"  Merged output  -> {merged_path}")
    print(f"  Compact output -> {actuals_path}")
    if args.detailed_review:
        print("  Sheets (review)-> Review | Summary | Field Audit | Address Legend")
    print("  Merged sheet   -> Final Merged Output")
    print("  Compact sheet  -> Actuals Review")
    print("=" * 65)


if __name__ == "__main__":
    main()
