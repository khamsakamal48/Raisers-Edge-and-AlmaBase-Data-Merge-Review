"""
Column Discovery Tool  v3.0
============================
Run this BEFORE review.py to preview how your 750+ columns will be grouped
and which DB1/DB2 columns will be matched to each field.

What it does:
  - Reads ONLY the header row of all files (zero data rows — fast and safe)
  - Groups similar column names using fuzzy matching
    e.g. firstname1.1, firstname1.2, first_name  =>  canonical: firstname
  - Shows which DB1 and DB2 columns match each actual field, with match score
  - Auto-detects normalization type: phone / email / url / date / number / text
  - Saves a fields_config.json you can edit to fix any wrong mappings
  - Pass that config to review.py --config to run the full review

Security:
  - Reads only column headers (row 0), zero data loaded
  - No network calls — purely offline string operations
  - Same security guard as review.py

Usage:
    python discover_columns.py --actual actual.xlsx --db1 db1_dump.xlsx --db2 db2_dump.xlsx

    python discover_columns.py --actual actual.xlsx --db1 db1.xlsx --db2 db2.xlsx ^
        --output ..\\config\\fields_config.json

    python discover_columns.py --actual actual.xlsx --db1 education_dump.xlsx --db2 db2.xlsx ^
        --fuzzy-threshold 0.70 --output ..\\config\\education_config.json
"""

import argparse
import json
import os
import re
import sys
from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd

try:
    from rapidfuzz import process as _rf_process, fuzz as _rf_fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False


# ═══════════════════════════════════════════════════════════════════════════════
#  SECURITY GUARD  (identical to review.py)
# ═══════════════════════════════════════════════════════════════════════════════
_BANNED_MODULES = {
    "requests", "urllib3", "httpx", "aiohttp", "http.client",
     "ftplib", "smtplib", "paramiko", "boto3",
    "google.cloud", "azure", "dropbox", "onedrive",
}

def _security_guard():
    loaded = set(sys.modules.keys())
    hits   = loaded & _BANNED_MODULES
    if hits:
        raise RuntimeError(
            f"SECURITY VIOLATION: Network-capable module(s) detected: {hits}\n"
            "This script must run offline. Aborting."
        )

_security_guard()


# ═══════════════════════════════════════════════════════════════════════════════
#  NORMALIZATION TYPE DETECTION  (column-name based, no data needed)
# ═══════════════════════════════════════════════════════════════════════════════
_PHONE_KEYS  = {"phone","mobile","cell","contact","tel","telephone","ph","mob","landline"}
_EMAIL_KEYS  = {"email","e-mail","mail","emailid","emailaddress"}
_URL_KEYS    = {"url","website","web","site","link","http","www","homepage","portal"}
_DATE_KEYS   = {"date","dob","birth","doj","joining","expiry","expiration","anniversary",
                 "from","to","since","until","start","end"}
_NUMBER_KEYS = {"amount","salary","income","balance","score","count","number","no","num",
                 "code","pin","pincode","zip","age","year","rate","percent","pct"}

def detect_norm_type(col_name: str) -> str:
    words = set(re.findall(r'[a-z]+', col_name.lower()))
    if words & _EMAIL_KEYS:  return "email"
    if words & _URL_KEYS:    return "url"
    if words & _PHONE_KEYS:  return "phone"
    if words & _DATE_KEYS:   return "date"
    if words & _NUMBER_KEYS: return "number"
    return "text"


# ═══════════════════════════════════════════════════════════════════════════════
#  FUZZY COLUMN MATCHING  (identical logic to review.py)
# ═══════════════════════════════════════════════════════════════════════════════

@lru_cache(maxsize=None)
def normalize_col_name(name: str) -> str:
    """Lowercase, collapse separators, strip trailing version numbers."""
    n = str(name).lower().strip()
    n = re.sub(r'[\s_\-\.]+', ' ', n)
    n = re.sub(r'(\s*\d+(\.\d+)*)+$', '', n).strip()
    return n


@lru_cache(maxsize=None)
def _seq_ratio(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def col_similarity(a: str, b: str) -> float:
    return _seq_ratio(normalize_col_name(a), normalize_col_name(b))


def find_matching_columns(target: str, columns: list, threshold: float) -> list:
    """Return (col_name, score) pairs for columns that fuzzy-match target."""
    t_norm = normalize_col_name(target)
    results = [(c, round(_seq_ratio(t_norm, normalize_col_name(c)), 3)) for c in columns]
    results = [(c, s) for c, s in results if s >= threshold]
    results.sort(key=lambda x: -x[1])
    return results


def group_variants(columns: list, threshold: float) -> dict:
    """
    Cluster column list into canonical groups.
    e.g. ['firstname1.1', 'firstname1.2', 'first_name']
         => {'firstname': ['firstname1.1', 'firstname1.2', 'first_name']}
    """
    used   = set()
    groups = {}
    norm_map = {c: normalize_col_name(c) for c in columns}
    for col in columns:
        if col in used:
            continue
        norm = norm_map[col]
        if norm not in groups:
            groups[norm] = []
        groups[norm].append(col)
        used.add(col)
        col_norm = norm_map[col]
        for other in columns:
            if other in used:
                continue
            if _seq_ratio(col_norm, norm_map[other]) >= threshold:
                groups[norm].append(other)
                used.add(other)
    return groups


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════════

def discover(actual_path, db1_path, db2_path, threshold, output_json):

    print("=" * 70)
    print("  Column Discovery Tool v3.0 — 100% Local, Zero Network")
    print("=" * 70)

    # ── Read headers only (nrows=0 = no data rows loaded) ──
    print("\nReading column headers only (no data loaded) ...")

    def _read_header(p):
        try:
            return pd.read_excel(p, nrows=0, engine="calamine").columns.tolist()
        except Exception:
            return pd.read_excel(p, nrows=0, engine="openpyxl").columns.tolist()

    actual_cols = _read_header(actual_path)
    db1_cols    = _read_header(db1_path)
    db2_cols    = _read_header(db2_path)

    actual_cols = [str(c).strip() for c in actual_cols]
    db1_cols    = [str(c).strip() for c in db1_cols]
    db2_cols    = [str(c).strip() for c in db2_cols]

    print(f"\n  Actual file : {len(actual_cols):>4} columns  ({os.path.basename(actual_path)})")
    print(f"  DB1 dump    : {len(db1_cols):>4} columns  ({os.path.basename(db1_path)})")
    print(f"  DB2 dump    : {len(db2_cols):>4} columns  ({os.path.basename(db2_path)})")
    print(f"  Threshold   : {threshold}")

    # ── Group actual columns into canonical clusters ──
    print(f"\nGrouping {len(actual_cols)} actual columns into canonical fields ...")
    actual_groups = group_variants(actual_cols, threshold)
    print(f"  -> {len(actual_groups)} canonical field groups found")

    # ── Print discovery table ──
    print()
    print(f"{'#':<5} {'CANONICAL FIELD':<32} {'NORM TYPE':<10} "
          f"{'ACTUAL COLS':<35} {'DB1 MATCHES (score)':<38} {'DB2 MATCHES (score)'}")
    print("-" * 160)

    output = {
        "fuzzy_threshold": threshold,
        "generated":       str(pd.Timestamp.now())[:19],
        "actual_file":     os.path.basename(actual_path),
        "db1_file":        os.path.basename(db1_path),
        "db2_file":        os.path.basename(db2_path),
        "fields":          []
    }

    warnings = []

    for idx, (canonical, a_cols) in enumerate(sorted(actual_groups.items()), 1):
        db1_matches = find_matching_columns(canonical, db1_cols, threshold)
        db2_matches = find_matching_columns(canonical, db2_cols, threshold)

        norm_type = detect_norm_type(canonical)

        # Terminal display (truncated for readability)
        a_str   = ", ".join(a_cols[:4]) + (" ..." if len(a_cols) > 4 else "")
        db1_str = ", ".join(f"{c}({s})" for c, s in db1_matches[:3]) or "— NOT FOUND —"
        db2_str = ", ".join(f"{c}({s})" for c, s in db2_matches[:3]) or "— NOT FOUND —"

        flag = ""
        if not db1_matches and not db2_matches:
            flag = " <-- NOT IN EITHER DB"
            warnings.append(f"  [{idx:>3}] '{canonical}' — not found in DB1 or DB2")
        elif not db1_matches:
            flag = " <-- missing from DB1"
        elif not db2_matches:
            flag = " <-- missing from DB2"

        print(f"{idx:<5} {canonical:<32} {norm_type:<10} "
              f"{a_str:<35} {db1_str:<38} {db2_str}{flag}")

        output["fields"].append({
            "canonical":    canonical,
            "norm_type":    norm_type,
            "actual_cols":  a_cols,
            "db1_cols":     [c for c, _ in db1_matches],
            "db2_cols":     [c for c, _ in db2_matches],
            "db1_scores":   {c: s for c, s in db1_matches},
            "db2_scores":   {c: s for c, s in db2_matches},
            "note":         ""   # ← edit this manually to add override notes
        })

    # ── Save config JSON ──
    os.makedirs(os.path.dirname(os.path.abspath(output_json)), exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    # ── Print warnings and summary ──
    print()
    print("=" * 70)
    print(f"  DISCOVERY SUMMARY")
    print("=" * 70)
    total       = len(output["fields"])
    found_both  = sum(1 for f in output["fields"] if f["db1_cols"] and f["db2_cols"])
    found_db1   = sum(1 for f in output["fields"] if f["db1_cols"] and not f["db2_cols"])
    found_db2   = sum(1 for f in output["fields"] if not f["db1_cols"] and f["db2_cols"])
    found_none  = sum(1 for f in output["fields"] if not f["db1_cols"] and not f["db2_cols"])

    print(f"  Total canonical fields  : {total}")
    print(f"  Found in BOTH DB1 & DB2 : {found_both}")
    print(f"  Found in DB1 only       : {found_db1}")
    print(f"  Found in DB2 only       : {found_db2}")
    print(f"  NOT found in either DB  : {found_none}  {'<-- review these manually' if found_none else ''}")

    if warnings:
        print()
        print("  Fields NOT found in either DB dump (check resource_location.xlsx):")
        for w in warnings:
            print(w)

    print()
    print(f"  Config saved to : {output_json}")
    print()
    print("  NEXT STEPS:")
    print(f"  1. Open {output_json} in Notepad and check the mappings look correct.")
    print("  2. For any wrong mapping, edit the 'db1_cols' or 'db2_cols' list manually.")
    print("  3. Use your resource_location.xlsx to confirm the right column names.")
    print("  4. Then run review.py with:")
    print(f"       python review.py --actual ... --db1 ... --db2 ... --config {output_json} --output ..\\output\\review.xlsx")
    print("=" * 70)


# ═══════════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Column Discovery Tool v3.0 — previews column groupings before running review.py"
    )
    parser.add_argument("--actual",          required=True,  help="Path to actual data Excel file")
    parser.add_argument("--db1",             required=True,  help="Path to DB1 dump Excel file")
    parser.add_argument("--db2",             required=True,  help="Path to DB2 dump Excel file")
    parser.add_argument("--fuzzy-threshold", type=float, default=0.80,
                        help="Column name similarity threshold 0-1 (default 0.80). "
                             "Lower = more lenient, Higher = stricter.")
    parser.add_argument("--output",          default="fields_config.json",
                        help="Where to save the JSON config (default: fields_config.json)")
    args = parser.parse_args()

    discover(
        actual_path  = args.actual,
        db1_path     = args.db1,
        db2_path     = args.db2,
        threshold    = args.fuzzy_threshold,
        output_json  = args.output,
    )
