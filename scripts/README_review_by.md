# Data Review Tool — `review_by.py`

Validates an "actual" data file against two reference dumps (DB1, DB2), row by
row, field by field. Produces colour-coded Excel review workbooks showing which
values match, fuzzy-match, are missing, or mismatch.

This is the merged script: your original `review.py` + your colleague's changes
(`review - Nirbhaya.py`), with all features from both combined.

---

## Quick start

```bash
python3 review_by.py \
  --actual ../input/actual_data.xlsx \
  --db1    ../input/db1_dump.xlsx \
  --db2    ../input/db2_dump.xlsx \
  --config ../config/fields_config.json \
  --output ../output/review_output.xlsx \
  --workers 4
```

Run from the `scripts/` directory. Paths can be relative or absolute.

---

## Arguments

### Required

| Argument | Description |
|----------|-------------|
| `--actual PATH` | Actual data file to review (the file being checked). |
| `--db1 PATH`    | First reference dump (DB1). |
| `--db2 PATH`    | Second reference dump (DB2). |

Each file may be `.xlsx`, `.xls`, **`.csv`, `.tsv`, or `.txt`**. CSV/TSV are read
as text with a UTF-8 → Latin-1 fallback. Excel is read via `calamine`, falling
back to `openpyxl`.

### ID columns (how rows in actual link to the dumps)

| Argument | Default | Description |
|----------|---------|-------------|
| `--actual-db1-id NAME` | `db1_id` | Column in **actual** holding the DB1 key. |
| `--actual-db2-id NAME` | `db2_id` | Column in **actual** holding the DB2 key. |
| `--db1-id NAME` | `id` | Key column in **DB1**. |
| `--db2-id NAME` | `id` | Key column in **DB2**. |

Names are matched fuzzily (≥0.65), so close header names still resolve. The
resolved name is printed in step `[2]`. A cell with multiple IDs (e.g.
`"123,456"`) is split and each ID is tried until a value is found.

### Field mapping

| Argument | Default | Description |
|----------|---------|-------------|
| `--config PATH` | _(none)_ | JSON config mapping fields → DB columns. **When supplied, it is the sole authority** — auto-discovery is disabled and only configured mappings are used. See [Config format](#config-format). |
| `--fields A B ...` | _(none)_ | Only used when **no** `--config` is given: restricts auto-discovery to these field names. |
| `--always-green-config PATH` | `../config/always_green_columns.json` | JSON listing actual columns to force GREEN with **no validation** (skips DB lookup / fuzzy / combined match). Silently ignored if the file does not exist. See [format](#always-green-format). |

### Match thresholds

| Argument | Default | Description |
|----------|---------|-------------|
| `--fuzzy-threshold F` | `0.90` | Similarity (0–1) required for a fuzzy text match. |
| `--addr-high F` | `0.90` | Address similarity ≥ this → GREEN. |
| `--addr-low F`  | `0.70` | Address similarity ≥ this → YELLOW; below → RED. |

### Output & performance

| Argument | Default | Description |
|----------|---------|-------------|
| `--output PATH` | `review_output_<DD-Mon-YYYY>.xlsx` | Main review workbook path. The other two output files are named from this (see [Outputs](#outputs)). |
| `--workers N` | half of CPU cores | Worker threads for lookup + workbook writing. |

---

## Outputs

Three files are written (names derived from `--output`, e.g. `review_output.xlsx`):

1. **Main review** — `review_output.xlsx`
   Sheets: `Review` | `Summary` | `Field Audit` | `Address Legend`.
   Full detail: actual value, each DB value, per-DB review, diffs, match %.

2. **Final merged** — `final_merged_output.xlsx` (+ `final_merged_output.pkl`)
   Only mapped actual columns with match colours, plus two trailing columns:
   `Red Cell Count` and `Red Fields (Mismatch)` per row. The `.pkl` companion
   lets a batch runner stitch many batches into one file.
   *(Name strips a leading `review_` prefix: `review_batch_001.xlsx` → `final_merged_batch_001.xlsx`.)*

3. **Compact actuals** — `review_output_actuals.xlsx`
   Sheet `Actuals Review`: IDs + each actual column coloured, plus per-row
   `Red cells` and `Yellow cells` counts.

---

## Colour key

| Colour | Meaning |
|--------|---------|
| GREEN | Exact match, combined/contains match, blank actual, or boolean/unmapped column. |
| FUZZY_GREEN (medium green) | Matched only by fuzzy similarity, no exact match. |
| YELLOW | Value present in actual but missing in the DB (`MISSING`). |
| RED | No match found in either DB. |
| Address: HIGH→green, MID→yellow, LOW→red | Per address-similarity thresholds. |

**Match order per field:** (1) exact → (2) fuzzy → (3) combined/contains
(substring + encoding-cleaned substring + token-overlap, for text fields only;
dates/numbers excluded). Combined matches count as GREEN, same as exact.

---

## Config format

`--config` JSON. Only the `fields` array matters at runtime; other keys are
metadata. Each entry:

```json
{
  "fields": [
    {
      "canonical": "account owner",
      "norm_type": "text",
      "actual_cols": ["Account_Owner"],
      "db1_cols": ["Account Owner"],
      "db2_cols": []
    }
  ]
}
```

- `actual_cols` — column header(s) in the actual file to review.
- `db1_cols` / `db2_cols` — columns to compare against. `[]` or absent = no
  mapping for that DB. Names are resolved **case-insensitively** and
  whitespace-trimmed, so `"abc xyz"` matches `"ABC XYZ"`. Unfound columns print
  a `⚠ DB1/DB2 column NOT FOUND` warning.
- `norm_type` — optional override: `text`, `date`, `number`, `phone`, `email`,
  `url`, `address`. Auto-detected if omitted.

The mapping plan is printed as a table at startup — check it before trusting
the run.

<a name="always-green-format"></a>
## Always-green format

`always_green_columns.json`:

```json
{
  "description": "...",
  "columns": ["Permanent_Verification_Source", "account owner"]
}
```

Entries are either exact actual-file headers **or** canonical names from the
config (which expand to their `actual_cols`). Listed columns are forced GREEN
and skip all validation.

---

## Notes / gotchas

- All `lru_cache`s are cleared at startup, so config/column edits always take
  effect on the next run (no stale cached matches in a long-lived shell).
- Config is loaded as `utf-8-sig`, so a BOM in the JSON is handled.
- 100% local — no network calls (enforced by `_security_guard()`).
- Always-green columns never appear RED in the merged file's red counts,
  because they are not validated.

---

## Examples

Auto-discover fields (no config):
```bash
python3 review_by.py --actual a.xlsx --db1 d1.xlsx --db2 d2.xlsx --output out.xlsx
```

Config-driven, custom thresholds, 8 threads:
```bash
python3 review_by.py \
  --actual ../input/actual_data.xlsx --db1 ../input/db1_dump.xlsx --db2 ../input/db2_dump.xlsx \
  --config ../config/fields_config.json \
  --fuzzy-threshold 0.85 --addr-high 0.92 --addr-low 0.72 \
  --workers 8 --output ../output/review_output.xlsx
```

CSV inputs, custom ID columns:
```bash
python3 review_by.py \
  --actual data.csv --db1 dump1.tsv --db2 dump2.csv \
  --actual-db1-id member_id --db1-id id \
  --actual-db2-id alt_id --db2-id id \
  --output out.xlsx
```
