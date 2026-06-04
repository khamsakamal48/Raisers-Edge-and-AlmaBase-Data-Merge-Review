# Data Review — scripts overview

This folder has **four** Python files. Only one is the script you run for a full
review. This page tells you which is which, **which one to run**, and the
arguments each takes.

---

## TL;DR — which script do I run?

> **Run `review_by.py`.** It is the current, merged script.

| File | What it is | Run it? |
|------|------------|---------|
| **`review_by.py`** | **The current script.** Merge of your `review.py` + Nirbhaya's `review - Nirbhaya.py`, with all features from both. | ✅ **Yes — this is the one.** |
| `review.py` | Your previous version. Kept for reference / comparison. | ⚠️ Old — superseded by `review_by.py`. |
| `review - Nirbhaya.py` | Nirbhaya's version (the one you were given). Kept for reference. | ⚠️ Old — superseded by `review_by.py`. |
| `discover_columns.py` | Helper. Previews column groupings and **generates** a starter `fields_config.json`. Run it *before* the review only if you need to build/refresh the config. | Optional helper. |

`review_by.py` has its own detailed guide in **`README_review_by.md`** (config
format, outputs, colour key, gotchas). This page is the short overview.

---

## `review_by.py` — arguments

Run from the `scripts/` directory. Paths can be relative or absolute.

```bash
python3 review_by.py \
  --actual ../input/actual_data.xlsx \
  --db1    ../input/db1_dump.xlsx \
  --db2    ../input/db2_dump.xlsx \
  --config ../config/fields_config.json \
  --output ../output/review_output.xlsx \
  --workers 4
```

### Required

| Argument | Description |
|----------|-------------|
| `--actual PATH` | Actual data file being reviewed. |
| `--db1 PATH` | First reference dump (DB1). |
| `--db2 PATH` | Second reference dump (DB2). |

Files may be `.xlsx`, `.xls`, `.csv`, `.tsv`, or `.txt`.

### ID columns (link actual rows to the dumps)

| Argument | Default | Description |
|----------|---------|-------------|
| `--actual-db1-id NAME` | `db1_id` | Column in **actual** holding the DB1 key. |
| `--actual-db2-id NAME` | `db2_id` | Column in **actual** holding the DB2 key. |
| `--db1-id NAME` | `id` | Key column in **DB1**. |
| `--db2-id NAME` | `id` | Key column in **DB2**. |

Names match fuzzily, so close header names still resolve. Resolved name printed
in step `[2]`.

### Field mapping

| Argument | Default | Description |
|----------|---------|-------------|
| `--config PATH` | _(none)_ | JSON mapping fields → DB columns. **When supplied it is the sole authority** — auto-discovery disabled. |
| `--fields A B ...` | _(none)_ | Only used when **no** `--config`: restricts auto-discovery to these fields. |
| `--always-green-config PATH` | `../config/always_green_columns.json` | Columns forced GREEN with no validation. Ignored if file missing. |

### Match thresholds

| Argument | Default | Description |
|----------|---------|-------------|
| `--fuzzy-threshold F` | `0.90` | Similarity (0–1) for a fuzzy text match. |
| `--addr-high F` | `0.90` | Address similarity ≥ this → GREEN. |
| `--addr-low F` | `0.70` | Address similarity ≥ this → YELLOW; below → RED. |

### Output & performance

| Argument | Default | Description |
|----------|---------|-------------|
| `--output PATH` | `review_output_<DD-Mon-YYYY>.xlsx` | Main workbook. Other two outputs derive their names from this. |
| `--workers N` | half of CPU cores | Worker threads. |

> `review.py` and `review - Nirbhaya.py` take the **same arguments** (except the
> two old versions may lack `--always-green-config`). You don't need to run them.

---

## `discover_columns.py` — arguments (optional helper)

Generates a starter `fields_config.json` by guessing actual↔DB column groupings.
Run only when building or refreshing the config.

```bash
python3 discover_columns.py \
  --actual ../input/actual_data.xlsx \
  --db1    ../input/db1_dump.xlsx \
  --db2    ../input/db2_dump.xlsx \
  --output ../config/fields_config.json
```

| Argument | Required | Default | Description |
|----------|----------|---------|-------------|
| `--actual PATH` | yes | — | Actual data file. |
| `--db1 PATH` | yes | — | DB1 dump. |
| `--db2 PATH` | yes | — | DB2 dump. |
| `--fuzzy-threshold F` | no | `0.80` | Column-name similarity (lower = more lenient). |
| `--output PATH` | no | `fields_config.json` | Where to write the generated config. |

**Always review the generated config by hand** before feeding it to
`review_by.py` — auto-grouping is a starting point, not final truth.

---

## Typical workflow

1. *(optional, once)* `discover_columns.py` → generate `fields_config.json`, then
   hand-correct it.
2. `review_by.py` with `--config ../config/fields_config.json` → produces the
   colour-coded review workbooks in `output/`.

See `README_review_by.md` for config format, output files, and colour meanings.
