# Monthly data update

The analyzer ships its dataset embedded in `index.html` as a single `let RAW = {...}`
line. To refresh it from a new BG Fund historical workbook:

```bash
# 1. Dry-run: compare the workbook against what's currently embedded.
python3 scripts/gen_raw.py "/path/to/BG Fund - Historical data ....xlsx"

# 2. If the diff looks right, write it in.
python3 scripts/gen_raw.py "/path/to/BG Fund - Historical data ....xlsx" --write
```

The dry-run prints, per section (nav, perf, exp, ear, greeks, lev, cash, ifrs):

- row counts old vs new,
- **added dates** — expected: the new month(s),
- **REVISED** cells — a prior month's value changed in the workbook (e.g. a
  provisional NAV finalised); expected and harmless,
- **BLOCKING** lines — a date or key present in the old data disappeared. This
  usually means the workbook's column layout changed. `--write` refuses to run
  while any BLOCKING problem exists; fix the column mapping in `gen_raw.py` first.

Note the workbook often lags a month on the **Cash** and **IFRS13** sheets, so those
two sections legitimately end one month earlier than the rest.

Requires `openpyxl` (`pip install openpyxl`).
