# Merging Vulnerability Scan Exports into `RAW_File.xlsx` Format

## 1. Background

Each source file (e.g. `Cisco_Meraki_Switch_*.csv`, `Endpoint_Dealer_system_*.csv`,
`Active_Directory_*.csv`) is a per-target export from the same scanner (Nessus/Tenable
style output), all sharing **one identical 17-column schema**. `RAW_File.xlsx` is simply
every one of these exports stacked on top of each other into a single sheet, with the
header row appearing only once.

The script's job: take a folder containing any number of these exports (`.csv` and/or
`.xlsx`), validate they share the expected schema, concatenate their rows in file order,
and write out one combined `.xlsx` that matches `RAW_File.xlsx`'s structure exactly.

## 2. Expected Input Schema

Every input file must contain exactly these 17 columns, in this order:

| #   | Column               | Type in output                                         |
| --- | -------------------- | ------------------------------------------------------ |
| 1   | Plugin ID            | integer                                                |
| 2   | CVE                  | string (blank if empty)                                |
| 3   | CVSS v2.0 Base Score | float (blank if empty)                                 |
| 4   | Risk                 | string                                                 |
| 5   | Host                 | string                                                 |
| 6   | Protocol             | string                                                 |
| 7   | Port                 | integer                                                |
| 8   | Name                 | string                                                 |
| 9   | Synopsis             | string                                                 |
| 10  | Description          | string (may contain embedded newlines)                 |
| 11  | Solution             | string                                                 |
| 12  | See Also             | string (blank if empty)                                |
| 13  | Plugin Output        | string (blank if empty, may contain embedded newlines) |
| 14  | CVSS v4.0 Base Score | float (blank if empty)                                 |
| 15  | CVSS v3.0 Base Score | float (blank if empty)                                 |
| 16  | VPR Score            | float (blank if empty)                                 |
| 17  | EPSS Score           | float (blank if empty)                                 |

Notes observed from the sample files:

- All numeric-looking fields are double-quoted in the raw CSVs (e.g. `"10107"`), but
  should still land as real numbers (int/float) in the output — not text — matching
  `RAW_File.xlsx`.
- Empty quoted fields (`""`) should become truly empty cells (`None`/NaN), not the
  literal string `"nan"` or `"None"`.
- **Critical gotcha:** the `Risk` column legitimately contains the literal text
  `"None"` as one of its severity levels (None / Low / Medium / High / Critical), and
  `Solution` legitimately contains the literal text `"n/a"` for many findings. Both of
  these are real data, not missing values — but pandas' default NA-detection treats
  `"None"`, `"n/a"`, `"NA"`, `"NULL"`, `"nan"`, etc. as missing by default and will
  silently blank them out on read. The script below disables that default behavior
  (`keep_default_na=False, na_values=[]`) and decides "empty" for itself instead.
- `Description` and `Plugin Output` routinely contain embedded newlines inside quoted
  CSV fields — a standards-compliant CSV reader (Python's `csv` module / pandas) handles
  this correctly; do not attempt to parse the file line-by-line.
- Files may arrive as either `.csv` or `.xlsx` — the script should accept both.

## 3. Algorithm

1. Accept an input folder path and an output file path from the command line.
2. Find every `.csv` and `.xlsx` file in the input folder (skip the output file itself
   and any temporary `~$` lock files).
3. For each file:
   - Load it into a DataFrame (`pandas.read_csv` for `.csv`, `pandas.read_excel` for
     `.xlsx`).
   - Strip whitespace from column headers and confirm they exactly match the 17
     expected column names (order-independent check, then reorder to the canonical
     order). If a file's columns don't match, skip it and warn the user rather than
     silently corrupting the merge.
4. Concatenate all validated DataFrames in filename order (or the order given on the
   command line), keeping only **one** header row overall.
5. **By default, keep every row exactly as-is — including exact duplicates.** If the
   same finding appears twice, three times, or N times across source files (or within
   one), all copies are preserved in the output; nothing is deduplicated or dropped
   unless the user explicitly opts in (see `--dedupe` below, off by default).
6. Write the combined DataFrame to a single-sheet `.xlsx` using `pandas.ExcelWriter`
   (engine `openpyxl`), letting pandas' natural dtype inference produce ints/floats for
   the numeric columns and blank cells for missing values — this reproduces
   `RAW_File.xlsx`'s formatting without extra type-coercion code.
7. Print a short summary: number of files merged, number of rows contributed by each,
   and the total row count in the output.

## 4. Reference Implementation

```python
#!/usr/bin/env python3
"""
merge_scan_exports.py

Merge multiple vulnerability-scan export files (.csv and/or .xlsx), all sharing the
same 17-column Nessus-style schema, into a single combined .xlsx matching the
RAW_File.xlsx layout.

Usage:
    python merge_scan_exports.py --input-dir ./exports --output ./RAW_File_merged.xlsx
    python merge_scan_exports.py --input-dir ./exports --output ./out.xlsx --dedupe
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

EXPECTED_COLUMNS = [
    "Plugin ID",
    "CVE",
    "CVSS v2.0 Base Score",
    "Risk",
    "Host",
    "Protocol",
    "Port",
    "Name",
    "Synopsis",
    "Description",
    "Solution",
    "See Also",
    "Plugin Output",
    "CVSS v4.0 Base Score",
    "CVSS v3.0 Base Score",
    "VPR Score",
    "EPSS Score",
]


TEXT_COLUMNS = [
    "CVE", "Risk", "Host", "Protocol", "Name", "Synopsis",
    "Description", "Solution", "See Also", "Plugin Output",
]
INT_COLUMNS = ["Plugin ID", "Port"]
FLOAT_COLUMNS = [
    "CVSS v2.0 Base Score", "CVSS v4.0 Base Score",
    "CVSS v3.0 Base Score", "VPR Score", "EPSS Score",
]


def load_one_file(path: Path) -> pd.DataFrame | None:
    """Load a single export file and validate/normalize its columns.

    IMPORTANT: this scan format uses the literal text values "None" (a Risk
    level) and "n/a" (a Solution placeholder) as real, meaningful data — not
    as missing-value markers. pandas' default NA-detection treats both of
    those strings (along with "NA", "NULL", "nan", etc.) as missing data and
    will silently blank them out. We disable that default behavior entirely
    (`keep_default_na=False, na_values=[]`) and only decide what counts as
    "empty" ourselves, afterwards.
    """
    try:
        read_kwargs = dict(keep_default_na=False, na_values=[], dtype=str)
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path, **read_kwargs)
        elif path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, **read_kwargs)
        else:
            return None
    except Exception as exc:
        print(f"  [SKIP] Could not read {path.name}: {exc}", file=sys.stderr)
        return None

    df.columns = [str(c).strip() for c in df.columns]

    if set(df.columns) != set(EXPECTED_COLUMNS):
        missing = set(EXPECTED_COLUMNS) - set(df.columns)
        extra = set(df.columns) - set(EXPECTED_COLUMNS)
        print(
            f"  [SKIP] {path.name} does not match expected schema. "
            f"Missing: {missing or 'none'} | Unexpected: {extra or 'none'}",
            file=sys.stderr,
        )
        return None

    # Reorder to the canonical column order
    df = df[EXPECTED_COLUMNS]

    # Only a genuinely empty cell ("") should become blank. "None" and "n/a"
    # are real values in this dataset and must survive untouched.
    for col in TEXT_COLUMNS:
        df[col] = df[col].apply(lambda v: None if v == "" else v)

    # Numeric columns: an empty string legitimately means "no score" -> blank.
    # pd.to_numeric's errors="coerce" turns "" (and any non-numeric junk) into
    # NaN, which is exactly what we want here.
    for col in INT_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    for col in FLOAT_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input-dir", required=True, type=Path,
        help="Folder containing the .csv/.xlsx exports to merge",
    )
    parser.add_argument(
        "--output", required=True, type=Path,
        help="Path to write the combined .xlsx file",
    )
    parser.add_argument(
        "--dedupe", action="store_true",
        help="Drop exact duplicate rows across the merged files",
    )
    args = parser.parse_args()

    input_files = sorted(
        p for p in args.input_dir.iterdir()
        if p.suffix.lower() in (".csv", ".xlsx", ".xls")
        and not p.name.startswith("~$")
        and p.resolve() != args.output.resolve()
    )

    if not input_files:
        print(f"No .csv/.xlsx files found in {args.input_dir}", file=sys.stderr)
        sys.exit(1)

    frames = []
    print(f"Found {len(input_files)} candidate file(s):")
    for path in input_files:
        df = load_one_file(path)
        if df is not None:
            print(f"  [OK]   {path.name}: {len(df)} rows")
            frames.append(df)

    if not frames:
        print("No valid files to merge — nothing written.", file=sys.stderr)
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)

    if args.dedupe:
        before = len(combined)
        combined = combined.drop_duplicates()
        print(f"Deduped: {before} -> {len(combined)} rows")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Sheet1")

    print(f"\nDone. Wrote {len(combined)} rows to {args.output}")


if __name__ == "__main__":
    main()
```

## 5. Usage

```bash
pip install pandas openpyxl

python merge_scan_exports.py \
    --input-dir ./exports \
    --output ./RAW_File_merged.xlsx
```

By default **no rows are ever removed** — if a finding appears multiple times across
your source files, it will appear that many times in the output too. Only add
`--dedupe` if you specifically want identical rows (same Plugin ID + Host + Port +
everything else) collapsed into one; leave it off (the default) to preserve every row
untouched, duplicates included.

## 6. Edge Cases to Keep in Mind

- **Extra/renamed columns**: files with a mismatched schema are skipped with a warning
  rather than aborting the whole merge — check the console output for `[SKIP]` lines.
- **Encoding**: if a source file isn't UTF-8, pass `encoding="latin-1"` (or the correct
  encoding) to `pd.read_csv` for that file.
- **Very large merges**: for tens of thousands of rows, this in-memory pandas approach
  is fine; beyond a few hundred thousand rows, consider `csv.DictWriter` streaming
  instead of holding everything in a DataFrame.
- **Row order**: files are merged in alphabetical filename order by default — pass an
  explicit file list to `load_one_file` in a loop if a specific order is required
  instead.
