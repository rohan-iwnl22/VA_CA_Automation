# 12 — Testing Checklist

## 1. Ingestion / Schema Validation

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Valid raw file loads | Sample `RAW_file.xlsx` | 19,958 data rows loaded, 17 columns, sheet `RAW File` recognized | Row/column count mismatch, sheet not found |
| Missing sheet | File with sheet renamed to `Sheet1` | Clear, specific error naming the expected sheet | Silent failure or generic exception |
| Extra/renamed column | File with `Risk` renamed to `Severity` | Clear schema-mismatch error | Silent misalignment of columns |
| Unknown Risk value | Row with `Risk = "Informational"` | Logged warning, row excluded from both pipelines, count reported | Row silently included in VA or CA output |
| Whitespace/casing variants | `Risk = " none "`, `"CRITICAL"` | Correctly normalized and filtered/classified as if `"None"`/`"Critical"` | Row incorrectly treated as unknown or misrouted |

## 2. Filtering (VA)

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Exclude `Risk = None` | Sample raw file | 0 rows with `Risk = None` in `va_candidates_df` | Any `None` row present |
| Retain Critical/High/Medium/Low | Sample raw file | All 4 severities present in `va_candidates_df` | Any severity missing without cause |
| Exclude CA rows | Sample raw file | 0 rows with `Risk ∈ {PASSED, FAILED, WARNING}` in VA pipeline output | Any CA-type row leaks into VA output |

## 3. Deduplication

### 3a. Stage 1 — Exact dedup

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Exact duplicate removed | Two rows identical on Name+Description+Risk+Host | Only 1 row remains | 2 rows remain |
| Distinct rows retained | Two rows same Name+Description+Risk but different Host | Both rows remain | One row incorrectly dropped |
| Dedup key edge case: same finding, different Port | Two rows same Name/Description/Risk/Host, different Port | **Currently would be merged into 1 row under the exact-match key** — this is expected/confirmed behavior since Port is not part of the key | Behavior silently changes without updating this test |

### 3b. Stage 2 — Version-collapse (CONFIRMED business rule)

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Higher version kept | Two rows, same Host+Risk, `Name` = `"Adobe Flash Player <= 32.0.0.387 ..."` and `"Adobe Flash Player <= 32.0.0.390 ..."` | Only the `32.0.0.390` row remains | Wrong row kept, or both rows kept |
| Older-version row fully dropped, not merged | Same as above, older row has a unique `CVE`/`Port` not present on the newer row | That `CVE`/`Port` is **not** present anywhere in the final output (confirmed: no merging) | Older row's CVE/Port leaks into the kept row |
| Three-or-more version variants | Three rows, same Host+Risk, versions `1.0`, `1.2`, `1.1.5` | Only `1.2` row remains (correct numeric tuple comparison, not string comparison) | String-sort picks `1.2` incorrectly relative to `1.1.5`, or `1.10` sorts below `1.2` due to naive string comparison |
| No version token present | Row with `Name = "Fortinet Fortigate Format String Bug in cli command (FG-IR-23-137)"` (advisory code only, no dotted version number) | Row passes through Stage 2 unchanged | Row incorrectly grouped/dropped based on the advisory code being mistaken for a version |
| Different base titles not falsely collapsed | Two rows, same Host+Risk, different products each with their own version (e.g., `"Adobe Flash Player <= 32.0.0.387..."` and `"Adobe AIR <= 33.0.0.1..."`) | Both rows retained (different `base_title` after version-stripping) | Rows incorrectly collapsed into one due to overly loose base-title matching |
| Padded version comparison | Versions `7.2` vs `7.2.1` in the same group | `7.2.1` correctly identified as higher (via zero-padded tuple comparison) | `7.2` incorrectly treated as equal or higher due to naive comparison |
| Full pipeline row count matches sample | Full sample raw file through VA pipeline (Stage 1 + Stage 2 + filtering) | 145 final rows (Critical 21 / High 32 / Medium 77 / Low 15) | Count deviates from sample ground truth |

## 4. Sorting / Numbering

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Risk order within a host | Rows for one host with mixed severities | Ordered Critical, High, Medium, Low | Alphabetical or raw-file order used instead |
| Sr. no sequential, no gaps | Final sorted dataset | `Sr. no` = 1, 2, 3 ... N with no skips/duplicates | Gaps, duplicate numbers, or off-by-one |
| Stable tiebreak | Two rows same Host + same Risk | Original raw relative order preserved | Non-deterministic ordering across runs |

## 5. Column Mapping

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Raw→template mapping complete | `va_sorted_df` | All template columns (`Vulnerbility Title` ... `CVE`) populated per `06_Data_Transformation.md` | Any mapped field blank when source had data |
| Excluded columns absent | `va_sorted_df` written to sheet | `Plugin Output`, `Synopsis`, CVSS scores not present in `VA Report` sheet | Sensitive/unused raw data leaks into output |
| Null handling | Row with blank `See Also` | `Reference` cell is truly blank, not string `"None"` | Cell literally contains text `"None"` |

## 6. Formatting Validation

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Font family/size | Any written data row | Cambria, 11pt | Different font/size applied |
| Header row untouched | Generated workbook | `VA Report!A13:I13` text and style byte-identical to template | Header text/style altered |
| Borders applied | Any written data row | Thin border, all 4 sides | Missing/partial borders |
| Wrap text enabled | Long `Description` cell | Text wraps, row height accommodates content | Text overflow/truncated visually |
| Column widths unchanged | Generated workbook | Widths match `04_Report_Template_Analysis.md` table exactly | Any column resized |
| No unintended color-coding | `Risk` column cells | No fill color variation by severity (matches sample's plain-text behavior) unless business explicitly requests this feature later | Unexpected conditional coloring introduced |

## 7. Template Validation

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Logo preserved | Generated workbook | `image1.png` present, both anchors at original coordinates, unmodified pixel data | Logo missing, resized, or repositioned |
| Introduction legend untouched | Generated workbook | `Introduction!A3:B10` byte-identical to template | Any legend text altered |
| No leftover data from prior run | Generated workbook from a 2nd run on different input | Old run's rows/pivot tables completely absent | Any stale row/pivot table survives (regression test directly targeting the observed sample defect) |
| Merged cell ranges intact | Generated workbook | Same merged ranges as template (`A1:B1`, `I1:XFD3`, `A5:C6`, `J7:O7`) | Any merge broken/altered |

## 8. Summary / Pivot / Chart Validation

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Pivot Grand Total matches row count | 145-row processed dataset | Pivot `Grand Total` = 145 | Mismatch |
| Pivot per-severity counts correct | 145-row processed dataset | Critical 21 / High 32 / Medium 77 / Low 15 | Any count wrong |
| Only one pivot table present | Generated workbook | Exactly 1 pivot table in `Summary` sheet | 2+ pivot tables (regression test for the stale-pivot defect) |
| Chart reflects current data | Modified input producing different counts | Chart values update to match new counts | Chart shows stale/cached values |
| Chart type is pie | Generated workbook | `<c:pieChart>` present in chart XML | Wrong chart type generated |
| Scope table lists all distinct hosts | Processed dataset with N distinct hosts | N rows in `Summary!A8:A<7+N>` | Missing or duplicate host rows |

## 9. Naming Convention Validation

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| Filename matches pattern | Sample metadata (VA, Server, First, Trust Investment Advisors..., 2026, V1.0) | `VA_Server_First_Audit_Report_Trust_Investment_Advisors_Private_Limited_..._2026_V1_0.xlsx` | Token missing/misordered |
| Special characters sanitized | Client name with `&`, `/`, etc. | Filename-safe characters only | Invalid filename characters cause a save failure |
| No accidental overwrite | Existing file with the same generated name | Confirmation prompt or auto-increment, not silent overwrite | Existing report silently replaced |

## 10. Regression Tests

| Test | Purpose |
|---|---|
| Re-run full pipeline on unchanged input twice | Output is byte-for-byte identical (idempotency) except for any run-timestamp metadata explicitly expected to change |
| Re-run after template file update (e.g., new logo) | New template's static elements propagate correctly without needing pipeline code changes |
| Full end-to-end sample test | Given the exact supplied `RAW_file.xlsx` + the supplied engagement metadata inferred from the sample header block, output matches the supplied final report's row count, per-severity counts, and column content as closely as achievable given the documented open questions |

## 11. CA Pipeline Tests `[DRAFT — pending real CA sample per 08_CA_Workflow.md]`

| Test | Input | Expected Output | Failure Condition |
|---|---|---|---|
| CA row routing | Sample raw file | 5,652 rows routed to CA candidate set (2,857 FAILED + 2,757 PASSED + 38 WARNING) | Any CA row misrouted to VA pipeline or dropped |
| CA filtering policy | Depends on business decision (all states vs. exceptions-only) | Matches whichever policy is confirmed | Output doesn't match confirmed policy — **cannot be finalized until Open Question is resolved** |

## Test Data

- Use the supplied `RAW_file_-_Copy.xlsx` as the canonical regression fixture (ground truth: 145
  final VA rows, Critical 21/High 32/Medium 77/Low 15).
- Use the supplied final report workbook as the "golden file" for formatting/structure
  comparison tests.
- Construct additional small synthetic fixtures (5–10 rows) covering edge cases (duplicate rows,
  unknown Risk values, blank optional fields, multiple hosts with identical findings) for fast
  unit-level testing, since the full 20K-row file is too large for quick unit tests.
