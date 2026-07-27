# 04 — Report Template Analysis

Analysis performed by extracting the raw OOXML (`unzip`) of the sample workbook and inspecting
`xl/worksheets/*.xml`, `xl/drawings/*.xml`, `xl/charts/chart1.xml`, `xl/pivotTables/*.xml`, and
`xl/pivotCache/*.xml`, cross-checked against `openpyxl` cell/style introspection.

## Workbook Structure

| Sheet | Sheet ID | Visibility | Dimensions | Zoom | Gridlines |
|---|---|---|---|---|---|
| `Introduction` | 1 | Visible | A1:B20 | Default | Default |
| `VA Report` | 2 | Visible | A1:XFD158 (effective data area A1:I158) | 70% | Default |
| `Summary` | 3 | Visible | A1:O922 (effective data area A1:O29) | 85% | Default |

No hidden sheets were found in the sample. **If a real production template contains hidden
sheets (e.g., a hidden calculation/staging sheet), that must be reconfirmed — see Open
Questions.**

## Purpose of Each Sheet

- **`Introduction`** — Static "how to read this report" legend page plus scanner and report-owner
  identity fields. Almost entirely static content; only two small fields change per engagement.
- **`VA Report`** — The core deliverable: a formatted, filtered, deduplicated table of
  vulnerability findings, one row per finding-per-host.
- **`Summary`** — Executive summary: which IPs were tested (with scan type/device type), plus a
  pivot table and pivot pie chart showing the count of findings by risk severity.

## Section Breakdown

### `Introduction` sheet

| Range | Content | Static/Dynamic |
|---|---|---|
| A1 (merged A1:B1) | Title "Introduction" | Static |
| A3 | "How to read the Report" | Static |
| A4:B10 | Legend rows explaining each report column (Vulnerability Title, Vulnerability Description, Risk, Host, Recommendation, Deadline for Corrective Action, Asset Owner Response) | Static |
| A13 | "Security Scanner Details" | Static |
| A14 / B14 | "Security Scanner Name" / value (`Nessus`) | Label static, value **dynamic** (scanner may change) |
| A15 / B15 | "Security Scanner Signature Version" / value (`10.11.4`) | Label static, value **dynamic** (per engagement) |
| A18 | "Security Testing Team" | Static |
| A19 / B19 | "Report owner :-" / value (`John Jacob`) | Label static, value **dynamic** (per engagement) |
| A20 | "Information Security Consultant (Secunatix Consultants Private Limited)" | Static — this is the vendor company name, not client-specific |

### `VA Report` sheet

| Range | Content | Static/Dynamic |
|---|---|---|
| I1:XFD3 (merged) | Logo image anchor area (top-right) | Static (image only, no text) |
| B5 / C5 | "Client Name:" / value | Label static, value **dynamic** |
| B6 / C6 | "Security Tester:" / value | Label static, value **dynamic** |
| B7 / C7 | "Reviewed By:" / value | Label static, value **dynamic** |
| B8 / C8 | "Report Date:" / value (date) | Label static, value **dynamic** |
| B9 / C9 | "Report Version:" / value (e.g., `1.2`) | Label static, value **dynamic** |
| B10 / C10 | "Scanner:" / value (`Nessus`) | Label static, value **dynamic** |
| A13:I13 | Column headers: `Sr. no`, `Vulnerbility Title` *(sic — preserve typo)*, `Description`, `Risk`, `Host`, `Port`, `Recommendation ` *(trailing space — preserve)*, `Reference`, `CVE` | Static (never change wording/spelling) |
| A14:I158 (sample) | Data rows | **Fully dynamic** — this is the primary write target for automation |

### `Summary` sheet

| Range | Content | Static/Dynamic |
|---|---|---|
| O1:Q3 (approx., via drawing anchor) | Logo image (second placement) | Static |
| A5:C6 (merged) | "List of Ips in scope" heading | Static |
| A7:C7 | Column headers `IP Address`, `Scan Type`, `Device Type` | Static |
| A8:C26 (sample) | Scope table rows (one per distinct scanned IP) | **Dynamic** — requires distinct hosts + externally supplied scan-type/device-type metadata |
| J7:O7 (merged) | "Vulnerability Chart" heading | Static |
| E11:F16 | **Stale/leftover pivot summary table** from a prior engagement (Grand Total 194) | Should be treated as a defect to eliminate, not reproduced — see `13_Open_Questions.md` |
| E18:F23 | Current pivot summary table (`Row Labels` = Critical/High/Medium/Low/Grand Total; `Count of Host` values) | **Dynamic** — regenerated from `VA Report` data |
| H8:P29 (approx., via `drawing2.xml` anchor) | Pivot **pie chart**, bound to `PivotTable1` | **Dynamic** — regenerated alongside the pivot table |

## Tables

| Table | Location | Type |
|---|---|---|
| VA findings table | `VA Report!A13:I158` | Plain formatted range (not an Excel "Table" object — confirmed no `xl/tables/*.xml` present); header row + manually bordered data rows |
| Scope table | `Summary!A7:C26` | Plain formatted range |
| Pivot summary table(s) | `Summary!E11:F16` and `Summary!E18:F23` | Native Excel PivotTable (`xl/pivotTables/pivotTable1.xml`, cache `xl/pivotCache/pivotCacheDefinition1.xml` + `pivotCacheRecords1.xml`) |

## Charts

| Chart | Type | Source | Location |
|---|---|---|---|
| Chart 2 | **Pie chart** (`<c:pieChart>` in `xl/charts/chart1.xml`) | `PivotTable1` on `Summary` sheet (pivot-backed chart, `<c:pivotSource>`) | Anchored `Summary!H8` to `P29` approx. (via `xl/drawings/drawing2.xml`, columns 7–15, rows 7–28) |

The chart is a **PivotChart**, not a static chart — meaning in the original manual workflow, it
is expected to auto-update whenever the underlying PivotTable is refreshed. Automation should
either (a) rebuild the pivot table + pivot chart natively so Excel's own refresh mechanics apply,
or (b) generate an equivalent static chart from the same aggregated data if native pivot-chart
generation is not practical in the chosen Python library — this is a design decision for
`11_Project_Architecture.md` and should be explicitly resolved, not assumed.

## Images / Logo Placement

| Image | File | Used on | Anchor |
|---|---|---|---|
| Company logo | `xl/media/image1.png` | `VA Report` sheet (top-right, via `drawing1.xml`, `rId1`) and `Summary` sheet (top-right, via `drawing2.xml`, `rId2`) | Both anchors are `twoCellAnchor` / `oneCell`-edit type, non-resizing, fixed pixel offsets — **must be preserved exactly**, not regenerated or re-scaled |

There is only **one embedded image** (`image1.png`) in the sample workbook, reused twice. No
watermark, background image, or secondary branding element was found.

## Headers and Footers

- Excel print header/footer (`oddHeader`/`oddFooter`) is **empty/unset** on all three sheets in
  the sample file. This means the visible "header block" (client name, tester, etc. on rows 5–10
  of `VA Report`) is an **in-sheet content block**, not an Excel page header — it will appear
  once at the top of the sheet, not repeated on every printed page unless print titles / repeat
  rows are separately configured (not currently configured in the sample — flagged in Open
  Questions if repeat-on-every-page behavior is actually required for printing).

## Merged Cells

| Sheet | Merged ranges |
|---|---|
| `Introduction` | `A1:B1` |
| `VA Report` | `I1:XFD3` (logo area) |
| `Summary` | `A5:C6` (scope heading), `J7:O7` (chart heading) |

## Print Settings

- Page margins: default Excel values on all sheets (left/right 0.7", top/bottom 0.75",
  header/footer 0.3") — **not customized**, meaning the sample was likely never fine-tuned for
  print/PDF export. No defined print area, no forced orientation, no fit-to-page scaling.
- **This should be flagged to the business owner**: if the report is regularly printed or
  exported to PDF, a real print layout (print area, orientation, fit-to-width) is expected but
  was not present in the sample. Automation should not silently invent print settings — see
  `13_Open_Questions.md`.

## Named Ranges

None found (`Defined Names` collection is empty in both workbooks).

## Formatting

| Element | `VA Report` header row (13) | `VA Report` data rows (14+) |
|---|---|---|
| Font | Cambria | Cambria |
| Size | 16pt | 11pt |
| Bold | Yes | No |
| Fill | Solid, theme color (Accent, theme index 7) | No fill (white/default) — **no per-severity color coding was found**; all `Risk` cells share the same style index regardless of Critical/High/Medium/Low value |
| Alignment | Center, wrap text | Left (default), wrap text on |
| Borders | (not explicitly checked on header, assume consistent) | Thin border, all 4 sides, every cell |
| Row height | 39pt | ~171pt (tall, to fit wrapped long-form `Description`/`Recommendation` text) |

### Column Widths (`VA Report` sheet)

| Column | Header | Width |
|---|---|---|
| A | Sr. no | 12.33 |
| B | Vulnerbility Title | 55.11 |
| C | Description | 40.89 |
| D | Risk | 11.0 |
| E | Host | 15.0 |
| F | Port | 8.44 |
| G | Recommendation | 37.66 |
| H | Reference | 43.33 |
| I | CVE | 20.89 |

### Color Palette

- Header fill uses an Excel **theme** color (theme index 7 = typically an accent color in the
  workbook's theme, not a hard-coded RGB) — automation should read/reuse the template's theme
  rather than hard-coding a hex value, so that any future theme/branding change to the template
  automatically propagates.
- **No conditional formatting rules were found anywhere in the workbook** (`conditional_formatting`
  collections are empty on all sheets). Despite common industry convention (red/orange/yellow/
  green severity coloring), this specific template does **not** color-code the `Risk` column.
  This is an important, counter-intuitive finding — flagged explicitly in
  `13_Open_Questions.md` in case the business actually wants that behavior added, since it is
  common in this report category but is not what the sample demonstrates.

## Cell Protection

No sheet protection or cell-locking was found enabled on any sheet in the sample (standard
Excel default: all cells "locked" at the style level but sheet protection itself is off, so this
has no practical effect). If the production template does enable sheet protection, automation
must either unprotect/reprotect programmatically or write via a method that respects protected
sheets — flagged for confirmation.

## Hidden Sheets

None in the sample.

## Elements That Must Never Be Modified by Automation

1. The logo image and both its drawing anchors.
2. The `Introduction` sheet's legend text (rows 3–10).
3. The exact column header text on `VA Report!A13:I13`, including the "Vulnerbility Title" typo
   and the trailing space in "Recommendation ".
4. Font family (Cambria), header fill theme color, border style, and column widths.
5. The vendor company name ("Secunatix Consultants Private Limited") on `Introduction!A20`
   unless explicitly instructed otherwise per engagement.
