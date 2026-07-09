# 03 — Input Files

## File 1: Raw Scanner Export

| Attribute | Value |
|---|---|
| Sample file name | `RAW_file_-_Copy.xlsx` |
| Source | Nessus vulnerability scanner export (combined VA + CA/compliance results), exported to `.xlsx` |
| Purpose | Primary raw data source for both VA and CA report pipelines |
| Sheet(s) | Single sheet, named exactly **`RAW File`** |
| Header row | Row 1 |
| Data rows | Rows 2 – 19,959 (19,958 data rows in sample) |
| Column count | 17 (A–Q) |

### Column Schema

| Col | Header | Data Type | Mandatory | Notes |
|---|---|---|---|---|
| A | `Plugin ID` | Integer | Yes | Unique Nessus plugin identifier. Used as a stable dedup/reference key. |
| B | `CVE` | String (nullable) | No | May be blank; may contain multiple CVEs (needs delimiter check — see Validation Rules) |
| C | `CVSS v2.0 Base Score` | Float (nullable) | No | Legacy scoring; often blank for compliance-only plugins |
| D | `Risk` | String (categorical) | **Yes** | One of: `Critical`, `High`, `Medium`, `Low`, `None` (VA) **or** `PASSED`, `FAILED`, `WARNING` (CA). This single column is the primary routing key between the VA and CA pipelines. |
| E | `Host` | String (IP address) | **Yes** | IPv4 address of scanned asset. Sample file has been anonymized to `192.x.x.N` — real exports will contain full IPs. |
| F | `Protocol` | String | No | `tcp` / `udp` |
| G | `Port` | Integer | No | `0` used for host-level (non-port-specific) findings |
| H | `Name` | String | **Yes** | Vulnerability / check title — maps to template's `Vulnerbility Title` |
| I | `Synopsis` | String | No | One-line summary (not used directly in final report — see `06_Data_Transformation.md`) |
| J | `Description` | String (long text) | **Yes** | Full finding description — maps to template's `Description` |
| K | `Solution` | String (long text) | **Yes** | Remediation text — maps to template's `Recommendation` |
| L | `See Also` | String (nullable, multi-line URLs) | No | Reference links — maps to template's `Reference` |
| M | `Plugin Output` | String (long text, nullable) | No | Raw scanner evidence output (e.g., SMB share listings, NetBIOS names). **Not copied into the final report** — contains highly sensitive raw evidence (file listings, share names, etc.) and must be excluded from client deliverable by design. |
| N | `CVSS v4.0 Base Score` | Float (nullable) | No | |
| O | `CVSS v3.0 Base Score` | Float (nullable) | No | |
| P | `VPR Score` | Float (nullable) | No | Tenable Vulnerability Priority Rating — not used in the sample final report but potentially useful for future prioritization features |
| Q | `EPSS Score` | Float (nullable) | No | Exploit Prediction Scoring System — same as above |

### Relationships

- `Host` + `Port` + `Plugin ID` together uniquely identify a single scanner check result.
- `Risk` value determines pipeline routing (VA vs. CA) — see `05_Business_Rules.md`.
- `Name` + `Description` + `Risk` + `Host` together form the **deduplication key** used when
  building the final VA Report (see `02_Current_Workflow.md`, Step 6).

### Validation Rules

| Rule | Reason |
|---|---|
| `Risk` must be non-null and one of the 8 known values | Unknown values cannot be routed to VA or CA pipeline; should be logged and flagged for manual review rather than silently dropped |
| `Host` must match IPv4 pattern (or the anonymized `192.x.x.N` pattern seen in this sample) | Downstream `Summary` sheet's IP list depends on distinct, valid host values |
| `Name` and `Description` must be non-empty for any row entering the VA pipeline | These map directly to required template columns |
| Sheet name must be exactly `RAW File` | Automation should fail fast with a clear error if the sheet is renamed/missing, rather than guessing |
| Header row must exactly match the 17 expected column names, in order | Protects against a scanner export format change going unnoticed |

### Possible Errors / Failure Modes

- Export contains additional/renamed columns (Nessus version upgrade changes export schema).
- `CVE` field contains multiple CVEs in one cell (comma or newline separated) — needs a defined
  splitting/formatting rule before use (flagged in `13_Open_Questions.md`).
- Extremely large `Plugin Output` cells slow down load if that column is not explicitly skipped.
- Duplicate header row accidentally included as a data row (row 1 repeated further down after a
  bad export/merge).
- Mixed-case or trimmed/untrimmed `Risk` values (`"critical "`, `"CRITICAL"`) causing routing
  failures.

---

## File 2: Branded Report Template / Sample Final Report

| Attribute | Value |
|---|---|
| Sample file name | `VA_Server_First_Audit_Report_Trust_Investment_Advisors_Private_Limited_TSS__SCPL_2026_V1_0_-_Copy.xlsx` |
| Source | Company-branded Excel report template, already populated with one completed engagement's data (this is both **the template** to clone and **a worked example** of correct final output) |
| Purpose | (a) Defines the exact visual/structural target the automation must reproduce; (b) provides ground-truth data to validate the automated pipeline's row counts and summary numbers |
| Sheets | `Introduction`, `VA Report`, `Summary` (see `04_Report_Template_Analysis.md` for full breakdown) |

### Expected Structure Per Sheet

| Sheet | Mandatory sections | Optional/variable sections |
|---|---|---|
| `Introduction` | "How to read the Report" legend (static), Security Scanner Details, Security Testing Team / report owner | None observed |
| `VA Report` | Header metadata block (rows 5–10), column headers (row 13), data rows (row 14+) | Row count varies per engagement |
| `Summary` | "List of IPs in scope" table, one pivot table (`Count of Host` by `Risk`), one pivot pie chart | A second, stale pivot table was observed in the sample — see `13_Open_Questions.md`; this should **not** be treated as a required structural element, it is a data-hygiene defect to avoid replicating |

### Relationships to File 1 (Raw Export)

- `VA Report` data rows are a filtered, deduplicated, remapped, resorted, renumbered subset of
  File 1's VA-type rows (see `02_Current_Workflow.md` and `06_Data_Transformation.md`).
- `Summary` sheet's pivot table is a count aggregation of `VA Report!Risk` — it is **not**
  computed from File 1 directly; it is downstream of the already-processed `VA Report` sheet.
- `Summary` sheet's IP list (`Scan Type`, `Device Type`) is **not derivable from File 1 at all**
  — it requires separate engagement metadata supplied by the analyst (see `13_Open_Questions.md`).

### Validation Rules

| Rule | Reason |
|---|---|
| Template file must retain sheets named exactly `Introduction`, `VA Report`, `Summary` | Automation writes to these sheets by name |
| Logo image (`image1.png`) and its two drawing anchors must be preserved unmodified | Required for visual fidelity — see `04_Report_Template_Analysis.md` |
| Header row (row 13 of `VA Report`) must not be altered (including the existing typo "Vulnerbility Title") | Any deviation from the client's established template breaks visual fidelity, which is the core success criterion |
| Pivot table/pivot chart definitions must be regenerated from current data, never left stale | See defect noted in `02_Current_Workflow.md`, Step 13 |

### File 2b: Blank/Pristine Master Template — ✅ SUPPLIED (2026-07-09)

| Attribute | Value |
|---|---|
| File name | `Blank_Template_VA_Server_First_Audit_Report_Trust_Investment_Advisors_Private_Limited_TSS__SCPL_2026_V1_0.xlsx` |
| Purpose | This is the actual pristine starting file automation should clone for every new run — resolves the earlier open question about a missing blank master (`13_Open_Questions.md` Q2) |

Verified structure (directly inspected):

| Sheet | Confirmed state |
|---|---|
| `Introduction` | Legend text (rows 3–10) present and static, matching the completed sample exactly. `B14`/`B15` (Scanner Name / Scanner Signature Version) come **pre-filled with default values `"Nessus "` / `"10.11.4"`** rather than blank — treat these as default values that engagement metadata may override, not as empty required fields. `B19` (Report owner) is blank as expected. |
| `VA Report` | Header block labels (rows 5–10, column B) present, values (column C) blank. Column headers on row 13 present and byte-identical to the completed sample, including the `"Vulnerbility Title"` typo and the `"Recommendation "` trailing space. **Row 14 onward is completely empty — confirmed zero data rows.** |
| `Summary` | `"List of Ips in scope"` heading and `IP Address`/`Scan Type`/`Device Type` column headers present; scope table rows (8+) empty. **Confirmed: no PivotTable, no PivotCache, and no chart exist anywhere in this file** — only the two logo drawing anchors (`drawing1.xml`, `drawing2.xml`, `image1.png`) are present. |

Logo image (`image1.png`) is byte-identical between the blank template and the completed sample.
Column widths, fonts (Cambria), and header styling are also identical between the two files.

**Key implication for the build:** because the pristine template has no pivot table/chart at
all, automation must **create the PivotTable and PivotChart from scratch** on every run — this
is not a "refresh an existing pivot" operation as might be assumed from the completed sample
alone. This directly informs the architecture decision in `11_Project_Architecture.md`.

### Validation Rules (Blank Template)

| Rule | Reason |
|---|---|
| Template file must retain sheets named exactly `Introduction`, `VA Report`, `Summary` | Automation writes to these sheets by name |
| Logo image (`image1.png`) and its two drawing anchors must be preserved unmodified | Required for visual fidelity — see `04_Report_Template_Analysis.md` |
| Header row (row 13 of `VA Report`) must not be altered (including the existing typo "Vulnerbility Title") | Any deviation from the client's established template breaks visual fidelity, which is the core success criterion |
| `VA Report` rows 14+ and `Summary` rows 8+ must be empty in the master template used for cloning | Confirmed baseline state — if a future "blank" template ever contains leftover rows/pivot artifacts, treat it as a data-hygiene defect (per the stale-pivot issue found in the earlier completed sample), not a structural feature to replicate |

### Possible Errors / Failure Modes

- Template file is edited/re-saved by a newer Excel version and changes internal drawing/style
  XML structure in a way that breaks the automation's direct-write approach (mitigation:
  automation should always clone from this verified blank master, and re-verify its structure
  automatically at startup rather than assuming it never changes).
- A future replacement blank template accidentally ships with leftover data rows or a stale
  pivot table (as happened in the original completed sample) — automation's template-loading
  step should assert zero data rows and zero pivot tables/charts before proceeding, failing loudly
  if that assumption doesn't hold.
