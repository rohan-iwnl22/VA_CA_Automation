# 11 — Project Architecture

> **Build-phase scope note (2026-07-09):** Implementation is proceeding **core pipeline
> functions first** (ingestion through save, per `07_VA_Workflow.md`), with the `ui/` layer
> explicitly deferred to a later phase — see `10_UI_Requirements.md` and
> `13_Open_Questions.md` Q12. The `src/` structure below is the current build target; `ui/` is
> shown for completeness/extensibility but should not be started yet.

## Design Principles

1. **Mapping-driven, not hard-coded.** Cell ranges, column mappings, and business rules live in
   configuration (per `09_Excel_Template_Mapping.md` and `05_Business_Rules.md`), not scattered
   as magic strings through code, so a new client template or report type can be added by adding
   config, not rewriting logic.
2. **Template-clone-first.** The pipeline always starts from a pristine copy of the branded
   template and never mutates a template file that might already contain a previous engagement's
   data — this directly prevents the stale-pivot-table defect observed in the manual process.
3. **Values, not formulas.** Matches the manual process's observed behavior of pasting flattened
   values rather than live cross-sheet formulas.
4. **Deterministic and idempotent.** Same input + same metadata → same output, every time.
5. **Fail loud, not silent.** Unknown `Risk` values, missing sheets, or schema drift should raise
   clear, specific errors/warnings rather than being silently dropped.

## Suggested Folder Structure

```
va_ca_automation/
├── config/
│   ├── column_mappings.yaml        # raw -> template column maps (VA + CA)
│   ├── template_ranges.yaml        # cell/range map per 09_Excel_Template_Mapping.md
│   ├── business_rules.yaml         # risk order, dedup keys, filters (05_Business_Rules.md)
│   └── naming_convention.yaml      # filename pattern tokens
├── templates/
│   ├── va_report_template.xlsx     # pristine, blank-data-rows master template
│   └── ca_report_template.xlsx     # pristine CA master template (once confirmed, see 08)
├── src/
│   ├── ingestion/
│   │   ├── raw_file_loader.py      # loads & validates RAW File sheet against schema
│   │   └── schema_validator.py
│   ├── pipelines/
│   │   ├── va_pipeline.py          # filter -> dedup -> map -> sort -> renumber
│   │   └── ca_pipeline.py
│   ├── transform/
│   │   ├── filters.py
│   │   ├── dedup.py
│   │   ├── column_mapper.py
│   │   └── sorter.py
│   ├── excel_writer/
│   │   ├── template_cloner.py      # copies pristine template to a working file
│   │   ├── data_writer.py          # writes rows into dynamic ranges, cloning row styles
│   │   ├── summary_builder.py      # scope table + pivot table regeneration
│   │   └── chart_builder.py        # pivot pie chart regeneration
│   ├── metadata/
│   │   └── engagement_metadata.py  # dataclass/schema for client/tester/date/etc.
│   ├── naming/
│   │   └── filename_builder.py
│   └── logging/
│       └── pipeline_logger.py      # stage-by-stage row-count logging for auditability
├── ui/                              # DEFERRED - not part of current build phase, see 10_UI_Requirements.md
├── tests/
│   └── (see 12_Testing_Checklist.md)
└── output/                          # default generated-report destination
```

## Module Responsibilities

| Module | Responsibility |
|---|---|
| `ingestion.raw_file_loader` | Load `RAW File` sheet; verify 17-column schema; return a validated DataFrame or raise a schema error |
| `ingestion.schema_validator` | Column presence/order check, `Risk` value whitelist check, logs unknown values |
| `pipelines.va_pipeline` | Orchestrates filter → route → dedup → map → sort → renumber for VA, per `07_VA_Workflow.md` |
| `pipelines.ca_pipeline` | Same for CA, per `08_CA_Workflow.md` (marked draft pending real CA sample) |
| `transform.filters` | Risk-based filtering, `None` exclusion, CA inclusion-policy toggle |
| `transform.dedup` | Applies configurable dedup key (default: Name+Description+Risk+Host) |
| `transform.column_mapper` | Raw→template column renaming per `06_Data_Transformation.md` |
| `transform.sorter` | Host grouping + risk-weight ordering + Sr. no assignment |
| `excel_writer.template_cloner` | Copies the pristine template to a working output path; strips any leftover data rows/pivot artifacts as a safety net even on a "pristine" template |
| `excel_writer.data_writer` | Writes VA/CA rows into the dynamic range, cloning cell style from a template master row (font, border, wrap, row height) rather than hardcoding style objects |
| `excel_writer.summary_builder` | Computes scope table and risk-count aggregation; writes/regenerates the native PivotTable |
| `excel_writer.chart_builder` | Regenerates the pivot pie chart bound to the fresh pivot table, reusing the template's existing chart anchor/position/theme |
| `metadata.engagement_metadata` | Typed structure for all per-engagement dynamic fields (client name, tester, dates, per-host scan/device type) |
| `naming.filename_builder` | Constructs the output filename from metadata + naming convention config |
| `logging.pipeline_logger` | Emits structured, stage-by-stage counts (raw rows → filtered → deduped → final) for the UI log panel and for QA/testing |

## Configuration Files

- `column_mappings.yaml` — one section per report type (`va`, `ca`), each a list of
  `{raw_column, template_column}` pairs, directly derived from `06_Data_Transformation.md`.
- `template_ranges.yaml` — one section per sheet, each entry `{range, static_or_dynamic,
  source}`, directly derived from `09_Excel_Template_Mapping.md`. Automation code should read
  this file rather than hard-coding cell references, so a template layout change only requires a
  config update.
- `business_rules.yaml` — dedup key, risk order weight map, filename token order, filter policy
  (e.g., CA inclusion policy once confirmed). Should expose the currently-`INFERRED` rules as
  named, overridable settings rather than baked-in constants, given several are still pending
  business confirmation (`13_Open_Questions.md`).
- `naming_convention.yaml` — token order and separators for filename generation.

## Logging

- Structured (JSON-lines or similar) log per run, capturing:
  - Input file name/hash, timestamp.
  - Row counts at each pipeline stage (raw → risk-filtered → deduped → final).
  - Any unknown `Risk` values encountered (value + row count).
  - Any host in the data with no Scan Type/Device Type metadata supplied.
  - Output file name/path and generation duration.
- This log is the backbone of both the UI's live progress log (`10_UI_Requirements.md`) and the
  automated test suite's assertions (`12_Testing_Checklist.md`).

## Template Handling

- Never write directly into a template file that might be reused across runs — always clone to
  a working copy first (`excel_writer.template_cloner`).
- Prefer reusing the template's existing native Excel objects (PivotTable, PivotChart, cell
  styles, drawing anchors) over generating new ones from scratch wherever the chosen Excel
  library allows it, since re-creating a PivotChart from scratch in a non-Excel library
  (e.g., `openpyxl`) has historically limited fidelity for pivot-bound charts — see the open
  architecture decision below.

## Validation

- Schema validation on ingestion (column names/order, `Risk` value whitelist).
- Post-write validation: re-open the generated workbook and assert:
  - Row count in `VA Report` matches the processed dataset's row count.
  - Pivot table `Grand Total` matches the row count.
  - No orphaned/stale pivot tables exist in the workbook.
  - Static ranges (header text, logo anchors) are byte-identical to the template's originals.

## Error Recovery

- Any pipeline failure should leave the **template file untouched** (fail before/without saving)
  and produce the working copy in a temp/staging location, not directly in the final output
  folder, so a partial/corrupt file is never mistaken for a finished report.
- Retry should be safe to call repeatedly without side effects (idempotency principle).

## Output Generation

- Final save always goes through `naming.filename_builder`, never a user-typed free-text
  filename, to guarantee the naming convention (`05_Business_Rules.md` §11) is always followed.

## Extensibility

- New report types (beyond VA/CA) should only require: a new `column_mappings.yaml` section, a
  new `template_ranges.yaml` section, a new pristine template file, and a new pipeline module
  implementing the same interface as `va_pipeline`/`ca_pipeline`.
- New client-specific branded templates should be supportable by pointing
  `template_ranges.yaml`/`templates/` at an alternate file, provided the alternate template's
  sheet names and dynamic-range structure match the documented mapping (or an additional mapping
  profile is added).

## Open Architecture Decision: Native Pivot vs. Static Aggregation

`09_Excel_Template_Mapping.md` and `07_VA_Workflow.md` both flag this decision point.

**Update (2026-07-09):** the verified blank master template contains **zero** pivot
tables/charts — only the two logo drawing anchors exist in the pristine file. This confirms the
pivot table and chart must be **created from scratch on every run**, not refreshed from an
existing object. This removes one prior consideration in favor of Option A (there is no existing
pivot-cache XML to inherit/preserve), but the underlying trade-off below is otherwise unchanged.

| Option | Pros | Cons |
|---|---|---|
| A. Build a true native Excel PivotTable + PivotChart from scratch (e.g., via a COM/Excel-automation
  approach on Windows, or a library with genuine pivot-cache write support) | Exact behavioral
  and visual fidelity; end user can right-click "Refresh" like the manual file | Higher
  implementation complexity; may require platform-specific tooling |
| B. Write a static aggregated table + a static (non-pivot) pie chart styled to look identical | Simpler, fully cross-platform, easier to test | Loses the "live refreshable pivot" behavior; a sufficiently observant client could notice it's not a real PivotTable if they inspect the file |

**Recommendation for the build phase:** default to Option B for a first working version (it
satisfies the stated visual-fidelity success criterion), while keeping the
`excel_writer.summary_builder`/`chart_builder` interfaces abstracted so Option A can be swapped
in later without touching the rest of the pipeline. Chart anchor position and theme-accent colors
should be hard-coded/configured from the earlier completed sample's coordinates (`H8:P29`
approx.), since the blank master has no existing chart object to copy positioning from. This
should be confirmed with the business — see `13_Open_Questions.md`.

---

## Appendix: Concrete Config File Contents (VA-only, Build Phase)

These are the literal starting contents for the four `config/*.yaml` files referenced above.
Every value is traceable to a specific source document — no new business logic is invented here.
CA sections are stubbed but commented out / marked `enabled: false` per the VA-only build-phase
scope note in `01_Project_Overview.md`; they exist so the CA phase only requires un-commenting
and confirming Q1, not restructuring the config schema.

### `config/column_mappings.yaml`

```yaml
# Raw RAW-File column -> VA Report template column, per 06_Data_Transformation.md
va:
  sheet_name: "VA Report"
  header_row: 13
  data_start_row: 14
  columns:
    # order matters - this is the literal left-to-right column order in the template
    - template_column: "Sr. no"
      raw_column: null          # computed: sequential 1..N after sort, not copied from raw
      dtype: int
    - template_column: "Vulnerbility Title"   # sic - preserve template's existing typo
      raw_column: "Name"
      dtype: str
    - template_column: "Description"
      raw_column: "Description"
      dtype: str
    - template_column: "Risk"
      raw_column: "Risk"
      dtype: str
    - template_column: "Host"
      raw_column: "Host"
      dtype: str
    - template_column: "Port"
      raw_column: "Port"
      dtype: int
    - template_column: "Recommendation "   # trailing space - preserve exactly
      raw_column: "Solution"
      dtype: str
    - template_column: "Reference"
      raw_column: "See Also"
      dtype: str
      null_as_blank: true       # empty string -> None, never the text "None"
    - template_column: "CVE"
      raw_column: "CVE"
      dtype: str
      null_as_blank: true

  # columns present in the raw file but intentionally never written to the report
  dropped_raw_columns:
    - "Synopsis"
    - "Plugin Output"           # sensitive raw evidence - must never reach the client deliverable
    - "CVSS v2.0 Base Score"
    - "CVSS v3.0 Base Score"
    - "CVSS v4.0 Base Score"
    - "VPR Score"
    - "EPSS Score"
    - "Protocol"

  # retained internally for dedup/audit only, never written to a visible cell
  internal_only_columns:
    - "Plugin ID"

ca:
  enabled: false   # 08_CA_Workflow.md is draft - do not build against this section yet
  sheet_name: "CA Report"        # [ASSUMED]
  header_row: 13                 # [ASSUMED, by analogy to VA]
  data_start_row: 14             # [ASSUMED]
  columns:
    - template_column: "Sr. no"
      raw_column: null
      dtype: int
    - template_column: "Control / Check Title"   # [ASSUMED]
      raw_column: "Name"
    - template_column: "Description"
      raw_column: "Description"
    - template_column: "Status"                  # [ASSUMED] - not "Risk", these are compliance states
      raw_column: "Risk"
    - template_column: "Host"
      raw_column: "Host"
    - template_column: "Port"
      raw_column: "Port"
    - template_column: "Recommendation"
      raw_column: "Solution"
    - template_column: "Reference"
      raw_column: "See Also"
    # CVE column intentionally omitted for CA - [ASSUMED] not applicable to most compliance checks
```

### `config/template_ranges.yaml`

```yaml
# Cell-range map per 09_Excel_Template_Mapping.md.
# static ranges must never be written to; dynamic ranges are the only automation write targets.
introduction:
  static:
    - range: "A1:B1"
      purpose: "Title 'Introduction' (merged)"
    - range: "A3:B10"
      purpose: "'How to read the Report' legend"
    - range: "A13"
      purpose: "'Security Scanner Details' heading"
    - range: "A18"
      purpose: "'Security Testing Team' heading"
    - range: "A20"
      purpose: "Vendor company name"
  dynamic:
    - cell: "B14"
      source: "engagement_metadata.scanner_name"
      default: "Nessus "        # pre-filled in the verified blank master; treat as overridable default
    - cell: "B15"
      source: "engagement_metadata.scanner_version"
      default: "10.11.4"        # pre-filled in the verified blank master
    - cell: "B19"
      source: "engagement_metadata.report_owner"

va_report:
  static:
    - range: "I1:XFD3"
      purpose: "Logo image anchor (do not touch drawing XML)"
    - range: "A13:I13"
      purpose: "Column headers - exact text including typo/trailing space, never altered"
  dynamic:
    header_block:
      - cell: "C5"
        source: "engagement_metadata.client_name"
      - cell: "C6"
        source: "engagement_metadata.security_tester"
      - cell: "C7"
        source: "engagement_metadata.reviewed_by"
      - cell: "C8"
        source: "engagement_metadata.report_date"
        excel_type: "date"
      - cell: "C9"
        source: "engagement_metadata.report_version"
        excel_type: "float"
      - cell: "C10"
        source: "engagement_metadata.scanner_name"
    data_table:
      start_row: 14
      start_column: "A"
      end_column: "I"
      source: "va_sorted_df"
      style_source: "clone header-adjacent master row style (Cambria 11, thin border, wrap_text)"
      clear_before_write: true   # always clear/reset any leftover rows first - prevents stale-data defect

summary:
  static:
    - range: "O1:Q3"
      purpose: "Logo image anchor (2nd placement)"
    - range: "A5:C6"
      purpose: "'List of Ips in scope' heading (merged)"
    - range: "A7:C7"
      purpose: "Scope table headers"
    - range: "J7:O7"
      purpose: "'Vulnerability Chart' heading (merged)"
  forbidden_ranges:
    # confirmed absent from the blank master - automation must never create these
    - range: "E11:F16"
      reason: "Stale/leftover pivot artifact observed in the earlier completed sample only"
  dynamic:
    scope_table:
      start_row: 8
      start_column: "A"
      end_column: "C"
      source: "scope_df"
      style_source: "clone existing row style"
    risk_pivot:
      anchor: "E18"              # match completed-sample layout; created fresh every run
      created_fresh_every_run: true
      source: "risk_summary_df"
      row_labels: ["Critical", "High", "Medium", "Low", "Grand Total"]
    pie_chart:
      anchor: "H8:P29"           # approx, from completed-sample drawing2.xml graphicFrame anchor
      created_fresh_every_run: true
      bound_to: "risk_pivot"
      style: "theme-accent-ordered colors, matching completed sample"

never_touch:
  - "Introduction!A1:B10"
  - "Introduction!A13"
  - "Introduction!A18"
  - "Introduction!A20"
  - "VA Report!I1:XFD3"
  - "VA Report!A13:I13"
  - "Summary!A5:C6"
  - "Summary!A7:C7"
  - "Summary!J7:O7"
  - "Summary drawing/logo anchors (both placements)"
```

### `config/business_rules.yaml`

```yaml
# Filtering, dedup, sort, and naming rules per 05_Business_Rules.md.
# Status tags (confirmed/inferred) are preserved as comments for traceability.

risk_routing:
  va_values: ["Critical", "High", "Medium", "Low", "None"]   # CONFIRMED
  ca_values: ["PASSED", "FAILED", "WARNING"]                  # CONFIRMED
  unknown_value_policy: "log_and_exclude"                     # INFERRED (safety default)

filtering:
  va:
    exclude_risk_values: ["None"]     # CONFIRMED
  ca:
    enabled: false                    # deferred - Q1 unresolved (13_Open_Questions.md)
    inclusion_policy: null            # "all_states" | "exceptions_only" - set once Q1 is resolved

deduplication:
  stage1_exact:
    key_fields: ["Name", "Description", "Risk", "Host"]   # CONFIRMED by business
    keep: "first"
    scope: "per_host"                 # same finding on 2 hosts = 2 rows, not merged
  stage2_version_collapse:
    enabled: true                     # CONFIRMED by business
    group_key_fields: ["_base_title", "Risk", "Host"]
    version_pattern: '(?:<=|<|=|version)?\s*(\d+(?:\.\d+){1,4})'
    version_token_selection: "last_match"
    keep: "highest_version"
    on_no_version_token: "pass_through_unchanged"   # advisory-code-only titles unaffected - Q4

sorting:
  group_by: "host"                    # CONFIRMED
  host_block_order: "first_seen"      # INFERRED default - see Q11, ascending-IP is the alternative
  within_host_order:
    field: "risk"
    weights: {Critical: 0, High: 1, Medium: 2, Low: 3}   # CONFIRMED - not alphabetical
  tiebreaker: "stable_original_row_order"   # INFERRED

renumbering:
  field: "Sr. no"
  start: 1
  step: 1

summary:
  pivot_source: "processed_va_report_data"   # never the raw file - CONFIRMED
  regenerate_every_run: true                 # CONFIRMED - never reuse/refresh stale pivot
  chart_type: "pie"                          # CONFIRMED
  chart_color_source: "theme_accent_order"   # CONFIRMED - no fixed severity color mapping
  scope_table_includes_zero_finding_hosts: null   # INFERRED - open, see Q10

formatting:
  font_family: "Cambria"
  header_font_size: 16
  data_font_size: 11
  header_bold: true
  data_bold: false
  border: "thin_all_sides"
  wrap_text_columns: ["Vulnerbility Title", "Description", "Recommendation ", "Reference"]
  severity_color_coding: false        # CONFIRMED absent in sample - flagged as Q7, toggle-able later

write_mode: "values_only"             # CONFIRMED - flatten to pasted values, never live formulas
```

### `config/naming_convention.yaml`

```yaml
# Filename pattern per 05_Business_Rules.md §11.
# Pattern: <ReportType>_<Scope>_<Phase>_Audit_Report_<ClientLegalName>_<EntityCode(s)>_<Year>_V<Major>.<Minor>
pattern_tokens:
  - name: "report_type"
    source: "static per pipeline"
    values: {va: "VA", ca: "CA"}
  - name: "scope"
    source: "engagement_metadata.scope_label"     # e.g. "Server", "Firewall"
  - name: "phase"
    source: "engagement_metadata.phase_label"      # e.g. "First", "Retest"
  - name: "literal"
    value: "Audit_Report"
  - name: "client_legal_name"
    source: "engagement_metadata.client_name"
    transform: "spaces_to_underscores"
  - name: "entity_codes"
    source: "engagement_metadata.entity_codes"      # optional, free-text tokens - see Q5
    optional: true
  - name: "year"
    source: "engagement_metadata.report_date.year"
  - name: "version"
    source: "engagement_metadata.report_version"
    transform: "period_to_underscore"    # e.g. 1.2 -> V1_2 in the filename (independent of the in-sheet numeric field - see Q13)
    prefix: "V"

separator: "_"
extension: ".xlsx"

sanitization:
  strip_characters: ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '&']
  replace_with: ""

overwrite_policy: "prompt_or_auto_increment"   # never silently overwrite - per 11_Project_Architecture.md
```

**Note on scope:** these four files fully cover the VA pipeline as specified in
`07_VA_Workflow.md`. The `ca:` sections above are placeholders only — they should not be treated
as implementation-ready until `13_Open_Questions.md` Q1 is resolved and a real CA template
sample is supplied, per `08_CA_Workflow.md`'s recommendation.
