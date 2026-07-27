# 01 — Project Overview

> **Build-phase scope note (2026-07-09):** Implementation is proceeding **VA-only** for the
> first phase. `08_CA_Workflow.md` remains a draft specification for a later phase and should
> not block VA delivery. All other documents in this set describe the full target system;
> where a document covers both VA and CA, the VA-specific portions are the current priority.
>
> **Build-sequencing note (2026-07-09):** The build is further scoped to **core pipeline
> functions first, UI later**. `10_UI_Requirements.md` is a fast-follow spec, not a v1
> deliverable — the first implementation milestone is a working, testable VA pipeline
> (ingestion → filter → dedup → sort → write → summary/chart → save) invocable as a library or
> CLI, with no desktop UI wrapper. This resolves `13_Open_Questions.md` Q12 in favor of
> "core pipeline first, UI as fast-follow."

## 1. Purpose

Secunatix Consultants Private Limited (the security testing team) manually converts raw
vulnerability-scanner exports (Nessus) into polished, client-ready **Vulnerability Assessment
(VA)** and **Configuration Audit (CA)** reports for clients such as *Trust Investment Advisors
Private Limited*. This project documents that manual process in enough depth that a second AI
coding agent can build a **Python automation tool** that ingests the same raw scanner export and
produces a finished Excel workbook that is visually and structurally indistinguishable from a
report a human analyst built by hand in Excel.

This document set is **specification-only**. No code is written here. Every other numbered
document in this folder is an input to the build phase.

## 2. Business Problem

- Analysts receive a raw Nessus export (`RAW_file.xlsx`) containing **~20,000 rows** covering
  both vulnerability findings (Risk = Critical/High/Medium/Low/None) and configuration/compliance
  audit results (Risk = PASSED/FAILED/WARNING) for every scanned host.
- The analyst manually filters, deduplicates, sorts, reformats, and copies this data into a
  branded Word-quality Excel template, then builds summary pivot tables and a pie chart by hand.
- This is done **per client, per engagement, per scan batch**, which is repetitive, slow, and
  error-prone (copy/paste mistakes, inconsistent formatting, stale leftover pivot tables reused
  from a previous report — observed directly in the sample workbook, see
  `13_Open_Questions.md`).
- Report turnaround time and formatting consistency both suffer as a result.

## 3. Existing (Manual) Workflow — Summary

1. Run/export a Nessus scan to Excel (`RAW file`).
2. Open the raw export in Excel.
3. Filter/inspect rows by `Risk` column.
4. Remove informational rows (`Risk = None`) and, for CA-type rows, handle
   `PASSED` / `FAILED` / `WARNING` separately from VA-type rows.
5. Identify duplicate findings (same vulnerability affecting the same host, or the same
   plugin firing more than once) and remove exact duplicates.
6. Copy the cleaned records into the branded report template (`VA Report` sheet), aligning
   raw columns to the template's business columns (`Sr. no`, `Vulnerability Title`,
   `Description`, `Risk`, `Host`, `Port`, `Recommendation`, `Reference`, `CVE`).
7. Sort the records (by host, then by business risk ranking — see `05_Business_Rules.md`).
8. Manually re-number the `Sr. no` column sequentially.
9. Build/refresh the `Summary` sheet: list of in-scope IPs with scan type and device type,
   plus a pivot table (`Count of Host` by `Risk`) and a pivot pie chart.
10. Fill in the `Introduction` sheet metadata (client name, tester, reviewer, report date,
    report version, scanner name/version).
11. Save the workbook using the company naming convention (see `05_Business_Rules.md`).

The complete, click-by-click reconstruction is in `02_Current_Workflow.md`.

## 4. Desired (Automated) Workflow

1. User provides the raw Nessus export (`.xlsx`) plus a small set of engagement metadata
   (client name, tester, reviewer, report date, version, scope type — Server / Firewall / etc.).
2. Python application validates the raw file structure (see `03_Input_Files.md`).
3. Application applies all business rules (filtering, deduplication, sorting, numbering) —
   see `05_Business_Rules.md` and `06_Data_Transformation.md`.
4. Application splits processing into two parallel pipelines — VA and CA — see
   `07_VA_Workflow.md` and `08_CA_Workflow.md`.
5. Application writes the result into a **copy of the branded template**, preserving every
   static visual element (logo, fonts, colors, borders, column widths, merged cells) and only
   touching the dynamic cell ranges identified in `09_Excel_Template_Mapping.md`.
6. Application regenerates the `Summary` sheet's IP list, pivot table, and pivot pie chart from
   the freshly processed data (not left over from a prior report).
7. Application saves the output file using the enforced naming convention.
8. (Optional) A desktop UI wraps this pipeline — see `10_UI_Requirements.md`.

## 5. Inputs

| # | Input | Format | Description |
|---|-------|--------|--------------|
| 1 | Raw scanner export | `.xlsx` (single sheet `RAW File`) | Combined Nessus VA + CA export, ~20K rows, 17 columns |
| 2 | Report template | `.xlsx` (3 sheets: `Introduction`, `VA Report`, `Summary`) | Branded, pre-formatted workbook shell |
| 3 | Engagement metadata | Form fields / config | Client name, tester, reviewer, date, version, scanner name/version, scan type per IP, device type per IP |
| 4 | Company logo | `.png` (already embedded in template) | Reused as-is; not regenerated |

## 6. Outputs

| # | Output | Format | Description |
|---|--------|--------|--------------|
| 1 | Final VA report | `.xlsx` | 3-sheet workbook matching the manual report exactly |
| 2 | Final CA report | `.xlsx` | Equivalent workbook for configuration-audit findings (structure inferred — see `08_CA_Workflow.md` and Open Questions) |
| 3 | Processing log | `.log` / `.txt` | Record of filtering, dedup, and row counts for QA |

## 7. Functional Goals

- Reproduce 100% of the manual filtering, deduplication, and sorting logic.
- Reproduce the `VA Report` sheet's exact column layout, formatting, fonts, and row structure.
- Regenerate the `Summary` sheet's IP list, pivot table, and pivot pie chart automatically
  from the processed dataset (not manually).
- Preserve all static template elements — logo, headers, Introduction sheet content — untouched.
- Support both VA and CA report generation from a single raw export.
- Enforce the observed file-naming convention automatically.

## 8. Non-Functional Goals

- **Fidelity**: output must be visually indistinguishable from a manually produced report
  (same fonts — Cambria; same column widths; same borders; same logo placement; same chart type
  — pie chart on a pivot table).
- **Performance**: process a ~20,000-row raw export in well under a minute.
- **Idempotency**: re-running the tool on the same raw file produces the same output (no
  stale/leftover data carried from a previous run, unlike the observed manual process).
- **Auditability**: every filtering/dedup decision should be traceable in a log.
- **Extensibility**: the mapping-driven design (see `09_Excel_Template_Mapping.md`) should
  allow new report types or new client templates to be added without rewriting core logic.
- **Safety**: the tool must never corrupt or silently overwrite the template's static
  formatting (logo, fonts, headers, print layout).

## 9. Success Criteria

- Given the sample raw file, the automated pipeline reproduces the sample final report's:
  - Row count (145 rows body, matching `Critical 21 / High 32 / Medium 77 / Low 15`).
  - Column order and headers.
  - Risk-count summary numbers.
  - Pie chart categories and values.
- A business user cannot tell, from formatting alone, whether a report was built manually or
  by the tool.
- No manual Excel steps remain in the reporting workflow other than final review/sign-off.

## 10. Out of Scope (for this documentation phase)

- Actual Python implementation (explicitly excluded per project instructions).
- Nessus scan execution / raw-export generation.
- Report delivery/distribution (email, portal upload, etc.) — unless later added to
  `10_UI_Requirements.md`.
