# 05 — Business Rules

Each rule states its **status**: `CONFIRMED` (directly verified against file evidence),
`STATED` (given directly by the business/user prompt, not independently re-derivable from the
sample data), or `INFERRED` (best-effort reconstruction — needs sign-off).

## 1. Report Type Routing

| Rule | Status |
|---|---|
| A raw export row belongs to the **VA pipeline** if `Risk ∈ {Critical, High, Medium, Low, None}` | CONFIRMED |
| A raw export row belongs to the **CA pipeline** if `Risk ∈ {PASSED, FAILED, WARNING}` | CONFIRMED |
| Any other `Risk` value is invalid and must be logged, not silently dropped or silently routed | INFERRED (safety default) |

## 2. Filtering

| Rule | Status |
|---|---|
| VA report **always excludes** `Risk = "None"` rows | CONFIRMED — 5,262 `None` rows in raw file, zero appear in final report |
| VA report **never includes** raw scanner evidence (`Plugin Output` column) | CONFIRMED — column does not appear in final template at all |
| CA report is expected to include `PASSED`, `FAILED`, and `WARNING` (i.e., not filtered the same way as VA) — but no sample CA final report was provided to confirm whether all three states are shown or only `FAILED`/`WARNING` | INFERRED — see `13_Open_Questions.md` |

## 3. Deduplication

Deduplication is a **two-stage** process, confirmed directly by the business (2026-07-08).

### Stage 1 — Exact duplicate removal

| Rule | Status |
|---|---|
| Duplicate key = exact match on (`Name`, `Description`, `Risk`, `Host`) | **CONFIRMED by business.** Rows that are byte-identical on all four fields are duplicates; keep one, drop the rest |
| Deduplication happens **per host** — i.e., the same vulnerability on two different hosts produces two separate report rows (not merged into one row with a host list) | CONFIRMED — no comma-separated/multi-value `Host` cells were found anywhere in the final report |

### Stage 2 — Version-collapse (same finding, different embedded version number)

| Rule | Status |
|---|---|
| Some findings repeat with a **version number embedded directly in the `Name`/title text** (e.g. `Adobe Flash Player <= 32.0.0.387 Multiple Vulnerabilities (APSB19-53)` vs. `...<= 32.0.0.390...`). These rows survive Stage 1 because `Name` (and usually `Description`) differ by the version substring, so they are not byte-identical | **CONFIRMED by business** — this pattern is visible in the raw file (e.g., repeated `Adobe AIR`/`Adobe Flash Player` findings across many rows) |
| Resolution rule: group rows by **base title with the version substring stripped out**, plus `Risk` and `Host`. Within each group, parse the numeric version token(s) found in `Name`, compare them, and **keep only the row with the highest (most recent) version**. All other rows in that group are dropped entirely (not merged) | **CONFIRMED by business.** "Most recent" = highest embedded numeric version string, e.g. `32.0.0.390` > `32.0.0.387`. Older-version row's data (CVE, Port, Reference, etc.) is discarded, not merged into the kept row |
| Version parsing pattern: dotted numeric version strings (e.g. `X`, `X.Y`, `X.Y.Z`, `X.Y.Z.W`), typically preceded by a comparison qualifier such as `<=`, `<`, or appearing directly after the product name | INFERRED implementation detail — see `06_Data_Transformation.md` for the exact extraction/comparison approach. Titles with **advisory/date codes instead of version numbers** (e.g. Fortinet's `FG-IR-25-756`) are a **separate, distinct case** — see Open Question in `13_Open_Questions.md` (Q4) regarding whether Stage 2 applies to them too, since the business confirmed the rule specifically for embedded version numbers, not advisory codes |
| Execution order | Stage 1 (exact dedup) always runs **before** Stage 2 (version-collapse) |

## 4. Sorting / Grouping — ✅ CONFIRMED by business (2026-07-09)

| Rule | Status |
|---|---|
| Rows are grouped into contiguous blocks by `Host` | **CONFIRMED by business.** Each unique IP forms its own contiguous block in the output |
| Within each host's block, rows are ordered by severity: **Critical > High > Medium > Low** (not alphabetical) | **CONFIRMED by business**, explicitly re-stated: "for one unique IP the list gets sorted Critical > High > Medium > Low, and similarly for the next set of IPs." This is now the authoritative rule for the whole report — every host block independently follows this same internal ordering |
| Secondary/tiebreaker sort order within the same host + risk level | INFERRED as original raw row order (i.e., a **stable sort**), pending confirmation |

## 5. Summary Calculations

| Rule | Status |
|---|---|
| Summary pivot table counts `Host` occurrences grouped by `Risk`, using the final (post-filter, post-dedup) `VA Report` data as its source — **never** the raw file | CONFIRMED — pivot cache is bound to a named range on the `Summary`/`VA Report` sheets, not the `RAW File` sheet |
| Summary pivot table must be regenerated fresh on every run; a previous run's stale pivot table must not be left in the workbook | CONFIRMED AS A DEFECT TO AVOID — sample workbook contains exactly this defect (an old pivot table with Grand Total 194 sitting alongside the current one with Grand Total 145) |
| "Grand Total" in the pivot table must always equal the total row count of the `VA Report` data table | CONFIRMED (for the *current*, non-stale pivot table: 145 = 145) |

## 6. IP / Host Handling

| Rule | Status |
|---|---|
| The `Summary!A7:C26` scope table lists each **distinct** host that appears anywhere in the VA Report data | INFERRED (reasonable, not independently falsifiable from the anonymized sample since every row happens to have a unique host in this file) |
| `Scan Type` (e.g., `Authenticated`) and `Device Type` (e.g., `Server`, `Firewall`) per host are **not derivable from the raw scanner export** and must be supplied as external engagement metadata | CONFIRMED — no column in the raw file maps to either of these values |

## 7. Scope Generation

| Rule | Status |
|---|---|
| "Scope" for a given report = the distinct set of hosts that produced at least one VA finding after filtering/dedup | INFERRED |
| Hosts scanned but producing **zero** findings after filtering are still expected to appear in the scope table (a clean host is still "in scope") — **cannot be confirmed from the sample**, since the sample's scope table (19 hosts) and finding-hosts (also apparently ~19-20 distinct hosts in the "current" pivot table's total-145 group) roughly line up but this could be coincidental | INFERRED — flagged for confirmation |

## 8. Scan Type (VA vs. CA) Naming

| Rule | Status |
|---|---|
| File name embeds the report type and scope, e.g. `VA_Server_First_Audit_Report_...` = **V**ulnerability **A**ssessment + scope `Server` + phase `First` | INFERRED from filename pattern |
| Client entity is represented by both a full legal name and short entity codes, e.g., `Trust Investment Advisors Private Limited` + `TSS`/`SCPL` — exact meaning of the two entity codes (are these two related group companies both in scope, or an internal engagement code + a separate abbreviation?) is **not confirmed** | OPEN QUESTION — see `13_Open_Questions.md` |

## 9. Risk Ordering (Explicit Restatement)

Per direct business instruction, the canonical severity ranking used throughout is:

```
Critical  >  High  >  Medium  >  Low
```

This is **not** alphabetical (`Critical, High, Low, Medium` would be alphabetical) and must be
implemented as an explicit ordinal mapping (e.g., `{Critical: 0, High: 1, Medium: 2, Low: 3}`),
never a default string sort.

## 10. Chart Generation

| Rule | Status |
|---|---|
| One pie chart per report, sourced from the risk-count pivot table, categories = `Critical/High/Medium/Low`, values = count of findings | CONFIRMED |
| Chart colors follow the workbook's theme accent colors (pivot-format-driven, `<c:pivotFmt>` entries in `chart1.xml` assign a distinct accent/solid color per category index) | CONFIRMED |
| No conditional/severity-specific fixed color mapping (e.g., "Critical is always red") was found encoded in the chart XML — colors come from theme accent order, not a fixed severity-to-color dictionary | CONFIRMED — flagged in case business actually wants fixed severity colors (common convention) rather than theme-accent-order colors; see Open Questions |

## 11. Naming Conventions

### Report file naming pattern (inferred from sample filename)

```
<ReportType>_<Scope>_<Phase>_Audit_Report_<ClientLegalName>_<EntityCode(s)>_<Year>_V<Major>.<Minor>
```

Example (sample): `VA_Server_First_Audit_Report_Trust_Investment_Advisors_Private_Limited_TSS__SCPL_2026_V1_0`

| Token | Sample value | Notes |
|---|---|---|
| `ReportType` | `VA` | Expected alternate value: `CA` |
| `Scope` | `Server` | Other likely values: `Firewall`, `Network`, `Web`, etc. — inferred from `Summary` sheet's `Device Type` values |
| `Phase` | `First` | Suggests engagements can have multiple phases/rounds (e.g., "First", "Retest", "Second") |
| `ClientLegalName` | `Trust Investment Advisors Private Limited` | Spaces replaced with underscores |
| `EntityCode(s)` | `TSS`, `SCPL` (double underscore before `SCPL` — likely a typo/artifact, not intentional) | Meaning unconfirmed — see Open Questions |
| `Year` | `2026` | |
| Version | `V1_0` | Underscore used instead of a period in the filename (though the in-sheet cell value for Report Version is `1.2`, a numeric, showing the file was renamed at least once without updating the in-sheet version field — another data-hygiene defect worth flagging) |

### Report Date Formatting

| Rule | Status |
|---|---|
| `Report Date` cell (`VA Report!C8`) is stored as a native Excel date (`datetime`), not a text string | CONFIRMED |
| Display format was not independently re-derived (number format string not captured in this pass) — assume default Excel short date or the workbook's existing number format is preserved as-is when automation writes to this cell | INFERRED |

## 12. Everything Else Observed

- The manual process pastes **values**, not live formulas, from the raw file into the `VA
  Report` sheet (no cross-sheet formulas were found referencing `RAW File` anywhere in the
  workbook). Automation should replicate this "flatten to values" behavior rather than writing
  live formulas, to preserve exact fidelity with how the client is used to receiving the file.
- The `Report Version` field (`1.2`) is a plain float, not a `"V1.2"` formatted string — but the
  filename uses the `V1_0` convention. These two version representations are **not the same
  field** and must be tracked/entered separately by automation, not derived from one another.
