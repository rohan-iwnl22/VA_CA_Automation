# 13 — Open Questions

Every item below could not be conclusively inferred from the supplied files/recording. Each is
tracked as: Question / Reason / Possible Assumptions / Impact / Priority.

---

### Q1. CA report filtering policy — include all statuses, or exceptions-only?

- **Reason**: No finished CA report sample was supplied, only raw `PASSED`/`FAILED`/`WARNING`
  rows. VA excludes its "clean" state (`None`); it's unclear if CA should analogously exclude
  `PASSED`, or intentionally include it for a full compliance posture view.
- **Possible assumptions**: (a) Include all 3 states (full compliance report); (b) Include only
  `FAILED`/`WARNING` (exceptions-only, mirroring VA's exclusion pattern).
- **Impact**: Changes report size (~5,652 vs. ~2,895 rows) and overall tone/purpose of the CA
  deliverable. Blocks `08_CA_Workflow.md` from being finalized.
- **Priority**: **High** — blocks CA pipeline build.

### Q2. Blank/pristine master template — ✅ RESOLVED (2026-07-09)

- **Resolution**: Business supplied the actual blank master template
  (`Blank_Template_VA_Server_First_Audit_Report_Trust_Investment_Advisors_Private_Limited_TSS__SCPL_2026_V1_0.xlsx`).
  Directly verified: same 3-sheet structure, byte-identical logo, identical fonts/column widths;
  `VA Report` rows 14+ empty; `Summary` scope table rows empty; **zero pivot tables and zero
  charts anywhere in the file** (only the two logo drawing anchors exist).
- **New finding surfaced by this file (not previously knowable from the completed sample
  alone)**: because the pristine template has no pivot/chart objects at all, the pivot table and
  pie chart must be **created from scratch on every run**, not refreshed from an existing
  object. Updated in `07_VA_Workflow.md`, `09_Excel_Template_Mapping.md`, and
  `11_Project_Architecture.md`.
- **Minor nuance**: `Introduction!B14`/`B15` (Scanner Name / Scanner Signature Version) are
  **pre-filled with default values** (`"Nessus "` / `"10.11.4"`) in the blank master rather than
  empty — treat these as overridable defaults, not required inputs, in the engagement metadata
  form (`10_UI_Requirements.md`).
- **Priority**: Resolved — no longer a blocker.

### Q3. True sort key — ✅ RESOLVED (2026-07-09)

- **Resolution**: Business explicitly re-confirmed the exact behavior: *"the custom sort happens
  in the form like after the sort, for one unique IP the list gets sorted in the form Critical >
  High > Medium > Low, and similar for the next set of IPs."* This matches the originally stated
  rule exactly — each host forms a contiguous block, and every block independently follows
  Critical → High → Medium → Low internally.
- The sample file's apparent inconsistency is attributed to the `Host` column having been
  anonymized/renumbered (`192.x.x.N`, sequential by row) before the sample was shared, which
  disturbed the visible grouping without changing the underlying rule.
- Implemented in `05_Business_Rules.md` §4, `06_Data_Transformation.md` (Sorting / Ordering
  Pipeline), and `07_VA_Workflow.md` (Sorting Recap).
- **Remaining spin-off (new, low priority):** the order of the **host blocks themselves**
  (first-seen order in the filtered/deduped data vs. ascending numeric IP order) is not yet
  confirmed — see Q11 below, which absorbs this remaining nuance.

### Q4. Exact deduplication key — ✅ RESOLVED (2026-07-09)

- **Resolution**: Business confirmed a **two-stage** dedup process:
  1. Stage 1 (exact): key = `(Name, Description, Risk, Host)`.
  2. Stage 2 (version-collapse): for rows sharing `(base_title_with_version_stripped, Risk,
     Host)`, keep only the row with the highest embedded numeric version found in `Name`; drop
     the rest entirely (no merging of `CVE`/`Port`/`Reference`).
- See `05_Business_Rules.md` §3 and `06_Data_Transformation.md` for the full specification and
  worked example.
- **Remaining sub-question (new, spun off from this one):** Stage 2 currently only fires when
  `Name` contains a parsable **numeric version token**. Titles that instead vary by an
  **advisory/bulletin code** (e.g., Fortinet's `FG-IR-25-756` vs. `FG-IR-23-137` — these are
  genuinely different advisories, not versions, in the sample data) are explicitly **out of
  scope** per the business's answer ("version number in the title," not "advisory code"). This
  is documented as intentional, not a gap — but should be revisited if a client's raw export
  turns out to have finding titles that repeat via advisory-code drift only (no version number)
  in a way that should also be collapsed. **Priority: Low** (confirmed out-of-scope unless new
  evidence emerges).

### Q5. Meaning of the `TSS` / `SCPL` entity codes in the filename

- **Reason**: Filename contains `Trust_Investment_Advisors_Private_Limited_TSS__SCPL` — unclear
  whether these represent two related group companies both in scope of one combined report, an
  internal project/engagement code, or something else. The double-underscore before `SCPL` may
  also be an accidental typo in the original filename rather than an intentional token separator.
- **Possible assumptions**: Treat as two free-text "entity code" tokens supplied by the user at
  generation time, with no further business logic attached.
- **Impact**: Low functional risk (filename cosmetic only) but could matter if entity codes
  drive any downstream routing/branding logic not yet observed.
- **Priority**: **Low**.

### Q6. Print layout / PDF export requirements

- **Reason**: The sample workbook has no custom print area, orientation, or fit-to-page
  settings — meaning either printing was never a requirement, or the sample simply wasn't
  finalized for print before being handed over.
- **Possible assumptions**: No print-specific setup is required for v1; reports are consumed as
  `.xlsx` only.
- **Impact**: If clients actually receive a printed/PDF version, missing print setup would
  produce a poorly paginated document.
- **Priority**: **Medium**.

### Q7. Fixed severity color-coding — is it actually wanted?

- **Reason**: No conditional formatting or per-severity fill exists in the sample, which is
  unusual for this report category (most VA reports color Critical=red, High=orange, etc.).
  This could mean the business genuinely doesn't want it (matches sample exactly), or that the
  colorization step exists in the real manual process but happened to be stripped/lost before
  this "Copy" file was saved.
- **Possible assumptions**: No coloring, per documented evidence.
- **Impact**: A visually "flatter" report than industry norm if the assumption is wrong; a
  wrongly-added feature not matching the client's established branding if added without
  confirmation.
- **Priority**: **Medium** — easy to add later as a config toggle if confirmed wanted (see
  `11_Project_Architecture.md`, `business_rules.yaml`).

### Q8. Native PivotTable/PivotChart vs. static equivalent

- **Reason**: See `11_Project_Architecture.md`'s architecture decision — building a truly
  native, Excel-refreshable pivot in a non-Excel Python library carries real implementation risk
  and cross-platform limitations.
- **Possible assumptions**: Ship a static aggregation + static pie chart styled to match for v1.
- **Impact**: A sufficiently technical client could tell it's not a "real" pivot if they inspect
  the file (right-click won't show "Refresh"), though visually it will be indistinguishable.
- **Priority**: **Medium**.

### Q9. Multi-CVE cell formatting

- **Reason**: Raw `CVE` column may contain multiple CVE IDs in one cell; exact expected
  delimiter/line-break convention in the final report column wasn't confirmed from the sample
  (most sample rows have a single CVE or none).
- **Possible assumptions**: Preserve raw cell content verbatim, no reformatting.
- **Priority**: **Low**.

### Q10. Scope table completeness for zero-finding hosts

- **Reason**: Unclear whether a scanned host that produced **zero** VA findings after
  filtering/dedup should still appear in the `Summary` scope table (a "clean" host is still
  technically in scope of testing).
- **Possible assumptions**: Yes, include all scanned hosts regardless of finding count — this
  requires the raw file (or metadata) to identify the **full scanned host list**, not just hosts
  that happen to have surviving findings, which may need an additional data source not yet
  identified.
- **Impact**: If wrong, the scope table under-represents what was actually tested.
- **Priority**: **Medium**.

### Q11. Host grouping order in the scope table and report body

- **Reason**: Unclear if hosts should be ordered by first-appearance in the filtered data, or by
  ascending numeric IP order, in both the `VA Report` body and the `Summary` scope table.
- **Possible assumptions**: Ascending numeric IP order (most common convention for this report
  type), pending confirmation.
- **Priority**: **Low-Medium**.

### Q12. Desktop UI vs. CLI-only for v1 — ✅ RESOLVED (2026-07-09)

- **Resolution**: Business confirmed sequencing: **implement core pipeline functions first**
  (ingestion, filtering, dedup, sort, template write, summary/chart, save — per
  `07_VA_Workflow.md` / `11_Project_Architecture.md`), with the desktop UI (`10_UI_Requirements.md`)
  explicitly deferred to a later phase, built only once the pipeline is functionally complete
  and tested. `10_UI_Requirements.md` remains fully specified so no re-analysis is needed when
  that phase starts, but nothing in it should be implemented now.
- **Impact**: No change to pipeline logic; affects only project sequencing. `ui/` folder in
  `11_Project_Architecture.md`'s structure stays a placeholder until the UI phase begins.
- **Priority**: Resolved — no longer open.

### Q13. Report Date / Report Version display formats

- **Reason**: Excel's underlying number format string for the `Report Date` cell was not
  captured in this analysis pass; the `Report Version` field is a numeric (`1.2`) while the
  filename separately uses a different convention (`V1_0`) — it's unclear if these two are
  meant to always stay in sync or are genuinely independent fields (the sample shows them
  **out of sync**, `1.2` vs. `V1_0`, which may itself be a data-entry mistake in the sample
  rather than an intentional distinction).
- **Possible assumptions**: Treat as two independently entered fields; do not attempt to
  auto-derive one from the other.
- **Priority**: **Low**.

---

## Summary Priority Table

| # | Question | Priority |
|---|---|---|
| Q1 | CA filtering policy | High |
| Q2 | Blank master template | ✅ Resolved |
| Q3 | Sort-key inconsistency in sample | ✅ Resolved (residual nuance folded into Q11) |
| Q4 | Exact dedup key | ✅ Resolved (see Q4 entry for a low-priority spin-off sub-question) |
| Q5 | Meaning of TSS/SCPL codes | Low |
| Q6 | Print/PDF layout requirements | Medium |
| Q7 | Fixed severity color-coding | Medium |
| Q8 | Native pivot vs. static chart | Medium |
| Q9 | Multi-CVE cell formatting | Low |
| Q10 | Zero-finding hosts in scope table | Medium |
| Q11 | Host ordering convention | Low-Medium |
| Q12 | UI vs. CLI-only for v1 | ✅ Resolved |
| Q13 | Date/version format sync | Low |

**Recommendation:** Per business direction (2026-07-09), **implementation is proceeding VA-only
for now** — Q1 (CA filtering policy) is deferred until the CA phase begins. **Q2, Q3, Q4, and
Q12 are all now resolved.** Per Q12's resolution, the build is further sequenced as **core
pipeline functions first, UI later** — `10_UI_Requirements.md` is explicitly out of scope for
the current build phase. No remaining high-priority blockers exist for starting the VA pipeline
build. The only open items are low/medium priority (Q6, Q7, Q8, Q9, Q10, Q11, Q13) and can be
resolved iteratively during build/QA without blocking the start of development. The VA pipeline
specification (`07_VA_Workflow.md`, `05_Business_Rules.md`, `06_Data_Transformation.md`,
`09_Excel_Template_Mapping.md`, `11_Project_Architecture.md`) is considered ready for
implementation to begin.
