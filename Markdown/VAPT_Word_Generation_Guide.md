# VAPT Report Generation Guide (v1.1 — corrected)

> **Change log (this revision):** Added §4.4.1, a mandatory single-insertion
> guardrail for the summary table. Diagnosed from a real failure: the
> generator called the table-building routine twice for this section,
> producing two identical `Category/Value` tables both floating to the exact
> same page anchor (`tblpX=10645, tblpY=4133`), which rendered as an
> overlapping, misaligned mess with a stray page break between them. See
> `Pie_Chart_Placement_Spec.md` v1.1 for the matching chart-side fix — that
> revision addresses a *second*, independent bug found in the same run: the
> pie chart was generated as a PNG image and its relationship was registered
> but never actually referenced by any `<w:drawing>`, so no chart appeared at
> all. Both bugs were present simultaneously in the same generated document;
> fix both, and verify both, every time.

General instructions for turning uploaded **VA (Vulnerability Assessment)** and
**CA (Configuration Audit)** Excel workbooks into the final Word VAPT report.
This is a reusable spec — it does not assume any particular client, IP range,
or finding counts. Apply it every time new Excel files are uploaded.

---

## 1. Expected input files

For a given engagement, you may receive VA and/or CA workbooks for one or more
asset types, typically named with a pattern like:

```
VA_<AssetType>_...xlsx                     e.g. VA_Network, VA_Server, VA_Endpoint
Configuration_Audit_<AssetType>_...xlsx    e.g. CA_Network, CA_Server, CA_Endpoint
```

Asset types commonly seen: `Network` (network devices), `Server`, `Endpoint` —
but the same rules apply to any asset‑type grouping supplied.

Each workbook has three sheets:

| Sheet | Purpose |
|---|---|
| `Introduction` | Legend / how‑to‑read notes — not report content, ignore |
| `VA Report` (VA files) or `CA_Report` (CA files) | The finding‑by‑finding detail table |
| `Summary` | The in‑scope IP list + a severity/status pivot table |

---

## 2. Reading the detail‑table sheet (`VA Report` / `CA_Report`)

- Rows near the top contain a metadata block (`Client Name:`, `Security Tester:`,
  `Reviewed By:`, `Report Date:`, `Report Version:`, `Scanner:`) — **ignore this**.
- Find the real header row by searching for the header text, not by assuming a
  fixed row number (it has been row 13 in practice, but confirm each time).
- Data rows continue below the header until the first fully blank row.

### VA Report → Word columns

`Sr. no → Sr no`, `Vulnerbility Title`, `Description`, `Risk`, `Host`, `Port`,
`Recommendation`, `Reference`, `CVE`.

### CA_Report → Word columns

`Sr. no → Sr no`, `Title`, `Description`, `Host`, `Recommendation`, `Status`
(value is `FAILED` or `WARNING`).

---

## 3. Reading the `Summary` sheet (IP scope + pie‑chart data)

Every `Summary` sheet contains two blocks, but **their row/column position is
not fixed and varies from file to file.** Never hard‑code cell addresses —
always locate blocks by their label text.

### 3.1 IP‑in‑scope block

1. Search the sheet for the text `"List of Ips in scope"` (case‑insensitive).
2. Below it, find the real header row by searching for a cell containing
   `"IP Address"`. Use that column as the anchor.
3. Read rows downward — columns are `Sr no | IP Address | Device Type` — until
   the `IP Address` cell is blank.
4. `Device Type` is often only filled on the **first row of each group**
   (blank on subsequent rows of the same type) — carry the last non‑blank value
   forward to blank rows below it.
5. Don't assume a fixed starting column — some files start the IP block in
   column A, others may be shifted right by one or more columns.
6. Don't assume exact spelling/casing of `Device Type` values (e.g. typos like
   "Enpoint" instead of "Endpoint" can occur in source data) — treat them as
   free text and reproduce as‑is rather than validating against a fixed list.

### 3.2 Severity / status pivot block (feeds the summary table + pie chart)

1. Search the sheet for a cell containing exactly `"Row Labels"`. The column
   immediately to its right is `Count of Host`.
2. Read rows downward until the label is blank or reads `Grand Total`.
3. For VA files, possible labels: `Critical`, `High`, `Medium`, `Low`, `Info`.
   For CA files, possible labels: `FAILED`, `WARNING`.
4. **Any label not present in the pivot = a count of 0** for that category —
   pivot tables omit rows whose count is zero, so a missing label is not an
   error, it's a zero.
5. A `Grand Total` row may or may not be present — don't rely on it; compute
   your own sum from the individual category rows instead.

---

## 4. Word report structure to populate

### 4.1 Engagement Scope section

Build two side‑by‑side tables per audit type (VA and CA), each with:

- **Server + Endpoint table** — IP list rows from the Server‑type workbook's
  Summary sheet, followed by the Endpoint‑type workbook's Summary sheet.
- **Network Devices table** — IP list rows from the Network‑type workbook's
  Summary sheet.

Repeat once for VA workbooks and once for CA workbooks. Columns shown: `IP
Address`, `Device Type` (Sr no is not required in this section).

### 4.2 Vulnerabilities Details section (normally the last section)

Fixed sequence of elements, **each element inserted exactly once**:

1. Bullet: **"Vulnerability Assessment and Configuration Audit of Assets"**
2. Auto‑generated summary sentence (§4.3)
3. **Pie chart + summary table**, placed side by side (§4.4) — one table, one
   chart, no repeats. See §4.4.1.
4. Bullet: **"The below table shows the detailed report of the VA scan done on assets."**
5. VA detail table — all VA findings from every VA workbook, concatenated (§4.5)
6. Bullet: **"The below table shows the detailed report of the Compliance scan done on Assets."**
7. CA detail table — all CA findings from every CA workbook, concatenated (§4.5)

### 4.3 Summary sentence template

> The audit reported **{Critical} Critical, {High} High, {Medium} Medium,
> {Low} Low risk, {Info} informational** vulnerabilities and **{Failed} Failed,
> {Warning} Warning** compliance issue on the target asset. Following table and
> graph summarizes overall risk of target infrastructure.

Each `{value}` = the sum of that category across **all** VA (or CA) workbooks
uploaded for this engagement (see §5).

### 4.4 Summary table + pie chart

Summary table rows, in this order:

| Category | Value |
|---|---|
| Critical | combined VA Critical count |
| High | combined VA High count |
| Medium | combined VA Medium count |
| Low | combined VA Low count |
| Info | combined VA Info count |
| Failed | combined CA FAILED count |
| Compliance | combined CA WARNING count |
| Total | sum of all rows above |

The **pie chart must be a real embedded Word/Office chart object** (native
OOXML pie chart with its own chart part and data), not a picture or
screenshot. Its slices use the same category values as the summary table
(`Critical, High, Medium, Low, Failed, Compliance`; `Info` may be included as a
0‑value slice if not zero). Place it directly beside/after the summary table.
Colour slices/legend consistently with whatever shading the report template
uses for each severity, so the chart and table read as a matched pair.
Category labels in the chart legend must match the summary table's category
names **verbatim** (e.g. `Compliance`, not `Warning`) — the two are read
side by side and a label mismatch reads as an error even when the numbers
are right.

#### 4.4.1 Insert this table and chart exactly once — mandatory guardrail

**This is the single most common failure mode in this pipeline and must be
actively guarded against, not just avoided by care.**

A prior run of this pipeline produced a document with the summary table
inserted **twice in a row**, both copies floating to the identical page
anchor (`tblpX`/`tblpY` — see `Pie_Chart_Placement_Spec.md` §3), with a
spurious extra page break between them. The two tables had byte‑identical
content. Root cause was a code path that called the "build summary
table + pair with chart" routine more than once for what is supposed to be a
single section — e.g. once per aggregation pass and again during
chart‑pairing, or once per workbook type instead of once for the whole merged
totals.

To prevent recurrence:

- The summary table and its paired chart belong to the merged, engagement‑wide
  totals (§5) — **not** to any individual workbook or asset type. There is
  exactly one such table and one such chart in the entire document, ever.
  Never call the table/chart insertion function inside a per‑workbook or
  per‑asset‑type loop.
- Treat table‑insertion as an idempotent, single‑call operation for this
  section: if your code has more than one call site that could plausibly
  insert this table (e.g. a "build layout" step and a separate "insert chart"
  step that also inserts the table it pairs with), consolidate them into one
  call, or add an explicit flag/assertion that raises an error if the routine
  runs a second time in the same generation pass.
- **After generation, programmatically count** `w:tblpPr` elements whose
  `tblpX`/`tblpY` match the summary‑table anchor coordinates in the output
  `document.xml`. There must be **exactly one**. Fail the build if the count
  is 0 or ≥2 — do not rely on visual inspection alone to catch this, since a
  perfectly-aligned single table and two identically-misaligned tables can
  look similar in a quick preview but are structurally very different.
- Do not insert extra manual page breaks around this section beyond the one
  specified in `Pie_Chart_Placement_Spec.md` §2/§4 (immediately after the
  chart paragraph). Extra page breaks before the table or between the table
  and chart are themselves a signal that something is being emitted more than
  once.

### 4.5 VA and CA detail tables

- **VA detail table columns**: `Sr no, Vulnerbility Title, Description, Risk,
  Host, Port, Recommendation, Reference, CVE`.
- **CA detail table columns**: `Sr no, Title, Description, Host,
  Recommendation, Status`.

Build each by reading every relevant workbook's detail sheet (§2) and
appending rows one after another, in a consistent asset‑type order (e.g.
Network → Server → Endpoint) applied the same way every time.

**Re‑number `Sr no` sequentially (1, 2, 3, …) across the merged table** — do
not keep each source workbook's own numbering, or the combined table will
contain duplicate `Sr no` values.

For the CA table, shade the `Status` cell background by value (e.g. one fixed
colour for `FAILED`, another for `WARNING`), applied consistently across the
whole merged table regardless of which source workbook a row came from.

---

## 5. Aggregation rules

- VA totals (`Critical, High, Medium, Low, Info`): sum each category's pivot
  count (§3.2) across **all** VA workbooks supplied for the engagement.
- CA totals (`FAILED, WARNING`): sum each category's pivot count across
  **all** CA workbooks supplied.
- `Total` = Critical + High + Medium + Low + Info + Failed + Warning.
- Cross‑check: `Total` should equal the total number of data rows across all
  VA detail sheets plus all CA detail sheets combined. If it doesn't match,
  re‑check the pivot read and the detail‑table read for missed rows.
- These totals are computed **once** for the whole engagement and feed **one**
  summary table and **one** chart (§4.4.1). Do not recompute-and-reinsert per
  workbook.

---

## 6. Common pitfalls to guard against

- **Never use fixed cell coordinates** for the `Summary` sheet blocks — always
  locate by label text (`"List of Ips in scope"`, `"IP Address"`, `"Row
  Labels"`). Row/column position varies file to file, including occasional
  whole‑column shifts.
- **Missing pivot rows mean zero**, not "skip this category" — every severity/
  status must still appear in the summary table and chart, even at 0.
- **A pie chart that is a static image is not acceptable** if the original
  template uses a native embedded chart object — check the target template
  and match its chart type (native OOXML vs. image). See
  `Pie_Chart_Placement_Spec.md` §1a for a real failure case where an image
  was generated (e.g. via a plotting library) and its relationship registered
  in the `.docx`, but the `<w:drawing>` that would actually place it in the
  document body was never written — resulting in **no visible chart at all**,
  not merely the wrong chart type. Registering a media relationship is not
  the same as inserting it; both steps are required, and only the native
  chart approach should be used regardless.
- **Don't carry forward per‑workbook `Sr no` values** into the merged detail
  tables — always renumber sequentially after merging.
- **Don't call the summary‑table/chart insertion logic more than once** for
  the Vulnerabilities Details section — see §4.4.1.
- Tolerate minor inconsistencies in source data (typos in `Device Type`,
  inconsistent header capitalisation like `Sr no` vs `Sr No`) rather than
  failing to parse — match on the semantic label, not exact string equality.

---

## 7. Validation checklist (run after generating any report)

- [ ] Every IP listed in each workbook's Summary "List of Ips in scope" block
      appears in the corresponding Engagement Scope table.
- [ ] Summary sentence numbers match the summary table exactly.
- [ ] Summary table `Total` equals the sum of its own rows and matches the
      cross‑check in §5.
- [ ] **Exactly one** `w:tblpPr` exists in the output `document.xml` at the
      summary‑table anchor coordinates — count them programmatically, not by
      eye. Zero or more than one is a build failure. (§4.4.1)
- [ ] A real chart object exists in the generated `.docx` (not an image),
      referenced from the document body — search for `<c:chart` in
      `document.xml` and confirm the referenced `word/charts/chartN.xml` part
      exists and is non-empty. A relationship/media file existing is **not**
      sufficient proof the chart is present; confirm the `<w:drawing>` that
      references it actually appears in `document.xml`.
- [ ] Pie chart slice values match the summary table row values exactly, and
      slice/legend labels match the summary table's category names verbatim.
- [ ] VA detail table row count = total data rows across all VA detail sheets.
- [ ] CA detail table row count = total data rows across all CA detail sheets.
- [ ] `Sr no` in both merged detail tables is sequential with no repeats.
