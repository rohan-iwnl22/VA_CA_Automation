# Pie Chart Placement — Build Spec (v1.1 — corrected)

> **Change log (this revision):** Added §1a. Diagnosed from a real generator
> run whose output (`current.docx`) had **zero** visible pie chart and two
> overlapping duplicate summary tables instead. Root cause on the chart side:
> the pie chart was rendered as a PNG image (via a plotting library), the PNG
> was saved into `word/media/` and a relationship (`r:id`) was registered for
> it in `document.xml.rels` — but no `<w:drawing>` anywhere in `document.xml`
> actually referenced that relationship. Registering a relationship is not
> the same as placing it in the document; the chart was silently dropped.
> This is a *different* bug from "chart is an image instead of native OOXML"
> (still also forbidden, see §1) — it's "chart is missing entirely, and
> the image-generation code path is why." See
> `VAPT_Word_Generation_Guide.md` §4.4.1 for the matching table-side bug
> (duplicate table insertion) found in the same document.

Scope: **only** the pie chart in the "Vulnerabilities Details" section — how many
to generate, where exactly it goes, and how it must relate to the summary table
next to it. Based on directly inspecting the OOXML (`document.xml`) of the
reference report, not a screenshot guess.

---

## 1. Diagnosis: why 2 charts are appearing (reference file)

The reference file itself actually contains **two chart objects**, but only
**one of them is the correct one**:

| Chart | Where it sits in the document | Verdict |
|---|---|---|
| The **correct** chart (paired with the summary table) | Inline, in the paragraph immediately after the summary table | ✅ Keep — this is the one visible beside the table |
| A **second, stray** chart | Sits earlier, wrapped in a run styled with **white font colour** and preceded by ~7 empty paragraphs, on its own effectively-blank page | ❌ Duplicate/mistake — do not replicate this |

If your generator is producing 2 pie charts, it is very likely reproducing both
of these instead of only the correct one. **Only ever insert one chart object**
for this section — the one built per §2 below.

## 1a. Diagnosis: why 0 charts appeared (a different, more severe failure)

In a separate generator run, the output had **no chart at all** — not two,
not one, zero — even though the generator clearly *attempted* to produce one:
`word/media/image2.png` existed in the package and contained a correctly
computed pie chart (right categories, right values), and
`word/_rels/document.xml.rels` had a valid image relationship pointing to it.
But grepping `document.xml` for that relationship's `r:id` found **no
matches** — no `<w:drawing>`, `<w:blip>`, or anything else in the document
body referenced it. The chart was fully computed and saved to the package,
then never wired into the visible document.

This is worse than generating the wrong chart type, and it's easy to miss in
review because the file size and relationship count look "normal" — the
missing piece is invisible unless you specifically check that every image/
chart relationship is actually referenced from `document.xml`.

**Do not build the pie chart this way at all.** Concretely:

- Do not generate the chart with a plotting/imaging library (matplotlib,
  PIL, etc.) and save it as a PNG/JPEG under `word/media/`, even as an
  intermediate step. That code path is what produced both the "chart is an
  image" defect (§1, forbidden by the main guide) and the "chart never got
  inserted" defect above — once a chart exists only as a media file, it is
  one missed `<w:drawing>` away from silently vanishing, and there is no
  visual signal in a quick preview that distinguishes "no chart was built"
  from "chart was built but not linked."
- Build the chart the same way the reference file does: as a native chart
  part (`word/charts/chartN.xml` + its `colorsN.xml`/`styleN.xml` +
  relationship), inserted as an inline `<w:drawing>` referencing
  `<c:chart r:id="...">` per §4 below, in the same operation that creates the
  chart part — never as a two-step "generate image, then (maybe) insert it
  later" process where the second step can be skipped.
- After generation, verify by relationship, not just by file presence: for
  every entry in `word/_rels/document.xml.rels` whose type is `image` or
  `chart`, confirm its `r:id` string literally appears inside a `<w:drawing>`
  in `document.xml`. An unreferenced relationship of either type means
  something was generated and abandoned — treat it as a build failure, not a
  harmless leftover.

---

## 2. Correct structure (in document order)

```
[Paragraph] "...compliance issue on the target asset. Following table and
             graph summarizes overall risk of target infrastructure."
[Table]      the Critical/High/Medium/Low/Info/Failed/Compliance/Total table
             — this table is a FLOATING table (see §3), so it does NOT take up
             a line in the normal text flow. Inserted exactly ONCE — see
             VAPT_Word_Generation_Guide.md §4.4.1 for a real duplicate-
             insertion failure to guard against.
[Paragraph]  contains exactly ONE inline chart object (the pie chart, §4)
[Page break] immediately after that same paragraph
[Bullet]     "The below table shows the detailed report of the VA scan done on assets."
```

Because the summary table is floating/anchored to an absolute page position
(not inline), the pie chart — placed as the very next normal paragraph — ends
up rendering in the remaining column to its left, which is what produces the
side‑by‑side "chart on the left, numbers table on the right" look. You do not
need a 2‑column layout table to achieve this; you need the table to float and
the chart to be the next inline paragraph.

**Do not insert any extra manual page breaks in this sequence** beyond the
single one shown above. Extra page breaks before the table, or between the
table and the chart paragraph, are a symptom of the table or chart being
emitted more than once (see §1a and the table-side guardrail in the main
guide) — they should not appear even as a "just in case" spacing fix.

---

## 3. Summary table — exact floating position

Reproduce these table properties so the table anchors at the same spot on the
page every time (values taken directly from the reference file's table
properties):

```xml
<w:tblPr>
  <w:tblpPr w:leftFromText="180" w:rightFromText="180"
            w:vertAnchor="page" w:horzAnchor="page"
            w:tblpX="10645" w:tblpY="4133"/>
  <w:tblW w:w="3235" w:type="dxa"/>
</w:tblPr>
```

- `horzAnchor="page"` / `vertAnchor="page"` → position is measured from the
  **page edge**, not the margin.
- `tblpX="10645"` twips ≈ 7.39in from the left page edge.
- `tblpY="4133"` twips ≈ 2.87in from the top page edge.
- `leftFromText="180"` / `rightFromText="180"` (180 twips ≈ 0.125in) → this is
  the wrap margin that pushes the chart paragraph's text/graphic away from the
  table, which is what creates the visual gap between chart and table.
- Table width `w:w="3235"` (twips) ≈ 2.25in — keep the table narrow; it's a
  small 2‑column key/value table (category name + count), not a wide table.
- Page for this section is **A4 landscape** (`w:pgSz w:w="16838" w:h="11906"
  w:orient="landscape"`) with 1in margins on top/left/right — the tblpX/tblpY
  values above assume this page size. If your template uses a different page
  size, keep the table anchored to the **top‑right of the printable area**
  rather than reusing the raw twip numbers verbatim.
- **These exact coordinates should appear on exactly one `w:tblpPr` element in
  the whole document.** If your validation finds two tables sharing this
  anchor, that is not "the table rendering twice as a visual effect" — it is
  two separate `<w:tbl>` elements in the XML, both fully populated, silently
  stacked on top of each other. Fix the call site that emits the table, don't
  try to reposition the duplicate.

If your automation builds this table as an ordinary inline table instead of a
floating one, that is likely also why the layout looks wrong even before you
get to the chart — fix the table to float first, then place the chart.

---

## 4. The pie chart itself

Insert exactly **one** native OOXML chart (`c:chart` referencing a real
`chart#.xml` part with its own embedded data — not a picture) as an **inline**
drawing in the single paragraph that comes right after the table:

```xml
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="5461000" cy="3441700"/>
        <wp:docPr id="..." name="Chart"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="rIdXX"/>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
  <w:r><w:br w:type="page"/></w:r>
</w:p>
```

Note that the `<c:chart r:id="rIdXX"/>` line above is precisely the element
that was missing in the failure described in §1a — the relationship existed,
but this reference to it, inside a `<w:drawing>` inside the document body,
did not. Treat this snippet as the thing to grep for post-generation, not
just the existence of a relationship or a media file.

- **Size:** `cx="5461000" cy="3441700"` EMU = **5.97in × 3.76in**. Use this
  exact size to match the reference layout.
- **Type:** pie chart (`<c:pieChart>`).
- **Categories/values:** must be the **same 6–7 rows and same numbers, with
  the same category names**, as the summary table next to it (`Critical,
  High, Medium, Low, [Info,] Failed, Compliance`) — the chart's embedded
  worksheet data and the table's cell values should never diverge, and
  category label text (e.g. `Compliance`, not a synonym like `Warning`)
  should match the table verbatim.
- **Page break:** put the manual page break in the same paragraph, right after
  the chart run, so the next bullet ("below table shows the detailed report of
  the VA scan...") always starts on a fresh page.

---

## 5. Checklist before finalizing

- [ ] Search the generated `.docx` for `<c:chart` — it must appear **exactly
      once** in this section (i.e. exactly one `word/charts/chartN.xml` part
      referenced from this part of `document.xml`).
- [ ] No stray chart run exists elsewhere with hidden/white‑coloured text or
      sitting after a run of empty paragraphs — that pattern is the duplicate
      bug, not a template feature.
- [ ] Every `image`/`chart` relationship in `word/_rels/document.xml.rels` has
      its `r:id` actually referenced by a `<w:drawing>` in `document.xml`. An
      orphaned relationship (registered but unreferenced) means a chart or
      image was generated and then dropped — this is a build failure, not
      inert leftover data. (§1a)
- [ ] The summary table has `tblpPr` (floating), not a plain inline table,
      and that exact anchor (`tblpX="10645" w:tblpY="4133"`, or your
      template's equivalent) appears on exactly **one** table in the whole
      document.
- [ ] Chart size is exactly `5461000 x 3441700` EMU (5.97in × 3.76in).
- [ ] Chart slice values match the summary table's values exactly, category
      for category, including matching label text.
- [ ] A manual page break follows immediately after the chart paragraph, and
      no additional/earlier page breaks have been inserted around the table
      or chart beyond that one.
