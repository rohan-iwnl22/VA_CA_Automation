# Task 02: Extend Word Report Builder to Populate All Sections

## Objective
Extend `build_word_report()` to populate Document Details, Document Change History, Document Distribution, Audit Team, and Testing Details sections in the Word template.

## Files to Modify
- `src/va_ca_automation/word_writer/word_report_builder.py`

## Current State
The Word builder currently only:
1. Replaces `"Client name"` placeholder
2. Populates executive summary tables (Table 11, Table 13)
3. Inserts pie chart on page 14
4. Inserts VA and CA data tables

## Sections to Populate

### 1. Document Details
Search for and populate these fields in the template:
- Document Title (based on First/Final)
- Document ID (`{client_short_name} | {report_number}`)
- Document Version (1.0 for First, report_number for Final)
- Prepared By (from `metadata.security_tester`)
- Reviewed By (from `metadata.reviewed_by`)
- Approved (default: "Default")
- Released Date (from `metadata.released_date`)

### 2. Document Change History
Find the change history table in the template and populate:
- Document Version
- Released Date
- First or Final

### 3. Document Distribution
Find the distribution table and populate:
- Spokesperson name (from `metadata.spokesperson_name`)
- Organisation (from `metadata.client_name`)
- Designation (from `metadata.spokesperson_designation`)
- Email ID (from `metadata.spokesperson_email`)

### 4. Audit Team Details
Find and populate:
- Prepared By: `metadata.security_tester`
- Senior Name: `metadata.senior_name`
- Fixed field (hardcoded in template)

### 5. Testing Details (Conditional)
If First:
- Start Date: `metadata.assessment_start_date`
- Finish Date: `metadata.assessment_finish_date`
- Retesting: "Revalidation not performed"

If Final:
- First Audit Dates: "NA"
- Final Retesting Date: `metadata.final_retesting_start` to `metadata.final_retesting_finish`

### 6. Table of Contents
Note: python-docx cannot auto-update TOC. Add a comment/instruction that user must update TOC manually (Ctrl+A, F9) after printing, or leave a placeholder note.

## Implementation Steps

1. Add new helper functions for each section:
   - `_populate_document_details(doc, metadata)` — Find placeholders and fill document details
   - `_populate_change_history(doc, metadata)` — Find change history table and fill row
   - `_populate_distribution(doc, metadata)` — Find distribution table and fill row
   - `_populate_audit_team(doc, metadata)` — Find and fill audit team section
   - `_populate_testing_details(doc, metadata)` — Conditional date logic

2. Each helper should search for placeholder text (e.g., "Prepared By", "Document ID", etc.) and replace with actual values using `_replace_text_in_paragraph()` or direct cell text assignment.

3. Call all new helpers in `build_word_report()` before step 5 (before inserting VA/CA tables).

4. Update `build_word_report()` signature to accept the additional metadata fields (or pass the full `EngagementMetadata` object which already has them).

## Helper Function Pattern

```python
def _populate_document_details(doc: Document, metadata: EngagementMetadata) -> None:
    """Populate the Document Details section."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "Document Title" in para.text:
                        _replace_text_in_paragraph(para, "Document Title", metadata.document_title)
                    # ... more replacements
```

## Acceptance Criteria
- [ ] Document Details section populated correctly
- [ ] Document Change History table has one row with correct data
- [ ] Document Distribution table populated with spokesperson info
- [ ] Audit Team section shows tester and senior name
- [ ] Testing Details conditional logic works for First and Final
- [ ] All existing functionality (client name, summary tables, pie chart, data tables) still works
- [ ] No regression in existing tests
