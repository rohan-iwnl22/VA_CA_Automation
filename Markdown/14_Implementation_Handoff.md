# 14 - Implementation Handoff

This document tracks what has been completed so far and gives a clean starting point for the next coding agent.

## Completed So Far

### Repo structure
- Created a `src/`-based Python package at `src/va_ca_automation/`.
- Added module folders for:
  - `ingestion`
  - `pipelines`
  - `transform`
  - `excel_writer`
  - `metadata`
  - `naming`
  - `logging`

### Project manifest
- Added `pyproject.toml` with:
  - `setuptools` build backend
  - package discovery from `src/`
  - CLI entry point `va-ca-automation`
  - initial dependencies:
    - `openpyxl`
    - `pandas`
    - `pyyaml`

### Runtime entrypoint
- Added a minimal application entrypoint in `src/va_ca_automation/app.py`.
- Added `src/va_ca_automation/__main__.py` so the package can run as a module.
- Confirmed the package imports correctly with version `0.1.0`.

### Test scaffolding
- Added `tests/test_smoke.py` as a basic import smoke test.
- Added `tests/conftest.py` so the `src/` layout resolves cleanly in tests.

### Repo hygiene
- Added `.gitignore`.
- Added placeholder `README.md`.
- Added placeholder folders:
  - `config/`
  - `templates/`
  - `output/`

## Current State

- The repository is now scaffolded but not yet functionally implemented.
- No VA processing logic exists yet.
- No Excel read/write logic exists yet.
- No template data, mapping data, or business-rule code has been implemented yet.

## Canonical Spec References

Use these docs as the source of truth for implementation:
- `01_Project_Overview.md`
- `03_Input_Files.md`
- `05_Business_Rules.md`
- `06_Data_Transformation.md`
- `07_VA_Workflow.md`
- `09_Excel_Template_Mapping.md`
- `11_Project_Architecture.md`
- `12_Testing_Checklist.md`

## Next Coding Instructions For The Other LLM Agent

### Priority 1
1. Implement raw workbook ingestion for the `RAW File` sheet.
2. Validate the 17-column schema exactly.
3. Normalize `Risk`, `Host`, and `Name` by trimming whitespace.
4. Classify rows into VA and CA routing buckets.
5. For the current build phase, only continue with the VA pipeline.

### Priority 2
1. Implement VA filtering:
   - keep `Critical`, `High`, `Medium`, `Low`
   - exclude `None`
2. Implement two-stage deduplication:
   - Stage 1 exact key: `Name`, `Description`, `Risk`, `Host`
   - Stage 2 version-collapse for embedded numeric versions in `Name`
3. Implement host grouping and severity ordering:
   - `Critical > High > Medium > Low`
   - stable tiebreak within the same host and severity
4. Assign `Sr. no` after sorting.

### Priority 3
1. Write processed VA rows into the template using values only.
2. Preserve all static template formatting.
3. Write only to the dynamic ranges defined in `09_Excel_Template_Mapping.md`.
4. Populate the `VA Report` header block and `Introduction` metadata cells.

### Priority 4
1. Build the `Summary` sheet scope table from distinct hosts.
2. Use engagement metadata for `Scan Type` and `Device Type`.
3. Recreate the risk summary as a fresh aggregation every run.
4. Recreate the pie chart from the same summary data.

### Priority 5
1. Add filename generation based on the naming convention.
2. Add structured logging for row counts and dedup decisions.
3. Add post-write validation checks.
4. Expand tests from smoke tests to unit and integration coverage.

## Implementation Rules

- Start from a pristine template copy every run.
- Do not mutate the template source file directly.
- Do not write formulas unless explicitly required by the spec.
- Do not touch static ranges or header text.
- Keep the implementation deterministic and idempotent.
- Treat the stale pivot/chart in the sample workbook as a defect, not a feature.
- Keep CA pipeline work deferred unless the business provides a confirmed CA sample and filtering decision.

## Open Technical Risks

- Native Excel pivot/chart creation may require a different approach than pure `openpyxl`.
- The safest first implementation path may be static aggregation plus a matching pie chart if native pivot creation is not practical.
- Host block ordering is still a lower-priority nuance if the business wants it finalized before full build.

## Suggested Immediate Next Step

Build `src/va_ca_automation/ingestion/raw_file_loader.py` and `src/va_ca_automation/ingestion/schema_validator.py` first, then wire them into a small `run_va_pipeline()` function stub.

