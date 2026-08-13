# Running the VA/CA Automation Project

## Prerequisites

- Python 3.11 or higher
- pip (Python package manager)
- The following files must be present:
  - A raw Nessus export workbook (`.xlsx`) with a sheet named exactly `RAW File`
  - A blank report template (`.xlsx`) with sheets `Introduction`, `VA Report`, and `Summary`

## Step 1: Install the Package

```bash
cd VA_CA_Automation
pip install -e .
```

This installs the `va-ca-automation` package in editable mode, making the CLI command available.

## Step 2: Prepare Your Input Files

1. **Raw Nessus Export**: Must be an `.xlsx` file with a single sheet named `RAW File` containing exactly 17 columns in the expected order:
   - Plugin ID, CVE, CVSS v2.0 Base Score, Risk, Host, Protocol, Port, Name, Synopsis, Description, Solution, See Also, Plugin Output, CVSS v4.0 Base Score, CVSS v3.0 Base Score, VPR Score, EPSS Score

2. **Blank Template**: Must be an `.xlsx` file with sheets `Introduction`, `VA Report`, and `Summary`. The template must be pristine (no leftover data rows or pivot tables). Place it at `templates/va_report_template.xlsx` or specify its path via `--template`.

## Step 3: Run the CLI

### Basic Usage

```bash
va-ca-automation path/to/RAW_file.xlsx \
    --client-name "Trust Investment Advisors Private Limited" \
    --tester "John Doe" \
    --reviewer "Jane Smith"
```

### Full Options

```bash
va-ca-automation path/to/RAW_file.xlsx \
    --template path/to/blank_template.xlsx \
    --output-dir output/ \
    --client-name "Trust Investment Advisors Private Limited" \
    --tester "John Doe" \
    --reviewer "Jane Smith" \
    --report-date 2026-07-09 \
    --version 1.2 \
    --scanner-name "Nessus " \
    --scanner-version "10.11.4" \
    --report-owner "Security Team" \
    --scope "Server" \
    --phase "First" \
    --entity-codes TSS SCPL \
    --host-metadata "192.168.1.1:Authenticated:Server" \
    --host-metadata "192.168.1.2:Authenticated:Firewall" \
    --log-file output/pipeline.log \
    -v
```

### CLI Arguments Reference

| Argument | Required | Default | Description |
|---|---|---|---|
| `raw_file` | Yes | — | Path to the raw Nessus export `.xlsx` |
| `--template` | No | `templates/va_report_template.xlsx` | Path to the blank template |
| `--output-dir` | No | `output/` | Directory for generated reports |
| `--client-name` | Yes | — | Client legal name |
| `--tester` | Yes | — | Security tester name |
| `--reviewer` | Yes | — | Report reviewer name |
| `--report-date` | No | Today | Report date (YYYY-MM-DD) |
| `--version` | No | 1.0 | Report version number |
| `--scanner-name` | No | `Nessus ` | Scanner name |
| `--scanner-version` | No | `10.11.4` | Scanner version |
| `--report-owner` | No | — | Report owner for Introduction sheet |
| `--scope` | No | `Server` | Scope label (Server, Firewall, etc.) |
| `--device-type` | No | — | Device type for all hosts in the summary table (e.g., Server, Firewall) |
| `--phase` | No | `First` | Phase label (First, Retest, etc.) |
| `--entity-codes` | No | — | Entity codes (e.g., TSS SCPL) |
| `--host-metadata` | No | — | Per-host metadata as `IP:ScanType:DeviceType` |
| `--log-file` | No | — | Path for structured log output |
| `-v, --verbose` | No | Off | Enable debug logging |

## Step 4: Check the Output

1. The generated report will be saved to the `--output-dir` directory (default: `output/`)
2. The filename follows the convention: `VA_<Scope>_<Phase>_Audit_Report_<Client>_<Entities>_<Year>_V<Version>.xlsx`
3. If a `--log-file` was specified, check it for pipeline stage counts and dedup details

## Step 5: Run the Tests

```bash
python -m pytest tests/ -v
```

This runs all 50 unit tests covering:
- Ingestion and schema validation
- Risk filtering and normalization
- Two-stage deduplication (exact + version-collapse)
- Column mapping
- Sorting and numbering
- Excel writing and formatting
- Summary/scope table building
- Filename generation
- Pipeline logging

## Architecture Overview

```
src/va_ca_automation/
├── app.py                      # CLI entrypoint with argument parsing
├── ingestion/
│   ├── raw_file_loader.py      # Load RAW File sheet with pandas
│   └── schema_validator.py     # Validate 17-column schema, normalize Risk
├── transform/
│   ├── filters.py              # VA risk filtering (exclude None)
│   ├── dedup.py                # Stage 1 exact dedup + Stage 2 version-collapse
│   ├── column_mapper.py        # Raw-to-template column renaming
│   └── sorter.py               # Host grouping + risk severity ordering
├── excel_writer/
│   ├── template_cloner.py      # Clone pristine template with validation
│   ├── data_writer.py          # Write VA rows with style cloning
│   ├── summary_builder.py      # Scope table + risk aggregation
│   └── chart_builder.py        # Pie chart creation
├── metadata/
│   └── engagement_metadata.py  # Typed dataclass for engagement fields
├── naming/
│   └── filename_builder.py     # Naming convention enforcement
├── logging/
│   └── pipeline_logger.py      # Structured JSONL logging
└── pipelines/
    └── va_pipeline.py          # Full VA pipeline orchestration
```

## Pipeline Flow

```
RAW File (19,958 rows)
    ↓ Load & Validate Schema
    ↓ Normalize whitespace + Risk casing
    ↓ Classify: VA vs CA vs Unknown
    ↓ Filter: exclude Risk=None
    ↓ Stage 1 Dedup: exact match on (Name, Description, Risk, Host)
    ↓ Stage 2 Dedup: version-collapse, keep highest version per group
    ↓ Map columns to template schema
    ↓ Sort: Host blocks → Risk severity (Critical > High > Medium > Low)
    ↓ Assign Sr. no sequentially
    ↓ Clone pristine template
    ↓ Write header metadata (C5:C10)
    ↓ Write data rows (A14:I<last>)
    ↓ Build Summary scope table
    ↓ Build risk aggregation + pie chart
    ↓ Write Introduction fields
    ↓ Save with naming convention
Final Report (.xlsx)
```

## Configuration Files

| File | Purpose |
|---|---|
| `config/column_mappings.yaml` | Raw-to-template column mapping |
| `config/template_ranges.yaml` | Cell range map (static vs dynamic) |
| `config/business_rules.yaml` | Filtering, dedup, sort rules |
| `config/naming_convention.yaml` | Filename pattern tokens |

## Troubleshooting

- **Sheet not found**: Ensure the raw export has a sheet named exactly `RAW File`
- **Schema mismatch**: The raw export must have exactly 17 columns in the expected order
- **Template not found**: Place the blank template at `templates/va_report_template.xlsx` or use `--template`
- **Unknown Risk values**: Check the log file for warnings about unexpected Risk values that were excluded
