# How to Run This Project (Simple Guide)

## What Is This?

This is a Python tool that takes a raw Nessus vulnerability scan export (an Excel file) and turns it into a polished, branded Excel report. It reads your scan data, cleans it up, removes duplicates, and writes it into a nice template.

---

## What Do You Need?

1. **Python 3.11 or newer** installed on your computer
   - Check by opening a terminal and typing: `python --version`
   - If you don't have it, download from https://www.python.org/downloads/

2. **A raw Nessus Excel file** (`.xlsx`) — this is what you exported from Nessus
   - It MUST have a sheet called exactly `RAW File`
   - That sheet MUST have exactly 17 columns in this order:
     `Plugin ID, CVE, CVSS v2.0 Base Score, Risk, Host, Protocol, Port, Name, Synopsis, Description, Solution, See Also, Plugin Output, CVSS v4.0 Base Score, CVSS v3.0 Base Score, VPR Score, EPSS Score`

3. **A blank report template** (`.xlsx`) — this is your branded template
   - It MUST have 3 sheets: `Introduction`, `VA Report`, `Summary`
   - Place it in the `templates` folder and name it `va_report_template.xlsx`

---

## Step 1: Install the Tool

Open a terminal (PowerShell, Command Prompt, or VS Code terminal) and run these commands one by one:

```bash
cd D:\DeployableProj\VA_CA_Automation
pip install -e .
```

- The `cd` command moves into the project folder
- The `pip install -e .` command installs the tool on your computer

If you see "Successfully installed", you're good to go.

---

## Step 2: Run the Tool

### The Simplest Way (just 3 required things)

```bash
va-ca-automation path\to\your\scan_file.xlsx --client-name "Client Name" --tester "Your Name" --reviewer "Reviewer Name"

va-ca-automation D:\DeployableProj\VA_CA_Automation\HOW_TO_RUN.md --client-name "D" --tester "P" --reviewer "R"

```

Replace:
- `path\to\your\scan_file.xlsx` with the actual path to your Nessus Excel export
- `"Client Name"` with the client's company name
- `"Your Name"` with who did the testing
- `"Reviewer Name"` with who reviewed the report

**Example:**
```bash
va-ca-automation "C:\Scans\acme_scan.xlsx" --client-name "Acme Corporation" --tester "John Smith" --reviewer "Jane Doe"
```

The report will be saved in the `output` folder.

### With More Options

```bash
va-ca-automation "C:\Scans\acme_scan.xlsx" ^
    --template "C:\Templates\my_template.xlsx" ^
    --output-dir "C:\Reports" ^
    --client-name "Acme Corporation" ^
    --tester "John Smith" ^
    --reviewer "Jane Doe" ^
    --report-date 2026-08-11 ^
    --version 1.0 ^
    --scope "Server" ^
    --phase "First" ^
    --entity-codes TSS SCPL ^
    -v
```

Note: On PowerShell, use backticks `` ` `` instead of `^` for line continuation. Or just put everything on one line.

---

## Step 3: Check Your Report

1. Open the `output` folder (or wherever you set `--output-dir`)
2. You'll find an Excel file with a name like:
   `VA_Server_First_Audit_Report_Acme_Corporation_TSS_SCPL_2026_V1.xlsx`
3. Open it and check that it looks correct

---

## Running the Tests

To make sure everything is working, run the tests:

```bash
python -m pytest tests/ -v
```

If all 50 tests pass, you're all set.

---

## All the Options You Can Use

| What it does | Flag | Required? | Default |
|---|---|---|---|
| Your scan file | (just put the path) | YES | — |
| Client name | `--client-name` | YES | — |
| Tester name | `--tester` | YES | — |
| Reviewer name | `--reviewer` | YES | — |
| Template file | `--template` | No | `templates/va_report_template.xlsx` |
| Output folder | `--output-dir` | No | `output/` |
| Report date | `--report-date` | No | Today's date |
| Version number | `--version` | No | 1.0 |
| Scanner name | `--scanner-name` | No | `Nessus ` |
| Scanner version | `--scanner-version` | No | `10.11.4` |
| Report owner | `--report-owner` | No | — |
| Scope (Server/Firewall/etc) | `--scope` | No | `Server` |
| Phase (First/Retest/etc) | `--phase` | No | `First` |
| Entity codes | `--entity-codes` | No | — |
| Host info | `--host-metadata` | No | — |
| Log file | `--log-file` | No | — |
| Verbose output | `-v` | No | Off |

---

## Common Errors and Fixes

| Error | What It Means | How to Fix |
|---|---|---|
| `'va-ca-automation' is not recognized` | The tool isn't installed | Run `pip install -e .` again |
| `Sheet not found: RAW File` | Your Excel file doesn't have the right sheet name | Make sure the sheet is named exactly `RAW File` (capital R, capital F) |
| `Expected 17 columns, found X` | Your Excel file has wrong columns | Make sure you exported from Nessus correctly |
| `Template not found` | Can't find the blank template | Put your template at `templates/va_report_template.xlsx` or use `--template` |
| `ModuleNotFoundError` | Python can't find the code | Make sure you ran `pip install -e .` from the project folder |

---

## Quick Reference

```bash
# 1. Go to project folder
cd D:\DeployableProj\VA_CA_Automation

# 2. Install (only need to do this once)
pip install -e .

# 3. Run the tool (minimum required arguments)
va-ca-automation "C:\path\to\scan.xlsx" --client-name "Client" --tester "Tester" --reviewer "Reviewer"

# 4. Check the output folder for your report
```
