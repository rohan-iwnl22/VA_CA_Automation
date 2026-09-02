import pandas as pd
import sys
sys.path.insert(0, '.')

from src.va_ca_automation.ingestion.raw_file_loader import load_raw_file
from src.va_ca_automation.ingestion.schema_validator import (
    classify_rows, normalize_whitespace_columns, validate_and_normalize_risk,
)
from src.va_ca_automation.transform.dedup import (
    stage1_exact_dedup, stage1b_name_host_dedup, stage2_version_collapse,
    _extract_identifier, _make_base_title,
)
from src.va_ca_automation.transform.filters import filter_va_candidates
from src.va_ca_automation.logging.pipeline_logger import PipelineLogger

raw = load_raw_file('servers.xlsx')
raw = normalize_whitespace_columns(raw, ["Risk", "Host", "Name"])
raw = validate_and_normalize_risk(raw, PipelineLogger())
va_rows, _, _ = classify_rows(raw)
va_filtered = filter_va_candidates(va_rows)
s1 = stage1_exact_dedup(va_filtered)
s1b = stage1b_name_host_dedup(s1)

# For host 192.168.33.49, show what version-collapse does
host = '192.168.33.49'
host_data = s1b[s1b['Host'] == host].copy()

# Show OpenSSL vulns
openssl = host_data[host_data['Name'].str.contains('OpenSSL 1.0.2', na=False)]
print(f"=== OpenSSL 1.0.2 vulns for {host} ({len(openssl)} rows) ===")
for _, row in openssl.iterrows():
    name = row['Name']
    ident = _extract_identifier(name)
    base = _make_base_title(name)
    print(f"  Name: {name}")
    print(f"    identifier: {ident}")
    print(f"    base_title: {base}")
    print()

# Show Oracle Java vulns
java = host_data[host_data['Name'].str.contains('Oracle Java', na=False)]
print(f"=== Oracle Java vulns for {host} ({len(java)} rows) ===")
for _, row in java.iterrows():
    name = row['Name']
    ident = _extract_identifier(name)
    base = _make_base_title(name)
    print(f"  Name: {name}")
    print(f"    identifier: {ident}")
    print(f"    base_title: {base}")
    print()

# Show what version-collapse produces
plogger = PipelineLogger()
s2 = stage2_version_collapse(s1b, plogger)
host_result = s2[s2['Host'] == host]
openssl_result = host_result[host_result['Name'].str.contains('OpenSSL 1.0.2', na=False)]
java_result = host_result[host_result['Name'].str.contains('Oracle Java', na=False)]
print(f"=== After version collapse for {host} ===")
print(f"OpenSSL 1.0.2 kept: {len(openssl_result)} rows")
for _, row in openssl_result.iterrows():
    print(f"  {row['Name']}")
print(f"Oracle Java kept: {len(java_result)} rows")
for _, row in java_result.iterrows():
    print(f"  {row['Name']}")
