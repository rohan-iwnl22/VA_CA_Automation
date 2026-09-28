"""Quick smoke test for the /api/final-report endpoint."""
import io
import openpyxl
from fastapi.testclient import TestClient
from jose import jwt
from va_ca_automation.api.main import create_app
from va_ca_automation.api.deps import JWT_SECRET, ALGORITHM

app = create_app()
client = TestClient(app)

token = jwt.encode({"sub": "admin", "role": "admin"}, JWT_SECRET, algorithm=ALGORITHM)
headers = {"Authorization": f"Bearer {token}"}

header_names = [
    "Sr. no", "Vulnerbility Title", "Description", "Risk",
    "Host", "Port", "Recommendation ", "Reference", "CVE", "Retest Status",
]

# First Audit
wb1 = openpyxl.Workbook()
ws1 = wb1.active
ws1.title = "VA Report"
for i, h in enumerate(header_names, 1):
    ws1.cell(row=13, column=i, value=h)
for r, row in enumerate([
    [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "r1", "CVE-1", ""],
    [2, "Outdated Apache", "Desc", "Critical", "10.0.0.2", "80", "Fix", "r2", "CVE-2", ""],
    [3, "SMB Vulnerability", "Desc", "High", "10.0.0.3", "445", "Fix", "r3", "CVE-3", ""],
], 14):
    for c, val in enumerate(row, 1):
        ws1.cell(row=r, column=c, value=val)
buf1 = io.BytesIO()
wb1.save(buf1)
buf1.seek(0)
wb1.close()

# Retest
wb2 = openpyxl.Workbook()
ws2 = wb2.active
ws2.title = "VA Report"
for i, h in enumerate(header_names, 1):
    ws2.cell(row=13, column=i, value=h)
for r, row in enumerate([
    [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "r1", "CVE-1", ""],
    [3, "SMB Vulnerability", "Desc", "High", "10.0.0.3", "445", "Fix", "r3", "CVE-3", ""],
], 14):
    for c, val in enumerate(row, 1):
        ws2.cell(row=r, column=c, value=val)
buf2 = io.BytesIO()
wb2.save(buf2)
buf2.seek(0)
wb2.close()

response = client.post(
    "/api/final-report",
    headers=headers,
    files={
        "first_audit": ("first.xlsx", buf1, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "retest_file": ("retest.xlsx", buf2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    },
)

print(f"Status: {response.status_code}")
print(f"Content-Type: {response.headers.get('content-type', 'N/A')}")
print(f"Content-Length: {len(response.content)} bytes")

if response.status_code == 200:
    wb_out = openpyxl.load_workbook(io.BytesIO(response.content))
    ws_out = wb_out.active
    print("\nOutput verification:")
    for row in range(14, ws_out.max_row + 1):
        title = ws_out.cell(row=row, column=2).value
        host = ws_out.cell(row=row, column=5).value
        status = ws_out.cell(row=row, column=10).value
        print(f"  {host:12s} | {title:20s} | Retest Status: {status}")
    wb_out.close()
else:
    print(f"Error: {response.text}")
