"""Build scope tables and summary aggregations."""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from ..metadata.engagement_metadata import EngagementMetadata

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

DATA_FONT = Font(name="Cambria", size=11)

RISK_ORDER = ["Critical", "High", "Medium", "Low"]

GRAND_TOTAL_FILL = PatternFill(start_color="AEAAAA", end_color="AEAAAA", fill_type="solid")


def build_scope_table(va_sorted_df: pd.DataFrame, metadata: EngagementMetadata) -> pd.DataFrame:
    """Build the Summary scope table from distinct hosts in the processed data."""
    distinct_hosts = va_sorted_df["Host"].unique()

    rows = []
    for ip in distinct_hosts:
        rows.append(
            {
                "IP Address": ip,
                "Scan Type": metadata.get_host_scan_type(ip),
                "Device Type": metadata.get_host_device_type(ip),
            }
        )

    return pd.DataFrame(rows)


def write_scope_table(ws, scope_df: pd.DataFrame) -> None:
    """Write the scope table to the Summary sheet starting at row 8."""
    for i, row in scope_df.iterrows():
        excel_row = 8 + i
        ws.cell(row=excel_row, column=1, value=row["IP Address"])
        ws.cell(row=excel_row, column=2, value=row["Scan Type"])
        ws.cell(row=excel_row, column=3, value=row["Device Type"])

        for col in range(1, 4):
            cell = ws.cell(row=excel_row, column=col)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER


def build_risk_summary(va_sorted_df: pd.DataFrame) -> dict[str, int]:
    """Build the risk count aggregation from processed VA data."""
    risk_counts = va_sorted_df["Risk"].value_counts()
    result = {}
    for risk in RISK_ORDER:
        result[risk] = int(risk_counts.get(risk, 0))
    result["Grand Total"] = sum(result.values())
    return result


def write_risk_summary_table(ws, risk_summary: dict[str, int]) -> None:
    """Write the risk summary table to the Summary sheet.

    Positions the table to match the completed-sample layout.
    """
    # Write header row
    ws.cell(row=18, column=5, value="Row Labels")
    ws.cell(row=18, column=6, value="Count of Host")

    # Apply border to header row
    for col in [5, 6]:
        cell = ws.cell(row=18, column=col)
        cell.font = DATA_FONT
        cell.border = THIN_BORDER

    # Write data rows
    row_offset = 19
    for i, (label, count) in enumerate(risk_summary.items()):
        ws.cell(row=row_offset + i, column=5, value=label)
        ws.cell(row=row_offset + i, column=6, value=count)

        for col in [5, 6]:
            cell = ws.cell(row=row_offset + i, column=col)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER

    # Bold the Grand Total row and apply background color
    grand_total_row = row_offset + len(risk_summary) - 1
    for col in [5, 6]:
        cell = ws.cell(row=grand_total_row, column=col)
        cell.font = Font(name="Cambria", size=11, bold=True)
        cell.fill = GRAND_TOTAL_FILL
