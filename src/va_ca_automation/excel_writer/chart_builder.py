"""Build the summary chart."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList


def build_pie_chart(
    ws,
    risk_summary: dict[str, int],
    data_anchor: str = "E18",
    chart_anchor: str = "H8",
) -> None:
    """Create a pie chart from the risk summary data.

    The chart is anchored at approximately H8 to match the completed sample layout.
    Data source is the risk summary table written at E18:F23.
    """
    chart = PieChart()
    chart.title = "Vulnerability Risk Distribution"
    chart.style = 10
    chart.width = 18
    chart.height = 14

    # Data source: column F (Count of Host), rows 19 to 23
    data_ref = Reference(ws, min_col=6, min_row=18, max_row=18 + len(risk_summary))

    # Categories: column E (Row Labels), rows 19 to 23
    cats_ref = Reference(ws, min_col=5, min_row=19, max_row=18 + len(risk_summary))

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    # Style the pie slices
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = False

    # Place chart at anchor position
    ws.add_chart(chart, chart_anchor)
