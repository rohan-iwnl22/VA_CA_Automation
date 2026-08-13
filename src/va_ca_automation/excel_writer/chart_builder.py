"""Build the summary chart."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint

RISK_COLORS = {
    "Critical": "C00000",
    "High": "FF0000",
    "Medium": "FFC000",
    "Low": "FFFF00",
}


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

    # Exclude Grand Total from chart data - only show Critical, High, Medium, Low
    risk_labels = [label for label in risk_summary.keys() if label != "Grand Total"]
    num_risk_levels = len(risk_labels)

    # Data source: column F (Count of Host), rows 19 to 22 (4 risk levels only)
    data_ref = Reference(ws, min_col=6, min_row=18, max_row=18 + num_risk_levels)

    # Categories: column E (Row Labels), rows 19 to 22
    cats_ref = Reference(ws, min_col=5, min_row=19, max_row=18 + num_risk_levels)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    # Apply custom colors to each slice based on risk level
    series = chart.series[0]
    for idx, label in enumerate(risk_labels):
        color = RISK_COLORS.get(label)
        if color:
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = color
            series.data_points.append(pt)

    # Data labels: only percentage, bold font
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showVal = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False

    # Bold font for data labels
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties

    chart.dataLabels.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=CharacterProperties(
                    b=True,
                    sz=1200,
                    solidFill="000000",
                )
            ),
            endParaRPr=CharacterProperties(
                b=True,
                sz=1200,
                solidFill="000000",
            )
        )]
    )

    # Place chart at anchor position
    ws.add_chart(chart, chart_anchor)
