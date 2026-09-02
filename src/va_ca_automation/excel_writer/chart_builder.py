"""Build the summary chart."""

from __future__ import annotations

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

RISK_COLORS = {
    "Critical": "C00000",
    "High": "FF0000",
    "Medium": "FFC000",
    "Low": "FFFF00",
}

CA_RISK_COLORS = {
    "FAILED": "4472C4",
    "WARNING": "ED7D31",
}


def build_pie_chart(
    ws,
    risk_summary: dict[str, int],
    data_anchor: str = "E18",
    chart_anchor: str = "I8",
    data_start_row: int = 18,
) -> None:
    """Create a pie chart from the risk summary data.

    The chart is anchored at I8 (below the orange merged cell I7:N7).
    Data source is the risk summary table written at columns D-E.

    Parameters
    ----------
    ws : worksheet
        The Summary worksheet.
    risk_summary : dict
        Risk level counts including "Grand Total".
    data_anchor : str
        Legacy parameter (unused, kept for backward compatibility).
    chart_anchor : str
        Cell where the pie chart is placed.
    data_start_row : int
        The row where the risk summary header starts (default 18).
    """
    chart = PieChart()
    chart.title = "Vulnerability Risk Distribution"
    chart.style = 10
    # Height: 9.87 cm, Width: 9.4 cm (openpyxl uses cm)
    chart.width = 9.4
    chart.height = 9.87

    # Exclude Grand Total from chart data - only show Critical, High, Medium, Low
    risk_labels = [label for label in risk_summary.keys() if label != "Grand Total"]
    num_risk_levels = len(risk_labels)

    # Data source: column F (Count of Host)
    # Header is at data_start_row, data starts at data_start_row + 1
    data_ref = Reference(ws, min_col=6, min_row=data_start_row + 1, max_row=data_start_row + num_risk_levels)

    # Categories: column E (Row Labels)
    cats_ref = Reference(ws, min_col=5, min_row=data_start_row + 1, max_row=data_start_row + num_risk_levels)

    chart.add_data(data_ref, titles_from_data=False)
    chart.set_categories(cats_ref)

    # Apply custom colors to each slice based on risk level
    series = chart.series[0]
    for idx, label in enumerate(risk_labels):
        color = RISK_COLORS.get(label)
        if color:
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = color
            # Add border (outline) to pie slice
            if pt.graphicalProperties.line is None:
                pt.graphicalProperties.line = LineProperties()
            pt.graphicalProperties.line.solidFill = "000000"
            pt.graphicalProperties.line.width = 10000  # 1pt border
            series.data_points.append(pt)

    # Data labels: only value, bold font
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = False
    chart.dataLabels.showVal = True
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


def build_ca_pie_chart(
    ws,
    risk_summary: dict[str, int],
    chart_anchor: str = "I8",
    data_start_row: int = 15,
) -> None:
    """Create a pie chart from the CA risk summary data.

    The chart is anchored at I8 (below the orange merged cell I7:N7).
    Data source is the risk summary table written at columns E-F.

    Parameters
    ----------
    ws : worksheet
        The Summary worksheet.
    risk_summary : dict
        Risk level counts including "Grand Total".
    chart_anchor : str
        Cell where the pie chart is placed.
    data_start_row : int
        The row where the risk summary header starts (default 15).
    """
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties

    chart = PieChart()
    chart.title = "Compliance Risk Distribution"
    chart.style = 10
    # Height: 9.87 cm, Width: 9.4 cm (openpyxl uses cm)
    chart.width = 9.4
    chart.height = 9.87

    # Add border around the chart frame
    chart.graphical_properties = GraphicalProperties()
    chart.graphical_properties.line = LineProperties()
    chart.graphical_properties.line.solidFill = "000000"
    chart.graphical_properties.line.width = 12700  # 1pt border

    # Exclude Grand Total from chart data - only show FAILED, WARNING
    risk_labels = [label for label in risk_summary.keys() if label != "Grand Total"]
    num_risk_levels = len(risk_labels)

    # Data source: column F (Count of Host)
    # Header is at data_start_row, data starts at data_start_row + 1
    data_ref = Reference(ws, min_col=6, min_row=data_start_row + 1, max_row=data_start_row + num_risk_levels)

    # Categories: column E (Row Labels)
    cats_ref = Reference(ws, min_col=5, min_row=data_start_row + 1, max_row=data_start_row + num_risk_levels)

    chart.add_data(data_ref, titles_from_data=False)
    chart.set_categories(cats_ref)

    # Apply custom colors to each slice based on risk level
    series = chart.series[0]
    for idx, label in enumerate(risk_labels):
        color = CA_RISK_COLORS.get(label)
        if color:
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = color
            if pt.graphicalProperties.line is None:
                pt.graphicalProperties.line = LineProperties()
            pt.graphicalProperties.line.solidFill = "000000"
            pt.graphicalProperties.line.width = 10000
            series.data_points.append(pt)

    # Data labels: only value, bold font
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = False
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False

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

    ws.add_chart(chart, chart_anchor)
