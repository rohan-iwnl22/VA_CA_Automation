"""Build the Word audit report from template and processed data."""

from __future__ import annotations

import io
import logging
import os
import shutil
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from ..metadata.engagement_metadata import EngagementMetadata

logger = logging.getLogger("va_ca_automation")

VA_COLUMNS = ["Sr. no", "Vulnerbility Title", "Description", "Risk", "Host", "Port", "Recommendation ", "Reference", "CVE"]
CA_COLUMNS = ["Sr.No.", "Title", "Host", "Description", "Solution", "Risk"]

RISK_COLORS = {
    "Critical": "#C00000",
    "High": "#FF0000",
    "Medium": "#FFC000",
    "Low": "#FFFF00",
    "Info": "#00B0F0",
    "Failed": "#4472C4",
    "Warning": "#ED7D31",
}

CA_RISK_COLORS = {
    "FAILED": "#4472C4",
    "WARNING": "#ED7D31",
}


def _set_cell_shading(cell, color_hex: str) -> None:
    """Set background shading on a table cell."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def _set_row_height(row, height_pt: float) -> None:
    """Set row height in points."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(height_pt * 20)))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)


def _repeat_table_header(row) -> None:
    """Mark a table row as a repeating header row for page breaks."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def _set_cell_borders(cell, color: str = "000000", size: str = "4") -> None:
    """Set thin borders on a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        tcBorders.append(element)
    tcPr.append(tcBorders)


def _insert_table_after_paragraph(doc: Document, para_index: int, rows: int, cols: int):
    """Insert a new table after the specified paragraph index.

    Returns the new table object.
    """
    para = doc.paragraphs[para_index]
    tbl = doc.add_table(rows=rows, cols=cols)
    # Move the table element to right after the paragraph
    para._p.addnext(tbl._tbl)
    return tbl


def _generate_pie_chart_image(
    risk_summary: dict[str, int],
    title: str,
    colors: dict[str, str],
    ca_risk_summary: dict[str, int] | None = None,
) -> io.BytesIO:
    """Generate a pie chart image and return as BytesIO.

    If ca_risk_summary is provided, FAILED and WARNING counts are included
    in the pie chart alongside the VA risk levels.
    """
    labels = []
    sizes = []
    chart_colors = []

    # Add VA risk levels (Critical, High, Medium, Low, Info)
    for label, count in risk_summary.items():
        if label == "Grand Total" or count == 0:
            continue
        labels.append(label)
        sizes.append(count)
        chart_colors.append(colors.get(label, "#999999"))

    # Add CA risk levels (FAILED, WARNING) if provided
    if ca_risk_summary:
        for label, count in ca_risk_summary.items():
            if label == "Grand Total" or count == 0:
                continue
            # Map FAILED -> Failed, WARNING -> Warning for display
            display_label = label.title() if label.isupper() else label
            labels.append(display_label)
            sizes.append(count)
            chart_colors.append(colors.get(display_label, colors.get(label, "#999999")))

    if not sizes:
        return None

    fig, ax = plt.subplots(figsize=(5, 4))
    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=labels,
        colors=chart_colors,
        autopct=lambda pct: f"{int(round(pct * sum(sizes) / 100))}",
        startangle=90,
        textprops={"fontsize": 10},
    )
    for t in autotexts:
        t.set_fontweight("bold")
    ax.set_title(title, fontsize=12, fontweight="bold")
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _replace_text_in_paragraph(para, old_text: str, new_text: str) -> None:
    """Replace text in a paragraph, handling text split across multiple runs."""
    full_text = para.text
    if old_text not in full_text:
        return

    # Strategy: clear all runs, set the replaced text on the first run
    new_full_text = full_text.replace(old_text, new_text)
    if para.runs:
        para.runs[0].text = new_full_text
        for run in para.runs[1:]:
            run.text = ""


def _replace_client_name(doc: Document, client_name: str) -> None:
    """Replace 'Client name' with the actual client name throughout the document."""
    # Replace in all paragraphs (handles text split across runs)
    for para in doc.paragraphs:
        if "Client name" in para.text or "client name" in para.text:
            _replace_text_in_paragraph(para, "Client name", client_name)
            _replace_text_in_paragraph(para, "client name", client_name)

    # Replace in all table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "Client name" in para.text or "client name" in para.text:
                        _replace_text_in_paragraph(para, "Client name", client_name)
                        _replace_text_in_paragraph(para, "client name", client_name)

    # Replace client name in footers (including tables in footers)
    for section in doc.sections:
        footer = section.footer
        # Access footer XML directly to find all text elements
        from lxml import etree
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for t_elem in footer._element.findall(".//w:t", nsmap):
            if t_elem.text and ("Client name" in t_elem.text or "client name" in t_elem.text):
                t_elem.text = t_elem.text.replace("Client name", client_name)
                t_elem.text = t_elem.text.replace("client name", client_name)


def _populate_executive_summary_table(doc: Document, va_risk: dict[str, int], ca_risk: dict[str, int]) -> None:
    """Populate Table 11 (Security Assessment summary) and Table 13 (Risk Classification)."""
    # Table 11: Security Assessment summary (5 rows x 9 cols)
    # Row 0: merged header "Security Assessment" / "Initial VAPT Scan"
    # Row 1: headers Sr. No., VAPT, Critical, High, Medium, Low, Info, Failed, Warning
    # Row 2: VA data
    # Row 3: CA data
    # Row 4: Total
    if len(doc.tables) > 11:
        tbl = doc.tables[11]
        # Row 2 (VA): fill Critical, High, Medium, Low
        va_row = tbl.rows[2]
        va_row.cells[2].text = str(va_risk.get("Critical", 0))
        va_row.cells[3].text = str(va_risk.get("High", 0))
        va_row.cells[4].text = str(va_risk.get("Medium", 0))
        va_row.cells[5].text = str(va_risk.get("Low", 0))
        va_row.cells[6].text = "0"
        va_row.cells[7].text = "-"
        va_row.cells[8].text = "-"

        # Row 3 (CA): fill Failed, Warning
        ca_row = tbl.rows[3]
        ca_row.cells[2].text = "-"
        ca_row.cells[3].text = "-"
        ca_row.cells[4].text = "-"
        ca_row.cells[5].text = "-"
        ca_row.cells[6].text = "0"
        ca_row.cells[7].text = str(ca_risk.get("FAILED", 0))
        ca_row.cells[8].text = str(ca_risk.get("WARNING", 0))

        # Row 4 (Total)
        total_row = tbl.rows[4]
        total_row.cells[2].text = str(va_risk.get("Critical", 0))
        total_row.cells[3].text = str(va_risk.get("High", 0))
        total_row.cells[4].text = str(va_risk.get("Medium", 0))
        total_row.cells[5].text = str(va_risk.get("Low", 0))
        total_row.cells[6].text = "0"
        total_row.cells[7].text = str(ca_risk.get("FAILED", 0))
        total_row.cells[8].text = str(ca_risk.get("WARNING", 0))

    # Table 13: Risk Classification (8 rows x 2 cols)
    if len(doc.tables) > 13:
        tbl13 = doc.tables[13]
        risk_map = {
            0: "Critical",
            1: "High",
            2: "Medium",
            3: "Low",
            4: "Info",
            5: "Failed",
            6: "Compliance",
            7: "Total",
        }
        for idx, label in risk_map.items():
            if label in ("Critical", "High", "Medium", "Low", "Info"):
                val = va_risk.get(label, 0)
            elif label == "Failed":
                val = ca_risk.get("FAILED", 0)
            elif label == "Compliance":
                val = ca_risk.get("WARNING", 0)
            elif label == "Total":
                val = sum(v for k, v in va_risk.items() if k != "Grand Total") + \
                      sum(v for k, v in ca_risk.items() if k != "Grand Total")
            else:
                val = 0
            tbl13.rows[idx].cells[1].text = str(val)


def _populate_executive_summary_text(doc: Document, va_risk: dict[str, int], ca_risk: dict[str, int]) -> None:
    """Update the executive summary paragraph with actual counts."""
    total_ca_failed = ca_risk.get("FAILED", 0)
    total_ca_warning = ca_risk.get("WARNING", 0)

    for para in doc.paragraphs:
        if "The audit reported" in para.text and "Critical" in para.text:
            new_text = (
                f"The audit reported {va_risk.get('Critical', 0)} Critical, "
                f"{va_risk.get('High', 0)} High, {va_risk.get('Medium', 0)} Medium, "
                f"{va_risk.get('Low', 0)} Low risk, 0 informational vulnerabilities and "
                f"{total_ca_failed} Failed, {total_ca_warning} Warning compliance issue "
                f"on the target asset. Following table and graph summarizes overall risk "
                f"of target infrastructure."
            )
            # Clear all runs and set text on first run
            if para.runs:
                para.runs[0].text = new_text
                for run in para.runs[1:]:
                    run.text = ""
            break


def _create_va_table(doc: Document, va_df: pd.DataFrame, para_index: int) -> None:
    """Insert the VA detailed report table after the specified paragraph."""
    num_rows = len(va_df) + 1  # +1 for header
    num_cols = len(VA_COLUMNS)

    tbl = _insert_table_after_paragraph(doc, para_index, num_rows, num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Style header row
    header_row = tbl.rows[0]
    _repeat_table_header(header_row)
    _set_row_height(header_row, 30)
    for j, col_name in enumerate(VA_COLUMNS):
        cell = header_row.cells[j]
        cell.text = col_name
        _set_cell_shading(cell, "FFC000")
        _set_cell_borders(cell)
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.name = "Cambria"

    # Write data rows
    for i, (_, row) in enumerate(va_df.iterrows()):
        data_row = tbl.rows[i + 1]
        _set_row_height(data_row, 60)
        for j, col_name in enumerate(VA_COLUMNS):
            cell = data_row.cells[j]
            value = row.get(col_name, "")
            if pd.isna(value) or value == "":
                cell.text = "N/A"
            else:
                cell.text = str(value)
            _set_cell_borders(cell)
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.name = "Cambria"
                if col_name in ("Sr. no", "Risk", "Host", "Port", "CVE"):
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_name in ("Description", "Recommendation ", "Reference"):
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _create_ca_table(doc: Document, ca_df: pd.DataFrame, para_index: int) -> None:
    """Insert the CA detailed report table after the specified paragraph."""
    num_rows = len(ca_df) + 1  # +1 for header
    num_cols = len(CA_COLUMNS)

    tbl = _insert_table_after_paragraph(doc, para_index, num_rows, num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Style header row
    header_row = tbl.rows[0]
    _repeat_table_header(header_row)
    _set_row_height(header_row, 30)
    for j, col_name in enumerate(CA_COLUMNS):
        cell = header_row.cells[j]
        cell.text = col_name
        _set_cell_shading(cell, "FFC000")
        _set_cell_borders(cell)
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.name = "Cambria"

    # Write data rows
    for i, (_, row) in enumerate(ca_df.iterrows()):
        data_row = tbl.rows[i + 1]
        _set_row_height(data_row, 60)
        for j, col_name in enumerate(CA_COLUMNS):
            cell = data_row.cells[j]
            value = row.get(col_name, "")
            if pd.isna(value) or value == "":
                cell.text = "N/A"
            else:
                cell.text = str(value)
            _set_cell_borders(cell)

            # Color Risk cells
            if col_name == "Risk":
                risk_val = str(value).strip().upper() if not pd.isna(value) else ""
                if risk_val == "FAILED":
                    _set_cell_shading(cell, "4472C4")
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                            run.font.bold = True
                elif risk_val == "WARNING":
                    _set_cell_shading(cell, "ED7D31")
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                            run.font.bold = True

            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.name = "Cambria"
                if col_name in ("Sr.No.", "Risk"):
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_name == "Host":
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_name in ("Description", "Solution"):
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _create_va_table_with_chart(doc: Document, va_df: pd.DataFrame, para_index: int, va_risk_summary: dict[str, int], ca_risk_summary: dict[str, int] | None = None) -> None:
    """Insert VA data table after the 'below table shows' text.

    Layout matches reference:
    - "The below table shows the detailed report of the VA scan done on assets."
    - VA data table (pie chart is on page 14 in Executive Summary section)
    """
    para = doc.paragraphs[para_index]

    # Insert the VA data table after the "The below table shows..." paragraph
    num_rows = len(va_df) + 1  # +1 for header
    num_cols = len(VA_COLUMNS)

    tbl = _insert_table_after_paragraph(doc, para_index, num_rows, num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Style header row
    header_row = tbl.rows[0]
    _repeat_table_header(header_row)
    _set_row_height(header_row, 30)
    for j, col_name in enumerate(VA_COLUMNS):
        cell = header_row.cells[j]
        cell.text = col_name
        _set_cell_shading(cell, "FFC000")
        _set_cell_borders(cell)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.name = "Cambria"

    # Write data rows
    for i, (_, row) in enumerate(va_df.iterrows()):
        data_row = tbl.rows[i + 1]
        _set_row_height(data_row, 60)
        for j, col_name in enumerate(VA_COLUMNS):
            cell = data_row.cells[j]
            value = row.get(col_name, "")
            if pd.isna(value) or value == "":
                cell.text = "N/A"
            else:
                cell.text = str(value)
            _set_cell_borders(cell)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)
                    run.font.name = "Cambria"
                if col_name in ("Sr. no", "Risk", "Host", "Port", "CVE"):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_name in ("Description", "Recommendation ", "Reference"):
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _find_paragraph_index(doc: Document, text_fragment: str, style_name: str | None = None) -> int:
    """Find the paragraph index containing the given text fragment.

    Parameters
    ----------
    text_fragment : str
        Text to search for.
    style_name : str, optional
        If provided, only match paragraphs with this style (e.g., "Heading 2").
    """
    for i, para in enumerate(doc.paragraphs):
        if text_fragment in para.text:
            if style_name is None or para.style.name == style_name:
                return i
    return -1


def _insert_paragraph_after_table(doc: Document, table_index: int) -> None:
    """Insert a new paragraph after the specified table index.

    This is used to place content (like pie charts) immediately after a table.
    """
    if table_index < len(doc.tables):
        tbl = doc.tables[table_index]
        # Get the last row's last cell's last paragraph
        last_para = tbl.rows[-1].cells[-1].paragraphs[-1]
        # Create a new paragraph after the table
        new_para = doc.add_paragraph()
        last_para._p.addnext(new_para._p)
        return new_para
    return None


def build_word_report(
    template_path: Path | str,
    output_path: Path | str,
    metadata: EngagementMetadata,
    va_df: pd.DataFrame,
    ca_df: pd.DataFrame,
    va_risk_summary: dict[str, int],
    ca_risk_summary: dict[str, int],
) -> Path:
    """Build the Word audit report.

    Parameters
    ----------
    template_path : Path
        Path to the Word template (.docx).
    output_path : Path
        Path where the output Word report will be saved.
    metadata : EngagementMetadata
        Engagement metadata.
    va_df : pd.DataFrame
        Processed VA data (normal report, with all columns).
    ca_df : pd.DataFrame
        Processed CA data (normal report, with all columns).
    va_risk_summary : dict
        VA risk level counts including "Grand Total".
    ca_risk_summary : dict
        CA risk level counts including "Grand Total".

    Returns
    -------
    Path
        Path to the saved Word report.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    # Copy template to temp location to avoid locking
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    shutil.copy2(template_path, tmp.name)
    doc = Document(tmp.name)

    # 1. Replace client name
    _replace_client_name(doc, metadata.client_name)

    # 2. Update executive summary text
    _populate_executive_summary_text(doc, va_risk_summary, ca_risk_summary)

    # 3. Populate summary tables
    _populate_executive_summary_table(doc, va_risk_summary, ca_risk_summary)

    # 4. Add pie chart on page 14 (centered in middle of page)
    va_chart_buf = _generate_pie_chart_image(
        va_risk_summary, "Vulnerability Risk Distribution", RISK_COLORS, ca_risk_summary
    )

    # Insert page break after Executive Summary table to move chart to page 14
    # Chart dimensions: Height 9.87 cm (3.886 in), Width 9.4 cm (3.701 in)
    if len(doc.tables) > 11:
        tbl11 = doc.tables[11]
        last_para = tbl11.rows[-1].cells[-1].paragraphs[-1]
        
        # Insert page break paragraph after the table
        page_break_para = doc.add_paragraph()
        last_para._p.addnext(page_break_para._p)
        run_break = page_break_para.add_run()
        run_break.add_break(WD_BREAK.PAGE)
        
        # Add vertical spacing to center chart on page (approx 8-10 empty paragraphs)
        for _ in range(8):
            spacing_para = doc.add_paragraph()
            page_break_para._p.addnext(spacing_para._p)
            page_break_para = spacing_para
        
        # Insert VA pie chart centered on page 14
        if va_chart_buf:
            chart_para = doc.add_paragraph()
            page_break_para._p.addnext(chart_para._p)
            chart_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = chart_para.add_run()
            run.add_picture(va_chart_buf, width=Inches(3.701), height=Inches(3.886))

    # 5. Find insertion points for VA and CA tables in VULNERABILITIES DETAILS
    va_para_idx = _find_paragraph_index(doc, "The below table shows the detailed report of the VA scan done on assets.")
    ca_para_idx = _find_paragraph_index(doc, "The below table shows the detailed report of the Compliance scan done on Assets.")

    # 6. Insert CA table first (higher index) so VA index stays valid
    if ca_para_idx >= 0 and not ca_df.empty:
        _create_ca_table(doc, ca_df, ca_para_idx)

    # 7. Insert VA table with pie chart on the left side (side-by-side layout)
    if va_para_idx >= 0 and not va_df.empty:
        # Insert pie chart before the VA table text, then insert VA table
        _create_va_table_with_chart(doc, va_df, va_para_idx, va_risk_summary, ca_risk_summary)
    elif va_para_idx >= 0 and va_df.empty:
        # If VA is empty, just insert chart
        va_chart_buf2 = _generate_pie_chart_image(va_risk_summary, "Vulnerability Risk Distribution", RISK_COLORS, ca_risk_summary)
        if va_chart_buf2:
            para = doc.paragraphs[va_para_idx]
            run = para.add_run()
            run.add_picture(va_chart_buf2, width=Inches(3.701), height=Inches(3.886))

    # 8. Save
    os.makedirs(output_path.parent, exist_ok=True)
    doc.save(str(output_path))

    # Cleanup temp file
    try:
        os.unlink(tmp.name)
    except OSError:
        pass

    logger.info("Word report saved: %s", output_path)
    return output_path
