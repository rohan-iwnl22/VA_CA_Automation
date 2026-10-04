"""Generated Word reports must be Cambria 12pt, with heading sizes preserved."""
from __future__ import annotations

import hashlib
from datetime import date
from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml.ns import qn

from va_ca_automation.metadata.engagement_metadata import EngagementMetadata
from va_ca_automation.word_writer.word_report_builder import (
    CA_COLUMNS,
    VA_COLUMNS,
    build_word_report,
)

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "Word file.docx"

# Parts whose runs are visible document text (matches builder normalization).
STORY_PREFIXES = ("word/document", "word/header", "word/footer", "word/footnotes", "word/endnotes")


def _metadata() -> EngagementMetadata:
    return EngagementMetadata(
        client_name="Test Client",
        security_tester="Tester",
        reviewed_by="Reviewer",
        report_date=date(2026, 1, 15),
        report_version="1.0",
        default_device_type="Server",
    )


def _va_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Sr. no": 1,
                "Vulnerbility Title": "Weak TLS",
                "Description": "desc",
                "Risk": "High",
                "Host": "10.0.0.1",
                "Port": "443",
                "Recommendation ": "fix",
                "Reference": "r1",
                "CVE": "CVE-1",
            }
        ]
    )


def _ca_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Sr.No.": 1,
                "Title": "Weak config",
                "Host": "10.0.0.1",
                "Description": "desc",
                "Solution": "fix",
                "Risk": "FAILED",
            }
        ]
    )


def _build(tmp_path: Path) -> Path:
    """Build a report and assert the template file itself was not modified."""
    template_hash = hashlib.sha256(TEMPLATE.read_bytes()).hexdigest()
    output = build_word_report(
        TEMPLATE,
        tmp_path / "report.docx",
        _metadata(),
        _va_df(),
        _ca_df(),
        {"Critical": 0, "High": 1, "Medium": 0, "Low": 0, "Info": 0, "Grand Total": 1},
        {"FAILED": 1, "WARNING": 0, "Grand Total": 1},
    )
    assert hashlib.sha256(TEMPLATE.read_bytes()).hexdigest() == template_hash
    return output


def _story_elements(doc: Document):
    for part in doc.part.package.parts:
        partname = str(part.partname).lstrip("/")
        if partname.startswith(STORY_PREFIXES):
            element = getattr(part, "element", None)
            if element is not None:
                yield element


def test_body_styles_are_cambria_12(tmp_path):
    output = _build(tmp_path)
    doc = Document(output)

    for name in ("Normal", "Body Text", "Table Grid", "List Paragraph", "toc 2"):
        style = doc.styles[name]
        assert style.font.name == "Cambria", f"{name} font = {style.font.name}"
        assert style.font.size is not None and style.font.size.pt == 12, (
            f"{name} size = {style.font.size}"
        )

    # Heading/title styles keep their structural size but switch to Cambria.
    for name, expected_pt in [("Heading 1", 16), ("Heading 2", 13), ("Title", 28)]:
        style = doc.styles[name]
        assert style.font.name == "Cambria", f"{name} font = {style.font.name}"
        assert style.font.size is not None and style.font.size.pt == expected_pt, (
            f"{name} size = {style.font.size}, expected {expected_pt}pt"
        )


def test_every_run_is_cambria_with_no_small_sizes(tmp_path):
    output = _build(tmp_path)
    doc = Document(output)

    run_count = 0
    heading_sized_runs = 0
    for element in _story_elements(doc):
        for run in element.iter(qn("w:r")):
            r_pr = run.find(qn("w:rPr"))
            assert r_pr is not None, "run missing w:rPr"

            r_fonts = r_pr.find(qn("w:rFonts"))
            assert r_fonts is not None, "run missing w:rFonts"
            assert r_fonts.get(qn("w:ascii")) == "Cambria", (
                f"run font = {r_fonts.get(qn('w:ascii'))}"
            )
            assert r_fonts.get(qn("w:asciiTheme")) is None, "theme font left on run"

            sz = r_pr.find(qn("w:sz"))
            if sz is not None:
                half_points = int(sz.get(qn("w:val")))
                assert half_points == 24 or half_points >= 26, (
                    f"run size {half_points / 2}pt is below the 12pt body size "
                    "or the structural >=13pt sizes"
                )
                if half_points >= 26:
                    heading_sized_runs += 1
            run_count += 1

    assert run_count > 100, f"expected many runs, found {run_count}"
    assert heading_sized_runs > 0, "no heading-sized runs survived; hierarchy lost"
