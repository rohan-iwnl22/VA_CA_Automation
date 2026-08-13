"""Validate raw workbook schema and routing values."""

from __future__ import annotations

import logging
from collections import Counter

import pandas as pd

from ..logging.pipeline_logger import PipelineLogger

logger = logging.getLogger("va_ca_automation")

KNOWN_VA_RISKS = {"Critical", "High", "Medium", "Low", "None"}
KNOWN_CA_RISKS = {"PASSED", "FAILED", "WARNING"}
ALL_KNOWN_RISKS = KNOWN_VA_RISKS | KNOWN_CA_RISKS


def normalize_whitespace_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Strip leading/trailing whitespace and normalize casing for specified columns.

    Returns a copy; does not mutate the input DataFrame.
    """
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df


def validate_and_normalize_risk(df: pd.DataFrame, plogger: PipelineLogger) -> pd.DataFrame:
    """Validate Risk values, normalize casing, and log unknown values.

    Returns a copy with normalized Risk column.
    """
    df = df.copy()
    df["Risk"] = df["Risk"].astype(str).str.strip()

    # Normalize casing: VA risks are Title Case, CA risks are UPPERCASE
    raw_lower = df["Risk"].str.lower()
    va_risk_lower = {v.lower() for v in KNOWN_VA_RISKS}
    ca_risk_lower = {v.lower() for v in KNOWN_CA_RISKS}

    va_mask = raw_lower.isin(va_risk_lower)
    ca_mask = raw_lower.isin(ca_risk_lower)

    df.loc[va_mask, "Risk"] = df.loc[va_mask, "Risk"].str.title()
    df.loc[ca_mask, "Risk"] = df.loc[ca_mask, "Risk"].str.upper()

    risk_counts = Counter(df["Risk"])
    for risk_val, count in risk_counts.items():
        if risk_val not in ALL_KNOWN_RISKS:
            plogger.log_unknown_risk(risk_val, count)

    return df


def classify_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Classify rows into VA candidates, CA candidates, and unknown.

    Returns three DataFrames: (va_rows, ca_rows, unknown_rows).
    """
    va_mask = df["Risk"].isin(KNOWN_VA_RISKS)
    ca_mask = df["Risk"].isin(KNOWN_CA_RISKS)
    unknown_mask = ~(va_mask | ca_mask)

    return df[va_mask].copy(), df[ca_mask].copy(), df[unknown_mask].copy()
