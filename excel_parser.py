from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
import re
from typing import BinaryIO
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

import pandas as pd

from filename_utils import extract_timestamp_from_filename, file_sha256


COLUMN_ALIASES = {
    "rule_name": ["rule name", "rulename", "rule", "rule_name"],
    "validation_name": ["validation name", "validation", "validation_name", "check name"],
    "table_name": ["table name", "tablename", "table", "table_name"],
    "fallout_count": ["count of fallouts", "fallouts", "fallout count", "failure count", "count"],
    "customer_count": [
        "count of customers",
        "customers",
        "customer count",
        "impacted customers",
        "count of impacted customers",
    ],
}


@dataclass
class ParsedReport:
    filename: str
    file_hash: str
    report_timestamp: datetime | None
    sheets: dict[str, pd.DataFrame]
    warnings: list[str] | None = None


def _clean_column_name(column: object) -> str:
    value = str(column or "").strip().lower()
    value = re.sub(r"[\n\r\t]+", " ", value)
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _canonical_column(clean_name: str) -> str:
    for canonical, aliases in COLUMN_ALIASES.items():
        if clean_name in aliases:
            return canonical
    if "fallout" in clean_name and ("count" in clean_name or clean_name == "fallouts"):
        return "fallout_count"
    if "customer" in clean_name and ("count" in clean_name or "impact" in clean_name):
        return "customer_count"
    if "validation" in clean_name:
        return "validation_name"
    if clean_name == "rule" or "rule name" in clean_name:
        return "rule_name"
    if clean_name == "table" or "table name" in clean_name:
        return "table_name"
    return clean_name.replace(" ", "_")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized = df.copy()
    seen: dict[str, int] = {}
    columns: list[str] = []
    for column in normalized.columns:
        base = _canonical_column(_clean_column_name(column))
        seen[base] = seen.get(base, 0) + 1
        columns.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    normalized.columns = columns
    for numeric in ("fallout_count", "customer_count"):
        if numeric in normalized.columns:
            normalized[numeric] = pd.to_numeric(normalized[numeric], errors="coerce").fillna(0).astype(int)
    for text_col in ("rule_name", "validation_name", "table_name"):
        if text_col in normalized.columns:
            normalized[text_col] = normalized[text_col].fillna("").astype(str).str.strip()
    return normalized.dropna(how="all")


def _xlsx_without_styles(data: bytes) -> bytes:
    source = BytesIO(data)
    repaired = BytesIO()
    with ZipFile(source, "r") as zin, ZipFile(repaired, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.lower() == "xl/styles.xml":
                continue
            zout.writestr(item, zin.read(item.filename))
    return repaired.getvalue()


def _open_workbook(data: bytes, filename: str) -> tuple[pd.ExcelFile, list[str]]:
    warnings: list[str] = []
    try:
        return pd.ExcelFile(BytesIO(data)), warnings
    except Exception as exc:
        message = str(exc).lower()
        can_retry = filename.lower().endswith((".xlsx", ".xlsm")) and (
            "stylesheet" in message or "styles.xml" in message or "style" in message
        )
        if not can_retry:
            raise
        try:
            repaired_data = _xlsx_without_styles(data)
            warnings.append(
                "Workbook stylesheet was invalid, so formatting was ignored and sheet data was parsed from a repaired in-memory copy."
            )
            return pd.ExcelFile(BytesIO(repaired_data), engine="openpyxl"), warnings
        except (BadZipFile, Exception) as retry_exc:
            raise ValueError(f"Unable to read workbook after stylesheet repair: {retry_exc}") from exc


def parse_excel_report(file_obj: str | bytes | BinaryIO, filename: str) -> ParsedReport:
    if isinstance(file_obj, bytes):
        data = file_obj
    elif isinstance(file_obj, str):
        data = open(file_obj, "rb").read()
    else:
        data = file_obj.read()

    workbook, warnings = _open_workbook(data, filename)
    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheet_names:
        try:
            raw = pd.read_excel(workbook, sheet_name=sheet_name)
            sheets[sheet_name] = normalize_columns(raw)
        except Exception as exc:
            sheets[sheet_name] = pd.DataFrame({"parse_error": [str(exc)]})

    return ParsedReport(
        filename=filename,
        file_hash=file_sha256(data),
        report_timestamp=extract_timestamp_from_filename(filename),
        sheets=sheets,
        warnings=warnings,
    )
