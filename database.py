from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from config import DB_PATH, ensure_data_dirs
from excel_parser import ParsedReport


def connect(db_path: Path = DB_PATH) -> sqlite3.Connection:
    ensure_data_dirs()
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                file_hash TEXT NOT NULL UNIQUE,
                report_timestamp TEXT,
                ingested_at TEXT NOT NULL,
                source TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id INTEGER NOT NULL,
                sheet_name TEXT NOT NULL,
                row_count INTEGER NOT NULL,
                data_json TEXT NOT NULL,
                FOREIGN KEY(report_id) REFERENCES reports(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS scan_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                scanned_at TEXT NOT NULL,
                mode TEXT NOT NULL,
                keywords TEXT,
                matched_count INTEGER NOT NULL,
                downloaded_count INTEGER NOT NULL,
                details_json TEXT
            );
            """
        )


def find_report_by_filename(filename: str) -> dict[str, Any] | None:
    init_db()
    with connect() as conn:
        row = conn.execute("SELECT * FROM reports WHERE filename = ? ORDER BY id DESC LIMIT 1", (filename,)).fetchone()
        return dict(row) if row else None


def save_report(report: ParsedReport, source: str = "manual", overwrite_same_filename: bool = False) -> int:
    init_db()
    with connect() as conn:
        if overwrite_same_filename:
            filename_rows = conn.execute("SELECT id FROM reports WHERE filename = ?", (report.filename,)).fetchall()
            for row in filename_rows:
                conn.execute("DELETE FROM sheets WHERE report_id = ?", (row["id"],))
                conn.execute("DELETE FROM reports WHERE id = ?", (row["id"],))
        existing = conn.execute("SELECT id FROM reports WHERE file_hash = ?", (report.file_hash,)).fetchone()
        if existing:
            return int(existing["id"])
        cursor = conn.execute(
            """
            INSERT INTO reports(filename, file_hash, report_timestamp, ingested_at, source)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                report.filename,
                report.file_hash,
                report.report_timestamp.isoformat() if report.report_timestamp else None,
                datetime.now().isoformat(timespec="seconds"),
                source,
            ),
        )
        report_id = int(cursor.lastrowid)
        for sheet_name, df in report.sheets.items():
            clean_df = df.where(pd.notna(df), None)
            conn.execute(
                "INSERT INTO sheets(report_id, sheet_name, row_count, data_json) VALUES (?, ?, ?, ?)",
                (report_id, sheet_name, len(clean_df), clean_df.to_json(orient="records", date_format="iso")),
            )
        return report_id


def list_reports() -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT r.*, COUNT(s.id) AS sheet_count, COALESCE(SUM(s.row_count), 0) AS row_count
            FROM reports r
            LEFT JOIN sheets s ON s.report_id = r.id
            GROUP BY r.id
            ORDER BY COALESCE(r.report_timestamp, r.ingested_at) DESC
            """
        ).fetchall()
        return [dict(row) for row in rows]


def load_report(report_id: int) -> ParsedReport:
    init_db()
    with connect() as conn:
        report_row = conn.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
        if not report_row:
            raise ValueError(f"Report id {report_id} was not found.")
        sheets: dict[str, pd.DataFrame] = {}
        for row in conn.execute("SELECT sheet_name, data_json FROM sheets WHERE report_id = ?", (report_id,)):
            sheets[row["sheet_name"]] = pd.DataFrame(json.loads(row["data_json"]))
        timestamp = datetime.fromisoformat(report_row["report_timestamp"]) if report_row["report_timestamp"] else None
        return ParsedReport(
            filename=report_row["filename"],
            file_hash=report_row["file_hash"],
            report_timestamp=timestamp,
            sheets=sheets,
        )


def delete_report(report_id: int) -> None:
    init_db()
    with connect() as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("DELETE FROM reports WHERE id = ?", (report_id,))


def save_scan_history(mode: str, keywords: str, matched_count: int, downloaded_count: int, details: list[dict[str, Any]]) -> None:
    init_db()
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO scan_history(scanned_at, mode, keywords, matched_count, downloaded_count, details_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (datetime.now().isoformat(timespec="seconds"), mode, keywords, matched_count, downloaded_count, json.dumps(details, default=str)),
        )


def list_scan_history(limit: int = 50) -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        rows = conn.execute("SELECT * FROM scan_history ORDER BY scanned_at DESC LIMIT ?", (limit,)).fetchall()
        return [dict(row) for row in rows]
