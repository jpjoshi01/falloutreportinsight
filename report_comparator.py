from __future__ import annotations

import pandas as pd

from excel_parser import ParsedReport


def _validation_label(row: pd.Series) -> str:
    for column in ("validation_name", "rule_name"):
        value = str(row.get(column, "") or "").strip()
        if value and value.lower() != "nan":
            return value
    return "Unknown validation"


def aggregate_report(report: ParsedReport) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet_name, df in report.sheets.items():
        if df.empty or "parse_error" in df.columns:
            continue
        work = df.copy()
        work["validation_key"] = work.apply(_validation_label, axis=1).str.lower().str.strip()
        work["validation_name_display"] = work.apply(_validation_label, axis=1)
        if "fallout_count" not in work.columns:
            work["fallout_count"] = 0
        if "customer_count" not in work.columns:
            work["customer_count"] = 0
        if "table_name" not in work.columns:
            work["table_name"] = "Unknown"
        work["sheet_name"] = sheet_name
        frames.append(work[["validation_key", "validation_name_display", "table_name", "sheet_name", "fallout_count", "customer_count"]])

    if not frames:
        return pd.DataFrame(columns=["validation_key", "validation_name", "table_name", "fallout_count", "customer_count"])

    combined = pd.concat(frames, ignore_index=True)
    grouped = (
        combined.groupby(["validation_key", "validation_name_display", "table_name"], dropna=False)
        .agg(fallout_count=("fallout_count", "sum"), customer_count=("customer_count", "sum"))
        .reset_index()
        .rename(columns={"validation_name_display": "validation_name"})
    )
    return grouped


def compare_reports(old_report: ParsedReport, new_report: ParsedReport) -> dict[str, pd.DataFrame | dict[str, int]]:
    old = aggregate_report(old_report)
    new = aggregate_report(new_report)
    merged = old.merge(
        new,
        on=["validation_key", "validation_name", "table_name"],
        how="outer",
        suffixes=("_old", "_new"),
    ).fillna({"fallout_count_old": 0, "fallout_count_new": 0, "customer_count_old": 0, "customer_count_new": 0})

    for column in ("fallout_count_old", "fallout_count_new", "customer_count_old", "customer_count_new"):
        merged[column] = pd.to_numeric(merged[column], errors="coerce").fillna(0).astype(int)
    merged["fallout_delta"] = merged["fallout_count_new"] - merged["fallout_count_old"]
    merged["customer_delta"] = merged["customer_count_new"] - merged["customer_count_old"]
    is_new = (merged["fallout_count_old"] == 0) & (merged["fallout_count_new"] > 0)
    is_resolved = (merged["fallout_count_old"] > 0) & (merged["fallout_count_new"] == 0)
    merged["status"] = "unchanged"
    merged.loc[(merged["fallout_delta"] > 0) & ~is_new, "status"] = "increased"
    merged.loc[(merged["fallout_delta"] < 0) & ~is_resolved, "status"] = "reduced"
    merged.loc[is_new, "status"] = "new"
    merged.loc[is_resolved, "status"] = "resolved"

    table_summary = (
        merged.groupby("table_name", dropna=False)
        .agg(
            fallout_count_old=("fallout_count_old", "sum"),
            fallout_count_new=("fallout_count_new", "sum"),
            customer_count_old=("customer_count_old", "sum"),
            customer_count_new=("customer_count_new", "sum"),
            validations=("validation_key", "nunique"),
        )
        .reset_index()
    )
    table_summary["fallout_delta"] = table_summary["fallout_count_new"] - table_summary["fallout_count_old"]
    table_summary["customer_delta"] = table_summary["customer_count_new"] - table_summary["customer_count_old"]

    totals = {
        "total_fallout_old": int(merged["fallout_count_old"].sum()),
        "total_fallout_new": int(merged["fallout_count_new"].sum()),
        "total_customer_old": int(merged["customer_count_old"].sum()),
        "total_customer_new": int(merged["customer_count_new"].sum()),
        "fallout_delta": int(merged["fallout_delta"].sum()),
        "customer_delta": int(merged["customer_delta"].sum()),
        "new_validations": int((merged["status"] == "new").sum()),
        "resolved_validations": int((merged["status"] == "resolved").sum()),
    }

    return {
        "totals": totals,
        "comparison": merged.sort_values("fallout_count_new", ascending=False),
        "new_validations": merged[merged["status"] == "new"].sort_values("fallout_count_new", ascending=False),
        "resolved_validations": merged[merged["status"] == "resolved"].sort_values("fallout_count_old", ascending=False),
        "top_fallout": merged.sort_values("fallout_count_new", ascending=False).head(10),
        "top_customers": merged.sort_values("customer_count_new", ascending=False).head(10),
        "increased_fallout": merged[merged["fallout_delta"] > 0].sort_values("fallout_delta", ascending=False).head(10),
        "reduced_fallout": merged[merged["fallout_delta"] < 0].sort_values("fallout_delta", ascending=True).head(10),
        "increased_customers": merged[merged["customer_delta"] > 0].sort_values("customer_delta", ascending=False).head(10),
        "reduced_customers": merged[merged["customer_delta"] < 0].sort_values("customer_delta", ascending=True).head(10),
        "table_summary": table_summary.sort_values("fallout_count_new", ascending=False),
    }
