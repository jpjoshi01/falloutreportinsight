from __future__ import annotations

from typing import Any

import pandas as pd
import requests


PROVIDER_DEFAULTS = {
    "Kimi": "https://api.moonshot.ai/v1",
    "OpenAI": "https://api.openai.com/v1",
    "Claude": "https://api.anthropic.com/v1",
    "Custom": "",
}


def _ai_request(method: str, url: str, **kwargs: Any) -> requests.Response:
    """Call AI providers without inheriting broken local proxy environment vars."""
    session = requests.Session()
    session.trust_env = False
    try:
        return session.request(method, url, **kwargs)
    finally:
        session.close()


def sanitize_error(exc: Exception, api_key: str | None = None) -> str:
    message = str(exc)
    if api_key:
        message = message.replace(api_key, "[redacted]")
    lower = message.lower()
    if "proxyerror" in lower or "unable to connect to proxy" in lower or "127.0.0.1" in lower:
        return (
            "Could not reach the AI provider because this laptop is configured to use a local proxy, "
            "but that proxy is not running. Check Windows proxy/VPN settings or unset HTTP_PROXY/HTTPS_PROXY, "
            "then test the connection again."
        )
    if "winerror 10061" in lower or "actively refused" in lower:
        return (
            "Network connection was refused before reaching the AI provider. "
            "Check VPN/proxy/firewall settings, then try Test API connection again."
        )
    return message


def _compact_frame(df: pd.DataFrame, rows: int = 10) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    return df.head(rows).to_dict(orient="records")


def build_prompt(results: dict[str, Any]) -> str:
    if results.get("mode") == "current":
        payload = {
            "mode": "current_report",
            "report_name": results.get("report_name", ""),
            "totals": results.get("totals", {}),
            "top_major_fallout": _compact_frame(results.get("top_major_fallout")),
            "top_minor_fallout": _compact_frame(results.get("top_minor_fallout")),
            "top_customers": _compact_frame(results.get("top_customers")),
            "table_summary": _compact_frame(results.get("table_summary")),
            "current_validations": _compact_frame(results.get("current_validations")),
        }
        return (
            "You are summarizing a single current migration fallout report for executives. "
            "Use only this report data. Provide: 1) executive summary, "
            "2) top current risks based on major fallouts first, 3) suggested actions, 4) watchlist including minor fallouts only as secondary context. "
            "Ignore rows named Unknown validation or rows with missing validation names when identifying risks. "
            "Do not describe increases, reductions, new validations, or resolved validations unless comparison data is provided.\n\n"
            f"DATA:\n{payload}"
        )

    totals = results.get("totals", {})
    payload = {
        "mode": "comparison",
        "totals": totals,
        "new_validations": _compact_frame(results.get("new_validations")),
        "resolved_validations": _compact_frame(results.get("resolved_validations")),
        "top_major_fallout": _compact_frame(results.get("top_major_fallout")),
        "top_minor_fallout": _compact_frame(results.get("top_minor_fallout")),
        "increased_fallout": _compact_frame(results.get("increased_fallout")),
        "table_summary": _compact_frame(results.get("table_summary")),
    }
    return (
        "You are summarizing a migration fallout comparison for executives. "
        "Use only this aggregate data. Provide: 1) executive summary, "
        "2) top risks based on major fallouts first, 3) suggested actions, 4) watchlist including minor fallouts only as secondary context. "
        "Ignore rows named Unknown validation or rows with missing validation names when identifying risks.\n\n"
        f"DATA:\n{payload}"
    )


def _compact_report(report: Any, rows_per_sheet: int = 5) -> dict[str, Any]:
    if report is None:
        return {}
    sheets: dict[str, Any] = {}
    for sheet_name, df in getattr(report, "sheets", {}).items():
        if df is None:
            continue
        columns = [str(column) for column in getattr(df, "columns", [])]
        sheets[str(sheet_name)] = {
            "rows": int(len(df)),
            "columns": columns[:20],
            "sample_rows": _compact_frame(df, rows_per_sheet),
        }
    return {
        "filename": getattr(report, "filename", ""),
        "report_timestamp": str(getattr(report, "report_timestamp", "") or ""),
        "sheets": sheets,
    }


def build_question_prompt(question: str, results: dict[str, Any] | None, current_report: Any = None) -> str:
    payload = {
        "question": question,
        "latest_report": _compact_report(current_report),
        "comparison": {
            "mode": (results or {}).get("mode", "comparison" if results else ""),
            "totals": (results or {}).get("totals", {}),
            "current_validations": _compact_frame((results or {}).get("current_validations")),
            "top_major_fallout": _compact_frame((results or {}).get("top_major_fallout")),
            "top_minor_fallout": _compact_frame((results or {}).get("top_minor_fallout")),
            "top_customers": _compact_frame((results or {}).get("top_customers")),
            "new_validations": _compact_frame((results or {}).get("new_validations")),
            "resolved_validations": _compact_frame((results or {}).get("resolved_validations")),
            "increased_fallout": _compact_frame((results or {}).get("increased_fallout")),
            "reduced_fallout": _compact_frame((results or {}).get("reduced_fallout")),
            "table_summary": _compact_frame((results or {}).get("table_summary")),
        },
    }
    return (
        "You are Fallout Insight Engine for a migration fallout dashboard. "
        "Answer the user's question using only the latest report and comparison data provided below. "
        "When discussing risk or top fallout, prioritize top_major_fallout before minor fallout. "
        "Ignore rows named Unknown validation or rows with missing validation names when identifying risks. "
        "Be concise, business-friendly, and specific. If the data is not present, say what is missing and suggest where to look in the dashboard. "
        "Do not invent validation names, counts, tables, or report facts.\n\n"
        f"DATA:\n{payload}"
    )


def generate_ai_insights(
    provider: str,
    api_key: str,
    base_url: str,
    model: str,
    results: dict[str, Any],
    timeout: int = 60,
) -> str:
    if not api_key:
        raise ValueError("API key is required when AI insights are enabled.")
    if not model:
        raise ValueError("Model name is required when AI insights are enabled.")
    base = (base_url or PROVIDER_DEFAULTS.get(provider, "")).rstrip("/")
    if not base:
        raise ValueError("API base URL is required for the selected provider.")

    prompt = build_prompt(results)
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    if provider == "Claude" and "anthropic" in base:
        headers["anthropic-version"] = "2023-06-01"
        response = _ai_request(
            "POST",
            f"{base}/messages",
            headers=headers,
            json={
                "model": model,
                "max_tokens": 1200,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=timeout,
        )
        response.raise_for_status()
        data = response.json()
        return "\n".join(block.get("text", "") for block in data.get("content", []) if block.get("type") == "text")

    response = _ai_request(
        "POST",
        f"{base}/chat/completions",
        headers=headers,
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": "You produce concise migration fallout intelligence summaries."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.2,
        },
        timeout=timeout,
    )
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"]


def answer_report_question(
    provider: str,
    api_key: str,
    base_url: str,
    model: str,
    question: str,
    results: dict[str, Any] | None = None,
    current_report: Any = None,
    timeout: int = 60,
) -> str:
    if not question.strip():
        raise ValueError("Ask a question before sending.")
    if not api_key:
        raise ValueError("API key is required when AI insights are enabled.")
    if not model:
        raise ValueError("Model name is required when AI insights are enabled.")
    base = (base_url or PROVIDER_DEFAULTS.get(provider, "")).rstrip("/")
    if not base:
        raise ValueError("API base URL is required for the selected provider.")

    prompt = build_question_prompt(question, results, current_report)
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    if provider == "Claude" and "anthropic" in base:
        headers["anthropic-version"] = "2023-06-01"
        response = _ai_request(
            "POST",
            f"{base}/messages",
            headers=headers,
            json={
                "model": model,
                "max_tokens": 900,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=timeout,
        )
        response.raise_for_status()
        data = response.json()
        return "\n".join(block.get("text", "") for block in data.get("content", []) if block.get("type") == "text")

    response = _ai_request(
        "POST",
        f"{base}/chat/completions",
        headers=headers,
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": "You answer questions about migration fallout reports using only provided dashboard data."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.1,
        },
        timeout=timeout,
    )
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"]


def test_ai_connection(provider: str, api_key: str, base_url: str, model: str, timeout: int = 20) -> tuple[bool, str]:
    if not api_key:
        return False, "API key is required."
    base = (base_url or PROVIDER_DEFAULTS.get(provider, "")).rstrip("/")
    if not base:
        return False, "API base URL is required."
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        if provider == "Claude" and "anthropic" in base:
            headers["anthropic-version"] = "2023-06-01"
            response = _ai_request("GET", f"{base}/models", headers=headers, timeout=timeout)
        else:
            response = _ai_request("GET", f"{base}/models", headers=headers, timeout=timeout)
        if response.status_code < 400:
            return True, "Connection successful."
        if response.status_code in {404, 405} and model:
            return True, "Base URL reached. Model listing is not supported by this provider."
        return False, f"Connection failed with HTTP {response.status_code}."
    except Exception as exc:
        return False, sanitize_error(exc, api_key)
