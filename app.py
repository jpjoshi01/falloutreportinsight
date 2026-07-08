from __future__ import annotations

from datetime import date, time, timedelta
from datetime import datetime
from html import escape
from io import BytesIO
import json
from pathlib import Path
import re
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import streamlit.components.v1 as components

import database
from ai_insights import PROVIDER_DEFAULTS, answer_report_question, generate_ai_insights, sanitize_error, test_ai_connection
from config import APP_NAME, ATTACHMENT_DIR, DEFAULT_EMAIL_KEYWORDS, EXCEL_EXTENSIONS, ensure_data_dirs
from credential_store import (
    CredentialStoreError,
    delete_api_key,
    get_api_key,
    get_provider_preference,
    has_saved_api_key,
    mask_api_key,
    save_api_key,
    save_provider_preference,
)
from excel_parser import ParsedReport, parse_excel_report
from email_matcher import flexible_match
from filename_utils import decide_latest_report
from outlook_scanner import date_range_from_preset, scan_outlook
from report_comparator import aggregate_report, compare_reports


st.set_page_config(page_title=APP_NAME, layout="wide")

THEME_OPTIONS = ("Dark", "Light")
THEME_QUERY_KEY = "mf_theme"
NAV_OPTIONS = (
    "Dashboard",
    "Compare Reports",
    "Fallouts",
    "Outlook Scanner",
    "Stored Reports",
    "AI Insights",
    "Fallout Insight Engine",
    "Settings",
    "About",
)
AI_PROVIDER_DEFAULTS = {
    "Kimi": {"base_url": PROVIDER_DEFAULTS.get("Kimi", ""), "model": "kimi-k2"},
    "OpenAI": {"base_url": PROVIDER_DEFAULTS.get("OpenAI", ""), "model": "gpt-4o-mini"},
    "Claude": {"base_url": PROVIDER_DEFAULTS.get("Claude", ""), "model": "claude-3-5-sonnet-latest"},
    "Custom": {"base_url": "", "model": ""},
}
MIGRATION_SUMMARY_LABELS = {
    "customers",
    "customers in collection",
    "customers without products",
    "dtv/idtv lines",
    "fixed line",
    "internet product",
    "mobile lines",
    "tv lines",
    "bundles",
}


def _normalize_theme(value: Any) -> str:
    if isinstance(value, list):
        value = value[0] if value else ""
    theme = str(value or "").strip().title()
    return theme if theme in THEME_OPTIONS else "Dark"


def initialize_theme() -> str:
    query_theme = st.query_params.get(THEME_QUERY_KEY)
    if query_theme is not None:
        st.session_state["theme_mode"] = _normalize_theme(query_theme)
        st.session_state["_theme_from_default"] = False
    elif "theme_mode" not in st.session_state:
        st.session_state["theme_mode"] = "Dark"
        st.session_state["_theme_from_default"] = True
    return st.session_state["theme_mode"]


def apply_theme_class(theme: str, save_preference: bool = False) -> None:
    storage_line = (
        'window.parent.localStorage.setItem("migration_fallout_theme", theme);'
        if save_preference
        else ""
    )
    components.html(
        f"""
        <script>
        const theme = "{theme.lower()}";
        const root = window.parent.document.documentElement;
        const body = window.parent.document.body;
        root.classList.remove("theme-dark", "theme-light");
        body.classList.remove("theme-dark", "theme-light");
        root.classList.add(`theme-${{theme}}`);
        body.classList.add(`theme-${{theme}}`);
        {storage_line}
        </script>
        """,
        height=0,
        width=0,
    )


def persist_theme_preference(theme: str) -> None:
    st.query_params[THEME_QUERY_KEY] = theme.lower()
    st.session_state["_theme_from_default"] = False
    apply_theme_class(theme, save_preference=True)


def sync_theme_from_local_storage() -> None:
    components.html(
        """
        <script>
        const params = new URLSearchParams(window.parent.location.search);
        const storedTheme = window.parent.localStorage.getItem("migration_fallout_theme");
        if (!params.has("mf_theme") && (storedTheme === "dark" || storedTheme === "light")) {
            params.set("mf_theme", storedTheme);
            window.parent.location.search = params.toString();
        }
        </script>
        """,
        height=0,
        width=0,
    )


def apply_theme(mode: str) -> None:
    dark = mode == "Dark"
    theme_class = "theme-dark" if dark else "theme-light"
    primary_scale = ["#38BDF8", "#3B82F6"] if dark else ["#60A5FA", "#2563EB"]
    st.session_state["chart_theme"] = {
        "primary_scale": primary_scale,
        "plot_bg": "rgba(15,23,42,0.35)" if dark else "rgba(241,245,249,0.82)",
        "font": "#CBD5E1" if dark else "#0F172A",
        "title": "#F8FAFC" if dark else "#0F172A",
        "grid": "rgba(51,65,85,0.45)" if dark else "rgba(203,213,225,0.82)",
        "status": {
            "new": "background-color: rgba(37, 99, 235, 0.20); color: #BFDBFE; font-weight: 800;"
            if dark
            else "background-color: rgba(37, 99, 235, 0.12); color: #1D4ED8; font-weight: 800;",
            "resolved": "background-color: rgba(34, 197, 94, 0.18); color: #BBF7D0; font-weight: 800;"
            if dark
            else "background-color: rgba(22, 163, 74, 0.12); color: #166534; font-weight: 800;",
            "increased": "background-color: rgba(239, 68, 68, 0.20); color: #FECACA; font-weight: 800;"
            if dark
            else "background-color: rgba(220, 38, 38, 0.12); color: #991B1B; font-weight: 800;",
            "reduced": "background-color: rgba(34, 197, 94, 0.18); color: #BBF7D0; font-weight: 800;"
            if dark
            else "background-color: rgba(22, 163, 74, 0.12); color: #166534; font-weight: 800;",
            "unchanged": "background-color: rgba(148, 163, 184, 0.14); color: #CBD5E1; font-weight: 800;"
            if dark
            else "background-color: rgba(100, 116, 139, 0.10); color: #475569; font-weight: 800;",
            "default": "color: #CBD5E1; font-weight: 700;" if dark else "color: #475569; font-weight: 700;",
        },
    }
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root,
        .theme-dark {{
            --mf-bg: #0B1120;
            --mf-panel: #111827;
            --mf-panel-2: #1E293B;
            --mf-border: #334155;
            --mf-primary: #38BDF8;
            --mf-primary-2: #3B82F6;
            --mf-success: #22C55E;
            --mf-warning: #F59E0B;
            --mf-error: #EF4444;
            --mf-text: #F8FAFC;
            --mf-muted: #94A3B8;
            --mf-sidebar-bg: linear-gradient(180deg, #020617 0%, #0F172A 55%, #111827 100%);
            --mf-sidebar-text: #CBD5E1;
            --mf-brand-title: #F8FAFC;
            --mf-brand-subtitle: #94A3B8;
            --mf-header-bg: linear-gradient(135deg, rgba(17, 24, 39, 0.96), rgba(30, 41, 59, 0.86));
            --mf-card-bg: linear-gradient(135deg, rgba(56, 189, 248, 0.10), transparent 38%), linear-gradient(180deg, rgba(30, 41, 59, 0.94), rgba(17, 24, 39, 0.95));
            --mf-section-bg: rgba(17, 24, 39, 0.68);
            --mf-soft-bg: rgba(15, 23, 42, 0.72);
            --mf-control-bg: #0F172A;
            --mf-control-alt: #1E293B;
            --mf-table-header-bg: #1E293B;
            --mf-table-header-text: #F8FAFC;
            --mf-table-header-border: rgba(56, 189, 248, 0.28);
            --mf-button-bg: linear-gradient(180deg, #1E293B, #0F172A);
            --mf-button-text: #E2E8F0;
            --mf-shadow: 0 18px 50px rgba(2, 8, 23, 0.35);
            --mf-card-shadow: 0 16px 38px rgba(2, 8, 23, 0.28);
            --mf-primary-soft: rgba(56, 189, 248, 0.14);
            --mf-primary-line: rgba(56, 189, 248, 0.36);
            --mf-border-soft: rgba(51, 65, 85, 0.78);
            --mf-app-glow-a: rgba(56, 189, 248, 0.14);
            --mf-app-glow-b: rgba(59, 130, 246, 0.10);
            --mf-radius: 18px;
        }}

        .theme-light {{
            --mf-bg: #F8FAFC;
            --mf-panel: #FFFFFF;
            --mf-panel-2: #F1F5F9;
            --mf-border: #CBD5E1;
            --mf-primary: #2563EB;
            --mf-primary-2: #1D4ED8;
            --mf-success: #16A34A;
            --mf-warning: #D97706;
            --mf-error: #DC2626;
            --mf-text: #0F172A;
            --mf-muted: #475569;
            --mf-sidebar-bg: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 52%, #EEF2FF 100%);
            --mf-sidebar-text: #475569;
            --mf-brand-title: #0F172A;
            --mf-brand-subtitle: #475569;
            --mf-header-bg: linear-gradient(135deg, rgba(255, 255, 255, 0.98), rgba(241, 245, 249, 0.96));
            --mf-card-bg: linear-gradient(135deg, rgba(37, 99, 235, 0.06), transparent 42%), #FFFFFF;
            --mf-section-bg: rgba(255, 255, 255, 0.88);
            --mf-soft-bg: rgba(241, 245, 249, 0.92);
            --mf-control-bg: #FFFFFF;
            --mf-control-alt: #F1F5F9;
            --mf-table-header-bg: #DBEAFE;
            --mf-table-header-text: #0F172A;
            --mf-table-header-border: rgba(37, 99, 235, 0.22);
            --mf-button-bg: linear-gradient(180deg, #FFFFFF, #F1F5F9);
            --mf-button-text: #0F172A;
            --mf-shadow: 0 18px 46px rgba(15, 23, 42, 0.10);
            --mf-card-shadow: 0 16px 34px rgba(15, 23, 42, 0.09);
            --mf-primary-soft: rgba(37, 99, 235, 0.10);
            --mf-primary-line: rgba(37, 99, 235, 0.34);
            --mf-border-soft: rgba(203, 213, 225, 0.92);
            --mf-app-glow-a: rgba(37, 99, 235, 0.08);
            --mf-app-glow-b: rgba(22, 163, 74, 0.05);
            --mf-radius: 18px;
        }}

        html:not(.theme-light):not(.theme-dark),
        body:not(.theme-light):not(.theme-dark) {{
            color-scheme: {"dark" if dark else "light"};
        }}

        html.{theme_class},
        body.{theme_class} {{
            color-scheme: {"dark" if dark else "light"};
        }}

        html:not(.theme-light):not(.theme-dark) {{
            --mf-bg: {"#0B1120" if dark else "#F8FAFC"};
            --mf-panel: {"#111827" if dark else "#FFFFFF"};
            --mf-panel-2: {"#1E293B" if dark else "#F1F5F9"};
            --mf-border: {"#334155" if dark else "#CBD5E1"};
            --mf-primary: {"#38BDF8" if dark else "#2563EB"};
            --mf-primary-2: {"#3B82F6" if dark else "#1D4ED8"};
            --mf-success: {"#22C55E" if dark else "#16A34A"};
            --mf-warning: {"#F59E0B" if dark else "#D97706"};
            --mf-error: {"#EF4444" if dark else "#DC2626"};
            --mf-text: {"#F8FAFC" if dark else "#0F172A"};
            --mf-muted: {"#94A3B8" if dark else "#475569"};
            --mf-sidebar-bg: {"linear-gradient(180deg, #020617 0%, #0F172A 55%, #111827 100%)" if dark else "linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 52%, #EEF2FF 100%)"};
            --mf-sidebar-text: {"#CBD5E1" if dark else "#475569"};
            --mf-brand-title: {"#F8FAFC" if dark else "#0F172A"};
            --mf-brand-subtitle: {"#94A3B8" if dark else "#475569"};
            --mf-header-bg: {"linear-gradient(135deg, rgba(17, 24, 39, 0.96), rgba(30, 41, 59, 0.86))" if dark else "linear-gradient(135deg, rgba(255, 255, 255, 0.98), rgba(241, 245, 249, 0.96))"};
            --mf-card-bg: {"linear-gradient(135deg, rgba(56, 189, 248, 0.10), transparent 38%), linear-gradient(180deg, rgba(30, 41, 59, 0.94), rgba(17, 24, 39, 0.95))" if dark else "linear-gradient(135deg, rgba(37, 99, 235, 0.06), transparent 42%), #FFFFFF"};
            --mf-section-bg: {"rgba(17, 24, 39, 0.68)" if dark else "rgba(255, 255, 255, 0.88)"};
            --mf-soft-bg: {"rgba(15, 23, 42, 0.72)" if dark else "rgba(241, 245, 249, 0.92)"};
            --mf-control-bg: {"#0F172A" if dark else "#FFFFFF"};
            --mf-control-alt: {"#1E293B" if dark else "#F1F5F9"};
            --mf-table-header-bg: {"#1E293B" if dark else "#DBEAFE"};
            --mf-table-header-text: {"#F8FAFC" if dark else "#0F172A"};
            --mf-table-header-border: {"rgba(56, 189, 248, 0.28)" if dark else "rgba(37, 99, 235, 0.22)"};
            --mf-button-bg: {"linear-gradient(180deg, #1E293B, #0F172A)" if dark else "linear-gradient(180deg, #FFFFFF, #F1F5F9)"};
            --mf-button-text: {"#E2E8F0" if dark else "#0F172A"};
            --mf-shadow: {"0 18px 50px rgba(2, 8, 23, 0.35)" if dark else "0 18px 46px rgba(15, 23, 42, 0.10)"};
            --mf-card-shadow: {"0 16px 38px rgba(2, 8, 23, 0.28)" if dark else "0 16px 34px rgba(15, 23, 42, 0.09)"};
            --mf-primary-soft: {"rgba(56, 189, 248, 0.14)" if dark else "rgba(37, 99, 235, 0.10)"};
            --mf-primary-line: {"rgba(56, 189, 248, 0.36)" if dark else "rgba(37, 99, 235, 0.34)"};
            --mf-border-soft: {"rgba(51, 65, 85, 0.78)" if dark else "rgba(203, 213, 225, 0.92)"};
            --mf-app-glow-a: {"rgba(56, 189, 248, 0.14)" if dark else "rgba(37, 99, 235, 0.08)"};
            --mf-app-glow-b: {"rgba(59, 130, 246, 0.10)" if dark else "rgba(22, 163, 74, 0.05)"};
            --mf-radius: 18px;
        }}

        html, body, [class*="css"] {{
            font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }}
        .stApp {{
            background:
                radial-gradient(circle at 20% 0%, var(--mf-app-glow-a), transparent 30rem),
                radial-gradient(circle at 85% 10%, var(--mf-app-glow-b), transparent 28rem),
                var(--mf-bg);
            color: var(--mf-text);
        }}
        .block-container {{
            max-width: 1480px;
            padding-top: 2rem;
            padding-bottom: 3rem;
        }}
        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
        .stApp p, .stApp label, .stApp span, .stApp div {{
            color: var(--mf-text);
        }}
        .stCaptionContainer, .stCaptionContainer * {{
            color: var(--mf-muted) !important;
        }}

        /* Keep Streamlit's sidebar expand/collapse controls visible after custom theming. */
        [data-testid="stHeader"] {{
            background: color-mix(in srgb, var(--mf-bg) 88%, transparent) !important;
        }}
        [data-testid="stSidebarCollapseButton"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="stSidebarCollapseButton"] button,
        [data-testid="stSidebarCollapsedControl"] button {{
            opacity: 1 !important;
            color: var(--mf-primary) !important;
            background: var(--mf-control-bg) !important;
            border: 1px solid var(--mf-border-soft) !important;
            border-radius: 0.75rem !important;
            box-shadow: var(--mf-card-shadow) !important;
        }}
        [data-testid="stSidebarCollapseButton"]:hover,
        [data-testid="stSidebarCollapsedControl"]:hover,
        [data-testid="stSidebarCollapseButton"] button:hover,
        [data-testid="stSidebarCollapsedControl"] button:hover {{
            background: var(--mf-primary-soft) !important;
            border-color: var(--mf-primary-line) !important;
        }}
        [data-testid="stSidebarCollapseButton"] svg,
        [data-testid="stSidebarCollapsedControl"] svg,
        button[aria-label*="sidebar" i] svg {{
            opacity: 1 !important;
            color: var(--mf-primary) !important;
            fill: var(--mf-primary) !important;
            stroke: var(--mf-primary) !important;
        }}
        button[aria-label*="sidebar" i] {{
            opacity: 1 !important;
            color: var(--mf-primary) !important;
        }}

        /* Sidebar layout and navigation styling. */
        [data-testid="stSidebar"] {{
            background: var(--mf-sidebar-bg);
            border-right: 1px solid var(--mf-border-soft);
            box-shadow: 10px 0 30px rgba(2, 8, 23, 0.22);
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] span {{
            color: var(--mf-sidebar-text) !important;
        }}
        .mf-brand {{
            display: flex;
            align-items: center;
            gap: 0.8rem;
            padding: 0.85rem;
            margin-bottom: 0.8rem;
            border: 1px solid var(--mf-primary-line);
            border-radius: var(--mf-radius);
            background: linear-gradient(135deg, var(--mf-primary-soft), rgba(59, 130, 246, 0.08));
        }}
        .mf-logo {{
            width: 2.35rem;
            height: 2.35rem;
            display: grid;
            place-items: center;
            border-radius: 0.85rem;
            background: linear-gradient(135deg, var(--mf-primary), var(--mf-primary-2));
            color: #FFFFFF;
            font-weight: 800;
            box-shadow: 0 12px 28px rgba(56, 189, 248, 0.22);
        }}
        .mf-brand-title {{ font-size: 0.96rem; font-weight: 800; color: var(--mf-brand-title) !important; line-height: 1.1; }}
        .mf-brand-subtitle {{ font-size: 0.76rem; color: var(--mf-brand-subtitle) !important; margin-top: 0.15rem; }}
        [data-testid="stSidebar"] .stRadio > label {{
            font-weight: 850 !important;
            color: var(--mf-brand-title) !important;
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] {{
            display: grid;
            gap: 0.42rem;
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"][aria-label="Theme"] {{
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"][aria-label="Theme"] label {{
            min-height: 2.6rem;
            justify-content: center;
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] label {{
            min-height: 2.9rem;
            padding: 0.5rem 0.75rem;
            border: 1px solid transparent;
            border-radius: 0.85rem;
            background: transparent;
            transition: all 160ms ease;
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:hover {{
            background: var(--mf-primary-soft);
            border-color: var(--mf-primary-line);
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:has(input:checked) {{
            background: linear-gradient(90deg, var(--mf-primary-soft), rgba(59, 130, 246, 0.08));
            border-color: var(--mf-primary-line);
        }}
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:has(input:checked) p {{
            color: var(--mf-brand-title) !important;
            font-weight: 850 !important;
        }}

        /* Header, cards, badges, and page scaffolding. */
        .mf-header {{
            position: relative;
            display: grid;
            grid-template-columns: minmax(12rem, 0.62fr) minmax(18rem, 1fr) minmax(15rem, 0.78fr);
            align-items: center;
            gap: 1rem;
            padding: 1.1rem 1.15rem;
            margin-bottom: 1.05rem;
            border: 1px solid var(--mf-border-soft);
            border-radius: 22px;
            background:
                radial-gradient(circle at 14% 22%, color-mix(in srgb, var(--mf-primary) 22%, transparent), transparent 20rem),
                radial-gradient(circle at 74% 32%, color-mix(in srgb, var(--mf-success) 14%, transparent), transparent 18rem),
                linear-gradient(135deg, color-mix(in srgb, var(--mf-control-bg) 88%, transparent), color-mix(in srgb, var(--mf-panel-2) 72%, transparent));
            box-shadow: var(--mf-shadow);
            overflow: hidden;
        }}
        .mf-header::before {{
            content: "";
            position: absolute;
            inset: 0;
            pointer-events: none;
            background:
                linear-gradient(90deg, transparent, color-mix(in srgb, var(--mf-primary) 10%, transparent), transparent),
                radial-gradient(circle at 35% 100%, color-mix(in srgb, var(--mf-primary-2) 12%, transparent), transparent 18rem);
            opacity: 0.75;
        }}
        .mf-header > * {{
            position: relative;
            z-index: 1;
        }}
        .mf-hero-visual {{
            min-height: 8.6rem;
            border-radius: 16px;
            padding: 0.78rem;
            border: 1px solid color-mix(in srgb, var(--mf-primary-line) 80%, transparent);
            background:
                linear-gradient(180deg, color-mix(in srgb, var(--mf-panel-2) 56%, transparent), color-mix(in srgb, var(--mf-control-bg) 82%, transparent));
            box-shadow: inset 0 1px 0 color-mix(in srgb, var(--mf-primary) 16%, transparent);
        }}
        .mf-mini-chart {{
            height: 4.2rem;
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            align-items: end;
            gap: 0.42rem;
            padding: 0.5rem 0.25rem 0.1rem;
            border-bottom: 1px solid var(--mf-border-soft);
            background:
                linear-gradient(color-mix(in srgb, var(--mf-border) 42%, transparent) 1px, transparent 1px) 0 10% / 100% 36%,
                transparent;
        }}
        .mf-mini-bar {{
            border-radius: 0.35rem 0.35rem 0 0;
            background: linear-gradient(180deg, var(--mf-success), var(--mf-primary-2));
            box-shadow: 0 0 20px color-mix(in srgb, var(--mf-primary) 22%, transparent);
            opacity: 0.9;
        }}
        .mf-mini-line {{
            height: 1.85rem;
            margin: 0.35rem 0 0.1rem;
            border-radius: 999px;
            background:
                linear-gradient(135deg, transparent 8%, var(--mf-primary) 8% 11%, transparent 11% 30%, var(--mf-primary-2) 30% 33%, transparent 33% 55%, var(--mf-success) 55% 58%, transparent 58%),
                color-mix(in srgb, var(--mf-primary) 5%, transparent);
            border: 1px solid color-mix(in srgb, var(--mf-primary) 15%, transparent);
        }}
        .mf-hero-stat {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.7rem;
            margin-top: 0.55rem;
        }}
        .mf-donut {{
            width: 2.75rem;
            height: 2.75rem;
            border-radius: 50%;
            background: conic-gradient(var(--mf-primary) 0 72%, color-mix(in srgb, var(--mf-panel-2) 78%, transparent) 72% 100%);
            position: relative;
            box-shadow: 0 0 22px color-mix(in srgb, var(--mf-primary) 18%, transparent);
        }}
        .mf-donut::after {{
            content: "";
            position: absolute;
            inset: 0.62rem;
            border-radius: 50%;
            background: var(--mf-control-bg);
        }}
        .mf-hero-kicker {{
            color: var(--mf-primary) !important;
            font-size: 0.7rem;
            font-weight: 850;
            letter-spacing: 0.12em;
            text-transform: uppercase;
        }}
        .mf-eyebrow {{
            color: var(--mf-primary) !important;
            font-size: 0.76rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }}
        .mf-title {{
            color: var(--mf-text) !important;
            font-size: clamp(1.5rem, 2vw, 2.15rem);
            line-height: 1.08;
            font-weight: 900;
            margin-top: 0.25rem;
        }}
        .mf-gradient-word {{
            color: transparent !important;
            background: linear-gradient(90deg, var(--mf-primary), var(--mf-success));
            -webkit-background-clip: text;
            background-clip: text;
        }}
        .mf-subtitle {{
            color: var(--mf-muted) !important;
            max-width: 29rem;
            font-size: 0.8rem;
            margin-top: 0.5rem;
            line-height: 1.45;
        }}
        .mf-hero-features {{
            display: grid;
            grid-template-columns: repeat(2, minmax(7.6rem, 1fr));
            gap: 0.55rem;
            margin-top: 0.8rem;
            max-width: 21rem;
        }}
        .mf-feature-chip {{
            display: flex;
            align-items: center;
            gap: 0.45rem;
            padding: 0.42rem 0.54rem;
            border-radius: 0.8rem;
            background: color-mix(in srgb, var(--mf-control-bg) 78%, transparent);
            border: 1px solid color-mix(in srgb, var(--mf-border) 70%, transparent);
            color: var(--mf-muted) !important;
            font-size: 0.72rem;
            font-weight: 750;
            min-width: 0;
            white-space: nowrap;
        }}
        .mf-feature-icon {{
            width: 1.3rem;
            height: 1.3rem;
            min-width: 1.3rem;
            display: inline-grid;
            place-items: center;
            border-radius: 0.45rem;
            color: var(--mf-primary) !important;
            background: var(--mf-primary-soft);
            font-size: 0.72rem;
            font-weight: 900;
        }}
        .mf-header-actions {{
            display: grid;
            grid-template-columns: 1fr;
            align-items: stretch;
            gap: 0.6rem;
            min-width: 0;
            align-self: center;
        }}
        .mf-report-pill {{
            padding: 0.56rem 0.66rem;
            border-radius: 0.9rem;
            border: 1px solid var(--mf-primary-line);
            background: var(--mf-primary-soft);
        }}
        .mf-report-label {{
            color: var(--mf-muted) !important;
            font-size: 0.68rem;
            font-weight: 800;
            text-transform: uppercase;
            margin-bottom: 0.22rem;
        }}
        .mf-report-name {{
            color: var(--mf-primary) !important;
            font-size: 0.74rem;
            font-weight: 850;
            line-height: 1.25;
            overflow-wrap: anywhere;
            word-break: break-word;
        }}
        .mf-status-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            justify-content: flex-start;
            align-items: center;
        }}
        .mf-chip, .mf-badge {{
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.36rem 0.5rem;
            border-radius: 999px;
            border: 1px solid var(--mf-border-soft);
            background: var(--mf-soft-bg);
            color: var(--mf-muted) !important;
            font-size: 0.72rem;
            font-weight: 700;
            max-width: 100%;
            white-space: normal;
        }}
        @media (max-width: 1050px) {{
            .mf-header {{
                grid-template-columns: 1fr;
            }}
            .mf-hero-visual {{
                max-width: 28rem;
            }}
        }}
        .mf-badge.success {{ color: var(--mf-success) !important; border-color: color-mix(in srgb, var(--mf-success) 36%, transparent); background: color-mix(in srgb, var(--mf-success) 12%, transparent); }}
        .mf-badge.warning {{ color: var(--mf-warning) !important; border-color: color-mix(in srgb, var(--mf-warning) 36%, transparent); background: color-mix(in srgb, var(--mf-warning) 12%, transparent); }}
        .mf-badge.info {{ color: var(--mf-primary) !important; border-color: var(--mf-primary-line); background: var(--mf-primary-soft); }}
        .mf-card {{
            padding: 1rem;
            border: 1px solid var(--mf-border-soft);
            border-radius: var(--mf-radius);
            background: var(--mf-card-bg);
            box-shadow: var(--mf-card-shadow);
            min-height: 9.2rem;
            overflow: hidden;
        }}
        .mf-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(13rem, 1fr));
            gap: 1rem;
            margin: 1.2rem 0 1.4rem;
            align-items: stretch;
        }}
        .mf-kpi-pair {{
            display: grid;
            gap: 1rem;
        }}
        .mf-card-top {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
        }}
        .mf-card-icon {{
            width: 2.35rem;
            height: 2.35rem;
            min-width: 2.35rem;
            display: grid;
            place-items: center;
            border-radius: 0.9rem;
            background: var(--mf-primary-soft);
            border: 1px solid var(--mf-primary-line);
            color: var(--mf-primary) !important;
            font-size: 1.05rem;
            font-weight: 850;
        }}
        .mf-card-label {{ color: var(--mf-muted) !important; font-size: 0.76rem; font-weight: 800; text-transform: uppercase; line-height: 1.25; }}
        .mf-card-value {{ color: var(--mf-text) !important; font-size: 1.8rem; font-weight: 850; margin-top: 0.85rem; line-height: 1.05; }}
        .mf-card-foot {{ color: var(--mf-muted) !important; font-size: 0.82rem; margin-top: 0.45rem; line-height: 1.35; }}
        .mf-ai-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(10rem, 1fr));
            gap: 0.9rem;
            margin: 1rem 0 1.25rem;
        }}
        .mf-ai-kpi-card {{
            min-height: 8.4rem;
        }}
        .mf-ai-kpi-card .mf-card-value {{
            font-size: clamp(1rem, 1.45vw, 1.45rem);
            overflow-wrap: anywhere;
        }}
        .mf-ai-section-head {{
            display: flex;
            gap: 0.5rem;
            margin-bottom: 0.8rem;
        }}
        [data-testid="stExpander"] {{
            border: 1px solid var(--mf-border-soft) !important;
            border-radius: var(--mf-radius) !important;
            background: var(--mf-section-bg) !important;
            box-shadow: var(--mf-card-shadow);
            overflow: hidden;
            margin-bottom: 0.9rem;
        }}
        [data-testid="stExpander"] summary {{
            color: var(--mf-text) !important;
            font-weight: 850 !important;
        }}
        [data-testid="stExpander"] summary *,
        [data-testid="stExpander"] summary svg {{
            color: var(--mf-text) !important;
            fill: var(--mf-text) !important;
            stroke: var(--mf-text) !important;
        }}
        .theme-light [data-testid="stExpander"] {{
            background: #FFFFFF !important;
            border-color: rgba(203, 213, 225, 0.95) !important;
            box-shadow: 0 14px 30px rgba(15, 23, 42, 0.08) !important;
        }}
        .theme-light [data-testid="stExpander"] summary {{
            background: linear-gradient(180deg, #F8FAFC, #EEF2FF) !important;
            color: #0F172A !important;
            border-bottom: 1px solid rgba(203, 213, 225, 0.72);
        }}
        .theme-light [data-testid="stExpander"] summary *,
        .theme-light [data-testid="stExpander"] summary svg {{
            color: #0F172A !important;
            fill: #0F172A !important;
            stroke: #0F172A !important;
        }}
        .mf-ai-risk-separator {{
            height: 1px;
            margin: 0.95rem 0 1.05rem;
            background: var(--mf-border-soft);
            opacity: 0.7;
        }}
        @media (max-width: 1200px) {{
            .mf-ai-kpi-grid {{ grid-template-columns: repeat(2, minmax(12rem, 1fr)); }}
        }}
        @media (max-width: 720px) {{
            .mf-ai-kpi-grid {{ grid-template-columns: 1fr; }}
        }}
        .mf-action-toolbar {{
            display: flex;
            align-items: center;
            gap: 0.85rem;
            margin: 0.95rem 0 1rem;
            max-width: 32rem;
        }}
        .mf-action-toolbar [data-testid="column"] {{
            width: 9.6rem !important;
            min-width: 9.6rem !important;
            flex: 0 0 9.6rem !important;
        }}
        .mf-action-toolbar [data-testid="stButton"] button,
        .mf-action-toolbar [data-testid="stDownloadButton"] button {{
            min-height: 2.25rem !important;
            height: 2.25rem !important;
            padding: 0.35rem 0.85rem !important;
            border-radius: 0.55rem !important;
            font-size: 0.78rem !important;
            font-weight: 800 !important;
            justify-content: center !important;
            white-space: nowrap !important;
            min-width: 9rem !important;
            background: color-mix(in srgb, var(--mf-control-bg) 82%, transparent) !important;
            border: 1px solid var(--mf-border-soft) !important;
            color: var(--mf-text) !important;
            box-shadow: 0 8px 20px rgba(2, 8, 23, 0.14) !important;
        }}
        .mf-action-toolbar [data-testid="stButton"] button *,
        .mf-action-toolbar [data-testid="stDownloadButton"] button *,
        .mf-action-toolbar [data-testid="stButton"] button p,
        .mf-action-toolbar [data-testid="stDownloadButton"] button p {{
            white-space: nowrap !important;
            word-break: keep-all !important;
            overflow-wrap: normal !important;
            line-height: 1 !important;
        }}
        .mf-action-toolbar [data-testid="stButton"] button:hover,
        .mf-action-toolbar [data-testid="stDownloadButton"] button:hover {{
            border-color: var(--mf-primary-line) !important;
            background: var(--mf-primary-soft) !important;
        }}
        [data-testid="stButton"] button p,
        [data-testid="stDownloadButton"] button p,
        [data-testid="stButton"] button span,
        [data-testid="stDownloadButton"] button span {{
            white-space: nowrap !important;
            word-break: keep-all !important;
            overflow-wrap: normal !important;
        }}
        .mf-section {{
            padding: 1rem;
            border: 1px solid var(--mf-border-soft);
            border-radius: var(--mf-radius);
            background: var(--mf-section-bg);
            box-shadow: 0 14px 32px rgba(2, 8, 23, 0.16);
        }}

        /* Streamlit controls: polished forms, uploaders, buttons, tabs, alerts. */
        [data-testid="stTabs"] [role="tablist"] {{
            gap: 0.4rem;
            border-bottom: 1px solid var(--mf-border-soft);
            padding-bottom: 0.35rem;
        }}
        [data-testid="stTabs"] [role="tab"] {{
            color: var(--mf-muted) !important;
            background: var(--mf-soft-bg) !important;
            border: 1px solid var(--mf-border-soft) !important;
            border-radius: 999px !important;
            padding: 0.58rem 0.86rem !important;
            transition: all 160ms ease;
        }}
        [data-testid="stTabs"] [role="tab"] p {{ color: inherit !important; font-weight: 700; }}
        [data-testid="stTabs"] [role="tab"]:hover {{
            color: var(--mf-text) !important;
            background: var(--mf-primary-soft) !important;
            border-color: var(--mf-primary-line) !important;
        }}
        [data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
            color: var(--mf-text) !important;
            background: linear-gradient(135deg, var(--mf-primary-soft), color-mix(in srgb, var(--mf-primary-2) 14%, transparent)) !important;
            border-color: var(--mf-primary-line) !important;
            box-shadow: 0 10px 26px rgba(56, 189, 248, 0.12);
        }}
        [data-testid="stTabs"] button[aria-label],
        [data-testid="stTabs"] [data-testid="stTabScrollButton"] {{
            color: var(--mf-text) !important;
            background: var(--mf-control-alt) !important;
            border: 1px solid var(--mf-border) !important;
            border-radius: 0.7rem !important;
        }}
        .stButton > button,
        .stDownloadButton > button {{
            min-height: 2.55rem;
            color: var(--mf-button-text) !important;
            background: var(--mf-button-bg) !important;
            border: 1px solid var(--mf-border) !important;
            border-radius: 0.8rem !important;
            box-shadow: 0 8px 22px rgba(2, 8, 23, 0.18);
            font-weight: 800 !important;
            transition: all 160ms ease;
        }}
        .stButton > button:hover,
        .stDownloadButton > button:hover {{
            color: var(--mf-text) !important;
            transform: translateY(-1px);
            border-color: var(--mf-primary-line) !important;
            box-shadow: 0 12px 28px rgba(56, 189, 248, 0.12);
        }}
        .theme-light .stButton > button,
        .theme-light .stDownloadButton > button {{
            color: #0F172A !important;
            background: linear-gradient(180deg, #FFFFFF, #F8FAFC) !important;
            border-color: #CBD5E1 !important;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08) !important;
        }}
        .theme-light .stButton > button *,
        .theme-light .stDownloadButton > button * {{
            color: #0F172A !important;
        }}
        .theme-light .stButton > button:hover,
        .theme-light .stDownloadButton > button:hover {{
            color: #0F172A !important;
            background: linear-gradient(180deg, #EFF6FF, #FFFFFF) !important;
            border-color: rgba(37, 99, 235, 0.42) !important;
        }}
        html.theme-light button[data-testid^="baseButton-secondary"],
        html.theme-light button[data-testid^="baseButton-minimal"],
        html.theme-light button[data-testid^="baseButton-tertiary"],
        body.theme-light button[data-testid^="baseButton-secondary"],
        body.theme-light button[data-testid^="baseButton-minimal"],
        body.theme-light button[data-testid^="baseButton-tertiary"],
        html.theme-light [data-testid="stButton"] button:not([kind="primary"]),
        html.theme-light [data-testid="stDownloadButton"] button:not([kind="primary"]),
        body.theme-light [data-testid="stButton"] button:not([kind="primary"]),
        body.theme-light [data-testid="stDownloadButton"] button:not([kind="primary"]) {{
            color: #0F172A !important;
            background: linear-gradient(180deg, #FFFFFF, #F8FAFC) !important;
            border: 1px solid #CBD5E1 !important;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08) !important;
        }}
        html.theme-light button[data-testid^="baseButton-secondary"] *,
        html.theme-light button[data-testid^="baseButton-minimal"] *,
        html.theme-light button[data-testid^="baseButton-tertiary"] *,
        html.theme-light [data-testid="stButton"] button:not([kind="primary"]) *,
        html.theme-light [data-testid="stDownloadButton"] button:not([kind="primary"]) *,
        body.theme-light button[data-testid^="baseButton-secondary"] *,
        body.theme-light button[data-testid^="baseButton-minimal"] *,
        body.theme-light button[data-testid^="baseButton-tertiary"] *,
        body.theme-light [data-testid="stButton"] button:not([kind="primary"]) *,
        body.theme-light [data-testid="stDownloadButton"] button:not([kind="primary"]) * {{
            color: #0F172A !important;
            fill: #0F172A !important;
            stroke: #0F172A !important;
        }}
        html.theme-light [data-testid="stFileUploader"] button,
        body.theme-light [data-testid="stFileUploader"] button {{
            color: #0F172A !important;
            background: linear-gradient(180deg, #EEF2FF, #FFFFFF) !important;
            border: 1px solid #CBD5E1 !important;
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.08) !important;
        }}
        html.theme-light [data-testid="stFileUploader"] button *,
        body.theme-light [data-testid="stFileUploader"] button * {{
            color: #0F172A !important;
            fill: #0F172A !important;
            stroke: #0F172A !important;
        }}
        html.theme-light [data-testid="stStatusWidget"],
        body.theme-light [data-testid="stStatusWidget"] {{
            color: #0F172A !important;
            background: rgba(255, 255, 255, 0.86) !important;
            border-color: #CBD5E1 !important;
        }}
        html.theme-light [data-testid="stStatusWidget"] *,
        body.theme-light [data-testid="stStatusWidget"] * {{
            color: #0F172A !important;
            fill: #2563EB !important;
            stroke: #2563EB !important;
        }}
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {{
            color: #FFFFFF !important;
            background: linear-gradient(135deg, var(--mf-primary), var(--mf-primary-2)) !important;
            border-color: var(--mf-primary-line) !important;
        }}
        .stButton > button:focus,
        .stDownloadButton > button:focus,
        input:focus, textarea:focus {{
            outline: 2px solid var(--mf-primary-line) !important;
            outline-offset: 2px !important;
        }}
        .stTextInput input,
        .stTextArea textarea,
        .stSelectbox div[data-baseweb="select"] > div,
        .stDateInput input {{
            color: var(--mf-text) !important;
            background: var(--mf-control-bg) !important;
            border: 1px solid var(--mf-border) !important;
            border-radius: 0.85rem !important;
        }}
        .stFileUploader section {{
            background:
                linear-gradient(135deg, var(--mf-primary-soft), transparent 45%),
                var(--mf-soft-bg) !important;
            border: 1px dashed var(--mf-primary-line) !important;
            border-radius: 1rem !important;
            padding: 1rem !important;
        }}
        .stFileUploader section:hover {{
            border-color: var(--mf-primary) !important;
            background: var(--mf-control-alt) !important;
        }}
        [data-testid="stAlert"] {{
            border-radius: 1rem !important;
            border: 1px solid var(--mf-border-soft) !important;
            box-shadow: 0 12px 30px rgba(2, 8, 23, 0.14);
        }}

        /* Enterprise table polish for st.dataframe containers. */
        [data-testid="stDataFrame"] {{
            border: 1px solid var(--mf-border-soft);
            border-radius: var(--mf-radius);
            overflow: hidden;
            box-shadow: 0 16px 36px rgba(2, 8, 23, 0.18);
            background: var(--mf-control-bg) !important;
            --gdg-bg-header: var(--mf-table-header-bg);
            --gdg-bg-header-hovered: var(--mf-table-header-bg);
            --gdg-bg-header-has-focus: var(--mf-table-header-bg);
            --gdg-text-header: var(--mf-table-header-text);
            --gdg-text-group-header: var(--mf-table-header-text);
            --gdg-bg-cell: var(--mf-control-bg);
            --gdg-bg-cell-medium: var(--mf-control-alt);
            --gdg-text-dark: var(--mf-text);
            --gdg-text-medium: var(--mf-muted);
            --gdg-border-color: var(--mf-border-soft);
        }}
        [data-testid="stDataFrame"] div[role="grid"] {{
            background: var(--mf-control-bg) !important;
        }}
        [data-testid="stDataFrame"] [role="gridcell"],
        [data-testid="stDataFrame"] [role="rowheader"] {{
            background: var(--mf-control-bg) !important;
            color: var(--mf-text) !important;
            border-color: var(--mf-border-soft) !important;
        }}
        [data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] {{
            background: color-mix(in srgb, var(--mf-control-alt) 68%, var(--mf-control-bg)) !important;
        }}
        [data-testid="stDataFrame"] [role="columnheader"] {{
            background: var(--mf-table-header-bg) !important;
            background-color: var(--mf-table-header-bg) !important;
            color: var(--mf-table-header-text) !important;
            border-right: 1px solid var(--mf-table-header-border) !important;
            border-bottom: 1px solid var(--mf-table-header-border) !important;
            font-weight: 900 !important;
            letter-spacing: 0 !important;
        }}
        [data-testid="stDataFrame"] [role="columnheader"] p,
        [data-testid="stDataFrame"] [role="columnheader"] span,
        [data-testid="stDataFrame"] [role="columnheader"] div {{
            color: var(--mf-table-header-text) !important;
            font-weight: 900 !important;
        }}
        [data-testid="stDataFrame"] [role="row"]:hover {{
            background: var(--mf-primary-soft) !important;
        }}
        [data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"] {{
            background: var(--mf-primary-soft) !important;
        }}
        div[role="dialog"] {{
            width: min(96vw, 1500px) !important;
            max-width: 96vw !important;
            background:
                linear-gradient(135deg, var(--mf-primary-soft), transparent 34%),
                var(--mf-panel) !important;
            border: 1px solid var(--mf-border-soft) !important;
            color: var(--mf-text) !important;
        }}
        div[role="dialog"] *,
        div[role="dialog"] h1,
        div[role="dialog"] h2,
        div[role="dialog"] h3,
        div[role="dialog"] p,
        div[role="dialog"] span {{
            color: var(--mf-text) !important;
        }}
        div[role="dialog"] [data-testid="stDataFrame"] {{
            max-height: 78vh;
        }}

        @media (max-width: 1100px) {{
            .mf-header {{ align-items: flex-start; }}
            .mf-header-actions {{ justify-content: flex-start; }}
            .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
            .mf-kpi-grid {{ grid-template-columns: 1fr; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def style_plotly_chart(
    fig: go.Figure,
    height: int,
    xaxis_title: str | None = None,
    yaxis_title: str | None = None,
    xaxis_tickangle: int | None = None,
    showlegend: bool | None = None,
    margin: dict[str, int] | None = None,
) -> go.Figure:
    chart_theme = st.session_state.get("chart_theme", {})
    light = st.session_state.get("theme_mode") == "Light"
    font_color = "#0F172A" if light else chart_theme.get("font", "#CBD5E1")
    title_color = "#0F172A" if light else chart_theme.get("title", "#F8FAFC")
    grid = "rgba(100,116,139,0.34)" if light else chart_theme.get("grid", "rgba(51,65,85,0.45)")
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(248,250,252,0.78)" if light else chart_theme.get("plot_bg", "rgba(15,23,42,0.35)"),
        font={"color": font_color, "family": "Inter, Segoe UI, sans-serif"},
        title={"font": {"color": title_color, "size": 18}},
        legend={"font": {"color": font_color}, "title": {"font": {"color": font_color}}},
        margin=margin or {"l": 36, "r": 24, "t": 56, "b": 58},
    )
    if showlegend is not None:
        fig.update_layout(showlegend=showlegend)
    fig.update_xaxes(
        title={"text": xaxis_title, "font": {"color": title_color}},
        tickfont={"color": font_color},
        gridcolor="rgba(0,0,0,0)",
        zerolinecolor=grid,
        linecolor=grid,
    )
    fig.update_yaxes(
        title={"text": yaxis_title, "font": {"color": title_color}},
        tickfont={"color": font_color},
        gridcolor=grid,
        zerolinecolor=grid,
        linecolor=grid,
    )
    if xaxis_tickangle is not None:
        fig.update_xaxes(tickangle=xaxis_tickangle)
    fig.update_traces(textfont={"color": title_color})
    return fig


def dataframe_download(df: pd.DataFrame, label: str, filename: str) -> None:
    st.download_button(
        label,
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=filename,
        mime="text/csv",
        use_container_width=True,
    )


@st.dialog("Full table details", width="large")
def render_table_dialog(title: str, df: pd.DataFrame, filename: str) -> None:
    st.markdown(
        f"""
        <div class="mf-section">
            <div class="mf-eyebrow">Full table details</div>
            <div class="mf-title" style="font-size:1.35rem;">{escape(title)}</div>
            <div class="mf-subtitle">{len(df):,} row(s), {len(df.columns):,} column(s)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.dataframe(_dataframe_theme_style(df), use_container_width=True, hide_index=True, height=680)
    dataframe_download(df, "Download full table CSV", filename)


def _dataframe_theme_style(df: pd.DataFrame, styled_df: Any | None = None) -> Any:
    styler = styled_df if styled_df is not None else df.style
    dark = st.session_state.get("theme_mode", "Dark") == "Dark"
    if dark:
        return (
            styler.set_properties(
                **{
                    "background-color": "#0F172A",
                    "color": "#E2E8F0",
                    "border-color": "#334155",
                }
            )
            .set_table_styles(
                [
                    {
                        "selector": "th",
                        "props": [
                            ("background-color", "#1E293B"),
                            ("color", "#F8FAFC"),
                            ("font-weight", "900"),
                            ("border-color", "#334155"),
                        ],
                    },
                    {
                        "selector": "tbody tr:nth-child(even) td",
                        "props": [("background-color", "#111827")],
                    },
                    {
                        "selector": "tbody tr:hover td",
                        "props": [("background-color", "#1D4ED8"), ("color", "#FFFFFF")],
                    },
                ]
            )
        )
    return (
        styler.set_properties(
            **{
                "background-color": "#FFFFFF",
                "color": "#0F172A",
                "border-color": "#CBD5E1",
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#DBEAFE"),
                        ("color", "#0F172A"),
                        ("font-weight", "900"),
                        ("border-color", "#CBD5E1"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(even) td",
                    "props": [("background-color", "#F8FAFC")],
                },
                {
                    "selector": "tbody tr:hover td",
                    "props": [("background-color", "#EFF6FF")],
                },
            ]
        )
    )


def render_dataframe(
    df: pd.DataFrame,
    key: str,
    title: str,
    filename: str,
    styled_df: Any | None = None,
    compact_height: int = 340,
) -> None:
    if st.button("Maximize details", key=f"{key}_open_table"):
        render_table_dialog(title, df, filename)
    st.dataframe(
        _dataframe_theme_style(df, styled_df),
        use_container_width=True,
        hide_index=True,
        height=compact_height,
    )


def _fmt_number(value: Any) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value or "0")


def _latest_report_status() -> tuple[str, str]:
    current = st.session_state.get("new_report")
    if current:
        stamp = current.report_timestamp.isoformat(sep=" ", timespec="minutes") if current.report_timestamp else "No filename timestamp"
        return current.filename, stamp
    reports = database.list_reports()
    if reports:
        row = reports[0]
        return row["filename"], row.get("report_timestamp") or row.get("ingested_at") or "Stored locally"
    return "No report loaded", "Upload or select reports to begin"


def render_app_header() -> None:
    report_name, report_status = _latest_report_status()
    ai_enabled = st.session_state.get("ai_enabled")
    ai_status = "AI Enabled" if ai_enabled else "AI Off"
    report_display = report_name if len(report_name) <= 48 else f"{report_name[:45]}..."
    st.markdown(
        f"""
        <div class="mf-header">
            <div class="mf-hero-visual" aria-hidden="true">
                <div class="mf-mini-line"></div>
                <div class="mf-mini-chart">
                    <div class="mf-mini-bar" style="height: 32%;"></div>
                    <div class="mf-mini-bar" style="height: 48%;"></div>
                    <div class="mf-mini-bar" style="height: 38%;"></div>
                    <div class="mf-mini-bar" style="height: 62%;"></div>
                    <div class="mf-mini-bar" style="height: 54%;"></div>
                    <div class="mf-mini-bar" style="height: 76%;"></div>
                    <div class="mf-mini-bar" style="height: 92%;"></div>
                </div>
                <div class="mf-hero-stat">
                    <div class="mf-donut"></div>
                    <div>
                        <div class="mf-hero-kicker">Live Comparison</div>
                        <div class="mf-report-name">Impact, risk, and progress in one view</div>
                    </div>
                </div>
            </div>
            <div class="mf-hero-copy">
                <div class="mf-eyebrow">Real insights. Better decisions.</div>
                <div class="mf-title">Migration Fallout<br><span class="mf-gradient-word">Intelligence Dashboard</span></div>
                <div class="mf-subtitle">
                    Get a 360 degree view of migration fallout, track impact, uncover trends, and turn data into confident business decisions.
                </div>
                <div class="mf-hero-features">
                    <span class="mf-feature-chip"><span class="mf-feature-icon">C</span>Compare Reports</span>
                    <span class="mf-feature-chip"><span class="mf-feature-icon">T</span>Track Impact</span>
                    <span class="mf-feature-chip"><span class="mf-feature-icon">R</span>Detect Issues</span>
                    <span class="mf-feature-chip"><span class="mf-feature-icon">I</span>Drive Insights</span>
                </div>
            </div>
            <div class="mf-header-actions">
                <div class="mf-report-pill">
                    <div class="mf-report-label">Current Report</div>
                    <div class="mf-report-name" title="{escape(report_name)}">{escape(report_display)}</div>
                </div>
                <div class="mf-status-row">
                    <span class="mf-chip">{escape(report_status)}</span>
                    <span class="mf-badge {'success' if ai_enabled else 'warning'}">{ai_status}</span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_actions() -> None:
    st.markdown('<div class="mf-action-toolbar">', unsafe_allow_html=True)
    cols = st.columns([1.7, 2.0, 1.1, 3.2])
    with cols[0]:
        if st.button(
            "Refresh",
            key="dashboard_refresh",
            use_container_width=False,
            help="Reload the dashboard and refresh current data from session/database.",
        ):
            st.rerun()
    with cols[1]:
        if st.button(
            "Refresh DB",
            key="dashboard_repair_db",
            use_container_width=False,
            help="Initialize or repair the local SQLite database used by stored reports.",
        ):
            database.init_db()
            st.success("Database is ready.")
    with cols[2]:
        st.download_button(
            "Help",
            data="Use Compare Reports to upload files, then open AI Insights to generate a summary.",
            file_name="migration_dashboard_quick_help.txt",
            mime="text/plain",
            use_container_width=False,
            key="dashboard_export_help",
            help="Download a short guide for using this dashboard.",
        )
    st.markdown("</div>", unsafe_allow_html=True)


def render_sidebar_brand() -> None:
    st.markdown(
        """
        <div class="mf-brand">
            <div class="mf-logo">MF</div>
            <div>
                <div class="mf-brand-title">Migration Fallout</div>
                <div class="mf-brand-subtitle">Report Intelligence</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_card(label: str, value: Any, icon: str, foot: str = "", badge: str = "", badge_type: str = "info") -> None:
    st.markdown(
        f"""
        <div class="mf-card">
            <div class="mf-card-top">
                <div>
                    <div class="mf-card-label">{escape(label)}</div>
                    {f'<span class="mf-badge {badge_type}">{escape(badge)}</span>' if badge else ''}
                </div>
                <div class="mf-card-icon">{escape(icon)}</div>
            </div>
            <div class="mf-card-value">{escape(_fmt_number(value))}</div>
            <div class="mf-card-foot">{escape(foot)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def load_uploaded_report(uploaded, source: str = "manual", overwrite_same_filename: bool = False) -> ParsedReport | None:
    if uploaded is None:
        return None
    try:
        report = parse_excel_report(uploaded.getvalue(), uploaded.name)
        existing = database.find_report_by_filename(report.filename)
        if existing and not overwrite_same_filename:
            st.warning(
                f"Report '{report.filename}' already exists in storage. Enable overwrite and run again if you want to replace it."
            )
            return None
        database.save_report(report, source=source, overwrite_same_filename=overwrite_same_filename)
        for warning in report.warnings or []:
            st.warning(f"{uploaded.name}: {warning}")
        return report
    except Exception as exc:
        st.error(f"Could not parse {uploaded.name}: {exc}")
        return None


def report_options() -> dict[str, int]:
    reports = database.list_reports()
    return {
        f"#{row['id']} - {row['filename']} ({row.get('report_timestamp') or row.get('ingested_at')})": row["id"]
        for row in reports
    }


def _normalized_text(value: object) -> str:
    return str(value or "").strip().lower().replace(" ", "_")


def _find_sheet(report: ParsedReport, *required_parts: str) -> pd.DataFrame | None:
    required = [_normalized_text(part) for part in required_parts]
    for sheet_name, df in report.sheets.items():
        normalized_name = _normalized_text(sheet_name)
        if all(part in normalized_name for part in required):
            return df
    return None


def _format_percent_value(value: object) -> str:
    if value is None or str(value).strip() == "":
        return "N/A"
    if isinstance(value, str) and "%" in value:
        return value.strip()
    try:
        number = float(str(value).strip().replace("%", ""))
    except (TypeError, ValueError):
        return str(value)
    if 0 <= number <= 1:
        number *= 100
    return f"{number:.1f}%"


def _customer_success_rate(report: ParsedReport) -> str:
    df = _find_sheet(report, "sdb", "summary")
    if df is None or df.empty:
        return "N/A"
    quality_columns = [column for column in df.columns if "quality" in _normalized_text(column)]
    if not quality_columns:
        return "N/A"
    target = "sdb_customer_b2b"
    table_mask = df.astype(str).apply(
        lambda col: col.str.strip().str.lower().str.replace(" ", "_", regex=False).str.contains(target, na=False)
    ).any(axis=1)
    if not table_mask.any():
        return "N/A"
    value = df.loc[table_mask, quality_columns[0]].iloc[0]
    return _format_percent_value(value)


def _major_validation_count(report: ParsedReport) -> int:
    severity_lookup = _validation_severity_lookup(report)
    return sum(1 for severity in severity_lookup.values() if severity == "major")


def _validation_severity_lookup(*reports: ParsedReport | None) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for report in reports:
        if report is None:
            continue
        df = _find_sheet(report, "summary", "rules")
        if df is None or df.empty:
            continue
        severity_columns = []
        for column in df.columns:
            values = df[column].dropna().astype(str).str.strip().str.lower()
            if values.isin(["major", "minor"]).any():
                severity_columns.append(column)
        if not severity_columns:
            continue
        severity_column = severity_columns[0]
        validation_candidates = [
            column
            for column in ("validation_name", "rule_name", "major_validation")
            if column in df.columns and column != severity_column
        ]
        if not validation_candidates:
            validation_candidates = [
                column
                for column in df.columns
                if column != severity_column
                and ("validation" in _normalized_text(column) or "rule" in _normalized_text(column))
            ]
        if not validation_candidates:
            validation_candidates = [column for column in df.columns if column != severity_column]
        if not validation_candidates:
            continue
        validation_column = validation_candidates[0]
        for _, row in df[[validation_column, severity_column]].dropna(how="all").iterrows():
            validation_key = str(row.get(validation_column, "") or "").strip().lower()
            severity = str(row.get(severity_column, "") or "").strip().lower()
            if validation_key and validation_key != "nan" and severity in {"major", "minor"}:
                lookup[validation_key] = severity
    return lookup


def _with_validation_severity(df: pd.DataFrame, severity_lookup: dict[str, str]) -> pd.DataFrame:
    work = df.copy()
    if "validation_key" not in work.columns:
        return work
    work["validation_severity"] = work["validation_key"].fillna("").astype(str).str.lower().str.strip().map(
        lambda value: severity_lookup.get(value, "minor")
    )
    return work


def _usable_fallout_rows(df: pd.DataFrame, count_column: str = "fallout_count_new") -> pd.DataFrame:
    if df.empty:
        return df
    work = df.copy()
    if "validation_key" in work.columns:
        work = work[~work["validation_key"].fillna("").astype(str).str.lower().str.strip().isin({"", "unknown validation", "nan"})]
    if "validation_name" in work.columns:
        work = work[~work["validation_name"].fillna("").astype(str).str.lower().str.strip().isin({"", "unknown validation", "nan"})]
    if count_column in work.columns:
        work[count_column] = pd.to_numeric(work[count_column], errors="coerce").fillna(0)
        work = work[work[count_column] > 0]
    return work


def _add_fallout_segments(results: dict[str, Any], severity_lookup: dict[str, str]) -> dict[str, Any]:
    sort_column = "fallout_count_new"
    for key in (
        "comparison",
        "current_validations",
        "new_validations",
        "resolved_validations",
        "top_fallout",
        "top_customers",
        "increased_fallout",
        "reduced_fallout",
        "increased_customers",
        "reduced_customers",
    ):
        if isinstance(results.get(key), pd.DataFrame):
            results[key] = _with_validation_severity(results[key], severity_lookup)
    source = results.get("comparison")
    if not isinstance(source, pd.DataFrame):
        source = results.get("current_validations")
    if isinstance(source, pd.DataFrame) and not source.empty:
        source = _usable_fallout_rows(_with_validation_severity(source, severity_lookup), sort_column)
        major = source[source["validation_severity"] == "major"].sort_values(sort_column, ascending=False)
        minor = source[source["validation_severity"] == "minor"].sort_values(sort_column, ascending=False)
        results["top_major_fallout"] = major.head(10)
        results["top_minor_fallout"] = minor.head(10)
        results["top_fallout"] = results["top_major_fallout"] if not results["top_major_fallout"].empty else source.sort_values(sort_column, ascending=False).head(10)
    else:
        empty = pd.DataFrame()
        results["top_major_fallout"] = empty
        results["top_minor_fallout"] = empty
    return results


def build_current_report_results(report: ParsedReport) -> dict[str, Any]:
    aggregated = aggregate_report(report)
    severity_lookup = _validation_severity_lookup(report)
    if aggregated.empty:
        table_summary = pd.DataFrame(
            columns=["table_name", "fallout_count_new", "customer_count_new", "validations"]
        )
    else:
        table_summary = (
            aggregated.groupby("table_name", dropna=False)
            .agg(
                fallout_count_new=("fallout_count", "sum"),
                customer_count_new=("customer_count", "sum"),
                validations=("validation_key", "nunique"),
            )
            .reset_index()
            .sort_values("fallout_count_new", ascending=False)
        )
    current = aggregated.rename(
        columns={
            "fallout_count": "fallout_count_new",
            "customer_count": "customer_count_new",
        }
    )
    current = _with_validation_severity(current, severity_lookup)
    totals = {
        "total_fallout_new": int(current["fallout_count_new"].sum()) if not current.empty else 0,
        "total_customer_new": int(current["customer_count_new"].sum()) if not current.empty else 0,
        "total_validations": int(current["validation_key"].nunique()) if not current.empty else 0,
        "total_tables": int(current["table_name"].nunique()) if not current.empty else 0,
        "success_rate": _customer_success_rate(report),
        "major_validations": _major_validation_count(report),
    }
    results = {
        "mode": "current",
        "report_name": report.filename,
        "totals": totals,
        "current_validations": current.sort_values("fallout_count_new", ascending=False),
        "top_fallout": current.sort_values("fallout_count_new", ascending=False).head(10),
        "top_customers": current.sort_values("customer_count_new", ascending=False).head(10),
        "table_summary": table_summary,
    }
    return _add_fallout_segments(results, severity_lookup)


def get_current_results() -> dict[str, Any] | None:
    return st.session_state.get("comparison_results")


def get_active_analysis_results() -> dict[str, Any] | None:
    return st.session_state.get("comparison_results") or st.session_state.get("current_report_results")


def _session_key(provider: str) -> str:
    return f"session_api_key_{provider}"


def get_active_api_key(provider: str) -> str:
    if st.session_state.get("auto_load_api_key", True):
        try:
            saved_key = get_api_key(provider)
            if saved_key:
                st.session_state[_session_key(provider)] = saved_key
                return saved_key
        except CredentialStoreError as exc:
            st.session_state["keyring_warning"] = str(exc)
    return st.session_state.get(_session_key(provider), "")


def _wrap_text(text: str, width: int = 92) -> list[str]:
    lines: list[str] = []
    for raw_line in str(text or "").splitlines() or [""]:
        words = raw_line.split()
        if not words:
            lines.append("")
            continue
        current = ""
        for word in words:
            candidate = f"{current} {word}".strip()
            if len(candidate) > width and current:
                lines.append(current)
                current = word
            else:
                current = candidate
        lines.append(current)
    return lines


def _summary_doc_bytes(summary: str) -> bytes:
    body = escape(summary).replace("\n", "<br>")
    html = f"""<!doctype html>
<html>
<head><meta charset="utf-8"><title>Executive Summary</title></head>
<body style="font-family: Calibri, Arial, sans-serif; color:#111827;">
<h1>Migration Fallout Executive Summary</h1>
<div style="font-size: 12pt; line-height: 1.45;">{body}</div>
</body>
</html>"""
    return html.encode("utf-8")


def _pdf_escape(text: str) -> str:
    return str(text).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _summary_pdf_bytes(summary: str) -> bytes:
    lines = ["Migration Fallout Executive Summary", ""] + _wrap_text(summary, 88)
    stream_lines = ["BT", "/F1 12 Tf", "50 790 Td", "14 TL"]
    for index, line in enumerate(lines[:52]):
        prefix = "" if index == 0 else "T* "
        stream_lines.append(f"{prefix}({_pdf_escape(line)}) Tj")
    stream_lines.append("ET")
    content = "\n".join(stream_lines).encode("latin-1", "replace")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(content)).encode("ascii") + b" >>\nstream\n" + content + b"\nendstream",
    ]
    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for number, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{number} 0 obj\n".encode("ascii"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")
    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        f"trailer << /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF".encode("ascii")
    )
    return bytes(pdf)


def render_summary_actions(summary: str) -> None:
    st.text_area("Executive summary text", value=summary, height=220, key="ai_summary_copy_text")
    payload = json.dumps(summary)
    components.html(
        f"""
        <button id="copy-summary" style="
            background:#2563EB;color:white;border:0;border-radius:10px;
            padding:10px 16px;font:600 14px Arial;cursor:pointer;">
            Copy executive summary
        </button>
        <span id="copy-status" style="margin-left:10px;color:#94A3B8;font:14px Arial;"></span>
        <script>
        const button = document.getElementById('copy-summary');
        const status = document.getElementById('copy-status');
        button.addEventListener('click', async () => {{
            try {{
                await navigator.clipboard.writeText({payload});
                status.textContent = 'Copied';
            }} catch (err) {{
                status.textContent = 'Select the text box above and press Ctrl+C';
            }}
        }});
        </script>
        """,
        height=52,
    )
    c1, c2, c3 = st.columns(3)
    c1.download_button(
        "Download TXT",
        data=summary,
        file_name="executive_summary.txt",
        mime="text/plain",
        use_container_width=True,
    )
    c2.download_button(
        "Download Word",
        data=_summary_doc_bytes(summary),
        file_name="executive_summary.doc",
        mime="application/msword",
        use_container_width=True,
    )
    c3.download_button(
        "Download PDF",
        data=_summary_pdf_bytes(summary),
        file_name="executive_summary.pdf",
        mime="application/pdf",
        use_container_width=True,
    )


def _analysis_frame(results: dict[str, Any]) -> pd.DataFrame:
    source = results.get("comparison")
    if not isinstance(source, pd.DataFrame):
        source = results.get("current_validations")
    if not isinstance(source, pd.DataFrame) or source.empty:
        return pd.DataFrame()
    work = _usable_fallout_rows(source, "fallout_count_new")
    for column in ("fallout_count_new", "customer_count_new"):
        if column not in work.columns:
            work[column] = 0
        work[column] = pd.to_numeric(work[column], errors="coerce").fillna(0)
    work["priority_score"] = (work["fallout_count_new"] * 0.4) + (work["customer_count_new"] * 0.6)
    return work.sort_values("priority_score", ascending=False)


def _ai_section_text(summary: str, aliases: list[str]) -> str:
    pattern = r"(?im)^\s*(?:#{1,4}\s*)?(?:\d+[\.\)]\s*)?(" + "|".join(re.escape(alias) for alias in aliases) + r")\s*:?\s*$"
    matches = list(re.finditer(pattern, summary or ""))
    if not matches:
        return ""
    start = matches[0].end()
    following = [match.start() for match in matches[1:]]
    all_headings = list(re.finditer(r"(?im)^\s*(?:#{1,4}\s*)?(?:\d+[\.\)]\s*)?[A-Z][A-Za-z /&-]{3,}\s*:?\s*$", summary or ""))
    for heading in all_headings:
        if heading.start() > matches[0].start():
            following.append(heading.start())
    end = min(following) if following else len(summary or "")
    return (summary or "")[start:end].strip()


def _badge(label: str, tone: str) -> str:
    return f'<span class="mf-badge {tone}">{escape(label)}</span>'


def _status_badges(row: pd.Series) -> str:
    badges: list[str] = []
    severity = str(row.get("validation_severity", "")).lower()
    if str(row.get("validation_name", "")).strip().lower() == "unknown validation":
        badges.append(_badge("Unknown Validation", ""))
    if severity == "major":
        badges.append(_badge("Critical", "warning"))
    if float(row.get("customer_count_new", 0) or 0) >= 500:
        badges.append(_badge("High Customer Impact", "info"))
    status = str(row.get("status", "")).lower()
    if status in {"increased", "new"}:
        badges.append(_badge("Worsened", "warning"))
    elif status in {"reduced", "resolved"}:
        badges.append(_badge("Improved", "success"))
    return " ".join(badges)


def _top_impacted_table(results: dict[str, Any]) -> str:
    table_summary = results.get("table_summary")
    if not isinstance(table_summary, pd.DataFrame) or table_summary.empty or "fallout_count_new" not in table_summary.columns:
        return "N/A"
    row = table_summary.sort_values("fallout_count_new", ascending=False).iloc[0]
    return str(row.get("table_name", "N/A"))


def _highest_customer_validation(results: dict[str, Any]) -> str:
    df = results.get("top_customers")
    if not isinstance(df, pd.DataFrame) or df.empty:
        return "N/A"
    row = df.sort_values("customer_count_new", ascending=False).iloc[0]
    return str(row.get("validation_name", "N/A"))


def render_ai_kpi_cards(results: dict[str, Any]) -> None:
    totals = results.get("totals", {})
    current_totals = st.session_state.get("current_report_results", {}).get("totals", {})
    success_rate = totals.get("success_rate") or current_totals.get("success_rate") or "N/A"
    cards = [
        ("Total Fallout", totals.get("total_fallout_new", 0), "F", "Current fallout volume", "warning"),
        ("Customers Impacted", totals.get("total_customer_new", 0), "C", "Customer impact", "info"),
        ("Success Rate", success_rate, "%", "Percentage of Valid Customers", "success"),
        ("Top Impacted Table", _top_impacted_table(results), "T", "Highest fallout table", "warning"),
        ("Highest Customer Impact", _highest_customer_validation(results), "H", "Largest customer impact validation", "info"),
    ]
    html = ['<div class="mf-ai-kpi-grid">']
    for label, value, icon, foot, tone in cards:
        html.append(
            '<div class="mf-card mf-ai-kpi-card">'
            '<div class="mf-card-top">'
            '<div>'
            f'<div class="mf-card-label">{escape(label)}</div>'
            f'<span class="mf-badge {tone}">{escape(tone.title())}</span>'
            '</div>'
            f'<div class="mf-card-icon">{escape(icon)}</div>'
            '</div>'
            f'<div class="mf-card-value">{escape(_fmt_number(value)) if isinstance(value, (int, float)) else escape(str(value))}</div>'
            f'<div class="mf-card-foot">{escape(foot)}</div>'
            '</div>'
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def render_ai_charts(results: dict[str, Any]) -> None:
    work = _analysis_frame(results)
    if work.empty:
        st.info("No validation data available for AI insight charts.")
        return
    chart_theme = st.session_state.get("chart_theme", {})
    c1, c2 = st.columns(2)
    with c1:
        fallout = work.sort_values("fallout_count_new", ascending=False).head(10)
        fallout = fallout.copy()
        fallout["rank"] = [f"#{idx}" for idx in range(1, len(fallout) + 1)]
        fig = px.bar(
            fallout,
            x="rank",
            y="fallout_count_new",
            color="validation_severity" if "validation_severity" in fallout.columns else None,
            title="Top 10 validations by fallout count",
            text="fallout_count_new",
            custom_data=["validation_name", "table_name", "customer_count_new", "priority_score"],
            color_discrete_map={"major": "#EF4444", "minor": "#F59E0B"},
        )
        fig.update_traces(
            hovertemplate="<b>%{customdata[0]}</b><br>Table: %{customdata[1]}<br>Fallouts: %{y:,}<br>Customers: %{customdata[2]:,}<br>Priority Score: %{customdata[3]:,.0f}<extra></extra>",
            textposition="outside",
        )
        style_plotly_chart(
            fig,
            height=360,
            xaxis_title="Validation rank",
            yaxis_title="Fallout count",
            showlegend=True,
            margin={"l": 40, "r": 24, "t": 58, "b": 58},
        )
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        customers = work.sort_values("customer_count_new", ascending=False).head(10)
        customers = customers.copy()
        customers["rank"] = [f"#{idx}" for idx in range(1, len(customers) + 1)]
        fig = px.bar(
            customers,
            x="rank",
            y="customer_count_new",
            color="validation_severity" if "validation_severity" in customers.columns else None,
            title="Top 10 validations by impacted customers",
            text="customer_count_new",
            custom_data=["validation_name", "table_name", "fallout_count_new", "priority_score"],
            color_discrete_map={"major": "#EF4444", "minor": "#F59E0B"},
        )
        fig.update_traces(
            hovertemplate="<b>%{customdata[0]}</b><br>Table: %{customdata[1]}<br>Customers: %{y:,}<br>Fallouts: %{customdata[2]:,}<br>Priority Score: %{customdata[3]:,.0f}<extra></extra>",
            textposition="outside",
        )
        style_plotly_chart(
            fig,
            height=360,
            xaxis_title="Validation rank",
            yaxis_title="Impacted customers",
            showlegend=True,
            margin={"l": 40, "r": 24, "t": 58, "b": 58},
        )
        st.plotly_chart(fig, use_container_width=True)
    if results.get("mode") == "comparison" and "fallout_delta" in work.columns:
        delta = work.copy()
        delta["abs_delta"] = pd.to_numeric(delta["fallout_delta"], errors="coerce").fillna(0).abs()
        delta = delta.sort_values("abs_delta", ascending=False).head(10)
        fig = px.bar(
            delta,
            x="validation_name",
            y="fallout_delta",
            color="fallout_delta",
            title="Comparison trend and fallout delta",
            text="fallout_delta",
            color_continuous_scale=["#22C55E", "#94A3B8", "#EF4444"],
        )
        style_plotly_chart(
            fig,
            height=360,
            xaxis_title="Validation",
            yaxis_title="Fallout delta",
            xaxis_tickangle=-30,
            showlegend=False,
            margin={"l": 40, "r": 24, "t": 58, "b": 96},
        )
        st.plotly_chart(fig, use_container_width=True)


def render_ai_visual_cards(summary: str, results: dict[str, Any]) -> None:
    work = _analysis_frame(results)
    major_source = results.get("top_major_fallout")
    if isinstance(major_source, pd.DataFrame) and not major_source.empty:
        top = _analysis_frame({"current_validations": major_source}).head(6)
    else:
        top = work[work.get("validation_severity", "") == "major"].head(6) if "validation_severity" in work.columns else work.head(6)
    max_score = float(top["priority_score"].max()) if not top.empty else 0
    sections = {
        "Executive Summary": _ai_section_text(summary, ["Executive Summary"]) or summary,
        "Critical Risks": _ai_section_text(summary, ["Top Current Risks", "Top Risks", "Critical Risks"]),
        "Suggested Actions": _ai_section_text(summary, ["Suggested Actions", "Recommended Actions"]),
        "Watchlist": _ai_section_text(summary, ["Watchlist", "Watchlist Minor Fallouts"]),
    }
    card_meta = [
        ("Executive Summary", "Overview", "info", "E"),
        ("Critical Risks", f"Priority Score {max_score:,.0f}", "warning", "R"),
        ("Suggested Actions", "Recommended", "success", "A"),
        ("Watchlist", "Monitor", "", "W"),
    ]
    for title, badge, tone, icon in card_meta:
        with st.expander(f"{icon}  {title}  |  {badge}", expanded=title in {"Executive Summary", "Critical Risks"}):
            st.markdown(f'<div class="mf-ai-section-head">{_badge(badge, tone) if tone else _badge(badge, "")}</div>', unsafe_allow_html=True)
            if title == "Critical Risks" and not top.empty:
                for idx, (_, row) in enumerate(top.iterrows()):
                    label = str(row.get("validation_name", "Unknown validation"))
                    score = float(row.get("priority_score", 0) or 0)
                    st.markdown(
                        f"**{escape(label)}**  \n"
                        f"Priority Score: `{score:,.0f}` | Fallouts: `{int(row.get('fallout_count_new', 0)):,}` | Customers: `{int(row.get('customer_count_new', 0)):,}`",
                        unsafe_allow_html=False,
                    )
                    st.markdown(_status_badges(row), unsafe_allow_html=True)
                    st.caption(f"Recommended action: investigate table {row.get('table_name', 'Unknown')} and validate remediation owner.")
                    if idx < len(top) - 1:
                        st.markdown('<div class="mf-ai-risk-separator"></div>', unsafe_allow_html=True)
            text = sections.get(title, "")
            if text:
                st.markdown(text)


def render_ai_insights_dashboard(summary: str, results: dict[str, Any]) -> None:
    render_ai_kpi_cards(results)
    render_ai_charts(results)
    st.subheader("AI Insight Cards")
    render_ai_visual_cards(summary, results)
    with st.expander("Original AI text", expanded=False):
        st.markdown(summary)
        render_summary_actions(summary)


def render_about_page() -> None:
    st.subheader("About this tool")
    st.markdown(
        """
        This dashboard helps teams compare migration fallout reports, understand what changed between two report runs,
        inspect the latest workbook data, scan Outlook for report attachments, and generate executive-ready AI summaries.

        The core comparison logic is deterministic: counts, deltas, statuses, charts, tables, and CSV exports come from
        the uploaded or stored Excel reports. AI features are optional and only use prepared report/comparison context for
        summaries and natural-language answers.
        """
    )

    st.markdown("### Main workflow")
    st.markdown(
        """
        1. Open **Compare Reports** and upload two Excel reports, or select two stored reports.
        2. Review the high-level metrics and migration summary charts in **Dashboard**.
        3. Use **Fallouts** to inspect validations, sheets, new fallout, resolved fallout, and detailed tables.
        4. Use **AI Insights** or **Fallout Insight Engine** if AI is enabled and a provider key is available.
        5. Export data, summaries, or reports where needed.
        """
    )

    tab_details = [
        (
            "Dashboard",
            "The executive overview. It shows the current report context, fallout and customer KPIs, new/resolved validation counts, migration summary charts for previous versus new report, top failed validations, and the full comparison table.",
        ),
        (
            "Compare Reports",
            "The comparison starting point. Upload old/new Excel reports manually or pick stored reports from the local database. The app decides the latest report from filenames when enabled, parses sheets, normalizes common columns, and builds the comparison results.",
        ),
        (
            "Fallouts",
            "The detailed investigation area. It shows comparison slices such as new validations, resolved validations, top fallouts, customer impact, and an explorer for sheets in the latest workbook. The explorer uses sheet buttons plus search so users can inspect raw sheet-level details.",
        ),
        (
            "Outlook Scanner",
            "The email ingestion area. It can scan Outlook for report attachments using keyword/date filters, stop an active scan, and schedule recurring scans while the dashboard is open. This requires desktop Outlook/COM support on the machine.",
        ),
        (
            "Stored Reports",
            "The local report library. It lists reports already saved in the local SQLite database, supports downloading/inspecting stored metadata, and allows cleanup of reports that are no longer needed.",
        ),
        (
            "AI Insights",
            "The executive-summary generator. When AI is enabled in the sidebar and the API connection works, it sends aggregate comparison context to the selected provider and returns a structured executive summary. The summary can be copied or downloaded as TXT, Word, or PDF.",
        ),
        (
            "Fallout Insight Engine",
            "The natural-language assistant for the current/latest report and comparison differences. Users can ask questions like what changed most, which validations are riskiest, or which tables improved. It is gated by the same AI setting and API key as AI Insights.",
        ),
        (
            "Settings",
            "The operational settings page. It shows the database path, attachment folder, saved key status by provider, and database repair/initialization controls.",
        ),
    ]

    st.markdown("### Tabs")
    for title, description in tab_details:
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.write(description)

    st.markdown("### Sidebar configuration")
    st.markdown(
        """
        The sidebar controls the app theme and optional AI setup. API keys can be stored per provider in Windows
        Credential Manager when the `keyring` package is available. The selected AI provider, base URL, model, saved key,
        and test connection button all live in the sidebar so AI setup is not duplicated across pages.
        """
    )


def render_api_key_controls(
    provider: str,
    location: str,
    compact: bool = False,
    show_test: bool = False,
    base_url: str = "",
    model: str = "",
) -> None:
    st.session_state["auto_load_api_key"] = st.checkbox(
        "Load saved API key automatically",
        value=st.session_state.get("auto_load_api_key", True),
        key=f"{location}_autoload_{provider}",
    )

    saved_key: str | None = None
    keyring_available = True
    try:
        saved_key = get_api_key(provider)
    except CredentialStoreError as exc:
        keyring_available = False
        st.warning(f"{exc} Use session-only API key entry on this machine.")

    if keyring_available:
        st.caption(f"Saved Credential Manager key: {mask_api_key(saved_key)}")
        if st.session_state["auto_load_api_key"] and saved_key:
            st.session_state[_session_key(provider)] = saved_key

    entered_key = st.text_input(
        "API key",
        value="",
        type="password",
        placeholder="Enter key to use, save, or update",
        key=f"{location}_api_key_entry_{provider}",
    )
    if entered_key:
        st.session_state[_session_key(provider)] = entered_key

    if compact:
        save_clicked = st.button(
            "Save API key",
            key=f"{location}_save_key_{provider}",
            disabled=not keyring_available,
            use_container_width=True,
        )
        update_clicked = st.button(
            "Update API key",
            key=f"{location}_update_key_{provider}",
            disabled=not keyring_available,
            use_container_width=True,
        )
        delete_clicked = st.button(
            "Delete saved key",
            key=f"{location}_delete_key_{provider}",
            disabled=not keyring_available,
            use_container_width=True,
        )
    else:
        c1, c2, c3 = st.columns(3)
        save_clicked = c1.button("Save API key", key=f"{location}_save_key_{provider}", disabled=not keyring_available)
        update_clicked = c2.button("Update API key", key=f"{location}_update_key_{provider}", disabled=not keyring_available)
        delete_clicked = c3.button("Delete saved key", key=f"{location}_delete_key_{provider}", disabled=not keyring_available)

    if save_clicked:
        try:
            save_api_key(provider, st.session_state.get(_session_key(provider), ""))
            st.success("API key saved to Windows Credential Manager.")
        except CredentialStoreError as exc:
            st.warning(f"{exc} The key was kept only for this Streamlit session.")
    if update_clicked:
        try:
            save_api_key(provider, st.session_state.get(_session_key(provider), ""))
            st.success("Saved API key updated.")
        except CredentialStoreError as exc:
            st.warning(f"{exc} The key was kept only for this Streamlit session.")
    if delete_clicked:
        try:
            delete_api_key(provider)
            st.session_state.pop(_session_key(provider), None)
            st.success("Saved API key removed from Windows Credential Manager.")
        except CredentialStoreError as exc:
            st.warning(str(exc))

    if show_test:
        if st.button("Test API connection", key=f"{location}_test_connection_{provider}", use_container_width=True):
            api_key = get_active_api_key(provider)
            ok, message = test_ai_connection(provider, api_key, base_url, model)
            if ok:
                st.success(message)
            else:
                st.warning(message)


def render_kpis(results: dict[str, Any]) -> None:
    totals = results["totals"]
    card_groups = [
        [
            ("Fallouts Old", totals["total_fallout_old"], "F", "Previous baseline", "", "info"),
            (
                "Fallouts New",
                totals["total_fallout_new"],
                "N",
                "Current report",
                f"Delta {totals['fallout_delta']:+,}",
                "success" if totals["fallout_delta"] <= 0 else "warning",
            ),
        ],
        [
            ("Customers Old", totals["total_customer_old"], "C", "Previous impact", "", "info"),
            (
                "Customers New",
                totals["total_customer_new"],
                "I",
                "Current impact",
                f"Delta {totals['customer_delta']:+,}",
                "success" if totals["customer_delta"] <= 0 else "warning",
            ),
        ],
        [
            ("New Validations", totals["new_validations"], "+", "Detected in current", "Watch", "warning"),
            ("Resolved", totals["resolved_validations"], "OK", "No longer active", "Good", "success"),
        ],
    ]
    def card_markup(label: str, value: Any, icon: str, foot: str, badge: str, badge_type: str) -> str:
        badge_html = f'<span class="mf-badge {badge_type}">{escape(badge)}</span>' if badge else ""
        return (
            '<div class="mf-card">'
            '<div class="mf-card-top">'
            '<div>'
            f'<div class="mf-card-label">{escape(label)}</div>'
            f"{badge_html}"
            "</div>"
            f'<div class="mf-card-icon">{escape(icon)}</div>'
            "</div>"
            f'<div class="mf-card-value">{escape(_fmt_number(value))}</div>'
            f'<div class="mf-card-foot">{escape(foot)}</div>'
            "</div>"
        )
    group_html = []
    for group in card_groups:
        group_html.append('<div class="mf-kpi-pair">' + "".join(card_markup(*card) for card in group) + "</div>")
    st.markdown(f'<div class="mf-kpi-grid">{"".join(group_html)}</div>', unsafe_allow_html=True)


def render_current_report_kpis(results: dict[str, Any]) -> None:
    totals = results["totals"]
    cards = [
        ("Fallouts", totals["total_fallout_new"], "F", "Current report fallout count", "", "info"),
        ("Customers", totals["total_customer_new"], "C", "Current impacted customers", "", "info"),
        ("Validations", totals["total_validations"], "V", "Distinct validations in report", "", "info"),
        ("Tables", totals["total_tables"], "T", "Tables with fallout data", "", "info"),
        ("Success Rate", totals["success_rate"], "%", "Percentage of Valid Customers", "", "success"),
        ("Total Major Validations", totals["major_validations"], "M", "Distinct major validations", "", "info"),
    ]

    def card_markup(label: str, value: Any, icon: str, foot: str, badge: str, badge_type: str) -> str:
        badge_html = f'<span class="mf-badge {badge_type}">{escape(badge)}</span>' if badge else ""
        return (
            '<div class="mf-card">'
            '<div class="mf-card-top">'
            '<div>'
            f'<div class="mf-card-label">{escape(label)}</div>'
            f"{badge_html}"
            "</div>"
            f'<div class="mf-card-icon">{escape(icon)}</div>'
            "</div>"
            f'<div class="mf-card-value">{escape(_fmt_number(value))}</div>'
            f'<div class="mf-card-foot">{escape(foot)}</div>'
            "</div>"
        )

    groups = [cards[0:2], cards[2:4], cards[4:6]]
    group_html = ['<div class="mf-kpi-pair">' + "".join(card_markup(*card) for card in group) + "</div>" for group in groups]
    st.markdown(f'<div class="mf-kpi-grid">{"".join(group_html)}</div>', unsafe_allow_html=True)


def render_chart(df: pd.DataFrame, x: str, y: str, title: str) -> None:
    if df.empty:
        st.info("No data available for this chart.")
        return
    chart_theme = st.session_state.get("chart_theme", {})
    fig = px.bar(
        df,
        x=x,
        y=y,
        color=y,
        title=title,
        text=y,
        color_continuous_scale=chart_theme.get("primary_scale", ["#38BDF8", "#3B82F6"]),
    )
    style_plotly_chart(
        fig,
        height=420,
        xaxis_title=x.replace("_", " ").title(),
        yaxis_title=y.replace("_", " ").title(),
        xaxis_tickangle=-35,
        showlegend=False,
        margin={"l": 40, "r": 24, "t": 56, "b": 90},
    )
    st.plotly_chart(fig, use_container_width=True)


def _migration_summary_frame(report: ParsedReport | None) -> pd.DataFrame:
    if report is None:
        return pd.DataFrame()
    sheet = next(
        (
            df
            for name, df in report.sheets.items()
            if "migration summary" in name.strip().lower().replace("_", " ")
        ),
        None,
    )
    if sheet is None or sheet.empty or len(sheet.columns) < 4:
        return pd.DataFrame()
    label_col = sheet.columns[0]
    clean_columns = {
        column: str(column).strip().lower().replace("_", " ").replace("%", "").replace(",", " ").strip()
        for column in sheet.columns
    }
    valid_col = next(
        (column for column, clean in clean_columns.items() if clean in {"valid", "valid pct", "valid percent"}),
        sheet.columns[3],
    )
    failed_col = next(
        (column for column, clean in clean_columns.items() if clean in {"failed", "failed pct", "failed percent"}),
        sheet.columns[5] if len(sheet.columns) > 5 else None,
    )
    selected_columns = [label_col, valid_col] + ([failed_col] if failed_col is not None else [])
    work = sheet[selected_columns].copy()
    work.columns = ["label", "valid_pct"] + (["failed_pct"] if failed_col is not None else [])
    work["label"] = work["label"].fillna("").astype(str).str.strip()
    for pct_col in ("valid_pct", "failed_pct"):
        if pct_col not in work:
            continue
        work[pct_col] = pd.to_numeric(work[pct_col], errors="coerce").fillna(0)
    work["label_key"] = work["label"].str.lower().str.replace(r"\s+", " ", regex=True)
    work = work[work["label_key"].isin(MIGRATION_SUMMARY_LABELS)]
    if work.empty:
        return pd.DataFrame()
    for pct_col in ("valid_pct", "failed_pct"):
        if pct_col not in work:
            continue
        if work[pct_col].max() <= 1:
            work[pct_col] = work[pct_col] * 100
        elif work[pct_col].max() > 100:
            work[pct_col] = work[pct_col] / 100
        work[pct_col] = work[pct_col].clip(lower=0, upper=100)
    if "failed_pct" not in work:
        work["failed_pct"] = (100 - work["valid_pct"]).clip(lower=0, upper=100)
    work = work[(work["label"] != "") & ((work["valid_pct"] > 0) | (work["failed_pct"] > 0))]
    return work[["label", "valid_pct", "failed_pct"]]


def render_migration_summary_pies(old_report: ParsedReport | None, new_report: ParsedReport | None) -> None:
    chart_theme = st.session_state.get("chart_theme", {})
    previous = _migration_summary_frame(old_report)
    current = _migration_summary_frame(new_report)
    if previous.empty and current.empty:
        st.info("Migration Summary sheet was not found, or it does not have usable values in column D.")
        return
    metric = st.radio("Migration summary metric", ["Valid %", "Failed %"], horizontal=True, key="migration_summary_metric")
    value_col = "valid_pct" if metric == "Valid %" else "failed_pct"
    hover_label = "Valid" if metric == "Valid %" else "Failed"
    colors = (
        ["#16A34A", "#22C55E", "#38BDF8", "#2563EB", "#8B5CF6", "#14B8A6"]
        if metric == "Valid %"
        else ["#DC2626", "#D97706", "#F59E0B", "#EF4444", "#FB7185", "#F97316"]
    )
    previous_visible = previous[previous[value_col] > 0].copy() if not previous.empty else pd.DataFrame()
    current_visible = current[current[value_col] > 0].copy() if not current.empty else pd.DataFrame()
    if previous_visible.empty and current_visible.empty:
        st.info(f"No usable {metric} values found in Migration Summary.")
        return

    if previous_visible.empty and not current_visible.empty:
        labels = current_visible["label"].tolist()
        values = current_visible[value_col].tolist()
        fig = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    text=[f"{value:.1f}%" for value in values],
                    textinfo="text",
                    textposition="inside",
                    hole=0.45,
                    name="Current Migration Summary",
                    marker={"colors": colors, "line": {"color": chart_theme.get("plot_bg", "#0F172A"), "width": 2}},
                    hovertemplate=f"<b>%{{label}}</b><br>{hover_label}: %{{value:.2f}}%<extra>Current Migration Summary</extra>",
                    sort=False,
                )
            ]
        )
        fig.update_layout(
            height=430,
            title={"text": "Current Migration Summary", "x": 0.5, "font": {"size": 16, "color": chart_theme.get("title", "#F8FAFC")}},
            showlegend=False,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font={"color": chart_theme.get("font", "#CBD5E1"), "family": "Inter, Segoe UI, sans-serif"},
            margin={"l": 22, "r": 22, "t": 72, "b": 18},
        )
        st.plotly_chart(fig, use_container_width=True)
        return

    fig = make_subplots(
        rows=1,
        cols=2,
        specs=[[{"type": "domain"}, {"type": "domain"}]],
        subplot_titles=("Previous Migration Summary", "New Migration Summary"),
        horizontal_spacing=0.22,
    )
    for index, (title, df, col) in enumerate(
        (
            ("Previous Migration Summary", previous_visible, 1),
            ("New Migration Summary", current_visible, 2),
        )
    ):
        if df.empty:
            labels: list[str] = ["No data"]
            values: list[float] = [1]
            text: list[str] = [""]
        else:
            labels = df["label"].tolist()
            values = df[value_col].tolist()
            text = [f"{value:.1f}%" for value in values]
        fig.add_trace(
            go.Pie(
                labels=labels,
                values=values,
                text=text,
                textinfo="text",
                textposition="inside",
                hole=0.45,
                name=title,
                marker={"colors": colors, "line": {"color": chart_theme.get("plot_bg", "#0F172A"), "width": 2}},
                hovertemplate=f"<b>%{{label}}</b><br>{hover_label}: %{{value:.2f}}%<extra>{title}</extra>",
                sort=False,
            ),
            row=1,
            col=col,
        )
    fig.update_layout(
        height=430,
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"color": chart_theme.get("font", "#CBD5E1"), "family": "Inter, Segoe UI, sans-serif"},
        margin={"l": 22, "r": 22, "t": 72, "b": 18},
    )
    fig.update_annotations(font={"size": 15, "color": chart_theme.get("title", "#F8FAFC")}, y=0.98)
    chart_id = "migration-summary-linked-pies"
    html = fig.to_html(
        full_html=False,
        include_plotlyjs=True,
        div_id=chart_id,
        config={"displayModeBar": True, "responsive": True},
        post_script=f"""
        const chart = document.getElementById("{chart_id}");
        function pullsForTrace(trace, label) {{
            return (trace.labels || []).map((item) => item === label ? 0.12 : 0);
        }}
        chart.on("plotly_hover", function(event) {{
            const label = event.points && event.points[0] ? event.points[0].label : null;
            if (!label || label === "No data") return;
            const pulls = chart.data.map((trace) => pullsForTrace(trace, label));
            Plotly.restyle(chart, {{pull: pulls}}, [0, 1]);
        }});
        chart.on("plotly_unhover", function() {{
            Plotly.restyle(chart, {{pull: [[], []]}}, [0, 1]);
        }});
        """,
    )
    components.html(html, height=470)


def render_top_failed_chart(df: pd.DataFrame, title: str = "Top validations failed") -> None:
    if df.empty:
        st.info(f"No data available for {title}.")
        return
    chart_theme = st.session_state.get("chart_theme", {})
    work = df.head(10).copy().reset_index(drop=True)
    work["rank"] = [f"#{idx}" for idx in range(1, len(work) + 1)]
    has_comparison_columns = {"fallout_count_old", "fallout_delta"}.issubset(work.columns)
    custom_data = ["validation_name", "table_name"]
    if has_comparison_columns:
        custom_data.extend(["fallout_count_old", "fallout_delta"])
    fig = px.bar(
        work,
        x="rank",
        y="fallout_count_new",
        text="fallout_count_new",
        custom_data=custom_data,
        title=title,
    )
    hovertemplate = (
        "<b>%{customdata[0]}</b><br>"
        "Table: %{customdata[1]}<br>"
        "Current fallouts: %{y:,}<extra></extra>"
    )
    if has_comparison_columns:
        hovertemplate = (
            "<b>%{customdata[0]}</b><br>"
            "Table: %{customdata[1]}<br>"
            "Current fallouts: %{y:,}<br>"
            "Previous fallouts: %{customdata[2]:,}<br>"
            "Delta: %{customdata[3]:+,}<extra></extra>"
        )
    fig.update_traces(
        hovertemplate=hovertemplate,
        textposition="outside",
        marker_line_width=0,
        marker_color="#38BDF8",
    )
    style_plotly_chart(
        fig,
        height=360,
        xaxis_title="Validation rank",
        yaxis_title="Failed count",
        showlegend=False,
        margin={"l": 40, "r": 24, "t": 56, "b": 50},
    )
    grid = chart_theme.get("grid", "rgba(51,65,85,0.45)")
    fig.update_xaxes(
        gridcolor="rgba(0,0,0,0)",
        zerolinecolor=grid,
        type="category",
        tickmode="array",
        tickvals=work["rank"].tolist(),
        ticktext=work["rank"].tolist(),
    )
    fig.update_yaxes(gridcolor=grid, zerolinecolor=grid)
    st.plotly_chart(fig, use_container_width=True)


def render_major_minor_fallout_charts(results: dict[str, Any]) -> None:
    c1, c2 = st.columns(2)
    with c1:
        render_top_failed_chart(results.get("top_major_fallout", pd.DataFrame()), "Top 10 major fallouts")
    with c2:
        render_top_failed_chart(results.get("top_minor_fallout", pd.DataFrame()), "Top 10 minor fallouts")


def render_table_section(title: str, df: pd.DataFrame, filename: str) -> None:
    st.markdown(f'<div class="mf-eyebrow">{escape(title)}</div>', unsafe_allow_html=True)
    if not df.empty and ("status" in df.columns or any(column.endswith("_delta") for column in df.columns)):
        dark = st.session_state.get("theme_mode", "Dark") == "Dark"

        def status_style(value: object) -> str:
            colors = st.session_state.get("chart_theme", {}).get("status", {})
            return colors.get(str(value).lower(), colors.get("default", "color: #CBD5E1;"))

        def delta_style(value: object) -> str:
            try:
                number = float(value)
            except (TypeError, ValueError):
                number = 0
            if number > 0:
                return (
                    "background-color: rgba(239, 68, 68, 0.16); color: #FECACA; font-weight: 800;"
                    if dark
                    else "background-color: rgba(220, 38, 38, 0.10); color: #991B1B; font-weight: 800;"
                )
            if number < 0:
                return (
                    "background-color: rgba(34, 197, 94, 0.16); color: #BBF7D0; font-weight: 800;"
                    if dark
                    else "background-color: rgba(22, 163, 74, 0.10); color: #166534; font-weight: 800;"
                )
            return (
                "background-color: rgba(148, 163, 184, 0.10); color: #CBD5E1;"
                if dark
                else "background-color: rgba(100, 116, 139, 0.08); color: #475569;"
            )

        styled_df = df.style
        if "status" in df.columns:
            styled_df = styled_df.map(status_style, subset=["status"])
        delta_columns = [column for column in df.columns if column.endswith("_delta")]
        if delta_columns:
            styled_df = styled_df.map(delta_style, subset=delta_columns)

        render_dataframe(
            df,
            key=filename.replace(".", "_"),
            title=title,
            filename=filename,
            styled_df=styled_df,
        )
    else:
        render_dataframe(df, key=filename.replace(".", "_"), title=title, filename=filename)
    dataframe_download(df, f"Download {title} CSV", filename)


def choose_latest_pair(old_uploaded, new_uploaded) -> tuple[ParsedReport | None, ParsedReport | None]:
    reports = [report for report in [old_uploaded, new_uploaded] if report is not None]
    if len(reports) != 2:
        return old_uploaded, new_uploaded
    latest_name = decide_latest_report([report.filename for report in reports])
    latest = next(report for report in reports if report.filename == latest_name)
    older = next(report for report in reports if report.filename != latest_name)
    return older, latest


def compare_and_store(old_report: ParsedReport, new_report: ParsedReport) -> None:
    st.session_state["old_report"] = old_report
    st.session_state["new_report"] = new_report
    st.session_state["comparison_results"] = compare_reports(old_report, new_report)
    st.session_state["comparison_results"]["mode"] = "comparison"
    st.session_state["comparison_results"] = _add_fallout_segments(
        st.session_state["comparison_results"],
        _validation_severity_lookup(old_report, new_report),
    )
    st.session_state["current_report_results"] = build_current_report_results(new_report)
    st.session_state["analysis_mode"] = "comparison"
    st.session_state.pop("ai_executive_summary", None)
    st.success(f"Compared old report '{old_report.filename}' with latest report '{new_report.filename}'.")


def set_current_report(report: ParsedReport) -> None:
    st.session_state["old_report"] = None
    st.session_state["new_report"] = report
    st.session_state["comparison_results"] = None
    st.session_state["current_report_results"] = build_current_report_results(report)
    st.session_state["analysis_mode"] = "current"
    st.session_state.pop("ai_executive_summary", None)
    st.success(f"Current report loaded: {report.filename}")


def run_outlook_scan_once(
    preset: str,
    keywords: str,
    unread_only: bool,
    start_dt: datetime | None = None,
    end_dt: datetime | None = None,
    overwrite_same_filename: bool = False,
) -> tuple[int, int]:
    start, end = date_range_from_preset(preset, start_dt, end_dt)
    matches = scan_outlook(keywords, unread_only=unread_only, start=start, end=end)
    details = [match.__dict__ for match in matches]
    attachment_count = sum(len(m.attachments) for m in matches)
    stored_count = 0
    for match in matches:
        for attachment in match.attachments:
            report = parse_excel_report(Path(attachment).read_bytes(), Path(attachment).name)
            if database.find_report_by_filename(report.filename) and not overwrite_same_filename:
                for detail in details:
                    if attachment in detail.get("attachments", []):
                        detail["skipped"] = "Report with same filename already exists."
                continue
            database.save_report(report, source="outlook", overwrite_same_filename=overwrite_same_filename)
            stored_count += 1
    database.save_scan_history(preset, keywords, len(matches), stored_count, details)
    return len(matches), stored_count


def default_report_scan_folder() -> str:
    downloads = Path.home() / "Downloads"
    return str(downloads if downloads.exists() else ATTACHMENT_DIR)


def scan_report_folder(
    folder_path: str,
    keywords: str,
    preset: str,
    start_dt: datetime | None = None,
    end_dt: datetime | None = None,
    overwrite_same_filename: bool = False,
) -> tuple[int, int, list[dict[str, Any]]]:
    start, end = date_range_from_preset(preset, start_dt, end_dt)
    folder = Path(folder_path).expanduser()
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Folder was not found: {folder}")

    matches: list[dict[str, Any]] = []
    for file_path in folder.rglob("*"):
        if not file_path.is_file() or file_path.suffix.lower() not in EXCEL_EXTENSIONS:
            continue
        try:
            modified = datetime.fromtimestamp(file_path.stat().st_mtime)
        except OSError:
            continue
        if start and modified < start:
            continue
        if end and modified > end:
            continue
        is_match, matched_keywords = flexible_match(file_path.name, keywords)
        if keywords.strip() and not is_match:
            continue
        matches.append(
            {
                "subject": f"Local folder file: {file_path.name}",
                "sender": "Local folder",
                "received_time": modified.isoformat(timespec="seconds"),
                "matched_keywords": matched_keywords,
                "attachments": [str(file_path)],
            }
        )

    saved_count = 0
    for match in matches:
        for attachment in match["attachments"]:
            path = Path(attachment)
            try:
                report = parse_excel_report(path.read_bytes(), path.name)
                if database.find_report_by_filename(report.filename) and not overwrite_same_filename:
                    match["skipped"] = "Report with same filename already exists."
                    continue
                database.save_report(report, source="folder", overwrite_same_filename=overwrite_same_filename)
                saved_count += 1
            except Exception as exc:
                match["parse_error"] = str(exc)

    database.save_scan_history(f"Folder: {preset}", keywords, len(matches), saved_count, matches)
    return len(matches), saved_count, matches


def run_report_scan_once(
    source_mode: str,
    preset: str,
    keywords: str,
    unread_only: bool,
    folder_path: str,
    start_dt: datetime | None = None,
    end_dt: datetime | None = None,
    overwrite_same_filename: bool = False,
) -> tuple[str, int, int]:
    normalized = source_mode.lower()
    if normalized == "classic outlook":
        match_count, saved_count = run_outlook_scan_once(
            preset, keywords, unread_only, start_dt, end_dt, overwrite_same_filename
        )
        return "Classic Outlook", match_count, saved_count
    if normalized == "local folder":
        match_count, saved_count, _ = scan_report_folder(folder_path, keywords, preset, start_dt, end_dt, overwrite_same_filename)
        return "Local folder", match_count, saved_count

    try:
        match_count, saved_count = run_outlook_scan_once(
            preset, keywords, unread_only, start_dt, end_dt, overwrite_same_filename
        )
        return "Classic Outlook", match_count, saved_count
    except RuntimeError as exc:
        if "COM" not in str(exc) and "New Outlook" not in str(exc) and "registered" not in str(exc):
            raise
        match_count, saved_count, _ = scan_report_folder(folder_path, keywords, preset, start_dt, end_dt, overwrite_same_filename)
        st.info("Classic Outlook COM was not available, so the app scanned the configured local folder instead.")
        return "Local folder", match_count, saved_count


def calculate_next_scan_run(
    cadence: str,
    run_at: time,
    weekly_day: str = "Monday",
    monthly_day: int = 1,
    custom_every: int = 1,
    custom_unit: str = "days",
    now: datetime | None = None,
) -> datetime:
    now = now or datetime.now()
    candidate = datetime.combine(now.date(), run_at)
    if cadence == "Daily":
        return candidate if candidate > now else candidate + timedelta(days=1)
    if cadence == "Weekly":
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        target_weekday = weekdays.index(weekly_day)
        days_ahead = (target_weekday - now.weekday()) % 7
        candidate = datetime.combine(now.date() + timedelta(days=days_ahead), run_at)
        return candidate if candidate > now else candidate + timedelta(days=7)
    if cadence == "Monthly":
        day = max(1, min(int(monthly_day), 28))
        candidate = datetime.combine(date(now.year, now.month, day), run_at)
        if candidate > now:
            return candidate
        month = now.month + 1
        year = now.year + (1 if month == 13 else 0)
        month = 1 if month == 13 else month
        return datetime.combine(date(year, month, day), run_at)
    amount = max(1, int(custom_every))
    delta = timedelta(hours=amount) if custom_unit == "hours" else timedelta(days=amount)
    return now + delta


def main() -> None:
    ensure_data_dirs()
    database.init_db()
    sync_theme_from_local_storage()
    theme = initialize_theme()
    apply_theme(theme)
    if st.query_params.get(THEME_QUERY_KEY) is not None or not st.session_state.get("_theme_from_default", False):
        persist_theme_preference(theme)
    else:
        apply_theme_class(theme)

    with st.sidebar:
        render_sidebar_brand()
        st.caption("Quick controls")
        selected_theme = st.radio(
            "Theme",
            list(THEME_OPTIONS),
            index=THEME_OPTIONS.index(theme),
            horizontal=True,
            key="theme_mode",
        )
        if selected_theme != theme:
            persist_theme_preference(selected_theme)
            st.rerun()
        st.session_state["ai_enabled"] = st.toggle("Enable AI insights", value=st.session_state.get("ai_enabled", False))
        st.divider()
        selected_page = st.radio(
            "Navigation",
            NAV_OPTIONS,
            index=NAV_OPTIONS.index(st.session_state.get("selected_page", "Dashboard"))
            if st.session_state.get("selected_page", "Dashboard") in NAV_OPTIONS
            else 0,
            key="selected_page",
        )
        st.divider()
        st.header("AI configuration")
        st.caption("AI is optional and only uses aggregate comparison data.")
        if "ai_provider_select" not in st.session_state:
            try:
                st.session_state["ai_provider_select"] = get_provider_preference() or "OpenAI"
            except CredentialStoreError:
                st.session_state["ai_provider_select"] = "OpenAI"
        provider = st.selectbox("AI provider", list(AI_PROVIDER_DEFAULTS.keys()), key="ai_provider_select")
        if st.session_state.get("ai_provider_last") != provider:
            defaults = AI_PROVIDER_DEFAULTS[provider]
            st.session_state["ai_base_url_input"] = defaults["base_url"]
            st.session_state["ai_model_input"] = defaults["model"]
            st.session_state["ai_provider_last"] = provider
            try:
                save_provider_preference(provider)
            except CredentialStoreError as exc:
                st.session_state["keyring_warning"] = str(exc)
            get_active_api_key(provider)
        st.session_state["ai_provider"] = provider
        st.session_state["ai_base_url"] = st.text_input("API base URL", key="ai_base_url_input")
        st.session_state["ai_model"] = st.text_input("Model name", key="ai_model_input")
        render_api_key_controls(
            provider,
            location="sidebar",
            compact=True,
            show_test=True,
            base_url=st.session_state["ai_base_url"],
            model=st.session_state["ai_model"],
        )
        if st.session_state.get("keyring_warning"):
            st.warning(st.session_state["keyring_warning"])

    if selected_page == "About":
        render_about_page()

    elif selected_page == "Dashboard":
        render_app_header()
        render_dashboard_actions()
        results = get_active_analysis_results()
        if not results:
            st.info("Upload/select reports in Compare Reports, or analyze a current report only, to populate the dashboard.")
        elif results.get("mode") == "current":
            st.caption(f"Current report analysis: {results.get('report_name', 'Current report')}")
            render_current_report_kpis(results)
            render_migration_summary_pies(None, st.session_state.get("new_report"))
            render_major_minor_fallout_charts(results)
            render_table_section("Current Report Validations", results["current_validations"], "current_report_validations.csv")
        else:
            render_kpis(results)
            render_migration_summary_pies(st.session_state.get("old_report"), st.session_state.get("new_report"))
            render_major_minor_fallout_charts(results)
            render_table_section("Full Comparison", results["comparison"], "full_comparison.csv")

    elif selected_page == "Compare Reports":
        st.subheader("Report analysis")
        mode = st.radio(
            "Analysis source",
            ["Upload two Excel files", "Select stored reports", "Analyze current report only"],
            horizontal=True,
        )
        if mode == "Upload two Excel files":
            c1, c2 = st.columns(2)
            old_file = c1.file_uploader("Old report", type=["xlsx", "xlsm", "xls"], key="old_upload")
            new_file = c2.file_uploader("New report", type=["xlsx", "xlsm", "xls"], key="new_upload")
            auto_latest = st.checkbox("Decide latest report from filename timestamp", value=True)
            overwrite_uploads = st.checkbox(
                "Overwrite stored reports when the exact same filename already exists",
                value=False,
                help="If off, duplicate filenames are left unchanged and the upload is not stored.",
                key="overwrite_compare_uploads",
            )
            if st.button("Compare uploaded reports", type="primary"):
                old_report = load_uploaded_report(old_file, overwrite_same_filename=overwrite_uploads)
                new_report = load_uploaded_report(new_file, overwrite_same_filename=overwrite_uploads)
                if old_report and new_report:
                    if auto_latest:
                        old_report, new_report = choose_latest_pair(old_report, new_report)
                    compare_and_store(old_report, new_report)
        elif mode == "Select stored reports":
            options = report_options()
            if len(options) < 2:
                st.warning("Store at least two reports before comparing from SQLite.")
            else:
                c1, c2 = st.columns(2)
                old_id = c1.selectbox("Old stored report", options=list(options.keys()))
                new_id = c2.selectbox("New stored report", options=list(options.keys()), index=1)
                if st.button("Compare stored reports", type="primary"):
                    compare_and_store(database.load_report(options[old_id]), database.load_report(options[new_id]))
        else:
            source = st.radio("Current report source", ["Upload one Excel file", "Select one stored report"], horizontal=True)
            if source == "Upload one Excel file":
                current_file = st.file_uploader("Current/latest report", type=["xlsx", "xlsm", "xls"], key="current_upload")
                overwrite_current = st.checkbox(
                    "Overwrite stored report when the exact same filename already exists",
                    value=False,
                    help="If off, duplicate filenames are left unchanged and the upload is not stored.",
                    key="overwrite_current_upload",
                )
                if st.button("Analyze current report", type="primary"):
                    current_report = load_uploaded_report(current_file, overwrite_same_filename=overwrite_current)
                    if current_report:
                        set_current_report(current_report)
            else:
                options = report_options()
                if not options:
                    st.warning("No stored reports available yet.")
                else:
                    current_id = st.selectbox("Current stored report", options=list(options.keys()))
                    if st.button("Analyze selected report", type="primary"):
                        set_current_report(database.load_report(options[current_id]))

    elif selected_page == "Fallouts":
        results = get_current_results()
        active_results = get_active_analysis_results()
        fallout_options = (
            ["Top 10", "Customers", "Explorer"]
            if active_results and active_results.get("mode") == "current"
            else ["New", "Resolved", "Top 10", "Customers", "Explorer"]
        )
        if st.session_state.get("fallout_view") not in fallout_options:
            st.session_state["fallout_view"] = fallout_options[0]
        fallout_view = st.radio(
            "Fallout view",
            fallout_options,
            horizontal=True,
            key="fallout_view",
        )
        if fallout_view == "Explorer":
            current = st.session_state.get("new_report")
            if not current:
                reports = database.list_reports()
                if reports:
                    current = database.load_report(reports[0]["id"])
            if not current:
                st.info("No current report available. Upload or store a report first.")
            else:
                st.subheader(f"Current report: {current.filename}")
                sheet_names = list(current.sheets.keys())
                selected_sheet = st.session_state.get("explorer_selected_sheet", sheet_names[0])
                if selected_sheet not in current.sheets:
                    selected_sheet = sheet_names[0]
                    st.session_state["explorer_selected_sheet"] = selected_sheet
                for start in range(0, len(sheet_names), 4):
                    cols = st.columns(4)
                    for offset, sheet_name in enumerate(sheet_names[start : start + 4]):
                        button_type = "primary" if sheet_name == selected_sheet else "secondary"
                        if cols[offset].button(sheet_name, key=f"explorer_sheet_{start + offset}", type=button_type, use_container_width=True):
                            st.session_state["explorer_selected_sheet"] = sheet_name
                            selected_sheet = sheet_name
                df = current.sheets[selected_sheet]
                query = st.text_input("Search/filter selected sheet")
                if query:
                    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)
                    df = df[mask]
                render_dataframe(df, key=f"explorer_{selected_sheet}", title=f"Explorer - {selected_sheet}", filename=f"{selected_sheet}.csv")
                dataframe_download(df, "Download selected sheet CSV", f"{selected_sheet}.csv")
        elif not active_results:
            st.info("No comparison or current report analysis loaded.")
        elif active_results.get("mode") == "current":
            if fallout_view == "Top 10":
                render_table_section("Top 10 Major Fallouts", active_results["top_major_fallout"], "current_top_10_major_fallout.csv")
                render_table_section("Top 10 Minor Fallouts", active_results["top_minor_fallout"], "current_top_10_minor_fallout.csv")
            elif fallout_view == "Customers":
                render_table_section("Top 10 Current Validations by Customer Count", active_results["top_customers"], "current_top_10_customer_impact.csv")
        else:
            if fallout_view == "New":
                render_table_section("New Fallouts", results["new_validations"], "new_fallouts.csv")
            elif fallout_view == "Resolved":
                render_table_section("Resolved Fallouts", results["resolved_validations"], "resolved_fallouts.csv")
            elif fallout_view == "Top 10":
                render_table_section("Top 10 Major Fallouts", results["top_major_fallout"], "top_10_major_fallout.csv")
                render_table_section("Top 10 Minor Fallouts", results["top_minor_fallout"], "top_10_minor_fallout.csv")
                render_table_section("Top 10 Increased by Fallout Count", results["increased_fallout"], "top_10_increased_fallout.csv")
                render_table_section("Top 10 Reduced by Fallout Count", results["reduced_fallout"], "top_10_reduced_fallout.csv")
            elif fallout_view == "Customers":
                render_table_section("Top 10 Current Validations by Customer Count", results["top_customers"], "top_10_customer_impact.csv")
                render_table_section("Top 10 Increased by Customer Count", results["increased_customers"], "top_10_increased_customers.csv")
                render_table_section("Top 10 Reduced by Customer Count", results["reduced_customers"], "top_10_reduced_customers.csv")

    elif selected_page == "Outlook Scanner":
        st.subheader("Report scanner")
        st.caption("Manual scans run immediately. Scheduled scans run while this dashboard is open.")
        scan_source = st.radio(
            "Scan source",
            ["Auto", "Classic Outlook", "Local folder"],
            horizontal=True,
            help=(
                "Auto uses classic Outlook COM when available. If this machine only has New Outlook, "
                "it scans the local folder you configure below."
            ),
        )
        folder_path = st.text_input(
            "Local report folder for New Outlook or fallback scans",
            value=st.session_state.get("report_scan_folder", default_report_scan_folder()),
            help="For New Outlook machines, save/download report attachments into this folder and scan it.",
        )
        st.session_state["report_scan_folder"] = folder_path
        if scan_source == "Classic Outlook":
            st.info("Classic Outlook mode requires desktop Outlook COM/MAPI. New Outlook for Windows does not support this.")
        elif scan_source == "Local folder":
            st.info("Local folder mode works with New Outlook, Outlook Web, Teams downloads, shared folders, and manually saved reports.")
        else:
            st.info("Auto mode tries classic Outlook first and falls back to the local folder when COM is unavailable.")
        unread_only = st.checkbox("Scan unread emails only", value=True)
        preset = st.selectbox("Email date range to scan", ["Last day", "Last week", "Last month", "Custom range"])
        start_dt = end_dt = None
        if preset == "Custom range":
            c1, c2 = st.columns(2)
            start_dt = datetime.combine(c1.date_input("Start date"), datetime.min.time())
            end_dt = datetime.combine(c2.date_input("End date"), datetime.max.time())
        keywords = st.text_area("Subject/body/attachment keywords", value="\n".join(DEFAULT_EMAIL_KEYWORDS), height=120)
        overwrite_scanned_reports = st.checkbox(
            "Overwrite stored reports when the exact same filename already exists",
            value=False,
            help="If off, scan results with duplicate filenames are skipped and existing stored reports are kept.",
            key="overwrite_scanned_reports",
        )

        scan_cols = st.columns([1, 1, 4])
        if scan_cols[0].button("Scan now", type="primary", use_container_width=True, disabled=st.session_state.get("outlook_scan_running", False)):
            st.session_state["outlook_scan_running"] = True
            try:
                used_source, match_count, attachment_count = run_report_scan_once(
                    scan_source, preset, keywords, unread_only, folder_path, start_dt, end_dt, overwrite_scanned_reports
                )
                st.success(f"{used_source} scan found {match_count} matching item(s). Stored {attachment_count} report file(s).")
            except Exception as exc:
                st.error(f"Report scan failed: {exc}")
            finally:
                st.session_state["outlook_scan_running"] = False
        if scan_cols[1].button("Stop scan", use_container_width=True):
            st.session_state["outlook_scan_running"] = False
            st.session_state["outlook_schedule_enabled"] = False
            st.warning("Stopped scheduled scanning. A scan already in progress may finish its current request.")

        st.divider()
        st.subheader("Scheduled scan job")
        schedule_enabled = st.toggle(
            "Start scheduled scan job",
            value=st.session_state.get("outlook_schedule_enabled", False),
            key="outlook_schedule_enabled",
        )
        schedule_cols = st.columns(4)
        cadence = schedule_cols[0].selectbox("Run every", ["Daily", "Weekly", "Monthly", "Custom"], key="outlook_schedule_cadence")
        run_at = schedule_cols[1].time_input("At time", value=st.session_state.get("outlook_schedule_time", time(9, 0)), key="outlook_schedule_time")
        weekly_day = "Monday"
        monthly_day = 1
        custom_every = 1
        custom_unit = "days"
        if cadence == "Weekly":
            weekly_day = schedule_cols[2].selectbox(
                "Day",
                ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
                key="outlook_schedule_weekday",
            )
        elif cadence == "Monthly":
            monthly_day = schedule_cols[2].number_input("Day", min_value=1, max_value=28, value=1, key="outlook_schedule_month_day")
        elif cadence == "Custom":
            custom_every = schedule_cols[2].number_input("Every", min_value=1, max_value=365, value=1, key="outlook_schedule_every")
            custom_unit = schedule_cols[3].selectbox("Unit", ["hours", "days"], key="outlook_schedule_unit")

        if schedule_enabled and st.session_state.get("outlook_next_run") is None:
            st.session_state["outlook_next_run"] = calculate_next_scan_run(
                cadence, run_at, weekly_day, monthly_day, custom_every, custom_unit
            )
        if st.button("Update schedule", use_container_width=True):
            st.session_state["outlook_next_run"] = calculate_next_scan_run(
                cadence, run_at, weekly_day, monthly_day, custom_every, custom_unit
            )
            st.success(f"Next report scan: {st.session_state['outlook_next_run'].strftime('%Y-%m-%d %H:%M')}")
        if not schedule_enabled:
            st.session_state["outlook_next_run"] = None
            st.info("Scheduled scan job is stopped.")
        else:
            next_run = st.session_state.get("outlook_next_run")
            components.html(
                """
                <script>
                window.setTimeout(() => window.parent.location.reload(), 60000);
                </script>
                """,
                height=0,
                width=0,
            )
            if next_run:
                st.info(f"Scheduled scan job is active. Next scan: {next_run.strftime('%Y-%m-%d %H:%M')}")
            if next_run and datetime.now() >= next_run and not st.session_state.get("outlook_scan_running", False):
                st.session_state["outlook_scan_running"] = True
                try:
                    used_source, match_count, attachment_count = run_report_scan_once(
                        scan_source, preset, keywords, unread_only, folder_path, start_dt, end_dt, overwrite_scanned_reports
                    )
                    st.session_state["outlook_last_scheduled_scan"] = datetime.now()
                    st.session_state["outlook_next_run"] = calculate_next_scan_run(
                        cadence, run_at, weekly_day, monthly_day, custom_every, custom_unit
                    )
                    st.success(
                        f"Scheduled {used_source} scan finished. Found {match_count} item(s), stored {attachment_count} report file(s)."
                    )
                except Exception as exc:
                    st.error(f"Scheduled report scan failed: {exc}")
                    st.session_state["outlook_next_run"] = calculate_next_scan_run(
                        cadence, run_at, weekly_day, monthly_day, custom_every, custom_unit
                    )
                finally:
                    st.session_state["outlook_scan_running"] = False
            if st.session_state.get("outlook_last_scheduled_scan"):
                last_run = st.session_state["outlook_last_scheduled_scan"]
                st.caption(f"Last scheduled scan: {last_run.strftime('%Y-%m-%d %H:%M')}")

        st.subheader("Scan history")
        render_dataframe(
            pd.DataFrame(database.list_scan_history()),
            key="scan_history",
            title="Report Scan History",
            filename="report_scan_history.csv",
        )

    elif selected_page == "Stored Reports":
        st.subheader("Stored reports")
        reports_df = pd.DataFrame(database.list_reports())
        render_dataframe(reports_df, key="stored_reports", title="Stored Reports", filename="stored_reports.csv")
        options = report_options()
        if options:
            selected = st.selectbox("Delete stored report", list(options.keys()))
            if st.button("Delete selected report"):
                database.delete_report(options[selected])
                st.success("Deleted selected report. Refresh the page to update the table.")

    elif selected_page == "AI Insights":
        results = get_active_analysis_results()
        st.info(
            "To see AI insights: save or enter an API key, enable AI insights in the sidebar, then run a comparison or analyze one current report."
        )
        if not results:
            st.info("Run a comparison or analyze a current report before generating AI insights.")
        elif not st.session_state.get("ai_enabled"):
            st.info("Enable AI insights in the sidebar.")
        elif st.button("Generate executive summary", type="primary"):
            api_key = get_active_api_key(st.session_state["ai_provider"])
            try:
                summary = generate_ai_insights(
                    st.session_state["ai_provider"],
                    api_key,
                    st.session_state["ai_base_url"],
                    st.session_state["ai_model"],
                    results,
                )
                st.session_state["ai_executive_summary"] = summary
            except Exception as exc:
                st.error(f"AI insight generation failed: {sanitize_error(exc, api_key)}")
        summary = st.session_state.get("ai_executive_summary")
        if summary:
            st.divider()
            st.subheader("AI Insights Dashboard")
            render_ai_insights_dashboard(summary, results)

    elif selected_page == "Fallout Insight Engine":
        results = get_active_analysis_results()
        current_report = st.session_state.get("new_report")
        st.subheader("Fallout Insight Engine")
        st.caption("Ask natural-language questions about the latest report or the comparison differences.")
        if not st.session_state.get("ai_enabled"):
            st.info("Enable AI insights in the sidebar, choose your provider, and save or enter an API key to use Fallout Insight Engine.")
        elif not get_active_api_key(st.session_state["ai_provider"]):
            st.warning("Fallout Insight Engine needs a working API key. Add or load your key in the sidebar, then use Test API connection.")
        elif not current_report and not results:
            st.info("Run a comparison or analyze a current report first, then Fallout Insight Engine can answer questions about it.")
        else:
            if "insight_copilot_history" not in st.session_state:
                st.session_state["insight_copilot_history"] = []
            examples = [
                "What changed the most between the previous and latest report?",
                "Which validations are the biggest risk right now?",
                "Summarize the latest report for an operations lead.",
                "Which tables improved and which got worse?",
            ]
            example = st.selectbox("Example questions", [""] + examples, key="copilot_example")
            if example and st.session_state.get("copilot_last_example") != example:
                st.session_state["insight_copilot_question"] = example
                st.session_state["copilot_last_example"] = example
            question = st.text_area(
                "Ask about the latest report or comparison",
                placeholder="Example: Which validations increased the most and what should I investigate first?",
                height=110,
                key="insight_copilot_question",
            )
            c1, c2 = st.columns([1, 4])
            ask_clicked = c1.button("Ask Engine", type="primary", use_container_width=True)
            if c2.button("Clear conversation", use_container_width=True):
                st.session_state["insight_copilot_history"] = []
                st.rerun()
            if ask_clicked:
                api_key = get_active_api_key(st.session_state["ai_provider"])
                try:
                    answer = answer_report_question(
                        st.session_state["ai_provider"],
                        api_key,
                        st.session_state["ai_base_url"],
                        st.session_state["ai_model"],
                        question,
                        results,
                        current_report,
                    )
                    st.session_state["insight_copilot_history"].append({"question": question, "answer": answer})
                except Exception as exc:
                    st.error(f"Fallout Insight Engine could not answer yet: {sanitize_error(exc, api_key)}")

            for item in reversed(st.session_state["insight_copilot_history"]):
                with st.container(border=True):
                    st.markdown(f"**You asked:** {escape(item['question'])}")
                    st.markdown(item["answer"])

    elif selected_page == "Settings":
        st.subheader("Settings")
        st.write(f"Database: `{database.DB_PATH}`")
        st.write(f"Attachment folder: `{ATTACHMENT_DIR}`")
        st.write("API keys are stored per provider in Windows Credential Manager under `MigrationFalloutDashboard` when keyring is available.")
        provider_for_settings = st.selectbox("Credential provider", ["Kimi", "OpenAI", "Claude", "Custom"], key="settings_provider")
        try:
            st.write(f"Saved key status: `{ 'saved' if has_saved_api_key(provider_for_settings) else 'not saved' }`")
        except CredentialStoreError as exc:
            st.warning(f"{exc} Session-only API key entry is still available in the sidebar.")
        st.write("Important sheets are parsed automatically when present, and every other sheet is preserved for exploration.")
        if st.button("Initialize/repair local database"):
            database.init_db()
            st.success("Database is ready.")


if __name__ == "__main__":
    main()
