# Migration Fallout Report Intelligence Dashboard

Local Streamlit dashboard for comparing migration fallout Excel reports, storing parsed reports in SQLite, scanning Outlook for matching emails, and optionally generating executive summaries through a securely stored user-supplied AI API key.

## Features

- Upload two Excel reports and compare old vs new.
- Select two previously stored reports from SQLite.
- Automatically decide the latest report from filenames like `  Report_20260611113050.xlsx`.
- Parse every workbook sheet and preserve the full sheet data.
- Dynamically normalize columns such as `Rule Name`, `Validation Name`, `Table Name`, `Count of Fallouts`, `Count of customers`, and `Impacted customers`.
- Calculate new, resolved, increased, reduced, top fallout, customer impact, and table-wise summaries.
- Explore the most recent report sheet by sheet with search and CSV download.
- Scan Outlook unread or date-ranged emails using flexible keyword matching.
- Download matching Excel attachments and store them locally.
- Optional AI executive summary for Kimi, OpenAI, Claude, or any OpenAI-compatible custom endpoint.
- Secure API key storage per provider using Windows Credential Manager through `keyring`.
- Windows EXE build flow using PyInstaller.

## Project Structure

```text
migration_fallout_dashboard/
  app.py
  outlook_scanner.py
  excel_parser.py
  report_comparator.py
  database.py
  ai_insights.py
  credential_store.py
  filename_utils.py
  email_matcher.py
  config.py
  launcher.py
  build_windows.ps1
  requirements.txt
  README.md
```

## Developer Run Steps

```powershell
cd C:\Users\DELL\Desktop\Learning\AIWithPython\falloutanalyser\migration_fallout_dashboard
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

`requirements.txt` includes `keyring`, but you can install it directly if needed:

```powershell
pip install keyring
```

If PowerShell blocks activation, run this once for the current terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

## Outlook Notes

Outlook scanning requires:

- Windows.
- Microsoft Outlook desktop installed and configured.
- `pywin32` installed from `requirements.txt`.
- Outlook profile access allowed by your machine policy.

Use the Outlook Scanner tab to turn scanning on, choose unread/date range behavior, enter keywords, and download matching Excel attachments.

## AI Key Storage

AI is optional. No API key is hardcoded, written to SQLite, written to config files, included in builds, or printed intentionally in UI errors.

When `keyring` is available on Windows, API keys are stored in Windows Credential Manager:

- Service name: `MigrationFalloutDashboard`
- Account names: `Kimi:api_key`, `OpenAI:api_key`, `Claude:api_key`, `Custom:api_key`

Use the AI Insights or Settings tab to:

- Save API key.
- Load saved API key automatically.
- Test API connection.
- Update API key.
- Delete saved API key.

If Windows Credential Manager or the keyring backend is unavailable, the app shows a warning and falls back to session-only password entry. It does not fall back to plain-text storage.

Kimi and Custom use the OpenAI-compatible `/chat/completions` request format.

Core comparisons are deterministic Python logic; AI only summarizes aggregate comparison results.

## Build EXE Steps

Run this from Windows PowerShell:

```powershell
cd C:\Users\DELL\Desktop\Learning\AIWithPython\falloutanalyser\migration_fallout_dashboard
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build_windows.ps1 -Clean
```

The script will:

- Create `.venv` if needed.
- Install `requirements.txt`.
- Run PyInstaller.
- Create `dist\MigrationFalloutDashboard`.

The executable will be here:

```text
dist\MigrationFalloutDashboard\MigrationFalloutDashboard.exe
```

The build includes the app modules and local data folder structure. It excludes `.env`, API keys, build artifacts, and local SQLite database files through `.gitignore`.

## End-User Run Steps

Give the user the full folder:

```text
dist\MigrationFalloutDashboard
```

The user runs:

```text
MigrationFalloutDashboard.exe
```

The app starts a local Streamlit server and opens in the browser. End users can still:

- Upload and compare reports manually.
- Scan Outlook on Windows machines with Outlook configured.
- Store report history in local SQLite.
- Use saved Credential Manager API keys for AI insights.
- Switch between dark and light mode.

To remove a saved API key, open the Settings tab, choose the provider, and click `Delete saved key`.
