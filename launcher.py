from __future__ import annotations

import os
import sys
from pathlib import Path


def _bundle_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def main() -> None:
    bundle_dir = _bundle_dir()
    app_path = bundle_dir / "app.py"
    if not app_path.exists():
        app_path = Path(__file__).resolve().with_name("app.py")

    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    sys.path.insert(0, str(bundle_dir))
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
    ]

    from streamlit.web.cli import main as streamlit_main

    streamlit_main()


if __name__ == "__main__":
    main()
