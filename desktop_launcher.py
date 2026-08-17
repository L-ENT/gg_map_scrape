"""Windows desktop launcher for the local Streamlit clinic collector."""
import os
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path


def resource_path(filename: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


def open_local_app(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    for _ in range(45):
        try:
            urllib.request.urlopen(url, timeout=1)
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(1)


def main() -> None:
    from streamlit.web import cli as stcli

    port = 8501
    app_file = resource_path("app.py")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    threading.Thread(target=open_local_app, args=(port,), daemon=True).start()
    sys.argv = [
        "streamlit", "run", str(app_file),
        "--server.address", "127.0.0.1",
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.fileWatcherType", "none",
        "--browser.gatherUsageStats", "false",
        # A PyInstaller build must never inherit Streamlit's development mode:
        # that mode conflicts with the fixed localhost port used by the desktop
        # launcher and prevents the application from starting.
        "--global.developmentMode", "false",
    ]
    stcli.main()


if __name__ == "__main__":
    main()
