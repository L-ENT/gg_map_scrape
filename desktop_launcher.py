"""Cross-platform desktop launcher for the local Streamlit clinic collector."""
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path


def show_startup_error(message: str) -> None:
    """Show a visible startup error even in a windowed PyInstaller build."""
    try:
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, "Clinic Lead Collector", 0x10)
        elif sys.platform == "darwin":
            script = 'on run argv\ndisplay alert "Clinic Lead Collector" message (item 1 of argv) as critical\nend run'
            subprocess.run(["/usr/bin/osascript", "-e", script, "--", message], check=False)
        else:
            print(message, file=sys.stderr)
    except Exception:
        print(message, file=sys.stderr)


def resource_path(filename: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


def local_port_is_open(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def local_app_is_healthy(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}", timeout=2) as response:
            return 200 <= response.status < 500
    except Exception:
        return False


def open_local_app(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    for _ in range(45):
        try:
            with urllib.request.urlopen(url, timeout=1):
                pass
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(1)
    show_startup_error("App đã khởi động nhưng localhost không phản hồi. Tiến trình sẽ được đóng để không giữ file; hãy mở app lại.")
    os._exit(1)


def main() -> None:
    from streamlit.web import cli as stcli

    port = 8501
    if local_port_is_open(port):
        if local_app_is_healthy(port):
            webbrowser.open(f"http://127.0.0.1:{port}")
        else:
            show_startup_error("Cổng localhost 8501 đang bị một phiên app cũ bị treo giữ. Hãy đóng Clinic Lead Collector trong Task Manager rồi mở lại.")
        return
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
