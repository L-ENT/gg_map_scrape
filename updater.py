"""Small companion program that installs a downloaded app update safely.

It deliberately does not need a GitHub credential. The main app supplies a
public GitHub Release asset URL after the person has clicked the update button.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from urllib.request import urlretrieve


def wait_for_parent(pid: int, seconds: int = 30) -> None:
    """Wait briefly for the Streamlit launcher process to release its files."""
    if pid <= 0:
        return
    deadline = time.time() + seconds
    while time.time() < deadline:
        try:
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, check=False,
            )
            if str(pid) not in result.stdout:
                return
        except OSError:
            return
        time.sleep(1)


def replace_installation(source: Path, target: Path) -> None:
    """Copy everything except the updater binary currently in use."""
    target.mkdir(parents=True, exist_ok=True)
    for item in source.iterdir():
        if item.name.casefold() == "clinic lead updater.exe":
            continue
        destination = target / item.name
        if item.is_dir():
            shutil.copytree(item, destination, dirs_exist_ok=True)
        else:
            shutil.copy2(item, destination)


def main() -> int:
    parser = argparse.ArgumentParser(description="Install Clinic Lead Collector update")
    parser.add_argument("--target", required=True)
    parser.add_argument("--url", required=True)
    parser.add_argument("--parent-pid", type=int, default=0)
    args = parser.parse_args()
    target = Path(args.target).resolve()
    workspace = Path(tempfile.mkdtemp(prefix="clinic-lead-update-"))
    try:
        archive = workspace / "update.zip"
        urlretrieve(args.url, archive)
        with zipfile.ZipFile(archive) as zip_file:
            zip_file.extractall(workspace / "unzipped")
        package_root = workspace / "unzipped" / "Clinic Lead Collector"
        if not package_root.is_dir():
            raise RuntimeError("Gói cập nhật không có thư mục Clinic Lead Collector.")
        wait_for_parent(args.parent_pid)
        replace_installation(package_root, target)
        launcher = target / "Clinic Lead Collector.exe"
        if not launcher.exists():
            raise RuntimeError("Không tìm thấy Clinic Lead Collector.exe sau cập nhật.")
        # Windows keeps this updater executable locked while it runs. A detached
        # batch file copies that last file after this process exits, then starts
        # the newly installed main application.
        script = workspace / "finish-update.cmd"
        script.write_text(
            "@echo off\r\n"
            "timeout /t 2 /nobreak >nul\r\n"
            f'copy /y "{package_root / "Clinic Lead Updater.exe"}" "{target / "Clinic Lead Updater.exe"}" >nul\r\n'
            f'start "" "{launcher}"\r\n',
            encoding="utf-8",
        )
        subprocess.Popen(["cmd", "/c", str(script)], cwd=str(target), close_fds=True)
        return 0
    except Exception as exc:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Cập nhật không thành công:\n{exc}", "Clinic Lead Collector", 0x10)
        except Exception:
            print(f"Update failed: {exc}", file=sys.stderr)
        return 1
    finally:
        # On success the detached batch still reads this folder. Failed updates
        # are cleaned up immediately.
        if 'script' not in locals():
            shutil.rmtree(workspace, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
