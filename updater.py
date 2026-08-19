"""Small companion program that installs a downloaded app update safely.

It deliberately does not need a GitHub credential. The main app supplies a
public GitHub Release asset URL after the person has clicked the update button.
"""
from __future__ import annotations

import argparse
import os
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
            if sys.platform == "win32":
                result = subprocess.run(
                    ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                    stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, check=False,
                )
                if str(pid) not in result.stdout:
                    return
            else:
                os.kill(pid, 0)
        except ProcessLookupError:
            return
        except PermissionError:
            # The process exists but macOS will not let us signal it.
            pass
        except OSError:
            return
        time.sleep(1)


def extract_archive(archive: Path, destination: Path) -> None:
    """Extract a release without allowing paths to escape the workspace."""
    destination.mkdir(parents=True, exist_ok=True)
    destination_root = destination.resolve()
    with zipfile.ZipFile(archive) as zip_file:
        for member in zip_file.infolist():
            member_path = (destination / member.filename).resolve()
            try:
                member_path.relative_to(destination_root)
            except ValueError as exc:
                raise RuntimeError("Gói cập nhật chứa đường dẫn không an toàn.") from exc
    if sys.platform == "darwin":
        # ditto preserves executable modes, symlinks and macOS bundle metadata.
        subprocess.run(["/usr/bin/ditto", "-x", "-k", str(archive), str(destination)], check=True)
        return
    with zipfile.ZipFile(archive) as zip_file:
        zip_file.extractall(destination)


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


def replace_macos_application(source: Path, target: Path) -> None:
    """Replace one .app bundle with rollback if installation fails."""
    if target.suffix.casefold() != ".app":
        raise RuntimeError("Thư mục cài đặt macOS không phải là file .app.")
    backup = target.with_name(f"{target.name}.previous")
    if backup.exists():
        shutil.rmtree(backup)
    target.rename(backup)
    try:
        shutil.move(str(source), str(target))
    except Exception:
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)
        backup.rename(target)
        raise
    subprocess.Popen(["/usr/bin/open", str(target)], close_fds=True)
    shutil.rmtree(backup, ignore_errors=True)


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
        extracted = workspace / "unzipped"
        extract_archive(archive, extracted)
        package_name = "Clinic Lead Collector.app" if sys.platform == "darwin" else "Clinic Lead Collector"
        package_root = extracted / package_name
        if not package_root.is_dir():
            raise RuntimeError(f"Gói cập nhật không có {package_name}.")
        wait_for_parent(args.parent_pid)
        if sys.platform == "darwin":
            replace_macos_application(package_root, target)
            return 0
        replace_installation(package_root, target)
        launcher = target / "Clinic Lead Collector.exe"
        if not launcher.exists():
            raise RuntimeError("Không tìm thấy Clinic Lead Collector.exe sau cập nhật.")
        # Keep the running updater executable in place. The package also
        # contains "Clinic Lead Updater Payload.exe", which is copied above;
        # after this process exits, the restarted main app promotes that payload
        # over the old updater using the Unicode-safe Windows process API.
        subprocess.Popen([str(launcher)], cwd=str(target), close_fds=True)
        return 0
    except Exception as exc:
        try:
            if sys.platform == "darwin":
                script = 'on run argv\ndisplay alert "Clinic Lead Collector" message (item 1 of argv) as critical\nend run'
                subprocess.run(["/usr/bin/osascript", "-e", script, "--", f"Cập nhật không thành công:\n{exc}"], check=False)
            else:
                import ctypes
                ctypes.windll.user32.MessageBoxW(0, f"Cập nhật không thành công:\n{exc}", "Clinic Lead Collector", 0x10)
        except Exception:
            print(f"Update failed: {exc}", file=sys.stderr)
        return 1
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
