from pathlib import Path
import socket

import app
import desktop_launcher
import pytest
import updater


def test_macos_bundle_paths(tmp_path, monkeypatch):
    executable = tmp_path / "Clinic Lead Collector.app" / "Contents" / "MacOS" / "Clinic Lead Collector"
    executable.parent.mkdir(parents=True)
    executable.write_bytes(b"")
    version = executable.parents[1] / "Resources" / "version.txt"
    version.parent.mkdir()
    version.write_text("build-test", encoding="utf-8")
    monkeypatch.setattr(app.sys, "frozen", True, raising=False)
    monkeypatch.setattr(app.sys, "platform", "darwin")
    monkeypatch.setattr(app.sys, "executable", str(executable))

    bundle = executable.parents[2].resolve()
    assert app.installed_app_directory() == bundle
    assert app.installed_app_version() == "build-test"
    assert app.updater_executable() == bundle / "Contents" / "Resources" / "Clinic Lead Updater"


def test_macos_selects_arm64_release_asset(monkeypatch):
    class Response:
        status_code = 200

        @staticmethod
        def json():
            return {
                "tag_name": "build-test",
                "assets": [
                    {"name": "Clinic-Lead-Collector-windows.zip", "browser_download_url": "https://example.test/windows"},
                    {"name": "Clinic-Lead-Collector-macos-arm64.zip", "browser_download_url": "https://example.test/macos"},
                ],
            }

    monkeypatch.setattr(app.sys, "platform", "darwin")
    monkeypatch.setattr(app.requests, "get", lambda *args, **kwargs: Response())

    assert app.available_release() == {"version": "build-test", "url": "https://example.test/macos"}


def test_replace_macos_application_keeps_new_bundle(tmp_path, monkeypatch):
    target = tmp_path / "Applications" / "Clinic Lead Collector.app"
    source = tmp_path / "download" / "Clinic Lead Collector.app"
    target.mkdir(parents=True)
    source.mkdir(parents=True)
    (target / "version.txt").write_text("old", encoding="utf-8")
    (source / "version.txt").write_text("new", encoding="utf-8")
    launched = []
    monkeypatch.setattr(updater.subprocess, "Popen", lambda command, **kwargs: launched.append(command))

    updater.replace_macos_application(source, target)

    assert (target / "version.txt").read_text(encoding="utf-8") == "new"
    assert not target.with_name("Clinic Lead Collector.app.previous").exists()
    assert launched == [["/usr/bin/open", str(target)]]


def test_windows_archive_rejects_parent_path(tmp_path, monkeypatch):
    import zipfile

    archive = tmp_path / "unsafe.zip"
    with zipfile.ZipFile(archive, "w") as package:
        package.writestr("../outside.txt", "unsafe")
    monkeypatch.setattr(updater.sys, "platform", "win32")

    try:
        updater.extract_archive(archive, tmp_path / "output")
    except RuntimeError as exc:
        assert "không an toàn" in str(exc)
    else:
        raise AssertionError("Unsafe archive path was accepted")


def test_updater_refuses_to_continue_when_parent_does_not_exit():
    with pytest.raises(RuntimeError, match="chưa đóng hoàn toàn"):
        updater.wait_for_parent(12345, seconds=0)


def test_updater_refuses_to_continue_while_local_port_is_occupied():
    listener = socket.socket()
    listener.bind(("127.0.0.1", 0))
    listener.listen(1)
    port = listener.getsockname()[1]
    try:
        assert desktop_launcher.local_port_is_open(port) is True
        with pytest.raises(RuntimeError, match="vẫn đang bị"):
            updater.wait_for_port_release(port, seconds=0)
    finally:
        listener.close()
    updater.wait_for_port_release(port, seconds=1)
