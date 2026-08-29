import io

import pandas as pd
from openpyxl import load_workbook

from app import (
    EXPORT_COLUMNS,
    AI_MODE_MIN_EVIDENCE_CHARS,
    AI_MODE_PENDING_MIN_EVIDENCE_CHARS,
    ai_mode_has_pending_marker,
    ai_mode_evidence_is_ready,
    append_rows_preserving_template,
    extract_operation_time_from_text,
    filter_result,
    promote_updater_payload,
    run_job,
    should_keep_in_final_output,
    wait_for_manual_captcha,
)


def test_extract_operation_time_from_relevant_sentence():
    page_text = "Monday-Friday 9:00 AM - 5:00 PM. We offer counseling services."
    assert extract_operation_time_from_text(page_text) == "Monday-Friday 9:00 AM - 5:00 PM"


def test_ai_mode_rejects_streaming_and_short_answers():
    partial = "Đang tìm kiếm " + ("clinic information " * 10)
    short_header = "Here is a brief, scannable overview of the clinic in Provo, Utah:"
    assert len(partial) < AI_MODE_PENDING_MIN_EVIDENCE_CHARS
    assert ai_mode_has_pending_marker(partial) is True
    assert ai_mode_evidence_is_ready(partial) is False
    assert ai_mode_evidence_is_ready(short_header) is False


def test_ai_mode_accepts_complete_answer_without_pending_marker():
    complete = "Owner Jane Doe. Private group counseling practice. " + ("Verified clinic details and services. " * 10)
    assert len(complete) >= AI_MODE_MIN_EVIDENCE_CHARS
    assert ai_mode_evidence_is_ready(complete) is True


def test_incomplete_ai_mode_is_not_sent_to_gemini(monkeypatch):
    row = {
        "Practice name": "Incomplete Clinic",
        "location": "Provo",
        "operation time and days": "N/A",
    }
    monkeypatch.setattr("app.maps_search_urls", lambda *args, **kwargs: [("https://maps.test/place", "Incomplete Clinic")])
    monkeypatch.setattr("app.extract_maps_place_with_retry", lambda *args, **kwargs: dict(row))
    ai_calls = []
    def incomplete_ai(*args, **kwargs):
        ai_calls.append(True)
        return ""
    monkeypatch.setattr("app.google_ai_overview", incomplete_ai)

    def unexpected_gemini_call(*args, **kwargs):
        raise AssertionError("Gemini must not be called without complete AI Mode evidence")

    monkeypatch.setattr("app.gemini_metadata", unexpected_gemini_call)
    known = set()
    accepted, debug = run_job(object(), "Provo", "UT", "therapy", 10, known, lambda message: None, "api-key")

    assert accepted == []
    assert debug[0]["Filter result"] == "RETRY: AI Mode chưa hoàn tất"
    assert debug[0]["Gemini"] == "not called"
    assert ("incomplete clinic", "provo") not in known
    assert len(ai_calls) == 2


def test_ai_mode_failure_is_retried_after_the_main_list(monkeypatch):
    rows = {
        "First Clinic": {"Practice name": "First Clinic", "location": "Provo", "operation time and days": "N/A"},
        "Second Clinic": {"Practice name": "Second Clinic", "location": "Provo", "operation time and days": "N/A"},
    }
    monkeypatch.setattr("app.maps_search_urls", lambda *args, **kwargs: [("first", "First Clinic"), ("second", "Second Clinic")])
    monkeypatch.setattr("app.extract_maps_place_with_retry", lambda driver, url, city, name, *args: dict(rows[name]))
    ai_order = []
    def ai_result(driver, name, *args, **kwargs):
        ai_order.append(name)
        if name == "First Clinic" and ai_order.count(name) == 1:
            return ""
        return "Complete AI Mode evidence " * 20
    monkeypatch.setattr("app.google_ai_overview", ai_result)
    monkeypatch.setattr("app.gemini_metadata", lambda *args, **kwargs: {
        "status": "ok", "owner": "N/A", "operation_time_and_days": "N/A",
        "doctor_count": 2, "branch_count": 1, "is_solo": False,
        "is_collective": False, "nonprofit": False, "private_practice": True,
        "target_service": True, "red_flags": [], "disallowed_provider_title": False,
        "outdated_or_insufficient": False, "over_25_years": False, "has_board": False,
    })
    candidate_progress = []
    retry_waiting_changes = []

    accepted, debug = run_job(
        object(), "Provo", "UT", "therapy", 10, set(), lambda message: None, "api-key",
        on_candidate_progress=lambda done, total: candidate_progress.append((done, total)),
        on_retry_waiting=retry_waiting_changes.append,
    )

    assert ai_order == ["First Clinic", "Second Clinic", "First Clinic"]
    assert {row["Practice name"] for row in accepted} == {"First Clinic", "Second Clinic"}
    assert all(item["Filter result"] != "RETRY: AI Mode chưa hoàn tất" for item in debug)
    assert candidate_progress == [(1, 2), (2, 2)]
    assert retry_waiting_changes == [1, -1]


def test_captcha_wait_can_be_stopped_without_a_timeout(monkeypatch):
    messages = []
    monkeypatch.setattr("app.captcha_is_visible", lambda driver: True)

    assert wait_for_manual_captcha(object(), messages.append, lambda: True) is False
    assert messages == ["Đã dừng trong khi chờ xác minh CAPTCHA."]


def test_debug_callback_receives_each_processed_clinic(monkeypatch):
    row = {"Practice name": "Live Clinic", "location": "Provo", "operation time and days": "N/A"}
    monkeypatch.setattr("app.maps_search_urls", lambda *args, **kwargs: [("https://maps.test/place", "Live Clinic")])
    monkeypatch.setattr("app.extract_maps_place_with_retry", lambda *args, **kwargs: dict(row))
    monkeypatch.setattr("app.google_ai_overview", lambda *args, **kwargs: "")
    updates = []

    run_job(object(), "Provo", "UT", "therapy", 10, set(), lambda message: None, "api-key", on_debug=updates.append)

    assert len(updates) == 1
    assert updates[0]["Practice"] == "Live Clinic"


def test_keep_logic_and_debug_result_agree_for_matching_clinic():
    verification = {"Is_NonProfit_StateOwned": False, "Contains_Red_Flags": False, "Has_Multiple_Therapists": True, "Contains_Target_Services_Licenses": True, "Owner's name": "N/A"}
    assert should_keep_in_final_output(verification) is True
    assert filter_result(verification) == "KEEP"


def test_keep_logic_and_debug_result_agree_without_qualifying_evidence():
    verification = {"Is_NonProfit_StateOwned": False, "Contains_Red_Flags": False, "Has_Multiple_Therapists": False, "Contains_Target_Services_Licenses": False, "Owner's name": "N/A"}
    assert should_keep_in_final_output(verification) is False
    assert filter_result(verification) == "REJECT: insufficient qualifying evidence"


def test_append_rows_preserves_existing_data_and_adds_missing_columns():
    source = io.BytesIO()
    with pd.ExcelWriter(source, engine="openpyxl") as writer:
        pd.DataFrame([{"practice name": "Existing Clinic", "Phone Number": "111"}]).to_excel(writer, sheet_name="Provo, UT", index=False)
    new_rows = pd.DataFrame(
        [["New Clinic", "https://example.com", "222", "Provo", "9:00 AM - 5:00 PM", "4.8", "Jane Doe"]],
        columns=EXPORT_COLUMNS,
    )

    result = append_rows_preserving_template(source.getvalue(), {"Provo, UT": new_rows})
    workbook = load_workbook(io.BytesIO(result))
    sheet = workbook["Provo, UT"]
    headers = [sheet.cell(1, column).value for column in range(1, sheet.max_column + 1)]

    assert sheet.cell(2, 1).value == "Existing Clinic"
    assert sheet.cell(2, 2).value == "111"
    assert all(any(str(value).casefold() == header.casefold() for value in headers) for header in EXPORT_COLUMNS)
    assert sheet.max_row == 3


def test_updater_payload_replaces_installed_updater(tmp_path, monkeypatch):
    updater = tmp_path / "Clinic Lead Updater.exe"
    payload = tmp_path / "Clinic Lead Updater Payload.exe"
    updater.write_bytes(b"old updater")
    payload.write_bytes(b"new updater")
    monkeypatch.setattr("app.installed_app_directory", lambda: tmp_path)
    monkeypatch.setattr("app.sys.frozen", True, raising=False)
    monkeypatch.setattr("app.sys.platform", "win32")

    promote_updater_payload()

    assert updater.read_bytes() == b"new updater"
    assert not payload.exists()
