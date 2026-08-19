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
    monkeypatch.setattr("app.google_ai_overview", lambda *args, **kwargs: "")

    def unexpected_gemini_call(*args, **kwargs):
        raise AssertionError("Gemini must not be called without complete AI Mode evidence")

    monkeypatch.setattr("app.gemini_metadata", unexpected_gemini_call)
    known = set()
    accepted, debug = run_job(object(), "Provo", "UT", "therapy", 10, known, lambda message: None, "api-key")

    assert accepted == []
    assert debug[0]["Filter result"] == "RETRY: AI Mode chưa hoàn tất"
    assert debug[0]["Gemini"] == "not called"
    assert ("incomplete clinic", "provo") not in known


def test_keep_logic_and_debug_result_agree_for_matching_clinic():
    row = {"Practice name": "Bright Path Counseling", "website link": "https://example.com", "phone number": "123"}
    verification = {"Is_NonProfit_StateOwned": False, "Contains_Red_Flags": False, "Has_Multiple_Therapists": True, "Contains_Target_Services_Licenses": True, "Owner's name": "N/A"}
    assert should_keep_in_final_output(row, verification) is True
    assert filter_result(verification) == "KEEP"


def test_keep_logic_and_debug_result_agree_without_qualifying_evidence():
    verification = {"Is_NonProfit_StateOwned": False, "Contains_Red_Flags": False, "Has_Multiple_Therapists": False, "Contains_Target_Services_Licenses": False, "Owner's name": "N/A"}
    assert should_keep_in_final_output({}, verification) is False
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
