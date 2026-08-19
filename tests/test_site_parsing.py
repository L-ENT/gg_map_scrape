import io

import pandas as pd
from openpyxl import load_workbook

from app import (
    EXPORT_COLUMNS,
    append_rows_preserving_template,
    extract_operation_time_from_text,
    filter_result,
    promote_updater_payload,
    should_keep_in_final_output,
)


def test_extract_operation_time_from_relevant_sentence():
    page_text = "Monday-Friday 9:00 AM - 5:00 PM. We offer counseling services."
    assert extract_operation_time_from_text(page_text) == "Monday-Friday 9:00 AM - 5:00 PM"


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
