from app import extract_operation_time_from_text, extract_owner_and_time_from_pages, should_keep_in_final_output


def test_extract_operation_time_from_relevant_sentence():
    page_text = "Monday-Friday 9:00 AM - 5:00 PM. We offer counseling services."
    assert extract_operation_time_from_text(page_text) == "Monday-Friday 9:00 AM - 5:00 PM"


def test_extract_owner_and_time_from_pages_without_gemini():
    page_infos = [
        {"title": "About us", "text": "Owner: Jane Doe leads this counseling clinic."},
    ]
    result = extract_owner_and_time_from_pages(page_infos, api_key="")
    assert result["Owner's name"] == "Jane Doe"
    assert result["operation time and days"] == "N/A"


def test_should_keep_in_final_output_allows_matching_clinics():
    row = {"Practice name": "Bright Path Counseling", "website link": "https://example.com", "phone number": "123"}
    verification = {"Is_NonProfit_StateOwned": False, "Contains_Red_Flags": False, "Has_Multiple_Therapists": False, "Contains_Target_Services_Licenses": False, "Has_Team_Page": False}
    assert should_keep_in_final_output(row, verification) is True
