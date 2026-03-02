from app.spreadsheet_import import parse_adobe_directory_import_file


def test_parse_adobe_directory_import_csv_uses_office_column() -> None:
    raw = (
        "Email,First Name,Last Name,Office\n"
        "user1@example.com,Amy,Adams,Home Office\n"
        "USER2@example.com,Bob,Baker,Sugar Hill\n"
        "user2@example.com,Bob,Baker,Doraville\n"
    ).encode("utf-8")

    result = parse_adobe_directory_import_file("Adobe Cost Calc.csv", raw)

    assert result.source == "csv"
    assert len(result.rows) == 2
    assert result.rows[0]["email"] == "user1@example.com"
    assert result.rows[0]["branch"] == "Home Office"
    assert result.rows[1]["email"] == "user2@example.com"
    assert result.rows[1]["branch"] == "Doraville"


def test_parse_adobe_directory_import_csv_defaults_blank_branch() -> None:
    raw = "Email,First Name,Last Name,Office\nuser3@example.com,Chris,Cole,\n".encode("utf-8")

    result = parse_adobe_directory_import_file("users.csv", raw)

    assert len(result.rows) == 1
    assert result.rows[0]["branch"] == "Home Office"


def test_parse_adobe_directory_import_rejects_unsupported_file() -> None:
    result = parse_adobe_directory_import_file("users.txt", b"Email,Branch")

    assert result.rows == []
    assert result.warnings
