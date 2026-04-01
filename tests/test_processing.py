from decimal import Decimal
from types import SimpleNamespace
import sys

from app.processing import (
    ADOBE_ADJUSTMENT_LICENSE,
    ADOBE_HOME_OFFICE,
    INTEGRICOM_HOME_OFFICE,
    IntegricomInvoiceLine,
    IntegricomSupportBlock,
    apply_home_office_adjustment,
    build_adobe_user_allocations,
    build_breakdown,
    build_branch_totals,
    build_integricom_support_allocations,
    build_integricom_user_allocations,
    parse_adobe_csv,
    parse_adobe_export_csv,
    parse_integricom_export_csv,
    parse_integricom_invoice,
    parse_csv,
    parse_hexnode_csv,
    summary_to_csv,
)


def test_parse_csv_with_amount_column() -> None:
    raw = b"Branch,License,Amount\nA,Product A,100.50\nA,Product A,99.50\n"
    parsed = parse_csv("sample.csv", raw)

    assert parsed.rows_skipped == 0
    assert len(parsed.rows) == 2

    summary = build_breakdown(parsed.rows)
    assert len(summary) == 1
    assert summary[0]["total_amount"] == 200.0


def test_parse_csv_with_qty_and_unit_price() -> None:
    raw = b"Location,Cost Center,Product,Quantity,Unit Price\nB,HR,Plan X,3,12.00\n"
    parsed = parse_csv("sample2.csv", raw)

    assert parsed.rows_skipped == 0
    assert len(parsed.rows) == 1

    summary = build_breakdown(parsed.rows)
    assert summary[0]["total_amount"] == 36.0


def test_parse_hexnode_csv_maps_default_user_to_home_office() -> None:
    raw = (
        b"Device Name,Username,Department\n"
        b"Phone 1,Default User,\n"
        b"Phone 2,Destin Install,\n"
        b"Phone 3,Acworth,\n"
    )
    parsed = parse_hexnode_csv("hexnode.csv", raw)
    summary = build_breakdown(parsed.rows)

    assert parsed.rows_skipped == 0
    by_branch = {row["branch"]: row["total_amount"] for row in summary}
    assert by_branch["Home Office"] == 2.0
    assert by_branch["Destin Install"] == 2.0
    assert by_branch["Acworth"] == 2.0


def test_home_office_adjustment_adds_to_home_office_line() -> None:
    summary = [
        {
            "branch": "Home Office",
            "license": "Hexnode UEM Cloud Pro Edition",
            "total_amount": 50.0,
        },
        {
            "branch": "Acworth",
            "license": "Hexnode UEM Cloud Pro Edition",
            "total_amount": 10.0,
        },
    ]
    adjusted = apply_home_office_adjustment(summary, Decimal("30.00"))
    by_branch = {row["branch"]: row["total_amount"] for row in adjusted}
    assert by_branch["Home Office"] == 80.0
    assert by_branch["Acworth"] == 10.0


def test_parse_adobe_csv_allocates_by_mapping_and_license_prices() -> None:
    raw = (
        b"Email,First Name,Last Name,Admin Roles,User Groups,Team Products\n"
        b"user1@example.com,User,One,,,Acrobat Pro (DIRECT - ABC)\n"
        b'user2@example.com,User,Two,,,"Creative Cloud Pro (DIRECT - ABC),Photoshop (DIRECT - ABC)"\n'
    )
    per_license = {
        "Acrobat Pro": Decimal("23.99"),
        "Creative Cloud Pro": Decimal("99.99"),
        "Photoshop": Decimal("37.99"),
    }
    mapping = {
        "user1@example.com": {"branch": "Acworth"},
        "user2@example.com": {"branch": "Home Office"},
    }

    parsed_export = parse_adobe_export_csv("users.csv", raw)
    parsed = parse_adobe_csv("users.csv", parsed_export.users, mapping, per_license)
    summary = build_breakdown(parsed.rows)
    by_key = {(row["branch"], row["license"]): row["total_amount"] for row in summary}

    assert parsed.rows_skipped == 0
    assert by_key[("Acworth", "Acrobat Pro")] == 23.99
    assert by_key[("Home Office", "Creative Cloud Pro")] == 99.99


def test_adobe_adjustment_creates_adjustment_license_line() -> None:
    summary = [
        {
            "branch": "Home Office",
            "license": "Acrobat Pro",
            "total_amount": 100.0,
        }
    ]

    adjusted = apply_home_office_adjustment(
        summary,
        Decimal("25.00"),
        license_name=ADOBE_ADJUSTMENT_LICENSE,
        home_office_name=ADOBE_HOME_OFFICE,
    )

    row = [item for item in adjusted if item["license"] == ADOBE_ADJUSTMENT_LICENSE][0]
    assert row["branch"] == "Home Office"
    assert row["total_amount"] == 25.0


def test_summary_to_csv_includes_branch_totals_section() -> None:
    summary = [
        {"branch": "A", "license": "L1", "total_amount": 10.0},
        {"branch": "A", "license": "L2", "total_amount": 15.0},
        {"branch": "B", "license": "L3", "total_amount": 5.0},
    ]

    branch_totals = build_branch_totals(summary)
    by_branch = {row["branch"]: row["total_amount"] for row in branch_totals}
    assert by_branch["A"] == 25.0
    assert by_branch["B"] == 5.0

    csv_text = summary_to_csv(summary)
    assert csv_text.startswith("Branch,Total")
    assert "Branch,License,TotalAmount,BranchTotal" in csv_text
    assert "A,L1,10.0,25.0" in csv_text
    assert "A,25.0" in csv_text
    assert "Grand Total,,30.0" in csv_text


def test_parse_integricom_invoice_applies_credit_to_effective_total(monkeypatch) -> None:
    sample_text = """
Date Invoice
03/01/2026 26757
Products & Other Charges Quantity Price Amount
Sample Service 1.00 $25.00 $25.00
Invoice Subtotal: $25.00
Invoice Total: $25.00
Payments: $0.00
Credits: -$5.00
Balance Due: $20.00
"""

    class FakePage:
        def extract_text(self):
            return sample_text

    class FakePdfReader:
        def __init__(self, _stream):
            self.pages = [FakePage()]

    monkeypatch.setitem(sys.modules, "pypdf", SimpleNamespace(PdfReader=FakePdfReader))

    parsed = parse_integricom_invoice("invoice.pdf", b"%PDF")

    assert parsed.invoice_total == Decimal("20.00")
    assert parsed.credits_total == Decimal("-5.00")
    assert any("applied invoice credits" in warning.lower() for warning in parsed.warnings)


def test_build_adobe_user_allocations_returns_user_rows_and_unresolved() -> None:
    raw = (
        b"Email,First Name,Last Name,Admin Roles,User Groups,Team Products\n"
        b"user1@example.com,User,One,,,Acrobat Pro (DIRECT - ABC)\n"
        b'user2@example.com,User,Two,,,"Creative Cloud Pro (DIRECT - ABC),Photoshop (DIRECT - ABC)"\n'
    )
    per_license = {
        "Acrobat Pro": Decimal("23.99"),
        "Creative Cloud Pro": Decimal("99.99"),
        "Photoshop": Decimal("37.99"),
    }
    mapping = {
        "user1@example.com": {"branch": "Acworth"},
        "user2@example.com": {"branch": ""},
    }

    parsed_export = parse_adobe_export_csv("users.csv", raw)
    line_rows, user_rows, warnings, unresolved = build_adobe_user_allocations(
        parsed_export.users,
        mapping,
        per_license,
    )

    assert not warnings
    assert unresolved == ["user2@example.com"]
    assert len(line_rows) == 1
    assert line_rows[0]["branch"] == "Acworth"
    by_email = {row["email"]: row for row in user_rows}
    assert by_email["user1@example.com"]["license_list"] == "Acrobat Pro"
    assert by_email["user1@example.com"]["user_total"] == 23.99
    assert by_email["user2@example.com"]["license_list"] == "Creative Cloud Pro, Photoshop"
    assert by_email["user2@example.com"]["user_total"] == 137.98


def test_parse_integricom_export_csv_uses_office_and_filters_unlicensed() -> None:
    raw = (
        b"Display name,User principal name,First name,Last name,Department,Office,Licenses\n"
        b"User One,user1@example.com,User,One,Maintenance,Acworth,Microsoft 365 Business Premium\n"
        b"User Two,user2@example.com,User,Two,Office,Corporate,Exchange Online (Plan 1)\n"
        b"User Three,user3@example.com,User,Three,Office,Home Office,Unlicensed\n"
        b"External,user4#EXT#@example.com,User,Four,Office,Tampa,Microsoft 365 Business Premium\n"
    )
    parsed = parse_integricom_export_csv("m365.csv", raw)

    assert parsed.rows_skipped == 2
    assert len(parsed.users) == 2
    by_email = {user.email: user for user in parsed.users}
    assert by_email["user1@example.com"].default_branch == "Acworth"
    assert by_email["user2@example.com"].default_branch == "Home Office"


def test_parse_integricom_export_csv_construction_department_overrides_office() -> None:
    raw = (
        b"Display name,User principal name,First name,Last name,Department,Office,Licenses\n"
        b"User One,user1@example.com,User,One,Construction,Doraville,Microsoft 365 Business Premium\n"
        b"User Two,user2@example.com,User,Two,Construction Operations,Corporate,Exchange Online (Plan 1)\n"
    )

    parsed = parse_integricom_export_csv("m365.csv", raw)

    assert len(parsed.users) == 2
    assert parsed.users[0].default_branch == "Construction"
    assert parsed.users[1].default_branch == "Construction"


def test_build_integricom_user_allocations_applies_dynamic_and_home_office_remainder() -> None:
    raw = (
        b"Display name,User principal name,First name,Last name,Department,Office,Licenses\n"
        b"User One,user1@example.com,User,One,Maintenance,Acworth,Microsoft 365 Business Premium\n"
        b"User Two,user2@example.com,User,Two,Maintenance,Destin,Exchange Online (Plan 1)\n"
        b"User Three,user3@example.com,User,Three,Maintenance,Corporate,Microsoft 365 F3\n"
    )
    users = parse_integricom_export_csv("m365.csv", raw).users
    invoice_lines = [
        IntegricomInvoiceLine(
            description="Workstation",
            canonical_name="Workstation",
            quantity=Decimal("4.00"),
            unit_price=Decimal("25.00"),
            amount=Decimal("100.00"),
        ),
        IntegricomInvoiceLine(
            description="Office 365 Cloud Backup",
            canonical_name="Office 365 Cloud Backup",
            quantity=Decimal("1.00"),
            unit_price=Decimal("3.00"),
            amount=Decimal("3.00"),
        ),
        IntegricomInvoiceLine(
            description="Dropbox Business Standard",
            canonical_name="Dropbox Business Standard",
            quantity=Decimal("10.00"),
            unit_price=Decimal("3.00"),
            amount=Decimal("30.00"),
        ),
    ]

    rows, user_rows, non_user_rows, warnings, unresolved, unresolved_branch_prompts = build_integricom_user_allocations(
        users,
        {},
        invoice_lines,
    )
    summary = build_breakdown(rows)
    by_key = {(row["branch"], row["license"]): row["total_amount"] for row in summary}
    by_email = {row["email"]: row for row in user_rows}

    assert unresolved == []
    assert unresolved_branch_prompts == []
    assert any("Workstation" in warning for warning in warnings)
    assert any("Office 365 Cloud Backup" in warning for warning in warnings)
    assert by_key[("Acworth", "Workstation")] == 25.0
    assert by_key[("Destin", "Workstation")] == 25.0
    assert by_key[(INTEGRICOM_HOME_OFFICE, "Workstation")] == 50.0
    assert by_key[(INTEGRICOM_HOME_OFFICE, "Dropbox Business Standard")] == 30.0
    assert by_key[(INTEGRICOM_HOME_OFFICE, "Office 365 Cloud Backup")] == -3.0
    assert by_email["user1@example.com"]["user_total"] == 28.0
    assert by_email["user2@example.com"]["user_total"] == 28.0
    assert by_email["user3@example.com"]["user_total"] == 25.0
    non_user_lookup = {
        (row["branch"], row["license"], row["allocation_type"]): row["total_amount"] for row in non_user_rows
    }
    assert non_user_lookup[(INTEGRICOM_HOME_OFFICE, "Dropbox Business Standard", "Fixed Branch Item")] == 30.0
    assert non_user_lookup[(INTEGRICOM_HOME_OFFICE, "Workstation", "Invoice Delta")] == 25.0


def test_build_integricom_user_allocations_prefers_construction_over_saved_branch() -> None:
    raw = (
        b"Display name,User principal name,First name,Last name,Department,Office,Licenses\n"
        b"User One,user1@example.com,User,One,Construction,Doraville,Microsoft 365 Business Premium\n"
    )
    users = parse_integricom_export_csv("m365.csv", raw).users
    invoice_lines = [
        IntegricomInvoiceLine(
            description="Workstation",
            canonical_name="Workstation",
            quantity=Decimal("1.00"),
            unit_price=Decimal("25.00"),
            amount=Decimal("25.00"),
        ),
    ]

    rows, user_rows, _non_user_rows, _warnings, _unresolved, _prompts = build_integricom_user_allocations(
        users,
        {
            "user1@example.com": {
                "branch": "Doraville",
                "first_name": "User",
                "last_name": "One",
            }
        },
        invoice_lines,
    )

    assert rows[0]["branch"] == "Construction"
    assert user_rows[0]["branch"] == "Construction"


def test_integricom_branch_tethered_extra_quantity_requires_assignment_prompt() -> None:
    invoice_lines = [
        IntegricomInvoiceLine(
            description="Managed firewall",
            canonical_name="NetWatch360 Managed Firewall",
            quantity=Decimal("14.00"),
            unit_price=Decimal("10.00"),
            amount=Decimal("140.00"),
        )
    ]

    rows, _user_rows, _non_user_rows, warnings, _unresolved, unresolved_branch_prompts = build_integricom_user_allocations(
        [],
        {},
        invoice_lines,
    )
    summary = build_breakdown(rows)
    by_key = {(row["branch"], row["license"]): row["total_amount"] for row in summary}

    assert unresolved_branch_prompts
    assert unresolved_branch_prompts[0]["license"] == "NetWatch360 Managed Firewall"
    assert unresolved_branch_prompts[0]["line_key"].startswith("1:NetWatch360 Managed Firewall")
    assert any("extra branch assignment" in warning.lower() for warning in warnings)
    assert (INTEGRICOM_HOME_OFFICE, "NetWatch360 Managed Firewall") not in by_key

    updates = [
        {
            "line_key": unresolved_branch_prompts[0]["line_key"],
            "prompt_index": unresolved_branch_prompts[0]["prompt_index"],
            "branch": "Sugar Hill",
        }
    ]

    rows_after, _u2, _n2, _w2, _e2, unresolved_after = build_integricom_user_allocations(
        [],
        {},
        invoice_lines,
        branch_item_updates=updates,
    )
    summary_after = build_breakdown(rows_after)
    by_key_after = {(row["branch"], row["license"]): row["total_amount"] for row in summary_after}

    assert unresolved_after == []
    assert by_key_after[("Sugar Hill", "NetWatch360 Managed Firewall")] == 10.0
    assert (INTEGRICOM_HOME_OFFICE, "NetWatch360 Managed Firewall") not in by_key_after


def test_integricom_support_allocations_flag_low_confidence_rows_for_review() -> None:
    blocks = [
        IntegricomSupportBlock(
            row_key="1:abc",
            charge_summary="FW - CP1530 - Savannah - packet loss",
            billable_entries=2,
            billable_hours=Decimal("0.75"),
            amount=Decimal("123.75"),
        ),
        IntegricomSupportBlock(
            row_key="2:def",
            charge_summary="Action required: Review admin consent request",
            billable_entries=1,
            billable_hours=Decimal("0.50"),
            amount=Decimal("82.50"),
        ),
    ]

    rows, support_rows, warnings = build_integricom_support_allocations(blocks)
    summary = build_breakdown(rows)
    by_key = {(row["branch"], row["license"]): row["total_amount"] for row in summary}
    by_row = {row["row_key"]: row for row in support_rows}

    assert by_key[("Savannah", "Support: FW - CP1530 - Savannah - packet loss")] == 123.75
    assert by_key[(INTEGRICOM_HOME_OFFICE, "Support: Action required: Review admin consent request")] == 82.5
    assert by_row["1:abc"]["needs_review"] is False
    assert by_row["1:abc"]["confidence"] == "high"
    assert by_row["2:def"]["needs_review"] is True
    assert by_row["2:def"]["confidence"] == "low"
    assert any("defaulted to Home Office" in warning for warning in warnings)


def test_integricom_support_user_override_clears_review_flag() -> None:
    blocks = [
        IntegricomSupportBlock(
            row_key="2:def",
            charge_summary="Action required: Review admin consent request",
            billable_entries=1,
            billable_hours=Decimal("0.50"),
            amount=Decimal("82.50"),
        )
    ]

    rows, support_rows, warnings = build_integricom_support_allocations(
        blocks,
        support_updates=[{"row_key": "2:def", "branch": "Acworth"}],
    )
    summary = build_breakdown(rows)
    by_key = {(row["branch"], row["license"]): row["total_amount"] for row in summary}

    assert by_key[("Acworth", "Support: Action required: Review admin consent request")] == 82.5
    assert support_rows[0]["needs_review"] is False
    assert support_rows[0]["confidence"] == "user"
    assert not warnings
