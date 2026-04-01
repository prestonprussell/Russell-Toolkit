from __future__ import annotations

import csv
import hashlib
import io
import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any


HEXNODE_DEFAULT_COST = Decimal("2.00")
HEXNODE_DEFAULT_LICENSE = "Hexnode UEM Cloud Pro Edition"
HEXNODE_HOME_OFFICE = "Home Office"
HEXNODE_BRANCH_ALIASES: dict[str, str] = {
    # User-confirmed Hexnode remap rule.
    "Default User": "Home Office",
}

ADOBE_HOME_OFFICE = "Home Office"
ADOBE_ADJUSTMENT_LICENSE = "Adobe Invoice Adjustment"
ADOBE_PRODUCT_ALIASES: dict[str, str] = {
    "acrobat pro": "Acrobat Pro",
    "acrobat pro dc": "Acrobat Pro",
    "creative cloud pro": "Creative Cloud Pro",
    "creative cloud all apps": "Creative Cloud Pro",
    "creative cloud all apps - pro edition": "Creative Cloud Pro",
    "indesign": "InDesign",
    "indesign - pro edition": "InDesign",
    "illustrator": "Illustrator",
    "lightroom": "Lightroom",
    "lightroom single app plan with 1tb": "Lightroom",
    "photoshop": "Photoshop",
    "photoshop - pro edition": "Photoshop",
    "adobe stock - 40 assets a month": "Adobe Stock - 40 assets a month",
    "adobe stock – 40 assets a month": "Adobe Stock - 40 assets a month",
    "ai assistant for acrobat": "AI Assistant for Acrobat",
}

INTEGRICOM_HOME_OFFICE = "Home Office"
INTEGRICOM_ADJUSTMENT_LICENSE = "Integricom Invoice Adjustment"
INTEGRICOM_CREDIT_LICENSE = "Integricom Invoice Credit"
INTEGRICOM_BRANCH_ALIASES: dict[str, str] = {
    "": INTEGRICOM_HOME_OFFICE,
    "Corporate": INTEGRICOM_HOME_OFFICE,
    "Process Smart": INTEGRICOM_HOME_OFFICE,
}
INTEGRICOM_DISTRICT_BRANCHES: list[str] = [
    "Acworth",
    "Canton",
    "Charleston",
    "Cobb",
    "Color Burst",
    "Doraville",
    "Destin",
    "Fort Walton",
    "Pensacola",
    "Nashville",
    "Savannah",
    "St. Pete",
    "Tampa",
]
INTEGRICOM_MANAGED_INTERNET_BRANCHES: list[str] = [INTEGRICOM_HOME_OFFICE, *INTEGRICOM_DISTRICT_BRANCHES]
INTEGRICOM_KNOWN_BRANCHES: list[str] = [
    INTEGRICOM_HOME_OFFICE,
    *INTEGRICOM_DISTRICT_BRANCHES,
    "Construction",
    "Sugar Hill",
    "Grayson",
]

INTEGRICOM_LICENSE_BP = "Microsoft 365 Business Premium"
INTEGRICOM_LICENSE_P1 = "Exchange Online (Plan 1)"
INTEGRICOM_LICENSE_P2 = "Exchange Online (Plan 2)"
INTEGRICOM_LICENSE_F3 = "Microsoft 365 F3"
INTEGRICOM_LICENSE_TEAMS_ESSENTIALS = "Microsoft Teams Essentials"


HEADER_ALIASES: dict[str, list[str]] = {
    "branch": [
        "branch",
        "branch_name",
        "location",
        "site",
        "office",
        "store",
        "entity",
        "branch_code",
    ],
    "license": [
        "license",
        "licence",
        "product",
        "product_name",
        "subscription",
        "service",
        "sku",
        "plan",
        "item",
        "name",
    ],
    "amount": [
        "amount",
        "charge",
        "cost",
        "extended_cost",
        "line_total",
        "total",
        "price",
        "net_amount",
        "subtotal",
    ],
    "quantity": ["qty", "quantity", "seats", "licenses", "units", "count"],
    "unit_price": ["unit_price", "price_per_unit", "rate", "unit_cost", "cost_per_license"],
}


@dataclass
class ParseResult:
    filename: str
    rows: list[dict[str, Any]]
    rows_skipped: int
    warnings: list[str]


@dataclass
class InvoiceParseResult:
    filename: str
    invoice_number: str | None
    invoice_total: Decimal | None
    billed_device_count: int | None
    warnings: list[str]


@dataclass
class AdobeInvoiceParseResult:
    filename: str
    invoice_number: str | None
    invoice_total: Decimal | None
    per_license_cost: dict[str, Decimal]
    warnings: list[str]


@dataclass
class AdobeExportUser:
    source_file: str
    email: str
    first_name: str
    last_name: str
    products: list[str]


@dataclass
class AdobeExportParseResult:
    filename: str
    users: list[AdobeExportUser]
    rows_skipped: int
    warnings: list[str]


@dataclass
class IntegricomInvoiceLine:
    description: str
    canonical_name: str
    quantity: Decimal
    unit_price: Decimal
    amount: Decimal


@dataclass
class IntegricomInvoiceParseResult:
    filename: str
    invoice_number: str | None
    invoice_total: Decimal | None
    credits_total: Decimal
    line_items: list[IntegricomInvoiceLine]
    warnings: list[str]


@dataclass
class IntegricomExportUser:
    source_file: str
    email: str
    first_name: str
    last_name: str
    office: str
    default_branch: str
    licenses: list[str]


@dataclass
class IntegricomExportParseResult:
    filename: str
    users: list[IntegricomExportUser]
    rows_skipped: int
    warnings: list[str]


@dataclass
class IntegricomSupportBlock:
    row_key: str
    charge_summary: str
    billable_entries: int
    billable_hours: Decimal
    amount: Decimal


@dataclass
class IntegricomSupportInvoiceParseResult:
    filename: str
    invoice_number: str | None
    invoice_total: Decimal | None
    blocks: list[IntegricomSupportBlock]
    warnings: list[str]


def _normalize_header(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def _match_header(headers: list[str], aliases: list[str]) -> str | None:
    normalized = {_normalize_header(h): h for h in headers}
    for alias in aliases:
        if alias in normalized:
            return normalized[alias]

    for norm, original in normalized.items():
        for alias in aliases:
            if alias in norm or norm in alias:
                return original
    return None


def _decode_bytes(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _parse_decimal(value: str | None) -> Decimal | None:
    if value is None:
        return None
    cleaned = value.strip()
    if not cleaned:
        return None

    negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negative = True
        cleaned = cleaned[1:-1]

    cleaned = cleaned.replace("$", "").replace(",", "").strip()
    if cleaned.startswith("-"):
        negative = True
        cleaned = cleaned[1:]

    try:
        result = Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None

    return -result if negative else result


def _normalize_product_name(value: str) -> str:
    cleaned = value.strip()
    cleaned = re.sub(r"\s*\(DIRECT[^)]*\)\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("–", "-")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip().lower()


def _canonical_adobe_product(value: str) -> str | None:
    normalized = _normalize_product_name(value)
    if not normalized:
        return None

    if normalized in ADOBE_PRODUCT_ALIASES:
        return ADOBE_PRODUCT_ALIASES[normalized]

    # Fallback fuzzy contains checks for export variants.
    if "acrobat" in normalized and "assistant" not in normalized:
        return "Acrobat Pro"
    if "creative cloud" in normalized:
        return "Creative Cloud Pro"
    if "indesign" in normalized:
        return "InDesign"
    if "illustrator" in normalized:
        return "Illustrator"
    if "lightroom" in normalized:
        return "Lightroom"
    if "photoshop" in normalized:
        return "Photoshop"
    if "adobe stock" in normalized and "40 assets" in normalized:
        return "Adobe Stock - 40 assets a month"
    if "ai assistant for acrobat" in normalized:
        return "AI Assistant for Acrobat"

    return None


def parse_csv(filename: str, raw: bytes) -> ParseResult:
    text = _decode_bytes(raw)
    reader = csv.DictReader(io.StringIO(text))

    if reader.fieldnames is None:
        return ParseResult(
            filename=filename,
            rows=[],
            rows_skipped=0,
            warnings=[f"{filename}: no headers found; file skipped."],
        )

    headers = [h for h in reader.fieldnames if h is not None]
    branch_col = _match_header(headers, HEADER_ALIASES["branch"])
    license_col = _match_header(headers, HEADER_ALIASES["license"])
    amount_col = _match_header(headers, HEADER_ALIASES["amount"])
    qty_col = _match_header(headers, HEADER_ALIASES["quantity"])
    unit_col = _match_header(headers, HEADER_ALIASES["unit_price"])

    warnings: list[str] = []
    if license_col is None:
        warnings.append(f"{filename}: could not confidently identify a license/product column.")
    if amount_col is None and (qty_col is None or unit_col is None):
        warnings.append(
            f"{filename}: no amount column and no quantity+unit price pair found; rows may be skipped."
        )

    parsed_rows: list[dict[str, Any]] = []
    rows_skipped = 0

    for line_number, row in enumerate(reader, start=2):
        license_name = (row.get(license_col, "") if license_col else "").strip() if license_col else ""
        branch = (row.get(branch_col, "") if branch_col else "").strip() if branch_col else ""
        amount = _parse_decimal(row.get(amount_col)) if amount_col else None
        if amount is None and qty_col and unit_col:
            qty = _parse_decimal(row.get(qty_col))
            unit_price = _parse_decimal(row.get(unit_col))
            if qty is not None and unit_price is not None:
                amount = qty * unit_price

        if amount is None:
            rows_skipped += 1
            warnings.append(f"{filename}: row {line_number} skipped (amount missing or invalid).")
            continue

        parsed_rows.append(
            {
                "source_file": filename,
                "branch": branch or "UNMAPPED_BRANCH",
                "license": license_name or "UNMAPPED_LICENSE",
                "amount": amount,
            }
        )

    return ParseResult(
        filename=filename,
        rows=parsed_rows,
        rows_skipped=rows_skipped,
        warnings=warnings,
    )


def build_breakdown(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    for row in rows:
        key = (row["branch"], row["license"])
        grouped[key] += row["amount"]

    summary: list[dict[str, Any]] = []
    for (branch, license_name), total in sorted(grouped.items()):
        summary.append(
            {
                "branch": branch,
                "license": license_name,
                "total_amount": float(total.quantize(Decimal("0.01"))),
            }
        )
    return summary


def build_branch_totals(summary: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    for row in summary:
        grouped[row["branch"]] += Decimal(str(row["total_amount"]))

    totals: list[dict[str, Any]] = []
    for branch, total in sorted(grouped.items()):
        totals.append(
            {
                "branch": branch,
                "total_amount": float(total.quantize(Decimal("0.01"))),
            }
        )
    return totals


def summary_to_csv(summary: list[dict[str, Any]]) -> str:
    branch_totals = build_branch_totals(summary)
    branch_lookup = {row["branch"]: row["total_amount"] for row in branch_totals}

    buffer = io.StringIO()
    writer = csv.writer(buffer)
    # Put branch-level pivot totals first so exports open with the allocation rollup.
    writer.writerow(["Branch", "Total"])
    for row in branch_totals:
        writer.writerow([row["branch"], row["total_amount"]])

    grand_total = round(sum(row["total_amount"] for row in branch_totals), 2)
    writer.writerow(["Grand Total", "", grand_total])
    writer.writerow([])

    writer.writerow(["Branch", "License", "TotalAmount", "BranchTotal"])
    for row in summary:
        branch_total = branch_lookup.get(row["branch"], row["total_amount"])
        writer.writerow([row["branch"], row["license"], row["total_amount"], branch_total])
    return buffer.getvalue()


def _money_from_text(text: str, patterns: list[str]) -> Decimal | None:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if not match:
            continue
        value = _parse_decimal(match.group(1))
        if value is not None:
            return value.quantize(Decimal("0.01"))
    return None


def parse_hexnode_invoice(filename: str, raw: bytes) -> InvoiceParseResult:
    warnings: list[str] = []
    invoice_number: str | None = None
    invoice_total: Decimal | None = None
    billed_device_count: int | None = None

    try:
        from pypdf import PdfReader  # Imported lazily for environments that skip invoice parsing.
    except Exception:
        return InvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            billed_device_count=None,
            warnings=["Invoice parser unavailable (pypdf not installed)."],
        )

    try:
        reader = PdfReader(io.BytesIO(raw))
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        return InvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            billed_device_count=None,
            warnings=[f"Could not parse PDF text from {filename}: {exc}"],
        )

    invoice_match = re.search(r"Invoice:\s*#?\s*([A-Z0-9-]+)", text, flags=re.IGNORECASE)
    if invoice_match:
        invoice_number = invoice_match.group(1).strip()

    invoice_total = _money_from_text(
        text,
        patterns=[
            r"Total amount payable after discounts\s*\$?\s*([0-9][0-9,]*\.\d{2})",
            r"Amount Paid\s*\$?\s*([0-9][0-9,]*\.\d{2})",
            r"Sub Total\s*\$?\s*([0-9][0-9,]*\.\d{2})",
        ],
    )
    if invoice_total is None:
        warnings.append(f"{filename}: could not extract invoice total.")

    count_match = re.search(r"Total device count:\s*([0-9]+)", text, flags=re.IGNORECASE)
    if count_match:
        billed_device_count = int(count_match.group(1))

    return InvoiceParseResult(
        filename=filename,
        invoice_number=invoice_number,
        invoice_total=invoice_total,
        billed_device_count=billed_device_count,
        warnings=warnings,
    )


def parse_hexnode_csv(
    filename: str,
    raw: bytes,
    per_device_cost: Decimal = HEXNODE_DEFAULT_COST,
    branch_aliases: dict[str, str] | None = None,
) -> ParseResult:
    text = _decode_bytes(raw)
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        return ParseResult(
            filename=filename,
            rows=[],
            rows_skipped=0,
            warnings=[f"{filename}: no headers found; file skipped."],
        )

    headers = [h for h in reader.fieldnames if h is not None]
    username_col = _match_header(headers, ["username", "branch", "location", "site"])
    if username_col is None:
        return ParseResult(
            filename=filename,
            rows=[],
            rows_skipped=0,
            warnings=[f"{filename}: could not find Username/Branch column for Hexnode export."],
        )

    aliases = {**HEXNODE_BRANCH_ALIASES}
    if branch_aliases:
        aliases.update(branch_aliases)

    rows: list[dict[str, Any]] = []
    rows_skipped = 0
    warnings: list[str] = []

    for line_number, row in enumerate(reader, start=2):
        raw_branch = (row.get(username_col) or "").strip()
        if not raw_branch:
            rows_skipped += 1
            warnings.append(f"{filename}: row {line_number} skipped (blank Username/Branch).")
            continue

        mapped_branch = aliases.get(raw_branch, raw_branch)
        rows.append(
            {
                "source_file": filename,
                "branch": mapped_branch,
                "license": HEXNODE_DEFAULT_LICENSE,
                "amount": per_device_cost,
            }
        )

    return ParseResult(filename=filename, rows=rows, rows_skipped=rows_skipped, warnings=warnings)


def parse_adobe_invoice(filename: str, raw: bytes) -> AdobeInvoiceParseResult:
    warnings: list[str] = []
    invoice_number: str | None = None
    invoice_total: Decimal | None = None
    per_license_cost: dict[str, Decimal] = {}

    try:
        from pypdf import PdfReader
    except Exception:
        return AdobeInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            per_license_cost={},
            warnings=["Invoice parser unavailable (pypdf not installed)."],
        )

    try:
        reader = PdfReader(io.BytesIO(raw))
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        return AdobeInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            per_license_cost={},
            warnings=[f"Could not parse PDF text from {filename}: {exc}"],
        )

    invoice_match = re.search(r"([0-9]{6,})\s*Invoice Number", text, flags=re.IGNORECASE)
    if not invoice_match:
        invoice_match = re.search(r"Invoice Number\s*([0-9]{6,})", text, flags=re.IGNORECASE)
    if invoice_match:
        invoice_number = invoice_match.group(1).strip()

    total_match = re.search(r"GRAND TOTAL \(USD\)\s*([0-9][0-9,]*\.\d{2})", text, flags=re.IGNORECASE)
    if total_match:
        invoice_total = _parse_decimal(total_match.group(1))
    if invoice_total is None:
        warnings.append(f"{filename}: could not extract Adobe invoice grand total.")

    product_patterns = {
        "Illustrator": r"Illustrator\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "Acrobat Pro": r"Acrobat Pro\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "Creative Cloud Pro": r"Creative Cloud Pro\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "InDesign": r"InDesign\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "Lightroom": r"Lightroom\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "Photoshop": r"Photoshop\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "Adobe Stock - 40 assets a month": r"Adobe Stock\s*[–-]\s*40 assets a month\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
        "AI Assistant for Acrobat": r"AI Assistant for Acrobat\s+([0-9]+)\s+EA\s+[0-9,]+\.\d{2}\s+[0-9,]+\.\d{2}\s+[0-9.]+%\s+[0-9,]+\.\d{2}\s+([0-9,]+\.\d{2})",
    }

    for product, pattern in product_patterns.items():
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        qty = _parse_decimal(match.group(1))
        total = _parse_decimal(match.group(2))
        if qty is None or total is None or qty == Decimal("0"):
            continue
        per_license_cost[product] = (total / qty).quantize(Decimal("0.01"))

    if not per_license_cost:
        warnings.append(f"{filename}: no Adobe line-item pricing could be extracted.")

    return AdobeInvoiceParseResult(
        filename=filename,
        invoice_number=invoice_number,
        invoice_total=invoice_total,
        per_license_cost=per_license_cost,
        warnings=warnings,
    )


def parse_adobe_export_csv(
    filename: str,
    raw: bytes,
) -> AdobeExportParseResult:
    text = _decode_bytes(raw)
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        return AdobeExportParseResult(
            filename=filename,
            users=[],
            rows_skipped=0,
            warnings=[f"{filename}: no headers found; file skipped."],
        )

    headers = [h for h in reader.fieldnames if h is not None]
    email_col = _match_header(headers, ["email", "user_email"])
    first_name_col = _match_header(headers, ["first_name", "first", "given_name"])
    last_name_col = _match_header(headers, ["last_name", "last", "surname", "family_name"])
    team_products_col = _match_header(headers, ["team_products", "products", "product", "licenses", "license"])

    if email_col is None or team_products_col is None:
        return AdobeExportParseResult(
            filename=filename,
            users=[],
            rows_skipped=0,
            warnings=[f"{filename}: expected Adobe export columns (Email, Team Products) were not found."],
        )

    users: list[AdobeExportUser] = []
    rows_skipped = 0
    warnings: list[str] = []

    for line_number, row in enumerate(reader, start=2):
        email = (row.get(email_col) or "").strip().lower()
        if not email:
            rows_skipped += 1
            warnings.append(f"{filename}: row {line_number} skipped (missing email).")
            continue

        raw_products = (row.get(team_products_col) or "").strip()
        product_tokens = [token.strip() for token in raw_products.split(",") if token.strip()]
        users.append(
            AdobeExportUser(
                source_file=filename,
                email=email,
                first_name=(row.get(first_name_col) or "").strip() if first_name_col else "",
                last_name=(row.get(last_name_col) or "").strip() if last_name_col else "",
                products=product_tokens,
            )
        )

    return AdobeExportParseResult(filename=filename, users=users, rows_skipped=rows_skipped, warnings=warnings)


def parse_adobe_csv(
    filename: str,
    users: list[AdobeExportUser],
    user_directory: dict[str, dict[str, str]],
    per_license_cost: dict[str, Decimal],
) -> ParseResult:
    rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    warned_missing_user: set[str] = set()
    warned_unknown_product: set[str] = set()

    for user in users:
        if not user.products:
            continue

        profile = user_directory.get(user.email)
        if profile is None:
            if user.email not in warned_missing_user:
                warnings.append(
                    f"{filename}: {user.email} is not in the Adobe directory yet; user was skipped."
                )
                warned_missing_user.add(user.email)
            continue

        branch = (profile.get("branch") or "").strip() or ADOBE_HOME_OFFICE
        for token in user.products:
            canonical = _canonical_adobe_product(token)
            if canonical is None:
                if token not in warned_unknown_product:
                    warnings.append(f"{filename}: unrecognized Adobe product '{token}' skipped.")
                    warned_unknown_product.add(token)
                continue

            cost = per_license_cost.get(canonical)
            if cost is None:
                if canonical not in warned_unknown_product:
                    warnings.append(f"{filename}: no invoice price found for '{canonical}'; charges skipped.")
                    warned_unknown_product.add(canonical)
                continue

            rows.append(
                {
                    "source_file": filename,
                    "branch": branch,
                    "license": canonical,
                    "amount": cost,
                }
            )

    return ParseResult(filename=filename, rows=rows, rows_skipped=0, warnings=warnings)


def build_adobe_user_allocations(
    users: list[AdobeExportUser],
    user_directory: dict[str, dict[str, str]],
    per_license_cost: dict[str, Decimal],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str], list[str]]:
    line_rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    unresolved_emails: list[str] = []
    warned_unknown_product: set[str] = set()
    warned_missing_cost: set[str] = set()
    user_rows_map: dict[str, dict[str, Any]] = {}

    for user in users:
        email = user.email.strip().lower()
        if not email:
            continue

        profile = user_directory.get(email)
        branch = (profile.get("branch") if profile else "") or ""
        first_name = (user.first_name or "").strip()
        last_name = (user.last_name or "").strip()
        if profile:
            first_name = first_name or (profile.get("first_name") or "").strip()
            last_name = last_name or (profile.get("last_name") or "").strip()

        if email not in user_rows_map:
            user_rows_map[email] = {
                "email": email,
                "first_name": first_name,
                "last_name": last_name,
                "branch": branch,
                "licenses": [],
                "user_total": Decimal("0"),
                "known_user": bool(branch),
            }

        row_entry = user_rows_map[email]
        if branch and not row_entry["branch"]:
            row_entry["branch"] = branch
            row_entry["known_user"] = True
        if first_name and not row_entry["first_name"]:
            row_entry["first_name"] = first_name
        if last_name and not row_entry["last_name"]:
            row_entry["last_name"] = last_name

        for token in user.products:
            canonical = _canonical_adobe_product(token)
            if canonical is None:
                if token not in warned_unknown_product:
                    warnings.append(f"Unrecognized Adobe product '{token}' skipped.")
                    warned_unknown_product.add(token)
                continue

            if canonical not in row_entry["licenses"]:
                row_entry["licenses"].append(canonical)

            cost = per_license_cost.get(canonical)
            if cost is None:
                if canonical not in warned_missing_cost:
                    warnings.append(f"No invoice price found for '{canonical}'; charges skipped.")
                    warned_missing_cost.add(canonical)
                continue

            row_entry["user_total"] += cost
            if row_entry["branch"]:
                line_rows.append(
                    {
                        "source_file": user.source_file,
                        "branch": row_entry["branch"],
                        "license": canonical,
                        "amount": cost,
                    }
                )

        if not row_entry["branch"] and email not in unresolved_emails:
            unresolved_emails.append(email)

    user_rows: list[dict[str, Any]] = []
    for row in sorted(user_rows_map.values(), key=lambda item: (item["last_name"], item["first_name"], item["email"])):
        user_rows.append(
            {
                "email": row["email"],
                "first_name": row["first_name"],
                "last_name": row["last_name"],
                "branch": row["branch"],
                "license_list": ", ".join(row["licenses"]),
                "user_total": float(Decimal(str(row["user_total"])).quantize(Decimal("0.01"))),
                "known_user": bool(row["branch"]),
            }
        )

    return line_rows, user_rows, warnings, unresolved_emails


def _normalize_integricom_branch(office: str | None, department: str | None = None) -> str:
    department_raw = (department or "").strip()
    if "construction" in department_raw.lower():
        return "Construction"

    raw = (office or "").strip()
    return INTEGRICOM_BRANCH_ALIASES.get(raw, raw or INTEGRICOM_HOME_OFFICE)


def _normalize_integricom_text(value: str) -> str:
    cleaned = value.replace("–", "-")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def _is_integricom_section_header(line: str) -> bool:
    lowered = line.lower()
    return any(
        marker in lowered
        for marker in (
            "products & other charges quantity price amount",
            "netwatch360 limited:",
            "dataprotect 360 backup products:",
            "microsoft 365 products:",
            "dropbox products:",
            "cloud server:",
            "password manager:",
        )
    )


def _canonical_integricom_line(description: str) -> str:
    normalized = _normalize_integricom_text(description).lower()
    if "managed user/workstation" in normalized:
        return "Workstation"
    if "managed firewall" in normalized and "security subscription" not in normalized:
        return "NetWatch360 Managed Firewall"
    if "managed network device" in normalized:
        return "NetWatch360 Managed Network Device"
    if "managed internet" in normalized:
        return "NetWatch360 Managed Internet"
    if "firewall security subscription, main office" in normalized:
        return "Firewall Security Subscription Main Office"
    if "firewall security subscription, district office" in normalized:
        return "Firewall Security Subscription District Office"
    if "latest fw bought in 2025" in normalized or "fw bought in 2025" in normalized:
        return "Firewall Security Subscription Latest 2025"
    if "ticketing system user license" in normalized:
        return "Ticketing System User License"
    if "documentation system" in normalized:
        return "Documentation System License"
    if "monthly recurring block" in normalized or "monthly block hours" in normalized:
        return "Monthly Block Hours"
    if "dark web monitoring" in normalized:
        return "Dark Web Monitoring"
    if "it automation tool" in normalized:
        return "IT Automation Tool"
    if "teams rooms pro" in normalized:
        return "Teams Rooms Pro"
    if "netwatch360 mac" in normalized:
        return "NetWatch360 MAC"
    if "managed server" in normalized and "netwatch360" in normalized:
        return "NetWatch360 Managed Server"
    if "dropbox business standard" in normalized:
        return "Dropbox Business Standard"
    if "office 365 cloud backup" in normalized:
        return "Office 365 Cloud Backup"
    if "server image backup, cloud" in normalized:
        return "DP Server Image Backup Cloud"
    if "business premium" in normalized and "microsoft 365" in normalized:
        return "Microsoft Business Premium Annual"
    if "power bi pro" in normalized:
        return "Power BI Pro"
    if "project plan 3" in normalized:
        return "Project Plan 3"
    if "exchange online p1" in normalized:
        return "Exchange Online P1 Annual"
    if "microsoft f3" in normalized:
        return "Microsoft F3 Annual"
    if "exchange online plan 2" in normalized or "exchange online p2" in normalized:
        return "Exchange Online P2 Annual"
    if "teams essentials" in normalized:
        return "Microsoft Teams Essentials NCE Annual"
    if "microsoft e5" in normalized:
        return "M365 Microsoft E5"
    if "intune" in normalized:
        return "M365 Intune"
    if "prorated m365" in normalized:
        return "Prorated M365"
    if "teams audio conferencing" in normalized:
        return "Teams Audio Conferencing"
    if "aws cloud server" in normalized:
        return "AWS Cloud Server"
    if "keeper enterprise" in normalized or normalized == "keeper":
        return "Keeper Enterprise Password Manager"
    return _normalize_integricom_text(description)


def parse_integricom_invoice(filename: str, raw: bytes) -> IntegricomInvoiceParseResult:
    warnings: list[str] = []
    invoice_number: str | None = None
    invoice_total: Decimal | None = None
    line_items: list[IntegricomInvoiceLine] = []

    try:
        from pypdf import PdfReader
    except Exception:
        return IntegricomInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            credits_total=Decimal("0.00"),
            line_items=[],
            warnings=["Invoice parser unavailable (pypdf not installed)."],
        )

    try:
        reader = PdfReader(io.BytesIO(raw))
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        return IntegricomInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            credits_total=Decimal("0.00"),
            line_items=[],
            warnings=[f"Could not parse PDF text from {filename}: {exc}"],
        )

    number_match = re.search(r"Date\s+Invoice\s*[0-9/]+\s+([0-9]{3,})", text, flags=re.IGNORECASE)
    if number_match:
        invoice_number = number_match.group(1).strip()

    invoice_total = _money_from_text(
        text,
        patterns=[
            r"Invoice Total:\s*\$?\s*([0-9][0-9,]*\.\d{2})",
            r"Invoice Subtotal:\s*\$?\s*([0-9][0-9,]*\.\d{2})",
            r"Balance Due:\s*\$?\s*([0-9][0-9,]*\.\d{2})",
        ],
    )
    credits_total = _money_from_text(
        text,
        patterns=[
            r"Credits:\s*(-?\$?\s*[0-9][0-9,]*\.\d{2})",
            r"Credits:\s*\(?\$?\s*([0-9][0-9,]*\.\d{2})\)?",
        ],
    )
    credits_total = credits_total or Decimal("0.00")
    if invoice_total is not None and credits_total is not None and credits_total != Decimal("0.00"):
        invoice_total = (invoice_total + credits_total).quantize(Decimal("0.01"))
        warnings.append(
            f"{filename}: applied invoice credits of {credits_total} to the Home Office adjustment."
        )
    if invoice_total is None:
        warnings.append(f"{filename}: could not extract Integricom invoice total.")

    pattern = re.compile(
        r"^(?P<desc>.*?)(?P<qty>[0-9][0-9,]*\.[0-9]{2})\s+\$?(?P<price>[0-9][0-9,]*\.[0-9]{2})\s+\$?(?P<amount>[0-9][0-9,]*\.[0-9]{2})$"
    )
    description_buffer: list[str] = []
    for raw_line in text.splitlines():
        line = _normalize_integricom_text(raw_line)
        if not line:
            continue
        lowered = line.lower()
        if _is_integricom_section_header(line):
            description_buffer = []
            continue
        if lowered.startswith("total products & other"):
            description_buffer = []
            continue
        if lowered.startswith("invoice subtotal:") or lowered.startswith("sales tax:"):
            description_buffer = []
            continue
        if lowered.startswith("invoice total:") or lowered.startswith("payments:") or lowered.startswith("credits:"):
            description_buffer = []
            continue
        if lowered.startswith("balance due:") or lowered.startswith("please pay invoices at"):
            description_buffer = []
            continue

        match = pattern.match(line)
        if match:
            inline_desc = _normalize_integricom_text(match.group("desc"))
            if inline_desc:
                description = _normalize_integricom_text(" ".join([*description_buffer, inline_desc]))
            else:
                description = _normalize_integricom_text(" ".join(description_buffer))
            description_buffer = []
            if not description:
                continue

            qty = _parse_decimal(match.group("qty"))
            price = _parse_decimal(match.group("price"))
            amount = _parse_decimal(match.group("amount"))
            if qty is None or price is None or amount is None:
                warnings.append(f"{filename}: could not parse numeric amounts for line '{description}'.")
                continue

            canonical_name = _canonical_integricom_line(description)
            line_items.append(
                IntegricomInvoiceLine(
                    description=description,
                    canonical_name=canonical_name,
                    quantity=qty.quantize(Decimal("0.01")),
                    unit_price=price.quantize(Decimal("0.01")),
                    amount=amount.quantize(Decimal("0.01")),
                )
            )
            continue

        description_buffer.append(line)

    if not line_items:
        warnings.append(f"{filename}: no Integricom line items could be extracted from invoice.")

    return IntegricomInvoiceParseResult(
        filename=filename,
        invoice_number=invoice_number,
        invoice_total=invoice_total,
        credits_total=credits_total,
        line_items=line_items,
        warnings=warnings,
    )


def _integricom_support_summary_from_header(header: str) -> str:
    cleaned = _normalize_integricom_text(header)
    if " / " in cleaned:
        cleaned = cleaned.split(" / ", 1)[1]
    cleaned = re.sub(r"\s+Location:\s*.*$", "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


def _infer_integricom_support_branch(charge_summary: str) -> tuple[str, str, str]:
    summary_lower = charge_summary.lower()
    for branch in INTEGRICOM_KNOWN_BRANCHES:
        if branch == INTEGRICOM_HOME_OFFICE:
            continue
        if branch.lower() in summary_lower:
            return branch, "high", f"Found branch keyword '{branch}' in charge summary."

    return (
        INTEGRICOM_HOME_OFFICE,
        "low",
        "No explicit branch found in charge summary; defaulted to Home Office.",
    )


def parse_integricom_support_invoice(filename: str, raw: bytes) -> IntegricomSupportInvoiceParseResult:
    warnings: list[str] = []
    invoice_number: str | None = None
    invoice_total: Decimal | None = None
    blocks: list[IntegricomSupportBlock] = []

    try:
        from pypdf import PdfReader
    except Exception:
        return IntegricomSupportInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            blocks=[],
            warnings=["Invoice parser unavailable (pypdf not installed)."],
        )

    try:
        reader = PdfReader(io.BytesIO(raw))
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        return IntegricomSupportInvoiceParseResult(
            filename=filename,
            invoice_number=None,
            invoice_total=None,
            blocks=[],
            warnings=[f"Could not parse PDF text from {filename}: {exc}"],
        )

    number_match = re.search(r"Date\s+Invoice\s*[0-9/]+\s+([0-9]{3,})", text, flags=re.IGNORECASE)
    if number_match:
        invoice_number = number_match.group(1).strip()

    invoice_total = _money_from_text(
        text,
        patterns=[
            r"Invoice Total:\s*\$?\s*([0-9][0-9,]*\.\d{2})",
            r"Balance Due:\s*\$?\s*([0-9][0-9,]*\.\d{2})",
        ],
    )
    if invoice_total is None:
        warnings.append(f"{filename}: could not extract Integricom support invoice total.")

    block_pattern = re.compile(
        r"Charge To:\s*(?P<header>.*?)Date Staff Notes Bill Hours Rate Ext Amt(?P<body>.*?)(?=Charge To:|Total Hours:|Invoice Subtotal:|Please pay invoices at|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )

    for block_index, match in enumerate(block_pattern.finditer(text), start=1):
        header = _normalize_integricom_text(match.group("header"))
        body = match.group("body")

        billable_entries = len(re.findall(r"\bY\b", body))
        if billable_entries == 0:
            continue

        billable_hours = Decimal("0.00")
        for hours_match in re.finditer(r"\bY\s+([0-9]+\.[0-9]+)", body):
            parsed_hours = _parse_decimal(hours_match.group(1))
            if parsed_hours is not None:
                billable_hours += parsed_hours
        billable_hours = billable_hours.quantize(Decimal("0.01"))

        line_amount_total = Decimal("0.00")
        for amount_match in re.finditer(
            r"\bY\s+[0-9]+\.[0-9]+\s+[0-9]+\.[0-9]+\s+\$([0-9][0-9,]*\.\d{2})",
            body,
        ):
            parsed_amount = _parse_decimal(amount_match.group(1))
            if parsed_amount is not None:
                line_amount_total += parsed_amount
        line_amount_total = line_amount_total.quantize(Decimal("0.01"))

        subtotal = _money_from_text(body, [r"Subtotal:\s*\$?\s*([0-9][0-9,]*\.\d{2})"])
        if subtotal is None and line_amount_total == Decimal("0.00"):
            warnings.append(
                f"{filename}: block {block_index} has billable entries but no parseable subtotal/line amount; skipped."
            )
            continue

        # Block subtotal is the most stable source in this PDF layout; line captures can be truncated on wraps.
        amount = subtotal if subtotal is not None else line_amount_total
        if amount is None:
            continue

        charge_summary = _integricom_support_summary_from_header(header)
        row_seed = f"{block_index}:{charge_summary.lower()}"
        row_key = f"{block_index}:{hashlib.sha1(row_seed.encode('utf-8')).hexdigest()[:10]}"

        blocks.append(
            IntegricomSupportBlock(
                row_key=row_key,
                charge_summary=charge_summary or f"Support Block {block_index}",
                billable_entries=billable_entries,
                billable_hours=billable_hours,
                amount=amount.quantize(Decimal("0.01")),
            )
        )

    if not blocks:
        warnings.append(f"{filename}: no billable support blocks (Bill=Y) were extracted.")

    return IntegricomSupportInvoiceParseResult(
        filename=filename,
        invoice_number=invoice_number,
        invoice_total=invoice_total,
        blocks=blocks,
        warnings=warnings,
    )


def build_integricom_support_allocations(
    blocks: list[IntegricomSupportBlock],
    support_updates: list[dict[str, Any]] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    updates_by_key: dict[str, str] = {}
    for item in support_updates or []:
        row_key = (item.get("row_key") or "").strip()
        branch = (item.get("branch") or "").strip()
        if not row_key:
            continue
        updates_by_key[row_key] = branch

    warnings: list[str] = []
    line_rows: list[dict[str, Any]] = []
    support_rows: list[dict[str, Any]] = []

    unresolved_reviews = 0
    for block in blocks:
        inferred_branch, confidence, reason = _infer_integricom_support_branch(block.charge_summary)
        branch = inferred_branch
        confidence_label = confidence
        reason_label = reason

        if block.row_key in updates_by_key:
            branch = updates_by_key[block.row_key] or inferred_branch
            confidence_label = "user"
            reason_label = "Branch set by user."

        needs_review = confidence != "high" and block.row_key not in updates_by_key
        if needs_review:
            unresolved_reviews += 1

        line_rows.append(
            {
                "source_file": "invoice",
                "branch": branch,
                "license": f"Support: {block.charge_summary}",
                "amount": block.amount,
            }
        )
        support_rows.append(
            {
                "row_key": block.row_key,
                "charge_summary": block.charge_summary,
                "billable_entries": block.billable_entries,
                "billable_hours": float(block.billable_hours),
                "amount": float(block.amount),
                "branch": branch,
                "confidence": confidence_label,
                "assignment_reason": reason_label,
                "needs_review": needs_review,
            }
        )

    if unresolved_reviews:
        warnings.append(
            f"{unresolved_reviews} support block(s) defaulted to Home Office due to low-confidence matching. Review recommended."
        )

    support_rows.sort(key=lambda row: (not row["needs_review"], row["charge_summary"]))
    return line_rows, support_rows, warnings


def parse_integricom_export_csv(filename: str, raw: bytes) -> IntegricomExportParseResult:
    text = _decode_bytes(raw)
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        return IntegricomExportParseResult(
            filename=filename,
            users=[],
            rows_skipped=0,
            warnings=[f"{filename}: no headers found; file skipped."],
        )

    headers = [h for h in reader.fieldnames if h is not None]
    email_col = _match_header(
        headers,
        [
            "user_principal_name",
            "user_principal",
            "email",
            "user_email",
        ],
    )
    first_name_col = _match_header(headers, ["first_name", "first", "given_name"])
    last_name_col = _match_header(headers, ["last_name", "last", "surname", "family_name"])
    office_col = _match_header(headers, ["office", "branch", "location", "site"])
    department_col = _match_header(headers, ["department", "dept", "division"])
    licenses_col = _match_header(headers, ["licenses", "license", "products"])

    if email_col is None or licenses_col is None:
        return IntegricomExportParseResult(
            filename=filename,
            users=[],
            rows_skipped=0,
            warnings=[f"{filename}: expected columns (User principal name, Licenses) were not found."],
        )

    users: list[IntegricomExportUser] = []
    warnings: list[str] = []
    rows_skipped = 0
    for line_number, row in enumerate(reader, start=2):
        email = (row.get(email_col) or "").strip().lower()
        if not email:
            rows_skipped += 1
            warnings.append(f"{filename}: row {line_number} skipped (missing email).")
            continue
        if "#ext#" in email:
            rows_skipped += 1
            continue

        licenses_raw = (row.get(licenses_col) or "").strip()
        if not licenses_raw or licenses_raw.lower() == "unlicensed":
            rows_skipped += 1
            continue

        tokens = [token.strip() for token in licenses_raw.split("+") if token.strip()]
        if not tokens:
            rows_skipped += 1
            continue

        office = (row.get(office_col) or "").strip() if office_col else ""
        department = (row.get(department_col) or "").strip() if department_col else ""
        users.append(
            IntegricomExportUser(
                source_file=filename,
                email=email,
                first_name=(row.get(first_name_col) or "").strip() if first_name_col else "",
                last_name=(row.get(last_name_col) or "").strip() if last_name_col else "",
                office=office,
                default_branch=_normalize_integricom_branch(office, department),
                licenses=tokens,
            )
        )

    return IntegricomExportParseResult(
        filename=filename,
        users=users,
        rows_skipped=rows_skipped,
        warnings=warnings,
    )


def _integricom_user_matches_rule(user: IntegricomExportUser, canonical_line: str) -> bool:
    tokens = set(user.licenses)
    if canonical_line == "Workstation":
        return any(
            token in tokens
            for token in (
                INTEGRICOM_LICENSE_BP,
                INTEGRICOM_LICENSE_P1,
                INTEGRICOM_LICENSE_P2,
                INTEGRICOM_LICENSE_F3,
                INTEGRICOM_LICENSE_TEAMS_ESSENTIALS,
            )
        )
    if canonical_line == "Office 365 Cloud Backup":
        return INTEGRICOM_LICENSE_BP in tokens or INTEGRICOM_LICENSE_P1 in tokens
    if canonical_line == "Microsoft Business Premium Annual":
        return INTEGRICOM_LICENSE_BP in tokens
    if canonical_line == "Exchange Online P1 Annual":
        return INTEGRICOM_LICENSE_P1 in tokens
    if canonical_line == "Microsoft F3 Annual":
        return INTEGRICOM_LICENSE_F3 in tokens
    if canonical_line == "Exchange Online P2 Annual":
        return INTEGRICOM_LICENSE_P2 in tokens
    return False


def _allocate_integricom_fixed_line(
    line: IntegricomInvoiceLine,
    *,
    line_key: str,
    branch_assignment_updates: dict[tuple[str, int], str],
) -> tuple[list[dict[str, Any]], list[str], list[dict[str, Any]]]:
    rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    pending_branch_prompts: list[dict[str, Any]] = []
    qty_int = int(line.quantity)
    unit = line.unit_price.quantize(Decimal("0.01"))
    total = line.amount.quantize(Decimal("0.01"))

    def add_row(branch: str, amount: Decimal) -> None:
        amount = amount.quantize(Decimal("0.01"))
        if amount == Decimal("0.00"):
            return
        rows.append(
            {
                "source_file": "invoice",
                "branch": branch or INTEGRICOM_HOME_OFFICE,
                "license": line.canonical_name,
                "amount": amount,
            }
        )

    def add_branch_prompt(
        *,
        prompt_index: int,
        already_assigned_branches: list[str],
        submitted_branch: str,
        validation_error: str = "",
    ) -> None:
        available_branches = [branch for branch in INTEGRICOM_KNOWN_BRANCHES if branch not in already_assigned_branches]
        pending_branch_prompts.append(
            {
                "line_key": line_key,
                "prompt_index": prompt_index,
                "license": line.canonical_name,
                "unit_price": float(unit),
                "quantity": qty_int,
                "already_assigned_branches": already_assigned_branches,
                "available_branches": available_branches,
                "branch": submitted_branch,
                "validation_error": validation_error,
            }
        )

    def allocate_by_unit_sequence(branches: list[str]) -> None:
        assigned_units = min(max(qty_int, 0), len(branches))
        assigned_branch_order = branches[:assigned_units]
        assigned_branch_set = set(assigned_branch_order)
        for branch in assigned_branch_order:
            add_row(branch, unit)

        if qty_int < len(branches):
            warnings.append(
                f"{line.canonical_name}: invoice quantity is {qty_int}; template allocation used the first {assigned_units} branches."
            )

        extra_units = max(qty_int - len(branches), 0)
        for prompt_index in range(1, extra_units + 1):
            submitted_branch = (branch_assignment_updates.get((line_key, prompt_index)) or "").strip()
            if not submitted_branch:
                add_branch_prompt(
                    prompt_index=prompt_index,
                    already_assigned_branches=list(assigned_branch_order),
                    submitted_branch="",
                )
                continue

            if submitted_branch in assigned_branch_set:
                error_message = f"{submitted_branch} is already assigned for this charge."
                warnings.append(
                    f"{line.canonical_name}: branch '{submitted_branch}' is already assigned; choose a different branch for extra license {prompt_index}."
                )
                add_branch_prompt(
                    prompt_index=prompt_index,
                    already_assigned_branches=list(assigned_branch_order),
                    submitted_branch=submitted_branch,
                    validation_error=error_message,
                )
                continue

            add_row(submitted_branch, unit)
            assigned_branch_set.add(submitted_branch)
            assigned_branch_order.append(submitted_branch)

        if pending_branch_prompts:
            warnings.append(
                f"{line.canonical_name}: {len(pending_branch_prompts)} extra branch assignment(s) required before this charge can be finalized."
            )
            return

        assigned_amount = (unit * Decimal(len(assigned_branch_order))).quantize(Decimal("0.01"))
        remainder = (total - assigned_amount).quantize(Decimal("0.01"))
        if remainder != Decimal("0.00"):
            add_row(INTEGRICOM_HOME_OFFICE, remainder)

    fixed_home_office = {
        "Ticketing System User License",
        "Documentation System License",
        "Monthly Block Hours",
        "Dark Web Monitoring",
        "IT Automation Tool",
        "Teams Rooms Pro",
        "NetWatch360 MAC",
        "NetWatch360 Managed Server",
        "Dropbox Business Standard",
        "DP Server Image Backup Cloud",
        "Power BI Pro",
        "Microsoft Teams Essentials NCE Annual",
        "M365 Microsoft E5",
        "M365 Intune",
        "Prorated M365",
        "AWS Cloud Server",
        "Keeper Enterprise Password Manager",
        "Teams Audio Conferencing",
    }

    if line.canonical_name in fixed_home_office:
        add_row(INTEGRICOM_HOME_OFFICE, total)
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "NetWatch360 Managed Firewall":
        allocate_by_unit_sequence(INTEGRICOM_DISTRICT_BRANCHES)
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "NetWatch360 Managed Network Device":
        add_row(INTEGRICOM_HOME_OFFICE, total)
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "NetWatch360 Managed Internet":
        allocate_by_unit_sequence(INTEGRICOM_MANAGED_INTERNET_BRANCHES)
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "Firewall Security Subscription Main Office":
        sugar_hill_amount = Decimal("97.00")
        if total >= sugar_hill_amount:
            add_row("Sugar Hill", sugar_hill_amount)
            add_row(INTEGRICOM_HOME_OFFICE, total - sugar_hill_amount)
        else:
            add_row(INTEGRICOM_HOME_OFFICE, total)
            warnings.append(
                f"{line.canonical_name}: invoice amount was below expected split baseline; allocated entirely to Home Office."
            )
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "Firewall Security Subscription District Office":
        allocate_by_unit_sequence(
            [
                "Canton",
                "Cobb",
                "Doraville",
                "Destin",
                "Fort Walton",
                "Tampa",
                "Savannah",
                "Charleston",
                "Nashville",
                "Color Burst",
                "Acworth",
            ]
        )
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "Firewall Security Subscription Latest 2025":
        add_row("St. Pete", total)
        return rows, warnings, pending_branch_prompts

    if line.canonical_name == "Project Plan 3":
        add_row("Sugar Hill", total)
        return rows, warnings, pending_branch_prompts

    add_row(INTEGRICOM_HOME_OFFICE, total)
    warnings.append(
        f"{line.canonical_name}: no Integricom allocation rule configured; amount allocated to Home Office."
    )
    return rows, warnings, pending_branch_prompts


def build_integricom_user_allocations(
    users: list[IntegricomExportUser],
    user_directory: dict[str, dict[str, str]],
    invoice_lines: list[IntegricomInvoiceLine],
    branch_item_updates: list[dict[str, Any]] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[str], list[str], list[dict[str, Any]]]:
    line_rows: list[dict[str, Any]] = []
    non_user_rows_raw: list[dict[str, Any]] = []
    warnings: list[str] = []
    unresolved_emails: list[str] = []
    unresolved_branch_prompts: list[dict[str, Any]] = []
    user_rows_map: dict[str, dict[str, Any]] = {}
    branch_assignment_updates_map: dict[tuple[str, int], str] = {}

    for update in branch_item_updates or []:
        line_key = (update.get("line_key") or "").strip()
        branch = (update.get("branch") or "").strip()
        try:
            prompt_index = int(update.get("prompt_index"))
        except (TypeError, ValueError):
            continue
        if not line_key or prompt_index < 1:
            continue
        branch_assignment_updates_map[(line_key, prompt_index)] = branch

    for user in users:
        email = user.email.strip().lower()
        if not email:
            continue
        profile = user_directory.get(email)
        directory_branch = (profile.get("branch") if profile else "") or ""
        branch = (directory_branch or user.default_branch).strip()
        if user.default_branch == "Construction":
            branch = "Construction"
        first_name = (user.first_name or "").strip()
        last_name = (user.last_name or "").strip()
        if profile:
            first_name = first_name or (profile.get("first_name") or "").strip()
            last_name = last_name or (profile.get("last_name") or "").strip()

        user_rows_map[email] = {
            "email": email,
            "first_name": first_name,
            "last_name": last_name,
            "branch": branch,
            "licenses": [],
            "user_total": Decimal("0.00"),
            "known_user": bool(profile and (profile.get("branch") or "").strip()),
        }
        if not branch and email not in unresolved_emails:
            unresolved_emails.append(email)

    dynamic_licenses = {
        "Workstation",
        "Office 365 Cloud Backup",
        "Microsoft Business Premium Annual",
        "Exchange Online P1 Annual",
        "Microsoft F3 Annual",
        "Exchange Online P2 Annual",
    }

    for line_index, line in enumerate(invoice_lines, start=1):
        line_key = f"{line_index}:{line.canonical_name}"
        if line.canonical_name in dynamic_licenses:
            matched_users = [user for user in users if _integricom_user_matches_rule(user, line.canonical_name)]
            matched_count = len(matched_users)
            for user in matched_users:
                entry = user_rows_map.get(user.email)
                if entry is None:
                    continue
                if line.canonical_name not in entry["licenses"]:
                    entry["licenses"].append(line.canonical_name)
                entry["user_total"] += line.unit_price
                if entry["branch"]:
                    line_rows.append(
                        {
                            "source_file": "invoice",
                            "branch": entry["branch"],
                            "license": line.canonical_name,
                            "amount": line.unit_price,
                        }
                    )

            invoice_qty = int(line.quantity)
            if matched_count != invoice_qty:
                warnings.append(
                    f"{line.canonical_name}: invoice quantity is {invoice_qty}, matched users are {matched_count}; difference allocated to Home Office."
                )

            allocated_total = (line.unit_price * Decimal(matched_count)).quantize(Decimal("0.01"))
            remainder = (line.amount - allocated_total).quantize(Decimal("0.01"))
            if remainder != Decimal("0.00"):
                remainder_row = {
                    "source_file": "invoice",
                    "branch": INTEGRICOM_HOME_OFFICE,
                    "license": line.canonical_name,
                    "amount": remainder,
                }
                line_rows.append(remainder_row)
                non_user_rows_raw.append(
                    {
                        **remainder_row,
                        "allocation_type": "Invoice Delta",
                    }
                )
            continue

        fixed_rows, fixed_warnings, pending_branch_rows = _allocate_integricom_fixed_line(
            line,
            line_key=line_key,
            branch_assignment_updates=branch_assignment_updates_map,
        )
        line_rows.extend(fixed_rows)
        non_user_rows_raw.extend(
            [
                {
                    **row,
                    "allocation_type": "Fixed Branch Item",
                }
                for row in fixed_rows
            ]
        )
        warnings.extend(fixed_warnings)
        unresolved_branch_prompts.extend(pending_branch_rows)

    grouped_non_user: dict[tuple[str, str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    for row in non_user_rows_raw:
        key = (row["branch"], row["license"], row["allocation_type"])
        grouped_non_user[key] += row["amount"]

    non_user_rows: list[dict[str, Any]] = []
    for (branch, license_name, allocation_type), total in sorted(grouped_non_user.items()):
        non_user_rows.append(
            {
                "branch": branch,
                "license": license_name,
                "allocation_type": allocation_type,
                "total_amount": float(total.quantize(Decimal("0.01"))),
            }
        )

    user_rows: list[dict[str, Any]] = []
    for row in sorted(user_rows_map.values(), key=lambda item: (item["last_name"], item["first_name"], item["email"])):
        user_rows.append(
            {
                "email": row["email"],
                "first_name": row["first_name"],
                "last_name": row["last_name"],
                "branch": row["branch"],
                "license_list": ", ".join(row["licenses"]),
                "user_total": float(Decimal(str(row["user_total"])).quantize(Decimal("0.01"))),
                "known_user": bool(row["branch"]),
            }
        )

    return line_rows, user_rows, non_user_rows, warnings, unresolved_emails, unresolved_branch_prompts


def apply_home_office_adjustment(
    summary: list[dict[str, Any]],
    adjustment: Decimal,
    *,
    license_name: str = HEXNODE_DEFAULT_LICENSE,
    home_office_name: str = HEXNODE_HOME_OFFICE,
) -> list[dict[str, Any]]:
    if adjustment == Decimal("0"):
        return summary

    updated = [dict(row) for row in summary]
    updated_row = None
    for row in updated:
        if row["branch"] == home_office_name and row["license"] == license_name:
            updated_row = row
            break

    if updated_row is None:
        updated.append(
            {
                "branch": home_office_name,
                "license": license_name,
                "total_amount": 0.0,
            }
        )
        updated_row = updated[-1]

    current = Decimal(str(updated_row["total_amount"]))
    updated_row["total_amount"] = float((current + adjustment).quantize(Decimal("0.01")))
    updated.sort(key=lambda item: (item["branch"], item["license"]))
    return updated
