from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path
from typing import Any

from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .adobe_directory import (
    deactivate_adobe_users as deactivate_adobe_directory_users,
    find_missing_users,
    init_adobe_directory,
    list_adobe_users,
    touch_seen_users,
    upsert_adobe_users,
)
from .integricom_directory import (
    deactivate_integricom_users as deactivate_integricom_directory_users,
    find_missing_integricom_users,
    init_integricom_directory,
    list_integricom_users,
    touch_seen_integricom_users,
    upsert_integricom_users,
)
from .entra_graph import EntraSyncError, sync_integricom_users_from_entra
from .processing import (
    ADOBE_ADJUSTMENT_LICENSE,
    ADOBE_HOME_OFFICE,
    HEXNODE_DEFAULT_LICENSE,
    IntegricomExportUser,
    INTEGRICOM_ADJUSTMENT_LICENSE,
    INTEGRICOM_HOME_OFFICE,
    INTEGRICOM_KNOWN_BRANCHES,
    apply_home_office_adjustment,
    build_adobe_user_allocations,
    build_integricom_support_allocations,
    build_integricom_user_allocations,
    build_breakdown,
    parse_adobe_export_csv,
    parse_adobe_invoice,
    parse_integricom_export_csv,
    parse_integricom_invoice,
    parse_integricom_support_invoice,
    parse_csv,
    parse_hexnode_csv,
    parse_hexnode_invoice,
    summary_to_csv,
)
from .spreadsheet_import import parse_adobe_directory_import_file

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="AP Allocation App", version="0.3.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
def launcher() -> FileResponse:
    return FileResponse(STATIC_DIR / "launcher.html")


@app.get("/apps/invoice-analyzer")
def invoice_analyzer_app() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/apps/admin")
def admin_app() -> FileResponse:
    return FileResponse(STATIC_DIR / "admin.html")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/adobe/users/save")
def save_adobe_users(payload: list[dict[str, Any]] = Body(...)) -> dict[str, Any]:
    parsed = _parse_user_updates(json.dumps(payload), field_name="adobe_user_updates")
    rows_to_upsert = [
        {
            "email": item["email"],
            "first_name": item["first_name"],
            "last_name": item["last_name"],
            "branch": item["branch"],
        }
        for item in parsed
        if item["branch"]
    ]
    upsert_adobe_users(rows_to_upsert)
    return {
        "received": len(parsed),
        "saved": len(rows_to_upsert),
        "skipped_blank_branch": len(parsed) - len(rows_to_upsert),
    }


@app.get("/api/adobe/users")
def get_adobe_users(active_only: bool = True) -> dict[str, Any]:
    users = list_adobe_users(active_only=active_only)
    serialized = _serialize_directory_users(users)
    return {
        "vendor": "adobe",
        "active_only": active_only,
        "count": len(serialized),
        "users": serialized,
    }


@app.post("/api/adobe/users/deactivate")
def deactivate_adobe_users(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    emails = _parse_email_list(payload, field_name="emails")
    count = deactivate_adobe_directory_users(emails)
    return {
        "requested": len(emails),
        "deactivated": count,
    }


@app.post("/api/adobe/users/import")
async def import_adobe_users(mapping_file: UploadFile = File(...)) -> dict[str, Any]:
    if not mapping_file.filename:
        raise HTTPException(status_code=400, detail="Please choose a spreadsheet file to import.")

    raw = await mapping_file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="The uploaded spreadsheet is empty.")

    parsed = parse_adobe_directory_import_file(mapping_file.filename, raw)
    if not parsed.rows:
        raise HTTPException(status_code=400, detail=" | ".join(parsed.warnings) or "No importable users found.")

    upsert_adobe_users(parsed.rows)
    touch_seen_users(parsed.rows)

    return {
        "source": parsed.source,
        "filename": mapping_file.filename,
        "imported": len(parsed.rows),
        "warnings": parsed.warnings,
    }


@app.post("/api/integricom/users/save")
def save_integricom_users(payload: list[dict[str, Any]] = Body(...)) -> dict[str, Any]:
    parsed = _parse_user_updates(json.dumps(payload), field_name="integricom_user_updates")
    rows_to_upsert = [
        {
            "email": item["email"],
            "first_name": item["first_name"],
            "last_name": item["last_name"],
            "branch": item["branch"],
        }
        for item in parsed
        if item["branch"]
    ]
    upsert_integricom_users(rows_to_upsert)
    return {
        "received": len(parsed),
        "saved": len(rows_to_upsert),
        "skipped_blank_branch": len(parsed) - len(rows_to_upsert),
    }


@app.get("/api/integricom/users")
def get_integricom_users(active_only: bool = True) -> dict[str, Any]:
    users = list_integricom_users(active_only=active_only)
    serialized = _serialize_directory_users(users)
    return {
        "vendor": "integricom",
        "active_only": active_only,
        "count": len(serialized),
        "users": serialized,
    }


@app.post("/api/integricom/users/deactivate")
def deactivate_integricom_users(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    emails = _parse_email_list(payload, field_name="emails")
    count = deactivate_integricom_directory_users(emails)
    return {
        "requested": len(emails),
        "deactivated": count,
    }


@app.post("/api/integricom/sync/entra")
def sync_integricom_users_from_entra_endpoint() -> dict[str, Any]:
    try:
        result = sync_integricom_users_from_entra()
    except EntraSyncError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    upsert_integricom_users(result.users)
    touch_seen_integricom_users(result.users)

    return {
        "synced": len(result.users),
        "users_scanned": result.users_scanned,
        "users_with_supported_licenses": result.users_with_supported_licenses,
        "users_skipped_external": result.users_skipped_external,
        "users_skipped_unlicensed": result.users_skipped_unlicensed,
        "unknown_sku_parts": result.unknown_sku_parts,
        "warnings": result.warnings,
    }


def _parse_user_updates(raw: str | None, *, field_name: str) -> list[dict[str, str]]:
    if not raw:
        return []

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid {field_name} JSON: {exc}") from exc

    if not isinstance(payload, list):
        raise HTTPException(status_code=400, detail=f"{field_name} must be a JSON array.")

    parsed: list[dict[str, str]] = []
    for index, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail=f"{field_name} item {index} must be an object.")

        email = (item.get("email") or "").strip().lower()
        if not email:
            raise HTTPException(status_code=400, detail=f"{field_name} item {index} is missing email.")

        parsed.append(
            {
                "email": email,
                "first_name": (item.get("first_name") or "").strip(),
                "last_name": (item.get("last_name") or "").strip(),
                "branch": (item.get("branch") or "").strip(),
            }
        )
    return parsed


def _parse_email_list(payload: dict[str, Any], *, field_name: str) -> list[str]:
    raw_list = payload.get(field_name)
    if not isinstance(raw_list, list):
        raise HTTPException(status_code=400, detail=f"{field_name} must be a JSON array.")

    parsed: list[str] = []
    for index, value in enumerate(raw_list, start=1):
        if not isinstance(value, str):
            raise HTTPException(status_code=400, detail=f"{field_name} item {index} must be a string.")
        email = value.strip().lower()
        if not email:
            raise HTTPException(status_code=400, detail=f"{field_name} item {index} is blank.")
        parsed.append(email)
    return parsed


def _parse_integricom_branch_item_updates(raw: str | None) -> list[dict[str, Any]]:
    if not raw:
        return []

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid integricom_branch_item_updates JSON: {exc}") from exc

    if not isinstance(payload, list):
        raise HTTPException(status_code=400, detail="integricom_branch_item_updates must be a JSON array.")

    parsed: list[dict[str, Any]] = []
    for index, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail=f"integricom_branch_item_updates item {index} must be an object.")

        line_key = (item.get("line_key") or "").strip()
        if not line_key:
            raise HTTPException(
                status_code=400,
                detail=f"integricom_branch_item_updates item {index} is missing line_key.",
            )

        prompt_index_raw = item.get("prompt_index")
        try:
            prompt_index = int(prompt_index_raw)
        except (TypeError, ValueError) as exc:
            raise HTTPException(
                status_code=400,
                detail=f"integricom_branch_item_updates item {index} has an invalid prompt_index.",
            ) from exc
        if prompt_index < 1:
            raise HTTPException(
                status_code=400,
                detail=f"integricom_branch_item_updates item {index} has an invalid prompt_index.",
            )

        parsed.append(
            {
                "line_key": line_key,
                "prompt_index": prompt_index,
                "branch": (item.get("branch") or "").strip(),
            }
        )
    return parsed


def _parse_integricom_support_updates(raw: str | None) -> list[dict[str, str]]:
    if not raw:
        return []

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid integricom_support_updates JSON: {exc}") from exc

    if not isinstance(payload, list):
        raise HTTPException(status_code=400, detail="integricom_support_updates must be a JSON array.")

    parsed: list[dict[str, str]] = []
    for index, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail=f"integricom_support_updates item {index} must be an object.")

        row_key = (item.get("row_key") or "").strip()
        if not row_key:
            raise HTTPException(
                status_code=400,
                detail=f"integricom_support_updates item {index} is missing row_key.",
            )

        parsed.append(
            {
                "row_key": row_key,
                "branch": (item.get("branch") or "").strip(),
            }
        )
    return parsed


def _directory_to_profile_map(directory: dict[str, Any]) -> dict[str, dict[str, str]]:
    profiles: dict[str, dict[str, str]] = {}
    for email, user in directory.items():
        profiles[email] = {
            "branch": user.branch,
            "first_name": user.first_name,
            "last_name": user.last_name,
        }
    return profiles


def _serialize_directory_users(directory: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for email, user in sorted(directory.items(), key=lambda item: item[0]):
        rows.append(
            {
                "email": email,
                "first_name": user.first_name,
                "last_name": user.last_name,
                "branch": user.branch,
                "is_active": bool(user.is_active),
                "created_at": user.created_at,
                "updated_at": user.updated_at,
                "last_seen_at": user.last_seen_at,
            }
        )
    return rows


def _serialize_missing_adobe_users(current_emails: set[str]) -> list[dict[str, Any]]:
    missing = find_missing_users(current_emails)
    return [
        {
            "email": user.email,
            "first_name": user.first_name,
            "last_name": user.last_name,
            "branch": user.branch,
            "last_seen_at": user.last_seen_at,
        }
        for user in missing
    ]


def _serialize_missing_integricom_users(current_emails: set[str]) -> list[dict[str, Any]]:
    missing = find_missing_integricom_users(current_emails)
    return [
        {
            "email": user.email,
            "first_name": user.first_name,
            "last_name": user.last_name,
            "branch": user.branch,
            "last_seen_at": user.last_seen_at,
        }
        for user in missing
    ]


async def _analyze_adobe(
    csv_files: list[UploadFile],
    invoice_file: UploadFile | None,
    adobe_user_updates: str | None,
) -> dict[str, Any]:
    if invoice_file is None or not invoice_file.filename:
        raise HTTPException(status_code=400, detail="Adobe mode requires an invoice PDF upload.")

    warnings: list[str] = []
    file_summaries: list[dict[str, Any]] = []

    invoice_raw = await invoice_file.read()
    parsed_invoice = parse_adobe_invoice(invoice_file.filename, invoice_raw)
    warnings.extend(parsed_invoice.warnings)
    if not parsed_invoice.per_license_cost:
        raise HTTPException(status_code=400, detail="Could not parse Adobe invoice line-item pricing.")

    export_users = []
    for upload in csv_files:
        if not upload.filename:
            continue

        raw = await upload.read()
        if not raw:
            warnings.append(f"{upload.filename}: empty file skipped.")
            continue

        parsed_export = parse_adobe_export_csv(upload.filename, raw)
        export_users.extend(parsed_export.users)
        file_summaries.append(
            {
                "filename": upload.filename,
                "rows_ingested": len(parsed_export.users),
                "rows_skipped": parsed_export.rows_skipped,
            }
        )
        warnings.extend(parsed_export.warnings)

    init_adobe_directory()
    submitted_updates = _parse_user_updates(adobe_user_updates, field_name="adobe_user_updates")
    if submitted_updates:
        upsert_adobe_users(
            [
                {
                    "email": item["email"],
                    "first_name": item["first_name"],
                    "last_name": item["last_name"],
                    "branch": item["branch"],
                }
                for item in submitted_updates
                if item["branch"]
            ]
        )

    directory = list_adobe_users()
    directory_profiles = _directory_to_profile_map(directory)

    all_rows, adobe_user_rows, allocation_warnings, unresolved_emails = build_adobe_user_allocations(
        export_users,
        directory_profiles,
        parsed_invoice.per_license_cost,
    )
    warnings.extend(allocation_warnings)

    current_emails = {user.email for user in export_users if user.email}
    missing_users = _serialize_missing_adobe_users(current_emails)

    if unresolved_emails:
        unresolved_set = set(unresolved_emails)
        return {
            "vendor_type": "adobe",
            "needs_user_enrichment": True,
            "needs_non_user_branch_assignment": False,
            "message": "Some users are missing a branch. Enter branch values, then analyze again.",
            "new_users": [row for row in adobe_user_rows if row["email"] in unresolved_set],
            "adobe_user_rows": adobe_user_rows,
            "user_rows": adobe_user_rows,
            "non_user_rows": [],
            "non_user_branch_prompts": [],
            "missing_users": missing_users,
            "files": file_summaries,
            "summary": [],
            "totals": {"line_items": 0, "grand_total": 0.0},
            "reconciliation": None,
            "invoice": {
                "filename": invoice_file.filename,
                "size_bytes": len(invoice_raw),
                "invoice_number": parsed_invoice.invoice_number,
                "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total else None,
                "parsed_licenses": sorted(parsed_invoice.per_license_cost.keys()),
                "directory_users": len(directory),
            },
            "warnings": warnings,
            "breakdown_csv": "",
        }

    touch_seen_users(
        [
            {
                "email": user.email,
                "first_name": user.first_name,
                "last_name": user.last_name,
            }
            for user in export_users
        ]
    )

    summary = build_breakdown(all_rows)
    base_total = sum(Decimal(str(item["total_amount"])) for item in summary)

    reconciliation = None
    if parsed_invoice.invoice_total is not None:
        adjustment = (parsed_invoice.invoice_total - base_total).quantize(Decimal("0.01"))
        summary = apply_home_office_adjustment(
            summary,
            adjustment,
            license_name=ADOBE_ADJUSTMENT_LICENSE,
            home_office_name=ADOBE_HOME_OFFICE,
        )
        reconciliation = {
            "base_total": float(base_total.quantize(Decimal("0.01"))),
            "invoice_total": float(parsed_invoice.invoice_total),
            "home_office_adjustment": float(adjustment),
        }

    breakdown_csv = summary_to_csv(summary)
    grand_total = round(sum(item["total_amount"] for item in summary), 2)

    return {
        "vendor_type": "adobe",
        "needs_user_enrichment": False,
        "needs_non_user_branch_assignment": False,
        "new_users": [],
        "adobe_user_rows": adobe_user_rows,
        "user_rows": adobe_user_rows,
        "non_user_rows": [],
        "non_user_branch_prompts": [],
        "missing_users": missing_users,
        "invoice": {
            "filename": invoice_file.filename,
            "size_bytes": len(invoice_raw),
            "invoice_number": parsed_invoice.invoice_number,
            "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total else None,
            "parsed_licenses": sorted(parsed_invoice.per_license_cost.keys()),
            "directory_users": len(directory),
        },
        "files": file_summaries,
        "summary": summary,
        "totals": {
            "line_items": len(summary),
            "grand_total": grand_total,
        },
        "reconciliation": reconciliation,
        "warnings": warnings,
        "breakdown_csv": breakdown_csv,
    }


async def _analyze_integricom(
    csv_files: list[UploadFile],
    invoice_file: UploadFile | None,
    integricom_user_updates: str | None,
    integricom_branch_item_updates: str | None,
) -> dict[str, Any]:
    if invoice_file is None or not invoice_file.filename:
        raise HTTPException(status_code=400, detail="Integricom mode requires an invoice PDF upload.")

    warnings: list[str] = []
    file_summaries: list[dict[str, Any]] = []

    invoice_raw = await invoice_file.read()
    parsed_invoice = parse_integricom_invoice(invoice_file.filename, invoice_raw)
    warnings.extend(parsed_invoice.warnings)
    if not parsed_invoice.line_items:
        raise HTTPException(status_code=400, detail="Could not parse Integricom invoice line items.")

    export_users: list[IntegricomExportUser] = []
    csv_upload_requested = bool(csv_files)
    for upload in csv_files:
        if not upload.filename:
            continue
        raw = await upload.read()
        if not raw:
            warnings.append(f"{upload.filename}: empty file skipped.")
            continue

        parsed_export = parse_integricom_export_csv(upload.filename, raw)
        export_users.extend(parsed_export.users)
        file_summaries.append(
            {
                "filename": upload.filename,
                "rows_ingested": len(parsed_export.users),
                "rows_skipped": parsed_export.rows_skipped,
            }
        )
        warnings.extend(parsed_export.warnings)

    if not export_users:
        try:
            entra_sync = sync_integricom_users_from_entra()
        except EntraSyncError as exc:
            if csv_upload_requested:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "No usable Integricom users were found in uploaded CSVs and Entra fallback failed: "
                        f"{exc}"
                    ),
                ) from exc
            raise HTTPException(
                status_code=400,
                detail=(
                    "Integricom mode now allows no CSV uploads, but Entra sync is not ready. "
                    f"{exc}"
                ),
            ) from exc

        export_users = [
            IntegricomExportUser(
                source_file="entra",
                email=user.email,
                first_name=user.first_name,
                last_name=user.last_name,
                office=user.office,
                default_branch=user.default_branch,
                licenses=user.licenses,
            )
            for user in entra_sync.export_users
        ]
        if not export_users:
            raise HTTPException(
                status_code=400,
                detail=(
                    "No supported licensed users were found in Microsoft Entra for Integricom allocation. "
                    "Upload a CSV or verify Entra license assignments."
                ),
            )

        file_summaries.append(
            {
                "filename": "Microsoft Entra (Graph)",
                "rows_ingested": len(export_users),
                "rows_skipped": entra_sync.users_skipped_external + entra_sync.users_skipped_unlicensed,
            }
        )
        warnings.extend(entra_sync.warnings)
        if csv_upload_requested:
            warnings.append("CSV uploads produced no usable Integricom users; used Entra sync fallback.")
        else:
            warnings.append("No CSV uploaded; used Microsoft Entra sync for Integricom user licensing.")

    init_integricom_directory()
    submitted_updates = _parse_user_updates(integricom_user_updates, field_name="integricom_user_updates")
    submitted_branch_item_updates = _parse_integricom_branch_item_updates(integricom_branch_item_updates)
    if submitted_updates:
        upsert_integricom_users(
            [
                {
                    "email": item["email"],
                    "first_name": item["first_name"],
                    "last_name": item["last_name"],
                    "branch": item["branch"],
                }
                for item in submitted_updates
                if item["branch"]
            ]
        )

    directory = list_integricom_users()
    seed_users: list[dict[str, str]] = []
    for user in export_users:
        if user.email in directory:
            continue
        seed_users.append(
            {
                "email": user.email,
                "first_name": user.first_name,
                "last_name": user.last_name,
                "branch": user.default_branch,
            }
        )
    if seed_users:
        upsert_integricom_users(seed_users)
        directory = list_integricom_users()

    directory_profiles = _directory_to_profile_map(directory)
    (
        all_rows,
        user_rows,
        non_user_rows,
        allocation_warnings,
        unresolved_emails,
        unresolved_branch_prompts,
    ) = build_integricom_user_allocations(
        export_users,
        directory_profiles,
        parsed_invoice.line_items,
        branch_item_updates=submitted_branch_item_updates,
    )
    warnings.extend(allocation_warnings)

    current_emails = {user.email for user in export_users if user.email}
    missing_users = _serialize_missing_integricom_users(current_emails)

    needs_user_enrichment = bool(unresolved_emails)
    needs_branch_assignment = bool(unresolved_branch_prompts)
    if needs_user_enrichment or needs_branch_assignment:
        unresolved_set = set(unresolved_emails)
        message_parts: list[str] = []
        if needs_user_enrichment:
            message_parts.append("Some users are missing a branch.")
        if needs_branch_assignment:
            message_parts.append("Some branch-tethered charges need branch assignments.")
        message_parts.append("Enter missing values, then analyze again.")
        return {
            "vendor_type": "integricom",
            "needs_user_enrichment": needs_user_enrichment,
            "needs_non_user_branch_assignment": needs_branch_assignment,
            "message": " ".join(message_parts),
            "new_users": [row for row in user_rows if row["email"] in unresolved_set],
            "user_rows": user_rows,
            "integricom_user_rows": user_rows,
            "non_user_rows": non_user_rows,
            "integricom_non_user_rows": non_user_rows,
            "non_user_branch_prompts": unresolved_branch_prompts,
            "integricom_non_user_branch_prompts": unresolved_branch_prompts,
            "missing_users": missing_users,
            "files": file_summaries,
            "summary": [],
            "totals": {"line_items": 0, "grand_total": 0.0},
            "reconciliation": None,
            "invoice": {
                "filename": invoice_file.filename,
                "size_bytes": len(invoice_raw),
                "invoice_number": parsed_invoice.invoice_number,
                "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total else None,
                "parsed_licenses": sorted({line.canonical_name for line in parsed_invoice.line_items}),
                "directory_users": len(directory),
            },
            "warnings": warnings,
            "breakdown_csv": "",
        }

    touch_seen_integricom_users(
        [
            {
                "email": user.email,
                "first_name": user.first_name,
                "last_name": user.last_name,
            }
            for user in export_users
        ]
    )

    summary = build_breakdown(all_rows)
    base_total = sum(Decimal(str(item["total_amount"])) for item in summary)
    reconciliation = None
    if parsed_invoice.invoice_total is not None:
        adjustment = (parsed_invoice.invoice_total - base_total).quantize(Decimal("0.01"))
        summary = apply_home_office_adjustment(
            summary,
            adjustment,
            license_name=INTEGRICOM_ADJUSTMENT_LICENSE,
            home_office_name=INTEGRICOM_HOME_OFFICE,
        )
        reconciliation = {
            "base_total": float(base_total.quantize(Decimal("0.01"))),
            "invoice_total": float(parsed_invoice.invoice_total),
            "home_office_adjustment": float(adjustment),
        }

    breakdown_csv = summary_to_csv(summary)
    grand_total = round(sum(item["total_amount"] for item in summary), 2)

    return {
        "vendor_type": "integricom",
        "needs_user_enrichment": False,
        "needs_non_user_branch_assignment": False,
        "new_users": [],
        "user_rows": user_rows,
        "integricom_user_rows": user_rows,
        "non_user_rows": non_user_rows,
        "integricom_non_user_rows": non_user_rows,
        "non_user_branch_prompts": [],
        "integricom_non_user_branch_prompts": [],
        "missing_users": missing_users,
        "invoice": {
            "filename": invoice_file.filename,
            "size_bytes": len(invoice_raw),
            "invoice_number": parsed_invoice.invoice_number,
            "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total else None,
            "parsed_licenses": sorted({line.canonical_name for line in parsed_invoice.line_items}),
            "directory_users": len(directory),
        },
        "files": file_summaries,
        "summary": summary,
        "totals": {
            "line_items": len(summary),
            "grand_total": grand_total,
        },
        "reconciliation": reconciliation,
        "warnings": warnings,
        "breakdown_csv": breakdown_csv,
    }


async def _analyze_integricom_support(
    invoice_file: UploadFile | None,
    integricom_support_updates: str | None,
) -> dict[str, Any]:
    if invoice_file is None or not invoice_file.filename:
        raise HTTPException(status_code=400, detail="Integricom Support mode requires an invoice PDF upload.")

    warnings: list[str] = []
    invoice_raw = await invoice_file.read()
    parsed_invoice = parse_integricom_support_invoice(invoice_file.filename, invoice_raw)
    warnings.extend(parsed_invoice.warnings)
    if not parsed_invoice.blocks:
        raise HTTPException(status_code=400, detail="Could not parse billable support blocks (Bill=Y) from invoice.")

    submitted_updates = _parse_integricom_support_updates(integricom_support_updates)
    line_rows, support_rows, allocation_warnings = build_integricom_support_allocations(
        parsed_invoice.blocks,
        submitted_updates,
    )
    warnings.extend(allocation_warnings)

    summary = build_breakdown(line_rows)
    base_total = sum(Decimal(str(item["total_amount"])) for item in summary)
    reconciliation = None
    if parsed_invoice.invoice_total is not None:
        adjustment = (parsed_invoice.invoice_total - base_total).quantize(Decimal("0.01"))
        summary = apply_home_office_adjustment(
            summary,
            adjustment,
            license_name="Integricom Support Invoice Adjustment",
            home_office_name=INTEGRICOM_HOME_OFFICE,
        )
        reconciliation = {
            "base_total": float(base_total.quantize(Decimal("0.01"))),
            "invoice_total": float(parsed_invoice.invoice_total),
            "home_office_adjustment": float(adjustment),
        }

    needs_support_review = any(row["needs_review"] for row in support_rows)
    message = None
    if needs_support_review:
        message = (
            "Some support blocks were defaulted to Home Office with low confidence. "
            "Review the branch column, then analyze again."
        )

    breakdown_csv = summary_to_csv(summary)
    grand_total = round(sum(item["total_amount"] for item in summary), 2)

    return {
        "vendor_type": "integricom_support",
        "needs_user_enrichment": False,
        "needs_non_user_branch_assignment": False,
        "needs_support_review": needs_support_review,
        "message": message,
        "new_users": [],
        "adobe_user_rows": [],
        "user_rows": [],
        "non_user_rows": [],
        "non_user_branch_prompts": [],
        "support_rows": support_rows,
        "integricom_support_rows": support_rows,
        "support_branch_options": INTEGRICOM_KNOWN_BRANCHES,
        "missing_users": [],
        "invoice": {
            "filename": invoice_file.filename,
            "size_bytes": len(invoice_raw),
            "invoice_number": parsed_invoice.invoice_number,
            "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total else None,
            "billable_blocks": len(parsed_invoice.blocks),
        },
        "files": [],
        "summary": summary,
        "totals": {
            "line_items": len(summary),
            "grand_total": grand_total,
        },
        "reconciliation": reconciliation,
        "warnings": warnings,
        "breakdown_csv": breakdown_csv,
    }


@app.post("/api/analyze")
async def analyze(
    vendor_type: str = Form(default="generic"),
    csv_files: list[UploadFile] | None = File(default=None),
    invoice_file: UploadFile | None = File(default=None),
    adobe_user_updates: str | None = Form(default=None),
    integricom_user_updates: str | None = Form(default=None),
    integricom_branch_item_updates: str | None = Form(default=None),
    integricom_support_updates: str | None = Form(default=None),
) -> dict:
    uploads = csv_files or []

    vendor = vendor_type.strip().lower()
    if vendor not in {"generic", "hexnode", "adobe", "integricom", "integricom_support"}:
        raise HTTPException(
            status_code=400,
            detail=(
                "Unsupported vendor_type. Use 'generic', 'hexnode', 'adobe', 'integricom', "
                "or 'integricom_support'."
            ),
        )

    if vendor == "adobe":
        if not uploads:
            raise HTTPException(status_code=400, detail="At least one CSV export file is required.")
        return await _analyze_adobe(uploads, invoice_file, adobe_user_updates)
    if vendor == "integricom":
        return await _analyze_integricom(
            uploads,
            invoice_file,
            integricom_user_updates,
            integricom_branch_item_updates,
        )
    if vendor == "integricom_support":
        return await _analyze_integricom_support(invoice_file, integricom_support_updates)

    if not uploads:
        raise HTTPException(status_code=400, detail="At least one CSV export file is required.")

    all_rows = []
    file_summaries: list[dict] = []
    warnings: list[str] = []

    for upload in uploads:
        if not upload.filename:
            continue

        raw = await upload.read()
        if not raw:
            warnings.append(f"{upload.filename}: empty file skipped.")
            continue

        parsed = parse_hexnode_csv(upload.filename, raw) if vendor == "hexnode" else parse_csv(upload.filename, raw)
        all_rows.extend(parsed.rows)
        file_summaries.append(
            {
                "filename": upload.filename,
                "rows_ingested": len(parsed.rows),
                "rows_skipped": parsed.rows_skipped,
            }
        )
        warnings.extend(parsed.warnings)

    summary = build_breakdown(all_rows)
    breakdown_csv = summary_to_csv(summary)
    grand_total = round(sum(item["total_amount"] for item in summary), 2)
    reconciliation = None

    invoice_meta = None
    if invoice_file and invoice_file.filename:
        invoice_raw = await invoice_file.read()
        if vendor == "hexnode":
            parsed_invoice = parse_hexnode_invoice(invoice_file.filename, invoice_raw)
            warnings.extend(parsed_invoice.warnings)
            invoice_meta = {
                "filename": invoice_file.filename,
                "size_bytes": len(invoice_raw),
                "invoice_number": parsed_invoice.invoice_number,
                "invoice_total": float(parsed_invoice.invoice_total) if parsed_invoice.invoice_total is not None else None,
                "billed_device_count": parsed_invoice.billed_device_count,
            }

            if parsed_invoice.billed_device_count is not None and parsed_invoice.billed_device_count != len(all_rows):
                warnings.append(
                    f"Invoice says {parsed_invoice.billed_device_count} devices, but CSV has {len(all_rows)} rows."
                )

            if parsed_invoice.invoice_total is not None:
                base_total = sum(Decimal(str(item["total_amount"])) for item in summary)
                adjustment = (parsed_invoice.invoice_total - base_total).quantize(Decimal("0.01"))
                summary = apply_home_office_adjustment(
                    summary,
                    adjustment,
                    license_name=HEXNODE_DEFAULT_LICENSE,
                )
                breakdown_csv = summary_to_csv(summary)
                grand_total = round(sum(item["total_amount"] for item in summary), 2)
                reconciliation = {
                    "base_total": float(base_total.quantize(Decimal("0.01"))),
                    "invoice_total": float(parsed_invoice.invoice_total),
                    "home_office_adjustment": float(adjustment),
                }
        else:
            invoice_meta = {
                "filename": invoice_file.filename,
                "size_bytes": len(invoice_raw),
                "note": "Invoice uploaded as reference. Generic invoice parsing is not enabled yet.",
            }
    elif vendor == "hexnode":
        warnings.append("No invoice uploaded. Home Office add-on adjustment was not applied.")

    return {
        "vendor_type": vendor,
        "needs_user_enrichment": False,
        "needs_non_user_branch_assignment": False,
        "new_users": [],
        "adobe_user_rows": [],
        "user_rows": [],
        "non_user_rows": [],
        "non_user_branch_prompts": [],
        "missing_users": [],
        "invoice": invoice_meta,
        "files": file_summaries,
        "summary": summary,
        "totals": {
            "line_items": len(summary),
            "grand_total": grand_total,
        },
        "reconciliation": reconciliation,
        "warnings": warnings,
        "breakdown_csv": breakdown_csv,
    }
