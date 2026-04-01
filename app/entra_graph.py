from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Any

from .processing import (
    INTEGRICOM_LICENSE_BP,
    INTEGRICOM_LICENSE_F3,
    INTEGRICOM_LICENSE_P1,
    INTEGRICOM_LICENSE_P2,
    INTEGRICOM_LICENSE_TEAMS_ESSENTIALS,
    _normalize_integricom_branch,
)


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

SKU_PART_TO_INTEGRICOM_LICENSE: dict[str, str] = {
    # Microsoft 365 Business Premium variants
    "SPB": INTEGRICOM_LICENSE_BP,
    "O365_BUSINESS_PREMIUM": INTEGRICOM_LICENSE_BP,
    "SMB_BUSINESS_PREMIUM": INTEGRICOM_LICENSE_BP,
    # Exchange Online Plan 1 variants
    "EXCHANGESTANDARD": INTEGRICOM_LICENSE_P1,
    "EXCHANGE_S_STANDARD": INTEGRICOM_LICENSE_P1,
    # Exchange Online Plan 2
    "EXCHANGEENTERPRISE": INTEGRICOM_LICENSE_P2,
    # Microsoft 365 F3
    "SPE_F3": INTEGRICOM_LICENSE_F3,
    # Teams Essentials variants
    "TEAMS_ESSENTIALS_AAD": INTEGRICOM_LICENSE_TEAMS_ESSENTIALS,
    "TEAMS_ESSENTIALS": INTEGRICOM_LICENSE_TEAMS_ESSENTIALS,
}


class EntraSyncError(RuntimeError):
    pass


@dataclass
class EntraIntegricomSyncResult:
    users: list[dict[str, str]]
    export_users: list["EntraIntegricomExportUser"]
    users_scanned: int
    users_skipped_external: int
    users_skipped_unlicensed: int
    users_with_supported_licenses: int
    unknown_sku_parts: list[str]
    warnings: list[str]


@dataclass
class EntraIntegricomExportUser:
    email: str
    first_name: str
    last_name: str
    office: str
    default_branch: str
    licenses: list[str]


def _canonical_integricom_license_from_sku_part(sku_part: str | None) -> str | None:
    normalized = (sku_part or "").strip().upper()
    if not normalized:
        return None
    return SKU_PART_TO_INTEGRICOM_LICENSE.get(normalized)


def _read_required_env() -> tuple[str, str, str]:
    tenant_id = (os.getenv("ENTRA_TENANT_ID") or "").strip()
    client_id = (os.getenv("ENTRA_CLIENT_ID") or "").strip()
    client_secret = (os.getenv("ENTRA_CLIENT_SECRET") or "").strip()
    if not tenant_id or not client_id or not client_secret:
        raise EntraSyncError(
            "Missing Entra configuration. Set ENTRA_TENANT_ID, ENTRA_CLIENT_ID, and ENTRA_CLIENT_SECRET."
        )
    return tenant_id, client_id, client_secret


def _json_request(
    method: str,
    url: str,
    *,
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
) -> dict[str, Any]:
    request = urllib.request.Request(url=url, method=method, headers=headers or {}, data=body)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise EntraSyncError(f"Graph request failed ({exc.code}): {details}") from exc
    except urllib.error.URLError as exc:
        raise EntraSyncError(f"Graph request error: {exc.reason}") from exc
    except Exception as exc:  # pragma: no cover - unexpected transport issues
        raise EntraSyncError(f"Unexpected Graph request error: {exc}") from exc

    try:
        return json.loads(payload)
    except json.JSONDecodeError as exc:
        raise EntraSyncError("Graph response was not valid JSON.") from exc


def _acquire_graph_access_token() -> str:
    tenant_id, client_id, client_secret = _read_required_env()
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    form = urllib.parse.urlencode(
        {
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": GRAPH_SCOPE,
            "grant_type": "client_credentials",
        }
    ).encode("utf-8")
    payload = _json_request(
        "POST",
        token_url,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=form,
    )
    token = payload.get("access_token")
    if not isinstance(token, str) or not token.strip():
        raise EntraSyncError("Graph token response did not include an access token.")
    return token


def _graph_get_paginated(url: str, access_token: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    current_url: str | None = url
    while current_url:
        payload = _json_request(
            "GET",
            current_url,
            headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json"},
        )
        values = payload.get("value")
        if isinstance(values, list):
            rows.extend([item for item in values if isinstance(item, dict)])
        next_link = payload.get("@odata.nextLink")
        current_url = next_link if isinstance(next_link, str) and next_link else None
    return rows


def _get_subscribed_sku_map(access_token: str) -> dict[str, str]:
    rows = _graph_get_paginated(f"{GRAPH_BASE_URL}/subscribedSkus?$select=skuId,skuPartNumber", access_token)
    mapping: dict[str, str] = {}
    for row in rows:
        sku_id = str(row.get("skuId") or "").strip().lower()
        sku_part = str(row.get("skuPartNumber") or "").strip()
        if sku_id and sku_part:
            mapping[sku_id] = sku_part
    return mapping


def sync_integricom_users_from_entra() -> EntraIntegricomSyncResult:
    access_token = _acquire_graph_access_token()
    sku_id_to_part = _get_subscribed_sku_map(access_token)
    users_url = (
        f"{GRAPH_BASE_URL}/users"
        "?$select=givenName,surname,userPrincipalName,mail,officeLocation,department,assignedLicenses"
        "&$top=999"
    )
    graph_users = _graph_get_paginated(users_url, access_token)

    rows_to_upsert: list[dict[str, str]] = []
    export_users: list[EntraIntegricomExportUser] = []
    users_scanned = 0
    skipped_external = 0
    skipped_unlicensed = 0
    supported_users = 0
    unknown_sku_parts: set[str] = set()
    warnings: list[str] = []

    for user in graph_users:
        users_scanned += 1
        email = str(user.get("userPrincipalName") or user.get("mail") or "").strip().lower()
        if not email:
            skipped_unlicensed += 1
            continue
        if "#ext#" in email:
            skipped_external += 1
            continue

        assigned = user.get("assignedLicenses")
        if not isinstance(assigned, list) or not assigned:
            skipped_unlicensed += 1
            continue

        canonical_licenses: set[str] = set()
        for entry in assigned:
            if not isinstance(entry, dict):
                continue
            sku_id = str(entry.get("skuId") or "").strip().lower()
            if not sku_id:
                continue
            sku_part = sku_id_to_part.get(sku_id, "")
            canonical = _canonical_integricom_license_from_sku_part(sku_part)
            if canonical:
                canonical_licenses.add(canonical)
            elif sku_part:
                unknown_sku_parts.add(sku_part)

        if not canonical_licenses:
            skipped_unlicensed += 1
            continue

        supported_users += 1
        office = str(user.get("officeLocation") or "").strip()
        department = str(user.get("department") or "").strip()
        branch = _normalize_integricom_branch(office, department)
        rows_to_upsert.append(
            {
                "email": email,
                "first_name": str(user.get("givenName") or "").strip(),
                "last_name": str(user.get("surname") or "").strip(),
                "branch": branch,
            }
        )
        export_users.append(
            EntraIntegricomExportUser(
                email=email,
                first_name=str(user.get("givenName") or "").strip(),
                last_name=str(user.get("surname") or "").strip(),
                office=office,
                default_branch=branch,
                licenses=sorted(canonical_licenses),
            )
        )

    if unknown_sku_parts:
        warnings.append(
            "Unmapped Entra license SKUs were ignored: " + ", ".join(sorted(unknown_sku_parts)[:12])
        )

    return EntraIntegricomSyncResult(
        users=rows_to_upsert,
        export_users=export_users,
        users_scanned=users_scanned,
        users_skipped_external=skipped_external,
        users_skipped_unlicensed=skipped_unlicensed,
        users_with_supported_licenses=supported_users,
        unknown_sku_parts=sorted(unknown_sku_parts),
        warnings=warnings,
    )
