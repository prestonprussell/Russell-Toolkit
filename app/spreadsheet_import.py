from __future__ import annotations

import csv
import io
import re
import zipfile
from dataclasses import dataclass
from typing import Iterable
from xml.etree import ElementTree

ADOBE_HOME_OFFICE = "Home Office"


@dataclass
class AdobeDirectoryImportResult:
    rows: list[dict[str, str]]
    source: str
    warnings: list[str]


def _normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").strip().lower())


def _is_email(value: str) -> bool:
    raw = (value or "").strip()
    return "@" in raw and "." in raw.split("@")[-1]


def _split_csv_rows(raw: bytes) -> Iterable[list[str]]:
    text = raw.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    for row in reader:
        yield [str(cell or "").strip() for cell in row]


def _parse_xlsx_rows(raw: bytes) -> list[list[str]]:
    with zipfile.ZipFile(io.BytesIO(raw)) as workbook:
        names = set(workbook.namelist())
        workbook_xml = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
        ns = {
            "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        rel_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

        rels_xml = ElementTree.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
            for rel in rels_xml.findall("r:Relationship", rel_ns)
        }

        sheets = workbook_xml.find("x:sheets", ns)
        if sheets is None or not list(sheets):
            return []

        target_sheet_rel = ""
        for sheet in sheets:
            name = (sheet.attrib.get("name") or "").strip().lower()
            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            if name == "users":
                target_sheet_rel = rel_id
                break
            if not target_sheet_rel:
                target_sheet_rel = rel_id

        target = rel_map.get(target_sheet_rel, "")
        if not target:
            return []
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        if target not in names:
            return []

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in names:
            sst_xml = ElementTree.fromstring(workbook.read("xl/sharedStrings.xml"))
            for si in sst_xml.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                text = "".join(
                    token.text or ""
                    for token in si.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                )
                shared_strings.append(text.strip())

        def cell_value(cell: ElementTree.Element) -> str:
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                inline = cell.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is")
                if inline is None:
                    return ""
                return "".join(
                    token.text or ""
                    for token in inline.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                ).strip()

            value_node = cell.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
            if value_node is None:
                return ""

            raw_value = (value_node.text or "").strip()
            if cell_type == "s":
                try:
                    index = int(raw_value)
                except ValueError:
                    return ""
                if 0 <= index < len(shared_strings):
                    return shared_strings[index].strip()
                return ""
            return raw_value

        rows: list[list[str]] = []
        sheet_xml = ElementTree.fromstring(workbook.read(target))
        for row in sheet_xml.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
            values = [
                cell_value(cell)
                for cell in row.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c")
            ]
            rows.append(values)
        return rows


def _extract_adobe_rows(rows: Iterable[list[str]], source: str) -> AdobeDirectoryImportResult:
    warnings: list[str] = []
    parsed_rows = list(rows)
    if not parsed_rows:
        return AdobeDirectoryImportResult(rows=[], source=source, warnings=["The file appears to be empty."])

    header_idx = -1
    email_col = -1
    first_col = -1
    last_col = -1
    branch_col = -1

    for idx, row in enumerate(parsed_rows[:30]):
        normalized = [_normalize_header(value) for value in row]
        for col, header in enumerate(normalized):
            if header in {"email", "useremail", "emailaddress"}:
                header_idx = idx
                email_col = col
        if header_idx >= 0:
            first_col = next((i for i, h in enumerate(normalized) if h in {"firstname", "givenname"}), -1)
            last_col = next((i for i, h in enumerate(normalized) if h in {"lastname", "surname"}), -1)
            branch_col = next((i for i, h in enumerate(normalized) if h in {"branch", "office", "officelocation"}), -1)
            break

    if header_idx < 0 or email_col < 0:
        return AdobeDirectoryImportResult(
            rows=[],
            source=source,
            warnings=["Could not find an Email header in the uploaded file."],
        )

    by_email: dict[str, dict[str, str]] = {}
    skipped_invalid_email = 0
    for row in parsed_rows[header_idx + 1 :]:
        if email_col >= len(row):
            continue
        email = (row[email_col] or "").strip().lower()
        if not email:
            continue
        if not _is_email(email):
            skipped_invalid_email += 1
            continue

        first_name = row[first_col].strip() if 0 <= first_col < len(row) else ""
        last_name = row[last_col].strip() if 0 <= last_col < len(row) else ""
        branch = row[branch_col].strip() if 0 <= branch_col < len(row) else ""
        by_email[email] = {
            "email": email,
            "first_name": first_name,
            "last_name": last_name,
            "branch": branch or ADOBE_HOME_OFFICE,
        }

    if skipped_invalid_email:
        warnings.append(f"Skipped {skipped_invalid_email} row(s) with invalid email format.")

    if not by_email:
        warnings.append("No importable Adobe user rows were found.")

    return AdobeDirectoryImportResult(rows=sorted(by_email.values(), key=lambda item: item["email"]), source=source, warnings=warnings)


def parse_adobe_directory_import_file(filename: str, raw: bytes) -> AdobeDirectoryImportResult:
    lowered = (filename or "").strip().lower()
    if lowered.endswith(".csv"):
        return _extract_adobe_rows(_split_csv_rows(raw), source="csv")
    if lowered.endswith(".xlsx") or lowered.endswith(".xlsm"):
        rows = _parse_xlsx_rows(raw)
        return _extract_adobe_rows(rows, source="xlsx")
    return AdobeDirectoryImportResult(
        rows=[],
        source="unsupported",
        warnings=["Unsupported file type. Upload a .xlsx, .xlsm, or .csv file."],
    )
