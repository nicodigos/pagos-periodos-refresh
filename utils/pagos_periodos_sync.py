import os
import tempfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pymysql

from utils.ms_graph_excel import download_sharepoint_file_bytes, resolve_drive_id
from utils.settings import (
    get_bool_setting,
    get_int_setting,
    get_setting,
    materialize_secret_file,
)

PAGOS_PERIODOS_DEFAULT_WORKBOOK_PATH = (
    "General/12433087 CANADA INC-MASTER/09-Pagos Periodos/2025/"
    "Building & Contractor Pay List/Building Address & Contractor Pay List.xlsx"
)
PAGOS_PERIODOS_CONTRACTOR_SHEET_NAME = "contractor list"
PAGOS_PERIODOS_CONTRACTOR_COLUMNS = [
    {"db": "name", "headers": ["NAME"]},
    {"db": "hourly_rate", "headers": ["HOURLY RATE"]},
    {"db": "type_of_payment", "headers": ["TYPE OF PAYMENT"]},
    {"db": "employee_sub_grouping", "headers": ["EMPLOYEE SUB GROUPING", "EMPLOYEE/SUB GROUPING"]},
    {"db": "category", "headers": ["CATEGORY"]},
]
PAGOS_PERIODOS_BUILDING_SHEET_NAME = "building list"
PAGOS_PERIODOS_BUILDING_COLUMNS = [
    {"db": "name", "headers": ["BUILDING LIST"], "index": 0},
    {"db": "invoicing", "headers": ["INVOICING"], "index": 1},
    {"db": "expense_id", "headers": ["EXPENSE ID"], "index": 2},
    {"db": "building_number", "headers": ["BUILDING NUMBER"], "index": 3},
    {"db": "client", "headers": ["CLIENT"], "index": 4},
    {"db": "vendor", "headers": ["VENDOR"], "index": 5},
    {"db": "province", "headers": ["PROVINCE"], "index": 6},
    {"db": "budget", "headers": ["BUDGET"], "index": 7},
    {"db": "type_of_plan", "headers": ["TYPE OF PLAN", "TYPE OF WORK"], "index": 17},
    {"db": "po", "headers": ["PO"], "index": 9},
    {"db": "icn", "headers": ["ICN"], "index": 10},
]


def pagos_periodos_workbook_path() -> str:
    return str(get_setting("PAGOS_PERIODOS_WORKBOOK_PATH", PAGOS_PERIODOS_DEFAULT_WORKBOOK_PATH)).strip().strip("/")


def normalize_workbook_name(value: str) -> str:
    return " ".join(value.strip().lower().split())


def cell_column_index(ref: str) -> int:
    ref = ref.strip()
    if not ref:
        return -1

    col = 0
    for char in ref:
        if not char.isalpha():
            break
        col = col * 26 + (ord(char.upper()) - ord("A") + 1)
    return col - 1


def decode_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = (cell.attrib.get("t") or "").strip()
    value = (cell.findtext("{*}v") or "").strip()

    if cell_type == "s":
        if value.isdigit():
            idx = int(value)
            if 0 <= idx < len(shared_strings):
                return shared_strings[idx]
        return ""

    if cell_type == "inlineStr":
        return (cell.findtext("{*}is/{*}t") or "").strip()

    return value


def workbook_shared_strings(file_map: dict[str, bytes]) -> list[str]:
    data = file_map.get("xl/sharedStrings.xml")
    if not data:
        return []

    root = ET.fromstring(data)
    out: list[str] = []
    for item in root.findall("{*}si"):
        text = (item.findtext("{*}t") or "").strip()
        if not text:
            text = "".join(run.text or "" for run in item.findall("{*}r/{*}t")).strip()
        out.append(text)
    return out


def workbook_sheet_targets(file_map: dict[str, bytes]) -> dict[str, str]:
    workbook_xml = file_map.get("xl/workbook.xml")
    rels_xml = file_map.get("xl/_rels/workbook.xml.rels")
    if not workbook_xml:
        raise RuntimeError("xl/workbook.xml not found")
    if not rels_xml:
        raise RuntimeError("xl/_rels/workbook.xml.rels not found")

    workbook_root = ET.fromstring(workbook_xml)
    rels_root = ET.fromstring(rels_xml)

    rel_map: dict[str, str] = {}
    for rel in rels_root.findall("{*}Relationship"):
        rel_id = rel.attrib.get("Id", "")
        target = (rel.attrib.get("Target") or "").replace("\\", "/").lstrip("/")
        if target and not target.startswith("xl/"):
            target = f"xl/{target}"
        rel_map[rel_id] = target

    out: dict[str, str] = {}
    for sheet in workbook_root.findall("{*}sheets/{*}sheet"):
        name = normalize_workbook_name(sheet.attrib.get("name", ""))
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_map.get(rel_id)
        if name and target:
            out[name] = target
    return out


def workbook_rows(sheet_xml: bytes, shared_strings: list[str]) -> list[dict[int, str]]:
    root = ET.fromstring(sheet_xml)
    out: list[dict[int, str]] = []

    for row in root.findall("{*}sheetData/{*}row"):
        cells_by_column: dict[int, str] = {}
        for cell in row.findall("{*}c"):
            col_index = cell_column_index(cell.attrib.get("r", ""))
            cells_by_column[col_index] = decode_cell_value(cell, shared_strings).strip()
        out.append(cells_by_column)

    return out


def workbook_table_records(
    sheet_xml: bytes,
    shared_strings: list[str],
    column_specs: list[dict[str, Any]],
    primary_key: str = "name",
    min_header_matches: int = 1,
) -> list[dict[str, str]]:
    rows = workbook_rows(sheet_xml, shared_strings)
    if not rows:
        raise RuntimeError("sheet is empty")

    header_columns: dict[str, int] | None = None
    header_row_index = -1

    for row_index, row in enumerate(rows):
        row_headers = {normalize_workbook_name(value): col_index for col_index, value in row.items() if value.strip()}
        candidate_columns: dict[str, int] = {}
        matched_headers = 0

        for spec in column_specs:
            matched_index = None
            for header in spec.get("headers", []):
                normalized_header = normalize_workbook_name(header)
                if normalized_header in row_headers:
                    matched_index = row_headers[normalized_header]
                    matched_headers += 1
                    break
            if matched_index is None:
                matched_index = spec.get("index")
            if matched_index is None:
                continue
            candidate_columns[spec["db"]] = matched_index

        if matched_headers < min_header_matches:
            continue

        if len(candidate_columns) == len(column_specs):
            header_columns = candidate_columns
            header_row_index = row_index
            break

    if header_columns is None:
        expected = []
        for spec in column_specs:
            headers = spec.get("headers", [])
            if headers:
                expected.append(" / ".join(headers))
        raise RuntimeError(f"sheet headers not found: {', '.join(expected)}")

    records_by_pk: dict[str, dict[str, str]] = {}
    for row_index, row in enumerate(rows):
        if row_index <= header_row_index:
            continue

        record = {db_column: row.get(col_index, "").strip() for db_column, col_index in header_columns.items()}
        pk_value = record.get(primary_key, "").strip()
        if not pk_value:
            continue
        records_by_pk[pk_value] = record

    return sorted(records_by_pk.values(), key=lambda item: normalize_workbook_name(item[primary_key]))


def parse_pagos_periodos_workbook(content: bytes) -> dict[str, list[str]]:
    temp_path = write_temp_file("pagos_periodos.xlsx", content)
    try:
        with ZipFile(temp_path) as archive:
            file_map = {
                entry.filename.replace("\\", "/"): archive.read(entry)
                for entry in archive.infolist()
                if not entry.is_dir()
            }
    finally:
        Path(temp_path).unlink(missing_ok=True)

    sheet_targets = workbook_sheet_targets(file_map)
    shared_strings = workbook_shared_strings(file_map)

    contractor_sheet_path = sheet_targets.get(normalize_workbook_name(PAGOS_PERIODOS_CONTRACTOR_SHEET_NAME))
    building_sheet_path = sheet_targets.get(normalize_workbook_name(PAGOS_PERIODOS_BUILDING_SHEET_NAME))
    if not contractor_sheet_path:
        raise RuntimeError(f'sheet "{PAGOS_PERIODOS_CONTRACTOR_SHEET_NAME}" not found in workbook')
    if not building_sheet_path:
        raise RuntimeError(f'sheet "{PAGOS_PERIODOS_BUILDING_SHEET_NAME}" not found in workbook')

    return {
        "contractors": workbook_table_records(
            file_map[contractor_sheet_path],
            shared_strings,
            PAGOS_PERIODOS_CONTRACTOR_COLUMNS,
        ),
        "buildings": workbook_table_records(
            file_map[building_sheet_path],
            shared_strings,
            PAGOS_PERIODOS_BUILDING_COLUMNS,
            min_header_matches=3,
        ),
    }


def write_temp_file(filename: str, content: bytes) -> str:
    out_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    out_dir.mkdir(exist_ok=True)
    local = out_dir / filename
    local.write_bytes(content)
    return str(local)


def parse_mysql_connection_config() -> dict:
    local_dsn = str(get_setting("LOCAL_MYSQL_DSN", "")).strip()
    env = str(get_setting("ENV", "development")).strip()

    if env == "development" and local_dsn:
        return parse_local_mysql_dsn(local_dsn)

    host = str(get_setting("DB_HOST", "")).strip()
    port = get_int_setting("DB_PORT", 3306)
    user = str(get_setting("DB_USER", "")).strip()
    password = str(get_setting("DB_PASS", ""))
    database = str(get_setting("DB_NAME", "")).strip()
    if not all([host, user, password, database]):
        raise RuntimeError("Missing DB_HOST / DB_PORT / DB_USER / DB_PASS / DB_NAME")

    config = {
        "host": host,
        "port": port,
        "user": user,
        "password": password,
        "database": database,
        "charset": "utf8mb4",
        "autocommit": False,
        "cursorclass": pymysql.cursors.Cursor,
    }

    ca_path = str(get_setting("DO_MYSQL_CA_PATH", "")).strip()
    ca_pem = str(get_setting("DO_MYSQL_CA_PEM", ""))
    if ca_path:
        config["ssl"] = {"ca": ca_path}
    elif ca_pem.strip():
        ca_file = materialize_secret_file("DO_MYSQL_CA_PEM", "cnet_mysql_ca.pem")
        if ca_file:
            config["ssl"] = {"ca": ca_file}

    return config


def is_ssh_tunnel_enabled() -> bool:
    return get_bool_setting("SSH_TUNNEL_ENABLED", False)


def build_ssh_tunnel() -> Any:
    try:
        from sshtunnel import SSHTunnelForwarder
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "SSH tunneling is enabled, but the 'sshtunnel' package is not installed. "
            "Run `venv\\Scripts\\python.exe -m pip install -r requirements.txt`."
        ) from exc

    ssh_host = str(get_setting("SSH_TUNNEL_HOST", "")).strip()
    ssh_port = get_int_setting("SSH_TUNNEL_PORT", 22)
    ssh_user = str(get_setting("SSH_TUNNEL_USER", "")).strip()
    ssh_password = str(get_setting("SSH_TUNNEL_PASSWORD", ""))
    ssh_key_path = str(get_setting("SSH_TUNNEL_KEY_PATH", "")).strip()
    ssh_key_passphrase = str(get_setting("SSH_TUNNEL_KEY_PASSPHRASE", ""))
    ssh_private_key = str(get_setting("SSH_TUNNEL_PRIVATE_KEY", ""))
    remote_host = (
        str(get_setting("SSH_TUNNEL_REMOTE_HOST", "")).strip()
        or str(get_setting("DB_HOST", "")).strip()
        or "127.0.0.1"
    )
    remote_port = get_int_setting(
        "SSH_TUNNEL_REMOTE_PORT",
        get_int_setting("DB_PORT", 3306),
    )
    local_host = str(get_setting("SSH_TUNNEL_LOCAL_HOST", "127.0.0.1")).strip() or "127.0.0.1"

    if not all([ssh_host, ssh_user, remote_host]):
        raise RuntimeError("Missing SSH_TUNNEL_HOST / SSH_TUNNEL_USER / SSH_TUNNEL_REMOTE_HOST")

    tunnel_kwargs = {
        "ssh_address_or_host": (ssh_host, ssh_port),
        "ssh_username": ssh_user,
        "remote_bind_address": (remote_host, remote_port),
        "local_bind_address": (local_host, 0),
    }

    inline_key_path = None
    if ssh_private_key.strip():
        inline_key_path = materialize_secret_file("SSH_TUNNEL_PRIVATE_KEY", "ssh_tunnel_key.pem")

    if inline_key_path:
        tunnel_kwargs["ssh_pkey"] = inline_key_path
        if ssh_key_passphrase:
            tunnel_kwargs["ssh_private_key_password"] = ssh_key_passphrase
    elif ssh_key_path:
        tunnel_kwargs["ssh_pkey"] = ssh_key_path
        if ssh_key_passphrase:
            tunnel_kwargs["ssh_private_key_password"] = ssh_key_passphrase
    elif ssh_password:
        tunnel_kwargs["ssh_password"] = ssh_password
    else:
        raise RuntimeError(
            "Missing SSH_TUNNEL_PASSWORD, SSH_TUNNEL_KEY_PATH, or SSH_TUNNEL_PRIVATE_KEY for SSH tunnel authentication"
        )

    tunnel = SSHTunnelForwarder(**tunnel_kwargs)
    tunnel.start()
    return tunnel


@contextmanager
def mysql_connection():
    config = parse_mysql_connection_config()
    tunnel = None

    try:
        if is_ssh_tunnel_enabled():
            tunnel = build_ssh_tunnel()
            config["host"] = "127.0.0.1"
            config["port"] = int(tunnel.local_bind_port)

        connection = pymysql.connect(**config)
        try:
            yield connection
        finally:
            connection.close()
    finally:
        if tunnel:
            tunnel.stop()


def parse_local_mysql_dsn(dsn: str) -> dict:
    left, _, right = dsn.partition("@")
    if not right:
        raise RuntimeError("Invalid LOCAL_MYSQL_DSN")

    user, _, password = left.partition(":")
    protocol_part, _, db_part = right.partition("/")
    database = db_part.split("?", 1)[0].strip()
    if not user or not database:
        raise RuntimeError("Invalid LOCAL_MYSQL_DSN")

    config = {
        "user": user.strip(),
        "password": password,
        "database": database,
        "charset": "utf8mb4",
        "autocommit": False,
        "cursorclass": pymysql.cursors.Cursor,
    }

    protocol_part = protocol_part.strip()
    if protocol_part.startswith("tcp(") and protocol_part.endswith(")"):
        host_port = protocol_part[4:-1]
        host, _, port = host_port.partition(":")
        config["host"] = host.strip() or "127.0.0.1"
        config["port"] = int((port or "3306").strip())
    else:
        config["host"] = "127.0.0.1"
        config["port"] = 3306

    return config


def replace_lookup_table(
    cursor: pymysql.cursors.Cursor,
    table_name: str,
    columns: list[str],
    rows: list[dict[str, str]],
) -> None:
    cursor.execute(f"DELETE FROM {table_name}")
    if not rows:
        return

    column_names = ", ".join(columns)
    placeholders = ", ".join(["%s"] * len(columns))
    sql = f"INSERT INTO {table_name} ({column_names}) VALUES ({placeholders})"
    values = [tuple(row.get(column, "") for column in columns) for row in rows]
    cursor.executemany(sql, values)


def sync_pagos_periodos_lookup_tables(token: str) -> dict[str, int]:
    drive_id = resolve_drive_id(token)
    workbook_bytes = download_sharepoint_file_bytes(
        pagos_periodos_workbook_path(),
        token,
        drive_id=drive_id,
    )
    workbook_data = parse_pagos_periodos_workbook(workbook_bytes)

    with mysql_connection() as connection:
        try:
            with connection.cursor() as cursor:
                replace_lookup_table(
                    cursor,
                    "pagos_periodos_contractors",
                    [spec["db"] for spec in PAGOS_PERIODOS_CONTRACTOR_COLUMNS],
                    workbook_data["contractors"],
                )
                replace_lookup_table(
                    cursor,
                    "pagos_periodos_buildings",
                    [spec["db"] for spec in PAGOS_PERIODOS_BUILDING_COLUMNS],
                    workbook_data["buildings"],
                )
            connection.commit()
        except Exception:
            connection.rollback()
            raise

    return {
        "contractors": len(workbook_data["contractors"]),
        "buildings": len(workbook_data["buildings"]),
    }
