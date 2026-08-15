"""Firebird database access layer for Strelets Integral INTEGRAL.FDB."""

from __future__ import annotations

import os
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Any, Iterator

from dotenv import load_dotenv
from firebird.driver import connect

load_dotenv()

FDB_PATH = os.getenv("FDB_PATH", "/workspace/strelets-fdb-app/data/INTEGRAL.FDB")
FDB_USER = os.getenv("FDB_USER", "SYSDBA")
FDB_PASSWORD = os.getenv("FDB_PASSWORD", "masterkey")
FDB_HOST = os.getenv("FDB_HOST", "localhost")
READ_ONLY = os.getenv("READ_ONLY", "false").lower() in {"1", "true", "yes"}

STRELETS_APP_USER = os.getenv("STRELETS_APP_USER", "2047")
STRELETS_APP_PASSWORD = os.getenv("STRELETS_APP_PASSWORD", "1111")

TABLE_GROUPS = {
    "events": {
        "INT_EVENTS",
        "INT_EVENT_PARSERS",
        "INT_EVENT_SCHEME",
        "EVENTS_DEVICE",
    },
    "config": {
        "INT_CONFIGURATION",
        "INT_CONFIGURATIONDATA",
        "INT_SYSTEMCONFIG",
        "INT_HOSTCONFIG",
        "INT_SCHEMAS",
        "INT_SCHEMAS_LINKS",
        "INT_SCHEMAS_GEO",
    },
    "devices": {
        "INT_STATES",
        "INT_STATES2",
        "INT_STATES3",
        "INT_CURRENT_STATES",
        "PARTITIONS",
        "SEGMENTS",
    },
    "users": {"USERS"},
    "graphics": {
        "INT_GRAPH",
        "INT_GRAPH_POSITION",
        "INT_GRAPH_HISTORY",
        "INT_GEODATA",
    },
}


class DbError(Exception):
    """Raised when database operations fail."""


@dataclass
class ColumnInfo:
    name: str
    field_type: str
    nullable: bool
    position: int


@dataclass
class TableInfo:
    name: str
    group: str


def _dsn() -> str:
    return f"{FDB_HOST}:{FDB_PATH}"


@contextmanager
def get_connection() -> Iterator[Any]:
    conn = connect(_dsn(), user=FDB_USER, password=FDB_PASSWORD, charset="UTF8")
    try:
        yield conn
    finally:
        conn.close()


def _table_group(name: str) -> str:
    upper = name.upper()
    for group, names in TABLE_GROUPS.items():
        if upper in names:
            return group
    return "other"


def _quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def list_tables() -> list[TableInfo]:
    sql = """
        SELECT TRIM(RDB$RELATION_NAME)
        FROM RDB$RELATIONS
        WHERE COALESCE(RDB$SYSTEM_FLAG, 0) = 0
          AND RDB$VIEW_BLR IS NULL
        ORDER BY 1
    """
    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
    return [TableInfo(name=row[0], group=_table_group(row[0])) for row in rows]


def describe_table(table: str) -> list[ColumnInfo]:
    _assert_known_table(table)
    sql = """
        SELECT
            TRIM(rf.RDB$FIELD_NAME) AS field_name,
            TRIM(f.RDB$FIELD_TYPE) AS field_type,
            COALESCE(rf.RDB$NULL_FLAG, 0) AS null_flag,
            rf.RDB$FIELD_POSITION
        FROM RDB$RELATION_FIELDS rf
        JOIN RDB$FIELDS f ON f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE
        WHERE rf.RDB$RELATION_NAME = ?
        ORDER BY rf.RDB$FIELD_POSITION
    """
    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, (table.upper(),))
        rows = cur.fetchall()

    type_map = {
        7: "SMALLINT",
        8: "INTEGER",
        10: "FLOAT",
        12: "DATE",
        13: "TIME",
        14: "CHAR",
        16: "BIGINT",
        27: "DOUBLE",
        35: "TIMESTAMP",
        37: "VARCHAR",
        261: "BLOB",
    }
    columns: list[ColumnInfo] = []
    for name, field_type, null_flag, position in rows:
        columns.append(
            ColumnInfo(
                name=name,
                field_type=type_map.get(field_type, str(field_type)),
                nullable=null_flag == 0,
                position=position,
            )
        )
    return columns


def get_primary_key_columns(table: str) -> list[str]:
    _assert_known_table(table)
    sql = """
        SELECT TRIM(s.RDB$FIELD_NAME)
        FROM RDB$RELATION_CONSTRAINTS rc
        JOIN RDB$INDEX_SEGMENTS s ON s.RDB$INDEX_NAME = rc.RDB$INDEX_NAME
        WHERE rc.RDB$RELATION_NAME = ?
          AND rc.RDB$CONSTRAINT_TYPE = 'PRIMARY KEY'
        ORDER BY s.RDB$FIELD_POSITION
    """
    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, (table.upper(),))
        rows = cur.fetchall()
    return [row[0] for row in rows]


def fetch_row_by_pk(table: str, pk_values: dict[str, Any]) -> dict[str, Any]:
    _assert_known_table(table)
    pk_cols = get_primary_key_columns(table)
    if not pk_cols:
        raise DbError(f"Table {table} has no primary key.")

    missing = [col for col in pk_cols if col not in pk_values]
    if missing:
        raise DbError(f"Missing primary key values: {', '.join(missing)}")

    table_q = _quote_ident(table.upper())
    where_clause = " AND ".join(f"{_quote_ident(pk)} = ?" for pk in pk_cols)
    values = [_normalize_value(pk_values[pk]) for pk in pk_cols]
    sql = f"SELECT * FROM {table_q} WHERE {where_clause}"

    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, values)
        row = cur.fetchone()
        if row is None:
            raise DbError("Row not found.")
        columns = [desc[0].strip() for desc in cur.description]
    return {columns[i]: row[i] for i in range(len(columns))}


def fetch_rows(table: str, limit: int = 50, offset: int = 0) -> tuple[list[str], list[list[Any]], int]:
    _assert_known_table(table)
    limit = max(1, min(limit, 500))
    offset = max(0, offset)
    table_q = _quote_ident(table.upper())

    count_sql = f"SELECT COUNT(*) FROM {table_q}"
    data_sql = f"SELECT FIRST {limit} SKIP {offset} * FROM {table_q}"

    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(count_sql)
        total = int(cur.fetchone()[0])
        cur.execute(data_sql)
        columns = [desc[0].strip() for desc in cur.description]
        rows = [list(row) for row in cur.fetchall()]
    return columns, rows, total


def _assert_writable() -> None:
    if READ_ONLY:
        raise DbError("Database is in read-only mode (READ_ONLY=true).")


def _assert_known_table(table: str) -> None:
    known = {t.name.upper() for t in list_tables()}
    if table.upper() not in known:
        raise DbError(f"Unknown table: {table}")


def _normalize_value(value: Any) -> Any:
    if value == "":
        return None
    return value


def insert_row(table: str, data: dict[str, Any]) -> None:
    _assert_writable()
    _assert_known_table(table)
    columns = describe_table(table)
    allowed = {c.name.upper() for c in columns}
    fields = [k for k in data.keys() if k.upper() in allowed and data[k] not in (None, "")]
    if not fields:
        raise DbError("No valid columns provided for insert.")

    table_q = _quote_ident(table.upper())
    cols_q = ", ".join(_quote_ident(f) for f in fields)
    placeholders = ", ".join("?" for _ in fields)
    values = [_normalize_value(data[f]) for f in fields]
    sql = f"INSERT INTO {table_q} ({cols_q}) VALUES ({placeholders})"

    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, values)
        conn.commit()


def update_row(table: str, pk_values: dict[str, Any], data: dict[str, Any]) -> None:
    _assert_writable()
    _assert_known_table(table)
    pk_cols = get_primary_key_columns(table)
    if not pk_cols:
        raise DbError(f"Table {table} has no primary key; update is not supported.")

    columns = describe_table(table)
    allowed = {c.name.upper() for c in columns}
    set_fields = [k for k in data.keys() if k.upper() in allowed and k.upper() not in {p.upper() for p in pk_cols}]
    if not set_fields:
        raise DbError("No updatable columns provided.")

    table_q = _quote_ident(table.upper())
    set_clause = ", ".join(f"{_quote_ident(f)} = ?" for f in set_fields)
    where_clause = " AND ".join(f"{_quote_ident(pk)} = ?" for pk in pk_cols)
    values = [_normalize_value(data[f]) for f in set_fields]
    values.extend(_normalize_value(pk_values[pk]) for pk in pk_cols)
    sql = f"UPDATE {table_q} SET {set_clause} WHERE {where_clause}"

    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, values)
        if cur.rowcount == 0:
            raise DbError("No row matched the primary key.")
        conn.commit()


def delete_row(table: str, pk_values: dict[str, Any]) -> None:
    _assert_writable()
    _assert_known_table(table)
    pk_cols = get_primary_key_columns(table)
    if not pk_cols:
        raise DbError(f"Table {table} has no primary key; delete is not supported.")

    table_q = _quote_ident(table.upper())
    where_clause = " AND ".join(f"{_quote_ident(pk)} = ?" for pk in pk_cols)
    values = [_normalize_value(pk_values[pk]) for pk in pk_cols]
    sql = f"DELETE FROM {table_q} WHERE {where_clause}"

    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, values)
        if cur.rowcount == 0:
            raise DbError("No row matched the primary key.")
        conn.commit()


def test_connection() -> dict[str, Any]:
    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM RDB$RELATIONS WHERE COALESCE(RDB$SYSTEM_FLAG, 0) = 0")
        count = int(cur.fetchone()[0])
    return {
        "dsn": _dsn(),
        "user": FDB_USER,
        "tables": count,
        "read_only": READ_ONLY,
    }
