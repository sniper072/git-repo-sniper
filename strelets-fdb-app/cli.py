#!/usr/bin/env python3
"""CLI helper for Strelets Integral Firebird database."""

from __future__ import annotations

import argparse
import json
import sys

import db


def cmd_tables(_: argparse.Namespace) -> int:
    for table in db.list_tables():
        print(f"{table.group:10} {table.name}")
    return 0


def cmd_schema(args: argparse.Namespace) -> int:
    columns = db.describe_table(args.table)
    pk_cols = db.get_primary_key_columns(args.table)
    print(f"Table: {args.table}")
    if pk_cols:
        print(f"Primary key: {', '.join(pk_cols)}")
    for col in columns:
        null_flag = "NULL" if col.nullable else "NOT NULL"
        print(f"  {col.name:30} {col.field_type:12} {null_flag}")
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    columns, rows, total = db.fetch_rows(args.table, limit=args.limit, offset=args.offset)
    print(f"Table: {args.table} ({total} rows)")
    print("\t".join(columns))
    for row in rows:
        print("\t".join("" if v is None else str(v) for v in row))
    return 0


def cmd_info(_: argparse.Namespace) -> int:
    info = db.test_connection()
    print(json.dumps(info, indent=2))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Strelets Integral FDB CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("tables", help="List user tables").set_defaults(func=cmd_tables)
    sub.add_parser("info", help="Show connection info").set_defaults(func=cmd_info)

    schema = sub.add_parser("schema", help="Show table schema")
    schema.add_argument("table")
    schema.set_defaults(func=cmd_schema)

    show = sub.add_parser("show", help="Show table rows")
    show.add_argument("table")
    show.add_argument("--limit", type=int, default=20)
    show.add_argument("--offset", type=int, default=0)
    show.set_defaults(func=cmd_show)

    args = parser.parse_args()
    try:
        return args.func(args)
    except db.DbError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
