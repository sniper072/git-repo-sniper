"""FastAPI web UI for Strelets Integral Firebird database."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import FastAPI, Form, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import db

BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

app = FastAPI(title="Strelets Integral FDB Browser", version="1.0.0")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


def _format_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray, memoryview)):
        return f"<blob {len(value)} bytes>"
    return str(value)


def _row_pk_values(columns: list[str], row: list[Any], pk_cols: list[str]) -> dict[str, Any]:
    index = {name.upper(): i for i, name in enumerate(columns)}
    return {pk: row[index[pk.upper()]] for pk in pk_cols}


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    try:
        info = db.test_connection()
        tables = db.list_tables()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    grouped: dict[str, list[db.TableInfo]] = {}
    for table in tables:
        grouped.setdefault(table.group, []).append(table)

    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "info": info,
            "grouped": grouped,
            "strelets_user": db.STRELETS_APP_USER,
            "strelets_password": db.STRELETS_APP_PASSWORD,
        },
    )


@app.get("/api/tables")
def api_tables() -> list[dict[str, str]]:
    return [{"name": t.name, "group": t.group} for t in db.list_tables()]


@app.get("/tables/{table_name}", response_class=HTMLResponse)
def table_rows(
    request: Request,
    table_name: str,
    page: int = Query(1, ge=1),
    limit: int = Query(50, ge=1, le=500),
) -> HTMLResponse:
    offset = (page - 1) * limit
    try:
        columns, rows, total = db.fetch_rows(table_name, limit=limit, offset=offset)
        pk_cols = db.get_primary_key_columns(table_name)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    total_pages = max(1, (total + limit - 1) // limit)
    row_items = []
    for row in rows:
        pk = _row_pk_values(columns, row, pk_cols) if pk_cols else {}
        pk_json = json.dumps(pk, default=str)
        row_items.append(
            {
                "cells": [_format_cell(v) for v in row],
                "pk_query": quote(pk_json),
                "pk_json": pk_json,
            }
        )

    return templates.TemplateResponse(
        request,
        "table.html",
        {
            "table_name": table_name,
            "columns": columns,
            "rows": row_items,
            "page": page,
            "limit": limit,
            "total": total,
            "total_pages": total_pages,
            "pk_cols": pk_cols,
            "read_only": db.READ_ONLY,
        },
    )


@app.get("/tables/{table_name}/schema", response_class=HTMLResponse)
def table_schema(request: Request, table_name: str) -> HTMLResponse:
    try:
        columns = db.describe_table(table_name)
        pk_cols = db.get_primary_key_columns(table_name)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return templates.TemplateResponse(
        request,
        "schema.html",
        {"table_name": table_name, "columns": columns, "pk_cols": pk_cols},
    )


@app.get("/tables/{table_name}/new", response_class=HTMLResponse)
def new_row_form(request: Request, table_name: str) -> HTMLResponse:
    if db.READ_ONLY:
        raise HTTPException(status_code=403, detail="Read-only mode is enabled.")
    try:
        columns = db.describe_table(table_name)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return templates.TemplateResponse(
        request,
        "form.html",
        {
            "table_name": table_name,
            "columns": columns,
            "pk_query": None,
            "values": {},
            "action": "create",
        },
    )


@app.post("/tables/{table_name}/create")
async def create_row_post(table_name: str, request: Request) -> RedirectResponse:
    if db.READ_ONLY:
        raise HTTPException(status_code=403, detail="Read-only mode is enabled.")
    body = await request.form()
    data = {k: v for k, v in body.items()}
    try:
        db.insert_row(table_name, data)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return RedirectResponse(url=f"/tables/{table_name}", status_code=303)


@app.get("/tables/{table_name}/edit", response_class=HTMLResponse)
def edit_row_form(request: Request, table_name: str, pk: str) -> HTMLResponse:
    if db.READ_ONLY:
        raise HTTPException(status_code=403, detail="Read-only mode is enabled.")
    pk_values = json.loads(pk)
    try:
        columns_meta = db.describe_table(table_name)
        selected = db.fetch_row_by_pk(table_name, pk_values)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return templates.TemplateResponse(
        request,
        "form.html",
        {
            "table_name": table_name,
            "columns": columns_meta,
            "pk_query": pk,
            "values": selected,
            "action": "edit",
        },
    )


@app.post("/tables/{table_name}/edit")
async def edit_row_post(table_name: str, request: Request, pk: str = Form(...)) -> RedirectResponse:
    if db.READ_ONLY:
        raise HTTPException(status_code=403, detail="Read-only mode is enabled.")
    pk_values = json.loads(pk)
    body = await request.form()
    data = {k: v for k, v in body.items() if k != "pk"}
    try:
        db.update_row(table_name, pk_values, data)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return RedirectResponse(url=f"/tables/{table_name}", status_code=303)


@app.post("/tables/{table_name}/delete")
async def delete_row_post(table_name: str, pk: str = Form(...)) -> RedirectResponse:
    if db.READ_ONLY:
        raise HTTPException(status_code=403, detail="Read-only mode is enabled.")
    pk_values = json.loads(pk)
    try:
        db.delete_row(table_name, pk_values)
    except db.DbError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return RedirectResponse(url=f"/tables/{table_name}", status_code=303)
