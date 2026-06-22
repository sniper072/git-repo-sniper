#!/usr/bin/env python3
"""Generate personnel equipment templates per organizational unit."""

from __future__ import annotations

import argparse
import re
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Column aliases for auto-detection (Russian + English)
COLUMN_ALIASES = {
    "model": ["модель", "model", "наименование", "тип", "устройство"],
    "serial": ["серийный", "serial", "s/n", "sn", "сер. номер", "серийный номер"],
    "inventory": ["инвентар", "inventory", "инв. номер", "инвентарный", "инвентарный номер"],
    "location": ["местоположение", "location", "расположение", "кабинет", "этаж", "адрес"],
    "unit": ["подразделение", "отдел", "unit", "департамент", "служба", "филиал"],
    "employee": ["сотрудник", "фио", "employee", "ответственный", "пользователь"],
    "equipment_name": ["наименование", "оборудование", "название", "device", "принтер"],
}

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TABLE_HEADER_FILL = PatternFill("solid", fgColor="BDD7EE")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def detect_column(columns: list[str], field: str) -> str | None:
    aliases = COLUMN_ALIASES[field]
    for column in columns:
        normalized = normalize_header(column)
        if any(alias in normalized for alias in aliases):
            return column
    return None


def load_txt_context(txt_dir: Path) -> str:
    if not txt_dir.exists():
        return ""

    parts: list[str] = []
    for txt_file in sorted(txt_dir.glob("*.txt")):
        content = txt_file.read_text(encoding="utf-8", errors="replace").strip()
        if content:
            parts.append(content)
    return "\n\n".join(parts)


def read_equipment_workbook(xlsx_path: Path) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(xlsx_path)
    sheets: dict[str, pd.DataFrame] = {}

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, dtype=str)
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue
        df.columns = [str(col).strip() for col in df.columns]
        sheets[sheet_name.strip() or "Без названия"] = df

    return sheets


def map_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    columns = list(df.columns)
    mapping = {
        "model": detect_column(columns, "model"),
        "serial": detect_column(columns, "serial"),
        "inventory": detect_column(columns, "inventory"),
        "location": detect_column(columns, "location"),
        "unit": detect_column(columns, "unit"),
        "employee": detect_column(columns, "employee"),
        "equipment_name": detect_column(columns, "equipment_name"),
    }

    result = pd.DataFrame()
    for field, source in mapping.items():
        if source:
            result[field] = df[source].fillna("").astype(str).str.strip()
        else:
            result[field] = ""

    if result["model"].eq("").all() and not result["equipment_name"].eq("").all():
        result["model"] = result["equipment_name"]

    return result


def split_by_unit(df: pd.DataFrame, fallback_unit: str) -> dict[str, pd.DataFrame]:
    if df["unit"].str.strip().ne("").any():
        grouped: dict[str, pd.DataFrame] = {}
        for unit_name, group in df.groupby(df["unit"].str.strip(), dropna=False):
            key = str(unit_name).strip() or fallback_unit
            grouped[key] = group.reset_index(drop=True)
        return grouped

    return {fallback_unit: df.reset_index(drop=True)}


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", name.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:120] or "unit"


def style_cell(cell, *, bold: bool = False, fill: PatternFill | None = None, wrap: bool = False):
    cell.font = Font(name="Calibri", size=11, bold=bold)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill


def write_template(
    output_path: Path,
    unit_name: str,
    equipment_rows: pd.DataFrame,
    txt_context: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ведомость"

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 20

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    title = ws.cell(row=row, column=1, value="ВЕДОМОСТЬ ЗАКРЕПЛЕНИЯ ОРГТЕХНИКИ ЗА ПОДРАЗДЕЛЕНИЕМ")
    style_cell(title, bold=True, fill=HEADER_FILL)
    title.alignment = Alignment(horizontal="center", vertical="center")

    row += 2
    meta = [
        ("Подразделение:", unit_name),
        ("Дата составления:", date.today().strftime("%d.%m.%Y")),
    ]
    for label, value in meta:
        ws.cell(row=row, column=1, value=label)
        style_cell(ws.cell(row=row, column=1), bold=True)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        ws.cell(row=row, column=2, value=value)
        style_cell(ws.cell(row=row, column=2))
        row += 1

    if txt_context.strip():
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        note = ws.cell(row=row, column=1, value="Дополнительная информация")
        style_cell(note, bold=True, fill=HEADER_FILL)
        row += 1

        paragraphs = [p.strip() for p in txt_context.split("\n\n") if p.strip()]
        for paragraph in paragraphs[:3]:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
            cell = ws.cell(row=row, column=1, value=paragraph)
            style_cell(cell, wrap=True)
            row += 1

    row += 1
    headers = [
        "№",
        "Наименование",
        "Модель",
        "Серийный номер",
        "Инвентарный номер",
        "Местоположение",
        "Ответственный",
    ]

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        style_cell(cell, bold=True, fill=TABLE_HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row += 1
    start_data_row = row

    for index, item in equipment_rows.iterrows():
        name = item.get("equipment_name", "") or item.get("model", "")
        values = [
            index + 1,
            name,
            item.get("model", ""),
            item.get("serial", ""),
            item.get("inventory", ""),
            item.get("location", ""),
            item.get("employee", ""),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            style_cell(cell, wrap=True)
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

    if equipment_rows.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        cell = ws.cell(row=row, column=1, value="Оборудование не найдено")
        style_cell(cell)
        row += 1

    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="Передал: __________________ / __________________")
    style_cell(ws.cell(row=row, column=1))
    ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
    ws.cell(row=row, column=5, value="Принял: __________________ / __________________")
    style_cell(ws.cell(row=row, column=5))

    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    footer = ws.cell(row=row, column=1, value="Подписи. М.П.")
    style_cell(footer)

    ws.freeze_panes = f"A{start_data_row}"
    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].auto_size = False

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def generate_templates(
    xlsx_path: Path,
    txt_dir: Path,
    output_dir: Path,
    sample_only: str | None = None,
) -> list[Path]:
    txt_context = load_txt_context(txt_dir)
    sheets = read_equipment_workbook(xlsx_path)
    created: list[Path] = []

    for sheet_name, raw_df in sheets.items():
        mapped = map_dataframe(raw_df)
        units = split_by_unit(mapped, fallback_unit=sheet_name)

        for unit_name, unit_df in units.items():
            if sample_only and unit_name != sample_only and sheet_name != sample_only:
                continue

            filename = f"Ведомость_{sanitize_filename(unit_name)}.xlsx"
            output_path = output_dir / filename
            write_template(output_path, unit_name, unit_df, txt_context)
            created.append(output_path)

            if sample_only:
                return created

    return created


def build_sample_data() -> tuple[pd.DataFrame, str]:
    equipment = pd.DataFrame(
        [
            {
                "equipment_name": "МФУ",
                "model": "HP LaserJet Pro M428fdw",
                "serial": "VNC3K12345",
                "inventory": "INV-2024-001",
                "location": "3 этаж, каб. 305",
                "employee": "Иванов И.И.",
                "unit": "Отдел продаж",
            },
            {
                "equipment_name": "Принтер",
                "model": "Canon imageRUNNER 2625i",
                "serial": "QWE987654",
                "inventory": "INV-2024-002",
                "location": "3 этаж, каб. 307",
                "employee": "Петрова А.С.",
                "unit": "Отдел продаж",
            },
            {
                "equipment_name": "Сканер",
                "model": "Epson WorkForce DS-770",
                "serial": "EPS112233",
                "inventory": "INV-2024-003",
                "location": "3 этаж, архив",
                "employee": "",
                "unit": "Отдел продаж",
            },
        ]
    )

    txt_context = (
        "Шаблон составлен для закрепления оргтехники за сотрудниками подразделения.\n"
        "Ответственные лица обязаны обеспечить сохранность оборудования и своевременную инвентаризацию."
    )
    return equipment, txt_context


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate personnel equipment templates")
    parser.add_argument(
        "--xlsx",
        type=Path,
        default=Path("equipment-templates/input/Этажи покопийка AI.xlsx"),
        help="Source workbook with printers and office equipment",
    )
    parser.add_argument(
        "--txt-dir",
        type=Path,
        default=Path("equipment-templates/input/txt"),
        help="Directory with Qwen txt context files",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("equipment-templates/output"),
        help="Directory for generated templates",
    )
    parser.add_argument(
        "--sample-only",
        type=str,
        default=None,
        help="Generate only one unit (by unit or sheet name) for approval",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Generate a demo sample without source files",
    )
    args = parser.parse_args()

    if args.demo or not args.xlsx.exists():
        output = args.output_dir / "SAMPLE_Ведомость_для_согласования.xlsx"
        equipment, txt_context = build_sample_data()
        write_template(output, "Отдел продаж (образец)", equipment, txt_context)
        print(f"Created demo sample: {output}")
        if not args.xlsx.exists():
            print(f"Source file not found: {args.xlsx}")
            print("Upload source files to equipment-templates/input/ and rerun without --demo")
        return

    created = generate_templates(args.xlsx, args.txt_dir, args.output_dir, args.sample_only)
    if not created:
        raise SystemExit("No templates were generated. Check workbook sheets and columns.")

    for path in created:
        print(f"Created: {path}")


if __name__ == "__main__":
    main()
