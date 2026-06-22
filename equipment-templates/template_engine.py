"""Load Etagi txt templates and render them into Excel without splitting text."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

TEMPLATE_DIR = Path("equipment-templates/input/templates")

TEMPLATE_FILES = {
    "card": ["Etagi_equipment_Card.txt", "Etagi_equipment_card.txt"],
    "checklist": ["Etagi_checklist_template.txt"],
    "daily_report": ["Etagi_daily_report.txt", "Etagi_ daily_report.txt"],
}


def find_template(kind: str, template_dir: Path = TEMPLATE_DIR) -> Path:
    for name in TEMPLATE_FILES[kind]:
        path = template_dir / name
        if path.exists():
            return path
    matches = sorted(template_dir.glob(f"*{kind}*.txt")) if kind != "card" else sorted(
        template_dir.glob("*Card*.txt")
    )
    if matches:
        return matches[0]
    expected = ", ".join(TEMPLATE_FILES[kind])
    raise FileNotFoundError(
        f"Template not found for '{kind}' in {template_dir}.\n"
        f"Upload one of: {expected}"
    )


def load_template(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n")


def safe_text(value: object) -> str:
    text = str(value or "").strip()
    return text if text else ""


def build_context(device: pd.Series) -> dict[str, str]:
    values = {
        "date": date.today().strftime("%d.%m.%Y"),
        "row_number": safe_text(device.get("row_number", "")),
        "inventory": safe_text(device.get("inventory", "")),
        "model": safe_text(device.get("model", "")),
        "serial": safe_text(device.get("serial", "")),
        "location": safe_text(device.get("location", "")),
        "office": safe_text(device.get("office", "")),
        "unit": safe_text(device.get("unit", "")),
        "employee": safe_text(device.get("employee", "")),
        "prev_counter": safe_text(device.get("prev_counter", "")),
        "report_counter": safe_text(device.get("report_counter", "")),
        "print_count": safe_text(device.get("print_count", "")),
    }

    context: dict[str, str] = {}
    for key, value in values.items():
        context[key] = value
        context[key.upper()] = value
        context[key.replace("_", " ")] = value
        context[key.replace("_", "-")] = value

    russian = {
        "дата": values["date"],
        "номер": values["row_number"],
        "инвентарный_номер": values["inventory"],
        "инвентарный номер": values["inventory"],
        "модель": values["model"],
        "серийный_номер": values["serial"],
        "серийный номер": values["serial"],
        "местоположение": values["location"],
        "место_размещения": values["location"],
        "офис": values["office"],
        "подразделение": values["unit"],
        "ответственный": values["employee"],
        "счетчик_пред": values["prev_counter"],
        "счетчик_отч": values["report_counter"],
        "отпечатки": values["print_count"],
    }
    context.update(russian)
    return context


def render_template(template_text: str, device: pd.Series) -> str:
    context = build_context(device)
    rendered = template_text

    for key, value in sorted(context.items(), key=lambda item: len(item[0]), reverse=True):
        rendered = rendered.replace(f"{{{{{key}}}}}", value)
        rendered = rendered.replace(f"{{{key}}}", value)
        rendered = rendered.replace(f"<<{key}>>", value)
        rendered = rendered.replace(f"[{key}]", value)

    return rendered


def validate_xlsx(path: Path) -> None:
    workbook = load_workbook(path, read_only=True, data_only=True)
    workbook.close()


def write_template_document(output_path: Path, sheet_name: str, rendered_text: str) -> Path:
    """Write the full template text line-by-line (no splitting of template content)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name[:31]
    worksheet.column_dimensions["A"].width = 110

    text_font = Font(name="Calibri", size=11)
    wrap = Alignment(wrap_text=True, vertical="top")

    lines = rendered_text.split("\n")
    if not lines:
        lines = [""]

    for row_index, line in enumerate(lines, start=1):
        cell = worksheet.cell(row=row_index, column=1, value=line)
        cell.font = text_font
        cell.alignment = wrap

    workbook.save(output_path)
    validate_xlsx(output_path)
    return output_path


def write_template_bundle(output_path: Path, documents: dict[str, str]) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)

    text_font = Font(name="Calibri", size=11)
    wrap = Alignment(wrap_text=True, vertical="top")

    for sheet_name, rendered_text in documents.items():
        worksheet = workbook.create_sheet(title=sheet_name[:31])
        worksheet.column_dimensions["A"].width = 110
        for row_index, line in enumerate(rendered_text.split("\n"), start=1):
            cell = worksheet.cell(row=row_index, column=1, value=line)
            cell.font = text_font
            cell.alignment = wrap

    workbook.save(output_path)
    validate_xlsx(output_path)
    return output_path
