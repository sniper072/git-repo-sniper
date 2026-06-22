"""Load Etagi txt templates and render them into Excel without splitting text."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

INPUT_DIR = Path("equipment-templates/input")
TEMPLATE_DIR = INPUT_DIR / "templates"

TEMPLATE_FILES = {
    "card": ["Etagi_equipment_Card.txt", "Etagi_equipment_card.txt"],
    "checklist": ["Etagi_checklist_template.txt"],
    "daily_report": ["Etagi_daily_report.txt", "Etagi_ daily_report.txt"],
}

# Labels from Etagi templates -> device field keys
UNDERSCORE_LABELS: list[tuple[str, str]] = [
    ("Модель:", "model"),
    ("Серийный номер:", "serial"),
    ("Инвентарный номер:", "inventory"),
    ("Локация (офис/кабинет):", "location"),
    ("Ответственный у Заказчика:", "employee"),
    ("Общий пробег (Total):", "report_counter"),
    ("Пробег с последнего ТО:", "print_count"),
]


def template_search_dirs(template_dir: Path | None = None) -> list[Path]:
    dirs: list[Path] = []
    if template_dir is not None:
        dirs.append(template_dir)
    dirs.extend([INPUT_DIR, TEMPLATE_DIR])
    unique: list[Path] = []
    for path in dirs:
        if path not in unique:
            unique.append(path)
    return unique


def find_template(kind: str, template_dir: Path | None = None) -> Path:
    for directory in template_search_dirs(template_dir):
        for name in TEMPLATE_FILES[kind]:
            path = directory / name
            if path.exists():
                return path
        if kind == "card":
            matches = sorted(directory.glob("*Card*.txt"))
        elif kind == "checklist":
            matches = sorted(directory.glob("*checklist*.txt"))
        else:
            matches = sorted(directory.glob("*daily*report*.txt"))
        if matches:
            return matches[0]

    expected = ", ".join(TEMPLATE_FILES[kind])
    raise FileNotFoundError(
        f"Template not found for '{kind}'.\n"
        f"Upload one of: {expected}\n"
        f"To: equipment-templates/input/"
    )


def load_template(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n")


def safe_text(value: object) -> str:
    return str(value or "").strip()


def build_context(device: pd.Series) -> dict[str, str]:
    location = safe_text(device.get("location", ""))
    office = safe_text(device.get("office", ""))
    unit = safe_text(device.get("unit", ""))
    if location and office and office not in location:
        location_display = f"{location} / {office}"
    elif location:
        location_display = location
    elif office:
        location_display = office
    else:
        location_display = unit

    values = {
        "date": date.today().strftime("%d.%m.%Y"),
        "row_number": safe_text(device.get("row_number", "")),
        "inventory": safe_text(device.get("inventory", "")),
        "model": safe_text(device.get("model", "")),
        "serial": safe_text(device.get("serial", "")),
        "location": location_display,
        "office": office,
        "unit": unit,
        "employee": safe_text(device.get("employee", "")),
        "prev_counter": safe_text(device.get("prev_counter", "")),
        "report_counter": safe_text(device.get("report_counter", "")),
        "print_count": safe_text(device.get("print_count", "")),
    }

    context: dict[str, str] = {}
    for key, value in values.items():
        context[key] = value
        context[key.upper()] = value

    russian = {
        "дата": values["date"],
        "модель": values["model"],
        "серийный номер": values["serial"],
        "инвентарный номер": values["inventory"],
        "местоположение": values["location"],
    }
    context.update(russian)
    return context


def replace_underscores_after_label(line: str, label: str, value: str) -> str:
    if label not in line or not value:
        return line

    label_index = line.find(label)
    before = line[: label_index + len(label)]
    after = line[label_index + len(label) :]
    match = re.match(r"(\s*)(_+)(.*)$", after)
    if not match:
        return line

    spaces, underscores, suffix = match.groups()
    replacement = value[: len(underscores)].ljust(len(underscores))
    return f"{before}{spaces}{replacement}{suffix}"


def fill_special_patterns(text: str, context: dict[str, str]) -> str:
    text = re.sub(
        r"КАРТОЧКА АППАРАТА №\s*_+",
        f"КАРТОЧКА АППАРАТА № {context['inventory'] or '_______'}",
        text,
        count=1,
    )
    text = re.sub(
        r"(?m)(Дата:\s*)_+",
        rf"\g<1>{context['date']}",
        text,
    )
    text = re.sub(
        r"(?m)^(   )_+\s*$",
        rf"\g<1>{context['location']}",
        text,
        count=1,
    )
    return text


def fill_underscore_fields(text: str, context: dict[str, str]) -> str:
    lines = []
    for line in text.split("\n"):
        updated = line
        for label, field_key in UNDERSCORE_LABELS:
            updated = replace_underscores_after_label(updated, label, context.get(field_key, ""))
        lines.append(updated)
    return "\n".join(lines)


def render_template(template_text: str, device: pd.Series) -> str:
    context = build_context(device)
    rendered = template_text

    for key, value in sorted(context.items(), key=lambda item: len(item[0]), reverse=True):
        rendered = rendered.replace(f"{{{{{key}}}}}", value)
        rendered = rendered.replace(f"{{{key}}}", value)

    rendered = fill_special_patterns(rendered, context)
    rendered = fill_underscore_fields(rendered, context)
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

    for row_index, line in enumerate(rendered_text.split("\n"), start=1):
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
