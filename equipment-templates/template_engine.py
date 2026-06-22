"""Load Etagi txt templates and render them into Excel without splitting text."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

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


def parse_rendered_line(line: str) -> tuple[str, str, str, str]:
    """Split one template line into columns: type, field, value, mark."""
    raw = line.rstrip()
    stripped = raw.strip()

    if not stripped:
        return ("", "", "", "")

    if stripped.startswith("════") or stripped.startswith("────") or stripped.startswith("┌") or stripped.startswith("│") or stripped.startswith("└"):
        return ("Заголовок", stripped, "", "")

    if re.match(r"^\d+\.\s+", stripped):
        return ("Раздел", stripped, "", "")

    if re.match(r"^ШАГ\s+\d+", stripped, flags=re.IGNORECASE):
        return ("Шаг", stripped, "", "")

    checkbox_match = re.match(r"^(\[\s*\]|\[.\]|□)\s*(.+)$", stripped)
    if checkbox_match:
        return ("Проверка", checkbox_match.group(2).strip(), "", "")

    if "🟢" in stripped or "🟡" in stripped or "🔴" in stripped:
        label_match = re.match(r"^([🟢🟡🔴]\s*.+?:)\s*(.*)$", stripped)
        if label_match:
            value = label_match.group(2).strip()
            value = "" if re.fullmatch(r"_+\s*шт\.?", value) else value.replace("_", "").strip()
            return ("Поле", label_match.group(1).strip(), value, "")

    if "Дата:" in stripped and "Инженер:" in stripped:
        parts = re.findall(
            r"(Дата|Инженер):\s*([^А-Яа-яЁё]*[A-Za-zА-Яа-яЁё0-9.+:\-\s]*?)(?=\s+(?:Дата|Инженер):|$)",
            stripped,
        )
        if len(parts) >= 2:
            # Return first part only; caller handles multi-field lines via split helper
            pass

    multi_fields = re.findall(
        r"([A-Za-zА-Яа-яЁё0-9 №()/\-]+):\s*([^:]+?)(?=\s{2,}[A-Za-zА-Яа-яЁё0-9 №()/\-]+:|$)",
        stripped,
    )
    if len(multi_fields) > 1:
        label, value = multi_fields[0]
        return ("Поле", label.strip(), clean_value(value), "")

    colon_match = re.match(r"^(.+?):\s*(.*)$", stripped)
    if colon_match:
        label = colon_match.group(1).strip()
        value = clean_value(colon_match.group(2))
        if label.startswith("КАРТОЧКА АППАРАТА №"):
            inventory = label.replace("КАРТОЧКА АППАРАТА №", "").strip()
            if inventory:
                return ("Поле", "КАРТОЧКА АППАРАТА №", inventory, "")
        if label.startswith("Приоритет") or label.startswith("Срок"):
            return ("Поле", label, value, "")
        return ("Поле", label, value, "")

    if stripped.startswith("ПОДПИСИ") or stripped.startswith("Антон:") or stripped.startswith("Проверил:"):
        return ("Подпись", stripped, "", "")

    if stripped.startswith("Комментарий") or stripped.startswith("Краткое описание"):
        return ("Поле", stripped, "", "")

    if set(stripped) <= {"_"}:
        return ("", "", "", "")

    return ("Текст", stripped, "", "")


def clean_value(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    if re.fullmatch(r"_+", text):
        return ""
    text = re.sub(r"^_+\s*", "", text)
    text = re.sub(r"\s*_+$", "", text)
    if text in {"стр.", "шт."}:
        return ""
    return text.strip()


def expand_line_rows(line: str) -> list[tuple[str, str, str, str]]:
    stripped = line.strip()
    if "Дата:" in stripped and "Инженер:" in stripped:
        rows: list[tuple[str, str, str, str]] = []
        for label in ("Дата", "Инженер"):
            match = re.search(rf"{label}:\s*([^:]+?)(?=\s+(?:Дата|Инженер):|$)", stripped)
            if match:
                rows.append(("Поле", label, clean_value(match.group(1)), ""))
        if rows:
            return rows

    if stripped.startswith("КАРТОЧКА АППАРАТА") and "Дата:" in stripped:
        rows = []
        inv_match = re.search(r"КАРТОЧКА АППАРАТА №\s*(\S+)", stripped)
        date_match = re.search(r"Дата:\s*(\S+)", stripped)
        if inv_match:
            rows.append(("Поле", "КАРТОЧКА АППАРАТА №", inv_match.group(1), ""))
        if date_match:
            rows.append(("Поле", "Дата", date_match.group(1), ""))
        if rows:
            return rows

    return [parse_rendered_line(line)]


def rendered_text_to_dataframe(rendered_text: str) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    pending_location = False

    for line in rendered_text.split("\n"):
        stripped = line.strip()

        if pending_location and line.startswith("   ") and stripped and not stripped.startswith("["):
            rows.append({"Тип": "Поле", "Поле": "Локация дня", "Значение": stripped, "Отметка": ""})
            pending_location = False
            continue

        for row_type, field, value, mark in expand_line_rows(line):
            if not any([row_type, field, value, mark]):
                continue
            if row_type == "Раздел" and "ЛОКАЦИЯ ДНЯ" in field.upper():
                pending_location = True
            else:
                pending_location = False
            rows.append(
                {
                    "Тип": row_type,
                    "Поле": field,
                    "Значение": value,
                    "Отметка": mark,
                }
            )

    if not rows:
        rows.append({"Тип": "", "Поле": "", "Значение": "", "Отметка": ""})

    return pd.DataFrame(rows, columns=["Тип", "Поле", "Значение", "Отметка"]).fillna("")


def write_template_document(output_path: Path, sheet_name: str, rendered_text: str) -> Path:
    """Write template as columns: one value per cell for manual filling."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    frame = rendered_text_to_dataframe(rendered_text).astype(str).replace("nan", "")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        worksheet = writer.sheets[sheet_name[:31]]
        worksheet.column_dimensions["A"].width = 14
        worksheet.column_dimensions["B"].width = 48
        worksheet.column_dimensions["C"].width = 28
        worksheet.column_dimensions["D"].width = 14

    validate_xlsx(output_path)
    return output_path


def write_template_bundle(output_path: Path, documents: dict[str, str]) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, rendered_text in documents.items():
            frame = rendered_text_to_dataframe(rendered_text).astype(str).replace("nan", "")
            safe_name = sheet_name[:31]
            frame.to_excel(writer, sheet_name=safe_name, index=False)
            worksheet = writer.sheets[safe_name]
            worksheet.column_dimensions["A"].width = 14
            worksheet.column_dimensions["B"].width = 48
            worksheet.column_dimensions["C"].width = 28
            worksheet.column_dimensions["D"].width = 14

    validate_xlsx(output_path)
    return output_path
