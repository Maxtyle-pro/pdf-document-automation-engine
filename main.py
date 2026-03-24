from __future__ import annotations

import csv
import json
import os
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

DATA_EXTENSIONS = {".csv", ".json"}
TEMPLATE_EXTENSIONS = {".html", ".htm"}
INVOICE_ID_KEYS = (
    "invoice_id",
    "invoiceid",
    "invoice id",
    "invoice",
    "id",
)


def ensure_directories() -> None:
    for directory in (DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def collect_files(directory: Path, allowed_extensions: set[str]) -> list[Path]:
    if not directory.exists():
        return []
    return sorted(
        [
            file_path
            for file_path in directory.iterdir()
            if file_path.is_file() and file_path.suffix.lower() in allowed_extensions
        ],
        key=lambda path: path.name.lower(),
    )


def print_numbered_menu(title: str, options: list[str]) -> None:
    print()
    print(title)
    for index, option in enumerate(options, start=1):
        print(f"{index}. {option}")


def choose_from_menu(
    prompt: str,
    options: list[Any],
    label_getter: Callable[[Any], str],
    show_menu: bool = True,
) -> Any:
    if not options:
        raise ValueError("No options available.")

    if show_menu:
        print_numbered_menu(prompt, [label_getter(option) for option in options])
    else:
        print()
        print(prompt)

    while True:
        choice = input("Введите номер: ").strip()
        if not choice.isdigit():
            print("Нужно ввести номер из списка.")
            continue

        selected_index = int(choice) - 1
        if 0 <= selected_index < len(options):
            return options[selected_index]

        print("Такого номера нет в списке. Попробуйте снова.")


def read_json_file(file_path: Path) -> Any:
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with file_path.open("r", encoding=encoding) as file:
                return json.load(file)
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError("utf-8", b"", 0, 1, f"Could not decode {file_path.name}")


def read_csv_file(file_path: Path) -> list[dict[str, Any]]:
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with file_path.open("r", encoding=encoding, newline="") as file:
                reader = csv.DictReader(file)
                return [dict(row) for row in reader]
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError("utf-8", b"", 0, 1, f"Could not decode {file_path.name}")


def normalize_records(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]

    if isinstance(data, dict):
        if find_invoice_id(data):
            return [data]

        if isinstance(data.get("invoices"), list):
            return [item for item in data["invoices"] if isinstance(item, dict)]

        normalized: list[dict[str, Any]] = []
        for key, value in data.items():
            if isinstance(value, dict):
                record = dict(value)
                record.setdefault("invoice_id", key)
                normalized.append(record)
        if normalized:
            return normalized

    raise ValueError("Формат данных не поддерживается. Ожидается список объектов или словарь счетов.")


def normalize_key(value: str) -> str:
    return value.strip().lower().replace("_", " ")


def find_invoice_id(record: dict[str, Any]) -> str | None:
    normalized_map = {normalize_key(str(key)): value for key, value in record.items()}

    for key in INVOICE_ID_KEYS:
        value = normalized_map.get(normalize_key(key))
        if value not in (None, ""):
            return str(value)

    return None


def build_invoice_map(records: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    invoice_map: dict[str, dict[str, Any]] = {}
    grouped_rows: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)

    for record in records:
        invoice_id = find_invoice_id(record)
        if not invoice_id:
            continue

        items = record.get("items")
        if isinstance(items, list):
            invoice = dict(record)
            invoice["items"] = [item for item in items if isinstance(item, dict)]
            invoice_map[invoice_id] = invoice
        else:
            grouped_rows[invoice_id].append(dict(record))

    for invoice_id, rows in grouped_rows.items():
        if invoice_id in invoice_map:
            continue

        invoice_data = dict(rows[0])
        invoice_data["items"] = rows
        invoice_map[invoice_id] = invoice_data

    return dict(sorted(invoice_map.items(), key=lambda item: item[0]))


def load_data_file(file_path: Path) -> dict[str, dict[str, Any]]:
    if file_path.suffix.lower() == ".csv":
        raw_data = read_csv_file(file_path)
    elif file_path.suffix.lower() == ".json":
        raw_data = read_json_file(file_path)
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {file_path.suffix}")

    records = normalize_records(raw_data)
    invoice_map = build_invoice_map(records)

    if not invoice_map:
        raise ValueError("Не удалось найти ни одного invoice id в выбранном файле.")

    return invoice_map


def find_font_file() -> tuple[str | None, str]:
    candidate_names = ("DejaVuSans.ttf", "Roboto-Regular.ttf", "DejaVu Sans.ttf", "Roboto.ttf")
    search_directories = []

    if sys.platform.startswith("win"):
        search_directories.extend(
            [
                Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
                Path.home() / "AppData/Local/Microsoft/Windows/Fonts",
            ]
        )
    elif sys.platform == "darwin":
        search_directories.extend(
            [
                Path("/System/Library/Fonts"),
                Path("/Library/Fonts"),
                Path.home() / "Library/Fonts",
            ]
        )

    for directory in search_directories:
        if not directory.exists():
            continue
        for candidate_name in candidate_names:
            font_path = directory / candidate_name
            if font_path.exists():
                font_family = "DejaVu Sans" if "dejavu" in candidate_name.lower() else "Roboto"
                return font_path.as_uri(), font_family

    return None, "sans-serif"


def inject_font_css(html_content: str) -> str:
    font_uri, font_family = find_font_file()

    if font_uri:
        font_css = f"""
        <style>
            @font-face {{
                font-family: "{font_family}";
                src: url("{font_uri}");
            }}

            body {{
                font-family: "{font_family}", sans-serif;
            }}
        </style>
        """
    else:
        font_css = """
        <style>
            body {
                font-family: sans-serif;
            }
        </style>
        """
        print("Предупреждение: DejaVu Sans или Roboto не найдены в системе. Используется системный sans-serif.")

    lower_html = html_content.lower()
    head_close_index = lower_html.find("</head>")
    if head_close_index != -1:
        return html_content[:head_close_index] + font_css + html_content[head_close_index:]

    return font_css + html_content


def render_html(template_path: Path, invoice_id: str, invoice_data: dict[str, Any]) -> str:
    environment = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
    )
    template = environment.get_template(template_path.name)
    html_content = template.render(
        invoice_id=invoice_id,
        invoice=invoice_data,
        items=invoice_data.get("items", []),
        rows=invoice_data.get("items", []),
        generated_at=datetime.now(),
    )
    return inject_font_css(html_content)


def generate_pdf(template_path: Path, invoice_id: str, invoice_data: dict[str, Any]) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    safe_invoice_id = "".join(character if character.isalnum() or character in "-_." else "_" for character in invoice_id)
    safe_template_name = template_path.stem.replace(" ", "_")
    pdf_path = OUTPUT_DIR / f"{safe_invoice_id}_{safe_template_name}.pdf"

    html_content = render_html(template_path, invoice_id, invoice_data)
    HTML(string=html_content, base_url=str(template_path.parent)).write_pdf(str(pdf_path))

    return pdf_path


def open_pdf_file(pdf_path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(pdf_path))  # type: ignore[attr-defined]
        return

    if sys.platform == "darwin":
        subprocess.run(["open", str(pdf_path)], check=False)
        return

    subprocess.run(["xdg-open", str(pdf_path)], check=False)


def main() -> None:
    ensure_directories()

    data_files = collect_files(DATA_DIR, DATA_EXTENSIONS)
    template_files = collect_files(TEMPLATES_DIR, TEMPLATE_EXTENSIONS)

    if not data_files:
        print(f"В директории {DATA_DIR} не найдено ни одного CSV или JSON файла.")
        return

    if not template_files:
        print(f"В директории {TEMPLATES_DIR} не найдено ни одного HTML-шаблона.")
        return

    print("Доступные файлы с данными:")
    for index, file_path in enumerate(data_files, start=1):
        print(f"{index}. {file_path.name}")

    print()
    print("Доступные HTML-шаблоны:")
    for index, file_path in enumerate(template_files, start=1):
        print(f"{index}. {file_path.name}")

    selected_data_file = choose_from_menu(
        "Выберите файл с данными:",
        data_files,
        lambda path: path.name,
        show_menu=False,
    )
    selected_template = choose_from_menu(
        "Выберите HTML-шаблон:",
        template_files,
        lambda path: path.name,
        show_menu=False,
    )

    invoice_map = load_data_file(selected_data_file)
    invoice_ids = list(invoice_map.keys())

    print_numbered_menu("Доступные чеки (invoice id):", invoice_ids)
    selected_invoice_id = choose_from_menu(
        "Выберите invoice id:",
        invoice_ids,
        lambda invoice_id: invoice_id,
        show_menu=False,
    )

    pdf_path = generate_pdf(selected_template, selected_invoice_id, invoice_map[selected_invoice_id])

    print()
    print(f"PDF успешно создан: {pdf_path}")
    open_pdf_file(pdf_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print()
        print(f"Ошибка: {error}")
