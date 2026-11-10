#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Analizador de gastos de tarjeta

Lee un archivo TXT con una compra por línea, por ejemplo:

25/05/2026 WOMPI*TEXACO AHUACHAPAN AHUACHAPAN $29.00
31/05/2026 PEDIDOSYA RESTAURANTES -I SAN SALVADOR $12.28

Genera un archivo Excel con:
- Dashboard
- Movimientos
- Resumen comercios
- Resumen mensual
- Resumen categorías
- Líneas no reconocidas, si existen

Uso por consola:
    python analizador_gastos_tarjeta.py compras.txt
    python analizador_gastos_tarjeta.py compras.txt -o historial_gastos.xlsx

Si se ejecuta sin argumentos, abre una ventana para seleccionar los archivos.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
import tempfile
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print(
        "\nFalta la librería openpyxl.\n"
        "Instálala con este comando:\n\n"
        "    python -m pip install openpyxl\n",
        file=sys.stderr,
    )
    raise SystemExit(1)


HEADERS = [
    "ID",
    "Fecha",
    "Mes",
    "Descripción original",
    "Procesador",
    "Comercio",
    "Comercio agrupado",
    "Ubicación",
    "Categoría",
    "Monto",
    "Archivo de origen",
]

# Agrega aquí otras ciudades o municipios que aparezcan al final de tus consumos.
LOCATIONS = sorted(
    {
        "SAN SALVADOR",
        "SANTA TECLA",
        "ANTIGUO CUSCATLAN",
        "NUEVO CUSCATLAN",
        "SOYAPANGO",
        "MEJICANOS",
        "APOPA",
        "ILOPANGO",
        "CUSCATANCINGO",
        "CIUDAD DELGADO",
        "SAN MARCOS",
        "AYUTUXTEPEQUE",
        "TONACATEPEQUE",
        "NEJAPA",
        "AGUILARES",
        "QUEZALTEPEQUE",
        "LA LIBERTAD",
        "LA LIBERTAD COSTA",
        "ZARAGOZA",
        "LOURDES COLON",
        "COLON",
        "AHUACHAPAN",
        "SANTA ANA",
        "SONSONATE",
        "CHALATENANGO",
        "COJUTEPEQUE",
        "CUSCATLAN",
        "ZACATECOLUCA",
        "LA PAZ",
        "SAN VICENTE",
        "USULUTAN",
        "SAN MIGUEL",
        "MORAZAN",
        "SAN FRANCISCO GOTERA",
        "LA UNION",
        "EL SALVADOR",
    },
    key=len,
    reverse=True,
)

# Solo se consideran procesadores cuando aparecen al inicio y separados por *, : o /.
PROCESSORS = (
    "WOMPI",
    "PAGADITO",
    "PAYPAL",
    "NIU",
)

GROUP_RULES = [
    (("PEDIDOSYA",), "PEDIDOSYA"),
    (("TEXACO",), "TEXACO"),
    (("UNO", "ESTACION UNO"), "UNO"),
    (("SHELL",), "SHELL"),
    (("PUMA",), "PUMA"),
    (("UBER EATS", "UBEREATS"), "UBER EATS"),
    (("UBER",), "UBER"),
    (("TEMU","TEMU.COM","DLC TEMU "), "TEMU"),
    (("WALMART",), "WALMART"),
    (("SUPER SELECTOS", "SELECTOS"), "SUPER SELECTOS"),
    (("DESPENSA DE DON JUAN",), "DESPENSA DE DON JUAN"),
    (("PRICE SMART", "PRICESMART"), "PRICESMART"),
    (("SIMAN",), "SIMAN"),
    (("FARMACIAS ECONOMICAS",), "FARMACIAS ECONÓMICAS"),
    (("FARMACIA SAN NICOLAS", "SAN NICOLAS"), "FARMACIA SAN NICOLÁS"),
    (("FARMACIA LA BUENA",), "FARMACIA LA BUENA"),
    (("STARBUCKS",), "STARBUCKS"),
    (("MCDONALD",), "MCDONALD'S"),
    (("BURGER KING",), "BURGER KING"),
    (("KFC",), "KFC"),
    (("PIZZA HUT",), "PIZZA HUT"),
    (("TIGO",), "TIGO"),
    (("CLARO",), "CLARO"),
    (("AES", "CAESS", "CLESA", "EEO", "DEUSEM"), "ENERGÍA ELÉCTRICA"),
]

CATEGORY_RULES = [
    (("TEXACO", "SHELL", "PUMA", "ESTACION UNO", "GASOLINERA","DLC LAS AMAPOLAS"), "Combustible"),
    (("PEDIDOSYA RESTAURANTES","POLLO","CHICHARRON" ,"RESTAURANTE","CHINA WOK" ,"CAFE", "CAFÉ", "PIZZA", "WENDYS","BURGER", "KFC", "MCDONALD", "STARBUCKS","LAS DELICIAS DE LA ABUELA AHUACHAPAN"), "Restaurantes"),
    (("PEDIDOSYA FARMACIAS", "FARMACIA", "HOSPITAL", "CLINICA", "CLÍNICA", "LABORATORIO","FCIA ECO","FV27","FV08","FARMAVALUE"), "Salud y farmacia"),
    (("WALMART", "SUPER SELECTOS", "SELECTOS", "DESPENSA", "PRICESMART", "SUPERMERCADO"), "Supermercado"),
    (("UBER", "TAXI", "INDRIVE", "TRANSPORTE"), "Transporte"),
    (("TIGO", "CLARO", "TELEFON", "INTERNET"), "Telefonía e internet"),
    (("AES", "CAESS", "CLESA", "EEO", "DEUSEM", "ANDA"), "Servicios básicos"),
    (("HOTEL", "AIRBNB", "BOOKING"), "Viajes y alojamiento"),
    (("SIMAN", "ZARA", "BERSHKA", "CLOTHING", "ROPA", "TEMU" , "DLC TEMU"), "Ropa y tiendas"),
    (("NETFLIX", "SPOTIFY", "DISNEY", "HBO", "YOUTUBE"), "Suscripciones"),
    (("MURALLA CHINA"), "Tiendas Chinas"),
    (("EPA","VIDRI","N1CO*VIDRI","FREUND"), "Ferreterias"),
    (("CINEPOLIS"), "Entretenimiento"),


]


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def plain_upper(value: str) -> str:
    value = normalize_text(value).upper()
    return "".join(
        char
        for char in unicodedata.normalize("NFD", value)
        if unicodedata.category(char) != "Mn"
    )


def parse_amount(raw: str) -> float:
    value = raw.replace("$", "").replace(" ", "").strip()

    if "," in value and "." in value:
        # El último separador se interpreta como decimal.
        if value.rfind(".") > value.rfind(","):
            value = value.replace(",", "")
        else:
            value = value.replace(".", "").replace(",", ".")
    elif "," in value:
        parts = value.split(",")
        if len(parts[-1]) == 2:
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")

    return round(float(value), 2)


def extract_location(description: str) -> tuple[str, str]:
    remaining = normalize_text(description)
    found: list[str] = []

    # Permite eliminar una ubicación repetida, por ejemplo:
    # "TEXACO AHUACHAPAN AHUACHAPAN".
    for _ in range(2):
        upper_remaining = plain_upper(remaining)
        match = None

        for location in LOCATIONS:
            loc_plain = plain_upper(location)
            if upper_remaining == loc_plain or upper_remaining.endswith(" " + loc_plain):
                match = location
                break

        if not match:
            break

        found.append(match)
        remaining = remaining[: len(remaining) - len(match)].rstrip(" -_/")

    if not found:
        return remaining, ""

    unique = []
    for location in found:
        if location not in unique:
            unique.append(location)

    return remaining, " / ".join(unique)


def extract_processor(description: str) -> tuple[str, str]:
    value = normalize_text(description)

    for processor in PROCESSORS:
        pattern = rf"^{re.escape(processor)}\s*[\*:/]\s*"
        if re.match(pattern, value, flags=re.IGNORECASE):
            return re.sub(pattern, "", value, count=1, flags=re.IGNORECASE), processor

    return value, ""


def clean_merchant(value: str) -> str:
    merchant = normalize_text(value)

    # Elimina indicadores técnicos como "-I" o "-I-".
    merchant = re.sub(r"\s+-[A-Z0-9]{1,4}-?\s*$", "", merchant, flags=re.IGNORECASE)
    merchant = merchant.strip(" -_/")

    return normalize_text(merchant)


def merchant_group(merchant: str) -> str:
    comparable = plain_upper(merchant)

    for keywords, group in GROUP_RULES:
        if any(plain_upper(keyword) in comparable for keyword in keywords):
            return group

    return merchant


def categorize(merchant: str, group: str) -> str:
    comparable = plain_upper(f"{merchant} {group}")

    for keywords, category in CATEGORY_RULES:
        if any(plain_upper(keyword) in comparable for keyword in keywords):
            return category

    return "Sin categoría"


def make_transaction_id(date_value: datetime, description: str, amount: float) -> str:
    base = f"{date_value:%Y-%m-%d}|{plain_upper(description)}|{amount:.2f}"
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


def parse_line(line: str, source_file: str) -> dict[str, Any]:
    original = normalize_text(line)

    pattern = re.compile(
        r"^\s*(\d{2}/\d{2}/\d{4})\s+(.+?)\s+\$?\s*"
        r"([0-9][0-9.,]*)\s*$"
    )
    match = pattern.match(original)

    if not match:
        raise ValueError("No coincide con fecha + descripción + monto")

    date_text, description, amount_text = match.groups()

    try:
        date_value = datetime.strptime(date_text, "%d/%m/%Y")
    except ValueError as exc:
        raise ValueError("Fecha inválida; use DD/MM/AAAA") from exc

    amount = parse_amount(amount_text)
    without_location, location = extract_location(description)
    without_processor, processor = extract_processor(without_location)
    merchant = clean_merchant(without_processor)

    if not merchant:
        merchant = normalize_text(description)

    group = merchant_group(merchant)
    category = categorize(merchant, group)
    transaction_id = make_transaction_id(date_value, description, amount)

    return {
        "ID": transaction_id,
        "Fecha": date_value,
        "Mes": date_value.strftime("%Y-%m"),
        "Descripción original": normalize_text(description),
        "Procesador": processor,
        "Comercio": merchant,
        "Comercio agrupado": group,
        "Ubicación": location,
        "Categoría": category,
        "Monto": amount,
        "Archivo de origen": source_file,
    }


def read_txt(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = path.read_text(encoding="latin-1")

    transactions: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    occurrences: dict[str, int] = defaultdict(int)

    for line_number, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue

        try:
            item = parse_line(line, path.name)

            # Conserva dos compras idénticas dentro del mismo archivo.
            # El sufijo también permite omitirlas correctamente si se vuelve
            # a procesar exactamente el mismo estado de cuenta.
            base_id = str(item["ID"])
            occurrences[base_id] += 1
            item["ID"] = f"{base_id}-{occurrences[base_id]:03d}"

            transactions.append(item)
        except Exception as exc:
            rejected.append(
                {
                    "Línea": line_number,
                    "Contenido": line.strip(),
                    "Motivo": str(exc),
                    "Archivo": path.name,
                }
            )

    return transactions, rejected


def normalize_loaded_date(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value

    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return datetime(value.year, value.month, value.day)

    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                pass

    raise ValueError(f"No se pudo interpretar la fecha existente: {value!r}")


def load_existing_transactions(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []

    workbook = load_workbook(path, data_only=True, read_only=True)

    if "Movimientos" not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            "El Excel existente no tiene una hoja llamada 'Movimientos'. "
            "Seleccione un archivo creado por este programa o cambie el nombre de salida."
        )

    sheet = workbook["Movimientos"]
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        workbook.close()
        return []

    header_map = {str(value): index for index, value in enumerate(rows[0]) if value is not None}

    missing = [header for header in HEADERS if header not in header_map]
    if missing:
        workbook.close()
        raise ValueError(
            "El Excel existente no tiene el formato esperado. "
            f"Faltan columnas: {', '.join(missing)}"
        )

    result: list[dict[str, Any]] = []

    for row in rows[1:]:
        if not any(value is not None for value in row):
            continue

        item = {header: row[header_map[header]] for header in HEADERS}
        item["Fecha"] = normalize_loaded_date(item["Fecha"])
        item["Monto"] = float(item["Monto"])
        result.append(item)

    workbook.close()
    return result


def merge_transactions(
    existing: Iterable[dict[str, Any]],
    new_items: Iterable[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    merged: dict[str, dict[str, Any]] = {
        str(item["ID"]): item for item in existing
    }

    previous_count = len(merged)

    for item in new_items:
        merged[str(item["ID"])] = item

    transactions = sorted(
        merged.values(),
        key=lambda item: (
            item["Fecha"],
            str(item["Comercio agrupado"]),
            float(item["Monto"]),
        ),
    )

    return transactions, len(merged) - previous_count


def aggregate_sum(
    transactions: Iterable[dict[str, Any]],
    key_name: str,
) -> list[tuple[str, int, float]]:
    values: dict[str, list[float]] = defaultdict(list)

    for item in transactions:
        key = str(item.get(key_name) or "Sin dato")
        values[key].append(float(item["Monto"]))

    return sorted(
        (
            (key, len(amounts), round(sum(amounts), 2))
            for key, amounts in values.items()
        ),
        key=lambda row: (-row[2], row[0]),
    )


def aggregate_months(
    transactions: Iterable[dict[str, Any]],
) -> list[tuple[str, int, float]]:
    values: dict[str, list[float]] = defaultdict(list)

    for item in transactions:
        values[str(item["Mes"])].append(float(item["Monto"]))

    return sorted(
        (
            (month, len(amounts), round(sum(amounts), 2))
            for month, amounts in values.items()
        ),
        key=lambda row: row[0],
    )


def style_header(row_cells) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    border = Border(bottom=Side(style="thin", color="D9E2F3"))

    for cell in row_cells:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border


def add_table(sheet, reference: str, name: str) -> None:
    table = Table(displayName=name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def set_widths(sheet, widths: dict[str, float]) -> None:
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def write_summary_sheet(
    workbook: Workbook,
    title: str,
    rows: list[tuple[str, int, float]],
    table_name: str,
    first_header: str,
) -> None:
    sheet = workbook.create_sheet(title)
    sheet.append([first_header, "Cantidad de compras", "Total gastado"])

    for row in rows:
        sheet.append(list(row))

    style_header(sheet[1])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    set_widths(sheet, {"A": 34, "B": 20, "C": 18})

    for cell in sheet["C"][1:]:
        cell.number_format = '$#,##0.00'

    if sheet.max_row >= 2:
        add_table(sheet, f"A1:C{sheet.max_row}", table_name)
        sheet.conditional_formatting.add(
            f"C2:C{sheet.max_row}",
            DataBarRule(start_type="min", end_type="max", color="5B9BD5"),
        )


def write_workbook(
    output_path: Path,
    transactions: list[dict[str, Any]],
    rejected: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    dashboard = workbook.active
    dashboard.title = "Dashboard"

    movements = workbook.create_sheet("Movimientos")
    movements.append(HEADERS)

    for item in transactions:
        movements.append([item[header] for header in HEADERS])

    style_header(movements[1])
    movements.freeze_panes = "A2"
    movements.auto_filter.ref = movements.dimensions

    set_widths(
        movements,
        {
            "A": 18,
            "B": 13,
            "C": 11,
            "D": 48,
            "E": 15,
            "F": 32,
            "G": 25,
            "H": 22,
            "I": 22,
            "J": 14,
            "K": 26,
        },
    )

    for cell in movements["B"][1:]:
        cell.number_format = "dd/mm/yyyy"

    for cell in movements["J"][1:]:
        cell.number_format = '$#,##0.00'

    if movements.max_row >= 2:
        add_table(movements, f"A1:K{movements.max_row}", "TablaMovimientos")
        movements.conditional_formatting.add(
            f"J2:J{movements.max_row}",
            DataBarRule(start_type="min", end_type="max", color="70AD47"),
        )

    merchant_rows = aggregate_sum(transactions, "Comercio agrupado")
    month_rows = aggregate_months(transactions)
    category_rows = aggregate_sum(transactions, "Categoría")

    write_summary_sheet(
        workbook,
        "Resumen comercios",
        merchant_rows,
        "TablaResumenComercios",
        "Comercio",
    )
    write_summary_sheet(
        workbook,
        "Resumen mensual",
        month_rows,
        "TablaResumenMensual",
        "Mes",
    )
    write_summary_sheet(
        workbook,
        "Resumen categorías",
        category_rows,
        "TablaResumenCategorias",
        "Categoría",
    )

    if rejected:
        rejected_sheet = workbook.create_sheet("Líneas no reconocidas")
        rejected_sheet.append(["Línea", "Contenido", "Motivo", "Archivo"])

        for item in rejected:
            rejected_sheet.append(
                [
                    item["Línea"],
                    item["Contenido"],
                    item["Motivo"],
                    item["Archivo"],
                ]
            )

        style_header(rejected_sheet[1])
        rejected_sheet.freeze_panes = "A2"
        set_widths(rejected_sheet, {"A": 10, "B": 70, "C": 45, "D": 28})
        add_table(
            rejected_sheet,
            f"A1:D{rejected_sheet.max_row}",
            "TablaLineasNoReconocidas",
        )

    build_dashboard(
        dashboard,
        transactions,
        merchant_rows,
        month_rows,
        category_rows,
        workbook,
    )

    # Recalcular fórmulas al abrir el archivo.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Guardado atómico: primero crea un temporal y luego reemplaza el archivo.
    temp_fd, temp_name = tempfile.mkstemp(
        suffix=".xlsx",
        dir=str(output_path.parent),
    )
    os.close(temp_fd)

    try:
        workbook.save(temp_name)
        os.replace(temp_name, output_path)
    finally:
        if os.path.exists(temp_name):
            os.remove(temp_name)


def build_dashboard(
    sheet,
    transactions: list[dict[str, Any]],
    merchant_rows: list[tuple[str, int, float]],
    month_rows: list[tuple[str, int, float]],
    category_rows: list[tuple[str, int, float]],
    workbook: Workbook,
) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells("A1:N2")
    sheet["A1"] = "DASHBOARD DE GASTOS DE TARJETA"
    sheet["A1"].font = Font(size=20, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

    total = round(sum(float(item["Monto"]) for item in transactions), 2)
    count = len(transactions)
    average = round(total / count, 2) if count else 0.0

    latest_month = month_rows[-1][0] if month_rows else "-"
    latest_month_total = month_rows[-1][2] if month_rows else 0.0

    kpis = [
        ("Total histórico", total),
        ("Cantidad de compras", count),
        ("Compra promedio", average),
        (f"Total de {latest_month}", latest_month_total),
    ]

    sheet["A4"] = "INDICADORES"
    sheet["A4"].font = Font(bold=True, color="1F4E78")

    for index, (label, value) in enumerate(kpis, start=5):
        sheet[f"A{index}"] = label
        sheet[f"B{index}"] = value
        sheet[f"A{index}"].font = Font(bold=True)
        sheet[f"A{index}"].fill = PatternFill("solid", fgColor="D9EAF7")
        sheet[f"B{index}"].fill = PatternFill("solid", fgColor="EAF3F8")
        sheet[f"B{index}"].font = Font(size=12, bold=True)

        if index == 6:
            sheet[f"B{index}"].number_format = "0"
        else:
            sheet[f"B{index}"].number_format = '$#,##0.00'

    sheet["A11"] = "TOP 10 COMERCIOS"
    sheet["A11"].font = Font(bold=True, color="1F4E78")
    sheet.append([])

    sheet["A12"] = "Comercio"
    sheet["B12"] = "Compras"
    sheet["C12"] = "Total"
    style_header(sheet[12])

    for row_index, (group, quantity, amount) in enumerate(
        merchant_rows[:10],
        start=13,
    ):
        sheet[f"A{row_index}"] = group
        sheet[f"B{row_index}"] = quantity
        sheet[f"C{row_index}"] = amount
        sheet[f"C{row_index}"].number_format = '$#,##0.00'

    set_widths(
        sheet,
        {
            "A": 29,
            "B": 18,
            "C": 16,
            "D": 4,
            "E": 4,
            "F": 13,
            "G": 13,
            "H": 13,
            "I": 13,
            "J": 13,
            "K": 13,
            "L": 13,
            "M": 13,
            "N": 13,
        },
    )

    if merchant_rows:
        merchant_sheet = workbook["Resumen comercios"]
        max_row = min(merchant_sheet.max_row, 11)

        bar = BarChart()
        bar.type = "bar"
        bar.style = 10
        bar.title = "Comercios con mayor gasto"
        bar.y_axis.title = "Comercio"
        bar.x_axis.title = "Monto"
        bar.height = 8
        bar.width = 15

        data = Reference(merchant_sheet, min_col=3, min_row=1, max_row=max_row)
        categories = Reference(
            merchant_sheet,
            min_col=1,
            min_row=2,
            max_row=max_row,
        )
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(categories)
        sheet.add_chart(bar, "E4")

    if month_rows:
        month_sheet = workbook["Resumen mensual"]

        line = LineChart()
        line.style = 13
        line.title = "Evolución mensual del gasto"
        line.y_axis.title = "Monto"
        line.x_axis.title = "Mes"
        line.height = 7.5
        line.width = 15

        data = Reference(
            month_sheet,
            min_col=3,
            min_row=1,
            max_row=month_sheet.max_row,
        )
        categories = Reference(
            month_sheet,
            min_col=1,
            min_row=2,
            max_row=month_sheet.max_row,
        )
        line.add_data(data, titles_from_data=True)
        line.set_categories(categories)
        sheet.add_chart(line, "E20")

    if category_rows:
        category_sheet = workbook["Resumen categorías"]
        max_row = min(category_sheet.max_row, 8)

        pie = PieChart()
        pie.title = "Distribución por categoría"
        pie.height = 7.5
        pie.width = 9.5

        data = Reference(
            category_sheet,
            min_col=3,
            min_row=1,
            max_row=max_row,
        )
        labels = Reference(
            category_sheet,
            min_col=1,
            min_row=2,
            max_row=max_row,
        )
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        sheet.add_chart(pie, "A25")

    sheet.freeze_panes = "A4"


def process_file(
    input_path: Path,
    output_path: Path,
    keep_history: bool = True,
) -> dict[str, Any]:
    new_transactions, rejected = read_txt(input_path)

    if not new_transactions and not rejected:
        raise ValueError("El archivo TXT está vacío.")

    existing = (
        load_existing_transactions(output_path)
        if keep_history and output_path.exists()
        else []
    )

    transactions, added = merge_transactions(existing, new_transactions)
    write_workbook(output_path, transactions, rejected)

    return {
        "read": len(new_transactions),
        "added": added,
        "duplicates": len(new_transactions) - added,
        "rejected": len(rejected),
        "total_transactions": len(transactions),
        "output": str(output_path),
    }


def run_gui() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        raise SystemExit(
            "Tkinter no está disponible. Use el programa desde la consola."
        )

    root = tk.Tk()
    root.withdraw()

    input_name = filedialog.askopenfilename(
        title="Seleccione el archivo TXT de compras",
        filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")],
    )

    if not input_name:
        return

    default_output = Path(input_name).with_name("historial_gastos.xlsx")

    output_name = filedialog.asksaveasfilename(
        title=(
            "Seleccione el Excel de historial "
            "(puede escoger uno existente para agregar el nuevo mes)"
        ),
        initialdir=str(default_output.parent),
        initialfile=default_output.name,
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
    )

    if not output_name:
        return

    try:
        result = process_file(
            Path(input_name),
            Path(output_name),
            keep_history=True,
        )
    except PermissionError:
        messagebox.showerror(
            "No se pudo guardar",
            "Cierre el archivo Excel si está abierto y vuelva a intentarlo.",
        )
        return
    except Exception as exc:
        messagebox.showerror("Error", str(exc))
        return

    messagebox.showinfo(
        "Proceso completado",
        (
            f"Registros leídos: {result['read']}\n"
            f"Registros nuevos agregados: {result['added']}\n"
            f"Duplicados omitidos: {result['duplicates']}\n"
            f"Líneas no reconocidas: {result['rejected']}\n"
            f"Total del historial: {result['total_transactions']}\n\n"
            f"Archivo generado:\n{result['output']}"
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convierte compras de tarjeta desde TXT a un Excel analítico."
    )
    parser.add_argument(
        "archivo",
        nargs="?",
        help="Archivo TXT que contiene una compra por línea.",
    )
    parser.add_argument(
        "-o",
        "--salida",
        help="Archivo Excel de salida. Predeterminado: historial_gastos.xlsx",
    )
    parser.add_argument(
        "--reemplazar",
        action="store_true",
        help="No conservar el historial anterior; reemplazar el Excel.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if not args.archivo:
        run_gui()
        return

    input_path = Path(args.archivo).expanduser().resolve()

    if not input_path.exists():
        parser.error(f"No existe el archivo: {input_path}")

    output_path = (
        Path(args.salida).expanduser().resolve()
        if args.salida
        else input_path.with_name("historial_gastos.xlsx")
    )

    try:
        result = process_file(
            input_path,
            output_path,
            keep_history=not args.reemplazar,
        )
    except PermissionError:
        raise SystemExit(
            "No se pudo guardar el Excel. Ciérrelo si está abierto y vuelva a intentar."
        )
    except Exception as exc:
        raise SystemExit(f"Error: {exc}")

    print("\nProceso completado")
    print(f"  Registros leídos:          {result['read']}")
    print(f"  Registros nuevos:          {result['added']}")
    print(f"  Duplicados omitidos:       {result['duplicates']}")
    print(f"  Líneas no reconocidas:     {result['rejected']}")
    print(f"  Total del historial:       {result['total_transactions']}")
    print(f"  Excel generado:            {result['output']}")


if __name__ == "__main__":
    main()
