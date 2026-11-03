import datetime
import hmac
from io import BytesIO
import json
import mimetypes
import os
import tempfile
import threading
from copy import deepcopy
from pathlib import Path
from urllib.parse import urlsplit

import requests
from flask import Flask, jsonify, redirect, render_template, request, send_file, send_from_directory, session, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException


mimetypes.add_type("text/javascript", ".js")

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = Path(
    os.environ.get("LONDON_FOOD_EXCEL", str(BASE_DIR / "londoninc.xlsx"))
).expanduser().resolve()
WORKBOOK_LOCK = threading.RLock()

STAFF = [
    ["Nathan Heilbron Ruff", "Gerente"],
    ["Luis/Lucho", "Administrador"],
    ["Hanny", "Administradora"],
    ["Juand David", "Domicilios y atención"],
    ["Daniela", "Administración y caja"],
    ["Tania", "Cocina"],
    ["Asistente de cocina (por decidir)", "Asistente de cocina"],
    ["Mesero asistente (por decidir)", "Mesero"],
    ["Fercho", "Domiciliario"],
    ["Domiciliario (por decidir)", "Domicilios"],
]
LEGACY_STAFF = [["Nathan", "Administrador"], ["Hanny", "Cocina"], ["Jesus", "Domicilios"]]
PREVIOUS_STAFF = [
    ["Nathan Heilbron Ruff", "Representante legal, dueño y gerente; contaduría y pedidos de materia prima/productos"],
    ["Luis/Lucho", "Administrador de turno y jefe de pedidos en sitio"],
    ["Hanny", "Administradora de turno algunos días; asistente preparadora y contadora algunos días; administradora de London Inc Delicatessen"],
    ["Juand David", "Jefe de despacho de domicilios y mesero"],
    ["Daniela", "Administradora y cajera de pedidos en sitio y domicilios; contaduría y pedidos de materia prima/productos"],
    ["Tania", "Jefa de cocina"],
    ["Asistente de cocina (por decidir)", "Asistente de cocina"],
    ["Mesero asistente (por decidir)", "Mesero asistente"],
    ["Fercho", "Domiciliario"],
    ["Domiciliario (por decidir)", "Domiciliario"],
]

MENU = [
    ["Salchipapas", [["Salchipapa Casera Inglesa", 16000], ["Salchipapa de Chorizo Inglesa", 21000], ["Salchipapa Big Ben (para 4 personas)", 50000]]],
    ["Hamburguesas", [["Hamburguesa de Carne S", 15000], ["Hamburguesa de Carne M", 25000], ["Hamburguesa de Carne L", 38000], ["Hamburguesa de Carne XL", 50000], ["Hamburguesa de Pollo M", 26000], ["Hamburguesa Vegetariana M", 22000]]],
    ["Perros Calientes", [["Perro Sencillo", 8000], ["Perro Ranchero", 12000], ["Perro Suizo", 14000]]],
    ["Desgranados", [["Chuzos Desgranados de Pollo", 21000], ["Mazorcas Desgranadas", 23000]]],
    ["Especialidades de la Casa", [["Brochets Breaded with Chips (Pincho inglés con papas inglesas)", 20000], ["English Fish & Chips (Pescado apanado inglés con papas inglesas)", 30000]]],
    ["Menú Infantil (Little Angels Menu)", [["Mini Pinchos con porción de papas inglesa", 8000], ["Mini Salchipapas", 10000], ["Mini Pescado Inglés con papas inglesa", 4000], ["Hamburguesa de Carne XS", 10000], ["Hamburguesa de Pollo XS", 11000]]],
    ["Adicionales", [["Pinchos de Salchicha y Queso", 8000], ["Papas Inglesa", 8000], ["Papas a la Francesa", 6000], ["Queso Costeño", 2000], ["Queso Mozzarella", 3000]]],
    ["Bebidas", [["Agua Brisa Sabor Manzana", 3000], ["Coca Cola 400 ml Personal", 4000], ["Sprite Personal 400 ml", 4000], ["Coca Cola Zero Personal", 4000], ["Quatro Personal 400 ml", 4000], ["Jugos Hit Personal", 4000], ["Jugos de Sábila Personal", 4000], ["Coca Cola 1.5 L", 4000]]],
]

LEGACY_MENU = [
    ["Salchipapas", [["Salchipapa casera inglesa", 16000], ["Salchipapa de chorizo casera inglesa", 21000], ["Salchipapa Big Ben (4 personas)", 45000]]],
    ["Extras", [["Pinchos de salchicha y queso", 8000], ["Papas inglesas", 8000], ["Papas a la francesa", 6000]]],
    ["Adicional", [["Queso costeño", 2000], ["Queso mozzarella", 3000]]],
    ["Bebidas", [["Agua Brisa manzana", 3000], ["Coca-Cola 400 ml", 4000], ["Sprite personal", 4000], ["Coca-Cola Zero personal", 4000], ["Quatro personal", 4000], ["Jugos Hit personal", 4000], ["Jugo de sábila personal", 4000], ["Coca-Cola 1,7 L", 8000]]],
]


def migrate_legacy_menu(state):
    if state.get("menu") == LEGACY_MENU:
        state["menu"] = deepcopy(MENU)
        return True
    return False


def migrate_legacy_staff(state):
    if state.get("staff") == LEGACY_STAFF or state.get("staff") == PREVIOUS_STAFF:
        state["staff"] = deepcopy(STAFF)
        return True
    return False


class WorkbookLockedError(Exception):
    pass


def workbook_lock_path():
    return EXCEL_PATH.with_name(f"~${EXCEL_PATH.name}")


def is_workbook_locked():
    """Detecta si el archivo de Excel está abierto por el usuario u otro proceso."""
    if not EXCEL_PATH.exists():
        return False
    lock_file = workbook_lock_path()
    if lock_file.exists():
        try:
            with open(lock_file, "r+b"):
                pass
            lock_file.unlink(missing_ok=True)
        except (PermissionError, OSError):
            return True
    try:
        with open(EXCEL_PATH, "r+b"):
            pass
    except (PermissionError, OSError):
        return True
    return False


def format_time_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime.time):
        return value.strftime("%I:%M %p").replace("AM", "a. m.").replace("PM", "p. m.")
    if isinstance(value, datetime.datetime):
        return value.strftime("%I:%M %p").replace("AM", "a. m.").replace("PM", "p. m.")
    return str(value).strip()


def format_date_value(value, default=None):
    if default is None:
        default = datetime.date.today().strftime("%d/%m/%Y")
    if value is None:
        return default
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    return s if s else default


def parse_float(value, default=0.0):
    """Convierte cadenas con símbolos de moneda o formatos numéricos a float de forma segura."""
    if value is None or value == "":
        return float(default)
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace(" ", "")
    if "." in s and "," not in s:
        parts = s.split(".")
        if len(parts[-1]) == 3 and len(parts) > 1:
            s = "".join(parts)
    elif "," in s and "." not in s:
        parts = s.split(",")
        if len(parts[-1]) == 3 and len(parts) > 1:
            s = "".join(parts)
        else:
            s = s.replace(",", ".")
    elif "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return float(default)


def parse_int(value, default=0):
    return int(round(parse_float(value, default)))


def default_state():
    return {
        "tab": "menu",
        "cart": {},
        "f": {"n": "", "address": "", "t": "Mesa", "p": "Efectivo", "comments": ""},
        "orders": [],
        "pays": [],
        "moves": [],
        "open": 100000,
        "url": "https://londoninc.com/menu",
        "menu": deepcopy(MENU),
        "staff": deepcopy(STAFF),
        "balanceItems": [],
        "inv": [["Papa (kg)", 20, 8], ["Salchicha (und)", 80, 30], ["Chorizo (und)", 60, 20], ["Queso costeño (kg)", 3, 2], ["Queso mozzarella (kg)", 2, 2], ["Coca-Cola 400 ml (und)", 24, 10], ["Agua Brisa (und)", 18, 10], ["Aceite (L)", 10, 4]],
    }


def safe_cell(value):
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def set_rows(sheet, headers, rows):
    sheet.delete_rows(1, max(sheet.max_row, 1))
    sheet.append(headers)
    for row in rows:
        sheet.append([safe_cell(value) for value in row])
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B1F4D")
    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions
    else:
        sheet.auto_filter.ref = None
    for column in sheet.columns:
        if not column:
            continue
        letter = column[0].column_letter
        width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 45)
        sheet.column_dimensions[letter].width = width


def sheet_state(workbook):
    state = default_state()
    today_str = datetime.date.today().strftime("%d/%m/%Y")

    if "Menu" in workbook.sheetnames:
        grouped = {}
        for row in workbook["Menu"].iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            category = row[0] if len(row) > 0 else None
            name = row[1] if len(row) > 1 else None
            price = row[2] if len(row) > 2 else 0
            if category and name:
                grouped.setdefault(str(category).strip(), []).append([
                    str(name).strip(),
                    parse_int(price, 0),
                ])
        if grouped:
            state["menu"] = [[category, products] for category, products in grouped.items()]

    if "Pedidos" in workbook.sheetnames:
        state["orders"] = []
        sheet = workbook["Pedidos"]
        rows = list(sheet.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
            has_fecha = "fecha" in header

            for row in rows[1:]:
                if not row or not any(row):
                    continue
                if has_fecha:
                    order_id = row[0] if len(row) > 0 else None
                    fec = format_date_value(row[1] if len(row) > 1 else None, default=today_str)
                    hour = format_time_value(row[2] if len(row) > 2 else "")
                    customer = row[3] if len(row) > 3 else ""
                    order_type = row[4] if len(row) > 4 else "Mesa"
                    payment = row[5] if len(row) > 5 else "Efectivo"
                    items = row[6] if len(row) > 6 else "[]"
                    total = row[7] if len(row) > 7 else 0
                    status = row[8] if len(row) > 8 else "Pendiente"
                    comments_index = header.index("comentarios") if "comentarios" in header else 9
                    comments = row[comments_index] if len(row) > comments_index else ""
                    address_index = header.index("direccion") if "direccion" in header else 10
                    address = row[address_index] if len(row) > address_index else ""
                else:
                    order_id = row[0] if len(row) > 0 else None
                    oid_val = parse_int(order_id)
                    fec = "20/10/2026" if oid_val in (1, 2) else today_str
                    hour = format_time_value(row[1] if len(row) > 1 else "")
                    customer = row[2] if len(row) > 2 else ""
                    order_type = row[3] if len(row) > 3 else "Mesa"
                    payment = row[4] if len(row) > 4 else "Efectivo"
                    items = row[5] if len(row) > 5 else "[]"
                    total = row[6] if len(row) > 6 else 0
                    status = row[7] if len(row) > 7 else "Pendiente"
                    comments = ""
                    address = ""

                if order_id is not None and str(order_id).strip() != "":
                    parsed_items = []
                    if isinstance(items, list):
                        parsed_items = items
                    elif isinstance(items, str) and items.strip():
                        try:
                            parsed_items = json.loads(items)
                        except json.JSONDecodeError:
                            parsed_items = [[items.strip(), 1, parse_int(total)]]
                    state["orders"].append({
                        "id": parse_int(order_id, len(state["orders"]) + 1),
                        "fec": fec,
                        "h": hour,
                        "n": str(customer or "Cliente"),
                        "t": str(order_type or "Mesa"),
                        "p": str(payment or "Efectivo"),
                        "items": parsed_items,
                        "total": parse_float(total, 0.0),
                        "st": str(status or "Pendiente"),
                        "comments": str(comments or ""),
                        "address": str(address or ""),
                    })

    if "Pagos" in workbook.sheetnames:
        state["pays"] = []
        sheet = workbook["Pagos"]
        rows = list(sheet.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
            has_fecha = "fecha" in header
            type_index = header.index("clasificación") if "clasificación" in header else None
            for row in rows[1:]:
                if not row or not any(row):
                    continue
                if has_fecha:
                    fec = format_date_value(row[0] if len(row) > 0 else None, default=today_str)
                    hour = format_time_value(row[1] if len(row) > 1 else "")
                    employee = row[2] if len(row) > 2 else None
                    concept = row[3] if len(row) > 3 else "Pago"
                    amount = row[4] if len(row) > 4 else 0
                else:
                    hour = format_time_value(row[0] if len(row) > 0 else "")
                    employee = row[1] if len(row) > 1 else None
                    concept = row[2] if len(row) > 2 else "Pago"
                    amount = row[3] if len(row) > 3 else 0
                    fec = "20/10/2026" if str(employee).strip() == "Nathan" else today_str

                if employee is not None and str(employee).strip() != "":
                    state["pays"].append({
                        "fec": fec,
                        "h": hour,
                        "n": str(employee).strip(),
                        "c": str(concept or "Pago").strip(),
                        "a": parse_float(amount, 0.0),
                    })

    if "Gastos" in workbook.sheetnames:
        state["moves"] = []
        sheet = workbook["Gastos"]
        rows = list(sheet.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
            has_fecha = "fecha" in header
            for row in rows[1:]:
                if not row or not any(row):
                    continue
                if has_fecha:
                    fec = format_date_value(row[0] if len(row) > 0 else None, default=today_str)
                    hour = format_time_value(row[1] if len(row) > 1 else "")
                    concept = row[2] if len(row) > 2 else None
                    amount = row[3] if len(row) > 3 else 0
                else:
                    hour = format_time_value(row[0] if len(row) > 0 else "")
                    concept = row[1] if len(row) > 1 else None
                    amount = row[2] if len(row) > 2 else 0
                    fec = "20/10/2026" if str(concept).strip() == "ffett" else today_str

                if concept is not None and str(concept).strip() != "":
                    state["moves"].append({
                        "fec": fec,
                        "h": hour,
                        "c": str(concept).strip(),
                        "a": parse_float(amount, 0.0),
                        "type": (
                            str(row[type_index]).strip()
                            if type_index is not None and len(row) > type_index and row[type_index]
                            else "liability"
                        ),
                    })

    if "Inventario" in workbook.sheetnames:
        state["inv"] = []
        for row in workbook["Inventario"].iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            name = row[0] if len(row) > 0 else None
            stock = row[1] if len(row) > 1 else 0
            minimum = row[2] if len(row) > 2 else 0
            if name is not None and str(name).strip() != "":
                state["inv"].append([
                    str(name).strip(),
                    parse_float(stock, 0.0),
                    parse_float(minimum, 0.0),
                ])

    if "Balance" in workbook.sheetnames:
        state["balanceItems"] = []
        for row in workbook["Balance"].iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            concept = row[0] if len(row) > 0 else None
            category = row[1] if len(row) > 1 else ""
            value = row[2] if len(row) > 2 else 0
            impact = row[3] if len(row) > 3 else ""
            if concept is not None and str(concept).strip():
                state["balanceItems"].append({
                    "concept": str(concept).strip(),
                    "category": str(category or "Insumo").strip(),
                    "value": parse_float(value, 0.0),
                    "impact": str(impact or "").strip().lower() in {"suma al negocio", "true", "1", "si", "sí"},
                })

    if "Configuracion" in workbook.sheetnames:
        settings = {}
        for row in workbook["Configuracion"].iter_rows(min_row=2, values_only=True):
            if row and row[0] is not None:
                settings[str(row[0]).strip()] = row[1] if len(row) > 1 else None
        if "base_caja" in settings:
            state["open"] = parse_float(settings["base_caja"], state["open"])
        if "url_qr" in settings and settings["url_qr"]:
            state["url"] = str(settings["url_qr"]).strip()
        if settings.get("personal"):
            raw_staff = settings["personal"]
            if isinstance(raw_staff, list):
                state["staff"] = raw_staff
            elif isinstance(raw_staff, str):
                try:
                    state["staff"] = json.loads(raw_staff)
                except Exception:
                    pass
    return state


def write_state(workbook, state):
    today_str = datetime.date.today().strftime("%d/%m/%Y")

    menu_rows = [
        [category, name, price]
        for category, products in state["menu"]
        for name, price in products
    ]
    set_rows(workbook["Menu"], ["Categoría", "Producto", "Precio"], menu_rows)

    orders_rows = [
        [
            order.get("id"),
            order.get("fec", today_str),
            order.get("h", ""),
            order.get("n", "Cliente"),
            order.get("t", "Mesa"),
            order.get("p", "Efectivo"),
            json.dumps(order.get("items", []), ensure_ascii=False),
            order.get("total", 0),
            order.get("st", "Pendiente"),
            order.get("comments", ""),
            order.get("address", ""),
        ]
        for order in state["orders"]
    ]
    set_rows(
        workbook["Pedidos"],
        ["ID", "Fecha", "Hora", "Cliente", "Tipo", "Pago", "Productos (JSON)", "Total", "Estado", "Comentarios", "Dirección"],
        orders_rows,
    )

    pays_rows = [
        [
            payment.get("fec", today_str),
            payment.get("h", ""),
            payment.get("n", ""),
            payment.get("c", "Pago"),
            payment.get("a", 0),
        ]
        for payment in state["pays"]
    ]
    set_rows(
        workbook["Pagos"],
        ["Fecha", "Hora", "Empleado", "Concepto", "Monto"],
        pays_rows,
    )

    moves_rows = [
        [
            move.get("fec", today_str),
            move.get("h", ""),
            move.get("c", "Gasto"),
            move.get("a", 0),
            move.get("type", "liability"),
        ]
        for move in state["moves"]
    ]
    set_rows(
        workbook["Gastos"],
        ["Fecha", "Hora", "Concepto", "Monto", "Clasificación"],
        moves_rows,
    )

    set_rows(workbook["Inventario"], ["Producto", "Stock", "Mínimo"], state["inv"])
    set_rows(
        workbook["Balance"],
        ["Concepto / Ítem", "Categoría", "Valor ($)", "Impacto"],
        [
            [
                item.get("concept", ""),
                item.get("category", "Insumo"),
                item.get("value", 0),
                "Suma al negocio" if item.get("impact") is True else "No suma / Resta",
            ]
            for item in state.get("balanceItems", [])
        ],
    )
    set_rows(
        workbook["Configuracion"],
        ["Clave", "Valor"],
        [["base_caja", state["open"]], ["url_qr", state["url"]], ["personal", json.dumps(state["staff"], ensure_ascii=False)]],
    )
    app_state = workbook["AppState"]
    app_state.delete_rows(1, max(app_state.max_row, 1))
    app_state.append(["Estado JSON"])
    app_state.append([json.dumps(state, ensure_ascii=False)])
    app_state.sheet_state = "veryHidden"


def save_workbook(workbook):
    if is_workbook_locked():
        raise WorkbookLockedError(
            f"El libro {EXCEL_PATH.name} está abierto en Excel o bloqueado por el sistema. "
            "Cierra Excel para permitir que la plataforma guarde los cambios en el archivo."
        )
    handle = tempfile.NamedTemporaryFile(
        prefix=f".{EXCEL_PATH.stem}-",
        suffix=EXCEL_PATH.suffix,
        dir=EXCEL_PATH.parent,
        delete=False,
    )
    temporary_path = Path(handle.name)
    handle.close()
    try:
        workbook.save(temporary_path)
        try:
            os.replace(temporary_path, EXCEL_PATH)
        except (PermissionError, OSError) as exc:
            raise WorkbookLockedError(
                f"No se pudo guardar {EXCEL_PATH.name} porque está siendo utilizado por Excel. "
                "Cierra el archivo en Excel e intenta guardar de nuevo."
            ) from exc
    finally:
        temporary_path.unlink(missing_ok=True)


def load_or_create_workbook():
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        try:
            workbook = load_workbook(EXCEL_PATH)
        except Exception:
            workbook = load_workbook(EXCEL_PATH, data_only=True)
        for name in ["Menu", "Pedidos", "Pagos", "Gastos", "Inventario", "Balance", "Configuracion", "AppState"]:
            if name not in workbook.sheetnames:
                workbook.create_sheet(name)
        return workbook
    workbook = Workbook()
    workbook.active.title = "Menu"
    for name in ["Pedidos", "Pagos", "Gastos", "Inventario", "Balance", "Configuracion", "AppState"]:
        workbook.create_sheet(name)
    workbook["AppState"].sheet_state = "veryHidden"
    write_state(workbook, default_state())
    workbook["AppState"]["A2"] = None
    save_workbook(workbook)
    return workbook


def read_persisted_state(workbook):
    persisted = None
    if "AppState" in workbook.sheetnames:
        raw = workbook["AppState"]["A2"].value
        if raw:
            try:
                persisted = json.loads(raw)
            except Exception:
                persisted = None
    if persisted and isinstance(persisted, dict):
        today_str = datetime.date.today().strftime("%d/%m/%Y")
        for order in persisted.get("orders", []):
            if "fec" not in order or not order["fec"]:
                order["fec"] = "20/10/2026" if order.get("id") in (1, 2) else today_str
        for pay in persisted.get("pays", []):
            if "fec" not in pay or not pay["fec"]:
                pay["fec"] = today_str
        for move in persisted.get("moves", []):
            if "fec" not in move or not move["fec"]:
                move["fec"] = today_str
        return persisted, True
    return sheet_state(workbook), False


def enviar_alerta_discord(order):
    items = order.get("items")
    if not isinstance(items, list) or not items:
        raise ValueError("El pedido debe incluir al menos un producto.")
    product_lines = []
    for item in items:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            raise ValueError("El detalle de productos del pedido no es válido.")
        product_lines.append(f"- {item[1]} x {str(item[0])[:160]}")

    lines = [
        f"**Nuevo pedido #{str(order.get('id', ''))[:30]}**",
        f"**Cliente:** {str(order.get('n') or 'Cliente general')[:150]}",
        f"**Dirección:** {str(order.get('address') or 'No indicada')[:250]}",
        "**Productos:**",
        *product_lines,
        f"**Comentarios:** {str(order.get('comments') or 'Sin comentarios')[:500]}",
        f"**Total a pagar:** ${parse_float(order.get('total')):,.0f}",
    ]
    return enviar_mensaje_discord(_limitar_mensaje_discord("\n".join(lines)))


def _limitar_mensaje_discord(content):
    return content if len(content) <= 2000 else content[:1997] + "..."


def _resumen_productos(items):
    if not isinstance(items, list):
        raise ValueError("El detalle de productos del pedido no es válido.")
    products = {}
    for item in items:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            raise ValueError("El detalle de productos del pedido no es válido.")
        name = str(item[0])
        quantity = parse_float(item[1])
        price = parse_float(item[2]) if len(item) > 2 else 0
        products[name] = (quantity, price)
    return products


def mensaje_edicion_discord(previous_order, order):
    previous_products = _resumen_productos(previous_order.get("items", []))
    updated_products = _resumen_productos(order.get("items", []))
    changes = []

    fields = [
        ("Cliente", "n", "Cliente general"),
        ("Dirección", "address", "No indicada"),
        ("Tipo de pedido", "t", "Mesa"),
        ("Medio de pago", "p", "Efectivo"),
        ("Comentarios", "comments", "Sin comentarios"),
    ]
    for label, key, default in fields:
        previous = str(previous_order.get(key) or default)
        updated = str(order.get(key) or default)
        if previous != updated:
            changes.append(f"**{label}:** ~~{previous[:200]}~~ → **{updated[:200]}**")

    for name in sorted(previous_products.keys() | updated_products.keys()):
        previous = previous_products.get(name)
        updated = updated_products.get(name)
        if previous == updated:
            continue
        if previous is None:
            changes.append(
                f"**Producto agregado:** **{name[:140]}** "
                f"({updated[0]:g} x ${updated[1]:,.0f})"
            )
        elif updated is None:
            changes.append(
                f"**Producto eliminado:** ~~{name[:140]}~~ "
                f"({previous[0]:g} x ${previous[1]:,.0f})"
            )
        else:
            changes.append(
                f"**Producto modificado:** **{name[:140]}** "
                f"({previous[0]:g} x ${previous[1]:,.0f} → "
                f"{updated[0]:g} x ${updated[1]:,.0f})"
            )

    previous_total = parse_float(previous_order.get("total"))
    updated_total = parse_float(order.get("total"))
    if previous_total != updated_total:
        changes.append(f"**Total:** ~~${previous_total:,.0f}~~ → **${updated_total:,.0f}**")

    if not changes:
        changes.append("No se detectaron cambios en los datos del pedido.")
    lines = [
        f"📝 **Pedido #{str(order.get('id', ''))[:30]} editado**",
        *changes,
    ]
    return _limitar_mensaje_discord("\n".join(lines))


def enviar_mensaje_discord(content):
    webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
    if not webhook_url or not webhook_url.strip():
        print("Advertencia: DISCORD_WEBHOOK_URL no está configurada; no se enviará la alerta.")
        return None

    webhook_url = webhook_url.strip()
    parsed_url = urlsplit(webhook_url)
    if (
        parsed_url.scheme != "https"
        or parsed_url.hostname != "discord.com"
        or not parsed_url.path.startswith("/api/webhooks/")
        or parsed_url.query
        or parsed_url.fragment
    ):
        raise ValueError("DISCORD_WEBHOOK_URL no contiene una URL de webhook válida de Discord.")

    return requests.post(
        webhook_url,
        json={"content": content, "allowed_mentions": {"parse": []}},
        timeout=8,
        allow_redirects=False,
    )


app = Flask(__name__)
app.json.ensure_ascii = False
app.config.update(
    SECRET_KEY=os.getenv("SECRET_KEY"),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "true").lower() == "true",
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(hours=8),
)

if (os.getenv("RENDER") or os.getenv("RAILWAY_ENVIRONMENT_NAME")) and not os.getenv("LONDON_FOOD_EXCEL"):
    app.logger.warning(
        "No se configuró LONDON_FOOD_EXCEL. El libro de Excel quedará en el disco efímero del servicio."
    )


@app.before_request
def require_login():
    if request.endpoint in {"static", "login", "healthz", "service_worker"}:
        return None

    if not os.getenv("APP_PASSWORD") or not app.config.get("SECRET_KEY"):
        message = "Configura APP_PASSWORD y SECRET_KEY para habilitar el acceso seguro."
        if request.path.startswith("/api/"):
            return jsonify({"error": message}), 503
        return render_template("login.html", error=message, unavailable=True), 503

    if not session.get("authenticated"):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Inicia sesión para continuar."}), 401
        return redirect(url_for("login"))
    return None


@app.after_request
def secure_response_headers(response):
    if request.endpoint != "static":
        response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


@app.get("/")
def index():
    return render_template("index.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    password = os.getenv("APP_PASSWORD")
    if not password or not app.config.get("SECRET_KEY"):
        message = "Configura APP_PASSWORD y SECRET_KEY en las variables de entorno."
        return render_template("login.html", error=message, unavailable=True), 503

    error = None
    if request.method == "POST":
        supplied_password = request.form.get("password", "")
        if hmac.compare_digest(supplied_password, password):
            session.clear()
            session["authenticated"] = True
            session.permanent = True
            return redirect(url_for("index"))
        error = "La contraseña no es correcta."
    return render_template("login.html", error=error, unavailable=False)


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/healthz")
def healthz():
    return jsonify({"status": "ok"})


@app.get("/service-worker.js")
def service_worker():
    response = send_from_directory(app.static_folder, "service-worker.js")
    response.headers["Service-Worker-Allowed"] = "/"
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.post("/api/discord/order")
def notify_discord_order():
    order = request.get_json(silent=True)
    if not isinstance(order, dict):
        return jsonify({"error": "Los datos del pedido no son válidos."}), 400

    try:
        response = enviar_alerta_discord(order)
    except (TypeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    except requests.RequestException as error:
        app.logger.error("No se pudo enviar el pedido a Discord (%s).", type(error).__name__)
        return jsonify({"error": "No se pudo enviar la notificación a Discord."}), 502

    if response is None:
        return jsonify({"error": "Discord no está configurado en el servidor."}), 503
    if not 200 <= response.status_code < 300:
        app.logger.error("Discord rechazó la notificación (HTTP %s).", response.status_code)
        return jsonify({"error": "Discord rechazó la notificación del pedido."}), 502
    return jsonify({"sent": True})


@app.post("/api/discord/order/update")
def notify_discord_order_update():
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "Los datos de la edición no son válidos."}), 400

    previous_order = payload.get("previous")
    updated_order = payload.get("order")
    if not isinstance(previous_order, dict) or not isinstance(updated_order, dict):
        return jsonify({"error": "Falta el pedido original o el pedido actualizado."}), 400
    if previous_order.get("id") != updated_order.get("id"):
        return jsonify({"error": "El número del pedido editado no coincide."}), 400

    try:
        message = mensaje_edicion_discord(previous_order, updated_order)
        response = enviar_mensaje_discord(message)
    except (TypeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    except requests.RequestException as error:
        app.logger.error("No se pudo enviar la edición a Discord (%s).", type(error).__name__)
        return jsonify({"error": "No se pudo enviar la edición del pedido a Discord."}), 502

    if response is None:
        return jsonify({"error": "Discord no está configurado en el servidor."}), 503
    if not 200 <= response.status_code < 300:
        try:
            discord_error = response.json().get("message", "")
        except (ValueError, AttributeError):
            discord_error = ""
        app.logger.error(
            "Discord rechazó la edición (HTTP %s): %s",
            response.status_code,
            str(discord_error)[:300],
        )
        description = f" Discord: {str(discord_error)[:200]}" if discord_error else ""
        return jsonify({
            "error": f"Discord rechazó la notificación de edición (HTTP {response.status_code}).{description}"
        }), 502
    return jsonify({"sent": True})


@app.get("/api/health")
def health():
    try:
        with WORKBOOK_LOCK:
            locked = is_workbook_locked()
            exists = EXCEL_PATH.exists()
            return jsonify({
                "connected": True,
                "writable": not locked,
                "workbook": EXCEL_PATH.name,
                "exists": exists,
            })
    except (OSError, InvalidFileException) as error:
        app.logger.exception("No se pudo verificar el archivo de Excel")
        return jsonify({"error": f"No se pudo verificar el Excel: {error}"}), 500


@app.get("/api/state")
def get_state():
    workbook = None
    try:
        with WORKBOOK_LOCK:
            workbook = load_or_create_workbook()
            state, initialized = read_persisted_state(workbook)
            locked = is_workbook_locked()
            menu_migrated = migrate_legacy_menu(state)
            staff_migrated = migrate_legacy_staff(state)
            state_migrated = menu_migrated or staff_migrated
            if (not initialized or state_migrated) and not locked:
                try:
                    write_state(workbook, state)
                    save_workbook(workbook)
                    initialized = True
                except WorkbookLockedError:
                    pass
            return jsonify({
                "state": state,
                "initialized": initialized,
                "workbook": EXCEL_PATH.name,
                "writable": not locked,
            })
    except (OSError, InvalidFileException, ValueError, TypeError, KeyError) as error:
        app.logger.exception("No se pudo cargar el estado del Excel")
        return jsonify({"error": f"No se pudo cargar el Excel: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()


@app.post("/api/state")
def save_state():
    state = request.get_json(silent=True)
    if not isinstance(state, dict):
        return jsonify({"error": "El estado enviado no es válido."}), 400
    required = {"orders": list, "pays": list, "moves": list, "inv": list, "menu": list, "staff": list}
    if any(not isinstance(state.get(key), kind) for key, kind in required.items()):
        return jsonify({"error": "Faltan datos o tienen un formato incorrecto."}), 400
    normalized = default_state()
    normalized.update(state)
    workbook = None
    try:
        with WORKBOOK_LOCK:
            workbook = load_or_create_workbook()
            write_state(workbook, normalized)
            save_workbook(workbook)
    except WorkbookLockedError as error:
        return jsonify({"error": str(error)}), 423
    except (OSError, InvalidFileException) as error:
        app.logger.exception("No se pudo guardar el archivo de Excel")
        return jsonify({"error": f"No se pudo guardar el Excel: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()
    return jsonify({"saved": True, "workbook": EXCEL_PATH.name})


@app.post("/api/import")
def import_excel():
    workbook = None
    try:
        with WORKBOOK_LOCK:
            workbook = load_or_create_workbook()
            state = sheet_state(workbook)
            migrate_legacy_menu(state)
            migrate_legacy_staff(state)
            locked = is_workbook_locked()
            if not locked:
                try:
                    write_state(workbook, state)
                    save_workbook(workbook)
                except WorkbookLockedError:
                    locked = True
            return jsonify({
                "state": state,
                "workbook": EXCEL_PATH.name,
                "writable": not locked,
            })
    except (OSError, InvalidFileException, ValueError, TypeError, KeyError) as error:
        app.logger.exception("No se pudo importar el archivo de Excel")
        return jsonify({"error": f"No se pudo importar el Excel: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()


@app.post("/api/reset")
def reset_state():
    workbook = None
    try:
        with WORKBOOK_LOCK:
            workbook = load_or_create_workbook()
            state = default_state()
            write_state(workbook, state)
            save_workbook(workbook)
    except WorkbookLockedError as error:
        return jsonify({"error": str(error)}), 423
    except (OSError, InvalidFileException) as error:
        app.logger.exception("No se pudo restablecer el archivo de Excel")
        return jsonify({"error": f"No se pudo restablecer el Excel: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()
    return jsonify({"state": state, "workbook": EXCEL_PATH.name})


@app.get("/api/excel/download")
def download_excel():
    workbook = None
    try:
        with WORKBOOK_LOCK:
            if not EXCEL_PATH.exists():
                workbook = load_or_create_workbook()
                workbook.close()
                workbook = None
            return send_file(EXCEL_PATH, as_attachment=True, download_name=EXCEL_PATH.name)
    except (OSError, InvalidFileException) as error:
        app.logger.exception("No se pudo descargar el archivo de Excel")
        return jsonify({"error": f"No se pudo descargar el Excel: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()


def parse_report_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for date_format in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    return None


def report_workbook(state, selected_date):
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumen"
    details = workbook.create_sheet("Registros")
    order_details = workbook.create_sheet("Detalle pedidos")

    order_records = []
    payment_records = []
    expense_records = []
    for order in state.get("orders", []):
        record_date = parse_report_date(order.get("fec"))
        if selected_date is None or record_date == selected_date:
            order_records.append((record_date, order))
    for payment in state.get("pays", []):
        record_date = parse_report_date(payment.get("fec"))
        if selected_date is None or record_date == selected_date:
            payment_records.append((record_date, payment))
    for expense in state.get("moves", []):
        record_date = parse_report_date(expense.get("fec"))
        if selected_date is None or record_date == selected_date:
            expense_records.append((record_date, expense))

    sales_total = sum(parse_float(order.get("total")) for _, order in order_records)
    payments_total = sum(parse_float(payment.get("a")) for _, payment in payment_records)
    expenses_total = sum(parse_float(expense.get("a")) for _, expense in expense_records)
    period = selected_date.strftime("%d/%m/%Y") if selected_date else "Todo el historial"

    summary.append(["Informe de actividad"])
    summary.append(["Periodo", period])
    summary.append(["Tipo de registro", "Cantidad", "Total"])
    summary.append(["Pedidos", len(order_records), sales_total])
    summary.append(["Pagos al personal", len(payment_records), payments_total])
    summary.append(["Gastos", len(expense_records), expenses_total])
    summary.append(["Ventas menos pagos y gastos", "", sales_total - payments_total - expenses_total])

    details.append([
        "Fecha", "Hora", "Registro", "ID pedido", "Cliente o empleado",
        "Tipo o medio de pago", "Detalle", "Estado", "Ingreso", "Salida",
    ])
    order_details.append([
        "ID pedido", "Fecha", "Hora", "Cliente", "Dirección", "Tipo de pedido",
        "Medio de pago", "Producto", "Cantidad", "Precio unitario",
        "Subtotal producto", "Total pedido", "Estado", "Comentarios",
    ])
    rows = []
    for record_date, order in order_records:
        items = order.get("items", [])
        valid_items = [
            item for item in items
            if isinstance(item, (list, tuple)) and len(item) >= 2
        ] if isinstance(items, list) else []
        item_details = ", ".join(
            f'{item[1]} x {item[0]} @ {parse_float(item[2]):,.2f}'
            if len(item) > 2 else f"{item[1]} x {item[0]}"
            for item in valid_items
        )
        order_total = parse_float(order.get("total"))
        order_date = record_date.strftime("%d/%m/%Y") if record_date else ""
        if valid_items:
            for item in valid_items:
                quantity = parse_float(item[1])
                unit_price = parse_float(item[2]) if len(item) > 2 else None
                order_details.append([
                    order.get("id", ""),
                    order_date,
                    order.get("h", ""),
                    order.get("n", ""),
                    order.get("address", ""),
                    order.get("t", ""),
                    order.get("p", ""),
                    item[0],
                    quantity,
                    unit_price,
                    quantity * unit_price if unit_price is not None else None,
                    order_total,
                    order.get("st", ""),
                    order.get("comments", ""),
                ])
        else:
            order_details.append([
                order.get("id", ""),
                order_date,
                order.get("h", ""),
                order.get("n", ""),
                order.get("address", ""),
                order.get("t", ""),
                order.get("p", ""),
                "",
                None,
                None,
                None,
                order_total,
                order.get("st", ""),
                order.get("comments", ""),
            ])
        rows.append((
            record_date or datetime.date.min,
            str(order.get("h", "")),
            [
                order_date,
                order.get("h", ""),
                "Pedido",
                order.get("id", ""),
                order.get("n", ""),
                f'{order.get("t", "")} / {order.get("p", "")}',
                item_details,
                order.get("st", ""),
                order_total,
                0,
            ],
        ))
    for record_date, payment in payment_records:
        rows.append((
            record_date or datetime.date.min,
            str(payment.get("h", "")),
            [
                record_date.strftime("%d/%m/%Y") if record_date else "",
                payment.get("h", ""),
                "Pago al personal",
                "",
                payment.get("n", ""),
                "",
                payment.get("c", ""),
                "",
                0,
                parse_float(payment.get("a")),
            ],
        ))
    for record_date, expense in expense_records:
        rows.append((
            record_date or datetime.date.min,
            str(expense.get("h", "")),
            [
                record_date.strftime("%d/%m/%Y") if record_date else "",
                expense.get("h", ""),
                "Gasto",
                "",
                "",
                "",
                expense.get("c", ""),
                "",
                0,
                parse_float(expense.get("a")),
            ],
        ))
    for _, _, row in sorted(rows, key=lambda item: (item[0], item[1])):
        details.append(row)

    for sheet in (summary, details, order_details):
        header_rows = (1, 3) if sheet is summary else (1,)
        for row_number in header_rows:
            for cell in sheet[row_number]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="0B1F4D")
        sheet.freeze_panes = "A2" if sheet is details else "A4"
        if sheet is not summary and sheet.max_row > 1:
            sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            letter = column[0].column_letter
            width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 45)
            sheet.column_dimensions[letter].width = width
    for row in details.iter_rows(min_row=2, min_col=9, max_col=10):
        for cell in row:
            cell.number_format = '#,##0.00'
    for row in order_details.iter_rows(min_row=2, min_col=8, max_col=8):
        for cell in row:
            cell.number_format = '#,##0.##'
    for row in order_details.iter_rows(min_row=2, min_col=9, max_col=11):
        for cell in row:
            cell.number_format = '#,##0.00'
    for cell in summary["C"][3:]:
        cell.number_format = '#,##0.00'

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    workbook.close()
    return output


@app.get("/api/reports/download")
def download_report():
    raw_date = request.args.get("date", "").strip()
    if raw_date == "all":
        selected_date = None
    else:
        selected_date = parse_report_date(raw_date)
        if selected_date is None:
            return jsonify({"error": "Indica una fecha válida o selecciona todo el historial."}), 400

    workbook = None
    try:
        with WORKBOOK_LOCK:
            workbook = load_or_create_workbook()
            state, _ = read_persisted_state(workbook)
            report = report_workbook(state, selected_date)
        filename = (
            f"informe-{selected_date.isoformat()}.xlsx"
            if selected_date else "informe-historico.xlsx"
        )
        return send_file(
            report,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except (OSError, InvalidFileException, ValueError, TypeError, KeyError) as error:
        app.logger.exception("No se pudo generar el informe")
        return jsonify({"error": f"No se pudo generar el informe: {error}"}), 500
    finally:
        if workbook is not None:
            workbook.close()


if __name__ == "__main__":
    app.run(
        host=os.getenv("HOST", "0.0.0.0"),
        port=int(os.getenv("PORT", "7860")),
        debug=False,
    )
