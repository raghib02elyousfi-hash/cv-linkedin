#!/usr/bin/env python3
"""Generate BrillOPal_Gestion.xlsx"""

import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
import math

# ─── COLOUR PALETTE ───────────────────────────────────────────────────────────
BLUE_DARK  = "1F3864"
BLUE_MED   = "2E75B6"
BLUE_LIGHT = "D6E4F0"
GOLD       = "F4B942"
WHITE      = "FFFFFF"
GREY_LIGHT = "F2F2F2"
GREEN      = "70AD47"
RED        = "FF0000"
ORANGE     = "ED7D31"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color=WHITE, size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic)

def border_thin():
    s = Side(style='thin', color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal='center', vertical='center', wrap_text=True)

def right():
    return Alignment(horizontal='right', vertical='center')

def apply_header(ws, row, col_start, col_end, title, bg=BLUE_DARK):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start)
    cell.value = title
    cell.fill = fill(bg)
    cell.font = Font(bold=True, color=WHITE, size=12)
    cell.alignment = center()

def apply_subheader(ws, row, col_start, col_end, title):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start)
    cell.value = title
    cell.fill = fill(BLUE_MED)
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.alignment = center()

def col_header(ws, row, col, title, bg=BLUE_LIGHT):
    cell = ws.cell(row=row, column=col, value=title)
    cell.fill = fill(bg)
    cell.font = Font(bold=True, color=BLUE_DARK, size=10)
    cell.alignment = center()
    cell.border = border_thin()

def data_cell(ws, row, col, value, fmt=None, bg=None):
    cell = ws.cell(row=row, column=col, value=value)
    if bg:
        cell.fill = fill(bg)
    if fmt:
        cell.number_format = fmt
    cell.alignment = right()
    cell.border = border_thin()
    return cell

def label_cell(ws, row, col, value, bold=False, bg=None):
    cell = ws.cell(row=row, column=col, value=value)
    if bg:
        cell.fill = fill(bg)
    cell.font = Font(bold=bold, color=BLUE_DARK, size=10)
    cell.border = border_thin()
    return cell

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
EURO   = '#,##0.00 €'
PCT    = '0.00%'
HOURS  = '#,##0.00 "h"'

# Palencia 2025 conv. salary (14 pagas) + BrillOPal +3%
CATEGORIES = [
    ("I – Encargado General",    1375.00),
    ("II – Encargado de Zona",   1240.00),
    ("III – Especialista",       1165.00),
    ("IV – Limpiador/a",         1107.22),
    ("V – Peón Auxiliar",        1107.22),
]
PREMIUM = 0.03
PAGAS   = 14
HOURS_YEAR = 1687.5

# Employer SS
SS = {
    "Contingencias Comunes": 0.2360,
    "Desempleo (indef.)":    0.0550,
    "Formación Profesional": 0.0060,
    "FOGASA":                0.0020,
    "MEI":                   0.0058,
    "AT/EP (Limpieza)":      0.0130,
}
SS_TOTAL = sum(SS.values())  # 0.3178

# Direct variable costs per hour
DIRECT = {
    "b2b": {"Materiales/productos": 0.40, "Desplazamiento": 0.25},
    "b2c": {"Materiales/productos": 0.55, "Desplazamiento": 0.45},
}

# Fixed monthly structure
STRUCTURE = {
    "Oficina/Almacén":          300.00,
    "Suministros":               100.00,
    "Seguro RC Profesional":    150.00,
    "Seguro AT Empleados":       50.00,
    "Gestoría/Asesoría":        150.00,
    "Teléfono + Internet":        80.00,
    "Vehículo (cuota+seguro)":  350.00,
    "Material oficina/admin":    50.00,
    "Marketing / Web":           100.00,
    "Varios / imprevistos":      100.00,
}
STRUCTURE_TOTAL = sum(STRUCTURE.values())  # 1,430 €/month

# Owner (autónomo) cost
OWNER_SS_Y1 = 80.00    # tarifa plana primer año
OWNER_SS_Y2 = 160.00   # segundo semestre
OWNER_SS_STD = 320.00  # tarifa estándar
OWNER_DRAW = 1500.00   # monthly owner draw

# ─── SHEET 1: ESCANDALLO ──────────────────────────────────────────────────────

def build_escandallo(wb):
    ws = wb.create_sheet("1_Escandallo Costes")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 32
    for c in 'BCDEFGHI':
        ws.column_dimensions[c].width = 16

    r = 1
    apply_header(ws, r, 1, 9,
        "BRILLOPAL – Escandallo de Costes y Precio de Venta por Hora (2025/2026)")
    ws.row_dimensions[r].height = 30
    r += 1

    # ── 1A. Tabla salarial del convenio ──────────────────────────────────────
    apply_subheader(ws, r, 1, 9,
        "A. TABLA SALARIAL – Convenio Limpieza Edificios y Locales Palencia 2022-2025")
    r += 1
    ws.row_dimensions[r].height = 22

    headers_sal = ["Categoría", "Sal. Base/mes Conv.\n(€)", "Sal. Base/mes\n+3% BrillOPal",
                   "14 pagas/año Conv.", "14 pagas/año\n+3%", "Base SS mensual\n(anual/12)",
                   "SS empresa/año\n(31.78%)", "Coste total\nempresa/año"]
    for i, h in enumerate(headers_sal, 1):
        col_header(ws, r, i, h)
    r += 1

    for cat, sal_base in CATEGORIES:
        sal_brillopal = sal_base * (1 + PREMIUM)
        sal_anual_conv = sal_base * PAGAS
        sal_anual_brill = sal_brillopal * PAGAS
        ss_base_mensual = sal_anual_brill / 12
        ss_empresa_anual = ss_base_mensual * 12 * SS_TOTAL
        coste_total_anual = sal_anual_brill + ss_empresa_anual

        label_cell(ws, r, 1, cat)
        data_cell(ws, r, 2, sal_base, EURO)
        data_cell(ws, r, 3, sal_brillopal, EURO, GREY_LIGHT)
        data_cell(ws, r, 4, sal_anual_conv, EURO)
        data_cell(ws, r, 5, sal_anual_brill, EURO, GREY_LIGHT)
        data_cell(ws, r, 6, ss_base_mensual, EURO)
        data_cell(ws, r, 7, ss_empresa_anual, EURO)
        data_cell(ws, r, 8, coste_total_anual, EURO, BLUE_LIGHT)
        r += 1

    # SS breakdown
    r += 1
    apply_subheader(ws, r, 1, 4, "Desglose cotizaciones empresa (% s/base SS)")
    r += 1
    for concepto, pct in SS.items():
        label_cell(ws, r, 1, concepto)
        data_cell(ws, r, 2, pct, PCT)
        r += 1
    label_cell(ws, r, 1, "TOTAL COTIZACIÓN EMPRESA", bold=True)
    data_cell(ws, r, 2, SS_TOTAL, PCT, BLUE_LIGHT)
    r += 2

    # ── 1B. Escandallo por hora ───────────────────────────────────────────────
    apply_subheader(ws, r, 1, 9,
        "B. ESCANDALLO POR HORA – Categoría IV (Limpiador/a) · Referencia principal")
    r += 1

    sal_iv_brill = 1107.22 * (1 + PREMIUM)
    sal_anual_iv = sal_iv_brill * PAGAS
    ss_base_iv = sal_anual_iv / 12
    ss_emp_iv = ss_base_iv * 12 * SS_TOTAL
    coste_total_iv = sal_anual_iv + ss_emp_iv
    coste_hora_iv = coste_total_iv / HOURS_YEAR
    absent_iv = coste_hora_iv * 0.035  # 3.5% employer absenteeism cost (days 1-3)

    # B2B
    mat_b2b  = DIRECT["b2b"]["Materiales/productos"]
    desp_b2b = DIRECT["b2b"]["Desplazamiento"]
    direct_b2b = coste_hora_iv + absent_iv + mat_b2b + desp_b2b

    # B2C
    mat_b2c  = DIRECT["b2c"]["Materiales/productos"]
    desp_b2c = DIRECT["b2c"]["Desplazamiento"]
    direct_b2c = coste_hora_iv + absent_iv + mat_b2c + desp_b2c

    headers_esc = ["Concepto", "€/hora B2B", "€/hora B2C", "Notas"]
    for i, h in enumerate(headers_esc, 1):
        col_header(ws, r, i, h)
    r += 1

    rows_esc = [
        ("Coste laboral (salario+SS / h productiva)",
         coste_hora_iv, coste_hora_iv,
         f"={sal_anual_iv:.2f}+{ss_emp_iv:.2f}) / {HOURS_YEAR}h"),
        ("Provisión absentismo (3,5% s/labor)",
         absent_iv, absent_iv, "Días 1-3 baja a cargo empresa"),
        ("Materiales y productos de limpieza",
         mat_b2b, mat_b2c, "B2B: a granel; B2C: envases menores"),
        ("Desplazamiento / combustible",
         desp_b2b, desp_b2c, "B2B: rutas fijas; B2C: dispersión mayor"),
    ]
    for label, b2b, b2c, note in rows_esc:
        label_cell(ws, r, 1, label)
        data_cell(ws, r, 2, b2b, EURO)
        data_cell(ws, r, 3, b2c, EURO)
        label_cell(ws, r, 4, note)
        r += 1

    # Subtotals
    label_cell(ws, r, 1, "COSTE DIRECTO / HORA (sin estructura)", bold=True)
    data_cell(ws, r, 2, direct_b2b, EURO, GOLD)
    data_cell(ws, r, 3, direct_b2c, EURO, GOLD)
    r += 2

    # Structure allocation per employee (3 employees scenario as reference)
    for n_emp in [1, 3, 5]:
        struct_h = STRUCTURE_TOTAL / (n_emp * HOURS_YEAR / 12)
        total_b2b = direct_b2b + struct_h
        total_b2c = direct_b2c + struct_h
        pvp_b2b  = total_b2b * 1.30
        pvp_b2c  = total_b2c * 1.30
        margin_b2b = pvp_b2b - total_b2b
        margin_b2c = pvp_b2c - total_b2c

        apply_subheader(ws, r, 1, 9,
            f"  Escenario {n_emp} empleado{'s' if n_emp>1 else ''}  –  Estructura: {STRUCTURE_TOTAL:.0f}€/mes ÷ {n_emp} operarios")
        r += 1

        rows2 = [
            ("Coste directo / hora",       direct_b2b, direct_b2c),
            ("Estructura asignada / hora",  struct_h,   struct_h),
            ("COSTE TOTAL / HORA",          total_b2b,  total_b2c),
            ("PRECIO VENTA / HORA (+30%)",  pvp_b2b,    pvp_b2c),
            ("BENEFICIO / HORA",            margin_b2b, margin_b2c),
        ]
        for label, b2b, b2c in rows2:
            bold = label.startswith(("COSTE TOTAL", "PRECIO", "BENEFICIO"))
            bg = BLUE_LIGHT if "PRECIO" in label else (GREEN if "BENEFICIO" in label else None)
            label_cell(ws, r, 1, label, bold=bold)
            data_cell(ws, r, 2, b2b, EURO, bg)
            data_cell(ws, r, 3, b2c, EURO, bg)
            r += 1
        r += 1

    # ── 1C. Estructura mensual ────────────────────────────────────────────────
    apply_subheader(ws, r, 1, 9, "C. COSTES DE ESTRUCTURA FIJOS MENSUALES")
    r += 1
    col_header(ws, r, 1, "Concepto")
    col_header(ws, r, 2, "€/mes")
    r += 1

    for concepto, importe in STRUCTURE.items():
        label_cell(ws, r, 1, concepto)
        data_cell(ws, r, 2, importe, EURO)
        r += 1
    label_cell(ws, r, 1, "TOTAL ESTRUCTURA MENSUAL", bold=True)
    data_cell(ws, r, 2, STRUCTURE_TOTAL, EURO, BLUE_LIGHT)

    return ws


# ─── SHEET 2: CLIENTES Y CONTRATOS ───────────────────────────────────────────

def build_clientes(wb):
    ws = wb.create_sheet("2_Clientes y Contratos")
    ws.sheet_view.showGridLines = False

    apply_header(ws, 1, 1, 16, "BRILLOPAL – Registro de Clientes y Contratos")
    ws.row_dimensions[1].height = 25

    headers = [
        "ID", "Razón Social / Nombre", "Tipo\n(B2B/B2C)", "NIF/DNI",
        "Dirección", "C.P.", "Teléfono", "Email", "Contacto",
        "Tipo Servicio", "Frecuencia", "Horas/\nMes", "€/Hora\n(s/IVA)",
        "€/Mes\n(s/IVA)", "Fecha\nInicio", "Estado"
    ]
    widths = [6, 28, 8, 14, 30, 7, 14, 24, 18, 20, 14, 8, 10, 12, 12, 10]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        col_header(ws, 2, i, h)

    # Sample B2B rows
    b2b_samples = [
        ("B2B-001", "Industrias Cerealistas Norte SL", "B2B", "B-34001234",
         "Polígono La Mora, C/Cedros 12, Palencia", "34004", "979 123 456",
         "gerencia@cerealistas.com", "José M. García",
         "Limpieza nave industrial", "Diaria (L-V)", 80, 20.41,
         None, "01/07/2026", "Activo"),
        ("B2B-002", "Comunidad Propietarios C/Mayor 45", "B2B", "—",
         "C/Mayor 45, Palencia", "34001", "979 234 567",
         "presidente@cmajor45.es", "Marta López",
         "Limpieza zonas comunes", "3 días/sem", 50, 20.41,
         None, "01/08/2026", "Pendiente"),
        ("B2B-003", "Oficinas Plenilunio SL", "B2B", "B-34005678",
         "Avda. Santander 8, Palencia", "34003", "979 345 678",
         "admin@plenilunio.es", "Carmen Ruiz",
         "Limpieza oficinas", "Diaria (L-V)", 60, 20.41,
         None, "15/07/2026", "Propuesta"),
    ]
    b2c_samples = [
        ("B2C-001", "García Fernández, Antonio", "B2C", "12345678A",
         "C/Palomares 3, 2ºA, Palencia", "34002", "666 111 222",
         "antonio.garcia@email.com", "—",
         "Limpieza hogar", "Quincenal", 8, 22.50,
         None, "05/07/2026", "Activo"),
        ("B2C-002", "Consultoría ACME (pequeña of.)", "B2C", "B-34009999",
         "C/Gil de Fuentes 5, Palencia", "34001", "666 333 444",
         "info@acmeconsult.es", "—",
         "Limpieza oficina pequeña", "Semanal", 12, 22.50,
         None, "01/08/2026", "Propuesta"),
    ]

    r = 3
    for row in b2b_samples + b2c_samples:
        bg = GREY_LIGHT if r % 2 == 0 else WHITE
        label_cell(ws, r, 1, row[0], bg=bg)
        label_cell(ws, r, 2, row[1], bg=bg)
        for i, v in enumerate(row[2:], 3):
            if isinstance(v, float):
                data_cell(ws, r, i, v, EURO, bg)
            else:
                label_cell(ws, r, i, v, bg=bg)
        # compute monthly total
        horas = row[11]
        precio = row[12]
        ws.cell(row=r, column=14).value = horas * precio if horas and precio else None
        ws.cell(row=r, column=14).number_format = EURO
        ws.cell(row=r, column=14).fill = fill(GREEN if row[2] == "B2B" else GOLD)
        ws.cell(row=r, column=14).border = border_thin()
        r += 1

    # Blank rows for new entries
    for _ in range(20):
        bg = GREY_LIGHT if r % 2 == 0 else WHITE
        for i in range(1, 17):
            c = ws.cell(row=r, column=i, value="")
            c.fill = fill(bg)
            c.border = border_thin()
        r += 1

    return ws


# ─── SHEET 3: PLANIFICACIÓN ───────────────────────────────────────────────────

def build_planificacion(wb):
    ws = wb.create_sheet("3_Planificación Servicios")
    ws.sheet_view.showGridLines = False

    apply_header(ws, 1, 1, 9,
        "BRILLOPAL – Planificación Semanal de Servicios y Asignación de Empleados")
    ws.row_dimensions[1].height = 25

    # Week header
    days = ["Empleado", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "H. Total\nsemana", "Cliente\nPrincipal"]
    widths = [20, 22, 22, 22, 22, 22, 22, 10, 22]
    for i, (h, w) in enumerate(zip(days, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        col_header(ws, 2, i, h)

    employees = ["Empleado 1 (Limpiador/a)", "Empleado 2 (Limpiador/a)", "Empleado 3 (Especialista)"]
    shifts = [
        ["Ind. Cerealistas 6h", "Ind. Cerealistas 6h", "Ind. Cerealistas 6h", "Ind. Cerealistas 6h", "Ind. Cerealistas 6h", "—", 30, "Industrias Cerealistas"],
        ["Of. Plenilunio 4h\nCom. Mayor 3h", "Of. Plenilunio 4h\nCom. Mayor 3h", "Of. Plenilunio 4h\nCom. Mayor 3h", "Of. Plenilunio 4h", "Of. Plenilunio 4h\nCom. Mayor 3h", "Com. Mayor 3h", 37, "Plenilunio / C/Mayor"],
        ["Hogar B2C 3h\nOf. ACME 3h", "—", "Hogar B2C 3h", "—", "Hogar B2C 3h\nOf. ACME 3h", "—", 12, "Clientes B2C"],
    ]

    for i, (emp, shift) in enumerate(zip(employees, shifts)):
        r = 3 + i
        bg = [GREY_LIGHT, WHITE, GREY_LIGHT][i]
        label_cell(ws, r, 1, emp, bold=True, bg=bg)
        for j, v in enumerate(shift, 2):
            c = ws.cell(row=r, column=j, value=v)
            c.fill = fill(bg)
            c.border = border_thin()
            c.alignment = Alignment(wrap_text=True, vertical='center')

    # Colour legend
    r = 7
    ws.cell(row=r, column=1, value="Leyenda de colores:").font = Font(bold=True, color=BLUE_DARK)
    legend = [("B2B – Industria", GREEN), ("B2B – Oficinas/Comunidades", BLUE_MED), ("B2C – Hogares", GOLD), ("Libre/Vacío", WHITE)]
    for i, (txt, color) in enumerate(legend, 2):
        ws.cell(row=r, column=i, value=txt).fill = fill(color)
        ws.cell(row=r, column=i).border = border_thin()

    return ws


# ─── SHEET 4: FACTURACIÓN ─────────────────────────────────────────────────────

def build_facturacion(wb):
    ws = wb.create_sheet("4_Facturación y Cobros")
    ws.sheet_view.showGridLines = False

    apply_header(ws, 1, 1, 12, "BRILLOPAL – Registro de Facturación, Cobros y Antigüedad de Deuda")
    ws.row_dimensions[1].height = 25

    headers = ["Nº Factura", "Fecha Emisión", "Cliente", "Concepto",
               "Base\nImponible", "IVA 21%", "TOTAL\nFACTURA",
               "Fecha\nVencimiento", "Fecha\nCobro", "Estado",
               "Días\nPendiente", "Observaciones"]
    widths = [12, 14, 28, 30, 13, 10, 13, 14, 14, 12, 10, 24]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        col_header(ws, 2, i, h)

    samples = [
        ("FAC-001", "01/07/2026", "Industrias Cerealistas Norte SL",
         "Servicios limpieza nave – julio 2026",
         80*20.41, None, None, "31/07/2026", "31/07/2026", "Cobrada", 0, ""),
        ("FAC-002", "01/07/2026", "Of. Plenilunio SL",
         "Servicios limpieza oficinas – julio 2026",
         60*20.41, None, None, "31/07/2026", None, "Pendiente", None, ""),
        ("FAC-003", "01/07/2026", "García Fernández, Antonio",
         "Limpieza hogar julio 2026 (2 servicios)",
         8*22.50, None, None, "15/07/2026", "18/07/2026", "Cobrada", 3, ""),
    ]

    r = 3
    for row in samples:
        bg = GREY_LIGHT if r % 2 == 0 else WHITE
        base = row[4]
        iva  = round(base * 0.21, 2)
        total = round(base + iva, 2)
        values = list(row[:4]) + [base, iva, total] + list(row[7:])
        for i, v in enumerate(values, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.fill = fill(bg)
            c.border = border_thin()
            if i in (5, 6, 7):
                c.number_format = EURO
                c.alignment = right()
            if i == 10:
                c.fill = fill(GREEN if v == "Cobrada" else ORANGE)
                c.font = Font(bold=True, color=WHITE)
                c.alignment = center()
        r += 1

    # Empty rows
    for _ in range(40):
        bg = GREY_LIGHT if r % 2 == 0 else WHITE
        for i in range(1, 13):
            c = ws.cell(row=r, column=i, value="")
            c.fill = fill(bg)
            c.border = border_thin()
        r += 1

    return ws


# ─── SHEET 5: CUENTA RESULTADOS 18M ──────────────────────────────────────────

def build_cuenta_resultados(wb):
    ws = wb.create_sheet("5_Cuenta Resultados 18M")
    ws.sheet_view.showGridLines = False

    MONTHS = ["Jul'26","Ago'26","Sep'26","Oct'26","Nov'26","Dic'26",
              "Ene'27","Feb'27","Mar'27","Abr'27","May'27","Jun'27",
              "Jul'27","Ago'27","Sep'27","Oct'27","Nov'27","Dic'27"]

    # ── Scenario assumptions ─────────────────────────────────────────────────
    # Hours/month available per scenario
    AVAIL = {1: 140.625, 3: 421.875, 5: 703.125}  # 1,687.5/12

    # Occupancy ramp (% of available hours billed) – slow start
    RAMP = [0.35, 0.45, 0.55, 0.65, 0.70, 0.72,
            0.75, 0.75, 0.78, 0.80, 0.82, 0.85,
            0.85, 0.82, 0.87, 0.88, 0.90, 0.90]

    # B2B/B2C split of billed hours
    B2B_PCT = 0.70   # 70% B2B hours
    B2C_PCT = 0.30

    # Direct cost per hour (3-employee reference for the mix; scaled by actual)
    struct_per_h = {n: STRUCTURE_TOTAL / AVAIL[n] for n in [1,3,5]}

    # Labour cost per employee/month
    sal_iv = 1107.22 * 1.03 * PAGAS / 12
    ss_iv  = (1107.22 * 1.03 * PAGAS / 12) * SS_TOTAL
    LABOR_PER_EMP = sal_iv + ss_iv  # ~1,756 €/month

    # Owner cost per month (tarifa plana 1st year, then standard)
    owner_cost = [OWNER_SS_Y1]*6 + [OWNER_SS_Y2]*6 + [OWNER_SS_STD]*6

    # Revenue per hour
    rev_b2b = 20.41
    rev_b2c = 22.50
    dir_cost_h = 13.22 + 0.40 + 0.25  # labor+absentia+mat+desp (B2B ref.)

    # ── Write scenarios ───────────────────────────────────────────────────────
    scenarios = [(1, BLUE_DARK), (3, BLUE_MED), (5, GREEN)]
    col_offset = 0

    for s_idx, (n_emp, s_color) in enumerate(scenarios):
        c0 = 1 + s_idx * 22  # start column for this scenario

        apply_header(ws, 1, c0, c0+19,
            f"ESCENARIO {n_emp} EMPLEADO{'S' if n_emp>1 else ''}", s_color)

        # Month headers
        ws.cell(row=2, column=c0, value="Concepto").fill = fill(s_color)
        ws.cell(row=2, column=c0).font = Font(bold=True, color=WHITE)
        ws.cell(row=2, column=c0).border = border_thin()
        ws.cell(row=2, column=c0).alignment = center()
        for m_i, month in enumerate(MONTHS):
            c = ws.cell(row=2, column=c0+1+m_i, value=month)
            c.fill = fill(BLUE_LIGHT)
            c.font = Font(bold=True, color=BLUE_DARK, size=9)
            c.border = border_thin()
            c.alignment = center()
        ws.cell(row=2, column=c0+19, value="TOTAL 18M").fill = fill(GOLD)
        ws.cell(row=2, column=c0+19).font = Font(bold=True, color=BLUE_DARK)
        ws.cell(row=2, column=c0+19).border = border_thin()
        ws.cell(row=2, column=c0+19).alignment = center()

        rows_data = {}
        row_labels = [
            ("Horas B2B facturadas", "h_b2b"),
            ("Horas B2C facturadas", "h_b2c"),
            ("Ingresos B2B (s/IVA)", "rev_b2b"),
            ("Ingresos B2C (s/IVA)", "rev_b2c"),
            ("TOTAL INGRESOS", "total_rev"),
            ("— Coste laboral empleados", "labor"),
            ("— Cuota autónomo titular", "owner_ss"),
            ("— Materiales y transporte", "mat"),
            ("— Estructura fija", "struct"),
            ("TOTAL COSTES", "total_cost"),
            ("EBITDA / BENEFICIO NETO", "ebitda"),
            ("Margen EBITDA (%)", "margin"),
            ("Break-even (h/mes necesarias)", "bep"),
        ]

        monthly_vals = {k: [] for _, k in row_labels}

        for m_i in range(18):
            avail_h = AVAIL[n_emp]
            ramp = RAMP[m_i]
            billed = avail_h * ramp
            h_b2b = billed * B2B_PCT
            h_b2c = billed * B2C_PCT
            r_b2b = h_b2b * rev_b2b
            r_b2c = h_b2c * rev_b2c
            tot_rev = r_b2b + r_b2c

            labor  = n_emp * LABOR_PER_EMP
            o_ss   = owner_cost[m_i]
            mat    = billed * (DIRECT["b2b"]["Materiales/productos"] + DIRECT["b2b"]["Desplazamiento"])
            struct = STRUCTURE_TOTAL
            tot_cost = labor + o_ss + mat + struct
            ebitda = tot_rev - tot_cost
            margin = ebitda / tot_rev if tot_rev > 0 else 0
            # BEP: hours where rev covers fixed (struct+labor+ownercost)
            fixed = struct + labor + o_ss
            var_margin_h = (rev_b2b * B2B_PCT + rev_b2c * B2C_PCT) - dir_cost_h
            bep = fixed / var_margin_h if var_margin_h > 0 else 0

            for key, val in zip([k for _,k in row_labels],
                                 [h_b2b, h_b2c, r_b2b, r_b2c, tot_rev,
                                  -labor, -o_ss, -mat, -struct,
                                  -(labor+o_ss+mat+struct), ebitda, margin, bep]):
                monthly_vals[key].append(val)

        # Write rows
        for r_idx, (label, key) in enumerate(row_labels):
            row = 3 + r_idx
            is_total = label.startswith(("TOTAL", "EBITDA"))
            is_pct   = key == "margin"
            is_h     = key in ("h_b2b", "h_b2c", "bep")
            is_neg   = key in ("labor", "owner_ss", "mat", "struct", "total_cost")

            c = ws.cell(row=row, column=c0, value=label)
            c.fill = fill(BLUE_LIGHT if is_total else WHITE)
            c.font = Font(bold=is_total, color=BLUE_DARK, size=9)
            c.border = border_thin()

            vals = monthly_vals[key]
            total_val = sum(vals) if not is_pct and not is_h else vals[-1]

            for m_i, v in enumerate(vals):
                cell = ws.cell(row=row, column=c0+1+m_i, value=v)
                cell.border = border_thin()
                cell.alignment = right()
                if is_pct:
                    cell.number_format = PCT
                elif is_h:
                    cell.number_format = '#,##0.0 "h"'
                else:
                    cell.number_format = EURO
                if is_total and key == "ebitda":
                    cell.fill = fill(GREEN if v >= 0 else "FFCCCC")
                    cell.font = Font(bold=True, color=BLUE_DARK, size=9)

            # Total column
            tc = ws.cell(row=row, column=c0+19, value=total_val if not is_pct else vals[-1])
            tc.border = border_thin()
            tc.alignment = right()
            tc.fill = fill(GOLD)
            tc.font = Font(bold=True, color=BLUE_DARK, size=9)
            if is_pct:
                tc.number_format = PCT
            elif is_h:
                tc.number_format = '#,##0.0 "h"'
            else:
                tc.number_format = EURO

        ws.column_dimensions[get_column_letter(c0)].width = 30
        for m_i in range(19):
            ws.column_dimensions[get_column_letter(c0+1+m_i)].width = 10

    return ws


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    build_escandallo(wb)
    build_clientes(wb)
    build_planificacion(wb)
    build_facturacion(wb)
    build_cuenta_resultados(wb)

    path = "/home/user/cv-linkedin/brillopal/BrillOPal_Gestion.xlsx"
    wb.save(path)
    print(f"Saved: {path}")

if __name__ == "__main__":
    main()
