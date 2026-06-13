#!/usr/bin/env python3
"""Generate BrillOPal_Plan_Negocio.docx"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

BLUE = RGBColor(0x1F, 0x38, 0x64)
GOLD = RGBColor(0xF4, 0xB9, 0x42)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    run = p.runs[0]
    run.font.color.rgb = BLUE if level > 1 else RGBColor(0xFF, 0xFF, 0xFF)
    if level == 1:
        p.paragraph_format.space_before = Pt(14)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F3864')
        p._p.get_or_add_pPr().append(shd)
    return p

def add_para(doc, text, bold=False, italic=False, color=None, size=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    if size:
        run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'
    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(cell, '1F3864')
    # Data rows
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i+1].cells[c_i]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if r_i % 2 == 1:
                set_cell_bg(cell, 'D6E4F0')
    # Column widths
    if col_widths:
        for i, w in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    return table

def main():
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── PORTADA ──────────────────────────────────────────────────────────────
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("BRILLOPAL")
    r.bold = True
    r.font.size = Pt(36)
    r.font.color.rgb = BLUE

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = subtitle.add_run("PLAN DE NEGOCIO COMPLETO")
    r2.bold = True
    r2.font.size = Pt(18)
    r2.font.color.rgb = GOLD

    sub2 = doc.add_paragraph()
    sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub2.add_run("Servicios de Limpieza, Mantenimiento y Auxiliares\n"
                 "Palencia y Provincia · 2026").font.size = Pt(12)
    doc.add_paragraph()

    tagline = doc.add_paragraph()
    tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = tagline.add_run("\"La limpieza que tu empresa merece, sin la complejidad que no necesita\"")
    r3.bold = True
    r3.italic = True
    r3.font.size = Pt(13)
    r3.font.color.rgb = BLUE

    doc.add_page_break()

    # ── 1. RESUMEN EJECUTIVO ──────────────────────────────────────────────────
    add_heading(doc, "1. RESUMEN EJECUTIVO")
    add_para(doc,
        "BRILLOPAL es una empresa de servicios de limpieza y mantenimiento con sede en Palencia, "
        "orientada a la subcontratación B2B (línea principal) y al servicio directo B2C "
        "(mercado complementario). Su modelo de negocio se basa en prestar un servicio de calidad "
        "superior a la media del sector, con personal propio cualificado, retribuido un 3% por "
        "encima del Convenio Colectivo provincial, y con procesos estandarizados que garantizan "
        "consistencia y cumplimiento legal.", size=10)

    add_table(doc,
        ["Parámetro", "Dato"],
        [
            ("Localización", "Palencia y provincia (radio 40 km)"),
            ("Líneas de negocio", "B2B (subcontratación) + B2C (particulares/microempresas)"),
            ("Convenio aplicable", "Conv. Col. Limpieza Edif. y Locales Palencia (en ultraactividad 2026)"),
            ("Jornada laboral", "37,5 h/semana · 1.687,5 h/año"),
            ("Salario base de referencia", "1.107,22 €/mes × 14 pagas (Limpiador/a Cat. IV, 2025)"),
            ("Prima salarial BrillOPal", "+3% s/tablas convenio → 1.140,44 €/mes"),
            ("Coste empresa/empleado/año", "≈ 21.042 € (salario+SS 31,78%)"),
            ("Precio venta B2B/hora", "≈ 20,41 €/h (s/IVA) · Margen +30% s/coste directo"),
            ("Precio venta B2C/hora", "≈ 22,50 €/h (s/IVA)"),
            ("Forma jurídica recomendada", "Autónomo (inicio) → SL a partir de 3 empleados"),
            ("Break-even 1 empleado", "≈ 87 h/mes facturadas"),
            ("Break-even 3 empleados", "≈ 200 h/mes facturadas"),
        ],
        [5, 10])

    doc.add_page_break()

    # ── 2. DESCRIPCIÓN DEL NEGOCIO ────────────────────────────────────────────
    add_heading(doc, "2. DESCRIPCIÓN DEL NEGOCIO")

    add_heading(doc, "2.1 Misión y propuesta de valor", level=2)
    add_para(doc,
        "BRILLOPAL ofrece a empresas y particulares de Palencia servicios de limpieza, "
        "mantenimiento y auxiliares con tres pilares diferenciales:", size=10)
    for b in [
        "Fiabilidad contractual: servicio garantizado, sustituciones en 24h, acuerdos SLA.",
        "Transparencia total: escandallo de costes abierto, precios cerrados sin sorpresas.",
        "Cumplimiento legal pleno: PRL, CAE, subrogación, convenio actualizado.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "2.2 Servicios ofrecidos", level=2)
    add_table(doc,
        ["Línea", "Servicio", "Frecuencia típica", "Cliente objetivo"],
        [
            ("B2B", "Limpieza de naves industriales", "Diaria / Semanal", "PYME industrial Palencia"),
            ("B2B", "Limpieza de oficinas y despachos", "Diaria / Alterna", "Empresas medianas"),
            ("B2B", "Limpieza de comunidades de propietarios", "Diaria / 3/sem", "Administradoras de fincas"),
            ("B2B", "Fin de obra / limpieza post-construcción", "Puntual", "Constructoras / promotoras"),
            ("B2B", "Servicios auxiliares (conserje, recepción)", "Jornada completa", "Empresas con sede"),
            ("B2C", "Limpieza de hogar", "Quincenal / Semanal", "Familias Palencia capital"),
            ("B2C", "Limpieza de pequeñas oficinas (<200 m²)", "Semanal", "Autónomos / microempresas"),
        ],
        [1.5, 5.5, 3.5, 4.5])

    doc.add_page_break()

    # ── 3. MODELO B2B: SUBCONTRATACIÓN ────────────────────────────────────────
    add_heading(doc, "3. MODELO B2B: SUBCONTRATACIÓN DE SERVICIOS")

    add_heading(doc, "3.1 Cómo funciona la subcontratación en limpieza", level=2)
    add_para(doc,
        "En el modelo B2B, BRILLOPAL actúa como contratista de servicios: la empresa cliente "
        "(empresa principal) subcontrata la limpieza y externaliza la gestión del personal "
        "asociado. BRILLOPAL contrata directamente a los empleados, gestiona nóminas, "
        "SS, formación PRL y sustituciones. La empresa cliente paga una tarifa por horas "
        "de servicio y se desentiende de toda la carga laboral.", size=10)

    add_heading(doc, "3.2 Cómo captar empresas cliente B2B", level=2)
    for b in [
        "Visita comercial directa a polígonos industriales de Palencia (P.I. La Mora, P.I. Villalonquéjar).",
        "Alianza con administradores de fincas colegiados (CAFI Palencia): alto volumen de comunidades.",
        "Propuesta a constructoras y promotoras locales para contratos de fin de obra.",
        "Participación en la Cámara de Comercio de Palencia y eventos FEPAC.",
        "Google My Business + web básica con formulario de solicitud de presupuesto online.",
        "Ofrecer primeros 2 días gratis de prueba a clientes B2B potenciales (coste asumible: ~160€).",
        "Referidos: bonificación 50€ por empresa cliente referida que formalice contrato.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "3.3 Propuesta de valor frente a grandes empresas", level=2)
    add_table(doc,
        ["Factor", "Grandes empresas (EULEN, ISS, Clece…)", "BRILLOPAL"],
        [
            ("Agilidad", "Burocrática, procesos lentos", "Respuesta en 2h, gestor único"),
            ("Proximidad", "Oficina central Madrid/Barcelona", "Palencia, presencia física"),
            ("Flexibilidad", "Contratos mínimos 12 meses, 40h/mes", "Desde 10h/mes, contratos 3 meses"),
            ("Precio", "Similar o superior por overhead corporativo", "10-15% más económico s/calidad equivalente"),
            ("Subrogación", "Plantilla en distintas provincias, compleja", "Gestión local directa"),
            ("Control calidad", "Supervisor visita 1/mes", "Gestor en contacto semanal"),
            ("CAE/PRL", "Departamento centralizado lento", "Documentación lista en 48h"),
        ],
        [3.5, 6, 6])

    doc.add_page_break()

    # ── 4. MARCO LEGAL ────────────────────────────────────────────────────────
    add_heading(doc, "4. MARCO LEGAL Y OBLIGACIONES")

    add_heading(doc, "4.1 Convenio Colectivo aplicable", level=2)
    add_para(doc,
        "El convenio aplicable es el Convenio Colectivo Provincial de Limpieza de Edificios "
        "y Locales de Palencia (código 34000275011982), con vigencia 01/01/2022–31/12/2025. "
        "En 2026 se encuentra en situación de ultraactividad conforme al art. 86.3 ET: "
        "sus condiciones siguen siendo de aplicación íntegra hasta que se firme el nuevo "
        "convenio o se denuncie formalmente.", size=10)
    add_para(doc,
        "Categorías y salario base mensual 2025 (14 pagas):", bold=True, size=10)
    add_table(doc,
        ["Cat.", "Denominación", "Sal. base/mes\nConvenio 2025", "Sal. base/mes\n+3% BrillOPal", "Sal. anual\n+3%"],
        [
            ("I", "Encargado/a General", "1.375,00 €", "1.416,25 €", "19.827,50 €"),
            ("II", "Encargado/a de Zona", "1.240,00 €", "1.277,20 €", "17.880,80 €"),
            ("III", "Especialista (cristalero…)", "1.165,00 €", "1.199,95 €", "16.799,30 €"),
            ("IV", "Limpiador/a", "1.107,22 €", "1.140,44 €", "15.966,16 €"),
            ("V", "Peón Auxiliar", "1.107,22 €", "1.140,44 €", "15.966,16 €"),
        ],
        [1, 5, 3, 3, 3])

    add_heading(doc, "4.2 Subrogación de plantilla", level=2)
    add_para(doc,
        "El sector de limpieza tiene obligación legal de subrogación de personal cuando se "
        "produce cambio de contratista (art. 44 ET y cláusula específica del convenio). "
        "Esto implica:", size=10)
    for b in [
        "Cuando BrillOPal gana un contrato antes servido por otra empresa, debe asumir "
        "al personal adscrito a ese contrato en las mismas condiciones (salario, categoría, antigüedad).",
        "Al perder un contrato, el nuevo adjudicatario está obligado a absorber la plantilla de BrillOPal.",
        "Acción recomendada: en cada propuesta B2B, solicitar al cliente información sobre "
        "personal actualmente adscrito al servicio y antigüedad. Calcular el coste real del "
        "contrato con la subrogación antes de firmar.",
        "Riesgo clave: empleados con alta antigüedad tienen coste superior al convenio mínimo.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.3 Prevención de Riesgos Laborales (Ley 31/1995) y CAE", level=2)
    add_para(doc,
        "Obligaciones PRL para una empresa de limpieza con personal propio:", size=10)
    for b in [
        "Plan de Prevención de Riesgos Laborales obligatorio desde el 1er empleado.",
        "Evaluación de riesgos por puesto de trabajo (exposición a productos químicos, "
        "ergonomía, riesgos eléctricos, caídas).",
        "Formación PRL básica: mínimo 20 horas (Nivel Básico Sector Servicios) para todos "
        "los empleados, antes de comenzar a trabajar.",
        "Vigilancia de la salud: reconocimiento médico anual ofrecido (voluntario para el trabajador).",
        "Servicio de Prevención Ajeno (SPA): para empresas pequeñas se puede contratar con "
        "una mutua o SPA homologado por ~300-500€/año.",
        "CAE – Coordinación de Actividades Empresariales (RD 171/2004): cuando BrillOPal "
        "trabaja en las instalaciones de un cliente, debe intercambiar documentación CAE: "
        "evaluación de riesgos, póliza RC, fichas de seguridad de productos, formación de "
        "los trabajadores. La empresa cliente lo gestiona habitualmente via plataformas "
        "como Navisio, Ctaima o Ecoordina (coste: ~0 si el cliente lo asume, o ~200€/año "
        "si BrillOPal contrata plataforma propia).",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.4 REA – Registro de Empresas Acreditadas", level=2)
    add_para(doc,
        "El REA (Registro de Empresas Acreditadas) es obligatorio para empresas que "
        "subcontratan en el SECTOR DE LA CONSTRUCCIÓN (RD 1109/2007, Ley 32/2006). "
        "Para los servicios de limpieza de edificios en uso (Cnae 8121/8122), el REA "
        "NO es obligatorio salvo que se presten servicios de limpieza en obras de "
        "construcción activas. Si BrillOPal accede a contratos de fin de obra o limpieza "
        "en zonas de construcción, debe tramitar el REA: inscripción gratuita en la "
        "Consejería de Trabajo de CyL.", size=10)

    add_heading(doc, "4.5 Seguros obligatorios y recomendados", level=2)
    add_table(doc,
        ["Seguro", "¿Obligatorio?", "Cobertura", "Coste estimado/año"],
        [
            ("RC Profesional y Explotación", "Sí (exigido por clientes B2B)", "Daños a terceros durante el servicio · mínimo 300.000€", "800–1.200 €"),
            ("AT Empleados (Mutua)", "Sí (cotización SS)", "Accidentes de trabajo, incluido en cuota SS (AT/EP 1,30%)", "Incluido en SS"),
            ("RC Patronal", "Recomendado", "Daños al trabajador no cubiertos por SS", "200–400 €"),
            ("Seguro de vida/colectivo", "No (opcional)", "Fallecimiento o invalidez del trabajador", "300–600 €"),
            ("Seguro vehículo de empresa", "Obligatorio si hay vehículo", "Todo riesgo recomendado para vehículo de trabajo", "700–1.000 €"),
        ],
        [4, 2.5, 6.5, 3])

    doc.add_page_break()

    # ── 5. FORMA JURÍDICA ─────────────────────────────────────────────────────
    add_heading(doc, "5. ESTRUCTURA JURÍDICA RECOMENDADA")

    add_heading(doc, "5.1 Autónomo vs Sociedad Limitada", level=2)
    add_table(doc,
        ["Criterio", "Autónomo", "Sociedad Limitada (SL)"],
        [
            ("Capital mínimo", "0 €", "1 € (reforma 2023, antes 3.000 €)"),
            ("Responsabilidad", "ILIMITADA (patrimonio personal)", "Limitada al capital social"),
            ("Cuota SS titular", "80 €/mes (1er año tarifa plana) → hasta ~320 €/mes", "Autónomo societario: mínimo ~320 €/mes"),
            ("Retribución titular", "Libre (retirada de beneficios)", "Nómina fija + dividendos (más rígido)"),
            ("Tributación beneficios", "IRPF progresivo (hasta 47%)", "IS al 23% (pyme) o 15% primeros 2 años"),
            ("Imagen corporativa B2B", "Menor (percepción de pequeño)", "Mayor (inspira más confianza)"),
            ("Trámites constitución", "1 día, coste mínimo", "1-2 semanas, notaría ~300-600 €"),
            ("Contabilidad", "Libro de ingresos/gastos", "Contabilidad mercantil completa"),
            ("Recomendado cuando", "Inicio / 1-2 empleados", "≥3 empleados / beneficio >30.000 €/año"),
        ],
        [4.5, 5.5, 5.5])

    add_heading(doc, "5.2 Recomendación para BrillOPal", level=2)
    add_para(doc,
        "FASE 1 (meses 1-12): constituirse como AUTÓNOMO. La tarifa plana de 80 €/mes el "
        "primer año reduce drásticamente el coste fijo de estructura. El modelo B2B de "
        "subcontratación con personal contratado es perfectamente compatible con la figura "
        "del autónomo empleador. La responsabilidad ilimitada es el único riesgo: se mitiga "
        "con un seguro de RC adecuado.", size=10)
    add_para(doc,
        "FASE 2 (a partir de 3 empleados o beneficio >25.000 €/año): transformar en SL. "
        "Ventajas: tributación IS al 15% (primeros 2 años), protección de patrimonio personal, "
        "mejor imagen para contratos grandes B2B.", size=10)

    doc.add_page_break()

    # ── 6. OBLIGACIONES FISCALES Y LABORALES ──────────────────────────────────
    add_heading(doc, "6. OBLIGACIONES FISCALES Y LABORALES")

    add_heading(doc, "6.1 Obligaciones fiscales (autónomo)", level=2)
    add_table(doc,
        ["Impuesto/Modelo", "Periodicidad", "Descripción"],
        [
            ("Modelo 130 – IRPF pago fraccionado", "Trimestral (abr/jul/oct/ene)", "20% del rendimiento neto trimestral"),
            ("Modelo 303 – IVA autoliquidación", "Trimestral", "IVA repercutido – IVA soportado. Tipo 21%"),
            ("Modelo 390 – Resumen anual IVA", "Enero (año siguiente)", "Resumen de todas las liquidaciones trimestrales"),
            ("Modelo 100 – IRPF anual", "Mayo-Junio", "Declaración anual. Deduce los pagos fraccionados"),
            ("Modelo 111 – Retenciones IRPF trabajadores", "Trimestral", "Retenciones practicadas en nóminas"),
            ("Modelo 190 – Resumen anual retenciones", "Enero", "Resumen anual modelo 111"),
            ("Cuota autónomo SS", "Mensual (domiciliación)", "Según base de cotización elegida"),
        ],
        [5, 3.5, 7])

    add_heading(doc, "6.2 Obligaciones laborales", level=2)
    for b in [
        "Alta previa en Seguridad Social como empleador (TGSS) antes de contratar el primer empleado.",
        "Alta del trabajador en SS (TGSS) antes del primer día de trabajo (modalidad SEPE-SISPE).",
        "Contrato de trabajo por escrito: tipo recomendado → indefinido a tiempo parcial o completo "
        "(aporta estabilidad y reduce rotación, clave en sector limpieza).",
        "Nómina mensual: desglose salario base + pagas extras prorrateadas + SS trabajador descontada.",
        "TC1/TC2 (ahora Sistema RED): liquidación de cuotas SS cada mes, antes del último día.",
        "Comunicación al SEPE de contratos a tiempo parcial (≤10 días hábiles).",
        "Registro de jornada (obligatorio desde 2019, art. 34.9 ET): hora de entrada y salida diaria "
        "por escrito o digital.",
        "Calificación del riesgo de AT (epígrafe limpieza): comunicar a mutua colaboradora.",
    ]:
        add_bullet(doc, b)

    doc.add_page_break()

    # ── 7. PLAN DE CONTRATACIÓN ───────────────────────────────────────────────
    add_heading(doc, "7. PLAN DE CONTRATACIÓN 18 MESES")

    add_table(doc,
        ["Período", "Acción", "Coste mensual añadido"],
        [
            ("Mes 1-2", "Constitución. Sin empleados. Prueba de mercado con subcontratas externas "
             "para primeros servicios.", "0 € labor directa"),
            ("Mes 3", "Contratación Empleado 1 (Limpiador/a Cat. IV, 37,5h/sem). "
             "Primer contrato B2B firmado.", "+1.756 €/mes"),
            ("Mes 6-7", "Si ocupación >75%: contratación Empleado 2 (Limpiador/a). "
             "Ampliación cartera B2B.", "+1.756 €/mes"),
            ("Mes 10-12", "Si contratos B2B suman >400h/mes: contratación Empleado 3 "
             "(Especialista Cat. III o 2º Limpiador/a).", "+1.756-1.872 €/mes"),
            ("Mes 13-18", "Plantilla estabilizada 3-5 empleados. Valorar transformación en SL. "
             "Objetivo: ocupación 80-90%.", "Según crecimiento"),
        ],
        [2.5, 9.5, 3.5])

    add_heading(doc, "7.1 Tipo de contrato recomendado", level=2)
    add_para(doc,
        "Contrato INDEFINIDO ORDINARIO a jornada completa (37,5h/sem conforme convenio). "
        "Razones:", size=10)
    for b in [
        "El sector limpieza tiene alta rotación; el contrato indefinido fideliza y reduce "
        "coste de selección/formación.",
        "En subrogación: los contratos indefinidos del cedente pasan al nuevo adjudicatario "
        "(obligación legal); los temporales pueden quedar excluidos según convenio.",
        "Bonificación contratación indefinida jornada completa (SEPE 2026): hasta 138 €/mes "
        "si la persona proviene de situación de desempleo.",
    ]:
        add_bullet(doc, b)

    doc.add_page_break()

    # ── 8. PLAN FINANCIERO RESUMEN ────────────────────────────────────────────
    add_heading(doc, "8. PLAN FINANCIERO RESUMEN")

    add_heading(doc, "8.1 Estructura de costes y precio de venta", level=2)
    add_table(doc,
        ["Concepto", "B2B (€/h)", "B2C (€/h)", "% del precio"],
        [
            ("Coste laboral (salario + SS, Cat. IV)", "12,47", "12,47", "61%"),
            ("Provisión absentismo (3,5%)", "0,44", "0,44", "2%"),
            ("Materiales y productos", "0,40", "0,55", "2-3%"),
            ("Desplazamiento", "0,25", "0,45", "1-2%"),
            ("Estructura asignada (3 empleados)", "3,38", "3,38", "17%"),
            ("COSTE TOTAL / HORA", "16,94", "17,29", "83%"),
            ("BENEFICIO / HORA (+30%)", "5,08", "5,19", "25%"),
            ("PRECIO VENTA / HORA (s/IVA)", "22,02", "22,47", "100%"),
            ("PRECIO VENTA / HORA (c/IVA 21%)", "26,64", "27,19", "—"),
        ],
        [6, 3, 3, 3.5])

    add_heading(doc, "8.2 Punto de equilibrio (Break-Even)", level=2)
    add_table(doc,
        ["Escenario", "Costes fijos/mes", "Margen variable/h", "Break-Even h/mes"],
        [
            ("1 empleado", "3.186 €", "7,19 €/h", "≈ 87 h/mes (62% ocupación)"),
            ("3 empleados", "6.698 €", "7,19 €/h", "≈ 200 h/mes (48% ocupación)"),
            ("5 empleados", "10.211 €", "7,19 €/h", "≈ 287 h/mes (41% ocupación)"),
        ],
        [4, 4, 4, 4])
    add_para(doc, "Nota: margen variable/h = precio promedio – costes variables (labor+SS+mat+desp). "
             "Los costes fijos incluyen estructura + labor empleados + cuota autónomo.", italic=True, size=9)

    doc.add_page_break()

    # ── 9. ANÁLISIS DE RIESGOS ────────────────────────────────────────────────
    add_heading(doc, "9. ANÁLISIS DE RIESGOS Y MITIGACIÓN")

    add_table(doc,
        ["Riesgo", "Probabilidad", "Impacto", "Mitigación"],
        [
            ("Pérdida de contrato B2B principal", "Media", "Alto",
             "Nunca depender >40% de un solo cliente; diversificar desde el mes 6"),
            ("Rotación de empleados", "Alta", "Medio",
             "Prima salarial +3%, contratos indefinidos, buen clima laboral"),
            ("Subrogación con plantilla cara (antigüedad alta)", "Media", "Alto",
             "Due diligence previa; calcular sobrecoste antes de propuesta"),
            ("Morosidad clientes B2C", "Baja-Media", "Bajo",
             "Cobro domiciliado o anticipado; mínimo 2 servicios de depósito"),
            ("Accidente laboral sin SPA", "Baja", "Muy alto",
             "Contratar SPA desde el 1er empleado; formación PRL obligatoria"),
            ("Nuevo convenio con subida salarial >5%", "Media", "Medio",
             "Cláusula de revisión de precio en contratos B2B (anual s/IPC+conveio)"),
            ("Competencia precio de grandes empresas", "Alta", "Bajo",
             "Diferenciación por servicio, no por precio; target PYME, no multinacional"),
        ],
        [4, 2, 2, 7.5])

    doc.save("/home/user/cv-linkedin/brillopal/BrillOPal_Plan_Negocio.docx")
    print("Saved: BrillOPal_Plan_Negocio.docx")

if __name__ == "__main__":
    main()
