"""
╔══════════════════════════════════════════════════════════════════════════════╗
║           DISPOWER · Generador de Tarifas ZNI                               ║
║           Streamlit + ReportLab · Versión 1.0                               ║
╚══════════════════════════════════════════════════════════════════════════════╝

Requisitos:
    pip install streamlit pandas openpyxl reportlab pillow

Ejecución:
    streamlit run app_tarifas.py
"""

import io
import os
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTES Y CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

# Columnas obligatorias del Excel
REQUIRED_COLUMNS = [
    "AÑO", "MES", "DEPARTAMENTO", "MUNICIPIO", "TIPO SISTEMA",
    "ALMACENAMIENTO", "WHD", "IPP_0", "IPP_m-1",
    "AMGCnu_m", "AMGCvi_m", "AMGCau_m", "AMGCnf_m", "AMGCro_m", "AMGC_m",
    "Im", "CUm", "Comercializador SIN", "Tarifa SIN",
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

PAGE_W, PAGE_H = letter  # 612 × 792 pts
MARGIN = 1.8 * cm
CONTENT_W = PAGE_W - 2 * MARGIN


# ═══════════════════════════════════════════════════════════════════════════════
#  PALETA DE COLORES (con soporte para colores corporativos personalizados)
# ═══════════════════════════════════════════════════════════════════════════════

def build_palette(primary: str, secondary: str, accent: str) -> dict:
    """Construye la paleta de colores a partir de colores hexadecimales."""
    return {
        "primary":    colors.HexColor(primary),
        "secondary":  colors.HexColor(secondary),
        "accent":     colors.HexColor(accent),
        "dark":       colors.HexColor("#0E2841"),
        "white":      colors.white,
        "light_grey": colors.HexColor("#F2F5F8"),
        "border":     colors.HexColor("#D0D8E0"),
        "text_muted": colors.HexColor("#6B7280"),
        "text_dark":  colors.HexColor("#1A202C"),
        "green":      colors.HexColor("#196B24"),
    }


DEFAULT_PALETTE = build_palette("#156082", "#E97132", "#0F9ED5")


# ═══════════════════════════════════════════════════════════════════════════════
#  UTILIDADES
# ═══════════════════════════════════════════════════════════════════════════════

def fmt_num(v, decimals: int = 2) -> str:
    """Formatea un número con separadores de miles al estilo colombiano."""
    try:
        num = float(v)
        if decimals == 0:
            return f"{int(round(num)):,}".replace(",", ".")
        formatted = f"{num:,.{decimals}f}"
        # Convierte coma de miles y punto decimal al español
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(v)


def p(text: str, style: ParagraphStyle) -> Paragraph:
    """Shorthand para crear un Paragraph."""
    return Paragraph(str(text), style)


# ═══════════════════════════════════════════════════════════════════════════════
#  ESTILOS REPORTLAB
# ═══════════════════════════════════════════════════════════════════════════════

def build_styles(pal: dict) -> dict:
    """Genera el diccionario de estilos tipográficos para el PDF."""
    base = getSampleStyleSheet()

    def s(name, **kw) -> ParagraphStyle:
        return ParagraphStyle(name, parent=base["Normal"], **kw)

    return {
        "title": s("title",
            fontSize=20, textColor=pal["dark"], alignment=TA_CENTER,
            fontName="Helvetica-Bold", leading=26, spaceAfter=2),

        "subtitle": s("subtitle",
            fontSize=9, textColor=pal["primary"], alignment=TA_CENTER,
            fontName="Helvetica-Bold", leading=13),

        "legal": s("legal",
            fontSize=8, textColor=pal["text_dark"], fontName="Helvetica",
            alignment=TA_CENTER, leading=11),

        "section_hdr": s("section_hdr",
            fontSize=8, textColor=pal["white"], fontName="Helvetica-Bold",
            alignment=TA_CENTER, leading=12),

        "tbl_hdr": s("tbl_hdr",
            fontSize=7.5, textColor=pal["white"], fontName="Helvetica-Bold",
            alignment=TA_CENTER, leading=10),

        "tbl_val": s("tbl_val",
            fontSize=8, textColor=pal["text_dark"], fontName="Helvetica",
            alignment=TA_CENTER, leading=11),

        "tbl_val_bold": s("tbl_val_bold",
            fontSize=8.5, textColor=pal["text_dark"], fontName="Helvetica-Bold",
            alignment=TA_CENTER, leading=11),

        "footnote": s("footnote",
            fontSize=6.5, textColor=pal["text_muted"], fontName="Helvetica",
            alignment=TA_CENTER, leading=9),

        "formula": s("formula",
            fontSize=7.5, textColor=pal["text_dark"], fontName="Helvetica",
            alignment=TA_CENTER, leading=11),
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS PARA TABLAS
# ═══════════════════════════════════════════════════════════════════════════════

def hdr_cell(text: str, pal: dict, bg_key: str = "primary") -> Paragraph:
    st = ParagraphStyle("hc", parent=getSampleStyleSheet()["Normal"],
        fontSize=7.5, textColor=pal["white"], fontName="Helvetica-Bold",
        alignment=TA_CENTER, leading=10)
    return Paragraph(text, st)


def val_cell(text: str, bold: bool = False, color=None, fs: int = 8,
             align=TA_CENTER) -> Paragraph:
    fn = "Helvetica-Bold" if bold else "Helvetica"
    c = color or colors.HexColor("#1A202C")
    st = ParagraphStyle("vc", parent=getSampleStyleSheet()["Normal"],
        fontSize=fs, textColor=c, fontName=fn, alignment=align, leading=11)
    return Paragraph(str(text), st)


def section_banner(text: str, bg_color, content_w: float) -> Table:
    """Barra de sección con color de fondo."""
    st = ParagraphStyle("sb", parent=getSampleStyleSheet()["Normal"],
        fontSize=8, textColor=colors.white, fontName="Helvetica-Bold",
        alignment=TA_CENTER, leading=12)
    t = Table([[Paragraph(text, st)]], colWidths=[content_w])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg_color),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


def std_table_style(pal: dict, n_header_cols_alt: list = None) -> TableStyle:
    """Estilo estándar para tablas de datos."""
    return TableStyle([
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [pal["white"], pal["light_grey"]]),
        ("GRID", (0, 0), (-1, -1), 0.5, pal["border"]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ])


# ═══════════════════════════════════════════════════════════════════════════════
#  GENERACIÓN DE PÁGINA POR MUNICIPIO
# ═══════════════════════════════════════════════════════════════════════════════

def build_municipality_page(
    rows: pd.DataFrame,
    pal: dict,
    styles: dict,
    header_img_path: str | None,
    footer_img_path: str | None,
    show_subsidio: bool = False,
) -> list:
    """
    Genera la lista de Flowables de ReportLab para un municipio.

    Args:
        rows: DataFrame filtrado para un municipio.
        pal: Paleta de colores.
        styles: Diccionario de estilos tipográficos.
        header_img_path: Ruta al PNG de encabezado (o None).
        footer_img_path: Ruta al PNG de pie de página (o None).
        show_subsidio: Si True, incluye columnas de subsidio y tarifa usuario.

    Returns:
        Lista de Flowables para añadir al story principal.
    """
    story = []
    row0 = rows.iloc[0]
    mun  = row0["MUNICIPIO"]
    dept = row0["DEPARTAMENTO"]
    year = int(row0["AÑO"])
    mes  = MESES_ES.get(int(row0["MES"]), str(row0["MES"]))
    cw   = CONTENT_W

    # ── 1. IMAGEN ENCABEZADO ─────────────────────────────────────────────────
    if header_img_path:
        with PILImage.open(header_img_path) as im:
            iw, ih = im.size
        ratio = ih / iw
        story.append(RLImage(header_img_path, width=cw, height=cw * ratio))
        story.append(Spacer(1, 6))
    else:
        # Encabezado de texto si no hay imagen
        banner = Table([[
            p("⚡ DISPOWER SAS ESP", ParagraphStyle("bh",
                fontSize=14, textColor=pal["white"], fontName="Helvetica-Bold",
                alignment=TA_CENTER, leading=18)),
            p("Energía solar · ZNI", ParagraphStyle("bhs",
                fontSize=9, textColor=pal["light_grey"], fontName="Helvetica",
                alignment=TA_CENTER, leading=12)),
        ]], colWidths=[cw * 0.65, cw * 0.35])
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), pal["dark"]),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(banner)
        story.append(Spacer(1, 6))

    # ── 2. TÍTULO ────────────────────────────────────────────────────────────
    story.append(p("TARIFAS MÁXIMAS", styles["title"]))
    story.append(p(
        "CALCULADAS Y APLICADAS PARA DISPOWER E.S.P. · MUNICIPIOS Y LOCALIDADES",
        styles["subtitle"]))
    story.append(Spacer(1, 4))

    # ── 3. BANNER MUNICIPIO / PERIODO ────────────────────────────────────────
    banner_data = [[
        p(f"<b>{mun.upper()}</b>", ParagraphStyle("mun",
            fontSize=15, textColor=pal["white"], fontName="Helvetica-Bold",
            alignment=TA_CENTER, leading=20)),
        p(f"<b>Departamento:</b> {dept}<br/>"
          f"<b>Período:</b> {mes} {year}",
          ParagraphStyle("period",
            fontSize=8.5, textColor=pal["white"], fontName="Helvetica",
            alignment=TA_CENTER, leading=13)),
    ]]
    banner = Table(banner_data, colWidths=[cw * 0.65, cw * 0.35])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), pal["dark"]),
        ("BACKGROUND", (1, 0), (1, 0), pal["primary"]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(banner)
    story.append(Spacer(1, 6))

    # ── 4. MARCO LEGAL ───────────────────────────────────────────────────────
    legal_box = Table([[p(
        "<b>DISPOWER SAS ESP</b> informa a sus usuarios que se definen las tarifas para la "
        "prestación del servicio de energía mediante plantas solares en ZNI de acuerdo a la "
        "Resolución <b>CREG 101-026 DE 2022</b>.",
        styles["legal"])]], colWidths=[cw])
    legal_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pal["light_grey"]),
        ("BOX", (0, 0), (-1, -1), 0.8, pal["primary"]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(legal_box)
    story.append(Spacer(1, 7))

    # ── 5. PARÁMETROS BASE ───────────────────────────────────────────────────
    story.append(section_banner("PARÁMETROS BASE DEL SISTEMA", pal["dark"], cw))
    story.append(Spacer(1, 3))

    param_hdrs = [
        hdr_cell("WHD\n(Wh/día)", pal),
        hdr_cell("Tipo Sistema", pal),
        hdr_cell("Almacenamiento", pal),
        hdr_cell("IPP₀\n($/kWh)", pal),
        hdr_cell("IPPm-1\n($/kWh)", pal),
    ]
    param_rows = [param_hdrs]
    for _, r in rows.iterrows():
        param_rows.append([
            val_cell(fmt_num(r["WHD"], 0), bold=True),
            val_cell(r["TIPO SISTEMA"]),
            val_cell(r["ALMACENAMIENTO"]),
            val_cell(fmt_num(r["IPP_0"], 2)),
            val_cell(fmt_num(r["IPP_m-1"], 2)),
        ])

    t_params = Table(param_rows, colWidths=[cw / 5] * 5, repeatRows=1)
    ts = std_table_style(pal)
    ts.add("BACKGROUND", (0, 0), (-1, 0), pal["primary"])
    t_params.setStyle(ts)
    story.append(t_params)
    story.append(Spacer(1, 7))

    # ── 6. COMPONENTES AMGC ─────────────────────────────────────────────────
    story.append(section_banner(
        "COMPONENTES DEL COSTO UNITARIO — CARGO AMGC ($/usuario-mes)",
        pal["green"], cw))
    story.append(Spacer(1, 3))

    amgc_hdrs = [
        hdr_cell("AMGCnu_m\nNo Usuarios", pal),
        hdr_cell("AMGCvi_m\nVida Útil", pal),
        hdr_cell("AMGCau_m\nActualización", pal),
        hdr_cell("AMGCnf_m\nNo Facturación", pal),
        hdr_cell("AMGCro_m\nOtros", pal),
        hdr_cell("AMGC_m\nTOTAL", pal),
    ]
    amgc_rows = [amgc_hdrs]
    for _, r in rows.iterrows():
        amgc_rows.append([
            val_cell(fmt_num(r["AMGCnu_m"], 2)),
            val_cell(fmt_num(r["AMGCvi_m"], 2)),
            val_cell(fmt_num(r["AMGCau_m"], 2)),
            val_cell(fmt_num(r["AMGCnf_m"], 2)),
            val_cell(fmt_num(r["AMGCro_m"], 2)),
            val_cell(fmt_num(r["AMGC_m"], 2), bold=True, color=pal["dark"]),
        ])

    t_amgc = Table(amgc_rows, colWidths=[cw / 6] * 6, repeatRows=1)
    ts2 = std_table_style(pal)
    ts2.add("BACKGROUND", (0, 0), (-2, 0), pal["green"])
    ts2.add("BACKGROUND", (-1, 0), (-1, 0), pal["dark"])
    ts2.add("FONTNAME", (-1, 1), (-1, -1), "Helvetica-Bold")
    t_amgc.setStyle(ts2)
    story.append(t_amgc)
    story.append(Spacer(1, 7))

    # ── 7. COSTO UNITARIO Y TARIFAS ─────────────────────────────────────────
    story.append(section_banner(
        "COSTO UNITARIO Y TARIFA FINAL ($/usuario-mes)",
        pal["secondary"], cw))
    story.append(Spacer(1, 3))

    # Columnas base (siempre visibles)
    tarifa_hdrs = [
        hdr_cell("Im\n($/mes)", pal),
        hdr_cell("CUm\n($/usuario-mes)", pal),
        hdr_cell("Comercializador\nSIN", pal),
        hdr_cell("Tarifa SIN\n($/mes)", pal),
    ]
    col_widths = [cw / 4] * 4

    # Columnas opcionales de subsidio
    if show_subsidio:
        tarifa_hdrs += [
            hdr_cell("Subsidio\nAplicado", pal),
            hdr_cell("Tarifa\nUsuario", pal),
        ]
        col_widths = [cw / 6] * 6

    tarifa_rows = [tarifa_hdrs]
    for _, r in rows.iterrows():
        row_data = [
            val_cell(fmt_num(r["Im"], 2)),
            val_cell(fmt_num(r["CUm"], 2), bold=True, color=pal["secondary"]),
            val_cell(str(r["Comercializador SIN"])),
            val_cell(fmt_num(r["Tarifa SIN"], 2)),
        ]
        if show_subsidio:
            cum = float(r["CUm"])
            subsidio = max(0.0, cum - float(r["Tarifa SIN"]))
            row_data += [
                val_cell(fmt_num(subsidio, 2) if subsidio > 0 else "—"),
                val_cell("$ " + fmt_num(cum, 2), bold=True, color=pal["dark"]),
            ]
        tarifa_rows.append(row_data)

    t_tarifa = Table(tarifa_rows, colWidths=col_widths, repeatRows=1)
    ts3 = std_table_style(pal)
    ts3.add("BACKGROUND", (0, 0), (1, 0), pal["secondary"])
    ts3.add("BACKGROUND", (2, 0), (3, 0), pal["primary"])
    if show_subsidio:
        ts3.add("BACKGROUND", (4, 0), (4, 0), colors.HexColor("#6B46C1"))
        ts3.add("BACKGROUND", (5, 0), (5, 0), pal["dark"])
        ts3.add("BACKGROUND", (5, 1), (5, -1), colors.HexColor("#E8F5E9"))
        ts3.add("FONTNAME", (5, 1), (5, -1), "Helvetica-Bold")
    t_tarifa.setStyle(ts3)
    story.append(t_tarifa)
    story.append(Spacer(1, 7))

    # ── 8. CAJA FÓRMULA ──────────────────────────────────────────────────────
    formula_text = (
        "<b>Fórmula:</b>  CU<sub>m</sub> = I<sub>m</sub> + AMGC<sub>m</sub>"
        "&nbsp;&nbsp;·&nbsp;&nbsp;"
        "<b>Im:</b> Cargo máximo de inversión &nbsp;"
        "<b>AMGCm:</b> Cargo máximo de administración, mantenimiento y gestión comercial"
    )
    formula_box = Table([[p(formula_text, styles["formula"])]], colWidths=[cw])
    formula_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pal["light_grey"]),
        ("BOX", (0, 0), (-1, -1), 0.8, pal["border"]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(formula_box)
    story.append(Spacer(1, 5))

    # ── 9. NOTA AL PIE ───────────────────────────────────────────────────────
    story.append(p(
        f"Tarifas calculadas conforme a la Resolución CREG 101-026 de 2022 · "
        f"Departamento: {dept.upper()} · Zona No Interconectada (ZNI) · "
        f"Período: {mes} {year}",
        styles["footnote"]))
    story.append(Spacer(1, 5))

    # ── 10. IMAGEN PIE DE PÁGINA ─────────────────────────────────────────────
    if footer_img_path:
        with PILImage.open(footer_img_path) as im:
            iw, ih = im.size
        ratio = ih / iw
        story.append(RLImage(footer_img_path, width=cw, height=cw * ratio))
    else:
        footer = Table([[
            p("© DISPOWER SAS ESP · Todos los derechos reservados",
              ParagraphStyle("ft", fontSize=7, textColor=pal["white"],
                fontName="Helvetica", alignment=TA_CENTER, leading=10)),
        ]], colWidths=[cw])
        footer.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), pal["dark"]),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(footer)

    return story


# ═══════════════════════════════════════════════════════════════════════════════
#  GENERADOR DE PDF COMPLETO
# ═══════════════════════════════════════════════════════════════════════════════

def generate_pdf(
    df: pd.DataFrame,
    pal: dict,
    header_img_path: str | None = None,
    footer_img_path: str | None = None,
    municipios_filter: list | None = None,
    show_subsidio: bool = False,
) -> bytes:
    """
    Genera el PDF completo con una página por municipio.

    Args:
        df: DataFrame validado con todos los datos.
        pal: Paleta de colores.
        header_img_path: Ruta al PNG del encabezado.
        footer_img_path: Ruta al PNG del pie de página.
        municipios_filter: Lista de municipios a incluir (None = todos).
        show_subsidio: Si True, incluye columnas de subsidio/tarifa usuario.

    Returns:
        Bytes del PDF generado.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=0.4 * cm, bottomMargin=0.4 * cm,
    )
    styles = build_styles(pal)
    municipios = municipios_filter or df["MUNICIPIO"].unique().tolist()

    all_story = []
    for i, mun in enumerate(municipios):
        rows = df[df["MUNICIPIO"] == mun].reset_index(drop=True)
        if rows.empty:
            continue
        if i > 0:
            all_story.append(PageBreak())
        all_story.extend(build_municipality_page(
            rows, pal, styles,
            header_img_path, footer_img_path,
            show_subsidio=show_subsidio,
        ))

    doc.build(all_story)
    return buf.getvalue()


def generate_pdf_single(
    df: pd.DataFrame,
    municipio: str,
    pal: dict,
    header_img_path: str | None,
    footer_img_path: str | None,
    show_subsidio: bool = False,
) -> bytes:
    """Genera un PDF para un único municipio."""
    return generate_pdf(df, pal, header_img_path, footer_img_path,
                        municipios_filter=[municipio],
                        show_subsidio=show_subsidio)


# ═══════════════════════════════════════════════════════════════════════════════
#  VALIDACIONES
# ═══════════════════════════════════════════════════════════════════════════════

def validate_excel(df: pd.DataFrame) -> list[str]:
    """
    Valida el DataFrame y retorna lista de errores (vacía si todo OK).
    """
    errors = []

    # 1. Columnas requeridas
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        errors.append(f"❌ Columnas faltantes: {', '.join(missing)}")

    if errors:
        return errors  # No continuar si faltan columnas

    # 2. Municipios vacíos
    empty_mun = df["MUNICIPIO"].isna().sum()
    if empty_mun > 0:
        errors.append(f"❌ Hay {empty_mun} fila(s) con municipio vacío.")

    # 3. Datos por municipio
    for mun in df["MUNICIPIO"].dropna().unique():
        subset = df[df["MUNICIPIO"] == mun]
        if subset.empty:
            errors.append(f"❌ El municipio '{mun}' no tiene datos.")

    # 4. Componentes AMGC nulos
    amgc_cols = ["AMGCnu_m", "AMGCvi_m", "AMGCau_m", "AMGCnf_m", "AMGCro_m", "AMGC_m"]
    for col in amgc_cols:
        nulls = df[col].isna().sum()
        if nulls > 0:
            errors.append(f"⚠️ El componente AMGC '{col}' tiene {nulls} valor(es) nulo(s).")

    # 5. Valores numéricos en columnas clave
    num_cols = ["WHD", "IPP_0", "IPP_m-1", "Im", "CUm", "Tarifa SIN"]
    for col in num_cols:
        try:
            pd.to_numeric(df[col])
        except (ValueError, TypeError):
            errors.append(f"❌ La columna '{col}' contiene valores no numéricos.")

    return errors


# ═══════════════════════════════════════════════════════════════════════════════
#  INTERFAZ STREAMLIT
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Dispower · Generador de Tarifas ZNI",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── CSS personalizado ─────────────────────────────────────────────────────
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Sora:wght@600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        .main { background: #F8FAFC; }

        /* Encabezado principal */
        .hero {
            background: linear-gradient(135deg, #0E2841 0%, #156082 60%, #0F9ED5 100%);
            border-radius: 16px;
            padding: 32px 40px;
            margin-bottom: 28px;
            box-shadow: 0 8px 32px rgba(14,40,65,0.18);
        }
        .hero h1 {
            font-family: 'Sora', sans-serif;
            font-size: 2rem;
            font-weight: 700;
            color: white;
            margin: 0 0 6px 0;
            letter-spacing: -0.5px;
        }
        .hero p {
            color: rgba(255,255,255,0.75);
            font-size: 0.95rem;
            margin: 0;
        }
        .hero .badge {
            display: inline-block;
            background: rgba(255,255,255,0.15);
            color: white;
            border: 1px solid rgba(255,255,255,0.3);
            border-radius: 20px;
            padding: 3px 12px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-bottom: 12px;
            letter-spacing: 0.5px;
        }

        /* Cards de sección */
        .card {
            background: white;
            border-radius: 12px;
            padding: 24px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            margin-bottom: 16px;
        }
        .card-title {
            font-family: 'Sora', sans-serif;
            font-size: 0.9rem;
            font-weight: 700;
            color: #0E2841;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        /* Métricas de municipios */
        .mun-pill {
            display: inline-block;
            background: #EFF6FF;
            color: #156082;
            border: 1px solid #BFDBFE;
            border-radius: 20px;
            padding: 4px 14px;
            font-size: 0.78rem;
            font-weight: 600;
            margin: 3px;
        }

        /* Alertas */
        .alert-error {
            background: #FEF2F2;
            border-left: 4px solid #EF4444;
            border-radius: 8px;
            padding: 12px 16px;
            margin: 8px 0;
            font-size: 0.85rem;
            color: #991B1B;
        }
        .alert-success {
            background: #F0FDF4;
            border-left: 4px solid #22C55E;
            border-radius: 8px;
            padding: 12px 16px;
            font-size: 0.85rem;
            color: #166534;
        }

        /* Botones */
        .stButton > button {
            border-radius: 8px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stDownloadButton"] button {
            border-radius: 8px !important;
            font-weight: 600 !important;
        }

        /* Sidebar */
        section[data-testid="stSidebar"] {
            background: #0E2841;
        }
        section[data-testid="stSidebar"] * {
            color: white !important;
        }
        section[data-testid="stSidebar"] .stSelectbox label,
        section[data-testid="stSidebar"] .stCheckbox label,
        section[data-testid="stSidebar"] .stColorPicker label {
            color: rgba(255,255,255,0.85) !important;
            font-size: 0.85rem !important;
        }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            color: rgba(255,255,255,0.6) !important;
            font-size: 0.8rem !important;
        }
        .sidebar-section {
            border-top: 1px solid rgba(255,255,255,0.1);
            padding-top: 16px;
            margin-top: 16px;
        }

        /* Tabla preview */
        .dataframe { font-size: 0.82rem !important; }
    </style>
    """, unsafe_allow_html=True)

    # ── HERO ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="hero">
        <div class="badge">⚡ DISPOWER SAS ESP · ZNI</div>
        <h1>Generador de Tarifas</h1>
        <p>Automatización de piezas gráficas corporativas · Resolución CREG 101-026 de 2022</p>
    </div>
    """, unsafe_allow_html=True)

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configuración")

        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown("**🎨 Colores corporativos**")
        col_primary   = st.color_picker("Color primario",   "#156082")
        col_secondary = st.color_picker("Color secundario", "#E97132")
        col_accent    = st.color_picker("Color acento",     "#0F9ED5")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown("**📄 Opciones de contenido**")
        show_subsidio = st.checkbox("Incluir Subsidio y Tarifa Usuario", value=False)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown("**ℹ️ Columnas requeridas Excel**")
        for c in REQUIRED_COLUMNS:
            st.markdown(f"<small>· {c}</small>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Paleta desde los colores seleccionados
    pal = build_palette(col_primary, col_secondary, col_accent)

    # ── LAYOUT PRINCIPAL ─────────────────────────────────────────────────────
    col_left, col_right = st.columns([3, 2], gap="large")

    # ── COLUMNA IZQUIERDA: carga de archivos ──────────────────────────────────
    with col_left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📊 Archivo Excel (obligatorio)</div>',
                    unsafe_allow_html=True)

        excel_file = st.file_uploader(
            "Selecciona el archivo Excel con las tarifas",
            type=["xlsx", "xls"],
            help="El archivo debe contener todas las columnas listadas en el panel lateral.",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">🖼️ Plantilla gráfica (opcional)</div>',
                    unsafe_allow_html=True)
        st.caption("Si cargas imágenes, se usarán como encabezado y pie de página del PDF.")

        tcol1, tcol2 = st.columns(2)
        with tcol1:
            header_file = st.file_uploader(
                "Encabezado (PNG/JPG)", type=["png", "jpg", "jpeg"], key="header")
        with tcol2:
            footer_file = st.file_uploader(
                "Pie de página (PNG/JPG)", type=["png", "jpg", "jpeg"], key="footer")
        st.markdown('</div>', unsafe_allow_html=True)

    # ── COLUMNA DERECHA: estado y acciones ────────────────────────────────────
    with col_right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📋 Estado del sistema</div>',
                    unsafe_allow_html=True)

        if excel_file is None:
            st.info("⬅️ Carga el archivo Excel para comenzar.")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            # Leer Excel
            try:
                df = pd.read_excel(excel_file)
                df.columns = [str(c).strip() for c in df.columns]
            except Exception as e:
                st.error(f"Error al leer el Excel: {e}")
                st.markdown('</div>', unsafe_allow_html=True)
                st.stop()

            # Validar
            errors = validate_excel(df)

            if errors:
                for err in errors:
                    st.markdown(f'<div class="alert-error">{err}</div>',
                                unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                municipios = df["MUNICIPIO"].dropna().unique().tolist()
                year = int(df["AÑO"].iloc[0])
                mes  = MESES_ES.get(int(df["MES"].iloc[0]), "—")
                dept = df["DEPARTAMENTO"].iloc[0]

                st.markdown(f'<div class="alert-success">✅ Excel válido · {len(df)} filas cargadas</div>',
                            unsafe_allow_html=True)
                st.markdown(f"**Departamento:** {dept} &nbsp; · &nbsp; **Período:** {mes} {year}")
                st.markdown(f"**Municipios detectados ({len(municipios)}):**")
                pills = "".join(f'<span class="mun-pill">{m}</span>' for m in municipios)
                st.markdown(pills, unsafe_allow_html=True)

                # Selector de municipios
                st.markdown("---")
                st.markdown("**Filtrar municipios a incluir:**")
                selected_muns = st.multiselect(
                    "Selecciona municipios",
                    options=municipios,
                    default=municipios,
                    label_visibility="collapsed",
                )
                st.markdown('</div>', unsafe_allow_html=True)

                if not selected_muns:
                    st.warning("Selecciona al menos un municipio.")
                    st.stop()

                # Guardar imágenes en directorio temporal compatible con Windows/Mac/Linux
                import tempfile, pathlib
                _tmpdir = pathlib.Path(tempfile.gettempdir())
                header_path, footer_path = None, None
                if header_file:
                    hpath = str(_tmpdir / "dispower_header.png")
                    PILImage.open(header_file).save(hpath, "PNG")
                    header_path = hpath
                if footer_file:
                    fpath = str(_tmpdir / "dispower_footer.png")
                    PILImage.open(footer_file).save(fpath, "PNG")
                    footer_path = fpath

                # ── BOTONES DE GENERACIÓN ────────────────────────────────────
                st.markdown("---")
                st.markdown("### 📥 Descargar PDF")

                btn_col1, btn_col2 = st.columns(2)

                with btn_col1:
                    if st.button("🗂️ Generar PDF completo", use_container_width=True,
                                 type="primary"):
                        with st.spinner("Generando PDF..."):
                            pdf_bytes = generate_pdf(
                                df, pal,
                                header_img_path=header_path,
                                footer_img_path=footer_path,
                                municipios_filter=selected_muns,
                                show_subsidio=show_subsidio,
                            )
                        st.session_state["pdf_bytes"] = pdf_bytes
                        st.session_state["pdf_muns"] = selected_muns
                        st.success(f"✅ PDF generado · {len(selected_muns)} página(s)")

                with btn_col2:
                    if st.button("📦 Generar ZIP (1 PDF/municipio)",
                                 use_container_width=True):
                        with st.spinner("Generando archivos..."):
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                for mun in selected_muns:
                                    pdf_b = generate_pdf_single(
                                        df, mun, pal,
                                        header_img_path=header_path,
                                        footer_img_path=footer_path,
                                        show_subsidio=show_subsidio,
                                    )
                                    safe_name = mun.replace(" ", "_").replace("/", "-")
                                    zf.writestr(f"Tarifa_{safe_name}.pdf", pdf_b)
                            st.session_state["zip_bytes"] = zip_buf.getvalue()
                        st.success(f"✅ ZIP generado · {len(selected_muns)} archivos")

                # ── BOTONES DE DESCARGA ──────────────────────────────────────
                ts = datetime.now().strftime("%Y%m%d_%H%M")

                if "pdf_bytes" in st.session_state:
                    st.download_button(
                        label="⬇️ Descargar PDF completo",
                        data=st.session_state["pdf_bytes"],
                        file_name=f"Tarifas_Dispower_{ts}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

                if "zip_bytes" in st.session_state:
                    st.download_button(
                        label="⬇️ Descargar ZIP",
                        data=st.session_state["zip_bytes"],
                        file_name=f"Tarifas_Dispower_ZIP_{ts}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

    # ── VISTA PREVIA DE DATOS ─────────────────────────────────────────────────
    if excel_file is not None and "df" in dir() and not errors:
        st.markdown("---")
        st.markdown("### 👁️ Vista previa de datos")

        tab_all, tab_amgc, tab_tarifa = st.tabs(
            ["📊 Todos los datos", "🟢 Componentes AMGC", "🟠 Tarifas"])

        with tab_all:
            st.dataframe(df, use_container_width=True, hide_index=True)

        with tab_amgc:
            amgc_view = df[["MUNICIPIO", "AMGCnu_m", "AMGCvi_m",
                            "AMGCau_m", "AMGCnf_m", "AMGCro_m", "AMGC_m"]]
            st.dataframe(amgc_view.style.format({
                c: "{:,.2f}" for c in amgc_view.columns if c != "MUNICIPIO"
            }), use_container_width=True, hide_index=True)

        with tab_tarifa:
            tar_view = df[["MUNICIPIO", "WHD", "Im", "CUm",
                           "Comercializador SIN", "Tarifa SIN"]]
            st.dataframe(tar_view.style.format({
                c: "{:,.2f}" for c in tar_view.columns
                if c not in ["MUNICIPIO", "Comercializador SIN"]
            }), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
