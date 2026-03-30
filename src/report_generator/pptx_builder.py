"""
report_generator/pptx_builder.py — BDA Daily Report PowerPoint Generator.

Recreates the RESUMO DIÁRIO DOS MERCADOS template as closely as possible
in python-pptx, matching the orange BDA branding from the PDF reference.

Slide structure (11 slides):
  1.  Cover
  2.  Agenda
  3.  Sumário Executivo
  4.  Liquidez – Moeda Nacional (1/2)  — liquidity table + LUIBOR
  5.  Liquidez – Moeda Nacional (2/2)  — cash-flow + P&L
  6.  Liquidez – Moeda Estrangeira
  7.  Mercado Cambial
  8.  Mercado de Capitais – BODIVA     — segments + stocks
  9.  Mercado de Capitais – Operações BDA
  10. Informação de Mercados (1/2)     — indices + crypto
  11. Informação de Mercados (2/2)     — commodities + minerals

Usage:
    from src.report_generator.pptx_builder import BDAReportGenerator
    gen  = BDAReportGenerator(data)
    path = gen.build("output/bda_report_2026-03-30.pptx")
"""
from __future__ import annotations

import os
from datetime import date
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── BDA colour palette (derived from the PDF reference) ───────────────────────
#   Primary orange  : section headers, KPI values, titles
#   Orange dark     : table column headers, highlighted total rows
#   Orange light    : alternating table rows
#   Brown KPI       : KPI bubble fill (Sumário Executivo)
ORANGE_PRIMARY = RGBColor(0xE8, 0x75, 0x1A)   # #E8751A  — main BDA orange
ORANGE_DARK    = RGBColor(0xC0, 0x50, 0x00)   # #C05000  — table header orange
ORANGE_LIGHT   = RGBColor(0xFF, 0xF0, 0xE0)   # #FFF0E0  — alternating row tint
BROWN_KPI      = RGBColor(0x5C, 0x2D, 0x00)   # #5C2D00  — KPI bubble (dark brown)
WHITE          = RGBColor(0xFF, 0xFF, 0xFF)
BLACK          = RGBColor(0x00, 0x00, 0x00)
LIGHT_GREY     = RGBColor(0xF5, 0xF5, 0xF5)   # very light grey rows
MID_GREY       = RGBColor(0xD9, 0xD9, 0xD9)
DARK_GREY      = RGBColor(0x55, 0x55, 0x55)
GREEN_UP       = RGBColor(0x00, 0x7A, 0x33)   # positive variation
RED_DOWN       = RGBColor(0xCC, 0x00, 0x00)   # negative variation

# ── Slide dimensions: 16:9 widescreen ─────────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

FONT   = "Calibri"


# ─────────────────────────────────────────────────────────────────────────────
# Low-level shape helpers
# ─────────────────────────────────────────────────────────────────────────────

def _add_rect(slide, left, top, width, height,
              fill_color=None, line_color=None, line_width_pt: float = 0.5):
    from pptx.util import Pt as _Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa (unused, shapes constant below)
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = _Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def _add_text_box(slide, text: str, left, top, width, height,
                  font_size: int = 10, bold: bool = False,
                  italic: bool = False, color: RGBColor = BLACK,
                  align=PP_ALIGN.LEFT, word_wrap: bool = True,
                  font_name: str = FONT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name   = font_name
    return txBox


# ─────────────────────────────────────────────────────────────────────────────
# Reusable slide furniture
# ─────────────────────────────────────────────────────────────────────────────

def _slide_title(slide, title: str, date_str: str = ""):
    """
    Italic bold orange title — matching the PDF style where the section name
    appears as large italic orange text on a white background.
    A thin orange rule beneath separates it from the content.
    """
    # Thin orange top accent strip
    _add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.06), ORANGE_PRIMARY)
    # Italic orange title
    _add_text_box(slide, title,
                  Inches(0.25), Inches(0.08), Inches(9.8), Inches(0.6),
                  font_size=22, bold=True, italic=True,
                  color=ORANGE_PRIMARY, align=PP_ALIGN.LEFT)
    # Date — right-aligned, smaller, grey
    if date_str:
        _add_text_box(slide, date_str,
                      Inches(10.1), Inches(0.08), Inches(3.0), Inches(0.6),
                      font_size=11, color=DARK_GREY, align=PP_ALIGN.RIGHT)
    # Thin orange horizontal rule under title
    _add_rect(slide, Inches(0.25), Inches(0.68), SLIDE_W - Inches(0.5), Inches(0.03),
              ORANGE_PRIMARY)


def _section_bar(slide, label: str, left, top, width=None, height=Inches(0.28)):
    """Orange section header bar with white bold text — matches PDF segment headers."""
    w = width or (SLIDE_W - left - Inches(0.25))
    _add_rect(slide, left, top, w, height, ORANGE_DARK)
    _add_text_box(slide, label, left + Pt(4), top + Pt(1), w - Pt(8), height - Pt(2),
                  font_size=8, bold=True, color=WHITE, align=PP_ALIGN.LEFT)


def _footer(slide):
    """
    Orange footer strip — mirrors the 'UMA VISÃO DE FUTURO' band in the PDF.
    We can't embed the photo, so we use a solid orange strip.
    """
    footer_h = Inches(0.38)
    footer_top = SLIDE_H - footer_h
    _add_rect(slide, Inches(0), footer_top, SLIDE_W, footer_h, ORANGE_PRIMARY)
    _add_text_box(
        slide,
        "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona  |  Luanda – Angola",
        Inches(0.3), footer_top + Pt(2), SLIDE_W - Inches(3.5), footer_h - Pt(4),
        font_size=7, color=WHITE, align=PP_ALIGN.LEFT,
    )
    _add_text_box(
        slide, "UMA VISÃO DE FUTURO",
        SLIDE_W - Inches(3.2), footer_top + Pt(2), Inches(3.0), footer_h - Pt(4),
        font_size=7, bold=True, color=WHITE, align=PP_ALIGN.RIGHT,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Table helpers
# ─────────────────────────────────────────────────────────────────────────────

def _table_header_row(slide, cols: list[str], lefts: list, top,
                      height, widths: list, font_size: int = 8):
    """Orange header row — dark orange bg, white bold text."""
    for col, left, width in zip(cols, lefts, widths):
        _add_rect(slide, left, top, width, height, ORANGE_DARK, MID_GREY)
        _add_text_box(slide, str(col) if col else "",
                      left + Pt(2), top + Pt(1), width - Pt(4), height - Pt(2),
                      font_size=font_size, bold=True,
                      color=WHITE, align=PP_ALIGN.CENTER)


def _table_data_row(slide, cols: list[str], lefts: list, top, height, widths: list,
                    font_size: int = 8, bg=None,
                    highlight: bool = False):
    """
    Data row.  highlight=True renders an orange-accented row (e.g. LIQUIDEZ TOTAL).
    """
    fill = ORANGE_LIGHT if highlight else (bg or WHITE)
    txt  = ORANGE_DARK  if highlight else BLACK
    for i, (col, left, width) in enumerate(zip(cols, lefts, widths)):
        _add_rect(slide, left, top, width, height, fill, MID_GREY)
        _add_text_box(slide, str(col) if col else "",
                      left + Pt(2), top + Pt(1), width - Pt(4), height - Pt(2),
                      font_size=font_size, bold=highlight,
                      color=txt, align=PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT)


def _variation_color(val_str: str) -> RGBColor:
    """Return green/red/black based on sign prefix in string."""
    s = str(val_str).strip()
    if s.startswith("+") or (s.replace(".", "").replace(",", "").lstrip("0") and
                               not s.startswith("-") and s not in ("—", "0", "0,00%", "0.00%")):
        return GREEN_UP
    if s.startswith("-"):
        return RED_DOWN
    return DARK_GREY


def _kpi_bubble(slide, label: str, value: str, left, top,
                w=Inches(2.9), h=Inches(1.05)):
    """
    Orange-bordered KPI card — approximates the KPI display in the PDF
    Sumário Executivo (brown circles replaced by bordered rectangles for
    python-pptx compatibility).
    """
    _add_rect(slide, left, top, w, h, ORANGE_LIGHT, ORANGE_PRIMARY, 1.5)
    _add_text_box(slide, label,
                  left + Pt(4), top + Pt(3), w - Pt(8), Inches(0.3),
                  font_size=8, color=DARK_GREY, align=PP_ALIGN.LEFT)
    _add_text_box(slide, value,
                  left + Pt(4), top + Inches(0.32), w - Pt(8), Inches(0.62),
                  font_size=18, bold=True, color=ORANGE_PRIMARY, align=PP_ALIGN.LEFT)


# ─────────────────────────────────────────────────────────────────────────────
# Main Generator
# ─────────────────────────────────────────────────────────────────────────────

class BDAReportGenerator:
    """
    Builds the full BDA Daily Report PPTX.

    Pass a *data* dict with the keys listed below.  Any key not supplied will
    render dashes (—) so the slide structure is always complete even with partial
    data.

    Top-level data keys
    -------------------
    report_date         str      "30.03.2026"
    kpis                list     [{label, value, variation_str}]
    reembolso_credito   str      "17,62 M Kz"

    liquidez_mn_days    list     ["25/11","26/11","27/11","28/11","01/12"]
    liquidez_mn_rows    list     [{label, values:[5 str]}]
    transacoes_mn_rows  list     same schema
    luibor              dict     {ON,1M,3M,6M,9M,12M: str}
    luibor_variation    dict     {ON,1M,3M,6M,9M,12M: str}
    operacoes_vivas     list     [{contraparte,montante,taxa,residual,vencimento,juro_diario}]

    fluxos_mn_rows      list     [{label, values:[5 str]}]
    pl_rows             list     [{label, values:[5 str]}]
    pl_summary          list     [{label, n_ops, montante}]

    liquidez_me_rows    list     same schema as MN
    transacoes_me_rows  list
    fluxos_me_rows      list
    operacoes_vivas_me  list     [{contraparte,montante,taxa,residual,vencimento,juro_diario}]

    cambial             dict     {usd_akz, eur_akz, eur_usd, usd_akz_prev, eur_akz_prev}
    cambial_rows        list     [{par,anterior,atual,variacao}]
    mercado_rows        list     [{label,t0,t1,t2}]  (Transações do Mercado)

    bodiva_segments     dict     {key: str}
    bodiva_segment_rows list     [{segmento,anterior,atual,variacao}]
    bodiva_stocks       dict     {code: {name,volume,previous,current,change_pct}}

    carteira_titulos    list     [{carteira,cod,qty_d1,qty_d,nominal,taxa,
                                   montante,juros_anual,juro_diario}]
    bodiva_operacoes    list     [{tipo,data,cv,preco,quantidade,montante}]

    market_info         dict
        capital_markets list     [{indice,anterior,atual,variacao}]
        cm_commentary   str
        crypto          list     [{moeda,anterior,atual,variacao}]
        crypto_commentary str
        commodities     list     [{nome,anterior,atual,variacao}]
        commodities_commentary str
        minerais        list     [{nome,anterior,atual,variacao}]
        minerais_commentary str
    """

    def __init__(self, data=None):
        self.data = data or {}
        self.prs  = Presentation()
        self.prs.slide_width  = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self._blank = self.prs.slide_layouts[6]  # blank layout

    # ── Public ────────────────────────────────────────────────────────────────

    def build(self, output_path: str = "output/bda_report.pptx") -> str:
        """Generate all 11 slides and save to *output_path*. Returns the path."""
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        self._slide_cover()
        self._slide_agenda()
        self._slide_sumario_executivo()
        self._slide_liquidez_mn_1()
        self._slide_liquidez_mn_2()
        self._slide_liquidez_me()
        self._slide_mercado_cambial()
        self._slide_bodiva()
        self._slide_operacoes_bda()
        self._slide_market_info_1()
        self._slide_market_info_2()
        self.prs.save(output_path)
        return output_path

    # ── Slide 1: Cover ────────────────────────────────────────────────────────

    def _slide_cover(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", date.today().strftime("%d.%m.%Y"))

        # Upper two-thirds — orange background (stands in for the financial imagery)
        _add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(4.5), ORANGE_PRIMARY)

        # Angola map outline hint — thin white lines (simple geometric approximation)
        _add_rect(slide, Inches(0.3), Inches(0.2), Inches(0.5), Inches(2.8), None, WHITE, 1.5)

        # White panel for text
        _add_rect(slide, Inches(0), Inches(4.5), SLIDE_W, Inches(2.62), WHITE)

        # Thin orange separator line
        _add_rect(slide, Inches(0), Inches(4.5), Inches(0.12), Inches(2.2), ORANGE_PRIMARY)

        # Main title (italic)
        _add_text_box(slide, "Resumo Diário Dos Mercados",
                      Inches(0.35), Inches(4.55), Inches(11), Inches(0.85),
                      font_size=34, bold=True, italic=True,
                      color=BLACK, align=PP_ALIGN.LEFT)

        # Sub-title
        _add_text_box(slide, "DIRECÇÃO FINANCEIRA",
                      Inches(0.35), Inches(5.42), Inches(8), Inches(0.4),
                      font_size=14, bold=True, color=BLACK, align=PP_ALIGN.LEFT)

        # Date — orange
        _add_text_box(slide, date_str,
                      Inches(0.35), Inches(5.82), Inches(8), Inches(0.38),
                      font_size=14, bold=True, color=ORANGE_PRIMARY, align=PP_ALIGN.LEFT)

        # Address — small grey
        _add_text_box(
            slide,
            "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona  |  Luanda – Angola",
            Inches(0.35), Inches(6.95), Inches(10), Inches(0.28),
            font_size=7, color=DARK_GREY, align=PP_ALIGN.LEFT,
        )

        # BDA label — bottom right corner
        _add_text_box(slide, "BDA", Inches(11.5), Inches(6.6), Inches(1.6), Inches(0.6),
                      font_size=24, bold=True, color=ORANGE_PRIMARY, align=PP_ALIGN.RIGHT)
        _add_text_box(slide, "BANCO DE DESENVOLVIMENTO DE ANGOLA",
                      Inches(8.8), Inches(7.1), Inches(4.3), Inches(0.28),
                      font_size=6, color=DARK_GREY, align=PP_ALIGN.RIGHT)

        _footer(slide)

    # ── Slide 2: Agenda ───────────────────────────────────────────────────────

    def _slide_agenda(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "AGENDA", date_str)

        items = [
            ("1.", "Sumário Executivo"),
            ("2.", "Liquidez (MN)"),
            ("3.", "Liquidez (ME)"),
            ("4.", "Mercado Cambial"),
            ("5.", "Mercado Capitais"),
            ("6.", "Informação De Mercado"),
        ]
        # Two-column layout matching the PDF
        cols = [items[:3], items[3:]]
        col_lefts = [Inches(0.5), Inches(7.0)]
        for col_items, col_left in zip(cols, col_lefts):
            for i, (num, label) in enumerate(col_items):
                top = Inches(1.4) + i * Inches(1.5)
                # Number — orange bold
                _add_text_box(slide, num, col_left, top, Inches(0.5), Inches(0.5),
                              font_size=20, bold=True, color=ORANGE_PRIMARY,
                              align=PP_ALIGN.LEFT)
                # Label — black bold
                _add_text_box(slide, label,
                              col_left + Inches(0.5), top, Inches(5.5), Inches(0.55),
                              font_size=14, bold=True, color=BLACK, align=PP_ALIGN.LEFT)

        _footer(slide)

    # ── Slide 3: Sumário Executivo ────────────────────────────────────────────

    def _slide_sumario_executivo(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "Sumário Executivo", date_str)

        # Central KPI — Reembolso de Crédito (brown bubble)
        rc_val = self.data.get("reembolso_credito", "—")
        cx, cy = Inches(5.7), Inches(2.8)
        bw, bh = Inches(2.0), Inches(1.4)
        _add_rect(slide, cx, cy, bw, bh, BROWN_KPI, BROWN_KPI)
        _add_text_box(slide, rc_val, cx, cy + Pt(4), bw, Inches(0.55),
                      font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, "Reembolso de Crédito",
                      cx, cy + Inches(0.6), bw, Inches(0.5),
                      font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

        # Satellite KPI cards — arranged around the central bubble (4 left, 4 right)
        kpis = self.data.get("kpis", [
            {"label": "Liquidez MN",       "value": "—",  "variation_str": ""},
            {"label": "Liquidez ME",       "value": "—",  "variation_str": ""},
            {"label": "Posição Cambial",   "value": "—",  "variation_str": ""},
            {"label": "Carteira Títulos",  "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade MN",  "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade ME",  "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade Títulos", "value": "—", "variation_str": ""},
            {"label": "Reembolsos",        "value": "—",  "variation_str": ""},
        ])

        # Left column (4 items)
        for i, kpi in enumerate(kpis[:4]):
            top = Inches(0.82) + i * Inches(1.35)
            _kpi_bubble(slide, kpi["label"], kpi["value"], Inches(0.3), top,
                        w=Inches(2.8), h=Inches(1.0))
            if kpi.get("variation_str"):
                vc = _variation_color(kpi["variation_str"])
                _add_text_box(slide, kpi["variation_str"],
                              Inches(0.3), top + Inches(1.0), Inches(2.8), Inches(0.28),
                              font_size=8, bold=True, color=vc, align=PP_ALIGN.LEFT)

        # Right column (remaining items)
        for i, kpi in enumerate(kpis[4:8]):
            top = Inches(0.82) + i * Inches(1.35)
            _kpi_bubble(slide, kpi["label"], kpi["value"], Inches(10.25), top,
                        w=Inches(2.8), h=Inches(1.0))
            if kpi.get("variation_str"):
                vc = _variation_color(kpi["variation_str"])
                _add_text_box(slide, kpi["variation_str"],
                              Inches(10.25), top + Inches(1.0), Inches(2.8), Inches(0.28),
                              font_size=8, bold=True, color=vc, align=PP_ALIGN.RIGHT)

        _footer(slide)

    # ── Slide 4: Liquidez MN — tables + LUIBOR ────────────────────────────────

    def _slide_liquidez_mn_1(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "LIQUIDEZ – MOEDA NACIONAL", date_str)

        days       = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])
        row_h      = Inches(0.28)
        label_w    = Inches(3.4)
        col_w      = Inches(1.72)
        left0      = Inches(0.3)
        lefts      = [left0] + [left0 + label_w + i * col_w for i in range(5)]
        widths     = [label_w] + [col_w] * 5

        # ── Section: Liquidez MN table ────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Liquidez MN  (Em Milhares)", left0, top, SLIDE_W - Inches(0.5))
        top += Inches(0.28)
        _table_header_row(slide, [""] + days, lefts, top, row_h, widths)

        lmn_rows = self.data.get("liquidez_mn_rows", [
            {"label": "Posição Reservas Livres BNA", "values": ["—"] * 5},
            {"label": "Posição DO B. Comerciais",    "values": ["—"] * 5},
            {"label": "Posição DP B. Comerciais",    "values": ["—"] * 5},
            {"label": "Posição OMAs",                "values": ["—"] * 5},
            {"label": "LIQUIDEZ BDA",                "values": ["—"] * 5},
        ])
        for i, row in enumerate(lmn_rows):
            top += row_h
            is_total = "LIQUIDEZ" in row["label"].upper()
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide, [row["label"]] + row["values"],
                            lefts, top, row_h, widths, highlight=is_total, bg=bg)

        # ── Section: Transações (OMA) ─────────────────────────────────────────
        top += row_h + Inches(0.12)
        _section_bar(slide, "Transações", left0, top, SLIDE_W - Inches(0.5))
        tx_heads = ["Tipo", "Contraparte", "Taxa", "Montante", "Maturidade", "Juros"]
        tx_widths = [Inches(1.5), Inches(2.0), Inches(1.2), Inches(2.5), Inches(2.0), Inches(3.83)]
        tx_lefts  = [left0]
        for w in tx_widths[:-1]:
            tx_lefts.append(tx_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, tx_heads, tx_lefts, top, row_h, tx_widths)

        tx_rows = self.data.get("transacoes_mn_rows", [
            {"label": "OMA  Cedencia  BNA  10%  8 768 000 000  1  2 402 192",
             "values": [""] * 5},
        ])
        # Accept both simple list-of-dicts and flat rows
        raw_tx = self.data.get("transacoes_mn_raw", [])
        for i, row in enumerate(raw_tx):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            vals = [
                row.get("tipo", "OMA"),
                row.get("contraparte", "—"),
                row.get("taxa", "—"),
                row.get("montante", "—"),
                row.get("maturidade", "—"),
                row.get("juros", "—"),
            ]
            _table_data_row(slide, vals, tx_lefts, top, row_h, tx_widths, bg=bg)
        if not raw_tx:
            top += row_h
            _table_data_row(slide, ["—"] * 6, tx_lefts, top, row_h, tx_widths)

        # ── Section: Operações Vivas ──────────────────────────────────────────
        top += row_h + Inches(0.12)
        _section_bar(slide, "Operações Vivas", left0, top, SLIDE_W - Inches(0.5))
        op_heads = ["Tipo", "Contraparte", "Montante", "Taxa", "Residual (Dias)", "Vencimento", "Juro Diário"]
        n_op = len(op_heads)
        op_w = (SLIDE_W - Inches(0.5)) / n_op
        op_lefts  = [left0 + i * op_w for i in range(n_op)]
        op_widths = [op_w] * n_op
        top += Inches(0.28)
        _table_header_row(slide, op_heads, op_lefts, top, row_h, op_widths, font_size=7)

        ops = self.data.get("operacoes_vivas", [])
        for i, op in enumerate(ops):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            vals = [
                op.get("tipo", "DP"),
                op.get("contraparte", "—"),
                op.get("montante", "—"),
                op.get("taxa", "—"),
                str(op.get("residual", "—")),
                op.get("vencimento", "—"),
                op.get("juro_diario", "—"),
            ]
            _table_data_row(slide, vals, op_lefts, top, row_h, op_widths, font_size=7, bg=bg)
        if not ops:
            top += row_h
            _table_data_row(slide, ["—"] * n_op, op_lefts, top, row_h, op_widths, font_size=7)

        # ── Section: LUIBOR ───────────────────────────────────────────────────
        top += row_h + Inches(0.12)
        _section_bar(slide, "Taxas LUIBOR", left0, top, SLIDE_W - Inches(0.5))
        tenors  = ["LUIBOR O/N", "LUIBOR 1M", "LUIBOR 3M", "LUIBOR 6M", "LUIBOR 9M", "LUIBOR 12M"]
        lu_heads = ["Maturidade", "Anterior (D-2)", "Anterior (D-1)", "Actual (D)", "Var (%)"]
        lu_n    = len(lu_heads)
        lu_w    = (SLIDE_W - Inches(0.5)) / lu_n
        lu_lefts  = [left0 + i * lu_w for i in range(lu_n)]
        lu_widths = [lu_w] * lu_n
        top += Inches(0.28)
        _table_header_row(slide, lu_heads, lu_lefts, top, row_h, lu_widths)

        luibor = self.data.get("luibor", {})
        luibor_var = self.data.get("luibor_variation", {})
        for i, t in enumerate(tenors):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            rate = luibor.get(t, "—")
            var  = luibor_var.get(t, "—")
            # Prev D-2 and D-1: use same value if not provided separately
            _table_data_row(
                slide,
                [t, rate, rate, rate, var],
                lu_lefts, top, row_h, lu_widths, bg=bg,
            )

        _footer(slide)

    # ── Slide 5: Liquidez MN — Cash-flow + P&L ────────────────────────────────

    def _slide_liquidez_mn_2(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "LIQUIDEZ – MOEDA NACIONAL", date_str)

        days   = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])
        row_h  = Inches(0.28)
        lbl_w  = Inches(3.8)
        col_w  = Inches(1.72)
        left0  = Inches(0.3)
        lefts  = [left0] + [left0 + lbl_w + i * col_w for i in range(5)]
        widths = [lbl_w] + [col_w] * 5

        def _render_section(label, key, default_labels, start_top):
            _section_bar(slide, label, left0, start_top, SLIDE_W - Inches(0.5))
            t = start_top + Inches(0.28)
            _table_header_row(slide, [""] + days, lefts, t, row_h, widths)
            rows = self.data.get(key, [
                {"label": lbl, "values": ["—"] * 5} for lbl in default_labels
            ])
            for i, row in enumerate(rows):
                t += row_h
                is_total = any(x in row["label"].upper() for x in ("TOTAL", "GAP", "LÍQUIDO", "RESULTADO"))
                bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
                _table_data_row(slide, [row["label"]] + row["values"],
                                lefts, t, row_h, widths,
                                highlight=is_total, bg=bg)
            return t + row_h

        top = Inches(0.78)

        # Cash-flow section
        cf_defaults = [
            "Fluxos de Entradas (Cash in flow)",
            "Recebimentos de cupão de títulos",
            "Reembolsos de crédito (+)",
            "Reembolsos de OMA-O/N + Juros",
            "Transferencia a favor conta BNA",
            "Fluxos de Saídas (Cash out flow)",
            "(-) Juros, comissões e outros",
            "Custos Com Pessoal",
            "Fornecimentos e Serviços",
            "Desembolso de crédito (-)",
            "Impostos",
            "Aplicação em OMA",
            "GAP de Liquidez",
        ]
        top = _render_section("Fluxos de Caixa MN", "fluxos_mn_rows", cf_defaults, top)

        # P&L Control — compact table (Nº Ops + Montante)
        top += Inches(0.12)
        _section_bar(slide, "P&L Control", left0, top, SLIDE_W - Inches(0.5))
        pl_heads = ["Categoria", "Nº Operações", "Montante"]
        pl_w = [(SLIDE_W - Inches(0.5)) * r for r in (0.5, 0.25, 0.25)]
        pl_lefts = [left0, left0 + pl_w[0], left0 + pl_w[0] + pl_w[1]]
        top += Inches(0.28)
        _table_header_row(slide, pl_heads, pl_lefts, top, row_h, pl_w)

        pl_summary = self.data.get("pl_summary", [
            {"label": "Reembolso de Crédito", "n_ops": "—", "montante": "—"},
            {"label": "Fornecedores",          "n_ops": "—", "montante": "—"},
            {"label": "Desembolso de Crédito", "n_ops": "—", "montante": "—"},
        ])
        for i, row in enumerate(pl_summary):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["label"], str(row["n_ops"]), str(row["montante"])],
                            pl_lefts, top, row_h, pl_w, bg=bg)

        _footer(slide)

    # ── Slide 6: Liquidez ME ──────────────────────────────────────────────────

    def _slide_liquidez_me(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "LIQUIDEZ – MOEDA ESTRANGEIRA", date_str)

        days   = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])
        row_h  = Inches(0.28)
        lbl_w  = Inches(3.4)
        col_w  = Inches(1.72)
        left0  = Inches(0.3)
        lefts  = [left0] + [left0 + lbl_w + i * col_w for i in range(5)]
        widths = [lbl_w] + [col_w] * 5
        top    = Inches(0.78)

        # Liquidez ME table
        _section_bar(slide, "Liquidez ME  (Em Milhões USD)", left0, top, SLIDE_W - Inches(0.5))
        top += Inches(0.28)
        _table_header_row(slide, [""] + days, lefts, top, row_h, widths)

        lme_rows = self.data.get("liquidez_me_rows", [
            {"label": "SALDO D.O Estrangeiros", "values": ["—"] * 5},
            {"label": "DPs ME",                 "values": ["—"] * 5},
            {"label": "COLATERAL CDI",           "values": ["—"] * 5},
            {"label": "LIQUIDEZ BDA",            "values": ["—"] * 5},
        ])
        for i, row in enumerate(lme_rows):
            top += row_h
            is_total = "LIQUIDEZ" in row["label"].upper()
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide, [row["label"]] + row["values"],
                            lefts, top, row_h, widths,
                            highlight=is_total, bg=bg)

        # Transações ME
        top += row_h + Inches(0.12)
        _section_bar(slide, "Transações ME", left0, top, SLIDE_W - Inches(0.5))
        tx_me_heads = ["Tipo", "Moeda", "Contraparte", "Taxa", "Montante", "Maturidade", "Juros"]
        n = len(tx_me_heads)
        tx_me_w = (SLIDE_W - Inches(0.5)) / n
        tx_me_lefts = [left0 + i * tx_me_w for i in range(n)]
        tx_me_widths = [tx_me_w] * n
        top += Inches(0.28)
        _table_header_row(slide, tx_me_heads, tx_me_lefts, top, row_h, tx_me_widths, font_size=7)

        tx_me = self.data.get("transacoes_me_raw", [])
        for i, row in enumerate(tx_me):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row.get(k, "—") for k in
                             ("tipo", "moeda", "contraparte", "taxa", "montante", "maturidade", "juros")],
                            tx_me_lefts, top, row_h, tx_me_widths, font_size=7, bg=bg)
        if not tx_me:
            top += row_h
            _table_data_row(slide, ["—"] * n, tx_me_lefts, top, row_h, tx_me_widths, font_size=7)

        # Operações Vivas ME
        top += row_h + Inches(0.12)
        _section_bar(slide, "Operações Vivas ME", left0, top, SLIDE_W - Inches(0.5))
        op_me_heads = ["Contraparte", "Montante", "Taxa", "Residual (Dias)", "Vencimento", "Juro Diário"]
        n2 = len(op_me_heads)
        op_me_w = (SLIDE_W - Inches(0.5)) / n2
        op_me_lefts = [left0 + i * op_me_w for i in range(n2)]
        op_me_widths = [op_me_w] * n2
        top += Inches(0.28)
        _table_header_row(slide, op_me_heads, op_me_lefts, top, row_h, op_me_widths, font_size=7)

        ops_me = self.data.get("operacoes_vivas_me", [])
        for i, op in enumerate(ops_me):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [op.get(k, "—") for k in
                             ("contraparte", "montante", "taxa", "residual", "vencimento", "juro_diario")],
                            op_me_lefts, top, row_h, op_me_widths, font_size=7, bg=bg)
        if not ops_me:
            top += row_h
            _table_data_row(slide, ["—"] * n2, op_me_lefts, top, row_h, op_me_widths, font_size=7)

        # Fluxos ME
        top += row_h + Inches(0.12)
        _section_bar(slide, "Fluxos de Caixa ME", left0, top, SLIDE_W - Inches(0.5))
        top += Inches(0.28)
        _table_header_row(slide, [""] + days, lefts, top, row_h, widths)
        fluxos_me = self.data.get("fluxos_me_rows", [
            {"label": "Fluxos de entradas (Cash in flow)", "values": ["—"] * 5},
            {"label": "Outros recebimentos",               "values": ["—"] * 5},
            {"label": "Reembolsos de DP + Juros",          "values": ["—"] * 5},
            {"label": "Fluxos de Saídas (Cash out flow)",  "values": ["—"] * 5},
            {"label": "Aplicação em DP ME",                "values": ["—"] * 5},
            {"label": "GAP de Liquidez",                   "values": ["—"] * 5},
        ])
        for i, row in enumerate(fluxos_me):
            top += row_h
            is_total = "GAP" in row["label"].upper() or "TOTAL" in row["label"].upper()
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide, [row["label"]] + row["values"],
                            lefts, top, row_h, widths,
                            highlight=is_total, bg=bg)

        _footer(slide)

    # ── Slide 7: Mercado Cambial ──────────────────────────────────────────────

    def _slide_mercado_cambial(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "MERCADO CAMBIAL", date_str)

        cambial  = self.data.get("cambial", {})
        row_h    = Inches(0.28)
        left0    = Inches(0.3)
        table_w  = Inches(6.2)

        # ── Cambiais rates table (3 date columns + %) ─────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Cambiais", left0, top, table_w)
        c_heads  = ["Par", "Anterior (D-1)", "Anterior", "Actual (D)", "(%)"]
        c_w      = [Inches(1.4), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.2)]
        c_lefts  = [left0]
        for w in c_w[:-1]:
            c_lefts.append(c_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, c_heads, c_lefts, top, row_h, c_w)

        cambial_rows = self.data.get("cambial_rows", [
            {"par": "USD/AKZ", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
            {"par": "EUR/AKZ", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
            {"par": "EUR/USD", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(cambial_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(
                slide,
                [row["par"], row.get("anterior2", "—"), row.get("anterior", "—"),
                 row.get("atual", "—"), row.get("variacao", "—")],
                c_lefts, top, row_h, c_w, bg=bg,
            )

        # ── Transações BDA table ──────────────────────────────────────────────
        top += row_h + Inches(0.15)
        _section_bar(slide, "Transações BDA", left0, top, table_w)
        tb_heads  = ["C/V", "Par de moeda", "Montante Debt", "Câmbio", "P/L AKZ"]
        tb_w      = [Inches(0.7), Inches(1.5), Inches(1.5), Inches(1.3), Inches(1.2)]
        tb_lefts  = [left0]
        for w in tb_w[:-1]:
            tb_lefts.append(tb_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, tb_heads, tb_lefts, top, row_h, tb_w, font_size=7)

        bda_tx = self.data.get("transacoes_bda_rows", [])
        for i, row in enumerate(bda_tx):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(
                slide,
                [row.get(k, "—") for k in ("cv", "par", "montante", "cambio", "pl")],
                tb_lefts, top, row_h, tb_w, font_size=7, bg=bg,
            )
        if not bda_tx:
            top += row_h
            _table_data_row(slide, ["—"] * 5, tb_lefts, top, row_h, tb_w, font_size=7)

        # ── Transações do Mercado (T+0, T+1, T+2) ────────────────────────────
        top += row_h + Inches(0.15)
        _section_bar(slide, "Transações do Mercado", left0, top, table_w)
        tm_heads = ["Liquidação", "Montante USD", "Mínimo", "Máximo"]
        tm_w = [Inches(1.2), Inches(1.8), Inches(1.6), Inches(1.6)]
        tm_lefts = [left0]
        for w in tm_w[:-1]:
            tm_lefts.append(tm_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, tm_heads, tm_lefts, top, row_h, tm_w)

        mercado_rows = self.data.get("mercado_rows", [
            {"label": "T+0", "montante": "—", "min": "—", "max": "—"},
            {"label": "T+1", "montante": "—", "min": "—", "max": "—"},
            {"label": "T+2", "montante": "—", "min": "—", "max": "—"},
        ])
        for i, row in enumerate(mercado_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(
                slide,
                [row.get("label", "—"), row.get("montante", "—"),
                 row.get("min", "—"), row.get("max", "—")],
                tm_lefts, top, row_h, tm_w, bg=bg,
            )

        # ── Right panel — KPI summary bubbles ────────────────────────────────
        right_x = Inches(7.0)
        _kpi_bubble(slide, "Transações (USD)",
                    cambial.get("vol_total_usd", "—"), right_x, Inches(0.82),
                    w=Inches(2.9), h=Inches(1.0))
        _kpi_bubble(slide, "Posição Cambial (Kz)",
                    cambial.get("posicao_cambial", "—"), right_x + Inches(3.1), Inches(0.82),
                    w=Inches(2.9), h=Inches(1.0))

        _footer(slide)

    # ── Slide 8: Mercado de Capitais – BODIVA ─────────────────────────────────

    def _slide_bodiva(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "MERCADO DE CAPITAIS", date_str)

        row_h  = Inches(0.28)
        left0  = Inches(0.3)

        # ── Segmentado por Produtos ───────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Segmentado Por Produtos", left0, top, SLIDE_W - Inches(0.5))

        sp_heads  = ["Segmento", "Anterior", "Actual", "(%)"]
        sp_w      = [Inches(4.5), Inches(2.5), Inches(3.0), Inches(2.83)]
        sp_lefts  = [left0]
        for w in sp_w[:-1]:
            sp_lefts.append(sp_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, sp_heads, sp_lefts, top, row_h, sp_w)

        seg_rows = self.data.get("bodiva_segment_rows", [
            {"segmento": "Obrigações De Tesouro",    "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Bilhetes Do Tesouro",       "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Obrigações Privadas",       "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Unidades De Participações", "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Acções",                    "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Repos",                     "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Total",                     "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(seg_rows):
            top += row_h
            is_total = "TOTAL" in row["segmento"].upper()
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["segmento"], row["anterior"], row["atual"], row["variacao"]],
                            sp_lefts, top, row_h, sp_w,
                            highlight=is_total, bg=bg)

        # KPI summary bubble
        total_atual = self.data.get("bodiva_total_transacoes", "—")
        _kpi_bubble(slide, "Kz  Transações", total_atual,
                    SLIDE_W - Inches(3.5), Inches(0.82), w=Inches(3.0), h=Inches(1.0))

        # ── Mercado de Bolsas de Acções ───────────────────────────────────────
        top += row_h + Inches(0.18)
        _section_bar(slide, "Mercado de Bolsas de Acções", left0, top, SLIDE_W - Inches(0.5))

        stk_heads  = ["Código", "Vol. Transacc.", "Preço Anterior", "Preço Actual", "Variação", "Cap. Bolsista"]
        stk_w      = [Inches(1.8), Inches(1.8), Inches(2.0), Inches(2.0), Inches(1.8), Inches(3.73)]
        stk_lefts  = [left0]
        for w in stk_heads[:-1]:
            stk_lefts.append(stk_lefts[-1] + stk_w[len(stk_lefts) - 1])
        top += Inches(0.28)
        _table_header_row(slide, stk_heads, stk_lefts, top, row_h, stk_w)

        stocks = self.data.get("bodiva_stocks", {})
        for i, (code, info) in enumerate(stocks.items()):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            chg = info.get("change_pct")
            chg_str = f"{chg:+.2f}%" if isinstance(chg, (int, float)) else str(chg or "—")
            _table_data_row(
                slide,
                [code,
                 str(info.get("volume", "—")),
                 str(info.get("previous", "—")),
                 str(info.get("current", "—")),
                 chg_str,
                 str(info.get("cap_bolsista", "—"))],
                stk_lefts, top, row_h, stk_w, bg=bg,
            )
        if not stocks:
            top += row_h
            _table_data_row(slide, ["Dados não disponíveis (BODIVA)"] + ["—"] * 5,
                            stk_lefts, top, row_h, stk_w)

        _footer(slide)

    # ── Slide 9: Operações BDA ────────────────────────────────────────────────

    def _slide_operacoes_bda(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "MERCADO DE CAPITAIS – OPERAÇÕES BDA", date_str)

        row_h = Inches(0.26)
        left0 = Inches(0.3)

        # ── Transações section ────────────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Transações", left0, top, SLIDE_W - Inches(0.5))
        tx_heads  = ["Tipo de Operação", "Data Contrat.", "C/V", "Preço", "Quantidades", "Montante"]
        n_tx = len(tx_heads)
        tx_w = (SLIDE_W - Inches(0.5)) / n_tx
        tx_lefts  = [left0 + i * tx_w for i in range(n_tx)]
        tx_widths = [tx_w] * n_tx
        top += Inches(0.28)
        _table_header_row(slide, tx_heads, tx_lefts, top, row_h, tx_widths, font_size=7)

        ops = self.data.get("bodiva_operacoes", [])
        for i, row in enumerate(ops):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(
                slide,
                [row.get(k, "—") for k in ("tipo", "data", "cv", "preco", "quantidade", "montante")],
                tx_lefts, top, row_h, tx_widths, font_size=7, bg=bg,
            )
        if not ops:
            top += row_h
            _table_data_row(slide, ["—"] * n_tx, tx_lefts, top, row_h, tx_widths, font_size=7)

        # KPI bubbles — Transações + Juros Diário
        tx_kpi_val  = self.data.get("bodiva_transacoes_valor", "0,00 mM Kz")
        jd_kpi_val  = self.data.get("bodiva_juros_diario",    "—")
        _kpi_bubble(slide, "Kz  Transações",  tx_kpi_val,  Inches(3.0),  Inches(1.8),
                    w=Inches(2.8), h=Inches(0.95))
        _kpi_bubble(slide, "Kz  Juros Diário", jd_kpi_val, Inches(6.2),  Inches(1.8),
                    w=Inches(2.8), h=Inches(0.95))

        # ── Carteira de Títulos ───────────────────────────────────────────────
        top = Inches(3.1)
        _section_bar(slide, "Carteira De Títulos", left0, top, SLIDE_W - Inches(0.5))

        ct_heads  = ["Carteira", "Cód. Neg.", "Qtd D-1", "Qtd D", "Val. Nominal",
                     "Taxa", "Montante D", "Juros Anual", "Juros Diário D"]
        n_ct = len(ct_heads)
        ct_w = (SLIDE_W - Inches(0.5)) / n_ct
        ct_lefts  = [left0 + i * ct_w for i in range(n_ct)]
        ct_widths = [ct_w] * n_ct
        top += Inches(0.28)
        _table_header_row(slide, ct_heads, ct_lefts, top, row_h, ct_widths, font_size=6)

        carteira = self.data.get("carteira_titulos", [])
        for i, row in enumerate(carteira):
            top += row_h
            is_total = row.get("cod", "").upper() == "TOTAL" or row.get("carteira", "").upper() == "TOTAL"
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(
                slide,
                [row.get(k, "—") for k in
                 ("carteira", "cod", "qty_d1", "qty_d", "nominal",
                  "taxa", "montante", "juros_anual", "juro_diario")],
                ct_lefts, top, row_h, ct_widths, font_size=6,
                highlight=is_total, bg=bg,
            )
        if not carteira:
            top += row_h
            _table_data_row(slide, ["—"] * n_ct, ct_lefts, top, row_h, ct_widths, font_size=6)

        _footer(slide)

    # ── Slide 10: Informação de Mercados (1/2) — Indices + Crypto ─────────────

    def _slide_market_info_1(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "INFORMAÇÃO DE MERCADOS", date_str)

        market = self.data.get("market_info", {})
        row_h  = Inches(0.28)
        left0  = Inches(0.3)
        # Split: left table ~47 %, right commentary ~50 %
        tbl_w  = Inches(6.0)
        com_x  = Inches(6.6)
        com_w  = SLIDE_W - com_x - Inches(0.25)

        # ── Capital Markets table ─────────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Capital Markets", left0, top, tbl_w)
        cm_heads = ["Índice", "Anterior", "Actual", "(%)"]
        cm_w     = [Inches(2.4), Inches(1.2), Inches(1.2), Inches(1.2)]
        cm_lefts = [left0]
        for w in cm_w[:-1]:
            cm_lefts.append(cm_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, cm_heads, cm_lefts, top, row_h, cm_w)

        cm_rows = market.get("capital_markets", [
            {"indice": "S&P500",             "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "Dow Jones",           "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "NASDAQ",              "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "NIKKEI 225",          "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "IBOVESPA",            "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "SHANGHAI COMPOSITE",  "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "EUROSTOX",            "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "Bolsa de Londres",    "anterior": "—", "atual": "—", "variacao": "—"},
            {"indice": "PSI 20",              "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(cm_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["indice"], row["anterior"], row["atual"], row["variacao"]],
                            cm_lefts, top, row_h, cm_w, bg=bg)

        # Commentary — right panel
        cm_comment = market.get("cm_commentary", "")
        if cm_comment:
            _add_rect(slide, com_x, Inches(0.78), com_w, Inches(3.0), ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cm_comment,
                          com_x + Pt(6), Inches(0.82), com_w - Pt(12), Inches(2.9),
                          font_size=8, color=BLACK, word_wrap=True)

        # ── Criptomoedas ──────────────────────────────────────────────────────
        top = Inches(4.4)
        _section_bar(slide, "Criptomoedas", left0, top, tbl_w)
        cr_heads = ["Moeda", "Anterior", "Actual", "(%)"]
        cr_w     = [Inches(2.4), Inches(1.2), Inches(1.2), Inches(1.2)]
        cr_lefts = [left0]
        for w in cr_w[:-1]:
            cr_lefts.append(cr_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, cr_heads, cr_lefts, top, row_h, cr_w)

        cr_rows = market.get("crypto", [
            {"moeda": "BITCOIN (BTC)",  "anterior": "—", "atual": "—", "variacao": "—"},
            {"moeda": "ETHEREUM (ETH)", "anterior": "—", "atual": "—", "variacao": "—"},
            {"moeda": "XRP (XRP)",      "anterior": "—", "atual": "—", "variacao": "—"},
            {"moeda": "USDC",           "anterior": "—", "atual": "—", "variacao": "—"},
            {"moeda": "TETHER",         "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(cr_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["moeda"], row["anterior"], row["atual"], row["variacao"]],
                            cr_lefts, top, row_h, cr_w, bg=bg)

        # Crypto commentary — right panel
        cr_comment = market.get("crypto_commentary", "")
        if cr_comment:
            _add_rect(slide, com_x, Inches(4.4), com_w, Inches(2.6), ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cr_comment,
                          com_x + Pt(6), Inches(4.44), com_w - Pt(12), Inches(2.5),
                          font_size=8, color=BLACK, word_wrap=True)

        # "Nota" tag — matches PDF
        _add_rect(slide, com_x, top + row_h + Inches(0.05), Inches(1.0), Inches(0.25),
                  ORANGE_PRIMARY)
        _add_text_box(slide, "Nota", com_x + Pt(2), top + row_h + Inches(0.06),
                      Inches(0.96), Inches(0.22),
                      font_size=8, bold=True, italic=True,
                      color=WHITE, align=PP_ALIGN.CENTER)

        _footer(slide)

    # ── Slide 11: Informação de Mercados (2/2) — Commodities + Minerals ───────

    def _slide_market_info_2(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "INFORMAÇÃO DE MERCADOS", date_str)

        market = self.data.get("market_info", {})
        row_h  = Inches(0.28)
        left0  = Inches(0.3)
        tbl_w  = Inches(6.0)
        com_x  = Inches(6.6)
        com_w  = SLIDE_W - com_x - Inches(0.25)

        # ── Commodities ───────────────────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Commodities", left0, top, tbl_w)
        cmd_heads = ["Commodity", "Anterior", "Actual", "(%)"]
        cmd_w     = [Inches(2.4), Inches(1.2), Inches(1.2), Inches(1.2)]
        cmd_lefts = [left0]
        for w in cmd_heads[:-1]:
            cmd_lefts.append(cmd_lefts[-1] + cmd_w[len(cmd_lefts) - 1])
        top += Inches(0.28)
        _table_header_row(slide, cmd_heads, cmd_lefts, top, row_h, cmd_w)

        cmd_rows = market.get("commodities", [
            {"nome": "PETRÓLEO (BRENT)",      "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "MILHO (USD/BU)",         "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "SOJA (USD/BU)",          "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "TRIGO (USD/LBS)",        "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "CAFÉ (USD/LBS)",         "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "AÇÚCAR (USD/LBS)",       "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "ÓLEO DE PALMA (USD/LBS)","anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "ALGODÃO (USD/LBS)",      "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "BANANA (USD/LBS)",       "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(cmd_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["nome"], row["anterior"], row["atual"], row["variacao"]],
                            cmd_lefts, top, row_h, cmd_w, bg=bg)

        # Commodities commentary
        cmd_comment = market.get("commodities_commentary", "")
        if cmd_comment:
            _add_rect(slide, com_x, Inches(0.78), com_w, Inches(3.0), ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cmd_comment,
                          com_x + Pt(6), Inches(0.82), com_w - Pt(12), Inches(2.9),
                          font_size=8, color=BLACK, word_wrap=True)

        # ── Minerais ──────────────────────────────────────────────────────────
        top += row_h + Inches(0.18)
        _section_bar(slide, "Minerais", left0, top, tbl_w)
        min_heads = ["Mineral", "Anterior", "Actual", "(%)"]
        top += Inches(0.28)
        _table_header_row(slide, min_heads, cmd_lefts, top, row_h, cmd_w)

        min_rows = market.get("minerais", [
            {"nome": "OURO",     "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "FERRO",    "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "COBRE",    "anterior": "—", "atual": "—", "variacao": "—"},
            {"nome": "MANGANÊS", "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        for i, row in enumerate(min_rows):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row["nome"], row["anterior"], row["atual"], row["variacao"]],
                            cmd_lefts, top, row_h, cmd_w, bg=bg)

        # Minerals commentary
        min_comment = market.get("minerais_commentary", "")
        if min_comment:
            _add_rect(slide, com_x, top - Inches(4 * 0.28), com_w, Inches(2.6),
                      ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, min_comment,
                          com_x + Pt(6), top - Inches(4 * 0.28) + Pt(4),
                          com_w - Pt(12), Inches(2.5),
                          font_size=8, color=BLACK, word_wrap=True)

        # "Nota" tag
        _add_rect(slide, com_x, top + row_h + Inches(0.05), Inches(1.0), Inches(0.25),
                  ORANGE_PRIMARY)
        _add_text_box(slide, "Nota", com_x + Pt(2), top + row_h + Inches(0.06),
                      Inches(0.96), Inches(0.22),
                      font_size=8, bold=True, italic=True,
                      color=WHITE, align=PP_ALIGN.CENTER)

        _footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# Backward-compatible alias + quick smoke-test
# ─────────────────────────────────────────────────────────────────────────────

PPTXBuilder = BDAReportGenerator   # legacy alias

if __name__ == "__main__":
    gen  = BDAReportGenerator()
    path = gen.build("output/bda_report_test.pptx")
    print(f"Test report saved: {path}")
