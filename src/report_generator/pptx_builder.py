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
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Brand asset paths (extracted from original PPTX) ──────────────────────────
_ASSETS_DIR = Path(__file__).parent.parent.parent / "assets" / "images"
IMG_FOOTER_BANNER = str(_ASSETS_DIR / "footer_banner.png")     # "UMA VISÃO DE FUTURO" strip — every slide
IMG_COVER_HEADER  = str(_ASSETS_DIR / "cover_header.png")      # orange hero — cover
IMG_INNER_HEADER  = str(_ASSETS_DIR / "inner_header.png")      # orange hero — inner slides
IMG_AGENDA_BG     = str(_ASSETS_DIR / "agenda_bg.jpg")         # full background — agenda

# Per-slide icons (all extracted from original PPTX at exact positions)
IMG_ICON_S3_LMN       = str(_ASSETS_DIR / "icon_s3_liquidez_mn.png")       # s3 coin+arrow  → Liquidez MN KPI
IMG_ICON_S3_LME       = str(_ASSETS_DIR / "icon_s3_liquidez_me.png")       # s3 people      → Liquidez ME KPI
IMG_ICON_S3_RENTA     = str(_ASSETS_DIR / "icon_s3_rentabilidade.png")     # s3 bar-chart   → Rentabilidade cards
IMG_ICON_S3_VBAR      = str(_ASSETS_DIR / "icon_s3_vbar.png")              # s3 vertical bar divider
IMG_ICON_S3_CAMBIAL   = str(_ASSETS_DIR / "icon_s3_posicao_cambial.png")   # s3 people      → Posição Cambial
IMG_ICON_S3_REEMB     = str(_ASSETS_DIR / "icon_s3_reembolsos.png")        # s3 coins stack → Reembolsos
IMG_ICON_S3_REPORT    = str(_ASSETS_DIR / "icon_report_color.png")         # s3 coloured report icon bottom-right
IMG_ICON_CALCULATOR   = str(_ASSETS_DIR / "icon_calculator.png")           # s4/s6 calculator icon
IMG_ICON_GEAR         = str(_ASSETS_DIR / "icon_gear.png")                 # s4 hand+gear icon
IMG_ICON_KZ_CIRCLE    = str(_ASSETS_DIR / "icon_kz_circle.png")            # s5 "Kz" circle
IMG_ICON_REPORT_MONEY = str(_ASSETS_DIR / "icon_report_money.png")         # s7/s10/s11 report+money icon
IMG_ICON_FX_EXCHANGE  = str(_ASSETS_DIR / "icon_fx_exchange.png")          # s7 EUR/USD exchange icon
IMG_ICON_GLOBE        = str(_ASSETS_DIR / "icon_globe.png")                # s11 world globe
IMG_ICON_REPORT_BW    = str(_ASSETS_DIR / "icon_report_bw.png")            # s2 agenda BW report icon

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

def _place_img(slide, path, left, top, width=None, height=None):
    """Place an image if the file exists; silently skip if missing."""
    if os.path.isfile(path):
        slide.shapes.add_picture(path, left, top, width, height)


def _add_rect(slide, left, top, width, height,
              fill_color=None, line_color=None, line_width_pt: float = 0.5):
    from pptx.util import Pt as _Pt
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
    Inner slide header — matches reference exactly:
      • Small orange left accent block  (object 131: L=0 T=0.401 W=0.277 H=0.397)
      • Bold title text                 (object 132: L=0.396 T=0.224 W=11.321 H=0.351 20pt)
      • Thin orange horizontal rule     (Straight Connector at T=0.607)
    NO full-width image — the reference inner slides have none.
    """
    # Small orange left accent marker
    _add_rect(slide, Inches(0), Inches(0.401), Inches(0.277), Inches(0.397), ORANGE_PRIMARY)

    # Bold section title
    _add_text_box(slide, title,
                  Inches(0.396), Inches(0.224), Inches(11.321), Inches(0.351),
                  font_size=20, bold=True, color=BLACK, align=PP_ALIGN.LEFT)

    # Date — right-aligned, smaller, grey
    if date_str:
        _add_text_box(slide, date_str,
                      Inches(11.5), Inches(0.224), Inches(1.6), Inches(0.351),
                      font_size=10, color=DARK_GREY, align=PP_ALIGN.RIGHT)

    # Thin orange horizontal rule under title (T=0.607 matches reference connector)
    _add_rect(slide, Inches(0.396), Inches(0.607), SLIDE_W - Inches(0.6), Inches(0.015),
              ORANGE_PRIMARY)


def _section_bar(slide, label: str, left, top, width=None, height=Inches(0.28)):
    """Orange section header bar with white bold text — matches PDF segment headers."""
    w = width or (SLIDE_W - left - Inches(0.25))
    _add_rect(slide, left, top, w, height, ORANGE_DARK)
    _add_text_box(slide, label, left + Pt(4), top + Pt(1), w - Pt(8), height - Pt(2),
                  font_size=8, bold=True, color=WHITE, align=PP_ALIGN.LEFT)


def _footer(slide):
    """
    Footer strip — uses the real 'UMA VISÃO DE FUTURO' brand image extracted from
    the original PPTX.  Falls back to a solid orange strip if the image is missing.
    """
    footer_h   = Inches(0.72)
    footer_top = SLIDE_H - footer_h
    if os.path.isfile(IMG_FOOTER_BANNER):
        slide.shapes.add_picture(IMG_FOOTER_BANNER,
                                 Inches(0), footer_top, SLIDE_W, footer_h)
    else:
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


def _inner_header(slide):
    """Orange header banner for inner slides (extracted from original PPTX)."""
    header_h = Inches(0.52)
    if os.path.isfile(IMG_INNER_HEADER):
        slide.shapes.add_picture(IMG_INNER_HEADER, Inches(0), Inches(0), SLIDE_W, header_h)
    else:
        _add_rect(slide, Inches(0), Inches(0), SLIDE_W, header_h, ORANGE_PRIMARY)


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


def _summary_oval(slide, label: str, value: str, left, top, width, height,
                  fill=None, text_color=WHITE):
    """
    Brown-filled oval KPI — matches original PPTX summary ovals on slides 4, 6, 7, 8, 9.
    Exact sizes come from the coordinate dump of the source PPTX.
    """
    if fill is None:
        fill = BROWN_KPI
    oval = slide.shapes.add_shape(9, left, top, width, height)
    oval.fill.solid()
    oval.fill.fore_color.rgb = fill
    oval.line.fill.background()
    lbl_h = height * 0.42
    val_h = height - lbl_h
    _add_text_box(slide, label,
                  left + Inches(0.04), top + Inches(0.04),
                  width - Inches(0.08), lbl_h - Inches(0.04),
                  font_size=7, bold=True, color=text_color, align=PP_ALIGN.CENTER)
    _add_text_box(slide, value,
                  left + Inches(0.04), top + lbl_h,
                  width - Inches(0.08), val_h,
                  font_size=9, bold=True, color=text_color, align=PP_ALIGN.CENTER)


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

    # ── Pie chart helper (Slide 5) ────────────────────────────────────────────

    def _add_pie_charts_mn(self, slide):
        """
        Generates and embeds the DESEMBOLSOS and REEMBOLSOS 3D-style pie charts
        that appear in Slide 5 (Liquidez MN 2/2), matching the original design.

        Data keys used:
          desembolsos_pie  : float  — total desembolsos value (single slice when 0)
          reembolsos_pie   : list   — [{"label": str, "valor": float}, …]
        """
        import io
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import matplotlib.patches as mpatches
        except ImportError:
            return  # matplotlib not installed — skip silently

        PIE_COLORS = ["#E8751A", "#C05000", "#A03800", "#FF9A3C", "#7A3000",
                      "#FFC57A", "#5C2D00", "#FFAA55"]

        # ── DESEMBOLSOS chart (left, 6.21",3.34", 2.69x2.53in) ──────────────
        desembolso_val = float(str(self.data.get("desembolsos_total", 0)
                                   ).replace(",", ".").replace(" ", "") or 0)
        fig_d, ax_d = plt.subplots(figsize=(2.69, 2.53), subplot_kw=dict(aspect="equal"))
        fig_d.patch.set_alpha(0)
        ax_d.set_facecolor("none")
        if desembolso_val <= 0:
            # Single full orange circle = zero desembolsos (matches original)
            wedge = mpatches.Wedge((0, 0), 1, 0, 360,
                                   facecolor="#E8751A", edgecolor="#C05000", linewidth=1.5)
            ax_d.add_patch(wedge)
            # 3D shadow ellipse
            shadow = mpatches.Ellipse((0, -0.12), 2.0, 0.35,
                                      facecolor="#C05000", alpha=0.6, zorder=0)
            ax_d.add_patch(shadow)
            ax_d.set_xlim(-1.3, 1.3); ax_d.set_ylim(-0.6, 1.3)
        else:
            ax_d.pie([desembolso_val], colors=["#E8751A"],
                     wedgeprops={"edgecolor": "#C05000", "linewidth": 1.5})
        ax_d.set_title("DESEMBOLSOS", fontsize=9, fontweight="bold",
                       color="black", pad=4,
                       bbox=dict(boxstyle="round,pad=0.2", fc="#C05000",
                                 ec="#C05000"))
        ax_d.axis("off")
        buf_d = io.BytesIO()
        fig_d.savefig(buf_d, format="png", bbox_inches="tight",
                      transparent=True, dpi=150)
        plt.close(fig_d)
        buf_d.seek(0)
        slide.shapes.add_picture(buf_d, Inches(6.21), Inches(3.34),
                                  Inches(2.69), Inches(2.53))

        # ── Kz circle icon between charts ────────────────────────────────────
        _place_img(slide, IMG_ICON_KZ_CIRCLE, Inches(6.39), Inches(5.98),
                   Inches(0.56), Inches(0.41))

        # ── REEMBOLSOS chart (right, 8.52",3.45", 4.09x2.4in) ────────────────
        reemb_items = self.data.get("reembolsos_pie", [])
        if not reemb_items:
            # Empty — just render matching blank orange circle
            fig_r, ax_r = plt.subplots(figsize=(4.09, 2.4),
                                        subplot_kw=dict(aspect="equal"))
            fig_r.patch.set_alpha(0); ax_r.set_facecolor("none")
            ax_r.pie([1], colors=["#E8751A"],
                     wedgeprops={"edgecolor": "#C05000", "linewidth": 1.5})
            ax_r.set_title("REEMBOLSOS", fontsize=9, fontweight="bold",
                           color="black", pad=4,
                           bbox=dict(boxstyle="round,pad=0.2", fc="#C05000",
                                     ec="#C05000"))
            ax_r.axis("off")
            buf_r = io.BytesIO()
            fig_r.savefig(buf_r, format="png", bbox_inches="tight",
                          transparent=True, dpi=150)
            plt.close(fig_r)
        else:
            labels = [r["label"] for r in reemb_items]
            values = [float(str(r["valor"]).replace(",", ".")) for r in reemb_items]
            colors = PIE_COLORS[:len(values)]
            fig_r, ax_r = plt.subplots(figsize=(4.09, 2.4),
                                        subplot_kw=dict(aspect="equal"))
            fig_r.patch.set_alpha(0); ax_r.set_facecolor("none")
            wedges, _texts, autotexts = ax_r.pie(
                values, labels=None, colors=colors, autopct="%1.0f%%",
                pctdistance=0.75, startangle=90,
                wedgeprops={"edgecolor": "white", "linewidth": 1},
            )
            for at in autotexts:
                at.set_fontsize(8); at.set_color("white"); at.set_fontweight("bold")
            # Legend
            ax_r.legend(wedges, labels, loc="center right",
                        bbox_to_anchor=(1.55, 0.5), fontsize=7,
                        frameon=False)
            ax_r.set_title("REEMBOLSOS", fontsize=9, fontweight="bold",
                           color="black", pad=4,
                           bbox=dict(boxstyle="round,pad=0.2", fc="#C05000",
                                     ec="#C05000"))
            ax_r.axis("off")
            buf_r = io.BytesIO()
            fig_r.savefig(buf_r, format="png", bbox_inches="tight",
                          transparent=True, dpi=150)
            plt.close(fig_r)
        buf_r.seek(0)
        slide.shapes.add_picture(buf_r, Inches(8.52), Inches(3.45),
                                  Inches(4.09), Inches(2.4))

    # ── Cambial charts helper (Slide 7) ──────────────────────────────────────

    def _add_cambial_charts(self, slide):
        """
        Embeds two charts on the right half of Slide 7:
          1. Bar chart — Posição Cambial (Activos vs Passivos in M USD)
          2. Line chart — Taxa de Câmbio USD/AKZ over D-2, D-1, D
        """
        import io
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
        except ImportError:
            return

        cambial = self.data.get("cambial", {})
        bar_colors = ["#E8751A", "#5C2D00"]

        # ── Bar chart: Activos vs Passivos ────────────────────────────────────
        activos  = float(str(cambial.get("activos_usd",  0)).replace(",", ".") or 0)
        passivos = float(str(cambial.get("passivos_usd", 0)).replace(",", ".") or 0)

        fig1, ax1 = plt.subplots(figsize=(2.8, 2.2))
        fig1.patch.set_alpha(0)
        ax1.set_facecolor("none")
        bars = ax1.bar(["Activos", "Passivos"], [activos, passivos],
                       color=bar_colors, edgecolor="white", width=0.5)
        for bar, val in zip(bars, [activos, passivos]):
            if val:
                ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                         f"{val:.1f}", ha="center", va="bottom", fontsize=7, color="#333333")
        ax1.set_title("Posição Cambial (M USD)", fontsize=8, fontweight="bold", color="#333333", pad=4)
        ax1.tick_params(axis="both", labelsize=7)
        ax1.spines[["top", "right"]].set_visible(False)
        ax1.set_ylabel("M USD", fontsize=6, color="#555555")
        buf1 = io.BytesIO()
        fig1.savefig(buf1, format="png", bbox_inches="tight", transparent=True, dpi=150)
        plt.close(fig1)
        buf1.seek(0)
        slide.shapes.add_picture(buf1, Inches(6.9), Inches(0.78), Inches(2.9), Inches(2.3))

        # ── Line chart: USD/AKZ over 3 periods ───────────────────────────────
        cambial_rows = self.data.get("cambial_rows", [])
        usd_row = next((r for r in cambial_rows if "USD" in r.get("par", "").upper()
                        and "EUR" not in r.get("par", "").upper()), None)

        if usd_row:
            try:
                d2  = float(str(usd_row.get("anterior2", "0")).replace(",", "."))
                d1  = float(str(usd_row.get("anterior",  "0")).replace(",", "."))
                d   = float(str(usd_row.get("atual",     "0")).replace(",", "."))
                rates = [d2, d1, d]
            except ValueError:
                rates = [0, 0, 0]
        else:
            rates = [0, 0, 0]

        fig2, ax2 = plt.subplots(figsize=(2.8, 2.2))
        fig2.patch.set_alpha(0)
        ax2.set_facecolor("none")
        ax2.plot(["D-2", "D-1", "D"], rates, color="#E8751A", linewidth=2,
                 marker="o", markersize=5, markerfacecolor="#5C2D00")
        ax2.fill_between(["D-2", "D-1", "D"], rates, alpha=0.15, color="#E8751A")
        ax2.set_title("Taxa USD/AKZ", fontsize=8, fontweight="bold", color="#333333", pad=4)
        ax2.tick_params(axis="both", labelsize=7)
        ax2.spines[["top", "right"]].set_visible(False)
        ax2.set_ylabel("AKZ", fontsize=6, color="#555555")
        buf2 = io.BytesIO()
        fig2.savefig(buf2, format="png", bbox_inches="tight", transparent=True, dpi=150)
        plt.close(fig2)
        buf2.seek(0)
        slide.shapes.add_picture(buf2, Inches(10.0), Inches(0.78), Inches(2.9), Inches(2.3))

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

        # Upper orange header image (Angola map + financial imagery)
        if os.path.isfile(IMG_COVER_HEADER):
            slide.shapes.add_picture(IMG_COVER_HEADER,
                                     Inches(0), Inches(0), SLIDE_W, Inches(3.75))
        else:
            _add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(3.75), ORANGE_PRIMARY)

        # White panel for text below header
        _add_rect(slide, Inches(0), Inches(3.75), SLIDE_W, Inches(3.03), WHITE)

        # Thin orange vertical separator
        _add_rect(slide, Inches(0.28), Inches(3.75), Inches(0.08), Inches(2.2), ORANGE_PRIMARY)

        # Main title (italic, red-orange matching original #D44B36)
        _add_text_box(slide, "Resumo Diário Dos Mercados",
                      Inches(0.35), Inches(4.51), Inches(11), Inches(0.75),
                      font_size=34, bold=True, italic=False,
                      color=RGBColor(0xD4, 0x4B, 0x36), align=PP_ALIGN.LEFT)

        # Sub-title
        _add_text_box(slide, "DIRECÇÃO FINANCEIRA",
                      Inches(0.56), Inches(5.12), Inches(8), Inches(0.38),
                      font_size=14, bold=True, color=BLACK, align=PP_ALIGN.LEFT)

        # Date — orange-red
        _add_text_box(slide, date_str,
                      Inches(0.56), Inches(5.52), Inches(8), Inches(0.36),
                      font_size=14, bold=True,
                      color=RGBColor(0xD4, 0x4B, 0x36), align=PP_ALIGN.LEFT)

        # Address box (small text bottom-left)
        _add_text_box(
            slide,
            "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona\nLuanda - Angola",
            Inches(0.28), Inches(6.45), Inches(2.65), Inches(0.29),
            font_size=7, bold=True, color=BLACK, align=PP_ALIGN.LEFT,
        )

        _footer(slide)

    # ── Slide 2: Agenda ───────────────────────────────────────────────────────

    def _slide_agenda(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")

        # Full-slide background image
        if os.path.isfile(IMG_AGENDA_BG):
            slide.shapes.add_picture(IMG_AGENDA_BG,
                                     Inches(0), Inches(0), SLIDE_W, SLIDE_H)
        # Inner header banner on top
        if os.path.isfile(IMG_INNER_HEADER):
            slide.shapes.add_picture(IMG_INNER_HEADER,
                                     Inches(0), Inches(0), SLIDE_W, Inches(3.37))

        # "AGENDA" label
        _add_text_box(slide, "AGENDA",
                      Inches(0.4), Inches(3.81), Inches(1.82), Inches(0.43),
                      font_size=25, bold=True,
                      color=RGBColor(0xD3, 0x4A, 0x36), align=PP_ALIGN.LEFT)

        items = [
            ("1.", "Sumário Executivo"),
            ("2.", "Liquidez (MN)"),
            ("3.", "Liquidez (ME)"),
            ("4.", "Mercado Cambial"),
            ("5.", "Mercado\nCapitais"),
            ("6.", "Informação\nDe Mercado"),
        ]
        col_lefts = [Inches(0.40), Inches(2.01), Inches(3.53), Inches(5.11), Inches(6.42), Inches(7.97)]
        for i, ((num, label), col_left) in enumerate(zip(items, col_lefts)):
            top = Inches(4.38)
            _add_text_box(slide, num, col_left, top, Inches(0.35), Inches(0.5),
                          font_size=19, bold=True,
                          color=RGBColor(0xEB, 0x8B, 0x34), align=PP_ALIGN.LEFT)
            _add_text_box(slide, label,
                          col_left, top + Inches(0.38), Inches(1.7), Inches(1.2),
                          font_size=14, bold=True, color=BLACK, align=PP_ALIGN.LEFT)

        # BW report icon — bottom-right (12.51",5.39") 0.51x0.51in
        _place_img(slide, IMG_ICON_REPORT_BW, Inches(12.51), Inches(5.39),
                   Inches(0.51), Inches(0.51))

        _footer(slide)

    # ── Slide 3: Sumário Executivo ────────────────────────────────────────────

    def _slide_sumario_executivo(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "Sumário Executivo", date_str)

        # ── Chevron arrows (header area, exact original positions) ────────────
        # Two overlapping chevrons at top-left, orange fill
        for left_in in (0.07, 0.39):
            sh = slide.shapes.add_shape(13, Inches(left_in), Inches(0.76),
                                        Inches(0.47), Inches(0.38))  # 13 = chevron right
            sh.fill.solid(); sh.fill.fore_color.rgb = ORANGE_PRIMARY
            sh.line.fill.background()

        # ── Central oval KPI ──────────────────────────────────────────────────
        rc_val = self.data.get("reembolso_credito", "—")
        cx, cy = Inches(5.741), Inches(3.009)
        bw, bh = Inches(1.779), Inches(1.491)
        # Draw oval using shape type 9 (oval)
        oval = slide.shapes.add_shape(9, cx, cy, bw, bh)
        oval.fill.solid(); oval.fill.fore_color.rgb = BROWN_KPI
        oval.line.fill.background()
        _add_text_box(slide, rc_val,
                      cx + Inches(0.1), cy + Inches(0.1), bw - Inches(0.2), Inches(0.65),
                      font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, "Juros de DP",
                      cx + Inches(0.1), cy + Inches(0.75), bw - Inches(0.2), Inches(0.4),
                      font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # ── Satellite KPI cards ───────────────────────────────────────────────
        # Original layout: 8 cards arranged around central oval.
        # Each card: value box (orange bg) + label + vertical bar divider + % change box.
        # Icons are placed above/beside each card.
        kpis = self.data.get("kpis", [
            {"label": "Liquidez MN",           "value": "—",  "variation_str": ""},
            {"label": "Liquidez ME",            "value": "—",  "variation_str": ""},
            {"label": "Posição Cambial",        "value": "—",  "variation_str": ""},
            {"label": "Carteira Títulos",       "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade MN",       "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade ME",       "value": "—",  "variation_str": ""},
            {"label": "Rentabilidade Títulos",  "value": "—",  "variation_str": ""},
            {"label": "Reembolsos",             "value": "—",  "variation_str": ""},
        ])

        # Map KPI index → (left_x, top_y, icon_path, icon_w, icon_h)
        # Positions derived from original group shape coordinates
        kpi_layout = [
            # left cards (index 0-3) — original x ~ 1.76"–4.86"
            (Inches(1.76), Inches(2.01), IMG_ICON_S3_LMN,     Inches(0.44), Inches(0.55)),  # Liquidez MN
            (Inches(6.33), Inches(2.21), IMG_ICON_S3_LME,     Inches(0.43), Inches(0.38)),  # Liquidez ME
            (Inches(0.49), Inches(4.15), IMG_ICON_S3_CAMBIAL, Inches(0.56), Inches(0.44)),  # Posição Cambial
            (Inches(6.87), Inches(4.02), IMG_ICON_S3_REEMB,   Inches(0.55), Inches(0.43)),  # Reembolsos
            # right cards (index 4-7)
            (Inches(8.03), Inches(1.47), IMG_ICON_S3_RENTA,   Inches(0.46), Inches(0.30)),  # Rentabilidade MN
            (Inches(8.96), Inches(2.33), IMG_ICON_S3_RENTA,   Inches(0.46), Inches(0.30)),  # Rentabilidade ME
            (Inches(8.02), Inches(1.77), IMG_ICON_S3_RENTA,   Inches(0.46), Inches(0.30)),  # Rentabilidade Títulos
            (Inches(9.48), Inches(4.19), IMG_ICON_S3_REEMB,   Inches(0.55), Inches(0.43)),  # Desembolsos
        ]

        for i, kpi in enumerate(kpis[:8]):
            if i >= len(kpi_layout):
                break
            vl, vt, icon_path, iw, ih = kpi_layout[i]
            # Place icon
            _place_img(slide, icon_path, vl, vt, iw, ih)
            # Value card (orange bg, bold)
            card_w, card_h = Inches(2.58), Inches(1.21)
            card_left = vl + Inches(0.5)
            _add_rect(slide, card_left, vt, card_w, card_h,
                      ORANGE_LIGHT, ORANGE_PRIMARY, 1.0)
            _add_text_box(slide, kpi["label"],
                          card_left + Pt(4), vt + Pt(3), card_w - Pt(8), Inches(0.32),
                          font_size=9, color=DARK_GREY, align=PP_ALIGN.LEFT)
            _add_text_box(slide, kpi["value"],
                          card_left + Pt(4), vt + Inches(0.35), card_w - Pt(8), Inches(0.5),
                          font_size=15, bold=True, color=ORANGE_PRIMARY, align=PP_ALIGN.LEFT)
            # Variation — placed at bottom of card (avoids going off-screen)
            if kpi.get("variation_str"):
                vc = _variation_color(kpi["variation_str"])
                _add_text_box(slide, kpi["variation_str"],
                              card_left + Pt(4), vt + Inches(0.85), card_w - Pt(8), Inches(0.28),
                              font_size=9, bold=True, color=vc, align=PP_ALIGN.LEFT)

        # Enquadramento and Breve Conclusão are intentionally excluded from dashboard layout

        # Coloured report icon — bottom-right (original: 0.35x0.35" @ 10.74",6.42")
        _place_img(slide, IMG_ICON_S3_REPORT, Inches(10.74), Inches(6.42),
                   Inches(0.35), Inches(0.35))

        _footer(slide)

    # ── Slide 4: Liquidez MN — tables + LUIBOR ────────────────────────────────

    def _slide_liquidez_mn_1(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "LIQUIDEZ – MOEDA NACIONAL", date_str)

        days    = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])
        row_h   = Inches(0.28)
        # Reference: L=0.407 W=8.557 — leave right panel (L>9") for KPI graphics
        left0   = Inches(0.407)
        tbl_w   = Inches(8.557)
        label_w = Inches(3.0)
        col_w   = (tbl_w - label_w) / 5
        lefts   = [left0] + [left0 + label_w + i * col_w for i in range(5)]
        widths  = [label_w] + [col_w] * 5

        # ── Section: Liquidez MN table ─────────────────────────────────────────
        # Reference: Objeto 28 at T=0.7 H=1.417
        top = Inches(0.70)
        _section_bar(slide, "Liquidez MN  (Em Milhares)", left0, top, tbl_w)
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

        # ── Section: Transações (OMA) — reference Objeto 29 at T=2.204 H=0.741
        top = Inches(2.204)
        _section_bar(slide, "Transações", left0, top, tbl_w)
        tx_heads  = ["Tipo", "Contraparte", "Taxa", "Montante", "Maturidade", "Juros"]
        tx_w_list = [Inches(1.1), Inches(1.8), Inches(0.9), Inches(1.8), Inches(1.6), Inches(1.357)]
        tx_lefts  = [left0]
        for w in tx_w_list[:-1]:
            tx_lefts.append(tx_lefts[-1] + w)
        top += Inches(0.28)
        _table_header_row(slide, tx_heads, tx_lefts, top, row_h, tx_w_list)

        raw_tx = self.data.get("transacoes_mn_raw", [])
        for i, row in enumerate(raw_tx):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            vals = [
                row.get("tipo", "OMA"), row.get("contraparte", "—"),
                row.get("taxa", "—"),   row.get("montante", "—"),
                row.get("maturidade", "—"), row.get("juros", "—"),
            ]
            _table_data_row(slide, vals, tx_lefts, top, row_h, tx_w_list, bg=bg)
        if not raw_tx:
            top += row_h
            _table_data_row(slide, ["—"] * 6, tx_lefts, top, row_h, tx_w_list)

        # ── Section: Operações Vivas — reference Objeto 31 at T=3.012 H=2.116
        top = Inches(3.012)
        _section_bar(slide, "Operações Vivas", left0, top, tbl_w)
        op_heads  = ["Tipo", "Contraparte", "Montante", "Taxa", "Residual", "Vencimento", "Juro Diário"]
        n_op      = len(op_heads)
        op_w      = tbl_w / n_op
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
        _section_bar(slide, "Taxas LUIBOR", left0, top, tbl_w)
        tenors    = ["LUIBOR O/N", "LUIBOR 1M", "LUIBOR 3M", "LUIBOR 6M", "LUIBOR 9M", "LUIBOR 12M"]
        lu_heads  = ["Maturidade", "Anterior (D-2)", "Anterior (D-1)", "Actual (D)", "Var (%)"]
        lu_n      = len(lu_heads)
        lu_w      = tbl_w / lu_n
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
            _table_data_row(
                slide,
                [t, rate, rate, rate, var],
                lu_lefts, top, row_h, lu_widths, bg=bg,
            )

        # ── Summary ovals (Liquidez Total + Juros Diário MN) — exact original positions
        lmn_rows = self.data.get("liquidez_mn_rows", [])
        total_mn = lmn_rows[-1]["values"][-1] if lmn_rows else "—"
        juros_mn = self.data.get("juros_diario_mn", "—")
        # Oval 4: L=2.680 T=5.384 W=1.096 H=0.865
        _summary_oval(slide, "Liquidez Total", total_mn,
                      Inches(2.680), Inches(5.384), Inches(1.096), Inches(0.865))
        # Oval 5: L=4.202 T=5.403 W=1.096 H=0.865
        _summary_oval(slide, "Juros Diário", juros_mn,
                      Inches(4.202), Inches(5.403), Inches(1.096), Inches(0.865))

        # Icons — exact original positions
        _place_img(slide, IMG_ICON_GEAR,       Inches(6.37), Inches(5.45), Inches(0.81), Inches(0.78))
        _place_img(slide, IMG_ICON_CALCULATOR, Inches(7.74), Inches(5.59), Inches(0.57), Inches(0.67))

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

        # ── Pie charts — DESEMBOLSOS and REEMBOLSOS ───────────────────────────
        self._add_pie_charts_mn(slide)

        _footer(slide)

    # ── Slide 6: Liquidez ME ──────────────────────────────────────────────────
    # Reference layout: 3 tables on left (W=7.784), Fluxos ME on right (L=8.23)

    def _slide_liquidez_me(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "LIQUIDEZ – MOEDA ESTRANGEIRA", date_str)

        days  = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])
        row_h = Inches(0.25)
        fs    = 7

        # ── LEFT side (W=7.784) — 3 tables stacked ───────────────────────────
        left0  = Inches(0.396)
        tbl_w  = Inches(7.784)
        lbl_w  = Inches(2.3)
        col_w  = (tbl_w - lbl_w) / 5
        lefts  = [left0] + [left0 + lbl_w + i * col_w for i in range(5)]
        widths = [lbl_w] + [col_w] * 5

        # Liquidez ME — T=0.677 H≈1.0 (reference Objeto 5)
        top = Inches(0.677)
        _section_bar(slide, "Liquidez ME  (M USD)", left0, top, tbl_w)
        top += Inches(0.25)
        _table_header_row(slide, [""] + days, lefts, top, row_h, widths, font_size=fs)

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
                            lefts, top, row_h, widths, font_size=fs,
                            highlight=is_total, bg=bg)

        # Transações ME — T=1.838 H≈0.892 (reference Objeto 13)
        top = Inches(1.838)
        _section_bar(slide, "Transações ME", left0, top, tbl_w)
        tx_me_heads = ["Tipo", "Moeda", "Contraparte", "Taxa", "Montante", "Maturidade", "Juros"]
        n_tx  = len(tx_me_heads)
        tx_w  = tbl_w / n_tx
        tx_lefts  = [left0 + i * tx_w for i in range(n_tx)]
        tx_widths = [tx_w] * n_tx
        top += Inches(0.25)
        _table_header_row(slide, tx_me_heads, tx_lefts, top, row_h, tx_widths, font_size=fs)

        tx_me = self.data.get("transacoes_me_raw", [])
        for i, row in enumerate(tx_me):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [row.get(k, "—") for k in
                             ("tipo", "moeda", "contraparte", "taxa", "montante", "maturidade", "juros")],
                            tx_lefts, top, row_h, tx_widths, font_size=fs, bg=bg)
        if not tx_me:
            top += row_h
            _table_data_row(slide, ["—"] * n_tx, tx_lefts, top, row_h, tx_widths, font_size=fs)

        # Operações Vivas ME — T=2.819 H≈2.45 (reference Objeto 19)
        top = Inches(2.819)
        _section_bar(slide, "Operações Vivas ME", left0, top, tbl_w)
        op_me_heads = ["Contraparte", "Montante", "Taxa", "Residual (Dias)", "Vencimento", "Juro Diário"]
        n_op = len(op_me_heads)
        op_w = tbl_w / n_op
        op_lefts  = [left0 + i * op_w for i in range(n_op)]
        op_widths = [op_w] * n_op
        top += Inches(0.25)
        _table_header_row(slide, op_me_heads, op_lefts, top, row_h, op_widths, font_size=fs)

        ops_me = self.data.get("operacoes_vivas_me", [])
        for i, op in enumerate(ops_me):
            top += row_h
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide,
                            [op.get(k, "—") for k in
                             ("contraparte", "montante", "taxa", "residual", "vencimento", "juro_diario")],
                            op_lefts, top, row_h, op_widths, font_size=fs, bg=bg)
        if not ops_me:
            top += row_h
            _table_data_row(slide, ["—"] * n_op, op_lefts, top, row_h, op_widths, font_size=fs)

        # ── RIGHT side — Fluxos de Caixa ME (reference Objeto 21 at L=8.23 T=3.078)
        R_left  = Inches(8.23)
        R_width = Inches(5.021)
        fx_lbl_w = Inches(1.8)
        fx_col_w = (R_width - fx_lbl_w) / 5
        fx_lefts  = [R_left] + [R_left + fx_lbl_w + i * fx_col_w for i in range(5)]
        fx_widths = [fx_lbl_w] + [fx_col_w] * 5

        r_top = Inches(3.078)
        _section_bar(slide, "Fluxos de Caixa ME", R_left, r_top, R_width)
        r_top += Inches(0.25)
        _table_header_row(slide, [""] + days, fx_lefts, r_top, row_h, fx_widths, font_size=fs)

        fluxos_me = self.data.get("fluxos_me_rows", [
            {"label": "Cash in flow",      "values": ["—"] * 5},
            {"label": "Outros receb.",     "values": ["—"] * 5},
            {"label": "Remb. DP + Juros",  "values": ["—"] * 5},
            {"label": "Cash out flow",     "values": ["—"] * 5},
            {"label": "Aplicação DP ME",   "values": ["—"] * 5},
            {"label": "GAP de Liquidez",   "values": ["—"] * 5},
        ])
        for i, row in enumerate(fluxos_me):
            r_top += row_h
            is_total = "GAP" in row["label"].upper() or "TOTAL" in row["label"].upper()
            bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
            _table_data_row(slide, [row["label"]] + row["values"],
                            fx_lefts, r_top, row_h, fx_widths, font_size=fs,
                            highlight=is_total, bg=bg)

        # ── Summary ovals — exact reference: T=5.260 (Agrupar 16 at L=2.025)
        lme_data = self.data.get("liquidez_me_rows", [])
        total_me = lme_data[-1]["values"][-1] if lme_data else "—"
        juros_me = self.data.get("juros_diario_me", "—")
        _summary_oval(slide, "Liquidez Total", total_me,
                      Inches(2.817), Inches(5.260), Inches(1.176), Inches(0.941))
        _summary_oval(slide, "Juros Diário", juros_me,
                      Inches(4.350), Inches(5.260), Inches(1.175), Inches(0.941))
        _place_img(slide, IMG_ICON_CALCULATOR, Inches(5.884), Inches(5.607),
                   Inches(0.615), Inches(0.712))

        _footer(slide)

    # ── Slide 7: Mercado Cambial ──────────────────────────────────────────────

    def _slide_mercado_cambial(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "MERCADO CAMBIAL", date_str)

        cambial  = self.data.get("cambial", {})
        row_h    = Inches(0.28)
        # Reference: left tables W=6.112 (Objeto 4 at L=0.396 W=6.112)
        left0    = Inches(0.396)
        table_w  = Inches(6.112)

        # ── Cambiais rates table — reference T=0.705 H=0.899
        top = Inches(0.705)
        _section_bar(slide, "Cambiais", left0, top, table_w)
        c_heads  = ["Par", "Anterior (D-2)", "Anterior (D-1)", "Actual (D)", "(%)"]
        c_w      = [Inches(1.4), Inches(1.178), Inches(1.178), Inches(1.178), Inches(1.178)]
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

        # ── Transações BDA table — reference T=1.865 H=0.7
        top = Inches(1.865)
        _section_bar(slide, "Transações BDA", left0, top, table_w)
        tb_heads  = ["C/V", "Par de moeda", "Montante Debt", "Câmbio", "P/L AKZ"]
        tb_w      = [Inches(0.7), Inches(1.412), Inches(1.5), Inches(1.3), Inches(1.2)]
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

        # ── Transações do Mercado — reference Objeto 19 at T=3.047 W=5.417
        top = Inches(3.047)
        _section_bar(slide, "Transações do Mercado", left0, top, Inches(5.417))
        tm_heads = ["Liquidação", "Montante USD", "Mínimo", "Máximo"]
        tm_w = [Inches(1.2), Inches(1.539), Inches(1.339), Inches(1.339)]
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

        # ── Right panel — KPI ovals at original positions + charts below
        # Oval USD: reference L=8.502 T=1.422 W=1.241 H=0.976
        _summary_oval(slide, "Transações (USD)", cambial.get("vol_total_usd", "—"),
                      Inches(8.502), Inches(1.422), Inches(1.241), Inches(0.976))
        # Oval Kz: reference L=10.308 T=1.454 W=1.241 H=0.954
        _summary_oval(slide, "Posição Cambial (Kz)", cambial.get("posicao_cambial", "—"),
                      Inches(10.308), Inches(1.454), Inches(1.241), Inches(0.954))

        # Charts: Posição Cambial bar (L=6.212 T=3.26) + Taxa de Câmbio line (bottom)
        self._add_cambial_charts(slide)

        _place_img(slide, IMG_ICON_FX_EXCHANGE, Inches(12.089), Inches(0.667),
                   Inches(0.473), Inches(0.467))
        _place_img(slide, IMG_ICON_REPORT_MONEY, Inches(6.554), Inches(2.453),
                   Inches(0.543), Inches(0.536))

        _footer(slide)

    # ── Slide 8: Mercado de Capitais – BODIVA ─────────────────────────────────

    def _slide_bodiva(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "MERCADO DE CAPITAIS", date_str)

        row_h  = Inches(0.28)
        # Reference: segment table L=0.396 T=0.721 W=7.494; stocks L=0.396 T=4.622 W=8.825
        left0  = Inches(0.396)

        # ── Segmentado por Produtos — W=7.494 matching reference
        top = Inches(0.721)
        seg_tbl_w = Inches(7.494)
        _section_bar(slide, "Segmentado Por Produtos", left0, top, seg_tbl_w)

        sp_heads  = ["Segmento", "Anterior", "Actual", "(%)"]
        sp_w      = [Inches(3.0), Inches(1.498), Inches(1.498), Inches(1.498)]
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

        # KPI summary oval — exact original: L=4.721 T=3.159 W=1.473 H=1.058
        total_atual = self.data.get("bodiva_total_transacoes", "—")
        _summary_oval(slide, "Kz  Transações", total_atual,
                      Inches(4.721), Inches(3.159), Inches(1.473), Inches(1.058))

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

        # ── Carteira de Títulos — split into Custo Amortizado / Justo Valor ──
        top = Inches(1.85)
        carteira = self.data.get("carteira_titulos", [])

        custo_rows = [r for r in carteira if "CUSTO" in r.get("carteira", "").upper()]
        justo_rows = [r for r in carteira if "JUSTO" in r.get("carteira", "").upper()
                      or "VALOR" in r.get("carteira", "").upper()]
        # If no categorisation present, split evenly for display
        if not custo_rows and not justo_rows:
            mid = max(1, len(carteira) // 2)
            custo_rows = carteira[:mid]
            justo_rows = carteira[mid:]

        ct_heads  = ["Cód.", "Qtd D-1", "Qtd D", "Nominal", "Taxa", "Montante", "J. Anual", "J. Diário"]
        n_ct      = len(ct_heads)
        ct_w_each = (SLIDE_W - Inches(0.5)) / n_ct
        ct_lefts  = [left0 + i * ct_w_each for i in range(n_ct)]
        ct_widths = [ct_w_each] * n_ct

        def _render_carteira_section(label, rows, start_top):
            _section_bar(slide, label, left0, start_top, SLIDE_W - Inches(0.5))
            t = start_top + Inches(0.24)
            _table_header_row(slide, ct_heads, ct_lefts, t, row_h, ct_widths, font_size=6)
            if rows:
                for j, row in enumerate(rows):
                    t += row_h
                    is_total = row.get("cod", "").upper() == "TOTAL"
                    bg = ORANGE_LIGHT if j % 2 == 0 else WHITE
                    _table_data_row(
                        slide,
                        [row.get(k, "—") for k in
                         ("cod", "qty_d1", "qty_d", "nominal", "taxa", "montante", "juros_anual", "juro_diario")],
                        ct_lefts, t, row_h, ct_widths, font_size=6,
                        highlight=is_total, bg=bg,
                    )
            else:
                t += row_h
                _table_data_row(slide, ["—"] * n_ct, ct_lefts, t, row_h, ct_widths, font_size=6)
            return t + row_h

        top = _render_carteira_section("Custo Amortizado", custo_rows, top)
        top += Inches(0.08)
        _render_carteira_section("Justo Valor", justo_rows, top)

        # KPI ovals — bottom of slide
        tx_kpi_val = self.data.get("bodiva_transacoes_valor", "0,00 mM Kz")
        jd_kpi_val = self.data.get("bodiva_juros_diario", "—")
        _summary_oval(slide, "Kz  Transações",   tx_kpi_val,
                      Inches(4.50), Inches(5.85), Inches(1.30), Inches(0.88))
        _summary_oval(slide, "Kz  Juros Diário", jd_kpi_val,
                      Inches(6.00), Inches(5.85), Inches(1.30), Inches(0.88))

        _footer(slide)

    # ── Slide 10: Informação de Mercados (1/2) — Indices + Crypto ─────────────

    def _slide_market_info_1(self):
        slide    = self.prs.slides.add_slide(self._blank)
        date_str = self.data.get("report_date", "")
        _slide_title(slide, "INFORMAÇÃO DE MERCADOS", date_str)

        market = self.data.get("market_info", {})
        row_h  = Inches(0.28)
        left0  = Inches(0.3)
        # Left table width, right commentary panel — exact original positions
        tbl_w  = Inches(6.0)
        # Commentary box: L=7.134 T=0.662 W=5.702 H=2.077  (original Rectangle 15)
        com_x  = Inches(7.134)
        com_w  = Inches(5.702)

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

        # Commentary — right panel: L=7.134 T=0.662 W=5.702 H=2.077
        cm_comment = market.get("cm_commentary", "")
        if cm_comment:
            _add_rect(slide, com_x, Inches(0.662), com_w, Inches(2.077),
                      ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cm_comment,
                          com_x + Pt(6), Inches(0.69), com_w - Pt(12), Inches(2.0),
                          font_size=8, color=BLACK, word_wrap=True)

        # "Nota" tag — exact original: L=12.340 T=3.327 W=0.598 H=0.315
        _add_rect(slide, Inches(12.340), Inches(3.327), Inches(0.598), Inches(0.315),
                  ORANGE_PRIMARY)
        _add_text_box(slide, "Nota", Inches(12.344), Inches(3.331),
                      Inches(0.590), Inches(0.307),
                      font_size=8, bold=True, italic=True,
                      color=WHITE, align=PP_ALIGN.CENTER)

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

        # Crypto commentary — exact original: L=7.083 T=4.068 W=5.804 H=1.386
        cr_comment = market.get("crypto_commentary", "")
        if cr_comment:
            _add_rect(slide, Inches(7.083), Inches(4.068), Inches(5.804), Inches(1.386),
                      ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cr_comment,
                          Inches(7.10), Inches(4.09), Inches(5.760), Inches(1.340),
                          font_size=8, color=BLACK, word_wrap=True)

        # Icon — report+money: exact original L=6.619 T=1.404 (from chevron cluster area)
        _place_img(slide, IMG_ICON_REPORT_MONEY, Inches(6.43), Inches(0.75),
                   Inches(0.39), Inches(0.39))

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
        # Commentary panels — exact original positions
        # Commodities box: L=7.052 T=0.723 W=5.572 H=1.572
        # Minerals box:    L=7.187 T=3.792 W=5.465 H=1.413

        # ── Commodities ───────────────────────────────────────────────────────
        top = Inches(0.78)
        _section_bar(slide, "Commodities", left0, top, tbl_w)
        cmd_heads = ["Commodity", "Anterior", "Actual", "(%)"]
        cmd_w     = [Inches(2.4), Inches(1.2), Inches(1.2), Inches(1.2)]
        cmd_lefts = [left0]
        for _h in cmd_w[:-1]:
            cmd_lefts.append(cmd_lefts[-1] + _h)
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

        # Commodities commentary — exact original: L=7.052 T=0.723 W=5.572 H=1.572
        cmd_comment = market.get("commodities_commentary", "")
        if cmd_comment:
            _add_rect(slide, Inches(7.052), Inches(0.723), Inches(5.572), Inches(1.572),
                      ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, cmd_comment,
                          Inches(7.08), Inches(0.75), Inches(5.53), Inches(1.52),
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

        # Minerals commentary — exact original: L=7.187 T=3.792 W=5.465 H=1.413
        min_comment = market.get("minerais_commentary", "")
        if min_comment:
            _add_rect(slide, Inches(7.187), Inches(3.792), Inches(5.465), Inches(1.413),
                      ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
            _add_text_box(slide, min_comment,
                          Inches(7.22), Inches(3.82), Inches(5.42), Inches(1.37),
                          font_size=8, color=BLACK, word_wrap=True)

        # "Nota" — dedicated shaded box at bottom of slide
        nota_text = market.get("commodities_nota", market.get("minerais_commentary", ""))
        nota_top = Inches(5.60)
        nota_h   = Inches(0.85)
        _add_rect(slide, Inches(0.25), nota_top, SLIDE_W - Inches(0.5), nota_h,
                  ORANGE_LIGHT, ORANGE_PRIMARY, 0.8)
        _add_text_box(slide, f"Nota:  {nota_text}" if nota_text else "Nota:",
                      Inches(0.35), nota_top + Pt(4), SLIDE_W - Inches(0.7), nota_h - Pt(8),
                      font_size=7, color=BLACK, word_wrap=True)

        _place_img(slide, IMG_ICON_REPORT_MONEY, Inches(6.28), Inches(0.90),
                   Inches(0.39), Inches(0.39))

        _footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# Backward-compatible alias + quick smoke-test
# ─────────────────────────────────────────────────────────────────────────────

PPTXBuilder = BDAReportGenerator   # legacy alias

if __name__ == "__main__":
    gen  = BDAReportGenerator()
    path = gen.build("output/bda_report_test.pptx")
    print(f"Test report saved: {path}")
