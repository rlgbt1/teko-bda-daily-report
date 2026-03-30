"""
report_generator/pdf_builder.py — BDA Daily Report PDF Generator.

Produces a styled, BDA-branded PDF from the same data dict used by
BDAReportGenerator (pptx_builder.py).  Uses reportlab only — no system
tools (LibreOffice, etc.) required.

Structure mirrors the PPTX:
  Page 1:  Cover
  Page 2:  Agenda
  Page 3:  Sumário Executivo (KPIs)
  Page 4:  Liquidez – Moeda Nacional (1/2)
  Page 5:  Liquidez – Moeda Nacional (2/2)
  Page 6:  Liquidez – Moeda Estrangeira
  Page 7:  Mercado Cambial
  Page 8:  Mercado de Capitais – BODIVA
  Page 9:  Mercado de Capitais – Operações BDA
  Page 10: Informação de Mercados (1/2)
  Page 11: Informação de Mercados (2/2)

Usage:
    from src.report_generator.pdf_builder import BDAReportPDF
    gen  = BDAReportPDF(data)
    path = gen.build("output/bda_report_2026-03-30.pdf")
"""
from __future__ import annotations

import os
from datetime import date
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── BDA colour palette (matches pptx_builder.py) ─────────────────────────────
ORANGE_PRIMARY = colors.HexColor("#E8751A")
ORANGE_DARK    = colors.HexColor("#C05000")
ORANGE_LIGHT   = colors.HexColor("#FFF0E0")
BROWN_KPI      = colors.HexColor("#5C2D00")
WHITE          = colors.white
BLACK          = colors.black
LIGHT_GREY     = colors.HexColor("#F5F5F5")
MID_GREY       = colors.HexColor("#D9D9D9")
DARK_GREY      = colors.HexColor("#555555")
GREEN_UP       = colors.HexColor("#007A33")
RED_DOWN       = colors.HexColor("#CC0000")

# ── Page setup ────────────────────────────────────────────────────────────────
PAGE_SIZE  = landscape(A4)   # 297 × 210 mm — matches 16:9 slide proportions
MARGIN     = 1.2 * cm

# ── Styles ────────────────────────────────────────────────────────────────────
_base = getSampleStyleSheet()

def _style(name, **kwargs) -> ParagraphStyle:
    s = ParagraphStyle(name, parent=_base["Normal"], **kwargs)
    return s

STYLE_TITLE   = _style("BDA_Title",   fontSize=26, textColor=ORANGE_PRIMARY,
                        fontName="Helvetica-BoldOblique", spaceAfter=4)
STYLE_SECTION = _style("BDA_Section", fontSize=10, textColor=WHITE,
                        fontName="Helvetica-Bold",         backColor=ORANGE_DARK,
                        leftIndent=4, spaceAfter=2, spaceBefore=6)
STYLE_BODY    = _style("BDA_Body",    fontSize=8,  textColor=BLACK,
                        fontName="Helvetica",              spaceAfter=2)
STYLE_SMALL   = _style("BDA_Small",   fontSize=7,  textColor=DARK_GREY,
                        fontName="Helvetica")
STYLE_KPI_LBL = _style("BDA_KpiLbl", fontSize=7,  textColor=DARK_GREY,
                        fontName="Helvetica")
STYLE_KPI_VAL = _style("BDA_KpiVal", fontSize=16, textColor=ORANGE_PRIMARY,
                        fontName="Helvetica-Bold")
STYLE_COMMENT = _style("BDA_Comment", fontSize=7.5, textColor=BLACK,
                        fontName="Helvetica",  backColor=ORANGE_LIGHT,
                        leftIndent=4, rightIndent=4, spaceAfter=4, spaceBefore=4)


# ─────────────────────────────────────────────────────────────────────────────
# Reusable flowable helpers
# ─────────────────────────────────────────────────────────────────────────────

def _section_header(label: str) -> list:
    """Orange bar + white bold label."""
    return [
        Table(
            [[Paragraph(f"<b>{label}</b>", STYLE_SECTION)]],
            colWidths=[PAGE_SIZE[0] - 2 * MARGIN],
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), ORANGE_DARK),
                ("TEXTCOLOR",  (0, 0), (-1, -1), WHITE),
                ("TOPPADDING",    (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ]),
        ),
    ]


def _page_title(title: str, date_str: str = "") -> list:
    """Italic orange title + thin rule."""
    row = [[Paragraph(f"<i><b>{title}</b></i>", STYLE_TITLE),
            Paragraph(date_str, _style("dt", fontSize=9, textColor=DARK_GREY,
                                       fontName="Helvetica", alignment=TA_RIGHT))]]
    return [
        Table(row,
              colWidths=[PAGE_SIZE[0] * 0.75 - MARGIN, PAGE_SIZE[0] * 0.25 - MARGIN],
              style=TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")])),
        HRFlowable(width="100%", thickness=1.5, color=ORANGE_PRIMARY, spaceAfter=6),
    ]


def _data_table(headers: list[str], rows: list[list[str]],
                col_widths=None,
                font_size: int = 8) -> Table:
    """
    Standard BDA data table:
      - Orange header row (white bold text)
      - Alternating light-orange / white data rows
    """
    available = PAGE_SIZE[0] - 2 * MARGIN
    n = len(headers)
    if col_widths is None:
        col_widths = [available / n] * n

    data = [headers] + rows

    style = TableStyle([
        # Header
        ("BACKGROUND",   (0, 0), (-1, 0),  ORANGE_DARK),
        ("TEXTCOLOR",    (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",     (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, 0),  font_size),
        ("ALIGN",        (0, 0), (-1, 0),  "CENTER"),
        ("TOPPADDING",   (0, 0), (-1, 0),  3),
        ("BOTTOMPADDING",(0, 0), (-1, 0),  3),
        # Data rows
        ("FONTNAME",     (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",     (0, 1), (-1, -1), font_size),
        ("ALIGN",        (1, 1), (-1, -1), "CENTER"),
        ("ALIGN",        (0, 1), (0, -1),  "LEFT"),
        ("TOPPADDING",   (0, 1), (-1, -1), 2),
        ("BOTTOMPADDING",(0, 1), (-1, -1), 2),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID",         (0, 0), (-1, -1), 0.4, MID_GREY),
    ])

    # Alternating rows
    for i, row in enumerate(rows):
        bg = ORANGE_LIGHT if i % 2 == 0 else WHITE
        style.add("BACKGROUND", (0, i + 1), (-1, i + 1), bg)
        # Highlight total rows
        text = str(row[0]).upper() if row else ""
        if any(k in text for k in ("TOTAL", "LIQUIDEZ BDA", "GAP", "LÍQUIDO", "RESULTADO")):
            style.add("BACKGROUND", (0, i + 1), (-1, i + 1), ORANGE_LIGHT)
            style.add("FONTNAME",   (0, i + 1), (-1, i + 1), "Helvetica-Bold")
            style.add("TEXTCOLOR",  (0, i + 1), (-1, i + 1), ORANGE_DARK)

    return Table(data, colWidths=col_widths, style=style, repeatRows=1)


def _kpi_row(kpis: list[dict]) -> Table:
    """
    Renders a row of KPI cards:  [{label, value, variation_str}]
    Max 4 per row looks good on landscape A4.
    """
    n = min(len(kpis), 4)
    card_w = (PAGE_SIZE[0] - 2 * MARGIN) / n

    cells = []
    for kpi in kpis[:n]:
        var = kpi.get("variation_str", "")
        if var:
            var_color = "green" if not str(var).startswith("-") else "red"
            var_para = Paragraph(
                f'<font color="{var_color}"><b>{var}</b></font>',
                _style("vp", fontSize=7, fontName="Helvetica"),
            )
        else:
            var_para = Paragraph("", STYLE_SMALL)

        cell = [
            Paragraph(kpi.get("label", ""), STYLE_KPI_LBL),
            Paragraph(str(kpi.get("value", "—")), STYLE_KPI_VAL),
            var_para,
        ]
        cells.append(cell)

    t = Table([cells],
              colWidths=[card_w] * n,
              style=TableStyle([
                  ("BOX",        (0, 0), (-1, -1), 1.2, ORANGE_PRIMARY),
                  ("INNERGRID",  (0, 0), (-1, -1), 0.5, ORANGE_LIGHT),
                  ("BACKGROUND", (0, 0), (-1, -1), ORANGE_LIGHT),
                  ("VALIGN",     (0, 0), (-1, -1), "TOP"),
                  ("TOPPADDING",    (0, 0), (-1, -1), 5),
                  ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                  ("LEFTPADDING",   (0, 0), (-1, -1), 6),
              ]))
    return t


def _footer_content(date_str: str) -> str:
    return (
        "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona  |  Luanda – Angola"
        f"{'   |   ' + date_str if date_str else ''}"
        "   |   UMA VISÃO DE FUTURO"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Page template with orange top stripe + footer
# ─────────────────────────────────────────────────────────────────────────────

def _make_page_template(doc, date_str: str) -> PageTemplate:
    frame = Frame(
        MARGIN, MARGIN + 0.6 * cm,
        PAGE_SIZE[0] - 2 * MARGIN,
        PAGE_SIZE[1] - 2 * MARGIN - 1.0 * cm,
        id="main",
    )

    def _on_page(canvas, doc):
        canvas.saveState()
        w, h = PAGE_SIZE
        # Top orange stripe
        canvas.setFillColor(ORANGE_PRIMARY)
        canvas.rect(0, h - 0.35 * cm, w, 0.35 * cm, fill=1, stroke=0)
        # Bottom footer bar
        canvas.rect(0, 0, w, 0.6 * cm, fill=1, stroke=0)
        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica", 6.5)
        canvas.drawString(MARGIN, 0.18 * cm, _footer_content(date_str))
        # Page number
        canvas.drawRightString(w - MARGIN, 0.18 * cm, f"{doc.page}")
        canvas.restoreState()

    return PageTemplate(id="BDA", frames=[frame], onPage=_on_page)


# ─────────────────────────────────────────────────────────────────────────────
# Main generator
# ─────────────────────────────────────────────────────────────────────────────

class BDAReportPDF:
    """
    Generates an 11-page BDA-branded PDF from the same data dict as
    BDAReportGenerator.  See that class's docstring for the full data schema.
    """

    def __init__(self, data=None):
        self.data = data or {}
        self._date_str = self.data.get("report_date", date.today().strftime("%d.%m.%Y"))

    # ── Public ────────────────────────────────────────────────────────────────

    def build(self, output_path: str = "output/bda_report.pdf") -> str:
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        doc = BaseDocTemplate(
            output_path,
            pagesize=PAGE_SIZE,
            leftMargin=MARGIN, rightMargin=MARGIN,
            topMargin=MARGIN + 0.5 * cm, bottomMargin=MARGIN + 0.6 * cm,
        )
        doc.addPageTemplates([_make_page_template(doc, self._date_str)])

        story: list = []
        story += self._page_cover()
        story += self._page_agenda()
        story += self._page_sumario_executivo()
        story += self._page_liquidez_mn_1()
        story += self._page_liquidez_mn_2()
        story += self._page_liquidez_me()
        story += self._page_mercado_cambial()
        story += self._page_bodiva()
        story += self._page_operacoes_bda()
        story += self._page_market_info_1()
        story += self._page_market_info_2()

        doc.build(story)
        return output_path

    # ── Page builders ─────────────────────────────────────────────────────────

    def _page_cover(self) -> list:
        from reportlab.platypus import PageBreak
        d = self._date_str
        items: list = []

        # Full-width orange panel (simulated via a table with orange background)
        items.append(Spacer(1, 2 * cm))
        items.append(
            Table(
                [[Paragraph("<b><i>Resumo Diário Dos Mercados</i></b>",
                             _style("cv_title", fontSize=32, textColor=WHITE,
                                    fontName="Helvetica-BoldOblique",
                                    alignment=TA_CENTER))]],
                colWidths=[PAGE_SIZE[0] - 2 * MARGIN],
                style=TableStyle([
                    ("BACKGROUND",    (0, 0), (-1, -1), ORANGE_PRIMARY),
                    ("TOPPADDING",    (0, 0), (-1, -1), 30),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 30),
                ]),
            )
        )
        items.append(Spacer(1, 0.6 * cm))
        items.append(
            Table(
                [[Paragraph("<b>DIRECÇÃO FINANCEIRA</b>",
                             _style("cv_sub", fontSize=14, textColor=BLACK,
                                    fontName="Helvetica-Bold")),
                  Paragraph(f"<b><font color='#E8751A'>{d}</font></b>",
                             _style("cv_date", fontSize=14, textColor=ORANGE_PRIMARY,
                                    fontName="Helvetica-Bold", alignment=TA_RIGHT))]],
                colWidths=[(PAGE_SIZE[0] - 2 * MARGIN) * 0.6,
                            (PAGE_SIZE[0] - 2 * MARGIN) * 0.4],
                style=TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]),
            )
        )
        items.append(Spacer(1, 0.4 * cm))
        items.append(HRFlowable(width="100%", thickness=1.5, color=ORANGE_PRIMARY))
        items.append(Spacer(1, 0.3 * cm))
        items.append(Paragraph(
            "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona  |  Luanda – Angola",
            STYLE_SMALL,
        ))
        items.append(PageBreak())
        return items

    def _page_agenda(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("AGENDA", self._date_str)
        agenda = [
            ("1.", "Sumário Executivo"),
            ("2.", "Liquidez (MN)"),
            ("3.", "Liquidez (ME)"),
            ("4.", "Mercado Cambial"),
            ("5.", "Mercado Capitais"),
            ("6.", "Informação De Mercado"),
        ]
        rows = [[
            Paragraph(f"<b><font color='#E8751A'>{num}</font></b>",
                      _style("an", fontSize=18, fontName="Helvetica-Bold")),
            Paragraph(f"<b>{lbl}</b>",
                      _style("al", fontSize=13, fontName="Helvetica-Bold")),
        ] for num, lbl in agenda]

        items.append(
            Table(rows,
                  colWidths=[1.2 * cm, PAGE_SIZE[0] - 2 * MARGIN - 1.2 * cm],
                  style=TableStyle([
                      ("TOPPADDING",    (0, 0), (-1, -1), 8),
                      ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                      ("LEFTPADDING",   (0, 0), (-1, -1), 4),
                      ("LINEBELOW",     (0, 0), (-1, -1), 0.5, ORANGE_LIGHT),
                      ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                  ]))
        )
        items.append(PageBreak())
        return items

    def _page_sumario_executivo(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("Sumário Executivo", self._date_str)

        # Reembolso central highlight
        rc = self.data.get("reembolso_credito", "—")
        items.append(
            Table(
                [[Paragraph(f"<b>{rc}</b>",
                             _style("rc_v", fontSize=22, textColor=WHITE,
                                    fontName="Helvetica-Bold", alignment=TA_CENTER)),
                  Paragraph("<b>Reembolso de Crédito</b>",
                             _style("rc_l", fontSize=9, textColor=WHITE,
                                    fontName="Helvetica-Bold", alignment=TA_CENTER))]],
                colWidths=[(PAGE_SIZE[0] - 2 * MARGIN) * 0.4,
                            (PAGE_SIZE[0] - 2 * MARGIN) * 0.6],
                style=TableStyle([
                    ("BACKGROUND",    (0, 0), (-1, -1), BROWN_KPI),
                    ("TOPPADDING",    (0, 0), (-1, -1), 10),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                    ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                ]),
            )
        )
        items.append(Spacer(1, 0.3 * cm))

        kpis = self.data.get("kpis", [])
        if kpis:
            # Render 4 per row
            for chunk_start in range(0, len(kpis), 4):
                chunk = kpis[chunk_start:chunk_start + 4]
                items.append(_kpi_row(chunk))
                items.append(Spacer(1, 0.2 * cm))

        items.append(PageBreak())
        return items

    def _page_liquidez_mn_1(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("LIQUIDEZ – MOEDA NACIONAL (1/2)", self._date_str)
        days = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])

        # Liquidez MN table
        items += _section_header("Liquidez MN  (Em Milhares)")
        lmn_rows = self.data.get("liquidez_mn_rows", [
            {"label": "Posição Reservas Livres BNA", "values": ["—"] * 5},
            {"label": "Posição DO B. Comerciais",    "values": ["—"] * 5},
            {"label": "Posição DP B. Comerciais",    "values": ["—"] * 5},
            {"label": "Posição OMAs",                "values": ["—"] * 5},
            {"label": "LIQUIDEZ BDA",                "values": ["—"] * 5},
        ])
        items.append(_data_table(
            [""] + days,
            [[r["label"]] + r["values"] for r in lmn_rows],
        ))
        items.append(Spacer(1, 0.2 * cm))

        # Operações Vivas
        items += _section_header("Operações Vivas")
        ops = self.data.get("operacoes_vivas", [])
        op_rows = [
            [op.get("tipo","DP"), op.get("contraparte","—"), op.get("montante","—"),
             op.get("taxa","—"), str(op.get("residual","—")),
             op.get("vencimento","—"), op.get("juro_diario","—")]
            for op in ops
        ] or [["—"] * 7]
        items.append(_data_table(
            ["Tipo", "Contraparte", "Montante", "Taxa", "Residual", "Vencimento", "Juro Diário"],
            op_rows, font_size=7,
        ))
        items.append(Spacer(1, 0.2 * cm))

        # LUIBOR
        items += _section_header("Taxas LUIBOR")
        tenors = ["LUIBOR O/N", "LUIBOR 1M", "LUIBOR 3M", "LUIBOR 6M", "LUIBOR 9M", "LUIBOR 12M"]
        luibor = self.data.get("luibor", {})
        luibor_var = self.data.get("luibor_variation", {})
        lu_rows = [[t, luibor.get(t, "—"), luibor.get(t, "—"), luibor.get(t, "—"), luibor_var.get(t, "—")]
                   for t in tenors]
        items.append(_data_table(
            ["Maturidade", "Anterior (D-2)", "Anterior (D-1)", "Actual (D)", "Var (%)"],
            lu_rows,
        ))
        items.append(PageBreak())
        return items

    def _page_liquidez_mn_2(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("LIQUIDEZ – MOEDA NACIONAL (2/2)", self._date_str)
        days = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])

        items += _section_header("Fluxos de Caixa MN")
        fluxos = self.data.get("fluxos_mn_rows", [
            {"label": "Fluxos de Entradas (Cash in flow)", "values": ["—"] * 5},
            {"label": "Reembolsos de crédito (+)",          "values": ["—"] * 5},
            {"label": "Reembolsos de OMA-O/N + Juros",      "values": ["—"] * 5},
            {"label": "Fluxos de Saídas (Cash out flow)",   "values": ["—"] * 5},
            {"label": "Aplicação em OMA",                   "values": ["—"] * 5},
            {"label": "GAP de Liquidez",                    "values": ["—"] * 5},
        ])
        items.append(_data_table(
            [""] + days,
            [[r["label"]] + r["values"] for r in fluxos],
        ))
        items.append(Spacer(1, 0.3 * cm))

        items += _section_header("P&L Control")
        pl_summary = self.data.get("pl_summary", [
            {"label": "Reembolso de Crédito", "n_ops": "—", "montante": "—"},
            {"label": "Fornecedores",          "n_ops": "—", "montante": "—"},
            {"label": "Desembolso de Crédito", "n_ops": "—", "montante": "—"},
        ])
        items.append(_data_table(
            ["Categoria", "Nº Operações", "Montante"],
            [[r["label"], str(r["n_ops"]), str(r["montante"])] for r in pl_summary],
        ))
        items.append(PageBreak())
        return items

    def _page_liquidez_me(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("LIQUIDEZ – MOEDA ESTRANGEIRA", self._date_str)
        days = self.data.get("liquidez_mn_days", ["D-4", "D-3", "D-2", "D-1", "D"])

        items += _section_header("Liquidez ME  (Em Milhões USD)")
        lme = self.data.get("liquidez_me_rows", [
            {"label": "SALDO D.O Estrangeiros", "values": ["—"] * 5},
            {"label": "DPs ME",                 "values": ["—"] * 5},
            {"label": "COLATERAL CDI",           "values": ["—"] * 5},
            {"label": "LIQUIDEZ BDA",            "values": ["—"] * 5},
        ])
        items.append(_data_table([""] + days, [[r["label"]] + r["values"] for r in lme]))
        items.append(Spacer(1, 0.2 * cm))

        items += _section_header("Fluxos de Caixa ME")
        fluxos_me = self.data.get("fluxos_me_rows", [
            {"label": "Fluxos de entradas (Cash in flow)", "values": ["—"] * 5},
            {"label": "Reembolsos de DP + Juros",          "values": ["—"] * 5},
            {"label": "Fluxos de Saídas (Cash out flow)",  "values": ["—"] * 5},
            {"label": "Aplicação em DP ME",                "values": ["—"] * 5},
            {"label": "GAP de Liquidez",                   "values": ["—"] * 5},
        ])
        items.append(_data_table([""] + days, [[r["label"]] + r["values"] for r in fluxos_me]))
        items.append(PageBreak())
        return items

    def _page_mercado_cambial(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("MERCADO CAMBIAL", self._date_str)

        cambial = self.data.get("cambial", {})

        # KPI summary
        fx_kpis = [
            {"label": "USD/AKZ", "value": cambial.get("usd_akz", "—"), "variation_str": ""},
            {"label": "EUR/AKZ", "value": cambial.get("eur_akz", "—"), "variation_str": ""},
            {"label": "EUR/USD", "value": cambial.get("eur_usd", "—"), "variation_str": ""},
            {"label": "Posição Cambial (Kz)",
             "value": cambial.get("posicao_cambial", "—"), "variation_str": ""},
        ]
        items.append(_kpi_row(fx_kpis))
        items.append(Spacer(1, 0.2 * cm))

        items += _section_header("Cambiais")
        cambial_rows = self.data.get("cambial_rows", [
            {"par": "USD/AKZ", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
            {"par": "EUR/AKZ", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
            {"par": "EUR/USD", "anterior2": "—", "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        items.append(_data_table(
            ["Par", "Anterior (D-1)", "Anterior", "Actual (D)", "(%)"],
            [[r["par"], r.get("anterior2","—"), r.get("anterior","—"),
              r.get("atual","—"), r.get("variacao","—")] for r in cambial_rows],
        ))
        items.append(Spacer(1, 0.2 * cm))

        items += _section_header("Transações do Mercado")
        mercado_rows = self.data.get("mercado_rows", [
            {"label": "T+0", "montante": "—", "min": "—", "max": "—"},
            {"label": "T+1", "montante": "—", "min": "—", "max": "—"},
            {"label": "T+2", "montante": "—", "min": "—", "max": "—"},
        ])
        items.append(_data_table(
            ["Liquidação", "Montante USD", "Mínimo", "Máximo"],
            [[r.get("label","—"), r.get("montante","—"), r.get("min","—"), r.get("max","—")]
             for r in mercado_rows],
        ))
        items.append(PageBreak())
        return items

    def _page_bodiva(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("MERCADO DE CAPITAIS", self._date_str)

        items += _section_header("Segmentado Por Produtos")
        seg_rows = self.data.get("bodiva_segment_rows", [
            {"segmento": "Obrigações De Tesouro",    "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Bilhetes Do Tesouro",       "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Obrigações Privadas",       "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Unidades De Participações", "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Acções",                    "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Repos",                     "anterior": "—", "atual": "—", "variacao": "—"},
            {"segmento": "Total",                     "anterior": "—", "atual": "—", "variacao": "—"},
        ])
        items.append(_data_table(
            ["Segmento", "Anterior", "Actual", "(%)"],
            [[r["segmento"], r["anterior"], r["atual"], r["variacao"]] for r in seg_rows],
        ))
        items.append(Spacer(1, 0.25 * cm))

        items += _section_header("Mercado de Bolsas de Acções")
        stocks = self.data.get("bodiva_stocks", {})
        stk_rows = [
            [code,
             str(info.get("volume",   "—")),
             str(info.get("previous", "—")),
             str(info.get("current",  "—")),
             f"{info['change_pct']:+.2f}%" if isinstance(info.get("change_pct"), (int,float)) else "—",
             str(info.get("cap_bolsista", "—"))]
            for code, info in stocks.items()
        ] or [["Dados não disponíveis (BODIVA)"] + ["—"] * 5]
        items.append(_data_table(
            ["Código", "Vol. Transacc.", "Preço Anterior", "Preço Actual", "Variação", "Cap. Bolsista"],
            stk_rows, font_size=7,
        ))
        items.append(PageBreak())
        return items

    def _page_operacoes_bda(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("MERCADO DE CAPITAIS – OPERAÇÕES BDA", self._date_str)

        items += _section_header("Transações")
        ops = self.data.get("bodiva_operacoes", [])
        op_rows = [
            [r.get("tipo","—"), r.get("data","—"), r.get("cv","—"),
             r.get("preco","—"), r.get("quantidade","—"), r.get("montante","—")]
            for r in ops
        ] or [["—"] * 6]
        items.append(_data_table(
            ["Tipo Operação", "Data Contrat.", "C/V", "Preço", "Quantidades", "Montante"],
            op_rows, font_size=7,
        ))
        items.append(Spacer(1, 0.3 * cm))

        items += _section_header("Carteira De Títulos")
        carteira = self.data.get("carteira_titulos", [])
        ct_rows = [
            [r.get("carteira","—"), r.get("cod","—"), r.get("qty_d1","—"),
             r.get("qty_d","—"),    r.get("nominal","—"), r.get("taxa","—"),
             r.get("montante","—"), r.get("juros_anual","—"), r.get("juro_diario","—")]
            for r in carteira
        ] or [["—"] * 9]
        items.append(_data_table(
            ["Carteira","Cód. Neg.","Qtd D-1","Qtd D","Val. Nominal",
             "Taxa","Montante D","Juros Anual","Juros Diário"],
            ct_rows, font_size=6,
        ))
        items.append(PageBreak())
        return items

    def _page_market_info_1(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("INFORMAÇÃO DE MERCADOS (1/2)", self._date_str)
        market = self.data.get("market_info", {})
        w = PAGE_SIZE[0] - 2 * MARGIN

        items += _section_header("Capital Markets")
        cm_rows = market.get("capital_markets", [
            {"indice": "S&P500",            "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "Dow Jones",          "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "NASDAQ",             "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "NIKKEI 225",         "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "IBOVESPA",           "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "SHANGHAI COMPOSITE", "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "EUROSTOX",           "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "Bolsa de Londres",   "anterior":"—","atual":"—","variacao":"—"},
            {"indice": "PSI 20",             "anterior":"—","atual":"—","variacao":"—"},
        ])

        comment = market.get("cm_commentary", "")
        if comment:
            tbl_w = w * 0.45
            com_w = w * 0.52
            tbl = _data_table(
                ["Índice", "Anterior", "Actual", "(%)"],
                [[r["indice"], r["anterior"], r["atual"], r["variacao"]] for r in cm_rows],
                col_widths=[tbl_w * 0.46, tbl_w * 0.18, tbl_w * 0.18, tbl_w * 0.18],
            )
            com = Table(
                [[Paragraph(comment, STYLE_COMMENT)]],
                colWidths=[com_w],
                style=TableStyle([
                    ("BACKGROUND", (0,0),(-1,-1), ORANGE_LIGHT),
                    ("BOX",        (0,0),(-1,-1), 0.8, ORANGE_PRIMARY),
                    ("TOPPADDING",    (0,0),(-1,-1), 6),
                    ("BOTTOMPADDING", (0,0),(-1,-1), 6),
                ]),
            )
            items.append(Table([[tbl, com]], colWidths=[tbl_w, com_w],
                               style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")])))
        else:
            items.append(_data_table(
                ["Índice", "Anterior", "Actual", "(%)"],
                [[r["indice"], r["anterior"], r["atual"], r["variacao"]] for r in cm_rows],
            ))

        items.append(Spacer(1, 0.25 * cm))

        items += _section_header("Criptomoedas")
        cr_rows = market.get("crypto", [
            {"moeda":"BITCOIN (BTC)","anterior":"—","atual":"—","variacao":"—"},
            {"moeda":"ETHEREUM (ETH)","anterior":"—","atual":"—","variacao":"—"},
            {"moeda":"XRP (XRP)",    "anterior":"—","atual":"—","variacao":"—"},
            {"moeda":"USDC",         "anterior":"—","atual":"—","variacao":"—"},
            {"moeda":"TETHER",       "anterior":"—","atual":"—","variacao":"—"},
        ])
        cr_comment = market.get("crypto_commentary", "")
        if cr_comment:
            tbl_w = w * 0.45
            com_w = w * 0.52
            tbl = _data_table(
                ["Moeda","Anterior","Actual","(%)"],
                [[r["moeda"],r["anterior"],r["atual"],r["variacao"]] for r in cr_rows],
                col_widths=[tbl_w*0.46, tbl_w*0.18, tbl_w*0.18, tbl_w*0.18],
            )
            com = Table([[Paragraph(cr_comment, STYLE_COMMENT)]], colWidths=[com_w],
                        style=TableStyle([("BACKGROUND",(0,0),(-1,-1),ORANGE_LIGHT),
                                          ("BOX",(0,0),(-1,-1),0.8,ORANGE_PRIMARY)]))
            items.append(Table([[tbl, com]], colWidths=[tbl_w, com_w],
                               style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")])))
        else:
            items.append(_data_table(
                ["Moeda","Anterior","Actual","(%)"],
                [[r["moeda"],r["anterior"],r["atual"],r["variacao"]] for r in cr_rows],
            ))

        items.append(PageBreak())
        return items

    def _page_market_info_2(self) -> list:
        from reportlab.platypus import PageBreak
        items: list = _page_title("INFORMAÇÃO DE MERCADOS (2/2)", self._date_str)
        market = self.data.get("market_info", {})
        w = PAGE_SIZE[0] - 2 * MARGIN

        items += _section_header("Commodities")
        cmd_rows = market.get("commodities", [
            {"nome":"PETRÓLEO (BRENT)",       "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"MILHO (USD/BU)",          "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"SOJA (USD/BU)",           "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"TRIGO (USD/LBS)",         "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"CAFÉ (USD/LBS)",          "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"AÇÚCAR (USD/LBS)",        "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"ÓLEO DE PALMA (USD/LBS)", "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"ALGODÃO (USD/LBS)",       "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"BANANA (USD/LBS)",        "anterior":"—","atual":"—","variacao":"—"},
        ])
        cmd_comment = market.get("commodities_commentary","")
        if cmd_comment:
            tbl_w = w * 0.45
            com_w = w * 0.52
            tbl = _data_table(
                ["Commodity","Anterior","Actual","(%)"],
                [[r["nome"],r["anterior"],r["atual"],r["variacao"]] for r in cmd_rows],
                col_widths=[tbl_w*0.46, tbl_w*0.18, tbl_w*0.18, tbl_w*0.18],
            )
            com = Table([[Paragraph(cmd_comment, STYLE_COMMENT)]], colWidths=[com_w],
                        style=TableStyle([("BACKGROUND",(0,0),(-1,-1),ORANGE_LIGHT),
                                          ("BOX",(0,0),(-1,-1),0.8,ORANGE_PRIMARY)]))
            items.append(Table([[tbl, com]], colWidths=[tbl_w, com_w],
                               style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")])))
        else:
            items.append(_data_table(
                ["Commodity","Anterior","Actual","(%)"],
                [[r["nome"],r["anterior"],r["atual"],r["variacao"]] for r in cmd_rows],
            ))
        items.append(Spacer(1, 0.25 * cm))

        items += _section_header("Minerais")
        min_rows = market.get("minerais", [
            {"nome":"OURO",    "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"FERRO",   "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"COBRE",   "anterior":"—","atual":"—","variacao":"—"},
            {"nome":"MANGANÊS","anterior":"—","atual":"—","variacao":"—"},
        ])
        min_comment = market.get("minerais_commentary","")
        if min_comment:
            tbl_w = w * 0.45
            com_w = w * 0.52
            tbl = _data_table(
                ["Mineral","Anterior","Actual","(%)"],
                [[r["nome"],r["anterior"],r["atual"],r["variacao"]] for r in min_rows],
                col_widths=[tbl_w*0.46, tbl_w*0.18, tbl_w*0.18, tbl_w*0.18],
            )
            com = Table([[Paragraph(min_comment, STYLE_COMMENT)]], colWidths=[com_w],
                        style=TableStyle([("BACKGROUND",(0,0),(-1,-1),ORANGE_LIGHT),
                                          ("BOX",(0,0),(-1,-1),0.8,ORANGE_PRIMARY)]))
            items.append(Table([[tbl, com]], colWidths=[tbl_w, com_w],
                               style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")])))
        else:
            items.append(_data_table(
                ["Mineral","Anterior","Actual","(%)"],
                [[r["nome"],r["anterior"],r["atual"],r["variacao"]] for r in min_rows],
            ))

        items.append(PageBreak())
        return items
