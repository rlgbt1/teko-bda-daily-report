from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd
import os
from datetime import datetime

BANK_COLOR_RGB = (255, 140, 0)
WHITE = (255, 255, 255)
DARK_TEXT = (33, 33, 33)
BANK_ADDRESS = "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona, Luanda - Angola"
REPORT_TITLE = "Resumo Diário dos Mercados"


def _rgb(t):
    return RGBColor(*t)


class ReportBuilder:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.33)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

    def _add_header(self, slide, title: str, date_str: str):
        bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(1.2))
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(BANK_COLOR_RGB)
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.1), Inches(10), Inches(1))
        tf = tb.text_frame
        tf.text = title
        p = tf.paragraphs[0]
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = _rgb(WHITE)

        dtb = slide.shapes.add_textbox(Inches(10.5), Inches(0.3), Inches(2.5), Inches(0.6))
        dtf = dtb.text_frame
        dtf.text = date_str
        dp = dtf.paragraphs[0]
        dp.font.size = Pt(14)
        dp.font.color.rgb = _rgb(WHITE)
        dp.alignment = PP_ALIGN.RIGHT

    def _add_footer(self, slide):
        fb = slide.shapes.add_textbox(Inches(0.3), Inches(7.1), Inches(12), Inches(0.3))
        ff = fb.text_frame
        ff.text = BANK_ADDRESS
        fp = ff.paragraphs[0]
        fp.font.size = Pt(8)
        fp.font.color.rgb = _rgb((150, 150, 150))

    def add_title_slide(self, date_str: str):
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = _rgb(BANK_COLOR_RGB)

        tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(2))
        tf = tb.text_frame
        tf.text = REPORT_TITLE.upper()
        p = tf.paragraphs[0]
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = _rgb(WHITE)
        p.alignment = PP_ALIGN.CENTER

        db = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(11), Inches(1))
        df = db.text_frame
        df.text = f"Direcção Financeira  |  {date_str}"
        dp = df.paragraphs[0]
        dp.font.size = Pt(22)
        dp.font.color.rgb = _rgb(WHITE)
        dp.alignment = PP_ALIGN.CENTER

        tb2 = slide.shapes.add_textbox(Inches(10.5), Inches(6.5), Inches(2.5), Inches(0.7))
        tf2 = tb2.text_frame
        tf2.text = "Powered by TEKO"
        p2 = tf2.paragraphs[0]
        p2.font.size = Pt(10)
        p2.font.color.rgb = _rgb(WHITE)
        p2.alignment = PP_ALIGN.RIGHT

    def add_dataframe_slide(self, title: str, date_str: str, df: pd.DataFrame, ai_summary: str = ""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_header(slide, title, date_str)
        self._add_footer(slide)

        if df is None or df.empty:
            return

        rows, cols = len(df) + 1, len(df.columns)
        table_height = Inches(3.8) if ai_summary else Inches(4.5)

        tbl = slide.shapes.add_table(rows, cols, Inches(0.3), Inches(1.4), Inches(12.7), table_height).table

        for ci, col in enumerate(df.columns):
            cell = tbl.cell(0, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(BANK_COLOR_RGB)
            cell.text = str(col)
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.color.rgb = _rgb(WHITE)
            cell.text_frame.paragraphs[0].font.size = Pt(11)

        for ri, row in df.iterrows():
            for ci, val in enumerate(row):
                cell = tbl.cell(ri + 1, ci)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(10)
                p.font.color.rgb = _rgb(DARK_TEXT)
                try:
                    if "Var" in str(df.columns[ci]) or "%" in str(df.columns[ci]):
                        num = float(str(val).replace("%", "").replace(",", "."))
                        p.font.color.rgb = RGBColor(200, 0, 0) if num < 0 else RGBColor(0, 150, 0)
                except Exception:
                    pass

        if ai_summary:
            sb = slide.shapes.add_textbox(Inches(0.3), Inches(5.4), Inches(12.7), Inches(1.5))
            sf = sb.text_frame
            sf.word_wrap = True
            sf.text = f"💬 {ai_summary}"
            sp = sf.paragraphs[0]
            sp.font.size = Pt(10)
            sp.font.color.rgb = _rgb(DARK_TEXT)
            sp.font.italic = True

    def save(self, path: str):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.prs.save(path)
        return path

