"""
IB Presentation Generator - Gera apresentações PowerPoint com formatação
de investment banking a partir de dados estruturados em JSON.
"""

import os
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData


class IBPresentationGenerator:
    """Gera apresentações .pptx com estilo investment banking."""

    NAVY = RGBColor(0, 40, 85)
    DARK_BLUE = RGBColor(0, 75, 135)
    ACCENT_BLUE = RGBColor(0, 120, 190)
    DARK_GRAY = RGBColor(55, 55, 55)
    MEDIUM_GRAY = RGBColor(120, 120, 120)
    LIGHT_GRAY = RGBColor(235, 235, 235)
    LIGHTER_GRAY = RGBColor(245, 245, 245)
    WHITE = RGBColor(255, 255, 255)
    GOLD = RGBColor(175, 145, 45)
    GREEN = RGBColor(30, 130, 50)
    RED = RGBColor(190, 30, 30)

    CHART_COLORS = [
        RGBColor(0, 40, 85),
        RGBColor(0, 120, 190),
        RGBColor(175, 145, 45),
        RGBColor(120, 120, 120),
        RGBColor(0, 75, 135),
        RGBColor(55, 55, 55),
    ]

    FONT = "Calibri"

    SLIDE_W = Inches(13.333)
    SLIDE_H = Inches(7.5)

    MARGIN_L = Inches(0.75)
    MARGIN_R = Inches(0.75)
    CONTENT_W = Inches(11.833)  # SLIDE_W - MARGIN_L - MARGIN_R
    TITLE_TOP = Inches(0.4)
    TITLE_H = Inches(0.65)
    SEPARATOR_TOP = Inches(1.1)
    CONTENT_TOP = Inches(1.35)
    CONTENT_H = Inches(5.5)
    FOOTER_TOP = Inches(7.05)

    def __init__(self, template_path: str | None = None):
        if template_path and os.path.exists(template_path):
            self.prs = Presentation(template_path)
        else:
            self.prs = Presentation()
            self.prs.slide_width = self.SLIDE_W
            self.prs.slide_height = self.SLIDE_H
        self._page_num = 0

    def generate(self, data: dict) -> bytes:
        slides = data.get("slides", [])
        if not slides:
            raise ValueError("Nenhum slide fornecido no JSON")

        for slide_data in slides:
            layout_name = slide_data.get("layout", "content")
            handler = getattr(self, f"_add_{layout_name}_slide", None)
            if handler is None:
                handler = self._add_content_slide
            handler(slide_data)

        buf = BytesIO()
        self.prs.save(buf)
        buf.seek(0)
        return buf.getvalue()

    # ------------------------------------------------------------------ #
    #  SLIDE TYPES                                                        #
    # ------------------------------------------------------------------ #

    def _add_cover_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.NAVY)

        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(3.45), self.SLIDE_W, Inches(0.04),
            fill_color=self.GOLD, line_color=self.GOLD,
        )

        title = data.get("title", "")
        self._add_textbox(
            slide, title,
            self.MARGIN_L, Inches(1.8), self.CONTENT_W, Inches(1.4),
            font_size=Pt(38), font_color=self.WHITE, bold=True,
            alignment=PP_ALIGN.LEFT,
        )

        subtitle = data.get("subtitle", "")
        if subtitle:
            self._add_textbox(
                slide, subtitle,
                self.MARGIN_L, Inches(3.65), self.CONTENT_W, Inches(0.8),
                font_size=Pt(20), font_color=RGBColor(180, 200, 220),
                alignment=PP_ALIGN.LEFT,
            )

        date_text = data.get("date", "")
        if date_text:
            self._add_textbox(
                slide, date_text,
                self.MARGIN_L, Inches(6.5), self.CONTENT_W, Inches(0.5),
                font_size=Pt(14), font_color=self.MEDIUM_GRAY,
                alignment=PP_ALIGN.LEFT,
            )

    def _add_content_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.WHITE)
        self._add_title_bar(slide, data.get("title", ""))
        self._add_footer(slide)

        bullets = data.get("bullets", [])
        if not bullets:
            return

        top = self.CONTENT_TOP
        for bullet in bullets:
            tb = slide.shapes.add_textbox(
                self.MARGIN_L + Inches(0.15), top,
                self.CONTENT_W - Inches(0.3), Inches(0.7),
            )
            tf = tb.text_frame
            tf.word_wrap = True

            p = tf.paragraphs[0]
            p.space_before = Pt(4)
            p.space_after = Pt(4)

            run_bullet = p.add_run()
            run_bullet.text = "\u25CF  "
            run_bullet.font.size = Pt(10)
            run_bullet.font.color.rgb = self.ACCENT_BLUE

            run_text = p.add_run()
            run_text.text = str(bullet)
            run_text.font.size = Pt(16)
            run_text.font.name = self.FONT
            run_text.font.color.rgb = self.DARK_GRAY

            top += Inches(0.72)

    def _add_two_columns_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.WHITE)
        self._add_title_bar(slide, data.get("title", ""))
        self._add_footer(slide)

        col_w = Inches(5.6)
        gap = Inches(0.6)

        for i, col_key in enumerate(["left_column", "right_column"]):
            col_data = data.get(col_key, {})
            left = self.MARGIN_L + i * (col_w + gap)

            sub = col_data.get("subtitle", "")
            if sub:
                self._add_textbox(
                    slide, sub,
                    left, self.CONTENT_TOP, col_w, Inches(0.45),
                    font_size=Pt(16), font_color=self.DARK_BLUE, bold=True,
                )
                self._add_shape(
                    slide, MSO_SHAPE.RECTANGLE,
                    left, self.CONTENT_TOP + Inches(0.5), col_w, Inches(0.02),
                    fill_color=self.LIGHT_GRAY,
                )

            bullets = col_data.get("bullets", [])
            top = self.CONTENT_TOP + Inches(0.65)
            for bullet in bullets:
                tb = slide.shapes.add_textbox(left + Inches(0.1), top, col_w - Inches(0.2), Inches(0.6))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.space_before = Pt(3)
                run_b = p.add_run()
                run_b.text = "\u25CF  "
                run_b.font.size = Pt(9)
                run_b.font.color.rgb = self.ACCENT_BLUE
                run_t = p.add_run()
                run_t.text = str(bullet)
                run_t.font.size = Pt(14)
                run_t.font.name = self.FONT
                run_t.font.color.rgb = self.DARK_GRAY
                top += Inches(0.62)

        divider_x = self.MARGIN_L + col_w + gap / 2
        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            divider_x - Inches(0.01), self.CONTENT_TOP, Inches(0.02), Inches(5.2),
            fill_color=self.LIGHT_GRAY,
        )

    def _add_table_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.WHITE)
        self._add_title_bar(slide, data.get("title", ""))
        self._add_footer(slide)

        tbl_data = data.get("table", {})
        headers = tbl_data.get("headers", [])
        rows = tbl_data.get("rows", [])
        if not headers:
            return

        n_rows = len(rows) + 1
        n_cols = len(headers)
        col_w = self.CONTENT_W / n_cols

        tbl_height = min(Inches(0.45) * n_rows, Inches(5.2))
        shape = slide.shapes.add_table(
            n_rows, n_cols,
            self.MARGIN_L, self.CONTENT_TOP + Inches(0.15),
            self.CONTENT_W, tbl_height,
        )
        table = shape.table

        for ci, header in enumerate(headers):
            cell = table.cell(0, ci)
            cell.text = str(header)
            self._style_cell(cell, font_size=Pt(13), bold=True,
                             font_color=self.WHITE, fill_color=self.NAVY)

        for ri, row in enumerate(rows):
            bg = self.LIGHTER_GRAY if ri % 2 == 0 else self.WHITE
            for ci, val in enumerate(row):
                cell = table.cell(ri + 1, ci)
                cell.text = str(val)
                fc = self.DARK_GRAY
                if isinstance(val, str):
                    if val.startswith("+"):
                        fc = self.GREEN
                    elif val.startswith("-") and "%" in val:
                        fc = self.RED
                self._style_cell(cell, font_size=Pt(12), font_color=fc, fill_color=bg)

    def _add_chart_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.WHITE)
        self._add_title_bar(slide, data.get("title", ""))
        self._add_footer(slide)

        chart_type_str = data.get("chart_type", "bar")
        chart_map = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "stacked_bar": XL_CHART_TYPE.COLUMN_STACKED,
        }
        xl_type = chart_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

        cd = data.get("chart_data", {})
        categories = cd.get("categories", [])
        series_list = cd.get("series", [])
        if not categories or not series_list:
            return

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for s in series_list:
            chart_data.add_series(s.get("name", ""), s.get("values", []))

        chart_left = self.MARGIN_L + Inches(0.5)
        chart_top = self.CONTENT_TOP + Inches(0.2)
        chart_w = self.CONTENT_W - Inches(1.0)
        chart_h = Inches(4.8)

        chart_shape = slide.shapes.add_chart(
            xl_type, chart_left, chart_top, chart_w, chart_h, chart_data,
        )
        chart = chart_shape.chart

        for i, s in enumerate(chart.series):
            if i < len(self.CHART_COLORS):
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = self.CHART_COLORS[i]

        chart.has_legend = len(series_list) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(11)
            chart.legend.font.name = self.FONT

        if xl_type != XL_CHART_TYPE.PIE:
            cat_axis = chart.category_axis
            cat_axis.tick_labels.font.size = Pt(11)
            cat_axis.tick_labels.font.name = self.FONT
            cat_axis.tick_labels.font.color.rgb = self.DARK_GRAY
            cat_axis.has_major_gridlines = False

            val_axis = chart.value_axis
            val_axis.tick_labels.font.size = Pt(10)
            val_axis.tick_labels.font.name = self.FONT
            val_axis.tick_labels.font.color.rgb = self.MEDIUM_GRAY
            val_axis.major_gridlines.format.line.color.rgb = self.LIGHT_GRAY

    def _add_key_metrics_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.WHITE)
        self._add_title_bar(slide, data.get("title", ""))
        self._add_footer(slide)

        metrics = data.get("metrics", [])
        if not metrics:
            return

        n = len(metrics)
        max_per_row = min(n, 4)
        gap = Inches(0.35)
        total_gap = gap * (max_per_row - 1)
        card_w = (self.CONTENT_W - total_gap) / max_per_row
        card_h = Inches(2.4)
        start_y = self.CONTENT_TOP + Inches(0.6)

        for i, metric in enumerate(metrics):
            row = i // max_per_row
            col = i % max_per_row
            left = self.MARGIN_L + col * (card_w + gap)
            top = start_y + row * (card_h + gap)

            self._add_shape(
                slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                left, top, card_w, card_h,
                fill_color=self.LIGHTER_GRAY,
                line_color=self.LIGHT_GRAY,
            )

            self._add_shape(
                slide, MSO_SHAPE.RECTANGLE,
                left, top, card_w, Inches(0.06),
                fill_color=self.ACCENT_BLUE,
            )

            label = metric.get("label", "")
            self._add_textbox(
                slide, label,
                left + Inches(0.2), top + Inches(0.3),
                card_w - Inches(0.4), Inches(0.4),
                font_size=Pt(13), font_color=self.MEDIUM_GRAY,
                bold=False, alignment=PP_ALIGN.CENTER,
            )

            value = metric.get("value", "")
            self._add_textbox(
                slide, value,
                left + Inches(0.2), top + Inches(0.8),
                card_w - Inches(0.4), Inches(0.7),
                font_size=Pt(28), font_color=self.NAVY, bold=True,
                alignment=PP_ALIGN.CENTER,
            )

            variation = metric.get("variation", "")
            if variation:
                v_color = self.GREEN if "+" in variation else self.RED if "-" in variation else self.MEDIUM_GRAY
                self._add_textbox(
                    slide, variation,
                    left + Inches(0.2), top + Inches(1.65),
                    card_w - Inches(0.4), Inches(0.4),
                    font_size=Pt(14), font_color=v_color, bold=True,
                    alignment=PP_ALIGN.CENTER,
                )

    def _add_closing_slide(self, data: dict):
        slide = self._new_blank_slide()
        self._set_solid_bg(slide, self.NAVY)

        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(3.45), self.SLIDE_W, Inches(0.04),
            fill_color=self.GOLD, line_color=self.GOLD,
        )

        title = data.get("title", "Obrigado")
        self._add_textbox(
            slide, title,
            self.MARGIN_L, Inches(2.2), self.CONTENT_W, Inches(1.2),
            font_size=Pt(42), font_color=self.WHITE, bold=True,
            alignment=PP_ALIGN.CENTER,
        )

        subtitle = data.get("subtitle", "")
        if subtitle:
            self._add_textbox(
                slide, subtitle,
                self.MARGIN_L, Inches(3.8), self.CONTENT_W, Inches(0.7),
                font_size=Pt(18), font_color=RGBColor(160, 180, 200),
                alignment=PP_ALIGN.CENTER,
            )

    # ------------------------------------------------------------------ #
    #  HELPERS                                                            #
    # ------------------------------------------------------------------ #

    def _new_blank_slide(self):
        layout = self.prs.slide_layouts[6]  # blank layout
        slide = self.prs.slides.add_slide(layout)
        self._page_num += 1
        return slide

    def _set_solid_bg(self, slide, color: RGBColor):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    def _add_title_bar(self, slide, title: str):
        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0), self.SLIDE_W, Inches(1.0),
            fill_color=self.NAVY,
        )

        self._add_textbox(
            slide, title,
            self.MARGIN_L, Inches(0.18), self.CONTENT_W, Inches(0.65),
            font_size=Pt(24), font_color=self.WHITE, bold=True,
            alignment=PP_ALIGN.LEFT,
        )

        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(1.0), self.SLIDE_W, Inches(0.04),
            fill_color=self.GOLD, line_color=self.GOLD,
        )

    def _add_footer(self, slide):
        self._add_shape(
            slide, MSO_SHAPE.RECTANGLE,
            Inches(0), self.FOOTER_TOP, self.SLIDE_W, Inches(0.015),
            fill_color=self.LIGHT_GRAY,
        )

        self._add_textbox(
            slide, str(self._page_num),
            self.SLIDE_W - Inches(1.2), self.FOOTER_TOP + Inches(0.05),
            Inches(0.6), Inches(0.35),
            font_size=Pt(10), font_color=self.MEDIUM_GRAY,
            alignment=PP_ALIGN.RIGHT,
        )

    def _add_textbox(self, slide, text, left, top, width, height,
                     font_size=Pt(14), font_color=None, bold=False,
                     alignment=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None

        p = tf.paragraphs[0]
        p.text = text
        p.alignment = alignment
        p.font.size = font_size
        p.font.name = self.FONT
        p.font.bold = bold
        if font_color:
            p.font.color.rgb = font_color
        return tb

    def _add_shape(self, slide, shape_type, left, top, width, height,
                   fill_color=None, line_color=None):
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill_color
        else:
            shape.fill.background()
        if line_color:
            shape.line.color.rgb = line_color
        else:
            shape.line.fill.background()
        return shape

    def _style_cell(self, cell, font_size=Pt(12), bold=False,
                    font_color=None, fill_color=None):
        if fill_color:
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill_color

        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)

        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = font_size
            paragraph.font.name = self.FONT
            paragraph.font.bold = bold
            if font_color:
                paragraph.font.color.rgb = font_color
