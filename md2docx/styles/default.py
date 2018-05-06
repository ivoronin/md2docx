"""
Default style
"""
from contextlib import contextmanager
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

value_of = contextmanager(lambda v: (yield v))

class Style:
    """Default style class"""
    @classmethod
    def apply(cls, doc):
        """Apply style do document"""
        with value_of(doc.sections[0]) as section:
            # Page size A4
            section.page_height = Mm(297)
            section.page_width = Mm(210)

            # Margins
            section.top_margin = Pt(60)
            section.bottom_margin = Pt(60)
            section.left_margin = Pt(60)
            section.right_margin = Pt(60)

        # Normal
        with value_of(doc.styles['Normal']) as normal:
            normal.font.name = 'Calibri Light'
            normal.font.size = Pt(10)
            normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            normal.paragraph_format.space_after = Pt(10)

        # Quote
        with value_of(doc.styles['Quote']) as quote:
            quote.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for heading_level in range(1, 4): # 1..3
            with value_of(doc.styles[f'Heading {heading_level}']) as heading:
                heading.font.color.rgb = None
                heading.font.name = 'Calibri'
                heading.font.size = Pt(10 + 2 * (3 - heading_level))
                heading.font.bold = True
                heading.font.small_caps = True
                heading.paragraph_format.space_before = None
