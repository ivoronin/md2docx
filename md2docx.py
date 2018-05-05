import argparse
import io
import sys

from docx import Document
from docx.shared import Inches
import mistune

class DocXMarkdown(mistune.Markdown):
    def parse(self, text):
        s = io.BytesIO()
        super().parse(text)
        self.renderer.doc.save(s)
        return s.getvalue()
        

class DocXRenderer(mistune.Renderer):
    def __init__(self, *largs, **kargs):
        self.doc = Document()
        self.clist = None
        super().__init__(*largs, **kargs)

    def list_item(self, text):
        return f"{text}\n"
      
    def list(self, body, ordered=True):
        if ordered:
            style = "ListNumber3"
        else:
            style = "ListBullet"
        for i in body.rstrip().split("\n"):
            self.doc.add_paragraph(i, style)
        return ''
        
    def header(self, text, level, raw=None):
        self.doc.add_heading(text, level)
        return ''

    def paragraph(self, text):
        self.doc.add_paragraph(text, style=None)
        return ''

def parse_args(args):
    parser = argparse.ArgumentParser(description="Converts a markdown document into docx")
    parser.add_argument("input", help="Markdown input file")
    parser.add_argument("output", help="docx output file (will overwrite if it already exists)")
    args  = parser.parse_args()
    return args
                            

def main():
    renderer = DocXRenderer()
    markdown = DocXMarkdown(renderer = renderer)
    args = parse_args(sys.argv[1:])
    with open(args.input) as f:
        with open(args.output, "wb") as g:
            g.write(markdown(f.read()))


if __name__ == '__main__':
    main()
    
