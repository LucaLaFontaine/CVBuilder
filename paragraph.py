import skills
import yaml
from math import ceil

from docx import Document
from docx.shared import Mm, Pt
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

class paragraph:
    def __init__(self, doc, info,  paragraph=None):
        self.doc = doc
        self.info = info

        if paragraph == None:
            self.para = self.doc.add_paragraph()
        else: 
            self.para = paragraph

        

        if isinstance(info['runs'], str):
            info['runs'] = [info['runs']]

        for runText in info['runs']:
            self.addRun(runText)

        self.formatPara()

    def formatPara(self):
        self.para.style = self.info['para_style'] if 'para_style' in self.info else None
        para_format = self.para.paragraph_format
        para_format.space_after = eval(self.info['para_space_after'])
        para_format.line_spacing = self.info['para_line_spacing']
        para_format.alignment = eval(self.info['para_align'])

    def addRun(self, runText):
        run = self.para.add_run(runText)
        # run.style = self.info['para_style'] if 'para_style' in self.info else None
        font = run.font
        
        font.name = self.info['font-name'] if 'font-name' in self.info else None
        font.size = eval(self.info['font-size']) if 'font-size' in self.info else None
        font.color.rgb = eval(self.info['font-colour']) if 'font-colour' in self.info else None
        font.bold = self.info['font-bold'] if 'font-bold' in self.info else None


# if __name__ == '__main__':