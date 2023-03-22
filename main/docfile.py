import os
from typing import Optional
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


class DocFile(object):
    def __init__(self, 
                 input_path: str, 
                 output_path: str, 
                 file: str) -> None:
        self.input_path = input_path
        self.output_path = output_path
        self.file = file
        self.doc = docx.Document(os.path.abspath(self.input_path + file))

    def add_start_text(self, start_text: Optional[str]) -> None:
        if start_text:
            section = self.doc.sections[0]
            section.different_first_page_header_footer = True
            existing_text = '\n'.join(p.text for p in section.header.paragraphs)
            first_page_header = section.first_page_header
            first_paragraph = first_page_header.paragraphs[0]
            run = first_paragraph.add_run(start_text)
            first_page_header.add_paragraph()
            first_page_header.paragraphs[1].add_run(existing_text)

            paragraph_format = first_paragraph.paragraph_format
            paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            section.header_distance = Pt(24)
            paragraph_format.space_after = Pt(6)
            
            font = run.font
            font.name = 'Times New Roman'
            font.size = Pt(11)
            font.italic = True
            font.underline = True

    def save_file(self) -> None:
        self.doc.save(os.path.abspath(self.output_path + self.file))