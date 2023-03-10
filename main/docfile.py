import win32com.client
import pythoncom
import re
import datetime

from file import File


replace_dict = {
    ' - ': ' — ',
    ' )': ')',
    '( ': '(',
    'м2': 'м²',
    'м3': 'м³'
    }


months = {
    'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4,
    'мая': 5, 'июня': 6, 'июля': 7, 'августа': 8,
    'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
}


class DocFile(File):
    word = win32com.client.Dispatch('Word.Application', pythoncom.CoInitialize())
    word.Visible = False

    def __init__(self, input_path, output_path, file):
        super().__init__()
        self.doc = self.word.Documents.Open(self.input_path + '/' + self.file)
    
    def allowed_file(self, filename):
        super().allowed_file(filename)

    def upload_file(self):
        super().upload_file()

    def show_changes(self):
        self.doc.Activate()
        self.word.ActiveDocument.TrackRevisions = True
        self.doc.ShowRevisions = 0
        
    def add_start_text(self):
        start_text = 'Перевод с английского языка на русский язык\n'

        fline = self.doc.Range(0, 0)
        fline.InsertBefore(start_text)
        fline.Font.Name = 'Times New Roman'
        fline.Font.Size = 11
        fline.Font.Italic = True
        fline.Font.Underline = 2
        fline.Font.Bold = False
        fline.Paragraphs.Alignment = win32com.client.constants.wdAlignParagraphRight

    def replace_text(self):
        wdFindContinue = 1
        wdReplaceAll = 2

        for key, value in replace_dict.items():
            _ = self.word.Selection.Find.Execute(
                FindText=key, 
                MatchCase=False, 
                MatchWholeWord=False,
                MatchWildcards=False, 
                MatchSoundsLike=False,
                MatchAllWordForms=False, 
                Forward=True,
                Wrap=wdFindContinue, 
                Format=False, 
                ReplaceWith=value,
                Replace=wdReplaceAll
            )

    def replace_regex(self):
        old_regex = r"\"(.*?)\""
        new_regex = r"«\1»"

        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(old_regex, current_text):
                paragraph.Range.Text = re.sub(old_regex, new_regex, current_text)

    def edit_dates(self):
        regexp = r'\d{2}[-]\d{2}[-]\d{4}'

        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(regexp, current_text):
                paragraph.Range.Text = re.search(regexp, current_text).group().replace('-', '.')

    def format_dates(self):
        month_re = r"\d{1,2} (?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря) \d{4}"
        date_symbol = '.'
        date_format = '%d.%m.%Y'

        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(month_re, current_text):
                match = re.search(month_re, current_text).group()
                parse_date = datetime.datetime.strptime(self.convert_date(match, date_symbol), date_format)
                new_date = parse_date.strftime(date_format)
                paragraph.Range.Text = re.sub(match, new_date, current_text)

    def convert_date(self, text, date_symbol):
        integers = []
        for m in text.split(' '):
            if m in months:
                integers.append(months[m])
            else:
                integers.append(m)
        
        integers = date_symbol.join(str(x) for x in integers)
        return integers

    def edit_header_footer(self):
        header_primary = self.word.ActiveDocument.Sections(1).Headers(win32com.client.constants.wdHeaderFooterPrimary)
        header_fp = self.word.ActiveDocument.Sections(1).Headers(win32com.client.constants.wdHeaderFooterFirstPage)

        footer_primary = self.word.ActiveDocument.Sections(1).Footers(win32com.client.constants.wdHeaderFooterPrimary)
        footer_fp = self.word.ActiveDocument.Sections(1).Footers(win32com.client.constants.wdHeaderFooterFirstPage)

        def edit_element(element):
            old_regex = r"\"(.*?)\""
            new_regex = r"«\1»"

            for key, value in replace_dict.items():
                element.Range.Text = element.Range.Text.replace(key, value)
            element.Range.Text = re.sub(old_regex, new_regex, element.Range.Text)

        edit_element(header_primary)
        edit_element(header_fp)
        edit_element(footer_primary)
        edit_element(footer_fp)

    def accept_changes(self):
        self.word.ActiveDocument.Revisions.AcceptAll()
        if self.word.ActiveDocument.Comments.Count > 0:
            self.word.ActiveDocument.DeleteAllComments()

    def close_doc(self):
        self.doc.SaveAs(self.output_path + '/' + self.file)
        self.doc.Close()

    def download_file(self):
        super().download_file()