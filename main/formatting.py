import win32com.client
import re
import datetime

from helpers import replace_dict, months


class Formatting(object):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    def __init__(self, input_path, output_path, doc_name, old_regex, new_regex):
        self.input_path = input_path
        self.output_path = output_path
        self.doc_name = doc_name
        self.doc = self.word.Documents.Open(input_path + doc_name)
        self.old_regex = old_regex
        self.new_regex = new_regex

    def show_changes(self):
        self.doc.Activate()
        self.word.ActiveDocument.TrackRevisions = True
        self.doc.ShowRevisions = 0
        
    def add_start_text(self, start_text):
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
        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(self.old_regex, current_text):
                paragraph.Range.Text = re.sub(self.old_regex, self.new_regex, current_text)

    def edit_dates(self, regexp, old_symbol, new_symbol):
        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(regexp, current_text):
                paragraph.Range.Text = re.search(regexp, current_text).group().replace(old_symbol, new_symbol)

    def format_dates(self, month_re, date_format, date_symbol):
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
            for key, value in replace_dict.items():
                element.Range.Text = element.Range.Text.replace(key, value)
            element.Range.Text = re.sub(self.old_regex, self.new_regex, element.Range.Text)

        edit_element(header_primary)
        edit_element(header_fp)
        edit_element(footer_primary)
        edit_element(footer_fp)

    def accept_changes(self):
        self.word.ActiveDocument.Revisions.AcceptAll()
        if self.word.ActiveDocument.Comments.Count > 0:
            self.word.ActiveDocument.DeleteAllComments()

    def close_doc(self):
        self.doc.SaveAs(self.output_path + self.doc_name)
        self.doc.Close()


def edit_docs(doc_name):
    start_text_en_ru = 'Перевод с английского языка на русский язык\n'
    start_text_fr_ru = 'Перевод с французского языка на русский язык\n'
    start_text_other = 'Перевод с английского и итальянского языков на русский язык\n'

    month_re = r"\d{1,2} (?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря) \d{4}"
    date_symbol = '.'
    date_format = '%d.%m.%Y'
    date_with_dash = r'\d{2}[-]\d{2}[-]\d{4}'
    date_with_slash = r'\d{2}[/]\d{2}[/]\d{4}'

    d = Formatting(doc_name=doc_name,
                    old_regex=r"\"(.*?)\"",
                    new_regex=r"«\1»")
    d.show_changes()
    d.add_start_text(start_text_en_ru)
    d.replace_text()
    d.replace_regex()
    d.edit_dates(regexp=date_with_dash, old_symbol='-', new_symbol=date_symbol)
    d.format_dates(month_re, date_format, date_symbol)
    d.edit_header_footer()
    print(d.doc_name)
    d.close_doc()