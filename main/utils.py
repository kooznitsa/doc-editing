import calendar
import datetime
import locale
from queue import Empty, Queue
import re
import time
from typing import Optional

from docx.text.paragraph import Paragraph # type: ignore


FORMATS = (
    '%d %B %Y', '%d.%m.%Y', '%d-%m-%Y',
    '%d/%m/%Y', '%Y.%m.%d', '%Y-%m-%d', '%Y/%m/%d',
)

REPL_DICT = {
    '  ': ' ',
    ' - ': ' — ',
    '- ': '— ',
    ' )': ')',
    '( ': '(',
    ' , ': ', ',
    'м2': 'м²',
    'м3': 'м³',
    'm2': 'm²',
    'm3': 'm³',
}


def get_locale(language: str):
    if language == 'Russian':
        locale.setlocale(locale.LC_ALL, 'ru-Ru')
        MONTHS = 'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря'
        new_quotes = r'«\1»'

    elif language == 'English':
        locale.setlocale(locale.LC_ALL, 'en-US')
        MONTHS = '|'.join(list(calendar.month_name)[1:])
        new_quotes = r'“\1”'

    return MONTHS, new_quotes


def countdown(seconds: int) -> None:
    while seconds >= 0:
        m, s = divmod(seconds, 60)
        timer = f'{m:02d}:{s:02d}'
        print('Time until file gets deleted:', timer, end='\r')
        time.sleep(1)
        seconds -= 1


def execute_queue(queue: Queue) -> None:
    while True:
        try:
            q = queue.get()
            q()
        except Empty:
            pass
        

def paragraph_replace_text(
    paragraph: Paragraph, 
    regex: re.Pattern, 
    replace_str: str
) -> Paragraph:
    while True:
        text = paragraph.text
        match = regex.search(text)
        if not match:
            break

        runs = iter(paragraph.runs)
        start, end = match.start(), match.end()

        for run in runs:
            run_len = len(run.text)
            if start < run_len:
                break
            start, end = start - run_len, end - run_len

        run_text = run.text
        run_len = len(run_text)
        run.text = '%s%s%s' % (run_text[:start], replace_str, run_text[end:])
        end -= run_len

        for run in runs:
            if end <= 0:
                break
            run_text = run.text
            run_len = len(run_text)
            run.text = run_text[end:]
            end -= run_len

    return paragraph


def correct_month(date: str, MONTHS: str) -> str:
    """Correct months:
    Март —> марта
    """
    wrong_months = list(calendar.month_name)[1:]
    correct_months = dict(zip(wrong_months, MONTHS.split('|')))
    try:
        wrong_month = date.split()[1]
        date = date.replace(wrong_month, correct_months[wrong_month])
        return date
    except:
        return date
    

def get_current_date(date_format: str) -> str:
    MONTHS, _ = get_locale(language='English')
    return correct_month(datetime.datetime.today().strftime(date_format), MONTHS)


def convert_dates(
    text: str, 
    MONTHS: str, 
    date_format: Optional[str] = None
) -> Optional[str]:
    """Accept and return the following date formats:
    '%d %B %Y'   (24 марта 2023)
    '%d.%m.%Y'   (24.03.2023)
    '%d-%m-%Y'   (24-03-2023)
    '%d/%m/%Y'   (24/03/2023)
    '%Y.%m.%d'   (2023.03.24)
    '%Y-%m-%d'   (2023-03-24)
    '%Y/%m/%d'   (2023/03/24)
    """
    if date_format:
        patterns = [
            r'(\d{1,2} (?:' + MONTHS + ') \d{4})',
            r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})',
            r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})',
        ]

        NUMBERS = '01|02|03|04|05|06|07|08|09|10|11|12'
        months = dict(zip(MONTHS.split('|'), NUMBERS.split('|')))

        dates = re.findall('|'.join(patterns), text)
        dates = [list(filter(None, d))[0] for d in dates]

        get_symbol = lambda x: [i for i in x if i in '/-. '][0]

        for date in dates:
            if re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})', date):
                year, month, day = date.split(get_symbol(date))
            else:
                day, month, year = date.split(get_symbol(date))
            month = months[month] if not any(x in '.-/' for x in date) else month
            new_date = datetime.date(int(year), int(month), int(day))
            new_date = correct_month(new_date.strftime(date_format), MONTHS)

            text = re.sub(date, new_date, text)
    return text