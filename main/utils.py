import calendar
import datetime
from datetime import datetime
import locale
import re
import time
from queue import Empty, Queue

from docx.text.paragraph import Paragraph


locale.setlocale(locale.LC_ALL, 'ru-Ru')

MONTHS = 'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря'
NUMBERS = '01|02|03|04|05|06|07|08|09|10|11|12'

FORMATS = (
    '%d %B %Y', '%d.%m.%Y', '%d-%m-%Y',
    '%d/%m/%Y', '%Y.%m.%d', '%Y-%m-%d', '%Y/%m/%d'
)

REPL_DICT = {
    '  ': ' ',
    ' - ': ' — ',
    ' )': ')',
    '( ': '(',
    ' , ': ', ',
    'м2': 'м²',
    'м3': 'м³',
}


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
        

def paragraph_replace_text(paragraph: Paragraph, 
                           regex: str, 
                           replace_str: str) -> Paragraph:
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


def get_current_date(date_format: str) -> str:
    return datetime.today().strftime(date_format)


def correct_month(date: str) -> str:
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


def convert_dates(text: str, date_format: str) -> str:
    """Accept and return the following date formats:
    '%d %B %Y'   (24 марта 2023)
    '%d.%m.%Y'   (24.03.2023)
    '%d-%m-%Y'   (24-03-2023)
    '%d/%m/%Y'   (24/03/2023)
    '%Y.%m.%d'   (2023.03.24)
    '%Y-%m-%d'   (2023-03-24)
    '%Y/%m/%d'   (2023/03/24)
    """
    patterns = [
        r'(\d{1,2} (?:' + MONTHS + ') \d{4})',
        r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})'
    ]

    months = dict(zip(MONTHS.split('|'), NUMBERS.split('|')))

    dates = re.findall('|'.join(patterns), text)
    dates = [list(filter(None, d))[0] for d in dates]

    def get_symbol(x): return [i for i in x if i in '/-. '][0]

    for date in dates:
        day, month, year = date.split(get_symbol(date))
        month = months[month] if not any(x in '.-/' for x in date) else month
        new_date = datetime.date(int(year), int(month), int(day))
        new_date = correct_month(new_date.strftime(date_format))

        text = re.sub(date, new_date, text)
    return text