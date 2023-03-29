import pytest
from main.utils import convert_dates, get_locale


class TestDates:
    MONTHS, _ = get_locale(language='English')

    @pytest.mark.parametrize('pattern, dates, expected_result', [
        (
            '%d %B %Y', 
            ('18.09.1987', '18-09-1987', '18/09/1987', '1987/09/18', '1987.09.18', '1987-09-18'), 
            '18 September 1987'
        ),
        (
            '%d.%m.%Y', 
            ('18-09-1987', '18/09/1987', '18 September 1987', '1987/09/18', '1987.09.18', '1987-09-18'), 
            '18.09.1987'
        ),
        (
            '%d-%m-%Y', 
            ('18.09.1987', '18/09/1987', '18 September 1987', '1987/09/18', '1987.09.18', '1987-09-18'), 
            '18-09-1987'
        ),
        (
            '%d/%m/%Y', 
            ('18.09.1987', '18-09-1987', '18 September 1987', '1987/09/18', '1987.09.18', '1987-09-18'), 
            '18/09/1987'
        ),
        (
            '%Y.%m.%d', 
            ('18.09.1987', '18-09-1987', '18/09/1987', '18 September 1987'), 
            '1987.09.18'
        ),
        (
            '%Y-%m-%d', 
            ('18.09.1987', '18-09-1987', '18/09/1987', '18 September 1987'), 
            '1987-09-18'
        ),
        (
            '%Y/%m/%d', 
            ('18.09.1987', '18-09-1987', '18/09/1987', '18 September 1987'), 
            '1987/09/18'
        ),
    ])

    def test_convert_dates(self, pattern, dates, expected_result):
        for date in dates:
            assert convert_dates(date, self.MONTHS, pattern) == expected_result