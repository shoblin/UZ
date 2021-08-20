import unittest

import sys
from datetime import datetime, date
import UZ.fa as fa

from datetime import datetime,date


class Test_Fa(unittest.TestCase):
    def test_correct_fa_get_date(self):
        answer1, answer2 = fa.get_date('01-01-2021')
        correct_answer = datetime(2021, 1, 1, 0, 0)
        self.assertEqual(answer1, True)
        self.assertEqual(answer2, correct_answer)

    def test_error_fa_get_date_future_date(self):
        answer1, answer2 = fa.get_date('01-01-2022')
        self.assertEqual(answer1, False)

    def test_correct_fa_new_file_name(self):
        test_template = "Template_KOE"
        to_date = date(year=2021, month=8, day=4 )
        correct_answer = "KOE_08_2021.xlsx"
        self.assertEqual(correct_answer, fa.new_file_name(test_template, to_date))


if __name__ == '__main__':
    unittest.main()
