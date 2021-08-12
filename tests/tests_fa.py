import unittest

import sys
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


if __name__ == '__main__':
    unittest.main()
