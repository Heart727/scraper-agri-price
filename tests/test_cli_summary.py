import unittest

from main import format_report_summary


class CliSummaryTests(unittest.TestCase):
    def test_no_detail_summary_numbers_only_sheets_that_exist(self):
        lines = format_report_summary([
            ("涨跌幅排行TOP20", 20, "条"),
            ("分类品种批发价", 0, "条"),
            ("批发价格指数(日)", 90, "天"),
            ("分类价格指数", 4, "条"),
            ("今日报送市场", 158, "家"),
        ])

        self.assertEqual(lines, [
            "1. 涨跌幅排行TOP20 —— 20 条",
            "2. 批发价格指数(日) —— 90 天",
            "3. 分类价格指数 —— 4 条",
            "4. 今日报送市场 —— 158 家",
        ])

    def test_empty_summary_has_no_numbering_gaps(self):
        self.assertEqual(format_report_summary([]), [])
