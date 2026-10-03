import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

import report


class ReportStylingTests(unittest.TestCase):
    def test_report_sheet_has_bulletin_layout_and_filterable_price_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "daily.xlsx"
            report.build_excel(
                output_path,
                "2026-10-03",
                {
                    "涨跌幅排行": (
                        [
                            {"品种": "番茄", "涨跌幅(%)": 2.5},
                            {"品种": "白菜", "涨跌幅(%)": -1.2},
                        ],
                        "全国市场价格变化",
                    )
                },
            )

            worksheet = load_workbook(output_path)["涨跌幅排行"]

            self.assertEqual(worksheet.freeze_panes, "A3")
            self.assertEqual(worksheet.auto_filter.ref, "A2:B4")
            self.assertFalse(worksheet.sheet_view.showGridLines)
            self.assertTrue(worksheet["A1"].fill.fgColor.rgb.endswith("182A38"))
            self.assertEqual(worksheet["A1"].font.color.rgb, "00FFFFFF")
            self.assertTrue(worksheet["A2"].fill.fgColor.rgb.endswith("D9A441"))
            self.assertTrue(worksheet["A3"].fill.fgColor.rgb.endswith("FFFDFA"))
            self.assertTrue(worksheet["A4"].fill.fgColor.rgb.endswith("F5F0E5"))
            self.assertEqual(len(worksheet.conditional_formatting), 1)


if __name__ == "__main__":
    unittest.main()
