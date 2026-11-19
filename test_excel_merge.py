import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_merge import (
    BLACK, HEADERS, LIGHT_BROWN, LIGHT_RED, RED, YELLOW,
    build_summary, interpretation, merge_lis, save_summary,
)


VITEK_HEADERS = ["报告类型", "菌株编号", "卡片", "菌名", "完成时间", "抗菌剂", "MIC", "解释"]
LIS_HEADERS = ["条码号", "样本号", "菌株编号", "检验目的", "菌名", "抗生素", "MIC", "解释", "完成时间"]


class ExcelMergeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.vitek = self.folder / "vitek.xlsx"
        self.lis = self.folder / "lis.xlsx"

    def write_book(self, path, headers, rows):
        book = Workbook()
        book.active.append(headers)
        for row in rows:
            book.active.append(row)
        book.save(path)
        book.close()

    def make_summary(self, laboratory="<= 1", chart="<= 1", explanation="S"):
        base = ["001-1", "CARD", "Species", "2026-10-20 12:00 CST", "Drug A"]
        self.write_book(self.vitek, VITEK_HEADERS, [
            ["实验室报告", *base, laboratory, explanation],
            ["微生物图表报告", *base, chart, explanation],
            ["实验室报告", *base[:4], "Drug B", "8", "R"],
        ])
        return build_summary(self.vitek)

    def lis_row(self, drug="Drug A", mic="<=1", explanation="敏感", date="2026-10-20 09:00:00"):
        return ["barcode", "sample", "001-1", "purpose", "species", drug, mic, explanation, date]

    def test_vitek_mapping_and_report_pairing(self):
        sheet = self.make_summary().active
        self.assertEqual(tuple(cell.value for cell in sheet[1]), HEADERS)
        self.assertEqual(sheet.max_row, 3)
        self.assertEqual([sheet.cell(2, column).value for column in range(1, 10)], [
            "001-1", "CARD", "Species", "2026-10-20 12:00 CST", "Drug A", "<= 1", "S", "<= 1", "S",
        ])

    def test_matching_is_per_drug_and_mic_spacing_is_ignored(self):
        summary = self.make_summary()
        self.write_book(self.lis, LIS_HEADERS, [self.lis_row(), self.lis_row("Drug B", "8", "耐药")])
        result, stats = merge_lis(summary, self.lis)
        self.assertEqual(result.active["J2"].value, "<=1")
        self.assertEqual(result.active["K2"].value, "S")
        self.assertEqual(result.active["K3"].value, "R")
        self.assertEqual(stats.matched_rows, 2)
        self.assertEqual(result.active["F2"].fill.patternType, None)
        self.assertEqual(result.active["D2"].fill.patternType, None)
        self.assertEqual(result.active["F3"].fill.fgColor.rgb, YELLOW)

    def test_mismatch_colors_and_date_prefix(self):
        summary = self.make_summary(chart="2")
        self.write_book(self.lis, LIS_HEADERS, [self.lis_row(mic="4", explanation="耐药", date="2026-10-21")])
        result, stats = merge_lis(summary, self.lis)
        for column in ("F", "H", "J"):
            self.assertEqual(result.active[f"{column}2"].fill.fgColor.rgb, YELLOW)
            self.assertEqual(result.active[f"{column}2"].font.color.rgb, RED)
        for column in ("G", "I", "K"):
            self.assertEqual(result.active[f"{column}2"].fill.fgColor.rgb, YELLOW)
            self.assertEqual(result.active[f"{column}2"].font.color.rgb, BLACK)
        self.assertEqual(result.active["D2"].fill.fgColor.rgb, LIGHT_BROWN)
        self.assertEqual(result["匹配明细"]["F2"].fill.fgColor.rgb, LIGHT_BROWN)
        self.assertEqual(stats.date_differences, 1)

    def test_all_interpretation_mappings(self):
        for source, target in [
            ("耐药", "R"), ("敏感", "S"), ("中间", "I"), ("中介", "I"),
            ("非敏感", "NS"), (None, "NULL"), (" ", "NULL"), ("SDD", "SDD"), ("+", "+"),
        ]:
            with self.subTest(source=source):
                self.assertEqual(interpretation(source), target)

    def test_duplicates_are_preserved_and_unmatched_rows_are_audited(self):
        summary = self.make_summary()
        self.write_book(self.lis, LIS_HEADERS, [
            self.lis_row(), self.lis_row(mic="2"), self.lis_row(drug="Other"),
        ])
        result, stats = merge_lis(summary, self.lis)
        self.assertEqual(result.active.max_row, 4)
        self.assertEqual([result.active[f"J{row}"].value for row in (2, 3)], ["<=1", "2"])
        self.assertEqual(stats.duplicate_keys, 1)
        self.assertEqual(stats.unmatched_lis_rows, 1)
        self.assertEqual(stats.missing_rows, 1)
        self.assertEqual(result["匹配明细"].max_row, 5)
        self.assertEqual(summary.active.max_row, 3)

    def test_new_lis_strains_are_appended_in_source_order(self):
        summary = self.make_summary()
        first = self.lis_row(mic="16", explanation="耐药")
        first[2] = "002-1"
        second = self.lis_row(drug="Drug B", explanation=None)
        second[2] = "002-1"
        third = self.lis_row(drug="Drug C", explanation="SDD")
        third[2] = "003-1"
        self.write_book(self.lis, LIS_HEADERS, [
            self.lis_row(), first, second, third, self.lis_row(drug="Other"),
        ])
        original = self.lis.read_bytes()
        result, stats = merge_lis(summary, self.lis)
        sheet = result.active
        self.assertEqual(sheet.max_row, 6)
        self.assertEqual(stats.appended_lis_rows, 3)
        self.assertEqual(stats.output_rows, 5)
        self.assertEqual(stats.unmatched_lis_rows, 1)
        self.assertEqual(stats.matched_rows, 1)
        self.assertEqual(stats.missing_rows, 1)
        self.assertEqual([sheet.cell(4, column).value for column in range(1, 12)], [
            "002-1", None, "species", None, "Drug A", None, None, None, None, "16", "R",
        ])
        self.assertEqual(sheet["A5"].value, "002-1")
        self.assertEqual(sheet["K5"].value, "NULL")
        self.assertEqual(sheet["A6"].value, "003-1")
        self.assertEqual(sheet["K6"].value, "SDD")
        for row in sheet.iter_rows(min_row=4, max_row=6, max_col=11):
            self.assertTrue(all(cell.fill.fgColor.rgb == LIGHT_RED for cell in row))
        audit = result["匹配明细"]
        appended_audit = [row for row in audit.iter_rows(min_row=2, values_only=True) if row[6] == "菌株编号不存在，已追加"]
        self.assertEqual([row[0] for row in appended_audit], [4, 5, 6])
        self.assertEqual([row[1] for row in appended_audit], [3, 4, 5])
        self.assertEqual(original, self.lis.read_bytes())
        self.assertEqual(summary.active.max_row, 3)
        repeated, repeated_stats = merge_lis(summary, self.lis)
        self.assertEqual(repeated.active.max_row, sheet.max_row)
        self.assertEqual(repeated_stats.appended_lis_rows, 3)

    def test_appended_row_styles_survive_save_reload(self):
        summary = self.make_summary()
        row = self.lis_row(explanation="中介")
        row[2] = "00002-1"
        self.write_book(self.lis, LIS_HEADERS, [row])
        result, _ = merge_lis(summary, self.lis)
        output = self.folder / "result.xlsx"
        save_summary(result, output, [self.vitek, self.lis])
        book = load_workbook(output)
        try:
            self.assertEqual(book.active["A4"].value, "00002-1")
            self.assertEqual(book.active["K4"].value, "I")
            for cell in book.active[4]:
                self.assertEqual(cell.fill.patternType, "solid")
                self.assertEqual(cell.fill.fgColor.rgb, LIGHT_RED)
        finally:
            book.close()

    def test_save_reload_and_source_protection(self):
        summary = self.make_summary()
        original = self.vitek.read_bytes()
        with self.assertRaises(ValueError):
            save_summary(summary, self.vitek, [self.vitek])
        output = self.folder / "summary_result.xlsx"
        save_summary(summary, output, [self.vitek])
        book = load_workbook(output)
        self.assertEqual(book.active["A2"].value, "001-1")
        book.close()
        self.assertEqual(original, self.vitek.read_bytes())

    def test_invalid_filename_and_headers(self):
        with self.assertRaises(ValueError):
            build_summary(self.folder / "wrong.xlsx")
        self.write_book(self.vitek, ["wrong"], [])
        with self.assertRaisesRegex(ValueError, "报告类型"):
            build_summary(self.vitek)

    def test_empty_and_duplicate_vitek_are_rejected(self):
        self.write_book(self.vitek, VITEK_HEADERS, [])
        with self.assertRaisesRegex(ValueError, "没有"):
            build_summary(self.vitek)
        row = ["实验室报告", "001", "CARD", "Species", "2026-10-20", "Drug A", "1", "S"]
        self.write_book(self.vitek, VITEK_HEADERS, [row, row])
        with self.assertRaisesRegex(ValueError, "重复"):
            build_summary(self.vitek)


if __name__ == "__main__":
    unittest.main()