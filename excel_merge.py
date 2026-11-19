from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


HEADERS = (
    "菌种编号", "卡片", "菌名", "完成时间", "抗菌剂",
    "实验室报告-MIC", "实验室报告-解释", "微生物图表报告-MIC",
    "微生物图表报告-解释", "LIS 终报告结果", "LIS 终报告解释",
)
YELLOW = "FFFFFF00"
RED = "FFFF0000"
BLACK = "FF000000"
LIGHT_BROWN = "FFE2C4A5"
LIGHT_RED = "FFFFC7CE"
INTERPRETATIONS = {
    "耐药": "R", "敏感": "S", "中间": "I", "中介": "I",
    "非敏感": "NS", "": "NULL", "SDD": "SDD",
}
REPORT_COLUMNS = {"实验室报告": 6, "微生物图表报告": 8}


@dataclass
class MergeStats:
    output_rows: int = 0
    matched_rows: int = 0
    missing_rows: int = 0
    duplicate_keys: int = 0
    unmatched_lis_rows: int = 0
    appended_lis_rows: int = 0
    mic_differences: int = 0
    interpretation_differences: int = 0
    date_differences: int = 0


def text(value):
    return "" if value is None else str(value).strip()


def interpretation(value):
    value = text(value)
    return INTERPRETATIONS.get(value, value)


def validate_path(path, expected_name):
    path = Path(path).expanduser().resolve()
    if path.stem.lower() != expected_name or path.suffix.lower() != ".xlsx":
        raise ValueError(f"请选择名为 {expected_name}.xlsx 的文件。仅支持 .xlsx 格式。")
    if not path.is_file():
        raise ValueError(f"文件不存在：{path}")
    return path


def read_rows(path, expected_headers):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book.active
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows, ())
        for column, expected in expected_headers.items():
            actual = text(headers[column - 1]) if len(headers) >= column else ""
            if actual != expected:
                raise ValueError(
                    f"{Path(path).name} 的 {get_column_letter(column)} 列应为“{expected}”，"
                    f"实际为“{actual}”。"
                )
        return [
            (number, tuple(row))
            for number, row in enumerate(rows, start=2)
            if any(text(value) for value in row)
        ]
    finally:
        book.close()


def format_sheet(sheet, widths):
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 30
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="FF176B5B")
        cell.font = Font(name="Calibri", bold=True, color="FFFFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def append_values(sheet, values):
    cells = []
    for value in values:
        cell = Cell(sheet, value=value)
        if isinstance(cell.value, str):
            cell.data_type = "s"
        cells.append(cell)
    sheet.append(cells)
    return cells


def build_summary(vitek_path):
    vitek_path = validate_path(vitek_path, "vitek")
    rows = read_rows(vitek_path, {
        1: "报告类型", 2: "菌株编号", 3: "卡片", 4: "菌名",
        5: "完成时间", 6: "抗菌剂", 7: "MIC", 8: "解释",
    })
    book = Workbook()
    sheet = book.active
    sheet.title = "summary"
    sheet.append(HEADERS)
    positions = {}
    seen_reports = set()
    for number, row in rows:
        report = text(row[0])
        if report not in REPORT_COLUMNS:
            raise ValueError(f"VITEK 第 {number} 行存在未知报告类型：{report}")
        key = tuple(text(value) for value in row[1:6])
        if not key[0] or not key[4]:
            raise ValueError(f"VITEK 第 {number} 行的菌株编号或抗菌剂为空。")
        if (key, report) in seen_reports:
            raise ValueError(f"VITEK 第 {number} 行存在重复的同类报告，请先核对源数据。")
        seen_reports.add((key, report))
        if key not in positions:
            append_values(sheet, [*row[1:6], None, None, None, None, None, None])
            positions[key] = sheet.max_row
        target = positions[key]
        column = REPORT_COLUMNS[report]
        for offset, value in enumerate(row[6:8]):
            cell = sheet.cell(target, column + offset, value)
            if isinstance(value, str):
                cell.data_type = "s"
    if not positions:
        raise ValueError("VITEK 文件没有可汇总的数据。")
    format_sheet(sheet, [20, 16, 30, 26, 28, 20, 20, 24, 24, 22, 22])
    return book


def comparable(value, mic=False):
    value = text(value)
    if value.upper() == "NULL":
        return ""
    if mic:
        return "".join(value.split()).replace("\u2264", "<=").replace("\u2265", ">=")
    return value


def highlight(sheet, row, columns, color):
    for column in columns:
        cell = sheet.cell(row, column)
        cell.fill = PatternFill("solid", fgColor=YELLOW)
        font = copy(cell.font)
        font.color = color
        cell.font = font


def merge_lis(summary_book, lis_path):
    lis_path = validate_path(lis_path, "lis")
    rows = read_rows(lis_path, {
        3: "菌株编号", 5: "菌名", 6: "抗生素", 7: "MIC", 8: "解释", 9: "完成时间",
    })
    records = defaultdict(list)
    for number, row in rows:
        records[(text(row[2]), text(row[5]))].append((number, row))
    result = Workbook()
    sheet = result.active
    sheet.title = "summary"
    sheet.append(HEADERS)
    audit = result.create_sheet("匹配明细")
    audit.append([
        "summary 行号", "LIS 行号", "菌种编号", "抗菌剂", "VITEK 完成时间",
        "LIS 完成时间", "匹配状态", "LIS 原始解释", "LIS 转换后解释",
    ])
    stats = MergeStats()
    used_lis_rows = set()
    matched_keys = set()
    summary_ids = set()
    for source in summary_book.active.iter_rows(min_row=2, max_col=11, values_only=True):
        summary_ids.add(text(source[0]))
        key = (text(source[0]), text(source[4]))
        matches = records.get(key, [])
        if matches:
            matched_keys.add(key)
        for match in matches or [None]:
            values = list(source[:9]) + [None, None]
            if match is not None:
                lis_number, lis_row = match
                values[9:11] = [lis_row[6], interpretation(lis_row[7])]
                used_lis_rows.add(lis_number)
                stats.matched_rows += 1
            else:
                stats.missing_rows += 1
            append_values(sheet, values)
            target = sheet.max_row
            stats.output_rows += 1
            if len({comparable(values[column - 1], mic=True) for column in (6, 8, 10)}) > 1:
                highlight(sheet, target, (6, 8, 10), RED)
                stats.mic_differences += 1
            if len({comparable(values[column - 1]) for column in (7, 9, 11)}) > 1:
                highlight(sheet, target, (7, 9, 11), BLACK)
                stats.interpretation_differences += 1
            if match is None:
                append_values(audit, [target, None, source[0], source[4], source[3], None, "未找到 LIS", None, None])
                continue
            date_differs = text(lis_row[8])[:10] != text(source[3])[:10]
            if date_differs:
                sheet.cell(target, 4).fill = PatternFill("solid", fgColor=LIGHT_BROWN)
                stats.date_differences += 1
            status = "日期不同" if date_differs else "日期相同"
            if len(matches) > 1:
                status += "；同键多条，全部保留"
            append_values(audit, [
                target, lis_number, source[0], source[4], source[3], lis_row[8], status,
                lis_row[7], interpretation(lis_row[7]),
            ])
            if date_differs:
                for column in (5, 6):
                    audit.cell(audit.max_row, column).fill = PatternFill("solid", fgColor=LIGHT_BROWN)
    for number, row in rows:
        if number in used_lis_rows:
            continue
        if text(row[2]) not in summary_ids:
            cells = append_values(sheet, [
                row[2], None, row[4], None, row[5], None, None, None, None,
                row[6], interpretation(row[7]),
            ])
            for cell in cells:
                cell.fill = PatternFill("solid", fgColor=LIGHT_RED)
            stats.appended_lis_rows += 1
            stats.output_rows += 1
            target = stats.output_rows + 1
            append_values(audit, [
                target, number, row[2], row[5], None, row[8],
                "菌株编号不存在，已追加", row[7], interpretation(row[7]),
            ])
        else:
            stats.unmatched_lis_rows += 1
            append_values(audit, [
                None, number, row[2], row[5], None, row[8],
                "菌株编号存在，抗菌剂未匹配", row[7], interpretation(row[7]),
            ])
    stats.duplicate_keys = sum(len(records[key]) > 1 for key in matched_keys)
    format_sheet(sheet, [20, 16, 30, 26, 28, 20, 20, 24, 24, 22, 22])
    format_sheet(audit, [18, 16, 20, 28, 28, 30, 38, 20, 22])
    return result, stats


def save_summary(book, output_path, source_paths):
    output_path = Path(output_path).expanduser().resolve()
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("结果文件必须使用 .xlsx 扩展名。")
    if output_path in {Path(path).expanduser().resolve() for path in source_paths}:
        raise ValueError("结果文件不能覆盖输入文件，请另选保存位置。")
    book.save(output_path)
    return output_path