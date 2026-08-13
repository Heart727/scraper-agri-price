# -*- coding: utf-8 -*-
"""report.py —— Excel 报表生成层

把解析好的数据写成一份带格式的 Excel 日报：
  - 多工作表：涨跌幅排行 / 分类品种批发价 / 价格指数 / 分类指数 / 报送市场
  - 表头加粗、带底色，列宽自动调整，前两行冻结（方便滚动查看）
  - 文件名带数据日期，例如：农产品批发价格日报_2026-08-13.xlsx

用到的库：pandas（把数据组织成表格）+ openpyxl（写 Excel 并加格式）
"""
import pandas as pd

import config


def build_excel(output_path, report_date, sheets):
    """把多个工作表的数据一次性写进一个 Excel 文件。

    参数：
      output_path：输出文件路径（由 main.py 拼接好）
      report_date：数据日期字符串（如 "2026-08-13"），写进标题行
      sheets：{工作表名: (列表数据, 备注)} 的有序字典，
              列表数据会被 pandas 转成表格，备注写在表头上一行。
    """
    # 创建一个 Excel 写入器，engine="openpyxl" 表示用 openpyxl 引擎（支持格式）
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # 逐个工作表写入
        for index, (sheet_name, (rows, note)) in enumerate(sheets.items()):
            # 列表转成 DataFrame（pandas 的"表格"对象）
            frame = pd.DataFrame(rows)
            # startrow=1：第 1 行空出来写"数据日期"标题
            frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            _format_sheet(writer, sheet_name, frame, report_date, note)

        # 如果没有任何工作表（数据全为空），写一个提示页，保证文件不是空文件
        if not sheets:
            pd.DataFrame({"提示": ["本次没有抓取到任何数据"]}).to_excel(
                writer, sheet_name="空报表", index=False
            )

    print(f"✅ 报表已生成：{output_path}")


def _format_sheet(writer, sheet_name, frame, report_date, note):
    """给某个工作表加格式：标题行、表头样式、列宽、冻结窗格。"""
    worksheet = writer.sheets[sheet_name]

    # ---- 第 1 行：报表标题（数据日期 + 备注说明）----
    title = f"数据日期：{report_date}"
    if note:
        title += f"　｜　{note}"
    worksheet.cell(row=1, column=1, value=title)
    # 标题加粗、放大一点、给个浅色底
    title_cell = worksheet.cell(row=1, column=1)
    title_cell.font = _font(bold=True, size=11, color="1F4E79")
    # 把标题行占用的宽度合并起来（合并到最后一列），视觉上更像报表标题
    max_col = max(len(frame.columns), 1)
    if max_col > 1:
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)

    # ---- 第 2 行：表头（加粗 + 深色底 + 白字）----
    for col_index, column_name in enumerate(frame.columns, start=1):
        cell = worksheet.cell(row=2, column=col_index)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = _fill("2E75B6")  # 深蓝色底
        cell.alignment = _center()

    # ---- 数据行：数字列统一保留两位小数，让表格更整齐 ----
    for row_index in range(3, 3 + len(frame)):
        for col_index, column_name in enumerate(frame.columns, start=1):
            cell = worksheet.cell(row=row_index, column=col_index)
            if "元/公斤" in column_name or "(%)" in column_name or "指数" in column_name:
                # 数字列：设置两位小数的显示格式
                cell.number_format = "0.00"

    # ---- 列宽：按"表头长度"和"该列内容最大长度"自动调整 ----
    for col_index, column_name in enumerate(frame.columns, start=1):
        # pandas 从第 3 行开始放数据（前面空了两行），逐行找最长的字符串
        max_length = len(str(column_name))
        for row_index in range(3, 3 + len(frame)):
            value = worksheet.cell(row=row_index, column=col_index).value
            if value is not None:
                # 中文字符占两个显示宽度，所以长度按 2 倍算
                text = str(value)
                length = sum(2 if ord(ch) > 127 else 1 for ch in text)
                max_length = max(max_length, length)
        # 宽度 = 最长内容 + 留白，同时设一个上限防止某列过宽
        # 注：列字母用 get_column_letter 换算（第 1 行是合并单元格，不能直接读它）
        from openpyxl.utils import get_column_letter

        worksheet.column_dimensions[get_column_letter(col_index)].width = min(max_length + 2, 60)

    # ---- 冻结窗格：滚动数据时，标题和表头始终可见 ----
    worksheet.freeze_panes = "A3"


def _font(bold=False, size=10, color="000000"):
    """生成 openpyxl 的字体对象（避免在代码里反复写一长串）。"""
    from openpyxl.styles import Font

    return Font(name="微软雅黑", size=size, bold=bold, color=color)


def _fill(color):
    """生成 openpyxl 的填充对象（给单元格上底色）。"""
    from openpyxl.styles import PatternFill

    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _center():
    """生成居中对齐样式。"""
    from openpyxl.styles import Alignment

    return Alignment(horizontal="center", vertical="center")
