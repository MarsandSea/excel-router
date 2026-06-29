"""
生成可复现的演示样本（脱敏、通用），用于试跑工具和 README 演示。

运行：python examples/make_sample.py
会在 examples/ 下生成两个样本文件，两者都含「销售部」——
拿它们试跑「按『部门』拆分 + 跨文件合并」，能直观看到两个表的销售部被合并到同一个文件。

Copyright (c) 2026 Abelin
MIT License
"""
import os

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
HEADERS = ["工号", "姓名", "部门", "城市", "本月业绩"]


def _make(path, title, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"

    # 第1行：大标题（合并单元格，测试「表头不在第一行」也能识别）
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = title
    t.font = Font(size=14, bold=True)
    t.alignment = Alignment(horizontal="center")

    # 第2行：表头（带底色）
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDE", (10, 10, 12, 10, 12)):
        ws.column_dimensions[col].width = w

    # 数据行
    for i, r in enumerate(rows, start=3):
        for c, v in enumerate(r, 1):
            ws.cell(row=i, column=c, value=v)

    wb.save(path)
    print("已生成:", path)


def main():
    _make(os.path.join(HERE, "sample_团队A.xlsx"), "A 组业绩明细", [
        ["A001", "张三", "销售部", "北京", 12000],
        ["A002", "李四", "销售部", "北京", 9800],
        ["A003", "王五", "销售部", "上海", 15300],
        ["A004", "赵六", "技术部", "北京", 0],
        ["A005", "孙七", "技术部", "深圳", 0],
        ["A006", "周八", "市场部", "广州", 7600],
        ["", "合计", "", "", 44700],          # 合计行：部门为空，会被自动跳过
    ])
    _make(os.path.join(HERE, "sample_团队B.xlsx"), "B 组业绩明细", [
        ["B001", "吴九", "销售部", "广州", 11200],   # 与 A 表同为「销售部」-> 合并演示
        ["B002", "郑十", "销售部", "深圳", 8800],
        ["B003", "钱多", "市场部", "上海", 9100],
        ["B004", "孙少", "技术部", "杭州", 0],
    ])
    print("\n试跑建议：输入文件夹选 examples/，主拆分列选「部门」，开始处理。")
    print("打开结果里的「销售部.xlsx」，应能看到 A、B 两表的销售部 5 人被合并在一起。")


if __name__ == "__main__":
    main()
