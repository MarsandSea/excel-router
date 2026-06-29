"""core/splitter.py 的集成测试：表头识别、列枚举、跨文件合并、二级拆分、格式保留。"""
import os

import openpyxl
from openpyxl.styles import Font, PatternFill
import pytest

from core.splitter import (
    detect_header_row_auto, find_header_row, resolve_header_row,
    list_columns, list_values, run_split,
)


# ---------- 样本构造 ----------

HEADERS = ["工号", "姓名", "部门", "城市", "金额"]


def _make_book(path, rows, title="月度报表"):
    """表头在第2行、第1行是大标题（用于测自动表头识别）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = title
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for i, r in enumerate(rows, start=3):
        for c, v in enumerate(r, 1):
            ws.cell(row=i, column=c, value=v)
    wb.save(path)


def _count_rows(path, start=3):
    wb = openpyxl.load_workbook(path)
    total = 0
    for ws in wb.worksheets:
        for r in range(start, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                total += 1
    return total


@pytest.fixture
def two_books(tmp_path):
    inp = tmp_path / "in"
    out = tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    _make_book(inp / "A.xlsx", [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "销售部", "北京", 200],
        ["003", "王五", "销售部", "上海", 150],
        ["004", "赵六", "技术部", "北京", 300],
        ["", "合计", "", "", 750],
    ])
    _make_book(inp / "B.xlsx", [
        ["005", "钱七", "销售部", "广州", 120],
        ["006", "孙八", "技术部", "深圳", 220],
    ])
    return str(inp), str(out)


def _base_cfg(inp, out, **over):
    cfg = {
        "input_path": inp, "output_path": out,
        "header_mode": "auto", "header_row": 1,
        "grid_keys": [], "id_keys": [],
        "split_column": "部门", "selected_values": [],
        "person_column": "", "to_person": False, "person_file_filter": [],
        "value_alias_map": {},
        "skip_values": ["合计", "小计", "总计", ""],
        "merge_across_files": True, "make_zip": True, "exact_match": True,
        "preserve_format": True, "auto_open_output": False,
    }
    cfg.update(over)
    return cfg


# ---------- 表头识别 ----------

def test_detect_header_row_auto():
    rows = [["月度报表", None, None], ["工号", "姓名", "部门"], ["001", "张三", "销售部"]]
    assert detect_header_row_auto(rows) == 2


def test_resolve_header_modes():
    rows = [["标题"], ["网格", "工号", "姓名"], ["A", "001", "张三"]]
    assert resolve_header_row(rows, {"header_mode": "auto"}) == 2
    assert resolve_header_row(rows, {"header_mode": "row", "header_row": 2}) == 2
    assert resolve_header_row(
        rows, {"header_mode": "keyword", "grid_keys": ["网格"], "id_keys": ["工号"]}) == 2
    assert find_header_row(rows, ["不存在"], ["工号"]) == -1


# ---------- 列名 / 取值枚举 ----------

def test_list_columns_and_values(two_books):
    inp, _ = two_books
    cfg = {"header_mode": "auto"}
    a = os.path.join(inp, "A.xlsx")
    assert list_columns(a, cfg) == HEADERS
    assert list_values(a, cfg, "部门") == ["技术部", "销售部"]   # 合计行被跳过


# ---------- 跨文件合并（核心）----------

def test_merge_across_files(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out), log_fn=lambda m: None)
    sales = os.path.join(res, "销售部.xlsx")
    tech = os.path.join(res, "技术部.xlsx")
    assert os.path.exists(sales) and os.path.exists(tech)
    # A 的 3 行销售部 + B 的 1 行销售部 = 4（证明合并而非覆盖）
    assert _count_rows(sales) == 4
    assert _count_rows(tech) == 2


def test_no_merge_keeps_separate(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, merge_across_files=False), log_fn=lambda m: None)
    # 不合并：销售部/汇总 下按源文件分文件
    folder = os.path.join(res, "销售部", "汇总")
    assert os.path.isdir(folder)
    files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
    assert len(files) == 2   # A、B 各一个
    assert os.path.exists(os.path.join(res, "销售部.zip"))   # 每主取值打包


def test_single_file_input(two_books):
    inp, _ = two_books
    # 输入单个文件 -> 扁平输出，无文件夹、无 ZIP
    one = os.path.join(inp, "A.xlsx")
    out = os.path.join(os.path.dirname(inp), "single_out")
    os.makedirs(out, exist_ok=True)
    res = run_split(_base_cfg(one, out), log_fn=lambda m: None)
    assert _count_rows(os.path.join(res, "销售部.xlsx")) == 3   # A 表销售部 3 行
    assert _count_rows(os.path.join(res, "技术部.xlsx")) == 1
    assert not os.path.isdir(os.path.join(res, "销售部"))
    assert not os.path.exists(os.path.join(res, "销售部.zip"))


# ---------- 格式保留 ----------

def test_header_format_preserved(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out), log_fn=lambda m: None)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    ws = wb.active
    assert ws["A2"].value == "工号"
    assert ws["A2"].font.bold is True


# ---------- 到人（双产出 + 范围过滤）----------

def test_to_person_all(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, to_person=True, person_column="姓名",
                              merge_across_files=False), log_fn=lambda m: None)
    # 汇总照常产出
    assert os.path.isdir(os.path.join(res, "销售部", "汇总"))
    # 到人：销售部里 A 的张三/李四/王五 + B 的钱七 都各一个
    for name in ("张三", "李四", "王五", "钱七"):
        assert os.path.exists(os.path.join(res, "销售部", "到人", f"{name}.xlsx")), name
    # 技术部到人：A 赵六 + B 孙八
    assert os.path.exists(os.path.join(res, "技术部", "到人", "赵六.xlsx"))
    assert os.path.exists(os.path.join(res, "技术部", "到人", "孙八.xlsx"))


def test_to_person_scope_filter(two_books):
    inp, out = two_books
    # 仅文件名含 "A" 的表参与到人
    res = run_split(_base_cfg(inp, out, to_person=True, person_column="姓名",
                              person_file_filter=["A"], merge_across_files=False),
                    log_fn=lambda m: None)
    assert os.path.exists(os.path.join(res, "销售部", "到人", "张三.xlsx"))   # A 表
    assert not os.path.exists(os.path.join(res, "销售部", "到人", "钱七.xlsx"))  # B 表，不在范围
    # 但 B 表仍进了汇总
    assert os.path.exists(os.path.join(res, "销售部", "汇总", "B.xlsx"))


# ---------- 只拆部分取值 ----------

def test_selected_values_subset(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, selected_values=["销售部"]), log_fn=lambda m: None)
    assert os.path.exists(os.path.join(res, "销售部.xlsx"))
    assert not os.path.exists(os.path.join(res, "技术部.xlsx"))


# ---------- 取值归并映射 ----------

def test_value_alias_map(tmp_path):
    inp = tmp_path / "in"; out = tmp_path / "out"
    inp.mkdir(); out.mkdir()
    _make_book(inp / "C.xlsx", [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "销售一部", "上海", 200],
    ])
    cfg = _base_cfg(str(inp), str(out),
                    value_alias_map={"销售": ["销售部", "销售一部"]})
    res = run_split(cfg, log_fn=lambda m: None)
    merged = os.path.join(res, "销售.xlsx")
    assert os.path.exists(merged)
    assert _count_rows(merged) == 2                   # 两个别名归并成一组
