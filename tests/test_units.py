"""純粋関数のユニットテスト。

main のトップレベル（行1-542）に定義された関数は import だけで副作用なく
到達できる。ここではそれらを直接検証する。加えて run() 実行後にモジュールへ
公開される is_holiday / get_avail_code は harness フィクスチャ経由で検証する。
"""

import os
import sys

import numpy as np
import pandas as pd
import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import main  # import は副作用なし（run()化済み）


# ---------------------------------------------------------------------------
# normalize_name
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw,expected", [
    ("山田", "山田"),
    (" 山田 ", "山田"),
    ("山　田", "山田"),          # 全角スペース除去
    ("山 田", "山田"),           # 半角スペース除去
    ("  山　田  ", "山田"),
])
def test_normalize_name_strips_spaces(raw, expected):
    assert main.normalize_name(raw) == expected


def test_normalize_name_nan_returns_empty():
    assert main.normalize_name(np.nan) == ""
    assert main.normalize_name(None) == ""


# ---------------------------------------------------------------------------
# make_unique
# ---------------------------------------------------------------------------

def test_make_unique_disambiguates_duplicates():
    assert main.make_unique(["a", "b", "a", "a", "b"]) == ["a", "b", "a_2", "a_3", "b_2"]


def test_make_unique_handles_nan():
    out = main.make_unique([np.nan, np.nan])
    assert out[0] == "" and out[1] == "_2"


# ---------------------------------------------------------------------------
# safe_str
# ---------------------------------------------------------------------------

def test_safe_str():
    assert main.safe_str(np.nan) == ""
    assert main.safe_str("  x  ") == "x"
    assert main.safe_str(3) == "3"


# ---------------------------------------------------------------------------
# is_slot_value
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("v", [1, 1.0, "1", "〇", "○", "◯", "◎"])
def test_is_slot_value_true(v):
    assert main.is_slot_value(v) is True


@pytest.mark.parametrize("v", [0, 0.0, 2, 2.0, "", "医師名", "×", None, np.nan])
def test_is_slot_value_false(v):
    assert main.is_slot_value(v) is False


# ---------------------------------------------------------------------------
# find_sheet_name
# ---------------------------------------------------------------------------

class _FakeXls:
    def __init__(self, names):
        self.sheet_names = names


def test_find_sheet_name_exact():
    xls = _FakeXls(["sheet1", "Sheet2", "Sheet3"])
    assert main.find_sheet_name(xls, "sheet1") == "sheet1"


def test_find_sheet_name_case_insensitive():
    xls = _FakeXls(["SHEET1", "sheet2"])
    assert main.find_sheet_name(xls, "sheet1") == "SHEET1"
    assert main.find_sheet_name(xls, "Sheet2") == "sheet2"


def test_find_sheet_name_whitespace():
    xls = _FakeXls([" sheet1 "])
    assert main.find_sheet_name(xls, "sheet1") == " sheet1 "


def test_find_sheet_name_missing_returns_none():
    xls = _FakeXls(["a", "b"])
    assert main.find_sheet_name(xls, "sheet1") is None


# ---------------------------------------------------------------------------
# parse_sheet4_from_grid
# ---------------------------------------------------------------------------

def test_parse_sheet4_from_grid_basic():
    # header=None で読み込んだ想定のグリッド（先頭に注記行があるケース）
    grid = pd.DataFrame([
        ["メモ行", None, None, None],
        ["氏名", "属性", "全合計", "大学合計"],
        ["山田", "1", "3", "2"],
        [" 佐藤 ", "2", "4", "1"],
        [None, None, None, None],      # 空行は除去される
    ])
    out = main.parse_sheet4_from_grid(grid)
    assert list(out["氏名"]) == ["山田", "佐藤"]           # strip 済み
    assert list(out["属性"]) == ["1", "2"]                 # 属性は文字列保持
    assert list(out["全合計"]) == [3.0, 4.0]               # 数値化
    assert list(out["大学合計"]) == [2.0, 1.0]


def test_parse_sheet4_missing_name_header_raises():
    grid = pd.DataFrame([
        ["col_a", "col_b"],
        ["x", "y"],
    ])
    with pytest.raises(ValueError):
        main.parse_sheet4_from_grid(grid)


def test_parse_sheet4_numeric_coercion_fills_zero():
    grid = pd.DataFrame([
        ["氏名", "全合計"],
        ["山田", "notnum"],   # 数値化できない → 0
    ])
    out = main.parse_sheet4_from_grid(grid)
    assert out["全合計"].iloc[0] == 0.0


# ---------------------------------------------------------------------------
# 祝日判定（run() 実行後に公開される is_holiday）
# ---------------------------------------------------------------------------

def test_builtin_holiday_2026_03_20(harness):
    """内蔵祝日表: 2026-03-20（春分の日）は祝日と判定される。"""
    is_holiday = harness.main.is_holiday
    assert is_holiday(pd.Timestamp("2026-03-20")) is True


def test_non_holiday_weekday(harness):
    is_holiday = harness.main.is_holiday
    # 2026-03-17 (火) は平日
    assert is_holiday(pd.Timestamp("2026-03-17")) is False


# ---------------------------------------------------------------------------
# 可否コード解釈（get_avail_code, run() 実行後に公開）
# ---------------------------------------------------------------------------

def _pick_no_travel_doctor(m):
    travel = getattr(m, "doctor_travel_day", {})
    for doc in m.doctor_names:
        if doc in m.availability_df.columns and not travel.get(doc, ""):
            return doc
    return m.doctor_names[0]


def test_get_avail_code_interpretation(harness):
    """空欄→可(1) / 0.0→不可(0) / 0.5→解釈不能は可(1)扱い（読込時に警告対象）。"""
    m = harness.main
    doc = _pick_no_travel_doctor(m)
    # 出張なし医師なので出張前日ロジックは発火しない。有効な日付を1つ選ぶ。
    valid_dates = [d for d in m.availability_df.index if isinstance(d, pd.Timestamp) and not pd.isna(d)]
    assert valid_dates, "availability_df に有効な日付がありません"
    date = valid_dates[len(valid_dates) // 2]

    original = m.availability_df.at[date, doc]
    try:
        # 空欄（NaN）→ 1
        m.availability_df.at[date, doc] = np.nan
        assert m.get_avail_code(date, doc) == 1

        # 0.0 → 0（不可）: セルが実際に読まれていることの担保
        m.availability_df.at[date, doc] = 0.0
        assert m.get_avail_code(date, doc) == 0

        # 0.5 → 解釈不能 → 可(1)扱い（旧実装の int 切り捨て 0.5→0 とは異なる）
        m.availability_df.at[date, doc] = 0.5
        assert m.get_avail_code(date, doc) == 1
    finally:
        m.availability_df.at[date, doc] = original
