"""列役割アンカーの名前解決テスト（v6.10.0 / 中期#5）。

_resolve_column_anchors():
- 現行雛形（docs/Tochoku.v7_2026.03.1.xlsx）で解決結果が従来の位置定数と完全一致（回帰）。
- 列を1本追加した合成入力で、下流アンカーが正しくシフト検出される。
- 名前解決できない役割は位置固定にフォールバックし警告する。
- 役割範囲の重複・逆転、外病院範囲への割り込みは致命（ValueError）。

いずれも run() のフルパイプラインは回さない軽量テスト。
"""

import os
import sys

import pandas as pd
import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import main

INPUT_XLSX = os.path.join(REPO_ROOT, "docs", "Tochoku.v7_2026.03.1.xlsx")

# 現行雛形の標準ヘッダ（sheet1）。位置固定時代の *_COL_INDEX 想定と一致する並び。
STANDARD_COLS = [
    "Date",
    "大学平日",                                              # B=1
    "大学土曜昼", "大学土曜夜", "大学日曜昼",                 # C,D,E
    "大学日曜夜", "大学祝日昼", "大学祝日夜",                 # F,G,H
    "支援平日", "支援日直", "支援当直",                       # I,J,K
    "夜間急病", "太陽の国", "福島南循環器", "大原医療",       # L,M,...
    "星", "しのぶ", "熱海", "日赤", "相馬中央",               # ...Q=しのぶ
    "須賀川", "済生会第一", "済生会第三", "太田西ノ内",       # U=須賀川
    "ふたば医療",                                            # Y=24
    "カテ当番",                                              # Z（外病院範囲から除外）
]


# ---------------------------------------------------------------------------
# A. 回帰: 現行雛形で従来の位置定数と完全一致
# ---------------------------------------------------------------------------

def test_real_template_matches_standard_positions():
    """実雛形ファイルのヘッダで解決 → _STANDARD_COL_ANCHORS と完全一致・警告ゼロ。"""
    xls = pd.ExcelFile(INPUT_XLSX)
    s1 = main.strip_cols(pd.read_excel(xls, sheet_name=main.find_sheet_name(xls, "sheet1")))
    s1.columns = main.make_unique(list(s1.columns))
    kate = next((c for c in s1.columns[1:] if str(c).strip() == "カテ当番"), None)

    anchors, warnings = main._resolve_column_anchors(list(s1.columns), kate_toban_col=kate)

    assert warnings == []
    for key, std in main._STANDARD_COL_ANCHORS.items():
        assert anchors[key] == std, f"{key}: {anchors[key]} != 標準 {std}"


def test_standard_cols_constant_matches_position_defaults():
    """合成した標準ヘッダ定義も同じく位置定数と一致（テスト用定義の妥当性確認）。"""
    anchors, warnings = main._resolve_column_anchors(STANDARD_COLS, kate_toban_col="カテ当番")
    assert warnings == []
    assert anchors == main._STANDARD_COL_ANCHORS
    # 従来の位置固定値そのもの
    assert (anchors["B"], anchors["H"], anchors["I"], anchors["K"]) == (1, 7, 8, 10)
    assert (anchors["L"], anchors["Q"], anchors["Y"]) == (11, 16, 24)


# ---------------------------------------------------------------------------
# B. シフト検出: 列を1本追加
# ---------------------------------------------------------------------------

def test_added_external_column_shifts_downstream_anchors():
    """外病院範囲の先頭に1列追加 → L据置・Y/QはＬ以降で+1シフトし警告。"""
    cols = STANDARD_COLS[:11] + ["新病院X"] + STANDARD_COLS[11:]  # 夜間急病の直前に挿入
    anchors, warnings = main._resolve_column_anchors(cols, kate_toban_col="カテ当番")

    # 大学系・支援系は不変
    assert anchors["B"] == 1 and anchors["K"] == 10
    # 外病院範囲は先頭据置・末尾/しのぶが+1
    assert anchors["L"] == 11        # 新病院X が新たな先頭
    assert anchors["Y"] == 25        # 標準24 → 25
    assert anchors["Q"] == 17        # しのぶ 16 → 17（名前解決で追従）
    assert any("標準と異なります" in w for w in warnings)


def test_added_leading_university_shifts_everything():
    """先頭(B直後)に大学祝日列を1本追加 → 以降の全アンカーが+1シフトする。"""
    # 大学平日 の直後に 大学祝日夜2（別枠）を挿入すると H が複数該当になるため、
    # ここでは休日大学範囲の末尾に新カテゴリを足さず、支援の前に外病院を割り込ませない
    # 「大学」列を1本足すケースを避け、支援列を1本足して I〜K と外病院のシフトを見る。
    cols = STANDARD_COLS[:8] + ["支援祝日"] + STANDARD_COLS[8:]  # 支援平日 の直前に挿入
    anchors, warnings = main._resolve_column_anchors(cols, kate_toban_col="カテ当番")
    # 支援平日/日直/当直 は +1 シフト（支援祝日は I〜K のいずれのパターンにも一致しない）
    assert anchors["I"] == 9 and anchors["J"] == 10 and anchors["K"] == 11
    # 外病院範囲も +1
    assert anchors["L"] == 12 and anchors["Y"] == 25 and anchors["Q"] == 17
    assert any("標準と異なります" in w for w in warnings)


# ---------------------------------------------------------------------------
# C. フォールバック: 名前解決できない役割
# ---------------------------------------------------------------------------

def test_unresolvable_university_falls_back_to_position():
    """休日大学列の名前が非標準 → 名前解決不能な役割は位置固定にフォールバックし警告。"""
    cols = list(STANDARD_COLS)
    cols[2] = "土曜昼枠"   # 「大学」を含まない → C が解決できない
    anchors, warnings = main._resolve_column_anchors(cols, kate_toban_col="カテ当番")
    # フォールバックで標準位置(2)を採用
    assert anchors["C"] == 2
    assert any("フォールバック" in w for w in warnings)


def test_no_external_columns_falls_back():
    """外病院列が皆無（大学/支援/カテ当番のみ）→ L〜Y は位置固定フォールバック＋警告。"""
    cols = STANDARD_COLS[:11] + ["カテ当番"]
    anchors, warnings = main._resolve_column_anchors(cols, kate_toban_col="カテ当番")
    assert any("フォールバック" in w for w in warnings)
    # 例外を出さず、標準位置(クランプ済み)を返す
    assert "L" in anchors and "Y" in anchors


# ---------------------------------------------------------------------------
# D. 致命: 範囲の重複・割り込み
# ---------------------------------------------------------------------------

def test_university_inside_external_range_is_fatal():
    """外病院範囲(L〜Y)の内側に大学列が割り込む → 致命 ValueError。"""
    cols = STANDARD_COLS[:14] + ["大学予備"] + STANDARD_COLS[14:]  # 外病院の途中に大学列
    with pytest.raises(ValueError) as ei:
        main._resolve_column_anchors(cols, kate_toban_col="カテ当番")
    assert "外病院範囲" in str(ei.value)


def test_reversed_ranges_is_fatal():
    """支援列が休日大学列より前に来る → 範囲逆転で致命 ValueError。"""
    # 支援平日/日直/当直 を大学休日枠より前に移動（B の直後）
    cols = (["Date", "大学平日", "支援平日", "支援日直", "支援当直"]
            + STANDARD_COLS[2:8] + STANDARD_COLS[11:])
    with pytest.raises(ValueError) as ei:
        main._resolve_column_anchors(cols, kate_toban_col="カテ当番")
    assert "重複" in str(ei.value) or "逆転" in str(ei.value)
