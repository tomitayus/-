"""プリフライト検証 + 氏名突合厳格化のテスト（v6.9.0向け）。

- build_prev_name_matcher: NAME_ALIASES適用後の完全一致のみ。
  同姓プレフィックス（佐藤彰/佐藤悠/佐藤勇）が誤マッチしないことを担保する。
- preflight_validate: 正常雛形はパス、日付ズレ・重複・欠落・枠数超過は致命（ValueError）、
  氏名不一致・祝日枠不整合は警告（続行）。
- run() レベル: 必須シート欠落は読込直後に致命停止する。

いずれも run() のフルパイプラインは回さない軽量テスト（harness不使用）。
"""

import os
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import main

INPUT_XLSX = os.path.join(REPO_ROOT, "docs", "Tochoku.v7_2026.03.1.xlsx")


# ---------------------------------------------------------------------------
# A. 氏名突合の厳格化（build_prev_name_matcher / apply_name_alias）
# ---------------------------------------------------------------------------

def test_matcher_exact_match():
    m = main.build_prev_name_matcher(["山田", "佐藤彰"])
    assert m("山田") == "山田"
    assert m("佐藤彰") == "佐藤彰"


def test_matcher_same_surname_prefix_never_mismatches():
    """同姓医師（佐藤彰/佐藤悠/佐藤勇）: 完全一致のみで、姓だけでは誰にもマッチしない。"""
    m = main.build_prev_name_matcher(["佐藤彰", "佐藤悠", "佐藤勇"])
    assert m("佐藤彰") == "佐藤彰"
    assert m("佐藤悠") == "佐藤悠"
    assert m("佐藤勇") == "佐藤勇"
    assert m("佐藤") is None          # 旧実装は候補3件→Noneだったが、1件なら誤マッチしていた


def test_matcher_prefix_no_longer_matches_either_direction():
    """旧実装の双方向startswithを撤去: 前方一致はどちら向きにもマッチしない。"""
    # sheet4側が姓のみ・sheet2側がフルネーム（旧実装では誤マッチ）
    m = main.build_prev_name_matcher(["佐藤"])
    assert m("佐藤彰") is None
    # sheet4側がフルネーム・sheet2側が姓のみ（旧実装では誤マッチ）
    m2 = main.build_prev_name_matcher(["佐藤彰"])
    assert m2("佐藤") is None


def test_matcher_normalizes_spaces():
    m = main.build_prev_name_matcher(["佐藤 彰"])
    assert m("佐藤彰") == "佐藤 彰"   # 返り値はsheet4側の生表記（name_to_rowのキー）


def test_matcher_applies_name_aliases(monkeypatch):
    """config.NAME_ALIASES の表記ゆれマップはどちら向きに書いても突合できる。"""
    monkeypatch.setattr(main._cfg, "NAME_ALIASES", {"冨田": "富田"}, raising=False)
    m = main.build_prev_name_matcher(["富田"])
    assert m("冨田") == "富田"
    m2 = main.build_prev_name_matcher(["冨田"])
    assert m2("富田") == "冨田"


def test_matcher_no_alias_no_match(monkeypatch):
    monkeypatch.setattr(main._cfg, "NAME_ALIASES", {}, raising=False)
    m = main.build_prev_name_matcher(["富田"])
    assert m("冨田") is None


# ---------------------------------------------------------------------------
# B. preflight_validate（合成入力）
# ---------------------------------------------------------------------------

DOCTORS = ["山田", "佐藤彰", "佐藤悠"]


def _make_inputs(dates, doctors=DOCTORS, hospital_cols=("病院A",), slot_value=1):
    """最小構成の shift_df / availability_df を合成する。"""
    dates = pd.to_datetime(list(dates)).normalize()
    data = {"Date": dates}
    for h in hospital_cols:
        data[h] = [slot_value] * len(dates)
    shift_df = pd.DataFrame(data)
    # 病院列を object 型にしておく（後から医師名の事前割当を代入してもdtypeで弾かれないように）
    for h in hospital_cols:
        shift_df[h] = shift_df[h].astype(object)
    avail = pd.DataFrame(index=pd.DatetimeIndex(dates), columns=list(doctors), dtype=float)
    name_match = {d: d for d in doctors}
    return shift_df, "Date", avail, list(doctors), list(hospital_cols), name_match


def _run_preflight(shift_df, date_col, avail, doctors, hosp_cols, name_match,
                   holidays=frozenset(), **kw):
    return main.preflight_validate(shift_df, date_col, avail, doctors,
                                   hosp_cols, set(holidays), name_match, **kw)


def test_preflight_normal_synthetic_passes(capsys):
    dates = pd.date_range("2026-03-01", periods=12, freq="D")
    args = _make_inputs(dates)
    fatals, warnings = _run_preflight(*args)
    assert fatals == []
    assert warnings == []
    out = capsys.readouterr().out
    assert "プリフライト検証" in out
    assert "一致3人 / 不一致0人" in out
    assert "致命チェックOK" in out


def test_preflight_real_template_passes(capsys):
    """実雛形（docs/Tochoku.v7_2026.03.1.xlsx）は致命・警告ゼロでパスする。"""
    xls = pd.ExcelFile(INPUT_XLSX)
    s1 = main.strip_cols(pd.read_excel(xls, sheet_name=main.find_sheet_name(xls, "sheet1")))
    s1.columns = main.make_unique(list(s1.columns))
    s2 = main.strip_cols(pd.read_excel(xls, sheet_name=main.find_sheet_name(xls, "sheet2")))
    dcol = s1.columns[0]
    s1[dcol] = pd.to_datetime(s1[dcol], errors="coerce").dt.normalize()
    acol = s2.columns[0]
    s2[acol] = pd.to_datetime(s2[acol], errors="coerce").dt.normalize()
    avail = s2.set_index(acol)
    doctors = [main.normalize_name(x) for x in s2.columns[1:]]
    hosp_cols = [c for c in s1.columns[1:] if str(c).strip() != "カテ当番"]

    s4_name = main.find_sheet_name(xls, "sheet4") or main.find_sheet_name(xls, "Sheet4") \
        or main.find_sheet_name(xls, "sheet3")
    s4 = main.parse_sheet4_from_grid(pd.read_excel(xls, sheet_name=s4_name, header=None))
    matcher = main.build_prev_name_matcher(list(s4["氏名"]))
    name_match = {d: matcher(d) for d in doctors}

    fatals, warnings = _run_preflight(
        s1, dcol, avail, doctors, hosp_cols, name_match,
        holidays={pd.Timestamp("2026-03-20")})
    assert fatals == []
    assert warnings == []
    assert "不一致0人" in capsys.readouterr().out


def test_preflight_sheet2_missing_date_is_fatal():
    """sheet1にあってsheet2に無い日付 → 致命（旧: 黙って全員可(1)フォールバック）。"""
    dates = pd.date_range("2026-03-01", periods=12, freq="D")
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(dates)
    avail = avail.drop(index=pd.Timestamp("2026-03-05"))
    with pytest.raises(ValueError) as ei:
        _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert "2026-03-05" in str(ei.value)
    assert "sheet2" in str(ei.value)


def test_preflight_sheet1_duplicate_date_is_fatal():
    dates = list(pd.date_range("2026-03-01", periods=12, freq="D"))
    dates.append(pd.Timestamp("2026-03-07"))  # 重複
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(dates)
    avail = avail[~avail.index.duplicated()]
    with pytest.raises(ValueError) as ei:
        _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert "重複" in str(ei.value)
    assert "2026-03-07" in str(ei.value)


def test_preflight_sheet1_missing_day_is_fatal():
    dates = [d for d in pd.date_range("2026-03-01", periods=12, freq="D")
             if d != pd.Timestamp("2026-03-06")]  # 期間中に抜け
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(dates)
    with pytest.raises(ValueError) as ei:
        _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert "欠落" in str(ei.value)
    assert "2026-03-06" in str(ei.value)


def test_preflight_capacity_overflow_is_fatal():
    """総枠数 > gap>=3理論上限 → 致命（物理的に埋まらない）。"""
    dates = pd.date_range("2026-03-01", periods=6, freq="D")
    # 3列×6日=18枠に対し、3医師×gap3上限2回=6回しか埋められない
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(
        dates, hospital_cols=("病院A", "病院B", "病院C"))
    with pytest.raises(ValueError) as ei:
        _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert "理論上限" in str(ei.value)


def test_preflight_unmatched_names_warn_but_continue(capsys):
    dates = pd.date_range("2026-03-01", periods=12, freq="D")
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(dates)
    nm = dict(nm)
    nm["山田"] = None  # 突合失敗
    fatals, warnings = _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert fatals == []
    assert any("山田" in w and "0扱い" in w for w in warnings)
    out = capsys.readouterr().out
    assert "一致2人 / 不一致1人" in out
    assert "NAME_ALIASES" in out


def test_preflight_holiday_column_on_plain_weekday_warns(capsys):
    """祝日ゼロ検出なのに平日に祝日枠（列名「祝日」）→ 警告（続行）。"""
    dates = pd.date_range("2026-03-01", periods=12, freq="D")
    # 病院A(12枠)+大学祝日昼(1枠)=13枠。3医師では gap>=3 理論上限12を超えて
    # 枠数超過の致命が先に出てしまうため、この検証用に4医師で余裕を持たせる
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(
        dates, doctors=["山田", "佐藤彰", "佐藤悠", "田中"],
        hospital_cols=("病院A", "大学祝日昼"))
    # 大学祝日昼は 3/3(火) のみ枠あり、祝日セットは空
    shift_df["大学祝日昼"] = [np.nan] * len(shift_df)
    shift_df.loc[shift_df["Date"] == pd.Timestamp("2026-03-03"), "大学祝日昼"] = 1
    fatals, warnings = _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert fatals == []
    assert any("祝日" in w and "2026-03-03" in w for w in warnings)


def test_preflight_invalid_marks_and_missing_cols_warn(capsys):
    dates = pd.date_range("2026-03-01", periods=12, freq="D")
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(dates)
    fatals, warnings = _run_preflight(
        shift_df, dcol, avail, doctors, hosp, nm,
        invalid_avail_marks=[(pd.Timestamp("2026-03-02"), "山田", "x")],
        missing_avail_cols=["佐藤勇"])
    assert fatals == []
    assert any("解釈できないマーク" in w for w in warnings)
    assert any("佐藤勇" in w and "列が見つからない" in w for w in warnings)


def test_preflight_preassigned_cell_counts_as_slot():
    """事前割当（セルに医師名）も枠として数えられ、上限計算にも寄与する。"""
    dates = pd.date_range("2026-03-01", periods=6, freq="D")
    shift_df, dcol, avail, doctors, hosp, nm = _make_inputs(
        dates, hospital_cols=("病院A",), slot_value=np.nan)
    shift_df.loc[0, "病院A"] = "山田"  # 枠は事前割当1つのみ
    fatals, warnings = _run_preflight(shift_df, dcol, avail, doctors, hosp, nm)
    assert fatals == []


# ---------------------------------------------------------------------------
# run() レベル: 必須シート欠落は読込直後に致命停止
# ---------------------------------------------------------------------------

def test_run_missing_sheet_is_fatal(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append(["Date", "病院A"])
    ws.append(["2026-03-01", 1])
    path = tmp_path / "broken.xlsx"
    wb.save(path)
    with pytest.raises(ValueError) as ei:
        main.run(str(path), output_dir=str(tmp_path))
    assert "必要なシートが見つかりません" in str(ei.value)
    assert "sheet2" in str(ei.value)
