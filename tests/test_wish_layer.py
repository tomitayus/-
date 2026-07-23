# -*- coding: utf-8 -*-
"""希望ソフト制約 v1（避け専用）のテスト（CP-SAT経路）。

合成入力を openpyxl でその場生成し、任意の「希望」シート（×1/×2/×3）を付けて
solver_cpsat.solve で機械検証する。予算 WISH_AVOID_BUDGET は monkeypatch で操作
（実装は build 時に読む設計）。

検証項目:
- parse_wish_mark: ×1/×2/×3・記号のみ・x表記→段階、範囲外/数値/空→None
- wish_weights: 予算を段階合計で正規化（多く出すほど1件が薄まる）
- CP-SAT が feasible な避けを叶える（避け日から外れる）
- 公平が勝つ（避けで公平を崩す状況では避けが叶わない）
- 「希望」シート無し／予算0 は従来動作（wish_marks空・no-op）
- 後方互換（希望シート無しで既存テスト骨格が変わらない）
"""

import datetime
import os
import subprocess
import sys

import openpyxl
import pandas as pd
import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import solver_cpsat  # noqa: E402
import main as main_mod  # noqa: E402

SHEET1_COLS = [
    "Date", "大学平日", "大学土曜昼", "大学土曜夜", "大学日曜昼", "大学日曜夜",
    "大学祝日昼", "大学祝日夜", "支援平日", "支援日直", "支援当直",
    "病院X", "病院Y", "病院Z", "カテ当番",
]


def build_wish_xlsx(path, doctors, dates, slots, wish=None, avail=None):
    """合成入力（sheet1/Sheet2/Sheet3[+希望]）を生成。

    doctors: 医師名リスト
    dates: datetime のリスト（連続日）
    slots: {day(int): [列名,...]} 枠を立てる
    wish: {医師: {day(int): "×N" or stage文字列}}（省略=希望シート無し）
    avail: {医師: {day(int): コード}}（省略=全日可）
    """
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "sheet1"
    ws1.append(SHEET1_COLS)
    for dt in dates:
        row = [dt] + [None] * (len(SHEET1_COLS) - 1)
        for col in slots.get(dt.day, []):
            row[SHEET1_COLS.index(col)] = 1
        ws1.append(row)

    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["Date"] + doctors)
    for dt in dates:
        row = [dt]
        for doc in doctors:
            row.append((avail or {}).get(doc, {}).get(dt.day, 1))
        ws2.append(row)

    ws3 = wb.create_sheet("Sheet3")
    ws3.append(["氏名", "属性", "カテ当番", "出張日", "出張先",
                "全合計", "大学合計", "外病院合計", "平日", "休日合計"])
    for doc in doctors:
        ws3.append([doc, 2, None, None, None, 0, 0, 0, 0, 0])

    if wish is not None:
        wsw = wb.create_sheet("希望")
        wsw.append(["Date"] + doctors)
        for dt in dates:
            row = [dt]
            for doc in doctors:
                row.append(wish.get(doc, {}).get(dt.day))
            wsw.append(row)

    wb.save(path)
    return path


def _by_slot(data, solution):
    out = {}
    for si, doc in solution["assign"].items():
        slot = data.slots[si]
        out[(datetime.date(slot["date"].year, slot["date"].month, slot["date"].day),
             slot["hosp"])] = doc
    return out


# ---------------------------------------------------------------------------
# A. parse_wish_mark
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("value,expected", [
    ("×1", 1), ("×2", 2), ("×3", 3),
    ("×", 1), ("✕2", 2), ("x3", 3), ("X1", 1), ("　×2　", 2),
    ("×4", None), ("×0", None), ("○1", None), ("1", None),
    (1, None), (2.0, None), ("", None), (None, None), ("あ", None),
])
def test_parse_wish_mark(value, expected):
    assert solver_cpsat.parse_wish_mark(value) == expected


# ---------------------------------------------------------------------------
# B. wish_weights（予算正規化 = 多く出すほど薄まる）
# ---------------------------------------------------------------------------
def test_wish_weights_normalization(tmp_path):
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 15)]
    # 医A: ×3を1日だけ / 医B: ×3を6日 / 医C: 本命×3(1日)＋おまけ×1(4日)
    wish = {
        "医A": {2: "×3"},
        "医B": {d: "×3" for d in (2, 3, 4, 5, 6, 9)},
        "医C": {2: "×3", 3: "×1", 4: "×1", 5: "×1", 6: "×1"},
    }
    path = str(build_wish_xlsx(tmp_path / "w.xlsx", ["医A", "医B", "医C"], dates,
                               {d: ["病院X"] for d in range(1, 15)}, wish=wish))
    data = solver_cpsat.InputData(path, verbose=False)
    w = data.wish_weights(6)
    # 医A: 1件に全予算 → 6
    assert list(w["医A"].values()) == [6]
    # 医B: 6件に薄まる → 各 round(6*3/18)=1
    assert set(w["医B"].values()) == {1}
    assert len(w["医B"]) == 6
    # 医C: 本命(×3)は薄まっても最大、おまけ(×1)より重い
    d2 = datetime.datetime(2026, 3, 2)
    d3 = datetime.datetime(2026, 3, 3)
    assert w["医C"][d2] > w["医C"][d3]


def test_wish_weights_budget_zero_is_empty(tmp_path):
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 8)]
    wish = {"医A": {2: "×3"}}
    path = str(build_wish_xlsx(tmp_path / "z.xlsx", ["医A", "医B"], dates,
                               {2: ["病院X"], 6: ["病院X"]}, wish=wish))
    data = solver_cpsat.InputData(path, verbose=False)
    assert data.wish_weights(0) == {}


# ---------------------------------------------------------------------------
# C. CP-SAT が feasible な避けを叶える
# ---------------------------------------------------------------------------
def test_cpsat_honors_feasible_avoid(tmp_path, monkeypatch):
    """医A が day2 を避けたい → 公平を崩さず day6 に回せるので day2 は医B。"""
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 8)]
    slots = {2: ["病院X"], 6: ["病院X"]}       # 2枠・2医師 → 公平は1回ずつ
    wish = {"医A": {2: "×3"}}
    path = str(build_wish_xlsx(tmp_path / "avoid.xlsx", ["医A", "医B"], dates,
                               slots, wish=wish))
    monkeypatch.setattr(solver_cpsat, "WISH_AVOID_BUDGET", 1000)  # soft帯で支配的
    data, sols = solver_cpsat.solve(path, n_solutions=1, verbose=False)
    by = _by_slot(data, sols[0])
    d2 = datetime.date(2026, 3, 2)
    d6 = datetime.date(2026, 3, 6)
    assert by[(d2, "病院X")] == "医B", f"医Aが避けたい day2 に入っている: {by}"
    assert by[(d6, "病院X")] == "医A"


def test_cpsat_wish_cannot_override_feasibility(tmp_path, monkeypatch):
    """避けても day1 を埋められるのが医Aだけなら、避けは叶わず医Aが入る。

    3医師3枠（各日1枠）で count==1 が全員に強制される。day1 は医B・医C 不可なので
    医Aしか入れない → 医Aは day1 を避けたくても入るしかない（避けは feasibility を覆さない）。
    """
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 12)]
    slots = {1: ["病院X"], 5: ["病院Y"], 9: ["病院Z"]}
    wish = {"医A": {1: "×3"}}
    avail = {"医B": {1: 0}, "医C": {1: 0}}   # day1 は医Aのみ可
    path = str(build_wish_xlsx(tmp_path / "fair.xlsx", ["医A", "医B", "医C"], dates,
                               slots, wish=wish, avail=avail))
    monkeypatch.setattr(solver_cpsat, "WISH_AVOID_BUDGET", 1000)
    data, sols = solver_cpsat.solve(path, n_solutions=1, verbose=False)
    by = _by_slot(data, sols[0])
    assert by[(datetime.date(2026, 3, 1), "病院X")] == "医A"


# ---------------------------------------------------------------------------
# D. 後方互換（希望シート無し / 予算0）
# ---------------------------------------------------------------------------
def test_no_wish_sheet_is_noop(tmp_path):
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 8)]
    path = str(build_wish_xlsx(tmp_path / "none.xlsx", ["医A", "医B"], dates,
                               {2: ["病院X"], 6: ["病院X"]}, wish=None))
    data = solver_cpsat.InputData(path, verbose=False)
    assert data.wish_marks == {}
    assert data.wish_col_of is None
    # solve は通常どおり解ける
    _, sols = solver_cpsat.solve(path, n_solutions=1, verbose=False)
    assert sols[0]["status"] in ("OPTIMAL", "FEASIBLE")


# ---------------------------------------------------------------------------
# E. 両エンジンのパース整合（main.py と solver_cpsat.py は独立にパースする）
# ---------------------------------------------------------------------------
def test_main_and_cpsat_parse_wish_identically(tmp_path):
    """main.py と solver_cpsat.py が同じ希望シートから同じ重みを出す（二重パース整合）。"""
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 15)]
    wish = {
        "医A": {2: "×3"},
        "医B": {d: "×2" for d in (3, 4, 5)},
        "医C": {6: "×1", 9: "×3"},
    }
    path = str(build_wish_xlsx(tmp_path / "cons.xlsx", ["医A", "医B", "医C"], dates,
                               {d: ["病院X"] for d in range(1, 15)}, wish=wish))
    # main.py 経路
    xls = pd.ExcelFile(path)
    marks, invalid = main_mod.parse_wish_sheet(xls)
    main_w = main_mod.wish_day_weights(marks, 6)
    # solver_cpsat 経路
    data = solver_cpsat.InputData(path, verbose=False)
    cp_w = data.wish_weights(6)
    assert invalid == []
    assert main_w == cp_w and main_w != {}


def test_main_parse_wish_flags_invalid(tmp_path):
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 8)]
    wish = {"医A": {2: "×3", 3: "×5", 4: "○1"}}  # ×5 と ○1 は不正
    path = str(build_wish_xlsx(tmp_path / "inv.xlsx", ["医A", "医B"], dates,
                               {2: ["病院X"]}, wish=wish))
    xls = pd.ExcelFile(path)
    marks, invalid = main_mod.parse_wish_sheet(xls)
    assert marks == {"医A": {pd.Timestamp(2026, 3, 2): 3}}
    assert len(invalid) == 2  # ×5 と ○1


# ---------------------------------------------------------------------------
# F. main.run エンドツーエンド（Greedy経路 + プリフライト + サマリー）
# ---------------------------------------------------------------------------
E2E_DOCTORS = ["医A", "医B", "医C", "医D", "医E", "医F"]
E2E_SLOTS = {
    2: ["病院X"], 3: ["大学平日"], 4: ["病院Y"], 7: ["大学土曜昼", "病院Z"],
    8: ["大学日曜夜"], 9: ["病院X"], 10: ["大学平日"], 11: ["病院Y"], 14: ["病院Z"],
    16: ["病院X"], 17: ["支援当直"], 18: ["病院Y"], 21: ["病院Z"],
    22: ["支援日直", "病院Z"], 25: ["病院X"], 26: ["病院Y"], 28: ["大学土曜夜"],
}


def test_main_run_greedy_with_wish_pipeline(tmp_path):
    """Greedy経路: 希望シート付きで完走し、プリフライト警告とサマリー行が出る。"""
    dates = [datetime.datetime(2026, 3, d) for d in range(1, 29)]
    # 有効な避け(医A ×3) + 不正マーク(×5) + 名簿に無い列
    wish = {"医A": {2: "×3", 9: "×5"}, "医Z不在": {3: "×1"}}
    input_path = str(build_wish_xlsx(tmp_path / "e2e.xlsx", E2E_DOCTORS, dates,
                                     E2E_SLOTS, wish=wish))
    out_dir = str(tmp_path / "out")
    code = (
        "import config; config.SOLVER='greedy'; "
        "import main; "
        f"p = main.run({input_path!r}, output_dir={out_dir!r}, num_patterns=5); "
        "print('OUTPUT::' + p)"
    )
    proc = subprocess.run([sys.executable, "-c", code], cwd=REPO_ROOT,
                          capture_output=True, text=True, timeout=300)
    assert proc.returncode == 0, f"main.run失敗:\n{proc.stdout[-3000:]}\n{proc.stderr[-2000:]}"
    # プリフライト警告（不正マーク or 名簿不一致）が出る
    assert "希望" in proc.stdout
    out_path = next(line.split("OUTPUT::", 1)[1]
                    for line in proc.stdout.splitlines() if "OUTPUT::" in line)
    wb = openpyxl.load_workbook(out_path, read_only=True, data_only=True)
    assert "pattern_01_summary" in wb.sheetnames
    cells = [c for row in wb["pattern_01_summary"].iter_rows(values_only=True)
             for c in row if isinstance(c, str)]
    wb.close()
    assert any("避け希望 実現" in c for c in cells), "サマリーに避け希望の実現行が無い"
