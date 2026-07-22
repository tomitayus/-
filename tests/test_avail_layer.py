"""v6.11.0 可否レイヤ3件の回帰テスト（提案#3/#4/#6）。

- A: 曜日固定外勤の前日+当日NG自動生成（複数曜日パース・追加列の寛容読み）
- B: 病院別の同日/近接当直の例外ペアテーブル（GAIKIN_EXCEPTIONS による列限定緩和）
- C: 枠種別粒度のチーム制約（TEAM_SLOT_RESTRICTIONS）

docs/Tochoku.v7_2026.03.1.xlsx は変更禁止のため、合成入力を openpyxl で生成して検証する。
合成入力は test_grade_quota.py と同じ列構成（2026-03-01(日)〜03-28(土)・医師6人・20枠）だが、
既定経路（目標列なし=クォータ無効）で動かす。
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

import main as main_mod  # noqa: E402
import solver_cpsat  # noqa: E402

DOCTORS = ["医A", "医B", "医C", "医D", "医E", "医F"]
ATTRS = {"医A": 2, "医B": 2, "医C": 2, "医D": 2, "医E": 1, "医F": 1}

SHEET1_COLS = [
    "Date", "大学平日", "大学土曜昼", "大学土曜夜", "大学日曜昼", "大学日曜夜",
    "大学祝日昼", "大学祝日夜", "支援平日", "支援日直", "支援当直",
    "病院X", "病院Y", "病院Z", "カテ当番",
]
UNIV_COLS = set(SHEET1_COLS[1:11])

# 日付 -> 枠を立てる列（2026-03-20(金)は祝日=春分の日）
SLOTS = {
    2: ["病院X"], 3: ["大学平日"], 4: ["病院Y"],
    7: ["大学土曜昼", "病院Z"], 8: ["大学日曜夜"], 9: ["病院X"],
    10: ["大学平日"], 11: ["病院Y"], 14: ["病院Z"],
    16: ["病院X"], 17: ["支援当直"], 18: ["病院Y"],
    20: ["大学祝日昼"], 21: ["病院Z"], 22: ["支援日直", "病院Z"],
    25: ["病院X"], 26: ["病院Y"], 28: ["大学土曜夜"],
}
DATES = [datetime.datetime(2026, 3, d) for d in range(1, 29)]


def build_xlsx(path, travel=None, travel2=None, dest=None):
    """合成入力を生成する。

    Args:
        travel:  {医師: 出張日セル文字列}（Sheet3「出張日」）
        travel2: {医師: 出張日セル文字列}（Sheet3 追加列「外勤曜日2」・寛容読みテスト用）
        dest:    {医師: 出張先文字列}（Sheet3「出張先」）
    """
    travel = travel or {}
    travel2 = travel2 or {}
    dest = dest or {}
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "sheet1"
    ws1.append(SHEET1_COLS)
    for dt in DATES:
        row = [dt] + [None] * (len(SHEET1_COLS) - 1)
        for col in SLOTS.get(dt.day, []):
            row[SHEET1_COLS.index(col)] = 1
        ws1.append(row)

    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["Date"] + DOCTORS)
    for dt in DATES:
        ws2.append([dt] + [1] * len(DOCTORS))

    ws3 = wb.create_sheet("Sheet3")
    header = ["氏名", "属性", "カテ当番", "出張日", "出張先",
              "全合計", "大学合計", "外病院合計", "平日", "休日合計"]
    if travel2:
        header.append("外勤曜日2")
    ws3.append(header)
    for doc in DOCTORS:
        row = [doc, ATTRS[doc], None, travel.get(doc), dest.get(doc), 0, 0, 0, 0, 0]
        if travel2:
            row.append(travel2.get(doc))
        ws3.append(row)

    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 1. 純粋関数（main.py / solver_cpsat.py 両実装の同一挙動）
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
@pytest.mark.parametrize("raw,expected", [
    ("木", {3}),
    ("木・金", {3, 4}),
    ("木,金", {3, 4}),
    ("木、金", {3, 4}),
    ("木/金", {3, 4}),
    ("木曜・金曜日", {3, 4}),
    ("日", {6}),          # 「日」単体（曜日サフィックス除去で消えないこと）
    ("月 水 金", {0, 2, 4}),
    ("", set()),
    ("nan", set()),
    ("不明", set()),
    (None, set()),
])
def test_parse_travel_weekdays(mod, raw, expected):
    assert mod.parse_travel_weekdays(raw) == expected


@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
def test_find_travel_columns_tolerant(mod):
    cols = ["氏名", "属性", "出張日", "外勤曜日2", "出張先", "大学目標", "出張日程"]
    assert mod.find_travel_weekday_columns(cols) == ["出張日", "外勤曜日2"]
    assert mod.find_travel_dest_columns(cols) == ["出張先"]


@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
def test_build_travel_restriction_map_basic(mod):
    """出張=金 → 木(前日)・金(当日)が全面不可（空frozenset）。例外なし。"""
    m = mod.build_travel_restriction_map(
        {"医A": {4}}, {"医A": ""}, ["病院X", "病院Y"], {}, {})
    assert set(m["医A"].keys()) == {3, 4}
    assert m["医A"][3] == frozenset() and m["医A"][4] == frozenset()


@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
def test_build_travel_restriction_map_exception_pair(mod):
    """例外ペア: 当日(offset0)のみ病院Xを許容 → 金=限定緩和 / 木(前日)=全面不可。"""
    ex = {"南相馬": [{"offset": 0, "allow": ["病院X"]}]}
    m = mod.build_travel_restriction_map(
        {"医A": {4}}, {"医A": "南相馬市立総合病院"}, ["病院X", "病院Y"], ex, {})
    assert m["医A"][4] == frozenset({"病院X"})   # 当日は病院Xに限り可
    assert m["医A"][3] == frozenset()            # 前日は例外なし=全面不可


@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
def test_build_travel_restriction_map_group_and_intersection(mod):
    """グループ名展開 + 連続外勤（木金）の重複日は許容列の積集合。"""
    groups = {"市内": ["病院X", "病院Y"]}
    ex = {"太陽の国": [{"offset": 0, "allow": ["市内"]},
                      {"offset": -1, "allow": ["病院Y"]}]}
    m = mod.build_travel_restriction_map(
        {"医A": {3, 4}}, {"医A": "太陽の国"}, ["病院X", "病院Y", "病院Z"], ex, groups)
    # 水=金曜外勤とは無関係、木曜外勤の前日 → offset -1 のみ → 病院Y
    assert m["医A"][2] == frozenset({"病院Y"})
    # 木=木曜の当日(市内={X,Y}) かつ 金曜の前日({Y}) → 積集合 {Y}
    assert m["医A"][3] == frozenset({"病院Y"})
    # 金=金曜の当日のみ → 市内={X,Y}
    assert m["医A"][4] == frozenset({"病院X", "病院Y"})


@pytest.mark.parametrize("mod", [main_mod, solver_cpsat])
def test_normalize_team_slot_rules(mod):
    rules = mod.normalize_team_slot_rules([
        {"doctors": ["清水", "佐藤彰"], "weekday": 5, "slot_type": "日直", "range": "external"},
        {"doctors": ["医A"], "weekday": "土曜", "slot_type": "当直", "range": "大学"},
        {"doctors": ["医B"]},                # weekday/slot_type/range 省略
        {"doctors": []},                     # 医師なし → 無視
    ])
    assert len(rules) == 3
    assert rules[0]["doctors"] == {"清水", "佐藤彰"}
    assert rules[0]["weekday"] == 5 and rules[0]["slot_type"] == "日直"
    assert rules[0]["range"] == "external"
    assert rules[1]["weekday"] == 5 and rules[1]["range"] == "university"
    assert rules[2]["weekday"] is None and rules[2]["slot_type"] is None
    assert rules[2]["range"] == "all"


# ---------------------------------------------------------------------------
# 2. CP-SAT経路（InputData / slot_allowed / solve）
# ---------------------------------------------------------------------------

def test_cpsat_same_day_and_prev_day_auto_ng(tmp_path):
    """A: 出張=金 → 前日(木)に加えて当日(金)も自動0（v6.11.0で当日を追加）。"""
    path = str(tmp_path / "travel.xlsx")
    build_xlsx(path, travel={"医A": "金"})
    data = solver_cpsat.InputData(path, verbose=False)
    assert data.avail_code("2026-03-05", "医A") == 0   # 木=前日
    assert data.avail_code("2026-03-06", "医A") == 0   # 金=当日（新規）
    assert data.avail_code("2026-03-04", "医A") == 1   # 水=制限なし
    assert data.avail_code("2026-03-06", "医B") == 1   # 他医師は無関係


def test_cpsat_multi_weekday_and_extra_column(tmp_path):
    """A: セル内複数曜日「木・金」+ 追加列「外勤曜日2」も統合して読む。"""
    path = str(tmp_path / "multi.xlsx")
    build_xlsx(path, travel={"医A": "木・金"}, travel2={"医A": "月"})
    data = solver_cpsat.InputData(path, verbose=False)
    assert data.doctor_travel_wds["医A"] == {0, 3, 4}
    # 月木金の前日+当日 → 日月水木金 が全面不可（火土のみ可）
    assert data.avail_code("2026-03-01", "医A") == 0  # 日=月曜の前日
    assert data.avail_code("2026-03-02", "医A") == 0  # 月=当日
    assert data.avail_code("2026-03-03", "医A") == 1  # 火=制限なし
    assert data.avail_code("2026-03-04", "医A") == 0  # 水=木曜の前日
    assert data.avail_code("2026-03-07", "医A") == 1  # 土=制限なし
    wd_map = data.travel_restriction["医A"]
    assert set(wd_map.keys()) == {6, 0, 2, 3, 4}  # 日月水木金
    assert all(v == frozenset() for v in wd_map.values())


def test_cpsat_exception_pair_relaxation(tmp_path, monkeypatch):
    """B: 例外ペアで当日は許容列に限り配置可（slot_allowed が列単位で判定）。"""
    monkeypatch.setattr(solver_cpsat, "GAIKIN_EXCEPTIONS",
                        {"南相馬": [{"offset": 0, "allow": ["市内"]}]})
    monkeypatch.setattr(solver_cpsat, "GAIKIN_HOSPITAL_GROUPS",
                        {"市内": ["病院X"]})
    path = str(tmp_path / "exception.xlsx")
    build_xlsx(path, travel={"医A": "金"}, dest={"医A": "南相馬"})
    data = solver_cpsat.InputData(path, verbose=False)
    # 当日(金)は自動0ではなくなる（列限定緩和）
    assert data.avail_code("2026-03-06", "医A") == 1
    # 前日(木)は例外なし → 従来どおり全面不可
    assert data.avail_code("2026-03-05", "医A") == 0
    # 金曜(3/6)の枠: 病院X=可 / それ以外の列=GAIKIN-EX
    fri = pd.Timestamp("2026-03-06")
    slot_x = dict(date=fri, hidx=data.hosp_idx["病院X"], hosp="病院X")
    slot_y = dict(date=fri, hidx=data.hosp_idx["病院Y"], hosp="病院Y")
    slot_univ = dict(date=fri, hidx=data.hosp_idx["大学平日"], hosp="大学平日")
    assert data.slot_allowed(slot_x, "医A") is None
    assert data.slot_allowed(slot_y, "医A") == "GAIKIN-EX"
    assert data.slot_allowed(slot_univ, "医A") == "GAIKIN-EX"
    # 他医師には影響しない
    assert data.slot_allowed(slot_y, "医B") is None


def test_cpsat_team_slot_restriction(tmp_path, monkeypatch):
    """C: チーム枠種制約 — 医師集合×曜日×枠種別(日直)×列範囲(大学)の禁止。"""
    monkeypatch.setattr(solver_cpsat, "TEAM_SLOT_RESTRICTIONS", [
        {"doctors": ["医A", "医B"], "weekday": 5, "slot_type": "日直", "range": "university"},
    ])
    path = str(tmp_path / "team.xlsx")
    build_xlsx(path)
    data = solver_cpsat.InputData(path, verbose=False)
    sat = pd.Timestamp("2026-03-07")  # 土曜
    slot_day = dict(date=sat, hidx=data.hosp_idx["大学土曜昼"], hosp="大学土曜昼")   # 日直
    slot_night = dict(date=sat, hidx=data.hosp_idx["大学土曜夜"], hosp="大学土曜夜")  # 当直
    slot_ext = dict(date=sat, hidx=data.hosp_idx["病院Z"], hosp="病院Z")             # 外病院
    assert data.slot_allowed(slot_day, "医A") == "TEAM-001"
    assert data.slot_allowed(slot_day, "医B") == "TEAM-001"
    assert data.slot_allowed(slot_day, "医C") is None       # 対象外医師
    assert data.slot_allowed(slot_night, "医A") is None     # 枠種別(当直)は対象外
    assert data.slot_allowed(slot_ext, "医A") is None       # 列範囲(external)は対象外
    # 平日(金曜)の日直相当列は無い＝曜日不一致でも禁止されない
    fri_slot = dict(date=pd.Timestamp("2026-03-22"), hidx=data.hosp_idx["支援日直"], hosp="支援日直")
    assert data.slot_allowed(fri_slot, "医A") is None       # 3/22は日曜=weekday6≠5


def test_cpsat_solve_respects_travel_ng(tmp_path):
    """A: 求解結果に出張前日/当日の割当が現れない（エンドツーエンド）。"""
    path = str(tmp_path / "solve_travel.xlsx")
    build_xlsx(path, travel={"医A": "金", "医B": "水"})
    data, solutions = solver_cpsat.solve(path, n_solutions=1, verbose=False)
    assert solutions
    ng_wd = {"医A": {3, 4}, "医B": {1, 2}}  # 前日+当日の曜日
    for si, doc in solutions[0]["assign"].items():
        slot = data.slots[si]
        if doc in ng_wd:
            assert slot["date"].weekday() not in ng_wd[doc], (
                f"{doc} が出張前日/当日に割当: {slot['date'].date()} {slot['hosp']}")


# ---------------------------------------------------------------------------
# 3. main.py 側（グリッド読み・可否レイヤの構成関数）
# ---------------------------------------------------------------------------

def test_main_parse_sheet4_keeps_extra_travel_columns_as_string(tmp_path):
    """追加列「外勤曜日2」が数値化で壊れない（寛容読みの前提）。"""
    path = str(tmp_path / "grid.xlsx")
    build_xlsx(path, travel={"医A": "木・金"}, travel2={"医A": "月"})
    grid = pd.read_excel(pd.ExcelFile(path), sheet_name="Sheet3", header=None)
    data = main_mod.parse_sheet4_from_grid(grid)
    row = data[data["氏名"] == "医A"].iloc[0]
    assert row["出張日"] == "木・金"
    assert row["外勤曜日2"] == "月"


def test_main_run_travel_ng_end_to_end(tmp_path):
    """main.run(): 前日+当日NG + プリフライトの自動不可サマリ表示（サブプロセス実行）。"""
    input_path = str(tmp_path / "travel_e2e.xlsx")
    build_xlsx(input_path, travel={"医A": "金"})
    out_dir = str(tmp_path / "out")
    code = (
        "import main; "
        f"p = main.run({input_path!r}, output_dir={out_dir!r}); "
        "print('OUTPUT::' + p)"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code], cwd=REPO_ROOT,
        capture_output=True, text=True, timeout=300)
    assert proc.returncode == 0, f"main.run失敗:\n{proc.stdout[-3000:]}\n{proc.stderr[-3000:]}"
    assert "外勤由来の自動不可" in proc.stdout  # プリフライトのサマリ集計
    out_path = next(line.split("OUTPUT::", 1)[1]
                    for line in proc.stdout.splitlines() if "OUTPUT::" in line)

    wb = openpyxl.load_workbook(out_path, read_only=True, data_only=True)
    ws = wb["pattern_01"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = list(rows[1])
    col_index = {name: j for j, name in enumerate(header)}
    date_j = 0
    for row in rows[2:]:
        d = pd.to_datetime(str(row[date_j]).split(" (")[0])
        for col in SHEET1_COLS[1:-1]:
            v = row[col_index[col]]
            if isinstance(v, str) and v.strip() == "医A":
                assert d.weekday() not in (3, 4), (
                    f"医A が出張前日/当日に割当: {d.date()} {col}")


def test_main_run_greedy_travel_ng(tmp_path):
    """Greedy経路でも出張前日/当日NGが効く（config.SOLVER='greedy' で実行）。"""
    input_path = str(tmp_path / "travel_greedy.xlsx")
    build_xlsx(input_path, travel={"医A": "金"})
    out_dir = str(tmp_path / "out")
    code = (
        "import config; config.SOLVER='greedy'; "
        "import main; "
        f"p = main.run({input_path!r}, output_dir={out_dir!r}, num_patterns=5); "
        "print('OUTPUT::' + p)"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code], cwd=REPO_ROOT,
        capture_output=True, text=True, timeout=300)
    assert proc.returncode == 0, f"main.run失敗:\n{proc.stdout[-3000:]}\n{proc.stderr[-3000:]}"
    assert "Greedy" in proc.stdout
    out_path = next(line.split("OUTPUT::", 1)[1]
                    for line in proc.stdout.splitlines() if "OUTPUT::" in line)

    wb = openpyxl.load_workbook(out_path, read_only=True, data_only=True)
    ws = wb["pattern_01"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = list(rows[1])
    col_index = {name: j for j, name in enumerate(header)}
    for row in rows[2:]:
        d = pd.to_datetime(str(row[0]).split(" (")[0])
        for col in SHEET1_COLS[1:-1]:
            v = row[col_index[col]]
            if isinstance(v, str) and v.strip() == "医A":
                assert d.weekday() not in (3, 4), (
                    f"医A が出張前日/当日に割当(greedy): {d.date()} {col}")
