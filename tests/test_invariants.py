"""出力xlsxに対する不変条件（制約充足）テスト。

セッション共有フィクスチャ `harness` が main.run() を1回実行し、生成された
出力xlsxを独立に読み直す。ここでは xlsx 由来の割当値を、main が算出した参照値
（可否コード・TARGET_CAP・EXTRA_ALLOWED 等）と突き合わせて検証する。

対象パターン: ABS制約系は絶対禁忌クリア（バナー ✅）のシートに対してのみ検証する。
出力は abs_valid なパターンを優先選択するため通常 pattern_01 はクリア。万一
フォールバック（全パターン違反）の場合は skip して偽陽性を避ける。
"""

import pandas as pd
import pytest

PRIMARY = "pattern_01"


def _assignments(harness, sheet=PRIMARY):
    return harness.patterns[sheet]["assignments"]


def _is_abs_clear(harness, sheet):
    banner = harness.patterns[sheet].get("banner") or ""
    return isinstance(banner, str) and banner.startswith("✅")


def _hidx(harness, colname):
    """出力上の枠列名 -> shift_df 上の列インデックス（main の列範囲定義と整合）。"""
    return harness.main.shift_df.columns.get_loc(colname)


# ---------------------------------------------------------------------------
# 枠の充足
# ---------------------------------------------------------------------------

def test_slot_count_is_98(harness):
    assert len(harness.slots) == 98, f"枠数が98ではありません: {len(harness.slots)}"


def test_all_pattern_sheets_present(harness):
    sheets = harness.pattern_sheets()
    assert len(sheets) >= 1, "pattern_XX シートが出力されていません"
    assert PRIMARY in sheets


@pytest.mark.parametrize("nth", [0, 1, 2])
def test_all_slots_filled_with_real_doctor(harness, nth):
    sheets = harness.pattern_sheets()
    if nth >= len(sheets):
        pytest.skip(f"パターンが{len(sheets)}個のみ")
    sheet = sheets[nth]
    unfilled = [
        (rp, col, a["raw"])
        for (rp, col), a in harness.patterns[sheet]["assignments"].items()
        if a["doc"] is None
    ]
    assert not unfilled, f"{sheet}: 未割当/不明医師の枠が {len(unfilled)}件: {unfilled[:5]}"


def test_data_row_count_matches_input(harness):
    n_input_dates = len(harness.dates)
    for sheet in harness.pattern_sheets():
        assert harness.patterns[sheet]["n_data_rows"] == n_input_dates, (
            f"{sheet}: データ行数 {harness.patterns[sheet]['n_data_rows']} != 入力日数 {n_input_dates}"
        )


# ---------------------------------------------------------------------------
# 可否コード（ABS-001 / ABS-002 / ABS-003）
# ---------------------------------------------------------------------------

def test_code0_days_have_no_assignment(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない（フォールバック出力）")
    get_code = harness.main.get_avail_code
    viol = []
    for (rp, col), a in _assignments(harness).items():
        if a["pre"] or a["doc"] is None:
            continue
        if get_code(a["date"], a["doc"]) == 0:
            viol.append((a["date"], a["doc"], col))
    assert not viol, f"可否コード0の日に割当あり（ABS-001）: {viol[:5]}"


def test_code2_column_restriction(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない")
    m = harness.main
    get_code = m.get_avail_code
    checked = 0
    viol = []
    for (rp, col), a in _assignments(harness).items():
        if a["pre"] or a["doc"] is None:
            continue
        if get_code(a["date"], a["doc"]) == 2:
            checked += 1
            idx = _hidx(harness, col)
            if not (m.B_COL_INDEX <= idx <= m.Q_COL_INDEX):
                viol.append((a["date"], a["doc"], col, idx))
    assert not viol, f"コード2の列制約違反（ABS-002, B〜Q列のみ）: {viol[:5]}"
    # このデータセットにはコード2が存在する前提（検証が空振りしていないことを担保）
    assert checked > 0, "コード2の割当が1件も検出されず、テストが空振りしています"


def test_code3_column_restriction(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない")
    m = harness.main
    get_code = m.get_avail_code
    viol = []
    for (rp, col), a in _assignments(harness).items():
        if a["pre"] or a["doc"] is None:
            continue
        if get_code(a["date"], a["doc"]) == 3:
            idx = _hidx(harness, col)
            if not (m.L_COL_INDEX <= idx <= m.L_Y_END_INDEX):
                viol.append((a["date"], a["doc"], col, idx))
    assert not viol, f"コード3の列制約違反（ABS-003, L〜Y列のみ）: {viol[:5]}"


# ---------------------------------------------------------------------------
# 同日重複 / gap / 同一外病院重複（ABS-006 / ABS-007 / ABS-008）
# ---------------------------------------------------------------------------

def test_no_same_day_duplicate(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない")
    by_date = {}
    for (rp, col), a in _assignments(harness).items():
        if a["doc"] is None:
            continue
        by_date.setdefault(a["date"], []).append((a["doc"], a["pre"]))
    viol = []
    for date, lst in by_date.items():
        # 固定割当絡み以外での同日重複を検出
        free_docs = [doc for doc, pre in lst if not pre]
        seen = set()
        for doc in free_docs:
            if doc in seen:
                viol.append((date, doc))
            seen.add(doc)
    assert not viol, f"同日重複あり（ABS-006）: {viol[:5]}"


def test_gap_at_least_3_days(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない")
    by_doc = {}
    for (rp, col), a in _assignments(harness).items():
        if a["doc"] is None:
            continue
        by_doc.setdefault(a["doc"], []).append((a["date"], a["pre"]))
    viol = []
    for doc, lst in by_doc.items():
        lst.sort(key=lambda t: t[0])
        for i in range(1, len(lst)):
            (d0, pre0), (d1, pre1) = lst[i - 1], lst[i]
            if pre0 or pre1:
                continue  # 固定割当絡みは対象外
            if (d1 - d0).days < 3:
                viol.append((doc, d0.date(), d1.date(), (d1 - d0).days))
    assert not viol, f"gap<3 の連続当直あり（ABS-007, 固定割当以外）: {viol[:5]}"


def test_no_external_hospital_duplicate(harness):
    if not _is_abs_clear(harness, PRIMARY):
        pytest.skip("pattern_01 が絶対禁忌クリアではない")
    m = harness.main
    # 外病院 = L〜Y列（idx 11..24）
    by_doc_hosp = {}
    for (rp, col), a in _assignments(harness).items():
        if a["doc"] is None or a["pre"]:
            continue
        idx = _hidx(harness, col)
        if m.L_COL_INDEX <= idx <= m.L_Y_END_INDEX:
            by_doc_hosp.setdefault((a["doc"], col), 0)
            by_doc_hosp[(a["doc"], col)] += 1
    viol = [(doc, col, cnt) for (doc, col), cnt in by_doc_hosp.items() if cnt > 1]
    assert not viol, f"同一外病院への重複割当あり（ABS-008）: {viol[:5]}"


# ---------------------------------------------------------------------------
# TARGET_CAP / EXTRA枠（ABS-010 + EXTRA運用ルール）
# ---------------------------------------------------------------------------

def _count_per_doctor(harness, sheet=PRIMARY):
    counts = {}
    for (rp, col), a in harness.patterns[sheet]["assignments"].items():
        if a["doc"] is None:
            continue
        counts[a["doc"]] = counts.get(a["doc"], 0) + 1
    return counts


def test_target_cap_not_exceeded(harness):
    m = harness.main
    counts = _count_per_doctor(harness)
    viol = [(doc, c, m.TARGET_CAP.get(doc, 0)) for doc, c in counts.items() if c > m.TARGET_CAP.get(doc, 0)]
    assert not viol, f"TARGET_CAP超過（ABS-010）: {viol[:5]}"


def test_extra_allowed_are_attr1_tail(harness):
    """EXTRA(+1)は『属性1 → 名簿末尾（若手）から』という意図的運用ルール。"""
    m = harness.main
    if m.EXTRA_SLOTS <= 0:
        pytest.skip("EXTRA枠なし")
    attr1_sorted = [d for d in m.active_sorted_by_index if m.doctor_attribute.get(d, "") == "1"]
    assert len(attr1_sorted) >= m.EXTRA_SLOTS, "属性1医師が EXTRA_SLOTS より少ない（このデータでは想定外）"
    expected = set(attr1_sorted[-m.EXTRA_SLOTS:])
    assert m.EXTRA_ALLOWED == expected, (
        f"EXTRA_ALLOWED が属性1末尾と一致しない: got={m.EXTRA_ALLOWED} expected={expected}"
    )
    # 全 EXTRA が属性1
    assert all(m.doctor_attribute.get(d, "") == "1" for d in m.EXTRA_ALLOWED)


def test_over_base_target_only_extra_doctors(harness):
    """BASE_TARGET を超えて割り当てられるのは EXTRA_ALLOWED の医師だけ。"""
    m = harness.main
    counts = _count_per_doctor(harness)
    viol = [
        (doc, c) for doc, c in counts.items()
        if c > m.BASE_TARGET and doc not in m.EXTRA_ALLOWED
    ]
    assert not viol, f"EXTRA以外がBASE_TARGET超過: {viol[:5]}"


# ---------------------------------------------------------------------------
# 出力体裁（バナー / サマリー）
# ---------------------------------------------------------------------------

def test_pattern01_banner_and_header(harness):
    p = harness.patterns[PRIMARY]
    banner = p["banner"]
    assert isinstance(banner, str) and banner.strip(), "A1にバナー文字列がありません"
    assert ("絶対禁忌" in banner) or banner.startswith("✅") or banner.startswith("⚠️")
    # ヘッダは2行目（先頭列 = 'Date'）
    assert p["header"][0] == "Date", f"2行目がヘッダ（Date始まり）ではありません: {p['header'][:3]}"


def test_summary_total_score_row_is_numeric(harness):
    grid = harness.read_summary(f"{PRIMARY}_summary")
    found = None
    for row in grid:
        for j, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip() == "総合スコア（raw）":
                # 隣接セル（値列）を取得
                val = row[j + 1] if j + 1 < len(row) else None
                found = (cell, val)
                break
        if found:
            break
    assert found is not None, "サマリーに『総合スコア（raw）』行が見つかりません"
    _, val = found
    assert isinstance(val, (int, float)) and not pd.isna(val), f"総合スコア（raw）が数値ではありません: {val!r}"
