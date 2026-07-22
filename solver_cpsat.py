#!/usr/bin/env python3
"""
solver_cpsat.py — 当直くん CP-SAT厳密解ソルバー（併走版 Phase 1）

既存 main.py（Greedy+fixパイプライン）とは独立した比較用CLI。
入力Excel（sheet1=枠 / Sheet2=可否コード / Sheet3or4=医師属性+前月累積）を
自力でパースし、Google OR-Tools CP-SATで厳密解を求めて
`<入力名>_cpsat_v1.xlsx` を出力する。

制約の正本: docs/CONSTRAINT_RULES.md
実装範囲:
  - ABS-001〜015（絶対禁忌。ABS-009=全枠割当を含む）
  - HARD-001/002（B/I列1回まで・C-H/J-K列1回まで。カテ保有×属性1は緩和）
  - SEMI-001（B列カテ表必須。属性1は週1回まで緩和可＝違反数を第1目的で最小化）
目的関数（辞書式）:
  1. SEMI-001違反数最小化
  2. 月内公平性 max(割当数)-min(割当数) 最小化
  3. SOFT簡易版（外病院0回 / BG-HT差 / コード1.2大学0回）

使い方:
    python3 solver_cpsat.py <input.xlsx> [--output-dir DIR]
                            [--patterns N] [--time-limit SEC] [--min-diff N]

main.py には一切依存しない（import しない）。
"""

import argparse
import os
import sys
import time
from collections import defaultdict

import numpy as np
import pandas as pd
from ortools.sat.python import cp_model

SOLVER_VERSION = "cpsat_v1"

# =========================
# 列インデックス定義（main.py と同一のテンプレ依存定義: B〜Y列）
# =========================
B_COL = 1
C_COL = 2
D_COL = 3
E_COL = 4
F_COL = 5
G_COL = 6
H_COL = 7
I_COL = 8
J_COL = 9
K_COL = 10
L_COL = 11
Q_COL = 16
Y_COL = 24

# sheet1 の「枠」扱いする入力値（main.py SLOT_MARKERS と同一）
SLOT_MARKERS = {1, 1.0, "1", "〇", "○", "◯", "◎"}

WEEKDAY_MAP = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}

# 内蔵祝日表（main.py v6.8.0 の _BUILTIN_HOLIDAYS を移植）
BUILTIN_HOLIDAYS = [
    # 2026年（内閣府発表の国民の祝日・休日）
    "2026-01-01", "2026-01-12", "2026-02-11", "2026-02-23", "2026-03-20",
    "2026-04-29", "2026-05-03", "2026-05-04", "2026-05-05", "2026-05-06",
    "2026-07-20", "2026-08-11", "2026-09-21", "2026-09-22", "2026-09-23",
    "2026-10-12", "2026-11-03", "2026-11-23",
    # 2027年
    "2027-01-01", "2027-01-11", "2027-02-11", "2027-02-23", "2027-03-21",
    "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
    "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11",
    "2027-11-03", "2027-11-23",
]

# config.py と同一の既定値（main.py 非依存のため config.py があれば読む）
try:
    import config as _cfg
    WED_FORBIDDEN_DEFAULT = list(getattr(_cfg, "WED_FORBIDDEN_DOCTORS", []))
    CFG_HOLIDAYS = list(getattr(_cfg, "HOLIDAYS", []))
except Exception:
    WED_FORBIDDEN_DEFAULT = ["金城", "山田", "野寺"]
    CFG_HOLIDAYS = []


# =========================
# ユーティリティ（main.py から移植・同一挙動）
# =========================
def normalize_name(name):
    if pd.isna(name):
        return ""
    return str(name).strip().replace(" ", "").replace("　", "")


def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def norm_attr(v):
    """属性値の正規化: 2.0 -> '2' など（Excel float化対策）"""
    s = safe_str(v)
    if s.lower() in ("nan", "none"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s


def make_unique(names):
    seen = {}
    out = []
    for n in names:
        n = "" if pd.isna(n) else n
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
    return out


def find_sheet_name(xls, target):
    if target in xls.sheet_names:
        return target
    low_map = {s.lower(): s for s in xls.sheet_names}
    if target.lower() in low_map:
        return low_map[target.lower()]
    for s in xls.sheet_names:
        if s.strip().lower() == target.strip().lower():
            return s
    return None


def is_slot_value(v):
    if isinstance(v, str):
        return v.strip() in SLOT_MARKERS
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v) == 1.0
    return False


def norm_date(d):
    return pd.to_datetime(d).normalize().tz_localize(None)


# =========================
# 入力データ（main.py の読込ロジックと同一挙動で自力パース）
# =========================
class InputData:
    def __init__(self, path, verbose=True):
        self.path = path
        self.verbose = verbose
        self._log = print if verbose else (lambda *a, **k: None)
        self._parse()

    # ---- パース本体 ----
    def _parse(self):
        xls = pd.ExcelFile(self.path)
        sheet1 = find_sheet_name(xls, "sheet1")
        sheet2 = find_sheet_name(xls, "sheet2")
        sheet3 = find_sheet_name(xls, "sheet3")
        sheet4 = find_sheet_name(xls, "sheet4") or find_sheet_name(xls, "Sheet4")
        # v6.5.0: Sheet4がなければSheet3を医師情報として使用（旧カテ表は廃止）
        if sheet4 is None and sheet3 is not None:
            sheet4 = sheet3
            sheet3 = None
        missing = [k for k, v in [("sheet1", sheet1), ("sheet2", sheet2),
                                  ("sheet4/医師情報", sheet4)] if v is None]
        if missing:
            raise ValueError(f"必要なシートが見つかりません: {missing} / 実際: {xls.sheet_names}")

        shift_df = pd.read_excel(xls, sheet_name=sheet1)
        shift_df.columns = make_unique(
            [c.strip() if isinstance(c, str) else c for c in shift_df.columns])
        avail_raw = pd.read_excel(xls, sheet_name=sheet2)
        avail_raw.columns = make_unique(
            [c.strip() if isinstance(c, str) else c for c in avail_raw.columns])

        self.shift_df = shift_df
        self.date_col = shift_df.columns[0]
        shift_df[self.date_col] = pd.to_datetime(
            shift_df[self.date_col], errors="coerce").dt.normalize().dt.tz_localize(None)

        date_col_avail = avail_raw.columns[0]
        avail_raw[date_col_avail] = pd.to_datetime(
            avail_raw[date_col_avail], errors="coerce").dt.normalize().dt.tz_localize(None)
        self.avail_df = avail_raw.set_index(date_col_avail)
        # 列名（医師名）を正規化した対応表
        self.avail_col_of = {}
        for c in avail_raw.columns[1:]:
            self.avail_col_of[normalize_name(c)] = c

        self.doctor_names = [normalize_name(x) for x in avail_raw.columns[1:]]
        self.doctor_col_index = {d: i for i, d in enumerate(self.doctor_names)}

        # --- 祝日（jpholiday試行 → 内蔵祝日表。config.HOLIDAYSは常に追加） ---
        self.holidays = set()
        sched_dates = shift_df[self.date_col].dropna()
        target_ym = {(d.year, d.month) for d in sched_dates}
        holiday_source = "なし"
        if target_ym:
            try:
                import jpholiday
                for y, m in sorted(target_ym):
                    for hd, _hname in jpholiday.month_holidays(y, m):
                        self.holidays.add(pd.Timestamp(hd))
                holiday_source = "jpholiday（自動）"
            except ImportError:
                for hs in BUILTIN_HOLIDAYS:
                    ht = pd.Timestamp(hs)
                    if (ht.year, ht.month) in target_ym:
                        self.holidays.add(ht)
                holiday_source = "内蔵祝日表(2026-2027)"
                if any(y > 2027 for y, _m in target_ym):
                    self._log("⚠️ 2028年以降の祝日は内蔵表にありません")
        for h in CFG_HOLIDAYS:
            self.holidays.add(pd.Timestamp(h))
        self.holidays = {norm_date(d) for d in self.holidays}
        self._log(f"📅 祝日ソース: {holiday_source} / 対象期間内: "
                  + (", ".join(h.strftime('%m/%d') for h in sorted(self.holidays)
                               if (h.year, h.month) in target_ym) or "なし"))

        # --- Sheet1: カテ当番列・病院列 ---
        all_cols = list(shift_df.columns[1:])
        self.kate_col = None
        for c in all_cols:
            if str(c).strip() == "カテ当番":
                self.kate_col = c
                break
        self.hospital_cols = [c for c in all_cols if c != self.kate_col]
        self.hosp_idx = {h: shift_df.columns.get_loc(h) for h in self.hospital_cols}
        self.ly_end = min(Y_COL, len(shift_df.columns) - 1)

        self.kate_team_by_date = {}
        self.expected_team_codes = set()
        if self.kate_col is not None:
            for ridx in shift_df.index:
                d = shift_df.at[ridx, self.date_col]
                if pd.isna(d):
                    continue
                v = shift_df.at[ridx, self.kate_col]
                if pd.notna(v):
                    s = str(v).strip()
                    if s and s.lower() != "nan":
                        self.kate_team_by_date[norm_date(d)] = s
                        self.expected_team_codes.add(s)

        # --- Sheet4（=新構造ではSheet3）: 医師属性 ---
        grid = pd.read_excel(xls, sheet_name=sheet4, header=None)
        sheet4_data = self._parse_sheet4(grid)
        self.sheet4_data = sheet4_data
        name_to_row = {normalize_name(r["氏名"]): r for _, r in sheet4_data.iterrows()}
        prev_names = list(name_to_row.keys())

        def match_prev(doc):
            if doc in name_to_row:
                return doc
            ms = [p for p in prev_names if str(p).startswith(doc) or doc.startswith(str(p))]
            return ms[0] if len(ms) == 1 else None

        name_match = {d: match_prev(d) for d in self.doctor_names}

        def get_str(doc, col):
            p = name_match.get(doc)
            if p and p in name_to_row:
                return norm_attr(name_to_row[p].get(col, ""))
            return ""

        # カテチーム列の選択（main.py v6.5.1 と同一: カテ当番→属性→カテ→チーム）
        kate_team_col = None
        for cand in ["カテ当番", "属性", "カテ", "チーム"]:
            if cand not in sheet4_data.columns:
                continue
            col_vals = {norm_attr(v) for v in sheet4_data[cand] if norm_attr(v)}
            if self.expected_team_codes and col_vals & self.expected_team_codes:
                kate_team_col = cand
                break
            elif not self.expected_team_codes and col_vals:
                kate_team_col = cand
                break
        if kate_team_col:
            self.doctor_kate_team = {d: get_str(d, kate_team_col) for d in self.doctor_names}
        else:
            self.doctor_kate_team = {d: "" for d in self.doctor_names}
            self._log("⚠️ カテチーム列が見つかりません")

        if "属性" in sheet4_data.columns:
            self.doctor_attribute = {d: get_str(d, "属性") for d in self.doctor_names}
        else:
            self.doctor_attribute = {d: "" for d in self.doctor_names}

        travel_col = next((c for c in ["出張日", "出張曜日"] if c in sheet4_data.columns), None)
        if travel_col:
            self.doctor_travel_day = {d: get_str(d, travel_col) for d in self.doctor_names}
        else:
            self.doctor_travel_day = {d: "" for d in self.doctor_names}

        self.wed_forbidden = {normalize_name(d) for d in WED_FORBIDDEN_DEFAULT}

        # --- 枠の抽出 ---
        self.slots = []  # dict: ridx,hosp,hidx,date,fixed,doc
        self.preassigned_count = {d: 0 for d in self.doctor_names}
        for ridx in shift_df.index:
            d = shift_df.at[ridx, self.date_col]
            if pd.isna(d):
                continue
            d = norm_date(d)
            for hosp in self.hospital_cols:
                val = shift_df.at[ridx, hosp]
                val_str = normalize_name(val) if isinstance(val, str) else ""
                if val_str in self.doctor_names:
                    self.slots.append(dict(ridx=ridx, hosp=hosp, hidx=self.hosp_idx[hosp],
                                           date=d, fixed=True, doc=val_str))
                    self.preassigned_count[val_str] += 1
                elif is_slot_value(val):
                    self.slots.append(dict(ridx=ridx, hosp=hosp, hidx=self.hosp_idx[hosp],
                                           date=d, fixed=False, doc=None))
        self.total_slots = len(self.slots)
        self.all_shift_dates = sorted(
            {norm_date(d) for d in shift_df[self.date_col].dropna()})

        # --- 医師分類 ---
        self.inactive_doctors = [d for d in self.doctor_names
                                 if self._is_always_unavailable(d)]
        self.active_doctors = [d for d in self.doctor_names
                               if d not in self.inactive_doctors]
        if not self.active_doctors:
            raise ValueError("当月に割り当て可能な医師がいません")

        self.code2_doctors = {d for d in self.doctor_names
                              if any(self.avail_code(dt, d) == 2 for dt in self.all_shift_dates)}
        self.code12_doctors = {d for d in self.doctor_names
                               if any(self.avail_code(dt, d) == 1.2 for dt in self.all_shift_dates)}

        # カテ保有医師（Sheet4カテチーム属性がSheet1:Zのコードと一致）
        self.schedule_code_holders = set()
        for d in self.doctor_names:
            team = self.doctor_kate_team.get(d, "")
            if not team:
                continue
            if self.expected_team_codes:
                if team in self.expected_team_codes:
                    self.schedule_code_holders.add(d)
            elif team.upper() in ("A", "B", "C", "D", "E", "F", "G"):
                self.schedule_code_holders.add(d)
        self.no_kate_doctors = {d for d in self.doctor_names
                                if d not in self.schedule_code_holders}

        # sheet3「1」相当（新構造では属性1で代替 = SEMI-001平日緩和対象）
        self.sheet3_code1 = {d for d in self.doctor_names
                             if self.doctor_attribute.get(d, "") == "1"}

        # 比率バランス除外（Sheet2全日コード3の外病院専門医師）
        self.ratio_exempt = set()
        for d in self.doctor_names:
            col = self.avail_col_of.get(d)
            if col is None:
                continue
            vals = self.avail_df[col].dropna()
            if len(vals) > 0:
                try:
                    if all(abs(float(v) - 3.0) < 0.01 for v in vals):
                        self.ratio_exempt.add(d)
                except (TypeError, ValueError):
                    pass

        # --- TARGET_CAP（main.py と同一アルゴリズム） ---
        self._compute_target_cap()

    @staticmethod
    def _parse_sheet4(grid):
        g = grid.dropna(how="all").reset_index(drop=True)
        if len(g) == 0:
            raise ValueError("医師情報シートが空です")
        header_row = None
        for i in range(min(50, len(g))):
            row = g.iloc[i].astype(str).str.strip()
            if (row == "氏名").any():
                header_row = i
                break
        if header_row is None:
            raise ValueError("医師情報シートに '氏名' 列が見つかりません")
        headers = [safe_str(x) for x in g.iloc[header_row].tolist()]
        headers = [h if (h != "" and h.lower() != "nan") else f"Unnamed_{j}"
                   for j, h in enumerate(headers)]
        headers = make_unique(headers)
        data = g.iloc[header_row + 1:].reset_index(drop=True)
        data.columns = headers
        data["氏名"] = data["氏名"].astype(str).str.strip()
        data = data[(data["氏名"] != "") & (data["氏名"].str.lower() != "nan")].reset_index(drop=True)
        return data

    # ---- 可否コード（main.py get_avail_code と同一挙動） ----
    def avail_code(self, date, doc):
        date = norm_date(date)
        # 出張曜日の前日は不可（v6.5.0）
        travel = self.doctor_travel_day.get(doc, "")
        if travel and travel in WEEKDAY_MAP:
            if date.weekday() == (WEEKDAY_MAP[travel] - 1) % 7:
                return 0
        code = None
        col = self.avail_col_of.get(doc)
        if col is not None:
            try:
                v = self.avail_df.at[date, col]
                if isinstance(v, pd.Series):
                    v = v.iloc[0]
                if pd.notna(v):
                    f = float(v)
                    if abs(f - 1.2) < 0.01:
                        code = 1.2
                    elif f in (0.0, 1.0, 2.0, 3.0):
                        code = int(f)
            except (KeyError, TypeError, ValueError):
                pass
        if code is None:
            code = 1  # 空欄・解釈不能は「可」（v6.5.9仕様）
        return code

    # ---- カテ表コード（main.py get_sched_code の新構造分） ----
    def sched_code(self, date, doc):
        date = norm_date(date)
        team_on_duty = self.kate_team_by_date.get(date)
        doc_team = self.doctor_kate_team.get(doc, "")
        if team_on_duty and doc_team and team_on_duty == doc_team:
            return team_on_duty
        return None

    def is_holiday(self, date):
        return norm_date(date) in self.holidays

    def slot_is_holiday(self, date, hidx):
        """平日/休日分類（main.py recompute_stats と同一ロジック）"""
        dow = norm_date(date).weekday()
        weekday = dow < 5
        return (self.is_holiday(date) or dow >= 5
                or (weekday and hidx in (C_COL, D_COL, F_COL, G_COL)))

    def is_ch_slot(self, hidx):
        return C_COL <= hidx <= H_COL

    def is_eligible_for_ch(self, doc, date):
        if doc in self.no_kate_doctors:
            return True
        return bool(self.sched_code(date, doc))

    def _is_always_unavailable(self, doc):
        if self.preassigned_count.get(doc, 0) > 0:
            return False
        return all(self.avail_code(d, doc) == 0 for d in self.all_shift_dates)

    def max_gap3_assignments(self, doc):
        avail = sorted(d for d in self.all_shift_dates if self.avail_code(d, doc) != 0)
        cnt, last = 0, None
        for d in avail:
            if last is None or (d - last).days >= 3:
                cnt += 1
                last = d
        return cnt

    def _compute_target_cap(self):
        active = self.active_doctors
        self.base_target = self.total_slots // len(active)
        self.extra_slots = self.total_slots - self.base_target * len(active)
        active_sorted = sorted(active, key=lambda d: self.doctor_col_index[d])

        attr1 = [d for d in active_sorted if self.doctor_attribute.get(d, "") == "1"]
        if self.extra_slots > 0 and len(attr1) >= self.extra_slots:
            extra_allowed = set(attr1[-self.extra_slots:])
        elif self.extra_slots > 0 and attr1:
            remaining = self.extra_slots - len(attr1)
            non_attr1 = [d for d in active_sorted if d not in attr1]
            extra_allowed = set(attr1) | set(non_attr1[-remaining:])
        else:
            extra_allowed = set(active_sorted[-self.extra_slots:] if self.extra_slots > 0 else [])
        self.extra_allowed = extra_allowed

        cap = {d: 0 for d in self.doctor_names}
        for d in active:
            cap[d] = self.base_target
        for d in extra_allowed:
            cap[d] = self.base_target + 1
        for d in self.doctor_names:
            if self.preassigned_count.get(d, 0) > cap.get(d, 0):
                cap[d] = self.preassigned_count[d]

        # gap>=3物理上限で切下げ
        for d in active:
            m = self.max_gap3_assignments(d)
            if m < cap[d]:
                cap[d] = m

        # 切下げ分の再配分（属性1優先→末尾=若手から）
        shortage = self.total_slots - sum(cap[d] for d in active)
        if shortage > 0:
            redist = ([d for d in active_sorted if self.doctor_attribute.get(d, "") == "1"]
                      + [d for d in active_sorted if self.doctor_attribute.get(d, "") != "1"])
            for d in reversed(redist):
                if shortage <= 0:
                    break
                if cap[d] < self.max_gap3_assignments(d):
                    cap[d] += 1
                    shortage -= 1
            if shortage > 0:
                self._log(f"⚠️ {shortage}枠の再配分先なし（全医師がgap3上限）")
        self.target_cap = cap

    # ---- 枠に対する静的な配置可否（ABS-001〜005/013/015 + SEMI属性2） ----
    def slot_allowed(self, slot, doc):
        """理由文字列 or None（None=配置可）"""
        date, hidx = slot["date"], slot["hidx"]
        if doc in self.inactive_doctors:
            return "inactive"
        code = self.avail_code(date, doc)
        if code == 0:
            return "ABS-001"
        if code == 2 and not (B_COL <= hidx <= Q_COL):
            return "ABS-002"
        if code == 3 and not (L_COL <= hidx <= self.ly_end):
            return "ABS-003"
        if L_COL <= hidx <= self.ly_end:
            if self.sched_code(date, doc):
                return "ABS-004"
            if norm_date(date).weekday() == 2 and doc in self.wed_forbidden:
                return "ABS-005"
        if self.is_ch_slot(hidx) and not self.is_eligible_for_ch(doc, date):
            return "ABS-013"
        # B列 × カテ保有 × カテ表なし: 属性2は禁止（ABS-015/SEMI-001緩和不可）
        if (hidx == B_COL and doc in self.schedule_code_holders
                and not self.sched_code(date, doc)
                and self.doctor_attribute.get(doc, "") == "2"):
            return "ABS-015"
        return None

    def is_semi001_relax_slot(self, slot, doc):
        """B列×カテ保有×カテ表なし×(属性2以外) = SEMI-001緩和使用となる配置か"""
        if slot["hidx"] != B_COL:
            return False
        if doc not in self.schedule_code_holders:
            return False
        if self.sched_code(slot["date"], doc):
            return False
        return self.doctor_attribute.get(doc, "") != "2"

    def hard_exempt(self, doc):
        """HARD-001/002の緩和対象（カテ保有×sheet3「1」相当）か"""
        return doc in self.schedule_code_holders and doc in self.sheet3_code1


def monday_week_start(date):
    d = norm_date(date)
    return d - pd.Timedelta(days=d.weekday())


# =========================
# CP-SATモデル構築・求解
# =========================
class CpSatScheduler:
    W_SEMI = 10 ** 7      # 第1優先: SEMI-001違反数
    W_FAIR = 10 ** 5      # 第2優先: 月内公平性 max-min
    W_SOFT_HT0 = 300      # SOFT: 外病院0回
    W_SOFT_BGHT = 100     # SOFT: BG/HT差3以上（超過分×100）
    W_SOFT_C12 = 150      # SOFT: コード1.2医師の大学0回

    def __init__(self, data: InputData):
        self.data = data
        self.model = cp_model.CpModel()
        self.x = {}          # (slot_index, doc) -> BoolVar
        self.free_idx = []   # 自由枠のslot index
        self._build()

    def _build(self):
        data = self.data
        m = self.model

        # --- 変数生成（静的ABSフィルタ通過分のみ） ---
        for si, slot in enumerate(data.slots):
            if slot["fixed"]:
                continue
            self.free_idx.append(si)
            allowed = [d for d in data.active_doctors
                       if data.slot_allowed(slot, d) is None]
            if not allowed:
                raise RuntimeError(
                    f"候補医師ゼロの枠: {slot['date'].date()} {slot['hosp']}")
            for d in allowed:
                self.x[(si, d)] = m.NewBoolVar(f"x_{si}_{d}")
            # ABS-009: 全枠割当（各枠ちょうど1人）
            m.AddExactlyOne(self.x[(si, d)] for d in allowed)

        docs = data.active_doctors

        # --- 医師×日 の割当集計 ---
        self.day_terms = {d: defaultdict(list) for d in docs}   # 変数
        self.day_fixed = {d: defaultdict(int) for d in docs}    # 固定割当数
        self.bg_day_terms = {d: defaultdict(list) for d in docs}
        self.bg_day_fixed = {d: defaultdict(int) for d in docs}
        cnt_terms = {d: [] for d in docs}
        cnt_fixed = {d: 0 for d in docs}
        bg_terms = {d: [] for d in docs}
        bg_fixed = {d: 0 for d in docs}
        ht_terms = {d: [] for d in docs}
        ht_fixed = {d: 0 for d in docs}
        wd_terms = {d: [] for d in docs}
        wd_fixed = {d: 0 for d in docs}
        we_terms = {d: [] for d in docs}
        we_fixed = {d: 0 for d in docs}
        bi_terms = {d: [] for d in docs}
        bi_fixed = {d: 0 for d in docs}
        chjk_terms = {d: [] for d in docs}
        chjk_fixed = {d: 0 for d in docs}
        hosp_terms = {d: defaultdict(list) for d in docs}
        hosp_fixed = {d: defaultdict(int) for d in docs}
        semi_terms = {d: defaultdict(list) for d in docs}  # doc -> week -> vars
        all_semi_vars = []

        for si, slot in enumerate(data.slots):
            date, hidx, hosp = slot["date"], slot["hidx"], slot["hosp"]
            is_bg = B_COL <= hidx <= K_COL
            is_ht = L_COL <= hidx <= data.ly_end
            is_hol = data.slot_is_holiday(date, hidx)
            is_bi = hidx in (B_COL, I_COL)
            is_chjk = (C_COL <= hidx <= H_COL) or (J_COL <= hidx <= K_COL)
            if slot["fixed"]:
                d = slot["doc"]
                if d not in data.active_doctors:
                    continue
                self.day_fixed[d][date] += 1
                cnt_fixed[d] += 1
                hosp_fixed[d][hosp] += 1
                if is_bg:
                    bg_fixed[d] += 1
                    self.bg_day_fixed[d][date] += 1
                if is_ht:
                    ht_fixed[d] += 1
                if is_hol:
                    we_fixed[d] += 1
                else:
                    wd_fixed[d] += 1
                if is_bi:
                    bi_fixed[d] += 1
                if is_chjk:
                    chjk_fixed[d] += 1
            else:
                for d in docs:
                    v = self.x.get((si, d))
                    if v is None:
                        continue
                    self.day_terms[d][date].append(v)
                    cnt_terms[d].append(v)
                    hosp_terms[d][hosp].append(v)
                    if is_bg:
                        bg_terms[d].append(v)
                        self.bg_day_terms[d][date].append(v)
                    if is_ht:
                        ht_terms[d].append(v)
                    if is_hol:
                        we_terms[d].append(v)
                    else:
                        wd_terms[d].append(v)
                    if is_bi:
                        bi_terms[d].append(v)
                    if is_chjk:
                        chjk_terms[d].append(v)
                    if data.is_semi001_relax_slot(slot, d):
                        semi_terms[d][monday_week_start(date)].append(v)
                        all_semi_vars.append(v)

        # --- ABS-006: 同日重複禁止 ---
        for d in docs:
            for date, terms in self.day_terms[d].items():
                fx = self.day_fixed[d].get(date, 0)
                if fx >= 1:
                    for v in terms:
                        self.model.Add(v == 0)
                elif len(terms) > 1:
                    m.AddAtMostOne(terms)

        # --- ABS-007: gap>=3（日数差1,2の日ペアで排他） ---
        dates = data.all_shift_dates
        for d in docs:
            active_dates = [dt for dt in dates
                            if self.day_terms[d].get(dt) or self.day_fixed[d].get(dt, 0)]
            for i in range(len(active_dates)):
                for j in range(i + 1, len(active_dates)):
                    diff = (active_dates[j] - active_dates[i]).days
                    if diff >= 3:
                        break
                    t1 = self.day_terms[d].get(active_dates[i], [])
                    t2 = self.day_terms[d].get(active_dates[j], [])
                    f = (self.day_fixed[d].get(active_dates[i], 0)
                         + self.day_fixed[d].get(active_dates[j], 0))
                    if f >= 2:
                        raise RuntimeError(f"固定割当がgap<3で矛盾: {d}")
                    m.Add(sum(t1) + sum(t2) + f <= 1)

        # --- ABS-008: 同一病院重複禁止 ---
        for d in docs:
            for hosp, terms in hosp_terms[d].items():
                fx = hosp_fixed[d].get(hosp, 0)
                if fx >= 1:
                    for v in terms:
                        m.Add(v == 0)
                elif len(terms) > 1:
                    m.AddAtMostOne(terms)

        # --- ABS-010: TARGET_CAP ---
        self.count_expr = {}
        for d in docs:
            cnt = m.NewIntVar(0, data.target_cap.get(d, 0), f"cnt_{d}")
            m.Add(cnt == sum(cnt_terms[d]) + cnt_fixed[d])
            self.count_expr[d] = cnt

        # --- ABS-011: 大学系2回まで ---
        for d in docs:
            m.Add(sum(bg_terms[d]) + bg_fixed[d] <= 2)

        # --- ABS-012: 大学系7日間隔 ---
        for d in docs:
            bg_dates = [dt for dt in dates
                        if self.bg_day_terms[d].get(dt) or self.bg_day_fixed[d].get(dt, 0)]
            for i in range(len(bg_dates)):
                for j in range(i + 1, len(bg_dates)):
                    diff = (bg_dates[j] - bg_dates[i]).days
                    if diff >= 7:
                        break
                    t1 = self.bg_day_terms[d].get(bg_dates[i], [])
                    t2 = self.bg_day_terms[d].get(bg_dates[j], [])
                    f = (self.bg_day_fixed[d].get(bg_dates[i], 0)
                         + self.bg_day_fixed[d].get(bg_dates[j], 0))
                    if f >= 2:
                        raise RuntimeError(f"固定割当がBG gap<7で矛盾: {d}")
                    m.Add(sum(t1) + sum(t2) + f <= 1)

        # --- ABS-014: 平日/休日偏り差<=1 ---
        for d in docs:
            wd = sum(wd_terms[d]) + wd_fixed[d]
            we = sum(we_terms[d]) + we_fixed[d]
            m.Add(wd - we <= 1)
            m.Add(we - wd <= 1)

        # --- HARD-001/002（カテ保有×属性1は緩和対象＝制約なし） ---
        for d in docs:
            if data.hard_exempt(d):
                continue
            m.Add(sum(bi_terms[d]) + bi_fixed[d] <= 1)        # HARD-001
            m.Add(sum(chjk_terms[d]) + chjk_fixed[d] <= 1)    # HARD-002

        # --- SEMI-001: 週1回まで（月曜始まり週）＋違反数を目的で最小化 ---
        for d in docs:
            for week, terms in semi_terms[d].items():
                if len(terms) > 1:
                    m.AddAtMostOne(terms)
        self.semi_total = m.NewIntVar(0, len(all_semi_vars) or 1, "semi_total")
        m.Add(self.semi_total == (sum(all_semi_vars) if all_semi_vars else 0))

        # --- 公平性: max-min ---
        cap_max = max(data.target_cap.get(d, 0) for d in docs)
        t_max = m.NewIntVar(0, cap_max, "t_max")
        t_min = m.NewIntVar(0, cap_max, "t_min")
        for d in docs:
            m.Add(t_max >= self.count_expr[d])
            m.Add(t_min <= self.count_expr[d])
        self.fair_span = m.NewIntVar(0, cap_max, "fair_span")
        m.Add(self.fair_span == t_max - t_min)

        # --- SOFT簡易版 ---
        soft_terms = []
        for d in docs:
            # SOFT: 外病院0回（外病院に入りうる医師のみ）
            if ht_terms[d] or ht_fixed[d]:
                h0 = m.NewBoolVar(f"ht0_{d}")
                m.Add(sum(ht_terms[d]) + ht_fixed[d] >= 1 - h0)
                soft_terms.append(self.W_SOFT_HT0 * h0)
            # SOFT: BG/HT差3以上（比率除外医師以外、超過分ペナルティ）
            if d not in data.ratio_exempt:
                bg = sum(bg_terms[d]) + bg_fixed[d]
                ht = sum(ht_terms[d]) + ht_fixed[d]
                diff = m.NewIntVar(0, cap_max, f"bght_{d}")
                m.Add(diff >= bg - ht)
                m.Add(diff >= ht - bg)
                over = m.NewIntVar(0, cap_max, f"bghtov_{d}")
                m.Add(over >= diff - 2)
                soft_terms.append(self.W_SOFT_BGHT * over)
            # SOFT: コード1.2医師の大学0回
            if d in data.code12_doctors and (bg_terms[d] or bg_fixed[d]):
                b0 = m.NewBoolVar(f"bg0_{d}")
                m.Add(sum(bg_terms[d]) + bg_fixed[d] >= 1 - b0)
                soft_terms.append(self.W_SOFT_C12 * b0)

        m.Minimize(self.W_SEMI * self.semi_total
                   + self.W_FAIR * self.fair_span
                   + sum(soft_terms))

    def solve_multi(self, n_solutions=3, time_limit=60, min_diff=6):
        """逐次求解＋Hamming距離制約で多様な解をn個得る"""
        solutions = []
        for k in range(n_solutions):
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit
            solver.parameters.num_workers = 8
            solver.parameters.random_seed = 42 + k
            t0 = time.time()
            status = solver.Solve(self.model)
            elapsed = time.time() - t0
            if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                if k == 0:
                    raise RuntimeError(
                        f"解が見つかりません: {solver.StatusName(status)}")
                print(f"   pattern_{k+1:02d}: 追加解なし"
                      f"（{solver.StatusName(status)}、多様性制約min_diff={min_diff}）")
                break
            assign = {}   # slot_index -> doc
            sel_vars = []
            for (si, d), v in self.x.items():
                if solver.Value(v):
                    assign[si] = d
                    sel_vars.append(v)
            solutions.append(dict(
                assign=assign,
                status=solver.StatusName(status),
                objective=solver.ObjectiveValue(),
                semi=solver.Value(self.semi_total),
                fair_span=solver.Value(self.fair_span),
                time=elapsed,
            ))
            # 次解: 今回の割当と最低min_diff枠は異なること
            if k < n_solutions - 1:
                self.model.Add(sum(sel_vars) <= len(self.free_idx) - min_diff)
        return solutions


# =========================
# ライブラリAPI（main.py 本統合用・v6.9.0）
# =========================
def solve(input_data, n_solutions=3, min_diff=6, time_limit=60, verbose=False):
    """入力Excelをパースしてn個のCP-SAT厳密解を返すライブラリ関数。

    CLI（main()）とは独立に main.py から import して利用する本統合用の公開API。

    Args:
        input_data: 入力Excelのパス（str）。既存 InputData でパースする。
        n_solutions: 生成する解の数（既定3）。
        min_diff: 解間の最小Hamming距離（枠単位・既定6）。
        time_limit: 1解あたりの求解秒数上限（既定60）。
        verbose: InputDataパースのログ出力可否。

    Returns:
        (data, solutions):
            data      … InputData（TARGET_CAP・EXTRA等の突合に利用可）
            solutions … list[dict]。各要素は solve_multi の戻り
                        （assign=slot_index->doc, status, objective, semi,
                          fair_span, time）。

    Raises:
        ImportError   … ortools未導入（モジュールimport時点で送出されうる）
        RuntimeError  … INFEASIBLE等で1解も得られない場合
        その他例外    … パース失敗等。呼び出し側でフォールバック判定に用いる。
    """
    data = InputData(input_data, verbose=verbose)
    sched = CpSatScheduler(data)
    solutions = sched.solve_multi(
        n_solutions=n_solutions, time_limit=time_limit, min_diff=min_diff)
    return data, solutions


# =========================
# 出力Excel生成
# =========================
def build_pattern_df(data: InputData, assign):
    """入力sheet1と同型のグリッドに医師名を書き込んだDataFrameを返す"""
    out = data.shift_df.copy()
    for hosp in data.hospital_cols:
        out[hosp] = out[hosp].astype(object)
    for si, slot in enumerate(data.slots):
        if slot["fixed"]:
            out.at[slot["ridx"], slot["hosp"]] = slot["doc"]
        else:
            out.at[slot["ridx"], slot["hosp"]] = assign.get(si, "UNASSIGNED")
    out[data.date_col] = pd.to_datetime(out[data.date_col]).dt.strftime("%Y-%m-%d")
    return out


def build_summary_df(data: InputData, assign):
    stats = {d: dict(total=0, bg=0, ht=0, wd=0, we=0, semi=0) for d in data.doctor_names}
    for si, slot in enumerate(data.slots):
        d = slot["doc"] if slot["fixed"] else assign.get(si)
        if d not in stats:
            continue
        s = stats[d]
        s["total"] += 1
        hidx = slot["hidx"]
        if B_COL <= hidx <= K_COL:
            s["bg"] += 1
        elif L_COL <= hidx <= data.ly_end:
            s["ht"] += 1
        if data.slot_is_holiday(slot["date"], hidx):
            s["we"] += 1
        else:
            s["wd"] += 1
        if not slot["fixed"] and data.is_semi001_relax_slot(slot, d):
            s["semi"] += 1
    rows = []
    for d in data.doctor_names:
        s = stats[d]
        rows.append({
            "氏名": d,
            "割当計": s["total"],
            "大学(B-K)": s["bg"],
            "外病院(L-Y)": s["ht"],
            "平日": s["wd"],
            "休日": s["we"],
            "TARGET_CAP": data.target_cap.get(d, 0),
            "EXTRA(+1)": "◯" if d in data.extra_allowed else "",
            "属性": data.doctor_attribute.get(d, ""),
            "カテ班": data.doctor_kate_team.get(d, ""),
            "SEMI-001緩和": s["semi"],
            "状態": "inactive" if d in data.inactive_doctors else "active",
        })
    return pd.DataFrame(rows)


# =========================
# 独立チェッカー（出力xlsxを読み直して全制約を再検証）
# =========================
def verify_output(input_path, output_path, pattern_sheets):
    """出力Excelのpatternシートを読み直し、入力から再構築した制約で全検証。
    solverの内部変数は一切使わない。
    Returns: {sheet_name: [violation dicts]}, results_df
    """
    data = InputData(input_path, verbose=False)  # 入力を再パース（独立再構築）
    all_results = {}
    rows = []

    check_ids = [
        "ABS-001", "ABS-002", "ABS-003", "ABS-004", "ABS-005", "ABS-006",
        "ABS-007", "ABS-008", "ABS-009", "ABS-010", "ABS-011", "ABS-012",
        "ABS-013", "ABS-014", "ABS-015", "HARD-001", "HARD-002", "SEMI-001週1",
    ]
    desc = {
        "ABS-001": "可否コード0禁止", "ABS-002": "コード2はB-Q列のみ",
        "ABS-003": "コード3はL-Y列のみ", "ABS-004": "カテ当番日の外病院禁止",
        "ABS-005": "水曜L-Y禁止医師", "ABS-006": "同日重複禁止",
        "ABS-007": "gap>=3日", "ABS-008": "同一病院重複禁止",
        "ABS-009": "全枠割当（未割当/不明医師なし）", "ABS-010": "TARGET_CAP厳守",
        "ABS-011": "大学系2回まで", "ABS-012": "大学系7日間隔",
        "ABS-013": "C-H列カテ当番必須", "ABS-014": "平日/休日差<=1",
        "ABS-015": "属性2のB列カテ表必須", "HARD-001": "B/I列1回まで",
        "HARD-002": "C-H/J-K列1回まで", "SEMI-001週1": "B列カテ緩和は週1回まで",
    }

    for sheet in pattern_sheets:
        out_df = pd.read_excel(output_path, sheet_name=sheet)
        out_df.columns = make_unique(
            [c.strip() if isinstance(c, str) else c for c in out_df.columns])
        out_dates = pd.to_datetime(out_df[out_df.columns[0]], errors="coerce")
        viols = []

        # 出力グリッドから割当を復元（行位置＋日付一致で照合）
        assignments = []  # (slot, doc)
        for slot in data.slots:
            ridx = slot["ridx"]
            od = out_dates.iloc[ridx]
            if pd.isna(od) or norm_date(od) != slot["date"]:
                viols.append(("ABS-009", f"行{ridx}: 出力の日付が入力と不一致"))
                continue
            val = out_df.at[ridx, slot["hosp"]]
            doc = normalize_name(val) if isinstance(val, str) else ""
            if doc not in data.doctor_names:
                viols.append(("ABS-009",
                              f"{slot['date'].date()} {slot['hosp']}: 未割当/不明 ({val!r})"))
                continue
            assignments.append((slot, doc))

        # 集計
        day_count = defaultdict(int)          # (doc,date)
        hosp_count = defaultdict(int)         # (doc,hosp)
        counts = defaultdict(int)
        bg_counts = defaultdict(int)
        wd_counts = defaultdict(int)
        we_counts = defaultdict(int)
        bi_counts = defaultdict(int)
        chjk_counts = defaultdict(int)
        doc_dates = defaultdict(set)
        bg_dates = defaultdict(set)
        semi_week = defaultdict(int)          # (doc,week)

        for slot, doc in assignments:
            date, hidx = slot["date"], slot["hidx"]
            code = data.avail_code(date, doc)
            # ABS-001〜003
            if code == 0:
                viols.append(("ABS-001", f"{doc} {date.date()} {slot['hosp']}"))
            if code == 2 and not (B_COL <= hidx <= Q_COL):
                viols.append(("ABS-002", f"{doc} {date.date()} {slot['hosp']}"))
            if code == 3 and not (L_COL <= hidx <= data.ly_end):
                viols.append(("ABS-003", f"{doc} {date.date()} {slot['hosp']}"))
            # ABS-004/005（外病院）
            if L_COL <= hidx <= data.ly_end:
                if data.sched_code(date, doc):
                    viols.append(("ABS-004", f"{doc} {date.date()} {slot['hosp']}"))
                if date.weekday() == 2 and doc in data.wed_forbidden:
                    viols.append(("ABS-005", f"{doc} {date.date()} {slot['hosp']}"))
            # ABS-013（固定割当は許容: main.py validateと同一）
            if (not slot["fixed"] and data.is_ch_slot(hidx)
                    and not data.is_eligible_for_ch(doc, date)):
                viols.append(("ABS-013", f"{doc} {date.date()} {slot['hosp']}"))
            # ABS-015（main.py validateと同一: EXTRA対象は除外）
            if (hidx == B_COL and doc in data.schedule_code_holders
                    and data.doctor_attribute.get(doc, "") == "2"
                    and not data.sched_code(date, doc)
                    and doc not in data.extra_allowed):
                viols.append(("ABS-015", f"{doc} {date.date()} {slot['hosp']}"))
            # SEMI-001週1
            if not slot["fixed"] and data.is_semi001_relax_slot(slot, doc):
                semi_week[(doc, monday_week_start(date))] += 1

            day_count[(doc, date)] += 1
            hosp_count[(doc, slot["hosp"])] += 1
            counts[doc] += 1
            doc_dates[doc].add(date)
            if B_COL <= hidx <= K_COL:
                bg_counts[doc] += 1
                bg_dates[doc].add(date)
            if data.slot_is_holiday(date, hidx):
                we_counts[doc] += 1
            else:
                wd_counts[doc] += 1
            if hidx in (B_COL, I_COL):
                bi_counts[doc] += 1
            if (C_COL <= hidx <= H_COL) or (J_COL <= hidx <= K_COL):
                chjk_counts[doc] += 1

        # ABS-006
        for (doc, date), c in day_count.items():
            if c > 1:
                viols.append(("ABS-006", f"{doc} {date.date()} ({c}回)"))
        # ABS-007
        for doc, ds in doc_dates.items():
            sd = sorted(ds)
            for i in range(1, len(sd)):
                gap = (sd[i] - sd[i - 1]).days
                if gap < 3:
                    viols.append(("ABS-007", f"{doc} gap={gap}日 ({sd[i-1].date()}→{sd[i].date()})"))
        # ABS-008
        for (doc, hosp), c in hosp_count.items():
            if c > 1:
                viols.append(("ABS-008", f"{doc} {hosp} ({c}回)"))
        # ABS-010
        for doc, c in counts.items():
            if c > data.target_cap.get(doc, 0):
                viols.append(("ABS-010", f"{doc} {c}回 (上限{data.target_cap.get(doc, 0)})"))
        # ABS-011
        for doc, c in bg_counts.items():
            if c > 2:
                viols.append(("ABS-011", f"{doc} 大学{c}回"))
        # ABS-012
        for doc, ds in bg_dates.items():
            sd = sorted(ds)
            for i in range(1, len(sd)):
                gap = (sd[i] - sd[i - 1]).days
                if gap < 7:
                    viols.append(("ABS-012", f"{doc} BG gap={gap}日"))
        # ABS-014
        for doc in data.active_doctors:
            diff = abs(wd_counts.get(doc, 0) - we_counts.get(doc, 0))
            if diff >= 2:
                viols.append(("ABS-014",
                              f"{doc} 平日{wd_counts.get(doc,0)}/休日{we_counts.get(doc,0)}"))
        # HARD-001/002
        for doc in data.active_doctors:
            if data.hard_exempt(doc):
                continue
            if bi_counts.get(doc, 0) > 1:
                viols.append(("HARD-001", f"{doc} B/I列{bi_counts[doc]}回"))
            if chjk_counts.get(doc, 0) > 1:
                viols.append(("HARD-002", f"{doc} C-H/J-K列{chjk_counts[doc]}回"))
        # SEMI-001 週1回制限
        for (doc, week), c in semi_week.items():
            if c > 1:
                viols.append(("SEMI-001週1", f"{doc} 週{week.date()} {c}回"))

        all_results[sheet] = viols
        by_id = defaultdict(list)
        for cid, detail in viols:
            by_id[cid].append(detail)
        semi_used = sum(semi_week.values())
        for cid in check_ids:
            n = len(by_id.get(cid, []))
            rows.append({
                "パターン": sheet,
                "制約ID": cid,
                "内容": desc[cid],
                "違反数": n,
                "判定": "✅ OK" if n == 0 else "❌ NG",
                "詳細": "; ".join(by_id.get(cid, [])[:5]),
            })
        rows.append({
            "パターン": sheet, "制約ID": "SEMI-001使用数", "内容": "B列カテ緩和の使用回数（参考・違反ではない）",
            "違反数": semi_used, "判定": "参考", "詳細": "",
        })

    return all_results, pd.DataFrame(rows)


# =========================
# メイン
# =========================
def main():
    ap = argparse.ArgumentParser(description="当直くん CP-SAT厳密解ソルバー（併走版）")
    ap.add_argument("input", help="入力Excel（sheet1/Sheet2/Sheet3構造）")
    ap.add_argument("--output-dir", default=os.path.expanduser("~/Downloads"))
    ap.add_argument("--patterns", type=int, default=3, help="出力パターン数（既定3）")
    ap.add_argument("--time-limit", type=float, default=60, help="1解あたりの秒数上限")
    ap.add_argument("--min-diff", type=int, default=6,
                    help="パターン間で最低限異なる枠数（Hamming距離）")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        print(f"エラー: 入力ファイルが見つかりません: {args.input}")
        sys.exit(1)

    print("=" * 60)
    print(f"  当直くん CP-SAT併走ソルバー ({SOLVER_VERSION})")
    print("=" * 60)
    t_start = time.time()

    data = InputData(args.input)
    print(f"✅ 読込: 医師{len(data.doctor_names)}人"
          f"（active {len(data.active_doctors)} / inactive {len(data.inactive_doctors)}）"
          f" | 枠{data.total_slots}（固定{sum(1 for s in data.slots if s['fixed'])}）")
    extra_names = sorted(data.extra_allowed, key=lambda d: data.doctor_col_index[d])
    print(f"   BASE_TARGET={data.base_target} EXTRA_SLOTS={data.extra_slots}"
          f" | EXTRA(+1回): {', '.join(extra_names) or 'なし'}")

    print("\n🔧 CP-SATモデル構築中...")
    sched = CpSatScheduler(data)
    n_vars = len(sched.x)
    print(f"   変数: {n_vars} | 自由枠: {len(sched.free_idx)}")

    print(f"\n🚀 求解（{args.patterns}パターン、Hamming最小差={args.min_diff}枠）...")
    solutions = sched.solve_multi(
        n_solutions=args.patterns, time_limit=args.time_limit, min_diff=args.min_diff)
    for i, sol in enumerate(solutions):
        print(f"   pattern_{i+1:02d}: {sol['status']} obj={sol['objective']:.0f}"
              f" SEMI違反={sol['semi']} 公平span={sol['fair_span']}"
              f" ({sol['time']:.2f}秒)")

    # --- 出力 ---
    base = os.path.splitext(os.path.basename(args.input))[0]
    os.makedirs(args.output_dir, exist_ok=True)
    out_path = os.path.join(args.output_dir, f"{base}_{SOLVER_VERSION}.xlsx")

    pattern_sheets = []
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for i, sol in enumerate(solutions):
            psheet = f"pattern_{i+1:02d}"
            pattern_sheets.append(psheet)
            build_pattern_df(data, sol["assign"]).to_excel(writer, sheet_name=psheet, index=False)
            build_summary_df(data, sol["assign"]).to_excel(
                writer, sheet_name=f"summary_{i+1:02d}", index=False)
    print(f"\n💾 出力: {out_path}")

    # --- 独立チェッカー（出力xlsxを読み直して全制約再検証） ---
    print("\n🔍 独立検証（出力Excelを読み直して再チェック）...")
    all_viols, results_df = verify_output(args.input, out_path, pattern_sheets)

    # 検証シートを追記
    from openpyxl import load_workbook
    wb = load_workbook(out_path)
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as writer:
        results_df.to_excel(writer, sheet_name="検証", index=False)

    total_bad = 0
    for sheet, viols in all_viols.items():
        n = len(viols)
        total_bad += n
        mark = "✅ 違反0" if n == 0 else f"❌ 違反{n}件"
        print(f"   {sheet}: {mark}")
        for cid, detail in viols[:10]:
            print(f"      - [{cid}] {detail}")

    # --- レポート ---
    elapsed = time.time() - t_start
    print(f"\n⏱ 総実行時間: {elapsed:.1f}秒")
    if solutions:
        summ = build_summary_df(data, solutions[0]["assign"])
        act = summ[summ["状態"] == "active"]
        dist = act["割当計"].value_counts().sort_index()
        print(f"📊 pattern_01 医師別割当数分布: "
              + ", ".join(f"{k}回×{v}人" for k, v in dist.items()))
    print(f"👥 EXTRA(+1回)対象: {', '.join(extra_names) or 'なし'}")
    sys.exit(0 if total_bad == 0 else 2)


if __name__ == "__main__":
    main()
