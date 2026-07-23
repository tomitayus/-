#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_wish_from_form.py — Google Form の希望回答を当直くん入力に反映する。

Google Form（`docs/WISH_FORM_DESIGN.md`）で集めた各医師の希望を読み、当月の
`入力.xlsx` に

  - **当直できない日** → `Sheet2`（可否）の該当セル = `0`（絶対不可・ハード）
  - **避けたい日**     → 「希望」シートの該当セル = `×N`（既定 ×2）
  - **やりたい日**     → 「希望」シートの該当セル = `○N`（既定 ○2）

を書き込んだコピーを出力する。「希望」シートが無ければ `Sheet2` と同形で新規作成。
反映ロジック（×N/○N の意味・予算正規化）は `docs/WISH_CONSTRAINT_SPEC.md`。

- v1（Option A）: 強度なし＝避け一律 `×2`・やりたい一律 `○2`。件数だけで自己抑制が
  効く（予算を選んだ日数で割るため）。`--avoid-stage/--want-stage` で既定段階を変更可。
- 同一日で「できない」と「避け/やりたい」が重複したら **できない(0)が優先**（×/○は書かない）。
- 氏名は `normalize_name`（空白除去）+ 完全一致で名簿（Sheet2の列）に突合。不一致は警告。
- 未回答者は空欄のまま（＝可・希望なし）。回答者/未回答者を一覧表示。

CLI:
  python3 build_wish_from_form.py <form_responses.csv|xlsx> <入力.xlsx> -o <出力.xlsx>
      [--avoid-stage 2] [--want-stage 2]

注意: 本スクリプトは main.py / solver_cpsat.py に依存しない（独立実装）。
氏名正規化・シート探索は prepare_next_month の共通ヘルパを再利用する。
"""

import argparse
import os
import re
import sys

import openpyxl
import pandas as pd

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import prepare_next_month as pnm  # normalize_name / _find_sheet / _find_header_row / _parse_grid_date

# 回答列をヘッダ文字列のキーワードで特定（Google Forms は質問文がそのまま列名）
NAME_KEYS = ("氏名", "名前", "お名前")
CANT_KEYS = ("できない", "当直不可", "不可", "休み")
AVOID_KEYS = ("避け",)
WANT_KEYS = ("入りたい", "やりたい")


# ----------------------------------------------------------------------------
# 回答パース（純粋関数）
# ----------------------------------------------------------------------------
def find_col(columns, keys):
    """列名にキーワードを含む最初の列を返す（無ければ None）。"""
    for c in columns:
        s = str(c)
        if any(k in s for k in keys):
            return c
    return None


def parse_labels(cell):
    """チェックボックス回答セル → 日付ラベルのリスト（カンマ/読点/改行区切り）。"""
    if cell is None:
        return []
    if isinstance(cell, float) and pd.isna(cell):
        return []
    s = str(cell).strip()
    if not s or s.lower() == "nan":
        return []
    parts = re.split(r"[,、;；\n]+", s)
    return [p.strip() for p in parts if p.strip()]


def label_to_date(label, by_ymd, by_md, by_d):
    """日付ラベル('9/1(火)'・'2026/9/1'・'1'等) → 入力の該当 date。無ければ None。

    数字を抽出し、3個→(年,月,日)・2個→(月,日)・1個→(日)で入力の日付集合に照合する。
    """
    nums = re.findall(r"\d+", str(label))
    if len(nums) >= 3:
        return by_ymd.get((int(nums[0]), int(nums[1]), int(nums[2])))
    if len(nums) == 2:
        return by_md.get((int(nums[0]), int(nums[1])))
    if len(nums) == 1:
        return by_d.get(int(nums[0]))
    return None


# ----------------------------------------------------------------------------
# 入力xlsx の索引（Sheet2）
# ----------------------------------------------------------------------------
def index_sheet2(ws):
    """Sheet2 → (date→row, 正規化医師名→col, dates一覧, header_row, date_col_1based)。"""
    header_row = pnm._find_header_row(ws)
    doc_cols = {}
    for c in range(2, ws.max_column + 1):
        h = ws.cell(row=header_row, column=c).value
        if h is None or str(h).strip() == "":
            continue
        doc_cols[pnm.normalize_name(h)] = c
    date_rows = {}
    dates = []
    for r in range(header_row + 1, ws.max_row + 1):
        d = pnm._parse_grid_date(ws.cell(row=r, column=1).value)
        if d is None:
            continue
        date_rows[d] = r
        dates.append(d)
    return date_rows, doc_cols, dates, header_row


def ensure_wish_sheet(wb, ws2, header_row, doc_cols, date_rows):
    """「希望」シートを取得（無ければ Sheet2 と同形で新規作成）。戻り値 (ws, date→row)。"""
    ws = pnm._find_sheet(wb, "希望")
    if ws is not None:
        wr = {}
        hr = pnm._find_header_row(ws)
        for r in range(hr + 1, ws.max_row + 1):
            d = pnm._parse_grid_date(ws.cell(row=r, column=1).value)
            if d is not None:
                wr[d] = r
        return ws, wr
    # 新規作成: Sheet2 のヘッダ（Date + 医師名）と日付列をコピー
    ws = wb.create_sheet("希望")
    for c in range(1, ws2.max_column + 1):
        ws.cell(row=1, column=c, value=ws2.cell(row=header_row, column=c).value)
    wr = {}
    for i, (d, r2) in enumerate(sorted(date_rows.items()), start=2):
        ws.cell(row=i, column=1, value=ws2.cell(row=r2, column=1).value)
        wr[d] = i
    return ws, wr


# ----------------------------------------------------------------------------
# 変換本体
# ----------------------------------------------------------------------------
def read_responses(path):
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    return pd.read_csv(path)


def build(form_path, input_path, out_path, avoid_stage=2, want_stage=2):
    """Form回答を入力.xlsxに反映したコピーを out_path に保存し、結果 dict を返す。"""
    df = read_responses(form_path)
    name_col = find_col(df.columns, NAME_KEYS)
    if name_col is None:
        raise ValueError(f"回答に氏名列が見つかりません（{NAME_KEYS} を含む列が必要）: {list(df.columns)}")
    cant_col = find_col(df.columns, CANT_KEYS)
    avoid_col = find_col(df.columns, AVOID_KEYS)
    want_col = find_col(df.columns, WANT_KEYS)

    wb = openpyxl.load_workbook(input_path)
    ws2 = pnm._find_sheet(wb, "sheet2") or pnm._find_sheet(wb, "Sheet2")
    if ws2 is None:
        raise ValueError(f"入力に Sheet2(可否) がありません: {wb.sheetnames}")
    date_rows, doc_cols, dates, header_row = index_sheet2(ws2)

    # 日付照合テーブル
    by_ymd = {(d.year, d.month, d.day): d for d in dates}
    by_md = {(d.month, d.day): d for d in dates}
    # 「日のみ」ラベルは、その日が入力期間内で一意のときだけ照合可能にする
    _day_count = {}
    for d in dates:
        _day_count[d.day] = _day_count.get(d.day, 0) + 1
    by_d = {d.day: d for d in dates if _day_count[d.day] == 1}

    wsw, wish_rows = ensure_wish_sheet(wb, ws2, header_row, doc_cols, date_rows)

    n_cant = n_avoid = n_want = 0
    unmatched_names = []
    unknown_labels = []
    seen = set()
    dup_names = []

    for _, row in df.iterrows():
        raw_name = row[name_col]
        if raw_name is None or (isinstance(raw_name, float) and pd.isna(raw_name)):
            continue
        norm = pnm.normalize_name(raw_name)
        col = doc_cols.get(norm)
        if col is None:
            unmatched_names.append(str(raw_name))
            continue
        if norm in seen:
            dup_names.append(str(raw_name))  # 後の回答で上書き
        seen.add(norm)

        def _dates(colname):
            out = []
            if colname is None:
                return out
            for lab in parse_labels(row[colname]):
                d = label_to_date(lab, by_ymd, by_md, by_d)
                if d is None:
                    unknown_labels.append((str(raw_name), lab))
                else:
                    out.append(d)
            return out

        cant = set(_dates(cant_col))
        avoid = set(_dates(avoid_col)) - cant           # できない が優先
        want = set(_dates(want_col)) - cant - avoid     # できない>避け>やりたい
        for d in cant:
            ws2.cell(row=date_rows[d], column=col, value=0)
            n_cant += 1
        for d in avoid:
            if d in wish_rows:
                wsw.cell(row=wish_rows[d], column=col, value=f"×{avoid_stage}")
                n_avoid += 1
        for d in want:
            if d in wish_rows:
                wsw.cell(row=wish_rows[d], column=col, value=f"○{want_stage}")
                n_want += 1

    responders = set(seen)
    non_responders = [d for d in doc_cols if d not in responders]

    wb.save(out_path)
    return {
        "out_path": out_path,
        "n_cant": n_cant, "n_avoid": n_avoid, "n_want": n_want,
        "unmatched_names": unmatched_names,
        "unknown_labels": unknown_labels,
        "dup_names": dup_names,
        "responders": sorted(responders),
        "non_responders": sorted(non_responders),
        "name_col": name_col, "cant_col": cant_col,
        "avoid_col": avoid_col, "want_col": want_col,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Google Form の希望回答を入力.xlsx（Sheet2:0 + 希望:×/○）に反映")
    ap.add_argument("form", help="Form回答 csv/xlsx（1行1回答者）")
    ap.add_argument("input_xlsx", help="当月の入力 xlsx（sheet1/Sheet2/Sheet3）")
    ap.add_argument("-o", "--output", required=True, help="反映後の出力先 xlsx")
    ap.add_argument("--avoid-stage", type=int, default=2, choices=(1, 2, 3),
                    help="避けたい日に入れる×段階（既定2）")
    ap.add_argument("--want-stage", type=int, default=2, choices=(1, 2, 3),
                    help="やりたい日に入れる○段階（既定2）")
    args = ap.parse_args(argv)

    print("=" * 60)
    print("  当直くん 希望取込: build_wish_from_form")
    print("=" * 60)
    res = build(args.form, args.input_xlsx, args.output,
                args.avoid_stage, args.want_stage)
    print(f"回答列: 氏名={res['name_col']!r} / できない={res['cant_col']!r} / "
          f"避け={res['avoid_col']!r} / やりたい={res['want_col']!r}")
    print(f"反映: 不可(0)={res['n_cant']} / 避け(×)={res['n_avoid']} / やりたい(○)={res['n_want']}")
    print(f"回答者 {len(res['responders'])}人 / 未回答 {len(res['non_responders'])}人: "
          + ("、".join(res["non_responders"]) if res["non_responders"] else "なし"))
    if res["unmatched_names"]:
        print(f"⚠️ 名簿に無い氏名（無視）: {res['unmatched_names']} ※Formのプルダウン/NAME_ALIASESを確認")
    if res["unknown_labels"]:
        _u = res["unknown_labels"]
        _s = "、".join(f"{n}:{lab}" for n, lab in _u[:5]) + (f" 他{len(_u) - 5}件" if len(_u) > 5 else "")
        print(f"⚠️ 日付として解釈できないラベル（無視）: {_s}")
    if res["dup_names"]:
        print(f"⚠️ 重複回答（後の回答で上書き）: {res['dup_names']}")
    print(f"出力: {res['out_path']}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
