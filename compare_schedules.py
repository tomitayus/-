#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""compare_schedules.py — 当直表2枚の突合ヘルパー（当直くん 併走検証用）

同一の雛形から作られた2枚の当直表（例: 当直くん出力の pattern_01 と、
草野先生の完成版を pattern_01 化したもの）を読み、

  (a) セル単位の相違  … 同じ日・同じ枠で誰が違うか
  (b) 医師別の回数差  … 大学(B-K) / 外病院(L-Y) / 合計 の A−B

を算出して Markdown レポートに書き出す。併走検証（ソルバー出力 vs 手作り
実データ）を数字で締めるための唯一のピース。

グリッド解釈・氏名正規化・列分類は prepare_next_month.py を再利用するため、
prepare_next_month（= main.py recompute_stats 準拠）と挙動が完全一致する。

CLI:
  python3 compare_schedules.py <A.xlsx> <B.xlsx> \
      [--pattern N] [--label-a ラベル] [--label-b ラベル] [-o 突合.md]

- 第1引数 A / 第2引数 B とも「pattern_NN シートを持つワークブック」。
  A/B は対等（差は A−B で表す）。慣例では A=当直くん出力、B=完成版。
- 完成版を B に使うときは、シート名を `pattern_01`、A1 セルを `日付`/`Date`
  にしておくこと（詳細は docs/HEISOU_GUIDE.md）。
- -o 省略時は stdout にサマリーのみ出力。
"""

import argparse
import collections
import os
import sys

import openpyxl

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

# グリッド解釈・分類ロジックは prepare_next_month に一本化（写経しない）
import prepare_next_month as pnm


# ----------------------------------------------------------------------------
# グリッド読み込み
# ----------------------------------------------------------------------------
class Grid:
    """1枚の pattern シートのグリッド表現。

    cells: {date -> {j(0-based grid index) -> (表示名, 正規化名)}}
    headers: ヘッダ行の値リスト（0-based。0=Date列）
    kate_j / last_hosp_j: collect_month_stats と同一定義の列境界
    """

    def __init__(self, sheet_name, cells, headers, kate_j, last_hosp_j):
        self.sheet_name = sheet_name
        self.cells = cells
        self.headers = headers
        self.kate_j = kate_j
        self.last_hosp_j = last_hosp_j

    @property
    def dates(self):
        return sorted(self.cells.keys())


def _resolve_pattern_sheet(wb, pattern_num):
    """collect_month_stats と同じ pattern シート解決（桁/大小揺れ救済つき）。"""
    sheet_name = f"pattern_{pattern_num:02d}"
    if sheet_name not in wb.sheetnames:
        alt = f"pattern_{pattern_num}"
        cand = [s for s in wb.sheetnames
                if s.lower() == sheet_name.lower() or s.lower() == alt.lower()]
        if not cand:
            raise ValueError(
                f"❌ '{sheet_name}' シートがありません。"
                f"存在するシート: {wb.sheetnames}"
            )
        sheet_name = cand[0]
    return wb[sheet_name], sheet_name


def _hospital_bounds(headers):
    """カテ当番列(kate_j)と病院列右端(last_hosp_j)を返す。

    collect_month_stats (prepare_next_month L205-219) と同一定義。
    """
    kate_j = None
    for j, h in enumerate(headers):
        if isinstance(h, str) and h.strip() == "カテ当番":
            kate_j = j
            break
    if kate_j is not None:
        last_hosp_j = kate_j - 1
    else:
        last_hosp_j = 0
        for j, h in enumerate(headers):
            if h is not None and str(h).strip() != "" and not str(h).startswith("【"):
                last_hosp_j = j
    return kate_j, last_hosp_j


def load_grid(path, pattern_num=1):
    """pattern_NN シートを Grid として読む。医師名セルのみ保持。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws, sheet_name = _resolve_pattern_sheet(wb, pattern_num)

    header_row = pnm._find_header_row(ws)
    headers = [ws.cell(row=header_row, column=c).value
               for c in range(1, ws.max_column + 1)]
    kate_j, last_hosp_j = _hospital_bounds(headers)

    cells = {}
    for r in range(header_row + 1, ws.max_row + 1):
        d = pnm._parse_grid_date(ws.cell(row=r, column=1).value)
        if d is None:
            continue
        row = {}
        for j in range(1, last_hosp_j + 1):
            if kate_j is not None and j == kate_j:
                continue
            val = ws.cell(row=r, column=j + 1).value  # openpyxl は 1-based
            if not isinstance(val, str):
                continue
            norm = pnm.normalize_name(val)
            if not norm or norm in pnm._NON_DOCTOR_STRINGS:
                continue
            row[j] = (val.strip(), norm)
        cells[d] = row  # 空の枠なし日も key は残す（日付カバレッジ判定用）
    return Grid(sheet_name, cells, headers, kate_j, last_hosp_j)


# ----------------------------------------------------------------------------
# (a) セル単位の相違
# ----------------------------------------------------------------------------
def _col_label(ga, gb, j):
    """列 j の表示名（A優先→B→列番号）。"""
    for g in (ga, gb):
        if j < len(g.headers):
            h = g.headers[j]
            if h is not None and str(h).strip():
                return str(h).strip()
    return f"列{j}"


def diff_cells(ga, gb):
    """同じ (日付, 列) で割当医師が異なるセルの一覧を返す。

    片側だけ割当ありも相違として拾う（一方が空欄＝欠員/余剰）。
    戻り値: [{date, weekday, col_index, col_label, a, b}] を日付・列順。
    """
    diffs = []
    all_dates = sorted(set(ga.cells) | set(gb.cells))
    for d in all_dates:
        row_a = ga.cells.get(d, {})
        row_b = gb.cells.get(d, {})
        for j in sorted(set(row_a) | set(row_b)):
            a = row_a.get(j)
            b = row_b.get(j)
            na = a[1] if a else ""
            nb = b[1] if b else ""
            if na != nb:
                diffs.append({
                    "date": d,
                    "weekday": pnm._WD_JP[d.weekday()],
                    "col_index": j,
                    "col_label": _col_label(ga, gb, j),
                    "a": a[0] if a else "—",
                    "b": b[0] if b else "—",
                })
    return diffs


# ----------------------------------------------------------------------------
# (b) 医師別の回数差
# ----------------------------------------------------------------------------
def count_by_doctor(g):
    """正規化名 → {total, bg, ht, display}。分類は列位置ベース（BG=1-10, HT=11-24）。"""
    ht_end = min(pnm._HT_END_CAP, g.last_hosp_j)
    counts = collections.defaultdict(
        lambda: {"total": 0, "bg": 0, "ht": 0, "display": ""}
    )
    for _d, row in g.cells.items():
        for j, (disp, norm) in row.items():
            c = counts[norm]
            c["display"] = disp
            c["total"] += 1
            if pnm._BG_START <= j <= pnm._BG_END:
                c["bg"] += 1
            elif pnm._HT_START <= j <= ht_end:
                c["ht"] += 1
    return dict(counts)


def diff_counts(ca, cb):
    """医師別の A−B 差分。戻り値は |Δ合計| 降順→氏名順の list。"""
    rows = []
    for norm in set(ca) | set(cb):
        a = ca.get(norm, {"total": 0, "bg": 0, "ht": 0, "display": ""})
        b = cb.get(norm, {"total": 0, "bg": 0, "ht": 0, "display": ""})
        rows.append({
            "display": a["display"] or b["display"] or norm,
            "norm": norm,
            "a_total": a["total"], "b_total": b["total"],
            "d_total": a["total"] - b["total"],
            "a_bg": a["bg"], "b_bg": b["bg"], "d_bg": a["bg"] - b["bg"],
            "a_ht": a["ht"], "b_ht": b["ht"], "d_ht": a["ht"] - b["ht"],
        })
    rows.sort(key=lambda r: (-abs(r["d_total"]), r["norm"]))
    return rows


# ----------------------------------------------------------------------------
# レポート生成
# ----------------------------------------------------------------------------
def _fmt_delta(n):
    return f"+{n}" if n > 0 else str(n)


def build_report(ga, gb, label_a, label_b, cell_diffs, count_rows):
    """Markdown レポート文字列を返す。"""
    dates_a, dates_b = set(ga.cells), set(gb.cells)
    common = dates_a & dates_b
    only_a = sorted(dates_a - dates_b)
    only_b = sorted(dates_b - dates_a)

    # 突合対象スロット総数（両側の割当セルの和集合）
    total_slots = 0
    for d in dates_a | dates_b:
        total_slots += len(set(ga.cells.get(d, {})) | set(gb.cells.get(d, {})))
    n_diff = len(cell_diffs)
    n_match = total_slots - n_diff
    rate = (n_match / total_slots * 100) if total_slots else 100.0

    L = []
    L.append("# 当直表 突合レポート")
    L.append("")
    L.append(f"- **A**: {label_a}（シート `{ga.sheet_name}`）")
    L.append(f"- **B**: {label_b}（シート `{gb.sheet_name}`）")
    L.append("")
    L.append("## サマリー")
    L.append("")
    L.append(f"- 突合スロット数: **{total_slots}**")
    L.append(f"- 一致: **{n_match}** / 相違: **{n_diff}**")
    L.append(f"- 一致率: **{rate:.1f}%**")
    if only_a or only_b:
        L.append("- ⚠️ **日付カバレッジ不一致**: "
                 f"A のみ {len(only_a)}日 / B のみ {len(only_b)}日 "
                 "（同一雛形から作っていない可能性）")
        if only_a:
            L.append("  - A のみ: " + ", ".join(d.strftime("%m/%d") for d in only_a))
        if only_b:
            L.append("  - B のみ: " + ", ".join(d.strftime("%m/%d") for d in only_b))
    else:
        L.append(f"- 日付カバレッジ: 一致（{len(common)}日）")
    L.append("")

    # (a) セル単位の相違
    L.append("## (a) セル単位の相違")
    L.append("")
    if not cell_diffs:
        L.append("相違なし（全枠で割当が一致）。")
    else:
        L.append(f"同じ日・同じ枠で割当が異なるセル: **{n_diff}件**")
        L.append("")
        L.append(f"| 日付 | 曜 | 枠 | A: {label_a} | B: {label_b} |")
        L.append("|---|---|---|---|---|")
        for d in cell_diffs:
            L.append(
                f"| {d['date'].strftime('%m/%d')} | {d['weekday']} | "
                f"{d['col_label']} | {d['a']} | {d['b']} |"
            )
    L.append("")

    # (b) 医師別の回数差
    L.append("## (b) 医師別の回数差（A − B）")
    L.append("")
    nonzero = [r for r in count_rows if r["d_total"] or r["d_bg"] or r["d_ht"]]
    if not nonzero:
        L.append("回数差なし（全医師で 大学/外/合計 が一致）。")
    else:
        L.append(f"差のある医師: **{len(nonzero)}名**（Δ = A − B、|Δ合計| 降順）")
        L.append("")
        L.append("| 氏名 | Δ合計 | Δ大学 | Δ外 | "
                 f"合計({label_a}/{label_b}) | 大学 | 外 |")
        L.append("|---|---|---|---|---|---|---|")
        for r in nonzero:
            L.append(
                f"| {r['display']} | {_fmt_delta(r['d_total'])} | "
                f"{_fmt_delta(r['d_bg'])} | {_fmt_delta(r['d_ht'])} | "
                f"{r['a_total']}/{r['b_total']} | "
                f"{r['a_bg']}/{r['b_bg']} | {r['a_ht']}/{r['b_ht']} |"
            )
    L.append("")
    return "\n".join(L)


def compare(path_a, path_b, pattern_num=1, label_a=None, label_b=None):
    """2ファイルを突合し、レポート文字列と主要数値を返す。"""
    ga = load_grid(path_a, pattern_num)
    gb = load_grid(path_b, pattern_num)
    label_a = label_a or os.path.splitext(os.path.basename(path_a))[0]
    label_b = label_b or os.path.splitext(os.path.basename(path_b))[0]

    cell_diffs = diff_cells(ga, gb)
    count_rows = diff_counts(count_by_doctor(ga), count_by_doctor(gb))
    report = build_report(ga, gb, label_a, label_b, cell_diffs, count_rows)
    return {
        "report": report,
        "cell_diffs": cell_diffs,
        "count_rows": count_rows,
        "grid_a": ga,
        "grid_b": gb,
        "label_a": label_a,
        "label_b": label_b,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="当直表2枚の突合（セル相違＋医師別回数差）"
    )
    ap.add_argument("a_xlsx", help="A: 当直表 xlsx（pattern_NN 入り。慣例=当直くん出力）")
    ap.add_argument("b_xlsx", help="B: 当直表 xlsx（pattern_NN 入り。慣例=完成版）")
    ap.add_argument("--pattern", type=int, default=1,
                    help="使用する pattern 番号（既定 1）")
    ap.add_argument("--label-a", default=None, help="A の表示ラベル（既定=ファイル名）")
    ap.add_argument("--label-b", default=None, help="B の表示ラベル（既定=ファイル名）")
    ap.add_argument("-o", "--output", default=None,
                    help="Markdown レポートの出力先（省略時は stdout サマリーのみ）")
    args = ap.parse_args(argv)

    res = compare(args.a_xlsx, args.b_xlsx, args.pattern, args.label_a, args.label_b)

    print("=" * 60)
    print("  当直くん 突合ヘルパー: compare_schedules")
    print("=" * 60)
    n_diff = len(res["cell_diffs"])
    n_cnt = len([r for r in res["count_rows"]
                 if r["d_total"] or r["d_bg"] or r["d_ht"]])
    print(f"A: {res['label_a']}")
    print(f"B: {res['label_b']}")
    print(f"セル相違: {n_diff}件 / 回数差のある医師: {n_cnt}名")
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(res["report"])
        print(f"レポート出力: {args.output}")
    else:
        print("-" * 60)
        print(res["report"])
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
