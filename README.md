# 当直くん - 医師当直スケジュール自動生成ツール

> **正本はリポジトリ直下の `main.py`**（厳密解ソルバは `solver_cpsat.py`）。`backend/` `frontend/` は将来のWeb化まで凍結、`docs/archive/` は旧Colab版のアーカイブです。

## 使い方

### ターミナルにコピー＆ペーストするだけ

**初回:**

```bash
git clone https://github.com/tomitayus/Tochoku-kun.git && cd Tochoku-kun && bash run.sh
```

**2回目以降:**

```bash
cd Tochoku-kun && bash run.sh
```

自動的にExcelファイルの選択ダイアログが表示されるので、入力ファイルを選んでください。
結果は `~/Downloads/` に保存されます（`<入力名>_v<version>.xlsx`）。

## 実行フロー（v6.9.0〜）

```
入力Excel
  │
  ├─ ① プリフライト検証  … 日付整合・可否シート欠落・枠数上限・氏名突合を一括チェック
  │        └ 致命は停止（ValueError）／警告は表示して続行
  │
  ├─ ② ソルバー実行
  │      ├─ 既定: CP-SAT厳密解（solver_cpsat.py）… 全制約違反ゼロの3解を数秒で生成
  │      └─ フォールバック: Greedy + fixパイプライン（main.py）
  │           ※ ortools未導入 / INFEASIBLE / 整合性不一致 のとき自動切替（stdoutに明示）
  │
  └─ ③ 評価・出力  … 両エンジン共通の evaluate/summary/バナーで Excel 出力
           （pattern_XX / pattern_XX_summary シート）
```

**ソルバーの選択**は `config.py` の `SOLVER`（`"cpsat"` 既定 / `"greedy"`）で切替。詳細は `docs/CONSTRAINT_RULES.md §6`。
> CP-SATを使うには `ortools` が必要（`pip install ortools`）。未導入なら自動的にGreedyへフォールバックする。

### 入力Excelの構造（3シート）

| シート名 | 内容 |
|---------|------|
| sheet1 | シフト枠（日付 + 病院列、枠は`1`／医師名で固定割当）＋ Z列「カテ当番」 |
| Sheet2 | 医師の可否（空欄=可, 0=不可, 1=可, 1.2=大学優先, 2=大学系のみ, 3=外病院のみ） |
| Sheet3 | 医師情報（氏名・属性・カテ当番チーム・出張・前月累積） |

> **記入方法の詳細は `docs/TEMPLATE_GUIDE.md`**（可否コード・属性・カテ当番・累計列・エラー対応表・NAME_ALIASES）。

### 作成者実運用ルールの反映（v6.10.0〜v6.12.0 要点）

- **学年クォータ制（v6.10.0）**: Sheet3に「大学目標」「外目標」の両列を入れると、医師別に
  大学系=大学目標・総数=大学目標+外目標の**等式**で配分（BASE_TARGET自動計算とEXTRA(+1)を全置換）。
  ローテ者は月毎の目標入力で回数を変えられる。Σが枠数と合わないと差分表示付きで停止。
- **大学当直は暦週1回（v6.10.0）**: ABS-012をローリング7日から**暦週（日曜始まり〜土曜）で1回**に変更。
- **日直は月1回（v6.10.0）**: 土曜昼/日曜昼/祝日昼+支援日直は医師別に月1回まで（ABS-016）。
- **外勤の前日+当日は自動NG（v6.11.0）**: Sheet3「出張日」（複数曜日は`木・金`）から自動生成。
  例外は `GAIKIN_EXCEPTIONS`（外勤先×相対日×許容当直先）で列限定緩和。
- **二層公平（v6.12.0）**: 単月は学年傾斜・**年度の外病院累計は全員均等**（`W_FAIR_CUM_HT`、
  恒常例外は `CUM_FAIRNESS_EXEMPT`）。summaryの「外病院累計spread」行で確認。
- **ソフト回避・カテ番合わせ（v6.12.0）**: `SOFT_AVOID_DOCTORS`（割当1回ごと軽ペナルティ）、
  カテ当番日の平日大学枠にカテ当番医師が入ると加点（`W_KATE_WEEKDAY_BONUS`）。

### 翌月雛形の準備

前月の出力から累計を自動加算して翌月雛形ドラフトを作る（手転記の氏名照合事故を防止）:

```bash
python3 prepare_next_month.py <当月の出力.xlsx> <当月の雛形.xlsx> --pattern 1 -o <翌月ドラフト.xlsx>
```

### 併走検証（当直くん vs 完成版）

当直くん出力と手作り完成版を突き合わせ、セル相違・医師別回数差・一致率を出す:

```bash
python3 compare_schedules.py <当直くん.xlsx> <完成版.xlsx> --label-a 当直くん --label-b 完成版 -o <突合.md>
```

> 完成版は `pattern_01` シート（A1=`日付`・列並びは雛形 sheet1 と同じ）にするだけで突合対象になる。
> 月次の運用フロー・フォルダ構成の詳細は **`docs/HEISOU_GUIDE.md`**。

### 希望の収集・反映（Google Form）

医局員の希望（避けたい/やりたい/当直不可）を Google Form で集め、入力に反映:

```bash
python3 build_wish_from_form.py <form_responses.csv> <入力.xlsx> -o <出力.xlsx>
```

> 「希望」シート（`×1〜×3`避け / `○1〜○3`やりたい）＋ `Sheet2` の `0`（不可）を自動生成。
> 反映ロジックは `docs/WISH_CONSTRAINT_SPEC.md`、Form設計は `docs/WISH_FORM_DESIGN.md`。

### config.py の編集

必要に応じて `config.py` を更新してください。

```python
SOLVER = "cpsat"                     # ソルバー（"cpsat"既定 / "greedy"）
HOLIDAYS = []                        # 祝日の手動追加（通常は不要: v6.8.0から自動取得）
WED_FORBIDDEN_DOCTORS = ["金城"]     # 水曜外病院禁止医師
NAME_ALIASES = {}                    # 氏名の表記ゆれマップ（例 {"冨田": "富田"}）
NUM_PATTERNS = 1000                  # 生成パターン数（cpsat時は無視・常に3解）
GAIKIN_EXCEPTIONS = {}               # 外勤前日/当日NGの例外ペア（v6.11.0）
W_FAIR_CUM_HT = 20                   # 外病院の年度累計均等化の重み（v6.12.0・0で無効）
CUM_FAIRNESS_EXEMPT = []             # 累計均等化の対象外医師（恒常的外勤過多など）
SOFT_AVOID_DOCTORS = []              # なるべく当直を減らしたい医師（軽ペナルティ）
W_KATE_WEEKDAY_BONUS = 10            # カテ番×平日大学の「できれば合わせる」加点
```

**祝日は自動取得されます（v6.8.0〜）**: jpholiday がインストールされていればライブラリから、
なければ内蔵祝日表（2026〜2027年）から対象月の祝日を自動設定します。
実行時に `📅 祝日: ...` と表示されるので、対象月の祝日が正しいか確認してください。
2028年以降を扱う場合は `pip install jpholiday` を推奨します。
