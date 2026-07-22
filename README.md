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

### 翌月雛形の準備

前月の出力から累計を自動加算して翌月雛形ドラフトを作る（手転記の氏名照合事故を防止）:

```bash
python3 prepare_next_month.py <当月の出力.xlsx> <当月の雛形.xlsx> --pattern 1 -o <翌月ドラフト.xlsx>
```

### config.py の編集

必要に応じて `config.py` を更新してください。

```python
SOLVER = "cpsat"                     # ソルバー（"cpsat"既定 / "greedy"）
HOLIDAYS = []                        # 祝日の手動追加（通常は不要: v6.8.0から自動取得）
WED_FORBIDDEN_DOCTORS = ["金城"]     # 水曜外病院禁止医師
NAME_ALIASES = {}                    # 氏名の表記ゆれマップ（例 {"冨田": "富田"}）
NUM_PATTERNS = 1000                  # 生成パターン数（cpsat時は無視・常に3解）
```

**祝日は自動取得されます（v6.8.0〜）**: jpholiday がインストールされていればライブラリから、
なければ内蔵祝日表（2026〜2027年）から対象月の祝日を自動設定します。
実行時に `📅 祝日: ...` と表示されるので、対象月の祝日が正しいか確認してください。
2028年以降を扱う場合は `pip install jpholiday` を推奨します。
