# 当直くん - 医師当直スケジュール自動生成ツール

Google Colab版の当直表自動生成コードをローカル実行用に変換したプロジェクトです。

## 概要

医師の当直スケジュールを、複数の制約条件を考慮して自動生成するツールです。

### 主な機能

- 複数の制約を考慮した自動割当（可否、カテ表、gap、cap等）
- Greedy + 局所探索による最適化
- 複数パターン生成（100〜10000パターン）
- TOP3パターンの出力（公平性/gap回避/バランスの3軸評価）
- 診断情報の自動生成（gap違反、重複、偏り等）
- 前月累積データの考慮

## 前提条件

- Python 3.10 以上
- pip（パッケージ管理ツール）

## セットアップ手順

### 1. リポジトリのクローン

```bash
git clone <リポジトリURL>
cd Tochoku-kun
```

### 2. 仮想環境の作成

```bash
python3 -m venv .venv

# macOS/Linux:
source .venv/bin/activate

# Windows:
.venv\Scripts\activate
```

### 3. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 4. データの配置

`data/` フォルダに当直表のExcelファイルを配置してください。

```
data/
└── Tochoku.xlsx   ← ここに配置
```

**注意:** `data/` フォルダ内のExcel/CSVファイルはGitに含まれません（個人情報保護のため）。
手動で配置する必要があります。

### Excelファイルの構造

入力Excelファイルには以下のシートが必要です：

| シート名 | 内容 |
|---------|------|
| sheet1 | シフト表（日付 + 病院列、枠は`1`で表記） |
| sheet2 | 医師の可否（0=不可, 1=可, 2=B〜Mのみ, 3=H〜Uのみ） |
| sheet3 | カテ表（省略可：Sheet1:Z + Sheet4:属性で代替） |
| sheet4 | 前月までの累積データ（氏名、全合計、大学合計、外病院合計 等） |

`data/sample_staff.csv` にsheet4の構造のサンプルがあります。

## 実行方法

### 基本的な実行

```bash
python main.py
```

`config.py` で設定した `INPUT_FILE`（デフォルト: `data/Tochoku.xlsx`）を読み込み、
結果を `data/output/` に出力します。

### 入力ファイルを指定して実行

```bash
python main.py data/2026年4月当直表.xlsx
```

### 出力

`data/output/` に以下の形式のExcelファイルが生成されます：

```
data/output/Tochoku_v6.5.6.xlsx
```

ファイル内のシート構成：
- `pattern_01`: スケジュールパターン1（公平性重視）
- `pattern_01_summary`: パターン1のサマリー・診断
- `pattern_02`: スケジュールパターン2（gap回避重視）
- `pattern_02_summary`: パターン2のサマリー・診断
- `pattern_03`: スケジュールパターン3（バランス重視）
- `pattern_03_summary`: パターン3のサマリー・診断

## config.py の編集方法

月次運用では `config.py` のみ編集すればOKです。

### 主な設定項目

```python
# 入力ファイルパス
INPUT_FILE = "data/Tochoku.xlsx"

# 祝日（対象月のものを追加）
HOLIDAYS = [
    "2026-04-29",  # 昭和の日
]

# 水曜H〜U禁止医師
WED_FORBIDDEN_DOCTORS = ["金城", "山田", "野寺"]

# 生成パターン数（推奨: 1000）
NUM_PATTERNS = 1000
```

詳細は `config.py` のコメントを参照してください。

## ディレクトリ構成

```
Tochoku-kun/
├── main.py                    # メイン実行スクリプト
├── config.py                  # 設定ファイル（月ごとに編集）
├── requirements.txt           # Python依存パッケージ
├── .gitignore                 # Git除外設定
├── README.md                  # このファイル
├── colab_duty_scheduler.py    # 元のColabコード（参照用）
├── VERSION_HISTORY.md         # バージョン履歴
├── data/
│   ├── .gitkeep
│   ├── sample_staff.csv       # サンプルデータ（sheet4の構造参考）
│   └── output/                # 出力先（自動作成）
│       └── .gitkeep
├── backend/                   # Web版バックエンド（開発中）
├── frontend/                  # Web版フロントエンド（開発中）
└── docs/                      # ドキュメント
```

## 制約体系

### 絶対禁忌（ABS）: 配置不可
- ABS-001: 可否コード0禁止
- ABS-002: コード2の列制約（B〜M列のみ）
- ABS-003: コード3の列制約（H〜U列のみ）
- ABS-004: カテ当番日の外病院禁止
- ABS-005: 水曜日外病院禁止医師
- ABS-006: 同日重複禁止
- ABS-007: gap3日未満禁止
- ABS-008: 同一外病院重複禁止
- ABS-013: C-H列カテ当番必須

### ソフト制約（SOFT）: ペナルティ
- 公平性（max-min最小化）
- コード1.2の大学系最低1回
- 大学/外病院バランス

詳細は `docs/CONSTRAINT_RULES.md` を参照してください。

## トラブルシューティング

### 「入力ファイルが見つかりません」エラー
→ `data/` フォルダにExcelファイルを配置し、`config.py` の `INPUT_FILE` パスを確認してください。

### 「必要なシートが見つかりません」エラー
→ Excelファイル内のシート名が `sheet1`, `sheet2`, `sheet3`, `sheet4` であることを確認してください（大文字小文字は問いません）。

### tqdmがインストールされていない警告
→ `pip install tqdm` を実行してください。進捗バーなしでも動作します。

## Web版について

`backend/` と `frontend/` ディレクトリにWeb版のコードがありますが、現在開発中です。
ローカルCLI版（`main.py`）の使用を推奨します。
