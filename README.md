# 当直くん - 医師当直スケジュール自動生成ツール

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
結果は `~/Downloads/` に保存されます。

### 入力Excelの構造（sheet1〜3）

| シート名 | 内容 |
|---------|------|
| sheet1 | シフト枠（日付 + 病院列、枠は`1`で表記） |
| sheet2 | 医師の可否（0=不可, 1=可, 1.2=大学優先, 2=大学系のみ, 3=外病院のみ） |
| sheet3 | 医師情報（氏名、累積データ、属性 等） |

### config.py の編集

月ごとに `config.py` の祝日・禁止医師を更新してください。

```python
HOLIDAYS = ["2026-04-29"]           # 祝日
WED_FORBIDDEN_DOCTORS = ["金城"]     # 水曜外病院禁止医師
NUM_PATTERNS = 1000                  # 生成パターン数
```
