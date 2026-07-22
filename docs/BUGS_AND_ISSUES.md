# 当直くん - バグと問題点

> 最終更新: 2026-07-22（v6.9.0時点）。解決済み項目は末尾の「✅ 解決済み」節へ分離した。
> 本体ロジックの正本は `main.py`、厳密解エンジンは `solver_cpsat.py`。

## 🔴 未解決（要対応）

### U-1. 重複列名処理の副作用（`make_unique`）
```python
headers = make_unique(headers)
```
**問題**: sheet4（医師情報シート）で「大学病院」が2列あると「大学病院_2」になり、集計列名と不一致になりうる。
**影響**: 累計列の読み違い。ただし現行の v7 雛形では発生していない。
**対応案**: 元データの重複列名をプリフライトで検出して警告／停止する。

### U-2. Greedyエンジンの性能（`recompute_stats` の全走査）
```python
# recompute_stats() が毎回全DataFrameを走査（O(n)）
```
**問題**: Greedy経路では fix ループ・局所探索の各反復で全再計算が走る。
**影響**: `SOLVER="greedy"` かつ `NUM_PATTERNS` を大きくすると遅い。
**現状**: 既定は CP-SAT（3解を数秒）。Greedyはフォールバック専用のため実害は限定的。
**対応案**: 差分更新（delta update）。ただしGreedy退役方針（`docs/ARCHITECTURE_PLAN.md`）と重複投資になるため優先度は低い。

### U-3. 平日/祝日判定のドメイン依存
```python
if weekday and idx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX):
    holi = True
```
**問題**: 列インデックスに基づく祝日扱いがドメイン固有で、列構成変更に弱い。
**現状**: 列役割は名前解決（`_resolve_column_anchors`, v6.10.0）に移行し標準位置とのズレは警告される。判定ロジック自体のコメント整備は未了。
**対応案**: 列役割ベースの判定へ統一し、根拠コメントを追加。

### U-4. マジックナンバー／ハードコードの残り
```python
W_UNASSIGNED = 500        # config化済み
# 一方 evaluate 内の 300/150/120/100/80/50 はハードコード
```
**問題**: `evaluate_schedule_with_raw` 内の SOFT ペナルティ重み（300/150/120/100/80/50）はハードコードのまま。config.py の `W_CODE_12_UNIV`/`W_BG_HT_DIFF` は greedy スコアに反映されない（`docs/CONSTRAINT_RULES.md §2.4` 注記）。
**対応案**: 重みを一箇所（config or dataclass）へ集約。Greedy退役時に一括整理予定。

### U-5. Web化時のセキュリティ（`backend/` 凍結中）
- ファイルアップロードのサイズ制限なし
- 悪意あるExcel（XXE等）の処理
- ユーザーデータの保存場所・期間の未定義

**現状**: `backend/` `frontend/` は将来のWeb化まで凍結（`backend/FROZEN.md`）。着手時に対応する。

---

## ✅ 解決済み

| ID | 項目 | 解決バージョン | 備考 |
|----|------|---------------|------|
| R-1 | `date_doc_count` の KeyError リスク | v2.1 | `defaultdict` 化で解消 |
| R-2 | 同日重複チェックの不完全性 | v2.1 / v6.0.0 | ABS-006として `validate_absolute_constraints` で最終検査 |
| R-3 | 列インデックスのハードコーディング | v6.10.0 | `_resolve_column_anchors` でヘッダ名から5アンカー解決・標準位置とのズレを警告（`tests/test_column_anchors.py`） |
| R-4 | 医師名の完全一致依存（表記ゆれ） | v2.1 / v6.9.0 | `normalize_name`（空白除去）＋ sheet4照合に `config.NAME_ALIASES`（完全一致＋手動エイリアス） |
| R-5 | タイムゾーン問題 | v2.1 | 全日付処理で `normalize().tz_localize(None)` |
| R-6 | sheet4 ヘッダ検出の脆弱性 | v2.1 | 検索範囲 30→50行、エラーメッセージ改善 |
| R-7 | 非active医師への cap 設定 | （実装済み） | `preassigned_count > TARGET_CAP` で補正 |
| R-8 | エラーハンドリング不足（空ファイル/列欠落/日付パース） | v6.9.0 | `preflight_validate()` が致命=停止・警告=続行で一括検証 |
| R-9 | cap超過フォールバックの未文書化 | v6.0.0〜 | 段階的制約緩和（relax_semi→hard→abs）として仕様化（CONSTRAINT_RULES.md §2.3） |
| R-10 | スコアリングの透明性不足 | v6.1.0〜 | `{pattern}_summary` 診断シートにスコア内訳・違反一覧・医師別偏りを出力 |
| R-11 | メモリ使用量（NUM_PATTERNS=10000で全保持） | v3.2 / v6.9.0 | 既定 NUM_PATTERNS=1000＋TOP_KEEP保持。CP-SAT既定化で greedy 大量生成自体が不要に |
| R-12 | ハードコードされた病院名/医師名 | v6.x | `WED_FORBIDDEN_DOCTORS`・`NAME_ALIASES`・`HOLIDAYS` を config.py へ外出し |
| R-13 | テストケース不足 | v6.9.0 | `tests/`（不変条件・プリフライト・列アンカー・月次準備で計87件・全緑） |
| R-14 | ABS-013 の検証漏れ | v6.5.9 | `validate_absolute_constraints` にC-H列カテ当番チェックを追加 |
| R-15 | sheet2 空欄の「月初マーク伝播」バグ | v6.5.9 | 空欄は常に「1（可）」扱いへ修正・解釈不能マークは読込時警告 |
| R-16 | 公平化が固定割当を上書きするバグ | v6.5.9 | `fix_fairness_imbalance` に `is_preassigned_slot()` チェック追加 |
| R-17 | sheet1にあってsheet2に無い日付が黙って全員「可」化 | v6.9.0 | プリフライトで致命として検知・停止 |
| R-18 | 同姓医師（佐藤彰/悠/勇）の氏名誤マッチ | v6.9.0 | 双方向startswithを撤去し完全一致のみに |
