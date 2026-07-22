# 当直くん モジュール分割設計書（ARCHITECTURE_PLAN）

> ステータス: **設計のみ**。実装は着手しない。
> 前提: 大型リファクタ（モジュール分割）は **Greedyエンジン退役後** に行うのが正。
> CP-SAT（`solver_cpsat.py`）が既定になった今、`main.py` に残る Greedy+fix 群は「併走検証中の旧実装」であり、
> 退役が確定するまで分割の主対象（`main.py`）が動き続けるため、先に分割すると二重メンテになる。
> 本書は「退役の判断基準」と「退役後に行う分割の到達点」を先に固定しておくためのもの。

---

## 1. 現状の構成と問題

### 1.1 ファイル

| ファイル | 行数(目安) | 役割 | 状態 |
|----------|-----------|------|------|
| `main.py` | ~7,130 | 入力パース＋Greedy生成＋fixパイプライン＋evaluate＋Excel出力＋`run()`統括＋CP-SAT合流 | 肥大・要分割 |
| `solver_cpsat.py` | ~1,200 | `InputData` パース／`CpSatScheduler`／`solve()`／CLI／`verify_output`／出力生成 | 既にモジュール化済（良い手本） |
| `config.py` | ~90 | 月次パラメータ | OK |
| `prepare_next_month.py` | ~600 | 翌月雛形の自動生成（独立実装） | OK |
| `backend/scheduler.py` | ~500 | Web化用スケルトン（`SchedulerConfig` dataclass 等） | **凍結**。分割の参考実装 |
| `tests/` | 87件 | 不変条件・プリフライト・列アンカー・月次準備 | 全緑を維持 |

### 1.2 `main.py` の中核問題: ステートフルな `run()`

`run(input_path, output_dir, num_patterns)` は関数化されている（import副作用ゼロ）が、
**内部で全束縛名を `global` 宣言する**設計になっている。`run()` 冒頭に多数の `global ...` があり、
パース結果（`shift_df` `availability_df` `doctor_names` `TARGET_CAP` `EXTRA_ALLOWED` `SCHEDULE_CODE_HOLDERS` …）と
ネスト定義された関数群（`evaluate_schedule_with_raw`, `validate_absolute_constraints`, 全 `fix_*`）が
モジュールグローバルを暗黙に共有している。

**帰結:**
- ユニットテストが `run()` 全実行に依存しがち（部分テストが難しい）。
- 同一プロセスで2回 `run()` すると状態が持ち越される危険。
- CP-SAT合流部（`_cpsat_to_pattern_df` 等）も `bg_cat` 等のグローバルへ代入してから評価する必要がある。

→ 分割の一丁目一番地は **「グローバル共有」を `ScheduleContext` dataclass に置き換える**こと。

---

## 2. 分割境界案（退役後の到達点）

`solver_cpsat.py` の `InputData` 分離が既に成功例。同じ発想で `main.py` を機能境界で割る。

```
tochoku/
├── config.py                # 既存（月次パラメータ）
├── context.py    (新)       # ScheduleContext dataclass（パース結果＋派生値の器）
├── io_excel.py   (新)       # 入力読込（sheet1/2/3）＋プリフライト＋Excel出力・summary
├── columns.py    (新)       # 列役割の名前解決（_resolve_column_anchors 等・v6.10.0で名前解決化済）
├── constraints.py(新)       # ABS/SEMI判定（validate_absolute_constraints・is_valid_full_assignment・eligibility）
├── scoring.py    (新)       # evaluate_schedule_with_raw（SOFT重みの一元管理）
├── solver_cpsat.py          # 既存（厳密解・既定）
├── solver_greedy.py(新→退役)# Greedy生成＋fix群＋局所探索（退役時に丸ごと削除）
├── pipeline.py   (新)       # run(): パース→ソルバ選択→評価→出力の統括
└── cli.py                   # エントリポイント（Finderダイアログ／引数）
```

### 2.1 各モジュールの責務

| モジュール | 公開関数(案) | 依存 |
|-----------|-------------|------|
| `context.py` | `ScheduleContext`（後述 §3） | config, pandas |
| `io_excel.py` | `load_input(path)->ScheduleContext` / `preflight_validate(ctx)` / `write_output(ctx, solutions, out_dir)` | context, columns |
| `columns.py` | `resolve_column_anchors(headers, kate_col)->(anchors, warnings)` | — |
| `constraints.py` | `validate_absolute(ctx, pattern_df)` / `is_valid_full_assignment(...)` / `is_eligible_for_ch_slot(...)` | context, columns |
| `scoring.py` | `evaluate(ctx, pattern_df)->(score, raw, metrics)` | context, constraints |
| `solver_cpsat.py` | `solve(ctx, n_solutions, min_diff)->solutions` | context, constraints |
| `pipeline.py` | `run(path, out_dir, num_patterns)` | 全部 |

> **原則**: `constraints.py` と `scoring.py` を CP-SAT と Greedy の**両方から呼べる純関数**にする。
> 現在 CP-SAT は「解を pattern_df に転写してから main.py の evaluate/validate を呼ぶ」二段構えだが、
> 分割後は両エンジンが同じ `constraints`/`scoring` を直接参照でき、二重定義が消える。

---

## 3. ScheduleContext dataclass 設計

`main.py::run()` が現在グローバルで持ち回っている状態を1つの immutable 寄り dataclass に集約する。
`backend/scheduler.py` の `SchedulerConfig` / `DutyScheduler` フィールド群を土台にする。

```python
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd

@dataclass
class ColumnAnchors:
    """列役割 → 0始まりインデックス（v6.10.0の名前解決結果）。"""
    B: int; C: int; D: int; E: int; F: int; G: int; H: int
    I: int; J: int; K: int; L: int; M: int; Q: int; U: int; Y: int
    # 由来: main.py::_resolve_column_anchors / _STANDARD_COL_ANCHORS

@dataclass
class ScheduleContext:
    # ---- 生入力（パース済み）----
    shift_df: pd.DataFrame                 # sheet1（枠グリッド）
    availability_df: pd.DataFrame          # Sheet2（可否・DatetimeIndex）
    doctor_info_df: pd.DataFrame           # Sheet3（医師情報＝旧Sheet4役割）
    date_col_shift: str
    hospital_cols: list[str]
    anchors: ColumnAnchors

    # ---- 医師集合 ----
    doctor_names: list[str]
    active_doctors: list[str]
    inactive_doctors: list[str]

    # ---- 派生パラメータ（現状グローバル）----
    target_cap: dict[str, int]             # TARGET_CAP
    extra_allowed: set[str]                # EXTRA_ALLOWED（属性1優先→名簿末尾）
    base_target: int
    extra_slots: int

    # ---- 属性・カテ ----
    doctor_attribute: dict[str, str]       # "1"/"2"
    schedule_code_holders: set[str]        # カテ保有医師（SCHEDULE_CODE_HOLDERS）
    no_kate_doctors: set[str]              # カテなし医師
    code_2_doctors: set[str]
    code_1_2_doctors: set[str]
    ratio_exempt_doctors: set[str]         # コード3（比率計算除外）
    wed_forbidden_doctors: set[str]

    # ---- 前月累積（照合後）----
    name_match: dict[str, Optional[str]]   # doc -> sheet4側氏名 or None
    prev_total: dict[str, float] = field(default_factory=dict)
    prev_bg: dict[str, float] = field(default_factory=dict)
    prev_ht: dict[str, float] = field(default_factory=dict)
    prev_weekday: dict[str, float] = field(default_factory=dict)
    prev_weekend: dict[str, float] = field(default_factory=dict)

    # ---- 環境 ----
    holidays: set[pd.Timestamp] = field(default_factory=set)
    slot_meta: dict = field(default_factory=dict)   # (ridx,hosp)->(date, is_fixed)

    # ---- スコア重み（config由来・一元管理）----
    weights: dict[str, float] = field(default_factory=dict)
```

**設計上の狙い:**
- `run()` の `global` 群 → `ctx` の1引数渡しに置換。関数の入出力が明示化され部分テスト可能に。
- CP-SAT の `InputData.target_cap` / `extra_allowed` と `ScheduleContext` を**同一パーサから生成**すれば、
  現在 v6.9.0 で assert している「二重パース整合性チェック」自体が不要になる（パースが1本化されるため）。
- `weights` を dict 一元管理し、`evaluate` 内のハードコード重み（300/150/120/100/80/50）を集約
  （`docs/CONSTRAINT_RULES.md §2.4` 注記の解消）。

---

## 4. 列役割の全面名前解決（状況）

- **完了（v6.10.0 / 6f15ecd）**: `_resolve_column_anchors` が sheet1 ヘッダ名から B/C/D/E/F/G/H/I/J/K/L/M/Q/U/Y の
  15アンカーを解決。標準位置とのズレは警告、範囲重複・逆転は致命 ValueError。位置固定は解決不能時のフォールバックのみ。
- **残**: 解決結果を `ColumnAnchors` dataclass 化し、`B_COL_INDEX` 等のグローバル定数参照を全廃する（分割時に併せて実施）。
  現在は名前解決した値をグローバル定数へ再代入している経路が残るため、`ctx.anchors` 直参照へ寄せる。

---

## 5. Greedy退役の判断基準

Greedy+fixパイプラインを削除してよいと判断する条件（**全て満たす**）:

1. **併走差分ゼロ**: 実運用の月次入力で、CP-SAT が **N≥2ヶ月連続**（推奨3ヶ月）で
   「全ABS違反0・SEMI違反0・fair_span 許容内」の3解を返し、Greedyフォールバックが一度も発火しない。
2. **フォールバック未使用の確認**: 期間中 `🧩 ソルバーエンジン: CP-SAT` のみが出力され、
   ImportError/INFEASIBLE/整合性assert失敗のいずれも記録されない。
3. **性能の許容**: 想定最大規模（医師~35・枠~110）で CP-SAT の3解生成が実用時間（数秒〜数十秒）に収まる。
4. **テスト移管完了**: Greedy専用でない不変条件テスト（ABS/SOFT定義・プリフライト・列アンカー・月次準備）が
   CP-SAT経路で全緑。Greedy固有テストは退役と同時に削除 or CP-SATへ読み替え。
5. **ユーザー承認**: 運用者（Yusuke）が「Greedyフォールバックを外す」ことに明示同意。

> INFEASIBLE が現実の入力で起きうる場合（過密月・可否0が多い月）は、Greedyフォールバックを
> **残す**判断も妥当。その場合は「退役」ではなく「Greedyを縮退（fix群の一部のみ保持）」を選ぶ。

---

## 6. 削除対象一覧（退役が確定した場合）

`main.py` から削除できるコード（すべて Greedy 経路専用）。概算 **約2,500行**。

### 6.1 fix関数群（14本・約2,300行 / main.py 3575〜5888行）

| # | 関数 | 開始行(目安) |
|---|------|-------------|
| 1 | `fix_hard_constraint_violations` | 3575 |
| 2 | `fix_target_cap_violations` | 3749 |
| 3 | `fix_code_2_extra_violations` | 3928 |
| 4 | `fix_university_minimum_requirement` | 4102 |
| 5 | `fix_ch_kate_violations` | 4217 |
| 6 | `fix_bg_ht_imbalance_violations` | 4409 |
| 7 | `fix_gap_violations` | 4538 |
| 8 | `fix_external_hospital_dup_violations` | 4699 |
| 9 | `fix_university_over_2_violations` | 4863 |
| 10 | `fix_weekly_bg_violations` | 5092 |
| 11 | `fix_university_weekday_balance_violations` | 5266 |
| 12 | `fix_weekday_weekend_balance` | 5465 |
| 13 | `fix_fairness_imbalance` | 5595 |
| 14 | `fix_unassigned_slots` | 5768 |

＋ `safe_fix` ラッパー（6096〜）・収束ループ本体（6356〜、`OPTIMIZATION_ENABLED` 分岐）。

### 6.2 局所探索（約180行 / main.py 2930〜3110）

- `local_search_swap`（既定 `LOCAL_SEARCH_ENABLED=False` のため現状も未使用）。
- 補助: `count_doc_in_column` 等の swap 専用ヘルパー。

### 6.3 Greedy生成本体

- `build_schedule_pattern`（2119〜）・`collect_candidates` の緩和フラグ経路（relax_semi/hard/abs）・
  `run()` 内 `if _use_greedy:` ブロック（6274〜）と `NUM_PATTERNS`/`TOP_KEEP`/`REFINE_TOP` 系定数。

### 6.4 削除しないもの（両エンジン共用・分割後は `constraints.py`/`scoring.py` へ移設）

- `validate_absolute_constraints`（ABS最終検査）
- `evaluate_schedule_with_raw`（SOFTスコア）
- `recompute_stats` / eligibility 判定（`is_eligible_for_ch_slot` 等）
- 入力パース・プリフライト・Excel出力

> **注意**: `fix_unassigned_slots`(#14) は ABS-009 の最終セーフティネットでもある。CP-SAT が
> 未割当0を保証する前提が崩れる入力（フォールバック時）を残すなら、この1本だけ `constraints`/補修側へ
> 退避する選択肢がある。

---

## 7. 実施時期の前提条件（まとめ）

| 段階 | 前提 | 内容 |
|------|------|------|
| いま | — | 本設計書を凍結。CP-SAT併走を継続し、フォールバック発火をログで監視 |
| Greedy退役 | §5 の1〜5を全充足 | §6 の約2,500行を削除。Greedy固有テストを整理 |
| モジュール分割 | 退役完了後 | §2 の境界で `main.py` を分割。§3 の `ScheduleContext` を導入し `global` を全廃。CP-SATの整合性assertを撤去（パース1本化） |
| 仕上げ | 分割後 | `evaluate` のハードコード重みを `ctx.weights` へ集約（CONSTRAINT_RULES.md §2.4 の3系統ID分裂を解消） |

**リスク**: 分割中はテスト（現87件）を安全網にする。1モジュールずつ切り出し、各ステップで
`python3 -m pytest tests/ -q` 緑・実雛形フル実走の3解緑バナーを確認してから次へ進む。
