# 当直くん 制約仕様書（v6.10.0時点）

> 本ドキュメントは制約ルールの完全な定義を提供します。
> 実装コード（正本）:
> - **制約ID登録簿・可否コード判定・スコア重み**: `main.py`（現行版数は `VERSION` 定数＝6.9.0）
> - **絶対禁忌(ABS)の最終定義**: `main.py::validate_absolute_constraints()`（ABS-001〜015を実際に検査する関数）
> - **ソフト制約(SOFT)の適用重み**: `main.py::evaluate_schedule_with_raw()`（penalty加算そのもの）
> - **既定ソルバ（v6.9.0〜）**: `solver_cpsat.py`（CP-SAT厳密解）。ABS/SEMIをハード制約として直接モデル化し、
>   本書のSOFT重みで目的関数を最小化する。Greedyエンジンはフォールバック（後述 §6）。
> ※旧Colab版 `colab_duty_scheduler.py`(v6.7.0) は `docs/archive/` にアーカイブ済み。参照しないこと。

---

## §1 定義

### 1.1 用語定義

| 用語 | 定義 | コード上の表現 |
|------|------|----------------|
| **active医師** | 当月に割り当て可能な医師。sheet2で少なくとも1日は可否コード≠0、または事前割当がある医師 | `active_doctors` |
| **inactive医師** | 当月に割り当て不可の医師。sheet2で全日が可否コード0かつ事前割当なし。**解析処理から除外し、出力データのみに記載** | `inactive_doctors` |
| **BASE_TARGET** | 基本割当回数。`全枠数 ÷ active医師数`（端数切り捨て） | `BASE_TARGET` |
| **EXTRA_SLOTS** | 余り枠数。`全枠数 - (BASE_TARGET × active医師数)` | `EXTRA_SLOTS` |
| **TARGET_CAP** | 医師別の上限回数。通常は`BASE_TARGET`、余り対象医師は`BASE_TARGET + 1` | `TARGET_CAP[doc]` |
| **BG** | 大学系病院（B〜K列）の略称。Big Group の意 | `bg_counts` |
| **HT** | 外病院（L〜Y列）の略称。Hospital (外) の意 | `ht_counts` |
| **カテ当番** | その日にカテ表でアルファベット（A,B,C,CC,D,E等）が入っている状態。v6.5.0以降はSheet1:Z列+Sheet4:カテチーム属性で判定 | `get_sched_code()` |
| **カテ保有医師** | Sheet4のカテチーム属性がSheet1:Zのチームコードと一致する医師 | `SCHEDULE_CODE_HOLDERS` |
| **カテなし医師** | カテチーム属性を持たない医師（自由に配置可能） | `NO_KATE_DOCTORS` |
| **CC** | 大型連休特別シフト。公平性・重複計算から除外される | `is_cc_assignment()` |
| **gap** | 当直間隔。連続する2つの当直日の日数差 | `gap = date2 - date1` |
| **属性** | Sheet4の「属性」列。1=緩和可、2=緩和不可。SEMI-001の緩和判定に使用 | `doctor_attribute` |

### 1.2 列定義（Column Classification）

| 列範囲 | インデックス | 分類 | 説明 | カテ制約 |
|--------|-------------|------|------|---------|
| **B列** | 1 | 大学平日① | 平日の大学系当直（昼） | 緩和可 |
| **C-H列** | 2-7 | 大学休日 | 土日祝の大学系当直 | **カテ当番必須（ABS-013）** |
| **I列** | 8 | 大学平日② | 平日の大学系当直（昼） | 緩和可 |
| **J-K列** | 9-10 | 大学平日③ | 平日の大学系当直 | 緩和可 |
| **L-Y列** | 11-24 | 外病院 | 外部病院への派遣当直 | カテ当番日禁止 |

**時間帯分類（C-H列内）:**
- **昼（Day）**: C列、E列、F列
- **夜（Night）**: D列、G列

**グループ定義:**
```
大学系（B-K列）= B + C-H + I-K = インデックス 1-10
  ├─ 平日大学系 = B列 + I-K列（カテ制約緩和可）
  └─ 休日大学系 = C-H列（カテ当番必須）
外病院（L-Y列）= インデックス 11-24
```

**設計背景（J-K列が平日扱いの理由）:**
> J-K列は運用上、平日の補助枠として使用されている。休日（C-H列）はカテ当番との連携が重要だが、J-K列は平日業務の延長であるためカテ制約を緩和している。将来的にJ-K列の運用が変更される場合は、本定義の見直しが必要。

### 1.3 可否コード定義（Sheet2）

| コード | 意味 | 配置可能列 | 備考 |
|--------|------|-----------|------|
| **0** | 不可 | なし | その日は割り当て禁止（絶対禁忌） |
| **1** | 可 | B〜Y全て | 通常の配置対象 |
| **1.2** | 大学優先 | B〜Y全て | 大学系（B-K）を最低1回必須（準ハード） |
| **2** | 大学専用 | B〜Q列のみ | 外病院不可、EXTRA枠対象外 |
| **3** | 外病院専用 | L〜Y列のみ | 大学系不可、比率バランス除外 |

**特記事項:**
- **全日0の医師**: `inactive_doctors`として解析処理から完全除外。出力Excelには0回として記載
- **コード2医師**: v6.0.5以降EXTRA枠対象に含む。ただしTARGET_CAPは厳守

### 1.4 パラメータ定義

#### TARGET_CAP算出方法
```python
BASE_TARGET = total_slots // len(active_doctors)
EXTRA_SLOTS = total_slots - BASE_TARGET * len(active_doctors)

# 余り枠は属性1の医師を優先、不足時はSheet2末尾からフォールバック
# v6.0.5: CODE_2医師もEXTRA対象に含む
attr1_doctors = [d for d in active_sorted if doctor_attribute[d] == "1"]
EXTRA_ALLOWED = attr1_doctors[-EXTRA_SLOTS:]  # 属性1優先

TARGET_CAP[doc] = BASE_TARGET                    # 通常医師
TARGET_CAP[doc] = BASE_TARGET + 1                # EXTRA対象医師
TARGET_CAP[doc] = max(TARGET_CAP[doc], 事前割当数)  # 事前割当優先

# v6.1.0: gap>=3制約を満たせる物理的上限でCAPを制限
max_gap3 = compute_max_gap3_assignments(doc)
TARGET_CAP[doc] = min(TARGET_CAP[doc], max_gap3)
```

#### gap（間隔）の計算方法
```
間隔 = 次回当直日 - 前回当直日 の日数差

例:
  1/1 → 1/2 : gap = 1日 → NG（3日未満）
  1/1 → 1/3 : gap = 2日 → NG（3日未満）
  1/1 → 1/4 : gap = 3日 → OK（3日以上）

判定: gap < 3 → 違反
```

### 1.5 カテ表コード定義

v6.5.0以降、カテ当番の判定は以下の2段階で行われる:

1. **Sheet1:Z列「カテ当番」** - 日付ごとのチームコード（A, B, C, D, E等）
2. **Sheet4:カテチーム属性** - 医師ごとのチーム所属

| コード | 意味 | 配置への影響 |
|--------|------|-------------|
| **A,B,C,D,E等** | カテ当番（アルファベット） | その日L-Y禁止、C-H優先配置 |
| **CC** | 大型連休特別シフト | 公平性・重複計算から除外 |
| **1** | 平日大学緩和フラグ / 属性1 | B/I-K列でカテなしでも配置可 |
| **3** | 比率除外フラグ | BG/HT比率計算から除外 |
| **0 / 空白** | カテなし | C-H自由配置可、カテ制約なし |

---

## §2 制約一覧

> **制約ID体系**: `ABS`=絶対禁忌, `HARD`=ハード, `SEMI`=準ハード, `SOFT`=ソフト

### 2.1 絶対禁忌（違反=即却下・配置不可）

| ID | 制約名 | 条件 | 実装 |
|----|--------|------|------|
| ABS-001 | 可否コード0禁止 | `get_avail_code(date, doc) == 0` | `collect_candidates`で除外 |
| ABS-002 | コード2の列制約 | コード2医師がR〜Y列に配置 | `collect_candidates`で除外 |
| ABS-003 | コード3の列制約 | コード3医師がB〜K列に配置 | `collect_candidates`で除外 |
| ABS-004 | カテ当番日の外病院禁止 | カテ保有医師がカテ当番日にL〜Y列配置 | `collect_candidates`で除外 |
| ABS-005 | 水曜日L〜Y禁止医師 | 指定医師が水曜日にL〜Y列配置 | `WED_FORBIDDEN_DOCTORS`で除外 |
| ABS-006 | 同日重複禁止 | 同一医師が同一日に2枠以上 | ロジックで排除 |
| ABS-007 | gap>=3日必須 | 連続当直の間隔が3日未満 | `collect_candidates`で除外（v6.0.0: ABS格上げ） |
| ABS-008 | 同一病院重複禁止 | 初期生成時は全列、fix関数では外病院のみ | `collect_candidates` / `is_valid_full_assignment` |
| ABS-009 | 未割当禁止 | slot_metaの枠に医師が入っていない／不明医師名 | `fix_unassigned_slots`で最終補填・`validate_absolute_constraints`で検査 |
| ABS-010 | TARGET_CAP厳守 | 割当回数がTARGET_CAPを超過 | `collect_candidates`で除外（v6.0.0: ABS格上げ） |
| ABS-011 | 大学系上限 | B-K列の合計が上限を超過。既定2回・**学年クォータ制（v6.10.0）有効時は医師別「大学目標」** | `collect_candidates`で除外（`UNIV_CAP[doc]`） |
| ABS-012 | 大学系は暦週1回まで | 同一暦週（**日曜始まり〜土曜**）に大学系2回以上（v6.10.0: ローリング7日間隔から変更） | `collect_candidates`で除外 |
| ABS-013 | C-H列カテ当番必須 | 休日大学系にカテ当番日以外で配置（v6.5.3、旧SEMI-002格上げ） | `collect_candidates`で除外 |
| ABS-014 | 平日/休日偏り差≤1 | 1医師の `|平日回数 − 休日回数| ≥ 2` | `is_valid_full_assignment` / 生成時フィルタ（v6.8.0）・`validate_absolute_constraints`で検査 |
| ABS-015 | 属性2のB列カテ表必須 | 属性2医師がカテ表なしでB列に配置（v6.5.6） | `collect_candidates`で除外 |
| ABS-016 | 日直は月1回まで | 日直枠（昼系: 土曜昼C/日曜昼E/祝日昼G ＋ 支援日直J）に同一医師が月2回以上（v6.10.0） | `collect_candidates`で除外 |
| GAIKIN-EX | 外勤例外ペアの列限定 | 外勤（Sheet3出張日）の**前日+当日は自動不可(0)**。`config.GAIKIN_EXCEPTIONS` に該当する(外勤先, 相対日)のみ許容列に限り配置可（v6.11.0/提案#3+#4） | `get_avail_code`=0（全面不可）＋ `is_travel_col_forbidden`（列限定）で除外 |
| TEAM-001 | チーム枠種制約 | `config.TEAM_SLOT_RESTRICTIONS` の医師集合×曜日×枠種別(日直/当直)×列範囲(大学/外病院)に配置（v6.11.0/提案#6。ABS-005はこの一般機構の特殊例で共存） | `is_team_slot_forbidden`で除外 |

**補足:**
- **ABS-009**（未割当禁止）は最も深刻な違反。fix_unassigned_slots が最終セーフティネットとして緊急フォールバックで埋める。物理的に埋まらない場合のみ未割当のまま残り、警告表示される。
- **ABS-014**（平日/休日偏り差≤1）は evaluate 側では W=300 のソフト残留（`wd_we_imbalance_violations`）としても計上される（実質ハード）。

**v5.2からの変更点:**
- ABS-007, ABS-008, ABS-010: v6.0.0でSOFT/HARDからABS（絶対禁忌）に格上げ
- ABS-009: v6.0.4で未割当を最終パスで直接補填する扱いに（`validate_absolute_constraints`で検査）
- ABS-011, ABS-012: v6.0.0/v6.5.0で追加
- ABS-013: v6.5.3でSEMI-002から格上げ
- ABS-014: v6.8.0で生成時フィルタを追加（平日/休日偏り）
- ABS-015: v6.5.6で追加
- ABS-016: v6.10.0で追加（日直月1回）
- ABS-011/ABS-012: v6.10.0で改定（学年クォータ制の医師別上限 / 暦週(日〜土)1回ルール）
- GAIKIN-EX/TEAM-001: v6.11.0で追加（可否レイヤ3件）
  - 出張日の自動不可は**前日のみ→前日+当日**に拡張。セル内複数曜日（「木・金」等）と
    追加列（「外勤曜日2」等、`(出張|外勤)(日|曜日)N` パターン）に対応
  - `GAIKIN_EXCEPTIONS`/`GAIKIN_HOSPITAL_GROUPS`（config.py）で外勤先×相対日(offset 0/-1)×許容当直先の緩和
  - 連続外勤で同一日が複数オフセットに該当する場合は許容列の**積集合**のみ許容
  - 固定割当（sheet1に医師名を直書き）は意図的配置として GAIKIN-EX / 外勤由来ABS-001 の検査から除外
  - 枠種別「日直」は C/E/G/J 列のみ存在（外病院列に日直/当直の区別は無い）。
    `range: "external"` と組み合わせる場合は `slot_type` を省略するか「当直」を指定する

> **コード正本**: 上表の完全な検査実装は `main.py::validate_absolute_constraints()`（ABS-001〜016を1関数で走査）。
> `CONSTRAINT_ABS_00X` 定数（main.py §「制約ID定義」）は静的チェック用の一部IDのみを保持する（ABS-001〜006, 013, 015）。

### 2.2 ハード制約（レガシー・ほぼABSへ吸収済み）

v6.0.0以降、旧「HARD」制約の実体は **ほぼ全てABS（絶対禁忌）へ格上げ**された。
`main.py` の `CONSTRAINT_HARD_00X` 定数は歴史的経緯で残るラベルであり、スコアリングには使われない。
現行コードで「HARD」IDが指す内容（`CONSTRAINT_HARD_001〜004`, main.py §制約ID定義）は以下の通りだが、
いずれも実際にはABS側で強制される:

| ID | 定数上の意味 | 実際の強制先 |
|----|-------------|-------------|
| HARD-001 | TARGET_CAP超過 | ABS-010 |
| HARD-002 | gap違反 | ABS-007 |
| HARD-003 | 未割当枠 | ABS-009（`fix_unassigned_slots`） |
| HARD-004 | CODE_2のn+1違反 | ABS-010（TARGET_CAPベース判定・SOFT `code_2_extra_violations` W=300でも計上） |

> つまり運用上「ハード制約」という独立階層は事実上消滅している。B/I列・C-H/J-K列の「グループ1回まで」は
> 現行では ABS-011（大学系2回まで）＋ ABS-012（大学系7日間隔）＋ evaluate側の `bg_over_2`/`bg_weekday_over` ペナルティで担保する。

### 2.3 準ハード制約（条件付き緩和可）

選択肢がない場合に限り緩和される:

| ID | 制約名 | 緩和条件 | 備考 |
|----|--------|----------|------|
| SEMI-001 | B列カテ表コード必須 | 属性1の医師は週1回まで許容。属性2は緩和不可（ABS該当） | v6.5.6: 属性で緩和可否を判定 |

**v5.2からの変更点:**
- SEMI-002: v6.5.3でABS-013に格上げ
- SEMI-003（gap制約）: v6.0.0でABS-007に格上げ（削除）
- SEMI-004（大学最低1回）: `fix_university_minimum_requirement`で対応（削除）

**段階的制約緩和フロー（v6.3.0）:**
```
1. 全制約適用でチェック
   ↓ 候補なし
2. relax_semi=True で再試行（SEMI-001緩和）
   ↓ 候補なし
3. relax_hard=True で再試行（HARD制約緩和）
   ↓ 候補なし
4. relax_abs=True で再試行（ABS-007/008/010/011/012緩和）
   ※ ABS-001〜006, ABS-013は緩和不可
```

### 2.4 ソフト制約（ペナルティ加算）

> **コード正本 = `main.py::evaluate_schedule_with_raw()` の penalty 加算（main.py 2737〜2761行）**。
> `raw_score = 100 − penalty`、`score = max(raw_score, 0)`。
> 下表は実際に加算される全項目を「適用重み」で列挙したもの。
> **重要**: main.py の `CONSTRAINT_SOFT_00X` 定数コメントの重み、config.py のコメント、本表は
> 歴史的経緯で **ID割当が3系統に分裂している**。動作を規定するのは本表の「適用重み」列のみ。
> ID列は参考（main.py定数ブロックの登録簿に対応、無いものは「—」）。

| 適用重み | メトリクス（変数名） | 対象 | 定数/ID | 由来 |
|---------|---------------------|------|---------|------|
| **500** | `unassigned_slots` | 未割当スロット数 | `W_UNASSIGNED` | ABS-009（最優先。事実上ハード） |
| **300** | `code_2_extra_violations` | CODE_2医師のTARGET_CAP超過 | ハードコード | HARD-004→ABS-010 |
| **300** | `bg_over_2_violations` | 大学系3回以上（CC除外）の超過分 | ハードコード | SOFT-002相当・大学3回以上 |
| **300** | `ht_0_violations` | 外病院0回かつ大学≥1回 | ハードコード | SOFT-001相当・外病院0回 |
| **300** | `weekly_bg_violations` | 大学系7日間隔違反（残留分） | ハードコード | ABS-012残留 |
| **300** | `wd_we_imbalance_violations` | 平日/休日差≥2の超過分 | ハードコード | ABS-014残留 |
| **300** | `we_0_violations` | 休日0回かつ総≥1回 | ハードコード | 休日0回（事実上ハード） |
| **150** | `code_1_2_violations` | CODE_1.2医師が大学系0回 | ハードコード / `W_CODE_12_UNIV`(=150, greedy未使用) | 大学最低1回未達 |
| **120** | `ch_kate_violations` | C-H列にカテ当番日以外で配置 | ハードコード | 休日大学系カテ当番（ABS-013の残留・固定割当は許容） |
| **100** | `bg_ht_imbalance_violations` | 大学/外病院の差≥3（CC除外）超過分 | ハードコード / `W_BG_HT_DIFF`(=100, greedy未使用) | BG/HT不均衡 |
| **80** | `bg_weekday_over_violations` | 大学の平日2回以上の超過分 | ハードコード | 大学平日偏り |
| **50** | `bg_weekday_weekend_imbalance` | 大学2回で平日1+休日1でない | ハードコード | 大学2回バランス |
| **30** | `fairness_penalty` | 全合計公平性（active・CC除外、差≥2で2倍） | `W_FAIR_TOTAL` | 公平性 |
| **2** | `bk_ly_imbalance` | B-K/L-Y比率の偏り合計（コード3除外） | `W_BK_LY_BALANCE` | 比率バランス |
| **0** | `cum_total_spread` | 累計（前月+今月）全合計spread | `W_FAIR_CUM`(既定0) | 累計公平性（>0で有効化・表示のみ） |
| **0** | `gap_violations` | gap<3日 | `W_GAP` | ABS-007で強制済み |
| **0** | `hosp_dup_violations` | 大学系同一病院重複 | `W_HOSP_DUP` | 許容 |
| **0** | `external_hosp_dup_violations` | 外病院同一病院重複 | `W_EXTERNAL_HOSP_DUP` | ABS-008で強制済み |
| **0** | `cap_violations` | TARGET_CAP超過 | `W_CAP` | ABS-010で強制済み |
| **0** | `bg_spread`/`ht_spread`/`wd_spread`/`we_spread` | 累計ばらつき | `W_BG_SPREAD`他 | 簡略化で削除 |

**注記:**
- `code_1_2_violations`・`bg_ht_imbalance_violations` は evaluate 内でハードコード重み（150 / 100）が使われており、config.py の `W_CODE_12_UNIV`・`W_BG_HT_DIFF` は**現状 greedy スコアには反映されない**（値は同じ150/100なので実害なし。将来の整理対象）。
- W=300群は「事実上のハード制約」。ABSで生成時に排除されるため通常0だが、緩和フォールバック経路で残った場合にスコアを強く下げる。

### 重み定数一覧（main.py 383〜399行・config.py で上書き可）

```python
# v6.0.0: GAP/CAP/外病院重複はABS（絶対禁忌）に格上げされペナルティ重み0
W_FAIR_TOTAL   = getattr(_cfg, 'W_FAIR_TOTAL', 30)     # 公平性（max-min最小化）
W_FAIR_CUM     = getattr(_cfg, 'W_FAIR_CUM', 0)        # v6.5.9: 累計公平性（既定0=表示のみ）
W_CODE_12_UNIV = getattr(_cfg, 'W_CODE_12_UNIV', 150)  # ※greedyスコアはハードコード150を使用
W_BG_HT_DIFF   = getattr(_cfg, 'W_BG_HT_DIFF', 100)    # ※greedyスコアはハードコード100を使用
W_GAP                = 0    # ABS-007で対応
W_HOSP_DUP           = 0    # 大学重複（許容）
W_EXTERNAL_HOSP_DUP  = 0    # ABS-008で対応
W_UNASSIGNED   = getattr(_cfg, 'W_UNASSIGNED', 500)    # 未割当ペナルティ
W_CAP                = 0    # ABS-010で対応
W_BG_SPREAD = W_HT_SPREAD = W_WD_SPREAD = W_WE_SPREAD = 0  # 削除（簡略化）
W_BK_LY_BALANCE = getattr(_cfg, 'W_BK_LY_BALANCE', 2)  # B-K/L-Y比率バランス
```

---

## §3 配置ロジック

> 注意: 本セクションは制約定義ではなく、配置アルゴリズムの仕様です。

### 3.1 配置優先順序

```
1. 大学休日（C-H列）← カテ当番制約があるため先に配置
2. 大学平日（B列、I-K列）
3. 外病院（L-Y列）
```

### 3.2 休日大学系（C-H列）の配置ルール

**配置可能条件（OR条件）:**
1. その日にカテ当番がある（Sheet1:Zのチームコードが医師のチーム属性と一致）
2. カテなし医師（カテチーム属性を持たない）

> **関連制約**: ABS-013（v6.5.3でSEMI-002から格上げ。緩和不可）

### 3.3 平日大学系（B列 + I-K列）の配置ルール

**配置可能条件（OR条件）:**
1. カテなし医師
2. その日にカテ当番がある
3. 属性1の医師 / sheet3で「1」を持つ医師（SEMI-001緩和対象）

> **関連制約**: SEMI-001
> **属性による緩和可否（v6.5.6）:**
> - 属性1: 緩和可（週1回までカテ表なしでB列配置を許容）
> - 属性2: 緩和不可（ABS-015として扱い、カテ表必須）

### 3.4 修正パイプライン（実行順序）

```
#1  fix_hard_constraint_violations()        # 絶対禁忌の修正
#2  fix_code_2_extra_violations()           # CODE_2のTARGET_CAP超過
#3  fix_target_cap_violations()             # TARGET_CAP超過
#4  fix_university_minimum_requirement()    # 大学最低1回必須
#5  fix_ch_kate_violations()                # C-Hカテ当番違反
#6  fix_bg_ht_imbalance_violations()        # BG/HT不均衡
#7  fix_gap_violations()                    # gap違反
#8  fix_external_hospital_dup_violations()  # 外病院重複
#9  fix_university_over_2_violations()      # 大学3回以上
#10 fix_university_weekday_balance()        # 大学平日偏り
#11 fix_fairness_imbalance()                # 公平性
#12 fix_unassigned_slots()                  # 未割当枠
```

**順序設計の根拠:**

| フェーズ | 修正関数 | 依存関係・理由 |
|---------|----------|---------------|
| **Phase 1: 絶対禁忌** | #1 | 最優先。他の全修正の前提条件 |
| **Phase 2: CAP系** | #2-#3 | CAPが正しくないと他の均衡計算が狂う |
| **Phase 3: 必須配置** | #4-#5 | 大学最低1回・カテ制約を満たす |
| **Phase 4: バランス** | #6-#10 | BG/HT, gap, 重複を順次調整 |
| **Phase 5: 微調整** | #11-#12 | 公平性調整と残枠埋め（最後） |

**収束ループ（v6.0.3）:**
各fix関数は `safe_fix` ラッパーで実行され、収束するまで最大3回ループする。
fix関数がスコアを悪化させた場合は変更をrevertし、元のパターンを保持する。

---

## §4 inactive医師の扱い

### 定義
```python
def is_always_unavailable(doc):
    # 事前割当があれば対象外
    if preassigned_count.get(doc, 0) > 0:
        return False
    # 全日コード0ならinactive
    return all(get_avail_code(d, doc) == 0 for d in all_shift_dates)

inactive_doctors = [d for d in doctor_names if is_always_unavailable(d)]
```

### 処理方針

| フェーズ | inactive医師の扱い |
|----------|-------------------|
| **解析処理** | 完全除外（候補に含めない） |
| **TARGET_CAP計算** | 除外（active医師のみで計算） |
| **パターン生成** | 除外 |
| **最適化** | 除外 |
| **出力Excel** | 0回として記載（氏名は表示） |

---

## §5 クイックリファレンス

### 制約ID一覧

| カテゴリ | ID範囲 | 説明 |
|---------|--------|------|
| **ABS** | ABS-001〜015（009/014含む・欠番なし） | 絶対禁忌（配置不可）。定義=`validate_absolute_constraints()` |
| **HARD** | HARD-001〜004（レガシー） | 実体はABSへ吸収済み（§2.2） |
| **SEMI** | SEMI-001 | 準ハード制約（緩和可）。SEMI-002はABS-013へ格上げ済 |
| **SOFT** | 適用重みで規定（§2.4） | ID体系は3系統に分裂・重みは `evaluate_schedule_with_raw()` が正本 |

### 列インデックス対応表

| 列名 | Index | 分類 | カテ制約 | 関連制約ID |
|------|-------|------|---------|-----------|
| B | 1 | 平日大学系 | 緩和可 | SEMI-001, ABS-015(属性2) |
| C-H | 2-7 | 休日大学系 | カテ当番必須 | ABS-013 |
| I-K | 8-10 | 平日大学系 | 緩和可 | SEMI-001 |
| L-Y | 11-24 | 外病院 | カテ当番日禁止 | ABS-004 |

### 医師分類早見表

| 分類 | 判定条件 | 主な制約ID |
|------|----------|-----------|
| active | sheet2で1日でも≠0 | 全制約対象 |
| inactive | sheet2全日=0 | 解析除外、出力のみ |
| CODE_2 | sheet2にコード2あり | ABS-002 |
| CODE_3 | sheet2にコード3あり | ABS-003 |
| CODE_1.2 | sheet2にコード1.2あり | 大学最低1回必須 |
| カテ保有 | Sheet4のカテチーム属性がSheet1:Zと一致 | ABS-004, ABS-013 |
| カテなし | カテチーム属性なし | - |
| 属性1 | Sheet4属性=1 | SEMI-001緩和可 |
| 属性2 | Sheet4属性=2 | ABS-015（B列カテ表必須） |

---

## §6 ソルバーエンジンの切替（v6.9.0〜）

`config.SOLVER` で当直表生成エンジンを選択する。

| 値 | エンジン | 制約の扱い | 生成数 |
|----|---------|-----------|--------|
| `"cpsat"`（既定） | CP-SAT厳密解（`solver_cpsat.py`） | ABS/SEMIをハード制約として直接モデル化し、本書§2.4のSOFT重みで目的関数を最小化 | 常に3解（`NUM_PATTERNS`無視） |
| `"greedy"` | 旧Greedy + fixパイプライン（`main.py`内） | ABSは候補選定/fixで強制、SOFTはevaluateでペナルティ加算 | `NUM_PATTERNS`個から上位3 |

**フォールバック**: `SOLVER="cpsat"` でも次のいずれかで自動的にGreedyへ切替（stdoutに明示）:
- `ortools` 未導入（ImportError）
- INFEASIBLE / CP-SAT内部例外
- CP-SATとmain.pyの `TARGET_CAP` / `EXTRA_ALLOWED` 不一致（二重パース整合性assert失敗）

**合流点**: CP-SATの割当は main.py の `shift_df` 同型グリッド（`pattern_df`）へ転写され、
以降の evaluate / summary / バナー / Excel出力は**両エンジンで共通**。したがってスコア・診断シートの意味は
どちらのエンジンでも本書§2.4と一致する（CP-SAT解は通常 penalty が最小＝全ABS違反0）。
実装: `main.py::run()` 6180〜6273行。

---

## 更新履歴

| 日付 | バージョン | 変更内容 |
|------|-----------|---------|
| 2026-07-22 | v6.10.0 | 作成者実運用ルールの反映（コア配分3件）: ①学年クォータ制（Sheet3「大学目標」「外目標」列の直読み・ABS-010/011を医師別目標に置換・Σ検算プリフライト） ②ABS-012をローリング7日→暦週(日曜始まり〜土曜)1回に変更（月初第1週の前月重複は手動確認の警告） ③ABS-016新設（日直=昼系C/E/G+支援日直Jは月1回まで）。CP-SAT/Greedy両エンジンに同一実装 |
| 2026-07-22 | v6.9.0 | コード正本へ全面同期: §2.4 SOFTを `evaluate_schedule_with_raw()` の実適用重みで全面書き換え（3系統ID分裂を明記）、ABS-009/ABS-014を§2.1表に追加、§2.2 HARDをABS吸収済みとして整理、§6にCP-SAT/Greedy切替を追記、ヘッダのコード正本参照を更新（colab版はアーカイブ） |
| 2026-02-12 | v6.5 | v6.x系の変更を反映: ABS-007/008/010のABS格上げ、ABS-011〜015追加、SEMI-002〜004削除（格上げ/統合）、重み定数をv6.0.0仕様に更新（GAP/CAP/外病院重複=0）、段階的制約緩和（relax_abs）、修正パイプライン更新、属性による緩和可否、safe_fixラッパー |
| 2026-01-31 | v5.2 | 制約ID体系（ABS/HARD/SEMI/SOFT-XXX）を導入。緩和フラグ命名の歴史的経緯を注記。§1.2にJ-K列の設計背景を追記。§4.3の表記を「平日大学系」に統一。§6に制約ID一覧を追加 |
| 2026-01-31 | v5.1 | ハード/ソフト境界を明確化（§2.2, §2.4再構成）。§1.5にSheet3コード定義追加。水曜日L〜Y禁止制約を§2.1に追加。CODE_1.2と準ハード#4の関係性を整理。修正パイプラインの順序根拠を追記 |
| 2026-01-30 | v5.0 | 制約仕様書を新規作成。用語定義、列分類、ペナルティスケールを明確化 |
