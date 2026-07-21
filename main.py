# 当直くん v6.5.6 - ローカル実行版
# 元のGoogle Colabノートブックをローカル実行用に変換したものです。
# 詳細なバージョン履歴は VERSION_HISTORY.md を参照してください。
# 修正内容:
# v6.5.6 (2026-02-06):
# - SEMI-001の緩和可否を属性で判定
#   - カテ持ち + 属性1: 緩和可（週1回まで許容）
#   - カテ持ち + 属性2: 緩和不可（ABS該当、カテ表必須）
#   - カテなし: その他のルールに従う（B列禁止ではない）
#   - Sheet4の「属性」列を別途読み込み（doctor_attribute辞書）
# v6.5.5 (2026-02-06):
# - カテチーム判定をSheet1:Zのチームコード(A,B,C,D,E等)に限定
#   - 属性列の"2"などはカテ持ちとして扱わない
#   - expected_team_codesに含まれる値のみがカテチーム
#   - これによりSEMI-001は本当のカテ持ち医師にのみ適用
# v6.5.4 (2026-02-06):
# - SEMI-001（B列カテ表なし）を月曜始まりの週で1回まで許容
#   - 土日は同じ週（NG）、日月は別の週（OK）
#   - semi001_violation_weeksで週ごとの違反を追跡
# v6.5.3 (2026-02-06):
# - SEMI-002をABS-013に格上げ（C-H列休日大学系カテ当番必須）
#   - C-H列（休日大学系）でのカテ当番チェックを絶対禁忌に変更
#   - 緩和不可の制約として厳守（カテ当番なしで土日大学系には入れない）
# - 出力summaryの「今月サマリー」と「医師ごとの偏り」を1テーブルに統合
# v6.5.2 (2026-02-06):
# - カテチーム列の検出ロジックを改善
#   - Sheet1:Zのチームコード（A,B,C,D,E等）と一致する値を持つ列を自動検出
#   - 「属性」列に別のデータ（例: "2"）が入っている場合でも「カテ当番」列を正しく使用
#   - 検出優先順位: カテ当番 > 属性 > カテ > チーム
# v6.5.1 (2026-02-05):
# - SEMI制約の違反検出バグを修正
#   - SEMI-001: B列のみ（従来はB-K列全体を誤検出していた）
#   - SEMI-002: C-H列のみ（I-K列は対象外として正しく除外）
#   - これにより大量のSEMI違反が誤検出される問題を解消
# v6.5.0 (2026-02-05):
# - Excel構造変更に対応
#   - Sheet1:Z列「カテ当番」を病院列から除外し、チーム当番日として使用
#   - Sheet4:B列「属性」からカテチーム属性を読み込み
#   - get_sched_code()をSheet1:Z + Sheet4:属性で判定（Sheet3でオーバーライド）
#   - Sheet2の休み希望（1,2,3優先度）に対応（空欄=制約なし、0=不可、1=最優先休み、2=できれば休み、3=可能なら休み）
#   - 出張曜日の前日を自動的に不可（0）として扱う
# - 大学系7日間隔ルールに変更（ABS-012改）
#   - 日曜〜土曜の週単位から「7日以内に2回禁止」に変更
#   - 3/11と3/15のような週をまたぐケースも検出
# v6.4.0 (2026-02-04):
# - 大学系週1回ルール（ABS-012）を実装
#   - 日曜〜土曜の週単位で大学系（B-K列）は1回まで
#   - get_bg_week_start()で週の開始日（日曜日）を計算
#   - assigned_bg_weeksで週別の大学系割当をトラッキング
#   - collect_candidatesにABS-012チェックを追加
#   - fix_weekly_bg_violations()で週1違反を修正
#   - build_weekly_bg_details()で週1違反の詳細を診断
#   - 最適化パイプラインに週1違反修正を追加
# v6.3.0 (2026-02-04):
# - 未割当を許容せず、ABS制約も段階的に緩和して必ず割当
#   - 第4段階緩和: relax_abs=True を追加
#   - gap>=3、同一病院重複、TARGET_CAP、大学系2回制限を緩和可能に
#   - 同日重複(ABS-006)とコード0(ABS-001)のみ絶対禁止として維持
# - Excel出力のシート整理
#   - sheet1-4（元データ）を出力から削除
#   - 今月/累計/diagを1シートに統合（{pattern}_summary）
# v6.2.0 (2026-02-04):
# - 固定割当（事前割当）を制約違反の検出・ペナルティ計算から除外
#   - sheet1に医師名が直接記載された固定割当は意図的な配置であるため、
#     ABS/SEMI/SOFT全ての制約チェック対象から除外する
#   - 固定割当が片方でも含まれるgap<3はペナルティ対象外
#   - 固定割当による同一病院重複もペナルティ対象外
#   - build_hard_constraint_violationsで固定割当スロットをスキップ
#   - C-H列カテ当番違反チェックでも固定割当をスキップ
# - fix関数が固定割当を移動・削除しないよう保護
#   - is_preassigned_slot()ヘルパー関数を追加
#   - 全fix関数（gap/dup/cap/code2/code12/imbalance/over2/weekday）で
#     固定割当スロットを移動対象から除外
#   - 固定割当は意図的な配置のため、最適化で変更してはならない
# v6.1.0 (2026-02-04):
# - gap>=3上限によるTARGET_CAP自動調整
#   - 各医師の利用可能日分布からgap>=3で可能な最大割当数を計算
#   - CAPが物理上限を超えている場合は切り下げ（gap違反の根本原因を除去）
#   - 余った枠は他の余裕ある医師に再配分
# - safe_fix改善: gap(ABS-007)/dup(ABS-008)の個別増加もrevert対象
#   - 他のfix関数がgap/dupを犠牲にして別制約を直す「トレード」を防止
# - 収束ループ後にgap/外病院重複の最終パスを追加（safe_fix不使用）
#   - safe_fixがrevertしたgap/dup修正を最終的に解消
#   - 順序: gap修正 → 外病院重複修正 → 未割当修正（依存関係順）
# - diagシート再構成:
#   - 1.医師ごとの偏り 2.制約違反一覧 3.スコアサマリー の順序に変更
#   - メトリクスを日本語の「スコアサマリー」に置換（項目/値/説明の3列）
#   - 医師テーブルに「gap3上限」「利用可能日数」列を追加
# v6.0.5 (2026-02-04):
# - CODE_2医師もEXTRA_ALLOWED対象に含める（枠不足解消）
#   - CODE_2除外だと他医師の制約(gap/dup等)で物理的に枠が足りなくなる
#   - CODE_2医師のn+1回目はB〜Q列（大学系）に割り当て可能
#   - code_2_extra_violationsをTARGET_CAPベースに変更（BASE_TARGET→TARGET_CAP）
#   - fix_code_2_extra_violationsも同様にTARGET_CAPベースに変更
# v6.0.4 (2026-02-04):
# - 収束ループ後にfix_unassigned_slotsの最終パスを追加（safe_fix不使用）
#   - safe_fixがrevertした未割当てや、他fix関数が作った未割当てを確実に埋める
#   - ABS-009（未割当て）は最も深刻な違反のため、最終パスでは直接実行
# v6.0.3 (2026-02-03):
# - 根本改善: safe_fixラッパー + 収束ループで最適化を再有効化
#   - is_valid_full_assignment(): 全ABS制約を統合チェックする関数を追加
#   - safe_fix(): fix関数実行後にABS違反が増えたらrevertする安全ラッパー
#   - 全fix関数をsafe_fixで実行（ABS違反の連鎖を構造的に防止）
#   - 収束するまで最大3ラウンド繰り返す
#   - OPTIMIZATION_ENABLED=True に復帰
# v6.0.2 (2026-02-03):
# - パターン多様性の向上
#   - tie-breakをdeterministic(右側優先)からrandom.choiceに変更
#   - 出力パターン3つを3軸評価（公平性・gap回避・バランス）で選択
#   - 同じseedの重複を排除して異なるパターンを提示
# - fix_unassigned_slots内にABS-010/ABS-011チェック追加
# - fix_unassigned_slots内の追跡変数を正確に更新
# - W_UNASSIGNED=500に設定
# v6.0.1 (2026-02-02):
# - 段階的制約緩和を実装（ABS-009回避優先）
#   1. 全制約適用 → 候補なし
#   2. SEMI緩和 → 候補なし
#   3. HARD緩和 → 候補なし
#   4. 緊急フォールバック（ABSのみ）
# - 全フォールバック関数にABS-004, ABS-005チェック追加
# - gap制約をgap>=3に統一（以前はgap<2だった箇所を修正）
# v6.0.0 (2026-02-02):
# - 制約体系を全面改定
# - 絶対禁忌(ABS): 12項目
#   - ABS-001〜009: 既存の絶対禁忌
#   - ABS-010: TARGET_CAP遵守（n超過禁止）
#   - ABS-011: 大学系2回まで（B-K列合計）
#   - ABS-013: C-H列カテ当番必須（v6.5.3で追加、旧SEMI-002）
# - ハード制約(HARD): 3項目
#   - HARD-001: B/I列1回まで（グループA）
#   - HARD-002: C-H/J-K列1回まで（グループB）
#   - HARD-003: 外病院1回以上（L-Y列）
# - 準ハード制約(SEMI): 1項目
#   - SEMI-001: B列のみカテ表コード必須（sheet3「1」は例外）
# - ソフト制約(SOFT): 3項目
#   - SOFT-001: 公平性（max-min最小化）
#   - SOFT-002: コード1.2優先（大学系0回ペナルティ）
#   - SOFT-003: 大学/外病院差（差3以上ペナルティ）
# - 不要な制約を削除（HARD/ABSで吸収）
# - ABS-001（コード0禁止）修正
#   - fix_hard_constraint_violations()の緊急フォールバックでコード0チェック追加
#   - 最終手段でもコード0医師を除外
#   - 全員コード0の場合は未割当のまま（違反割当より優先）
# - inactive医師処理にドキュメント参照を追加
#   - CONSTRAINT_RULES.md §5準拠コメント
# v4.7 (2026-01-31):
# - CONSTRAINT_RULES.md v5.2仕様に基づく整備
# v4.3 (2026-01-30):
# - C-H列カテ当番制約をソフト制約に変更（ハード制約から除外）
#   - 適格医師不足時のパターン除外を防止
#   - ペナルティ(120)とfix関数は維持
#   - 修正不可でもパターン選択可能に
# - recompute_stats呼び出しのunpacking修正（*_追加）
# v4.2 (2026-01-30):
# - 平日大学系(B,I-K列)の制約を緩和
#   - sheet3で「1」を持つ医師はカテ当番が合わなくても許容
#   - カテ当番なし医師も配置可能
#   - is_weekday_university_slot()、is_eligible_for_weekday_university_slot()関数を追加
#   - SHEET3_CODE_1_DOCTORSセットを追加
# - CC（大型連休特別シフト）を特別カウントとして扱う
#   - is_cc_assignment()関数を追加（CCかどうか判定）
#   - recompute_statsでCC別カウント（cc_counts, cc_bg_counts, cc_ht_counts）を追跡
#   - 以下の計算からCCを除外:
#     - 公平性計算（fairness_penalty）
#     - BG/HT不均衡計算（bg_ht_imbalance_violations）
#     - 外病院重複計算（external_hosp_dup_violations）
#     - 大学3回以上違反（bg_over_2_violations）
#   - 各fix関数でもCC除外を適用
# v4.1 (2026-01-30):
# - 枠決定順序を最適化
#   - 大学休日(C-H列) → 大学平日(B,I-K列) → 外病院(L-Y列) の順に割り当て
#   - C-H列はカテ当番制約があるため先に埋めることで制約を満たしやすくする
#   - slot_priority関数を追加して優先度順にソート
# - ハード制約チェックにC-H列カテ当番違反を追加
#   - ch_kate_violations > 0 のパターンを除外
# v4.0 (2026-01-30):
# - C-H列（休日大学系）のカテ当番制約を追加
#   - C-H列はカテ当番のある医師（その日にアルファベットあり）または
#     カテ当番が一回もない医師（sheet3に1つもアルファベットなし）のみ割り当て可能
#   - NO_KATE_DOCTORSセットを追加（カテ当番なし医師の集合）
#   - is_ch_slot()、is_eligible_for_ch_slot()関数を追加
#   - collect_candidatesにrelax_ch_kateパラメータを追加
#   - fix_ch_kate_violations()関数を追加（違反修正用）
#   - ch_kate_violationsメトリクスを追加（ペナルティ120）
#   - 最適化パイプライン#4.5に追加（大学最低1回の後、gap違反の前）
# v3.9 (2026-01-30):
# - print出力の最適化
#   - tqdmによる進捗バー表示（パターン生成、局所探索）
#   - セクション区切りの統一（=== ===形式）
#   - 階層構造表示（├─/└─）
#   - TOPパターン評価をテーブル形式で表示
#   - 冗長な出力を削減
# - 出力ファイル名にバージョンを反映（filename_v3.9.xlsx）
# - VERSION定数を追加
# v3.8 (2026-01-30):
# - 外病院最低1回をハード制約として追加（大学3回以上を防止）
#   - fix_university_over_2_violationsを拡張して外病院0回も検出・修正
#   - ht_0_violationsメトリクスを追加（ペナルティ300）
#   - ハード制約チェックにbg_over_2_violations、ht_0_violationsを追加
# - 最適化パイプライン順序を修正
#   - BG/HT不均衡(#6) → 外病院重複(#7) の順序に変更
# - 処理番号の表示を追加 [X/15]
# v3.7 (2026-01-30):
# - build_hard_constraint_violationsのreturn文欠落バグを修正
# - CODE_2医師のn+1回違反を最適化後にチェック・修正する機能を追加
#   - fix_code_2_extra_violations関数を追加（ハード制約として修正）
#   - evaluate_schedule_with_rawにcode_2_extra_violationsメトリクスを追加
#   - ハード制約チェックにCODE_2 n+1違反を追加
#   - 最適化パイプラインの#2に追加（ハード制約直後、TARGET_CAP前）
# v3.6 (2026-01-30):
# - 可否コード2の医師をEXTRA枠（n+1回）対象から除外（ハード制約）
#   - has_code_2_anywhere関数を追加（sheet2でいずれかの日に2を持つ医師を判定）
#   - CODE_2_DOCTORSリストを作成
#   - EXTRA_ALLOWEDの計算時にCODE_2_DOCTORSを除外
#   - 可否コード2の医師は大学系のみ可能なため、外病院枠増加は不適切
#   - 出力に「可否コード2医師（EXTRA枠対象外）」を表示
# v3.4 (2026-01-29):
# - TARGET_CAP違反の厳格化
#   - パターン選択時に cap_violations > 0 のパターンを除外
#   - gap_violations > 0、unassigned_slots > 0 も除外
#   - ハード制約を満たすパターンのみを候補として選択
# - 多軸スコアリングシステムを実装
#   - 公平性重視: TARGET_CAP、医師間の割当回数の公平性を最優先
#   - 連続当直回避重視: gap違反、外病院重複を最優先
#   - バランス重視: 大学/外病院バランス、平日/休日バランスを最優先
#   - 各軸から最良パターンを1つずつ選択し、合計3パターン出力
# - 公平性の強制修正機能を追加
#   - fix_fairness_imbalance: 最大割当回数と最小割当回数の差を縮める
#   - 4回の医師から1回の医師にシフトを移動し、公平性を達成
#   - 差が1以下になるまで修正（例: 1回と4回 → 2回と3回）
#   - 最適化パイプラインに統合（修正順序の最後に実行）
# v3.3 (2026-01-28):
# - 大学病院3回以上を禁止（不満が高い）
#   - bg_over_2_violations: 大学3回以上の違反を検出（ペナルティ150）
#   - fix_university_over_2_violations: 最適化後に大学3回以上を修正
#   - 大学の割当を外病院に移動、または削除して2回以下に制限
# - 大学病院の平日偏り制約を追加（平日2回以上は不満）
#   - bg_weekday_over_violations: 大学の平日2回以上の違反を検出（ペナルティ80）
#   - fix_university_weekday_balance_violations: 最適化後に平日偏りを修正
#   - 大学平日の割当を外病院に移動、または削除
# - 全体の公平性を強化（2回の医師がいるなら4回の医師から渡す）
#   - fairness_penalty計算を強化: diff_total >= 2の場合、2倍のペナルティ
#   - W_FAIR_TOTAL: 10 → 30（公平性の重要度を上げる）
#   - min=2, max=4のような差が大きい場合に強く制約
# - 修正パイプラインを拡充
#   - 最適化後に全ての制約違反を強制的に修正
#   - 順序: ハード制約 → TARGET_CAP（優先1） → 大学最低1回（準ハード、優先2） → gap（優先3） → 外病院DUP（優先4） → BG/HT → 大学3+ → 大学平日偏り → 公平性
# v3.2 (2026-01-28):
# - 生成パターン数をデフォルト100に戻す（処理時間の最適化）
#   - NUM_PATTERNS: 10000 → 100
#   - TOP_KEEP: 100 → 20
#   - REFINE_TOP: 20 → 15
# - 大学病院2回の場合、平日1回+休日1回のバランス制約を追加
#   - bg_weekday_weekend_imbalance 違反を検出
#   - ペナルティ: 50
# - 優先順位を厳格化（TARGET_CAP > gap > DUP を死守）
#   - W_CAP = 200（優先度1位）
#   - W_GAP = 100（優先度2位、3→100に強化）
#   - W_EXTERNAL_HOSP_DUP = 70（優先度3位）
# v3.1 (2026-01-28):
# - 外病院（L～Y列）重複を厳格化、大学病院（B～K列）重複は許容
#   - 評価関数で外病院重複と大学病院重複を分離
#   - W_EXTERNAL_HOSP_DUP=70（優先度3位：TARGET_CAP > gap > 外病院DUP）
#   - fix_external_hospital_dup_violations関数で最適化後に外病院重複を修正
#   - 同じ日の他の外病院に移動、または割当削除で修正
# v3.0 (2026-01-28):
# - gap違反（3日未満の間隔での割当）を完全に排除
#   - 初期パターン生成時にgap違反0個の候補のみ選択
#   - 局所探索でgap違反1以上になるswapを拒否
#   - fix_gap_violations関数で最適化後にgap違反を強制修正
#   - 移動先が見つからない場合は割当を削除して違反を解消
#   - 同じ病院だけでなく他の病院の空き枠も探索
# v2.8 (2026-01-24):
# - 大学系と外病院の差が3未満になる制約を追加
#   - 評価関数に差が3以上の場合のペナルティ追加（重み100）
#   - fix_bg_ht_imbalance_violations関数で最適化後に差3以上を修正
# - recompute_stats関数のBG/HT範囲を修正（B〜G→B〜K、H〜U→L〜Y）
#   - これにより出力Excelの今月/累計の大学合計・外病院合計が正しく計算される
# - 可否コード1.2の医師が大学系最低1回の制約を追加
#   - get_avail_code関数で1.2を認識できるよう修正
#   - fix_code_1_2_violations関数で最適化後に大学系0回を修正
#   - 評価関数に1.2の医師が大学系0回の場合のペナルティ追加（重み150）
# - TARGET_CAP違反の強制修正機能を追加
#   - fix_target_cap_violations関数で最適化後にTARGET_CAP超過を修正
#   - 上位医師（小林、及川等）の割当を下位医師（大河内、猪股等）に移動
#   - W_CAPペナルティを50→200に強化
# - 余り枠の割当ロジックを修正（昇順ソート→最後のEXTRA_SLOTS人を選択）
# - デバッグ情報追加：+1回対象の医師名、TARGET_CAP設定値、1.2対象医師を表示
# - デフォルトパターン数を100に変更（環境変数で上書き可能）
# v2.7 (2026-01-24):
# - ハード制約違反の自動修正機能を実装
#   - fix_hard_constraint_violations関数を追加
#   - 局所探索完了後に全ての違反を自動検出・修正
#   - 代替医師が見つからない場合は未割当として警告表示
#   - 修正後にスコアを再評価して最終結果に反映
# v2.6 (2026-01-23):
# - get_sched_code関数の重大なバグを修正
#   - "0"と"3"を有効なカテ表コードとして扱わないように変更
#   - "0"はデータなし、"3"は可否コードであり、カテ表コードではない
#   - これにより、カテ表コード保有医師がその日に"0"や"3"の場合、L〜Y列への割当が正しく許可される
# v2.5 (2026-01-21):
# - カテ表コードと列の制約を修正
#   - sheet3で少なくとも1つのカテ表コード（A,B,C,CC,D,E等）を持つ医師を特定
#   - カテ表コード保有医師: その日にコードがある場合のみB〜K列可、コードがない日は割当なし
#   - カテ表コード保有医師: その日にコードがある場合はL〜Y列禁止
#   - カテ表コード非保有医師: B〜K列に自由に割り当て可能
# - 出力パターンを1個から3個に変更（TOP3候補を提示）
# v2.4 (2026-01-21):
# - 列構造の変更対応（B〜Y列）
#   - 可否コード2: B〜Q列のみ可（従来B〜M列）
#   - 可否コード3: L〜Y列のみ可（従来H〜U列）
#   - カテ表制約: L〜Y列禁止（従来H〜U列）
# - B〜H列の2回上限制約を実装
# - 診断シートにB〜H列2回超過違反検出を追加
# v2.3 (2026-01-21):
# - B〜K列のカテ表要件をハード制約に変更（relax_scheduleで緩和不可）
# - B〜K列カテ表コード欠如違反の検出機能を追加
# - can_assign_doc_to_slot関数にB〜K列カテ表チェックを追加（局所探索でも適用）
# v2.2 (2026-01-21):
# - ハード制約違反の修正（可否コード0、カテ表+外病院、コード2/3違反）
# - collect_candidates関数でコード0を常に除外するよう修正
# - カテ表がある日のH〜U列割当を絶対禁止に変更
# - ハード制約違反チェック機能を診断シートに追加
# v2.1:
# - タイムゾーン問題の修正
# - 医師名の正規化（空白除去）
# - sheet4ヘッダ検出範囲の拡大（30→50行）
# - date_doc_countのKeyError対策（defaultdict化）
# - 同日重複チェックの強化
# - エラーハンドリングの改善

import io
import sys
import os
import pandas as pd
import numpy as np
from collections import defaultdict
import random

# バージョン定数
VERSION = "6.5.9"

# tqdmのインポート（進捗バー用）
try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    def tqdm(iterable, **kwargs):
        return iterable

# ローカル実行専用（Google Colab固有コードを除去）
COLAB_AVAILABLE = False

# =========================
# config.py から設定を読み込み
# =========================
import config as _cfg

HOLIDAYS = set()  # 後で pd.Timestamp に正規化する
BG_DAY_COLS = set()
BG_NIGHT_COLS = set()

WED_FORBIDDEN_DOCTORS = set(_cfg.WED_FORBIDDEN_DOCTORS)

NUM_PATTERNS = _cfg.NUM_PATTERNS

# sheet1 の「枠」扱いする入力値（1以外の記号も許容したい場合）
SLOT_MARKERS = {1, 1.0, "1", "〇", "○", "◯", "◎"}

# --- ローカル探索（入替）設定 ---
LOCAL_SEARCH_ENABLED = _cfg.LOCAL_SEARCH_ENABLED
OPTIMIZATION_ENABLED = _cfg.OPTIMIZATION_ENABLED
TOP_KEEP = 20                 # greedyで残す候補数（100パターンから上位20候補を保持）
REFINE_TOP = 15               # ローカル探索をかける候補数（上位15候補を最適化）
LOCAL_MAX_ITERS = 3000        # 1候補あたりの入替試行回数
LOCAL_PATIENCE = 1200         # 改善が出ない試行がこの回数続いたら打ち切り
LOCAL_REFRESH_EVERY = 200     # 問題医師（gap/重複）を再抽出する間隔

# v6.0.0 スコア重み（ソフト制約のみ）
# 絶対禁忌(ABS)とハード制約(HARD)は候補選定時にチェック済み
W_FAIR_TOTAL = getattr(_cfg, 'W_FAIR_TOTAL', 30)
W_FAIR_CUM = getattr(_cfg, 'W_FAIR_CUM', 10)  # v6.5.9: 累計（前月+今月）公平性
W_CODE_12_UNIV = getattr(_cfg, 'W_CODE_12_UNIV', 150)
W_BG_HT_DIFF = getattr(_cfg, 'W_BG_HT_DIFF', 100)
# 以下は絶対禁忌のためペナルティ不要（v6.0.0）
W_GAP = 0                  # ABS-007で対応
W_HOSP_DUP = 0             # ABS-008で対応
W_EXTERNAL_HOSP_DUP = 0    # ABS-008で対応
W_UNASSIGNED = getattr(_cfg, 'W_UNASSIGNED', 500)
W_CAP = 0                  # ABS-010で対応
W_BG_SPREAD = 0            # 削除（簡略化）
W_HT_SPREAD = 0            # 削除（簡略化）
W_WD_SPREAD = 0            # 削除（簡略化）
W_WE_SPREAD = 0            # 削除（簡略化）
W_BK_LY_BALANCE = getattr(_cfg, 'W_BK_LY_BALANCE', 2)

# =========================
# 制約ID定義（v5.2仕様書準拠）
# =========================
# 絶対禁忌（ABS: 配置不可）
CONSTRAINT_ABS_001 = "ABS-001"  # 可否コード0禁止
CONSTRAINT_ABS_002 = "ABS-002"  # コード2の列制約
CONSTRAINT_ABS_003 = "ABS-003"  # コード3の列制約
CONSTRAINT_ABS_004 = "ABS-004"  # カテ当番日の外病院禁止
CONSTRAINT_ABS_005 = "ABS-005"  # 水曜日L〜Y禁止医師
CONSTRAINT_ABS_006 = "ABS-006"  # 同日重複禁止
CONSTRAINT_ABS_013 = "ABS-013"  # v6.5.3: C-H列（休日大学系）カテ当番必須
CONSTRAINT_ABS_015 = "ABS-015"  # 属性2のB列カテ表コード欠如（緩和不可）

# ハード制約（HARD: パターン除外）
CONSTRAINT_HARD_001 = "HARD-001"  # TARGET_CAP超過
CONSTRAINT_HARD_002 = "HARD-002"  # gap違反
CONSTRAINT_HARD_003 = "HARD-003"  # 未割当枠
CONSTRAINT_HARD_004 = "HARD-004"  # CODE_2のn+1違反

# 準ハード制約（SEMI: 緩和可）
CONSTRAINT_SEMI_001 = "SEMI-001"  # 平日大学系カテ要件
CONSTRAINT_SEMI_002 = "SEMI-002"  # 休日大学系カテ当番（※ABS-013に格上げ、互換性のため残す）

# ソフト制約（SOFT: ペナルティ）
CONSTRAINT_SOFT_001 = "SOFT-001"  # 外病院0回 (W=300)
CONSTRAINT_SOFT_002 = "SOFT-002"  # 大学3回以上 (W=150)
CONSTRAINT_SOFT_003 = "SOFT-003"  # 外病院同一病院重複 (W=150)
CONSTRAINT_SOFT_004 = "SOFT-004"  # CODE_1.2大学0回 (W=150)
CONSTRAINT_SOFT_005 = "SOFT-005"  # gap違反 (W=100)
CONSTRAINT_SOFT_006 = "SOFT-006"  # BG/HT差3以上 (W=100)
CONSTRAINT_SOFT_007 = "SOFT-007"  # 大学平日2回以上 (W=80)
CONSTRAINT_SOFT_008 = "SOFT-008"  # 公平性 (W=30)
CONSTRAINT_SOFT_009 = "SOFT-009"  # 大学ばらつき (W=3)
CONSTRAINT_SOFT_010 = "SOFT-010"  # 外病院ばらつき (W=3)
CONSTRAINT_SOFT_011 = "SOFT-011"  # 休日ばらつき (W=3)
CONSTRAINT_SOFT_012 = "SOFT-012"  # 平日ばらつき (W=2)
CONSTRAINT_SOFT_013 = "SOFT-013"  # B-K/L-Y比率バランス (W=2)
CONSTRAINT_SOFT_014 = "SOFT-014"  # 大学同一病院重複 (W=0)

# =========================
# ユーティリティ
# =========================
def strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df

def make_unique(names):
    """重複列名を _2, _3 ... でユニーク化"""
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

def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

# 🔧 FIX: 医師名の正規化関数を追加
def normalize_name(name):
    """医師名を正規化（全角/半角スペース除去）"""
    if pd.isna(name):
        return ""
    return str(name).strip().replace(" ", "").replace("　", "")

def find_sheet_name(xls: pd.ExcelFile, target: str):
    """sheet名の大小・表記ゆれに耐える"""
    if target in xls.sheet_names:
        return target
    low_map = {s.lower(): s for s in xls.sheet_names}
    if target.lower() in low_map:
        return low_map[target.lower()]
    for s in xls.sheet_names:
        if s.strip().lower() == target.strip().lower():
            return s
    return None

def is_slot_value(v) -> bool:
    if isinstance(v, str):
        return v.strip() in SLOT_MARKERS
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v) == 1.0
    return False

# =========================
# sheet4 読み込み（ヘッダ行自動検出＋重複耐性）
# 🔧 FIX: 検索範囲を30→50行に拡大
# =========================
def parse_sheet4_from_grid(grid: pd.DataFrame) -> pd.DataFrame:
    g = grid.copy()
    g = g.dropna(how="all").reset_index(drop=True)

    if len(g) == 0:
        raise ValueError("sheet4 が空です")

    # ヘッダ行を探す（'氏名' がある行）
    header_row_idx = None
    search_limit = min(50, len(g))  # 🔧 FIX: 30→50に拡大
    for i in range(search_limit):
        row = g.iloc[i].astype(str).str.strip()
        if (row == "氏名").any():
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(f"sheet4 のヘッダ行に '氏名' 列が見つかりません（先頭{search_limit}行を検索）")

    headers = [safe_str(x) for x in g.iloc[header_row_idx].tolist()]
    headers = [h if (h != "" and h.lower() != "nan") else f"Unnamed_{j}" for j, h in enumerate(headers)]
    headers = make_unique(headers)

    data = g.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = headers

    if "氏名" not in data.columns:
        raise ValueError("sheet4 のヘッダ行に '氏名' 列が見つかりません（sheet4の形式を確認してください）")

    # 氏名の空行削除
    data["氏名"] = data["氏名"].astype(str).str.strip()
    data = data[(data["氏名"].notna()) & (data["氏名"] != "") & (data["氏名"].str.lower() != "nan")].reset_index(drop=True)

    # v6.5.0: 属性・出張日は文字列として保持
    STRING_COLS = {"属性", "カテ当番", "出張日", "出張先"}

    # 数値化（氏名・文字列列以外）
    for col in data.columns:
        if col == "氏名" or col in STRING_COLS:
            # 文字列として保持
            data[col] = data[col].astype(str).str.strip()
            data[col] = data[col].replace(["nan", "None", ""], "")
            continue
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    return data

# =========================
# 入力ファイルの読み込み（ローカル実行）
# =========================
print("\n" + "="*60)
print(f"  当直くん v{VERSION} (ローカル実行)")
print("="*60)

# 入力ファイルパスの決定（Finderダイアログ > コマンドライン引数 > config.py）
if len(sys.argv) > 1:
    input_path = sys.argv[1]
else:
    try:
        import subprocess as _sp
        _result = _sp.run(
            ["osascript", "-e",
             'POSIX path of (choose file with prompt "入力Excelファイルを選択してください" of type {"xlsx", "xls"})'],
            capture_output=True, text=True, timeout=300,
        )
        if _result.returncode == 0 and _result.stdout.strip():
            input_path = _result.stdout.strip()
        else:
            print("\nファイルが選択されませんでした。終了します。")
            sys.exit(0)
    except Exception:
        input_path = _cfg.INPUT_FILE

if not os.path.exists(input_path):
    print(f"\nエラー: 入力ファイルが見つかりません: {input_path}")
    print(f"data/ フォルダにExcelファイルを配置してください。")
    print(f"使い方: python main.py [入力ファイルパス]")
    sys.exit(1)

print(f"\n入力ファイル: {input_path}")

uploaded_filename = os.path.basename(input_path)

try:
    xls = pd.ExcelFile(input_path)
except Exception as e:
    raise ValueError(f"Excelファイルの読み込みに失敗しました: {e}")

sheet1_name = find_sheet_name(xls, "sheet1")
sheet2_name = find_sheet_name(xls, "sheet2")
sheet3_name = find_sheet_name(xls, "sheet3")
sheet4_name = find_sheet_name(xls, "sheet4") or find_sheet_name(xls, "Sheet4")

# v6.5.0: 新しいExcel構造対応
# Sheet4がない場合はSheet3を医師情報シートとして使用（旧Sheet3のカテ表は廃止）
if sheet4_name is None and sheet3_name is not None:
    sheet4_name = sheet3_name
    sheet3_name = None  # 旧カテ表は使用しない

missing = [k for k, v in [("sheet1", sheet1_name), ("sheet2", sheet2_name), ("sheet4/医師情報", sheet4_name)] if v is None]
if missing:
    raise ValueError(f"必要なシートが見つかりません: {missing}\n実際のシート名: {xls.sheet_names}")

# --------- Excel 読み込み ---------
shift_df = strip_cols(pd.read_excel(xls, sheet_name=sheet1_name))
availability_raw = strip_cols(pd.read_excel(xls, sheet_name=sheet2_name))

# v6.5.0: 旧カテ表(Sheet3)がない場合は空のDataFrameを使用
if sheet3_name is not None:
    schedule_raw = strip_cols(pd.read_excel(xls, sheet_name=sheet3_name))
    schedule_raw.columns = make_unique(list(schedule_raw.columns))
else:
    # カテ表はSheet1:Z + Sheet4:属性で代替するため空でOK
    schedule_raw = pd.DataFrame()

shift_df.columns = make_unique(list(shift_df.columns))
availability_raw.columns = make_unique(list(availability_raw.columns))

# sheet4 は「出力用」と「解析用（header=None）」を分ける
sheet4_raw_out = strip_cols(pd.read_excel(xls, sheet_name=sheet4_name))
sheet4_raw_out.columns = make_unique(list(sheet4_raw_out.columns))

sheet4_grid = pd.read_excel(xls, sheet_name=sheet4_name, header=None)
sheet4_data = parse_sheet4_from_grid(sheet4_grid)

# =========================
# 日付列の整形
# 🔧 FIX: タイムゾーン問題の修正
# =========================
date_col_shift = shift_df.columns[0]
shift_df[date_col_shift] = pd.to_datetime(shift_df[date_col_shift], errors="coerce").dt.normalize().dt.tz_localize(None)  # 🔧 FIX
if shift_df[date_col_shift].isna().all():
    raise ValueError("❌ sheet1 の先頭列が日付として解釈できません（列の形式を確認してください）")

date_col_avail = availability_raw.columns[0]
availability_raw[date_col_avail] = pd.to_datetime(availability_raw[date_col_avail], errors="coerce").dt.normalize().dt.tz_localize(None)  # 🔧 FIX
availability_df = availability_raw.set_index(date_col_avail)

# v6.5.0: schedule_rawが空の場合（新Excel構造）は空のDataFrameを使用
if len(schedule_raw.columns) > 0:
    date_col_sched = schedule_raw.columns[0]
    schedule_raw[date_col_sched] = pd.to_datetime(schedule_raw[date_col_sched], errors="coerce").dt.normalize().dt.tz_localize(None)
    schedule_df = schedule_raw.set_index(date_col_sched)
else:
    schedule_df = pd.DataFrame()
    date_col_sched = None

# config.py の祝日をセットに追加
for _h in _cfg.HOLIDAYS:
    HOLIDAYS.add(pd.Timestamp(_h))
# 🔧 FIX: 祝日もタイムゾーン正規化
HOLIDAYS = {pd.to_datetime(d).normalize().tz_localize(None) for d in HOLIDAYS}

def is_holiday(date):
    return pd.to_datetime(date).normalize().tz_localize(None) in HOLIDAYS

# =========================
# 基本情報
# 🔧 FIX: 医師名を正規化
# =========================
doctor_names = [normalize_name(x) for x in list(availability_raw.columns[1:])]  # 🔧 FIX
doctor_col_index = {doc: idx for idx, doc in enumerate(doctor_names)}

# 🔧 FIX: 禁止医師名も正規化
WED_FORBIDDEN_DOCTORS = {normalize_name(d) for d in WED_FORBIDDEN_DOCTORS}  # 🔧 FIX

# v6.5.0: Sheet1:Z列「カテ当番」を病院列から除外
# カテ当番列の検出（「カテ当番」という名前の列を探す）
KATE_TOBAN_COL = None
all_cols = list(shift_df.columns[1:])
for col in all_cols:
    if str(col).strip() == "カテ当番":
        KATE_TOBAN_COL = col
        break

if KATE_TOBAN_COL is not None:
    hospital_cols = [c for c in all_cols if c != KATE_TOBAN_COL]
else:
    hospital_cols = all_cols
    print("⚠️ Sheet1に「カテ当番」列が見つかりません（従来方式を使用）")

n_cols = len(shift_df.columns)

# 列インデックス（テンプレ依存：B〜Y を想定）
B_COL_INDEX = 1
C_COL_INDEX = 2
D_COL_INDEX = min(3, n_cols - 1)
E_COL_INDEX = min(4, n_cols - 1)
F_COL_INDEX = min(5, n_cols - 1)
G_COL_INDEX = min(6, n_cols - 1)
H_COL_INDEX = min(7, n_cols - 1)
I_COL_INDEX = min(8, n_cols - 1)
J_COL_INDEX = min(9, n_cols - 1)
K_COL_INDEX = min(10, n_cols - 1)
L_COL_INDEX = min(11, n_cols - 1)
M_COL_INDEX = min(12, n_cols - 1)
Q_COL_INDEX = min(16, n_cols - 1)
U_COL_INDEX = min(20, n_cols - 1)
Y_COL_INDEX = min(24, n_cols - 1)

# 列範囲定義
B_H_START_INDEX = B_COL_INDEX  # 大学系前半（2回まで）
B_H_END_INDEX = H_COL_INDEX
I_K_START_INDEX = I_COL_INDEX  # 大学系後半
I_K_END_INDEX = K_COL_INDEX
B_K_START_INDEX = B_COL_INDEX  # 大学系全体
B_K_END_INDEX = K_COL_INDEX
L_Y_START_INDEX = L_COL_INDEX  # 外病院
L_Y_END_INDEX = min(Y_COL_INDEX, n_cols - 1)

print(f"\n✅ Excel読込完了: 医師{len(doctor_names)}人 | 病院{len(hospital_cols)}列 | {len(shift_df)}日間")

# =========================
# sheet2 可否コード
# =========================
# v6.5.9: 空欄は常に「1（可）」として扱う。
# 旧実装は列の最初のマークを全空欄日のフォールバックにしていたため、
# 月初の 0/2/3 マークが全月へ伝播し、休み希望・列制限が意図せず拡大していた
# （例: 1日だけ0を書いた医師が全月不可=inactive扱いになる）。
# 解釈できないマーク（非数値・未定義コード）は読込時に警告する。
_invalid_avail_marks = []
if isinstance(availability_df.index, pd.DatetimeIndex):
    _missing_avail_cols = [d for d in doctor_names if d not in availability_df.columns]
    if _missing_avail_cols:
        print(f"⚠️ WARNING: sheet2に列が見つからない医師がいます（全日「可(1)」として扱われます）: {_missing_avail_cols}")
    for doc in doctor_names:
        if doc not in availability_df.columns:
            continue
        for _dt, _v in availability_df[doc].items():
            # NaT行（日付として解釈できない凡例・注記行）は対象外
            if pd.isna(_v) or pd.isna(_dt):
                continue
            try:
                _f = float(_v)
                _ok = abs(_f - 1.2) < 0.01 or _f in (0.0, 1.0, 2.0, 3.0)
            except Exception:
                _ok = False
            if not _ok:
                _invalid_avail_marks.append((_dt, doc, _v))
if _invalid_avail_marks:
    print(f"⚠️ WARNING: sheet2に解釈できないマークが{len(_invalid_avail_marks)}件あります → 全て「可(1)」として扱われます:")
    for _dt, _doc, _v in _invalid_avail_marks[:10]:
        print(f"   - {pd.to_datetime(_dt).strftime('%Y-%m-%d')} {_doc}: {_v!r}")
    if len(_invalid_avail_marks) > 10:
        print(f"   ... 他{len(_invalid_avail_marks) - 10}件")

def get_avail_code(date, doctor):
    """可否コードを取得
    v6.5.0: 出張曜日の前日は自動的に0（不可）として扱う
    """
    date_norm = pd.to_datetime(date).normalize().tz_localize(None)

    # v6.5.0: 出張曜日の前日は0（不可）として扱う
    try:
        if 'doctor_travel_day' in globals() and 'WEEKDAY_MAP' in globals():
            travel_str = doctor_travel_day.get(doctor, "")
            if travel_str and travel_str in WEEKDAY_MAP:
                travel_wd = WEEKDAY_MAP[travel_str]
                pre_travel_wd = (travel_wd - 1) % 7
                if date_norm.weekday() == pre_travel_wd:
                    return 0  # 出張前日は不可
    except Exception:
        pass

    code = None
    raw_value = None
    if isinstance(availability_df.index, pd.DatetimeIndex):
        try:
            value = availability_df.at[date_norm, doctor]
            if isinstance(value, pd.Series):
                value = value.iloc[0]
            if pd.notna(value):
                raw_value = float(value)
                # 1.2は特別扱い：大学系優先
                if abs(raw_value - 1.2) < 0.01:
                    code = 1.2
                elif raw_value in (0.0, 1.0, 2.0, 3.0):
                    code = int(raw_value)
                # v6.5.9: 上記以外（0.5等）は解釈不能 → None（可(1)扱い、読込時に警告済み）
                # 旧実装のint切り捨ては 0.5→不可(0) 等、警告文と矛盾する挙動だった
        except Exception:
            pass
    if code is None:
        code = 1  # v6.5.9: 空欄・解釈不能は「可」（旧: 月初マークへのフォールバックは伝播バグのため廃止）
    if code not in (0, 1, 1.2, 2, 3):
        code = 1
    return code

def get_sched_code(date, doctor):
    """その日の有効なカテ表コードを取得
    v6.5.0: Sheet3に直接記載があればそれを優先、なければSheet1:Z + Sheet4:属性で判定
    """
    date_norm = pd.to_datetime(date).normalize().tz_localize(None)

    # 1. Sheet3に直接記載があればそれを優先（特別シフト: CC1, CC2等）
    if doctor in schedule_df.columns:
        try:
            value = schedule_df.at[date_norm, doctor]
            if isinstance(value, pd.Series):
                value = value.iloc[0]
            if pd.notna(value):
                code_str = str(value).strip()
                # 0と3は有効なカテ表コードではない（0=データなし、3=可否コード）
                # ただし「1」は平日緩和フラグとして有効
                if code_str and code_str != "0" and code_str != "3":
                    return code_str
        except Exception:
            pass

    # 2. Sheet1:Z + Sheet4:属性 で判定（v6.5.0 新機能）
    # kate_team_by_date と doctor_kate_team はグローバル変数として後で設定される
    try:
        if 'kate_team_by_date' in globals() and 'doctor_kate_team' in globals():
            team_on_duty = kate_team_by_date.get(date_norm)
            doc_team = doctor_kate_team.get(doctor, "")
            if team_on_duty and doc_team:
                # チームが一致すればカテ当番
                if team_on_duty == doc_team:
                    return team_on_duty
    except Exception:
        pass

    return None

# sheet2 と sheet3 の医師列がズレていないか（ズレてても動くが、制約が弱くなる）
# v6.5.0: schedule_rawが空の場合（新Excel構造）はスキップ
if len(schedule_raw.columns) > 1:
    sched_doctors = [normalize_name(x) for x in list(schedule_raw.columns[1:])]
    if doctor_names != sched_doctors:
        print("⚠️ WARNING: sheet2(可否) と sheet3(カテ表) の医師列が一致していません。")
        only2 = [d for d in doctor_names if d not in sched_doctors]
        only3 = [d for d in sched_doctors if d not in doctor_names]
        if only2:
            print(f"   sheet2 only (先頭10): {only2[:10]}")
        if only3:
            print(f"   sheet3 only (先頭10): {only3[:10]}")
        print("   ※H〜U の『カテ表あり不可』制約が一部の医師で効かない可能性があります。")
else:
    sched_doctors = []
    # Sheet3の医師列なし（EXTRA順序はSheet2にフォールバック）

# =========================
# sheet4 前月まで累積
# =========================
name_to_row = {row["氏名"]: row for _, row in sheet4_data.iterrows()}
prev_names = list(sheet4_data["氏名"])

def match_prev_name(doc):
    if doc in name_to_row:
        return doc
    ms = [p for p in prev_names if str(p).startswith(doc) or doc.startswith(str(p))]
    return ms[0] if len(ms) == 1 else None

name_match = {doc: match_prev_name(doc) for doc in doctor_names}
unmatched = [d for d in doctor_names if name_match.get(d) is None]
if unmatched:
    print(f"⚠️ WARNING: sheet4(累積)で名前が一致しない医師がいます（累積が0扱いになります）: {unmatched}")

def prev_get(doc, colname):
    pname = name_match.get(doc)
    if pname and pname in name_to_row:
        row = name_to_row[pname]
        v = row.get(colname, 0)
        try:
            return float(v or 0)
        except Exception:
            return 0.0
    return 0.0

prev_total   = {d: prev_get(d, "全合計")   for d in doctor_names}
prev_bg      = {d: prev_get(d, "大学合計") for d in doctor_names}
prev_ht      = {d: prev_get(d, "外病院合計") for d in doctor_names}
prev_weekday = {d: prev_get(d, "平日")     for d in doctor_names}
prev_weekend = {d: prev_get(d, "休日合計") for d in doctor_names}

# =========================
# v6.5.0: Sheet4から属性・出張曜日を取得
# =========================
def prev_get_str(doc, colname):
    """文字列列の値を取得"""
    pname = name_match.get(doc)
    if pname and pname in name_to_row:
        row = name_to_row[pname]
        v = row.get(colname, "")
        return str(v).strip() if pd.notna(v) else ""
    return ""

# v6.5.1: カテチーム属性の取得（Sheet1:Zと一致するチームコードを持つ列を探す）
# まずSheet1:Zから期待されるチームコードを取得（後でkate_team_by_dateが設定されるので先に取得）
expected_team_codes = set()
if KATE_TOBAN_COL is not None:
    for ridx in shift_df.index:
        team_val = shift_df.at[ridx, KATE_TOBAN_COL]
        if pd.notna(team_val):
            team_str = str(team_val).strip()
            if team_str and team_str.lower() != "nan":
                expected_team_codes.add(team_str)

# 列候補をチェックして、Sheet1:Zのチームコードと一致する値を持つ列を探す
kate_team_col_name = None
for col_candidate in ["カテ当番", "属性", "カテ", "チーム"]:  # カテ当番を優先
    if col_candidate not in sheet4_data.columns:
        continue
    # この列の値がSheet1:Zのチームコードと一致するかチェック
    col_values = set()
    for idx, row in sheet4_data.iterrows():
        v = row.get(col_candidate, "")
        if pd.notna(v):
            val_str = str(v).strip()
            if val_str and val_str.lower() not in ("nan", "none", ""):
                col_values.add(val_str)
    # 期待されるチームコードと1つでも一致すればこの列を使用
    if expected_team_codes and col_values & expected_team_codes:
        kate_team_col_name = col_candidate
        # Sheet4:kate_team_col_name列使用
        break
    elif not expected_team_codes and col_values:
        # Sheet1:Zがない場合は最初に見つかった列を使用
        kate_team_col_name = col_candidate
        # Sheet4:kate_team_col_name列使用（フォールバック）
        break

if kate_team_col_name:
    doctor_kate_team = {d: prev_get_str(d, kate_team_col_name) for d in doctor_names}
else:
    doctor_kate_team = {d: "" for d in doctor_names}
    print("⚠️ Sheet4にカテチーム列（カテ当番/属性）が見つからないか、Sheet1:Zのチームコードと一致しません")

# v6.5.6: 属性列（1, 2など）を別途取得（SEMI-001緩和可否の判定に使用）
# 属性=1: 緩和可、属性=2: 緩和不可
doctor_attribute = {}
if "属性" in sheet4_data.columns:
    for doc in doctor_names:
        attr_val = prev_get_str(doc, "属性")
        doctor_attribute[doc] = attr_val
    attr0 = sum(1 for v in doctor_attribute.values() if v == "0")
    attr1 = sum(1 for v in doctor_attribute.values() if v == "1")
    attr2 = sum(1 for v in doctor_attribute.values() if v == "2")
    attr3 = sum(1 for v in doctor_attribute.values() if v == "3")
    attr_none = len(doctor_attribute) - attr0 - attr1 - attr2 - attr3
    parts = []
    if attr0: parts.append(f"0:{attr0}")
    if attr1: parts.append(f"1:{attr1}")
    if attr2: parts.append(f"2:{attr2}")
    if attr3: parts.append(f"3:{attr3}")
    if attr_none: parts.append(f"未設定:{attr_none}")
    print(f"   属性: {len(doctor_attribute)}人（{', '.join(parts)}）")
else:
    doctor_attribute = {d: "" for d in doctor_names}
    print("⚠️ Sheet4に属性列が見つかりません")

# 出張曜日の取得（「出張日」列から）
travel_col_name = None
for col_candidate in ["出張日", "出張曜日"]:
    if col_candidate in sheet4_data.columns:
        travel_col_name = col_candidate
        break

if travel_col_name:
    doctor_travel_day = {d: prev_get_str(d, travel_col_name) for d in doctor_names}
else:
    doctor_travel_day = {d: "" for d in doctor_names}

# デバッグ: Sheet4の列名を表示
# Sheet4列名（デバッグ用）: print(f"📋 Sheet4列名: {list(sheet4_data.columns)}")

# 曜日名から曜日番号へのマッピング（月曜=0, ..., 日曜=6）
WEEKDAY_MAP = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}

def get_travel_weekday(doc):
    """医師の出張曜日を数値で返す（なければNone）"""
    travel_str = doctor_travel_day.get(doc, "")
    if travel_str and travel_str in WEEKDAY_MAP:
        return WEEKDAY_MAP[travel_str]
    return None

# 出張曜日の前日に当たる日付を計算
def get_pre_travel_dates(doc, all_dates):
    """出張曜日の前日の日付リストを返す"""
    travel_wd = get_travel_weekday(doc)
    if travel_wd is None:
        return set()
    # 前日の曜日（0-6）
    pre_travel_wd = (travel_wd - 1) % 7
    return {d for d in all_dates if d.weekday() == pre_travel_wd}

# 属性情報の表示
doc_with_attr = [(d, doctor_kate_team[d]) for d in doctor_names if doctor_kate_team[d]]
doc_with_travel = [(d, doctor_travel_day[d]) for d in doctor_names if doctor_travel_day[d]]
if not doc_with_attr:
    print("⚠️ カテチーム属性を持つ医師が0人です")

# =========================
# v6.5.0: Sheet1からカテ当番日（チーム）を取得
# =========================
kate_team_by_date = {}  # 日付 -> チームコード (A, B, C, D, E等)
if KATE_TOBAN_COL is not None:
    for ridx in shift_df.index:
        date = shift_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        team_val = shift_df.at[ridx, KATE_TOBAN_COL]
        if pd.notna(team_val):
            team_str = str(team_val).strip()
            if team_str and team_str.lower() != "nan":
                kate_team_by_date[date] = team_str
    if kate_team_by_date:
        unique_teams = set(kate_team_by_date.values())
    else:
        print("⚠️ Sheet1:Z列にカテ当番チームのデータがありません")
else:
    print("⚠️ Sheet1に「カテ当番」列がないため、Sheet3/Sheet4のカテ表を使用します")

# =========================
# 全枠数カウント + slots_by_date 前計算
# =========================
slots_by_date = defaultdict(lambda: {"preassigned": [], "free": []})
preassigned_count = {d: 0 for d in doctor_names}
total_slots = 0

for ridx in shift_df.index:
    date = shift_df.at[ridx, date_col_shift]
    if pd.isna(date):
        continue
    date = pd.to_datetime(date).normalize().tz_localize(None)  # 🔧 FIX

    for hosp in hospital_cols:
        val = shift_df.at[ridx, hosp]

        # 固定割当（セルに医師名が入っている）
        val_str = normalize_name(val) if isinstance(val, str) else ""  # 🔧 FIX
        if val_str in doctor_names:
            doc = val_str
            slots_by_date[date]["preassigned"].append((ridx, hosp, doc))
            preassigned_count[doc] += 1
            total_slots += 1
            continue

        # 自動枠（1/〇など）
        if is_slot_value(val):
            slots_by_date[date]["free"].append((ridx, hosp))
            total_slots += 1

if len(doctor_names) == 0:
    raise ValueError("❌ sheet2 に医師名がありません")

all_dates = sorted(slots_by_date.keys())
all_shift_dates = sorted(pd.to_datetime(shift_df[date_col_shift].dropna()).dt.normalize().dt.tz_localize(None).unique())  # 🔧 FIX

# =========================
# cap設計：n回ベース＋余りは右側（下側）からn+1回
# inactive医師の扱い（v5.2仕様書§5準拠）:
#   - 解析処理から完全除外（候補に含めない）
#   - TARGET_CAP計算から除外（active医師のみで計算）
#   - 出力Excelには0回として記載（氏名は表示）
# =========================
def is_always_unavailable(doc):
    """inactive医師の判定: sheet2で全日=0かつ事前割当なし"""
    if preassigned_count.get(doc, 0) > 0:
        return False
    return all(get_avail_code(d, doc) == 0 for d in all_shift_dates)

inactive_doctors = [d for d in doctor_names if is_always_unavailable(d)]
active_doctors = [d for d in doctor_names if d not in inactive_doctors]
if len(active_doctors) == 0:
    raise ValueError("❌ 当月に割り当て可能な医師がいません")
if inactive_doctors:
    print(f"⚠️ inactive医師（解析除外、出力のみ）: {len(inactive_doctors)}人")

# 可否コード2の医師（大学系のみ可能）
# v6.0.5: CODE_2医師もEXTRA枠対象に含める（除外すると物理的に枠不足になりうる）
def has_code_2_anywhere(doc):
    """医師がsheet2でいずれかの日に可否コード2を持っているか"""
    if doc not in availability_df.columns:
        return False
    for date in all_shift_dates:
        code = get_avail_code(date, doc)
        if code == 2:
            return True
    return False

CODE_2_DOCTORS = {doc for doc in doctor_names if has_code_2_anywhere(doc)}

BASE_TARGET = total_slots // len(active_doctors)
EXTRA_SLOTS = total_slots - BASE_TARGET * len(active_doctors)

# 余り枠(EXTRA)は属性1の医師から優先的に選出する
# v6.5.9: 同一優先度内では前月累積（全合計）が最少の医師から選出
#   - 旧実装のSheet2末尾順は位置ベースで、累積最多の医師に+1が付き
#     月をまたいだ公平性が逆行するケースがあった
# v6.0.5: CODE_2医師もEXTRA対象に含める
#   - CODE_2除外だと、他医師の制約(gap/dup等)で枠が埋まらず未割当が発生する
#   - CODE_2医師のn+1回目はB〜Q列（大学系）に割り当てればよい
active_sorted_by_index = sorted(active_doctors, key=lambda d: doctor_col_index[d])

def _extra_priority(d):
    """EXTRA枠の選出順: 前月累積（全合計）が少ない順 → Sheet2列順"""
    return (prev_total.get(d, 0), doctor_col_index[d])

# 属性1の医師をEXTRA候補として優先選出（前月累積最少順）
attr1_doctors = [d for d in active_sorted_by_index if doctor_attribute.get(d, "") == "1"]
attr1_by_prev = sorted(attr1_doctors, key=_extra_priority)
non_attr1_by_prev = sorted(
    [d for d in active_sorted_by_index if d not in attr1_doctors], key=_extra_priority
)
if EXTRA_SLOTS > 0 and len(attr1_by_prev) >= EXTRA_SLOTS:
    # 属性1の医師で十分 → 前月累積が少ない順にEXTRA_SLOTS人を選出
    EXTRA_ALLOWED = set(attr1_by_prev[:EXTRA_SLOTS])
elif EXTRA_SLOTS > 0 and attr1_by_prev:
    # 属性1だけでは不足 → 属性1全員 + 残りを前月累積最少の非属性1から補充
    remaining = EXTRA_SLOTS - len(attr1_by_prev)
    EXTRA_ALLOWED = set(attr1_by_prev) | set(non_attr1_by_prev[:remaining])
else:
    # 属性1がいない場合は前月累積最少からフォールバック
    _all_by_prev = sorted(active_sorted_by_index, key=_extra_priority)
    EXTRA_ALLOWED = set(_all_by_prev[:EXTRA_SLOTS] if EXTRA_SLOTS > 0 else [])

TARGET_CAP = {d: 0 for d in doctor_names}
for d in active_doctors:
    TARGET_CAP[d] = BASE_TARGET
for d in EXTRA_ALLOWED:
    TARGET_CAP[d] = BASE_TARGET + 1
for d in doctor_names:
    if preassigned_count.get(d, 0) > TARGET_CAP.get(d, 0):
        TARGET_CAP[d] = preassigned_count[d]

# v6.1.0: gap>=3を満たせる最大割当数でTARGET_CAPを制限
# 利用可能日の分布が偏っている医師は、CAPまで割当するとgap違反が不可避になる
# → 事前に物理的上限を計算してCAPを切り下げる
def compute_max_gap3_assignments(doc):
    """gap>=3を満たしつつ割当可能な最大回数を貪欲法で計算"""
    avail_dates = sorted([d for d in all_shift_dates if get_avail_code(d, doc) != 0])
    if not avail_dates:
        return 0
    count = 0
    last_assigned = None
    for d in avail_dates:
        if last_assigned is None or (d - last_assigned).days >= 3:
            count += 1
            last_assigned = d
    return count

gap3_cap_adjusted = 0
for d in active_doctors:
    max_gap3 = compute_max_gap3_assignments(d)
    if max_gap3 < TARGET_CAP[d]:
        gap3_cap_adjusted += 1
        TARGET_CAP[d] = max_gap3

if gap3_cap_adjusted > 0:
    print(f"   ⚠️ gap>=3制約により{gap3_cap_adjusted}人のTARGET_CAPを切り下げ")

# CAPを切り下げた分、余った枠を他の医師に再配分
total_cap = sum(TARGET_CAP[d] for d in active_doctors)
shortage = total_slots - total_cap
if shortage > 0:
    # 属性1の医師に優先的に再配分、同一優先度内は前月累積が少ない順（v6.5.9）
    _redist_candidates = attr1_by_prev + non_attr1_by_prev
    for d in _redist_candidates:
        if shortage <= 0:
            break
        max_gap3 = compute_max_gap3_assignments(d)
        if TARGET_CAP[d] < max_gap3:
            TARGET_CAP[d] += 1
            shortage -= 1
    if shortage > 0:
        print(f"   ⚠️ {shortage}枠の再配分先なし（全医師がgap3上限）")

floor_shifts = BASE_TARGET

total_cap_final = sum(TARGET_CAP[d] for d in active_doctors)
extra_names = [d for d in active_sorted_by_index if d in EXTRA_ALLOWED]
extra_attr1_count = sum(1 for d in EXTRA_ALLOWED if doctor_attribute.get(d, "") == "1")
gap3_limited = [(d, compute_max_gap3_assignments(d)) for d in active_doctors if compute_max_gap3_assignments(d) < BASE_TARGET]

print(f"\n✅ 割当: {len(active_doctors)}人 × {BASE_TARGET}回 + {len(EXTRA_ALLOWED)}人×1回 = {total_cap_final}/{total_slots}枠")
if extra_names:
    print(f"   +1回: {', '.join(extra_names)}")
if gap3_limited:
    print(f"   gap3制限: {', '.join(d for d, _ in gap3_limited)}")

# =========================
# B-K / L-Y 比率バランス（sheet3で「3」記載の医師は除外）
# sheet3でカテ表コード保有医師の特定
# =========================
def has_sheet3_code_3(doc):
    """医師がsheet3でコード「3」を持つか、またはSheet2で全日コード3（外病院専門）か"""
    # 旧構造: Sheet3にコード「3」がある
    if doc in schedule_df.columns:
        values = schedule_df[doc].dropna()
        if any(str(v).strip() == "3" for v in values):
            return True
    # v6.5.8: 新構造ではSheet2の可否コードで判定（全日がコード3なら外病院専門）
    if doc in availability_df.columns:
        avail_vals = availability_df[doc].dropna()
        if len(avail_vals) > 0 and all(str(v).strip() == "3" for v in avail_vals):
            return True
    return False

def has_any_schedule_code(doc):
    """医師がカテ当番を持っているか
    v6.5.5: Sheet1:Zのチームコード（A,B,C,D,E等）に一致する場合のみTrue
    """
    # v6.5.5: Sheet4のカテチーム属性がSheet1:Zのチームコードと一致すればTrue
    if 'doctor_kate_team' in globals() and doctor_kate_team:
        team = doctor_kate_team.get(doc, "")
        if team:
            # expected_team_codesが定義されていて、チームがその中にあればTrue
            if 'expected_team_codes' in globals() and expected_team_codes:
                if team in expected_team_codes:
                    return True
            # expected_team_codesがない場合は従来方式（A-E等の1文字アルファベット）
            elif team.upper() in ('A', 'B', 'C', 'D', 'E', 'F', 'G'):
                return True

    # 従来方式: Sheet3のカテ表コードをチェック
    if len(schedule_df.columns) > 0 and doc in schedule_df.columns:
        values = schedule_df[doc].dropna()
        for v in values:
            s = str(v).strip()
            if s and s != "0" and s != "3":  # 0と3以外のコードがあればTrue
                return True
    return False

RATIO_EXEMPT_DOCTORS = {doc for doc in doctor_names if has_sheet3_code_3(doc)}
if RATIO_EXEMPT_DOCTORS:
    print(f"   比率バランス除外（sheet3に3あり）: {sorted(RATIO_EXEMPT_DOCTORS)}")

SCHEDULE_CODE_HOLDERS = {doc for doc in doctor_names if has_any_schedule_code(doc)}
NO_KATE_DOCTORS = {doc for doc in doctor_names if not has_any_schedule_code(doc)}
print(f"   カテ保有: {len(SCHEDULE_CODE_HOLDERS)}人 | なし: {len(NO_KATE_DOCTORS)}人")

if len(SCHEDULE_CODE_HOLDERS) == 0:
    sample_teams = [(d, doctor_kate_team.get(d, "")) for d in list(doctor_names)[:5]]
    print(f"   ⚠️ カテ表保有者0人 - {sample_teams}")

# sheet3で「1」を持つ医師（平日大学系でカテ当番不一致を許容）
# v6.5.8: 新構造ではschedule_dfが空のため、属性1をフォールバックとして使用
def has_sheet3_code_1(doc):
    """医師がsheet3で少なくとも1つの「1」コードを持っているか（新構造では属性1で代替）"""
    # 旧構造: Sheet3にコード「1」がある
    if doc in schedule_df.columns:
        values = schedule_df[doc].dropna()
        if any(str(v).strip() == "1" for v in values):
            return True
    # 新構造: 属性1の医師を「平日緩和」対象として扱う
    if doctor_attribute.get(doc, "") == "1":
        return True
    return False

SHEET3_CODE_1_DOCTORS = {doc for doc in doctor_names if has_sheet3_code_1(doc)}
if SHEET3_CODE_1_DOCTORS:
    print(f"   平日緩和: {len(SHEET3_CODE_1_DOCTORS)}人")

def is_ch_slot(col_idx):
    """C-H列（休日大学系、インデックス2-7）かどうか"""
    return C_COL_INDEX <= col_idx <= H_COL_INDEX

def is_weekday_university_slot(col_idx):
    """B列またはI-K列（平日大学系）かどうか"""
    return col_idx == B_COL_INDEX or (I_COL_INDEX <= col_idx <= K_COL_INDEX)

def is_eligible_for_ch_slot(doc, date):
    """C-H列（休日大学系）に割り当て可能かどうか
    条件：その日にカテ当番あり OR カテ当番が一回もない医師
    """
    # カテ当番が一回もない医師はOK
    if doc in NO_KATE_DOCTORS:
        return True
    # カテ当番保有医師は、その日にカテ表コードがあればOK
    if get_sched_code(date, doc):
        return True
    return False

def is_eligible_for_weekday_university_slot(doc, date):
    """B列/I-K列（平日大学系）に割り当て可能かどうか
    条件：カテ当番なし医師 OR その日にカテ当番あり OR sheet3で「1」を持つ医師
    「1」を持つ医師はカテ当番が合わなくても許容
    """
    # カテ当番が一回もない医師はOK
    if doc in NO_KATE_DOCTORS:
        return True
    # その日にカテ表コードがあればOK
    sched_code = get_sched_code(date, doc)
    if sched_code:
        return True
    # sheet3で「1」を持つ医師はカテ当番なしでも許容（平日大学系のみ）
    if doc in SHEET3_CODE_1_DOCTORS:
        return True
    return False

def is_cc_assignment(date, doc):
    """その日のその医師の割り当てがCC（大型連休特別シフト）かどうか"""
    sched_code = get_sched_code(date, doc)
    return sched_code == "CC" if sched_code else False

def has_any_cc_assignment(doc, pattern_df):
    """医師がCC割り当てを持っているかどうかをpattern_dfから判定"""
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        for hosp in hospital_cols:
            val = pattern_df.at[ridx, hosp]
            if isinstance(val, str) and normalize_name(val) == doc:
                if is_cc_assignment(date, doc):
                    return True
    return False

# 可否コード1.2の医師（大学系最低1回必須）
def has_code_1_2(doc):
    """医師がsheet2で少なくとも1つの1.2コードを持っているか"""
    if doc not in availability_df.columns:
        return False
    for date in all_shift_dates:
        code = get_avail_code(date, doc)
        if code == 1.2:
            return True
    return False

CODE_1_2_DOCTORS = {doc for doc in doctor_names if has_code_1_2(doc)}

# 大学系最低1回必須の医師（準ハード制約：コード3以外の全医師）
# コード3は外病院専門なので除外
UNIVERSITY_MINIMUM_REQUIRED_DOCTORS = {doc for doc in active_doctors if doc not in RATIO_EXEMPT_DOCTORS}

# =========================
# 大学(B〜G)の昼夜判定 & 7分類
# =========================
def is_bg_day_shift(hosp_name, col_idx):
    if hosp_name in BG_DAY_COLS:
        return True
    if hosp_name in BG_NIGHT_COLS:
        return False
    # デフォルト：B,C,E,F=昼 / D,G=夜
    if col_idx in (B_COL_INDEX, C_COL_INDEX, E_COL_INDEX, F_COL_INDEX):
        return True
    if col_idx in (D_COL_INDEX, G_COL_INDEX):
        return False
    mid = (B_COL_INDEX + G_COL_INDEX) // 2
    return col_idx <= mid

def is_bk_slot(col_idx):
    return B_K_START_INDEX <= col_idx <= B_K_END_INDEX

def is_ly_slot(col_idx):
    return L_Y_START_INDEX <= col_idx <= L_Y_END_INDEX

def classify_bg_category(date, hosp_name):
    idx = shift_df.columns.get_loc(hosp_name)
    is_day = is_bg_day_shift(hosp_name, idx)
    dow = pd.to_datetime(date).weekday()
    holi = is_holiday(date)

    weekday = dow < 5
    # 平日かつ C,D,F,G は祝日扱い
    if weekday and idx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX):
        holi = True

    if holi and weekday:
        base = "祝日"
    elif dow == 5:
        base = "土曜"
    elif dow == 6:
        base = "日曜"
    else:
        return "平日"

    return base + ("昼" if is_day else "夜")

# =========================
# 1枠の医師選択（greedy）
# =========================
def choose_doctor_for_slot(
    date,
    hospital_name,
    assigned_count,
    assigned_dates,
    assigned_bg,
    assigned_ht,
    assigned_weekday,
    assigned_weekend,
    assigned_be,
    assigned_fg,
    assigned_bk,
    assigned_ly,
    assigned_bh,
    assigned_bi,      # v6.0.0: B/I列の合計（HARD-001）
    assigned_chjk,    # v6.0.0: C-H/J-K列の合計（HARD-002）
    assigned_hosp_count,
    assigned_bg_dates,  # v6.5.0: 大学系割当日（ABS-012改: 7日間隔ルール）
    semi001_violation_weeks,  # v6.5.4: SEMI-001違反週（月曜始まり）
):
    idx = shift_df.columns.get_loc(hospital_name)
    is_BE = B_COL_INDEX <= idx <= E_COL_INDEX
    is_BG = B_COL_INDEX <= idx <= K_COL_INDEX
    is_BH = B_H_START_INDEX <= idx <= B_H_END_INDEX
    is_LY_range = L_COL_INDEX <= idx <= L_Y_END_INDEX
    is_BK = is_bk_slot(idx)
    is_LY = is_ly_slot(idx)
    # v6.0.0: 新しい列グループ
    is_B_or_I = (idx == B_COL_INDEX or idx == I_COL_INDEX)  # グループA
    is_CH_or_JK = ((C_COL_INDEX <= idx <= H_COL_INDEX) or (J_COL_INDEX <= idx <= K_COL_INDEX))  # グループB
    is_B_only = (idx == B_COL_INDEX)  # SEMI-001対象
    is_CH_only = (C_COL_INDEX <= idx <= H_COL_INDEX)  # SEMI-002対象
    dow = pd.to_datetime(date).weekday()
    weekday = dow < 5

    def collect_candidates(
        relax_semi=False,  # v6.0.0: SEMI制約を緩和（sheet3「1」以外も許容）
        relax_hard=False,  # v6.0.1: HARD制約を緩和（ABS-009回避のため）
        relax_abs=False,   # v6.3.0: ABS制約を緩和（未割当回避のため）
    ):
        candidates = []
        for doc in doctor_names:
            # === 絶対禁忌（ABS）===
            # v6.3.0: relax_abs=Trueでも緩和しない制約（物理的に不可能）

            # ABS-006: 同日重複禁止（絶対不可）
            if date in assigned_dates[doc]:
                continue

            code = get_avail_code(date, doc)

            # ABS-001: コード0は全列禁止（絶対不可）
            if code == 0:
                continue

            # ABS-002: コード2はB〜Q列のみ（R-Y列禁止）
            if code == 2 and not (B_COL_INDEX <= idx <= Q_COL_INDEX):
                continue

            # ABS-003: コード3はL〜Y列のみ（大学系禁止）
            if code == 3 and not (L_COL_INDEX <= idx <= L_Y_END_INDEX):
                continue

            # ABS-004: カテ表コードありの日はL〜Y列不可
            if L_COL_INDEX <= idx <= L_Y_END_INDEX:
                if get_sched_code(date, doc):
                    continue

            # ABS-005: 水曜日L〜Y列禁止医師
            if dow == 2 and is_LY_range:
                if doc in WED_FORBIDDEN_DOCTORS:
                    continue

            # ABS-013: C-H列（休日大学系）カテ当番必須（v6.5.3）
            # カテ当番保有医師がC-H列に入るには、その日にカテ当番が必要
            if is_CH_only:
                if not is_eligible_for_ch_slot(doc, date):
                    continue

            # === v6.3.0: 以下のABS制約はrelax_abs=Trueで緩和可能 ===

            # ABS-007: gap >= 3日必須
            if not relax_abs and assigned_dates[doc]:
                min_gap = min(abs((pd.to_datetime(date) - x).days) for x in assigned_dates[doc])
                if min_gap < 3:
                    continue

            # ABS-008: 同一病院重複禁止（初期生成時は全列で禁止、fix関数では外病院のみ）
            if not relax_abs and assigned_hosp_count[doc].get(hospital_name, 0) >= 1:
                continue

            # ABS-010: TARGET_CAP遵守（n超過禁止）
            if not relax_abs and assigned_count[doc] >= TARGET_CAP.get(doc, 0):
                continue

            # ABS-011: 大学系2回まで（B-K列合計）
            if not relax_abs and is_BG and assigned_bg[doc] >= 2:
                continue

            # ABS-012改: 大学系は7日間隔必須（v6.5.0）
            if not relax_abs and is_BG:
                if assigned_bg_dates[doc]:
                    min_bg_gap = min(abs((pd.to_datetime(date) - x).days) for x in assigned_bg_dates[doc])
                    if min_bg_gap < 7:
                        continue

            # === ハード制約（HARD）: カテなし医師は必須遵守 ===

            # カテ当番の有無を判定
            is_kate_holder = doc in SCHEDULE_CODE_HOLDERS
            is_sheet3_one = doc in SHEET3_CODE_1_DOCTORS
            # v6.5.6: 属性による緩和可否（属性=1なら緩和可、属性=2なら緩和不可）
            doc_attr = doctor_attribute.get(doc, "")
            is_attr_1 = (doc_attr == "1")  # 属性1は緩和可
            is_attr_2 = (doc_attr == "2")  # 属性2は緩和不可

            # HARD-001: B/I列1回まで（グループA）
            if not relax_hard and is_B_or_I and assigned_bi[doc] >= 1:
                # カテなし医師は必須遵守
                if not is_kate_holder:
                    continue
                # カテあり医師でもsheet3「1」以外は遵守
                if is_kate_holder and not is_sheet3_one:
                    continue

            # HARD-002: C-H/J-K列1回まで（グループB）
            if not relax_hard and is_CH_or_JK and assigned_chjk[doc] >= 1:
                # カテなし医師は必須遵守
                if not is_kate_holder:
                    continue
                # カテあり医師でもsheet3「1」以外は遵守
                if is_kate_holder and not is_sheet3_one:
                    continue

            # === 準ハード制約（SEMI）: 属性1は緩和対象、属性2は緩和不可 ===

            # SEMI-001: B列のみカテ表コード必須（v6.5.6: 属性で緩和可否を判定）
            # カテ持ち + 属性1: 緩和可（週1回まで許容）
            # カテ持ち + 属性2: 緩和不可（ABS該当、カテ表必須）
            # カテなし: その他のルールに従う
            if is_B_only and is_kate_holder and not get_sched_code(date, doc):
                if is_attr_2:
                    # 属性2は緩和不可（ABS該当）- カテ表なしでB列配置不可
                    continue
                elif is_attr_1:
                    # 属性1は緩和可（週1回まで許容）
                    if not relax_semi:
                        week_start = get_monday_week_start(date)
                        if week_start in semi001_violation_weeks[doc]:
                            continue
                        # 1回目は許容（選ばれた場合、後で週を記録）
                else:
                    # 属性未設定の場合はsheet3「1」をフォールバック
                    if not relax_semi and not is_sheet3_one:
                        week_start = get_monday_week_start(date)
                        if week_start in semi001_violation_weeks[doc]:
                            continue

            # (SEMI-002はABS-013に格上げ済み - v6.5.3)

            candidates.append(doc)
        return candidates

    # v6.3.0: 段階的制約緩和（未割当回避を最優先）
    # 1. 全制約適用
    candidates = collect_candidates()
    # 2. SEMI緩和
    if not candidates:
        candidates = collect_candidates(relax_semi=True)
    # 3. HARD緩和（SEMI緩和済み）
    if not candidates:
        candidates = collect_candidates(relax_semi=True, relax_hard=True)
    # 4. ABS緩和（SEMI/HARD緩和済み）- 未割当を絶対に回避
    if not candidates:
        candidates = collect_candidates(relax_semi=True, relax_hard=True, relax_abs=True)

    if not candidates:
        # 全制約緩和後も候補なし（コード0または同日重複のみ）
        return None

    any_under_floor = any(assigned_count[d] < floor_shifts for d in active_doctors)
    if any_under_floor:
        under_floor = [d for d in candidates if assigned_count[d] < floor_shifts]
        if under_floor:
            candidates = under_floor

    # ★ C-H列（土日大学）はカテ当番医師を最優先（カテなし医師は最後の手段）
    # カテ当番医師がいる場合、カテなし医師より優先して配置
    if is_ch_slot(idx) and candidates:
        kate_docs_on_day = [d for d in candidates if get_sched_code(date, d)]
        if kate_docs_on_day:
            candidates = kate_docs_on_day

    # gap
    gaps = {}
    for d in candidates:
        if not assigned_dates[d]:
            gaps[d] = 999
        else:
            gaps[d] = min(abs((pd.to_datetime(date) - x).days) for x in assigned_dates[d])

    # 優先順位: 7,4,5,比率(B-K/L-Y),2,3,6,8,1(>=4),10

    # 7 全体（前月+今月）
    metric_total = {d: prev_total[d] + assigned_count[d] for d in candidates}
    min_total = min(metric_total.values())
    candidates = [d for d in candidates if metric_total[d] == min_total]

    # 4 大学/外病院偏り（前月+今月）
    if is_BG:
        metric_bg = {d: prev_bg[d] + assigned_bg[d] for d in candidates}
        mb = min(metric_bg.values())
        candidates = [d for d in candidates if metric_bg[d] == mb]
    elif is_LY_range:
        metric_ht = {d: prev_ht[d] + assigned_ht[d] for d in candidates}
        mh = min(metric_ht.values())
        candidates = [d for d in candidates if metric_ht[d] == mh]

    # 5 B〜E / F〜G
    if is_BG:
        if is_BE:
            mbe = min(assigned_be[d] for d in candidates)
            candidates = [d for d in candidates if assigned_be[d] == mbe]
        else:
            mfg = min(assigned_fg[d] for d in candidates)
            candidates = [d for d in candidates if assigned_fg[d] == mfg]

    # B-K / L-Y の比率バランス（除外医師以外）
    if (is_BK or is_LY) and candidates:
        def imbalance_score(doc):
            if doc in RATIO_EXEMPT_DOCTORS:
                return 0
            bk = assigned_bk[doc] + (1 if is_BK else 0)
            ly = assigned_ly[doc] + (1 if is_LY else 0)
            return abs(bk - ly)

        min_imbalance = min(imbalance_score(d) for d in candidates)
        candidates = [d for d in candidates if imbalance_score(d) == min_imbalance]

    # (削除: 同一病院重複は絶対禁忌として collect_candidates でチェック済み)

    # 3 B〜G はカテ表あり優先（ソフト優先）
    if is_BG:
        with_sched = [d for d in candidates if get_sched_code(date, d)]
        if with_sched:
            candidates = with_sched

    # 6 平日/休日偏り（前月+今月）
    holi_flag = (
        is_holiday(date)
        or dow >= 5
        or (weekday and idx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX))
    )
    if holi_flag:
        metric_we = {d: prev_weekend[d] + assigned_weekend[d] for d in candidates}
        mwe = min(metric_we.values())
        candidates = [d for d in candidates if metric_we[d] == mwe]
    else:
        metric_wd = {d: prev_weekday[d] + assigned_weekday[d] for d in candidates}
        mwd = min(metric_wd.values())
        candidates = [d for d in candidates if metric_wd[d] == mwd]

    # 6.5 全体の平日/休日差を最小化（ハード制約: 差2以上は違反）
    def _wd_we_diff(d):
        wd = prev_weekday[d] + assigned_weekday[d] + (0 if holi_flag else 1)
        we = prev_weekend[d] + assigned_weekend[d] + (1 if holi_flag else 0)
        return abs(wd - we)
    min_diff = min(_wd_we_diff(d) for d in candidates)
    balanced = [d for d in candidates if _wd_we_diff(d) == min_diff]
    if balanced:
        candidates = balanced

    # 8 floor未満優先
    under_floor = [d for d in candidates if assigned_count[d] < floor_shifts]
    if under_floor:
        candidates = under_floor

    # (削除: gap >= 3 は絶対禁忌として collect_candidates でチェック済み)

    # 9 TARGET_CAP未達の医師を優先（v6.5.7: relax_absで超過した医師を抑制）
    under_cap = [d for d in candidates if assigned_count[d] < TARGET_CAP.get(d, 0)]
    if under_cap:
        candidates = under_cap

    # 10 同点ならランダム選択（パターン多様性のため）
    # v6.0.2: deterministic tie-break から random.choice に変更
    return random.choice(candidates)

# =========================
# 1ヶ月分生成（greedy）
# =========================
def build_schedule_pattern(seed=0):
    random.seed(seed)

    df = shift_df.copy()
    for col in hospital_cols:
        df[col] = df[col].astype(object)

    assigned_count = {d: 0 for d in doctor_names}
    assigned_dates = {d: set() for d in doctor_names}

    assigned_bg = {d: 0 for d in doctor_names}
    assigned_ht = {d: 0 for d in doctor_names}
    assigned_weekday = {d: 0 for d in doctor_names}
    assigned_weekend = {d: 0 for d in doctor_names}
    assigned_be = {d: 0 for d in doctor_names}
    assigned_fg = {d: 0 for d in doctor_names}
    assigned_bk = {d: 0 for d in doctor_names}
    assigned_ly = {d: 0 for d in doctor_names}
    assigned_bh = {d: 0 for d in doctor_names}  # B〜H列の割当回数（旧: 2回まで）
    assigned_bi = {d: 0 for d in doctor_names}  # v6.0.0: B/I列の合計（HARD-001: 1回まで）
    assigned_chjk = {d: 0 for d in doctor_names}  # v6.0.0: C-H/J-K列の合計（HARD-002: 1回まで）
    assigned_hosp_count = {d: defaultdict(int) for d in doctor_names}
    bg_cat = {d: defaultdict(int) for d in doctor_names}
    assigned_bg_dates = {d: set() for d in doctor_names}  # v6.5.0: 大学系割当日（ABS-012改: 7日間隔ルール）
    semi001_violation_weeks = {d: set() for d in doctor_names}  # v6.5.4: SEMI-001違反週（月曜始まり）

    def _update_counts(doc, date, hosp):
        """割当カウントの一括更新ヘルパー"""
        assigned_count[doc] += 1
        assigned_dates[doc].add(date)
        assigned_hosp_count[doc][hosp] += 1

        hidx = shift_df.columns.get_loc(hosp)
        if B_COL_INDEX <= hidx <= K_COL_INDEX:
            assigned_bg[doc] += 1
            assigned_bg_dates[doc].add(date)
            if B_COL_INDEX <= hidx <= E_COL_INDEX:
                assigned_be[doc] += 1
            elif F_COL_INDEX <= hidx <= G_COL_INDEX:
                assigned_fg[doc] += 1
            bg_cat[doc][classify_bg_category(date, hosp)] += 1
        elif L_COL_INDEX <= hidx <= L_Y_END_INDEX:
            assigned_ht[doc] += 1

        if B_H_START_INDEX <= hidx <= B_H_END_INDEX:
            assigned_bh[doc] += 1
        if hidx == B_COL_INDEX or hidx == I_COL_INDEX:
            assigned_bi[doc] += 1
        if (C_COL_INDEX <= hidx <= H_COL_INDEX) or (J_COL_INDEX <= hidx <= K_COL_INDEX):
            assigned_chjk[doc] += 1

        dow = date.weekday()
        weekday = dow < 5
        holi_flag = (
            is_holiday(date)
            or dow >= 5
            or (weekday and hidx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX))
        )
        if holi_flag:
            assigned_weekend[doc] += 1
        else:
            assigned_weekday[doc] += 1

        if is_bk_slot(hidx):
            assigned_bk[doc] += 1
        elif is_ly_slot(hidx):
            assigned_ly[doc] += 1

    # 固定当直
    for date in all_dates:
        for ridx, hosp, doc in slots_by_date[date]["preassigned"]:
            df.at[ridx, hosp] = doc
            _update_counts(doc, date, hosp)

    # 自動割当
    # 枠決定順序: 大学休日(C-H) → 大学平日(B,I-K) → 外病院(L-Y)
    # C-H列はカテ当番制約があるため先に埋める
    def slot_priority(slot_tuple):
        ridx, hosp = slot_tuple
        hidx = shift_df.columns.get_loc(hosp)
        # C-H列（休日大学系）: 優先度0（最初）
        if C_COL_INDEX <= hidx <= H_COL_INDEX:
            return (0, hidx)
        # B列、I-K列（平日大学系）: 優先度1
        elif hidx == B_COL_INDEX or (I_COL_INDEX <= hidx <= K_COL_INDEX):
            return (1, hidx)
        # L-Y列（外病院）: 優先度2（最後）
        else:
            return (2, hidx)

    for date in all_dates:
        free_slots = slots_by_date[date]["free"].copy()
        # 優先度順にソート後、同一優先度内でシャッフル
        random.shuffle(free_slots)  # まずシャッフルしてランダム性を確保
        free_slots.sort(key=slot_priority)  # 安定ソートで優先度順に

        for ridx, hosp in free_slots:
            chosen = choose_doctor_for_slot(
                date=date,
                hospital_name=hosp,
                assigned_count=assigned_count,
                assigned_dates=assigned_dates,
                assigned_bg=assigned_bg,
                assigned_ht=assigned_ht,
                assigned_weekday=assigned_weekday,
                assigned_weekend=assigned_weekend,
                assigned_be=assigned_be,
                assigned_fg=assigned_fg,
                assigned_bk=assigned_bk,
                assigned_ly=assigned_ly,
                assigned_bh=assigned_bh,
                assigned_bi=assigned_bi,        # v6.0.0
                assigned_chjk=assigned_chjk,    # v6.0.0
                assigned_hosp_count=assigned_hosp_count,
                assigned_bg_dates=assigned_bg_dates,  # v6.5.0
                semi001_violation_weeks=semi001_violation_weeks,  # v6.5.4
            )
            if chosen is None:
                # v6.0.1 フォールバック: 絶対禁忌(ABS)をすべてチェック
                hidx = shift_df.columns.get_loc(hosp)
                is_bg_slot_here = B_COL_INDEX <= hidx <= K_COL_INDEX
                is_ly_slot_here = L_COL_INDEX <= hidx <= L_Y_END_INDEX
                day_of_week = pd.to_datetime(date).weekday()

                def is_valid_fallback(d):
                    code = get_avail_code(date, d)
                    # ABS-001: コード0禁止
                    if code == 0:
                        return False
                    # ABS-002: コード2はB〜Q列のみ
                    if code == 2 and not (B_COL_INDEX <= hidx <= Q_COL_INDEX):
                        return False
                    # ABS-003: コード3はL〜Y列のみ
                    if code == 3 and not is_ly_slot_here:
                        return False
                    # ABS-004: カテ表コードありの日はL〜Y列不可
                    if is_ly_slot_here and get_sched_code(date, d):
                        return False
                    # ABS-005: 水曜日L〜Y列禁止医師
                    if day_of_week == 2 and is_ly_slot_here and d in WED_FORBIDDEN_DOCTORS:
                        return False
                    # ABS-006: 同日重複禁止
                    if date in assigned_dates[d]:
                        return False
                    # ABS-007: gap >= 3日必須
                    if assigned_dates[d]:
                        min_gap = min(abs((pd.to_datetime(date) - x).days) for x in assigned_dates[d])
                        if min_gap < 3:
                            return False
                    # ABS-008: 同一病院重複禁止
                    if assigned_hosp_count[d].get(hosp, 0) >= 1:
                        return False
                    # ABS-010: TARGET_CAP厳守
                    if assigned_count[d] >= TARGET_CAP.get(d, 0):
                        return False
                    # ABS-011: 大学系2回まで
                    if is_bg_slot_here and assigned_bg[d] >= 2:
                        return False
                    # ABS-012改: 大学系は7日間隔必須（v6.5.0）
                    if is_bg_slot_here and assigned_bg_dates[d]:
                        min_bg_gap = min(abs((pd.to_datetime(date) - x).days) for x in assigned_bg_dates[d])
                        if min_bg_gap < 7:
                            return False
                    return True

                remaining = [d for d in doctor_names if is_valid_fallback(d)]
                if remaining:
                    fallback_doc = min(remaining, key=lambda d: (assigned_count[d], doctor_col_index[d]))
                else:
                    # 全員が絶対禁忌に該当する場合は未割当のまま（None）
                    fallback_doc = None
                if fallback_doc is not None:
                    df.at[ridx, hosp] = fallback_doc
                    chosen = fallback_doc
                # fallback_doc が None の場合は未割当のまま（後続処理をスキップ）
            else:
                df.at[ridx, hosp] = chosen

            # chosen が None でなければカウント更新
            if chosen is not None:
                _update_counts(chosen, date, hosp)

                # v6.5.4: SEMI-001違反週の記録（B列でカテ表なし）
                hidx = shift_df.columns.get_loc(hosp)
                if hidx == B_COL_INDEX:
                    if chosen in SCHEDULE_CODE_HOLDERS and not get_sched_code(date, chosen):
                        week_start = get_monday_week_start(date)
                        semi001_violation_weeks[chosen].add(week_start)

    return (
        df,
        assigned_count,
        assigned_bg,
        assigned_ht,
        assigned_weekday,
        assigned_weekend,
        assigned_bk,
        assigned_ly,
        bg_cat,
    )

# =========================
# slot_meta / movable_positions（ローカル探索用）
# =========================
slot_meta = {}  # (ridx,hosp) -> (date, fixed)
movable_positions = []  # (ridx,hosp,date)

for date in all_dates:
    for ridx, hosp, doc in slots_by_date[date]["preassigned"]:
        slot_meta[(ridx, hosp)] = (date, True)
    for ridx, hosp in slots_by_date[date]["free"]:
        slot_meta[(ridx, hosp)] = (date, False)
        movable_positions.append((ridx, hosp, date))

def is_preassigned_slot(ridx, hosp):
    """v6.2.0: そのスロットが固定割当（事前割当）かどうかを判定"""
    meta = slot_meta.get((ridx, hosp))
    if meta is None:
        return False
    return meta[1]  # fixed flag

def get_bg_week_start(date):
    """v6.4.0: 日曜始まりの週の開始日（日曜日）を返す
    大学系週1回ルール(ABS-012)で使用
    """
    d = pd.to_datetime(date).normalize()
    # タイムゾーン情報を削除
    if d.tz is not None:
        d = d.tz_localize(None)
    # 日曜日=0として計算（Python: 月曜日=0なので調整）
    days_since_sunday = (d.weekday() + 1) % 7
    return d - pd.Timedelta(days=days_since_sunday)

def get_monday_week_start(date):
    """v6.5.4: 月曜始まりの週の開始日（月曜日）を返す
    SEMI-001週1回ルールで使用
    例: 土曜と日曜は同じ週、日曜と月曜は別の週
    """
    d = pd.to_datetime(date).normalize()
    if d.tz is not None:
        d = d.tz_localize(None)
    # Python weekday(): 月曜=0, ..., 日曜=6
    days_since_monday = d.weekday()
    return d - pd.Timedelta(days=days_since_monday)

# =========================
# パターン統計再計算（pattern_df から）
# =========================
def recompute_stats(pattern_df):
    counts = {d: 0 for d in doctor_names}
    bg_counts = {d: 0 for d in doctor_names}
    ht_counts = {d: 0 for d in doctor_names}
    wd_counts = {d: 0 for d in doctor_names}
    we_counts = {d: 0 for d in doctor_names}
    bk_counts = {d: 0 for d in doctor_names}
    ly_counts = {d: 0 for d in doctor_names}
    bg_cat = {d: defaultdict(int) for d in doctor_names}
    assigned_hosp_count = {d: defaultdict(int) for d in doctor_names}
    doc_assignments = {d: [] for d in doctor_names}  # (date,hosp)
    unassigned = []  # (date,hosp,ridx)
    # CC（大型連休特別シフト）カウント - 各種バランス計算から除外用
    cc_counts = {d: 0 for d in doctor_names}
    cc_bg_counts = {d: 0 for d in doctor_names}  # CCのうち大学系
    cc_ht_counts = {d: 0 for d in doctor_names}  # CCのうち外病院

    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = pattern_df.at[ridx, hosp]

        # 未割り当てチェック（None, NaN, 非医師名の場合）
        if pd.isna(val):
            unassigned.append((date, hosp, ridx))
            continue
        if not isinstance(val, str):
            # 数値など（1など）は未割り当て
            unassigned.append((date, hosp, ridx))
            continue
        v = normalize_name(val)  # 🔧 FIX
        if v not in doctor_names:
            # 医師名でない文字列（"UNASSIGNED"や"1"など）も未割り当て
            unassigned.append((date, hosp, ridx))
            continue

        doc = v
        counts[doc] += 1
        assigned_hosp_count[doc][hosp] += 1
        doc_assignments[doc].append((date, hosp))

        # CC判定
        is_cc = is_cc_assignment(date, doc)
        if is_cc:
            cc_counts[doc] += 1

        hidx = shift_df.columns.get_loc(hosp)
        # 大学系はB〜K列（B_COL_INDEX=1 〜 K_COL_INDEX=10）
        if B_COL_INDEX <= hidx <= B_K_END_INDEX:
            bg_counts[doc] += 1
            bg_cat[doc][classify_bg_category(date, hosp)] += 1
            if is_cc:
                cc_bg_counts[doc] += 1
        # 外病院はL〜Y列（L_COL_INDEX=11 〜 Y_COL_INDEX=24）
        elif L_COL_INDEX <= hidx <= L_Y_END_INDEX:
            ht_counts[doc] += 1
            if is_cc:
                cc_ht_counts[doc] += 1

        dow = date.weekday()
        weekday = dow < 5
        holi_flag = (
            is_holiday(date)
            or dow >= 5
            or (weekday and hidx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX))
        )
        if holi_flag:
            we_counts[doc] += 1
        else:
            wd_counts[doc] += 1

        if is_bk_slot(hidx):
            bk_counts[doc] += 1
        elif is_ly_slot(hidx):
            ly_counts[doc] += 1

    return (
        counts,
        bg_counts,
        ht_counts,
        wd_counts,
        we_counts,
        bk_counts,
        ly_counts,
        bg_cat,
        assigned_hosp_count,
        doc_assignments,
        unassigned,
        cc_counts,
        cc_bg_counts,
        cc_ht_counts,
    )

# =========================
# スコア評価（raw_scoreも保持して 0 で潰れないように）
# =========================
def evaluate_schedule_with_raw(
    pattern_df,
    assigned_count,
    assigned_bg,
    assigned_ht,
    assigned_weekday,
    assigned_weekend,
    assigned_bk,
    assigned_ly,
):
    # UNASSIGNED - slot_metaに登録されたスロットのうち、医師名が入っていないものをカウント
    # None, NaN, 非医師名（1, 〇など）も未割り当てとしてカウント
    unassigned_slots = 0
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        v = pattern_df.at[ridx, hosp]
        # 医師名でない場合は未割り当て
        if pd.isna(v):
            unassigned_slots += 1
        elif isinstance(v, str):
            v_norm = normalize_name(v)
            if v_norm not in doctor_names:
                unassigned_slots += 1
        else:
            # 数値など（1など）は未割り当て
            unassigned_slots += 1

    # CC（大型連休特別シフト）カウント - バランス計算から除外用
    cc_counts = {d: 0 for d in doctor_names}
    cc_bg_counts = {d: 0 for d in doctor_names}
    cc_ht_counts = {d: 0 for d in doctor_names}
    cc_hosp_counts = {d: defaultdict(int) for d in doctor_names}  # CC分の病院別カウント
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        for hosp in hospital_cols:
            val = pattern_df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc not in doctor_names:
                continue
            if is_cc_assignment(date, doc):
                cc_counts[doc] += 1
                cc_hosp_counts[doc][hosp] += 1
                hidx = shift_df.columns.get_loc(hosp)
                if B_COL_INDEX <= hidx <= B_K_END_INDEX:
                    cc_bg_counts[doc] += 1
                elif L_COL_INDEX <= hidx <= L_Y_END_INDEX:
                    cc_ht_counts[doc] += 1

    # cap違反
    cap_violations = 0
    for doc in doctor_names:
        cap = TARGET_CAP.get(doc, 0)
        if assigned_count.get(doc, 0) > cap:
            cap_violations += (assigned_count[doc] - cap)

    # CODE_2医師のTARGET_CAP超過チェック
    # v6.0.5: CODE_2もEXTRA対象のため、TARGET_CAPベースで判定（BASE_TARGETではなく）
    code_2_extra_violations = 0
    for doc in CODE_2_DOCTORS:
        if doc not in active_doctors:
            continue
        cap = TARGET_CAP.get(doc, 0)
        if assigned_count.get(doc, 0) > cap:
            code_2_extra_violations += (assigned_count[doc] - cap)

    # 全合計公平性（activeのみ、CC除外）
    # CCは大型連休特別シフトなので公平性計算から除外
    active_counts_no_cc = [assigned_count.get(d, 0) - cc_counts.get(d, 0) for d in active_doctors]
    max_c = max(active_counts_no_cc) if active_counts_no_cc else 0
    min_c = min(active_counts_no_cc) if active_counts_no_cc else 0
    diff_total = max_c - min_c
    # 差が2以上の場合、不満が高いので強いペナルティ
    # 例: min=2, max=4の場合、4回の医師から2回の医師に渡すべき
    if diff_total >= 2:
        fairness_penalty = diff_total * 2  # 2倍のペナルティ
    else:
        fairness_penalty = max(0, diff_total - 1)

    # v6.5.9: 累計（前月+今月）全合計の公平性（SOFT-009）
    # 前月までの累積を含めた通算回数の偏りを最適化対象にする
    # （従来は前月累積がgreedy候補ソートと表示にしか使われず、重み0だった）
    cum_totals = [
        prev_total.get(d, 0) + assigned_count.get(d, 0) - cc_counts.get(d, 0)
        for d in active_doctors
    ]
    cum_total_spread = (max(cum_totals) - min(cum_totals)) if cum_totals else 0

    # gap(4日未満) と 同一病院重複
    # v6.2.0: 各割当が固定割当かどうかも記録（gap/dup計算で固定割当を除外するため）
    dates_by_doc = defaultdict(list)  # doc -> [(date, is_preassigned), ...]
    hosp_counts_by_doc = {doc: defaultdict(int) for doc in doctor_names}
    hosp_preassigned_counts = {doc: defaultdict(int) for doc in doctor_names}  # 固定割当分

    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)  # 🔧 FIX
        for hosp in hospital_cols:
            val = pattern_df.at[ridx, hosp]
            val_norm = normalize_name(val) if isinstance(val, str) else ""  # 🔧 FIX
            if val_norm in doctor_names:
                is_pre = is_preassigned_slot(ridx, hosp)
                dates_by_doc[val_norm].append((date, is_pre))
                hosp_counts_by_doc[val_norm][hosp] += 1
                if is_pre:
                    hosp_preassigned_counts[val_norm][hosp] += 1

    gap_violations = 0
    for doc, date_flags in dates_by_doc.items():
        sorted_dates = sorted(date_flags, key=lambda x: x[0])
        for i in range(1, len(sorted_dates)):
            if (sorted_dates[i][0] - sorted_dates[i - 1][0]).days < 3:
                # v6.2.0: 固定割当が片方でも含まれるgapは許容（意図的な配置）
                if sorted_dates[i][1] or sorted_dates[i - 1][1]:
                    continue
                gap_violations += 1

    hosp_dup_violations = 0
    external_hosp_dup_violations = 0  # 外病院重複（厳しく扱う）
    for doc, hdict in hosp_counts_by_doc.items():
        for hosp, c in hdict.items():
            # CC分を除外（CCは特別シフトなので重複カウントから除外）
            c_no_cc = c - cc_hosp_counts.get(doc, {}).get(hosp, 0)
            # v6.2.0: 固定割当分も除外（固定割当による重複は許容）
            c_no_cc -= hosp_preassigned_counts.get(doc, {}).get(hosp, 0)
            if c_no_cc > 1:
                # 病院が外病院（L～Y列）かどうかを判定
                hidx = shift_df.columns.get_loc(hosp)
                if L_COL_INDEX <= hidx <= L_Y_END_INDEX:
                    external_hosp_dup_violations += (c_no_cc - 1)
                else:
                    hosp_dup_violations += (c_no_cc - 1)

    # 偏り（累計：前月+今月）の spread
    bg_vals = [prev_bg[d] + assigned_bg.get(d, 0) for d in active_doctors]
    ht_vals = [prev_ht[d] + assigned_ht.get(d, 0) for d in active_doctors]
    wd_vals = [prev_weekday[d] + assigned_weekday.get(d, 0) for d in active_doctors]
    we_vals = [prev_weekend[d] + assigned_weekend.get(d, 0) for d in active_doctors]

    bg_spread = (max(bg_vals) - min(bg_vals)) if bg_vals else 0
    ht_spread = (max(ht_vals) - min(ht_vals)) if ht_vals else 0
    wd_spread = (max(wd_vals) - min(wd_vals)) if wd_vals else 0
    we_spread = (max(we_vals) - min(we_vals)) if we_vals else 0

    bk_ly_imbalance = 0
    for doc in active_doctors:
        if doc in RATIO_EXEMPT_DOCTORS:
            continue
        bk_val = assigned_bk.get(doc, 0)
        ly_val = assigned_ly.get(doc, 0)
        bk_ly_imbalance += abs(bk_val - ly_val)

    # 可否コード1.2の医師が大学系0回の場合のペナルティ
    code_1_2_violations = 0
    for doc in CODE_1_2_DOCTORS:
        if assigned_bg.get(doc, 0) == 0:
            code_1_2_violations += 1

    # 大学系と外病院の差が3以上の場合のペナルティ（CC除外）
    bg_ht_imbalance_violations = 0
    for doc in active_doctors:
        # CCは大型連休特別シフトなのでバランス計算から除外
        bg = assigned_bg.get(doc, 0) - cc_bg_counts.get(doc, 0)
        ht = assigned_ht.get(doc, 0) - cc_ht_counts.get(doc, 0)
        diff = abs(bg - ht)
        if diff >= 3:
            bg_ht_imbalance_violations += (diff - 2)  # 差が3以上の超過分をカウント

    # 大学病院2回の場合、平日1回+休日1回のバランス違反
    bg_weekday_weekend_imbalance = 0
    bg_over_2_violations = 0  # 大学3回以上の違反（不満が高い）
    ht_0_violations = 0  # 外病院0回の違反（ハード制約：大学3回以上を防ぐ）
    bg_weekday_over_violations = 0  # 大学の平日偏り（平日2回以上は不満）
    for doc in active_doctors:
        if doc in RATIO_EXEMPT_DOCTORS:  # コード3は外病院専門なので除外
            continue
        # CC除外（大型連休特別シフトはバランス計算から除外）
        bg_total_no_cc = assigned_bg.get(doc, 0) - cc_bg_counts.get(doc, 0)
        ht_total_no_cc = assigned_ht.get(doc, 0) - cc_ht_counts.get(doc, 0)
        # 元の値（ht_0_violationsなどハード制約用）
        bg_total = assigned_bg.get(doc, 0)
        ht_total = assigned_ht.get(doc, 0)
        weekday_count = bg_cat[doc].get("平日", 0)

        # 大学3回以上は不可（CC除外）
        if bg_total_no_cc >= 3:
            bg_over_2_violations += (bg_total_no_cc - 2)

        # 外病院0回かつ大学1回以上はハード制約違反（CCは除外しない：ハード制約）
        if ht_total == 0 and bg_total >= 1:
            ht_0_violations += 1

        # 大学2回の場合、平日1回+休日1回が理想（CC除外）
        if bg_total_no_cc == 2:
            if weekday_count == 0 or weekday_count == 2:
                bg_weekday_weekend_imbalance += 1

        # 大学の平日が2回以上は不満（CC除外）
        # 注：weekday_countからCCを除外するには追加トラッキングが必要
        # 現時点ではweekday_countはそのまま使用（大型連休は平日カウントされにくい）
        if weekday_count >= 2:
            bg_weekday_over_violations += (weekday_count - 1)

    # 全体の平日/休日偏り違反（差が2以上はハード制約違反）
    wd_we_imbalance_violations = 0
    we_0_violations = 0  # 休日0回の違反（ハード制約）
    for doc in active_doctors:
        wd = wd_counts.get(doc, 0)
        we = we_counts.get(doc, 0)
        total = counts.get(doc, 0)
        diff = abs(wd - we)
        if diff >= 2:
            wd_we_imbalance_violations += (diff - 1)
        if we == 0 and total >= 1:
            we_0_violations += 1

    # v6.5.0: 大学系7日間隔違反（7日以内に2回以上）- ABS-012改
    weekly_bg_violations = 0
    bg_dates_by_doc = {doc: [] for doc in doctor_names}  # doc -> [date list]
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize()
        if date.tz is not None:
            date = date.tz_localize(None)
        for hosp in hospital_cols:
            hidx = shift_df.columns.get_loc(hosp)
            # 大学病院（B～K列）か
            if not (B_COL_INDEX <= hidx <= K_COL_INDEX):
                continue
            val = pattern_df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc not in doctor_names:
                continue
            # 固定割当の場合は許容（v6.2.0互換）
            if is_preassigned_slot(ridx, hosp):
                continue
            bg_dates_by_doc[doc].append(date)
    # 7日以内に2回以上の違反をカウント
    for doc in active_doctors:
        dates = sorted(bg_dates_by_doc[doc])
        for i in range(1, len(dates)):
            gap = abs((dates[i] - dates[i-1]).days)
            if gap < 7:
                weekly_bg_violations += 1

    # C-H列（休日大学系）カテ当番違反
    # カテ当番保有医師がその日にカテ当番なしでC-H列に割り当てられている場合
    # v6.2.0: 固定割当は許容（意図的な配置のため）
    ch_kate_violations = 0
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        for hosp in hospital_cols:
            # v6.2.0: 固定割当は許容
            if is_preassigned_slot(ridx, hosp):
                continue
            idx = shift_df.columns.get_loc(hosp)
            if not is_ch_slot(idx):
                continue
            val = pattern_df.at[ridx, hosp]
            val_norm = normalize_name(val) if isinstance(val, str) else ""
            if val_norm in doctor_names:
                if not is_eligible_for_ch_slot(val_norm, date):
                    ch_kate_violations += 1

    penalty = 0
    penalty += fairness_penalty * W_FAIR_TOTAL
    penalty += max(0, cum_total_spread - 1) * W_FAIR_CUM  # v6.5.9: 累計公平性
    penalty += gap_violations * W_GAP
    penalty += hosp_dup_violations * W_HOSP_DUP
    penalty += external_hosp_dup_violations * W_EXTERNAL_HOSP_DUP  # 外病院重複は厳格
    penalty += unassigned_slots * W_UNASSIGNED
    penalty += cap_violations * W_CAP
    penalty += code_2_extra_violations * 300  # CODE_2医師のn+1違反は厳格（ハード制約）
    penalty += code_1_2_violations * 150  # 1.2の医師が大学系0回の場合、大きなペナルティ
    penalty += bg_ht_imbalance_violations * 100  # 大学系と外病院の差が3以上の場合、大きなペナルティ
    penalty += bg_weekday_weekend_imbalance * 50  # 大学病院2回の平日/休日バランス違反
    penalty += bg_over_2_violations * 300  # 大学3回以上の違反（ハード制約）
    penalty += ht_0_violations * 300  # 外病院0回の違反（ハード制約）
    penalty += bg_weekday_over_violations * 80  # 大学の平日偏り（平日2回以上は不満）
    penalty += ch_kate_violations * 120  # C-H列カテ当番違反（優先度高）
    penalty += weekly_bg_violations * 300  # v6.4.0: 大学系週1違反（ABS-012）
    penalty += wd_we_imbalance_violations * 300  # 全体の平日/休日偏り（ハード制約）
    penalty += we_0_violations * 300  # 休日0回（ハード制約）

    penalty += max(0, bg_spread - 1) * W_BG_SPREAD
    penalty += max(0, ht_spread - 1) * W_HT_SPREAD
    penalty += max(0, wd_spread - 1) * W_WD_SPREAD
    penalty += max(0, we_spread - 1) * W_WE_SPREAD
    penalty += bk_ly_imbalance * W_BK_LY_BALANCE

    raw_score = 100 - penalty
    score = max(raw_score, 0)

    metrics = {
        "raw_score": float(raw_score),
        "penalty_total": float(penalty),
        "max_minus_min_total_active": int(diff_total),
        "gap_violations": int(gap_violations),
        "hospital_dup_violations": int(hosp_dup_violations),
        "external_hosp_dup_violations": int(external_hosp_dup_violations),
        "unassigned_slots": int(unassigned_slots),
        "cap_violations": int(cap_violations),
        "code_2_extra_violations": int(code_2_extra_violations),
        "code_1_2_violations": int(code_1_2_violations),
        "bg_ht_imbalance_violations": int(bg_ht_imbalance_violations),
        "bg_weekday_weekend_imbalance": int(bg_weekday_weekend_imbalance),
        "bg_over_2_violations": int(bg_over_2_violations),
        "ht_0_violations": int(ht_0_violations),
        "bg_weekday_over_violations": int(bg_weekday_over_violations),
        "ch_kate_violations": int(ch_kate_violations),
        "weekly_bg_violations": int(weekly_bg_violations),  # v6.4.0: 大学系週1違反（ABS-012）
        "wd_we_imbalance_violations": int(wd_we_imbalance_violations),  # 全体の平日/休日偏り（差>=2）
        "we_0_violations": int(we_0_violations),  # 休日0回違反
        "total_spread_cum": float(cum_total_spread),  # v6.5.9: 累計全合計のmax-min
        "bg_spread_cum": float(bg_spread),
        "ht_spread_cum": float(ht_spread),
        "weekday_spread_cum": float(wd_spread),
        "weekend_spread_cum": float(we_spread),
        "bk_ly_imbalance": int(bk_ly_imbalance),
    }
    return score, raw_score, metrics

# =========================
# ローカル探索（入替 swap）
# 🔧 FIX: date_doc_count を完全な defaultdict(lambda: defaultdict(int)) に変更
# 🔧 FIX: 同日重複チェックの強化
# =========================
def can_assign_doc_to_slot(doc, date, hosp):
    """静的制約のみ（可否コード、カテ表、水曜禁止）"""
    idx = shift_df.columns.get_loc(hosp)
    dow = pd.to_datetime(date).weekday()

    code = get_avail_code(date, doc)
    if code == 0:
        return False
    # 可否コード2 → B〜Q列のみ可
    if code == 2 and not (B_COL_INDEX <= idx <= Q_COL_INDEX):
        return False
    # 可否コード3 → L〜Y列のみ可
    if code == 3 and not (L_COL_INDEX <= idx <= L_Y_END_INDEX):
        return False
    # その日にカテ表コードあり → L〜Y列不可
    if L_COL_INDEX <= idx <= L_Y_END_INDEX:
        if get_sched_code(date, doc):
            return False
    # B〜K列はカテ表コード保有医師のみカテ表コードが必要（EXTRA医師は例外）
    if B_COL_INDEX <= idx <= B_K_END_INDEX:
        if doc in SCHEDULE_CODE_HOLDERS and not get_sched_code(date, doc) and doc not in EXTRA_ALLOWED:
            return False
    # 水曜日L〜Y列禁止医師
    if dow == 2 and L_COL_INDEX <= idx <= L_Y_END_INDEX and doc in WED_FORBIDDEN_DOCTORS:
        return False
    return True


def is_valid_full_assignment(doc, date, hosp, doc_assignments, counts, bg_counts, assigned_hosp_count, already_on_date=None):
    """
    v6.0.3: 全ABS制約を統合チェック（静的+動的）
    全てのfix関数はこの関数を使って候補医師の妥当性を判定する。
    これにより、fix関数が他の制約を壊すことを防止する。

    Args:
        doc: 医師名
        date: 日付
        hosp: 病院列名
        doc_assignments: {doc: [(date, hosp), ...]} 現在の割当状態
        counts: {doc: int} 現在の割当回数
        bg_counts: {doc: int} 大学系割当回数
        assigned_hosp_count: {doc: {hosp: int}} 病院別割当回数
        already_on_date: set of doc names already assigned on this date (optional, computed if None)
    """
    # === 静的制約（ABS-001〜005） ===
    if not can_assign_doc_to_slot(doc, date, hosp):
        return False

    idx = shift_df.columns.get_loc(hosp)
    is_bg = B_COL_INDEX <= idx <= K_COL_INDEX
    is_external = L_COL_INDEX <= idx <= L_Y_END_INDEX

    # === ABS-006: 同日重複禁止 ===
    if already_on_date is not None:
        if doc in already_on_date:
            return False
    else:
        # already_on_dateが提供されない場合、doc_assignmentsから推定
        doc_dates = [d for d, h in doc_assignments.get(doc, [])]
        if date in doc_dates:
            return False

    # === ABS-007: gap >= 3日必須 ===
    doc_date_list = [d for d, h in doc_assignments.get(doc, [])]
    if doc_date_list:
        min_gap = min(abs((pd.to_datetime(date) - d).days) for d in doc_date_list)
        if min_gap < 3:
            return False

    # === ABS-008: 同一病院重複禁止（fix関数では外病院のみ。初期生成時は全列で禁止） ===
    if is_external and assigned_hosp_count.get(doc, {}).get(hosp, 0) >= 1:
        return False

    # === ABS-010: TARGET_CAP厳守 ===
    if counts.get(doc, 0) >= TARGET_CAP.get(doc, 0):
        return False

    # === ABS-011: 大学系2回まで ===
    if is_bg and bg_counts.get(doc, 0) >= 2:
        return False

    return True

def build_date_doc_count(pattern_df):
    """date -> doc -> count（同日複数割当検出も兼ねる）"""
    # 🔧 FIX: 完全な nested defaultdict に変更
    date_doc_count = defaultdict(lambda: defaultdict(int))
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = pattern_df.at[ridx, hosp]
        if isinstance(val, str):
            v = normalize_name(val)  # 🔧 FIX
            if v in doctor_names:
                date_doc_count[date][v] += 1
    return date_doc_count

def collect_violation_docs_from_assignments(doc_assignments, assigned_hosp_count):
    bad = set()
    # gap
    for doc, assigns in doc_assignments.items():
        dlist = sorted([d for d, _ in assigns])
        for i in range(1, len(dlist)):
            if (dlist[i] - dlist[i - 1]).days < 3:
                bad.add(doc)
                break
    # hospital dup
    for doc, hdict in assigned_hosp_count.items():
        if any(c > 1 for c in hdict.values()):
            bad.add(doc)
    return bad

def is_better_raw(new_raw, new_metrics, cur_raw, cur_metrics):
    if new_raw > cur_raw:
        return True
    if new_raw < cur_raw:
        return False
    # tie-break（重要度順）
    keys = [
        "unassigned_slots",
        "cap_violations",
        "gap_violations",
        "hospital_dup_violations",
        "max_minus_min_total_active",
        "bk_ly_imbalance",
        "bg_spread_cum",
        "ht_spread_cum",
        "weekday_spread_cum",
        "weekend_spread_cum",
    ]
    return tuple(new_metrics.get(k, 0) for k in keys) < tuple(cur_metrics.get(k, 0) for k in keys)

def local_search_swap(pattern_df, max_iters=2000, patience=800, refresh_every=200, seed=0):
    """入替（swap）局所探索：preassignedは動かさず、free枠のみを対象に改善する"""
    if not movable_positions:
        # 動かせる枠が無い（全部固定など）
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, *_ = recompute_stats(pattern_df)
        score, raw_score, metrics = evaluate_schedule_with_raw(
            pattern_df,
            counts,
            bg_counts,
            ht_counts,
            wd_counts,
            we_counts,
            bk_counts,
            ly_counts,
        )
        return pattern_df.copy(), score, raw_score, metrics

    rng = random.Random(seed)
    df = pattern_df.copy()

    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)
    cur_score, cur_raw, cur_metrics = evaluate_schedule_with_raw(
        df,
        counts,
        bg_counts,
        ht_counts,
        wd_counts,
        we_counts,
        bk_counts,
        ly_counts,
    )
    date_doc_count = build_date_doc_count(df)

    no_improve = 0
    bad_positions = None

    for it in range(1, max_iters + 1):
        if no_improve >= patience:
            break

        if it == 1 or it % refresh_every == 0:
            bad_docs = collect_violation_docs_from_assignments(doc_assignments, assigned_hosp_count)
            if bad_docs:
                tmp = []
                for (ridx, hosp, date) in movable_positions:
                    v = df.at[ridx, hosp]
                    if isinstance(v, str) and normalize_name(v) in bad_docs:  # 🔧 FIX
                        tmp.append((ridx, hosp, date))
                bad_positions = tmp if tmp else None
            else:
                bad_positions = None

        p1 = rng.choice(bad_positions if bad_positions is not None else movable_positions)
        p2 = rng.choice(movable_positions)
        if p1 == p2:
            no_improve += 1
            continue

        r1, h1, d1 = p1
        r2, h2, d2 = p2

        v1 = df.at[r1, h1]
        v2 = df.at[r2, h2]
        if not (isinstance(v1, str) and isinstance(v2, str)):
            no_improve += 1
            continue

        doc1 = normalize_name(v1)  # 🔧 FIX
        doc2 = normalize_name(v2)  # 🔧 FIX
        if doc1 not in doctor_names or doc2 not in doctor_names:
            no_improve += 1
            continue
        if doc1 == doc2:
            no_improve += 1
            continue

        # 🔧 FIX: 同日重複を作らない（同じ日同士のswapもチェック）
        if d1 == d2:
            # 同じ日のslot同士をswapする場合、元々同じdoctorがいなければOK
            # （既にdoc1!=doc2を確認済みなので、追加チェック不要）
            pass
        else:
            # 異なる日のswap
            if date_doc_count[d2][doc1] > 0:  # doc1がd2に既にいる
                no_improve += 1
                continue
            if date_doc_count[d1][doc2] > 0:  # doc2がd1に既にいる
                no_improve += 1
                continue

        # ハード制約
        if not can_assign_doc_to_slot(doc1, d2, h2):
            no_improve += 1
            continue
        if not can_assign_doc_to_slot(doc2, d1, h1):
            no_improve += 1
            continue

        # swap（in-place）
        df.at[r1, h1], df.at[r2, h2] = doc2, doc1

        # 🔧 FIX: date_doc_count 更新（defaultdictなのでKeyErrorなし）
        if d1 != d2:
            date_doc_count[d1][doc1] -= 1
            if date_doc_count[d1][doc1] <= 0:
                del date_doc_count[d1][doc1]
            date_doc_count[d2][doc2] -= 1
            if date_doc_count[d2][doc2] <= 0:
                del date_doc_count[d2][doc2]
            date_doc_count[d1][doc2] += 1
            date_doc_count[d2][doc1] += 1

        # 再評価（全再計算）
        counts2, bg2, ht2, wd2, we2, bk2, ly2, bg_cat2, assigned_hosp_count2, doc_assignments2, unassigned2, *_ = recompute_stats(df)
        new_score, new_raw, new_metrics = evaluate_schedule_with_raw(
            df,
            counts2,
            bg2,
            ht2,
            wd2,
            we2,
            bk2,
            ly2,
        )

        # 絶対禁忌チェック: gap違反または外病院重複があれば拒否
        new_gap_violations = new_metrics.get("gap_violations", 0)
        new_external_hosp_dup = new_metrics.get("external_hosp_dup_violations", 0)
        if new_gap_violations > 0 or new_external_hosp_dup > 0:
            # revert
            df.at[r1, h1], df.at[r2, h2] = doc1, doc2
            if d1 != d2:
                date_doc_count[d1][doc2] -= 1
                if date_doc_count[d1][doc2] <= 0:
                    del date_doc_count[d1][doc2]
                date_doc_count[d2][doc1] -= 1
                if date_doc_count[d2][doc1] <= 0:
                    del date_doc_count[d2][doc1]
                date_doc_count[d1][doc1] += 1
                date_doc_count[d2][doc2] += 1
            no_improve += 1
        elif is_better_raw(new_raw, new_metrics, cur_raw, cur_metrics):
            cur_score, cur_raw, cur_metrics = new_score, new_raw, new_metrics
            counts, bg_counts, ht_counts, wd_counts, we_counts, bg_cat = counts2, bg2, ht2, wd2, we2, bg_cat2
            assigned_hosp_count, doc_assignments = assigned_hosp_count2, doc_assignments2
            no_improve = 0
        else:
            # revert
            df.at[r1, h1], df.at[r2, h2] = doc1, doc2
            if d1 != d2:
                date_doc_count[d1][doc2] -= 1
                if date_doc_count[d1][doc2] <= 0:
                    del date_doc_count[d1][doc2]
                date_doc_count[d2][doc1] -= 1
                if date_doc_count[d2][doc1] <= 0:
                    del date_doc_count[d2][doc1]
                date_doc_count[d1][doc1] += 1
                date_doc_count[d2][doc2] += 1
            no_improve += 1

    return df, cur_score, cur_raw, cur_metrics

# =========================
# サマリー列（Sheet4 の列を基準に自動生成）
# =========================
META_COLS_SHEET4 = {"カテ当番", "出張日", "出張先"}
BASE_SUMMARY_COLS = {"全合計", "大学合計", "外病院合計", "平日", "休日合計"}

UNIV7_SET = {"大学平日", "大学土曜昼", "大学土曜夜", "大学日曜昼", "大学日曜夜", "大学祝日昼", "大学祝日夜"}
UNIV7_ORDER = [c for c in sheet4_raw_out.columns if c in UNIV7_SET]
if not UNIV7_ORDER:
    UNIV7_ORDER = ["大学土曜昼", "大学土曜夜", "大学日曜昼", "大学日曜夜", "大学祝日昼", "大学祝日夜", "大学平日"]

DETAIL_COLS = [
    c for c in sheet4_raw_out.columns
    if c not in META_COLS_SHEET4 and c not in ["氏名"] and c not in BASE_SUMMARY_COLS and c not in UNIV7_SET
]

SUMMARY_DETAIL_COLS = UNIV7_ORDER + DETAIL_COLS
SUMMARY_COLS = ["氏名", "全合計", "大学合計", "外病院合計", "平日", "休日合計"] + SUMMARY_DETAIL_COLS

def count_doc_in_column(df, colname, doc):
    if colname not in df.columns:
        return 0
    s = df[colname]
    cnt = 0
    for v in s:
        if isinstance(v, str) and normalize_name(v) == doc:  # 🔧 FIX
            cnt += 1
    return cnt

def build_summaries(pattern_df, counts, bg_counts, ht_counts, wd_counts, we_counts, bg_cat_local):
    rows_month = []
    rows_total = []

    uni_map = {
        "大学平日": "平日",
        "大学土曜昼": "土曜昼",
        "大学土曜夜": "土曜夜",
        "大学日曜昼": "日曜昼",
        "大学日曜夜": "日曜夜",
        "大学祝日昼": "祝日昼",
        "大学祝日夜": "祝日夜",
    }

    for doc in doctor_names:
        pname = name_match.get(doc)
        base_row = name_to_row.get(pname) if pname and pname in name_to_row else None

        def prev_val(colname):
            if base_row is None:
                return 0.0
            v = base_row.get(colname, 0)
            try:
                return float(v or 0)
            except Exception:
                return 0.0

        # 今月
        row_m = {c: 0.0 for c in SUMMARY_COLS}
        row_m["氏名"] = doc
        row_m["全合計"] = float(counts.get(doc, 0))
        row_m["大学合計"] = float(bg_counts.get(doc, 0))
        row_m["外病院合計"] = float(ht_counts.get(doc, 0))
        row_m["平日"] = float(wd_counts.get(doc, 0))
        row_m["休日合計"] = float(we_counts.get(doc, 0))

        # 大学7分類
        for col in UNIV7_ORDER:
            cat = uni_map.get(col)
            if cat:
                row_m[col] = float(bg_cat_local[doc].get(cat, 0))

        # 病院列（Sheet4準拠）をそのまま数える
        for col in DETAIL_COLS:
            row_m[col] = float(count_doc_in_column(pattern_df, col, doc))

        # 累計（前月＋今月）
        row_t = {c: 0.0 for c in SUMMARY_COLS}
        row_t["氏名"] = doc
        for c in SUMMARY_COLS:
            if c == "氏名":
                continue
            row_t[c] = prev_val(c) + float(row_m.get(c, 0))

        rows_month.append(row_m)
        rows_total.append(row_t)

    return pd.DataFrame(rows_month)[SUMMARY_COLS], pd.DataFrame(rows_total)[SUMMARY_COLS]

# =========================
# 診断シート生成（偏り & gap違反一覧）
# =========================
def build_gap_details(doc_assignments):
    """gap違反（3日未満の間隔）の詳細リストを生成
    v6.2.0: 固定割当が含まれるgapは除外（意図的な配置のため）
    """
    # 各(date, hosp)が固定割当かどうかを逆引きするため、slot_metaからhosp逆引き表を構築
    def _is_preassigned_assignment(date, hosp):
        """doc_assignmentsの(date,hosp)が固定割当かを判定"""
        for (ridx, h), (d, fixed) in slot_meta.items():
            if h == hosp and d == date and fixed:
                return True
        return False

    rows = []
    for doc, assigns in doc_assignments.items():
        assigns_sorted = sorted(assigns, key=lambda x: (x[0], x[1]))
        for i in range(1, len(assigns_sorted)):
            d_prev, h_prev = assigns_sorted[i - 1]
            d_cur, h_cur = assigns_sorted[i]
            gap = (d_cur - d_prev).days
            if gap < 3:
                # v6.2.0: 固定割当が片方でも含まれるgapは許容
                if _is_preassigned_assignment(d_prev, h_prev) or _is_preassigned_assignment(d_cur, h_cur):
                    continue
                rows.append({
                    "氏名": doc,
                    "前回日付": d_prev,
                    "前回病院": h_prev,
                    "今回日付": d_cur,
                    "今回病院": h_cur,
                    "間隔(日)": gap,
                })
    cols = ["氏名", "前回日付", "前回病院", "今回日付", "今回病院", "間隔(日)"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["氏名", "今回日付"]).reset_index(drop=True)

def build_same_day_duplicates(doc_assignments):
    rows = []
    for doc, assigns in doc_assignments.items():
        by_date = defaultdict(list)
        for d, h in assigns:
            by_date[d].append(h)
        for d, hs in by_date.items():
            if len(hs) > 1:
                rows.append({"氏名": doc, "日付": d, "件数": len(hs), "病院": ", ".join(sorted(hs))})
    cols = ["氏名", "日付", "件数", "病院"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["日付", "氏名"]).reset_index(drop=True)

def build_hosp_dup_details(assigned_hosp_count):
    rows = []
    for doc, hdict in assigned_hosp_count.items():
        for hosp, c in hdict.items():
            if c > 1:
                rows.append({"氏名": doc, "病院": hosp, "回数": c, "超過": c - 1})
    cols = ["氏名", "病院", "回数", "超過"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["超過", "氏名"], ascending=[False, True]).reset_index(drop=True)

def build_weekly_bg_details(doc_assignments):
    """v6.5.0: 大学系7日間隔違反の詳細リストを生成
    ABS-012改: 大学系は7日間隔必須
    """
    rows = []
    # 医師ごとのBG割当を日付順に収集
    bg_by_doc = defaultdict(list)  # doc -> [(date, hosp), ...]

    for doc, assigns in doc_assignments.items():
        for date, hosp in assigns:
            try:
                hidx = shift_df.columns.get_loc(hosp)
            except KeyError:
                continue
            if not (B_COL_INDEX <= hidx <= K_COL_INDEX):
                continue
            bg_by_doc[doc].append((date, hosp))

    # 7日間隔違反を検出
    for doc, assigns in bg_by_doc.items():
        assigns_sorted = sorted(assigns, key=lambda x: x[0])
        for i in range(1, len(assigns_sorted)):
            curr_date, curr_hosp = assigns_sorted[i]
            prev_date, prev_hosp = assigns_sorted[i-1]
            gap = abs((curr_date - prev_date).days)
            if gap < 7:
                # 固定割当をチェック
                fixed_count = 0
                for date, hosp in [assigns_sorted[i-1], assigns_sorted[i]]:
                    for (ridx, h), (d, fixed) in slot_meta.items():
                        if h == hosp and d == date and fixed:
                            fixed_count += 1
                            break
                # 両方固定割当なら許容
                if fixed_count == 2:
                    continue
                rows.append({
                    "氏名": doc,
                    "日付1": prev_date,
                    "病院1": prev_hosp,
                    "日付2": curr_date,
                    "病院2": curr_hosp,
                    "間隔(日)": gap,
                    "詳細": f"{prev_date.strftime('%m/%d')}({prev_hosp}) → {curr_date.strftime('%m/%d')}({curr_hosp})",
                })

    cols = ["氏名", "日付1", "病院1", "日付2", "病院2", "間隔(日)", "詳細"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["日付1", "氏名"]).reset_index(drop=True)

def build_unassigned_details(unassigned):
    rows = [{"日付": d, "病院": hosp, "row_index": ridx} for d, hosp, ridx in unassigned]
    cols = ["日付", "病院", "row_index"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["日付", "病院"]).reset_index(drop=True)

def build_doctor_diag(counts, bg_counts, ht_counts, wd_counts, we_counts, doc_assignments, assigned_hosp_count):
    rows = []
    active_set = set(active_doctors)

    for doc in doctor_names:
        assigns = sorted(doc_assignments.get(doc, []), key=lambda x: x[0])
        gaps = [(assigns[i][0] - assigns[i - 1][0]).days for i in range(1, len(assigns))]
        gap_viol = sum(1 for g in gaps if g < 3)
        min_gap = min(gaps) if gaps else None
        hosp_excess = sum(max(0, c - 1) for c in assigned_hosp_count.get(doc, {}).values())

        # gap3上限: この医師がgap>=3を守りつつ割当可能な最大回数
        gap3_max = compute_max_gap3_assignments(doc) if doc in active_set else 0
        avail_days = sum(1 for dt in all_shift_dates if get_avail_code(dt, doc) != 0)

        row = {
            "氏名": doc,
            "active": 1 if doc in active_set else 0,
            "cap": TARGET_CAP.get(doc, 0),
            "gap3上限": gap3_max,
            "利用可能日数": avail_days,
            "preassigned": preassigned_count.get(doc, 0),

            "今月_全合計": counts.get(doc, 0),
            "累計_全合計": prev_total.get(doc, 0) + counts.get(doc, 0),

            "今月_大学合計": bg_counts.get(doc, 0),
            "累計_大学合計": prev_bg.get(doc, 0) + bg_counts.get(doc, 0),

            "今月_外病院合計": ht_counts.get(doc, 0),
            "累計_外病院合計": prev_ht.get(doc, 0) + ht_counts.get(doc, 0),

            "今月_平日": wd_counts.get(doc, 0),
            "累計_平日": prev_weekday.get(doc, 0) + wd_counts.get(doc, 0),

            "今月_休日合計": we_counts.get(doc, 0),
            "累計_休日合計": prev_weekend.get(doc, 0) + we_counts.get(doc, 0),

            "gap違反回数": gap_viol,
            "最小間隔(日)": min_gap if min_gap is not None else "",
            "同一病院重複超過": hosp_excess,
        }
        rows.append(row)

    df = pd.DataFrame(rows)

    # 偏り（active平均との差）も出しておく
    active_rows = df[df["active"] == 1]
    for col in ["累計_全合計", "累計_大学合計", "累計_外病院合計", "累計_平日", "累計_休日合計"]:
        if len(active_rows) > 0:
            mean_val = float(active_rows[col].mean())
            df[col + "_平均との差"] = df[col] - mean_val
        else:
            df[col + "_平均との差"] = 0.0

    return df.sort_values(["active", "累計_全合計"], ascending=[False, False]).reset_index(drop=True)

def build_metrics_df(score_clamped, raw_score, metrics):
    """スコアサマリーを日本語で生成（旧メトリクス）"""
    rows = [
        {"項目": "総合スコア（0〜100）", "値": float(score_clamped), "説明": "制約違反のペナルティを100から引いた値（高いほど良い）"},
        {"項目": "ペナルティ合計", "値": float(metrics.get("penalty_total", 0)), "説明": "全制約違反のペナルティ合計（低いほど良い）"},
        {"項目": "--- 制約違反 ---", "値": "", "説明": ""},
        {"項目": "未割当枠", "値": int(metrics.get("unassigned_slots", 0)), "説明": "医師が割り当てられていないスロット数"},
        {"項目": "TARGET_CAP超過", "値": int(metrics.get("cap_violations", 0)), "説明": "割当上限を超えた回数"},
        {"項目": "gap違反（3日未満）", "値": int(metrics.get("gap_violations", 0)), "説明": "当直間隔が3日未満の件数"},
        {"項目": "同一病院重複（全体）", "値": int(metrics.get("hospital_dup_violations", 0)), "説明": "同じ病院に2回以上割当された件数"},
        {"項目": "外病院重複", "値": int(metrics.get("external_hosp_dup_violations", 0)), "説明": "外病院(L〜Y列)の同一病院重複"},
        {"項目": "CODE_2医師CAP超過", "値": int(metrics.get("code_2_extra_violations", 0)), "説明": "可否コード2医師のTARGET_CAP超過"},
        {"項目": "大学系3回以上", "値": int(metrics.get("bg_over_2_violations", 0)), "説明": "大学系(B〜K列)に3回以上割当"},
        {"項目": "外病院0回", "値": int(metrics.get("ht_0_violations", 0)), "説明": "外病院に1回も割当されていない医師数"},
        {"項目": "C-Hカテ当番違反", "値": int(metrics.get("ch_kate_violations", 0)), "説明": "C〜H列にカテ当番日以外で割当"},
        {"項目": "--- 公平性 ---", "値": "", "説明": ""},
        {"項目": "割当回数の偏り（max-min）", "値": int(metrics.get("max_minus_min_total_active", 0)), "説明": "active医師間の最大-最小割当回数差"},
        {"項目": "公平性ペナルティ", "値": int(metrics.get("fairness_penalty", 0)), "説明": "偏りが2以上で発生するペナルティ"},
        {"項目": "--- 偏り（累計spread） ---", "値": "", "説明": ""},
        {"項目": "累計全合計spread", "値": float(metrics.get("total_spread_cum", 0)), "説明": "累計全合計回数のmax-min（前月+今月、v6.5.9）"},
        {"項目": "大学系spread", "値": float(metrics.get("bg_spread_cum", 0)), "説明": "累計大学回数のmax-min（前月+今月）"},
        {"項目": "外病院spread", "値": float(metrics.get("ht_spread_cum", 0)), "説明": "累計外病院回数のmax-min"},
        {"項目": "平日spread", "値": float(metrics.get("weekday_spread_cum", 0)), "説明": "累計平日回数のmax-min"},
        {"項目": "休日spread", "値": float(metrics.get("weekend_spread_cum", 0)), "説明": "累計休日回数のmax-min"},
    ]
    return pd.DataFrame(rows)

def build_hard_constraint_violations(pattern_df):
    """ハード制約違反の詳細リストを生成
    v6.2.0: 固定割当（事前割当）スロットは検査対象外（意図的な配置のため）
    """
    rows = []

    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        dow = date.weekday()

        for hosp in hospital_cols:
            # v6.2.0: 固定割当は意図的な配置のため、違反検出をスキップ
            if is_preassigned_slot(ridx, hosp):
                continue

            val = pattern_df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc not in doctor_names:
                continue

            idx = shift_df.columns.get_loc(hosp)
            code = get_avail_code(date, doc)
            sched_code = get_sched_code(date, doc)

            # 違反1: 可否コード0 (ABS-001)
            if code == 0:
                rows.append({
                    "制約ID": CONSTRAINT_ABS_001,
                    "違反種別": "可否コード0違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code if sched_code else "",
                    "詳細": f"[{CONSTRAINT_ABS_001}] コード0（不可）の日に割当",
                })

            # 違反2: 可否コード2違反（Q列より後に割当）(ABS-002)
            if code == 2 and not (B_COL_INDEX <= idx <= Q_COL_INDEX):
                rows.append({
                    "制約ID": CONSTRAINT_ABS_002,
                    "違反種別": "可否コード2違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code if sched_code else "",
                    "詳細": f"[{CONSTRAINT_ABS_002}] コード2はB〜Q列のみ可。列{idx}に割当",
                })

            # 違反3: 可否コード3違反（L〜Y列以外に割当）(ABS-003)
            if code == 3 and not (L_COL_INDEX <= idx <= L_Y_END_INDEX):
                rows.append({
                    "制約ID": CONSTRAINT_ABS_003,
                    "違反種別": "可否コード3違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code if sched_code else "",
                    "詳細": f"[{CONSTRAINT_ABS_003}] コード3はL〜Y列のみ可。列{idx}に割当",
                })

            # 違反4: カテ表コードあり＋L〜Y列違反 (ABS-004)
            if L_COL_INDEX <= idx <= L_Y_END_INDEX and sched_code:
                rows.append({
                    "制約ID": CONSTRAINT_ABS_004,
                    "違反種別": "カテ表+外病院違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code,
                    "詳細": f"[{CONSTRAINT_ABS_004}] カテ表（{sched_code}）がある日は外病院（L〜Y列）に割当不可。列{idx}に割当",
                })

            # 違反5: SEMI-001（B列のみ）/ ABS-013（C-H列）カテ表コードなし
            # ※I-K列は制約対象外（平日緩和のため）
            # v6.5.6: 属性2は緩和不可（ABS該当）、属性1は緩和可（SEMI）
            if doc in SCHEDULE_CODE_HOLDERS and not sched_code and doc not in EXTRA_ALLOWED:
                # SEMI-001: B列のみカテ表コード必須
                if idx == B_COL_INDEX:
                    doc_attr = doctor_attribute.get(doc, "")
                    if doc_attr == "2":
                        # 属性2は緩和不可（ABS-015）
                        rows.append({
                            "制約ID": CONSTRAINT_ABS_015,
                            "違反種別": "B列カテ表コード欠如（属性2:緩和不可）",
                            "日付": date,
                            "医師名": doc,
                            "病院": hosp,
                            "列番号": idx,
                            "可否コード": code,
                            "カテ表": "",
                            "詳細": f"[{CONSTRAINT_ABS_015}] B列割当にカテ表必須（属性2は緩和不可）",
                        })
                    else:
                        # 属性1または未設定は緩和可（週1回まで許容）
                        rows.append({
                            "制約ID": CONSTRAINT_SEMI_001,
                            "違反種別": "B列カテ表コード欠如",
                            "日付": date,
                            "医師名": doc,
                            "病院": hosp,
                            "列番号": idx,
                            "可否コード": code,
                            "カテ表": "",
                            "詳細": f"[{CONSTRAINT_SEMI_001}] B列（平日大学系）の割当にカテ表コードが必要",
                        })
                # ABS-013: C-H列カテ当番必須（v6.5.3でSEMI-002から格上げ）
                elif C_COL_INDEX <= idx <= H_COL_INDEX:
                    rows.append({
                        "制約ID": CONSTRAINT_ABS_013,
                        "違反種別": "C-H列カテ当番欠如",
                        "日付": date,
                        "医師名": doc,
                        "病院": hosp,
                        "列番号": idx,
                        "可否コード": code,
                        "カテ表": "",
                        "詳細": f"[{CONSTRAINT_ABS_013}] C-H列（休日大学系）の割当にはカテ当番が必須",
                    })

            # 違反6: 水曜日L〜Y列禁止医師 (ABS-006)
            if dow == 2 and L_COL_INDEX <= idx <= L_Y_END_INDEX and doc in WED_FORBIDDEN_DOCTORS:
                rows.append({
                    "制約ID": CONSTRAINT_ABS_006,
                    "違反種別": "水曜日L〜Y列禁止違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code if sched_code else "",
                    "詳細": f"[{CONSTRAINT_ABS_006}] {doc}は水曜日のL〜Y列禁止",
                })

    # B〜H列の2回超過違反をチェック
    bh_counts = defaultdict(list)
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)

        for hosp in hospital_cols:
            val = pattern_df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc not in doctor_names:
                continue

            idx = shift_df.columns.get_loc(hosp)
            if B_H_START_INDEX <= idx <= B_H_END_INDEX:
                bh_counts[doc].append((date, hosp, idx))

    # 違反7: B〜H列が2回超過 (SOFT-002: 大学3回以上)
    for doc, assignments in bh_counts.items():
        if len(assignments) > 2:
            for date, hosp, idx in assignments[2:]:  # 3回目以降
                code = get_avail_code(date, doc)
                sched_code = get_sched_code(date, doc)
                rows.append({
                    "制約ID": CONSTRAINT_SOFT_002,
                    "違反種別": "B-H列2回超過違反",
                    "日付": date,
                    "医師名": doc,
                    "病院": hosp,
                    "列番号": idx,
                    "可否コード": code,
                    "カテ表": sched_code if sched_code else "",
                    "詳細": f"[{CONSTRAINT_SOFT_002}] B〜H列は2回まで。{len(assignments)}回目の割当",
                })

    cols = ["制約ID", "違反種別", "日付", "医師名", "病院", "列番号", "可否コード", "カテ表", "詳細"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols].sort_values(["制約ID", "日付", "医師名"]).reset_index(drop=True)

def fix_hard_constraint_violations(pattern_df, max_attempts=50, verbose=True):
    """
    ハード制約違反を自動修正する

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数, 修正失敗数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    total_failed = 0

    for attempt in range(max_attempts):
        violations_df = build_hard_constraint_violations(df)

        if len(violations_df) == 0:
            if verbose and total_fixed > 0:
                print(f"   ✅ ハード制約違反を{total_fixed}件修正しました")
            return df, True, total_fixed, total_failed

        if attempt == 0 and verbose:
            print(f"   ⚠️ ハード制約違反を{len(violations_df)}件検出 → 自動修正を開始...")

        # 各違反を修正試行
        fixed_in_this_iteration = 0

        for _, violation in violations_df.iterrows():
            date = violation['日付']
            doc = violation['医師名']
            hosp = violation['病院']
            violation_type = violation['違反種別']

            # 該当行を探す
            ridx = None
            for idx in df.index:
                if pd.to_datetime(df.at[idx, date_col_shift]).normalize().tz_localize(None) == date:
                    ridx = idx
                    break

            if ridx is None:
                continue

            # 違反している割当を解除
            current_val = df.at[ridx, hosp]
            if not isinstance(current_val, str) or normalize_name(current_val) != doc:
                continue

            df.at[ridx, hosp] = None

            # 代替医師を探す
            col_idx = shift_df.columns.get_loc(hosp)
            dow = pd.to_datetime(date).weekday()

            # この日に既に割り当てられている医師を除外
            already_assigned_on_date = set()
            for h in hospital_cols:
                v = df.at[ridx, h]
                if isinstance(v, str):
                    already_assigned_on_date.add(normalize_name(v))

            # 候補医師を探す（ハード制約のみチェック）
            candidates = []
            for candidate_doc in doctor_names:
                # 同日重複チェック
                if candidate_doc in already_assigned_on_date:
                    continue

                # ハード制約チェック
                if can_assign_doc_to_slot(candidate_doc, date, hosp):
                    candidates.append(candidate_doc)

            if candidates:
                # 優先順位：全体合計が少ない医師を優先
                candidates.sort(key=lambda d: prev_total.get(d, 0) + len([1 for h in hospital_cols for ridx2 in df.index if isinstance(df.at[ridx2, h], str) and normalize_name(df.at[ridx2, h]) == d]))
                new_doc = candidates[0]
                df.at[ridx, hosp] = new_doc
                fixed_in_this_iteration += 1
                total_fixed += 1
            else:
                # 緊急フォールバック: 絶対禁忌をすべて回避
                # 医師の現在の割当日を取得（gap1チェック用）
                def get_doc_dates(d):
                    dates = []
                    for ridx2 in df.index:
                        dt2 = df.at[ridx2, date_col_shift]
                        if pd.isna(dt2):
                            continue
                        dt2 = pd.to_datetime(dt2).normalize().tz_localize(None)
                        for h in hospital_cols:
                            v = df.at[ridx2, h]
                            if isinstance(v, str) and normalize_name(v) == d:
                                dates.append(dt2)
                    return dates

                # 医師の外病院割当回数を取得
                def get_doc_hosp_count(d, target_hosp):
                    count = 0
                    for ridx2 in df.index:
                        v = df.at[ridx2, target_hosp]
                        if isinstance(v, str) and normalize_name(v) == d:
                            count += 1
                    return count

                is_external = L_COL_INDEX <= col_idx <= L_Y_END_INDEX
                day_of_week = pd.to_datetime(date).weekday()

                def is_valid_emergency(d):
                    # 同日重複禁止 (ABS-006)
                    if d in already_assigned_on_date:
                        return False
                    code = get_avail_code(date, d)
                    # ABS-001: コード0禁止
                    if code == 0:
                        return False
                    # ABS-002: コード2はB〜Q列のみ
                    if code == 2 and not (B_COL_INDEX <= col_idx <= Q_COL_INDEX):
                        return False
                    # ABS-003: コード3はL〜Y列のみ
                    if code == 3 and not is_external:
                        return False
                    # ABS-004: カテ表コードありの日はL〜Y列不可
                    if is_external and get_sched_code(date, d):
                        return False
                    # ABS-005: 水曜日L〜Y列禁止医師
                    if day_of_week == 2 and is_external and d in WED_FORBIDDEN_DOCTORS:
                        return False
                    # ABS-007: gap >= 3日必須
                    doc_dates = get_doc_dates(d)
                    if doc_dates:
                        min_gap = min(abs((date - dt).days) for dt in doc_dates)
                        if min_gap < 3:
                            return False
                    # ABS-008: 外病院重複禁止
                    if is_external and get_doc_hosp_count(d, hosp) >= 1:
                        return False
                    # ABS-010/011: TARGET_CAPとBG上限は静的チェック困難なためスキップ
                    return True

                emergency_candidates = [d for d in doctor_names if is_valid_emergency(d)]
                if emergency_candidates:
                    # 全体合計が最も少ない医師を選択
                    emergency_candidates.sort(key=lambda d: prev_total.get(d, 0) + len([1 for h in hospital_cols for ridx2 in df.index if isinstance(df.at[ridx2, h], str) and normalize_name(df.at[ridx2, h]) == d]))
                    new_doc = emergency_candidates[0]
                    df.at[ridx, hosp] = new_doc
                    fixed_in_this_iteration += 1
                    total_fixed += 1
                    if verbose:
                        print(f"   ⚠️ 緊急フォールバック: {date.strftime('%Y-%m-%d')} {hosp} → {new_doc}")
                else:
                    # 絶対禁忌を満たす候補がいない場合は未割当のまま
                    total_failed += 1
                    if verbose:
                        print(f"   ❌ 修正不可（絶対禁忌回避不可）: {date.strftime('%Y-%m-%d')} {hosp}")

        # 進捗がなければループ終了
        if fixed_in_this_iteration == 0:
            break

    # 最終チェック
    final_violations = build_hard_constraint_violations(df)
    success = len(final_violations) == 0

    if verbose:
        if success:
            print(f"   ✅ 全てのハード制約違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {len(final_violations)}件のハード制約違反が残っています（修正数: {total_fixed}, 失敗: {total_failed}）")

    return df, success, total_fixed, total_failed

def fix_target_cap_violations(pattern_df, max_attempts=100, verbose=True):
    """
    TARGET_CAP違反を修正する（上限超過とBASE_TARGET未達の両方）

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        # 現在の割当回数を再計算
        counts, *_ = recompute_stats(df)

        # cap超過・BASE_TARGET未達の医師を特定
        over_cap_docs = []  # cap超過
        under_base_docs = []  # BASE_TARGET未達（新規）
        at_base_docs = []  # BASE_TARGETちょうど
        over_base_docs = []  # BASE_TARGET超過だがcap以下

        for doc in active_doctors:
            current = counts.get(doc, 0)
            cap = TARGET_CAP.get(doc, 0)

            if current > cap:
                over_cap_docs.append((doc, current - cap))
            elif current < BASE_TARGET:
                under_base_docs.append((doc, BASE_TARGET - current))
            elif current == BASE_TARGET:
                at_base_docs.append(doc)
            else:  # BASE_TARGET < current <= cap
                over_base_docs.append((doc, current - BASE_TARGET))

        # 違反がなければ終了
        if not over_cap_docs and not under_base_docs:
            if verbose and total_fixed > 0:
                print(f"   ✅ TARGET_CAP違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            if over_cap_docs:
                print(f"   ⚠️ TARGET_CAP超過を{len(over_cap_docs)}件検出 → 自動修正を開始...")
            if under_base_docs:
                print(f"   ⚠️ BASE_TARGET未達を{len(under_base_docs)}件検出 → 自動修正を開始...")

        # 修正試行
        fixed_in_this_iteration = 0

        # 1. cap超過の修正（優先）
        for over_doc, excess in over_cap_docs:
            if excess <= 0:
                continue

            over_doc_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == over_doc:
                        # v6.2.0: 固定割当は移動対象外
                        if is_preassigned_slot(ridx, hosp):
                            continue
                        over_doc_positions.append((ridx, hosp, date))

            random.shuffle(over_doc_positions)

            for ridx, hosp, date in over_doc_positions[:min(excess, 3)]:
                already_assigned_on_date = set()
                for h in hospital_cols:
                    v = df.at[ridx, h]
                    if isinstance(v, str):
                        already_assigned_on_date.add(normalize_name(v))

                # BASE_TARGET未達の医師を優先、次にat_base
                candidates = []
                for under_doc, deficit in under_base_docs:
                    if deficit <= 0:
                        continue
                    if under_doc in already_assigned_on_date:
                        continue
                    if can_assign_doc_to_slot(under_doc, date, hosp):
                        candidates.append((under_doc, 0))  # 優先度0（最優先）

                for at_doc in at_base_docs:
                    if at_doc in already_assigned_on_date:
                        continue
                    if can_assign_doc_to_slot(at_doc, date, hosp):
                        if TARGET_CAP.get(at_doc, 0) > BASE_TARGET:
                            candidates.append((at_doc, 1))  # 優先度1

                if candidates:
                    candidates.sort(key=lambda x: (x[1], prev_total.get(x[0], 0) + counts.get(x[0], 0)))
                    new_doc = candidates[0][0]

                    df.at[ridx, hosp] = new_doc
                    fixed_in_this_iteration += 1
                    total_fixed += 1

                    # リストを更新
                    for i, (d, deficit) in enumerate(under_base_docs):
                        if d == new_doc:
                            under_base_docs[i] = (d, deficit - 1)
                            if deficit - 1 <= 0 and d in at_base_docs:
                                pass
                            elif deficit - 1 <= 0:
                                at_base_docs.append(d)
                            break
                    if new_doc in at_base_docs:
                        over_base_docs.append((new_doc, 1))
                    break

        # 2. BASE_TARGET未達の修正
        for under_doc, deficit in under_base_docs:
            if deficit <= 0:
                continue

            # over_base（BASE_TARGET超過だがcap以下）の医師から取る
            donor_candidates = over_base_docs[:3]

            for donor_doc, _ in donor_candidates:
                donor_positions = []
                for ridx in df.index:
                    date = df.at[ridx, date_col_shift]
                    if pd.isna(date):
                        continue
                    date = pd.to_datetime(date).normalize().tz_localize(None)

                    for hosp in hospital_cols:
                        val = df.at[ridx, hosp]
                        if isinstance(val, str) and normalize_name(val) == donor_doc:
                            donor_positions.append((ridx, hosp, date))

                random.shuffle(donor_positions)

                for ridx, hosp, date in donor_positions[:2]:
                    already_assigned_on_date = set()
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str):
                            already_assigned_on_date.add(normalize_name(v))

                    if under_doc in already_assigned_on_date:
                        continue
                    if not can_assign_doc_to_slot(under_doc, date, hosp):
                        continue

                    df.at[ridx, hosp] = under_doc
                    fixed_in_this_iteration += 1
                    total_fixed += 1
                    break

                if fixed_in_this_iteration > 0:
                    break

    # 最終確認
    counts, *_ = recompute_stats(df)
    remaining_over = sum(1 for doc in active_doctors if counts.get(doc, 0) > TARGET_CAP.get(doc, 0))
    remaining_under = sum(1 for doc in active_doctors if counts.get(doc, 0) < BASE_TARGET)

    if verbose:
        if remaining_over == 0 and remaining_under == 0:
            print(f"   ✅ 全てのTARGET_CAP違反を修正しました（修正数: {total_fixed}）")
        else:
            if remaining_over > 0:
                print(f"   ⚠️ {remaining_over}件のTARGET_CAP超過違反が残っています（修正数: {total_fixed}）")
            if remaining_under > 0:
                print(f"   ⚠️ {remaining_under}件のBASE_TARGET未達違反が残っています（修正数: {total_fixed}）")

    return df, (remaining_over == 0 and remaining_under == 0), total_fixed

def fix_code_2_extra_violations(pattern_df, max_attempts=100, verbose=True):
    """
    可否コード2医師のTARGET_CAP超過違反を修正する
    v6.0.5: CODE_2もEXTRA対象のため、TARGET_CAPベースで判定

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        # 現在の割当回数を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # CODE_2医師でTARGET_CAPを超えている医師を特定
        code_2_over_docs = []
        for doc in CODE_2_DOCTORS:
            if doc not in active_doctors:
                continue
            current = counts.get(doc, 0)
            cap = TARGET_CAP.get(doc, 0)
            if current > cap:
                code_2_over_docs.append((doc, current - cap))

        # 違反がなければ終了
        if not code_2_over_docs:
            if verbose and total_fixed > 0:
                print(f"   ✅ 可否コード2医師のn+1違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            over_docs_str = ", ".join([f"{d}({excess}回超過)" for d, excess in code_2_over_docs])
            print(f"   ⚠️ 可否コード2医師のn+1違反を{len(code_2_over_docs)}件検出 → 自動修正を開始...")
            print(f"      対象: {over_docs_str}")

        # 修正試行
        fixed_in_this_iteration = 0

        for over_doc, excess in code_2_over_docs:
            if excess <= 0:
                continue

            # over_docの割当位置を取得
            over_doc_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == over_doc:
                        # v6.2.0: 固定割当は移動対象外
                        if is_preassigned_slot(ridx, hosp):
                            continue
                        over_doc_positions.append((ridx, hosp, date))

            random.shuffle(over_doc_positions)

            for ridx, hosp, date in over_doc_positions[:min(excess, 3)]:
                # その日に既に割当されている医師
                already_assigned_on_date = set()
                for h in hospital_cols:
                    v = df.at[ridx, h]
                    if isinstance(v, str):
                        already_assigned_on_date.add(normalize_name(v))

                # 代替医師を探す（CODE_2以外の医師でBASE_TARGET未達または余裕のある医師）
                candidates = []
                for alt_doc in active_doctors:
                    if alt_doc == over_doc:
                        continue
                    if alt_doc in already_assigned_on_date:
                        continue
                    if alt_doc in CODE_2_DOCTORS:
                        continue  # CODE_2医師は代替にならない

                    alt_current = counts.get(alt_doc, 0)
                    alt_cap = TARGET_CAP.get(alt_doc, 0)

                    if alt_current >= alt_cap:
                        continue  # 既にcap到達

                    if can_assign_doc_to_slot(alt_doc, date, hosp):
                        # 優先度: BASE_TARGET未達 > ちょうど > 超過
                        priority = 0 if alt_current < BASE_TARGET else (1 if alt_current == BASE_TARGET else 2)
                        candidates.append((alt_doc, priority))

                if candidates:
                    candidates.sort(key=lambda x: x[1])
                    new_doc = candidates[0][0]
                    df.at[ridx, hosp] = new_doc
                    counts[over_doc] = counts.get(over_doc, 0) - 1
                    counts[new_doc] = counts.get(new_doc, 0) + 1
                    total_fixed += 1
                    fixed_in_this_iteration += 1
                else:
                    # 緊急フォールバック: 絶対禁忌をすべて回避
                    col_idx = shift_df.columns.get_loc(hosp)
                    is_external = L_COL_INDEX <= col_idx <= L_Y_END_INDEX
                    day_of_week = pd.to_datetime(date).weekday()

                    def is_valid_emergency_target(d):
                        # 同日重複禁止 (ABS-006)
                        if d in already_assigned_on_date:
                            return False
                        if d == over_doc:
                            return False
                        code = get_avail_code(date, d)
                        # ABS-001: コード0禁止
                        if code == 0:
                            return False
                        # ABS-002: コード2はB〜Q列のみ
                        if code == 2 and not (B_COL_INDEX <= col_idx <= Q_COL_INDEX):
                            return False
                        # ABS-003: コード3はL〜Y列のみ
                        if code == 3 and not is_external:
                            return False
                        # ABS-004: カテ表コードありの日はL〜Y列不可
                        if is_external and get_sched_code(date, d):
                            return False
                        # ABS-005: 水曜日L〜Y列禁止医師
                        if day_of_week == 2 and is_external and d in WED_FORBIDDEN_DOCTORS:
                            return False
                        # ABS-007: gap >= 3日必須
                        doc_dates = sorted([dt for dt, _ in doc_assignments.get(d, [])])
                        if doc_dates:
                            min_gap = min(abs((date - dt).days) for dt in doc_dates)
                            if min_gap < 3:
                                return False
                        # ABS-008: 外病院重複禁止
                        if is_external and assigned_hosp_count.get(d, {}).get(hosp, 0) >= 1:
                            return False
                        return True

                    emergency = [d for d in doctor_names if is_valid_emergency_target(d)]
                    if emergency:
                        emergency.sort(key=lambda d: counts.get(d, 0))
                        new_doc = emergency[0]
                        df.at[ridx, hosp] = new_doc
                        counts[over_doc] = counts.get(over_doc, 0) - 1
                        counts[new_doc] = counts.get(new_doc, 0) + 1
                        total_fixed += 1
                        fixed_in_this_iteration += 1
                        if verbose:
                            print(f"      ⚠️ {over_doc}→{new_doc}(緊急): {date.strftime('%m/%d')} {hosp}")
                    else:
                        # 最終手段: 元の医師を維持（削除しない）
                        if verbose:
                            print(f"      ⚠️ {over_doc}の{date.strftime('%m/%d')} {hosp}を維持（絶対禁忌回避不可）")

        if fixed_in_this_iteration == 0:
            break

    # 最終状態を確認
    counts, *_ = recompute_stats(df)
    remaining = sum(1 for doc in CODE_2_DOCTORS if doc in active_doctors and counts.get(doc, 0) > TARGET_CAP.get(doc, 0))

    if verbose:
        if remaining == 0:
            if total_fixed > 0:
                print(f"   ✅ 全ての可否コード2医師のTARGET_CAP違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining}件の可否コード2医師のTARGET_CAP違反が残っています（修正数: {total_fixed}）")

    return df, (remaining == 0), total_fixed

def fix_university_minimum_requirement(pattern_df, max_attempts=100, verbose=True):
    """
    大学系最低1回必須違反を修正する（準ハード制約：コード3以外の全医師）

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    if not UNIVERSITY_MINIMUM_REQUIRED_DOCTORS:
        return pattern_df, True, 0

    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        # 現在の割当回数を再計算
        counts, bg_counts, *_ = recompute_stats(df)

        # 大学系0回の医師を特定（コード3除外）
        zero_bg_docs = []
        for doc in UNIVERSITY_MINIMUM_REQUIRED_DOCTORS:
            if bg_counts.get(doc, 0) == 0:
                zero_bg_docs.append(doc)

        if not zero_bg_docs:
            if verbose and total_fixed > 0:
                print(f"   ✅ 大学系最低1回必須違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 大学系最低1回必須違反を{len(zero_bg_docs)}件検出 → 自動修正を開始...")

        # 修正試行
        fixed_in_this_iteration = 0

        for zero_doc in zero_bg_docs:
            # このzero_docの外病院（L〜Y）割当を探す
            zero_doc_ly_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    idx = shift_df.columns.get_loc(hosp)
                    # L〜Y列（外病院）のみ
                    if L_COL_INDEX <= idx <= L_Y_END_INDEX:
                        val = df.at[ridx, hosp]
                        if isinstance(val, str) and normalize_name(val) == zero_doc:
                            # v6.2.0: 固定割当は移動対象外
                            if is_preassigned_slot(ridx, hosp):
                                continue
                            zero_doc_ly_positions.append((ridx, hosp, date))

            # 外病院の割当を1つ大学系に変更
            random.shuffle(zero_doc_ly_positions)

            for ridx, hosp, date in zero_doc_ly_positions[:1]:  # 1つだけ試行
                # この日付のB〜K列（大学系）で空いている枠を探す
                for bg_hosp in hospital_cols:
                    bg_idx = shift_df.columns.get_loc(bg_hosp)
                    if not (B_COL_INDEX <= bg_idx <= B_K_END_INDEX):
                        continue

                    val = df.at[ridx, bg_hosp]
                    # 空き枠かどうか
                    if not is_slot_value(shift_df.at[ridx, bg_hosp]):
                        continue
                    if isinstance(val, str) and val in doctor_names:
                        continue  # 既に割当済み

                    # この日にzero_docが既に割り当てられていないかチェック
                    already_assigned = False
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str) and normalize_name(v) == zero_doc and h != hosp:
                            already_assigned = True
                            break

                    if already_assigned:
                        continue

                    # 制約チェック
                    if can_assign_doc_to_slot(zero_doc, date, bg_hosp):
                        # 外病院から削除、大学系に追加
                        df.at[ridx, hosp] = None  # 外病院を解除
                        df.at[ridx, bg_hosp] = zero_doc  # 大学系に割当
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        break  # 次のzero_docへ

                if fixed_in_this_iteration > 0:
                    break  # 次のzero_docへ

        # 進捗がなければループ終了
        if fixed_in_this_iteration == 0:
            break

    # 最終確認
    counts, bg_counts, *_ = recompute_stats(df)
    remaining_violations = sum(1 for doc in UNIVERSITY_MINIMUM_REQUIRED_DOCTORS if bg_counts.get(doc, 0) == 0)

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全ての大学系最低1回必須違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件の大学系最低1回必須違反が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_ch_kate_violations(pattern_df, max_attempts=100, verbose=True):
    """
    C-H列（休日大学系）のカテ当番違反を修正する

    条件：C-H列はカテ当番ありの日 OR カテ当番なし医師のみ
    カテ当番保有医師がその日にカテ当番なしでC-H列に割り当てられている場合、
    適格な医師と交換する

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        # C-H列の違反を検出
        violations = []
        for ridx in df.index:
            date = df.at[ridx, date_col_shift]
            if pd.isna(date):
                continue
            date = pd.to_datetime(date).normalize().tz_localize(None)
            for hosp in hospital_cols:
                idx = shift_df.columns.get_loc(hosp)
                if not is_ch_slot(idx):
                    continue
                val = df.at[ridx, hosp]
                if not isinstance(val, str):
                    continue
                doc = normalize_name(val)
                if doc not in doctor_names:
                    continue
                if not is_eligible_for_ch_slot(doc, date):
                    violations.append({
                        "ridx": ridx,
                        "hosp": hosp,
                        "date": date,
                        "doc": doc,
                        "idx": idx
                    })

        if not violations:
            break

        # 最初の違反を修正
        viol = violations[0]
        ridx, hosp, date, bad_doc, col_idx = viol["ridx"], viol["hosp"], viol["date"], viol["doc"], viol["idx"]

        # 1. C-H列に適格な医師を探す
        # 適格条件: is_eligible_for_ch_slot(doc, date) = True
        # 候補：現在違反がないスロットにいる適格医師
        swap_done = False

        # 同じ日の他のスロット（C-H以外）で適格な医師を探して交換
        for other_hosp in hospital_cols:
            if swap_done:
                break
            other_idx = shift_df.columns.get_loc(other_hosp)
            # C-H列以外のスロットを探す（L-Y列など）
            if is_ch_slot(other_idx):
                continue
            other_val = df.at[ridx, other_hosp]
            if not isinstance(other_val, str):
                continue
            other_doc = normalize_name(other_val)
            if other_doc not in doctor_names:
                continue
            # other_docがC-H列に適格かチェック
            if not is_eligible_for_ch_slot(other_doc, date):
                continue
            # other_docがhosp（C-H列）に割り当て可能かチェック（ABS-001含む）
            if not can_assign_doc_to_slot(other_doc, date, hosp):
                continue
            # bad_docがother_hospに割り当て可能かチェック
            if not can_assign_doc_to_slot(bad_doc, date, other_hosp):
                continue
            # 交換
            df.at[ridx, hosp] = other_doc
            df.at[ridx, other_hosp] = bad_doc
            total_fixed += 1
            swap_done = True
            if verbose:
                print(f"   [C-Hカテ当番修正] {date.strftime('%Y-%m-%d')} {hosp}列: {bad_doc} ⇔ {other_doc}")

        if swap_done:
            continue

        # 2. 別の日の適格医師と交換を試みる
        for other_ridx in df.index:
            if swap_done:
                break
            other_date = df.at[other_ridx, date_col_shift]
            if pd.isna(other_date):
                continue
            other_date = pd.to_datetime(other_date).normalize().tz_localize(None)
            if other_date == date:
                continue

            for other_hosp in hospital_cols:
                if swap_done:
                    break
                other_idx = shift_df.columns.get_loc(other_hosp)
                # C-H列以外のスロット
                if is_ch_slot(other_idx):
                    continue
                other_val = df.at[other_ridx, other_hosp]
                if not isinstance(other_val, str):
                    continue
                other_doc = normalize_name(other_val)
                if other_doc not in doctor_names:
                    continue
                # other_docがC-H列に適格かチェック
                if not is_eligible_for_ch_slot(other_doc, date):
                    continue
                # bad_docがother_hospに割り当て可能かチェック
                if not can_assign_doc_to_slot(bad_doc, other_date, other_hosp):
                    continue
                # other_docがdate, hospに割り当て可能かチェック
                if not can_assign_doc_to_slot(other_doc, date, hosp):
                    continue
                # ABS-006: 同日重複チェック（bad_docがother_dateに既に割当されていないか）
                bad_doc_on_other_date = False
                other_doc_on_date = False
                for h_check in hospital_cols:
                    v_check = df.at[other_ridx, h_check]
                    if isinstance(v_check, str) and normalize_name(v_check) == bad_doc:
                        bad_doc_on_other_date = True
                        break
                if bad_doc_on_other_date:
                    continue
                for ridx_check in df.index:
                    d_check = df.at[ridx_check, date_col_shift]
                    if pd.isna(d_check):
                        continue
                    if pd.to_datetime(d_check).normalize().tz_localize(None) != date:
                        continue
                    for h_check in hospital_cols:
                        v_check = df.at[ridx_check, h_check]
                        if isinstance(v_check, str) and normalize_name(v_check) == other_doc and h_check != hosp:
                            other_doc_on_date = True
                            break
                    if other_doc_on_date:
                        break
                if other_doc_on_date:
                    continue
                # 交換
                df.at[ridx, hosp] = other_doc
                df.at[other_ridx, other_hosp] = bad_doc
                total_fixed += 1
                swap_done = True
                if verbose:
                    print(f"   [C-Hカテ当番修正] {date.strftime('%Y-%m-%d')} {hosp}列: {bad_doc} → {other_doc}")

        if not swap_done:
            # 交換できなかった場合、次の違反を試す
            if verbose and attempt == 0:
                print(f"   ⚠️ C-H列カテ当番違反を修正できません: {date.strftime('%Y-%m-%d')} {hosp}列 {bad_doc}")
            break

    # 残り違反を再計算
    remaining_violations = 0
    for ridx in df.index:
        date = df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize().tz_localize(None)
        for hosp in hospital_cols:
            idx = shift_df.columns.get_loc(hosp)
            if not is_ch_slot(idx):
                continue
            val = df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc not in doctor_names:
                continue
            if not is_eligible_for_ch_slot(doc, date):
                remaining_violations += 1

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全てのC-H列カテ当番違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件のC-H列カテ当番違反が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_bg_ht_imbalance_violations(pattern_df, max_attempts=100, verbose=True):
    """
    大学系と外病院の差が3以上の違反を修正する

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        # 現在の割当回数を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        # 大学系と外病院の差が3以上の医師を特定（CC除外）
        imbalance_docs = []
        for doc in active_doctors:
            # CCは大型連休特別シフトなのでバランス計算から除外
            bg = bg_counts.get(doc, 0) - cc_bg_counts.get(doc, 0)
            ht = ht_counts.get(doc, 0) - cc_ht_counts.get(doc, 0)
            diff = abs(bg - ht)
            if diff >= 3:
                imbalance_docs.append((doc, bg, ht, diff))

        if not imbalance_docs:
            if verbose and total_fixed > 0:
                print(f"   ✅ 大学系と外病院の差3以上の違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 大学系と外病院の差3以上の違反を{len(imbalance_docs)}件検出 → 自動修正を開始...")

        # 修正試行
        fixed_in_this_iteration = 0

        for doc, bg, ht, diff in imbalance_docs:
            if diff < 3:
                continue

            # BGが多い場合: BG→HTに移動
            # HTが多い場合: HT→BGに移動
            if bg > ht:
                # BGの割当を1つHTに変更
                source_range = (B_COL_INDEX, B_K_END_INDEX)
                target_range = (L_COL_INDEX, L_Y_END_INDEX)
            else:
                # HTの割当を1つBGに変更
                source_range = (L_COL_INDEX, L_Y_END_INDEX)
                target_range = (B_COL_INDEX, B_K_END_INDEX)

            # source範囲の割当を探す
            source_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    idx = shift_df.columns.get_loc(hosp)
                    if source_range[0] <= idx <= source_range[1]:
                        val = df.at[ridx, hosp]
                        if isinstance(val, str) and normalize_name(val) == doc:
                            # v6.2.0: 固定割当は移動対象外
                            if is_preassigned_slot(ridx, hosp):
                                continue
                            source_positions.append((ridx, hosp, date))

            # 1つ移動を試みる
            random.shuffle(source_positions)

            for ridx, hosp, date in source_positions[:1]:
                # target範囲で空いている枠を探す
                for target_hosp in hospital_cols:
                    target_idx = shift_df.columns.get_loc(target_hosp)
                    if not (target_range[0] <= target_idx <= target_range[1]):
                        continue

                    val = df.at[ridx, target_hosp]
                    # 空き枠かどうか
                    if not is_slot_value(shift_df.at[ridx, target_hosp]):
                        continue
                    if isinstance(val, str) and val in doctor_names:
                        continue  # 既に割当済み

                    # この日にdocが既に割り当てられていないかチェック
                    already_assigned = False
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str) and normalize_name(v) == doc and h != hosp:
                            already_assigned = True
                            break

                    if already_assigned:
                        continue

                    # 制約チェック
                    if can_assign_doc_to_slot(doc, date, target_hosp):
                        # sourceから削除、targetに追加
                        df.at[ridx, hosp] = None
                        df.at[ridx, target_hosp] = doc
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        break  # 次のdocへ

                if fixed_in_this_iteration > 0:
                    break  # 次のdocへ

        # 進捗がなければループ終了
        if fixed_in_this_iteration == 0:
            break

    # 最終確認（CC除外で判定）
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)
    remaining_violations = sum(1 for doc in active_doctors if abs((bg_counts.get(doc, 0) - cc_bg_counts.get(doc, 0)) - (ht_counts.get(doc, 0) - cc_ht_counts.get(doc, 0))) >= 3)

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全ての大学系と外病院の差3以上の違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件の大学系と外病院の差3以上の違反が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_gap_violations(pattern_df, max_attempts=200, verbose=True):
    """
    gap違反（3日未満の間隔での割当）を修正する

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    for attempt in range(max_attempts):
        # 現在の割当状態を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # gap違反を検出
        gap_violation_list = []
        for doc, date_hosp_list in doc_assignments.items():
            dates = sorted([d for d, h in date_hosp_list])
            for i in range(1, len(dates)):
                gap = (dates[i] - dates[i-1]).days
                if gap < 3:
                    gap_violation_list.append((doc, dates[i-1], dates[i], gap))

        if not gap_violation_list:
            if verbose and total_fixed > 0:
                print(f"   ✅ gap違反（3日未満の間隔）を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ gap違反を{len(gap_violation_list)}件検出 → 自動修正を開始...")

        # 修正試行（1イテレーションで複数の違反を修正）
        fixed_in_this_iteration = 0

        for doc, date1, date2, gap in gap_violation_list:
            if gap >= 3:
                continue

            # date2の割当を探す
            positions_at_date2 = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)
                if date != date2:
                    continue

                for hosp in hospital_cols:
                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == doc:
                        # v6.2.0: 固定割当は移動対象外
                        if is_preassigned_slot(ridx, hosp):
                            continue
                        positions_at_date2.append((ridx, hosp, date))

            # 各positionに対して修正を試みる
            for ridx_src, hosp_src, date_src in positions_at_date2:
                moved = False

                # 移動先候補を探す
                for ridx_tgt in df.index:
                    date_tgt = df.at[ridx_tgt, date_col_shift]
                    if pd.isna(date_tgt):
                        continue
                    date_tgt = pd.to_datetime(date_tgt).normalize().tz_localize(None)

                    # date1とdate_tgtの間隔をチェック
                    gap_from_date1 = abs((date_tgt - date1).days)
                    if gap_from_date1 < 4:
                        continue

                    # docの他の割当とdate_tgtの間隔をチェック
                    doc_dates = sorted([d for d, h in doc_assignments[doc]])
                    doc_dates_without_date2 = [d for d in doc_dates if d != date2]

                    valid_gap = True
                    for existing_date in doc_dates_without_date2:
                        if abs((date_tgt - existing_date).days) < 4:
                            valid_gap = False
                            break

                    if not valid_gap:
                        continue

                    # その日にdocが既に割当られていないかチェック
                    already_assigned = False
                    for hosp_check in hospital_cols:
                        val = df.at[ridx_tgt, hosp_check]
                        if isinstance(val, str) and normalize_name(val) == doc:
                            already_assigned = True
                            break

                    if already_assigned:
                        continue

                    # 全ての病院で空き枠を探す
                    hospitals_to_try = [hosp_src] + [h for h in hospital_cols if h != hosp_src]
                    for hosp_tgt in hospitals_to_try:
                        if pd.isna(df.at[ridx_tgt, hosp_tgt]):
                            # ハード制約チェック
                            if not can_assign_doc_to_slot(doc, date_tgt, hosp_tgt):
                                continue

                            # 移動実行
                            df.at[ridx_src, hosp_src] = None
                            df.at[ridx_tgt, hosp_tgt] = doc
                            fixed_in_this_iteration += 1
                            total_fixed += 1
                            moved = True
                            break

                    if moved:
                        break

                # 移動先が見つからない場合は削除（積極的に実行）
                if not moved and attempt >= 5:  # 5回目以降は削除も検討
                    df.at[ridx_src, hosp_src] = None
                    fixed_in_this_iteration += 1
                    total_fixed += 1
                    if verbose and attempt < 10:
                        print(f"      移動先が見つからないため、{doc}の{date_src.strftime('%m/%d')}の割当を削除します")
                    break  # この違反の他のpositionは試さない

            # この違反を修正したら、doc_assignmentsを更新
            if fixed_in_this_iteration > 0:
                counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # 進捗チェック
        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        # 連続で20回修正できなければ諦める
        if consecutive_failures >= 20:
            break

    # 最終確認
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)
    remaining_violations = 0
    for doc, date_hosp_list in doc_assignments.items():
        dates = sorted([d for d, h in date_hosp_list])
        for i in range(1, len(dates)):
            if (dates[i] - dates[i-1]).days < 3:
                remaining_violations += 1

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全てのgap違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件のgap違反が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_external_hospital_dup_violations(pattern_df, max_attempts=150, verbose=True):
    """
    外病院（L～Y列）の重複を修正する（大学病院の重複は許容）

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    for attempt in range(max_attempts):
        # 現在の割当状態を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        # CC分の病院別カウントを計算（重複検出から除外用）
        cc_hosp_counts = {d: defaultdict(int) for d in doctor_names}
        for ridx in df.index:
            date = df.at[ridx, date_col_shift]
            if pd.isna(date):
                continue
            date = pd.to_datetime(date).normalize().tz_localize(None)
            for hosp in hospital_cols:
                val = df.at[ridx, hosp]
                if isinstance(val, str):
                    doc = normalize_name(val)
                    if doc in doctor_names and is_cc_assignment(date, doc):
                        cc_hosp_counts[doc][hosp] += 1

        # 外病院重複を検出（CC除外）
        external_dup_list = []
        for doc, hosp_dict in assigned_hosp_count.items():
            for hosp, count in hosp_dict.items():
                # CC分を除外
                count_no_cc = count - cc_hosp_counts.get(doc, {}).get(hosp, 0)
                if count_no_cc > 1:
                    # 外病院かどうかを判定
                    hidx = shift_df.columns.get_loc(hosp)
                    if L_COL_INDEX <= hidx <= L_Y_END_INDEX:
                        external_dup_list.append((doc, hosp, count_no_cc))

        if not external_dup_list:
            if verbose and total_fixed > 0:
                print(f"   ✅ 外病院重複を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 外病院重複を{len(external_dup_list)}件検出 → 自動修正を開始...")

        # 修正試行
        fixed_in_this_iteration = 0

        for doc, dup_hosp, count in external_dup_list:
            if count <= 1:
                continue

            # この医師のこの病院への割当を探す
            dup_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                val = df.at[ridx, dup_hosp]
                if isinstance(val, str) and normalize_name(val) == doc:
                    # v6.2.0: 固定割当は移動対象外
                    if is_preassigned_slot(ridx, dup_hosp):
                        continue
                    dup_positions.append((ridx, dup_hosp, date))

            # 重複のうち1つを残して、残りを別の病院に移動または削除
            random.shuffle(dup_positions)

            for ridx, hosp, date in dup_positions[1:]:  # 最初の1つは残す
                moved = False

                # 同じ日の他の外病院（L～Y列）の空き枠を探す
                for other_hosp in hospital_cols:
                    other_hidx = shift_df.columns.get_loc(other_hosp)
                    # 外病院かつ重複病院でない
                    if not (L_COL_INDEX <= other_hidx <= L_Y_END_INDEX):
                        continue
                    if other_hosp == dup_hosp:
                        continue

                    # この病院にこの医師が既に割当られていないか
                    if assigned_hosp_count[doc].get(other_hosp, 0) >= 1:
                        continue

                    # 空き枠があるか
                    if pd.isna(df.at[ridx, other_hosp]):
                        # ハード制約チェック
                        if not can_assign_doc_to_slot(doc, date, other_hosp):
                            continue

                        # 移動実行
                        df.at[ridx, dup_hosp] = None
                        df.at[ridx, other_hosp] = doc
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        moved = True
                        break

                # 移動先が見つからない場合は削除
                if not moved and attempt >= 5:
                    df.at[ridx, dup_hosp] = None
                    fixed_in_this_iteration += 1
                    total_fixed += 1
                    if verbose and attempt < 10:
                        print(f"      移動先が見つからないため、{doc}の{date.strftime('%m/%d')}の{dup_hosp}割当を削除します")
                    break  # この重複の他のpositionは次回

            if fixed_in_this_iteration > 0:
                counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # 進捗チェック
        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        # 連続で20回修正できなければ諦める
        if consecutive_failures >= 20:
            break

    # 最終確認（CC除外で判定）
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)
    # CC分の病院別カウントを再計算
    cc_hosp_counts_final = {d: defaultdict(int) for d in doctor_names}
    for ridx_check in pattern_df.index:
        date_check = df.at[ridx_check, date_col_shift]
        if pd.isna(date_check):
            continue
        date_check = pd.to_datetime(date_check).normalize().tz_localize(None)
        for hosp_check in hospital_cols:
            val_check = df.at[ridx_check, hosp_check]
            if not isinstance(val_check, str):
                continue
            doc_check = normalize_name(val_check)
            if doc_check in doctor_names and is_cc_assignment(date_check, doc_check):
                cc_hosp_counts_final[doc_check][hosp_check] += 1
    remaining_violations = 0
    for doc, hosp_dict in assigned_hosp_count.items():
        for hosp, count in hosp_dict.items():
            count_no_cc = count - cc_hosp_counts_final.get(doc, {}).get(hosp, 0)
            if count_no_cc > 1:
                hidx = shift_df.columns.get_loc(hosp)
                if L_COL_INDEX <= hidx <= L_Y_END_INDEX:
                    remaining_violations += (count_no_cc - 1)

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全ての外病院重複を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件の外病院重複が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_university_over_2_violations(pattern_df, max_attempts=150, verbose=True):
    """
    大学病院（B～K列）が3回以上の医師の違反を修正する
    また、外病院0回の医師がいる場合も大学→外病院への移動を試みる（ハード制約）

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    for attempt in range(max_attempts):
        # 現在の割当状態を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        # 大学3回以上の医師を検出（CC除外）
        over_2_list = []
        for doc in active_doctors:
            if doc in RATIO_EXEMPT_DOCTORS:  # コード3は外病院専門なので除外
                continue
            # CCは大型連休特別シフトなので除外
            bg_count_no_cc = bg_counts.get(doc, 0) - cc_bg_counts.get(doc, 0)
            if bg_count_no_cc >= 3:
                over_2_list.append((doc, bg_count_no_cc, "大学3回以上"))

        # 外病院0回の医師を検出（大学を外病院に移動する必要あり）
        # 注：これはハード制約なのでCCは除外しない
        for doc in active_doctors:
            if doc in RATIO_EXEMPT_DOCTORS:  # コード3は外病院専門なので対象外
                continue
            ht_count = ht_counts.get(doc, 0)
            bg_count = bg_counts.get(doc, 0)
            # 外病院0回かつ大学1回以上なら、大学→外病院への移動が必要
            if ht_count == 0 and bg_count >= 1:
                # 既にover_2_listに含まれていないかチェック
                if not any(d == doc for d, _, _ in over_2_list):
                    over_2_list.append((doc, bg_count, "外病院0回"))

        if not over_2_list:
            if verbose and total_fixed > 0:
                print(f"   ✅ 大学3回以上/外病院0回違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            over_3_count = sum(1 for _, _, reason in over_2_list if reason == "大学3回以上")
            ext_0_count = sum(1 for _, _, reason in over_2_list if reason == "外病院0回")
            if over_3_count > 0:
                print(f"   ⚠️ 大学3回以上違反を{over_3_count}件検出")
            if ext_0_count > 0:
                print(f"   ⚠️ 外病院0回違反を{ext_0_count}件検出")

        # 修正試行
        fixed_in_this_iteration = 0

        for doc, bg_count, reason in over_2_list:
            # 大学3回以上の場合は2回に減らす、外病院0回の場合は1回移動
            if reason == "大学3回以上" and bg_count < 3:
                continue
            if reason == "外病院0回" and bg_count < 1:
                continue

            # この医師の大学病院への割当を探す
            bg_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    hidx = shift_df.columns.get_loc(hosp)
                    # 大学病院（B～K列）か
                    if not (B_COL_INDEX <= hidx <= B_K_END_INDEX):
                        continue

                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == doc:
                        # v6.2.0: 固定割当は移動対象外
                        if is_preassigned_slot(ridx, hosp):
                            continue
                        bg_positions.append((ridx, hosp, date))

            # 移動数を決定
            if reason == "大学3回以上":
                excess = bg_count - 2  # 2回まで減らす
            else:  # 外病院0回
                excess = 1  # 1回だけ移動

            random.shuffle(bg_positions)

            for ridx, hosp, date in bg_positions[:excess]:
                moved = False

                # 同じ日の外病院（L～Y列）の空き枠に移動を試みる
                for other_hosp in hospital_cols:
                    other_hidx = shift_df.columns.get_loc(other_hosp)
                    # 外病院か
                    if not (L_COL_INDEX <= other_hidx <= L_Y_END_INDEX):
                        continue

                    # この病院にこの医師が既に割当られていないか
                    if assigned_hosp_count[doc].get(other_hosp, 0) >= 1:
                        continue

                    # 空き枠があるか
                    if pd.isna(df.at[ridx, other_hosp]):
                        # ハード制約チェック
                        if not can_assign_doc_to_slot(doc, date, other_hosp):
                            continue

                        # 移動実行
                        df.at[ridx, hosp] = None
                        df.at[ridx, other_hosp] = doc
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        moved = True
                        break

                # 移動先が見つからない場合は代替医師を探して割り当て（未割当防止）
                if not moved and attempt >= 5:
                    # この日に既に割り当てられている医師を取得
                    already_on_date = set()
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str):
                            already_on_date.add(normalize_name(v))
                    # 代替候補: 同日重複なし & 大学系2回未満の医師
                    replacement_candidates = [
                        d for d in doctor_names
                        if d not in already_on_date
                        and d != doc
                        and bg_counts.get(d, 0) < 2
                        and can_assign_doc_to_slot(d, date, hosp)
                    ]
                    if replacement_candidates:
                        replacement_candidates.sort(key=lambda d: bg_counts.get(d, 0))
                        new_doc = replacement_candidates[0]
                        df.at[ridx, hosp] = new_doc
                        if verbose and attempt < 10:
                            print(f"      {doc}→{new_doc}: {date.strftime('%m/%d')}の大学病院割当を交代")
                    else:
                        # 緊急フォールバック: 絶対禁忌をすべて回避
                        hosp_idx = shift_df.columns.get_loc(hosp)
                        is_external_hosp = L_COL_INDEX <= hosp_idx <= L_Y_END_INDEX
                        day_of_week = pd.to_datetime(date).weekday()

                        def is_valid_bg_emergency(d):
                            # 同日重複禁止 (ABS-006)
                            if d in already_on_date or d == doc:
                                return False
                            code = get_avail_code(date, d)
                            # ABS-001: コード0禁止
                            if code == 0:
                                return False
                            # ABS-002: コード2はB〜Q列のみ
                            if code == 2 and not (B_COL_INDEX <= hosp_idx <= Q_COL_INDEX):
                                return False
                            # ABS-003: コード3はL〜Y列のみ
                            if code == 3 and not is_external_hosp:
                                return False
                            # ABS-004: カテ表コードありの日はL〜Y列不可
                            if is_external_hosp and get_sched_code(date, d):
                                return False
                            # ABS-005: 水曜日L〜Y列禁止医師
                            if day_of_week == 2 and is_external_hosp and d in WED_FORBIDDEN_DOCTORS:
                                return False
                            # ABS-007: gap >= 3日必須
                            d_dates = sorted([dt for dt, _ in doc_assignments.get(d, [])])
                            if d_dates:
                                min_gap = min(abs((date - dt).days) for dt in d_dates)
                                if min_gap < 3:
                                    return False
                            # ABS-008: 外病院重複禁止
                            if is_external_hosp and assigned_hosp_count.get(d, {}).get(hosp, 0) >= 1:
                                return False
                            return True

                        emergency = [d for d in doctor_names if is_valid_bg_emergency(d)]
                        if emergency:
                            emergency.sort(key=lambda d: bg_counts.get(d, 0))
                            new_doc = emergency[0]
                            df.at[ridx, hosp] = new_doc
                            if verbose and attempt < 10:
                                print(f"      {doc}→{new_doc}(緊急): {date.strftime('%m/%d')}の大学病院割当を交代")
                        else:
                            # 最終手段: 元の医師を維持（削除しない）
                            df.at[ridx, hosp] = doc
                            if verbose and attempt < 10:
                                print(f"      {doc}: {date.strftime('%m/%d')}の割当維持（絶対禁忌回避不可）")
                    fixed_in_this_iteration += 1
                    total_fixed += 1

            if fixed_in_this_iteration > 0:
                counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # 進捗チェック
        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        # 連続で20回修正できなければ諦める
        if consecutive_failures >= 20:
            break

    # 最終確認（CC除外で判定）
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)
    remaining_over_2 = sum(1 for doc in active_doctors if doc not in RATIO_EXEMPT_DOCTORS and (bg_counts.get(doc, 0) - cc_bg_counts.get(doc, 0)) >= 3)
    remaining_ext_0 = sum(1 for doc in active_doctors if doc not in RATIO_EXEMPT_DOCTORS and ht_counts.get(doc, 0) == 0 and bg_counts.get(doc, 0) >= 1)
    remaining_violations = remaining_over_2 + remaining_ext_0

    if verbose:
        if remaining_violations == 0:
            if total_fixed > 0:
                print(f"   ✅ 全ての大学3回以上/外病院0回違反を修正しました（修正数: {total_fixed}）")
        else:
            if remaining_over_2 > 0:
                print(f"   ⚠️ {remaining_over_2}件の大学3回以上違反が残っています")
            if remaining_ext_0 > 0:
                print(f"   ⚠️ {remaining_ext_0}件の外病院0回違反が残っています")

    return df, remaining_violations == 0, total_fixed

def fix_weekly_bg_violations(pattern_df, max_attempts=150, verbose=True):
    """
    v6.5.0: 大学系7日間隔違反を修正する
    ABS-012改: 大学系は7日間隔必須

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    def get_bg_assignments():
        """各医師の大学系割当を取得"""
        bg_assign = {doc: [] for doc in doctor_names}  # doc -> [(date, hosp, ridx), ...]
        for ridx in df.index:
            date = df.at[ridx, date_col_shift]
            if pd.isna(date):
                continue
            date = pd.to_datetime(date).normalize()
            if date.tz is not None:
                date = date.tz_localize(None)

            for hosp in hospital_cols:
                hidx = shift_df.columns.get_loc(hosp)
                if not (B_COL_INDEX <= hidx <= K_COL_INDEX):
                    continue

                val = df.at[ridx, hosp]
                if not isinstance(val, str):
                    continue
                doc = normalize_name(val)
                if doc not in doctor_names:
                    continue

                bg_assign[doc].append((date, hosp, ridx))
        return bg_assign

    def find_7day_violations(bg_assign):
        """7日間隔違反を検出"""
        violations = []
        for doc in active_doctors:
            assignments = sorted(bg_assign[doc], key=lambda x: x[0])
            for i in range(1, len(assignments)):
                gap = abs((assignments[i][0] - assignments[i-1][0]).days)
                if gap < 7:
                    # 違反ペア: 後の割当を移動対象とする
                    violations.append((doc, assignments[i-1], assignments[i], gap))
        return violations

    for attempt in range(max_attempts):
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        bg_assign = get_bg_assignments()
        violations = find_7day_violations(bg_assign)

        if not violations:
            if verbose and total_fixed > 0:
                print(f"   ✅ 大学系7日間隔違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 大学系7日間隔違反を{len(violations)}件検出")

        fixed_in_this_iteration = 0

        for doc, prev_assign, curr_assign, gap in violations:
            # 両方の割当を移動候補として試行（v6.5.7: 2番目だけでなく1番目も試す）
            targets = [(curr_assign, prev_assign), (prev_assign, curr_assign)]

            moved = False
            for target_assign, other_assign in targets:
                date, hosp, ridx = target_assign

                # 固定割当はスキップ
                if is_preassigned_slot(ridx, hosp):
                    continue

                # 同じ日の外病院（L～Y列）の空き枠に移動を試みる
                for other_hosp in hospital_cols:
                    other_hidx = shift_df.columns.get_loc(other_hosp)
                    if not (L_COL_INDEX <= other_hidx <= L_Y_END_INDEX):
                        continue

                    if assigned_hosp_count[doc].get(other_hosp, 0) >= 1:
                        continue

                    if pd.isna(df.at[ridx, other_hosp]):
                        if not can_assign_doc_to_slot(doc, date, other_hosp):
                            continue

                        df.at[ridx, hosp] = None
                        df.at[ridx, other_hosp] = doc
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        moved = True
                        break

                if moved:
                    break

            # 別の医師と交換を試みる（v6.5.7: attempt >= 2 に前倒し、両方の割当で試行）
            if not moved and attempt >= 2:
                for target_assign, other_assign in targets:
                    date, hosp, ridx = target_assign

                    if is_preassigned_slot(ridx, hosp):
                        continue

                    already_on_date = set()
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str):
                            already_on_date.add(normalize_name(v))

                    # 代替候補: 同日重複なし & 7日間隔を満たす医師
                    replacement_candidates = []
                    for d in doctor_names:
                        if d in already_on_date or d == doc:
                            continue
                        if not can_assign_doc_to_slot(d, date, hosp):
                            continue
                        # 7日間隔チェック
                        d_assignments = bg_assign.get(d, [])
                        has_conflict = False
                        for d_date, _, _ in d_assignments:
                            if abs((date - d_date).days) < 7:
                                has_conflict = True
                                break
                        if has_conflict:
                            continue
                        replacement_candidates.append(d)

                    if replacement_candidates:
                        replacement_candidates.sort(key=lambda d: bg_counts.get(d, 0))
                        new_doc = replacement_candidates[0]
                        df.at[ridx, hosp] = new_doc
                        if verbose and attempt < 10:
                            print(f"      {doc}→{new_doc}: {date.strftime('%m/%d')}（{gap}日間隔違反）の大学割当を交代")
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        moved = True
                        break

            if fixed_in_this_iteration > 0:
                counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)
                bg_assign = get_bg_assignments()

        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        if consecutive_failures >= 20:
            break

    # 最終確認
    bg_assign = get_bg_assignments()
    remaining_violations = len(find_7day_violations(bg_assign))

    if verbose:
        if remaining_violations == 0:
            if total_fixed > 0:
                print(f"   ✅ 全ての大学系7日間隔違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件の大学系7日間隔違反が残っています")

    return df, remaining_violations == 0, total_fixed

def fix_university_weekday_balance_violations(pattern_df, max_attempts=150, verbose=True):
    """
    大学病院の平日偏り（平日2回以上）の違反を修正する

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    for attempt in range(max_attempts):
        # 現在の割当状態を再計算
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        # 大学の平日2回以上の医師を検出
        # 注：weekday_countからCCを除外するには追加トラッキングが必要
        # 現時点ではweekday_countはそのまま使用（大型連休は平日カウントされにくい）
        weekday_over_list = []
        for doc in active_doctors:
            weekday_count = bg_cat[doc].get("平日", 0)
            if weekday_count >= 2:
                weekday_over_list.append((doc, weekday_count, bg_counts.get(doc, 0)))

        if not weekday_over_list:
            if verbose and total_fixed > 0:
                print(f"   ✅ 大学平日偏り違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 大学平日偏り違反を{len(weekday_over_list)}件検出 → 自動修正を開始...")

        # 修正試行
        fixed_in_this_iteration = 0

        for doc, weekday_count, bg_total in weekday_over_list:
            if weekday_count < 2:
                continue

            # この医師の大学病院平日の割当を探す
            bg_weekday_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    hidx = shift_df.columns.get_loc(hosp)
                    # 大学病院（B～K列）か
                    if not (B_COL_INDEX <= hidx <= B_K_END_INDEX):
                        continue

                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == doc:
                        # v6.2.0: 固定割当は移動対象外
                        if is_preassigned_slot(ridx, hosp):
                            continue
                        # 平日か
                        category = classify_bg_category(date, hosp)
                        if category == "平日":
                            bg_weekday_positions.append((ridx, hosp, date))

            # 平日のうち1つを外病院に移動または削除
            random.shuffle(bg_weekday_positions)

            for ridx, hosp, date in bg_weekday_positions[:1]:  # 1つだけ試行
                moved = False

                # 同じ日の外病院（L～Y列）の空き枠に移動を試みる
                for other_hosp in hospital_cols:
                    other_hidx = shift_df.columns.get_loc(other_hosp)
                    # 外病院か
                    if not (L_COL_INDEX <= other_hidx <= L_Y_END_INDEX):
                        continue

                    # この病院にこの医師が既に割当られていないか
                    if assigned_hosp_count[doc].get(other_hosp, 0) >= 1:
                        continue

                    # 空き枠があるか
                    if pd.isna(df.at[ridx, other_hosp]):
                        # ハード制約チェック
                        if not can_assign_doc_to_slot(doc, date, other_hosp):
                            continue

                        # 移動実行
                        df.at[ridx, hosp] = None
                        df.at[ridx, other_hosp] = doc
                        fixed_in_this_iteration += 1
                        total_fixed += 1
                        moved = True
                        break

                # 移動先が見つからない場合は代替医師を探して割り当て（未割当防止）
                if not moved and attempt >= 5:
                    # この日に既に割り当てられている医師を取得
                    already_on_date = set()
                    for h in hospital_cols:
                        v = df.at[ridx, h]
                        if isinstance(v, str):
                            already_on_date.add(normalize_name(v))
                    # 代替候補: 同日重複なし & 大学平日未割当の医師
                    replacement_candidates = [
                        d for d in doctor_names
                        if d not in already_on_date
                        and d != doc
                        and wd_counts.get(d, 0) < we_counts.get(d, 0)  # 平日<休日の医師を優先
                        and can_assign_doc_to_slot(d, date, hosp)
                    ]
                    if replacement_candidates:
                        replacement_candidates.sort(key=lambda d: wd_counts.get(d, 0))
                        new_doc = replacement_candidates[0]
                        df.at[ridx, hosp] = new_doc
                        if verbose and attempt < 10:
                            print(f"      {doc}→{new_doc}: {date.strftime('%m/%d')}の大学平日割当を交代")
                    else:
                        # 緊急フォールバック: 絶対禁忌をすべて回避
                        hosp_idx = shift_df.columns.get_loc(hosp)
                        is_external_hosp = L_COL_INDEX <= hosp_idx <= L_Y_END_INDEX
                        day_of_week = pd.to_datetime(date).weekday()

                        def is_valid_weekday_emergency(d):
                            # 同日重複禁止 (ABS-006)
                            if d in already_on_date or d == doc:
                                return False
                            code = get_avail_code(date, d)
                            # ABS-001: コード0禁止
                            if code == 0:
                                return False
                            # ABS-002: コード2はB〜Q列のみ
                            if code == 2 and not (B_COL_INDEX <= hosp_idx <= Q_COL_INDEX):
                                return False
                            # ABS-003: コード3はL〜Y列のみ
                            if code == 3 and not is_external_hosp:
                                return False
                            # ABS-004: カテ表コードありの日はL〜Y列不可
                            if is_external_hosp and get_sched_code(date, d):
                                return False
                            # ABS-005: 水曜日L〜Y列禁止医師
                            if day_of_week == 2 and is_external_hosp and d in WED_FORBIDDEN_DOCTORS:
                                return False
                            # ABS-007: gap >= 3日必須
                            d_dates = sorted([dt for dt, _ in doc_assignments.get(d, [])])
                            if d_dates:
                                min_gap = min(abs((date - dt).days) for dt in d_dates)
                                if min_gap < 3:
                                    return False
                            # ABS-008: 外病院重複禁止
                            if is_external_hosp and assigned_hosp_count.get(d, {}).get(hosp, 0) >= 1:
                                return False
                            return True

                        emergency = [d for d in doctor_names if is_valid_weekday_emergency(d)]
                        if emergency:
                            emergency.sort(key=lambda d: counts.get(d, 0))
                            new_doc = emergency[0]
                            df.at[ridx, hosp] = new_doc
                            if verbose and attempt < 10:
                                print(f"      {doc}→{new_doc}(緊急): {date.strftime('%m/%d')}の大学平日割当を交代")
                        else:
                            # 最終手段: 元の医師を維持
                            df.at[ridx, hosp] = doc
                            if verbose and attempt < 10:
                                print(f"      {doc}: {date.strftime('%m/%d')}の割当維持（絶対禁忌回避不可）")
                    fixed_in_this_iteration += 1
                    total_fixed += 1
                    break

            if fixed_in_this_iteration > 0:
                counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

        # 進捗チェック
        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        # 連続で20回修正できなければ諦める
        if consecutive_failures >= 20:
            break

    # 最終確認
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, *_ = recompute_stats(df)
    remaining_violations = sum(1 for doc in active_doctors if bg_cat[doc].get("平日", 0) >= 2)

    if verbose:
        if remaining_violations == 0:
            print(f"   ✅ 全ての大学平日偏り違反を修正しました（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ {remaining_violations}件の大学平日偏り違反が残っています（修正数: {total_fixed}）")

    return df, remaining_violations == 0, total_fixed

def fix_weekday_weekend_balance(pattern_df, max_attempts=200, verbose=True):
    """
    全体の平日/休日偏り（差>=2）を修正する。
    平日偏重の医師の平日割当と、休日偏重の医師の休日割当を交換する。
    """
    df = pattern_df.copy()
    total_fixed = 0

    for attempt in range(max_attempts):
        counts, bg_counts, ht_counts, wd_counts, we_counts, *_ = recompute_stats(df)

        # 偏り医師を特定
        wd_heavy = []  # 平日過多: wd - we >= 2
        we_heavy = []  # 休日過多: we - wd >= 2
        for doc in active_doctors:
            wd = wd_counts.get(doc, 0)
            we = we_counts.get(doc, 0)
            if wd - we >= 2:
                wd_heavy.append((doc, wd - we))
            elif we - wd >= 2:
                we_heavy.append((doc, we - wd))

        if not wd_heavy and not we_heavy:
            if verbose and total_fixed > 0:
                print(f"   ✅ 平日/休日偏り違反を{total_fixed}件修正しました")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            print(f"   ⚠️ 平日/休日偏り違反: 平日過多{len(wd_heavy)}人, 休日過多{len(we_heavy)}人 → 自動修正...")

        fixed_this = False

        # 平日過多の医師と休日過多の医師間でスワップ
        for wd_doc, wd_diff in sorted(wd_heavy, key=lambda x: -x[1]):
            if fixed_this:
                break
            for we_doc, we_diff in sorted(we_heavy, key=lambda x: -x[1]):
                if fixed_this:
                    break

                # wd_docの平日割当を探す
                wd_slots = []
                we_slots = []
                for (ridx, hosp), (date, fixed_flag) in slot_meta.items():
                    if fixed_flag:
                        continue
                    val = df.at[ridx, hosp]
                    if not isinstance(val, str):
                        continue
                    hidx = shift_df.columns.get_loc(hosp)
                    dow = date.weekday()
                    weekday = dow < 5
                    is_holi = (
                        is_holiday(date)
                        or dow >= 5
                        or (weekday and hidx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX))
                    )
                    doc_name = normalize_name(val)
                    if doc_name == wd_doc and not is_holi:
                        wd_slots.append((ridx, hosp, date))
                    elif doc_name == we_doc and is_holi:
                        we_slots.append((ridx, hosp, date))

                random.shuffle(wd_slots)
                random.shuffle(we_slots)

                for wd_ridx, wd_hosp, wd_date in wd_slots:
                    if fixed_this:
                        break
                    for we_ridx, we_hosp, we_date in we_slots:
                        # スワップ可能かチェック
                        # wd_doc → we_hosp(休日), we_doc → wd_hosp(平日)
                        if not can_assign_doc_to_slot(wd_doc, we_date, we_hosp):
                            continue
                        if not can_assign_doc_to_slot(we_doc, wd_date, wd_hosp):
                            continue

                        # 同日重複チェック
                        def has_other_assignment(doc, ridx, exclude_hosp):
                            for h in hospital_cols:
                                if h == exclude_hosp:
                                    continue
                                v = df.at[ridx, h]
                                if isinstance(v, str) and normalize_name(v) == doc:
                                    return True
                            return False

                        if has_other_assignment(wd_doc, we_ridx, we_hosp):
                            continue
                        if has_other_assignment(we_doc, wd_ridx, wd_hosp):
                            continue

                        # gap チェック（スワップ後に3日未満にならないか）
                        def check_gap_ok(doc, new_date, old_date):
                            assigns = []
                            for (r, h), (d, _) in slot_meta.items():
                                v = df.at[r, h]
                                if isinstance(v, str) and normalize_name(v) == doc:
                                    if d != old_date:
                                        assigns.append(d)
                            assigns.append(new_date)
                            assigns.sort()
                            for i in range(1, len(assigns)):
                                if abs((assigns[i] - assigns[i-1]).days) < 3:
                                    return False
                            return True

                        if not check_gap_ok(wd_doc, we_date, wd_date):
                            continue
                        if not check_gap_ok(we_doc, wd_date, we_date):
                            continue

                        # スワップ実行
                        df.at[wd_ridx, wd_hosp] = we_doc
                        df.at[we_ridx, we_hosp] = wd_doc
                        total_fixed += 1
                        fixed_this = True
                        break

        if not fixed_this:
            break  # これ以上修正できない

    # 最終確認
    _, _, _, wd_counts, we_counts, *_ = recompute_stats(df)
    remaining = sum(1 for d in active_doctors if abs(wd_counts.get(d, 0) - we_counts.get(d, 0)) >= 2)
    if verbose and remaining > 0:
        print(f"   ⚠️ {remaining}人の平日/休日偏り違反が残っています（修正数: {total_fixed}）")

    return df, remaining == 0, total_fixed

def fix_fairness_imbalance(pattern_df, max_attempts=200, verbose=True):
    """
    active医師間の割当回数の公平性を強化する（最大と最小の差を縮める）

    Args:
        pattern_df: スケジュールDataFrame
        max_attempts: 最大試行回数
        verbose: ログ出力するか

    Returns:
        (修正後のDataFrame, 成功フラグ, 修正数)
    """
    df = pattern_df.copy()
    total_fixed = 0
    consecutive_failures = 0

    for attempt in range(max_attempts):
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, cc_counts, cc_bg_counts, cc_ht_counts = recompute_stats(df)

        # active医師の割当回数を確認（CC除外）
        # CCは大型連休特別シフトなので公平性計算から除外
        active_counts = [(doc, counts.get(doc, 0) - cc_counts.get(doc, 0)) for doc in active_doctors]
        if not active_counts:
            return df, True, total_fixed

        # 最大と最小を取得
        max_count = max(c for _, c in active_counts)
        min_count = min(c for _, c in active_counts)
        diff = max_count - min_count

        # 差が1以下なら公平性達成
        if diff <= 1:
            if verbose and total_fixed > 0:
                print(f"   ✅ 公平性違反を{total_fixed}件修正しました（max={max_count}, min={min_count}, diff={diff}）")
            return df, True, total_fixed

        if attempt == 0 and verbose:
            max_docs = [doc for doc, c in active_counts if c == max_count]
            min_docs = [doc for doc, c in active_counts if c == min_count]
            print(f"   ⚠️ 公平性違反を検出（max={max_count}, min={min_count}, diff={diff}） → 自動修正を開始...")

        fixed_in_this_iteration = 0

        # 最大回数の医師から最小回数の医師にシフトを移動
        max_docs = [doc for doc, c in active_counts if c == max_count]
        min_docs = [doc for doc, c in active_counts if c == min_count]

        random.shuffle(max_docs)
        random.shuffle(min_docs)

        # 最大回数の医師のシフトを探す
        for max_doc in max_docs[:3]:  # 最大3人まで試行
            # max_docの割当位置を取得
            max_doc_positions = []
            for ridx in df.index:
                date = df.at[ridx, date_col_shift]
                if pd.isna(date):
                    continue
                date = pd.to_datetime(date).normalize().tz_localize(None)

                for hosp in hospital_cols:
                    # v6.5.9: 固定割当（preassigned）は公平化の移動対象から除外
                    if is_preassigned_slot(ridx, hosp):
                        continue
                    val = df.at[ridx, hosp]
                    if isinstance(val, str) and normalize_name(val) == max_doc:
                        max_doc_positions.append((ridx, hosp, date))

            random.shuffle(max_doc_positions)

            # 各位置について、最小回数の医師と入れ替え可能か試す
            for ridx, hosp, date in max_doc_positions[:5]:  # 最大5個まで試行
                # この日に既に割り当てられている医師を除外
                already_assigned_on_date = set()
                for h in hospital_cols:
                    v = df.at[ridx, h]
                    if isinstance(v, str):
                        already_assigned_on_date.add(normalize_name(v))

                # 対象スロットの平日/休日分類（recompute_statsと同一ロジック）
                hosp_idx = shift_df.columns.get_loc(hosp)
                slot_dow = date.weekday()
                slot_is_holiday = (
                    is_holiday(date)
                    or slot_dow >= 5
                    or (slot_dow < 5 and hosp_idx in (C_COL_INDEX, D_COL_INDEX, F_COL_INDEX, G_COL_INDEX))
                )

                # 最小回数の医師の中から代替を探す
                for min_doc in min_docs:
                    # v6.5.9: is_valid_full_assignmentで全ABS制約を統合チェック
                    # （ABS-001〜006静的制約, ABS-007 gap, ABS-008外病院重複,
                    #   ABS-010 TARGET_CAP, ABS-011大学系2回）
                    # 移動がABS違反を新規に作るとsafe_fixが全戻しし公平化自体が
                    # 無効化されるため、移動時点で完全に検証する
                    if not is_valid_full_assignment(
                        min_doc, date, hosp,
                        doc_assignments, counts, bg_counts, assigned_hosp_count,
                        already_on_date=already_assigned_on_date,
                    ):
                        continue

                    # v6.5.9: ABS-013 C-H列（休日大学系）カテ当番必須
                    if is_ch_slot(hosp_idx) and not is_eligible_for_ch_slot(min_doc, date):
                        continue

                    # v6.5.9: ABS-012 大学系7日間隔（B-K列に移す場合）
                    if B_COL_INDEX <= hosp_idx <= B_K_END_INDEX:
                        bg_dates = [
                            d for d, h in doc_assignments.get(min_doc, [])
                            if B_COL_INDEX <= shift_df.columns.get_loc(h) <= B_K_END_INDEX
                        ]
                        if any(abs((date - d).days) < 7 for d in bg_dates):
                            continue

                    # v6.5.9: ABS-014 平日/休日偏り（移動元・移動先とも差<=1を維持）
                    if slot_is_holiday:
                        min_wd, min_we = wd_counts.get(min_doc, 0), we_counts.get(min_doc, 0) + 1
                        max_wd, max_we = wd_counts.get(max_doc, 0), we_counts.get(max_doc, 0) - 1
                    else:
                        min_wd, min_we = wd_counts.get(min_doc, 0) + 1, we_counts.get(min_doc, 0)
                        max_wd, max_we = wd_counts.get(max_doc, 0) - 1, we_counts.get(max_doc, 0)
                    if abs(min_wd - min_we) >= 2 or abs(max_wd - max_we) >= 2:
                        continue

                    # 入れ替え
                    df.at[ridx, hosp] = min_doc
                    fixed_in_this_iteration += 1
                    total_fixed += 1

                    if verbose and attempt < 3:
                        print(f"      {date.strftime('%m/%d')} {hosp}: {max_doc}({max_count}回) → {min_doc}({min_count}回)")

                    # doc_assignmentsを更新（次の反復のため）
                    if max_doc in doc_assignments:
                        doc_assignments[max_doc] = [(d, h) for d, h in doc_assignments[max_doc] if h != hosp or d != date]
                    if min_doc not in doc_assignments:
                        doc_assignments[min_doc] = []
                    doc_assignments[min_doc].append((date, hosp))

                    break  # min_docs loop

                if fixed_in_this_iteration > 0:
                    break  # max_doc_positions loop

            if fixed_in_this_iteration > 0:
                break  # max_docs loop

        # 進捗チェック
        if fixed_in_this_iteration == 0:
            consecutive_failures += 1
        else:
            consecutive_failures = 0

        # 連続で20回修正できなければ諦める
        if consecutive_failures >= 20:
            break

    # 最終確認
    counts, *_ = recompute_stats(df)
    active_counts = [(doc, counts.get(doc, 0)) for doc in active_doctors]
    max_count = max(c for _, c in active_counts)
    min_count = min(c for _, c in active_counts)
    diff = max_count - min_count

    if verbose:
        if diff <= 1:
            print(f"   ✅ 公平性を達成しました（max={max_count}, min={min_count}, diff={diff}）（修正数: {total_fixed}）")
        else:
            print(f"   ⚠️ 公平性違反が残っています（max={max_count}, min={min_count}, diff={diff}）（修正数: {total_fixed}）")

    return df, diff <= 1, total_fixed

def fix_unassigned_slots(pattern_df, verbose=True):
    """
    slot_metaに登録されたスロットで医師が割り当てられていないものを埋める
    これは最終セーフティネットとして、全てのスロットに医師を配置することを保証する
    """
    df = pattern_df.copy()
    total_fixed = 0

    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(df)

    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = df.at[ridx, hosp]

        # 既に医師が割り当てられている場合はスキップ
        if isinstance(val, str):
            v_norm = normalize_name(val)
            if v_norm in doctor_names:
                continue

        # 未割り当てスロットを発見
        # この日に既に割り当てられている医師を取得
        already_assigned_on_date = set()
        for h in hospital_cols:
            v = df.at[ridx, h]
            if isinstance(v, str):
                already_assigned_on_date.add(normalize_name(v))

        # 候補医師を探す（制約チェック付き）
        # v6.0.2: ABS-010/ABS-011も厳守（CAP+1まで、大学系3回まで）
        col_idx = shift_df.columns.get_loc(hosp)
        is_university = B_COL_INDEX <= col_idx <= K_COL_INDEX
        is_external = L_COL_INDEX <= col_idx <= L_Y_END_INDEX

        candidates = [
            d for d in doctor_names
            if d not in already_assigned_on_date
            and can_assign_doc_to_slot(d, date, hosp)
            and counts.get(d, 0) < TARGET_CAP.get(d, 0)        # ABS-010: TARGET_CAP厳守
            and (not is_university or bg_counts.get(d, 0) < 2)  # ABS-011: 大学系2回まで
        ]

        if candidates:
            # 割当回数が少ない医師を優先
            candidates.sort(key=lambda d: counts.get(d, 0))
            new_doc = candidates[0]
        else:
            # 緊急フォールバック: 絶対禁忌をすべて回避
            day_of_week = pd.to_datetime(date).weekday()

            def is_valid_unassigned_fallback(d):
                # 同日重複禁止 (ABS-006)
                if d in already_assigned_on_date:
                    return False
                code = get_avail_code(date, d)
                # ABS-001: コード0禁止
                if code == 0:
                    return False
                # ABS-002: コード2はB〜Q列のみ
                if code == 2 and not (B_COL_INDEX <= col_idx <= Q_COL_INDEX):
                    return False
                # ABS-003: コード3はL〜Y列のみ
                if code == 3 and not is_external:
                    return False
                # ABS-004: カテ表コードありの日はL〜Y列不可
                if is_external and get_sched_code(date, d):
                    return False
                # ABS-005: 水曜日L〜Y列禁止医師
                if day_of_week == 2 and is_external and d in WED_FORBIDDEN_DOCTORS:
                    return False
                # ABS-007: gap >= 3日必須
                d_dates = sorted([dt for dt, _ in doc_assignments.get(d, [])])
                if d_dates:
                    min_gap = min(abs((date - dt).days) for dt in d_dates)
                    if min_gap < 3:
                        return False
                # ABS-008: 外病院重複禁止
                if is_external and assigned_hosp_count.get(d, {}).get(hosp, 0) >= 1:
                    return False
                # ABS-010: TARGET_CAP厳守
                if counts.get(d, 0) >= TARGET_CAP.get(d, 0):
                    return False
                # ABS-011: 大学系2回まで
                if is_university and bg_counts.get(d, 0) >= 2:
                    return False
                return True

            emergency = [d for d in doctor_names if is_valid_unassigned_fallback(d)]
            if emergency:
                emergency.sort(key=lambda d: counts.get(d, 0))
                new_doc = emergency[0]
            else:
                # 絶対禁忌を満たす候補がいない場合は未割当のまま
                if verbose:
                    print(f"   ❌ 未割当: {date.strftime('%Y-%m-%d')} {hosp}（絶対禁忌回避不可）")
                continue

        df.at[ridx, hosp] = new_doc
        counts[new_doc] = counts.get(new_doc, 0) + 1
        # v6.0.2: 追跡変数も更新（連続割当の制約チェック精度向上）
        if is_university:
            bg_counts[new_doc] = bg_counts.get(new_doc, 0) + 1
        if new_doc not in assigned_hosp_count:
            assigned_hosp_count[new_doc] = {}
        assigned_hosp_count[new_doc][hosp] = assigned_hosp_count[new_doc].get(hosp, 0) + 1
        if new_doc not in doc_assignments:
            doc_assignments[new_doc] = []
        doc_assignments[new_doc].append((date, hosp))
        total_fixed += 1

        if verbose:
            print(f"   🔧 未割り当て修正: {date.strftime('%Y-%m-%d')} {hosp} → {new_doc}")

    if verbose:
        if total_fixed == 0:
            print("   ✅ 未割り当てスロットなし")
        else:
            print(f"   ✅ {total_fixed}件の未割り当てスロットを修正しました")

    # 修正後の残り未割当数を確認して成功判定
    _, _, _, _, _, _, _, _, _, _, remaining_unassigned, *_ = recompute_stats(df)
    return df, (len(remaining_unassigned) == 0), total_fixed

def validate_absolute_constraints(pattern_df, verbose=True):
    """
    絶対禁忌の最終検証（v6.0.0）

    チェック項目:
    - ABS-001: コード0割当禁止
    - ABS-002: コード2列制限（B〜Q列のみ）
    - ABS-003: コード3列制限（L〜Y列のみ）
    - ABS-006: 同日重複禁止
    - ABS-007: gap >= 3日必須
    - ABS-008: 同一病院重複禁止（全列）
    - ABS-009: 未割当禁止
    - ABS-010: TARGET_CAP遵守
    - ABS-011: 大学系2回まで
    - ABS-012: 大学系7日間隔（v6.5.8）
    - ABS-013: C-H列カテ当番必須（v6.5.9で検証追加・固定割当は許容）
    - ABS-014: 平日/休日偏り差<=1
    - ABS-015: 属性2のB列カテ表コード欠如

    Returns:
        (violations_list, is_valid)
    """
    violations = []

    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(pattern_df)

    # ABS-001, ABS-002, ABS-003: コード制限チェック
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = pattern_df.at[ridx, hosp]
        if isinstance(val, str):
            doc = normalize_name(val)
            if doc in doctor_names:
                code = get_avail_code(date, doc)
                hidx = shift_df.columns.get_loc(hosp)
                # ABS-001: コード0禁止
                if code == 0:
                    violations.append({
                        "type": "ABS-001",
                        "desc": f"コード0割当: {doc} → {date.strftime('%Y-%m-%d')} {hosp}"
                    })
                # ABS-002: コード2はB〜Q列のみ
                if code == 2 and not (B_COL_INDEX <= hidx <= Q_COL_INDEX):
                    violations.append({
                        "type": "ABS-002",
                        "desc": f"コード2列違反: {doc} → {date.strftime('%Y-%m-%d')} {hosp}"
                    })
                # ABS-003: コード3はL〜Y列のみ
                if code == 3 and not (L_COL_INDEX <= hidx <= L_Y_END_INDEX):
                    violations.append({
                        "type": "ABS-003",
                        "desc": f"コード3列違反: {doc} → {date.strftime('%Y-%m-%d')} {hosp}"
                    })

    # ABS-006: 同日重複チェック
    for date, doc_count in build_date_doc_count(pattern_df).items():
        for doc, count in doc_count.items():
            if count > 1:
                violations.append({
                    "type": "ABS-006",
                    "desc": f"同日重複: {doc} → {date.strftime('%Y-%m-%d')} ({count}回)"
                })

    # ABS-007: gap >= 3日チェック
    for doc, assigns in doc_assignments.items():
        dates = sorted([d for d, _ in assigns])
        for i in range(1, len(dates)):
            gap = (dates[i] - dates[i-1]).days
            if gap < 3:
                violations.append({
                    "type": "ABS-007",
                    "desc": f"gap違反: {doc} → gap={gap}日 (必須>=3)"
                })

    # ABS-008: 同一病院重複チェック
    for doc, hosp_dict in assigned_hosp_count.items():
        for hosp, count in hosp_dict.items():
            if count > 1:
                violations.append({
                    "type": "ABS-008",
                    "desc": f"病院重複: {doc} → {hosp} ({count}回)"
                })

    # ABS-009: 未割当枠チェック
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = pattern_df.at[ridx, hosp]
        if not isinstance(val, str):
            violations.append({
                "type": "ABS-009",
                "desc": f"未割当: {date.strftime('%Y-%m-%d')} {hosp}"
            })
        elif normalize_name(val) not in doctor_names:
            violations.append({
                "type": "ABS-009",
                "desc": f"不明医師: {date.strftime('%Y-%m-%d')} {hosp} → {val}"
            })

    # ABS-010: TARGET_CAP遵守チェック
    for doc, count in counts.items():
        cap = TARGET_CAP.get(doc, 0)
        if count > cap:
            violations.append({
                "type": "ABS-010",
                "desc": f"TARGET_CAP超過: {doc} → {count}回 (上限{cap})"
            })

    # ABS-011: 大学系2回までチェック
    for doc, bg_count in bg_counts.items():
        if bg_count > 2:
            violations.append({
                "type": "ABS-011",
                "desc": f"大学系3回以上: {doc} → {bg_count}回 (上限2)"
            })

    # v6.5.8: ABS-012: 大学系7日間隔チェック
    bg_dates_by_doc_v = {doc: [] for doc in doctor_names}
    for ridx in pattern_df.index:
        date = pattern_df.at[ridx, date_col_shift]
        if pd.isna(date):
            continue
        date = pd.to_datetime(date).normalize()
        if date.tz is not None:
            date = date.tz_localize(None)
        for hosp in hospital_cols:
            hidx = shift_df.columns.get_loc(hosp)
            if not (B_COL_INDEX <= hidx <= K_COL_INDEX):
                continue
            val = pattern_df.at[ridx, hosp]
            if not isinstance(val, str):
                continue
            doc = normalize_name(val)
            if doc in bg_dates_by_doc_v:
                bg_dates_by_doc_v[doc].append(date)
    for doc in active_doctors:
        dates = sorted(bg_dates_by_doc_v.get(doc, []))
        for i in range(1, len(dates)):
            gap = abs((dates[i] - dates[i-1]).days)
            if gap < 7:
                violations.append({
                    "type": "ABS-012",
                    "desc": f"大学系7日間隔違反: {doc} → gap={gap}日 (必須>=7)"
                })

    # v6.5.9: ABS-013 C-H列（休日大学系）カテ当番必須チェック
    # 従来この検証が欠落しており、fix関数がABS-013を壊しても
    # 「絶対禁忌クリア」と表示され得た。固定割当（fixed）は意図的な
    # 配置として許容（evaluate側のch_kate_violationsと同一の扱い）
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        if fixed:
            continue
        hidx = shift_df.columns.get_loc(hosp)
        if not is_ch_slot(hidx):
            continue
        val = pattern_df.at[ridx, hosp]
        if not isinstance(val, str):
            continue
        doc = normalize_name(val)
        if doc not in doctor_names:
            continue
        if not is_eligible_for_ch_slot(doc, date):
            violations.append({
                "type": "ABS-013",
                "desc": f"C-H列カテ当番違反: {doc} → {date.strftime('%Y-%m-%d')} {hosp}"
            })

    # ABS-014: 全体の平日/休日偏り（差>=2）チェック
    for doc in active_doctors:
        wd = wd_counts.get(doc, 0)
        we = we_counts.get(doc, 0)
        diff = abs(wd - we)
        if diff >= 2:
            violations.append({
                "type": "ABS-014",
                "desc": f"平日/休日偏り: {doc} → 平日{wd}/休日{we} (差{diff}, 許容<=1)"
            })

    # ABS-015: 属性2のB列カテ表コード欠如チェック
    for (ridx, hosp), (date, fixed) in slot_meta.items():
        val = pattern_df.at[ridx, hosp]
        if not isinstance(val, str):
            continue
        doc = normalize_name(val)
        if doc not in doctor_names:
            continue
        hidx = shift_df.columns.get_loc(hosp)
        if hidx == B_COL_INDEX and doc in SCHEDULE_CODE_HOLDERS and doctor_attribute.get(doc, "") == "2":
            if not get_sched_code(date, doc) and doc not in EXTRA_ALLOWED:
                violations.append({
                    "type": "ABS-015",
                    "desc": f"属性2 B列カテ表欠如: {doc} → {date.strftime('%Y-%m-%d')} {hosp}"
                })

    is_valid = len(violations) == 0

    if verbose:
        if is_valid:
            print("   ✅ 絶対禁忌チェック: 全てクリア")
        else:
            print(f"   ❌ 絶対禁忌違反: {len(violations)}件")
            for v in violations[:10]:  # 最大10件表示
                print(f"      - [{v['type']}] {v['desc']}")
            if len(violations) > 10:
                print(f"      ... 他 {len(violations) - 10}件")

    return violations, is_valid


def safe_fix(fix_func, df, verbose=False, **kwargs):
    """
    v6.0.3/v6.1.0: fix関数の安全ラッパー
    fix関数実行後にABS違反が増えた場合、またはgap/dup違反が新たに発生した場合はrevertする。
    v6.1.0: 総数だけでなく、gap(ABS-007)/dup(ABS-008)の個別増加もrevert対象に追加。
    これにより、他の制約を直す代わりにgap/dupが悪化する「トレード」を防止する。

    Returns:
        (fixed_df, success, fix_count) + fix_func固有の追加戻り値
    """
    def count_by_type(violations):
        counts = {}
        for v in violations:
            t = v.get("type", "")
            counts[t] = counts.get(t, 0) + 1
        return counts

    pre_violations, _ = validate_absolute_constraints(df, verbose=False)
    pre_count = len(pre_violations)
    pre_by_type = count_by_type(pre_violations)

    result = fix_func(df, verbose=verbose, **kwargs)
    fixed_df = result[0]
    fix_count = result[2] if len(result) > 2 else 0

    post_violations, _ = validate_absolute_constraints(fixed_df, verbose=False)
    post_count = len(post_violations)
    post_by_type = count_by_type(post_violations)

    # revert条件1: ABS違反の総数が増えた
    should_revert = post_count > pre_count

    # revert条件2: gap(ABS-007)またはdup(ABS-008)が増えた（トレード防止）
    # ただしfix_gap_violations/fix_external_hospital_dup_violations自体は除外
    if not should_revert and fix_func.__name__ not in ("fix_gap_violations", "fix_external_hospital_dup_violations"):
        for critical_type in ("ABS-007", "ABS-008"):
            if post_by_type.get(critical_type, 0) > pre_by_type.get(critical_type, 0):
                should_revert = True
                if verbose:
                    print(f"   ⚠️ {fix_func.__name__}: {critical_type}増加 → revert")
                break

    if should_revert:
        if verbose and post_count > pre_count:
            print(f"   ⚠️ {fix_func.__name__}: ABS違反増加({pre_count}→{post_count}) → revert")
        if len(result) == 4:
            return df, False, 0, 0
        else:
            return df, False, 0
    return result


def build_diagnostics(pattern_df):
    counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, assigned_hosp_count, doc_assignments, unassigned, *_ = recompute_stats(pattern_df)
    score, raw, metrics = evaluate_schedule_with_raw(
        pattern_df,
        counts,
        bg_counts,
        ht_counts,
        wd_counts,
        we_counts,
        bk_counts,
        ly_counts,
    )

    df_doctors = build_doctor_diag(counts, bg_counts, ht_counts, wd_counts, we_counts, doc_assignments, assigned_hosp_count)
    df_gap = build_gap_details(doc_assignments)
    df_same = build_same_day_duplicates(doc_assignments)
    df_hdup = build_hosp_dup_details(assigned_hosp_count)
    df_weekly_bg = build_weekly_bg_details(doc_assignments)  # v6.4.0: 大学系週1違反
    df_unass = build_unassigned_details(unassigned)
    df_metrics = build_metrics_df(score, raw, metrics)
    df_hard_violations = build_hard_constraint_violations(pattern_df)

    return df_doctors, df_gap, df_same, df_hdup, df_weekly_bg, df_unass, df_metrics, df_hard_violations

# =========================
# パターン探索（greedy → top候補に局所探索 → top3）
# =========================
print("\n" + "="*60)
print("  🚀 スケジュール生成")
print("="*60)

score_rows = []
candidates = []  # TOP_KEEPだけ保持

for i in tqdm(range(1, NUM_PATTERNS + 1), desc="   パターン生成", ncols=60, disable=not TQDM_AVAILABLE):

    (
        pattern_df,
        counts,
        bg_counts,
        ht_counts,
        wd_counts,
        we_counts,
        bk_counts,
        ly_counts,
        bg_cat,
    ) = build_schedule_pattern(seed=i)
    score, raw_score, metrics = evaluate_schedule_with_raw(
        pattern_df,
        counts,
        bg_counts,
        ht_counts,
        wd_counts,
        we_counts,
        bk_counts,
        ly_counts,
    )

    score_rows.append({"seed": i, "score": score, "raw_score": raw_score, **metrics})

    # gap違反が0個のパターンのみ採用（完全なgap制約遵守）
    gap_violations = metrics.get("gap_violations", 0)
    if gap_violations == 0:
        candidates.append({
            "seed": i,
            "score": score,
            "raw_score": raw_score,
            "metrics": metrics,
            "pattern_df": pattern_df,
        })

# gap違反0個の候補をスコア順にソート
candidates = sorted(candidates, key=lambda e: e["raw_score"], reverse=True)[:TOP_KEEP]

if len(candidates) == 0:
    print("\n⚠️  gap違反0個の候補なし → 制約緩和して続行")
    # gap違反の制約を緩和して再選択
    candidates = []
    for row in score_rows:
        candidates.append({
            "seed": row["seed"],
            "score": row["score"],
            "raw_score": row["raw_score"],
            "metrics": {k: v for k, v in row.items() if k not in ["seed", "score", "raw_score"]},
            "pattern_df": None,  # 再生成が必要
        })
    candidates = sorted(candidates, key=lambda e: e["raw_score"], reverse=True)[:TOP_KEEP]
    # パターンを再生成
    for cand in candidates:
        if cand["pattern_df"] is None:
            pattern_df, *_ = build_schedule_pattern(seed=cand["seed"])
            cand["pattern_df"] = pattern_df

# ローカル探索で候補を改善
refined = []
refine_list = candidates[:REFINE_TOP]
for idx, cand in enumerate(tqdm(refine_list, desc="   局所探索    ", ncols=60, disable=not TQDM_AVAILABLE), 1):
    if LOCAL_SEARCH_ENABLED:
        improved_df, sc2, raw2, met2 = local_search_swap(
            cand["pattern_df"],
            max_iters=LOCAL_MAX_ITERS,
            patience=LOCAL_PATIENCE,
            refresh_every=LOCAL_REFRESH_EVERY,
            seed=1000 + cand["seed"],
        )
    else:
        improved_df = cand["pattern_df"]
        sc2 = cand["score"]
        raw2 = cand["raw_score"]
        met2 = cand["metrics"]

    # v6.0.3: safe_fixラッパー + 収束ループで最適化
    # 各fix関数をsafe_fixで実行。ABS違反が増えたらrevertされる。
    # 収束するまで最大3ラウンド繰り返す。
    if OPTIMIZATION_ENABLED:
        current_df = improved_df
        total_fix_counts = {}
        MAX_ROUNDS = 3

        for opt_round in range(MAX_ROUNDS):
            round_fixed = 0

            # 1. ハード制約違反の自動修正
            result = safe_fix(fix_hard_constraint_violations, current_df, max_attempts=50)
            current_df, _, fc = result[0], result[1], result[2]
            fail_count = result[3] if len(result) > 3 else 0
            total_fix_counts["hard"] = total_fix_counts.get("hard", 0) + fc
            round_fixed += fc

            # 2. 可否コード2医師のn+1回違反を修正
            current_df, _, fc = safe_fix(fix_code_2_extra_violations, current_df, max_attempts=100)
            total_fix_counts["code2"] = total_fix_counts.get("code2", 0) + fc
            round_fixed += fc

            # 3. TARGET_CAP違反の自動修正
            current_df, _, fc = safe_fix(fix_target_cap_violations, current_df, max_attempts=100)
            total_fix_counts["cap"] = total_fix_counts.get("cap", 0) + fc
            round_fixed += fc

            # 4. 大学系最低1回必須違反を修正
            current_df, _, fc = safe_fix(fix_university_minimum_requirement, current_df, max_attempts=100)
            total_fix_counts["univ_min"] = total_fix_counts.get("univ_min", 0) + fc
            round_fixed += fc

            # 5. C-H列カテ当番違反を修正
            current_df, _, fc = safe_fix(fix_ch_kate_violations, current_df, max_attempts=100)
            total_fix_counts["ch_kate"] = total_fix_counts.get("ch_kate", 0) + fc
            round_fixed += fc

            # 6. gap違反を修正
            current_df, _, fc = safe_fix(fix_gap_violations, current_df, max_attempts=200)
            total_fix_counts["gap"] = total_fix_counts.get("gap", 0) + fc
            round_fixed += fc

            # 7. 大学系/外病院バランス修正
            current_df, _, fc = safe_fix(fix_bg_ht_imbalance_violations, current_df, max_attempts=100)
            total_fix_counts["bg_ht"] = total_fix_counts.get("bg_ht", 0) + fc
            round_fixed += fc

            # 8. 外病院重複を修正
            current_df, _, fc = safe_fix(fix_external_hospital_dup_violations, current_df, max_attempts=150)
            total_fix_counts["ext_dup"] = total_fix_counts.get("ext_dup", 0) + fc
            round_fixed += fc

            # 9. 大学3回以上違反を修正
            current_df, _, fc = safe_fix(fix_university_over_2_violations, current_df, max_attempts=150)
            total_fix_counts["univ_over2"] = total_fix_counts.get("univ_over2", 0) + fc
            round_fixed += fc

            # 10. v6.4.0: 大学系週1違反を修正（ABS-012）
            current_df, _, fc = safe_fix(fix_weekly_bg_violations, current_df, max_attempts=150)
            total_fix_counts["weekly_bg"] = total_fix_counts.get("weekly_bg", 0) + fc
            round_fixed += fc

            # 11. 大学平日偏り違反を修正
            current_df, _, fc = safe_fix(fix_university_weekday_balance_violations, current_df, max_attempts=150)
            total_fix_counts["univ_wd"] = total_fix_counts.get("univ_wd", 0) + fc
            round_fixed += fc

            # 11.5 全体の平日/休日偏り違反を修正
            current_df, _, fc = safe_fix(fix_weekday_weekend_balance, current_df, max_attempts=200)
            total_fix_counts["wd_we"] = total_fix_counts.get("wd_we", 0) + fc
            round_fixed += fc

            # 12. 公平性違反の修正
            current_df, _, fc = safe_fix(fix_fairness_imbalance, current_df, max_attempts=200)
            total_fix_counts["fairness"] = total_fix_counts.get("fairness", 0) + fc
            round_fixed += fc

            # 13. 未割り当てスロットを埋める（セーフティネット）
            current_df, _, fc = safe_fix(fix_unassigned_slots, current_df)
            total_fix_counts["unassigned"] = total_fix_counts.get("unassigned", 0) + fc
            round_fixed += fc

            # 14. 大学系週1違反の再修正（ステップ11-13で再発した違反をキャッチ）
            current_df, _, fc = safe_fix(fix_weekly_bg_violations, current_df, max_attempts=150)
            total_fix_counts["weekly_bg"] = total_fix_counts.get("weekly_bg", 0) + fc
            round_fixed += fc

            # 収束チェック: 修正がなければループ終了
            if round_fixed == 0:
                break

        # ── 収束ループ後の最終パス ──
        # safe_fixがrevertしたABS違反を最終的に解消する
        # ABS制約の優先度: gap/病院重複を先に修正 → 未割当を最後に埋める
        # safe_fixを通さず直接実行（ABS違反同士のトレードオフを許容）

        # 1) gap違反を修正（safe_fix不使用: 移動先が見つからず削除→未割当になっても許容）
        final_df, _, final_gap_fc = fix_gap_violations(current_df, max_attempts=200, verbose=False)
        total_fix_counts["gap"] = total_fix_counts.get("gap", 0) + final_gap_fc

        # 2) 外病院重複を修正
        final_df, _, final_dup_fc = fix_external_hospital_dup_violations(final_df, max_attempts=150, verbose=False)
        total_fix_counts["ext_dup"] = total_fix_counts.get("ext_dup", 0) + final_dup_fc

        # 2.5) 大学系週1違反を修正（gap/dup修正で発生した違反を含む）
        final_df, _, final_weekly_bg_fc = fix_weekly_bg_violations(final_df, max_attempts=150, verbose=False)
        total_fix_counts["weekly_bg"] = total_fix_counts.get("weekly_bg", 0) + final_weekly_bg_fc

        # 2.7) TARGET_CAP違反を修正（safe_fixでrevertされた分を含む）
        final_df, _, final_cap_fc = fix_target_cap_violations(final_df, max_attempts=100, verbose=False)
        total_fix_counts["cap"] = total_fix_counts.get("cap", 0) + final_cap_fc

        # 2.8) 平日/休日偏り違反を修正
        final_df, _, final_wd_we_fc = fix_weekday_weekend_balance(final_df, max_attempts=200, verbose=False)
        total_fix_counts["wd_we"] = total_fix_counts.get("wd_we", 0) + final_wd_we_fc

        # 3) 未割当スロットを埋める（gap/dup修正で発生した未割当を含む）
        final_df, _, final_unassigned_fc = fix_unassigned_slots(final_df, verbose=False)
        total_fix_counts["unassigned"] = total_fix_counts.get("unassigned", 0) + final_unassigned_fc

        # 4) 大学系週1違反の最終修正（未割当埋めで発生した違反をキャッチ）
        final_df, _, final_weekly_bg_fc2 = fix_weekly_bg_violations(final_df, max_attempts=150, verbose=False)
        total_fix_counts["weekly_bg"] = total_fix_counts.get("weekly_bg", 0) + final_weekly_bg_fc2

        fix_count = total_fix_counts.get("hard", 0)
        code_2_fix_count = total_fix_counts.get("code2", 0)
        cap_fix_count = total_fix_counts.get("cap", 0)
        univ_min_fix_count = total_fix_counts.get("univ_min", 0)
        ch_kate_fix_count = total_fix_counts.get("ch_kate", 0)
        gap_fix_count = total_fix_counts.get("gap", 0)
        bg_ht_fix_count = total_fix_counts.get("bg_ht", 0)
        ext_dup_fix_count = total_fix_counts.get("ext_dup", 0)
        univ_over_2_fix_count = total_fix_counts.get("univ_over2", 0)
        univ_weekday_fix_count = total_fix_counts.get("univ_wd", 0)
        fairness_fix_count = total_fix_counts.get("fairness", 0)
        unassigned_fix_count = total_fix_counts.get("unassigned", 0)

        # 修正後に再評価
        any_fixed = sum(total_fix_counts.values()) > 0
        if any_fixed:
            counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, *_ = recompute_stats(final_df)
            sc2, raw2, met2 = evaluate_schedule_with_raw(
                final_df, counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts
            )
            improved_df = final_df
        else:
            improved_df = final_df
    else:
        # 最適化無効時: fix_unassigned_slots のみ実行
        fix_count = fail_count = 0
        code_2_fix_count = cap_fix_count = univ_min_fix_count = 0
        ch_kate_fix_count = gap_fix_count = bg_ht_fix_count = 0
        ext_dup_fix_count = univ_over_2_fix_count = univ_weekday_fix_count = 0
        fairness_fix_count = 0

        final_df, unassigned_success, unassigned_fix_count = fix_unassigned_slots(
            improved_df, verbose=False
        )
        if unassigned_fix_count > 0:
            counts_tmp, bg_tmp, ht_tmp, wd_tmp, we_tmp, bk_tmp, ly_tmp, bg_cat_tmp, *_ = recompute_stats(final_df)
            sc2, raw2, met2 = evaluate_schedule_with_raw(
                final_df, counts_tmp, bg_tmp, ht_tmp, wd_tmp, we_tmp, bk_tmp, ly_tmp
            )

    # v5.7.1: 絶対禁忌の最終検証
    violations, is_valid = validate_absolute_constraints(final_df, verbose=False)

    refined.append({
        "seed": cand["seed"],
        "score_before": cand["score"],
        "raw_before": cand["raw_score"],
        "score_after": sc2,
        "raw_after": raw2,
        "metrics_after": met2,
        "pattern_df": final_df,  # v5.7.1: 最終パターンを使用
        "violations_fixed": fix_count,
        "violations_failed": fail_count,
        "code_2_violations_fixed": code_2_fix_count,
        "cap_violations_fixed": cap_fix_count,
        "univ_min_violations_fixed": univ_min_fix_count,
        "ch_kate_violations_fixed": ch_kate_fix_count,
        "bg_ht_imbalance_fixed": bg_ht_fix_count,
        "gap_violations_fixed": gap_fix_count,
        "external_dup_violations_fixed": ext_dup_fix_count,
        "univ_over_2_violations_fixed": univ_over_2_fix_count,
        "univ_weekday_violations_fixed": univ_weekday_fix_count,
        "fairness_violations_fixed": fairness_fix_count,
        "unassigned_slots_fixed": unassigned_fix_count,
        "absolute_constraints_valid": is_valid,  # v5.7.1: 絶対禁忌チェック結果
        "absolute_violations": violations,  # v5.7.1: 違反詳細
    })

# =========================
# v5.7.1: 絶対禁忌チェック結果の表示
# =========================
abs_valid_count = sum(1 for e in refined if e.get("absolute_constraints_valid", False))
abs_invalid_count = len(refined) - abs_valid_count
print(f"\n   絶対禁忌: {abs_valid_count}/{len(refined)} パターンクリア")

# ハード制約違反のないパターンのみ選択（TARGET_CAP、gap、未割当）
valid_patterns = []
excluded_count = 0
for e in refined:
    met = e["metrics_after"]
    cap_viol = met.get('cap_violations', 0)
    gap_viol = met.get('gap_violations', 0)
    unassigned = met.get('unassigned_slots', 0)
    code_2_viol = met.get('code_2_extra_violations', 0)
    bg_over_2_viol = met.get('bg_over_2_violations', 0)
    ht_0_viol = met.get('ht_0_violations', 0)
    wd_we_viol = met.get('wd_we_imbalance_violations', 0)
    we_0_viol = met.get('we_0_violations', 0)
    abs_valid = e.get("absolute_constraints_valid", False)
    # ch_kate_violationsはソフト制約（ペナルティのみ、ハード制約から除外）

    if not abs_valid:
        excluded_count += 1
    elif cap_viol > 0 or gap_viol > 0 or unassigned > 0 or code_2_viol > 0 or bg_over_2_viol > 0 or ht_0_viol > 0 or wd_we_viol > 0 or we_0_viol > 0:
        excluded_count += 1
    else:
        valid_patterns.append(e)

if not valid_patterns:
    print("   ⚠️ ハード制約を満たすパターンなし → 全パターンから選択")
    valid_patterns = refined
else:
    print(f"   ハード制約OK: {len(valid_patterns)}/{len(refined)} パターン")

# 評価軸1: 公平性重視（TARGET_CAP、公平性ペナルティを重視）
fairness_patterns = sorted(
    valid_patterns,
    key=lambda e: (
        -e["metrics_after"].get('cap_violations', 0) * 1000,  # TARGET_CAP違反を最優先で回避
        -e["metrics_after"].get('max_minus_min_total_active', 0) * 100,  # 公平性
        -e["metrics_after"].get('bg_ht_imbalance_violations', 0) * 50,
        e["raw_after"]
    ),
    reverse=True
)

# 評価軸2: gap違反回避重視（連続当直の間隔を重視）
gap_patterns = sorted(
    valid_patterns,
    key=lambda e: (
        -e["metrics_after"].get('gap_violations', 0) * 1000,
        -e["metrics_after"].get('external_hosp_dup_violations', 0) * 100,
        -e["metrics_after"].get('hospital_dup_violations', 0) * 50,
        e["raw_after"]
    ),
    reverse=True
)

# 評価軸3: バランス重視（大学/外病院、平日/休日のバランスを重視）
balance_patterns = sorted(
    valid_patterns,
    key=lambda e: (
        -e["metrics_after"].get('bg_ht_imbalance_violations', 0) * 1000,
        -e["metrics_after"].get('bg_weekday_weekend_imbalance', 0) * 100,
        -e["metrics_after"].get('bg_over_2_violations', 0) * 100,
        -e["metrics_after"].get('bg_weekday_over_violations', 0) * 100,
        e["raw_after"]
    ),
    reverse=True
)

# v6.0.2: 絶対禁忌クリアのパターンから3軸で多様な候補を選択
abs_valid_patterns = [e for e in valid_patterns if e.get("absolute_constraints_valid", False)]

if abs_valid_patterns:
    # 3軸評価で多様なパターンを選択（同じseedの重複を排除）
    # 軸1: 公平性重視
    fairness_abs = sorted(
        abs_valid_patterns,
        key=lambda e: (
            -e["metrics_after"].get('cap_violations', 0) * 1000,
            -e["metrics_after"].get('max_minus_min_total_active', 0) * 100,
            -e["metrics_after"].get('bg_ht_imbalance_violations', 0) * 50,
            e["raw_after"]
        ),
        reverse=True
    )
    # 軸2: gap・重複回避重視
    gap_abs = sorted(
        abs_valid_patterns,
        key=lambda e: (
            -e["metrics_after"].get('gap_violations', 0) * 1000,
            -e["metrics_after"].get('external_hosp_dup_violations', 0) * 100,
            -e["metrics_after"].get('hospital_dup_violations', 0) * 50,
            e["raw_after"]
        ),
        reverse=True
    )
    # 軸3: バランス重視
    balance_abs = sorted(
        abs_valid_patterns,
        key=lambda e: (
            -e["metrics_after"].get('bg_ht_imbalance_violations', 0) * 1000,
            -e["metrics_after"].get('bg_weekday_weekend_imbalance', 0) * 100,
            -e["metrics_after"].get('bg_over_2_violations', 0) * 100,
            e["raw_after"]
        ),
        reverse=True
    )

    axis_labels = ["公平性重視", "gap回避重視", "バランス重視"]
    axis_bests = [fairness_abs[0], gap_abs[0], balance_abs[0]]
    top_patterns = []
    seen_seeds = set()
    for label, best in zip(axis_labels, axis_bests):
        if best["seed"] not in seen_seeds:
            best["axis_label"] = label
            top_patterns.append(best)
            seen_seeds.add(best["seed"])
    # 3つに満たない場合はスコア順で補充
    if len(top_patterns) < 3:
        abs_valid_patterns.sort(key=lambda e: e["raw_after"], reverse=True)
        for e in abs_valid_patterns:
            if e["seed"] not in seen_seeds:
                e["axis_label"] = f"スコア上位"
                top_patterns.append(e)
                seen_seeds.add(e["seed"])
                if len(top_patterns) >= 3:
                    break
    print(f"   → {len(top_patterns)}パターンを出力")
else:
    # 絶対禁忌クリアのパターンがない場合は警告
    print(f"   ⚠️ 絶対禁忌クリアなし → 全パターンから上位3を選択（参考用）")
    valid_patterns.sort(key=lambda e: e["raw_after"], reverse=True)
    top_patterns = valid_patterns[:3]
    for i, p in enumerate(top_patterns):
        p["axis_label"] = f"参考{i+1}位（違反あり）"

# ソート済みリストも作成（後方互換性のため）
refined_sorted = sorted(valid_patterns, key=lambda e: e["raw_after"], reverse=True)
TOP_OUTPUT_PATTERNS = len(top_patterns)  # v6.0.0: 最大3パターン出力

scores_df = pd.DataFrame(score_rows).sort_values(["raw_score", "seed"], ascending=[False, True]).reset_index(drop=True)

refined_df = pd.DataFrame([
    {
        "seed": e["seed"],
        "score_before": e["score_before"],
        "raw_before": e["raw_before"],
        "score_after": e["score_after"],
        "raw_after": e["raw_after"],
        **{f"after_{k}": v for k, v in e["metrics_after"].items() if k not in ("raw_score", "penalty_total")},
    }
    for e in refined_sorted
]).sort_values(["raw_after", "seed"], ascending=[False, True]).reset_index(drop=True)

# =========================
# v6.0.0: 上位3パターン評価
# =========================
print("\n" + "="*60)
print("  📊 上位パターン評価")
print("="*60)

if top_patterns:
    print()
    for i, pattern in enumerate(top_patterns, 1):
        raw_score = pattern.get('raw_after', 0)
        fairness = pattern['metrics_after'].get('max_minus_min_total_active', 0)
        abs_valid = pattern.get('absolute_constraints_valid', False)
        abs_viols = len(pattern.get('absolute_violations', []))
        seed = pattern.get('seed', 0)
        abs_str = "ABS=OK" if abs_valid else f"ABS=NG({abs_viols}件)"
        print(f"  {i}位: スコア {raw_score:.0f} / 公平性 {fairness} / {abs_str} / seed={seed}")
else:
    print("\n  ⚠️ 有効なパターンが生成されませんでした")

# =========================
# 出力（pattern + summary + diagnostics）
# =========================
base_name = uploaded_filename.rsplit(".", 1)[0]
output_filename = f"{base_name}_v{VERSION}.xlsx"
_downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(_downloads_dir, exist_ok=True)
output_path = os.path.join(_downloads_dir, output_filename)

def _fmt_date_jp(d):
    """日付を 'YYYY/M/D (曜日)' 形式に変換"""
    _WD = ["月", "火", "水", "木", "金", "土", "日"]
    if pd.isna(d):
        return ""
    d = pd.to_datetime(d)
    return f"{d.year}/{d.month}/{d.day} ({_WD[d.weekday()]})"

def _fmt_date_cols(df):
    """DataFrame内の日付列を文字列フォーマットに変換"""
    df = df.copy()
    for col in df.columns:
        if "日付" in str(col):
            df[col] = df[col].apply(_fmt_date_jp)
    return df

def _str_display_width(s):
    """文字列の表示幅を推定（全角=2, 半角=1）"""
    import unicodedata
    width = 0
    for c in str(s):
        if unicodedata.east_asian_width(c) in ('F', 'W', 'A'):
            width += 2
        else:
            width += 1
    return width

def _auto_format_sheet(ws):
    """全セル中央揃え + 列幅を内容に合わせて自動調整"""
    from openpyxl.styles import Alignment
    center = Alignment(horizontal='center')
    col_max_width = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.alignment = center
                w = _str_display_width(cell.value)
                cl = cell.column_letter
                if w > col_max_width.get(cl, 0):
                    col_max_width[cl] = w
    for cl, w in col_max_width.items():
        ws.column_dimensions[cl].width = min(w + 2, 40)

def _format_summary_sheet(ws, sections):
    """サマリーシートに詳細なフォーマットを適用

    sections: list of dict, each with:
        - title_row: セクションタイトルの行番号 (1-indexed)
        - header_row: ヘッダー行番号
        - data_start: データ開始行
        - data_end: データ最終行
        - num_cols: 列数
        - section_type: 'summary' | 'violation' | 'score' | 'detail'
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # --- カラー定義 ---
    NAVY = "1F3864"
    LIGHT_BLUE_GREY = "D6E4F0"
    ZEBRA_GREY = "F2F2F2"
    WHITE = "FFFFFF"
    BORDER_GREY = "D9D9D9"
    VIOLATION_RED = "FFC7CE"
    SCORE_GREEN = "C6EFCE"
    SCORE_YELLOW = "FFEB9C"
    SCORE_RED = "FFC7CE"

    # --- スタイル定義 ---
    font_base = Font(name="MS Pゴシック", size=10)
    font_title = Font(name="MS Pゴシック", size=12, bold=True, color=WHITE)
    font_header = Font(name="MS Pゴシック", size=11, bold=True)
    fill_title = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    fill_header = PatternFill(start_color=LIGHT_BLUE_GREY, end_color=LIGHT_BLUE_GREY, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_GREY, end_color=ZEBRA_GREY, fill_type="solid")
    fill_white = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
    fill_violation = PatternFill(start_color=VIOLATION_RED, end_color=VIOLATION_RED, fill_type="solid")
    fill_score_green = PatternFill(start_color=SCORE_GREEN, end_color=SCORE_GREEN, fill_type="solid")
    fill_score_yellow = PatternFill(start_color=SCORE_YELLOW, end_color=SCORE_YELLOW, fill_type="solid")
    fill_score_red = PatternFill(start_color=SCORE_RED, end_color=SCORE_RED, fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color=BORDER_GREY),
        right=Side(style="thin", color=BORDER_GREY),
        top=Side(style="thin", color=BORDER_GREY),
        bottom=Side(style="thin", color=BORDER_GREY),
    )
    header_bottom_border = Border(
        left=Side(style="thin", color=BORDER_GREY),
        right=Side(style="thin", color=BORDER_GREY),
        top=Side(style="thin", color=BORDER_GREY),
        bottom=Side(style="medium", color="000000"),
    )

    # 全セクション最大列数
    max_cols = max((s["num_cols"] for s in sections), default=6)

    for sec in sections:
        title_row = sec["title_row"]
        header_row = sec["header_row"]
        data_start = sec["data_start"]
        data_end = sec["data_end"]
        num_cols = sec["num_cols"]
        sec_type = sec.get("section_type", "summary")

        # --- セクションタイトル行: ダークネイビー + 白文字 + 結合 ---
        for col_idx in range(1, num_cols + 1):
            cell = ws.cell(row=title_row, column=col_idx)
            cell.fill = fill_title
            cell.font = font_title
            cell.alignment = align_center
        if num_cols > 1:
            ws.merge_cells(start_row=title_row, start_column=1,
                           end_row=title_row, end_column=num_cols)

        # --- ヘッダー行: ライトブルーグレー + ボールド + 下線 ---
        for col_idx in range(1, num_cols + 1):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = header_bottom_border

        # --- データ行 ---
        for r in range(data_start, data_end + 1):
            row_idx = r - data_start  # 0-based index within data
            is_odd = (row_idx % 2 == 1)

            for col_idx in range(1, num_cols + 1):
                cell = ws.cell(row=r, column=col_idx)
                cell.font = font_base
                cell.border = thin_border

                # 氏名列（通常col=1）は左寄せ、それ以外は中央
                if col_idx == 1:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center

                # ゼブラストライプ（奇数行にグレー背景）
                if is_odd:
                    cell.fill = fill_zebra
                else:
                    cell.fill = fill_white

            # --- 違反テーブル: 違反がある行をハイライト ---
            if sec_type == "violation":
                # 違反行: 数値列で非0の値がある行は赤ハイライト
                has_violation = False
                for col_idx in range(2, num_cols + 1):
                    val = ws.cell(row=r, column=col_idx).value
                    if val is not None and str(val).strip() != "" and str(val).strip() != "0":
                        has_violation = True
                        break
                if has_violation:
                    for col_idx in range(1, num_cols + 1):
                        ws.cell(row=r, column=col_idx).fill = fill_violation

            # --- スコアテーブル: 値に応じた条件付き色分け ---
            if sec_type == "score":
                for col_idx in range(2, num_cols + 1):
                    cell = ws.cell(row=r, column=col_idx)
                    val = cell.value
                    if val is not None:
                        try:
                            num_val = float(val)
                            if num_val == 0:
                                cell.fill = fill_score_green
                            elif num_val <= 5:
                                cell.fill = fill_score_yellow
                            else:
                                cell.fill = fill_score_red
                        except (ValueError, TypeError):
                            pass

    # --- 列幅自動調整 ---
    col_max_width = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                w = _str_display_width(cell.value)
                cl = cell.column_letter
                if w > col_max_width.get(cl, 0):
                    col_max_width[cl] = w
    for cl, w in col_max_width.items():
        # 最小幅を確保しつつ、内容に合わせる
        auto_w = w + 3
        ws.column_dimensions[cl].width = max(min(auto_w, 50), 8)

    # A列（氏名）は広めに確保
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 14)

    # --- freeze_panes: 最初のセクションのヘッダー下で固定 ---
    if sections:
        first_data = sections[0].get("data_start", 3)
        ws.freeze_panes = f"A{first_data}"

def write_combined_summary_sheet(writer, sheet_name, df_month, df_total, diagnostics, df_doctors=None):
    """今月サマリー → 制約違反・スコア → 累計詳細 の順で配置（詳細フォーマット付き）"""
    ws = writer.book.create_sheet(sheet_name)
    writer.sheets[sheet_name] = ws

    sections = []  # フォーマット用セクション情報

    # === 1. 今月サマリー（Sheet2順）===
    COMPACT_COLS = ["氏名", "全合計", "大学合計", "外病院合計", "平日", "休日合計"]
    df_month_compact = df_month[COMPACT_COLS].copy()
    sheet2_order = {doc: i for i, doc in enumerate(doctor_names)}
    df_month_compact["_sort"] = df_month_compact["氏名"].map(sheet2_order)
    df_month_compact = df_month_compact.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

    title_row = 1
    ws.cell(row=title_row, column=1, value="【今月サマリー】")
    df_month_compact.to_excel(writer, sheet_name=sheet_name, startrow=title_row, index=False)
    header_row = title_row + 1  # to_excel writes header at startrow
    data_start = header_row + 1
    data_end = data_start + len(df_month_compact.index) - 1
    sections.append({
        "title_row": title_row,
        "header_row": header_row,
        "data_start": data_start,
        "data_end": data_end,
        "num_cols": len(COMPACT_COLS),
        "section_type": "summary",
    })
    startrow = data_end + 2

    # === 2. 制約違反 + スコアサマリー ===
    for title, df in diagnostics:
        if title == "【医師ごとの偏り】":
            continue
        title_row_cur = startrow
        ws.cell(row=title_row_cur, column=1, value=title)
        df_out = _fmt_date_cols(df)
        df_out.to_excel(writer, sheet_name=sheet_name, startrow=title_row_cur, index=False)
        header_row_cur = title_row_cur + 1
        data_start_cur = header_row_cur + 1
        data_end_cur = data_start_cur + len(df_out.index) - 1
        if len(df_out.index) == 0:
            data_end_cur = data_start_cur - 1

        # セクションタイプ判定
        if "スコア" in title:
            sec_type = "score"
        elif "違反" in title or "未割当" in title:
            sec_type = "violation"
        else:
            sec_type = "summary"

        sections.append({
            "title_row": title_row_cur,
            "header_row": header_row_cur,
            "data_start": data_start_cur,
            "data_end": data_end_cur,
            "num_cols": len(df_out.columns),
            "section_type": sec_type,
        })
        startrow = data_end_cur + 2

    # === 3. 累計詳細 ===
    detail_cols_available = [c for c in SUMMARY_DETAIL_COLS if c in df_total.columns]
    if detail_cols_available:
        df_detail = df_total[["氏名"] + detail_cols_available].copy()
        title_row_det = startrow
        ws.cell(row=title_row_det, column=1, value="【累計詳細】")
        df_detail.to_excel(writer, sheet_name=sheet_name, startrow=title_row_det, index=False)
        header_row_det = title_row_det + 1
        data_start_det = header_row_det + 1
        data_end_det = data_start_det + len(df_detail.index) - 1
        sections.append({
            "title_row": title_row_det,
            "header_row": header_row_det,
            "data_start": data_start_det,
            "data_end": data_end_det,
            "num_cols": len(df_detail.columns),
            "section_type": "detail",
        })

    # === 4. 詳細フォーマット適用 ===
    _format_summary_sheet(ws, sections)


with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # v6.3.0: sheet1-4は出力しない（元データ不要）

    # TOPパターン出力
    for rank, entry in enumerate(top_patterns, start=1):
        axis_label = entry.get('axis_label', '総合スコア')
        sheet_label = f"pattern_{rank:02d}"

        # パターンシート（Date列を曜日付き文字列に変換）
        pdf = entry["pattern_df"].copy()
        pdf[date_col_shift] = pd.to_datetime(pdf[date_col_shift]).apply(_fmt_date_jp)
        pdf.to_excel(writer, sheet_name=sheet_label, index=False)

        # シート名に軸ラベルを追加
        ws = writer.sheets[sheet_label]
        axis_short = {"公平性重視": "公平性", "連続当直回避重視": "gap回避", "バランス重視": "バランス", "総合スコア": "総合"}.get(axis_label, axis_label)
        ws.cell(row=1, column=len(entry["pattern_df"].columns) + 2, value=f"【{axis_short}】")

        # v6.3.0: 今月/累計/診断を1シートに統合
        counts, bg_counts, ht_counts, wd_counts, we_counts, bk_counts, ly_counts, bg_cat, *_ = recompute_stats(entry["pattern_df"])
        df_month, df_total = build_summaries(entry["pattern_df"], counts, bg_counts, ht_counts, wd_counts, we_counts, bg_cat)
        df_doctors, df_gap, df_same, df_hdup, df_weekly_bg, df_unass, df_metrics, df_hard_violations = build_diagnostics(entry["pattern_df"])

        write_combined_summary_sheet(
            writer,
            sheet_name=f"{sheet_label}_summary",
            df_month=df_month,
            df_total=df_total,
            diagnostics=[
                ("【医師ごとの偏り】", df_doctors),  # v6.5.3: 上部テーブルに統合されるためスキップ
                ("【制約違反: gap（3日未満）】", df_gap),
                ("【制約違反: 同日重複】", df_same),
                ("【制約違反: 同一病院重複】", df_hdup),
                ("【制約違反: 大学系週1違反】", df_weekly_bg),  # v6.4.0: ABS-012
                ("【制約違反: 未割当枠】", df_unass),
                ("【制約違反: 重要/推奨ルール】", df_hard_violations),
                ("【スコアサマリー】", df_metrics),
            ],
            df_doctors=df_doctors,
        )

    # 全シート: 中央揃え + 列幅自動調整（summaryシートは個別フォーマット済み）
    for ws in writer.book.worksheets:
        if not ws.title.endswith("_summary"):
            _auto_format_sheet(ws)

print("\n" + "="*60)
print("  完了")
print("="*60)
print(f"\n出力ファイル: {output_path}")
print("\n【内容】")
for rank in range(1, len(top_patterns) + 1):
    label = f"pattern_{rank:02d}"
    print(f"  {label}: スケジュール / {label}_summary: サマリー・診断")
print("="*60)
