# MLBカードベースボール ゲーム 引継ぎメモ

最終更新: このファイルは会話コンテキスト引継ぎ用。次の担当（AIアシスタント）が状況を即把握するための要約。

---

## 0. 重要な前提（毎回守ること）

- **ユーザーは非エンジニア**。専門用語を避け、平易な日本語で説明する。リスクは先回りで指摘する。
- ゲーム本体ファイルの場所: `C:\Users\unatt\OneDrive\デスクトップ\MLB\card_game\`
  - `index.html` … エントリ。`<link ... style.css?v=NN>` と `<script ... game.js?v=NN>` の **バージョン番号でキャッシュ制御**。
  - `game.js` … ゲーム全ロジック（巨大IIFE。関数は基本 private）。
  - `style.css` … スタイル。
  - `card_store.js` … カード保存（IndexedDB、localStorageフォールバック）。
- **このフォルダ（card_game）はGit管理外**。編集してもコミット対象にならない（ファイル保存のみ）。
- カード生成ツール: `C:\Users\unatt\Documents\GitHub\MLBテスト\card_generator.html`（**こちらはGitHub Private リポジトリ管理下**。コミット可）。

## 1. 変更時の必須手順（毎回）

1. `game.js` または `style.css` を編集したら、**`index.html` の該当 `?v=` を必ず +1**（HMR等は無いためキャッシュ更新に必須）。
   - 現在: **`game.js?v=295` / `style.css?v=134`**（次に上げるならそれぞれ +1）。
2. 検証は preview ツールで行う:
   - `preview_start(name="game-preview")`（server.js / port 3456。card_game フォルダを配信）。
   - `preview_eval` で `location.href='http://localhost:3456/index.html?bust='+Date.now()` してハードロード。
   - **private関数の検証は一時的に `window.__xxxTest = {...}` フックを game.js に追記** → preview_eval で呼ぶ → **検証後にフックを必ず削除**（`[TEMP]` コメント付きで追加し、最後に消す運用）。
   - `window.alert` 等はモーダルで preview_eval を固める → テスト前に `window.alert=function(){}` 等で無効化。
   - サーバは頻繁に落ちるので `Server not found` が出たら `preview_start` で再起動。
3. 完了報告では「ハード再読み込み（Ctrl+Shift+R）で v=NN を反映してください」と必ず添える。
4. **テストフックの消し忘れ厳禁**（`window.__...Test` をリリースに残さない）。

## 2. 主要データ構造（要点）

- `playerKey(p)` = `名前_年_チーム_種別(batter/pitcher)`。
- `normalizeTeam(t)` で別名（SFG/SFN→SF 等）を吸収。
- チーム編成保存: localStorage `mlb_team_build_v1_<TEAM>`（orders[3] に batters/batterOrder/pinchHitters、pitchers{starter,mop,middle,setup,closer,bench}）。
- シーズン状態: localStorage `mlb_season_v1`（`SEASON`）。**`format:2`** で2リーグ制を識別（旧1リーグ制は読み込み時に破棄）。
- カード写真の再利用キャッシュは **card_generator.html 側で IndexedDB(`mlb_photo_db`)** に保存（以前 localStorage を圧迫していたのを移行済み）。

## 3. このセッションで実装した主な変更（時系列の要点）

- カード写真キャッシュを localStorage → **IndexedDB へ移行**（card_generator.html、コミット済み `969b303`）。ゲーム側 saveTeamBuild は容量超過時に写真キャッシュを掃除して再試行。
- **休養/怪我**を多ポジション選手でも確実に反映（後述の連鎖補完）。表記は「休養」→**「休養/怪我」**。
- **延長戦**: 最大12回→**15回**。延長の守備で抑え/SUが残率70%以上なら **抑え→SU1→SU2** の順に勝ち優先継投。
- チーム編成の**守備プルダウン**: 控え・**他守備位置のスタメン**も候補に。選択で入替（不可なら元位置を空に）。**AI推奨順**で並び、**ホバーでDRS**表示。
- チーム編成中央下に**守備位置別の登録可能人数**（捕/一/二/三/遊/左/中/右/全）を表示。凡例とDH枠の重なりも解消（`.tb-field`+`.tb-diamond-legend`）。
- 捕手評価: レギュラー査定に **リード×5pt・阻止率/10pt**、守備固め評価は **DRS/5 + リード×5 + 阻止率/10**（`catcherAssessBonus` / `defenseRating`）。
- 投手控え疲労スワップ: 控え1〜4全員が残率30%未満で控え3,4を**余剰投手と一時入替**、全員60%回復で復帰（`seasonEffectiveBenchPitchers`、`SEASON.pitBenchSwap`）。
- **2リーグ制**化（最大の変更）。下記参照。
- アワード/成績の**リーグ別表示**＋**列幅レイアウト調整**（下記）。

## 4. レギュラーシーズン＝2リーグ制（現行仕様）

- `SEASON_LEAGUES = { AL:['NYY','BOS','TB','TOR','BAL'], NL:['ARI','COL','LAD','SD','SF'] }`、`SEASON_TEAMS`=10球団。`seasonLeagueOf(team)`。
- 日程 `seasonGenerateSchedule()`: 各チーム**162試合**（同リーグ122＋交流戦40＝約24.7%）。総計810試合。同リーグは各ペア30＋5角形サイクル+1、交流戦はAL×NL各ペア8試合。位置を均等分散。
- **勝敗表**: AL/NLを別表（`seasonStandingsTableFor`）。CSS `.season-standings`（table-layout:fixed、順位46px・チーム名156px固定）で両リーグの列を一致。
- **打撃/投手ベスト20**: リーグ別セクション（`seasonBatLeaders`/`seasonPitLeaders`）。名前=`srow-name`(min/固定120〜140)、チーム=`srow-team`(min-width 150px=最長「ダイヤモンドバックス」基準)で AL/NL の列を揃え、文字は切らない。
- **アワード**: リーグ別（`seasonAwardsHtml`→`seasonAwardsForLeague(lg,...)`）。AL=赤/NL=青の**バナー**＋各カード右上に**AL/NLチップ**（`.award-league`, `.award-lg-chip`）。個人タイトル・シルバースラッガーを各リーグで選出。
- 注意: **自動シーズンは全10球団のチーム編成が保存済み**である必要（未保存はその試合でエラー）。トップメニュー説明は「2リーグ10球団・162試合」。

## 5. オーダー設定の特例（現行しきい値）※ユーザーが何度も調整した

`playerAllowedInOrder(p, orderIdx)`（game.js ~1877）:
- 基本: 出場試合数 `ordersAllowedByGames(games)`（150+→全/125+→1,2/100+→1,3/80+→1/60+→2,3/40+→2/未満→3、不明→全）。
- **全オーダー可の特例**:
  - 総合力 **59以下**
  - **ユーティリティ（守備DRS3ポジション以上）かつ総合力63以下**
  - 捕手（C守備可）かつ総合力59以下

`tbYearAllowed`/`seasonYearAllowed`: 指定年度のみ。ただし総合力（オーダー1=70以下/2,3=65以下）なら2年前まで可。同名年代違いは最新年度のみ（`seasonHasNewerSameName`）。

## 6. 休養/怪我の条件（現行）※野手。`seasonRestProb(p, perTeam=162)`（game.js ~6695）

- 休養なし: 総合力59以下 / **ユーティリティ(3ポジション可)かつ総合力63以下** / 年度試合数不明 or 162以上。
- それ以外: 休養率 = `max(0, min(0.95, 1 - (g/162)/P))`、P=出場可能オーダーの使用率合計(1=0.6,2=0.3,3=0.1)。総合力60〜65は休養率を1/2、66以上は通常。
- 休養者はスタメンにも控えにも入らない。**空き守備位置は「控え/余剰で直接補完」→不可なら「他スタメンを移動し元位置を控えで補完」(連鎖補完)** で必ず9人埋める（`buildSeasonSideSetup` 内）。DP全再割当は重いので軽量な連鎖方式。

## 7. その他の現行ルール（投手疲労など）

- リリーフ: スタミナ残率25%未満は当該試合休養（中継/MUは控え2,3,4から補充）。
- 投手控え1〜4の疲労スワップ（第4章 §3末尾）。
- 自動編成 `autoFillTeamBuild`: `diamondScore`=総合力+起用ボーナス(試合数/10)+守備微加点+捕手査定。`bestAssignment`/`tbTrimForAssignment`。

## 8. 既知の注意点 / 今後の候補

- preview サーバが落ちやすい → 再起動運用。
- `document.body.innerHTML=''` で検証すると script タグが消えて誤判定するので、検証描画は `position:fixed` のオーバーレイ等に出す。
- h2h（対戦成績）は10×10で横長（必要なら将来リーグ別分割）。
- 旧6球団シーズンのセーブは2リーグ制と非互換 → 自動破棄。ユーザーは新シーズン開始が必要。

## 9. バージョン管理メモ

- 現在: **game.js v=295 / style.css v=134**。次の編集で必ず +1。
- card_generator.html の写真IndexedDB移行のみ GitHub にコミット済み（`969b303`）。それ以降のゲーム側変更は **未コミット（Git管理外フォルダ）**。
