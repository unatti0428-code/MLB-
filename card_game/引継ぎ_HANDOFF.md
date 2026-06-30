# MLBカードベースボール ゲーム 引継ぎメモ

最終更新: このファイルは会話コンテキスト引継ぎ用。次の担当（AIアシスタント）が状況を即把握するための要約。
※前回からの大きな差分: レギュラーシーズンを **30球団・6地区化＋ポストシーズン** に拡張、**手動試合への動画演出** を追加。

---

## 0. 重要な前提（毎回守ること）

- **ユーザーは非エンジニア**。専門用語を避け、平易な日本語で説明する。リスクは先回りで指摘する。
- ゲーム本体ファイルの場所: `C:\Users\unatt\OneDrive\デスクトップ\MLB\card_game\`
  - `index.html` … エントリ。`<link ... style.css?v=NN>` と `<script ... game.js?v=NN>` の **バージョン番号でキャッシュ制御**。
  - `game.js` … ゲーム全ロジック（巨大IIFE。関数は基本 private）。
  - `style.css` … スタイル。
  - `card_store.js` / `card_parser.js` / `card_view.js` … カード保存・解析・カード表示。
- **このフォルダ（card_game）は OneDrive 上で、長らくGit管理外だった**。今セッションで GitHub の Private リポジトリ `MLB-`（`C:\Users\unatt\Documents\GitHub\MLBテスト`、remote=unatti0428-code/MLB-）に **`card_game/` フォルダとして1回コミット済み**（commit 80e790b, branch=master）。**それ以降の変更は未コミット**（保存したい時はユーザーの指示で再コミット）。
- カード生成ツール: `C:\Users\unatt\Documents\GitHub\MLBテスト\card_generator.html`（GitHub管理下・コミット可）。所属コードが `/^ALL\d*$/` の選手はチーム名を「ALL」と表示。

## 1. 変更時の必須手順（毎回）

1. `game.js` または `style.css` を編集したら、**`index.html` の該当 `?v=` を必ず +1**。
   - 現在: **`game.js?v=318` / `style.css?v=145`**（次に上げるならそれぞれ +1）。
2. 検証は preview ツールで行う（`preview_start(name="game-preview")` / port 3456。card_game フォルダを配信）。
   - `preview_eval` で `location.href='http://localhost:3456/index.html?bust='+Date.now()` してハードロード。
   - **private関数の検証は一時的に `window.__xxxTest = {...}` フックを追記** → preview_eval で呼ぶ → **検証後にフックを必ず削除**（`[TEMP]` コメント付きで追加し、最後に消す運用）。検証後 `localStorage.removeItem('mlb_season_v1')` でテスト用シーズンを掃除。
   - サーバは落ちやすい。`preview_screenshot` がよくタイムアウトするので、`preview_eval` で `getComputedStyle`/DOM を読んで検証する手も使う。
3. 完了報告で「ハード再読み込み（Ctrl+Shift+R）で v=NN を反映してください」と必ず添える。
4. **テストフックの消し忘れ厳禁**（grep で `__` や `[TEMP]` を確認して消す）。
5. **動画演出はpreviewサーバでは再生できない**（`../douga` 等を配信しないため）。動画の実再生確認は「ユーザーが index.html を file:// で開く実機」で行う旨を伝える。ロジック（src生成・分岐）は preview で検証できる。

## 2. 主要データ構造（要点）

- `playerKey(p)` = `名前_年_チーム_種別(batter/pitcher)`。`normalizeTeam(t)` で別名吸収。
- チーム編成保存: localStorage `mlb_team_build_v1_<TEAM>`（orders[3] に batters/batterOrder/pinchHitters、pitchers{starter,mop,middle,setup,closer,bench}）。
- シーズン状態: localStorage `mlb_season_v1`（`SEASON`）。**`format:3`** で30球団制を識別（旧 format:2=10球団 / 1リーグ6球団は読込時に破棄）。
- カード写真キャッシュは card_generator.html 側で IndexedDB(`mlb_photo_db`)。

## 3. レギュラーシーズン＝30球団・2リーグ×3地区（現行 format:3）

- `SEASON_DIVISIONS`（6地区）/ `SEASON_DIV_ORDER` / `SEASON_LEAGUE_DIVS` を新設。地区割りは `MLB_DIVISIONS`(game.js:92) と同一。
  - AL東=NYY,BOS,TB,TOR,BAL / AL中=CWS,CLE,DET,KC,MIN / AL西=ATH,HOU,LAA,SEA,TEX
  - NL東=ATL,MIA,NYM,PHI,WSH / NL中=CHC,CIN,MIL,PIT,STL / NL西=ARI,COL,LAD,SD,SF
- `SEASON_LEAGUES.AL/NL` = 各15球団、`SEASON_TEAMS` = 30。`seasonLeagueOf` / 新規 `seasonDivisionOf`。
- 日程 `seasonGenerateSchedule()`: 各チーム**162試合** = 同地区52(各13)＋同リーグ他地区64(4相手×7+6相手×6)＋インターリーグ46(ライバル1組×4+他14×3)。総2430試合。
- 順位表 `seasonStandingsTable`→6地区別 `seasonStandingsTableFor(divKey)`。対戦成績 `seasonH2HTable`→リーグ別(AL/NL 15×15)。
- 打撃/投手ベスト20・アワードは従来通りリーグ別（15球団）。
- **未保存チームの自動生成** `seasonEnsureBuild(teamCode)`: 保存編成が無いチームは、`TB_STATE`/`renderTeamBuild` を一時退避し `autoFillTeamBuild` をヘッドレス実行→`saveTeamBuild`。試合開始時(`seasonStartCurrentGame`/`seasonStartPostseasonManual`/`playPostseasonGame`)で呼ぶ。カード不足チームは従来通りエラー。

## 4. ポストシーズン（MLB方式12球団）

- `SEASON.postseason = { seeds:{AL[6],NL[6]}, series:{}, champion, bat:{}, pit:{}, wsGames:[], mvp }`。`seasonPostseasonInit`/`Ensure`。
- シード: 各リーグ 地区優勝3＋ワイルドカード3。WC(3戦2勝)→地区(5戦3勝)→リーグ優勝(7戦4勝)→WS(7戦4勝)。
- ブラケット解決: `seasonPostseasonBracket()`（依存解決）/ `seasonPostseasonNextGame()`（次の1戦）/ `seasonPostseasonApplyResult()`（勝数加算・先取で完了・WS完了で champion+MVP）。
- 自動: `playPostseasonGame`(記録込み) / `seasonPostseasonStepAuto` / `seasonPostseasonPlayAll`。手動: `seasonStartPostseasonManual`（結果は `seasonAfterManualResult` のポストシーズン分岐で記録）。
- 成績記録: `seasonRecordPostseasonGame`（`seasonAccumBatting/Pitching` を store引数で `SEASON.postseason.bat/pit` に蓄積。`seasonUpdateStamina` も呼ぶので**試合またぎのスタミナ**が成立）。WSは `seasonCaptureWSGame` でボックス保存（line/box/HR号数/W-L-S）。`seasonComputeWSMVP`。
- 画面: `seasonPostseasonHtml`（サブタブ: トーナメント表 / ポストシーズン成績 / WS記録）。
  - **トーナメント表 = SVG** `seasonPsBracketSvg`（AL左/NL右・中央トロフィー＋WS。`.ps-svg{max-height:58vh}` でスクロール無しに収まる）。
  - **ポストシーズン成績** `seasonPsStatsHtml`: 列クリックでソート(`SEASON_PS_SORT`・`data-pssort`)、チーム別タブ(`SEASON_PS_TEAM`・`data-psteam`)、選手名は `seasonNameCell` でカードリンク化。
  - **WS記録** `seasonPsWSHtml`: 各戦ラインスコア＋ボックス。ポ打率/ポ防＋ポ本/ポ点/ポ勝/ポ敗/ポH/ポS/ポ奪三（赤字＝ポストシーズン通算、`ps-cum`/`ps-n`クラス）。HRはポストシーズン通算号数。選手名カードリンク。
- 表示用成績の切替: `seasonActiveStores()`（ポストシーズン文脈なら postseason 成績を返す）。`seasonBatStatObj`/`seasonPitStatObj`/`seasonBatStatLine`/`seasonPitStatLine`/`seasonLiveBatAgg`/`seasonLivePitAgg`/`seasonHrNumbers`/`seasonPitcherRecordText` が利用 → 手動ポストシーズン試合の編成画面・結果画面・ライブ表示がポストシーズン積み上げになる。

## 5. 手動試合の動画演出（今セッションの新機能）

- 手動で球種(`.pitch-btn`)を選ぶと、`playPitchVideo()` がゲーム左側(`.game-left`)に固定overlay(`#pitchVideoOverlay`, z2000)を出し、動画を再生→終了で撤去。
- **動画の保存先（重要・接頭辞で切替）**: `videoSrc(file)` →
  - `laa_*` … `C:\Users\unatt\OneDrive\デスクトップ\MLB\team_out\laa_out\`（相対 `../team_out/laa_out/`）。
  - `defopit_*` / `defobat_*` … `C:\...\MLB\douga\`（相対 `../douga/`）。
- 再生フロー（`batResultClip(res, runs)` が分岐）:
  - **イントロ無し・単独動画**（ヒット系/三振/エラー/ファインプレー）: HR=defobat_hr*, 2/3塁=defobat_2b*, 内野安打=defobat_if1b*, タイムリー(1B+得点)=defobat_rbi*, 通常ヒット=defobat_1b*, 四球=defobat_fb*, エラー=defobat_miss*, 三振=defopit_str*, ファインプレー=defopit_nice*。
  - **イントロ(defopit_tou*ランダム)→結果動画**（各種アウト, laa_*）: フォースアウト=laa_forth2/3/4_*, 併殺=laa_double*, 深い飛球(SAC_FLY)=laa_flybig*, 内野ポップ=laa_pop/pop1/pop2, ファール=laa_foul*, 外野フライ(守備位置別 左7/中8/右9)=laa_7/8/9out*, ライナー=laa_line*, 緩いゴロ(GO_SLOW/力ない)=laa_soft*(一soft/三soft1/遊soft2/二soft3), 強い(詰まった)ゴロ=laa_strong*(一/二/三/遊), 通常内野ゴロ(守備位置別)=laa_3/4/5/6out*。
- **連続再生の固まり対策**: 2本目は `ended` 内再生でブラウザ自動再生制限を受けるため、`playClip` 内で `play()` 失敗時に **muted で再生継続**（それでも不可なら次へ）。1本目(イントロ/単独)はクリック直後＝ユーザー操作内なので音あり再生可。
- 守備位置取得: `effectiveFielderPos(res)`。

## 6. フォースアウト（野手選択）機構（今セッション新規・ゲームロジック）

- 「強くも弱くもない内野ゴロ」(`outcome:'GO'` かつ flavor が `/内野ゴロ/` で `ボテボテ|力ない|詰まった` を含まない) かつ1塁走者あり・非併殺(DP35%判定の後)の時、`applyOutcome` の GO 処理で発動。
- 封塁チェーン(1塁から連続する埋まった塁)をリード(高い塁)から走力依存(`outProb=clamp(0.7-(spd-60)*0.012,0.15,0.9)`)で封殺判定。封殺できた走者の起点塁を `outBase` とし `lpr.forceOut = outBase+2`(2/3/4)。打者は1塁セーフ(`meRef`)、他走者は1つ進塁。全員セーフなら打者1塁アウト(通常)。
- 例(1・2塁): 2塁走者速い→三塁到達で二塁封殺/遅い→三塁封殺＋走者1,2塁/全員速い→打者1塁アウト＋走者2,3塁。満塁は本塁封殺(forth4)。
- DPは従来通り `outcome='GO_DP'`(35%)。フォースアウトは `outcome='GO'` のまま `lpr.forceOut` フラグで識別（renderBattedBall等を壊さないため）。

## 7. 既存の主要ルール（変更なし系・要点）

- 打席は1球完結。`decidePitchOutcome` が `res{outcome,flavor,fielderPos,fineplay,infieldFly,...}` を返し `applyOutcome` が反映。outcome種別: HR/3B/2B/1B/BB/K/GO/GO_SLOW/GO_DP/FO/LO/SAC_FLY/E。
- 休養/怪我(野手)・投手疲労スワップ・延長15回・捕手評価・オーダー特例しきい値などは従来通り。
- 自動編成 `autoFillTeamBuild`: マイナー選手(名前が「マイナー」始まり)は `MINOR_FILL_PENALTY` で大きく減点し通常選手を優先・穴埋めのみ。`isMinorPlayer`(名前ベース)/`isAllTeamPlayer`。所属ALLは全チーム・全年度で登録可。
- 総合力バッジ/カードのレアリティ色: `cardRarity`(UR100+/SSR90/SR80/RR70/R60/N50/N相当50未満)。チーム編成の `ovrBadge` も同色。

## 8. 既知の注意点 / 今後の候補

- preview サーバが落ちやすい → 再起動運用。screenshot タイムアウト時は eval で代替検証。
- 動画は実機(file://)でのみ再生確認可。音は1本目あり/連続2本目はミュートフォールバックの可能性。
- 30球団自動進行は全30球団分のカードが要る。カード不足チームはエラー（自動生成でも9人組めない）。
- ポストシーズンの旧セーブに新フィールドが無い場合は遅延初期化。確実に全記録を残すには「やり直す」。
- `card_game` の GitHub 反映は手動（ユーザー指示時に再コミット）。最後のコミットは 80e790b（30球団化＋ポストシーズン時点）。以降の動画・微修正は未コミット。

## 9. バージョン管理メモ

- 現在: **game.js v=318 / style.css v=145**。次の編集で必ず +1。
- GitHub: リポジトリ `MLB-`(Private) の `card_game/` に 80e790b で1回保存済み。**それ以降は未コミット**。
