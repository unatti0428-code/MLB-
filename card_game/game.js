// =============================================================
//  MLBカード野球ゲーム ロジック本体
//  - window.PLAYERS をデータソースとして読み込み (players.js)
//  - 試合進行：1球毎にストライク/ボール/打球結果を判定
//  - 既存ゲーム(yakyu.js)の能力値スキーマと整合した数式
// =============================================================

(function(){
'use strict';

// ============== ユーティリティ ==============
const $  = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
const rand  = ()  => Math.random();
const randI = (n) => Math.floor(Math.random() * n);
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

// 球種短縮 ⇔ 球種フル名 マッピング
// FB=フォーシーム / 2C=ツーシーム / CT=カットボール / SL=スライダー
// CB=カーブ / CH=チェンジアップ / SF=シンカー
const PITCH_SHORT_TO_FULL = {
  'FB': ['フォーシーム'],
  '2C': ['ツーシーム'],
  'CT': ['カットボール'],
  'SL': ['スライダー'],
  'CB': ['カーブ'],
  'CH': ['チェンジアップ','サークルチェンジ'],
  'SF': ['シンカー','スプリット','スプリッター','フォーク'],
};
const PITCH_FULL_TO_SHORT = (function(){
  const m = {};
  for (const sh in PITCH_SHORT_TO_FULL) {
    for (const f of PITCH_SHORT_TO_FULL[sh]) m[f] = sh;
  }
  return m;
})();
function fullPitchToShort(name) {
  for (const k in PITCH_FULL_TO_SHORT) {
    if (name.includes(k)) return PITCH_FULL_TO_SHORT[k];
  }
  return 'FB';
}
// 「魔球フォーシーム」「代名詞のカーブ」のような肩書/前置きを除いて
// 球種コア名 (フォーシーム/カーブ/スライダー…) のみを返す。
// 試合中の球種選択ボタンに表示する用。data-pitch には元の名前を残すこと。
function shortenPitchName(name) {
  if (!name) return name;
  for (const k in PITCH_FULL_TO_SHORT) {
    if (name.includes(k)) return k;
  }
  return name;
}
// カード表示用の年(シーズン): 「2008 PEAK」「2022 MVP」等から年だけを取り出す
function cardSeason(p) {
  if (p && p.year != null && p.year !== '') return p.year;
  const m = String((p && p.seasonLabel) || '').match(/\d{4}/);
  return m ? m[0] : '';
}
// 球種名から通称(括弧書き、例:「チェンジアップ（エアベンダー）」)を除去
function stripPitchAlias(name) {
  return String(name || '').replace(/[（(][^）)]*[）)]/g, '').trim();
}

// ============== グローバル状態 ==============
const G = {
  setup: {
    // 投手は配列（先発 + リリーフ）、スタミナも配列で管理
    away: { pitchers: [], pitcherStamina: [], pitcherMax: [], activeIdx: 0, batters: [], batterPos: [] },
    home: { pitchers: [], pitcherStamina: [], pitcherMax: [], activeIdx: 0, batters: [], batterPos: [] },
  },
  innings: 9,
  // 試合中
  inning: 1, top: true, outs: 0,
  bases: [null, null, null], // [1, 2, 3] それぞれ runner(player) or null
  score: { away: [], home: [] },  // score[side][inning-1] = runs
  hits: { away: 0, home: 0 },
  ks:   { away: 0, home: 0 },
  awayBatIdx: 0, homeBatIdx: 0,
  currentPitcher: null,
  currentBatter:  null,
  ended: false,
  awaitingResult: false,  // 試合終了後、結果画面へ進む前の一時停止状態
  // 追跡データ
  hrEvents: [],         // [{ inning, top, batter, pitcher, runs, side }]
  pitcherLog: { away: [], home: [] },  // 各投手の登板情報 [{pitcher, runsAllowed, batters, leadAtEnter, leadAtExit}]
  subLog: { away: [], home: [] },      // 代打/代走/守備固めで退いた選手の成績スナップショット
  leadHistory: [],      // [{inning, top, leadSide, diff}] 各打席後のリードチーム
};

// ============== 初期化 ==============
// MLB 30球団 (6地区) — チーム選択プルダウンを地区別 optgroup で生成する
const MLB_DIVISIONS = [
  { label: 'アメリカンリーグ東地区', teams: [['NYY','ヤンキース'],['BOS','レッドソックス'],['TB','レイズ'],['TOR','ブルージェイズ'],['BAL','オリオールズ']] },
  { label: 'アメリカンリーグ中地区', teams: [['MIN','ツインズ'],['CLE','ガーディアンズ'],['CWS','ホワイトソックス'],['KC','ロイヤルズ'],['DET','タイガース']] },
  { label: 'アメリカンリーグ西地区', teams: [['HOU','アストロズ'],['LAA','エンジェルス'],['ATH','アスレチックス'],['SEA','マリナーズ'],['TEX','レンジャーズ']] },
  { label: 'ナショナルリーグ東地区', teams: [['ATL','ブレーブス'],['MIA','マーリンズ'],['NYM','メッツ'],['PHI','フィリーズ'],['WSH','ナショナルズ']] },
  { label: 'ナショナルリーグ中地区', teams: [['CHC','カブス'],['CIN','レッズ'],['MIL','ブルワーズ'],['PIT','パイレーツ'],['STL','カージナルス']] },
  { label: 'ナショナルリーグ西地区', teams: [['ARI','ダイヤモンドバックス'],['COL','ロッキーズ'],['LAD','ドジャース'],['SD','パドレス'],['SF','ジャイアンツ']] },
];
// チーム選択肢の HTML (オリジナル + 地区別 optgroup)
function teamOptionsHtml(originalLabel) {
  let html = `<option value="original">${originalLabel}</option>`;
  for (const div of MLB_DIVISIONS) {
    html += `<optgroup label="${div.label}">`;
    for (const [code, name] of div.teams) html += `<option value="${code}">${name} (${code})</option>`;
    html += '</optgroup>';
  }
  return html;
}
// 3つのチーム選択(チーム編成 / 先攻 / 後攻)を 30球団 + 地区分け で構築
function populateTeamSelects() {
  const tb = document.querySelector('#tb-team-select');
  if (tb) tb.innerHTML = teamOptionsHtml('オリジナル(全選手)');
  document.querySelectorAll('.team-filter').forEach(sel => {
    sel.innerHTML = teamOptionsHtml('オリジナル（全選手）');
  });
}

function init() {
  populateTeamSelects();   // チーム選択肢 (30球団) を先に構築
  buildSetup();
  // スタート画面のモード選択ボタン
  const modeExh = $('#mode-exhibition');
  if (modeExh) modeExh.addEventListener('click', () => showScreen('setup'));
  const modeTeam = $('#mode-team');
  if (modeTeam) modeTeam.addEventListener('click', () => { showScreen('teambuild'); initTeamBuild(); });
  // レギュラーシーズン
  const modeSeason = $('#mode-season');
  if (modeSeason) modeSeason.addEventListener('click', openSeason);
  const seasonBody = $('#season-body');
  if (seasonBody) seasonBody.addEventListener('click', seasonHandleClick);
  if (seasonBody) seasonBody.addEventListener('change', seasonHandleChange);
  // 打順ドラッグ&ドロップ (⠿ ハンドルをドラッグ → 行へドロップで打順入替)
  if (seasonBody) {
    seasonBody.addEventListener('dragstart', e => {
      const h = e.target.closest('.season-linedrag'); if (!h) return;
      const row = h.closest('.season-linerow'); if (!row) return;
      SEASON_DRAG = { side: row.dataset.slinSide, fromIdx: +row.dataset.lineIdx };
      if (e.dataTransfer) { e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', 'x'); } catch (_) {} }
      row.classList.add('dragging');
    });
    seasonBody.addEventListener('dragover', e => {
      const row = e.target.closest('.season-linerow');
      if (row && SEASON_DRAG && row.dataset.slinSide === SEASON_DRAG.side) {
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
        row.classList.add('drag-over');
      }
    });
    seasonBody.addEventListener('dragleave', e => { const row = e.target.closest('.season-linerow'); if (row) row.classList.remove('drag-over'); });
    seasonBody.addEventListener('drop', e => {
      const row = e.target.closest('.season-linerow');
      if (!row || !SEASON_DRAG || row.dataset.slinSide !== SEASON_DRAG.side) { SEASON_DRAG = null; return; }
      e.preventDefault();
      seasonReorderLine(SEASON_DRAG.side, SEASON_DRAG.fromIdx, +row.dataset.lineIdx);
      SEASON_DRAG = null;
    });
    seasonBody.addEventListener('dragend', () => {
      SEASON_DRAG = null;
      document.querySelectorAll('.season-linerow.dragging, .season-linerow.drag-over').forEach(el => el.classList.remove('dragging', 'drag-over'));
    });
  }
  const seasonBack = $('#season-back');
  if (seasonBack) seasonBack.addEventListener('click', () => showScreen('start'));
  const seasonReset = $('#season-reset');
  if (seasonReset) seasonReset.addEventListener('click', () => {
    if (!confirm('シーズンの進行・成績をすべて初期化して、新しいシーズンを開始します。よろしいですか？')) return;
    SEASON = seasonNewState(); SEASON_MANUAL_SEL = null; saveSeason(); SEASON_VIEW = 'menu'; renderSeason();
  });
  const seasonNext = $('#season-next');
  if (seasonNext) seasonNext.addEventListener('click', () => seasonAfterManualResult(false));
  const seasonToHub = $('#season-tohub');
  if (seasonToHub) seasonToHub.addEventListener('click', () => seasonAfterManualResult(true));
  // エキシビション結果画面: モード選択(スタート)画面へ戻る
  const resultToStart = $('#result-to-start');
  if (resultToStart) resultToStart.addEventListener('click', () => { resetGame(false); showScreen('start'); });
  const backFromSetup = $('#backToStartFromSetup');
  if (backFromSetup) backFromSetup.addEventListener('click', () => showScreen('start'));
  const tbBack = $('#tb-back');
  if (tbBack) tbBack.addEventListener('click', () => showScreen('start'));
  const tbReset = $('#tb-reset');
  if (tbReset) tbReset.addEventListener('click', () => {
    // 2段階確認: 誤操作防止
    if (!confirm('チーム編成をリセットします。よろしいですか？')) return;
    if (!confirm('本当にすべての選手割当をクリアしますか？ (この操作は取り消せません)')) return;
    resetTeamBuild();
    saveTeamBuild();  // 空状態を保存して localStorage も同期
    renderTeamBuild();
  });
  const tbAuto = $('#tb-auto');
  if (tbAuto) tbAuto.addEventListener('click', () => {
    if (!confirm('投手陣と現在のオーダーの野手を自動編成で上書きします。\n(他のオーダーは変更しません) よろしいですか？')) return;
    autoFillTeamBuild({ batters: true, pitchers: true });
  });
  const tbAutoOrder = $('#tb-auto-order');
  if (tbAutoOrder) tbAutoOrder.addEventListener('click', () => {
    const n = (TB_STATE && TB_STATE.currentOrder != null ? TB_STATE.currentOrder : 0) + 1;
    if (!confirm('現在のオーダー' + n + 'の野手のみを自動編成で上書きします。\n(投手陣・他のオーダーは変更しません) よろしいですか？')) return;
    autoFillTeamBuild({ batters: true, pitchers: false });
  });
  const tbSave = $('#tb-save');
  if (tbSave) tbSave.addEventListener('click', () => {
    const ok = saveTeamBuild();
    tbSave.textContent = ok ? '✓ 保存しました' : '⚠ 保存失敗';
    if (!ok) alert('保存に失敗しました（ブラウザの保存容量超過の可能性）。\n\n【診断情報】\n' + tbStorageDiag());
    setTimeout(() => { tbSave.textContent = '💾 チームセット'; }, 1500);
  });
  const tbTeam = $('#tb-team-select');
  if (tbTeam) tbTeam.addEventListener('change', () => {
    // 現在のチーム状態を保存してから新チームをロード (他チームの編成を保護)
    if (TB_STATE) saveTeamBuild();
    loadTeamBuild(tbTeam.value);   // 該当チームの保存があれば復元、無ければ blank
    renderTeamBuild();
  });

  $('#autoFill').addEventListener('click', autoFill);
  $('#startGame').addEventListener('click', startGame);
  // セットアップの 動画ON/OFF トグル (赤=ON / 青=OFF)
  const vToggle = $('#videoToggle');
  if (vToggle) vToggle.addEventListener('click', () => {
    VIDEO_ON = !VIDEO_ON;
    updateVideoToggleBtn();
    updateAutoVideoBar();   // OFFにしたら自動再生バーを隠す/ONなら出す
  });
  // 試合画面・ダイヤモンド下の 自動再生トグル (動画再生中 ▶ / 動画停止中 ■)
  const avBtn = $('#autoVideoBtn');
  if (avBtn) avBtn.addEventListener('click', toggleAutoVideo);
  updateVideoToggleBtn();   // 初期表示 (動画ON=赤)
  // ダイヤモンド枠内の振り返りボタン (試合終了後)
  const hPrev = $('#bbHistPrev'); if (hPrev) hPrev.addEventListener('click', () => historyStep(-1));
  const hNext = $('#bbHistNext'); if (hNext) hNext.addEventListener('click', () => historyStep(+1));
  const hInn  = $('#bbHistInning'); if (hInn) hInn.addEventListener('change', () => historyJumpToInning(hInn.value));
  $('#resetGame').addEventListener('click', resetGame);
  $('#autoPlay').addEventListener('click', () => runAutoInning());
  $('#autoFinish').addEventListener('click', () => {
    // 試合終了後で結果待ち → 結果/戦評画面へ進む
    if (G.awaitingResult) { showResultScreen(); return; }
    // 走行中なら停止、そうでなければ試合終了まで自動進行
    if (G.autoToEnd) stopAutoToEnd();
    else runAutoToEnd();
  });
  $('#reliefBtn').addEventListener('click', showReliefDialog);
  const pinchBtn = $('#pinchBtn');
  if (pinchBtn) pinchBtn.addEventListener('click', showPinchDialog);
  $('#backToSetup').addEventListener('click', () => {
    showScreen('setup'); resetGame(true);
  });
  // 試合中の「セットアップに戻る」ボタン
  const backInGame = $('#backToSetupGame');
  if (backInGame) backInGame.addEventListener('click', () => {
    if (confirm('試合を中止してセットアップに戻りますか？')) {
      G.ended = true;
      showScreen('setup');
    }
  });
  // レギュラーシーズン手動試合中: シーズン画面 / スタート画面へ戻る
  const backSeasonG = $('#backToSeasonGame');
  if (backSeasonG) backSeasonG.addEventListener('click', () => {
    if (confirm('この試合を中断してレギュラーシーズン画面に戻りますか？\n（この試合の成績は記録されません）')) {
      G.ended = true; SEASON_VIEW = 'manual'; showScreen('season'); renderSeason();
    }
  });
  const backStartG = $('#backToStartGame');
  if (backStartG) backStartG.addEventListener('click', () => {
    if (confirm('この試合を中断してスタート画面に戻りますか？\n（この試合の成績は記録されません）')) {
      G.ended = true; showScreen('start');
    }
  });
  // 投手カード内の球種ボタンクリック (イベント委譲)
  document.body.addEventListener('click', e => {
    const pb = e.target.closest('.pitch-btn');
    if (pb) {
      if (G.autoVideoPlaying) return;   // 自動再生中はAIが投球 (手動投球は無効)
      const name = pb.dataset.pitch;
      if (!G.currentPitcher || !G.currentPitcher.pitches) return;
      const pi = G.currentPitcher.pitches.find(x => x.name === name);
      if (pi) {
        // 投球前の状況を控え、解決後に攻守交替/継投を判定して動画(結果→交替→投手登場→打者登場)を重ねる
        const before = { top: G.top, inning: G.inning, pitcher: G.currentPitcher };
        pitchOne(pi);
        playPitchVideo(before);
      }
      return;
    }
    // 選手名(ミニカード)クリックで詳細表示
    const el = e.target.closest('.player-link');
    if (!el) return;
    const name = el.dataset.playerName;
    const year = el.dataset.playerYear;
    const type = el.dataset.playerType;   // 'pitcher'/'batter' (二刀流の投手版/打者版を区別)
    const team = el.dataset.playerTeam;   // 任意
    const all = loadAllPlayers();
    // 種別・チームが指定されていればそれも一致させる (二刀流大谷の投手版/打者版を正しく表示)
    const p = all.find(x => x.fullNameTop === name
                         && String(x.year || '') === String(year || '')
                         && (!type || playerType(x) === type)
                         && (!team || String(x.team || '') === String(team || '')))
           || all.find(x => x.fullNameTop === name && String(x.year || '') === String(year || ''))
           || all.find(x => x.fullNameTop === name);
    if (p) window.CARD_VIEW.show(p);
  });
}

// ============== セットアップUI ==============
const EXTRAS_KEY = 'mlb_card_extras_v1';

// 古いフォーマット（"制　球" のような全角スペース入りキー）でも能力値が引けるよう
// 全選手データのキーから空白を除去して正規化する
// + rawHtml がある場合、DRS を再抽出して旧パーサーのバグを回避
function normalizePlayer(p) {
  if (!p || typeof p !== 'object') return p;
  function stripKeys(obj) {
    if (!obj || typeof obj !== 'object') return obj;
    const out = {};
    for (const k in obj) {
      const newKey = k.replace(/\s+/g, '');
      out[newKey] = obj[k];
    }
    return out;
  }
  if (p.stats)     p.stats     = stripKeys(p.stats);
  if (p.statsMini) p.statsMini = stripKeys(p.statsMini);
  if (p.record)    p.record    = stripKeys(p.record);
  // rawHtml があれば DRS を再抽出して上書き (旧 parserのバグ修復)
  if (p.rawHtml) {
    const allPoses = [];
    const allInns  = [];
    const allNums  = [];
    let m;
    const reP = /<span class="drs-pos">([\s\S]*?)<\/span>/g;
    while ((m = reP.exec(p.rawHtml)) !== null) {
      allPoses.push(m[1].replace(/<[^>]+>/g,'').trim());
    }
    const reI = /<span class="drs-inn">([\s\S]*?)<\/span>/g;
    while ((m = reI.exec(p.rawHtml)) !== null) {
      const txt = m[1].replace(/<[^>]+>/g,'').trim();
      const n = parseFloat(txt.replace(/[^\-0-9.]/g, ''));
      allInns.push(Number.isFinite(n) ? n : 0);
    }
    const reN = /<span class="drs-num[^"]*">([\s\S]*?)<\/span>/g;
    while ((m = reN.exec(p.rawHtml)) !== null) {
      const txt = m[1].replace(/<[^>]+>/g,'').trim();
      const n = parseFloat(txt.replace(/[^\-0-9.]/g, ''));
      allNums.push(Number.isFinite(n) ? n : 0);
    }
    if (allPoses.length > 0) {
      const newDrs = [];
      const seen = new Set();
      for (let i = 0; i < allPoses.length; i++) {
        const pos = allPoses[i];
        if (seen.has(pos)) continue;
        seen.add(pos);
        newDrs.push({ pos, value: allNums[i] ?? 0, innings: allInns[i] ?? 0 });
      }
      p.drs = newDrs;
    }
    // 捕手データ (リード/阻止率) の再抽出: 旧パーサーで取り込んだカードは p.catcher が無い。
    //   カード表示は rawHtml を直描画するため「画面には見えるのに内部データに無い」状態になり、
    //   リード査定・盗塁阻止・スタメン捕手選定がすべて効かなくなる → rawHtml から補完する。
    //   実カードの構造 (card_generator 出力):
    //     <div class="catcher-bar" style="…">          ← style属性あり
    //       <div class="ca-item"> …リード… </div>       ← 内側にdivがネスト
    //       <div class="ca-item"> …阻止率… </div>
    //     </div>
    //   属性付き開始タグを許容し、ネストで閉じタグ探索が壊れないよう
    //   開始位置から一定範囲のセグメントを走査して ca-label / ca-val を対で拾う。
    if (!p.catcher || (p.catcher['リード'] == null && p.catcher['阻止率'] == null)) {
      const mOpen = p.rawHtml.match(/<div class="catcher-bar"[^>]*>/);
      if (mOpen) {
        const seg = p.rawHtml.slice(mOpen.index, mOpen.index + 1500);
        const labels = [], vals = [];
        const reL = /<span class="ca-label"[^>]*>([\s\S]*?)<\/span>/g;
        while ((m = reL.exec(seg)) !== null) labels.push(m[1].replace(/<[^>]+>/g, '').trim().replace(/\s+/g, ''));
        const reV = /<span class="ca-val[^"]*"[^>]*>([\s\S]*?)<\/span>/g;
        while ((m = reV.exec(seg)) !== null) vals.push(m[1].replace(/<[^>]+>/g, '').trim());
        if (labels.length) {
          p.catcher = {};
          for (let i = 0; i < labels.length; i++) {
            const n = parseFloat(String(vals[i] ?? '').replace(/[^\-0-9.]/g, ''));
            p.catcher[labels[i]] = Number.isFinite(n) ? n : 0;
          }
        }
      }
    }
  }
  return p;
}

// 正規化済み選手リストのキャッシュ。normalizePlayer (rawHtmlのDRS再抽出=正規表現) は重いので、
// 呼び出しごとに全選手分やり直さず、ソース(window.PLAYERS / 追加カード)が変わらない限り再利用する。
let _allPlayersCache = null;
let _allPlayersCacheKey = null;
function loadAllPlayers() {
  let extras = [];
  // 追加カードは CardStore(IndexedDB) のメモリキャッシュから同期取得。
  // CardStore が無い/未初期化の場合は旧 localStorage を直接読む。
  const fromStore = !!(window.CardStore && window.CardStore.getCachedSync);
  if (fromStore) {
    extras = window.CardStore.getCachedSync();
  } else {
    try { extras = JSON.parse(localStorage.getItem(EXTRAS_KEY)) || []; } catch (e) {}
  }
  const base = window.PLAYERS || [];
  // キャッシュ判定: ベース配列の参照+件数、追加カードの参照(CardStore)or件数(localStorage) が同一なら再利用
  const key = [base, base.length, fromStore ? extras : null, extras.length];
  if (_allPlayersCache && _allPlayersCacheKey &&
      _allPlayersCacheKey[0] === key[0] && _allPlayersCacheKey[1] === key[1] &&
      _allPlayersCacheKey[2] === key[2] && _allPlayersCacheKey[3] === key[3]) {
    return _allPlayersCache;
  }
  _allPlayersCache = [...base, ...extras].map(normalizePlayer);
  _allPlayersCacheKey = key;
  return _allPlayersCache;
}
function getBatters() { return loadAllPlayers().filter(p => p.type === 'batter'); }
function getPitchers(){ return loadAllPlayers().filter(p => p.type === 'pitcher'); }

// 選手の種別 (投手/打者)。card.type 優先、無ければ球種の有無で判定。
function playerType(p) { return (p && p.type) ? p.type : ((p && p.pitches && p.pitches.length > 0) ? 'pitcher' : 'batter'); }
// 選手名が長い (9文字超) 場合に、枠内へ収まるよう文字サイズを少し小さくするための
// 追加クラスを返す。各表示箇所の名前要素の class に付与する。
function longNameClass(name) {
  const n = (name || '').length;
  if (n > 12) return ' name-xlong';   // 13文字以上 → さらに縮小
  if (n > 9)  return ' name-long';    // 10〜12文字 → 少し縮小
  return '';
}
// 選手の同一性キー: 名前 + 年 + チーム + 種別。
// 大谷翔平のような二刀流(同名・同年・同チーム)を「投手版」「打者版」で区別するため種別を含める。
function playerKey(p) {
  if (!p) return '';
  return (p.fullNameTop || '') + '_' + (p.year || '') + '_' + (p.team || '') + '_' + playerType(p);
}
// 保存用 ID
function playerIdOf(p) { return p ? { name: p.fullNameTop, year: p.year, team: p.team, type: playerType(p) } : null; }
// 保存 ID から選手を照合 (古い保存= type/team 無し にも寛容に対応)
function playerMatchesId(p, id) {
  if (!p || !id) return false;
  if (p.fullNameTop !== id.name) return false;
  if (String(p.year || '') !== String(id.year || '')) return false;
  if (id.type != null && playerType(p) !== id.type) return false;
  // チームは別名(SF/SFG/SFN等)を吸収して比較。保存IDのteamとカードのteam表記が違っても同一チームなら一致とみなす。
  if (id.team != null && id.team !== '' && normalizeTeam(p.team) !== normalizeTeam(id.team)) return false;
  return true;
}

// チーム編成画面で保存済みのチームデータを読み込み (なければ null)
// 'original'(全選手) もチーム編成で保存していれば、その編成を読み込んでセットアップに連携する。
// team の保存済み編成を取得。orderIdx(0〜2) で どの打順オーダーの野手を使うか指定する。
// 投手陣は全オーダー共通。orders 配列が無い旧形式は top-level(=オーダー1) を使用する。
function getSavedTeamBuild(team, orderIdx) {
  if (!team) return null;
  try {
    const raw = localStorage.getItem('mlb_team_build_v1_' + team);
    if (!raw) return null;
    const data = JSON.parse(raw);
    const all = [...getBatters(), ...getPitchers()];
    const lookup = id => id ? all.find(p => playerMatchesId(p, id)) : null;
    // 選択オーダーの野手データ (batters/batterOrder/pinchHitters)。
    const oi = Math.max(0, Math.min(2, orderIdx || 0));
    const orderSrc = (Array.isArray(data.orders) && data.orders[oi]) ? data.orders[oi] : data;
    const batters = {};
    for (const k of ['C','1B','2B','3B','SS','LF','CF','RF','DH']) {
      batters[k] = lookup((orderSrc.batters || {})[k]);
    }
    const pitchers = {};
    for (const r of ['starter','mop','middle','setup','closer','bench']) {
      pitchers[r] = ((data.pitchers || {})[r] || []).map(lookup).filter(Boolean);
    }
    // 保存内容が完全に空であれば null と同等
    const hasAny = Object.values(batters).some(Boolean) || Object.values(pitchers).some(arr => arr.length > 0);
    if (!hasAny) return null;
    return {
      team,
      batters,
      batterOrder: orderSrc.batterOrder || {},
      // 控えは「枠の位置(PH1/PH2/PH3/代走/守備/守備)」で役割が決まるため、
      // null を詰めずに保持する (filter(Boolean) すると役割スロットがズレて代走優先が壊れる)
      pinchHitters: (orderSrc.pinchHitters || []).map(lookup),
      pitchers,
    };
  } catch (e) { return null; }
}

// 投手スロットラベル (セットアップ画面の先発枠 — 1 のみ)
const PITCHER_SLOT_LABELS = ['先発'];
const PITCHER_SLOT_COUNT = PITCHER_SLOT_LABELS.length;
// セットアップのリリーフ構成 (中継/SU/抑え/モップ) — 空欄でも試合開始可
const SETUP_RELIEF_GROUPS = [
  { role: 'middle', label: '中継', count: 5 },
  { role: 'setup',  label: 'SU',   count: 2 },
  { role: 'closer', label: '抑え', count: 1 },
  { role: 'mop',    label: 'MU', count: 2 },
];
const SETUP_RELIEF_COUNT = SETUP_RELIEF_GROUPS.reduce((s, g) => s + g.count, 0); // 10
// 控え (ベンチ) ラベル — 空欄でも試合開始可
const SETUP_BENCH_LABELS = ['PH1','PH2','PH3','代走','守備','守備'];
// リリーフスロット index → 役割 (中継/SU/抑え/モップ)
function reliefRoleForSlot(slotIdx) {
  let n = 0;
  for (const g of SETUP_RELIEF_GROUPS) {
    if (slotIdx < n + g.count) return g.role;
    n += g.count;
  }
  return 'mop';
}
const PITCHER_ROLE_LABELS = { starter:'先発', middle:'中継', setup:'SU', closer:'抑え', mop:'MU' };

// 守備ポジション定義
const POSITIONS = {
  'C':  { label: '捕手',  jpPos: ['捕手'] },
  '1B': { label: '一塁',  jpPos: ['一塁手'] },
  '2B': { label: '二塁',  jpPos: ['二塁手'] },
  '3B': { label: '三塁',  jpPos: ['三塁手'] },
  'SS': { label: '遊撃',  jpPos: ['遊撃手'] },
  'LF': { label: '左翼',  jpPos: ['左翼手'] },
  'CF': { label: '中堅',  jpPos: ['中堅手'] },
  'RF': { label: '右翼',  jpPos: ['右翼手'] },
  'DH': { label: 'DH',    jpPos: [] }, // DHは誰でも
};
const POSITION_KEYS = ['C','1B','2B','3B','SS','LF','CF','RF','DH'];
// スコアボード等で使う守備位置の1文字略号 (Yahoo方式)
const POS_ABBR = { 'C':'捕','1B':'一','2B':'二','3B':'三','SS':'遊','LF':'左','CF':'中','RF':'右','DH':'指' };
// デフォルト打順別ポジション (1番〜9番) — DHは5番(クリーンアップの後ろ)
const DEFAULT_BATTER_POS = ['CF','SS','1B','3B','DH','LF','RF','2B','C'];

// 選手 b が pos キーを守れるか
// ルール: DH は誰でも可。それ以外は「その年の守備DRSに該当ポジション+出場イニング>0」
//   が記載されている選手のみ。position-badge(b.position)による救済はしない。
//   例: 2025年大谷翔平のように position-badge が守備位置を示していても、
//   当該年の守備DRSに出場記録がなければ DH のみで起用可。
function canPlay(b, pos) {
  if (!b) return false;
  if (pos === 'DH') return true;
  // 捕手(C)は、捕手データ(リード/阻止率のcatcher-bar)を持つカードなら守備DRSにC記載が無くても可。
  //   捕手データは「その年に捕手として出場した選手」にしか無い実データのため、
  //   position-badge による救済(上のルールで禁止)とは異なり適性の根拠になる。
  if (pos === 'C' && b.catcher && (b.catcher['リード'] != null || b.catcher['阻止率'] != null)) return true;
  if (!b.drs) return false;
  return b.drs.some(d => d.pos === pos && (Number(d.innings) || 0) > 0);
}
function filterBattersByPos(batters, pos) {
  return batters.filter(b => canPlay(b, pos));
}

// チームフィルタ: side ごとに現在の選択を取得
function getTeamFilter(side) {
  const sel = $('.team-filter[data-side="' + side + '"]');
  return sel ? sel.value : 'original';
}
// 打順オーダー(0=オーダー1/1=オーダー2/2=オーダー3): side ごとの選択を取得
function getOrderFilter(side) {
  const sel = $('.order-filter[data-side="' + side + '"]');
  return sel ? (parseInt(sel.value, 10) || 0) : 0;
}
// 同じ球団の表記ゆれを、チーム選択(MLB_DIVISIONS)の正式コードへ寄せる別名表。
//   例: カードが "SFG" でも編成のドロップダウンは "SF" → 一致しないと選手が出ず保存できないため吸収する。
const TEAM_ALIASES = {
  SFG: 'SF',  SFN: 'SF',  SAN: 'SF',     // San Francisco Giants (BBRef:SFG / Retrosheet:SFN / 略:SAN)
  TBR: 'TB',  TBD: 'TB',  TBA: 'TB',  TAM: 'TB',   // Tampa Bay Rays
  KCR: 'KC',  KCA: 'KC',  KAN: 'KC',     // Kansas City Royals
  SDP: 'SD',  SDN: 'SD',                 // San Diego Padres
  WSN: 'WSH', WAS: 'WSH',                // Washington Nationals
  CHW: 'CWS', CHA: 'CWS',                // Chicago White Sox
  CHN: 'CHC',                            // Chicago Cubs
  OAK: 'ATH',                            // Athletics
  ANA: 'LAA', CAL: 'LAA',                // LA Angels
  LAN: 'LAD', LOS: 'LAD',                // LA Dodgers
  NYN: 'NYM',                            // NY Mets
  NYA: 'NYY', NEW: 'NYY',                // NY Yankees
  SLN: 'STL',                            // St. Louis Cardinals
  AZ:  'ARI', ARZ: 'ARI',                // Arizona Diamondbacks
  FLA: 'MIA', FLO: 'MIA',                // Miami (旧Florida) Marlins
};
// チーム名の正規化: 末尾の所属チーム数(例 "LAD2"→"LAD")を除き、頭のローマ字のみにしたうえで別名を正式コードへ寄せる
function normalizeTeam(t) {
  const s = String(t || '').toUpperCase().replace(/[0-9]+$/, '');
  return TEAM_ALIASES[s] || s;
}
function applyTeamFilter(players, team) {
  if (team === 'original' || !team) return players;
  const tag = normalizeTeam(team);
  return players.filter(p => normalizeTeam(p.team) === tag);
}
// 所属が「ALL」の選手は、チーム編成のすべてのチーム・すべての年度で登録可能とする
function isAllTeamPlayer(p) { return normalizeTeam(p && p.team) === 'ALL'; }
// 選手名が「マイナー」始まりのカード = 穴埋め用のマイナー選手 (マイナーP1 / マイナー内野手 等)。
//   所属コードの表記揺れ(ALLでない等)に左右されないよう、判定は「名前がマイナー始まり」のみとする。
//   (マイナー以外の名前の選手は所属がALLでも対象外なので、特別なALL選手は影響を受けない)
//   ・守備位置別の登録可能人数カウント(捕/2・全/28 等)からは除外する。
//   ・自動編成では通常選手を必ず優先し、通常選手で埋まらない枠の穴埋めにのみ選ばれるようにする。
function isMinorPlayer(p) {
  return String((p && p.fullNameTop) || '').trim().startsWith('マイナー');
}
// 自動編成でマイナー選手を「通常選手がいない時の穴埋め」に留めるためのスコア減算量。
//   通常選手の評価値(総合力+各種ボーナス, 高々数百)を必ず下回るよう十分大きく取る(有限値なので穴埋めには選ばれる)。
const MINOR_FILL_PENALTY = 1e6;

function buildSetup() {
  buildSetupSide('away');
  buildSetupSide('home');
  // チーム/オーダー切り替え時に対象 side だけ再構築
  for (const side of ['away','home']) {
    const sel = $('.team-filter[data-side="' + side + '"]');
    if (sel) {
      sel.addEventListener('change', () => buildSetupSide(side));
    }
    const ordSel = $('.order-filter[data-side="' + side + '"]');
    if (ordSel) {
      ordSel.addEventListener('change', () => buildSetupSide(side));
    }
  }
}

// セレクト群の中で、他の枠で選択済みの選手を option から disable にする。
// 各セレクトは _pool (参照する選手配列) を持ち、option.value はその _pool のインデックス。
// 異なる枠は異なるプールを参照しうるため、値ではなく「選手オブジェクトの同一性」で重複判定する。
function dedupSelectGroup(selects) {
  // 使用中の選手オブジェクトを収集
  const used = new Set();
  selects.forEach(s => {
    const pool = s._pool;
    if (s.value !== '' && pool && pool[+s.value]) used.add(pool[+s.value]);
  });
  selects.forEach(s => {
    const pool = s._pool;
    const myPlayer = (s.value !== '' && pool) ? pool[+s.value] : null;
    Array.from(s.options).forEach(opt => {
      if (opt.value === '') { opt.disabled = false; return; }
      const p = pool ? pool[+opt.value] : null;
      if (!p || p === myPlayer) { opt.disabled = false; return; }
      opt.disabled = used.has(p);
    });
  });
}

// 同チーム内の選手重複を防ぐ: 他の枠で選択済みの選手 option を disable する
//   - 投手グループ: 先発 + リリーフ(中継/SU/抑え/モップ)
//   - 打者グループ: 打順9 + 控え(PH/代走/守備)
function refreshSelectionsForSide(side) {
  // 打者グループ (打順 + 控え) — 同一プールを共有
  const batSels = [
    ...$$('.batter-slots[data-side="'+side+'"] .sel-batter'),
    ...$$('.sel-bench-slot[data-side="'+side+'"]'),
  ];
  dedupSelectGroup(batSels);
  // 投手グループ (先発 + リリーフ)
  const pitSels = [
    ...$$('.sel-pitcher-slot[data-side="'+side+'"]'),
    ...$$('.sel-relief-slot[data-side="'+side+'"]'),
  ];
  dedupSelectGroup(pitSels);
}

// 同一グループ内で重複した選択を、優先度の低い枠から空にする
//   - 投手: 先発 > リリーフ(中継/SU/抑え/モップ)
//   - 打者: 打順 > 控え(PH/代走/守備)
// (自動編成後など、複数枠が同じ選手を指してしまった場合の保険)
function clearDuplicateSelections(side) {
  const clearGroup = (selects) => {
    const seen = new Set();
    selects.forEach(s => {
      const pool = s._pool;
      if (s.value === '' || !pool) return;
      const p = pool[+s.value];
      if (!p) return;
      if (seen.has(p)) s.value = '';  // 既出 → 後発の枠を空に
      else seen.add(p);
    });
  };
  clearGroup([
    ...$$('.sel-pitcher-slot[data-side="'+side+'"]'),
    ...$$('.sel-relief-slot[data-side="'+side+'"]'),
  ]);
  clearGroup([
    ...$$('.batter-slots[data-side="'+side+'"] .sel-batter'),
    ...$$('.sel-bench-slot[data-side="'+side+'"]'),
  ]);
}

// セットアップ画面のプールを一元管理 (buildSetupSide と readSetup で同一インデックスを共有)
function getSetupPools(side) {
  const team = getTeamFilter(side);
  const tbSave = getSavedTeamBuild(team, getOrderFilter(side));
  if (tbSave) {
    const starters = (tbSave.pitchers.starter || []).filter(Boolean);
    const mop      = (tbSave.pitchers.mop    || []).filter(Boolean);
    const bench    = (tbSave.pitchers.bench  || []).filter(Boolean);
    // 先発スロット用プール = 先発5名 + モップアップ + 控え投手 (チーム登録時のみ)。
    //  → 手動プルダウンでモップ/控えも先発に選べる。先発が必ず先頭に並ぶ。
    //  ※自動編成・起動時の自動選択は starterCount(=先発の人数)の範囲だけから選ぶので、
    //    ランダムでモップ/控えが先発に選ばれることはない。
    const starterPool = [...starters, ...mop, ...bench];
    const reliefPool = [
      ...(tbSave.pitchers.middle || []),
      ...(tbSave.pitchers.setup  || []),
      ...(tbSave.pitchers.closer || []),
      ...(tbSave.pitchers.mop    || []),
      ...(tbSave.pitchers.bench  || []),
    ];
    const dedup = [];
    const seen = new Set();
    [...Object.values(tbSave.batters), ...(tbSave.pinchHitters || [])].forEach(b => {
      if (!b) return;
      const k = b.fullNameTop + '_' + (b.year||'');
      if (seen.has(k)) return;
      seen.add(k);
      dedup.push(b);
    });
    return { tbSave, batterPool: dedup, starterPool, reliefPool, starterCount: starters.length };
  }
  const allP = applyTeamFilter(getPitchers(), team);
  return { tbSave: null, batterPool: applyTeamFilter(getBatters(), team), starterPool: allP, reliefPool: allP, starterCount: allP.length };
}
// スロット番号に対応する投手プール (slot 0=先発, それ以外=リリーフ)
function setupPitcherPool(pools, slotIdx) {
  return (pools.tbSave && slotIdx > 0) ? pools.reliefPool : pools.starterPool;
}

function buildSetupSide(side) {
  const team     = getTeamFilter(side);
  const pools    = getSetupPools(side);
  const tbSave   = pools.tbSave;
  const batters  = pools.batterPool;
  const pitchers = pools.starterPool;
  const reliefPitchers = pools.reliefPool;
  {
    // 先発スロット (1枠)
    const pol = $('.pitcher-slots[data-side="'+side+'"]');
    pol.innerHTML = '';
    {
      const li = document.createElement('li');
      li.dataset.num = '先発';
      const sel = document.createElement('select');
      sel.className = 'sel-pitcher-slot';
      sel.dataset.side = side;
      sel.dataset.idx = 0;
      sel._pool = pitchers;
      sel.innerHTML = '<option value="">-- 選択 --</option>' +
        pitchers.map((p,pi) => `<option value="${pi}">${labelOf(p)}</option>`).join('');
      // 保存済みチーム(LAD等)なら 先発5人からランダムに自動選択 (モップ/控えは含めない)。
      // pitchers(=starterPool) の先頭 starterCount 名が先発なので、その範囲からのみ選ぶ。
      const starterCount = (pools.starterCount != null) ? pools.starterCount : pitchers.length;
      if (tbSave && starterCount > 0) {
        sel.value = String(randI(starterCount));
      }
      sel.addEventListener('change', () => refreshSelectionsForSide(side));
      li.appendChild(sel);
      const ibtn = document.createElement('button');
      ibtn.type = 'button';
      ibtn.className = 'info-btn';
      ibtn.title = '選択中の投手の詳細カードを表示';
      ibtn.textContent = 'ℹ';
      ibtn.addEventListener('click', () => {
        const v = sel.value;
        if (v === '') { alert('先に投手を選択してください'); return; }
        window.CARD_VIEW.show(pitchers[+v]);
      });
      li.appendChild(ibtn);
      pol.appendChild(li);
    }

    // 打者セレクト 9枠 (各枠に守備ポジション+選手の2セレクタ) + ドラッグハンドル
    const ol = $('.batter-slots[data-side="'+side+'"]');
    ol.innerHTML = '';
    for (let i = 0; i < 9; i++) {
      const li = document.createElement('li');
      li.dataset.num = (i+1);
      li.dataset.side = side;
      li.dataset.idx = i;
      li.draggable = true;
      // ドラッグハンドル (アイコン)
      const handle = document.createElement('span');
      handle.className = 'drag-handle';
      handle.textContent = '⠿';
      handle.title = 'ドラッグして並び替え';
      li.appendChild(handle);
      // 守備ポジションセレクタ
      const posSel = document.createElement('select');
      posSel.className = 'sel-pos';
      posSel.dataset.side = side;
      posSel.dataset.idx = i;
      posSel.innerHTML = POSITION_KEYS.map(pk =>
        `<option value="${pk}"${DEFAULT_BATTER_POS[i] === pk ? ' selected' : ''}>${pk} ${POSITIONS[pk].label}</option>`
      ).join('');
      li.appendChild(posSel);
      // 選手セレクタ (ポジションで絞り込み)
      const sel = document.createElement('select');
      sel.className = 'sel-batter';
      sel.dataset.idx = i;
      sel._pool = batters;
      const renderPlayerOptions = () => {
        const pos = posSel.value;
        const filtered = filterBattersByPos(batters, pos);
        const prev = sel.value;
        sel.innerHTML = '<option value="">-- 選択 --</option>' +
          filtered.map(b => {
            const realIdx = batters.indexOf(b);
            return `<option value="${realIdx}">${labelOf(b)}</option>`;
          }).join('');
        if (prev && Array.from(sel.options).some(o => o.value === prev)) {
          sel.value = prev;
        }
      };
      renderPlayerOptions();
      posSel.addEventListener('change', () => {
        renderPlayerOptions();
        refreshSelectionsForSide(side);
      });
      sel.addEventListener('change', () => refreshSelectionsForSide(side));
      li._renderPlayerOptions = renderPlayerOptions;
      li.appendChild(sel);
      const ibtn = document.createElement('button');
      ibtn.type = 'button';
      ibtn.className = 'info-btn';
      ibtn.title = '選択中の打者の詳細カードを表示';
      ibtn.textContent = 'ℹ';
      ibtn.addEventListener('click', () => {
        const v = sel.value;
        if (v === '') { alert('先に打者を選択してください'); return; }
        window.CARD_VIEW.show(batters[+v]);
      });
      li.appendChild(ibtn);
      ol.appendChild(li);
    }
    // 保存済みチームがあれば 打順 1〜9 を pre-fill
    if (tbSave) {
      const lis = ol.querySelectorAll('li');
      // 1) 保存された (order, position, batter) を集めて、order→{pos, batter} のマップを作成
      const orderMap = {};
      for (const pos of Object.keys(tbSave.batterOrder)) {
        const ord = tbSave.batterOrder[pos];
        if (ord >= 1 && ord <= 9) {
          orderMap[ord] = { pos, batter: tbSave.batters[pos] };
        }
      }
      // 2) 既に使用中のポジションを除外したフォールバックポジションを用意
      //    (未割当の order スロットに対して 守備位置の重複を避ける)
      const usedPositions = new Set();
      Object.values(orderMap).forEach(e => { if (e.pos) usedPositions.add(e.pos); });
      const fallbackPositions = POSITION_KEYS.filter(pk => !usedPositions.has(pk));
      let fbIdx = 0;
      // 3) 全 9 スロットを上書き
      for (let ord = 1; ord <= 9; ord++) {
        const li = lis[ord - 1];
        const posSel = li.querySelector('.sel-pos');
        const batSel = li.querySelector('.sel-batter');
        const entry = orderMap[ord];
        // ポジション設定 (保存があればそれ、無ければ未使用ポジションから割当)
        const targetPos = entry?.pos || fallbackPositions[fbIdx++] || 'DH';
        posSel.value = targetPos;
        posSel.dispatchEvent(new Event('change'));
        // 選手設定 (保存あり かつ そのポジションを実際に守れる場合のみ)
        if (entry?.batter && canPlay(entry.batter, targetPos)) {
          const realIdx = batters.indexOf(entry.batter);
          if (realIdx >= 0) {
            batSel.value = String(realIdx);
          }
        } else {
          batSel.value = '';  // 守れない/保存無しなら一旦空に
        }
      }
      // 空きスロットを「そのポジションを守れる未使用選手」で自動補完
      // (不正セーブや欠落があっても試合を開始できるようにする)
      const usedKeys = new Set();
      lis.forEach(li => {
        const bv = li.querySelector('.sel-batter').value;
        if (bv !== '' && batters[+bv]) {
          const b = batters[+bv];
          usedKeys.add(b.fullNameTop + '_' + (b.year||''));
        }
      });
      lis.forEach(li => {
        const batSel = li.querySelector('.sel-batter');
        if (batSel.value !== '') return;
        const pos = li.querySelector('.sel-pos').value;
        const cand = batters.find(b => {
          const k = b.fullNameTop + '_' + (b.year||'');
          return !usedKeys.has(k) && canPlay(b, pos);
        });
        if (cand) {
          batSel.value = String(batters.indexOf(cand));
          usedKeys.add(cand.fullNameTop + '_' + (cand.year||''));
        }
      });
      refreshSelectionsForSide(side);
    }
    // ドラッグ&ドロップ並び替え設定
    enableLineupDragDrop(side);

    // ===== リリーフ枠 (中継/SU/抑え/モップ) — 空欄可 =====
    buildReliefSlots(side, reliefPitchers, tbSave);
    // ===== 控え枠 (PH/代走/守備) — 空欄可 =====
    buildBenchSlots(side, batters, tbSave);
    // 全枠ビルド後にまとめて重複オプションを disable
    // (投手: 先発+リリーフ / 打者: 打順+控え)
    refreshSelectionsForSide(side);
  }
}

// リリーフ投手スロット (中継4/SU2/抑え1/モップ2) を構築。保存があれば pre-fill
// 各役割スロットは、その役割の保存配列から個別に充填する (役割を跨いで詰めない)
function buildReliefSlots(side, reliefPool, tbSave) {
  const rol = $('.relief-slots[data-side="'+side+'"]');
  if (!rol) return;
  rol.innerHTML = '';
  let slotNo = 0;
  for (const group of SETUP_RELIEF_GROUPS) {
    // この役割の保存済み投手 (順序保持。空欄はそのまま空欄に)
    const savedForRole = tbSave ? (tbSave.pitchers[group.role] || []) : [];
    for (let g = 0; g < group.count; g++) {
      const li = document.createElement('li');
      li.dataset.num = group.label;
      const sel = document.createElement('select');
      sel.className = 'sel-relief-slot';
      sel.dataset.side = side;
      sel.dataset.idx = slotNo;
      sel._pool = reliefPool;
      sel.innerHTML = '<option value="">-- 任意 --</option>' +
        reliefPool.map((p, pi) => `<option value="${pi}">${labelOf(p)}</option>`).join('');
      // この役割の g 番目の保存選手を、reliefPool 内のインデックスで pre-fill
      const savedPlayer = savedForRole[g];
      if (savedPlayer) {
        const idx = reliefPool.indexOf(savedPlayer);
        if (idx >= 0) sel.value = String(idx);
      }
      sel.addEventListener('change', () => refreshSelectionsForSide(side));
      li.appendChild(sel);
      const ibtn = document.createElement('button');
      ibtn.type = 'button'; ibtn.className = 'info-btn'; ibtn.textContent = 'ℹ';
      ibtn.title = '選択中の投手の詳細カード';
      ibtn.addEventListener('click', () => {
        const v = sel.value;
        if (v === '') { alert('先に投手を選択してください'); return; }
        window.CARD_VIEW.show(reliefPool[+v]);
      });
      li.appendChild(ibtn);
      rol.appendChild(li);
      slotNo++;
    }
  }
  refreshSelectionsForSide(side);
}

// 控え (ベンチ) スロット (PH1-3/代走/守備) を構築。保存があれば PH を pre-fill
function buildBenchSlots(side, batterPool, tbSave) {
  const bol = $('.bench-slots[data-side="'+side+'"]');
  if (!bol) return;
  bol.innerHTML = '';
  // 既に打順に入っている選手は除く
  const usedInLineup = new Set();
  $$('.batter-slots[data-side="'+side+'"] .sel-batter').forEach(s => {
    if (s.value !== '' && batterPool[+s.value]) {
      const b = batterPool[+s.value];
      usedInLineup.add(b.fullNameTop + '_' + (b.year||''));
    }
  });
  const phSaved = tbSave ? (tbSave.pinchHitters || []) : [];
  for (let i = 0; i < SETUP_BENCH_LABELS.length; i++) {
    const li = document.createElement('li');
    li.dataset.num = SETUP_BENCH_LABELS[i];
    const sel = document.createElement('select');
    sel.className = 'sel-bench-slot';
    sel.dataset.side = side;
    sel.dataset.idx = i;
    sel._pool = batterPool;
    sel.innerHTML = '<option value="">-- 任意 --</option>' +
      batterPool.map((p, pi) => `<option value="${pi}">${labelOf(p)}</option>`).join('');
    // 保存があり PH があれば pre-fill (打順未使用の選手のみ)
    if (phSaved[i]) {
      const k = phSaved[i].fullNameTop + '_' + (phSaved[i].year||'');
      if (!usedInLineup.has(k)) {
        const idx = batterPool.indexOf(phSaved[i]);
        if (idx >= 0) sel.value = String(idx);
      }
    }
    sel.addEventListener('change', () => refreshSelectionsForSide(side));
    li.appendChild(sel);
    const ibtn = document.createElement('button');
    ibtn.type = 'button'; ibtn.className = 'info-btn'; ibtn.textContent = 'ℹ';
    ibtn.title = '選択中の選手の詳細カード';
    ibtn.addEventListener('click', () => {
      const v = sel.value;
      if (v === '') { alert('先に選手を選択してください'); return; }
      window.CARD_VIEW.show(batterPool[+v]);
    });
    li.appendChild(ibtn);
    bol.appendChild(li);
  }
}

// 打順スロットを D&D で入れ替え可能にする
// ポジション + 選手をセットで移動する
function enableLineupDragDrop(side) {
  const ol = $('.batter-slots[data-side="'+side+'"]');
  const lis = ol.querySelectorAll('li');
  lis.forEach(li => {
    li.addEventListener('dragstart', e => {
      // セレクトボックスからのドラッグは無視 (テキスト編集と干渉)
      if (e.target.tagName === 'SELECT' || e.target.tagName === 'OPTION') {
        e.preventDefault(); return;
      }
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', JSON.stringify({
        side: li.dataset.side,
        idx: li.dataset.idx,
      }));
      li.classList.add('dragging');
    });
    li.addEventListener('dragend', () => {
      li.classList.remove('dragging');
      ol.querySelectorAll('li').forEach(x => x.classList.remove('drag-over'));
    });
    li.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
    });
    li.addEventListener('dragenter', e => {
      e.preventDefault();
      if (!li.classList.contains('dragging')) li.classList.add('drag-over');
    });
    li.addEventListener('dragleave', () => {
      li.classList.remove('drag-over');
    });
    li.addEventListener('drop', e => {
      e.preventDefault();
      li.classList.remove('drag-over');
      let data;
      try { data = JSON.parse(e.dataTransfer.getData('text/plain') || '{}'); }
      catch (err) { return; }
      // 同じチーム同士のみ入れ替え
      if (data.side !== li.dataset.side) return;
      const srcIdx = parseInt(data.idx);
      const dstIdx = parseInt(li.dataset.idx);
      if (Number.isNaN(srcIdx) || Number.isNaN(dstIdx) || srcIdx === dstIdx) return;
      swapLineupSlots(li.dataset.side, srcIdx, dstIdx);
    });
  });
}

// 打順スロット同士を入れ替え (ポジション + 選手をセットで移動)
function swapLineupSlots(side, fromIdx, toIdx) {
  const lis = $$('.batter-slots[data-side="'+side+'"] li');
  const fromLi = lis[fromIdx], toLi = lis[toIdx];
  const fromPosSel = fromLi.querySelector('.sel-pos');
  const fromBatSel = fromLi.querySelector('.sel-batter');
  const toPosSel   = toLi.querySelector('.sel-pos');
  const toBatSel   = toLi.querySelector('.sel-batter');
  // 現在値を退避
  const fromPos = fromPosSel.value;
  const fromBat = fromBatSel.value;
  const toPos   = toPosSel.value;
  const toBat   = toBatSel.value;
  // ポジション変更 → 選手リストが再生成される (change イベント発火)
  // → その後で選手を復元する順序にする
  fromPosSel.value = toPos;
  fromPosSel.dispatchEvent(new Event('change'));
  fromBatSel.value = toBat;
  toPosSel.value = fromPos;
  toPosSel.dispatchEvent(new Event('change'));
  toBatSel.value = fromBat;
  // 視覚的フィードバック (一時的にハイライト)
  fromLi.classList.add('swapped');
  toLi.classList.add('swapped');
  setTimeout(() => {
    fromLi.classList.remove('swapped');
    toLi.classList.remove('swapped');
  }, 600);
}

// 選手の「総合力」(カード表面に表示される power-value の数値)。
//  1) パース済みの powerValue があればそれを使う
//  2) なければ rawHtml から power-value を再抽出 (div/span 両対応)
//  3) それも無ければ 6 つのゲームステータスの平均で代用する
// 結果は p._ovr にキャッシュして再計算を避ける。
function overallOf(p) {
  if (!p || typeof p !== 'object') return null;
  if (Number.isFinite(p._ovr)) return p._ovr;
  let v = null;
  if (Number.isFinite(p.powerValue)) v = p.powerValue;
  if (v === null && p.rawHtml) {
    const m = p.rawHtml.match(/class="power-value"[^>]*>\s*([0-9]+)/);
    if (m) v = parseInt(m[1], 10);
  }
  if (v === null && p.stats) {
    const keys = (p.type === 'pitcher')
      ? ['スタミナ','制球','緩急','精神','奪三振','球重']
      : ['ミート','パワー','スピード','チャンス','選球眼','三振耐性'];
    let sum = 0, n = 0;
    for (const k of keys) { const x = parseFloat(p.stats[k]); if (Number.isFinite(x)) { sum += x; n++; } }
    if (n > 0) v = Math.round(sum / n);
  }
  p._ovr = Number.isFinite(v) ? v : null;
  return p._ovr;
}
// 総合力バッジ HTML (ワンポイント表示用)。値が取れなければ空文字。
//   色はカードのレアリティ(cardRarity: R/RR/SR/SSR/UR)に合わせる。
function ovrBadge(p) {
  const v = overallOf(p);
  if (!Number.isFinite(v)) return '';
  return `<span class="ovr-badge ovr-rar-${cardRarity(p)}" title="総合力">${v}</span>`;
}

function labelOf(p) {
  const seas = p.year ? p.year : '';
  const v = overallOf(p);
  const ovr = Number.isFinite(v) ? ` 総合${v}` : '';
  return `${p.fullNameTop} (${seas} ${p.team}/${p.position})${ovr}`;
}

// =============================================================
// チーム編成モード (Team Build)
// =============================================================
const TB_DIAMOND_POSITIONS = ['SS','2B','3B','1B','LF','CF','RF','C','DH'];
// ダイヤモンド上の x%, y% 座標
// anchor が 'center' (デフォルト) は座標を中心とし、
// 'topLeft' は座標を左上角、'topRight' は座標を右上角に揃える
const TB_POS_LAYOUT = {
  'LF': { x: 18, y: 22, label: 'LF' },
  'CF': { x: 50, y: 10, label: 'CF' },
  'RF': { x: 82, y: 22, label: 'RF' },
  'SS': { x: 32, y: 42, label: 'SS' },
  '2B': { x: 68, y: 42, label: '2B' },
  // 1B: 枠の左端(開始点)を 1塁ベースの左端の真下に揃える (実測: 1塁ベース左端 = 66.2%)
  '1B': { x: 66.26, y: 61.9, label: '1B', anchor: 'topLeft' },
  // 3B: 枠の右端(終点)を 3塁ベースの右端の真下に揃える (1B と対称, 3塁ベース右端 = 33.8%)
  '3B': { x: 33.74, y: 61.9, label: '3B', anchor: 'topRight' },
  'C':  { x: 50, y: 82, label: 'C' },
  // DH: 1B に追従して 1B と同量だけ左へ移動 (従来どおり 1B の少し左下)
  'DH': { x: 66.16, y: 82.3, label: 'DH', anchor: 'topLeft', external: true },
};
// 投手陣スロット定義 (グリッド配置順: 左上=先発, 右上=中継ぎ, 左中=SU, 右中=抑え, 左下=モップ, 右下=控え)
const TB_PITCHER_SLOTS = [
  { role: 'starter',  label: '先発',  count: 5 },  // 左上
  { role: 'middle',   label: '中継',  count: 5 },  // 右上
  { role: 'setup',    label: 'SU',    count: 2 },  // 左中
  { role: 'closer',   label: '抑え',  count: 1 },  // 右中
  { role: 'mop',      label: 'MU', count: 2 }, // 左下
  { role: 'bench',    label: '控え',  count: 4 },  // 右下
];
const TB_PH_COUNT = 6;  // PH1, PH2, PH3, 代走, 守備, 守備
const TB_PH_LABELS = ['PH1', 'PH2', 'PH3', '代走', '守備', '守備'];

// チーム編成状態
let TB_STATE = null;
function resetTeamBuild() {
  // 現在のチームの編成のみ空にする (他チームは温存)
  const teamSel = $('#tb-team-select');
  const team = teamSel ? teamSel.value : (TB_STATE ? TB_STATE.team : 'LAD');
  TB_STATE = blankTeamState(team);
}
function initTeamBuild() {
  // 初回のみ最後に保存したチームを読み込む。
  // 一度開いた後にスタートへ戻って再入場した場合は、前回の画面状態
  // (チーム・オーダー・年度・未保存の編集) をそのまま保持して復元する。
  if (!TB_STATE) loadTeamBuild();
  const teamSel = $('#tb-team-select');
  if (teamSel && TB_STATE) teamSel.value = TB_STATE.team;
  renderTeamBuild();
}

// チーム編成の自動設定。指定ロジックで各ポジション・控え・投手を自動選択する。
// ===== AI打順決定 (現代MLBの監督采配・共通ロジック) =====
//   entries=[{pos, p}](最大9人) の各要素へ打順(1〜9)を割り当てた Map(entry→打順) を返す。
//   チーム編成の自動編成と、レギュラーシーズン手動試合の休養時AI打順の両方で使う。
//   現代MLB(セイバーメトリクス以降)の定石:
//     ・2番 = チーム最強の打者 (打席数と得点機会のバランスが最良の、現代MLBの主軸打順)
//     ・4番 = 長打と勝負強さで走者を還すスラッガー
//     ・1番 = 出塁力(選球眼・ミート)最優先＋足 (最多打席を無駄にしない)
//     ・3番 = 「2アウト走者なし」で回りやすい打順のため、2・4番より一段軽い好打者
//     ・5番 = 4番の後ろの保険となる長打力
//     ・9番 = 足のある選手がいれば「第2の1番」として上位へ繋ぐ (DH制の定石)。いなければ最弱打者
//     ・6〜8番 = 残りを打力の高い順
function aiAssignBattingSpots(entries) {
  const stat = (p, k) => (p && p.stats && Number.isFinite(p.stats[k])) ? p.stats[k] : 0;
  const mini = (p, k) => (p && p.statsMini && Number.isFinite(p.statsMini[k])) ? p.statsMini[k] : 0;
  const mSpeed = p => stat(p, 'スピード'), mSteal = p => mini(p, '盗塁能'), mEye = p => stat(p, '選球眼');
  const mPow = p => stat(p, 'パワー'), mMeat = p => stat(p, 'ミート'), mChan = p => stat(p, 'チャンス'), mHR = p => mini(p, 'HR能');
  const obp  = p => mEye(p) * 1.6 + mMeat(p);                          // 出塁力 (選球眼を重視)
  const slug = p => mPow(p) + mHR(p) * 8;                              // 長打力
  const bat  = p => mMeat(p) + mPow(p) + mEye(p) * 0.8 + mHR(p) * 5;   // 総合打力
  const remain = entries.slice();
  const spotOf = new Map();
  const assignSpot = (spotNum, scoreFn, opts) => {
    if (!remain.length) return;
    let cands = remain;
    if (opts && opts.filter) { const f = remain.filter(e => opts.filter(e.p)); if (f.length) cands = f; }
    const better = (opts && opts.min) ? (a, b) => scoreFn(b.p) < scoreFn(a.p) : (a, b) => scoreFn(b.p) > scoreFn(a.p);
    let pick = cands[0];
    for (const e of cands) if (better(pick, e)) pick = e;
    spotOf.set(pick, spotNum);
    remain.splice(remain.indexOf(pick), 1);
  };
  // (1) 2番: 総合打力+勝負強さ最大 = チーム最強打者を最も価値の高い打順へ
  assignSpot(2, p => bat(p) + mChan(p) * 0.5);
  // (2) 4番: 残りで長打力+勝負強さ最大 (走者を還す)
  assignSpot(4, p => slug(p) + mChan(p) * 2 + mMeat(p) * 0.5);
  // (3) 1番: 残りで出塁力+足最大 (選球眼55以上 または 俊足(スピード+盗塁能80以上) を優先)
  assignSpot(1, p => obp(p) + (mSpeed(p) + mSteal(p)) * 0.6, { filter: p => mEye(p) >= 55 || (mSpeed(p) + mSteal(p)) >= 80 });
  // (4) 3番: 残りで総合打力+勝負強さ (2・4番より一段軽い好打者)
  assignSpot(3, p => bat(p) + mChan(p));
  // (5) 5番: 残りで長打力寄りの好打者 (4番の後ろの保険)
  assignSpot(5, p => slug(p) + mMeat(p) * 0.7 + mChan(p));
  // (6) 9番: 残り4人のうち「最も打てる1人」を除いて足のある選手がいれば「第2の1番」。
  //     いなければ従来通り打力最小の選手を9番へ。
  const bestBatP = remain.length ? remain.reduce((m, e) => (bat(e.p) > bat(m.p) ? e : m), remain[0]).p : null;
  if (remain.some(e => e.p !== bestBatP && mSpeed(e.p) + mSteal(e.p) >= 70)) {
    assignSpot(9, p => mSpeed(p) + mSteal(p) + mEye(p), { filter: p => p !== bestBatP && mSpeed(p) + mSteal(p) >= 70 });
  } else {
    assignSpot(9, p => bat(p), { min: true });
  }
  // (7) 6・7・8番: 残りを打力の高い順
  remain.sort((a, b) => bat(b.p) - bat(a.p));
  [6, 7, 8].forEach((n, i) => { if (remain[i]) spotOf.set(remain[i], n); });
  remain.forEach((e, i) => { if (!spotOf.has(e)) spotOf.set(e, 6 + i); });   // 念のため未割当を埋める
  return spotOf;
}

function autoFillTeamBuild(opts) {
  if (!TB_STATE) return;
  // opts.batters: 野手(現在オーダーのみ)を自動編成 / opts.pitchers: 投手も自動編成
  const doBatters  = !opts || opts.batters !== false;
  const doPitchers = !opts || opts.pitchers === true;
  // 選抜スコアの総合力。マイナー選手(穴埋め用)は大きく減点し、通常選手を必ず優先する
  //   (有限値の減点なので、通常選手で埋まらない枠には穴埋めとして選ばれる)。野手diamondScore・投手選抜の両方で共用。
  const ov   = p => (overallOf(p) || 0) - (isMinorPlayer(p) ? MINOR_FILL_PENALTY : 0);
  const stat = (p, k) => (p && p.stats && Number.isFinite(p.stats[k])) ? p.stats[k] : 0;
  const meatPow = p => stat(p, 'ミート') + stat(p, 'パワー');
  // PH評価 (MLB監督采配): 代打の仕事は一発と勝負どころの一打 → 長打力(HR能)と勝負強さを重視、選球眼も少し見る
  const phScore = p => stat(p, 'パワー') + stat(p, 'ミート') + stat(p, 'チャンス') * 1.5
    + ((p && p.statsMini && Number.isFinite(p.statsMini['HR能'])) ? p.statsMini['HR能'] : 0) * 8
    + stat(p, '選球眼') * 0.3;
  const speedOf = p => stat(p, 'スピード');
  const drsCount = p => (p && p.drs) ? p.drs.filter(d => (Number(d.innings) || 0) > 0).length : 0;
  const drsSum   = p => (p && p.drs) ? p.drs.reduce((s, d) => s + ((Number(d.innings) || 0) > 0 ? (d.value || 0) : 0), 0) : 0;
  const pad = (arr, n) => { const a = arr.slice(0, n); while (a.length < n) a.push(null); return a; };

  // 年代違いの同一選手を重複登録しないよう、自動編成の候補は同名(同種別)につき最新年度のみへ間引く
  const allBat = tbDedupePoolNewest(tbBatterPool());
  const allPit = tbDedupePoolNewest(tbPitcherPool());

  // ===== 野手: 9守備位置 (現在のオーダーのみ) =====
  // MLB監督の采配: 総合力を主軸にしつつ、センターライン(SS/2B/CF)は守備力を「実加点」で査定する。
  //   守備の要は失点を直接減らすため、多少打力が落ちても守れる選手を優先する (捕手はリード査定で別途考慮)。
  //   コーナー(1B/3B/LF/RF)は打撃優先で、守備は小さめの加点に留める。
  if (doBatters) {
  const usedBat = new Set();
  const DEF_TB = new Set(['2B', 'SS', 'CF', 'C']);   // 守備DRSをタイブレークに使うポジション
  const curOrderIdx = (TB_STATE && TB_STATE.currentOrder) || 0;
  const DIAMOND = ['C','1B','2B','3B','SS','LF','CF','RF','DH'];
  // 現在のオーダーに出場可能な野手だけを候補にする (オーダー制約)
  const cand = allBat.filter(p => playerAllowedInOrder(p, curOrderIdx));
  // 守備の実加点: センターライン=守備DRS×0.6 (DRS+10の遊撃手は総合力6点ぶんの価値) /
  //   コーナー=×0.25 (打力優先) / C=0 (リード×4+阻止率÷5 の捕手査定側で評価) / DH=0。
  const DEF_POS_W = { 'SS': 0.6, '2B': 0.6, 'CF': 0.6, '1B': 0.25, '3B': 0.25, 'LF': 0.25, 'RF': 0.25 };
  const posBonus = (p, pos) => {
    const drs = tbDrsValue(p, pos);
    const real = (DEF_POS_W[pos] || 0) * (Number.isFinite(drs) ? drs : 0);
    const tb = DEF_TB.has(pos) ? drs : meatPow(p);
    return real + (Number.isFinite(tb) ? tb : 0) / 100000;   // 微小タイブレークは従来通り
  };
  // 各ポジションで出場可能な選手の最高ミート+パワー (起用ボーナスの安全弁に使用)
  const maxMeatPowAt = {};
  DIAMOND.forEach(pos => {
    let mx = -Infinity;
    cand.forEach(p => { if (canPlay(p, pos)) mx = Math.max(mx, meatPow(p)); });
    maxMeatPowAt[pos] = (mx === -Infinity) ? 0 : mx;
  });
  // 起用ボーナス: 年度試合数10につき1ポイント (実際の出場=打数を加味し、数ポイント差なら試合数が多い選手を起用)。
  //   ただし、その位置で守備への悪影響が大きい / ミート+パワーがその位置の最高より15以上低い 場合は
  //   このボーナスを与えない (総合力のみで評価)。
  //   捕手(C)の守備判定は生のDRSでなく捕手評価(リード×4+阻止率÷5+DRS÷10)で行う
  //   — 打球処理(DRS)が悪くてもリードの良い正捕手を門前払いしないため。
  const usageBonus = (p, pos) => {
    const g = getCardGamesOf(p);
    if (!g) return 0;
    if (pos === 'C') {
      if (defenseRating(p, pos) <= 0) return 0;           // 捕手評価が0以下 → 適用しない
    } else {
      const drs = tbDrsValue(p, pos);
      if (Number.isFinite(drs) && drs <= -10) return 0;   // 守備への悪影響大 → 適用しない
    }
    if (meatPow(p) <= (maxMeatPowAt[pos] || 0) - 15) return 0;  // 攻撃力の低下大 → 適用しない
    return g / 10;
  };
  // 捕手は「リード×4 + 阻止率÷5」をレギュラー査定に加味 (リード重視のMLB監督采配)。
  //   さらに、捕手適性のある選手を「C以外」(DH・一塁など) に置く場合は tbDhCatcherPenalty で減点する
  //   — 「リードの良い捕手はマスクに回し、他の枠には別の選手を置く」方が総合点が高くなり、
  //   捕手をDHや一塁で消費する編成 (例: ペレス一塁+守備型捕手スタメン) を避ける。
  const diamondScore = (p, pos) => canPlay(p, pos)
    ? (ov(p) + usageBonus(p, pos) + posBonus(p, pos)
       + (pos === 'C' ? catcherAssessBonus(p) : -tbDhCatcherPenalty(p)))
    : -Infinity;
  // bestAssignment が扱えるサイズへ削減 (カバレッジ保証)。試合数・捕手査定込みで絞り込み、守備型捕手が漏れないようにする。
  const players = tbTrimForAssignment(cand, DIAMOND, p => ov(p) + ((getCardGamesOf(p) || 0) / 10) + catcherAssessBonus(p), 18);
  // 9ポジション全充足の最大総合力割当。埋められなければ部分最大充足。
  let assign = bestAssignment(DIAMOND, players, diamondScore);
  if (!assign) assign = tbGreedyMaxFill(DIAMOND, players, diamondScore);
  const newBatters = { C:null,'1B':null,'2B':null,'3B':null,SS:null,LF:null,CF:null,RF:null,DH:null };
  DIAMOND.forEach((pos, k) => {
    const pi = assign ? assign[k] : -1;
    if (pi != null && pi >= 0 && players[pi]) { newBatters[pos] = players[pi]; usedBat.add(players[pi]); }
  });
  // ===== DH是正: DHに守備の良い選手が入っていたら守備位置へ移し、守備で劣る選手をDHにする =====
  //   DH選手が あるフィールド守備位置で 現在の守備者よりDRSが高ければ、その位置へ入れ替える。
  //   これを繰り返し、最終的に「守備で最も貢献しない選手」がDHに残るよう選定しなおす。
  {
    const FIELD_POS = ['C','1B','2B','3B','SS','LF','CF','RF'];
    let guard = 0;
    while (guard++ < 10) {
      const dhP = newBatters['DH'];
      if (!dhP) break;
      let bestPos = null, bestGain = 0;
      for (const pos of FIELD_POS) {
        const fp = newBatters[pos];
        if (!fp || !canPlay(dhP, pos)) continue;   // DH選手がそのポジションを守れること
        // 守備評価は defenseRating: 捕手(C)はリード×4+阻止率÷5+DRS÷10 の捕手評価で比較 (他はDRS)。
        // 捕手価値の放棄(tbDhCatcherPenalty)の交換は C への移動時のみ加味する:
        //   ・C なら dhP の放棄が解消し fp の放棄が発生 → 差し引きを加算
        //   ・C 以外 (一塁など) は移動してもDHでも放棄したままで相殺 → 純粋な守備DRS比較
        //   (C以外にも加えると「捕手をDHから一塁へ」の移動を不当に促してしまう)
        const gain = (pos === 'C')
          ? (defenseRating(dhP, pos) + tbDhCatcherPenalty(dhP)) - (defenseRating(fp, pos) + tbDhCatcherPenalty(fp))
          : defenseRating(dhP, pos) - defenseRating(fp, pos);
        if (gain > bestGain) { bestGain = gain; bestPos = pos; }   // 総合改善が最大の位置
      }
      if (!bestPos) break;   // どの守備位置でも現守備者を上回らない → DHのまま (既定どおり)
      const fieldP = newBatters[bestPos];
      newBatters[bestPos] = dhP;   // 守備の良いDH選手をフィールドへ
      newBatters['DH'] = fieldP;   // 守備で劣る選手をDHへ
    }
  }
  // ===== 守備位置の最適化: 選ばれた9人(と後で計算する打順)はそのまま、守備位置の
  //   組み合わせだけをチーム合計DRSが最大になるよう再割当する
  //   (例: 三塁チゾム/二塁カブレラ → 二塁チゾム/三塁カブレラ で合計DRSが上がるなら入替) =====
  {
    const chosen = DIAMOND.filter(pos => newBatters[pos]).map(pos => ({ pos, p: newBatters[pos] }));
    if (chosen.length === DIAMOND.length) {
      const curPos = new Map(chosen.map(e => [e.p, e.pos]));
      const defOpt = (p, pos) => {
        if (!canPlay(p, pos)) return -Infinity;
        // 捕手(C)はリード×4+阻止率÷5+DRS÷10 の捕手評価、他ポジションは守備DRS (defenseRating)。
        // C以外(DH・一塁など)に捕手適性持ちを置く場合は「捕手価値の放棄」を減点
        //   (ペレスの一塁DRSがプラスでも、捕手価値24点超を捨ててまで一塁へ回さない)。
        const base  = (pos === 'DH') ? 0 : defenseRating(p, pos);
        const waste = (pos === 'C') ? 0 : tbDhCatcherPenalty(p);
        return base - waste + (curPos.get(p) === pos ? 0.001 : 0);   // 合計評価が同じなら現状の配置を維持
      };
      const ps = chosen.map(e => e.p);
      const assign = bestAssignment(DIAMOND, ps, defOpt);
      if (assign && assign.every(j => j >= 0)) {
        DIAMOND.forEach((pos, k) => { newBatters[pos] = ps[assign[k]]; });
      }
    }
  }
  // ===== 打順編成: 確定した9人に現代MLBの采配ロジックで 1〜9番を割り当てる (aiAssignBattingSpots 共通) =====
  const newOrder = { C:null,'1B':null,'2B':null,'3B':null,SS:null,LF:null,CF:null,RF:null,DH:null };
  {
    const entries = Object.keys(newBatters).filter(pos => newBatters[pos]).map(pos => ({ pos, p: newBatters[pos] }));
    const spots = aiAssignBattingSpots(entries);
    entries.forEach(e => { newOrder[e.pos] = spots.get(e) || null; });
  }

  // ===== 控え (PH1 / PH2 / PH3 / 代走 / 守備 / 守備) =====
  const takeBest = (scoreFn) => {
    const rem = allBat.filter(p => !usedBat.has(p));
    if (!rem.length) return null;
    rem.sort((a, b) => scoreFn(b) - scoreFn(a));
    usedBat.add(rem[0]); return rem[0];
  };
  // 守れるポジションの分類 + 守備力(守備DRS合計)
  const FIELD8 = new Set(['C','1B','2B','3B','SS','LF','CF','RF']);
  const INF = new Set(['1B','2B','3B','SS']);   // 内野
  const OUF = new Set(['LF','CF','RF']);        // 外野
  const defInfo = (p) => {
    const set = new Set((p.drs || []).filter(d => (Number(d.innings) || 0) > 0).map(d => d.pos).filter(x => FIELD8.has(x)));
    return {
      count: set.size,
      inf: [...set].some(x => INF.has(x)),
      ouf: [...set].some(x => OUF.has(x)),
      ss2b: set.has('SS') || set.has('2B'),   // 特に SS/2B を守れるか
      def: drsSum(p),                          // 守備力 = 守備DRS合計
    };
  };
  const HIGH_DEF = 3;   // 「守備力が高め」の目安 (守備DRS合計)
  // 守備1の選定: 指定の優先度カスケード。pool から1人返す (該当なければ最後は総合力順)。
  const pickDefenseSub = (pool) => {
    if (!pool.length) return null;
    // SS/2B可 → 守備力 → 守れるポジション数 → 総合力
    const sortDef = (a, b) => { const ia = defInfo(a), ib = defInfo(b);
      return (ib.ss2b - ia.ss2b) || (ib.def - ia.def) || (ib.count - ia.count) || (ov(b) - ov(a)); };
    const sortOvr = (a, b) => (ov(b) - ov(a)) || (defInfo(b).ss2b - defInfo(a).ss2b);
    const tiers = [
      { f: p => { const i = defInfo(p); return i.count >= 3 && i.inf && i.ouf && i.def >= HIGH_DEF; }, s: sortDef }, // (1) 内外野3+・守備高
      { f: p => { const i = defInfo(p); return i.count >= 3 && i.inf && i.def >= HIGH_DEF; },          s: sortDef }, // (2) 内野3+・守備高
      { f: p => { const i = defInfo(p); return i.count >= 3 && i.inf && i.ouf; },                      s: sortDef }, // (3) 内外野3+
      { f: p => { const i = defInfo(p); return i.count >= 3 && i.inf; },                               s: sortDef }, // (4) 内野3+
      { f: p => { const i = defInfo(p); return i.count >= 2 && i.inf && i.ouf; },                      s: sortOvr }, // (5) 内外野2+・総合力高
    ];
    for (const t of tiers) {
      const c = pool.filter(t.f);
      if (c.length) { c.sort(t.s); return c[0]; }
    }
    // (6) SS,CF,2B の順に DRS≥10 → 次に DRS≥3 の選手を探す
    for (const thr of [10, 3]) {
      for (const pos of ['SS', 'CF', '2B']) {
        const c = pool.filter(p => canPlay(p, pos) && tbDrsValue(p, pos) >= thr);
        if (c.length) { c.sort((a, b) => (tbDrsValue(b, pos) - tbDrsValue(a, pos)) || (ov(b) - ov(a))); return c[0]; }
      }
    }
    // それでも該当なし → 総合力順
    return pool.slice().sort((a, b) => ov(b) - ov(a))[0] || null;
  };
  const takePlayer = (p) => { if (p) usedBat.add(p); return p || null; };

  const ph1  = takeBest(phScore);    // PH1: パワー+ミート+チャンス×2
  const def1 = takePlayer(pickDefenseSub(allBat.filter(p => !usedBat.has(p))));  // 守備1: ユーティリティ守備固め
  // 守備2: 控え捕手。総合力＋捕手査定(リード×4+阻止率÷5)が高い選手 (第2捕手も配球力を重視)。
  //   捕手が居なければ守備1のロジックに準じる。
  let def2;
  { const rem = allBat.filter(p => !usedBat.has(p));
    const catchers = rem.filter(p => canPlay(p, 'C'))
      .sort((a, b) => (ov(b) + catcherAssessBonus(b)) - (ov(a) + catcherAssessBonus(a)));
    def2 = takePlayer(catchers.length ? catchers[0] : pickDefenseSub(rem)); }
  const pinchRun = takeBest(speedValue);// 代走: 走力総合(スピード+盗塁能×2+盗塁実績) — 盗塁できる走者を優先 (MLB采配)
  const ph2  = takeBest(phScore);    // PH2
  const ph3  = takeBest(phScore);    // PH3
  // pinchHitters の並び: [PH1, PH2, PH3, 代走, 守備, 守備]
  const newPinch = [ph1, ph2, ph3, pinchRun, def1, def2];
  // 現在のオーダーへ反映 (他のオーダーは変更しない)
  const _ord = TB_STATE.orders[TB_STATE.currentOrder] || TB_STATE.orders[0];
  _ord.batters = newBatters;
  _ord.batterOrder = newOrder;
  _ord.pinchHitters = pad(newPinch, TB_PH_COUNT);
  }

  // ===== 投手 (全オーダー共有) =====
  if (doPitchers) {
  const usedPit = new Set();
  // 指定年度 (TB_STATE.year)。null のときは年度区別なし。
  const tbYear = (TB_STATE && TB_STATE.year != null) ? TB_STATE.year : null;
  // 年度違いの予備候補: 指定年度の過去2年 かつ 総合力73以内 (指定年度指定時のみ該当)。
  //   これらは「主戦」(先発/抑え/SU/中継/MU1)には一切使わず、控え1・MU2・控え2〜4 の予備枠だけに回す。
  const isYearDiffReserve = (p) => tbYear != null && !!p && Number.isFinite(p.year)
    && p.year >= tbYear - 2 && p.year < tbYear && (overallOf(p) || 0) <= 73;
  const starterScore = (p) => (ov(p) || 0) + getRecoveryOf(p) + getInningsOf(p) / 5;
  // カードの肩書きが「先発投手」の選手。先発系の役割(先発→控え1→SU→控え2)へ優先的に回し、
  // 中継/抑え/MU には(スタミナ等の理由で)安易に入れない。
  const isStarterCard = (p) => !!p && typeof p.position === 'string' && p.position.indexOf('先発') >= 0;
  // 主戦プール: 年度違い予備を除外 (指定年度ありなら実質「指定年度のみ」。年度指定なしは全候補)。
  const mainPool = allPit.filter(p => !isYearDiffReserve(p));
  // 予備プール: 通常プール + 過去2年・総合力73以内 (重複除外)。控え系の枠で使う。
  //   年代違いの同一選手の重複を避けるため、最後に同名(同種別)最新年度のみへ間引く。
  const reservePool = (function () {
    if (tbYear == null) return allPit.slice();
    const extras = getPitchers()
      .filter(p => normalizeTeam(p.team) === normalizeTeam(TB_STATE.team) && !allPit.includes(p) && isYearDiffReserve(p));
    return tbDedupePoolNewest(allPit.concat(extras));
  })();
  // 主戦スロットの選定 (年度違い予備は使わない)
  const takeMain = (filterFn, n, scoreFn) => {
    const cmp = scoreFn ? (a, b) => scoreFn(b) - scoreFn(a) : (a, b) => ov(b) - ov(a);
    const rem = mainPool.filter(p => !usedPit.has(p) && filterFn(p)).sort(cmp).slice(0, n);
    rem.forEach(p => usedPit.add(p)); return rem;
  };
  // 主力枠(先発1-4・中継1-4・SU1,2・抑え)には回復量4以上の投手のみ自動登録する。
  //   (先発5・中継5・MU・控え は回復量の制限なし)
  const okRec = (p) => getRecoveryOf(p) >= 4;
  // 先発(5): 先発1-4 は回復量4以上の「先発投手」カードを優先→不足はスタミナ50以上・回復量7超の非カード。先発5 は回復量制限なし。
  const starters = (function () {
    const picks = [];
    picks.push(...takeMain(p => isStarterCard(p) && okRec(p), 4, starterScore));
    if (picks.length < 4) picks.push(...takeMain(p => !isStarterCard(p) && getStaminaOf(p) >= 50 && getRecoveryOf(p) > 7, 4 - picks.length, starterScore));
    if (picks.length < 5) {
      let five = takeMain(p => isStarterCard(p), 1, starterScore);
      if (!five.length) five = takeMain(p => !isStarterCard(p) && getStaminaOf(p) >= 50 && getRecoveryOf(p) > 7, 1, starterScore);
      picks.push(...five);
    }
    return picks;
  })();
  // 先発に入りきらなかった「先発投手」カード = 控え1 → MU → 控え2… の順に充当する待ち行列。
  //   マイナー選手は除外する。先発カードのマイナーがこの待ち行列経由で控え等へ流れ込み、
  //   減点比較(reservePool)を素通りして通常の中継投手より優先されるのを防ぐ。
  const leftoverStarters = mainPool.filter(p => !usedPit.has(p) && isStarterCard(p) && !isMinorPlayer(p)).sort((a, b) => starterScore(b) - starterScore(a));
  // 抑え(1): 先発カードは除外・回復量4以上・セーブ数最大 (同数なら総合力)
  let closer = [];
  { const rem = mainPool.filter(p => !usedPit.has(p) && !isStarterCard(p) && okRec(p));
    // 通常選手を最優先(マイナーは穴埋めのみ) → セーブ数 → 総合力 の順
    rem.sort((a, b) => ((isMinorPlayer(a)?1:0) - (isMinorPlayer(b)?1:0)) || getSeasonSavesOf(b) - getSeasonSavesOf(a) || ov(b) - ov(a));
    if (rem.length) { closer = [rem[0]]; usedPit.add(rem[0]); } }
  // ===== 予備(年度違い等)を拾うヘルパ =====
  //   「年度違い予備(過去2年・総合力73以内)」を最優先(先発適性スコア順)で入れ、居なければ指定年度の余りで埋める。
  const pickReserve = (minSta, excludeStarterCard) => {
    const pool = reservePool.filter(p => !usedPit.has(p) && getStaminaOf(p) >= minSta && !(excludeStarterCard && isStarterCard(p)));
    if (!pool.length) return null;
    const diff = pool.filter(isYearDiffReserve);
    const cands = (diff.length ? diff : pool).sort((a, b) => starterScore(b) - starterScore(a));
    usedPit.add(cands[0]); return cands[0];
  };
  // 控え1: 余った先発カードを最優先。居なければ従来の予備(年度違い等・スタミナ51以上)。
  let benchHead = null;
  if (leftoverStarters.length) { benchHead = leftoverStarters.shift(); usedPit.add(benchHead); }
  else { benchHead = pickReserve(51, true); }
  // SU(2): スタミナ40以下・回復量4以上・総合力順。先発カードはSUには登録しない。
  const setup = takeMain(p => !isStarterCard(p) && getStaminaOf(p) <= 40 && okRec(p), 2);
  // 中継: 先発カードは除外。中継1-4=回復量4以上・スタミナ50以下 / 中継5=回復量制限なし(中継1-4が埋まっている場合のみ)。
  const middle = takeMain(p => !isStarterCard(p) && getStaminaOf(p) <= 50 && okRec(p), 4);
  if (middle.length === 4) middle.push(...takeMain(p => !isStarterCard(p) && getStaminaOf(p) <= 50, 1));
  // MU(2): 控え1の次に、余った先発カードを優先的に充当する (先発→控え1→MU→控え2)。
  //   不足分は従来ロジック (MU1=スタミナ50以上・先発適性スコア順 / MU2=予備・年度違い優先)。
  const mop = (function () {
    const picks = [];
    while (picks.length < 2 && leftoverStarters.length) { const p = leftoverStarters.shift(); usedPit.add(p); picks.push(p); }
    if (picks.length < 2) { const m1 = takeMain(p => !isStarterCard(p) && getStaminaOf(p) >= 50, 1, starterScore)[0]; if (m1) picks.push(m1); }
    if (picks.length < 2) { const m2 = pickReserve(50, true); if (m2) picks.push(m2); }
    // MU1,2 を2まで充足 (控え2,3,4より優先)。スタミナ制約を緩め、先発カード以外を総合力順で。
    if (picks.length < 2) {
      const more = reservePool.filter(p => !usedPit.has(p) && !isStarterCard(p)).sort((a, b) => ov(b) - ov(a)).slice(0, 2 - picks.length);
      more.forEach(p => usedPit.add(p)); picks.push(...more);
    }
    return picks;
  })();              // MU1=idx0 / MU2=idx1 (位置を保持)
  // 控え1・MU1・MU2(ローテ谷間先発の枠)の中で、回復量5未満の投手がいれば、
  //   その「谷間先発が可能(先発カード or スタミナ50以上)」な投手の中で回復量が最小の選手を控え1に置く。
  //   (登板間隔が最も空く控え1=谷間先発に、最も回復の遅い投手を充てる運用)
  {
    const spotCapable = (p) => !!p && (isStarterCard(p) || getStaminaOf(p) >= 50);
    const trio = [];
    if (spotCapable(benchHead)) trio.push({ slot: 'bench', p: benchHead });
    if (spotCapable(mop[0]))    trio.push({ slot: 'mop0', p: mop[0] });
    if (spotCapable(mop[1]))    trio.push({ slot: 'mop1', p: mop[1] });
    if (benchHead && trio.some(t => getRecoveryOf(t.p) < 5)) {
      let minT = trio[0];
      trio.forEach(t => { if (getRecoveryOf(t.p) < getRecoveryOf(minT.p)) minT = t; });
      if (minT.slot !== 'bench') {   // 最小回復量が控え1以外 → 控え1と入れ替える
        const prev = benchHead;
        benchHead = minT.p;
        if (minT.slot === 'mop0') mop[0] = prev; else mop[1] = prev;
      }
    }
  }
  // 控え2〜4 を埋める。ただし「中継を5まで充足」を「控え3,4の充足」より優先する:
  //   (1)控え2を先に確保 → (2)中継を5まで補充(控え3,4より優先) → (3)残りで控え3,4を埋める。
  const benchCap = benchHead ? 3 : 4;   // benchRest が埋める枠数 (控え2 以降)
  const benchRest = [];
  const fillBenchTo = (n) => {
    while (benchRest.length < n && leftoverStarters.length) { const p = leftoverStarters.shift(); usedPit.add(p); benchRest.push(p); }
    if (benchRest.length < n) {
      const rem = reservePool.filter(p => !usedPit.has(p)).sort((a, b) => ov(b) - ov(a)).slice(0, n - benchRest.length);
      rem.forEach(p => usedPit.add(p)); benchRest.push(...rem);
    }
  };
  fillBenchTo(Math.min(1, benchCap));   // (1) 控え2 を確保
  // (2) 中継を5まで補充 (控え3,4より優先)。中継1-4は回復量4以上、中継5は制限なし。スタミナ制約は緩める。
  while (middle.length < 5) {
    const recReq = middle.length < 4;   // 1〜4番手枠は回復量4以上が必要
    const cand = reservePool.filter(p => !usedPit.has(p) && !isStarterCard(p) && (!recReq || okRec(p))).sort((a, b) => ov(b) - ov(a))[0];
    if (!cand) break;
    usedPit.add(cand); middle.push(cand);
  }
  fillBenchTo(benchCap);   // (3) 残りで控え3,4 を埋める
  const benchP = benchHead ? [benchHead, ...benchRest] : benchRest;
  TB_STATE.pitchers = {
    starter: pad(starters, 5),
    mop:     pad(mop, 2),
    middle:  pad(middle, 5),
    setup:   pad(setup, 2),
    closer:  pad(closer, 1),
    bench:   pad(benchP, 4),
  };
  }

  tbAliasCurrentOrder();
  renderTeamBuild();
}

// 永続化: チームごとに別キーで localStorage 保存
// 各チームのデータを独立して保持し、チーム切替で他チームの編成が消えないようにする
const TB_STORAGE_PREFIX = 'mlb_team_build_v1_';
const TB_LAST_TEAM_KEY  = 'mlb_team_build_last';
function tbStorageKey(team) { return TB_STORAGE_PREFIX + (team || 'original'); }

function saveTeamBuild() {
  if (!TB_STATE) return false;
  tbInvalidateRatingCal();  // 保存編成が変わる → 戦力評価の校正基準を次回参照時に再計算
  const idOf = playerIdOf;  // 名前+年+チーム+種別 を保存 (二刀流の投手/打者を区別)
  const serOrder = (o) => ({
    batters: Object.fromEntries(Object.entries(o.batters).map(([k,v]) => [k, idOf(v)])),
    batterOrder: o.batterOrder,
    pinchHitters: o.pinchHitters.map(idOf),
  });
  const orders = (TB_STATE.orders || []).map(serOrder);
  const o0 = orders[0] || serOrder(blankOrder());
  const payload = {
    team: TB_STATE.team,
    year: (TB_STATE.year == null) ? null : TB_STATE.year,   // 年度フィルタも保持
    orders,                       // オーダー1/2/3 (野手のみ)
    currentOrder: TB_STATE.currentOrder || 0,
    pitchers: Object.fromEntries(Object.entries(TB_STATE.pitchers).map(([role, arr]) => [role, arr.map(idOf)])),
    // 後方互換 & セットアップ連携用: オーダー1 を top-level にも保持
    batters: o0.batters,
    batterOrder: o0.batterOrder,
    pinchHitters: o0.pinchHitters,
  };
  const doSave = () => {
    localStorage.setItem(tbStorageKey(TB_STATE.team), JSON.stringify(payload));
    localStorage.setItem(TB_LAST_TEAM_KEY, TB_STATE.team);
  };
  try { doSave(); return true; }
  catch (e) {
    // 容量超過からの自動回復: ゲームが使わないデータを空けて再試行する。
    const trySave = () => { try { doSave(); return true; } catch (_) { return false; } };
    try {
      // (1) idbモードなら冗長なカード控え(EXTRAS_KEY: idbモードでは読まれない)を削除。
      //     ※lsモードでは EXTRAS_KEY がカード本体なので削除しない (誤消去防止)。
      const idbMode = !!(window.CardStore && window.CardStore.mode && window.CardStore.mode() === 'idb');
      if (idbMode && localStorage.getItem(EXTRAS_KEY) != null) { localStorage.removeItem(EXTRAS_KEY); if (trySave()) return true; }
      // (2) カード生成ツールの写真キャッシュ(mlbcard_teamphoto_*)を必要な分だけ削除して空きを作る。
      //     これはゲームでは未使用で、各カードの写真はカード自体に保持されているため失われない。
      const photoKeys = [];
      for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k && k.indexOf('mlbcard_teamphoto_') === 0) photoKeys.push(k); }
      for (const pk of photoKeys) { localStorage.removeItem(pk); if (trySave()) return true; }
    } catch (e2) { console.error('saveTeamBuild recover', e2); }
    console.error('saveTeamBuild', e); return false;
  }
}
// localStorage の使用状況の内訳 (保存失敗時の診断用)
function tbStorageDiag() {
  let total = 0; const sizes = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      const sz = ((localStorage.getItem(k) || '').length + k.length) * 2;   // UTF-16 概算バイト
      total += sz; sizes.push([k, sz]);
    }
  } catch (e) {}
  sizes.sort((a, b) => b[1] - a[1]);
  const top = sizes.slice(0, 6).map(([k, sz]) => '  ' + k + ': ' + (sz / 1024).toFixed(0) + 'KB').join('\n');
  const mode = (window.CardStore && window.CardStore.mode) ? window.CardStore.mode() : '?';
  const nCards = (window.CardStore && window.CardStore.getCachedSync) ? window.CardStore.getCachedSync().length : '?';
  return '保存先モード: ' + mode + ' / カード枚数: ' + nCards + '\n'
       + 'localStorage合計: ' + (total / 1024).toFixed(0) + 'KB\n大きい項目:\n' + top;
}

const TB_ORDER_COUNT = 3;   // オーダー1/2/3 (野手のみ。投手は共有)
// 1オーダー分の野手データ (打順 + 控え)
function blankOrder() {
  return {
    batters:     { C:null,'1B':null,'2B':null,'3B':null,'SS':null,'LF':null,'CF':null,'RF':null,'DH':null },
    batterOrder: { C:null,'1B':null,'2B':null,'3B':null,'SS':null,'LF':null,'CF':null,'RF':null,'DH':null },
    pinchHitters: new Array(TB_PH_COUNT).fill(null),
  };
}
// TB_STATE.batters/batterOrder/pinchHitters を現在オーダーへのエイリアスにする
// (既存の描画・操作コードはこれらを参照・変更するだけで現在オーダーに反映される)
function tbAliasCurrentOrder() {
  if (!TB_STATE || !TB_STATE.orders) return;
  const o = TB_STATE.orders[TB_STATE.currentOrder] || TB_STATE.orders[0];
  TB_STATE.batters = o.batters;
  TB_STATE.batterOrder = o.batterOrder;
  TB_STATE.pinchHitters = o.pinchHitters;
}
// オーダー切り替え (野手の打順・控えのみ切替。投手は共有)
function tbSwitchOrder(i) {
  if (!TB_STATE || !TB_STATE.orders) return;
  TB_STATE.currentOrder = Math.max(0, Math.min(TB_ORDER_COUNT - 1, i));
  tbAliasCurrentOrder();
  renderTeamBuild();
}
// ===== チーム戦力の4カテゴリ評価 =====
//   「素点計算 (tbRawRatings)」+「全球団相対の自動校正 (tbCalibratedScore)」の2段構え。
//   校正: オリジナルとSTLを除く29球団の保存編成(オーダー1)から各カテゴリの分布を取り、
//     『中央値のチーム→71点(B) / 上位3チームの平均→100点(S帯の中心)』となる線形換算に載せる。
//     → 上位5チーム前後が100点前後のS評価になり、素点のばらつき(機動力115点等)の影響を受けない。
//   通常チームは15〜100点、オリジナル/年代不問オールスターは15〜120点にクランプ (100超=SS)。
//   校正データ不足(保存編成8チーム未満)の場合は素点をそのまま表示。
function tbRawRatings(batters, pitchers) {
  const st = (p, k, d) => (p && p.stats && Number.isFinite(p.stats[k])) ? p.stats[k] : d;
  const mi = (p, k) => (p && p.statsMini && Number.isFinite(p.statsMini[k])) ? p.statsMini[k] : 0;
  batters = batters || {};
  const POSL = ['C','1B','2B','3B','SS','LF','CF','RF','DH'];
  let offSum = 0, mobSum = 0, defPts = 0;
  for (const pos of POSL) {
    const p = batters[pos];
    if (!p) { offSum += 40; mobSum += 40; if (pos !== 'DH') defPts -= 5; continue; }
    offSum += st(p,'ミート',50)*0.35 + st(p,'パワー',50)*0.35 + st(p,'選球眼',50)*0.15 + st(p,'チャンス',50)*0.15 + mi(p,'HR能')*1.5;
    mobSum += st(p,'スピード',50) + mi(p,'盗塁能')*3;
    if (pos !== 'DH') defPts += defenseRating(p, pos);
  }
  const pit = pitchers || {};
  const avg = arr => { const a = (arr||[]).filter(Boolean).map(x => overallOf(x) || 0); return a.length ? a.reduce((s,v)=>s+v,0)/a.length : 40; };
  return {
    pitching: avg(pit.starter)*0.55 + avg([...(pit.setup||[]), ...(pit.closer||[])])*0.30 + avg(pit.middle)*0.15,
    offense:  offSum / 9,
    defense:  55 + defPts * 0.45,
    mobility: mobSum / 9,
  };
}
// MLB30球団(オリジナル除く)の保存編成から校正基準(5番目に強いチーム・中央値)を作る。遅延計算でキャッシュ。
let TB_RATING_CAL = null;
function tbInvalidateRatingCal() { TB_RATING_CAL = null; }
function tbComputeRatingCal() {
  const teams = MLB_DIVISIONS.flatMap(d => d.teams.map(t => t[0]));   // 30球団
  const cats = { pitching: [], offense: [], defense: [], mobility: [] };
  for (const code of teams) {
    let tb = null;
    try { tb = getSavedTeamBuild(code, 0); } catch (e) { tb = null; }
    if (!tb || !tb.batters) continue;
    const filled = Object.values(tb.batters).filter(Boolean).length;
    if (filled < 7) continue;   // ほぼ空の編成はノイズとして除外
    const r = tbRawRatings(tb.batters, tb.pitchers);
    for (const k in cats) cats[k].push(r[k]);
  }
  const cal = {};
  for (const k in cats) {
    const arr = cats[k].slice().sort((a, b) => b - a);   // 素点の高い順
    // p5 = 5番目に強いチームの素点 (S下限90のアンカー) / med = 中央値 (B下限70のアンカー)
    cal[k] = (arr.length >= 8)
      ? { p5: arr[Math.min(4, arr.length - 1)], med: arr[Math.floor(arr.length / 2)] }
      : null;   // データ不足 → 素点表示
  }
  return cal;
}
// 素点 → 校正済み点数。アンカー: 中央値のチーム→70点(B下限) / 5番目に強いチーム→90点(S下限)。
//   → 各項目で上位5チーム前後がS(90〜100)、中位はB、下位はC以下に分布する。
//   分布が密集していても換算が暴れないよう分母(p5-med)は最低6を確保 (傾きの上限≈3.3)。
//   クランプ: 通常チーム 15〜100 / オリジナル・年代不問オールスター 15〜120 (100超=SS)。
function tbCalibratedScore(cat, raw, allowOver) {
  if (!TB_RATING_CAL) TB_RATING_CAL = tbComputeRatingCal();
  const c = TB_RATING_CAL[cat];
  let v = raw;
  if (c) v = 70 + (raw - c.med) * 20 / Math.max(c.p5 - c.med, 6);
  return Math.max(15, Math.min(allowOver ? 120 : 100, v));
}
// 現在オーダーの4カテゴリ評価 (全球団相対の校正済み)。
//   オリジナルチーム or 年度フィルタ「指定なし」(=年代を問わないオールスター編成) は100超え(SS)を許可する。
function tbTeamRatings() {
  const raw = tbRawRatings((TB_STATE && TB_STATE.batters) || {}, (TB_STATE && TB_STATE.pitchers) || {});
  const allowOver = !!(TB_STATE && (TB_STATE.team === 'original' || TB_STATE.year == null));
  return {
    pitching: tbCalibratedScore('pitching', raw.pitching, allowOver),
    offense:  tbCalibratedScore('offense',  raw.offense, allowOver),
    defense:  tbCalibratedScore('defense',  raw.defense, allowOver),
    mobility: tbCalibratedScore('mobility', raw.mobility, allowOver),
  };
}
// 点数(整数表示値) → 評価。100超=SS(レインボー、オリジナル/オールスターのみ)。
//   S:90〜100 / A:80〜89 / B:70〜79 / C:60〜69 / D:50〜59 / E:40〜49 / F:30〜39 / G:1〜29。
function tbGradeOf(v) {
  return v >= 101 ? 'SS' : v >= 90 ? 'S' : v >= 80 ? 'A' : v >= 70 ? 'B' : v >= 60 ? 'C' : v >= 50 ? 'D' : v >= 40 ? 'E' : v >= 30 ? 'F' : 'G';
}

// 「全change」: オーダー1〜3の野手(控え込み)と共有の投手陣を、年度設定に従ってすべて自動編成し直す。
//   各オーダーで登録可能な選手の制約(年度・オーダー特例)が異なるため、
//   オーダーを順に切り替えながら1つずつ自動編成する。投手は共有なので最初の1回のみ。
function tbAutoFillAllOrders() {
  if (!TB_STATE || !TB_STATE.orders) return;
  const cur = TB_STATE.currentOrder || 0;
  for (let i = 0; i < TB_ORDER_COUNT; i++) {
    TB_STATE.currentOrder = i;
    tbAliasCurrentOrder();
    autoFillTeamBuild({ batters: true, pitchers: i === 0 });
  }
  TB_STATE.currentOrder = cur;   // 元のオーダー表示へ戻す
  tbAliasCurrentOrder();
  renderTeamBuild();
}

function blankTeamState(team) {
  const orders = [];
  for (let i = 0; i < TB_ORDER_COUNT; i++) orders.push(blankOrder());
  const st = {
    team: team || 'LAD',
    // 年度フィルタ: null=指定なし。オリジナルは指定なし、その他チームは既定2025年。
    year: (team === 'original' || !team) ? null : 2025,
    orders,
    currentOrder: 0,
    pitchers: {
      starter: new Array(5).fill(null),
      mop:     new Array(2).fill(null),
      middle:  new Array(5).fill(null),
      setup:   new Array(2).fill(null),
      closer:  new Array(1).fill(null),
      bench:   new Array(4).fill(null),
    },
  };
  // 現在オーダー(オーダー1)へのエイリアス
  st.batters = orders[0].batters;
  st.batterOrder = orders[0].batterOrder;
  st.pinchHitters = orders[0].pinchHitters;
  return st;
}

function loadTeamBuild(specificTeam) {
  // 引数指定 > 最後に使ったチーム > 'LAD' の順でロード
  const team = specificTeam || localStorage.getItem(TB_LAST_TEAM_KEY) || 'LAD';
  let payload = null;
  try { payload = JSON.parse(localStorage.getItem(tbStorageKey(team)) || 'null'); }
  catch (e) { /* ignore */ }
  // 保存データの team とキーの team が不一致なら、旧バージョン由来の混入データ
  // (例: オリジナルのキーに LAD の編成が入っている) とみなして破棄し、空から始める。
  // これによりチームごとに必ず独立した編成になる。
  // 別名(SF/SFG/SFN等)を吸収して比較。保存teamとキーteamが同一チームなら一致とみなし、誤って削除しない。
  const teamMatches = payload && normalizeTeam(payload.team) === normalizeTeam(team);
  if (payload && !teamMatches) {
    try { localStorage.removeItem(tbStorageKey(team)); } catch (e) {}
    payload = null;
  }
  if (!payload) {
    TB_STATE = blankTeamState(team);
    return false;
  }
  const all = [...getBatters(), ...getPitchers()];
  const lookup = (id) => {
    if (!id) return null;
    return all.find(p => playerMatchesId(p, id)) || null;
  };
  TB_STATE = blankTeamState(team);
  // 1オーダー分のデータ(src)を order オブジェクトへ読み込む
  const loadOrderInto = (order, src) => {
    if (!src) return;
    if (src.batterOrder) for (const k of Object.keys(order.batterOrder)) order.batterOrder[k] = src.batterOrder[k] ?? null;
    for (const k of Object.keys(order.batters)) order.batters[k] = lookup(src.batters && src.batters[k]);
    const ph = (src.pinchHitters || []).map(lookup);
    for (let i = 0; i < TB_PH_COUNT; i++) order.pinchHitters[i] = ph[i] || null;
  };
  if (Array.isArray(payload.orders)) {
    for (let i = 0; i < TB_ORDER_COUNT; i++) loadOrderInto(TB_STATE.orders[i], payload.orders[i]);
  } else {
    // 旧形式 (単一オーダー): top-level の野手データをオーダー1として読み込む
    loadOrderInto(TB_STATE.orders[0], payload);
  }
  // 投手 (全オーダー共有)
  for (const role of Object.keys(TB_STATE.pitchers)) {
    const saved = payload.pitchers && payload.pitchers[role];
    if (Array.isArray(saved)) {
      for (let i = 0; i < TB_STATE.pitchers[role].length; i++) {
        TB_STATE.pitchers[role][i] = lookup(saved[i]);
      }
    }
  }
  TB_STATE.currentOrder = Math.max(0, Math.min(TB_ORDER_COUNT - 1, payload.currentOrder || 0));
  // 年度フィルタの復元 (保存があればそれを使う。null=指定なしも有効値)
  if (Object.prototype.hasOwnProperty.call(payload, 'year')) TB_STATE.year = payload.year;
  tbAliasCurrentOrder();
  return true;
}

// 年度フィルタ: 選択年度(TB_STATE.year)に対し、その選手が選択候補になれるか。
//   ・指定なし(null) → 常に可
//   ・選択年度 Y → 同年(py===Y)は可
//   ・低総合力の選手は2年前まで遡って可 (Y-2 ≤ py < Y)
//       遡り可能な総合力しきい値: オーダー1は70以下 / オーダー2・3は65以下
//   ・未来の年度(py>Y) や 年度不明は不可
function tbYearAllowed(p) {
  if (!TB_STATE || TB_STATE.year == null) return true;
  if (isAllTeamPlayer(p)) return true;          // 所属ALLは全年度で登録可
  const y = TB_STATE.year;
  const py = p && p.year;
  if (!Number.isFinite(py)) return false;
  if (py === y) return true;
  const thresh = (TB_STATE.currentOrder === 0) ? 70 : 65;   // オーダー1のみ70以下
  if ((overallOf(p) || 0) <= thresh && py >= y - 2 && py < y) return true;
  return false;
}
// 年度選択を変更
function tbSetYear(val) {
  if (!TB_STATE) return;
  TB_STATE.year = (val === '' || val == null) ? null : parseInt(val, 10);
  renderTeamBuild();
}
// チーム + 年度でフィルタした選手リスト
function tbBatterPool() {
  let pool = getBatters();
  if (!TB_STATE) return pool;
  if (!(TB_STATE.team === 'original' || !TB_STATE.team)) {
    pool = pool.filter(p => isAllTeamPlayer(p) || normalizeTeam(p.team) === normalizeTeam(TB_STATE.team));
  }
  return pool.filter(tbYearAllowed);
}
function tbPitcherPool() {
  let pool = getPitchers();
  if (!TB_STATE) return pool;
  if (!(TB_STATE.team === 'original' || !TB_STATE.team)) {
    pool = pool.filter(p => isAllTeamPlayer(p) || normalizeTeam(p.team) === normalizeTeam(TB_STATE.team));
  }
  return pool.filter(tbYearAllowed);
}

// 重複チェック: 全スロット (打順/PH/投手) で選択済の選手キーを集める。
// キーは種別を含む playerKey なので、二刀流(投手版/打者版)は別選手として扱われ、
// 投手で登録しても打者での登録が妨げられない (逆も同様)。
function tbUsedKeys(excludeKey) {
  const keys = new Set();
  const addKey = (p) => { if (p) keys.add(playerKey(p)); };
  Object.values(TB_STATE.batters).forEach(addKey);
  TB_STATE.pinchHitters.forEach(addKey);
  Object.values(TB_STATE.pitchers).forEach(arr => arr.forEach(addKey));
  if (excludeKey) keys.delete(excludeKey);
  return keys;
}

// 年代違いの同一選手(同名・同種別)の重複を防いだ候補リストを返す共通処理。
//   pool   : 既にチーム/年度/守備位置/役割等で絞り込んだ候補配列。
//   curKey : この枠に現在いる選手のキー (常に候補へ残す)。
//   ルール: ・他スロットに同名(同種別)が登録済み → その名前は全年度を候補から除外
//           ・残った候補は同名(同種別)につき最新年度のみ残す (例: 2025があれば2024/2023は出さない)
//   ※二刀流(投手版/打者版)は種別が違うので別選手として扱う。
function tbSameNameKey(p) { return (p && p.fullNameTop || '') + '|' + playerType(p); }
function tbFilterCandidates(pool, curKey) {
  if (!TB_STATE) return pool;
  const used = tbUsedKeys(curKey);   // 他スロット登録済み(完全一致キー)。現選手は除外済み。
  const usedNames = new Set();       // 他スロット登録済みの「名前(同種別)」
  const collect = (p) => { if (p && playerKey(p) !== curKey) usedNames.add(tbSameNameKey(p)); };
  Object.values(TB_STATE.batters).forEach(collect);
  (TB_STATE.pinchHitters || []).forEach(collect);
  Object.values(TB_STATE.pitchers).forEach(arr => arr.forEach(collect));
  const newest = new Map();          // 名前(同種別) → 最新年度の選手
  for (const p of pool) {
    if (playerKey(p) === curKey) continue;          // 現選手は最後に必ず残す
    if (used.has(playerKey(p))) continue;           // 完全一致の重複
    const nk = tbSameNameKey(p);
    if (usedNames.has(nk)) continue;                // 同名(年代違い)が他枠で登録済み
    const c = newest.get(nk);
    if (!c || (Number(p.year) || 0) > (Number(c.year) || 0)) newest.set(nk, p);
  }
  const keep = new Set(newest.values());
  return pool.filter(p => playerKey(p) === curKey || keep.has(p));
}
// 自動編成用: 同名(同種別)につき最新年度のみへ間引いたプールを返す (重複登録防止)。
function tbDedupePoolNewest(pool) {
  const newest = new Map();
  for (const p of pool) {
    const nk = tbSameNameKey(p);
    const c = newest.get(nk);
    if (!c || (Number(p.year) || 0) > (Number(c.year) || 0)) newest.set(nk, p);
  }
  const keep = new Set(newest.values());
  return pool.filter(p => keep.has(p));
}

// プレイヤーから DRS 値 (該当ポジション) を取得
function tbDrsValue(player, pos) {
  if (!player || !player.drs) return 0;
  const d = player.drs.find(x => x.pos === pos);
  return d ? (d.value || 0) : 0;
}

// 捕手のレギュラー査定への加点: リード×4 + 阻止率÷5 (総合力換算)。
//   ゲームエンジン分析による重み: リード1点はヒット判定(hitPt)を1下げ、相手の全打席(≈38/試合)に
//   作用して約0.10失点/試合の防御価値。打者の総合力1点(自分の約4.3打席のみ)≈0.025点/試合なので
//   リード1点 ≈ 総合力4点。阻止率は盗塁企図(約1.2回/試合)×成功率変化(0.8%/点)＋抑止効果 ≈ ÷5。
//   打力型捕手を正捕手にできれば DH に別の強打者を置けてチーム力が上がる — この損得は
//   bestAssignment(全体最適)が自動で解決するので、ここでは捕手固有の価値だけを加点する。
function catcherAssessBonus(player) {
  if (!player || !player.catcher || !canPlay(player, 'C')) return 0;
  const lead = Number(player.catcher['リード']);
  const arm  = Number(player.catcher['阻止率']);
  return (Number.isFinite(lead) ? lead * 4 : 0) + (Number.isFinite(arm) ? arm / 5 : 0);
}
// 捕手適性のある選手を「C以外」(DH・一塁など他の守備位置) で使う時の
//   「捕手価値の放棄」ペナルティ (MLB監督の采配)。
//   リード/阻止率の良い捕手をC以外に置くと、正捕手として持つ価値(リード×4+阻止率÷5)を
//   捨てた上に他の選手の枠まで塞ぐ (彼をマスクに回せば他の枠に別の選手を置ける)。
//   本職捕手 (カードのポジションが「捕手」) は固定12点 + 査定×0.8 と特に重く減点し、
//   それ以外の捕手適性持ちは査定×0.6。査定のマイナス分は0扱い (下手な捕手の転用は自由)。
function tbDhCatcherPenalty(player) {
  const cab = Math.max(0, catcherAssessBonus(player));
  const primaryC = !!(player && typeof player.position === 'string'
    && player.position.indexOf('捕手') >= 0 && canPlay(player, 'C'));
  return primaryC ? (12 + cab * 0.8) : cab * 0.6;
}

// 守備の総合評価 (守備固め判定・守備位置最適化用)。基本は守備DRS。
// 捕手は打球処理(DRS)より配球が重要: 守備DRSは10につき1pt・リードは1につき4pt・阻止率は5につき1pt。
//   (リード×4/阻止率÷5 はゲームエンジンの失点影響分析による総合力換算値 — catcherAssessBonus 参照)
function defenseRating(player, pos) {
  if (!player) return -Infinity;
  const drs = tbDrsValue(player, pos) || 0;
  if (pos === 'C' && player.catcher) {
    const lead = Number(player.catcher['リード']);
    const arm  = Number(player.catcher['阻止率']);
    return (drs / 10)                                    // 守備DRS: 10につき1pt (捕手の打球処理は影響小)
      + (Number.isFinite(lead) ? lead * 4 : 0)           // リード: 1につき4pt (≈0.10失点/試合の防御価値)
      + (Number.isFinite(arm)  ? arm / 5 : 0);           // 阻止率: 5につき1pt
  }
  return drs;
}

// チーム総合力を算出する (チーム編成画面のダイヤモンドに表示)。
// 計算式:
//   打順1〜9の総合力の「和」
// + 控え(PH/代走/守備固め)の平均 [いなければ 打順1〜9の平均]
// + 先発の平均 × 3
// + 中継の平均 × 2 [中継いなければ SU平均、SUもいなければ 先発平均]
// + SU の平均 × 1   [SUいなければ 中継平均、中継もいなければ 先発平均]
// + 抑え の総合力    [抑えいなければ 中継平均、中継もいなければ 先発平均]
// + モップ+控え投手の平均 [いなければ 先発平均]
function tbTeamOverall() {
  if (!TB_STATE) return null;
  const ov = p => { const v = overallOf(p); return Number.isFinite(v) ? v : null; };
  const avg = arr => {
    const vals = (arr || []).map(ov).filter(v => v !== null);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  // 打順1〜9 の選手 (batterOrder が 1〜9 の守備位置の選手)
  const orderPlayers = [];
  for (let n = 1; n <= 9; n++) {
    for (const pos of TB_DIAMOND_POSITIONS) {
      if (TB_STATE.batterOrder[pos] === n && TB_STATE.batters[pos]) { orderPlayers.push(TB_STATE.batters[pos]); break; }
    }
  }
  const orderVals = orderPlayers.map(ov).filter(v => v !== null);
  const orderSum = orderVals.reduce((a, b) => a + b, 0);
  const orderAvg = orderVals.length ? orderSum / orderVals.length : null;
  // 控え打者 (PH/代走/守備固め)
  const benchBat = (TB_STATE.pinchHitters || []).filter(Boolean);
  const benchBatAvg = benchBat.length ? avg(benchBat) : orderAvg;
  // 投手陣
  const P = TB_STATE.pitchers || {};
  const starterAvg = avg(P.starter);
  const middleRaw  = avg(P.middle);
  const setupRaw   = avg(P.setup);
  const closerVals = (P.closer || []).map(ov).filter(v => v !== null);
  const mopBench   = [...(P.mop || []), ...(P.bench || [])].filter(Boolean);
  // フォールバック付き各値
  const middleAvg = (middleRaw !== null) ? middleRaw : (setupRaw !== null ? setupRaw : starterAvg);
  const setupAvg  = (setupRaw  !== null) ? setupRaw  : (middleRaw !== null ? middleRaw : starterAvg);
  const closerVal = closerVals.length ? closerVals[0] : (middleRaw !== null ? middleRaw : starterAvg);
  const mopBenchAvg = mopBench.length ? avg(mopBench) : starterAvg;
  const z = v => (v === null ? 0 : v);
  const total = orderSum
              + z(benchBatAvg)
              + z(starterAvg) * 3
              + z(middleAvg) * 2
              + z(setupAvg) * 1
              + z(closerVal)
              + z(mopBenchAvg);
  return Math.round(total);
}

// ====== オーダー制約 (カード裏面「試合」数による出場可能オーダーの限定) ======
// 試合数の取得結果はキャッシュ (playerKey 単位)。rawHtml 解析はやや重いため。
const _gamesPlayedCache = new Map();
// rawHtml(カード裏面の年度別成績テーブル)から、該当年の「試合」数を抽出する。
function extractGamesFromRawHtml(html, year) {
  try {
    const ystr = String(year);
    const trRe = /<tr>([\s\S]*?)<\/tr>/g;
    let gi = 2;   // 既定の「試合」列インデックス (年度/チーム/試合…)
    let m;
    while ((m = trRe.exec(html)) !== null) {
      const row = m[1];
      if (/<th/i.test(row)) {
        // ヘッダー行: 「試合」列のインデックスを特定
        const ths = (row.match(/<th[^>]*>([\s\S]*?)<\/th>/g) || [])
          .map(s => s.replace(/<[^>]*>/g, '').trim());
        const idx = ths.indexOf('試合');
        if (idx >= 0) gi = idx;
        continue;
      }
      if (!/class="cb-year"/.test(row)) continue;   // 年度別成績の行のみ
      const tds = (row.match(/<td[^>]*>([\s\S]*?)<\/td>/g) || [])
        .map(s => s.replace(/<[^>]*>/g, '').trim());
      if (tds[0] === ystr) {
        const n = parseInt(String(tds[gi] || '').replace(/[^0-9]/g, ''), 10);
        return Number.isFinite(n) ? n : null;
      }
    }
  } catch (e) {}
  return null;
}
// 選手の「該当年の試合数」を返す。構造化データ → rawHtml の順で探す。取得不能なら null。
function getGamesPlayedOf(p) {
  if (!p) return null;
  const key = playerKey(p);
  if (_gamesPlayedCache.has(key)) return _gamesPlayedCache.get(key);
  const tryNum = v => { const n = parseInt(String(v).replace(/[^0-9]/g, ''), 10); return Number.isFinite(n) ? n : null; };
  let g = null;
  if (p.gamesPlayed != null) g = tryNum(p.gamesPlayed);
  if (g == null && p.record && p.record['試合'] != null) g = tryNum(p.record['試合']);
  if (g == null && p.stats && p.stats['試合'] != null) g = tryNum(p.stats['試合']);
  if (g == null && p.rawHtml && p.year != null) g = extractGamesFromRawHtml(p.rawHtml, p.year);
  _gamesPlayedCache.set(key, g);
  return g;
}
// 試合数 → 出場可能なオーダーのインデックス集合 (0=オーダー1 / 1=オーダー2 / 2=オーダー3)。
// 試合数が不明(null)の場合は制約を課さず全オーダー可とする。
function ordersAllowedByGames(games) {
  if (games == null) return [0, 1, 2];
  if (games >= 150) return [0, 1, 2];   // オーダー1,2,3
  if (games >= 125) return [0, 1];      // オーダー1,2
  if (games >= 100) return [0, 2];      // オーダー1,3
  if (games >= 80)  return [0];         // オーダー1
  if (games >= 60)  return [1, 2];      // オーダー2,3
  if (games >= 40)  return [1];         // オーダー2
  return [2];                           // オーダー3
}
// 守備DRSを持つ(=守れる)守備位置の数。FIELD8(C/内野/外野)で innings>0 のもの。
function drsPositionCount(p) {
  if (!p || !p.drs) return 0;
  const FIELD = new Set(['C','1B','2B','3B','SS','LF','CF','RF']);
  const set = new Set(p.drs.filter(d => (Number(d.innings) || 0) > 0 && FIELD.has(d.pos)).map(d => d.pos));
  return set.size;
}
// 選手 p が orderIdx(0-2) のレギュラー候補になれるか
function playerAllowedInOrder(p, orderIdx) {
  const ovr = overallOf(p) || 0;
  // 制約の例外 (全オーダー可):
  //   ・総合力が59以下の選手
  //   ・守備DRSを3ポジション以上持つユーティリティ選手 で 総合力が63以下
  //   ・捕手 (C守備可) で 総合力が59以下 (第2捕手の確保用)
  if (ovr <= 59) return true;
  if (drsPositionCount(p) >= 3 && ovr <= 63) return true;
  if (canPlay(p, 'C') && ovr <= 59) return true;
  return ordersAllowedByGames(getGamesPlayedOf(p)).includes(orderIdx);
}

// 守備位置 pos 用に出せるバッターのリスト (チーム + DRS フィルタ + 重複除外 + オーダー制約)
function tbEligibleBatters(pos, exclude) {
  const pool = tbBatterPool();
  const oi = (TB_STATE && TB_STATE.currentOrder) || 0;
  const filtered = pool.filter(p => {
    const k = playerKey(p);
    if (!canPlay(p, pos)) return false;
    // オーダー制約: 現在のオーダーに出場可能な選手のみ。
    // ただし既にこの枠に割り当て済みの選手(exclude)は常に候補に残す。
    if (k !== exclude && !playerAllowedInOrder(p, oi)) return false;
    return true;
  });
  // 重複(完全一致)除外 + 年代違い同名の最新年度のみ + 現選手は常に残す
  return tbFilterCandidates(filtered, exclude);
}

// 守備位置 pos のプルダウンでの「AI推奨スコア」。総合力を主軸に、その位置の守備DRSを加味する。
//   捕手はリード・阻止率(catcherAssessBonus)も加味する。値が大きいほど上位に表示する。
function tbRecScore(p, pos) {
  let s = overallOf(p) || 0;
  if (pos === 'C') s += catcherAssessBonus(p);          // 捕手はリード・阻止率を加味
  const drs = tbDrsValue(p, pos);
  if (Number.isFinite(drs)) s += drs * 0.5;             // 守備適性(DRS)を加味
  return s;
}
// 守備位置 pos のプルダウン候補。控え(PH)・他守備位置のスタメン・フリー(未登録)をすべて含める。
//   返り値: [{ p, benchIdx, fromPos, drs }]
//     benchIdx != null … 控え枠の選手 (選択時は控えと入替)
//     fromPos  != null … 他守備位置のスタメン (選択時はポジション入替。入替不可なら元位置を空に)
//     どちらも null    … フリー(未登録)の選手 (選択時はそのまま配置)
//   現選手(curKey)は守備可否に関わらず必ず候補に残す。AI推奨順(推奨スコア→DRS)でソートして返す。
function tbDiamondCandidates(pos, curKey) {
  const oi = (TB_STATE && TB_STATE.currentOrder) || 0;
  const out = [];
  const seen = new Set();
  const tryAdd = (p, extra) => {
    if (!p) return;
    const k = playerKey(p);
    if (seen.has(k)) return;
    if (k !== curKey) {                                  // 現選手以外は「守れる + 現オーダー登録可」を要求
      if (!canPlay(p, pos) || !playerAllowedInOrder(p, oi)) return;
    }
    seen.add(k);
    out.push(Object.assign({ p, benchIdx: null, fromPos: null }, extra || {}));
  };
  // 1) 控え(PH)
  (TB_STATE.pinchHitters || []).forEach((p, idx) => { if (p && playerKey(p) !== curKey) tryAdd(p, { benchIdx: idx }); });
  // 2) 他の守備位置に配置済みのスタメン
  TB_DIAMOND_POSITIONS.forEach(op => { if (op !== pos && TB_STATE.batters[op] && playerKey(TB_STATE.batters[op]) !== curKey) tryAdd(TB_STATE.batters[op], { fromPos: op }); });
  // 3) フリー(未登録)。tbEligibleBatters は控え/他スタメン/投手を除外済み・同名重複除外・現選手保持。
  tbEligibleBatters(pos, curKey).forEach(p => tryAdd(p, {}));
  // 4) 現選手(この枠)は守備可否に関わらず必ず残す
  if (curKey) tryAdd(TB_STATE.batters[pos], {});
  out.forEach(c => { c.drs = tbDrsValue(c.p, pos); });
  out.sort((a, b) => (tbRecScore(b.p, pos) - tbRecScore(a.p, pos)) || ((b.drs || 0) - (a.drs || 0)));
  return out;
}

// 現在のオーダーで、各守備位置に「レギュラー登録可能な選手数」を集計する。
//   ・対象: チーム/年度フィルタ後の打者プール。同名(同種別)は最新年度1名に間引く。
//   ・登録制約: 現オーダーに出場できない選手(playerAllowedInOrder=false)はカウントしない。
//   ・複数ポジション守れる選手は各ポジションで重複してカウント。
//   ・全 = DHに登録可能な選手数(=登録可能な合計人数)。DHは誰でも可。
function tbPositionCounts() {
  const oi = (TB_STATE && TB_STATE.currentOrder) || 0;
  // ★順序が重要: 先にオーダー登録可で絞り、その後で同名(同種別)を最新年度1名に間引く。
  //   先に間引くと「最新年度版はオーダー不可だが旧年度版なら可」の選手を取りこぼし、
  //   守備プルダウン候補(tbEligibleBatters: 絞り→間引き) と人数がズレる原因になる。
  const pool = tbDedupePoolNewest(tbBatterPool().filter(p => playerAllowedInOrder(p, oi) && !isMinorPlayer(p)));
  const POS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF'];
  const counts = {};
  POS.forEach(pos => { counts[pos] = pool.filter(p => canPlay(p, pos)).length; });
  counts.all = pool.length;   // 全 = DH可 = 登録可能合計
  return counts;
}

// TB_STATE の不整合を自動修復する。
// 「選手がいないのに打順番号だけ残っている」守備位置は描画クラッシュの原因になるため、
// その打順番号を消して整合させる (孤立した打順の除去)。
function sanitizeTbState() {
  if (!TB_STATE) return;
  for (const pos of Object.keys(TB_STATE.batters)) {
    if (!TB_STATE.batters[pos] && TB_STATE.batterOrder[pos] != null) {
      TB_STATE.batterOrder[pos] = null;
    }
  }
}

function renderTeamBuild() {
  if (!TB_STATE) resetTeamBuild();
  sanitizeTbState();
  renderTbDiamond();
  renderTbLineupPanel();
  renderTbPitcherRow();
}

// ダイヤモンドビュー (9守備位置のセレクタ)
function renderTbDiamond() {
  const root = $('#tb-diamond');
  if (!root) return;
  let html = `<svg class="tb-field-svg" viewBox="0 0 200 180" preserveAspectRatio="xMidYMid meet" aria-hidden="true">
    <defs>
      <pattern id="tbGrass" x="0" y="0" width="14" height="14" patternUnits="userSpaceOnUse">
        <rect width="14" height="14" fill="#4ea84a"/>
        <rect x="0" width="7" height="14" fill="#3f9a3b"/>
      </pattern>
    </defs>
    <path d="M 100 170 L 10 95 Q 10 12 100 6 Q 190 12 190 95 Z" fill="url(#tbGrass)" stroke="#2b6f29" stroke-width="1"/>
    <path d="M 100 148 L 60 100 L 100 56 L 140 100 Z" fill="#caa177" stroke="#8a6a45" stroke-width="0.6"/>
    <circle cx="100" cy="100" r="22" fill="#5cb85c"/>
    <circle cx="100" cy="100" r="7" fill="#caa177" stroke="#8a6a45" stroke-width="0.5"/>
    <rect x="135" y="94" width="12" height="12" fill="#fff" stroke="#333" stroke-width="1" transform="rotate(45 141 100)"/>
    <rect x="94"  y="54" width="12" height="12" fill="#fff" stroke="#333" stroke-width="1" transform="rotate(45 100 60)"/>
    <rect x="53"  y="94" width="12" height="12" fill="#fff" stroke="#333" stroke-width="1" transform="rotate(45 59 100)"/>
    <polygon points="92,142 108,142 108,148 100,154 92,148" fill="#fff" stroke="#333" stroke-width="1"/>
  </svg>`;
  for (const pos of TB_DIAMOND_POSITIONS) {
    const layout = TB_POS_LAYOUT[pos];
    const cur = TB_STATE.batters[pos];
    const orderNum = TB_STATE.batterOrder[pos];
    const drs = cur ? tbDrsValue(cur, pos) : null;
    const curKey = cur ? playerKey(cur) : null;
    const eligible = tbDiamondCandidates(pos, curKey);
    // 打順インライン select (1-9)
    const orderOpts = ['<option value="">-</option>']
      .concat([1,2,3,4,5,6,7,8,9].map(n => `<option value="${n}"${n===orderNum?' selected':''}>${n}</option>`)).join('');
    const orderSel = `<select class="tb-order-sel ${orderNum?'':'empty'}" data-pos="${pos}" title="打順">${orderOpts}</select>`;
    const drsBox = (drs !== null) ? `<span class="tb-drs" title="DRS">${drs}</span>` : '<span class="tb-drs empty">-</span>';
    const ovrBox = cur ? ovrBadge(cur) : '';
    // プレイヤー名 / 追加ボタン (クリックで select 表示)
    let nameHtml;
    if (cur) {
      nameHtml = `<div class="tb-name-row" data-pos="${pos}">
        <span class="tb-name player-link${longNameClass(cur.fullNameTop)}" data-player-name="${cur.fullNameTop}" data-player-year="${cur.year||''}" data-player-type="${playerType(cur)}" data-player-team="${cur.team||''}">${cur.fullNameTop}</span>
        <button type="button" class="tb-change-btn" data-pos="${pos}" title="選手変更">▼</button>
      </div>`;
    } else {
      nameHtml = `<button type="button" class="tb-add-btn" data-pos="${pos}">＋ 追加</button>`;
    }
    // 隠し select (クリック時にだけ display: block)。候補はAI推奨順。ホバー(title)でDRSを表示。
    const POS_SHORT = { C: '捕', '1B': '一', '2B': '二', '3B': '三', SS: '遊', LF: '左', CF: '中', RF: '右', DH: 'DH' };
    const playerOpts = eligible.map((c, i) => {
      const p = c.p;
      const sel = (cur && p === cur) ? ' selected' : '';
      const tag = (c.benchIdx != null) ? '（控え）' : (c.fromPos != null ? `（${POS_SHORT[c.fromPos] || c.fromPos}）` : '');
      const drsTxt = Number.isFinite(c.drs) ? `DRS ${c.drs >= 0 ? '+' : ''}${c.drs}` : 'DRS なし';
      return `<option value="${i}" title="${drsTxt}"${sel}>${p.fullNameTop}${p.year||''}${tag}</option>`;
    }).join('');
    const clearOpt = cur ? '<option value="__clear__">— 外す（＋追加に戻す）—</option>' : '';
    const hiddenSel = `<select class="tb-select tb-select-hidden" data-pos="${pos}" style="display:none">
      <option value="">(キャンセル)</option>${clearOpt}${playerOpts}
    </select>`;
    // アンカー別の配置スタイル
    let posStyle;
    if (layout.anchor === 'topLeft') {
      posStyle = `left:${layout.x}%; top:${layout.y}%; transform:none;`;
    } else if (layout.anchor === 'topRight') {
      posStyle = `right:${(100 - layout.x).toFixed(2)}%; top:${layout.y}%; transform:none;`;
    } else {
      // 中心アンカー (デフォルト)
      posStyle = `left:${layout.x}%; top:${layout.y}%;`;
    }
    html += `<div class="tb-slot ${layout.external ? 'ext' : ''}" style="${posStyle}" data-pos="${pos}">
      <div class="tb-slot-head">
        ${orderSel}${drsBox}${ovrBox}
        <span class="tb-pos-label">${layout.label}</span>
      </div>
      ${nameHtml}
      ${hiddenSel}
    </div>`;
  }
  // ダイヤモンド左上: チーム総合力 (2cm大)
  const teamOvr = tbTeamOverall();
  html += `<div class="tb-team-ovr" title="チーム総合力">
    <span class="tto-label">チーム総合力</span>
    <span class="tto-value">${teamOvr != null ? teamOvr : '—'}</span>
  </div>`;
  // ダイヤモンド右上: オーダー切り替え + 年度選択
  const curOrd = TB_STATE.currentOrder || 0;
  let ordBtns = '';
  for (let i = 0; i < TB_ORDER_COUNT; i++) {
    ordBtns += `<button type="button" class="tb-order-btn${i === curOrd ? ' active' : ''}" data-order="${i}">${i + 1}</button>`;
  }
  // 年度プルダウン: データ中の年度 (降順) + 現在選択年度を必ず候補に含める
  const yearSet = new Set([...getBatters(), ...getPitchers()].map(p => p.year).filter(y => Number.isFinite(y)));
  if (Number.isFinite(TB_STATE.year)) yearSet.add(TB_STATE.year);
  const allYears = [...yearSet].sort((a, b) => b - a);
  const curYear = TB_STATE.year;
  const yearOpts = ['<option value="">指定なし</option>']
    .concat(allYears.map(y => `<option value="${y}"${curYear === y ? ' selected' : ''}>${y}年</option>`)).join('');
  html += `<div class="tb-topright">
    <div class="tb-order-switch" title="打順オーダーの切り替え (野手のみ)">
      <span class="tos-label">オーダー</span>${ordBtns}
    </div>
    <div class="tb-year-row" title="年度フィルタ (この年度のチーム所属選手のみ候補に表示。総合力65以下は2年前まで可)">
      <span class="tos-label">年度</span>
      <select class="tb-year-select">${yearOpts}</select>
    </div>
  </div>`;
  // ダイヤモンド左下 (三塁枠の下・凡例の上): 現在オーダーの戦力4カテゴリ評価
  const ratings = tbTeamRatings();
  const rRows = [['投手力', ratings.pitching], ['攻撃力', ratings.offense], ['守備力', ratings.defense], ['機動力', ratings.mobility]]
    .map(([lbl, v]) => { const disp = Math.round(v); const g = tbGradeOf(disp); return `<div class="tbr-row"><span class="tbr-label">${lbl}</span><span class="tbr-score">${disp}</span><span class="tbr-grade g-${g.toLowerCase()}">${g}</span></div>`; })
    .join('');
  html += `<div class="tb-team-ratings" title="現在のオーダーの戦力評価 (S:90〜100 A:80〜89 B:70〜79 C:60〜69 D:50〜59 E:40〜49 F:30〜39 G:〜29。100超SSはオリジナル/年代不問のオールスターのみ)">${rRows}</div>`;
  // ダイヤモンド下中央 (捕手枠の下・凡例の上): 全changeボタン
  html += `<button type="button" class="btn tb-allchange" title="オーダー1〜3の野手(控え込み)と投手陣を、年度設定に従ってすべて自動編成し直します">🔄 全change</button>`;
  // ここまでが守備フィールド本体 (SVG + 9枠 + 総合力 + オーダー/年度 + 戦力評価 + 全change)。
  //   凡例と人数はフィールドの「下」(枠外) に置き、DH枠などと重ならないようにする。
  html = `<div class="tb-field">${html}</div>`;
  // 打順/DRS 凡例 + このオーダーの守備位置別 登録可能人数
  const pc = tbPositionCounts();
  const PC_JP = { C: '捕', '1B': '一', '2B': '二', '3B': '三', SS: '遊', LF: '左', CF: '中', RF: '右' };
  const pcStr = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF'].map(p => `${PC_JP[p]}/${pc[p]}`).join('　') + `　全/${pc.all}`;
  html += `<div class="tb-diamond-legend">
    <div class="tb-legend-row">
      <span><span class="tb-legend-box red">N</span>打順</span>
      <span><span class="tb-legend-box yellow">N</span>DRS</span>
    </div>
    <div class="tb-poscount" title="このオーダーで各守備位置に登録可能な選手数（同名は最新年度1名／登録不可は除外、全=DH可=登録合計）">${pcStr}</div>
  </div>`;
  root.innerHTML = html;
  // オーダー切り替えボタン
  root.querySelectorAll('.tb-order-btn').forEach(btn => {
    btn.addEventListener('click', () => tbSwitchOrder(parseInt(btn.dataset.order, 10)));
  });
  // 全change: オーダー1〜3の野手・控え + 投手陣を年度設定に従い全自動編成
  const allChg = root.querySelector('.tb-allchange');
  if (allChg) allChg.addEventListener('click', () => {
    if (!confirm('オーダー1・2・3の野手(控え込み)と投手陣を、すべて自動編成で上書きします。\nよろしいですか？')) return;
    tbAutoFillAllOrders();
  });
  // 年度選択
  const yearSelEl = root.querySelector('.tb-year-select');
  if (yearSelEl) yearSelEl.addEventListener('change', () => tbSetYear(yearSelEl.value));
  // 打順 select change → 同番号の他選手と入れ替え
  root.querySelectorAll('.tb-order-sel').forEach(sel => {
    sel.addEventListener('change', () => {
      const pos = sel.dataset.pos;
      const v = sel.value;
      if (v === '') {
        TB_STATE.batterOrder[pos] = null;
      } else {
        const n = parseInt(v);
        // 同じ番号の他選手があれば入れ替え
        for (const otherPos of TB_DIAMOND_POSITIONS) {
          if (otherPos !== pos && TB_STATE.batterOrder[otherPos] === n) {
            TB_STATE.batterOrder[otherPos] = TB_STATE.batterOrder[pos] || null;
            break;
          }
        }
        TB_STATE.batterOrder[pos] = n;
      }
      renderTeamBuild();
    });
  });
  // 選手追加ボタン / 変更ボタンクリック → 隠し select を表示&フォーカス
  const openSelect = (pos) => {
    const slot = root.querySelector(`.tb-slot[data-pos="${pos}"]`);
    if (!slot) return;
    const sel = slot.querySelector('.tb-select-hidden');
    if (!sel) return;
    slot.classList.add('tb-slot-open');  // z-index 引き上げ
    sel.style.display = '';
    sel.focus();
    sel.size = Math.min(8, sel.options.length);
  };
  root.querySelectorAll('.tb-add-btn, .tb-change-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openSelect(btn.dataset.pos);
    });
  });
  root.querySelectorAll('.tb-select-hidden').forEach(sel => {
    const closeSelect = () => {
      sel.style.display = 'none';
      sel.removeAttribute('size');
      const slot = sel.closest('.tb-slot');
      if (slot) slot.classList.remove('tb-slot-open');
    };
    sel.addEventListener('change', () => {
      const pos = sel.dataset.pos;
      const v = sel.value;
      if (v === '') { closeSelect(); return; }
      if (v === '__clear__') {  // 外す → 空きへ戻す
        TB_STATE.batters[pos] = null;
        TB_STATE.batterOrder[pos] = null;
        renderTeamBuild();
        return;
      }
      const curKey = TB_STATE.batters[pos] ? playerKey(TB_STATE.batters[pos]) : null;
      const eligible = tbDiamondCandidates(pos, curKey);
      const chosen = eligible[+v];
      if (!chosen) { closeSelect(); return; }
      const curP = TB_STATE.batters[pos] || null;
      if (chosen.benchIdx != null) {
        // 控え選手をスタメンへ。元のスタメン選手は その控え枠へ移動 (入替)。
        //   空き枠だった場合は控え枠を空にする (null)。打順番号は守備位置側に保持。
        const benchP = TB_STATE.pinchHitters[chosen.benchIdx];
        TB_STATE.pinchHitters[chosen.benchIdx] = curP;
        TB_STATE.batters[pos] = benchP;
      } else if (chosen.fromPos != null) {
        // 他守備位置のスタメンを選択 → この位置へ移す。元の選手はその空いた位置へ:
        //   ポジションチェンジ可(その位置を守れる) → 入替 / 不可 → 元の位置は空にする。
        const posS = chosen.fromPos;
        TB_STATE.batters[pos] = chosen.p;
        TB_STATE.batters[posS] = (curP && canPlay(curP, posS)) ? curP : null;
        // 打順番号は守備位置側に保持 (空きになった位置の打順は sanitizeTbState が整理)
      } else {
        // フリー(未登録)の選手を配置 (従来挙動)
        TB_STATE.batters[pos] = chosen.p;
      }
      if (!TB_STATE.batterOrder[pos]) {
        const used = new Set(Object.values(TB_STATE.batterOrder).filter(n => n));
        for (let n = 1; n <= 9; n++) {
          if (!used.has(n)) { TB_STATE.batterOrder[pos] = n; break; }
        }
      }
      renderTeamBuild();
    });
    sel.addEventListener('blur', closeSelect);
  });
}

// 左ペイン: 打順 + PH リスト (ドラッグ&ドロップで順番入替)
function renderTbLineupPanel() {
  const list = $('#tb-lineup-list');
  if (!list) return;
  const rows = [];
  for (let n = 1; n <= 9; n++) {
    let foundPos = null;
    for (const pos of TB_DIAMOND_POSITIONS) {
      if (TB_STATE.batterOrder[pos] === n) { foundPos = pos; break; }
    }
    const p = foundPos ? TB_STATE.batters[foundPos] : null;
    if (foundPos && p) {
      rows.push(`<li draggable="true" data-order="${n}" data-pos="${foundPos}">
        <span class="tbl-handle">⠿</span>
        <span class="tbl-num">${n}</span>
        <span class="tbl-pos">${foundPos}</span>
        <span class="tbl-name player-link${longNameClass(p.fullNameTop)}" data-player-name="${p.fullNameTop}" data-player-year="${p.year||''}" data-player-type="${playerType(p)}" data-player-team="${p.team||''}">${p.fullNameTop}</span>
        ${ovrBadge(p)}
      </li>`);
    } else {
      rows.push(`<li draggable="false" data-order="${n}" class="empty-slot">
        <span class="tbl-handle">⠿</span>
        <span class="tbl-num">${n}</span>
        <span class="tbl-pos">-</span>
        <span class="tbl-name empty">未割当</span>
      </li>`);
    }
  }
  list.innerHTML = rows.join('');
  // D&D 設定
  enableLineupDnd(list);

  // PH リスト
  const phList = $('#tb-ph-list');
  if (!phList) return;
  const pool = tbBatterPool();
  const phRows = [];
  for (let i = 0; i < TB_PH_COUNT; i++) {
    const cur = TB_STATE.pinchHitters[i];
    const curKey = cur ? playerKey(cur) : null;
    const eligible = tbFilterCandidates(pool, curKey);
    let nameHtml;
    if (cur) {
      nameHtml = `<span class="tbl-name player-link${longNameClass(cur.fullNameTop)}" data-player-name="${cur.fullNameTop}" data-player-year="${cur.year||''}" data-player-type="${playerType(cur)}" data-player-team="${cur.team||''}">${cur.fullNameTop}</span>
        ${ovrBadge(cur)}
        <button type="button" class="tbl-change-btn" data-ph="${i}">▼</button>`;
    } else {
      nameHtml = `<button type="button" class="tbl-add-btn" data-ph="${i}">＋ 追加</button>`;
    }
    const opts = eligible.map((p, idx) => {
      const sel = cur && p === cur ? ' selected' : '';
      return `<option value="${idx}"${sel}>${p.fullNameTop}${p.year||''}</option>`;
    }).join('');
    const clearOpt = cur ? '<option value="__clear__">— 外す（＋追加に戻す）—</option>' : '';
    const hiddenSel = `<select class="tb-ph-select tb-select-hidden" data-idx="${i}" style="display:none">
      <option value="">(キャンセル)</option>${clearOpt}${opts}
    </select>`;
    phRows.push(`<li><span class="tbl-num">${TB_PH_LABELS[i] || ('控'+(i+1))}</span>${nameHtml}${hiddenSel}</li>`);
  }
  phList.innerHTML = phRows.join('');
  // PH 追加/変更ボタン
  phList.querySelectorAll('.tbl-add-btn, .tbl-change-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const li = btn.closest('li');
      const sel = li.querySelector('.tb-ph-select');
      sel.style.display = '';
      sel.size = Math.min(8, sel.options.length);
      sel.focus();
    });
  });
  // PH select の動作
  phList.querySelectorAll('.tb-ph-select').forEach(sel => {
    sel.addEventListener('change', () => {
      const i = parseInt(sel.dataset.idx);
      const v = sel.value;
      if (v === '') {
        sel.style.display = 'none';
        sel.removeAttribute('size');
        return;
      }
      if (v === '__clear__') {  // 外す → 空きへ戻す
        TB_STATE.pinchHitters[i] = null;
        renderTeamBuild();
        return;
      }
      const cur = TB_STATE.pinchHitters[i];
      const curKey = cur ? playerKey(cur) : null;
      const eligible = tbFilterCandidates(tbBatterPool(), curKey);
      TB_STATE.pinchHitters[i] = eligible[+v];
      renderTeamBuild();
    });
    sel.addEventListener('blur', () => {
      sel.style.display = 'none';
      sel.removeAttribute('size');
    });
  });
}

// 打順リストの D&D 設定 (順番入替え)
function enableLineupDnd(list) {
  const lis = list.querySelectorAll('li');
  lis.forEach(li => {
    li.addEventListener('dragstart', e => {
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', li.dataset.order);
      li.classList.add('dragging');
    });
    li.addEventListener('dragend', () => {
      li.classList.remove('dragging');
      list.querySelectorAll('li').forEach(x => x.classList.remove('drag-over'));
    });
    li.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
    li.addEventListener('dragenter', e => { e.preventDefault(); if (!li.classList.contains('dragging')) li.classList.add('drag-over'); });
    li.addEventListener('dragleave', () => li.classList.remove('drag-over'));
    li.addEventListener('drop', e => {
      e.preventDefault();
      li.classList.remove('drag-over');
      const srcOrder = parseInt(e.dataTransfer.getData('text/plain'));
      const dstOrder = parseInt(li.dataset.order);
      if (!srcOrder || !dstOrder || srcOrder === dstOrder) return;
      // batterOrder の値を入れ替え (該当ポジションを探して値を swap)
      let srcPos = null, dstPos = null;
      for (const pos of TB_DIAMOND_POSITIONS) {
        if (TB_STATE.batterOrder[pos] === srcOrder) srcPos = pos;
        if (TB_STATE.batterOrder[pos] === dstOrder) dstPos = pos;
      }
      if (srcPos) TB_STATE.batterOrder[srcPos] = dstOrder;
      if (dstPos) TB_STATE.batterOrder[dstPos] = srcOrder;
      // 片方しかいない場合 (dstが空) は src の打順だけ移動
      if (srcPos && !dstPos) TB_STATE.batterOrder[srcPos] = dstOrder;
      renderTeamBuild();
    });
  });
}

// 下ペイン: 投手陣 (先発5/モップ2/中継4/SU2/抑1/控4) — D&D 対応
function renderTbPitcherRow() {
  const root = $('#tb-pitcher-row');
  if (!root) return;
  const pool = tbPitcherPool();
  const html = TB_PITCHER_SLOTS.map(group => {
    const slots = TB_STATE.pitchers[group.role];
    // 役割別フィルタ: 先発は スタミナ≥50・回復量7超 (リリーフ用低スタミナ投手を除外)
    const rolePool = (group.role === 'starter')
      ? pool.filter(p => (p.stats?.['スタミナ'] || 0) >= 50 && getRecoveryOf(p) > 7)
      : pool;
    const slotHtml = slots.map((cur, idx) => {
      const curKey = cur ? playerKey(cur) : null;
      // 重複(完全一致)除外 + 年代違い同名の最新年度のみ + 現選手は常に残す
      const eligible = tbFilterCandidates(rolePool, curKey);
      let nameHtml;
      if (cur) {
        nameHtml = `<span class="tbp-name player-link${longNameClass(cur.fullNameTop)}" data-player-name="${cur.fullNameTop}" data-player-year="${cur.year||''}" data-player-type="${playerType(cur)}" data-player-team="${cur.team||''}">${cur.fullNameTop}</span>
          ${ovrBadge(cur)}
          <button type="button" class="tbp-change-btn" data-role="${group.role}" data-idx="${idx}">▼</button>`;
      } else {
        nameHtml = `<button type="button" class="tbp-add-btn" data-role="${group.role}" data-idx="${idx}">＋ 追加</button>`;
      }
      const opts = eligible.map((p, i) => {
        const sel = cur && p === cur ? ' selected' : '';
        return `<option value="${i}"${sel}>${p.fullNameTop}${p.year||''}</option>`;
      }).join('');
      const clearOpt = cur ? '<option value="__clear__">— 外す（＋追加に戻す）—</option>' : '';
      const hiddenSel = `<select class="tbp-select tb-select-hidden" data-role="${group.role}" data-idx="${idx}" style="display:none">
        <option value="">(キャンセル)</option>${clearOpt}${opts}
      </select>`;
      return `<div class="tbp-slot" draggable="${cur?'true':'false'}" data-role="${group.role}" data-idx="${idx}">
        <span class="tbp-handle">⠿</span>
        ${nameHtml}
        ${hiddenSel}
      </div>`;
    }).join('');
    return `<div class="tbp-group">
      <div class="tbp-group-title">${group.label} <span class="tbp-count">${group.count}</span></div>
      ${slotHtml}
    </div>`;
  }).join('');
  root.innerHTML = html;
  // 追加/変更ボタン → 隠し select を開く
  root.querySelectorAll('.tbp-add-btn, .tbp-change-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const slot = btn.closest('.tbp-slot');
      const sel = slot.querySelector('.tbp-select');
      slot.classList.add('tbp-slot-open');
      sel.style.display = '';
      sel.size = Math.min(8, sel.options.length);
      sel.focus();
    });
  });
  // セレクタ change
  root.querySelectorAll('.tbp-select').forEach(sel => {
    const closeSelect = () => {
      sel.style.display = 'none';
      sel.removeAttribute('size');
      const slot = sel.closest('.tbp-slot');
      if (slot) slot.classList.remove('tbp-slot-open');
    };
    sel.addEventListener('change', () => {
      const role = sel.dataset.role;
      const idx  = parseInt(sel.dataset.idx);
      const v = sel.value;
      if (v === '') { closeSelect(); return; }
      if (v === '__clear__') {  // 外す → 空きへ戻す
        TB_STATE.pitchers[role][idx] = null;
        renderTeamBuild();
        return;
      }
      const cur = TB_STATE.pitchers[role][idx];
      const curKey = cur ? playerKey(cur) : null;
      // 先発の rolePool は描画側と同条件 (スタミナ50以上・回復量7超) にする (候補indexを一致させるため)
      const rolePool = (role === 'starter')
        ? tbPitcherPool().filter(p => (p.stats?.['スタミナ']||0) >= 50 && getRecoveryOf(p) > 7)
        : tbPitcherPool();
      const eligible = tbFilterCandidates(rolePool, curKey);
      TB_STATE.pitchers[role][idx] = eligible[+v];
      renderTeamBuild();
    });
    sel.addEventListener('blur', closeSelect);
  });
  // D&D: スロット間の選手入替え (役割を跨いでもOK)
  enablePitcherDnd(root);
}

function enablePitcherDnd(root) {
  const slots = root.querySelectorAll('.tbp-slot');
  slots.forEach(slot => {
    slot.addEventListener('dragstart', e => {
      if (slot.getAttribute('draggable') !== 'true') return;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', JSON.stringify({ role: slot.dataset.role, idx: slot.dataset.idx }));
      slot.classList.add('dragging');
    });
    slot.addEventListener('dragend', () => {
      slot.classList.remove('dragging');
      root.querySelectorAll('.tbp-slot').forEach(x => x.classList.remove('drag-over'));
    });
    slot.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
    slot.addEventListener('dragenter', e => { e.preventDefault(); if (!slot.classList.contains('dragging')) slot.classList.add('drag-over'); });
    slot.addEventListener('dragleave', () => slot.classList.remove('drag-over'));
    slot.addEventListener('drop', e => {
      e.preventDefault();
      slot.classList.remove('drag-over');
      let data;
      try { data = JSON.parse(e.dataTransfer.getData('text/plain') || '{}'); }
      catch { return; }
      if (!data.role || data.idx === undefined) return;
      const srcRole = data.role, srcIdx = parseInt(data.idx);
      const dstRole = slot.dataset.role, dstIdx = parseInt(slot.dataset.idx);
      if (srcRole === dstRole && srcIdx === dstIdx) return;
      const tmp = TB_STATE.pitchers[srcRole][srcIdx];
      TB_STATE.pitchers[srcRole][srcIdx] = TB_STATE.pitchers[dstRole][dstIdx];
      TB_STATE.pitchers[dstRole][dstIdx] = tmp;
      renderTeamBuild();
    });
  });
}

// 投手の役割適性スコア
function getStaminaOf(p)     { return p.stats?.['スタミナ'] ?? 50; }
function getRecoveryOf(p)    { const r = parseFloat(p.stats?.['回復量'] ?? p.statsMini?.['回復量'] ?? 0); return Number.isFinite(r) ? r : 0; }
function getInningsOf(p)     { const i = parseFloat(p.record?.['イニング'] ?? 0); return Number.isFinite(i) ? i : 0; }
function getSeasonSavesOf(p) { const s = parseFloat(p.record?.['セーブ'] || '0'); return Number.isFinite(s) ? s : 0; }
function getSeasonERAOf(p)   { const e = parseFloat(p.record?.['防御率'] || '99'); return Number.isFinite(e) ? e : 99; }
function getSeasonHoldsOf(p) {
  // カードのrawHtml内、当該年度のHLDを正規表現で抽出 (なければ0)
  if (!p.rawHtml || !p.year) return 0;
  const re = new RegExp('data-statkey="hld"\\s+data-year="' + p.year + '"\\s+data-val="(\\d+)"');
  const m = p.rawHtml.match(re);
  return m ? parseInt(m[1]) : 0;
}
// 選手カードの「年度試合出場数」。取得不能なら null (休養ロジック対象外)。
//   オーダー制約用に実装済みの getGamesPlayedOf を流用する。
//   (getGamesPlayedOf は gamesPlayed / record['試合'] / stats['試合'] → rawHtml の
//    年度別成績テーブル(class="cb-year" 行の「試合」列) の順で取得し、結果をキャッシュする)
function getCardGamesOf(p) {
  if (!p) return null;
  const g = getGamesPlayedOf(p);
  return (g != null && g > 0) ? g : null;
}
function rankAsStarter(p) {
  // スタミナ50以上を強く優先、加えてERA低いほどボーナス
  const sta = getStaminaOf(p);
  const era = getSeasonERAOf(p);
  let score = sta;
  if (sta < 50) score -= 200;        // 50未満は大幅減点
  score += Math.max(0, (5.0 - era) * 5);  // ERA低いほどボーナス
  return score;
}
// 打者の打順別適性スコア (1〜9番)
function scoreForBattingOrder(b, orderNum) {
  const s = b.stats || {};
  const r = b.record || {};
  const m = b.statsMini || {};
  const pf = (v, d) => { const x = parseFloat(v); return Number.isFinite(x) ? x : d; };
  const obp  = pf(r['出塁率'], 0.300);
  const avg  = pf(r['打率'],   0.250);
  const ops  = pf(r['OPS'],    0.700);
  const hr   = pf(r['本塁打'], 0);
  const sb   = pf(r['盗塁'],   0);
  const meet  = s['ミート']    || 60;
  const power = s['パワー']    || 60;
  const speed = s['スピード']  || 60;
  const chance= s['チャンス']  || 60;
  const eye   = s['選球眼']    || 60;
  const kTol  = s['三振耐性']  || 60;
  const hrAb  = m['HR能']      || 0;
  const sbAb  = m['盗塁能']    || 0;
  switch (orderNum) {
    case 1: // リードオフ: 出塁率高 + 足速 + 選球眼
      return obp * 200 + speed * 0.6 + eye * 0.5 + meet * 0.2 + sb * 0.3 + sbAb * 2;
    case 2: // 繋ぎ: 出塁率高 + ミート + 選球眼
      return obp * 180 + meet * 0.6 + eye * 0.4 + speed * 0.2 + kTol * 0.2;
    case 3: // 好打者: 打率 + ミート + OPS
      return avg * 200 + ops * 50 + meet * 0.6 + power * 0.3 + chance * 0.2;
    case 4: // 主砲: パワー + HR能 + チャンス + OPS
      return power * 0.9 + hrAb * 8 + chance * 0.6 + ops * 60 + hr * 0.4;
    case 5: // 主砲補佐: パワー + チャンス + OPS
      return power * 0.7 + hrAb * 5 + chance * 0.5 + ops * 50 + hr * 0.3;
    case 6: // 中位パワー
      return power * 0.5 + chance * 0.3 + ops * 30 + meet * 0.2;
    case 7: // 平均
      return meet * 0.4 + ops * 25 + chance * 0.2;
    case 8: // 下位 (ミート控えめ + 出塁率)
      return meet * 0.3 + obp * 60 + kTol * 0.2;
    case 9: // 守備重視 (打撃低めを許容、ピッチャー枠的)
      return ops * 10 + meet * 0.2;
  }
  return 0;
}

function rankAsCloser(p) {
  // スタミナ低い + セーブ実績 or (ホールド + 低ERA) を優先
  const sta = getStaminaOf(p);
  const sv  = getSeasonSavesOf(p);
  const hld = getSeasonHoldsOf(p);
  const era = getSeasonERAOf(p);
  let score = 0;
  // セーブ実績で大ボーナス
  score += sv * 10;
  // ホールド実績で中ボーナス + 低ERAでさらに加算
  if (hld > 0) score += hld * 3 + Math.max(0, (4.0 - era) * 8);
  // スタミナが低いほどボーナス (短いイニング向き)
  score += Math.max(0, (60 - sta));
  // ERA単体でも低ければ少しボーナス
  score += Math.max(0, (4.5 - era) * 4);
  return score;
}

function autoFill() {
  for (const side of ['away','home']) {
    const team = getTeamFilter(side);
    const tbSave = getSavedTeamBuild(team, getOrderFilter(side));
    // === 保存済みチームがある場合: 先発1〜5 からランダム1人を slot 0 にセット、それ以外は維持 ===
    if (tbSave) {
      const starters = (tbSave.pitchers.starter || []).filter(Boolean);
      if (starters.length > 0) {
        const picked = starters[randI(starters.length)];
        const slot0 = $('.sel-pitcher-slot[data-side="'+side+'"][data-idx="0"]');
        if (slot0) {
          const idx = Array.from(slot0.options).findIndex(o =>
            o.textContent === labelOf(picked));
          slot0.value = idx > 0 ? String(idx - 1) : '0';  // skip placeholder
          // 上の findIndex は placeholder 込み index なので value(starters内 index) を直接設定
          slot0.value = String(starters.indexOf(picked));
        }
      }
      // 打順 (lineup) は buildSetupSide で pre-fill 済み — 触らない
      // 中継ぎ/抑えも pre-fill 済み
      clearDuplicateSelections(side);
      refreshSelectionsForSide(side);
      continue;
    }
    // === 保存無し: 従来通りランダム自動編成 ===
    const batters  = applyTeamFilter(getBatters(), team);
    const pitchers = applyTeamFilter(getPitchers(), team);
    // 投手の役割割り当て:
    //   index 0=先発, 1=中継ぎ1, 2=中継ぎ2, 3=抑え
    const pickByScore = (pool, scorer) => {
      const scored = pool.map(p => ({ p, s: scorer(p) + rand() * 8 }));
      scored.sort((a, b) => b.s - a.s);
      return scored[0].p;
    };
    // スタミナ50以上の候補から、上位7人を重み付き抽選で選ぶ (ランダム性強め)
    const pickStarter = (pool) => {
      const eligible = pool.filter(p => getStaminaOf(p) >= 50);
      const fallback = eligible.length > 0 ? eligible : pool;
      const scored = fallback.map(p => ({ p, s: rankAsStarter(p) }))
                             .sort((a, b) => b.s - a.s);
      // 上位7人まで (プール大きければより多様)
      const topN = scored.slice(0, Math.min(7, scored.length));
      // 重みを 0.8〜1.2 に強く圧縮 (上位は約1.5倍止まりで、ほぼフラット選出)
      const minS = topN[topN.length - 1].s;
      const maxS = topN[0].s;
      const range = Math.max(1, maxS - minS);
      const weights = topN.map(x => 0.8 + (x.s - minS) / range * 0.4);
      const total = weights.reduce((a, b) => a + b, 0);
      let r = rand() * total;
      for (let i = 0; i < topN.length; i++) {
        r -= weights[i];
        if (r <= 0) return topN[i].p;
      }
      return topN[topN.length - 1].p;
    };
    const pool = [...pitchers];
    // 1) 先発: スタミナ50以上から重み付き抽選 (ランダム性あり)
    const starter = pickStarter(pool);
    // 先発スロット (1) にセット
    const starterSel = $('.sel-pitcher-slot[data-side="'+side+'"][data-idx="0"]');
    if (starterSel) {
      const si = pitchers.indexOf(starter);
      starterSel.value = si >= 0 ? String(si) : '';
    }
    // 2) リリーフ枠: 先発以外をシャッフルして 中継/SU/抑え/モップ スロットに順次充填
    const reliefCandidates = pool.filter(p => p !== starter);
    for (let i = reliefCandidates.length - 1; i > 0; i--) {
      const j = randI(i + 1);
      [reliefCandidates[i], reliefCandidates[j]] = [reliefCandidates[j], reliefCandidates[i]];
    }
    const reliefSels = $$('.sel-relief-slot[data-side="'+side+'"]');
    reliefSels.forEach((sel, i) => {
      const cand = reliefCandidates[i];
      // reliefPool は no-save 時 pitchers と同一なので index を流用
      sel.value = cand ? String(pitchers.indexOf(cand)) : '';
    });
    // 打者: MLBセオリーに沿った打順 (1-2番OBP+足, 3-5番パワー, 9番守備重視)
    //       かつ各枠の守備ポジションを守れる選手のみに限定
    const slotsLi = $$('.batter-slots[data-side="'+side+'"] li');
    // 各スロットの情報を集約
    const slotInfo = Array.from(slotsLi).map((li, idx) => ({
      li, idx, orderNum: idx + 1,
      posSel: li.querySelector('.sel-pos'),
      batSel: li.querySelector('.sel-batter'),
    }));
    // 守れる選手が0のポジションでも自動でDHには変更しない
    // (defaultの 9 ポジション C/1B/2B/3B/SS/LF/CF/RF/DH を維持)
    // チーム選択で該当選手がいない場合は readSetup でエラー表示
    // スロット×打者 のスコア行列
    const matrix = slotInfo.map(si => {
      const pos = si.posSel.value;
      return batters.map(b => {
        if (!canPlay(b, pos)) return -Infinity;
        return scoreForBattingOrder(b, si.orderNum) + rand() * 3;
      });
    });
    // 候補が少ないスロットから先に割り当てる貪欲法
    const candidateCount = slotInfo.map(si =>
      filterBattersByPos(batters, si.posSel.value).length
    );
    const slotOrder = slotInfo.map((_, i) => i)
      .sort((a, b) => candidateCount[a] - candidateCount[b]);
    const usedBatters = new Set();
    const lineupAssigned = new Array(9).fill(null);
    // 重複なしで埋める。プールが足りなければ未割当のまま (-- 選択 --) にする
    for (const sIdx of slotOrder) {
      const row = matrix[sIdx];
      let bestB = -1, bestScore = -Infinity;
      for (let bi = 0; bi < row.length; bi++) {
        if (usedBatters.has(bi)) continue;
        if (row[bi] > bestScore) { bestScore = row[bi]; bestB = bi; }
      }
      if (bestB >= 0 && bestScore > -Infinity) {
        lineupAssigned[sIdx] = bestB;
        usedBatters.add(bestB);
      }
    }
    // 反映
    lineupAssigned.forEach((bi, sIdx) => {
      slotInfo[sIdx].batSel.value = bi !== null ? String(bi) : '';
    });
    // 重複した選択を後発枠から除去 → 重複オプションを disable
    clearDuplicateSelections(side);
    refreshSelectionsForSide(side);
  }
}

function readSetup() {
  for (const side of ['away','home']) {
    // buildSetupSide と同じプールを使用 (インデックスの整合性を保つ)
    const pools = getSetupPools(side);
    const batters = pools.batterPool;
    // 先発 (必須)
    const starterSel = $('.sel-pitcher-slot[data-side="'+side+'"][data-idx="0"]');
    const myPitchers = [];
    const myRoles = [];
    if (starterSel && starterSel.value !== '') {
      const p = pools.starterPool[+starterSel.value];
      if (p) { myPitchers.push(p); myRoles.push('starter'); }
    }
    if (myPitchers.length === 0) {
      alert(`${side === 'away' ? 'AWAY' : 'HOME'}の先発投手を選んでください`);
      return false;
    }
    // リリーフ (中継/SU/抑え/モップ) — 任意・空欄可。スロット位置から役割を判定
    $$('.sel-relief-slot[data-side="'+side+'"]').forEach((sel, slotIdx) => {
      if (sel.value === '') return;
      const p = pools.reliefPool[+sel.value];
      if (p) { myPitchers.push(p); myRoles.push(reliefRoleForSlot(slotIdx)); }
    });
    // 投手重複は除外 (役割を並行して保持)
    const pSeen = new Set();
    const dedupPitchers = [];
    const dedupRoles = [];
    for (let i = 0; i < myPitchers.length; i++) {
      const p = myPitchers[i];
      const key = p.fullNameTop + '_' + (p.year || '');
      if (pSeen.has(key)) continue;
      pSeen.add(key);
      dedupPitchers.push(p);
      dedupRoles.push(myRoles[i]);
    }
    G.setup[side].pitchers = dedupPitchers;
    G.setup[side].pitcherRoles = dedupRoles;
    // 打者と守備ポジションを読む
    const liList = $$('.batter-slots[data-side="'+side+'"] li');
    G.setup[side].batters = [];
    G.setup[side].batterPos = [];
    const bSeen = new Set();
    for (let i = 0; i < 9; i++) {
      const li = liList[i];
      const posSel = li.querySelector('.sel-pos');
      const batSel = li.querySelector('.sel-batter');
      const pos = posSel.value;
      const v = batSel.value;
      if (v === '') {
        alert(`${side === 'away' ? 'AWAY' : 'HOME'}の${i+1}番(${POSITIONS[pos].label})の打者を選んでください\n(チーム選択にその守備位置の選手がいない場合は、別の選手・ポジション・チームに変更してください)`);
        return false;
      }
      const b = batters[+v];
      if (!b) {
        alert(`${side}の${i+1}番打者の選択が無効です。チーム選択と整合しているか確認してください`);
        return false;
      }
      if (!canPlay(b, pos)) {
        alert(`${side}の${i+1}番打者 (${b.fullNameTop}) は ${POSITIONS[pos].label} を守れません`);
        return false;
      }
      const key = b.fullNameTop + '_' + (b.year || '');
      if (bSeen.has(key)) {
        alert(`${side === 'away' ? 'AWAY' : 'HOME'}の打順で「${b.fullNameTop}」が重複しています`);
        return false;
      }
      bSeen.add(key);
      G.setup[side].batters.push(b);
      G.setup[side].batterPos.push(pos);
    }
    // 控え (PH/代走/守備) — 任意・空欄可。打順や他の控えとの重複は除外する
    // (bSeen には打順9人のキーが入っているので、それを引き継いで重複判定)
    // 役割ラベル(SETUP_BENCH_LABELS: PH1/PH2/PH3/代走/守備/守備)も保持する。
    G.setup[side].bench = [];
    G.setup[side].benchRole = new Map();  // 選手オブジェクト → 役割 ('PH'/'代走'/'守備')
    $$('.sel-bench-slot[data-side="'+side+'"]').forEach(sel => {
      if (sel.value === '') return;
      const b = batters[+sel.value];
      if (!b) return;
      const key = b.fullNameTop + '_' + (b.year || '');
      if (bSeen.has(key)) return;  // 打順 or 既出の控えと重複 → スキップ
      bSeen.add(key);
      G.setup[side].bench.push(b);
      // 役割を正規化して記録 (PH1/PH2/PH3 → PH)
      const rawRole = SETUP_BENCH_LABELS[+sel.dataset.idx] || '';
      const role = /^PH/.test(rawRole) ? 'PH' : rawRole;  // '代走' / '守備' / 'PH'
      G.setup[side].benchRole.set(b, role);
    });
  }
  G.innings = 9;
  return true;
}

// ============== 試合開始/リセット ==============
function startGame() {
  if (!readSetup()) return;
  G.seasonMode = false;
  beginGame();
}
// 試合開始の本体 (セットアップ済みの G.setup から試合を開始)。
// エキシビジョン(startGame)とレギュラーシーズン(seasonStartCurrentGame)で共用。
function beginGame() {
  G.autoToEnd = false;
  G.awaitingResult = false;
  updateAutoFinishButton();
  // 投手陣ごとにスタミナを初期化（カードのスタミナ値を上限とする）
  for (const side of ['away','home']) {
    const arr = G.setup[side].pitchers;
    G.setup[side].pitcherMax = arr.map(p => (p?.stats?.['スタミナ']) || 70);
    G.setup[side].pitcherStamina = [...G.setup[side].pitcherMax];
    G.setup[side].activeIdx = 0;
    // 野手交代用: 使用可能な控え (代打/代走/守備固め) と守備固め予約
    G.setup[side].benchAvail = [...(G.setup[side].bench || [])];
    G.setup[side].pendingDef = [];
  }
  // シーズンモード: 持ち越したスタミナ残量で上書き (回復は試合終了時に処理)
  if (G.seasonMode) applySeasonStamina();
  // 試合状態リセット
  G.inning = 1; G.top = true; G.outs = 0;
  G.fieldInn = {};   // この試合の守備イニング集計 (シーズン用・playerKey→{pos:イニング})
  G.starterBonusGiven = {};   // 先発の好投スタミナボーナス付与済みフラグ (side_inning → true)
  G.maxInnings = 15;  // タイブレーク延長は最大15回まで (G.innings=9 は正規イニング数)
  G.bases = [null, null, null];
  G.score = { away: [], home: [] };
  G.hits  = { away: 0, home: 0 };
  G.ks    = { away: 0, home: 0 };
  for (let i = 0; i < G.innings; i++) {
    G.score.away.push(0);
    G.score.home.push(0);
  }
  G.awayBatIdx = 0; G.homeBatIdx = 0;
  G.ended = false;
  G.awaitingResult = false;
  G.homeSkipBottomIdx = null;  // 後攻が9回裏を攻撃せず勝った場合、その回(0始まり)を記録 → スコアは「X」表示
  G.homeWalkoffIdx = null;     // 後攻がサヨナラ勝ち(最終回裏に勝ち越し)した回(0始まり) → スコアは「得点+x」表示
  G.lastPitchResult = null;    // 前試合の打球結果(軌道・ラベル)をクリア
  G.lastPitchRuns = 0;
  // 打席履歴。先頭(0番目)はプレイボール状態 (振り返りで「0: プレイボール」から開始できる)
  G.pitchHistory = [{
    res: null, runs: 0, bases: [null, null, null], inning: 1, top: true, outs: 0, playball: true,
    info: `🏟️ プレイボール！ ${G.innings}回戦\n(AWAY ${labelTeam('away')} vs HOME ${labelTeam('home')})`,
  }];
  G.historyView = null;        // 履歴閲覧インデックス (null=ライブ表示)
  G.lastInfo = '';             // ダイヤ外野に出すインフォ(交代/盗塁/代打/終了等)の最新文
  G.infoNew = false;           // 通知が新規(まだ打席結果と一緒に表示していない)か
  G.hrEvents = [];
  G.pitcherLog = { away: [], home: [] };
  G.leadHistory = [];
  // 打撃成績: 各打順スロット毎に初期化 (重複選手も別枠で管理)
  // subRole: 交代で入った選手の役割 (null=スタメン / '代打' / '代走' / '守備')
  // fielded: その選手が守備に就いたか (代打/代走が守備位置を引き継いだ判定に使用)
  G.batterStats = { away: [], home: [] };
  G.subLog = { away: [], home: [] };  // 交代で退いた選手の成績スナップショット (登板順)
  for (const side of ['away','home']) {
    G.setup[side].batters.forEach((b, i) => {
      G.batterStats[side].push(newBatterStat(side, i, b, null, true));
    });
  }
  // 各先発投手の登板を記録開始 (詳細stats付き)
  for (const side of ['away','home']) {
    G.pitcherLog[side].push(newPitcherLog(G.setup[side].pitchers[0], side, 1, (side === 'home')));
  }
  G.currentPitcher = G.setup.home.pitchers[0];
  G.currentBatter  = G.setup.away.batters[0];

  // 実況ログをクリア (前試合の残りを消す)
  const logEl = $('#log');
  if (logEl) logEl.innerHTML = '';

  // 投手登場の演出管理: 守備側の投手が前回登板から替わった時だけ defopit_sta を流す。
  //   1回表は先発HOMEが初登板済みとして記録(開始演出で流す)。AWAYはまだ未登板(null)。
  G._lastIntroPitcher = { home: G.currentPitcher, away: null };
  G.autoVideoPlaying = false;   // 自動再生は停止状態で開始 (ユーザーがボタンで開始)
  G._videoActive = false;
  G._videoRect = null;          // 動画表示位置を再計算 (最初の動画で固定)
  clearAutoVideoCountdown();

  // サイレント(シーズン自動高速進行)時は画面切替・描画をスキップ
  if (!G.silent) {
    showScreen('game');
    updateGameTopbar();
    buildScoreboard();
    renderAll();
    logLine(`🏟️ プレイボール！ ${G.innings}回戦\n(AWAY ${labelTeam('away')} vs HOME ${labelTeam('home')})`, 'event-inning');
    // 試合開始演出: 投手登場(1回表の先発) → 打者登場
    playVideoOverlay([pickVideo(PITCHER_INTRO_VIDEOS), pickVideo(BATTER_INTRO_VIDEOS)]);
  }
}

// 試合画面の操作バー: レギュラーシーズン手動試合では「シーズン/スタートへ戻る」を表示し、
//   リセット/セットアップを隠す。それ以外(エキシビ等)は従来通り。
function updateGameTopbar() {
  const seasonManual = !!(G.seasonMode && G.seasonCtx && !G.seasonCtx.auto);
  const vis = { resetGame: !seasonManual, backToSetupGame: !seasonManual, backToSeasonGame: seasonManual, backToStartGame: seasonManual };
  for (const id in vis) { const el = document.querySelector('#' + id); if (el) el.hidden = !vis[id]; }
}

function resetGame(toSetup) {
  G.ended = true;
  if (toSetup) showScreen('setup');
}

function labelTeam(side) {
  // チーム名: 先発投手のチーム名を採用 (末尾の所属チーム数を除いて表示)
  const t = G.setup[side];
  return normalizeTeam(t.pitchers[0]?.team) || '???';
}

// 守備側 = 攻撃の反対側
function defenseSide() { return G.top ? 'home' : 'away'; }

// 新しい投手登板ログを作る (詳細stats込み)
function newPitcherLog(pitcher, side, inning, top) {
  return {
    pitcher, side,
    runsAllowed: 0, battersFaced: 0,
    enterInning: inning, enterTop: top,
    outs: 0, pitches: 0,
    hits: 0, HR: 0, K: 0, BB: 0, HBP: 0, balks: 0,
    earnedRuns: 0,
  };
}
// 現役投手と そのスタミナを取得
function getActivePitcherInfo() {
  const ds = defenseSide();
  const setup = G.setup[ds];
  const idx = setup.activeIdx;
  return {
    side: ds,
    pitcher: setup.pitchers[idx],
    stamina: setup.pitcherStamina[idx],
    maxStamina: setup.pitcherMax[idx],
    idx,
    setup,
  };
}

// ============== 画面切替 ==============
function showScreen(name) {
  ['start','teambuild','setup','game','result','season'].forEach(s => {
    const el = $('#'+s);
    if (el) el.classList.toggle('hidden', s !== name);
  });
  // ヘッダー内の回/アウトはゲーム画面でのみ表示
  document.body.classList.toggle('in-game', name === 'game');
}

// ============== スコアボード ==============
// 表示すべきイニング列数 (延長時は最大 maxInnings まで増える)
function scoreboardCols() {
  return Math.min(G.maxInnings || 15,
                  Math.max(G.innings || 9, G.score.away.length, G.score.home.length));
}

function buildScoreboard() {
  const cols = scoreboardCols();
  const table = document.querySelector('.scoreboard table');
  if (table) table.classList.toggle('extra', cols > 9);  // 延長時は各列を縮小
  // ヘッダー行をイニング数に合わせて再構築
  const headRow = document.querySelector('.scoreboard thead tr');
  if (headRow) {
    let h = '<th class="team-name">チーム</th>';
    for (let i = 1; i <= cols; i++) h += `<th class="ib" data-ib="${i}">${i}</th>`;
    h += '<th>計</th><th>H</th><th>K</th>';
    headRow.innerHTML = h;
  }
  const rowA = $('#row-away');
  const rowH = $('#row-home');
  rowA.innerHTML = `<th class="team-name">${labelTeam('away')}</th>`;
  rowH.innerHTML = `<th class="team-name">${labelTeam('home')}</th>`;
  for (let i = 0; i < cols; i++) {
    const tdA = document.createElement('td'); tdA.id = `sc-away-${i}`;
    const tdH = document.createElement('td'); tdH.id = `sc-home-${i}`;
    rowA.appendChild(tdA); rowH.appendChild(tdH);
  }
  // 計, H, K
  for (const lbl of ['total','h','k']) {
    const tdA = document.createElement('td'); tdA.id = `sc-away-${lbl}`;
    const tdH = document.createElement('td'); tdH.id = `sc-home-${lbl}`;
    rowA.appendChild(tdA); rowH.appendChild(tdH);
  }
}

function updateScoreboard() {
  const cols = scoreboardCols();
  // 列数が変わった (延長突入など) ならヘッダー・行を再構築
  const existing = document.querySelectorAll('#row-away td').length; // = cols + 3 (計/H/K)
  if (existing !== cols + 3) buildScoreboard();
  for (let i = 0; i < cols; i++) {
    const a = $('#sc-away-'+i), h = $('#sc-home-'+i);
    if (!a || !h) continue;
    const awayCur = !G.ended && i === G.inning - 1 && G.top;    // AWAYが今この回の表を攻撃中
    const homeCur = !G.ended && i === G.inning - 1 && !G.top;   // HOMEが今この回の裏を攻撃中
    // AWAY(先攻): その回の表が終了済みなら表示
    // - 後の回に進んだ (i < G.inning - 1)
    // - 現在この回の裏 (i === G.inning - 1 && !G.top)
    // - 試合終了済み
    const awayDone = G.ended || (i < G.inning - 1) || (i === G.inning - 1 && !G.top);
    // 攻撃中でも得点が入っていればリアルタイム表示。無得点なら従来通り回終了まで空欄。
    const awayShow = awayDone || (awayCur && (G.score.away[i] || 0) > 0);
    a.textContent = awayShow ? (G.score.away[i] ?? 0) : '';
    // HOME(後攻): その回の裏が終了済みなら表示
    // - 次の回に進んだ (i < G.inning - 1)
    // - 試合終了済み
    // 後攻が攻撃せず勝った回は「X」/ サヨナラ勝ちの回は「得点+x」(野球サイト方式)
    const homeDone = G.ended || (i < G.inning - 1);
    const homeShow = homeDone || (homeCur && (G.score.home[i] || 0) > 0);
    h.textContent = (G.homeSkipBottomIdx === i) ? 'X'
      : (G.homeWalkoffIdx === i) ? ((G.score.home[i] ?? 0) + 'x')
      : (homeShow ? (G.score.home[i] ?? 0) : '');
    // 現在攻撃中のセルをハイライト
    a.classList.toggle('current', awayCur);
    h.classList.toggle('current', homeCur);
  }
  $('#sc-away-total').textContent = G.score.away.reduce((a,b)=>a+b,0);
  $('#sc-home-total').textContent = G.score.home.reduce((a,b)=>a+b,0);
  $('#sc-away-h').textContent = G.hits.away;
  $('#sc-home-h').textContent = G.hits.home;
  $('#sc-away-k').textContent = G.ks.away;
  $('#sc-home-k').textContent = G.ks.home;
}

// ============== 試合状況パネル ==============
// G.bases[i] は { side, slotIdx } の参照。実際の選手は batterStats か setup から取得
function getRunnerPlayer(ref) {
  if (!ref) return null;
  // 履歴スナップショット等で確定済みの走者選手があればそれを優先
  // (後の交代でスロットの選手が変わっても、その時点の走者を正しく表示する)
  if (ref._player) return ref._player;
  return G.batterStats?.[ref.side]?.[ref.slotIdx]?.player
      || G.setup?.[ref.side]?.batters?.[ref.slotIdx]
      || null;
}
function renderState() {
  renderBbState(G.inning, G.top, G.outs);
  renderRunnerBadges(G.bases);
}
// ダイヤモンド右下の「回 + アウトカウント」表示。アウトは数字でなく赤丸の点灯数で示す(電光掲示板風)。
function renderBbState(inning, top, outs) {
  const innEl = $('#bbStateInning');
  if (innEl) innEl.textContent = `${inning}回${top ? '表' : '裏'}`;
  const outsEl = $('#bbStateOuts');
  if (outsEl) {
    const o = Math.min(2, Math.max(0, outs || 0));   // 表示は0〜2 (3つ目はチェンジで次の回へ)
    outsEl.innerHTML = [0, 1].map(i => `<span class="out-dot${i < o ? ' on' : ''}"></span>`).join('');
  }
}
// 上の実況ラベルの3行目に「NEXT チーム N番 選手名」を追記する。
//   G.currentBatter は打席解決後に既に「次の打者」を指す (チェンジ時は次チームのその回の先頭打者)。
//   ライブ表示中のみ表示 (履歴閲覧・試合終了/待機・未開始では出さない)。
function renderNextBatterLine() {
  const label = document.querySelector('#bbLabel');
  if (!label) return;
  const old = label.querySelector('.bb-next');
  if (old) old.remove();   // 既存の3行目を消してから付け直す
  const b = G.currentBatter;
  if (G.historyView != null || G.ended || G.awaitingResult || !G.setup || !b) return;
  const side = G.top ? 'away' : 'home';
  const batIdx = side === 'away' ? G.awayBatIdx : G.homeBatIdx;
  const team = labelTeam(side);
  const name = b.fullNameTop || b.nameJa || '';
  const span = document.createElement('span');
  span.className = 'bb-next';
  span.innerHTML = `<span class="bn-lbl">NEXT</span> <span class="bn-team">${team} ${batIdx + 1}番</span> <span class="bn-name">${name}</span>`;
  label.appendChild(span);
}
// 走者バッジを描画 (1塁=bases[0], 2塁=bases[1], 3塁=bases[2])。履歴閲覧でも使うため bases を引数で受け取る
function renderRunnerBadges(bases) {
  for (let i = 0; i < 3; i++) {
    const badge = $('#runner-' + (i + 1));
    if (!badge) continue;
    const player = getRunnerPlayer(bases[i]);
    if (player) {
      const speed = player.stats?.['スピード'];
      const name  = player.fullNameTop || player.nameJa || '走者';
      badge.querySelector('.rn-speed').textContent = (speed != null) ? speed : '-';
      badge.querySelector('.rn-name').textContent  = name;
      badge.classList.add('visible');
    } else {
      badge.classList.remove('visible');
    }
  }
}

// ============== カード ミニ表示 ==============
function rankClass(v) {
  if (v == null) return '';
  if (v >= 90) return 's-rank';
  if (v >= 80) return 'a-rank';
  if (v < 0)   return 'neg';
  return '';
}

// ===== レギュラーシーズン手動試合: シーズン累積成績(今試合の途中経過込み)のライブ表示 =====
//   SEASON(過去試合の累積) + 現在試合(G) を合算して「現時点のシーズン成績」を返す。
function isSeasonManualGame() { return !!(G.seasonMode && G.seasonCtx && !G.seasonCtx.auto && SEASON && SEASON.bat); }
function seasonLiveBatAgg(p) {
  const k = playerKey(p);
  const s = seasonActiveStores().bat[k] || {};
  const a = { AB: s.AB || 0, H: s.H || 0, dbl: s.dbl || 0, tpl: s.tpl || 0, HR: s.HR || 0, RBI: s.RBI || 0, BB: s.BB || 0, HBP: s.HBP || 0, SAC: s.SAC || 0, SB: s.SB || 0, PA: s.PA || 0 };
  if (G.batterStats) for (const side of ['away', 'home']) {
    const add = bs => {
      if (!bs || !bs.player || playerKey(bs.player) !== k) return;
      a.AB += bs.AB || 0; a.H += bs.H || 0; a.dbl += bs.doubles || 0; a.tpl += bs.triples || 0; a.HR += bs.HR || 0;
      a.RBI += bs.RBI || 0; a.BB += bs.BB || 0; a.HBP += bs.HBP || 0; a.SAC += bs.SAC || 0; a.SB += bs.SB || 0;
      a.PA += (bs.AB || 0) + (bs.BB || 0) + (bs.HBP || 0) + (bs.SAC || 0);
    };
    (G.batterStats[side] || []).forEach(add);
    (G.subLog[side] || []).forEach(add);
  }
  return a;
}
function seasonLivePitAgg(p) {
  const k = playerKey(p);
  const s = seasonActiveStores().pit[k] || {};
  const a = { outs: s.outs || 0, ER: s.ER || 0, W: s.W || 0, L: s.L || 0, S: s.S || 0, HLD: s.HLD || 0, K: s.K || 0, H: s.H || 0, BB: s.BB || 0 };
  if (G.pitcherLog) for (const side of ['away', 'home']) {
    (G.pitcherLog[side] || []).forEach(lg => {
      if (!lg || !lg.pitcher || playerKey(lg.pitcher) !== k) return;
      a.outs += lg.outs || 0; a.ER += lg.earnedRuns || 0; a.K += lg.K || 0; a.H += lg.hits || 0; a.BB += lg.BB || 0;
    });
  }
  return a;
}
function seasonLiveAvg(p) { const a = seasonLiveBatAgg(p); return fmtAvg(a.H, a.AB); }
function seasonLiveERA(p) { const a = seasonLivePitAgg(p); return a.outs > 0 ? fmtERA(a.ER, a.outs) : '-.--'; }
function seasonLiveBatRec(p) {
  const a = seasonLiveBatAgg(p);
  const obp = a.PA > 0 ? (a.H + a.BB + a.HBP) / a.PA : 0;
  const tb = a.H + a.dbl + 2 * a.tpl + 3 * a.HR;
  const slg = a.AB > 0 ? tb / a.AB : 0;
  const card = p.record || {};
  return { '打率': fmtAvg(a.H, a.AB), '本塁打': a.HR, '打点': a.RBI, '盗塁': a.SB, '出塁率': fmt3(obp), 'OPS': fmt3(obp + slg), 'WAR': card['WAR'] };
}
function seasonLivePitRec(p) {
  const a = seasonLivePitAgg(p);
  const card = p.record || {};
  return { '防御率': (a.outs > 0 ? fmtERA(a.ER, a.outs) : '-.--'), '勝利': a.W, '敗北': a.L, 'セーブ': a.S, 'イニング': fmtIP(a.outs), '奪三振': a.K, 'WAR': card['WAR'] };
}

function renderBatterCard(b) {
  const s = b.stats || {};
  const m = b.statsMini || {};
  const r = isSeasonManualGame() ? seasonLiveBatRec(b) : (b.record || {});
  const photoHtml = b.photo
    ? `<img src="${b.photo}" alt="${b.fullNameTop}">`
    : `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#666;font-size:11px;">No Photo</div>`;
  // 今日の成績 (試合中に蓄積された stats)
  const battingSide = G.top ? 'away' : 'home';
  const bIdx = G.top ? G.awayBatIdx : G.homeBatIdx;
  const todayStat = G.batterStats?.[battingSide]?.[bIdx];
  const tsHtml = (todayStat && !G.ended) ? `
    <div class="today-stats">
      <div class="ts-title">本日の成績</div>
      <div class="ts-row">
        <div><span class="ts-lbl">打数</span><span class="ts-val">${todayStat.AB}</span></div>
        <div><span class="ts-lbl">安打</span><span class="ts-val hl">${todayStat.H}</span></div>
        <div><span class="ts-lbl">本塁打</span><span class="ts-val pp">${todayStat.HR}</span></div>
        <div><span class="ts-lbl">打点</span><span class="ts-val hl">${todayStat.RBI}</span></div>
        <div><span class="ts-lbl">三振</span><span class="ts-val">${todayStat.K}</span></div>
      </div>
    </div>
  ` : '';
  // 対球種ポイント (FB/2C/CT/SL/CB/CH/SF) と 守備DRS — フルカードに近づける
  const pp = b.pitchPoints || {};
  const ppKeys = ['FB','2C','CT','SL','CB','CH','SF'];
  const fmtP = v => (v == null || v === 0) ? '±0' : (v > 0 ? '+' + v : '' + v);
  const pcl  = v => (v == null || v === 0) ? 'neu' : (v > 0 ? 'pos' : 'neg');
  const ppRow = keys => keys.map(k => `<span class="${pcl(pp[k])}">${k} <b>${fmtP(pp[k])}</b></span>`).join('');
  const ppHtml = ppKeys.some(k => pp[k] != null && pp[k] !== 0)
    ? `<div class="card-pp"><div class="ppr">${ppRow(['FB','2C','CT','SL'])}</div><div class="ppr">${ppRow(['CB','CH','SF'])}</div></div>`
    : '';
  const drsHtml = (b.drs && b.drs.length)
    ? `<div class="card-drs"><span class="dlbl">守備</span>${b.drs.map(d => `<span class="${d.value>0?'pos':(d.value<0?'neg':'neu')}">${d.pos}<b>${d.value>0?'+':''}${d.value??0}</b></span>`).join('')}</div>`
    : '';
  return `
    <div class="card-grid">
      <div class="card-left">
        <div class="card-photo">
          ${photoHtml}
          <span class="team-badge-mini">${b.team || '-'}</span>
          <div class="photo-fade"></div>
          <div class="name-overlay player-link${longNameClass(b.fullNameTop)}" data-player-name="${b.fullNameTop}" data-player-year="${b.year ?? ''}" data-player-type="${playerType(b)}" data-player-team="${b.team||''}" title="クリックで詳細カードを表示">${b.fullNameTop}</div>
        </div>
        <div class="mini-banner">
          <span class="season">${cardSeason(b)} / ${b.position} / ${b.hand || ''}</span>
        </div>
      </div>
      <div class="card-info">
        <div class="card-rec">
          <span><i>打率</i><b class="hl">${r['打率']||'-'}</b></span>
          <span><i>本</i><b class="pp">${r['本塁打']||'-'}</b></span>
          <span><i>点</i><b>${r['打点']||'-'}</b></span>
          <span><i>盗</i><b>${r['盗塁']||'-'}</b></span>
          <span><i>出塁</i><b class="hl">${r['出塁率']||'-'}</b></span>
          <span><i>OPS</i><b class="hl">${r['OPS']||'-'}</b></span>
          <span><i>WAR</i><b class="hl">${r['WAR']||'-'}</b></span>
        </div>
        <div class="mini-stats">
          <div class="msr msr4">
            <div class="stbox"><span class="sname">ミート</span><span class="sval ${rankClass(s['ミート'])}">${s['ミート']??'-'}</span></div>
            <div class="stbox"><span class="sname">パワー</span><span class="sval ${rankClass(s['パワー'])}">${s['パワー']??'-'}</span></div>
            <div class="stbox"><span class="sname">スピード</span><span class="sval ${rankClass(s['スピード'])}">${s['スピード']??'-'}</span></div>
            <div class="stbox"><span class="sname">チャンス</span><span class="sval ${rankClass(s['チャンス'])}">${s['チャンス']??'-'}</span></div>
          </div>
          <div class="msr msr-r2">
            <div class="stbox"><span class="sname">選球眼</span><span class="sval ${rankClass(s['選球眼'])}">${s['選球眼']??'-'}</span></div>
            <div class="stbox"><span class="sname">三振耐性</span><span class="sval ${rankClass(s['三振耐性'])}">${s['三振耐性']??'-'}</span></div>
            <div class="mssub3">
              <div class="stbox stbox-sm"><span class="sname">盗塁能</span><span class="sval ${(m['盗塁能']!=null && m['盗塁能']<0)?'neg':''}">${m['盗塁能']??0}</span></div>
              <div class="stbox stbox-sm"><span class="sname">対左</span><span class="sval ${(m['対左投手']!=null && m['対左投手']<0)?'neg':''}">${m['対左投手']??0}</span></div>
              <div class="stbox stbox-sm"><span class="sname">HR能</span><span class="sval ${(m['HR能']!=null && m['HR能']<0)?'neg':''}">${m['HR能']??0}</span></div>
            </div>
          </div>
        </div>
        ${ppHtml}
        ${drsHtml}
        ${tsHtml}
      </div>
    </div>`;
}

function renderPitcherCard(p) {
  const s = p.stats || {};
  const m = p.statsMini || {};
  const r = isSeasonManualGame() ? seasonLivePitRec(p) : (p.record || {});
  const photoHtml = p.photo
    ? `<img src="${p.photo}" alt="${p.fullNameTop}">`
    : `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#666;font-size:11px;">No Photo</div>`;
  // 現役投手なら残スタミナバーをカード内に表示
  let staminaHtml = '';
  const info = getActivePitcherInfo();
  if (info && info.pitcher === p && !G.ended) {
    // スタミナは負値もあり得るので 0..1 にクランプ
    const ratio = info.maxStamina > 0 ? Math.max(0, info.stamina / info.maxStamina) : 0;
    const cls = info.stamina <= 0 ? 'crit' : (info.stamina <= 15 ? 'low' : '');
    staminaHtml = `
      <div class="card-stamina">
        <span class="cs-lbl">スタミナ</span>
        <div class="cs-bar"><div class="cs-fill ${cls}" style="width:${(ratio*100).toFixed(1)}%"></div></div>
        <span class="cs-num">${info.stamina}/${info.maxStamina}</span>
      </div>
    `;
  }
  // 球種選択を HTMLカード風の表 (球種/球速/球威/割合) で表示。各行がクリック可能。
  let pitchBtns = '';
  if (p.pitches && p.pitches.length) {
    pitchBtns = '<div class="pitch-table-inline">'
      + '<div class="pti-head"><span class="pn">球種</span><span>球速</span><span>球威</span><span>割合</span></div>';
    for (const pi of p.pitches.slice(0, 6)) {   // 投手は最大6球種まで表示
      pitchBtns += `<button type="button" class="pitch-btn pti-row" data-pitch="${pi.name}" title="クリックで ${stripPitchAlias(pi.name)} を投げる">
        <span class="pn">${stripPitchAlias(pi.name)}</span>
        <span class="psp">${pi.speed ?? '-'}</span>
        <span class="ppw">${pi.power ?? '-'}</span>
        <span class="prt">${pi.ratio ?? '-'}%</span>
      </button>`;
    }
    pitchBtns += '</div>';
  }
  return `
    <div class="card-grid">
      <div class="card-left">
        <div class="card-photo">
          ${photoHtml}
          <span class="team-badge-mini">${p.team || '-'}</span>
          <div class="photo-fade"></div>
          <div class="name-overlay player-link${longNameClass(p.fullNameTop)}" data-player-name="${p.fullNameTop}" data-player-year="${p.year ?? ''}" data-player-type="${playerType(p)}" data-player-team="${p.team||''}" title="クリックで詳細カードを表示">${p.fullNameTop}</div>
        </div>
        <div class="mini-banner">
          <span class="season">${cardSeason(p)} / ${p.position} / ${p.hand || ''}</span>
        </div>
        ${staminaHtml}
      </div>
      <div class="card-info">
        <div class="card-rec">
          <span><i>防</i><b class="hl">${r['防御率']||'-'}</b></span>
          <span><i>勝</i><b>${r['勝利']||'-'}</b></span>
          <span><i>敗</i><b>${r['敗北']||'-'}</b></span>
          <span><i>S</i><b>${r['セーブ']||'-'}</b></span>
          <span><i>回</i><b>${r['イニング']||'-'}</b></span>
          <span><i>奪</i><b class="pp">${r['奪三振']||'-'}</b></span>
          <span><i>WAR</i><b class="hl">${r['WAR']||'-'}</b></span>
        </div>
        <div class="pitcher-stats-list">
          <div class="psr psr4">
            <div class="pst"><span class="lbl">スタミナ</span><span class="val ${rankClass(s['スタミナ'])}">${s['スタミナ']??'-'}</span></div>
            <div class="pst"><span class="lbl">制球</span><span class="val ${rankClass(s['制球'])}">${s['制球']??'-'}</span></div>
            <div class="pst"><span class="lbl">緩急</span><span class="val ${rankClass(s['緩急'])}">${s['緩急']??'-'}</span></div>
            <div class="pst"><span class="lbl">精神</span><span class="val ${rankClass(s['精神'])}">${s['精神']??'-'}</span></div>
          </div>
          <div class="psr psr-r2">
            <div class="pst"><span class="lbl">奪三振</span><span class="val ${rankClass(s['奪三振'])}">${s['奪三振']??'-'}</span></div>
            <div class="pst"><span class="lbl">球重</span><span class="val ${rankClass(s['球重'])}">${s['球重']??'-'}</span></div>
            <div class="psub3">
              <div class="pst pst-sm"><span class="lbl">対左</span><span class="val ${(m['対左']!=null && m['対左']<0)?'neg':''}">${m['対左']??0}</span></div>
              <div class="pst pst-sm"><span class="lbl">対盗塁</span><span class="val ${(m['対盗塁']!=null && m['対盗塁']<0)?'neg':''}">${m['対盗塁']??0}</span></div>
              <div class="pst pst-sm"><span class="lbl">回復量</span><span class="val">${s['回復量']??m['回復量']??'-'}</span></div>
            </div>
          </div>
        </div>
        ${pitchBtns}
      </div>
    </div>`;
}

function renderAll() {
  if (G.silent) return;   // シーズン自動高速進行中は描画しない
  renderState();
  renderBattedBall();
  renderNextBatterLine();   // 実況ラベル3行目: 次打者(NEXT)
  updateScoreboard();
  renderSideBoards();
  const pcEl = $('#card-pitcher'), bcEl = $('#card-batter');
  pcEl.innerHTML = renderPitcherCard(G.currentPitcher);
  bcEl.innerHTML = renderBatterCard(G.currentBatter);
  applyRarityClass(pcEl, G.currentPitcher);
  applyRarityClass(bcEl, G.currentBatter);
  renderHistoryBar();
  renderBbInfo();
  updateAutoVideoBar();   // 自動再生バー(動画ON時)の表示・状態を同期
}
// 総合力からレアリティ相当の背景色クラスを付与 (RR=銀 / SR=金 / SSR=UR特殊 / N=赤 / N相当=黒)
function applyRarityClass(el, p) {
  if (!el) return;
  el.classList.remove('rar-ur','rar-ssr','rar-sr','rar-rr','rar-r','rar-n','rar-n0');
  el.classList.add('rar-' + cardRarity(p));
}
// カードのレアリティを判定。UR は rawHtml の上部表記で検出 (UR専用の虹色背景)。
// それ以外は総合力で近似 (UR=100+ / SSR=90+ / SR=80+ / RR=70+ / R=60+ / N=50+ / N相当=50未満)。
function cardRarity(p) {
  const head = String((p && p.rawHtml) || '').slice(0, 700);
  if (/(?:^|[>\s"'（(\[])UR(?=[<\s★＊*"'）)\]]|$)/.test(head)) return 'ur';
  const v = overallOf(p);
  if (!Number.isFinite(v)) return 'rr';
  if (v >= 100) return 'ur';   // 総合力100以上は UR (虹色ホログラム)
  if (v >= 90)  return 'ssr';
  if (v >= 80)  return 'sr';
  if (v >= 70)  return 'rr';
  if (v >= 60)  return 'r';
  if (v >= 50)  return 'n';    // 50〜59: N (赤)
  return 'n0';                 // 50未満: N相当 (黒地に白文字)
}

// 打球の可視化 (ネットニュース「1球速報」風):
//   直近の打席結果(G.lastPitchResult)から、打球の方向と処理内容を球場に描画する。
//   守備位置(fielderPos)と結果種別(outcome)から軌道(ゴロ=バウンド / フライ・安打=放物線 / 本塁打=大きな弧)を生成。
function renderBattedBall() {
  const overlay = document.querySelector('#bbOverlay');
  const label   = document.querySelector('#bbLabel');
  if (!overlay || !label) return;
  const clear = () => { overlay.innerHTML = ''; label.textContent = ''; label.className = 'bb-label'; };
  // 履歴閲覧中(試合終了後の 戻る/進む)はその時点の結果を、通常はライブの直近結果を表示
  let res, runs;
  if (G.historyView != null && G.pitchHistory && G.pitchHistory[G.historyView]) {
    res  = G.pitchHistory[G.historyView].res;
    runs = G.pitchHistory[G.historyView].runs || 0;
  } else {
    res  = G.lastPitchResult;
    runs = G.lastPitchRuns || 0;
  }
  // まだ打球結果が無い (試合開始直後 or 振り返りの0番=プレイボール) → 「試合開始！！」を表示
  if (!res || !res.outcome || G.inning == null) {
    overlay.innerHTML = '';
    if (G.historyView != null || (!G.ended && !G.awaitingResult)) {
      label.className = 'bb-label out'; label.innerHTML = '<span class="bb-result">試合開始！！</span>';
    } else {
      label.textContent = ''; label.className = 'bb-label';
    }
    return;
  }
  // 最終結果種別を優先 (犠飛が成立しなかった大飛球は FO、ゲッツーは GO_DP 等)
  const outcome = res.finalOutcome || res.outcome;
  let pos = res.fielderPos || 'P';
  if (res._dispRand == null) res._dispRand = Math.random();   // この打席で安定した振り直し用乱数
  const rnd = res._dispRand;
  // アウトカウント/チェンジの接尾辞 (アウト系の結果に付加)。3アウト目は「チェンジ」表示。
  const outSuffix = () => {
    if (res.isThirdOut) return ' チェンジ';
    if (res.outsAfter != null) return ` ${res.outsAfter}アウト`;
    return '';
  };
  // ホームラン/ヒットの判定表現 (実況フレーバー由来): 通常HR=！！ / 一発=！？, 快打=！ / きわどい安打=？！
  const hrExcl  = (res.flavor && res.flavor.includes('！？')) ? '！？' : '！！';
  const hitExcl = (res.flavor && res.flavor.includes('？！')) ? '？！' : '！';

  // ファールフライ: 実況がファールフライの打球は、判定守備位置に応じた名称を表示する。
  //   外野(LF/RF/CF)=レフト/ライト、1B/2B=一塁、3B/SS=三塁、C/DH/その他=キャッチャー。
  //   ※この簡易フィールドではフェア/ファールを描き分けられないため、軌道画像は表示しない(ラベルのみ)。
  if (res.flavor && /ファール/.test(res.flavor)) {
    const fp = res.fielderPos || 'C';
    let foulLabel;
    if (fp === 'RF')                     foulLabel = 'ライトファールフライ';
    else if (fp === 'LF')                foulLabel = 'レフトファールフライ';
    else if (fp === 'CF')                foulLabel = (rnd < 0.5) ? 'レフトファールフライ' : 'ライトファールフライ';
    else if (fp === '1B' || fp === '2B') foulLabel = '一塁ファールフライ';
    else if (fp === '3B' || fp === 'SS') foulLabel = '三塁ファールフライ';
    else                                 foulLabel = 'キャッチャーファールフライ';
    overlay.innerHTML = '';   // ファールゾーンの軌道は描画しない
    const pLine = (res.dispPitcher && res.dispPitch) ? `${res.dispPitcher}の${res.dispPitch}を` : '';
    const bName = res.dispBatter || '';
    label.className = 'bb-label out';
    label.innerHTML = `${pLine ? `<span class="bb-pitch">${pLine}</span>` : ''}<span class="bb-result">${bName ? bName + ' ' : ''}${foulLabel}${outSuffix()}</span>`;
    return;
  }

  // 守備位置の座標 (viewBox 200x180, 本塁=100,150)
  const FIELDER = {
    'C':{x:100,y:140},'P':{x:100,y:106},'1B':{x:134,y:112},'2B':{x:120,y:86},
    'SS':{x:80,y:86},'3B':{x:66,y:112},'LF':{x:50,y:50},'CF':{x:100,y:34},'RF':{x:150,y:50},
  };
  const ABBR = {'C':'捕','P':'投','1B':'一','2B':'二','3B':'三','SS':'遊','LF':'左','CF':'中','RF':'右'};
  const isIF = (p) => ['C','P','1B','2B','3B','SS'].includes(p);
  const isOF = (p) => ['LF','CF','RF'].includes(p);

  // 打球の「深さ」を判定 (実況フレーバー=ログ表示と一致させ、不整合を防ぐ)。
  //   IF=内野ゴロ/内野フライ, OF=外野まで届く打球, LINER=ライナー, KEEP=捕球位置のまま(失策/ファインプレー等)
  const fl = res.flavor || '';
  let depth;
  if (res.fineplay)                  depth = 'KEEP';   // ファインプレー: 好守した野手の位置をそのまま
  else if (res.clumsy)               depth = 'KEEP';   // 拙守(安打): 処理した野手の位置をそのまま
  else if (outcome === 'E')          depth = 'KEEP';   // 失策: 処理した野手の位置をそのまま
  else if (/外野/.test(fl))           depth = 'OF';     // 「(大きな/浅い)外野フライ」→ 外野方向へ
  else if (/内野|ファール/.test(fl))   depth = 'IF';     // 「内野フライ/ゴロ/安打」「ファールフライ」→ 内野へ
  else {
    const oc = outcome;
    if (oc === 'GO' || oc === 'GO_SLOW' || oc === 'GO_DP') depth = 'IF';
    else if (oc === 'LO') depth = 'LINER';
    else if (oc === 'FO' && res.infieldFly) depth = 'IF';
    else if (oc === 'FO') depth = 'KEEP';                // 方向不明のフライ(ファインプレー等)は捕球位置のまま
    else depth = 'OF';                                    // 本塁打/二塁打/三塁打/安打 等
  }

  // 処理位置と深さが矛盾する場合のみ、現実的な守備位置へ振り直す
  if (depth === 'IF' && isOF(pos)) {
    // 外野指定だが内野の打球: レフト→遊/三, センター→遊/二, ライト→二/一
    const m = ({ LF:['SS','3B'], CF:['SS','2B'], RF:['2B','1B'] })[pos] || ['SS'];
    pos = m[rnd < 0.5 ? 0 : 1];
  } else if (depth === 'OF' && isIF(pos)) {
    // 内野指定だが外野まで届いた打球: 三→左, 遊→左/中, 二→中/右, 一→右, 捕投→中
    const m = ({ '3B':['LF'], 'SS':['LF','CF'], '2B':['CF','RF'], '1B':['RF'], 'C':['CF'], 'P':['CF'] })[pos] || ['CF'];
    pos = (m.length > 1) ? m[rnd < 0.5 ? 0 : 1] : m[0];
  }
  // 「捕手ライナー」は非現実的なので、捕手以外の内野手 (遊/二/三/一) へ振り分ける
  if (outcome === 'LO' && pos === 'C') {
    const opts = ['SS', '2B', '3B', '1B'];
    pos = opts[Math.min(opts.length - 1, Math.floor(rnd * opts.length))];
  }

  const dir = (pos === 'RF' || pos === '1B') ? '右' : (pos === 'LF' || pos === '3B') ? '左' : '中';
  const OF_LAND = { '右':{x:152,y:58}, '中':{x:100,y:44}, '左':{x:48,y:58} };
  const HR_LAND = { '右':{x:166,y:16}, '中':{x:100,y:10}, '左':{x:34,y:16} };

  const sx = 100, sy = 150;
  // 放物線 (打ち上げ): arch が大きいほど高く上がる
  const arc = (t, arch) => { const mx = (sx + t.x) / 2, my = (sy + t.y) / 2 - arch; return `M ${sx} ${sy} Q ${mx} ${my} ${t.x} ${t.y}`; };
  // ゴロ (小さなバウンドを繰り返しながら野手へ転がる)
  const ground = (t) => {
    let d = `M ${sx} ${sy}`; const n = 3;
    for (let i = 1; i <= n; i++) {
      const px = sx + (t.x - sx) * (i / n), py = sy + (t.y - sy) * (i / n);
      const pxp = sx + (t.x - sx) * ((i - 1) / n), pyp = sy + (t.y - sy) * ((i - 1) / n);
      const cx = (pxp + px) / 2, cy = (pyp + py) / 2 - 15 * (1 - (i - 1) / n * 0.5);
      d += ` Q ${cx} ${cy} ${px} ${py}`;
    }
    return d;
  };

  const runsTxt = (runs > 0) ? `(${runs}点)` : '';
  // 打球の強弱(実況フレーバー由来)をラベルに反映: ゴロ=緩/強, 外野フライ=浅/深
  let intens = '';
  if (!res.fineplay) {
    if (outcome === 'GO' || outcome === 'GO_SLOW') {
      if (/ボテボテ/.test(res.flavor || '')) intens = '(緩)';   // ボテボテのみ(緩)。通常ゴロは無印
    } else if (outcome === 'FO' && !res.infieldFly && !/内野/.test(res.flavor || '')) {
      if (/大きな|大飛球/.test(res.flavor || ''))  intens = '(深)';
      else if (/浅い/.test(res.flavor || ''))       intens = '(浅)';
    }
  }
  let d = '', color = '#fff', target = null, resultText = '', cls = 'out';
  switch (outcome) {
    case 'HR':      target = HR_LAND[dir]; d = arc(target, 72); color = '#ff66cc'; cls = 'hr';  resultText = `${dir}本塁打${runsTxt}${hrExcl}`; break;
    case '1B':
      if (res.clumsy) {   // 拙守によるヒット: 処理した野手の位置で安打、(○○拙守) を付記
        target = FIELDER[pos] || FIELDER.CF;
        d = (res.clumsyFrom === 'FO' || res.clumsyFrom === 'LO') ? arc(target, 26) : ground(target);
        resultText = `${dir}安打${runsTxt}(${res.clumsyFielder || '守備'}拙守)`;
      } else if (res.flavor && res.flavor.includes('内野')) {   // 内野安打 (足を生かした内野安打) は内野へのゴロ
        target = FIELDER[pos] || FIELDER.SS; d = ground(target); resultText = `${ABBR[pos] || ''}内野安打${runsTxt}！`;
      } else {
        target = OF_LAND[dir]; d = arc(target, 30); resultText = `${dir}安打${runsTxt}${hitExcl}`;
      }
      color = '#8aff66'; cls = 'hit'; break;
    case '2B':      target = OF_LAND[dir]; d = arc(target, 26); color = '#8aff66'; cls = 'hit'; resultText = `${dir}二塁打${runsTxt}${hitExcl}`; break;
    case '3B':      target = OF_LAND[dir]; d = arc(target, 24); color = '#8aff66'; cls = 'hit'; resultText = `${dir}三塁打${runsTxt}${hitExcl}`; break;
    case 'E': {
      const ferr = res.fielder ? (res.fielder.fullNameTop || res.fielder.nameJa || '守備') : '守備';
      target = FIELDER[pos] || FIELDER.SS;
      d = (res.errorFrom === 'FO' || res.errorFrom === 'LO') ? arc(target, 26) : ground(target);  // 飛球の失策は弧、ゴロの失策はバウンド
      color = '#ffcc44'; cls = 'err';
      resultText = `エラー出塁(${ferr}失策)${runsTxt}`; break;
    }
    case 'GO': case 'GO_SLOW': case 'GO_DP':
                    target = FIELDER[pos] || FIELDER.SS; d = ground(target); color = '#ffffff'; cls = 'out'; resultText = `${ABBR[pos] || ''}ゴロ${outcome === 'GO_DP' ? '併殺' : ''}${intens}${runsTxt}`; break;
    case 'SAC_FLY': target = OF_LAND[dir]; d = arc(target, 46); color = '#ffffff'; cls = 'out'; resultText = `${dir}犠飛${runsTxt}`; break;
    case 'FO':      target = FIELDER[pos] || FIELDER.CF; d = arc(target, 40); color = '#ffffff'; cls = 'out';
                    resultText = `${ABBR[pos] || ''}${(pos === 'LF' || pos === 'CF' || pos === 'RF') ? '飛' : 'フライ'}${intens}${runsTxt}`; break;
    case 'LO':      target = FIELDER[pos] || FIELDER.CF; d = arc(target, 12); color = '#ffffff'; cls = 'out'; resultText = `${ABBR[pos] || ''}ライナー${runsTxt}`; break;
    case 'K':       cls = 'out'; resultText = '三振'; break;               // 軌道なし (アウトカウントは後で付加)
    case 'BB': {                                                           // 軌道なし (押し出し時は得点も表示)
      cls = 'bb';
      const bbReason = (res.flavor && res.flavor.includes('選球')) ? '(選球)'
                     : (res.flavor && res.flavor.includes('制球')) ? '(制球)' : '';
      resultText = `四球${bbReason}${runsTxt}`;
      break;
    }
    default:        clear(); return;
  }
  // ファインプレー: 結果に好守した野手名を付記 (例: 中飛(Wiメイズ好守))
  if (res.fineplay && (outcome === 'GO' || outcome === 'FO')) {
    const fn = res.fielder ? (res.fielder.fullNameTop || res.fielder.nameJa || '守備') : '守備';
    resultText += `(${fn}好守)`;
  }
  // アウト系の結果にはアウトカウント (3アウト目は「チェンジ」) を付記
  const OUT_SET = new Set(['K','FO','LO','GO','GO_SLOW','GO_DP','SAC_FLY']);
  if (OUT_SET.has(outcome)) resultText += outSuffix();
  // 打球軌道 (空振り/四球は軌道なし)
  if (d && target) {
    overlay.innerHTML = `<svg viewBox="0 0 200 180" preserveAspectRatio="xMidYMid meet">
      <path d="${d}" fill="none" stroke="${color}" stroke-width="2.4" stroke-linecap="round" opacity="0.95"/>
      <circle cx="${target.x}" cy="${target.y}" r="3.4" fill="${color}" stroke="#000" stroke-width="0.6"/>
    </svg>`;
  } else {
    overlay.innerHTML = '';
  }
  // 結果ラベル (2行): 「投手の球速・球種を」 / 「打者　結果(得点)！」
  const pitchLine = (res.dispPitcher && res.dispPitch) ? `${res.dispPitcher}の${res.dispPitch}を` : '';
  const batter = res.dispBatter || '';
  label.className = 'bb-label ' + cls;
  label.innerHTML = `${pitchLine ? `<span class="bb-pitch">${pitchLine}</span>` : ''}<span class="bb-result">${batter ? batter + ' ' : ''}${resultText}</span>`;
}

// 試合終了後の振り返りバー (ダイヤモンド枠内の イニング選択 / 戻る / 進む)。試合終了(結果待ち)中のみ表示。
function renderHistoryBar() {
  const bar = document.querySelector('#bbHistory');
  if (!bar) return;
  const hist = G.pitchHistory || [];
  if ((!G.awaitingResult && !G.ended) || hist.length === 0) { bar.hidden = true; return; }
  bar.hidden = false;
  const idx = (G.historyView == null) ? hist.length - 1 : G.historyView;
  const countEl = document.querySelector('#bbHistCount');
  // 先頭(0)はプレイボール。0始まりで表示 (0=プレイボール, 1..N=各打席)
  if (countEl) countEl.textContent = `${idx} / ${hist.length - 1}`;
  const prev = document.querySelector('#bbHistPrev');
  const next = document.querySelector('#bbHistNext');
  if (prev) prev.disabled = (idx <= 0);
  if (next) next.disabled = (idx >= hist.length - 1);
  // イニング選択プルダウン: 出現した各ハーフイニングを列挙 (内容が変わった時のみ再構築)
  const sel = document.querySelector('#bbHistInning');
  if (sel) {
    const seen = new Set(), opts = [];
    hist.forEach(h => {
      const key = h.inning + '|' + h.top;
      if (!seen.has(key)) { seen.add(key); opts.push({ key, label: `${h.inning}回${h.top ? '表' : '裏'}` }); }
    });
    const sig = opts.map(o => o.key).join(',');
    if (sel._sig !== sig) {
      sel._sig = sig;
      sel.innerHTML = opts.map(o => `<option value="${o.key}">${o.label}</option>`).join('');
    }
    const e = hist[idx];
    if (e) sel.value = e.inning + '|' + e.top;   // 現在地のイニングを選択状態に
  }
}
// 指定インデックスの打席を表示する (戻る/進む/イニングジャンプ共通)
function historyGoto(idx) {
  const hist = G.pitchHistory || [];
  if (hist.length === 0) return;
  idx = Math.max(0, Math.min(hist.length - 1, idx));
  G.historyView = idx;
  const e = hist[idx];
  renderBattedBall();              // ラベル・軌道を履歴から再描画
  renderRunnerBadges(e.bases);     // その打席終了時点の塁状況
  renderBbState(e.inning, e.top, e.outs);  // その時点の回/アウト
  renderBbInfo(e.info || '');      // その打席時点の通知(継投/代打等)を時系列表示 (試合終了文字は出さない)
  renderHistoryBar();
}
// 履歴を delta(±1)だけ移動
function historyStep(delta) {
  const hist = G.pitchHistory || [];
  if (hist.length === 0) return;
  const cur = (G.historyView == null) ? hist.length - 1 : G.historyView;
  historyGoto(cur + delta);
}
// 指定したハーフイニング("inning|top")の先頭打席へジャンプ
function historyJumpToInning(key) {
  const hist = G.pitchHistory || [];
  const idx = hist.findIndex(e => (e.inning + '|' + e.top) === key);
  if (idx >= 0) historyGoto(idx);
}

// 両サイドの打順・投手ボードを描画 (試合中)
function renderSideBoards() {
  // 試合終了後は、投手ボードに勝敗/セーブ/ホールド(W/L/S/H)を表示するため判定を算出
  const decisionMap = G.ended ? computePitcherDecisions().pitcherRoles : null;
  for (const side of ['away','home']) {
    const stats = G.batterStats[side] || [];
    const teamName = labelTeam(side);
    const titleEl = $('#bt-' + side);
    if (titleEl) titleEl.textContent = teamName;
    // 現在打席中の打者スロット
    const isBattingTeam = (side === (G.top ? 'away' : 'home'));
    const curBatIdx = side === 'away' ? G.awayBatIdx : G.homeBatIdx;
    // 打順テーブル
    const rows = stats.map((s, i) => {
      const isBatting = isBattingTeam && i === curBatIdx && !G.ended;
      return `<tr class="${isBatting ? 'is-batting' : ''}">
        <td class="ln-num">${i+1}</td>
        <td class="ln-name player-link${longNameClass(s.player.fullNameTop)}" data-player-name="${s.player.fullNameTop}" data-player-year="${s.player.year ?? ''}" data-player-type="${playerType(s.player)}" data-player-team="${s.player.team||''}">${s.player.fullNameTop}</td>
        <td>${isSeasonManualGame() ? seasonLiveAvg(s.player) : fmtAvg(s.H, s.AB)}</td>
        <td class="ln-pos">${s.position || '-'}</td>
        <td>${s.AB}</td>
        <td>${s.H}</td>
        <td>${s.HR}</td>
        <td>${s.RBI}</td>
        <td>${s.BB}</td>
        <td>${s.SB}</td>
        <td>${s.E || 0}</td>
      </tr>`;
    }).join('');
    const lineupHtml = `
      <table>
        <thead>
          <tr><th>順</th><th>選手</th><th>率</th><th>守</th><th>打</th><th>安</th><th>本</th><th>点</th><th>四</th><th>盗</th><th>失</th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>`;
    const el = $('#lineup-' + side);
    if (el) el.innerHTML = lineupHtml;
    // 投手テーブル: 登板した投手のみを登板順に表示 (初期は先発のみ、交代で順次追加)
    const log = G.pitcherLog[side] || [];
    const setup = G.setup[side] || {};
    const pitchersArr = setup.pitchers || [];
    const roles = setup.pitcherRoles || [];
    const activePitcher = pitchersArr[setup.activeIdx ?? 0];
    const isDefendingTeam = !isBattingTeam;
    const pRows = [];
    for (const myLog of log) {
      const p = myLog.pitcher;
      if (!p) continue;
      // 投球していない投手(0打者0アウト)は表示しない。ただし試合中の現役投手(登板直後)は残す。
      const faced = (myLog.battersFaced || 0) > 0 || (myLog.outs || 0) > 0;
      if (!faced && !(!isBattingTeam && p === activePitcher && !G.ended)) continue;
      const pIdx = pitchersArr.indexOf(p);
      const slotLabel = PITCHER_ROLE_LABELS[roles[pIdx]] || (pIdx === 0 ? '先発' : '救援');
      const isPitching = isDefendingTeam && p === activePitcher && !G.ended;
      // 試合終了後: 勝敗/セーブ/ホールドのワンポイント(W/L/S/H)を役割の前に表示
      const dec = decisionMap ? decisionMap.get(myLog) : null;
      const decBadge = dec ? `<span class="role-badge role-${dec}" title="${({W:'勝利投手',L:'敗戦投手',H:'ホールド投手',S:'セーブ投手'}[dec])}">${dec}</span>` : '';
      pRows.push(`<tr class="${isPitching ? 'is-pitching' : ''}">
        <td class="pt-slot">${decBadge}${slotLabel}</td>
        <td class="pt-name player-link${longNameClass(p.fullNameTop)}" data-player-name="${p.fullNameTop}" data-player-year="${p.year ?? ''}" data-player-type="${playerType(p)}" data-player-team="${p.team||''}">${p.fullNameTop}</td>
        <td>${isSeasonManualGame() ? seasonLiveERA(p) : fmtERA(myLog.earnedRuns, myLog.outs)}</td>
        <td>${fmtIP(myLog.outs)}</td>
        <td>${myLog.hits}</td>
        <td>${myLog.K}</td>
        <td>${myLog.BB || 0}</td>
        <td>${myLog.runsAllowed}</td>
        <td>${myLog.earnedRuns}</td>
      </tr>`);
    }
    const pitcherHtml = `
      <table style="margin-top:6px;">
        <thead>
          <tr><th>枠</th><th>選手</th><th>防御</th><th>回</th><th>被安</th><th>奪三</th><th>四</th><th>失</th><th>自</th></tr>
        </thead>
        <tbody>${pRows.join('')}</tbody>
      </table>`;
    // 交代選手の成績 (代打・代走・守備固めで退いた選手) を投手表示の下に表示
    const subs = G.subLog[side] || [];
    let subHtml = '';
    if (subs.length) {
      const sRows = subs.map(s => `<tr>
        <td class="ln-num">${s.slotIdx + 1}</td>
        <td class="ln-pos">${stintPosLabel(s)}</td>
        <td class="ln-name player-link${longNameClass(s.player.fullNameTop)}" data-player-name="${s.player.fullNameTop}" data-player-year="${s.player.year ?? ''}" data-player-type="${playerType(s.player)}" data-player-team="${s.player.team||''}">${s.player.fullNameTop}</td>
        <td>${fmtAvg(s.H, s.AB)}</td>
        <td>${s.AB}</td>
        <td>${s.H}</td>
        <td>${s.HR}</td>
        <td>${s.RBI}</td>
        <td>${s.BB}</td>
        <td>${s.SB}</td>
        <td>${s.E || 0}</td>
      </tr>`).join('');
      subHtml = `
        <div class="sub-stats-block">
          <h5 class="sub-stats-title">🔁 交代選手</h5>
          <table class="sub-stats-table">
            <thead>
              <tr><th>順</th><th>区分</th><th>選手</th><th>率</th><th>打</th><th>安</th><th>本</th><th>点</th><th>四</th><th>盗</th><th>失</th></tr>
            </thead>
            <tbody>${sRows}</tbody>
          </table>
        </div>`;
    }
    const pe = $('#pitchers-' + side);
    if (pe) pe.innerHTML = pitcherHtml + subHtml;
  }
}

// ホームランサマリーをスコアボード下に表示
function updateHRSummary() {
  const wrap = $('#hr-list');
  if (!wrap) return;
  if (!G.hrEvents || G.hrEvents.length === 0) {
    wrap.innerHTML = '<span style="color:#888;">なし</span>';
    return;
  }
  wrap.innerHTML = G.hrEvents.map(h =>
    `<span class="hr-entry">${h.inning}回${h.top?'表':'裏'} ${h.batter} (${h.runs}点)</span>`
  ).join('');
}

// ============== 打席シミュレーション(1球完結) ==============
// 球種を選択 → 1球で打席結果が決まる(ボール/ストライクカウントなし)
// 守備が下手な選手(マイナスDRS)の エラー / 拙守。
// 対象は野手が捕る打球のアウト (内野ゴロ/外野フライ/ライナー)。三振や犠飛は対象外。守備DRSが0以上の選手は対象外。
// 「本来アウトになる打球がその野手に飛んだ場合」に、|DRS| に比例した確率で:
//   ・エラー (失策で出塁/得点は非自責) = |DRS| × 0.375%  (DRS-8 → 3.0%、上限 12%)
//   ・拙守   (本来アウト→安打扱い/得点は自責) = |DRS| × 0.125%  (DRS-8 → 1.0%、上限 4%)
// 失策＋拙守の合計は |DRS| × 0.5% (DRS-8 → 4.0%、上限 16%) で従来どおり。
//   内訳をエラー寄り(エラー:拙守 = 3:1)に変更 (エラー×1.5 / 拙守×0.5)。
function maybeFieldingError(res) {
  if (!res) return;
  const FIELDED_OUTS = new Set(['GO', 'GO_SLOW', 'FO', 'LO']);
  if (!FIELDED_OUTS.has(res.outcome)) return;
  if (res.fielderIsDH || !res.fielder) return;
  const drs = res.fielderDrs || 0;
  if (drs >= 0) return;                                       // 守備DRSがマイナスの選手のみ
  const ad = Math.abs(drs);
  const errProb    = Math.min(0.12, ad * 0.00375);            // エラー: |DRS|×0.375% (DRS-8 → 3.0%)
  const clumsyProb = Math.min(0.04, ad * 0.00125);            // 拙守:   |DRS|×0.125% (DRS-8 → 1.0%)
  const r = Math.random();
  const name = res.fielder.fullNameTop || res.fielder.nameJa || '守備陣';
  if (r < errProb) {
    // 失策で出塁 (本来のアウト → エラー)。得点は非自責。
    res.errorFrom = res.outcome;
    res.outcome = 'E';
    res.flavor = `${name}の失策で出塁！`;
    res.fineplay = false;
    res.staminaDelta = -1;
  } else if (r < errProb + clumsyProb) {
    // 拙守 (本来のアウト → 安打扱い)。失策は付かず、得点は自責。
    res.clumsyFrom = res.outcome;
    res.clumsy = true;
    res.clumsyFielder = name;
    res.outcome = '1B';
    res.flavor = `${name}の拙守で出塁`;
    res.fineplay = false;
    res.staminaDelta = -1;
  }
  // それ以外は通常アウトのまま
}

// 手動で球種を選んだ際、打席結果に応じた動画(MLB/douga 配下)をゲーム左側に重ねて再生する。
//   ・ヒット/三振/四球/エラー/ファインプレー(前回指定分) → 単独動画(イントロ無し)
//   ・各種アウト(ゴロ/フライ/DP/ライナー/ポップ/ファール/深い飛球 等) → defopit_tou(イントロ)→結果動画
//   動画が無い/読めない環境(プレビューサーバ等)でもゲームは継続する(error/タイムアウトで撤去)。
const PITCH_VIDEOS = ['defopit_tou', 'defopit_tou1', 'defopit_tou2'];
// 投手登場(ピッチャー交代/回頭の継投時)・打者登場(打席ごと)・攻守交替 の演出動画。すべて MLB/douga 配下。
const PITCHER_INTRO_VIDEOS = ['defopit_sta', 'defopit_sta1'];                 // 投手がマウンドに上がる
const BATTER_INTRO_VIDEOS  = ['defobat_box', 'defobat_box1', 'defobat_box2']; // 打者が打席に入る
const SIDE_CHANGE_VIDEO    = 'mlb_change';                                    // 3アウト攻守交替
const STEAL_OK_VIDEOS      = ['defobat_ste', 'defobat_ste1'];                 // 盗塁成功
const STEAL_NG_VIDEOS      = ['defobat_throw', 'defobat_throw1', 'defobat_throw2']; // 盗塁失敗(送球アウト)
const pickVideo = list => list[(Math.random() * list.length) | 0];
// 動画ON/OFF (セットアップのトグルで切替。OFFなら動画を一切再生せず、自動再生バーも出さない)。
let VIDEO_ON = true;
// 動画ファイルの参照先: laa_* は MLB/team_out/laa_out、それ以外(defopit_/defobat_) は MLB/douga。
function videoSrc(file) {
  return (file.indexOf('laa_') === 0 ? '../team_out/laa_out/' : '../douga/') + file + '.mp4';
}
// 打球の実効守備位置。renderBattedBall と同じ depth 補正(res._dispRand 共有)で表示と一致させる。
function effectiveFielderPos(res) {
  const outcome = res.finalOutcome || res.outcome;
  let pos = res.fielderPos || 'P';
  if (res._dispRand == null) res._dispRand = Math.random();
  const rnd = res._dispRand;
  const isIF = p => ['C', 'P', '1B', '2B', '3B', 'SS'].includes(p);
  const isOF = p => ['LF', 'CF', 'RF'].includes(p);
  const fl = res.flavor || '';
  let depth;
  if (res.fineplay || res.clumsy || outcome === 'E') depth = 'KEEP';
  else if (/外野/.test(fl)) depth = 'OF';
  else if (/内野|ファール/.test(fl)) depth = 'IF';
  else if (outcome === 'GO' || outcome === 'GO_SLOW' || outcome === 'GO_DP') depth = 'IF';
  else if (outcome === 'LO') depth = 'LINER';
  else if (outcome === 'FO' && res.infieldFly) depth = 'IF';
  else if (outcome === 'FO') depth = 'KEEP';
  else depth = 'OF';
  if (depth === 'IF' && isOF(pos)) { const m = ({ LF: ['SS', '3B'], CF: ['SS', '2B'], RF: ['2B', '1B'] })[pos] || ['SS']; pos = m[rnd < 0.5 ? 0 : 1]; }
  else if (depth === 'OF' && isIF(pos)) { const m = ({ '3B': ['LF'], 'SS': ['LF', 'CF'], '2B': ['CF', 'RF'], '1B': ['RF'], 'C': ['CF'], 'P': ['CF'] })[pos] || ['CF']; pos = (m.length > 1) ? m[rnd < 0.5 ? 0 : 1] : m[0]; }
  if (outcome === 'LO' && pos === 'C') { const opts = ['SS', '2B', '3B', '1B']; pos = opts[Math.min(opts.length - 1, Math.floor(rnd * opts.length))]; }
  return pos;
}
// 打席結果 → { intro: defopit_touを先に流すか, videos: 候補(ランダム1本) }。対象外は null。
function batResultClip(res, runs) {
  if (!res) return null;
  const oc = res.finalOutcome || res.outcome, fl = res.flavor || '', r2 = b => [b, b + '1'];
  // === 単独動画(イントロ無し) ===
  switch (oc) {
    case 'HR': return { intro: false, videos: ['defobat_hr', 'defobat_hr1'] };               // ホームラン
    case '2B': case '3B': return { intro: false, videos: ['defobat_2b', 'defobat_2b1'] };    // 2/3ベース
    case '1B':
      if (/内野/.test(fl)) return { intro: false, videos: ['defobat_if1b', 'defobat_if1b1'] };  // 内野安打
      if ((runs || 0) > 0) return { intro: false, videos: ['defobat_rbi', 'defobat_rbi1'] };    // タイムリー
      return { intro: false, videos: ['defobat_1b', 'defobat_1b1'] };                            // 通常ヒット
    case 'BB': return { intro: false, videos: ['defobat_fb', 'defobat_fb1'] };               // 四球
    case 'E': return { intro: false, videos: ['defobat_miss', 'defobat_miss1'] };            // エラー
    case 'K': return { intro: false, videos: ['defopit_str', 'defopit_str1', 'defopit_str2'] }; // 三振
  }
  if (res.fineplay) return { intro: false, videos: ['defopit_nice', 'defopit_nice1'] };      // ファインプレーアウト
  // === defopit_tou(イントロ) → 結果動画 (各種アウト) ===
  if (res.forceOut) {   // 走者フォースアウト (2=二塁/3=三塁/4=本塁)
    const fm = { 2: ['laa_forth2_1', 'laa_forth2_2', 'laa_forth2_3'], 3: ['laa_forth3_1', 'laa_forth3_2', 'laa_forth3_3'], 4: ['laa_forth4_1', 'laa_forth4_2', 'laa_forth4_3'] };
    if (fm[res.forceOut]) return { intro: true, videos: fm[res.forceOut] };
  }
  if (oc === 'GO_DP') return { intro: true, videos: ['laa_double', 'laa_double1'] };                       // ダブルプレー
  if (oc === 'SAC_FLY' || /大きな|大飛球/.test(fl)) return { intro: true, videos: ['laa_flybig', 'laa_flybig1', 'laa_flybig2'] }; // 深い飛球
  const pos = effectiveFielderPos(res);
  if (oc === 'FO' && (res.infieldFly || /内野|ファール/.test(fl))) {
    if (/ファール/.test(fl)) return { intro: true, videos: ['laa_foul', 'laa_foul1', 'laa_foul2'] };       // ファールフライ
    return { intro: true, videos: ({ '1B': ['laa_pop'], '2B': ['laa_pop1'], '3B': ['laa_pop2'] })[pos] || ['laa_pop1'] }; // 内野ポップ
  }
  if (oc === 'FO') {   // 外野フライ → 守備位置別 (左7/中8/右9)
    return { intro: true, videos: r2(({ 'LF': 'laa_7out', 'CF': 'laa_8out', 'RF': 'laa_9out' })[pos] || 'laa_8out') };
  }
  if (oc === 'LO') return { intro: true, videos: ['laa_line', 'laa_line1', 'laa_line2', 'laa_line3'] };    // 内野ライナー
  if (oc === 'GO_SLOW' || /ボテボテ|力ない/.test(fl)) {   // 緩いゴロ
    return { intro: true, videos: [({ '1B': 'laa_soft', '3B': 'laa_soft1', 'SS': 'laa_soft2', '2B': 'laa_soft3' })[pos] || 'laa_soft'] };
  }
  if (/詰まった/.test(fl)) {   // 強い(詰まった)ゴロ
    return { intro: true, videos: [({ '1B': 'laa_strong', '2B': 'laa_strong1', '3B': 'laa_strong2', 'SS': 'laa_strong3' })[pos] || 'laa_strong'] };
  }
  if (oc === 'GO') {   // 通常の内野ゴロ → 守備位置別 (一3/二4/三5/遊6)
    return { intro: true, videos: r2(({ '1B': 'laa_3out', '2B': 'laa_4out', '3B': 'laa_5out', 'SS': 'laa_6out' })[pos] || 'laa_6out') };
  }
  return null;
}
// 任意の動画ファイル列(seq)をゲーム左側に重ねて順番に再生する汎用関数。
//   1本目はユーザー操作直後なら音あり、2本目以降は自動再生制限でミュートにフォールバックする。
function playVideoOverlay(seq, onDone) {
  try {
    if (!VIDEO_ON) return;     // 動画OFF時は一切再生しない
    if (!seq || !seq.length) return;
    const game = document.querySelector('#game');
    if (!game || game.classList.contains('hidden')) return;   // ゲーム画面表示中のみ
    const anchor = document.querySelector('.game-left') || document.querySelector('.game-main') || game;
    const prev = document.getElementById('pitchVideoOverlay');
    if (prev) prev.remove();
    const ov = document.createElement('div');
    ov.id = 'pitchVideoOverlay';
    const v = document.createElement('video');
    v.autoplay = true; v.playsInline = true; v.setAttribute('playsinline', '');
    ov.appendChild(v);
    document.body.appendChild(ov);
    G._videoActive = true;     // 自動再生モードの「再生中か」判定に使用
    // 動画の表示位置はゲーム中で固定する。最初に動画を出した時の .game-left の位置を記録し、
    //   以降は投手/野手交代で .game-left の高さが変わっても同じ位置・サイズで再生する(ずれ防止)。
    const freeze = () => { const r = anchor.getBoundingClientRect(); G._videoRect = { left: r.left, top: r.top, width: r.width, height: r.height }; };
    if (!G._videoRect) freeze();
    const place = () => { const r = G._videoRect; if (!r) return; ov.style.left = r.left + 'px'; ov.style.top = r.top + 'px'; ov.style.width = r.width + 'px'; ov.style.height = r.height + 'px'; };
    place();
    const onResize = () => { freeze(); place(); };   // ウィンドウリサイズ時のみ位置を取り直す
    window.addEventListener('resize', onResize);
    let removed = false, safety = null;
    const cleanup = () => { if (removed) return; removed = true; window.removeEventListener('resize', onResize); if (safety) clearTimeout(safety); ov.remove(); G._videoActive = false; if (onDone) onDone(); else onVideoSequenceComplete(); };
    // 1本のクリップを再生し、終了/失敗/保険タイマーのいずれかで onClipEnd を一度だけ呼ぶ
    const playClip = (file, onClipEnd) => {
      if (safety) clearTimeout(safety);
      let fired = false;
      const fire = () => { if (fired) return; fired = true; if (safety) clearTimeout(safety); onClipEnd(); };
      v.onended = fire; v.onerror = fire;
      v.src = videoSrc(file);   // laa_* は team_out/laa_out、それ以外は douga
      v.load(); place();
      safety = setTimeout(fire, 20000);
      // 連続再生(2本目)はユーザー操作外のため、音ありだと自動再生がブロックされ固まる。
      //   ブロックされたらミュートで再生継続し、それでも不可なら次へ進める(固まらせない)。
      const tryPlay = () => { const pr = v.play(); if (pr && pr.catch) pr.catch(() => { if (!v.muted) { v.muted = true; tryPlay(); } else fire(); }); };
      tryPlay();
    };
    let i = 0;
    const next = () => { if (removed) return; if (i >= seq.length) { cleanup(); return; } playClip(seq[i++], next); };
    next();
  } catch (e) { /* 動画再生は補助機能。失敗してもゲームは継続 */ }
}
// 打席結果後、次打者までに挟む繋ぎ動画の種別列を順番に返す純粋関数。
//   before = 投球前 { top, inning, pitcher } / cur = 投球後 { top, inning, pitcher, ended } / lastIntro = {home,away}
//   返り値: 'change'(攻守交替) → 'pitcher'(投手登場) → 'batter'(打者登場) の順序付き配列。
//   ・3アウトで半回が替わった時のみ 'change'。
//   ・守備側の投手が前回登板から替わった(初登板含む)時のみ 'pitcher'(攻守交替時に限る)。
//   ・打者は毎打席替わるので常に 'batter'。
function batTransitionKinds(before, cur, lastIntro) {
  const kinds = [];
  if (!before || !cur || cur.ended) return kinds;
  const sideChanged = (before.top !== cur.top) || (before.inning !== cur.inning);
  if (sideChanged) kinds.push('change');
  const defSide = cur.top ? 'home' : 'away';
  if (sideChanged && cur.pitcher && lastIntro && lastIntro[defSide] !== cur.pitcher) kinds.push('pitcher');
  kinds.push('batter');
  return kinds;
}
// 手動の球種選択後に呼ぶ。打席結果動画 → (攻守交替) → (投手登場) → 打者登場 をまとめて再生する。
//   before = 投球前の { top, inning, pitcher } 。投球後の状態と比較して攻守交替/継投を検知する。
function playPitchVideo(before, opts) {
  try {
    const game = document.querySelector('#game');
    if (!game || game.classList.contains('hidden')) return;   // ゲーム画面表示中のみ
    const seq = [];
    // 1) 打席結果の動画 (対象外なら無し)。noResult=盗塁死3アウト等で投球が無かった場合はスキップ。
    if (!(opts && opts.noResult)) {
      const clip = batResultClip(G.lastPitchResult, G.lastPitchRuns);
      if (clip && clip.videos && clip.videos.length) {
        if (clip.intro) seq.push(pickVideo(PITCH_VIDEOS));   // アウトは投球イントロ→結果
        seq.push(pickVideo(clip.videos));
      }
    }
    // 2) 次打者への繋ぎ (試合継続中のみ)
    if (before && !G.ended) {
      if (!G._lastIntroPitcher) G._lastIntroPitcher = { home: null, away: null };
      const cur = { top: G.top, inning: G.inning, pitcher: G.currentPitcher, ended: G.ended };
      const defSide = G.top ? 'home' : 'away';
      for (const kind of batTransitionKinds(before, cur, G._lastIntroPitcher)) {
        if (kind === 'change') seq.push(SIDE_CHANGE_VIDEO);                       // 3アウト攻守交替
        else if (kind === 'pitcher') { seq.push(pickVideo(PITCHER_INTRO_VIDEOS)); G._lastIntroPitcher[defSide] = G.currentPitcher; } // 投手登場
        else seq.push(pickVideo(BATTER_INTRO_VIDEOS));                            // 次打者が打席へ
      }
    }
    playVideoOverlay(seq);
  } catch (e) { /* 動画再生は補助機能。失敗してもゲームは継続 */ }
}

// ============== 自動再生モード (動画ON時のみ) ==============
//   「動画再生中 ▶」= AIが自動進行。動画が終わる → 3秒カウントダウン → 0でAIが投球(投球確率で球種選択) → 動画。
//   「動画停止中 ■」= AI停止。手動で球種ボタンを押して進められる。
let autoVideoCountdownTimer = null;
// 自動再生バーを出す条件: 試合画面・動画ON・試合継続中
function autoVideoBarVisible() {
  if (G.silent || !VIDEO_ON || G.ended) return false;
  const game = document.querySelector('#game');
  return !!(game && !game.classList.contains('hidden'));
}
function clearAutoVideoCountdown() {
  if (autoVideoCountdownTimer) { clearInterval(autoVideoCountdownTimer); autoVideoCountdownTimer = null; }
  const el = document.querySelector('#autoVideoCountdown');
  if (el) el.textContent = '';
}
// 自動再生を止める (動画OFF/試合終了/停止ボタン/画面離脱)。
function stopAutoVideo() {
  G.autoVideoPlaying = false;
  clearAutoVideoCountdown();
}
// バーの表示/ラベル/色を現在状態に同期。表示条件を満たさなければ隠して自動再生も止める。
function updateAutoVideoBar() {
  const bar = document.querySelector('#autoVideoBar');
  if (!bar) return;
  if (!autoVideoBarVisible()) { stopAutoVideo(); bar.hidden = true; clearAutoVideoCountdown(); return; }
  bar.hidden = false;
  const btn = document.querySelector('#autoVideoBtn');
  if (btn) {
    btn.textContent = G.autoVideoPlaying ? '動画再生中 ▶' : '動画停止中 ■';
    btn.classList.toggle('is-playing', !!G.autoVideoPlaying);
  }
}
// 3秒カウントダウン → 0でAI投球。
function startAutoVideoCountdown() {
  clearAutoVideoCountdown();
  if (!G.autoVideoPlaying || G.ended) return;
  let n = 3;
  const el = document.querySelector('#autoVideoCountdown');
  if (el) el.textContent = String(n);
  autoVideoCountdownTimer = setInterval(() => {
    if (!G.autoVideoPlaying || G.ended) { clearAutoVideoCountdown(); return; }
    n--;
    if (el) el.textContent = (n > 0) ? String(n) : '';
    if (n <= 0) { clearAutoVideoCountdown(); autoVideoThrowNext(); }
  }, 1000);
}
// AIが投球確率で球種を選び投球 → 結果動画を再生 (終了時 onVideoSequenceComplete が次のカウントダウンを開始)。
function autoVideoThrowNext() {
  if (!G.autoVideoPlaying || G.ended) return;
  const p = autoPick();
  if (!p) return;
  const before = { top: G.top, inning: G.inning, pitcher: G.currentPitcher };
  pitchOne(p, true);
  const events = G._prePitchEvents || [];
  G._prePitchEvents = [];
  if (events.length && VIDEO_ON) {
    // 投球前イベント(代打/盗塁)がある: 各イベントのダイヤ表示+動画を順番に流し、
    //   全て終わってから投球結果を表示して投球結果動画へ繋ぐ。
    const noResult = events.some(e => e.type === 'steal' && e.thirdOut);  // 盗塁死3アウト=投球なし
    playPrePitchEvents(events, 0, before, noResult);
    return;
  }
  playPitchVideo(before);
  // 動画が再生されなかった場合(結果動画なし等)でも、終了コールバックが来ないのでループを継続させる
  if (!G._videoActive && G.autoVideoPlaying && !G.ended) startAutoVideoCountdown();
}
// 投球前イベント列を1つずつ (フレーム表示→動画) 再生し、最後に投球結果へ繋ぐ再帰関数。
function playPrePitchEvents(events, i, before, noResult) {
  if (i >= events.length) {
    // 全ての事前演出が完了 → 投球結果を表示して投球結果動画へ
    renderAll();
    playPitchVideo(before, { noResult });
    if (!G._videoActive && G.autoVideoPlaying && !G.ended) startAutoVideoCountdown();
    return;
  }
  const ev = events[i];
  const next = () => playPrePitchEvents(events, i + 1, before, noResult);
  if (ev.type === 'steal') {
    renderStealFrame(ev);
    playVideoOverlay([pickVideo(ev.success ? STEAL_OK_VIDEOS : STEAL_NG_VIDEOS)], next);
  } else if (ev.type === 'pinch') {
    renderPinchFrame(ev);
    playVideoOverlay([pickVideo(BATTER_INTRO_VIDEOS)], next);   // 代打も打者登場動画(defobat_box種)
  } else {
    next();
  }
}
// 盗塁の演出フレーム: ダイヤに盗塁直後(投球前)の状態を表示する (投球結果の打球は出さない)。
function renderStealFrame(steal) {
  if (!steal) return;
  renderBbState(steal.inning, steal.top, steal.outs);
  renderRunnerBadges(steal.bases);
  const overlay = document.querySelector('#bbOverlay');
  if (overlay) overlay.innerHTML = '';   // 打球軌道は描かない
  const label = document.querySelector('#bbLabel');
  if (label) {
    label.className = 'bb-label ' + (steal.success ? 'hit' : 'out');
    label.innerHTML = `<span class="bb-result">${steal.success ? '盗塁成功！' : '盗塁失敗！'}</span>`;
  }
  renderNextBatterLine();     // 実況ラベル3行目: 次打者(NEXT)
  renderBbInfo(steal.info);   // 🏃💨/🏃❌ の通知
}
// 代打の演出フレーム: ダイヤに「代打 選手名」を表示する (投球結果の打球は出さない)。
function renderPinchFrame(ev) {
  if (!ev) return;
  renderBbState(ev.inning, ev.top, ev.outs);
  renderRunnerBadges(ev.bases);
  const overlay = document.querySelector('#bbOverlay');
  if (overlay) overlay.innerHTML = '';   // 打球軌道は描かない
  const label = document.querySelector('#bbLabel');
  if (label) {
    label.className = 'bb-label bb';
    label.innerHTML = `<span class="bb-result">代打 ${ev.batterName}</span>`;
  }
  renderNextBatterLine();   // 実況ラベル3行目: 次打者(NEXT)
  renderBbInfo(ev.info);    // 🔁 代打: … の通知
}
// 動画(連続再生)が1区切り終わった時に呼ばれる。自動再生中なら次の投球までのカウントダウンを開始。
function onVideoSequenceComplete() {
  if (G.autoVideoPlaying && !G.ended) startAutoVideoCountdown();
}
// セットアップの動画ON/OFFボタンのラベル・色を同期 (ON=赤 / OFF=青)。
function updateVideoToggleBtn() {
  const btn = document.querySelector('#videoToggle');
  if (!btn) return;
  btn.textContent = VIDEO_ON ? '動画ON' : '動画OFF';
  btn.classList.toggle('video-on', VIDEO_ON);
  btn.classList.toggle('video-off', !VIDEO_ON);
}
// 「動画再生中/停止中」トグル。
function toggleAutoVideo() {
  if (!autoVideoBarVisible()) return;
  G.autoVideoPlaying = !G.autoVideoPlaying;
  updateAutoVideoBar();
  if (G.autoVideoPlaying) {
    // 再生開始。動画再生中ならその終了時に自動でカウントダウンへ。何も流れていなければ即カウントダウン。
    if (!G._videoActive) startAutoVideoCountdown();
  } else {
    clearAutoVideoCountdown();   // 停止 → AI投球を止め、手動操作に戻す
  }
}

function pitchOne(pitch, isAuto) {
  if (G.ended) return;
  // 打席開始時: 前打席で表示済みの外野通知(infoNew=false)とプレイボールをクリア。
  //   各通知は「その打席限り(1ターン)」で消える。今打席で出る通知(代打/交代/守備固め等)はこの後に積む。
  if (G.lastInfo && (!G.infoNew || G.lastInfo.indexOf('プレイボール') >= 0)) {
    G.lastInfo = '';
  }
  G._prePitchEvents = [];   // 今回の投球前イベント(代打/盗塁)。動画ONの自動再生で投球前演出に使用。
  // ある時点(投球前)の状態スナップショットを作る (動画演出でダイヤに再現するため)
  const snapPrePitch = () => ({
    inning: G.inning, top: G.top, outs: G.outs,
    bases: G.bases.map(b => b ? { ...b, _player: getRunnerPlayer(b) } : null),
    info: G.lastInfo || '',
  });
  // 打席前に野手交代 (代走→代打) を検討 (守備成立を保証した上で起用)
  // 自動試合(isAuto)時のみ AI が自動で代打/代走を実行する。手動操作時はオフ。
  if (isAuto) {
    const battingSide = G.top ? 'away' : 'home';
    maybePinchRun(battingSide);
    const batterBefore = G.currentBatter;
    maybePinchHit(battingSide);
    if (G.currentBatter !== batterBefore) {
      // 代打が送られた: この時点(投球前)の状態を控える。動画演出でダイヤに「代打」を表示。
      G._prePitchEvents.push(Object.assign({
        type: 'pinch',
        batterName: G.currentBatter ? (G.currentBatter.fullNameTop || G.currentBatter.nameJa || '代打') : '代打',
      }, snapPrePitch()));
    }
    // 自動盗塁: 投手VS打者結果の直前に判定。
    // 盗塁失敗で3アウトチェンジになった場合、この打席(投手VS打者)はノーカウント
    // (ヒット/アウト/スタミナ減少などを一切記録せず、そのままイニング交代)。
    const steal = maybeAutoSteal(battingSide);
    if (steal && steal.attempted) {
      G._prePitchEvents.push(Object.assign({
        type: 'steal', success: steal.success, thirdOut: steal.thirdOut,
      }, snapPrePitch()));
    }
    if (steal && steal.thirdOut) {
      switchInning();
      checkEnd();
      renderAll();
      G.infoNew = false;   // 表示済み → 次打席でクリア
      return;
    }
  }
  const P = G.currentPitcher, B = G.currentBatter;
  const res = decidePitchOutcome(pitch, P, B, isAuto);
  maybeFieldingError(res);  // 守備DRSがマイナスの野手: アウト→失策(出塁) を一定確率で
  G.lastPitchResult = res;  // applyOutcome / formatPlayByPlay から参照
  G.lastPitchRuns = 0;      // applyOutcome がこの打席の失点を上書き
  applyOutcome(res.outcome, pitch);
  // スタミナ調整: 球種・役別の基礎消費 + 失点 × 2 (ファインプレー時は +2 回復)
  applyPitchStaminaDelta(res, pitch);
  applyStarterStaminaBonus();   // 低スタミナ先発の好投ボーナス (継投判断の前に反映して延命させる)
  checkRelief();
  // この打席で出た全通知(継投/代打/守備固め等、switchInning/checkRelief後に確定)を履歴へ保存。
  if (G.pitchHistory && G.pitchHistory.length) {
    G.pitchHistory[G.pitchHistory.length - 1].info = G.lastInfo || '';
  }
  renderAll();
  G.infoNew = false;   // 表示済み → 次打席開始時にクリア(その打者限り1ターン)
}

// =============================================================
// 三振率モデル (通常試合・レギュラーシーズン共通)
//   アウトになる打席のうち何割を三振にするかを、投手「奪三振」と打者「三振耐性」から算出する。
//   ・投手「奪三振」 → アウトに占める三振割合 (アンカー: 50→25%, 75→35%, 85→40%)
//   ・打者「三振耐性」 → 同 (打数あたり目標 40→32%,60→23%,80→15% を、エンジン平均打率
//       ≒.24(=アウト/打数≒.75) で割り戻した値: 40→.428, 60→.308, 80→.200)
//   ・両者を Log5(オッズ合成) でリーグ平均 L=.3075 を基準に合成。
//     → 平均打者と対戦する投手は投手値、平均投手と対戦する打者は打者値 に概ね一致する。
// =============================================================
function pitcherKoutRate(k) {
  k = (k == null) ? 60 : k;
  if (k <= 50) return Math.max(0.08, 0.25 + (k - 50) * 0.004);
  if (k <= 75) return 0.25 + (k - 50) * (0.10 / 25);
  if (k <= 85) return 0.35 + (k - 75) * (0.05 / 10);
  return Math.min(0.60, 0.40 + (k - 85) * 0.005);
}
function batterKoutRate(t) {
  t = (t == null) ? 60 : t;
  const v = (t <= 60) ? (0.4278 + (t - 40) * ((0.3075 - 0.4278) / 20))
                      : (0.3075 + (t - 60) * ((0.2005 - 0.3075) / 20));
  return Math.max(0.05, Math.min(0.70, v));
}
function strikeoutOutRate(kAb, kTol) {
  const L = 0.3075;
  const a = pitcherKoutRate(kAb), b = batterKoutRate(kTol);
  const num = (a * b) / L, den = num + ((1 - a) * (1 - b)) / (1 - L);
  const p = den > 0 ? num / den : 0;
  return Math.max(0, Math.min(0.95, p));
}

// =============================================================
// 四球率モデル (通常試合・レギュラーシーズン共通)
//   打者「選球眼」を主因に、1打席あたりの四球確率を返す。
//   ・選球眼 → 四球/打数の目標 50→5%, 70→11%, 90→18%, 100→23% を
//     1打席確率 w = 目標/(1+目標) に変換した値で区分線形補間
//     (50→.0476, 70→.0991, 90→.1525, 100→.1870)。
//   ・投手「制球」が高いほど四球減 / 低いほど増 (制球60を基準)。
//   ・投手スタミナ切れ(suta0>1)でやや増。
// =============================================================
function batterWalkPA(eye) {
  eye = (eye == null) ? 60 : eye;
  let w;
  if (eye <= 70)      w = 0.04762 + (eye - 50) * 0.002574;
  else if (eye <= 90) w = 0.09910 + (eye - 70) * 0.002672;
  else                w = 0.15254 + (eye - 90) * 0.003445;
  return Math.max(0.01, Math.min(0.30, w));
}
function walkRatePA(eye, ctrl, suta0) {
  const cm = Math.max(0.4, Math.min(1.7, 1 + (60 - (ctrl == null ? 60 : ctrl)) * 0.01));   // 制球補正
  const sm = Math.min(1.8, Math.max(1, 1 + ((suta0 || 1) - 1) * 0.4));                       // スタミナ補正
  return Math.max(0, Math.min(0.6, batterWalkPA(eye) * cm * sm));
}

// =============================================================
// 投球結果の判定 (詳細スペック準拠)
// 返り値: { outcome, flavor, fielder, kire, displaySpeed, staminaDelta }
// =============================================================
function decidePitchOutcome(pitch, P, B, isAuto) {
  const ps = P.stats || {};
  const bs = B.stats || {};
  const bm = B.statsMini || {};

  // === 0. 捕手情報 (リード / 阻止率) ===
  const defSide = G.top ? 'home' : 'away';
  const setup = G.setup[defSide];
  const catIdx = setup.batterPos.findIndex(p => p === 'C');
  const catcher = catIdx >= 0 ? setup.batters[catIdx] : null;
  // 捕手「リード」は実際の値域が ±4〜+8 程度の小さな修正値。
  // データが無い場合は 0 (= 補正なし) にフォールバック。
  const rido = (catcher && catcher.catcher && catcher.catcher['リード'] != null)
    ? catcher.catcher['リード']
    : 0;
  // 阻止率 は 25-55 程度の値。データが無ければ平均的な 30 をフォールバック。
  const arm = (catcher && catcher.catcher && catcher.catcher['阻止率'] != null)
    ? catcher.catcher['阻止率']
    : 30;

  // === 1. スタミナ係数 ===
  const info = getActivePitcherInfo();
  const stamina = info.stamina;
  let suta0;
  if (stamina > 0)          suta0 = 1;
  else if (stamina >= -6)   suta0 = 1.5;
  else if (stamina >= -13)  suta0 = 2;
  else if (stamina >= -20)  suta0 = 4;
  else                       suta0 = 8;

  // === 2. dousyu (同投手・同イニング・同球種の連投ペナルティ) ===
  // 自動試合(isAuto)では適用しない。手動試合のみ、同じ投手・同じ回・同じ球種を連投すると蓄積。
  let dousyu = 0;
  if (isAuto) {
    G.dousyuTracker = null;   // 自動試合は連投ペナルティなし
  } else {
    const dk = G.dousyuTracker;
    if (!dk || dk.pitcherKey !== P.fullNameTop || dk.inning !== G.inning || dk.top !== G.top || dk.pitchName !== pitch.name) {
      G.dousyuTracker = { pitcherKey: P.fullNameTop, inning: G.inning, top: G.top, pitchName: pitch.name, count: 0 };
    } else {
      G.dousyuTracker.count += 10;
    }
    dousyu = G.dousyuTracker.count;
  }

  // === 3. 乱数群 ===
  const rand0 = Math.floor(Math.random()*60);
  // rand: 投手のキレを表す主乱数。仕様: rand = Math.floor(Math.random()*7/suta0)
  // 制球・緩急で再代入され、最終値が hit_Kyui / 強制HR(artisut) / 表示キレ で参照される
  let rand = Math.floor(Math.random()*7/suta0);
  const rand2 = Math.floor(Math.random()*283);
  const rand3 = Math.floor(Math.random()*200);
  const ctrlStat = ps['制球'] || 60;
  const rand4Max = Math.max(1, Math.floor((ctrlStat*18/Math.max(1,101-ctrlStat))/(suta0*suta0)));
  const rand4 = Math.floor(Math.random()*rand4Max);
  const rand5 = Math.floor(Math.random()*300);
  const rand6 = Math.floor(Math.random()*50);
  const rand7 = Math.floor(Math.random()*150);
  const rand8 = Math.floor(Math.random()*370);
  const rand9 = Math.floor(Math.random()*30);
  const rand10 = Math.floor(Math.random()*42);
  const rand11 = Math.floor(Math.random()*9);
  const rand12 = Math.floor(Math.random()*250);
  const rand13 = Math.floor(Math.random()*1100);
  const rand16 = Math.floor(Math.random()*19);
  const rand23 = Math.floor(Math.random()*400);
  const rand30 = Math.floor(Math.random()*100);

  // === 4. キレ調整 (制球・緩急) ===
  if (rand !== 6) {
    if (ctrlStat - 70 - rand6 > 0)         rand = rand + 1;
    else if ((ps['緩急']||60) - 70 - rand7 > 0) rand = 6;
    else if (rand !== 0 && (rand7 - 60 + ctrlStat) < 0) rand = rand - 1;
  }
  rand = Math.max(0, Math.min(6, rand));
  const kire = rand;  // 表示用エイリアス (displaySpeed / 実況表示の「キレN」で使用)

  // === 5. 球種・対球種ポイント ===
  const henkataiou = (B.pitchPoints || {})[fullPitchToShort(pitch.name)] || 0;
  const isLeftyP = (P.hand || '').includes('左投');
  const isLeftyB = (B.hand || '').includes('左打');
  const bat_taiHidari = isLeftyP ? (bm['対左投手'] ?? bm['対左'] ?? 0) : 0;
  const pit_taiHidari = isLeftyB ? (bm['対左'] ?? 0) : 0;

  // === 6. ピンチ・チャンス補正 ===
  let bat_tyan = 0, pit_pin = 0;
  if (G.bases[1] || G.bases[2]) {
    bat_tyan = (bs['チャンス']||60) - 70;
    pit_pin  = (ps['精神']||60) - 70;
  }

  // === 7. 投球威力 ===
  const pitchPower = pitch.power || 70;
  const adjPow = pitchPower - dousyu;

  // === 8. ヒット判定値 ===
  const hit_Kyui = Math.floor(adjPow/2 + (adjPow*rand/15) + (10*rand/4) - 55 - (dousyu/2));
  const rand_Mi = Math.floor((bs['ミート']||60) - hit_Kyui - ((ps['緩急']||60)/5 - 10));
  let hitPt = rand_Mi - rand2 + bat_tyan - pit_pin + bat_taiHidari - pit_taiHidari + (henkataiou*0.5) - rido;

  // === 9. 長打判定値 ===
  const rand_Pa = Math.floor((bs['パワー']||60) - 40 - (hit_Kyui/2) - (ctrlStat/12.5 - 4) - (((ps['球重']||60)-73)*0.6));
  const powerPt = rand_Pa - rand3 + (henkataiou*0.5);

  // === 10. 内野安打判定 ===
  const speedB = bs['スピード'] || 60;
  const naiAn = Math.floor((speedB/10)*(speedB/10)) - (rand0*30);

  // === 11. 奪三振 (アウトに占める三振割合は strikeoutOutRate で確率判定) ===
  const kAb = ps['奪三振'] || 60;
  const kTol = bs['三振耐性'] || 60;

  // === 12. 四球 ===
  const sisikyu2 = ((bs['選球眼']||60) - 25) - rand13;

  // === 13. アーティスト (HR能由来) ===
  const arthiSyuu  = ((ps['球重']||60) - 70 - rand10) > 0 ? 1 : 0;
  const arthiSyuu1 = (Math.floor(ctrlStat * (ps['精神']||60) / 800) - 4 - rand11) > 0 ? 1 : 0;
  const artisut = (bm['HR能']||0) - rand9 - arthiSyuu1 - arthiSyuu;

  // === 14. 救済アウト (低球威の救済) ===
  if (pitchPower < 79) {
    const r24max = Math.max(1, Math.floor(1000/(80-pitchPower)));
    const rand24 = Math.floor(Math.random()*r24max);
    if (rand24 < 5) hitPt = -35;
  }
  // 三振耐性救済 (極端低hitPtで打者の三振耐性なら -169 に押し戻し)
  if (hitPt < -170) {
    const doctorK = kAb > 71 ? (kAb/4 - 18) : 0;
    if (((kTol/5) - (rand11+14) - doctorK) > 0) hitPt = -169;
  }

  // === 15. 守備割り当て (打球の守備位置担当者を決定) ===
  let posKey;
  if      (rand30 < 18) posKey = 'SS';
  else if (rand30 < 34) posKey = '2B';
  else if (rand30 < 49) posKey = 'CF';
  else if (rand30 < 61) posKey = '3B';
  else if (rand30 < 73) posKey = 'RF';
  else if (rand30 < 83) posKey = 'LF';
  else if (rand30 < 92) posKey = '1B';
  else                  posKey = 'C';
  const defPosIdx = setup.batterPos.findIndex(p => p === posKey);
  const fielder = defPosIdx >= 0 ? setup.batters[defPosIdx] : null;
  const fielderIsDH = fielder && fielder.position === '指名打者';
  const drsEntry = fielder && fielder.drs && fielder.drs.find(d => d.pos === posKey);
  const fp_pt = drsEntry ? (drsEntry.value || 0) : 0;
  const syubiPt01 = rand12 + fp_pt * 2.5;

  // 表示用 球速
  const displaySpeed = (pitch.speed || 140) - 8 + Math.floor(rand * 4 / 3);
  const baseRes = { kire, displaySpeed, fielder, fielderIsDH, fielderDrs: fp_pt, fielderPos: posKey };

  // === 16. 結果判定 ===
  // (a) フォアボール: 打者「選球眼」(主) ＋ 投手「制球」・スタミナ から算出した確率で四球
  if (Math.random() < walkRatePA(bs['選球眼'] || 60, ctrlStat, suta0)) {
    return { ...baseRes, outcome: 'BB', flavor: 'フォアボール！', staminaDelta: -1 };
  }

  // (c) 内野安打ルート
  if (naiAn > 0) {
    if (syubiPt01 > 245 && Math.random() < 0.6) {   // ファインプレーは発生確率を3/5に抑制
      // ファインプレー (バッターアウト)
      const name = fielderIsDH ? P.fullNameTop : (fielder ? fielder.fullNameTop : '守備陣');
      return { ...baseRes, outcome: 'GO', flavor: `${name}のファインプレー！アウト`, fineplay: true, staminaDelta: +2 };
    }
    // 通常の内野安打 (拙守による安打の判定はここでは行わない — 失策は maybeFieldingError で別処理)
    // 打者のスピードが60以上なら「俊足を生かして」、59以下なら「ボテボテのあたりで」内野安打
    const naiFlavor = (speedB >= 60) ? '俊足を生かして内野安打！' : 'ボテボテのあたりで内野安打！';
    return { ...baseRes, outcome: '1B', flavor: naiFlavor, staminaDelta: -1 };
  }

  // (d) ホームラン (アーティスト or 強制HR)
  const power = bs['パワー'] || 60;
  const forcedHR = (
    (rand <= 1 && artisut > 0) ||
    (power > 54 && rand23 > 398 && pitchPower > 79) ||
    (power > 64 && rand23 > 397 && pitchPower > 79) ||
    (power > 74 && rand23 > 396 && pitchPower > 79) ||
    (power > 84 && rand23 > 395 && pitchPower > 79)
  );
  if (forcedHR) {
    return { ...baseRes, outcome: 'HR', flavor: '驚愕の一発！ホームラン！？', staminaDelta: -3 };
  }

  // (e) hitPt >= 0 ヒット枠
  if (hitPt >= 0) {
    if (syubiPt01 > 245 && Math.random() < 0.6) {   // ファインプレーは発生確率を3/5に抑制
      const name = fielderIsDH ? P.fullNameTop : (fielder ? fielder.fullNameTop : '守備陣');
      return { ...baseRes, outcome: 'FO', flavor: `${name}のファインプレー！アウト`, fineplay: true, staminaDelta: +2 };
    }
    if (powerPt > 0) {
      return { ...baseRes, outcome: 'HR', flavor: 'ホームラン！！', staminaDelta: -3 };
    }
    if (powerPt > -1) {
      return { ...baseRes, outcome: '3B', flavor: 'スリーベースヒット！', staminaDelta: -2 };
    }
    const sou4Pt = speedB - 40 - rand0;
    if (powerPt > -7 && sou4Pt > 0) {
      return { ...baseRes, outcome: '3B', flavor: '俊足を生かしてスリーベースヒット！', staminaDelta: -2 };
    }
    if (powerPt > -80) {
      return { ...baseRes, outcome: '2B', flavor: 'ツーベースヒット！', staminaDelta: -2 };
    }
    return { ...baseRes, outcome: '1B', flavor: 'ヒット！', staminaDelta: -1 };
  }

  // (f) hitPt < 0
  if (rand16 === 7) {
    return { ...baseRes, outcome: '1B', flavor: 'ヒット？！', staminaDelta: -0.5 };
  }
  // 三振判定: このアウト打席を、奪三振(投手)×三振耐性(打者)で決まる割合だけ三振にする
  if (Math.random() < strikeoutOutRate(kAb, kTol)) {
    return { ...baseRes, outcome: 'K', flavor: pickRand(['空振り三振！', '三振！', '見逃し三振']), staminaDelta: -1 };
  }
  // 各 hitPt × powerPt 帯による振り分け (非三振アウト)
  if (hitPt >= -35 && powerPt > -50) {
    return { ...baseRes, outcome: 'SAC_FLY', flavor: '大きな外野フライ', staminaDelta: -1 };
  }
  if (hitPt >= -35 && powerPt <= -100) {
    return { ...baseRes, outcome: 'FO', flavor: '外野フライアウト', staminaDelta: -1 };
  }
  if (hitPt >= -60 && powerPt > -20) {
    return { ...baseRes, outcome: 'FO', flavor: '浅い外野フライ', staminaDelta: -1 };
  }
  if (hitPt >= -60 && powerPt > -100) {
    return { ...baseRes, outcome: 'LO', flavor: 'ライナーアウト', staminaDelta: -1 };
  }
  if (hitPt >= -60 && powerPt <= -100) {
    return { ...baseRes, outcome: 'GO_SLOW', flavor: '内野ボテボテゴロ', staminaDelta: -1 };
  }
  if (hitPt >= -115 && powerPt > -20) {
    return { ...baseRes, outcome: 'LO', flavor: 'ライナーアウト', staminaDelta: -1 };
  }
  if (hitPt >= -115 && powerPt > -100) {
    return { ...baseRes, outcome: 'GO', flavor: '内野ゴロ', staminaDelta: -1 };
  }
  if (hitPt >= -170 && powerPt > -20) {
    return { ...baseRes, outcome: 'FO', flavor: '内野フライ', infieldFly: true, staminaDelta: -1 };
  }
  if (hitPt >= -170 && powerPt > -100) {
    return { ...baseRes, outcome: 'GO_SLOW', flavor: '内野ボテボテゴロ', staminaDelta: -1 };
  }
  if (hitPt >= -170 && powerPt <= -100) {
    return { ...baseRes, outcome: 'GO', flavor: '内野ゴロ', staminaDelta: -1 };
  }
  if (hitPt >= -220 && powerPt > -85) {
    return { ...baseRes, outcome: 'FO', flavor: 'ファールフライ', infieldFly: true, staminaDelta: -1 };
  }
  if (hitPt >= -220 && powerPt <= -85) {
    return { ...baseRes, outcome: 'GO', flavor: '詰まった内野ゴロ', staminaDelta: -1 };
  }
  return { ...baseRes, outcome: 'GO', flavor: '力ない内野ゴロ', staminaDelta: -1 };
}

// 球種カテゴリでスタミナ消費量(基礎)を決める
//   FB系 (フォーシーム/サイクロン/ライジング) と 2C系 (ツーシーム/シンカー/ワンシーム/ノーシーム): -2
//   先発投手: ナックル系 -2.5、それ以外の変化球 -3
//   中継ぎ・抑え: ナックル含む変化球すべて -2.5
function getStaminaDrainPerPitch(pitchName, isStarter) {
  if (!pitchName) return 2;
  const fb      = /フォーシーム|サイクロン|ライジング/.test(pitchName);
  const twoSeam = /ツーシーム|シンカー|ワンシーム|ノーシーム/.test(pitchName);
  if (fb || twoSeam) return 2;
  const knuckle = /ナックル/.test(pitchName);
  if (isStarter) return knuckle ? 2.5 : 3;
  return 2.5;
}

// 投球ごとのスタミナ反映
//   - ファインプレー時は +2 で回復 (例外処理)
//   - 通常時は 「球種・役割別の基礎消費 + 失点 × 2」を減算
function applyPitchStaminaDelta(res, pitch) {
  const info = getActivePitcherInfo();
  if (res && res.fineplay) {
    G.setup[info.side].pitcherStamina[info.idx] += 2;
    return;
  }
  const isStarter = (info.idx === 0);
  const baseDrain = getStaminaDrainPerPitch(pitch ? pitch.name : '', isStarter);
  const runs = G.lastPitchRuns || 0;
  const totalDrain = baseDrain + runs * 2;  // 失点 1 につき -2
  G.setup[info.side].pitcherStamina[info.idx] -= totalDrain;
}

// 低スタミナの先発が好投したときのスタミナ・ボーナス (全試合共通)。各回のチェックポイントで最大1回ずつ:
//   3回1アウトまで無失点   かつ 残量≤30 → +5
//   4回1アウトまで1失点以内 かつ 残量≤20 → +5
//   5回1アウトまで2失点以内 かつ 残量≤10 → +5
function applyStarterStaminaBonus() {
  const info = getActivePitcherInfo();
  if (!info || info.idx !== 0) return;            // 登板中が先発投手のときのみ
  if (G.outs < 1) return;                         // その回の1アウト到達時点で判定
  const CHECK = { 3: { maxRuns: 0, maxSta: 30 }, 4: { maxRuns: 1, maxSta: 20 }, 5: { maxRuns: 2, maxSta: 10 } };
  const c = CHECK[G.inning];
  if (!c) return;
  if (!G.starterBonusGiven) G.starterBonusGiven = {};
  const key = info.side + '_' + G.inning;
  if (G.starterBonusGiven[key]) return;           // この回のチェックポイントは評価済み
  G.starterBonusGiven[key] = true;
  const lg = G.pitcherLog[info.side] && G.pitcherLog[info.side][0];
  const runsAllowed = (lg && lg.runsAllowed) || 0;
  const sta = G.setup[info.side].pitcherStamina[0];
  if (runsAllowed <= c.maxRuns && sta <= c.maxSta) {
    G.setup[info.side].pitcherStamina[0] += 5;
    const nm = (G.setup[info.side].pitchers[0] || {}).fullNameTop || '先発投手';
    logLine(`💪 ${nm} 好投でスタミナ +5 (${G.inning}回1死・${runsAllowed}失点・残${Math.round(sta)})`, 'event-inning');
  }
}

function applyOutcome(outcome, pitch) {
  const B = G.currentBatter, side = G.top ? 'away' : 'home';
  const defSide = G.top ? 'home' : 'away';
  const P = G.currentPitcher;
  const bIdx = G.top ? G.awayBatIdx : G.homeBatIdx;
  const bStat = G.batterStats[side][bIdx];      // 現在の打席のスロット記録
  const pLog  = G.pitcherLog[defSide];
  const pStat = pLog[pLog.length - 1];           // 現役投手のログ
  let runs = 0;
  // 表示用 球速・キレ (decidePitchOutcome が G.lastPitchResult にセット)
  const lpr = G.lastPitchResult || {};
  const shownSpeed = lpr.displaySpeed != null ? lpr.displaySpeed : pitch.speed;
  const shownKire  = lpr.kire != null ? lpr.kire : null;
  const speedPart = shownSpeed ? `${shownSpeed}km/h` : '';
  const kirePart  = shownKire != null ? `(キレ${shownKire})` : '';
  const pitchInfo = `${speedPart}${kirePart}${speedPart || kirePart ? 'の' : ''}${pitch.name}`;
  const runnersInfo = { runners: [...G.bases] }; // 走者退避 (HR集計用)

  // 打席結果トラッキング用: ホームインした走者の {side, slotIdx} を集める
  let scoredRunners = [];
  // 自分自身が打席で走者を表す情報。respStint = この走者を出した(責任)投手の登板記録。
  //   → 引き継ぎ走者が生還しても、その失点・自責点は出した投手に帰属させる(野球ルール)。
  const meRef = { side, slotIdx: bIdx, respStint: pStat };

  switch (outcome) {
    case 'BB': {
      // 押し出し計算 (満塁時)
      if (G.bases[0]) {
        if (G.bases[1]) {
          if (G.bases[2]) { scoredRunners.push(G.bases[2]); }
          G.bases[2] = G.bases[1];
        }
        G.bases[1] = G.bases[0];
      }
      G.bases[0] = meRef;
      runs = scoredRunners.length;
      bStat.BB++;
      pStat.BB++;
      break;
    }
    case 'K': {
      G.outs++;
      G.ks[side]++;
      bStat.AB++; bStat.K++;
      pStat.K++; pStat.outs++;
      break;
    }
    case 'HR': {
      // 走者全員 + 打者がホームイン
      scoredRunners = G.bases.filter(b => b);
      const runners = scoredRunners.length;
      runs = 1 + runners;
      // 打者自身もR加算
      bStat.R++;
      bStat.AB++; bStat.H++; bStat.HR++; bStat.RBI += runs;
      pStat.hits++; pStat.HR++;
      G.bases = [null, null, null];
      G.hits[side]++;
      G.hrEvents.push({
        inning: G.inning, top: G.top, side,
        batter: `${B.fullNameTop}${B.year?'('+B.year+')':''}`,
        batterName: B.fullNameTop,
        batterKey: playerKey(B),
        batterTeam: B.team,
        pitcher: `${P.fullNameTop}${P.year?'('+P.year+')':''}`,
        runs, runners,
      });
      break;
    }
    case '3B': {
      scoredRunners = advanceRunnersOnHit(3);
      runs = scoredRunners.length;
      G.bases[2] = meRef;
      G.hits[side]++;
      bStat.AB++; bStat.H++; bStat.triples++; bStat.RBI += runs;
      pStat.hits++;
      break;
    }
    case '2B': {
      scoredRunners = advanceRunnersOnHit(2);
      runs = scoredRunners.length;
      G.bases[1] = meRef;
      G.hits[side]++;
      bStat.AB++; bStat.H++; bStat.doubles++; bStat.RBI += runs;
      pStat.hits++;
      break;
    }
    case '1B': {
      // 内野安打は走者を1つだけ進塁させる (足の遅い内野打球なので2塁→本塁のような余分な進塁はしない)。
      //   通常の外野へのヒットは従来通り走力に応じて追加進塁あり (advanceRunnersOnHit)。
      const infieldHit = /内野/.test((lpr && lpr.flavor) || '');
      scoredRunners = infieldHit ? pushRunners(1) : advanceRunnersOnHit(1);
      runs = scoredRunners.length;
      G.bases[0] = meRef;
      G.hits[side]++;
      bStat.AB++; bStat.H++; bStat.RBI += runs;
      pStat.hits++;
      break;
    }
    case 'E': {
      // 失策出塁: 打者は一塁へ。アウトにはならず、安打にも計上しない (打数のみ加算)。
      // 走者は単打と同様に1つ進む。失策による出塁・得点は自責点に含めない。
      scoredRunners = pushRunners(1);
      runs = scoredRunners.length;
      G.bases[0] = { side, slotIdx: bIdx, errorReach: true, respStint: pStat };  // 失策出塁=非自責。Rは責任投手へ
      bStat.AB++;        // 打数は加算 (安打・打点は付かない)
      // 失策を犯した野手 (守備側) の失策数(E)に加算
      const defPosArr = G.setup[defSide] && G.setup[defSide].batterPos;
      const eIdx = defPosArr ? defPosArr.findIndex(p => p === lpr.fielderPos) : -1;
      if (eIdx >= 0 && G.batterStats[defSide] && G.batterStats[defSide][eIdx]) {
        G.batterStats[defSide][eIdx].E = (G.batterStats[defSide][eIdx].E || 0) + 1;
      }
      break;
    }
    case 'FO': {
      G.outs++;
      pStat.outs++;
      // 外野フライ(内野フライ/ファール以外)はタッチアップ進塁を判定。走者が生還すれば犠飛扱い。
      const flFO = lpr.flavor || '';
      if (G.outs < 3 && /外野/.test(flFO) && !/内野/.test(flFO)) {
        scoredRunners = tagUpAdvance(flyDepth(flFO));
        if (scoredRunners.length > 0) {
          runs = scoredRunners.length;
          bStat.SAC = (bStat.SAC || 0) + 1;   // タッチアップ生還 → 犠飛 (打数なし)
          bStat.RBI += runs;
          outcome = 'SAC_FLY';
        } else {
          bStat.AB++;                          // 生還なし → 通常フライアウト (2塁→3塁の進塁はあり得る)
        }
      } else {
        bStat.AB++;
      }
      break;
    }
    case 'LO':
      G.outs++;
      bStat.AB++;
      pStat.outs++;
      break;
    case 'SAC_FLY': {
      // 大きな外野フライ — タッチアップ(深さ＋走力依存)で3塁走者が生還すれば犠牲フライ
      G.outs++;
      pStat.outs++;
      scoredRunners = tagUpAdvance('deep');
      if (scoredRunners.length > 0) {
        runs = scoredRunners.length;
        bStat.SAC = (bStat.SAC || 0) + 1;
        bStat.RBI += runs;
        outcome = 'SAC_FLY';
      } else {
        bStat.AB++;
        outcome = 'FO';  // 誰も還らなければ通常外野フライ扱い (2塁→3塁の進塁はあり得る)
      }
      break;
    }
    case 'GO_SLOW': {
      // 内野ボテボテ(緩い)ゴロ — 走者は走力依存で進みやすい (3B走者は生還しやすい)
      G.outs++;
      pStat.outs++;
      bStat.AB++;
      scoredRunners = advanceRunnersOnGrounder(true);
      runs = scoredRunners.length;
      bStat.RBI += runs;
      break;
    }
    case 'GO': {
      G.outs++;
      pStat.outs++;
      bStat.AB++;
      if (G.bases[0] && G.outs < 3 && rand() < 0.35) {
        // 併殺(ゲッツー): 1塁走者ありで成立。打者(1塁封殺)と1塁走者(2塁封殺)がアウト。残る走者は1つ進塁。
        //   例: 無死1・2塁 → 2塁走者は3塁へ進み、1塁走者と打者がゲッツー。
        G.outs++;
        pStat.outs++;
        outcome = 'GO_DP';
        const dpThirdOut = (G.outs >= 3);   // 併殺で3アウト目が成立した場合
        if (G.bases[2] && !dpThirdOut) scoredRunners.push(G.bases[2]);  // 3塁走者は生還 (3アウト時は無効)
        G.bases[2] = G.bases[1] || null;    // 2塁走者 → 3塁
        G.bases[1] = null;                  // 2塁は封殺でアウト → 空く
        G.bases[0] = null;                  // 打者は1塁でアウト
        runs = scoredRunners.length;        // 併殺打のため打点(RBI)は加算しない (野球規則)
      } else if (/内野ゴロ/.test((lpr && lpr.flavor) || '') && !/ボテボテ|力ない|詰まった/.test((lpr && lpr.flavor) || '') && G.bases[0] && G.outs < 3) {
        // フォースアウト(野手選択): 強くも弱くもない内野ゴロ。1塁から連続する封殺チェーンを
        //   リード(高い塁)から走力依存で封殺判定。封殺できればその走者がアウト・打者は1塁セーフ、
        //   全員セーフなら打者が1塁アウト(通常)。
        const chain = [0];
        if (G.bases[1]) chain.push(1);
        if (G.bases[1] && G.bases[2]) chain.push(2);
        const outProb = spd => Math.max(0.15, Math.min(0.9, 0.7 - ((spd || 60) - 60) * 0.012));  // 遅いほど封殺されやすい
        let outBase = null;
        for (let k = chain.length - 1; k >= 0; k--) {
          if (rand() < outProb(runnerSpeed(G.bases[chain[k]]))) { outBase = chain[k]; break; }
        }
        if (outBase != null) {
          const snap = G.bases.slice(), nb = [null, null, null];
          // チェーン外の3塁走者 (1・3塁時など): 走力で生還 or 残塁
          if (snap[2] && chain.indexOf(2) < 0) {
            if (rand() < Math.max(0.04, Math.min(0.92, 0.22 + ((runnerSpeed(snap[2]) || 60) - 60) * 0.011))) scoredRunners.push(snap[2]);
            else nb[2] = snap[2];
          }
          // チェーン内の走者を1つ進塁 (封殺された走者は除外)
          for (const b of chain) {
            if (b === outBase) continue;
            const r = snap[b]; if (!r) continue;
            if (b + 1 >= 3) scoredRunners.push(r); else nb[b + 1] = r;
          }
          nb[0] = meRef;            // 打者は1塁セーフ (アウトは封殺された走者)
          G.bases = nb;
          runs = scoredRunners.length;
          bStat.RBI += runs;
          lpr.forceOut = outBase + 2;   // 2=二塁/3=三塁/4=本塁 封殺 → 動画 laa_forth{N}
          lpr.flavor = `内野ゴロ、${({ 2: '二塁', 3: '三塁', 4: '本塁' })[lpr.forceOut]}フォースアウト！`;
        } else {
          scoredRunners = advanceRunnersOnGrounder(false);   // 全員セーフ → 打者1塁アウト
          runs = scoredRunners.length;
          bStat.RBI += runs;
        }
      } else if (G.outs < 3) {
        // 非併殺の内野ゴロ — 走者は走力依存で進塁 (1塁走者なしでも 3塁→本塁/2塁→3塁 あり。通常ゴロは緩いゴロより進みにくい)
        scoredRunners = advanceRunnersOnGrounder(false);
        runs = scoredRunners.length;
        bStat.RBI += runs;
      }
      break;
    }
    default:
      G.outs++;
      pStat.outs++;
      bStat.AB++;
  }

  // 打球可視化用に「最終的な結果種別」と表示用の投手/打者/球種を記録。
  //   (この後 endAtBat で打者が進むため、ここで打者名等を確定させておく)
  if (G.lastPitchResult) {
    G.lastPitchResult.finalOutcome = outcome;
    G.lastPitchResult.dispBatter  = B ? (B.fullNameTop || B.nameJa || '') : '';
    G.lastPitchResult.dispPitcher = P ? (P.fullNameTop || P.nameJa || '') : '';
    // 球種名から通称(括弧書き、例:「フォーシーム（ビッグユニット砲）」)を割愛してはみ出しを防ぐ
    const _pn = (pitch && pitch.name ? pitch.name : '').replace(/[（(][^）)]*[）)]/g, '').trim();
    // 投球のキレ(0〜6)を球種名の後ろに付記 (例: フォーシーム(6))
    const _kire = (G.lastPitchResult.kire != null) ? `(${G.lastPitchResult.kire})` : '';
    G.lastPitchResult.dispPitch   = `${shownSpeed ? shownSpeed + 'km/hの' : ''}${_pn}${_kire}`;
    // アウトカウント/チェンジ表示用 (この後 endAtBat でイニング交代=outsリセットされる前に確定しておく)
    G.lastPitchResult.outsAfter   = G.outs;
    G.lastPitchResult.isThirdOut  = (G.outs >= 3);
  }

  // ===== 詳細実況を1行で出力 =====
  const isThirdOut = G.outs >= 3;
  let cls = 'event-out';
  if (outcome === 'HR') cls = 'event-hr';
  else if (['1B','2B','3B'].includes(outcome)) cls = 'event-hit';
  else if (outcome === 'K')  cls = 'event-k';
  else if (outcome === 'BB') cls = 'event-bb';
  let line;
  if (outcome === 'GO_DP') {
    line = `${B.fullNameTop}: ${P.fullNameTop}の${pitchInfo}を引っかけて、ゲッツー（ダブルプレー）！${isThirdOut ? 'スリーアウトチェンジ！！' : ''}`;
  } else if (lpr.flavor) {
    // 打席結果に応じて接続詞を切り替え (自然な実況表現)
    let connector;
    if (outcome === 'BB')      connector = 'を選んで、';   // フォアボール: 打たずに選ぶ
    else if (outcome === 'K')  connector = 'に、';         // 三振: 球に三振
    else                        connector = 'を打って、';   // ヒット/長打/凡退
    // ホームランのラン数別表記 (ソロ/ツーラン/スリーラン/満塁) を flavor に反映
    let flavor = lpr.flavor;
    let runsTxt = runs > 0 ? ` (${runs}点)` : '';
    if (outcome === 'HR') {
      const hrPrefix = runs >= 4 ? '満塁' : runs === 3 ? 'スリーラン' : runs === 2 ? 'ツーラン' : 'ソロ';
      flavor = flavor.replace('ホームラン', `${hrPrefix}ホームラン`);
      runsTxt = '';  // 「スリーラン」「満塁」等にラン数が含まれるので末尾点数は省略
    }
    line = `${B.fullNameTop}: ${P.fullNameTop}の${pitchInfo}${connector}${flavor}${runsTxt}${isThirdOut ? '、スリーアウトチェンジ！！' : ''}`;
  } else {
    line = formatPlayByPlay(B, P, outcome, pitch, runs, isThirdOut);
  }
  logLine(line, cls);

  // ホームインした走者の R を加算
  for (const r of scoredRunners) {
    if (r && r.side && G.batterStats[r.side][r.slotIdx]) {
      G.batterStats[r.side][r.slotIdx].R++;
    }
  }
  G.lastPitchRuns = runs;  // スタミナ計算 (失点1につき-2) で参照
  // === 投手の失点(R)・自責点(ER) ===
  // 生還した走者は「その走者を出した投手(respStint)」に帰属させる。
  //   → 引き継いだ走者が生還した分は、前の(出した)投手の責任 (野球ルール)。
  // 非自責 (ER に含めない): タイブレーク走者(ghost) / 失策出塁(errorReach) / 失策プレー(outcome==='E')。
  for (const r of scoredRunners) {
    const stint = (r && r.respStint) ? r.respStint : pStat;   // 担当(責任)投手。不明なら現投手
    stint.runsAllowed = (stint.runsAllowed || 0) + 1;
    const unearned = (outcome === 'E') || !!(r && (r.ghost || r.errorReach));
    if (!unearned) stint.earnedRuns = (stint.earnedRuns || 0) + 1;
  }
  // 打者自身の生還 (本塁打) は現投手の責任 (自責点に計上)
  const batterRuns = runs - scoredRunners.length;
  if (batterRuns > 0) {
    pStat.runsAllowed = (pStat.runsAllowed || 0) + batterRuns;
    pStat.earnedRuns  = (pStat.earnedRuns  || 0) + batterRuns;
  }
  // 投手の投球数 +1, 対戦打者数 +1
  pStat.pitches++;
  pStat.battersFaced++;
  // イニング別打席結果テキストを追加
  const iIdx = G.inning - 1;
  const txt = outcomeShort(outcome);
  if (bStat.perInning[iIdx]) {
    bStat.perInning[iIdx] += '・' + txt;
  } else {
    bStat.perInning[iIdx] = txt;
  }

  if (runs > 0) G.score[side][G.inning - 1] += runs;
  // (失点・対戦打者数は switch 内で加算済み)
  // リード履歴を記録 (打席後のスコア差)
  const aSum = G.score.away.reduce((a,b)=>a+b,0);
  const hSum = G.score.home.reduce((a,b)=>a+b,0);
  G.leadHistory.push({
    inning: G.inning, top: G.top,
    awayScore: aSum, homeScore: hSum,
    leadSide: aSum > hSum ? 'away' : (hSum > aSum ? 'home' : null),
    diff: Math.abs(aSum - hSum),
    awayPitcher: G.setup.away.pitchers[G.setup.away.activeIdx]?.fullNameTop,
    homePitcher: G.setup.home.pitchers[G.setup.home.activeIdx]?.fullNameTop,
  });
  // 打席履歴を記録 (試合終了後に ダイヤ内の 戻る/進む で振り返れるように)。
  //   endAtBat() でイニング交代すると outs/塁がリセットされるため、この時点で確定保存する。
  //   サイレント(シーズン自動進行)では振り返りを使わないため記録しない (高速化)。
  if (!G.silent) {
    if (!G.pitchHistory) G.pitchHistory = [];
    G.pitchHistory.push({
      res: G.lastPitchResult,
      runs,
      // 走者は「その時点の選手」を確定保存 (後の守備固め等でスロットの選手が変わっても振り返りで正しく表示)
      bases: G.bases.map(b => (b ? { ...b, _player: getRunnerPlayer(b) } : null)),
      inning: G.inning, top: G.top, outs: G.outs,
      info: G.lastInfo || '',   // その打席時点の外野通知(継投/代打等) — 振り返りで時系列表示
    });
  }
  G.historyView = null;  // 新しい打席が進んだらライブ表示に戻す
  endAtBat();
}

// 走者を by 塁進める。ホームインした走者の {side, slotIdx} 配列を返す
function pushRunners(by) {
  const scoredRunners = [];
  for (let i = 2; i >= 0; i--) {
    if (G.bases[i]) {
      const newPos = i + by;
      if (newPos >= 3) {
        scoredRunners.push(G.bases[i]);
        G.bases[i] = null;
      } else {
        G.bases[newPos] = G.bases[i];
        G.bases[i] = null;
      }
    }
  }
  return scoredRunners;
}

// 安打時の走者進塁。基本は by 塁進むが、走力(スピード＋盗塁能を少し加味)に応じて
// 確率で1つ余分に進む（単打で1塁→3塁、二塁打で1塁→本塁 等）。
// 前を走る走者は追い越さない。ホームインした走者の {side, slotIdx} 配列を返す。
function advanceRunnersOnHit(by) {
  const scoredRunners = [];
  const twoOut = (G.outs >= 2);   // 2アウトは打球と同時にスタート → 追加進塁が大幅に増える
  // 追加進塁が起きる確率を走者の走力(＋アウトカウント・塁)から算出
  const extraBaseProb = (ref, fromBase) => {
    const { souru, touru } = getRunnerStealStats(getRunnerPlayer(ref));
    const eff = (souru || 60) + Math.min(15, (touru || 0) * 0.25); // 走力に盗塁能を少し加味
    let baseP;
    if (by >= 2)             baseP = 0.34;   // 二塁打以上での追加進塁 (1塁→本塁 等)
    else if (fromBase === 1) baseP = 0.40;   // 単打で2塁走者が本塁へ (MLBでは普通に多い)
    else                     baseP = 0.18;   // 単打で1塁走者が3塁へ
    if (twoOut) baseP += 0.25;               // 2アウト時は積極走塁で上げる
    return Math.max(0.02, Math.min(0.93, baseP + (eff - 60) * 0.011));  // 生還確率の上限は93%
  };
  let leadStop = 4; // 前を走る走者の到達塁。これ以上は進めない(4=制限なし=本塁まで可)
  for (let i = 2; i >= 0; i--) {
    const ref = G.bases[i];
    if (!ref) continue;
    let dest = i + by;
    // まだ本塁に達していない走者のみ、走力で追加進塁を検討 (三塁打は全員生還で対象外)
    if (by < 3 && dest < 3 && Math.random() < extraBaseProb(ref, i)) dest += 1;
    if (dest >= leadStop) dest = leadStop - 1; // 前の走者を追い越さない
    G.bases[i] = null;
    if (dest >= 3) { scoredRunners.push(ref); leadStop = 4; } // 生還(本塁は空くので後続も本塁可)
    else { G.bases[dest] = ref; leadStop = dest; }
  }
  return scoredRunners;
}

// 走者のスピード(0〜100)。不明なら60。
function runnerSpeed(ref) {
  const pl = getRunnerPlayer(ref);
  return (pl && pl.stats && Number.isFinite(pl.stats['スピード'])) ? pl.stats['スピード'] : 60;
}
// 外野フライの深さ (実況フレーバー由来): deep(大きな) / shallow(浅い) / medium(通常)
function flyDepth(flavor) {
  const fl = flavor || '';
  if (/大きな|大飛球/.test(fl)) return 'deep';
  if (/浅い/.test(fl))          return 'shallow';
  return 'medium';
}
// 外野フライでのタッチアップ進塁。3塁走者→本塁 / 2塁走者→3塁(空きがあれば)。
//   深さ＋走力で確率判定。生還した走者配列を返す。
function tagUpAdvance(depth) {
  const scored = [];
  if (G.outs >= 3) return scored;
  const prob = (spd, toHome) => {
    let base;
    if (toHome) base = (depth === 'deep') ? 0.85 : (depth === 'shallow') ? 0.12 : 0.45;
    else        base = (depth === 'deep') ? 0.45 : (depth === 'shallow') ? 0.04 : 0.18;
    return Math.max(0.02, Math.min(0.97, base + ((spd || 60) - 60) * 0.010));
  };
  if (G.bases[2] && Math.random() < prob(runnerSpeed(G.bases[2]), true)) {
    scored.push(G.bases[2]); G.bases[2] = null;          // 3塁走者 タッチアップ生還
  }
  if (G.bases[1] && !G.bases[2] && Math.random() < prob(runnerSpeed(G.bases[1]), false)) {
    G.bases[2] = G.bases[1]; G.bases[1] = null;          // 2塁走者 タッチアップで3塁へ
  }
  return scored;
}
// ゴロ(非併殺)での走者進塁。weak=ボテボテ(緩)ゴロか。打者は1塁でアウト前提。
//   野球規則のフォース(封塁)を反映: 打者が一塁へ走ることで押し出される走者は必ず1つ進塁する。
//     ・1塁走者 = 常にフォース(打者に押し出される) → 必ず2塁へ
//     ・2塁走者 = 1塁に走者がいる時のみフォース → 必ず3塁へ
//     ・3塁走者 = 満塁(1・2塁とも埋まる)の時のみフォース → 必ず生還
//   フォースでない走者は走力＋打球の緩急で「もう1つ進めるか」を確率判定する。生還走者を返す。
function advanceRunnersOnGrounder(weak) {
  const scored = [];
  if (G.outs >= 3) return scored;
  const r0 = G.bases[0], r1 = G.bases[1], r2 = G.bases[2];
  // フォース判定 (打者→1塁を起点に、直前の塁が埋まっていれば押し出される)
  const f0 = !!r0;            // 1塁走者は常にフォース
  const f1 = f0 && !!r1;      // 2塁走者は1塁に走者がいる時のみフォース
  const f2 = f1 && !!r2;      // 3塁走者は1・2塁が埋まっている時のみフォース
  const adv = (spd, toHome) => {
    const base = toHome ? (weak ? 0.55 : 0.22) : (weak ? 0.60 : 0.40);
    return Math.max(0.04, Math.min(0.95, base + ((spd || 60) - 60) * 0.011));
  };
  const nb = [null, null, null];
  // 3塁走者: フォースなら必ず生還、非フォースは走力で生還判定 (残れば3塁残留)
  if (r2) {
    if (f2 || Math.random() < adv(runnerSpeed(r2), true)) scored.push(r2);
    else nb[2] = r2;
  }
  // 2塁走者: フォースなら必ず3塁、非フォースは3塁が空いていれば走力で進塁判定
  if (r1) {
    if (f1 || (!nb[2] && Math.random() < adv(runnerSpeed(r1), false))) nb[2] = r1;
    else nb[1] = r1;
  }
  // 1塁走者: 常にフォース → 2塁へ (フォースなので2塁は必ず空く)
  if (r0) {
    if (!nb[1]) nb[1] = r0;
    else nb[0] = r0;   // 通常起こらない安全策 (前位の走者が残った場合は1塁残留)
  }
  G.bases = nb;
  return scored;
}

// 実況フォーマッタ: 打席結果を詳細にナレーションする
//   例: "Aaジャッジ: Clカーショーの158km/hフォーシームを捉えて、特大の本塁打！3点が入る！"
function pickRand(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
function formatPlayByPlay(B, P, outcome, pitch, runs, isThirdOut) {
  const speedDesc = pitch.speed ? `${pitch.speed}km/h` : '';
  const pitchDesc = `${P.fullNameTop}の${speedDesc ? speedDesc+'の' : ''}${pitch.name}`;
  let action, result;
  switch (outcome) {
    case 'HR':
      action = pickRand(['完璧に捉えて','フルスイングで',　'振り抜いて','芯で捉えて','打球を持ち上げ']);
      if (runs >= 4)      result = `満塁本塁打！グランドスラム${runs}点！`;
      else if (runs === 3) result = `3ランホームラン！3点が入る！`;
      else if (runs === 2) result = `2ランホームラン！2点が入る！`;
      else                 result = `ソロホームラン！1点が入る！`;
      break;
    case '3B':
      action = pickRand(['鋭く弾き返して','右中間を破る','左中間を割る','一二塁間を抜けて']);
      result = `三塁打${runs ? `・走者生還で${runs}点` : ''}`;
      break;
    case '2B':
      action = pickRand(['力強く打って','ライン際を破る','ワンバウンドで','弾き返して']);
      result = `二塁打${runs ? `・${runs}点が入る` : ''}`;
      break;
    case '1B':
      action = pickRand(['振り抜いて','コンパクトに弾き返して','弾き返し','つぼに入れて','逆らわず流して']);
      result = `センター前ヒット${runs ? `・${runs}点が入る` : ''}`;
      break;
    case 'FO':
      action = pickRand(['打ち上げて','詰まって高々と','振り遅れて']);
      result = pickRand(['センターフライアウト','レフトフライアウト','ライトフライアウト']);
      break;
    case 'LO':
      action = pickRand(['強い打球で','低く鋭く','ライナー性の打球で']);
      result = pickRand(['ショートライナーアウト','セカンドライナーアウト','サードライナーアウト']);
      break;
    case 'GO':
      action = pickRand(['引っかけて','詰まって','打ちつけて','差し込まれて']);
      result = pickRand(['ショートゴロでアウト','セカンドゴロでアウト','サードゴロでアウト','一塁ゴロでアウト','投手ゴロでアウト']);
      break;
    case 'K':
      action = pickRand(['バットが空を切り','タイミングが合わず','差し込まれ','見極められず']);
      result = pickRand(['空振り三振！','見逃し三振！','三振！']);
      break;
    case 'BB':
      action = pickRand(['ボール球をしっかり見極めて','ストライクゾーンを見極め']);
      result = `四球を選んで出塁${runs ? `・押し出し${runs}点` : ''}`;
      break;
    default:
      action = '';
      result = outcome;
  }
  const tail = isThirdOut ? '、スリーアウトチェンジ！！' : '';
  return `${B.fullNameTop}: ${pitchDesc}を${action}、${result}${tail}`;
}

// 打席結果の短縮テキスト (イニング別欄に表示)
function outcomeShort(outcome) {
  switch (outcome) {
    case 'BB': return '四球';
    case 'K':  return '三振';
    case 'HR': return '本塁打';
    case '3B': return '三塁打';
    case '2B': return '二塁打';
    case '1B': return '安打';
    case 'FO': return 'フライ';
    case 'LO': return 'ライナー';
    case 'GO': return 'ゴロ';
    case 'GO_DP':   return '併殺';
    case 'GO_SLOW': return 'ゴロ';
    case 'SAC_FLY': return '犠飛';
    case 'E':  return '失策';
    default:   return outcome;
  }
}

function endAtBat() {
  // 打者は3アウト時も含めて常に次へ進める。
  // これにより、凡退した打者の次の打者から次の回の打順が始まる (通常の野球ルール)。
  advanceBatter();
  if (G.outs >= 3) {
    switchInning();
  }
  checkEnd();
}

// ============== スタミナ管理・投手交代 ==============
// 旧API互換: 個別呼び出しが必要な場合のフォールバック (現在は applyPitchStaminaDelta が主)
function reduceStamina(outcome) {
  const info = getActivePitcherInfo();
  let drain = 1;
  if (outcome === 'HR') drain = 3;
  else if (outcome === '3B' || outcome === '2B') drain = 2;
  G.setup[info.side].pitcherStamina[info.idx] -= drain;  // 負値を許容
}

function checkRelief() {
  const info = getActivePitcherInfo();
  const isStarter = (info.idx === 0);
  const log = G.pitcherLog[info.side];
  const myLog = log.length > 0 ? log[log.length - 1] : null;

  // スタミナ切れ(0以下)で、相手に追いつかれ/逆転された場合は 先発・リリーフ問わず即交代。
  //   無失点や 0-0 の同点では粘る。相手が得点して 同点以上(リードを失った) になった時に発動。
  //   (例: 8回にスタミナ切れの先発が押し出しで同点/逆転 → ここで降板)
  const oppSide  = info.side === 'away' ? 'home' : 'away';
  const defScore = G.score[info.side].reduce((a, b) => a + b, 0);  // 自軍(守備側)得点
  const batScore = G.score[oppSide].reduce((a, b) => a + b, 0);    // 相手(攻撃側)得点
  if (info.stamina <= 0 && batScore > 0 && batScore >= defScore && myLog && !myLog._forcedRelief) {
    myLog._forcedRelief = true;   // この登板では1回だけ試行 (リリーフ枯渇時の連投スパムを防止)
    autoReliefSwitch(info.side, true);
    return;
  }

  if (isStarter) {
    // 先発: 失点5以上ならスタミナ余力があっても即降板(血止め)。
    // スタミナ残量に応じた降板は「回終了時点」= イニング頭の manageStarterWorkload が判断する。
    const runsAllowed = myLog ? (myLog.runsAllowed || 0) : 0;
    if (runsAllowed >= 5) {
      autoReliefSwitch(info.side, true);
    } else if (info.stamina <= 3 && myLog && !myLog._warned) {
      logLine(`⚠️ ${info.pitcher.fullNameTop}: スタミナ残り${info.stamina}・継投を検討`, 'event-inning');
      myLog._warned = true;
    }
  } else {
    // リリーフ: イニング途中でも スタミナが尽きたら(0以下)交代。
    // (基本の「1イニングで交代」はイニング頭の manageRelieverWorkload が担当。
    //  ここは長い回などでスタミナが切れた場合の安全網。限界まで投げさせない方針)
    if (info.stamina <= 0) {
      autoReliefSwitch(info.side, true);
    } else if (info.stamina <= 5 && myLog && !myLog._warned) {
      logLine(`⚠️ ${info.pitcher.fullNameTop}: スタミナ残り${info.stamina}`, 'event-inning');
      myLog._warned = true;
    }
  }
}

// 勝利投手の権利 (先発が5回=15アウト以上投げ、かつ自軍リード中)
function isWinEligible(side, myLog) {
  if (!myLog) return false;
  if ((myLog.outs || 0) < 15) return false;
  const aSum = G.score.away.reduce((a,b)=>a+b,0);
  const hSum = G.score.home.reduce((a,b)=>a+b,0);
  const myScore = side === 'away' ? aSum : hSum;
  const oppScore = side === 'away' ? hSum : aSum;
  return myScore > oppScore;
}

// 先発がスタミナ切れでも続投すべきかを判定
//   ケース1: あと 1〜3 アウトで 5 回到達 (= 勝ち投手権利確定) で自軍リード中
//   ケース2: 完投目前 (8回到達済) かつ 失点少・自軍リード
function shouldContinueDespiteFatigue(info, myLog) {
  if (!myLog) return false;
  const outsPitched = myLog.outs || 0;
  const side = info.side;
  const aSum = G.score.away.reduce((a, b) => a + b, 0);
  const hSum = G.score.home.reduce((a, b) => a + b, 0);
  const myScore = side === 'away' ? aSum : hSum;
  const oppScore = side === 'away' ? hSum : aSum;
  const leading = myScore > oppScore;
  if (!leading) return false;
  if (outsPitched >= 12 && outsPitched < 15) return true;
  if (outsPitched >= 24 && (myLog.runsAllowed || 0) <= 2) return true;
  return false;
}

// 状況に応じて最適なリリーフ投手の index を選ぶ (MLB采配)
//   9回・1〜3点リード        → 抑え
//   7,8回・3点差以内          → SU
//   5点以上リード             → 中継の後ろ (3,4番手) / モップ
//   3点以内 (その他)          → 中継の前 (1,2番手)
function pickReliever(setup, sideKey) {
  const roles = setup.pitcherRoles || [];
  const sta = setup.pitcherStamina;
  const aSum = G.score.away.reduce((a,b)=>a+b,0);
  const hSum = G.score.home.reduce((a,b)=>a+b,0);
  const myScore = sideKey === 'away' ? aSum : hSum;
  const oppScore = sideKey === 'away' ? hSum : aSum;
  const lead = myScore - oppScore;
  const inning = G.inning;
  const closeGame = Math.abs(lead) <= 3;   // 僅差 (3点差以内)
  const extraInnings = inning >= 10;       // 通常9回で決着せず延長 (タイブレーク) に突入

  // 既に登板した投手は再登板できない (野球ルール)。登板済み(=ログに残る)を除外する。
  const used = usedPitcherSet(sideKey);
  const avail = (role) => {
    // SU・抑えは「僅差(3点差以内)の緊迫した場面」でのみ起用。大差では中継/MUで回す。
    if ((role === 'setup' || role === 'closer') && !closeGame) return [];
    const arr = [];
    for (let i = 0; i < setup.pitchers.length; i++) {
      if (i === setup.activeIdx) continue;
      if (used.has(setup.pitchers[i])) continue;   // 登板済みは再登板不可
      if (roles[i] === role && (sta[i] ?? 0) > -10) arr.push(i);
    }
    return arr;
  };

  // モップアップ起用条件:
  //   (a) 先発が5回未満 (アウト15未満) で降板  (b) 5回以降で5点差以上負け
  const currentRole = roles[setup.activeIdx];
  const starterLog = (G.pitcherLog[sideKey] || [])[0];
  const starterOuts = starterLog ? (starterLog.outs || 0) : 0;
  const earlyStarterKO = (currentRole === 'starter') && starterOuts < 15;
  const blowoutLoss = (inning >= 5) && (lead <= -5);

  // モップ優先は序盤(6回未満)のみ。6,7回以降はモップを最後尾にして中継へ託す。
  const mopFirstOk = inning < 6;
  let order;
  if ((earlyStarterKO || blowoutLoss) && mopFirstOk) {
    order = ['mop','middle','setup','closer'];   // 序盤の敗戦処理: モップ優先
  } else if (inning >= 9 && lead >= 1 && lead <= 3) {
    order = ['closer','setup','middle','mop'];
  } else if (inning >= 9 && lead === 0 && sideKey === 'home') {
    // 後攻の9回以降・同点: セーブ機会は永遠に来ない(サヨナラ勝ちで終わる)ため、
    //   最高レバレッジの今、最強リリーフ(抑え)から注ぎ込む (MLBの定石)
    order = ['closer','setup','middle','mop'];
  } else if ((inning === 7 || inning === 8) && closeGame) {
    order = ['setup','middle','closer','mop'];
  } else if (Math.abs(lead) >= 5) {
    order = ['mop','middle','setup','closer'];   // 大差: モップ優先で中継ぎを温存 (SU/抑えはavailで除外)
  } else {
    order = ['middle','setup','closer','mop'];   // 接戦/6回以降: 中継優先、モップは最後
  }

  // 役割別の起用スタミナ残率しきい値: 中継/MU=75% / SU/抑え=65%
  //   残率 = 現在スタミナ / スタミナ上限 (シーズンの持ち越しスタミナを反映)
  const roleThresh = { middle: 0.75, mop: 0.75, setup: 0.65, closer: 0.65 };
  const ratioOf = (i) => { const m = setup.pitcherMax[i] || 70; return m > 0 ? (sta[i] ?? 0) / m : 0; };
  // シーズンの登板数 (起用を分散させる指標。エキシビション/未記録は0)
  const appsOf = (i) => {
    if (!G.seasonMode || !SEASON) return 0;
    const s = SEASON.pit[playerKey(setup.pitchers[i])];
    return s ? (s.G || 0) : 0;
  };
  // 起用分散の並び: 登板数が少ない順 → スタミナ残率が高い順 (総合力では決めない)
  const distSort = (a, b) => (appsOf(a) - appsOf(b)) || (ratioOf(b) - ratioOf(a));
  // 延長(通常9回で決着せず)の守備は勝ちにこだわる: 抑え/SUが残率70%以上なら、
  //   優秀な投手から (抑え→SU1→SU2) 投入する。70%未満なら以降の通常継投順に委ねる。
  if (extraInnings && closeGame) {
    const fresh = [];
    for (const role of ['closer', 'setup']) {   // 抑えを最優先、次にSU
      const idxs = avail(role).filter(i => ratioOf(i) >= 0.70);
      idxs.sort((a, b) => (overallOf(setup.pitchers[b]) || 0) - (overallOf(setup.pitchers[a]) || 0));  // 各役割内は優秀(総合力)順
      fresh.push(...idxs);
    }
    if (fresh.length) return fresh[0];
  }
  // パス1: 役割優先順に「残率しきい値を満たす役割合致投手」を、登板数が少なく休めている順で起用
  for (const role of order) {
    const th = roleThresh[role] ?? 0;
    const idxs = avail(role).filter(i => ratioOf(i) >= th);
    if (idxs.length === 0) continue;
    idxs.sort(distSort);
    return idxs[0];
  }
  // パス2(緩和): しきい値到達者がいない場合も、役割優先順で登板数が少なく休めている投手を起用
  for (const role of order) {
    const idxs = avail(role);
    if (idxs.length === 0) continue;
    idxs.sort(distSort);
    return idxs[0];
  }
  // どの役割も不可 → 登板未経験でスタミナある任意の投手 (登板済みは再登板不可)
  for (let i = 0; i < setup.pitchers.length; i++) {
    if (i === setup.activeIdx) continue;
    if (used.has(setup.pitchers[i])) continue;
    if ((sta[i] ?? 0) > -10) return i;
  }
  return -1;
}

// その試合で既に登板した投手の集合 (再登板の禁止に使用)
function usedPitcherSet(sideKey) {
  return new Set((G.pitcherLog[sideKey] || []).map(l => l.pitcher));
}

// 指定 index の投手へ交代
function switchToPitcher(sideKey, idx) {
  const setup = G.setup[sideKey];
  if (idx < 0 || idx >= setup.pitchers.length) return false;
  setup.activeIdx = idx;
  G.currentPitcher = setup.pitchers[idx];
  G.pitcherLog[sideKey].push(newPitcherLog(setup.pitchers[idx], sideKey, G.inning, G.top));
  const roleLabel = PITCHER_ROLE_LABELS[(setup.pitcherRoles||[])[idx]] || '救援';
  logLine(`🔄 投手交代: ${G.currentPitcher.fullNameTop} (${roleLabel}/スタミナ${setup.pitcherStamina[idx]}/${setup.pitcherMax[idx]})`, 'event-inning');
  return true;
}

function autoReliefSwitch(sideKey, force) {
  const setup = G.setup[sideKey];
  const idx = pickReliever(setup, sideKey);
  if (idx >= 0) return switchToPitcher(sideKey, idx);
  if (force) {
    logLine(`⚠️ ${setup.pitchers[setup.activeIdx].fullNameTop}: リリーフ枯渇のため続投`, 'event-inning');
  }
  return false;
}

// イニング頭でのリリーフ登板過多の抑制。
//   中継/SU/抑え: 基本1イニングで交代。回が浅く(≤6)スタミナ残18以上なら2イニング、
//                 5回以前(≤5)でスタミナ残18以上なら3イニングまで許容。
//   モップ      : 5回まではロングリリーフ可、6回以降は中継へ託す(1イニングで交代)。
//   いずれもスタミナ限界(残8未満)までは投げさせない。先発(谷間先発含む)は対象外。
function manageRelieverWorkload(sideKey) {
  const setup = G.setup[sideKey];
  if (!setup || !setup.pitcherRoles) return;
  const idx = setup.activeIdx;
  const role = setup.pitcherRoles[idx];
  if (role === 'starter') return;   // 先発は別ロジック (checkRelief)
  const logArr = G.pitcherLog[sideKey] || [];
  const myLog = logArr[logArr.length - 1];
  if (!myLog) return;
  const inningsPitched = Math.floor((myLog.outs || 0) / 3);
  const stamina = setup.pitcherStamina[idx];
  const inning = G.inning;          // これから始まる守備イニング
  let cap;
  if (role === 'mop') {
    cap = (inning >= 6) ? 1 : 4;    // 6回以降は中継へ託す / 5回まではロングリリーフ可
  } else {
    cap = 1;                                          // 基本1イニング
    if (inning <= 6 && stamina >= 18) cap = 2;        // 序盤+スタミナ余裕 → 2イニング
    if (inning <= 5 && stamina >= 18) cap = 3;        // さらに浅い+余裕 → 3イニング
  }
  if (inningsPitched >= cap || stamina < 8) {
    autoReliefSwitch(sideKey, false);                 // 枯渇時(false)はそのまま続投
  }
}

// 先発投手の降板判断 (イニング頭=回終了時点で評価)。
//   ・失点5以上 → スタミナ余力があっても降板
//   ・5回未満(outs<15): 勝ち越し中(勝利投手がかかる)なら残量-2まで、それ以外は残量5まで粘る
//   ・5回到達(outs>=15):
//       - 1失点以上 かつ 勝利投手の権利 かつ 6回終了以降(回頭>=7) → 残量15未満でSU/抑えへ託す
//       - 無失点 or 2失点以内の同点 → 残量3まで粘る
//       - それ以外 → 残量10未満で交代
function manageStarterWorkload(sideKey) {
  const setup = G.setup[sideKey];
  if (!setup || !setup.pitcherRoles) return;
  const idx = setup.activeIdx;
  if (setup.pitcherRoles[idx] !== 'starter') return;   // 先発のみ
  const logArr = G.pitcherLog[sideKey] || [];
  const myLog = logArr[logArr.length - 1];
  if (!myLog) return;
  const runs = myLog.runsAllowed || 0;
  const outs = myLog.outs || 0;
  const stamina = setup.pitcherStamina[idx];
  const aSum = G.score.away.reduce((a,b)=>a+b,0);
  const hSum = G.score.home.reduce((a,b)=>a+b,0);
  const myScore = sideKey === 'away' ? aSum : hSum;
  const oppScore = sideKey === 'away' ? hSum : aSum;
  const leading = myScore > oppScore;
  const tied = myScore === oppScore;
  const winEligible = (outs >= 15) && leading;   // 5回以上+リード=勝利投手の権利
  const inning = G.inning;                        // これから始まる回 (前の回終了時点)

  // 失点5以上 → 余力あっても降板
  if (runs >= 5) { autoReliefSwitch(sideKey, false); return; }

  // MLB的な降板判断の追加 (現代の継投マネジメント):
  //   ・7回以降の接戦(2点差以内)で3失点以上 → 打者3巡目の失速が始まる前にブルペンへ (QSで御の字)
  //   ・7回以降の同点でスタミナ20未満 → 疲れた先発に同点の終盤を任せず、フレッシュな腕へ
  const closeGm = Math.abs(myScore - oppScore) <= 2;
  if (inning >= 7 && closeGm && runs >= 3) { autoReliefSwitch(sideKey, false); return; }
  if (inning >= 7 && tied && stamina < 20) { autoReliefSwitch(sideKey, false); return; }

  let threshold;
  if (outs < 15) {
    // 5回未満: 勝ち越し中なら-2まで(5回到達=勝利投手の権利確保を狙う)、それ以外は5まで
    threshold = leading ? -2 : 5;
  } else if (runs >= 1 && winEligible && inning >= 7) {
    // 1失点以上+勝利投手の権利+6回終了以降 → 15未満でSU/抑えへ託す
    threshold = 15;
  } else if (runs === 0 || (runs <= 2 && tied)) {
    // 無失点 or 2失点以内の同点 → 3まで粘る
    threshold = 3;
  } else {
    threshold = 10;   // 5回以上の基本
  }
  if (stamina < threshold) autoReliefSwitch(sideKey, false);
}

// イニング頭の継投検討 (高レバレッジ起用)
//   9回・1〜3点リードなら、現投手が抑えでなければ抑えを投入 (セーブ機会)
function considerHighLeverageRelief(sideKey) {
  const setup = G.setup[sideKey];
  if (!setup.pitcherRoles) return;
  const aSum = G.score.away.reduce((a,b)=>a+b,0);
  const hSum = G.score.home.reduce((a,b)=>a+b,0);
  const myScore = sideKey === 'away' ? aSum : hSum;
  const oppScore = sideKey === 'away' ? hSum : aSum;
  const lead = myScore - oppScore;
  const curRole = setup.pitcherRoles[setup.activeIdx];
  const used = usedPitcherSet(sideKey);   // 登板済みは再登板不可
  // 9回・1〜3点リード → 抑え投入 (未登板 かつ スタミナ残率65%以上の抑えのみ。無理使いしない)
  if (G.inning >= 9 && lead >= 1 && lead <= 3 && curRole !== 'closer') {
    const okRatio = (i) => { const m = setup.pitcherMax[i] || 70; return m > 0 ? (setup.pitcherStamina[i] ?? 0) / m : 0; };
    const ci = setup.pitcherRoles.findIndex((r, i) =>
      r === 'closer' && i !== setup.activeIdx && !used.has(setup.pitchers[i]) && okRatio(i) >= 0.65);
    if (ci >= 0) switchToPitcher(sideKey, ci);
  }
}

// 手動投手交代UI
function showReliefDialog() {
  const info = getActivePitcherInfo();
  const setup = info.setup;
  const opts = setup.pitchers.map((p, i) => ({ p, i, sta: setup.pitcherStamina[i], max: setup.pitcherMax[i] }))
    .filter(o => o.i !== info.idx && o.sta > -20);
  if (opts.length === 0) {
    alert('交代可能な投手がいません');
    return;
  }
  const label = opts.map((o,k) => `${k+1}: ${o.p.fullNameTop} (スタミナ${o.sta}/${o.max})`).join('\n');
  const choice = prompt(`投手交代\n現在: ${info.pitcher.fullNameTop} (スタミナ${info.stamina}/${info.maxStamina})\n\n${label}\n\n番号を入力 (キャンセルで中止):`);
  if (!choice) return;
  const n = parseInt(choice) - 1;
  if (n < 0 || n >= opts.length) { alert('無効な選択です'); return; }
  setup.activeIdx = opts[n].i;
  G.currentPitcher = opts[n].p;
  G.pitcherLog[info.side].push(newPitcherLog(opts[n].p, info.side, G.inning, G.top));
  logLine(`🔄 投手交代: ${G.currentPitcher.fullNameTop} (スタミナ${opts[n].sta}/${opts[n].max})`, 'event-inning');
  // 投手交代の演出 (登場動画)。次の球は同じ打者なので打者登場は流さない。
  if (!G._lastIntroPitcher) G._lastIntroPitcher = { home: null, away: null };
  G._lastIntroPitcher[info.side] = G.currentPitcher;   // 回頭の重複再生を防ぐ
  playVideoOverlay([pickVideo(PITCHER_INTRO_VIDEOS)]);
  renderAll();
}

// ============== 盗塁 (1塁走者の二盗) ==============
// 能力値:
//   souru  = 走者「スピード」(走力)
//   touru  = 走者「盗塁能」
//   arm    = 守備側捕手「阻止率」(% を抜いた数値)
//   taoTou = 守備側投手「対盗塁」
// 成否は stealSuccessProb() の成功率に基づいて確率判定する
// (提示された 阻止率×走力 の成功率テーブルに合わせて算出)。
// 自動盗塁を「画策するか」は maybeAutoSteal() の条件式で判定 (rand19=0..19)。

// 守備側 (現在マウンドにいる守備チーム) の捕手「阻止率」と投手「対盗塁」を取得
function getStealDefenseStats() {
  const defSide = G.top ? 'home' : 'away';
  const setup = G.setup[defSide] || {};
  const catIdx = setup.batterPos ? setup.batterPos.findIndex(p => p === 'C') : -1;
  const catcher = (catIdx >= 0 && setup.batters) ? setup.batters[catIdx] : null;
  const arm = (catcher && catcher.catcher && catcher.catcher['阻止率'] != null)
    ? catcher.catcher['阻止率'] : 30;
  const pitcher = G.currentPitcher;
  const taoTou = (pitcher && pitcher.statsMini && pitcher.statsMini['対盗塁'] != null)
    ? pitcher.statsMini['対盗塁'] : 0;
  return { catcher, arm, pitcher, taoTou };
}

// 走者の盗塁能力 (souru=スピード, touru=盗塁能)
function getRunnerStealStats(runnerPlayer) {
  const s = (runnerPlayer && runnerPlayer.stats) || {};
  const m = (runnerPlayer && runnerPlayer.statsMini) || {};
  const souru = (s['スピード'] != null) ? s['スピード'] : 60;
  const touru = (m['盗塁能'] != null) ? m['盗塁能'] : 0;
  return { souru, touru };
}

// 盗塁成功率 (0..1) を能力値から算出する。
// 提示された成功率テーブル (守備=阻止率, 走者=走力) に線形近似でフィットさせたもの。
//        走60  走80  走100
//  阻20:  52%   70%   87%
//  阻30:  44%   59%   74%
//  阻40:  39%   52%   65%
//   成功% ≈ 0.76×走力 + 1.2×盗塁能 - 0.88×捕手阻止率 - 0.5×投手対盗塁 + 26.8
//   (最終的に 3%〜97% に収める)
// 旧実装は「走力(0〜100) - 乱数(最大29) - 対盗塁」だったため、走力が乱数より十分大きく
// 常にプラス→全員成功、になっていた。捕手の阻止率も判定に効いていなかった。
function stealSuccessProb(souru, touru, arm, taoTou) {
  const pct = 0.76 * (souru || 0)
            + 1.2  * (touru || 0)
            - 0.88 * (arm != null ? arm : 30)
            - 0.5  * (taoTou || 0)
            + 26.8;
  return Math.max(3, Math.min(97, pct)) / 100;
}

// 盗塁判定 (1回分)。成功率に基づき成否を決める。成功で true。
function judgeStealOnce(souru, touru, arm, taoTou) {
  return Math.random() < stealSuccessProb(souru, touru, arm, taoTou);
}

// 目安成功率(%): 手動ダイアログ表示用
function estimateStealRate(souru, touru, arm, taoTou) {
  return Math.round(stealSuccessProb(souru, touru, arm, taoTou) * 100);
}

// 1塁走者が二盗できる状況か (1塁に自軍走者あり かつ 2塁が空)
function canStealSecond(side) {
  const ref = G.bases[0];
  return !!(ref && ref.side === side && !G.bases[1]);
}

// 1塁走者の二盗を実行。戻り値 { attempted, success, thirdOut }
//   成功: 1塁→2塁へ進塁し SB++
//   失敗: 走者アウト (G.outs++ かつ 守備側投手のアウト数にも計上)
function executeSteal(side) {
  const ref = G.bases[0];
  if (!ref) return { attempted: false, success: false, thirdOut: false };
  const runnerPlayer = getRunnerPlayer(ref);
  const rn = runnerPlayer ? (runnerPlayer.fullNameTop || runnerPlayer.nameJa || '走者') : '走者';
  const { souru, touru } = getRunnerStealStats(runnerPlayer);
  const { arm, taoTou } = getStealDefenseStats();
  const success = judgeStealOnce(souru, touru, arm, taoTou);
  if (success) {
    G.bases[1] = ref;     // 1塁 → 2塁
    G.bases[0] = null;
    const st = G.batterStats && G.batterStats[ref.side] && G.batterStats[ref.side][ref.slotIdx];
    if (st) st.SB = (st.SB || 0) + 1;
    logLine(`🏃💨 ${rn} 二盗成功！`, 'event-hit');
    return { attempted: true, success: true, thirdOut: false };
  }
  // 失敗 (盗塁死)
  const stF = G.batterStats && G.batterStats[ref.side] && G.batterStats[ref.side][ref.slotIdx];
  if (stF) stF.CS = (stF.CS || 0) + 1;   // 盗塁死を走者に計上
  G.bases[0] = null;
  G.outs++;
  const defSide = G.top ? 'home' : 'away';
  const pLog = G.pitcherLog[defSide];
  const pStat = pLog && pLog[pLog.length - 1];
  if (pStat) pStat.outs = (pStat.outs || 0) + 1;  // イニング消化 (投手の投球回に計上)
  const thirdOut = G.outs >= 3;
  logLine(`🏃❌ ${rn} 盗塁失敗、二塁でタッチアウト！${thirdOut ? ' スリーアウトチェンジ！！' : ''}`, 'event-out');
  return { attempted: true, success: false, thirdOut };
}

// 手動盗塁ダイアログ: 能力値と目安成功率を開示し、実行可否を確認
function stealDialog(side) {
  if (!canStealSecond(side)) {
    alert('盗塁できる状況ではありません\n(1塁に自軍の走者がいて、2塁が空いているときのみ可能です)');
    return;
  }
  const ref = G.bases[0];
  const runnerPlayer = getRunnerPlayer(ref);
  const rn = runnerPlayer ? (runnerPlayer.fullNameTop || runnerPlayer.nameJa || '走者') : '走者';
  const { souru, touru } = getRunnerStealStats(runnerPlayer);
  const { catcher, arm, pitcher, taoTou } = getStealDefenseStats();
  const rate = estimateStealRate(souru, touru, arm, taoTou);
  const catName = catcher ? (catcher.fullNameTop || catcher.nameJa || '?') : '不明';
  const pitName = pitcher ? (pitcher.fullNameTop || pitcher.nameJa || '?') : '不明';
  const msg =
    `🏃 盗塁 (二盗) — 1塁走者 ${rn}\n\n` +
    `【走者の能力】\n` +
    `  スピード(走力): ${souru}\n` +
    `  盗塁能: ${touru}\n\n` +
    `【守備側】\n` +
    `  捕手 ${catName} 阻止率: ${arm}\n` +
    `  投手 ${pitName} 対盗塁: ${taoTou}\n\n` +
    `▶ 目安の成功率: 約 ${rate}%\n\n` +
    `盗塁を試みますか？\n(OK = 試みる / キャンセル = やめる)`;
  if (!confirm(msg)) return;
  const r = executeSteal(side);
  // 失敗で3アウトならイニング交代 (打者は次イニング先頭に残す)
  if (r.thirdOut) { switchInning(); checkEnd(); }
  renderAll();
}

// 自動盗塁: 1塁走者が二盗を画策するか判定し、するなら実行する。
// 画策条件 (いずれか成立で試行)。rand19=0..19。各行末は画策確率の目安。
// 戻り値: 実行した場合 { attempted, success, thirdOut } / 画策しない場合 null
function maybeAutoSteal(side) {
  if (!canStealSecond(side)) return null;
  const ref = G.bases[0];
  const runnerPlayer = getRunnerPlayer(ref);
  const { souru, touru } = getRunnerStealStats(runnerPlayer);
  const { arm, taoTou } = getStealDefenseStats();
  const rand19 = Math.floor(Math.random() * 20);  // 0..19
  const attempt =
    (souru + touru     - arm - 40  > 0 && rand19 > 16) ||  // 15%画策
    (souru + touru     - arm - 60  > 0 && rand19 > 14) ||  // 25%画策
    (souru + touru     - arm - 80  > 0 && rand19 > 12) ||  // 35%画策
    (souru + touru * 3 - arm - 130 > 0 && rand19 > 8)  ||  // 55%画策 (盗塁能依存)
    (souru + touru * 4 - arm - 240 - taoTou > 0 && rand19 > 4)  ||  // 75%画策 (盗塁能依存)
    (souru + touru     - arm - 100 - taoTou > 0 && rand19 > 10);    // 45%画策
  if (!attempt) return null;
  // 盗塁機能(スピード+盗塁能)が高い選手ほど企画数を削減し、シーズン盗塁数のバランスを取る
  const sbFunc = souru + touru;
  let cut = 0;
  if      (sbFunc >= 120) cut = 0.40;   // 40%減
  else if (sbFunc >= 110) cut = 0.30;   // 30%減
  else if (sbFunc >= 100) cut = 0.20;   // 20%減
  else if (sbFunc >= 90)  cut = 0.10;   // 10%減
  else if (sbFunc >= 80)  cut = 0.05;   // 5%減
  if (cut > 0 && Math.random() < cut) return null;   // 今回は企画を見送る
  return executeSteal(side);
}

// ============== 手動 代打 / 代走 UI ==============
// ボタン → モード選択 → 候補選択 (prompt式)。推奨選手を実況に表示し、
// canSecureDefenseForSlot で「次の守備が成立する」場合のみ起用を許可する。
function showPinchDialog() {
  if (G.ended) { alert('試合は終了しています'); return; }
  const side = G.top ? 'away' : 'home';
  const setup = G.setup[side];
  const teamName = labelTeam(side);
  const hasBench = !!(setup.benchAvail && setup.benchAvail.length > 0);
  const canSteal = canStealSecond(side);
  // 控えも盗塁機会も無ければ何もできない
  if (!hasBench && !canSteal) {
    alert(`${teamName} は交代可能な控え選手がおらず、盗塁できる走者もいません`);
    return;
  }
  const mode = prompt(
    `代打 / 代走 / 盗塁 (${teamName})\n\n` +
    `1: 代打 (現在の打者に代える)\n` +
    `2: 代走 (塁上の走者に代える)\n` +
    `3: 盗塁 (1塁走者が二盗を試みる)\n\n` +
    `番号を入力 (キャンセルで中止):`);
  if (!mode) return;
  const m = mode.trim();
  if (m === '1') pinchHitDialog(side);
  else if (m === '2') pinchRunDialog(side);
  else if (m === '3') stealDialog(side);
  else alert('無効な選択です');
}

function pinchHitDialog(side) {
  const setup = G.setup[side];
  const batIdx = side === 'away' ? G.awayBatIdx : G.homeBatIdx;
  const cur = setup.batters[batIdx];
  // 守備が成立する候補のみ (代打が守れる or DH or ベンチに守備固め要員がいる)
  const cands = (setup.benchAvail || []).filter(p => canSecureDefenseForSlot(side, batIdx, p));
  if (cands.length === 0) {
    alert('守備を埋められる代打候補がいません (守備位置が成立しないため起用不可)');
    return;
  }
  // 推奨: 打撃総合値が最大の控え
  let recIdx = 0, recVal = -Infinity;
  cands.forEach((p, k) => { const v = offValue(p); if (v > recVal) { recVal = v; recIdx = k; } });
  const rec = cands[recIdx];
  logLine(`📋 代打 推奨: ${rec.fullNameTop} (${batIdx + 1}番 ${cur ? cur.fullNameTop : '?'} に代えて)`, 'event-inning');
  const label = cands.map((p, k) => {
    const r = p.record || {};
    return `${k + 1}: ${k === recIdx ? '⭐ ' : '   '}${p.fullNameTop}  [OPS ${r['OPS'] ?? '-'} / ミート ${p.stats?.['ミート'] ?? '-'} / パワー ${p.stats?.['パワー'] ?? '-'}]`;
  }).join('\n');
  const choice = prompt(`代打 — ${batIdx + 1}番 ${cur ? cur.fullNameTop : '?'} に代えて\n(⭐ = 推奨)\n\n${label}\n\n番号を入力 (キャンセルで中止):`);
  if (!choice) return;
  const n = parseInt(choice) - 1;
  if (!(n >= 0 && n < cands.length)) { alert('無効な選択です'); return; }
  applyBatterSub(side, batIdx, cands[n], '代打');
  renderAll();
}

function pinchRunDialog(side) {
  const setup = G.setup[side];
  // 自軍の塁上走者を列挙
  const runners = [];
  for (let i = 0; i < 3; i++) {
    const ref = G.bases[i];
    if (ref && ref.side === side) {
      runners.push({ base: i + 1, slotIdx: ref.slotIdx, runner: setup.batters[ref.slotIdx] });
    }
  }
  if (runners.length === 0) { alert('塁上に自軍の走者がいません'); return; }
  let target;
  if (runners.length === 1) {
    target = runners[0];
  } else {
    const rlabel = runners.map((r, k) => `${k + 1}: ${r.base}塁 ${r.runner ? r.runner.fullNameTop : '?'} (スピード ${r.runner?.stats?.['スピード'] ?? '-'})`).join('\n');
    const rc = prompt(`代走 — どの走者に代えますか?\n\n${rlabel}\n\n番号を入力 (キャンセルで中止):`);
    if (!rc) return;
    const rn = parseInt(rc) - 1;
    if (!(rn >= 0 && rn < runners.length)) { alert('無効な選択です'); return; }
    target = runners[rn];
  }
  // 守備が成立する候補のみ (8守備位置を漏れなく補完できる起用のみ許可)
  const cands = (setup.benchAvail || []).filter(p => canSecureDefenseForSlot(side, target.slotIdx, p));
  if (cands.length === 0) {
    alert('守備を埋められる代走候補がいません (8守備位置が成立しないため起用不可)');
    return;
  }
  // 並び順 = 「代走」役割を優先 → スピード(走力)が速い順。先頭を推奨にする。
  const speedOf = p => (p && p.stats && p.stats['スピード'] != null) ? p.stats['スピード'] : 0;
  cands.sort((a, b) => {
    const ra = benchRoleOf(side, a) === '代走' ? 0 : 1;
    const rb = benchRoleOf(side, b) === '代走' ? 0 : 1;
    if (ra !== rb) return ra - rb;
    return speedOf(b) - speedOf(a);
  });
  const recIdx = 0;
  const rec = cands[recIdx];
  logLine(`📋 代走 推奨: ${rec.fullNameTop} (${target.base}塁 ${target.runner ? target.runner.fullNameTop : '?'} に代えて)`, 'event-inning');
  const label = cands.map((p, k) => {
    const role = benchRoleOf(side, p);
    const roleTag = role ? `《${role}》` : '';
    return `${k + 1}: ${k === recIdx ? '⭐ ' : '   '}${p.fullNameTop}${roleTag}  [スピード ${p.stats?.['スピード'] ?? '-'} / 盗塁能 ${p.statsMini?.['盗塁能'] ?? '-'}]`;
  }).join('\n');
  const choice = prompt(`代走 — ${target.base}塁 ${target.runner ? target.runner.fullNameTop : '?'} に代えて\n(⭐ = 推奨)\n\n${label}\n\n番号を入力 (キャンセルで中止):`);
  if (!choice) return;
  const n = parseInt(choice) - 1;
  if (!(n >= 0 && n < cands.length)) { alert('無効な選択です'); return; }
  applyBatterSub(side, target.slotIdx, cands[n], '代走');
  renderAll();
}

function advanceBatter() {
  if (G.top) {
    G.awayBatIdx = (G.awayBatIdx + 1) % 9;
    G.currentBatter = G.setup.away.batters[G.awayBatIdx];
  } else {
    G.homeBatIdx = (G.homeBatIdx + 1) % 9;
    G.currentBatter = G.setup.home.batters[G.homeBatIdx];
  }
}

// ============== 野手交代 (代打 / 代走 / 守備固め) ==============
// 設計方針: 「次の守備で守備位置不成立にならない」ことを常に保証する。
//   - 代打/代走で入る選手がその守備位置を守れる、または DH 枠 → そのまま起用
//   - 守れない場合 → ベンチに守れる守備固め要員がいる時のみ起用し、その要員を予約。
//     予約は次の守備イニング頭 (resolveDefense) で適用する。
// 成績保持: G.batterStats[side] は常に「各打順スロットの現役選手」を保持する。
//   交代時は退く選手の成績を G.subLog[side] に凍結保存し、入る選手に新しい成績枠を割り当てる。
//   これにより既存の集計コード(打席記録/走者解決/実況)は現役選手をそのまま参照できる。

// 打順スロット slotIdx に選手 player の新しい成績オブジェクトを作る
function newBatterStat(side, slotIdx, player, subRole, fielded) {
  return {
    slotIdx,
    position: G.setup[side].batterPos?.[slotIdx] || '-',
    player,
    subRole: subRole || null,    // null=スタメン / '代打' / '代走' / '守備'
    fielded: !!fielded,          // 守備に就いたか
    AB: 0, R: 0, H: 0, RBI: 0, K: 0, BB: 0, HBP: 0, SAC: 0, SB: 0, CS: 0, E: 0, HR: 0,
    doubles: 0, triples: 0,
    perInning: new Array(9).fill(''),
  };
}

// slotIdx の現役選手を incoming へ差し替え、退く選手の成績を subLog へ凍結保存する
function recordSubSwap(side, slotIdx, incoming, role) {
  const arr = G.batterStats[side];
  if (arr && arr[slotIdx]) G.subLog[side].push(arr[slotIdx]);  // 退く選手を凍結
  if (arr) arr[slotIdx] = newBatterStat(side, slotIdx, incoming, role, false);
}

// 1人分の成績スナップショット(現役 or subLog)から守備欄ラベルを作る (Yahoo方式)
//   スタメン      → 守備位置の略号 (例: 三)
//   代打→守備就く → 「打」+略号 (例: 打二)、守備に就かず退く → 「打」
//   代走→守備就く → 「走」+略号 (例: 走左)、守備に就かず退く → 「走」
//   守備固め      → 守備位置の略号
function stintPosLabel(stint) {
  const posKey = stint.position;
  const abbr = POS_ABBR[posKey] || posKey || '-';
  const role = stint.subRole;
  if (!role) return abbr;                       // スタメン
  if (role === '代打') return stint.fielded ? ('打' + abbr) : '打';
  if (role === '代走') return stint.fielded ? ('走' + abbr) : '走';
  return abbr;                                  // 守備固め
}

function leadFor(side) {
  const aSum = G.score.away.reduce((a, b) => a + b, 0);
  const hSum = G.score.home.reduce((a, b) => a + b, 0);
  return side === 'away' ? aSum - hSum : hSum - aSum;
}

// 打撃総合値 (代打候補の評価)
function offValue(b) {
  if (!b) return -Infinity;
  const s = b.stats || {}, r = b.record || {};
  const pf = (v, d) => { const x = parseFloat(v); return Number.isFinite(x) ? x : d; };
  const ops = pf(r['OPS'], 0.700);
  return ops * 100 + (s['ミート'] || 60) * 0.4 + (s['パワー'] || 60) * 0.4
       + (s['チャンス'] || 60) * 0.3 + (s['選球眼'] || 60) * 0.2;
}
// 控え選手の登録役割 ('PH' / '代走' / '守備') を返す。未登録なら ''。
function benchRoleOf(side, player) {
  const m = G.setup[side] && G.setup[side].benchRole;
  return (m && m.get(player)) || '';
}

// 走力総合値 (代走候補の評価)
function speedValue(b) {
  if (!b) return -Infinity;
  const s = b.stats || {}, r = b.record || {}, m = b.statsMini || {};
  const pf = (v, d) => { const x = parseFloat(v); return Number.isFinite(x) ? x : d; };
  return (s['スピード'] || 60) + (m['盗塁能'] || 0) * 2 + pf(r['盗塁'], 0) * 0.4;
}

// 二部マッチング (Kuhn法): positions[k] を、それを守れる players のいずれかに重複なく割当てる。
// 全 position を割当てられたら assign[k]=players index の配列、無理なら null を返す。
function matchPositionsToPlayers(positions, players) {
  const matchOfPlayer = new Array(players.length).fill(-1);  // players index → positions index
  const tryK = (k, visited) => {
    for (let pi = 0; pi < players.length; pi++) {
      if (visited[pi]) continue;
      if (!canPlay(players[pi], positions[k])) continue;
      visited[pi] = true;
      if (matchOfPlayer[pi] === -1 || tryK(matchOfPlayer[pi], visited)) {
        matchOfPlayer[pi] = k;
        return true;
      }
    }
    return false;
  };
  for (let k = 0; k < positions.length; k++) {
    const visited = new Array(players.length).fill(false);
    if (!tryK(k, visited)) return null;  // この守備位置を誰も守れない → 不成立
  }
  // matchOfPlayer から assign[k] を作る
  const assign = new Array(positions.length).fill(-1);
  for (let pi = 0; pi < players.length; pi++) {
    if (matchOfPlayer[pi] >= 0) assign[matchOfPlayer[pi]] = pi;
  }
  return assign;
}

// 重み付き割当: positions[k] を players のいずれかに重複なく割当て、合計スコア最大の割当を返す。
//   scoreFn(player, pos) は割当不可なら -Infinity。 返り値 assign[k]=players index、不能なら null。
//   ビットマスクDP (players数 M が小さい前提)。M が大きすぎる場合は重みなしマッチングへフォールバック。
function bestAssignment(positions, players, scoreFn) {
  const nPos = positions.length, M = players.length;
  if (M < nPos) return null;
  if (M > 18) return matchPositionsToPlayers(positions, players);  // 安全弁
  const popcount = x => { let c = 0; while (x) { c += x & 1; x >>>= 1; } return c; };
  const size = 1 << M;
  const dp = new Array(size).fill(-Infinity);
  const par = new Array(size).fill(-1);
  dp[0] = 0;
  for (let mask = 0; mask < size; mask++) {
    const cur = dp[mask];
    if (cur === -Infinity) continue;
    const posIdx = popcount(mask);
    if (posIdx >= nPos) continue;   // 既に全守備位置を割当済み
    for (let j = 0; j < M; j++) {
      if (mask & (1 << j)) continue;
      const sc = scoreFn(players[j], positions[posIdx]);
      if (sc === -Infinity) continue;
      const nm = mask | (1 << j);
      if (cur + sc > dp[nm]) { dp[nm] = cur + sc; par[nm] = j; }
    }
  }
  let best = -Infinity, bestMask = -1;
  for (let mask = 0; mask < size; mask++) {
    if (dp[mask] !== -Infinity && popcount(mask) === nPos && dp[mask] > best) {
      best = dp[mask]; bestMask = mask;
    }
  }
  if (bestMask === -1) return null;
  const assign = new Array(nPos).fill(-1);
  let mask = bestMask;
  for (let p = nPos - 1; p >= 0; p--) {
    const j = par[mask];
    if (j < 0) return null;
    assign[p] = j;
    mask ^= (1 << j);
  }
  return assign;
}

// 大規模プールを bestAssignment が扱えるサイズ(limit)へ削減する。
// 総合力(ov)上位を残しつつ、各守備位置(DH以外)を最低1人は守れるようカバレッジを保証する。
function tbTrimForAssignment(cand, positions, ov, limit) {
  if (cand.length <= limit) return cand.slice();
  const sorted = cand.slice().sort((a, b) => ov(b) - ov(a));
  const chosen = sorted.slice(0, limit);
  const inSet = new Set(chosen);
  const fielding = positions.filter(p => p !== 'DH');
  for (const pos of fielding) {
    if (chosen.some(p => canPlay(p, pos))) continue;            // 既にこの位置を守れる選手がいる
    const best = sorted.find(p => !inSet.has(p) && canPlay(p, pos));
    if (!best) continue;                                        // 誰も守れない位置はスキップ
    // 入替対象: 外しても他位置のカバレッジを崩さない、最も総合力の低い選手
    for (let i = chosen.length - 1; i >= 0; i--) {
      const v = chosen[i];
      const breaks = fielding.some(fp => canPlay(v, fp) && !chosen.some(p => p !== v && canPlay(p, fp)));
      if (!breaks) { inSet.delete(v); chosen.splice(i, 1); chosen.push(best); inSet.add(best); break; }
    }
  }
  return chosen;
}

// 全ポジションを埋められない場合に、できるだけ多くのポジションを埋める部分割当 (制約の厳しい位置から貪欲)。
// scoreFn(player,pos)=-Infinity は割当不可。返り値 assign[k]=players index (未割当は -1)。
function tbGreedyMaxFill(positions, players, scoreFn) {
  const assign = new Array(positions.length).fill(-1);
  const used = new Array(players.length).fill(false);
  // 候補数の少ない(=制約の厳しい)ポジションから先に埋める
  const order = positions
    .map((pos, k) => ({ k, pos, n: players.filter(p => scoreFn(p, pos) !== -Infinity).length }))
    .sort((a, b) => a.n - b.n);
  for (const { k, pos } of order) {
    let bestJ = -1, bestSc = -Infinity;
    for (let j = 0; j < players.length; j++) {
      if (used[j]) continue;
      const sc = scoreFn(players[j], pos);
      if (sc > bestSc) { bestSc = sc; bestJ = j; }
    }
    if (bestJ >= 0 && bestSc !== -Infinity) { assign[k] = bestJ; used[bestJ] = true; }
  }
  return assign;
}

// 守備割当を探索する。slotOverride/subOverride を渡すと「その打順枠にその選手が入った」想定で評価。
// 8守備位置を「現役8野手(DH除く) + 控え」で重複なく埋められるかを、ポジション入替も含めて精査する。
// excludeBench=true なら控えを使わず(スタメンのポジション入替のみで)充足できるかを調べる。
// 返り値: { slotList, positions, poolPlayers, benchCount, assign } または null(充足不能)
function findFieldingMatching(side, slotOverride, subOverride, excludeBench) {
  const setup = G.setup[side];
  if (!setup || !setup.batters) return null;
  // 守備に就く打順枠 (DH を除く) と、その占有者を集める
  const slotList = [];
  for (let i = 0; i < 9; i++) {
    const pos = setup.batterPos[i];
    if (!pos || pos === 'DH') continue;
    let player = setup.batters[i];
    if (i === slotOverride) player = subOverride;       // 代打/代走で入る選手を仮置き
    else {
      const pd = (setup.pendingDef || []).find(d => d.slotIdx === i);
      if (pd) player = pd.defender;                     // 予約済み守備固めは確定とみなす
    }
    slotList.push({ idx: i, pos, player });
  }
  const positions = slotList.map(s => s.pos);
  // 候補プール: 現役占有者 (+ 控え。予約済みは benchAvail から既に除外済み)
  const benchPool = excludeBench ? [] : (setup.benchAvail || []).filter(p => p && p !== subOverride);
  const poolPlayers = [...slotList.map(s => s.player), ...benchPool];
  const assign = matchPositionsToPlayers(positions, poolPlayers);
  if (!assign) return null;
  return { slotList, positions, poolPlayers, benchCount: benchPool.length, assign };
}

// sub を slotIdx に入れた場合、次の守備で 8守備位置を漏れなく補完できるか。
//   sub 自身が守る / 控え(守備・PH)が入る / スタメンとのポジション入替、すべてを加味して精査する。
function canSecureDefenseForSlot(side, slotIdx, sub) {
  return findFieldingMatching(side, slotIdx, sub) !== null;
}

// slotIdx の選手を sub へ交代 (塁上参照 {side,slotIdx} は自動で sub に解決される)
function applyBatterSub(side, slotIdx, sub, kind) {
  const setup = G.setup[side];
  const pos = setup.batterPos[slotIdx];
  const outgoing = setup.batters[slotIdx];
  // ベンチから除外
  const bi = (setup.benchAvail || []).indexOf(sub);
  if (bi >= 0) setup.benchAvail.splice(bi, 1);
  // ラインナップ & スタッツ表示を差し替え (退く選手の成績は subLog へ凍結)
  setup.batters[slotIdx] = sub;
  recordSubSwap(side, slotIdx, sub, kind);
  // 現在の打者なら currentBatter を更新
  const curIdx = side === 'away' ? G.awayBatIdx : G.homeBatIdx;
  if (slotIdx === curIdx && ((side === 'away') === G.top)) {
    G.currentBatter = sub;
  }
  logLine(`🔁 ${kind}: ${outgoing ? outgoing.fullNameTop : '?'} → ${sub.fullNameTop} (${slotIdx + 1}番)`, 'event-inning');
  // sub がそのポジションを守れない場合の守備手当て:
  //  1) まずスタメン同士のポジション入替だけで全守備位置を充足できるか確認 (控え温存)
  //     → 充足できるなら何も予約しない。実際の入替は resolveDefense が行う。
  //  2) 入替だけでは不足する場合のみ、その位置を守れる控えを1人予約 (次守備で投入)
  if (pos !== 'DH' && !canPlay(sub, pos)) {
    const swapOnlyOk = findFieldingMatching(side, undefined, undefined, true) !== null;
    if (!swapOnlyOk) {
      setup.pendingDef = setup.pendingDef || [];
      const reserved = new Set(setup.pendingDef.map(pd => pd.defender));
      const def = (setup.benchAvail || []).find(p => !reserved.has(p) && canPlay(p, pos));
      if (def) {
        setup.pendingDef.push({ slotIdx, defender: def, pos });
        const di = setup.benchAvail.indexOf(def);
        if (di >= 0) setup.benchAvail.splice(di, 1);
      }
    }
  }
}

// 代打検討 (MLB監督の采配):
//   ・7回以降の接戦(リード2点以内〜ビハインド)で検討。6点差以上の大敗では控えを温存する。
//   ・レバレッジで基準を変える: 通常は+10明確に上回る時のみ / 得点圏は+6 / 9回以降のビハインド(最後の勝負)は+4。
//   ・相手投手が左腕なら「対左」を加味して比較する (プラトーン起用)。
//   ・得点圏ではチャンス(勝負強さ)も加点して比較する。
function maybePinchHit(side) {
  const setup = G.setup[side];
  if (!setup.benchAvail || setup.benchAvail.length === 0) return;
  if (G.inning < 7) return;
  const lead = leadFor(side);
  if (lead > 2) return;      // リード3点以上: 打線をいじらない
  if (lead <= -6) return;    // 6点差以上の大敗: 消化試合に控えを使わない
  const batIdx = side === 'away' ? G.awayBatIdx : G.homeBatIdx;
  const cur = setup.batters[batIdx];
  const risp = !!(G.bases[1] || G.bases[2]);            // 得点圏に走者
  const lastChance = (G.inning >= 9 && lead < 0);       // 9回以降のビハインド = 最後の勝負
  const oppP = G.currentPitcher;
  const vsL = !!(oppP && typeof oppP.hand === 'string' && oppP.hand.indexOf('左投') >= 0);
  const phVal = (p) => {
    let v = offValue(p);
    if (vsL) v += ((p && p.statsMini && Number.isFinite(p.statsMini['対左投手'])) ? p.statsMini['対左投手'] : 0) * 1.2;  // 対左補正
    if (risp) v += ((p && p.stats && Number.isFinite(p.stats['チャンス'])) ? p.stats['チャンス'] : 50) * 0.25;             // 勝負強さ
    return v;
  };
  const margin = lastChance ? 4 : (risp ? 6 : 10);   // 勝負どころほど積極的に動く
  let best = null, bestVal = phVal(cur) + margin;
  for (const ph of (setup.benchAvail || [])) {
    const v = phVal(ph);
    if (v <= bestVal) continue;
    if (!canSecureDefenseForSlot(side, batIdx, ph)) continue;
    bestVal = v; best = ph;
  }
  if (best) applyBatterSub(side, batIdx, best, '代打');
}

// 代走検討 (MLB監督の采配): 8回以降の接戦(2点差以内)で「意味のある走者」だけを速い控えに代える。
//   ・1塁/2塁の走者が主対象 (盗塁・単打での生還など足が得点に直結する)。
//   ・3塁走者は既に得点圏の奥で足の価値が下がるため、+20以上の大幅改善時のみ。
//   ・9回以降の同点/ビハインドは走者1人が試合を決める → 基準を+10に緩めて積極的に動く。
function maybePinchRun(side) {
  const setup = G.setup[side];
  if (!setup.benchAvail || setup.benchAvail.length === 0) return;
  if (G.inning < 8) return;
  const lead = leadFor(side);
  if (Math.abs(lead) > 2) return;
  const clutch = (G.inning >= 9 && lead <= 0);   // 9回以降の同点/ビハインド
  for (let i = 0; i < 3; i++) {
    const ref = G.bases[i];
    if (!ref || ref.side !== side) continue;
    const slotIdx = ref.slotIdx;
    const runner = setup.batters[slotIdx];
    if (!runner) continue;
    const need = (i === 2) ? 20 : (clutch ? 10 : 15);   // 3塁走者は大幅改善時のみ / 勝負どころは積極的に
    const minSpeed = speedValue(runner) + need;
    // 候補から「代走」役割を優先 → 走力総合値が高い順 に1人選ぶ
    let best = null, bestRank = 9, bestVal = -Infinity;
    for (const pr of (setup.benchAvail || [])) {
      const v = speedValue(pr);
      if (v <= minSpeed) continue;
      if (!canSecureDefenseForSlot(side, slotIdx, pr)) continue;
      const rank = benchRoleOf(side, pr) === '代走' ? 0 : 1;
      if (rank < bestRank || (rank === bestRank && v > bestVal)) {
        bestRank = rank; bestVal = v; best = pr;
      }
    }
    if (best) applyBatterSub(side, slotIdx, best, '代走');
  }
}

// 守備につく side の守備を確定させる (予約適用 → 安全網 → リード保護の守備固め)
function resolveDefense(side) {
  const setup = G.setup[side];
  if (!setup || !setup.batters) return;
  // 1) 予約済みの守備固めを適用
  if (setup.pendingDef && setup.pendingDef.length) {
    for (const pd of setup.pendingDef) {
      const out = setup.batters[pd.slotIdx];
      setup.batters[pd.slotIdx] = pd.defender;
      recordSubSwap(side, pd.slotIdx, pd.defender, '守備');
      logLine(`🧤 守備固め: ${out ? out.fullNameTop : '?'} → ${pd.defender.fullNameTop} (${POSITIONS[pd.pos] ? POSITIONS[pd.pos].label : pd.pos})`, 'event-inning');
    }
    setup.pendingDef = [];
  }
  // 2) 守れない枠があれば、ポジション入替 + 控え投入で 8守備位置を漏れなく充足
  //    (現役占有者で守れない位置を、スタメン同士の守備位置入替や控えの守備固めで補完)
  const needsFix = (() => {
    for (let i = 0; i < 9; i++) {
      const pos = setup.batterPos[i];
      if (pos !== 'DH' && !canPlay(setup.batters[i], pos)) return true;
    }
    return false;
  })();
  if (needsFix) applyDefenseMatching(side);
  // 3) リード保護の守備固め (8回以降・1〜3点リード時の任意アップグレード)
  maybeDefensiveUpgrade(side);
  // 4) この守備に就いた現役選手を fielded=true に (代打/代走が守備位置を引き継いだ判定)
  for (let i = 0; i < 9; i++) {
    if (G.batterStats[side] && G.batterStats[side][i]) G.batterStats[side][i].fielded = true;
  }
}

// 8守備位置を充足する割当(マッチング)を実際に適用する。
//   - 残留する現役占有者: 担当守備位置を更新 (スタメン同士のポジション入替)
//   - マッチングから外れた占有者: 控え(守備固め)に置き換え
function applyDefenseMatching(side) {
  const setup = G.setup[side];
  const m = findFieldingMatching(side);
  if (!m) return false;  // 充足不能 (選択時チェックで防がれている想定)
  const { slotList, positions, poolPlayers } = m;
  const nSlots = slotList.length;
  // 守備DRSを最大化しつつ、無用な入替を避け、打力(スタメン)も温存する重み付き割当を求める
  const curPos = new Map();       // スタメン占有者 → 現在の守備位置
  const starterSet = new Set();   // スタメン(現役占有者)集合
  slotList.forEach(s => { curPos.set(s.player, s.pos); starterSet.add(s.player); });
  const scoreFn = (player, pos) => {
    if (!canPlay(player, pos)) return -Infinity;
    let s = defenseRating(player, pos) * 10;        // 守備の総合評価(捕手はリード・阻止率込み)が主軸
    if (starterSet.has(player)) s += 50;            // スタメン維持(打力温存)。控え投入は明確に守備が良い時のみ
    if (curPos.get(player) === pos) s += 15;        // 同じ守備位置を維持(無用な入替を防ぐ)
    return s;
  };
  let assign = bestAssignment(positions, poolPlayers, scoreFn);
  if (!assign) assign = m.assign;   // フォールバック(成立性のみのマッチング)
  if (!assign) return false;
  // 割当を分類: 残留占有者(slotListのidx→新守備位置) / 投入控え(選手+守備位置)
  const keptOcc = new Map();    // slotList index → 新守備位置
  const incomingBench = [];     // { player, pos }
  for (let k = 0; k < positions.length; k++) {
    const pi = assign[k];
    if (pi < nSlots) keptOcc.set(pi, positions[k]);
    else incomingBench.push({ player: poolPlayers[pi], pos: positions[k] });
  }
  // 残留占有者: 守備位置を更新 (打順スロットはそのまま=スワップ)
  for (const [si, pos] of keptOcc) {
    const slot = slotList[si].idx;
    const oldPos = setup.batterPos[slot];
    setup.batterPos[slot] = pos;
    if (G.batterStats[side] && G.batterStats[side][slot]) G.batterStats[side][slot].position = pos;
    if (oldPos !== pos) {
      const pl = setup.batters[slot];
      logLine(`🔁 守備位置変更: ${pl ? (pl.fullNameTop || pl.nameJa) : '?'} (${POSITIONS[oldPos] ? POSITIONS[oldPos].label : oldPos}→${POSITIONS[pos] ? POSITIONS[pos].label : pos})`, 'event-inning');
    }
  }
  // 外れた占有スロット (= 控えに置き換え)
  const freedSlots = [];
  for (let si = 0; si < nSlots; si++) if (!keptOcc.has(si)) freedSlots.push(slotList[si].idx);
  for (let j = 0; j < incomingBench.length && j < freedSlots.length; j++) {
    const slot = freedSlots[j];
    const { player, pos } = incomingBench[j];
    const out = setup.batters[slot];
    const bi = (setup.benchAvail || []).indexOf(player);
    if (bi >= 0) setup.benchAvail.splice(bi, 1);
    setup.batters[slot] = player;
    setup.batterPos[slot] = pos;
    recordSubSwap(side, slot, player, '守備');
    logLine(`🧤 守備固め: ${out ? out.fullNameTop : '?'} → ${player.fullNameTop} (${POSITIONS[pos] ? POSITIONS[pos].label : pos})`, 'event-inning');
  }
  return true;
}

// リードを守る場面 (8回以降・1〜3点リード) での守備固め検討。
// 守備が「総合的に」良くなり (捕手はリード・阻止率も込み)、かつ打力低下が見合う場合のみ
// 1件だけ交代する。守備が良くならない交代 (ダウングレード) や、主砲級の打者を
// わずかな守備改善で外すような不利な交代は行わない。
function maybeDefensiveUpgrade(side) {
  const setup = G.setup[side];
  if (!setup.benchAvail || setup.benchAvail.length === 0) return;
  if (G.inning < 8) return;
  const lead = leadFor(side);
  if (lead < 1 || lead > 3) return;
  // 明確なネット改善のみ採用。9回以降は外した打者に打席が回りにくい(打力低下のコストが小さい)ため基準を緩める
  let bestNet = (G.inning >= 9) ? 1.5 : 3, bestSlot = -1, bestDef = null;
  for (let i = 0; i < 9; i++) {
    const pos = setup.batterPos[i];
    if (pos === 'DH') continue;
    const cur = setup.batters[i];
    const curDef = defenseRating(cur, pos);
    const curOff = offValue(cur);
    for (const p of setup.benchAvail) {
      if (!canPlay(p, pos)) continue;
      const defGain = defenseRating(p, pos) - curDef;
      if (defGain <= 0) continue;                                 // 守備が良くならないなら却下(ダウングレード防止)
      const offLoss = Math.max(0, curOff - offValue(p)) * 0.12;   // 打力低下を守備換算で差し引く(主砲を守るための重み)
      const net = defGain - offLoss;
      if (net > bestNet) { bestNet = net; bestSlot = i; bestDef = p; }
    }
  }
  if (bestSlot >= 0 && bestDef) {
    const di = setup.benchAvail.indexOf(bestDef);
    setup.benchAvail.splice(di, 1);
    const out = setup.batters[bestSlot];
    const pos = setup.batterPos[bestSlot];
    setup.batters[bestSlot] = bestDef;
    recordSubSwap(side, bestSlot, bestDef, '守備');
    logLine(`🧤 守備固め: ${out ? out.fullNameTop : '?'} → ${bestDef.fullNameTop} (${POSITIONS[pos].label})`, 'event-inning');
  }
}

// これまでにプレーした (もしくは予定された) イニング数。延長戦に対応
function playedInnings() {
  return Math.max(G.innings || 9, G.score.away.length, G.score.home.length);
}
// スコア配列を指定イニング数まで 0 埋めで拡張 (延長戦用)
function ensureScoreInning(inning) {
  while (G.score.away.length < inning) G.score.away.push(0);
  while (G.score.home.length < inning) G.score.home.push(0);
}
// タイブレーク: 10回以降は無死2塁から開始。2塁走者は前の打者(打順を1つ戻した選手)
// この走者は失点しても投手の自責点には含めない (ghost フラグ)
function placeTiebreakRunnerIfNeeded() {
  if (G.inning < 10) return;
  const battingSide = G.top ? 'away' : 'home';
  const leadoffIdx  = G.top ? G.awayBatIdx : G.homeBatIdx;
  const ghostIdx    = (leadoffIdx - 1 + 9) % 9;  // 先頭打者の前の打者
  G.bases[1] = { side: battingSide, slotIdx: ghostIdx, ghost: true };
  const runner = getRunnerPlayer(G.bases[1]);
  const rn = runner ? (runner.fullNameTop || runner.nameJa || '走者') : '走者';
  logLine(`🏃 タイブレーク: 無死2塁 (走者 ${rn}) からスタート`, 'event-inning');
}

// この半イニングを守り終えた守備側の現役選手に、守備位置ごとのイニングを+1する。
// (DHは守備につかないが「出場イニング」の指標として同様に+1 → 無交代フル出場で1試合9回相当)
function creditFieldingInning() {
  if (!G.seasonMode || !G.fieldInn) return;
  const defSide = G.top ? 'home' : 'away';   // この半回を守っていた側
  const arr = G.batterStats[defSide] || [];
  for (let i = 0; i < arr.length; i++) {
    const st = arr[i];
    if (!st || !st.player) continue;
    const pos = st.position;
    if (!pos || pos === '-' || pos === 'P' || pos === '投手') continue;   // 投手は集計対象外
    const k = playerKey(st.player);
    if (!G.fieldInn[k]) G.fieldInn[k] = {};
    G.fieldInn[k][pos] = (G.fieldInn[k][pos] || 0) + 1;
  }
}

function switchInning() {
  creditFieldingInning();   // 守り終えた半回の守備イニングを記録 (シーズンのみ)
  G.outs = 0;
  G.bases = [null, null, null];
  if (G.top) {
    G.top = false;
    // 最終回(9回以降)の裏に入る時点で HOME がリードしていれば、裏は不要 → 試合終了。
    // (ここで AWAY の継投を行うと「投球しない投手」のログが残るため、その前に終了する)
    const sAx = G.score.away.reduce((a,b)=>a+b,0);
    const sHx = G.score.home.reduce((a,b)=>a+b,0);
    if (G.inning >= G.innings && sHx > sAx) { G.homeSkipBottomIdx = G.inning - 1; finishGame(); return; }
    // 裏は AWAY 投手陣の現役投手
    G.currentPitcher = G.setup.away.pitchers[G.setup.away.activeIdx];
    G.currentBatter  = G.setup.home.batters[G.homeBatIdx];
    logLine(`══ ${G.inning}回裏 ══`, 'event-inning');
    // 守備側 (AWAY) のイニング頭継投検討 + 守備確定 (守備固め/予約適用)
    manageStarterWorkload('away');      // 先発の降板判断 (回終了時点)
    manageRelieverWorkload('away');     // リリーフは基本1イニング (登板過多の抑制)
    considerHighLeverageRelief('away');
    resolveDefense('away');
    G.currentPitcher = G.setup.away.pitchers[G.setup.away.activeIdx];
    placeTiebreakRunnerIfNeeded();
  } else {
    // 裏終了時点の決着判定 (延長戦対応)
    const sA = G.score.away.reduce((a,b)=>a+b,0);
    const sH = G.score.home.reduce((a,b)=>a+b,0);
    if (G.inning >= G.innings) {
      // 正規回 (9回) 以降で勝負がついていれば終了
      if (sA !== sH) { finishGame(); return; }
      // 同点でも最大延長回 (15回) に到達していたら引き分けで終了
      if (G.inning >= G.maxInnings) { finishGame(); return; }
    }
    G.inning++;
    ensureScoreInning(G.inning);  // 延長回のスコア欄を確保
    G.top = true;
    G.currentPitcher = G.setup.home.pitchers[G.setup.home.activeIdx];
    G.currentBatter  = G.setup.away.batters[G.awayBatIdx];
    logLine(`══ ${G.inning}回表 ══`, 'event-inning');
    // 守備側 (HOME) のイニング頭継投検討 + 守備確定 (守備固め/予約適用)
    manageStarterWorkload('home');      // 先発の降板判断 (回終了時点)
    manageRelieverWorkload('home');     // リリーフは基本1イニング (登板過多の抑制)
    considerHighLeverageRelief('home');
    resolveDefense('home');
    G.currentPitcher = G.setup.home.pitchers[G.setup.home.activeIdx];
    placeTiebreakRunnerIfNeeded();
  }
}

function checkEnd() {
  // 最終回(以降)の裏に HOME が勝ち越したら試合終了 = サヨナラ勝ち。
  //   この回は「攻撃せず勝った(X)」ではなく実際に得点しているので、スコアは「得点+x」表示にする。
  const sA = G.score.away.reduce((a,b)=>a+b,0);
  const sH = G.score.home.reduce((a,b)=>a+b,0);
  if (!G.top && G.inning >= G.innings) {
    if (sH > sA) { G.homeWalkoffIdx = G.inning - 1; finishGame(); return; }
  }
}

// 試合終了。ただし結果画面へは直行せず、ここで一旦停止する。
// ユーザーが赤い「🏁 試合結果へ」ボタンを押したら showResultScreen() が結果/戦評へ進む。
function finishGame() {
  if (G.awaitingResult) return;  // 二重呼び出しガード
  G.ended = true;
  G.autoToEnd = false;
  stopAutoVideo();               // 自動再生モードを停止 (カウントダウンも消す)
  G.awaitingResult = true;
  logLine('🏁 試合終了！ 「🏁 試合結果へ」ボタンで結果・戦評を表示します', 'event-inning');
  updateAutoFinishButton();
  renderAll();
}

// 勝敗・セーブ・ホールドの判定をまとめて算出する。
// 返り値の pitcherRoles は「投手ログ(stint) → 'W'/'L'/'S'/'H'」の Map。
// 結果画面と試合画面の投手ボード(終了後のW/L/H/S表示)で共用する。
function computePitcherDecisions() {
  const sA = G.score.away.reduce((a,b)=>a+b,0);
  const sH = G.score.home.reduce((a,b)=>a+b,0);
  const isDraw = (sA === sH);
  const winSide  = sA > sH ? 'away' : (sH > sA ? 'home' : null);
  const loseSide = winSide === 'away' ? 'home' : (winSide === 'home' ? 'away' : null);
  let winPitcher = null, losePitcher = null, savePitcher = null;
  if (winSide) {
    // 勝者側がリードを取って最後まで維持した最初の打席 (決勝点) を leadHistory から探す
    let firstLeadIdx = -1;
    for (let i = G.leadHistory.length - 1; i >= 0; i--) {
      if (G.leadHistory[i].leadSide === winSide) {
        if (i === 0 || G.leadHistory[i-1].leadSide !== winSide) { firstLeadIdx = i; break; }
      }
    }
    if (firstLeadIdx >= 0) {
      const h = G.leadHistory[firstLeadIdx];
      const winPitcherName = winSide === 'away' ? h.awayPitcher : h.homePitcher;
      winPitcher = G.pitcherLog[winSide].find(p => p.pitcher.fullNameTop === winPitcherName) || G.pitcherLog[winSide][0];
      const losePitcherName = loseSide === 'away' ? h.awayPitcher : h.homePitcher;
      losePitcher = G.pitcherLog[loseSide].find(p => p.pitcher.fullNameTop === losePitcherName) || G.pitcherLog[loseSide][0];
    } else {
      winPitcher  = G.pitcherLog[winSide][0];
      losePitcher = G.pitcherLog[loseSide][0];
    }
    // セーブ投手: 勝者側の最終投手で、先発でなく、終了時のリード差が3点以内
    // 実際に登板した(打者と対戦した)投手のうち最後の1人をセーブ候補に
    const facedWin = G.pitcherLog[winSide].filter(l => (l.battersFaced || 0) > 0 || (l.outs || 0) > 0);
    const lastWin = facedWin[facedWin.length - 1];
    const finalDiff = Math.abs(sA - sH);
    if (lastWin && lastWin !== G.pitcherLog[winSide][0] && lastWin !== winPitcher && finalDiff <= 3) {
      savePitcher = lastWin;
    }
  }
  // ホールド投手: 勝者側の中継ぎで、勝利/セーブ/先発のいずれでもなく、1アウト以上投げた者
  const holdPitchers = [];
  if (winSide) {
    const starter = G.pitcherLog[winSide][0];
    for (const lg of G.pitcherLog[winSide]) {
      if (lg === starter || lg === winPitcher || lg === savePitcher) continue;
      if ((lg.outs || 0) < 1) continue;
      holdPitchers.push(lg);
    }
  }
  const pitcherRoles = new Map();
  if (winPitcher)  pitcherRoles.set(winPitcher,  'W');
  if (losePitcher) pitcherRoles.set(losePitcher, 'L');
  if (savePitcher) pitcherRoles.set(savePitcher, 'S');
  for (const h of holdPitchers) pitcherRoles.set(h, 'H');
  return { sA, sH, isDraw, winSide, loseSide, winPitcher, losePitcher, savePitcher, holdPitchers, pitcherRoles };
}

// ===== レギュラーシーズン用: 試合結果の付加表示ヘルパー =====
// 勝敗投手の通算成績「10勝4敗」/ セーブ投手「22セーブ」。集計は試合終了後(seasonRecordCurrentGame)に走るため、
// 結果画面時点では当該試合分を未反映 → 既存累計 + 今試合分(該当役割を+1) で“この試合後の通算”を表示する。
function seasonPitcherRecordText(logEntry, role) {
  if (!G.seasonMode || !logEntry || !logEntry.pitcher) return '';
  const k = playerKey(logEntry.pitcher);
  const s = seasonActiveStores().pit[k] || {};   // ポストシーズンはポストシーズン通算
  let W = s.W || 0, L = s.L || 0, Sv = s.S || 0;
  if (role === 'W') W += 1; else if (role === 'L') L += 1; else if (role === 'S') Sv += 1;
  if (role === 'S') return `${Sv}セーブ`;
  return `${W}勝${L}敗`;
}
// 本塁打の点数 → 丸囲み数字 (①ソロ / ②2ラン / ③3ラン / ④満塁)
const SEASON_HR_CIRCLE = ['', '①', '②', '③', '④'];
function seasonHrType(runs) { return runs >= 4 ? '満塁' : runs === 3 ? '3ラン' : runs === 2 ? '2ラン' : 'ソロ'; }
// この試合の各本塁打イベントに「通算号数」を割り当てる。
//   結果画面時点では当該試合分が SEASON.bat に未反映 → 既存通算 + 試合内の累積打順 で号数を確定。
function seasonHrNumbers() {
  const nums = [];
  const counter = {};
  const store = seasonActiveStores().bat;   // ポストシーズンはポストシーズン通算号数
  (G.hrEvents || []).forEach((e, i) => {
    const k = e.batterKey;
    const prior = (store[k] && store[k].HR) || 0;
    counter[k] = (counter[k] || 0) + 1;
    nums[i] = prior + counter[k];
  });
  return nums;
}
// 結果画面の本塁打一覧 (例: 大谷翔平33号(2回裏②))。シーズンのみ。
function seasonHrSummaryHtml() {
  if (!G.seasonMode || !G.hrEvents || !G.hrEvents.length) return '';
  const nums = seasonHrNumbers();
  const items = G.hrEvents.map((e, i) => {
    const c = SEASON_HR_CIRCLE[Math.min(4, e.runs || 1)] || '';
    return `<span class="rp-hr">${e.batterName || e.batter}${nums[i]}号(${e.inning}回${e.top ? '表' : '裏'}${c})</span>`;
  });
  return `<div class="result-hr-row"><span class="rp-hrlabel">🏟️ 本塁打</span>${items.join('')}</div>`;
}
// 結果画面の本塁打一覧 (エキシビション用。例: 大谷翔平(LAD) 2回裏②)。通算号数はシーズンでないため付けない。
function exhHrSummaryHtml() {
  if (!G.hrEvents || !G.hrEvents.length) return '';
  const items = G.hrEvents.map((e) => {
    const c = SEASON_HR_CIRCLE[Math.min(4, e.runs || 1)] || '';
    const team = e.batterTeam ? `(${e.batterTeam})` : '';
    return `<span class="rp-hr">${e.batterName || e.batter}${team} ${e.inning}回${e.top ? '表' : '裏'}${c}</span>`;
  });
  return `<div class="result-hr-row"><span class="rp-hrlabel">🏟️ 本塁打</span>${items.join('')}</div>`;
}

// 結果画面 (スコアボード / 戦評 / 打撃・投手成績) を構築して表示する
function showResultScreen() {
  G.awaitingResult = false;
  G.ended = true;
  updateAutoFinishButton();
  // シーズン(手動)なら結果画面のボタンを「セーブして次の試合 / シーズンへ戻る」に切替。
  //   エキシビション(通常試合)は「セットアップに戻る」+「モード選択へ戻る」を表示し、シーズン用2ボタンは隠す。
  const isSeasonManual = !!(G.seasonMode && G.seasonCtx && !G.seasonCtx.auto);
  const bSetup = document.querySelector('#backToSetup');
  const bStart = document.querySelector('#result-to-start');
  const bNext = document.querySelector('#season-next');
  const bHub = document.querySelector('#season-tohub');
  if (bSetup) bSetup.hidden = isSeasonManual;
  if (bStart) bStart.hidden = isSeasonManual;   // モード選択へ戻る: エキシビションのみ
  if (bNext) bNext.hidden = !isSeasonManual;
  if (bHub) bHub.hidden = !isSeasonManual;
  const { sA, sH, isDraw, winSide, loseSide, winPitcher, losePitcher, savePitcher, holdPitchers, pitcherRoles } = computePitcherDecisions();

  let title;
  if (winSide === 'away')      title = `🏆 ${labelTeam('away')} (AWAY) の勝利！`;
  else if (winSide === 'home') title = `🏆 ${labelTeam('home')} (HOME) の勝利！`;
  else                          title = `🤝 引き分け`;
  $('#result-title').textContent = title;
  // 結果点数を勝利行のセンターに配置 (例: "NYY 4-6 LAD")
  const scoreEl = $('#result-score-line');
  if (scoreEl) {
    scoreEl.innerHTML = `<span class="rsl-team">${labelTeam('away')}</span>` +
      `<span class="rsl-score">${sA}-${sH}</span>` +
      `<span class="rsl-team">${labelTeam('home')}</span>`;
  }

  // 勝利・敗戦・セーブ投手 (1行コンパクト)。シーズンは名前の後ろに通算成績「10勝4敗」/「22セーブ」を付ける。
  let pitcherHtml = '';
  if (!isDraw) {
    const rec = (logEntry, role) => { const t = seasonPitcherRecordText(logEntry, role); return t ? ` <span class="rp-rec">${t}</span>` : ''; };
    const winName  = winPitcher  ? `${winPitcher.pitcher.fullNameTop} (${winPitcher.pitcher.team})${rec(winPitcher, 'W')}`   : '';
    const loseName = losePitcher ? `${losePitcher.pitcher.fullNameTop} (${losePitcher.pitcher.team})${rec(losePitcher, 'L')}` : '';
    const saveName = savePitcher ? `${savePitcher.pitcher.fullNameTop} (${savePitcher.pitcher.team})${rec(savePitcher, 'S')}` : '';
    pitcherHtml = `
      <div class="result-pitcher-row">
        ${winName  ? `<span class="rp-item"><span class="tag tag-win">勝利</span>${winName}</span>` : ''}
        ${loseName ? `<span class="rp-item"><span class="tag tag-loss">敗戦</span>${loseName}</span>` : ''}
        ${saveName ? `<span class="rp-item"><span class="tag tag-save">セーブ</span>${saveName}</span>` : ''}
      </div>
    `;
  }
  // 本塁打一覧 (シーズン=通算号数付き / エキシビション=チーム名付き)
  const hrSummaryHtml = G.seasonMode ? seasonHrSummaryHtml() : exhHrSummaryHtml();

  // 戦評 (スポーツ新聞風) を生成
  const reviewText = generateGameReview({
    winSide, loseSide, sA, sH, isDraw,
    winPitcher, losePitcher, savePitcher, holdPitchers,
  });
  const reviewHtml = reviewText ? `
    <div class="game-review">
      <h3>📰 戦評</h3>
      <p>${reviewText}</p>
    </div>
  ` : '';

  $('#result-summary').innerHTML = `
    ${renderResultScoreboard()}
    ${pitcherHtml}
    ${hrSummaryHtml}
    ${reviewHtml}
    ${renderBattingTable('away')}
    ${renderBattingTable('home')}
    ${renderPitchingTable('away', pitcherRoles)}
    ${renderPitchingTable('home', pitcherRoles)}
  `;
  showScreen('result');
}

// 結果画面用のイニング別スコアボード (試合中のものと同じ形式)
function renderResultScoreboard() {
  const innings = playedInnings();
  let inningHeaders = '';
  for (let i = 1; i <= innings; i++) inningHeaders += `<th>${i}</th>`;
  const awayCells = [];
  const homeCells = [];
  for (let i = 0; i < innings; i++) {
    awayCells.push(`<td>${G.score.away[i] ?? 0}</td>`);
    // 後攻が攻撃せず勝った回は「X」/ サヨナラ勝ちの回は「得点+x」表示 (野球サイト方式)
    homeCells.push(`<td>${G.homeSkipBottomIdx === i ? 'X' : (G.homeWalkoffIdx === i ? ((G.score.home[i] ?? 0) + 'x') : (G.score.home[i] ?? 0))}</td>`);
  }
  const totalA = G.score.away.reduce((a, b) => a + b, 0);
  const totalH = G.score.home.reduce((a, b) => a + b, 0);
  return `
    <div class="result-scoreboard">
      <table class="${innings > 9 ? 'extra' : ''}">
        <thead>
          <tr>
            <th class="rsb-team">チーム</th>
            ${inningHeaders}
            <th class="rsb-total-th">計</th>
            <th>H</th>
            <th>K</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <th class="rsb-team">${labelTeam('away')} (AWAY)</th>
            ${awayCells.join('')}
            <td class="rsb-total">${totalA}</td>
            <td>${G.hits.away}</td>
            <td>${G.ks.away}</td>
          </tr>
          <tr>
            <th class="rsb-team">${labelTeam('home')} (HOME)</th>
            ${homeCells.join('')}
            <td class="rsb-total">${totalH}</td>
            <td>${G.hits.home}</td>
            <td>${G.ks.home}</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

// =============================================================
// 戦評生成 (スポーツ新聞風の試合振り返り文 — 3行程度の充実版)
// =============================================================
function generateGameReview({ winSide, loseSide, sA, sH, isDraw, winPitcher, losePitcher, savePitcher, holdPitchers }) {
  if (isDraw) {
    return `両軍譲らず、${sA}-${sH}の引き分けに終わった。両チームとも得点機を作りながら決定打を欠き、終盤までもつれる展開となった。スコアボードに刻まれた数字以上に、熱戦の余韻を残す一戦となった。`;
  }

  const winTeam  = labelTeam(winSide);
  const loseTeam = labelTeam(loseSide);
  const winRuns  = winSide === 'away' ? sA : sH;
  const loseRuns = winSide === 'away' ? sH : sA;
  const diff = winRuns - loseRuns;
  const winHits = G.hits[winSide] || 0;
  const winHRCount = (G.hrEvents || []).filter(e => e.side === winSide).length;

  // 各打席のリード状態を解析
  let wasBehind = false;
  let maxDeficit = 0;
  let leadChanges = 0, prevLead = null;
  for (const h of (G.leadHistory || [])) {
    if (h.leadSide && h.leadSide !== prevLead) { leadChanges++; prevLead = h.leadSide; }
    if (h.leadSide && h.leadSide !== winSide) {
      wasBehind = true;
      if ((h.diff || 0) > maxDeficit) maxDeficit = h.diff || 0;
    }
  }

  // ===== 特別な状況の検出 (タイブレーク延長 / サヨナラ勝ち) =====
  const finalInning = playedInnings();
  const extraInnings = finalInning > 9;                  // 延長(10回以降=タイブレーク)
  const finalIdx = finalInning - 1;
  const homeRunsFinal = G.score.home[finalIdx] || 0;     // 最終回裏のホーム得点
  const homeBeforeFinal = sH - homeRunsFinal;            // 最終回裏に入る前のホーム得点
  // サヨナラ: 後攻(home)が最終回裏に勝ち越して試合を終わらせた
  const walkOff = (winSide === 'home') && homeRunsFinal > 0 && homeBeforeFinal <= sA;
  const comebackWalkOff = walkOff && homeBeforeFinal < sA;  // 逆転サヨナラ (それまでビハインド)
  const totalRuns = sA + sH;

  // ===== 1. 試合パターンの導入文 (特別な状況を最優先で拾う) =====
  let opening;
  if (walkOff && extraInnings)           opening = `${winTeam}が延長タイブレークの死闘を制し、劇的なサヨナラ${comebackWalkOff ? '逆転' : ''}勝ちを飾った。`;
  else if (walkOff)                      opening = comebackWalkOff ? `${winTeam}が${finalInning}回裏、土壇場のサヨナラ逆転勝ちを収めた。` : `${winTeam}がサヨナラ勝ちで劇的な幕切れを呼び込んだ。`;
  else if (extraInnings)                 opening = `延長タイブレークまでもつれた死闘を、${winTeam}が${finalInning}回に制した。`;
  else if (wasBehind && maxDeficit >= 3) opening = `${winTeam}が劇的な逆転勝利を飾った。`;
  else if (wasBehind && diff <= 1)       opening = `${winTeam}が逆転で接戦を制した。`;
  else if (wasBehind)                    opening = `${winTeam}が逆転勝利を収めた。`;
  else if (loseRuns === 0)               opening = `${winTeam}が完封勝利。`;
  else if (diff >= 6)                    opening = `${winTeam}が${loseTeam}を圧倒した。`;
  else if (diff <= 1)                    opening = `${winTeam}が接戦をモノにした。`;
  else                                   opening = `${winTeam}が${loseTeam}を退けた。`;

  // 戦況コンテキスト
  let context = '';
  if (extraInnings || walkOff)         context = (totalRuns >= 8) ? '両軍譲らぬ点の取り合いとなり、' : '緊迫した攻防が最後まで続き、';
  else if (leadChanges >= 3)           context = '両軍リードを奪い合う一進一退の攻防となったが、';
  else if (!wasBehind && winRuns >= 5) context = '序盤から主導権を握り、';
  else if (winHRCount >= 2)            context = '本塁打攻勢で打線が爆発し、';
  else if (diff >= 4)                  context = '終始安定した試合運びで、';

  // ===== 2. 決定的イニング (= 勝者が最後までキープしたリードを取った瞬間) =====
  let turnInning = -1, turnTop = null;
  if (G.leadHistory) {
    for (let i = 0; i < G.leadHistory.length; i++) {
      const h = G.leadHistory[i];
      if (h.leadSide !== winSide) continue;
      const prev = i > 0 ? G.leadHistory[i-1].leadSide : null;
      if (prev === winSide) continue;
      let stayed = true;
      for (let j = i + 1; j < G.leadHistory.length; j++) {
        if (G.leadHistory[j].leadSide && G.leadHistory[j].leadSide !== winSide) {
          stayed = false; break;
        }
      }
      if (stayed) { turnInning = h.inning; turnTop = h.top; break; }
    }
  }

  // ヘルパー: あるイニングの勝者側英雄を抽出
  const collectHeroes = (side, iIdx) => {
    const heroes = [];
    // 現役選手 + 交代で退いた選手 (subLog) の両方を対象にする
    const stats = (G.batterStats[side] || []).concat(G.subLog[side] || []);
    for (const s of stats) {
      const txt = s.perInning[iIdx] || '';
      if (!txt) continue;
      const player = s.player.fullNameTop;
      if (txt.includes('本塁打'))      heroes.push({ name: player, type: '本塁打', priority: 3 });
      else if (txt.includes('三塁打')) heroes.push({ name: player, type: '三塁打', priority: 2 });
      else if (txt.includes('二塁打')) heroes.push({ name: player, type: '二塁打', priority: 2 });
      else if (txt.includes('安打'))   heroes.push({ name: player, type: '適時打', priority: 1 });
    }
    heroes.sort((a, b) => b.priority - a.priority);
    return heroes;
  };

  // ===== 3. 決定的イニングの英雄 / サヨナラの場面を文章化 =====
  let heroPart = '';
  if (walkOff) {
    // サヨナラ: 最終回裏の決勝打を主役として描写
    const heroes = collectHeroes(winSide, finalIdx);
    if (heroes.length >= 1) {
      const h0 = heroes[0];
      const sayonara = (h0.type === '本塁打') ? 'サヨナラ本塁打' : `サヨナラ${h0.type}`;
      heroPart = `${finalInning}回裏、${h0.name}の${sayonara}で試合に終止符を打った。`;
    } else {
      heroPart = `${finalInning}回裏、ついに均衡を破ってサヨナラ勝ちを決めた。`;
    }
  } else if (turnInning > 0) {
    const iIdx = turnInning - 1;
    const inningScore = G.score[winSide][iIdx] || 0;
    const tb = turnTop ? '表' : '裏';
    const heroes = collectHeroes(winSide, iIdx);
    if (heroes.length >= 2) {
      heroPart = `${turnInning}回${tb}、${heroes[0].name}の${heroes[0].type}と${heroes[1].name}の${heroes[1].type}などで${inningScore}得点を奪い、試合の流れを引き寄せた。`;
    } else if (heroes.length === 1) {
      heroPart = `${turnInning}回${tb}、${heroes[0].name}の${heroes[0].type}を皮切りに${inningScore}点を挙げてリードを奪った。`;
    } else if (inningScore > 0) {
      heroPart = `${turnInning}回${tb}に${inningScore}得点を奪ってリードを掴んだ。`;
    }
  }

  // ===== 4. 決定的イニング以外のビッグイニングがあれば追記 (サヨナラ時は突き放し表現を避ける) =====
  let bigPart = '';
  let bigInning = -1, bigInningRuns = 0;
  for (let i = 0; i < playedInnings(); i++) {
    const runs = G.score[winSide][i] || 0;
    if (runs >= 3 && runs > bigInningRuns && (i + 1) !== turnInning) {
      bigInning = i + 1; bigInningRuns = runs;
    }
  }
  if (bigInning > 0 && !walkOff) {
    const tb = winSide === 'away' ? '表' : '裏';
    if (bigInning > turnInning) {
      bigPart = `さらに${bigInning}回${tb}にも${bigInningRuns}点を加えて突き放した。`;
    } else {
      bigPart = `${bigInning}回${tb}には${bigInningRuns}点のビッグイニングを作る攻撃も光った。`;
    }
  }

  // ===== 5. 勝者側のトップヒッターを称賛 =====
  const topHitter = (side) => {
    const stats = (G.batterStats[side] || []).concat(G.subLog[side] || []);
    let best = null, score = -1;
    for (const s of stats) {
      if (s.H === 0 && s.RBI === 0 && s.HR === 0) continue;
      const sc = s.H * 10 + s.RBI * 6 + s.HR * 10;
      if (sc > score) { score = sc; best = s; }
    }
    return best;
  };
  let topBatterPart = '';
  const winHero = topHitter(winSide);
  if (winHero && (winHero.H >= 3 || winHero.HR >= 1 || winHero.RBI >= 3)) {
    const parts = [];
    if (winHero.H >= 3)   parts.push(`${winHero.H}安打`);
    if (winHero.HR >= 1)  parts.push(`${winHero.HR}本塁打`);
    if (winHero.RBI >= 2) parts.push(`${winHero.RBI}打点`);
    if (parts.length > 0) {
      topBatterPart = `打線では${winHero.player.fullNameTop}が${parts.join('・')}と気を吐き、勝利を引き寄せた。`;
    }
  }

  // ===== 6. 勝利投手の見どころ =====
  let winPitcherPart = '';
  if (winPitcher) {
    const isStarter = winPitcher === G.pitcherLog[winSide][0];
    const role = isStarter ? '先発' : '中継ぎ';
    const ip = fmtIP(winPitcher.outs);
    const er = winPitcher.earnedRuns;
    const k  = winPitcher.K || 0;
    let pitcherDesc;
    if (er === 0 && k >= 5)      pitcherDesc = `${ip}回無失点・${k}奪三振の圧巻投球`;
    else if (er === 0)           pitcherDesc = `${ip}回無失点の好投`;
    else if (k >= 6)             pitcherDesc = `${ip}回${er}失点ながら${k}奪三振の力投`;
    else                          pitcherDesc = `${ip}回${er}失点でゲームメイク`;
    if (G.seasonMode) {
      // シーズン: 「○回無失点で10勝目をマーク」のように勝ち数を具体化
      const winNum = (((SEASON && SEASON.pit && SEASON.pit[playerKey(winPitcher.pitcher)]) || {}).W || 0) + 1;
      const ipClean = (winPitcher.outs % 3 === 0) ? `${winPitcher.outs / 3}` : ip;   // 完全な回は「6回」表記
      const erTxt = (er === 0) ? '無失点' : `${er}失点`;
      const kTxt = (k >= 6) ? `・${k}奪三振` : '';
      winPitcherPart = `投げては${role}・${winPitcher.pitcher.fullNameTop}が${ipClean}回${erTxt}${kTxt}で${winNum}勝目をマークした。`;
    } else {
      winPitcherPart = `投げては${role}・${winPitcher.pitcher.fullNameTop}が${pitcherDesc}で勝ち星を手にした。`;
    }
  }

  // ===== 7. リリーフ陣の働き (ホールド + セーブ) =====
  let bullpenPart = '';
  const hCount = (holdPitchers || []).length;
  if (savePitcher && hCount > 0) {
    bullpenPart = `中継ぎ陣もリードを死守し、最後は${savePitcher.pitcher.fullNameTop}がセーブを締めた。`;
  } else if (savePitcher) {
    bullpenPart = `最後は${savePitcher.pitcher.fullNameTop}が試合を締めくくった。`;
  } else if (hCount > 0) {
    bullpenPart = `中継ぎ陣がしっかりとリードを守り抜いた。`;
  }

  // ===== 8. 敗戦投手 =====
  let losePart = '';
  if (losePitcher) {
    const isStarter = losePitcher === G.pitcherLog[loseSide][0];
    const role = isStarter ? '先発' : '中継ぎ';
    losePart = `敗れた${loseTeam}は${role}・${losePitcher.pitcher.fullNameTop}が${losePitcher.runsAllowed}失点と試合を作れなかった。`;
  }

  // ===== 9. 敗者側のせめてもの見せ場 (孤軍奮闘) =====
  let loseHeroPart = '';
  const loseHero = topHitter(loseSide);
  if (loseHero && (loseHero.HR >= 1 || loseHero.RBI >= 2 || loseHero.H >= 3)) {
    const lp = [];
    if (loseHero.H >= 2)  lp.push(`${loseHero.H}安打`);
    if (loseHero.HR >= 1) lp.push(`${loseHero.HR}本塁打`);
    if (loseHero.RBI >= 2) lp.push(`${loseHero.RBI}打点`);
    if (lp.length > 0) {
      loseHeroPart = `打線では${loseHero.player.fullNameTop}が${lp.join('・')}と一矢報いたが、反撃及ばず。`;
    }
  }

  // ===== 10. 特別な記録を拾う (ノーヒットノーラン / 満塁本塁打 / サイクル / 猛打賞 / 連続得点) =====
  const allStats = (s) => (G.batterStats[s] || []).concat(G.subLog[s] || []);
  const fullGame = playedInnings() >= 9;

  // ノーヒットノーラン: 相手を無安打無得点(9回以上)に抑えた
  let nohitPart = '';
  for (const defSide of ['away','home']) {
    const oppSide = defSide === 'away' ? 'home' : 'away';
    const oppRuns = oppSide === 'away' ? sA : sH;
    if (fullGame && (G.hits[oppSide] || 0) === 0 && oppRuns === 0) {
      const pl = G.pitcherLog[defSide] || [];
      nohitPart = (pl.length === 1)
        ? `${labelTeam(defSide)}・${pl[0].pitcher.fullNameTop}が${labelTeam(oppSide)}打線を無安打無得点に封じ、ノーヒットノーランの快挙を達成した。`
        : `${labelTeam(defSide)}投手陣が${labelTeam(oppSide)}を無安打無得点に抑え、継投でのノーヒットノーランを成し遂げた。`;
    }
  }

  // 満塁本塁打 (HRイベントで一度に4点)。シーズンは通算号数付き (例: 値千金の24号満塁本塁打)。
  let slamPart = '';
  const hrNums = G.seasonMode ? seasonHrNumbers() : null;
  const slamIdxs = [];
  (G.hrEvents || []).forEach((e, i) => { if ((e.runs || 0) >= 4) slamIdxs.push(i); });
  if (slamIdxs.length >= 2) {
    slamPart = `満塁本塁打が${slamIdxs.length}本も飛び出す一発攻勢となった。`;
  } else if (slamIdxs.length === 1) {
    const i = slamIdxs[0], s = G.hrEvents[i];
    const num = hrNums ? `${hrNums[i]}号` : '';
    slamPart = `${s.inning}回${s.top ? '表' : '裏'}には${s.batterName || s.batter}が値千金の${num}満塁本塁打を叩き込んだ。`;
  }
  // シーズン: 満塁の一発が無くても、大きな本塁打を号数付きで紹介 (例: 33号3ラン本塁打)
  let seasonHrPart = '';
  if (G.seasonMode && !slamPart && hrNums) {
    let best = -1, bestW = -1;
    (G.hrEvents || []).forEach((e, i) => {
      if ((e.runs || 0) < 2) return;   // ソロは対象外
      const w = (e.side === winSide ? 10 : 0) + (e.runs || 0);
      if (w > bestW) { bestW = w; best = i; }
    });
    if (best >= 0) {
      const e = G.hrEvents[best];
      seasonHrPart = `${e.inning}回${e.top ? '表' : '裏'}に飛び出した${e.batterName}の${hrNums[best]}号${seasonHrType(e.runs)}本塁打が試合を大きく動かした。`;
    }
  }

  // サイクルヒット (単打・二塁打・三塁打・本塁打を達成)
  let cyclePart = '';
  for (const side of ['away','home']) {
    for (const s of allStats(side)) {
      const singles = (s.H||0) - (s.doubles||0) - (s.triples||0) - (s.HR||0);
      if (singles >= 1 && (s.doubles||0) >= 1 && (s.triples||0) >= 1 && (s.HR||0) >= 1) {
        cyclePart = `${s.player.fullNameTop}が単打・二塁打・三塁打・本塁打を打ち分けるサイクルヒットの偉業を達成した。`;
      }
    }
  }

  // 猛打賞 (3安打以上)。ヘッドラインで既出の打者は除く
  const featuredNames = new Set();
  if (winHero && winHero.player)  featuredNames.add(winHero.player.fullNameTop);
  if (loseHero && loseHero.player) featuredNames.add(loseHero.player.fullNameTop);
  const mohit = [];
  for (const side of ['away','home']) {
    for (const s of allStats(side)) {
      if ((s.H||0) >= 3 && !featuredNames.has(s.player.fullNameTop)) mohit.push({ name: s.player.fullNameTop, H: s.H });
    }
  }
  let mohitPart = '';
  if (mohit.length) {
    mohit.sort((a,b)=>b.H-a.H);
    mohitPart = mohit.map(m=>`${m.name}(${m.H}安打)`).join('、') + 'も猛打賞をマークした。';
  }

  // 連続得点イニング (4イニング以上)
  let streakPart = '';
  for (const side of ['away','home']) {
    const arr = G.score[side] || [];
    let best = 0, bestEnd = -1, cur = 0;
    for (let i = 0; i < arr.length; i++) {
      if ((arr[i]||0) > 0) { cur++; if (cur > best) { best = cur; bestEnd = i; } }
      else cur = 0;
    }
    if (best >= 4) {
      streakPart = `${labelTeam(side)}は${bestEnd-best+2}回から${bestEnd+1}回まで${best}イニング連続で得点する勢いを見せた。`;
    }
  }

  const highlightsPart = slamPart + seasonHrPart + cyclePart + mohitPart + streakPart;

  return opening + context + heroPart + bigPart + highlightsPart + topBatterPart + winPitcherPart + nohitPart + bullpenPart + losePart + loseHeroPart;
}

// 打率を ".XXX" 形式で表示
function fmtAvg(h, ab) {
  if (!ab) return '.000';
  const a = h / ab;
  return (a >= 1 ? '1.000' : a.toFixed(3).replace(/^0/, ''));
}
// 投球イニング: 例 outs=10 -> 3.1
function fmtIP(outs) {
  if (!outs) return '0.0';
  return Math.floor(outs / 3) + '.' + (outs % 3);
}
// 防御率: ER*9/IP_decimal
function fmtERA(er, outs) {
  if (!outs) return '-.--';
  const ip = outs / 3;
  if (ip === 0) return '-.--';
  return (er * 9 / ip).toFixed(2);
}

function renderBattingTable(side) {
  const active = G.batterStats[side] || [];
  const subs = G.subLog[side] || [];
  const nInn = playedInnings();
  const teamLabel = labelTeam(side) + ' (' + (side === 'away' ? 'AWAY' : 'HOME') + ')';
  // 打順スロットごとに [退いた選手(subLog, 登板順)] → [現役選手] のチェーンを作り、
  // Yahoo方式で交代選手を元選手の下へ挿入表示する。
  const chains = [];
  for (let i = 0; i < 9; i++) {
    const chain = subs.filter(s => s.slotIdx === i);
    if (active[i]) chain.push(active[i]);
    chains.push(chain);
  }
  const allStints = [];
  chains.forEach(c => c.forEach(s => allStints.push(s)));
  // 合計算出 (死球/犠打/失策は除外。交代選手も含めた全選手を集計)
  const total = { AB:0, R:0, H:0, RBI:0, K:0, BB:0, SB:0, E:0, HR:0 };
  allStints.forEach(s => {
    total.AB += s.AB; total.R += s.R; total.H += s.H; total.RBI += s.RBI;
    total.K += s.K; total.BB += s.BB; total.SB += s.SB; total.E += (s.E || 0); total.HR += s.HR;
  });
  // 1選手分の行を生成 (isSub=交代選手なら控えめなスタイル + 「└」で挿入を表現)
  const renderStintRow = (s) => {
    const isSub = !!s.subRole;
    const cells = [];
    cells.push(`<td class="rt-pos">${stintPosLabel(s)}</td>`);
    cells.push(`<td class="rt-name player-link${longNameClass(s.player.fullNameTop)}" data-player-name="${s.player.fullNameTop}" data-player-year="${s.player.year ?? ''}" data-player-type="${playerType(s.player)}" data-player-team="${s.player.team||''}">${isSub ? '└ ' : ''}${s.player.fullNameTop}</td>`);
    cells.push(`<td>${isSeasonManualGame() ? seasonLiveAvg(s.player) : fmtAvg(s.H, s.AB)}</td>`);
    cells.push(`<td>${s.AB}</td>`);
    cells.push(`<td>${s.R}</td>`);
    cells.push(`<td>${s.H}</td>`);
    cells.push(`<td>${s.RBI}</td>`);
    cells.push(`<td>${s.K}</td>`);
    cells.push(`<td>${s.BB}</td>`);
    cells.push(`<td>${s.SB}</td>`);
    cells.push(`<td>${s.E || 0}</td>`);
    cells.push(`<td class="rt-hr">${s.HR}</td>`);
    for (let i = 0; i < nInn; i++) {
      const v = s.perInning[i] || '';
      let cls = '';
      if (v.includes('本塁打')) cls = 'cell-hr';
      else if (v.includes('安打') || v.includes('二塁打') || v.includes('三塁打')) cls = 'cell-hit';
      else if (v.includes('四球')) cls = 'cell-bb';
      else if (v.includes('三振')) cls = 'cell-k';
      cells.push(`<td class="rt-inning ${cls}">${v}</td>`);
    }
    return `<tr class="${isSub ? 'rt-sub' : ''}">${cells.join('')}</tr>`;
  };
  const rows = chains.map(chain => chain.map(renderStintRow).join('')).join('');
  const totalRow = `
    <tr class="rt-total">
      <td></td><td>合計</td><td>-</td>
      <td>${total.AB}</td><td>${total.R}</td><td>${total.H}</td>
      <td>${total.RBI}</td><td>${total.K}</td><td>${total.BB}</td>
      <td>${total.SB}</td><td>${total.E}</td><td class="rt-hr">${total.HR}</td>
      ${'<td></td>'.repeat(nInn)}
    </tr>`;
  // colgroup で列幅を上下チーム共通に固定
  const batColgroup = `
    <colgroup>
      <col style="width:42px">  <!-- 位置 -->
      <col style="width:110px"> <!-- 選手名 -->
      <col style="width:50px">  <!-- 打率 -->
      <col style="width:40px">  <!-- 打数 -->
      <col style="width:40px">  <!-- 得点 -->
      <col style="width:40px">  <!-- 安打 -->
      <col style="width:40px">  <!-- 打点 -->
      <col style="width:40px">  <!-- 三振 -->
      <col style="width:40px">  <!-- 四球 -->
      <col style="width:40px">  <!-- 盗塁 -->
      <col style="width:40px">  <!-- 失策 -->
      <col style="width:50px">  <!-- 本塁打 -->
      ${[...Array(nInn)].map(() => '<col>').join('')}
    </colgroup>`;
  return `
    <div class="result-section">
      <h3>🏏 打撃成績 — ${teamLabel}</h3>
      <div class="rt-wrap">
      <table class="result-table bat-table">
        ${batColgroup}
        <thead>
          <tr>
            <th>位置</th><th class="rt-name-th">選手名</th>
            <th>打率</th><th>打数</th><th>得点</th><th>安打</th>
            <th>打点</th><th>三振</th><th>四球</th>
            <th>盗塁</th><th>失策</th><th>本塁打</th>
            ${[...Array(nInn)].map((_,i) => `<th>${i+1}回</th>`).join('')}
          </tr>
        </thead>
        <tbody>${rows}${totalRow}</tbody>
      </table>
      </div>
    </div>
  `;
}

function renderPitchingTable(side, pitcherRoles) {
  const log = (G.pitcherLog[side] || []).filter(l => (l.battersFaced || 0) > 0 || (l.outs || 0) > 0);  // 登板していない投手は除外
  const setup = G.setup[side] || {};
  const teamLabel = labelTeam(side) + ' (' + (side === 'away' ? 'AWAY' : 'HOME') + ')';
  const rows = log.map(l => {
    const role = pitcherRoles && pitcherRoles.get(l);
    const roleBadge = role ? `<span class="role-badge role-${role}" title="${({W:'勝利投手',L:'敗戦投手',H:'ホールド投手',S:'セーブ投手'}[role])}">${role}</span>` : '';
    // スタミナ残量 / 基本(カード)スタミナ
    const pIdx = (setup.pitchers || []).indexOf(l.pitcher);
    const staVal = (pIdx >= 0 && setup.pitcherStamina) ? Math.round((setup.pitcherStamina[pIdx] ?? 0) * 10) / 10 : '-';
    const staMax = (pIdx >= 0 && setup.pitcherMax) ? setup.pitcherMax[pIdx] : '-';
    return `<tr>
      <td class="rt-badge">${roleBadge}</td>
      <td class="rt-name player-link${longNameClass(l.pitcher.fullNameTop)}" data-player-name="${l.pitcher.fullNameTop}" data-player-year="${l.pitcher.year ?? ''}" data-player-type="${playerType(l.pitcher)}" data-player-team="${l.pitcher.team||''}">${l.pitcher.fullNameTop}</td>
      <td>${isSeasonManualGame() ? seasonLiveERA(l.pitcher) : fmtERA(l.earnedRuns, l.outs)}</td>
      <td>${fmtIP(l.outs)}</td>
      <td class="rt-sta">${staVal}/${staMax}</td>
      <td>${l.battersFaced}</td>
      <td>${l.hits}</td>
      <td>${l.K}</td>
      <td>${l.BB || 0}</td>
      <td>${l.runsAllowed}</td>
      <td>${l.earnedRuns}</td>
      <td class="rt-spacer"></td>
    </tr>`;
  }).join('');
  // 打撃成績テーブルの「位置(42px)+選手名(110px)」列に揃え、
  // 投手成績側も「役割(42px)+選手名(88px)」で名前の書き出し位置を一致させる
  // 末尾に auto 幅のスペーサー列を追加して、宣言列が比例縮小されないようにする
  const pitColgroup = `
    <colgroup>
      <col style="width:42px">  <!-- 役割バッジ (打撃成績の 位置 列と同幅) -->
      <col style="width:88px">  <!-- 選手名 -->
      <col style="width:60px">  <!-- 防御率 -->
      <col style="width:54px">  <!-- 投球回 -->
      <col style="width:66px">  <!-- スタミナ(残量/基本) -->
      <col style="width:46px">  <!-- 打者 -->
      <col style="width:56px">  <!-- 被安打 -->
      <col style="width:56px">  <!-- 奪三振 -->
      <col style="width:46px">  <!-- 四球 -->
      <col style="width:46px">  <!-- 失点 -->
      <col style="width:56px">  <!-- 自責点 -->
      <col>                     <!-- auto: 残り幅を吸収 (列幅固定のため) -->
    </colgroup>`;
  return `
    <div class="result-section">
      <h3>⚾ 投手成績 — ${teamLabel}</h3>
      <div class="rt-wrap">
      <table class="result-table pit-table">
        ${pitColgroup}
        <thead>
          <tr>
            <th></th>
            <th class="rt-name-th">選手名</th>
            <th>防御率</th><th>投球回</th><th>スタミナ</th><th>打者</th>
            <th>被安打</th><th>奪三振</th><th>四球</th>
            <th>失点</th><th>自責点</th>
            <th class="rt-spacer-th"></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      </div>
    </div>
  `;
}

// ============== オートプレイ ==============
function autoPick() {
  // 投手のpitches から weighted random で1球選ぶ
  const ps = G.currentPitcher.pitches || [];
  if (!ps.length) return null;
  let totalW = 0;
  for (const p of ps) totalW += (p.ratio || 10);
  let r = rand() * totalW;
  for (const p of ps) {
    r -= (p.ratio || 10);
    if (r <= 0) return p;
  }
  return ps[0];
}

function runAutoInning() {
  if (G.ended) return;
  const startInning = G.inning, startTop = G.top;
  const step = () => {
    if (G.ended) return;
    if (G.inning !== startInning || G.top !== startTop) return; // イニングが変わったら停止
    const p = autoPick();
    if (!p) return;
    pitchOne(p, true);
    setTimeout(step, 50);
  };
  step();
}

function runAutoToEnd() {
  if (G.ended) return;
  G.autoToEnd = true;
  updateAutoFinishButton();
  const step = () => {
    // 試合終了 または ユーザーが停止を押したら抜ける
    if (G.ended || !G.autoToEnd) {
      G.autoToEnd = false;
      updateAutoFinishButton();
      return;
    }
    const p = autoPick();
    if (!p) {
      G.autoToEnd = false;
      updateAutoFinishButton();
      return;
    }
    pitchOne(p, true);
    setTimeout(step, 20);
  };
  step();
}
function stopAutoToEnd() {
  G.autoToEnd = false;
  updateAutoFinishButton();
}
function updateAutoFinishButton() {
  const btn = document.querySelector('#autoFinish');
  if (!btn) return;
  // 優先度: 結果待ち(赤) > 自動進行中(停止) > 通常
  if (G.awaitingResult) {
    btn.textContent = '🏁 試合結果へ';
    btn.classList.remove('is-running');
    btn.classList.add('is-result');
  } else if (G.autoToEnd) {
    btn.textContent = '■ 停止';
    btn.classList.remove('is-result');
    btn.classList.add('is-running');
  } else {
    btn.textContent = '⏩ 試合終了まで';
    btn.classList.remove('is-running', 'is-result');
  }
}

// ============== ログ ==============
function logLine(msg, cls) {
  const li = document.createElement('li');
  li.textContent = msg;
  if (cls) li.className = cls;
  const log = $('#log');
  log.appendChild(li);
  log.scrollTop = log.scrollHeight;
  // インフォメーション(イニング切替/投手交代/代打/盗塁/試合終了 等)は
  // ダイヤモンドの外野付近にもオレンジ太字で表示する。
  if (isInfoMessage(msg, cls)) {
    // 同じタイミング(同一打席)の通知は改行でまとめて全件表示する(複数交代/守備固め等)。
    // 次打席の開始時にまとめてクリアされる。
    if (G.infoNew && G.lastInfo) G.lastInfo = G.lastInfo + '\n' + msg;
    else G.lastInfo = msg;
    G.infoNew = true;
    renderBbInfo();
  }
}
// 実況のうち「インフォメーション」に該当するか判定 (打席結果の実況は対象外)。
// 回ヘッダー(══ N回 ══)は外野に出さない (回数は右下の bb-state に表示するため)。
function isInfoMessage(msg, cls) {
  const m = msg || '';
  if (/^══/.test(m)) return false;                    // 回ヘッダーは外野に出さない
  if (cls === 'event-inning') return true;            // 投手交代/代打/タイブレーク/試合終了 等
  return /^(🔄|🔁|🏃|🏟|🏁|⚠️)/.test(m);              // 盗塁(🏃💨/🏃❌)等 先頭マーカーで判定
}
// ダイヤモンド外野のインフォ表示を更新。text を渡すとそれを表示 (履歴閲覧の時系列表示用)。
function renderBbInfo(text) {
  const el = document.querySelector('#bbInfo');
  if (!el) return;
  let txt = (text != null) ? text : (G.lastInfo || '');
  // 通知が1件のときだけ括弧の直前で改行(見やすく)。複数件は各件1行のまま表示。
  if (txt && txt.indexOf('\n') < 0) txt = txt.replace(/\s*(?=[（(])/, '\n');
  el.textContent = txt;
  el.hidden = !txt;
}

// ============================================================
// レギュラーシーズン (30球団・2リーグ×3地区・各チーム162試合 ＋ ポストシーズン)
// ============================================================
const SEASON_KEY = 'mlb_season_v1';
// 2リーグ × 3地区 × 5球団 = 30球団 (実物MLB構成)
const SEASON_DIVISIONS = {
  ALE: { lg: 'AL', jp: 'アメリカンリーグ東地区', teams: ['NYY', 'BOS', 'TB', 'TOR', 'BAL'] },
  ALC: { lg: 'AL', jp: 'アメリカンリーグ中地区', teams: ['CWS', 'CLE', 'DET', 'KC', 'MIN'] },
  ALW: { lg: 'AL', jp: 'アメリカンリーグ西地区', teams: ['ATH', 'HOU', 'LAA', 'SEA', 'TEX'] },
  NLE: { lg: 'NL', jp: 'ナショナルリーグ東地区', teams: ['ATL', 'MIA', 'NYM', 'PHI', 'WSH'] },
  NLC: { lg: 'NL', jp: 'ナショナルリーグ中地区', teams: ['CHC', 'CIN', 'MIL', 'PIT', 'STL'] },
  NLW: { lg: 'NL', jp: 'ナショナルリーグ西地区', teams: ['ARI', 'COL', 'LAD', 'SD', 'SF'] },
};
const SEASON_DIV_ORDER = ['ALE', 'ALC', 'ALW', 'NLE', 'NLC', 'NLW'];
const SEASON_LEAGUE_DIVS = { AL: ['ALE', 'ALC', 'ALW'], NL: ['NLE', 'NLC', 'NLW'] };
const SEASON_LEAGUES = {
  AL: [].concat(...SEASON_LEAGUE_DIVS.AL.map(d => SEASON_DIVISIONS[d].teams)),
  NL: [].concat(...SEASON_LEAGUE_DIVS.NL.map(d => SEASON_DIVISIONS[d].teams)),
};
const SEASON_LEAGUE_JP = { AL: 'アメリカンリーグ', NL: 'ナショナルリーグ' };
const SEASON_TEAMS = [...SEASON_LEAGUES.AL, ...SEASON_LEAGUES.NL];   // 全30球団 (AL15 + NL15)
// チームの所属リーグ ('AL' | 'NL'、未所属は null)
function seasonLeagueOf(team) {
  const t = normalizeTeam(team);
  if (SEASON_LEAGUES.AL.includes(t)) return 'AL';
  if (SEASON_LEAGUES.NL.includes(t)) return 'NL';
  return null;
}
// チームの所属地区キー ('ALE'..'NLW'、未所属は null)
function seasonDivisionOf(team) {
  const t = normalizeTeam(team);
  for (const k of SEASON_DIV_ORDER) if (SEASON_DIVISIONS[k].teams.includes(t)) return k;
  return null;
}
const SEASON_TEAM_JP = (() => {
  const m = {};
  for (const d of MLB_DIVISIONS) for (const [c, n] of d.teams) m[c] = n;
  return m;
})();
let SEASON = null;          // 現在のシーズン状態 (localStorageと同期)
let SEASON_VIEW = 'menu';   // 'menu' | 'manual' | 'auto' | 'stats' | 'postseason' | 'postseason-manual'
let SEASON_STATS_TAB = 'standings';  // 'standings' | 'team' | 'bat' | 'pit'
let SEASON_PS_TAB = 'bracket';  // ポストシーズン画面のサブタブ: 'bracket' | 'stats' | 'ws'
const SEASON_PS_SORT = { bat: { key: 'ops', dir: -1 }, pit: { key: 'era', dir: 1 } };  // ポストシーズン成績の並び替え
let SEASON_PS_TEAM = '';  // ポストシーズン成績のチーム絞り込み ('' = 全体)
let SEASON_AUTORUN = null;  // 自動進行中の状態 { remaining, played, stop }
let SEASON_MANUAL_SEL = null;  // 手動モードの当該試合の選択状態 { cursor, away:{team,order,starterIdx,rest,picks,lineOrder}, home:{...} }
let SEASON_DRAG = null;        // 打順ドラッグ中の状態 { side, fromIdx }
let SEASON_OPEN_PICK = null;   // スタメン選択ドロップダウンの開閉状態 { side, pos }
let SEASON_OPEN_STARTER = null; // 先発選択ドロップダウンの開閉状態 { side }
// 成績表の並び替え状態 (列クリックで切替)。dir: -1=降順 / 1=昇順
const SEASON_SORT = {
  bat:  { key: 'avg',  dir: -1 },
  pit:  { key: 'era',  dir: 1 },
  tbat: { key: 'ops',  dir: -1 },
  tpit: { key: 'era',  dir: 1 },
};

function seasonTeamName(code) { return SEASON_TEAM_JP[normalizeTeam(code)] || code; }

// --- スケジュール生成: 30球団2リーグ×3地区。各チーム162試合 ---
//   同地区(相手4)       : 各ペア13試合 (4×13=52)
//   同リーグ他地区(相手10): 各チーム 4相手×7 + 6相手×6 = 64
//   インターリーグ(相手15): ライバル1組×4 + 他14×3 = 46
//   各カードの試合は季節全体へ均等分散して並べる (地区内/他地区/交流戦が偏らないように)。
function seasonGenerateSchedule() {
  const AL = SEASON_LEAGUES.AL, NL = SEASON_LEAGUES.NL;
  const queues = [];   // 各対戦カードの試合列 (後で均等に分散して結合)
  const addPair = (a, b, count, aHomeFirst) => {
    const q = [];
    for (let i = 0; i < count; i++) {
      const aHome = aHomeFirst ? (i % 2 === 0) : (i % 2 === 1);   // ホーム/アウェイを交互に
      q.push(aHome ? { home: a, away: b } : { home: b, away: a });
    }
    if (q.length) queues.push(q);
  };
  // 同地区: 各ペア13試合
  for (const dk of SEASON_DIV_ORDER) {
    const t = SEASON_DIVISIONS[dk].teams;
    for (let i = 0; i < t.length; i++)
      for (let j = i + 1; j < t.length; j++)
        addPair(t[i], t[j], 13, (i + j) % 2 === 0);
  }
  // 同リーグ他地区: 地区ペア(X,Y)ごとに Xのi番手は Yの {i, i+1}(巡回) を7試合・他を6試合。
  //   → 全球団が「4相手×7 + 6相手×6 = 64」を均等に満たす (対称性が保たれる)。
  for (const lg of ['AL', 'NL']) {
    const divs = SEASON_LEAGUE_DIVS[lg];
    for (let x = 0; x < divs.length; x++)
      for (let y = x + 1; y < divs.length; y++) {
        const TX = SEASON_DIVISIONS[divs[x]].teams, TY = SEASON_DIVISIONS[divs[y]].teams;
        for (let i = 0; i < TX.length; i++)
          for (let j = 0; j < TY.length; j++) {
            const seven = (j === i || j === (i + 1) % TX.length);
            addPair(TX[i], TY[j], seven ? 7 : 6, (i + j) % 2 === 0);
          }
      }
  }
  // インターリーグ: ライバル(AL[i]↔NL[i])のみ4試合、他は3試合 → 各チーム 1×4+14×3=46
  AL.forEach((a, ai) => NL.forEach((b, bi) => addPair(a, b, ai === bi ? 4 : 3, (ai + bi) % 2 === 0)));
  // 各カードの試合を [0,1) の相対位置へ均等配置し、位置順に並べて季節全体へ分散する
  const all = [];
  queues.forEach(q => { const n = q.length; q.forEach((g, i) => all.push({ g, pos: (i + 0.5) / n })); });
  all.sort((a, b) => a.pos - b.pos);
  return all.map(x => x.g);   // 計2430試合 (各チーム162)
}

function seasonNewState() {
  const s = {
    format: 3,   // 3 = 2リーグ×3地区30球団制 (旧10球団 format:2 とは非互換)
    teams: SEASON_TEAMS.slice(),
    schedule: seasonGenerateSchedule(),
    cursor: 0,
    standings: {}, h2h: {}, bat: {}, pit: {}, stamina: {}, staminaMax: {}, rotIdx: {},
    fieldInn: {},   // 守備イニング: { playerKey: { C, 1B, 2B, 3B, SS, LF, CF, RF, DH } } レギュラーシーズンで各守備位置に就いたイニング数
    createdAt: Date.now(),
  };
  SEASON_TEAMS.forEach(t => { s.standings[t] = { w: 0, l: 0, d: 0, rs: 0, ra: 0, hr: 0, sb: 0, e: 0 }; s.rotIdx[t] = 0; });
  SEASON_TEAMS.forEach(a => { s.h2h[a] = {}; SEASON_TEAMS.forEach(b => { if (a !== b) s.h2h[a][b] = { w: 0, l: 0, d: 0 }; }); });
  return s;
}
// 対戦成績(h2h)構造の保険初期化 (旧シーズンにも対応)
function seasonEnsureH2H() {
  if (!SEASON.h2h) SEASON.h2h = {};
  SEASON_TEAMS.forEach(a => {
    if (!SEASON.h2h[a]) SEASON.h2h[a] = {};
    SEASON_TEAMS.forEach(b => { if (a !== b && !SEASON.h2h[a][b]) SEASON.h2h[a][b] = { w: 0, l: 0, d: 0 }; });
  });
}
function loadSeason() {
  try { const r = localStorage.getItem(SEASON_KEY); SEASON = r ? JSON.parse(r) : null; }
  catch (e) { SEASON = null; }
  // 旧シーズン(1リーグ6球団 / 2リーグ10球団=format:2)は30球団制と非互換。破棄して新シーズンから開始させる。
  if (SEASON && SEASON.format !== 3) SEASON = null;
  return SEASON;
}
function saveSeason() {
  if (!SEASON) return;
  try { localStorage.setItem(SEASON_KEY, JSON.stringify(SEASON)); } catch (e) { console.error('saveSeason', e); }
}

// この選手がスタメンに入り得る試合の割合 P (自動: オーダー1=60%, 2=30%, 3=10% の使用率)。
//   出場可能オーダー(playerAllowedInOrder; 例外含む)の使用率を合計する。
function seasonInLineupP(p) {
  const usage = [0.60, 0.30, 0.10];
  let P = 0;
  for (let k = 0; k < 3; k++) { if (playerAllowedInOrder(p, k)) P += usage[k]; }
  return P > 0 ? P : 0.10;
}
// 野手の休養確率。シーズン出場数がカード年度試合数 g に概ね一致するよう、
//   「スタメンに入り得る試合(P)」の中で休む割合を逆算する:
//     目標出場割合 target = g / perTeam。 P の中で休む割合 = 1 - target/P。
//   → 出場数 ≈ P × (1 - rest) × perTeam = min(g, P×perTeam)。
//   例) 118試合・全オーダー可(P=1.0) → rest≒0.28 → 約118試合 (165フル出場しない)。
//      118試合・オーダー1,3のみ(P=0.7) → rest=0 → 約116試合 (オーダー制約で自然に減る)。
//   g 不明 or フル出場相当(>=perTeam) は休養なし。
//   総合力による調整: 59以下=休養なし / 60〜65=休養率を1/2 / 66以上=通常。
//   オーダー特例選手も休養なし: 守備DRS3ポジション以上のユーティリティで総合力63以下。
function seasonRestProb(p, perTeam) {
  const ovr = overallOf(p) || 0;
  if (ovr <= 59) return 0;                                  // 総合力59以下: 休養なし
  if (drsPositionCount(p) >= 3 && ovr <= 63) return 0;      // ユーティリティ(3ポジション可)で総合力63以下: 休養なし
  const g = getCardGamesOf(p);
  if (g == null || g >= perTeam) return 0;
  const P = seasonInLineupP(p);
  const target = g / perTeam;
  let rest = Math.max(0, Math.min(0.95, 1 - target / P));
  if (ovr <= 65) rest = rest / 2;             // 総合力60〜65: 休養率を従来の1/2に
  return rest;
}

// 投手控え(控え1〜4)の疲労スワップ。シーズン中、控え1〜4が全員スタミナ残率30%未満になったら
//   控え3,4 を余剰(登録外)の投手と一時的に入れ替える。控え1〜4が全員60%以上に回復したら元へ戻す。
//   ヒステリシス(30%で入替開始・60%で復帰)はチーム単位で SEASON.pitBenchSwap に保持する。
//   戻り値は実効的な控え配列 [控え1,控え2,控え3,控え4] (元の tb.pitchers.bench は変更しない)。
function seasonEffectiveBenchPitchers(teamCode, tb, orderIdx) {
  const orig = (tb.pitchers.bench || []).slice();
  if (!SEASON) return orig;
  const present = orig.slice(0, 4).filter(Boolean);
  if (present.length === 0) return orig;
  const ratios = present.map(seasonStaminaRatio);
  if (!SEASON.pitBenchSwap) SEASON.pitBenchSwap = {};
  const cur = !!SEASON.pitBenchSwap[teamCode];
  let swapped = cur;
  if (!cur && ratios.every(r => r < 0.30)) swapped = true;        // 控え1〜4 全員30%未満 → スワップ開始
  else if (cur && ratios.every(r => r >= 0.60)) swapped = false;  // 控え1〜4 全員60%以上 → 元に戻す
  SEASON.pitBenchSwap[teamCode] = swapped;
  if (!swapped) return orig;
  // 控え3,4 を余剰投手(登録外・年度可・同名重複除外)のフレッシュ(残率高)→総合力 順で差し替える
  const setYear = seasonTeamYear(teamCode);
  const registered = new Set();
  ['starter', 'mop', 'middle', 'setup', 'closer', 'bench'].forEach(r => (tb.pitchers[r] || []).forEach(p => { if (p) registered.add(playerKey(p)); }));
  const teamPit = applyTeamFilter(getPitchers(), teamCode);
  const surplus = teamPit
    .filter(p => p && !registered.has(playerKey(p)) && seasonYearAllowed(p, setYear, orderIdx) && !seasonHasNewerSameName(p, setYear, teamPit))
    .sort((a, b) => (seasonStaminaRatio(b) - seasonStaminaRatio(a)) || ((overallOf(b) || 0) - (overallOf(a) || 0)));
  const eff = orig.slice();
  if (surplus[0]) eff[2] = surplus[0];   // 控え3
  if (surplus[1]) eff[3] = surplus[1];   // 控え4
  return eff;
}

// 保存チーム + オーダー + 先発投手 から G.setup[side] を構築。OKなら null、不可ならエラーメッセージ。
//   customPicks ({pos: playerKey}) を渡すと、その守備位置の選手を差し替える (手動モードのスタメン編集)。
//   customLineOrder ([pos,...]) を渡すと、その打順並びで組む (手動モードのドラッグ&ドロップ打順)。
//   customPicks 指定時は、ユーザーが既に休養選手を外して選んでいるため自動の休養日入替は行わない。
function buildSeasonSideSetup(side, teamCode, orderIdx, starter, auto, customPicks, customLineOrder, customRest) {
  const SLOTS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];
  const cntBat = (t) => t ? SLOTS.filter(p => t.batters[p]).length : 0;
  let tb = getSavedTeamBuild(teamCode, orderIdx);
  // 指定オーダーが未編成(9人未満)なら オーダー1 にフォールバック
  if (cntBat(tb) < 9 && orderIdx !== 0) tb = getSavedTeamBuild(teamCode, 0);
  if (!tb) return '「' + seasonTeamName(teamCode) + '」のチーム編成が未保存です';
  // 同名(同種別)の選手は、セット年度以下で最も新しい年度版に統一する。
  //   保存編成に古い年度版(例:2023吉田)が残っていても、出場時は2025版に置き換える。
  const teamPoolAll = applyTeamFilter(getBatters(), teamCode);
  const dedupYear = seasonTeamYear(teamCode);
  const newestSameName = (p) => {
    if (!p || dedupYear == null || !Number.isFinite(p.year)) return p;
    let best = p;
    teamPoolAll.forEach(q => {
      if (q.fullNameTop === p.fullNameTop && playerType(q) === playerType(p) && Number.isFinite(q.year)
        && q.year <= dedupYear && q.year > best.year) best = q;
    });
    return best;
  };
  // customPicks 解決用 (チームの野手プールから playerKey で引く)
  const cpPool = customPicks ? teamPoolAll : null;
  const resolvePick = (pos, fallback) => {
    let p = fallback;
    if (customPicks && customPicks[pos]) { const cp = cpPool.find(pp => playerKey(pp) === customPicks[pos]); if (cp) p = cp; }
    return newestSameName(p);   // 同名の最新年度版に統一
  };
  const ordered = [];
  if (Array.isArray(customLineOrder) && customLineOrder.length) {
    // 手動で指定した打順並び (ドラッグ&ドロップ結果)
    for (const pos of customLineOrder) {
      if (ordered.length >= 9) break;
      const p = resolvePick(pos, tb.batters[pos]);
      if (p) ordered.push({ pos, p });
    }
  } else {
    for (let spot = 1; spot <= 9; spot++) {
      const pos = SLOTS.find(p => tb.batterOrder[p] === spot && tb.batters[p]);
      if (pos) ordered.push({ pos, p: resolvePick(pos, tb.batters[pos]) });
    }
    for (const pos of SLOTS) {                  // 打順未設定の穴埋め
      if (ordered.length >= 9) break;
      if (tb.batters[pos] && !ordered.some(o => o.pos === pos)) ordered.push({ pos, p: resolvePick(pos, tb.batters[pos]) });
    }
  }
  if (ordered.length < 9) return '「' + seasonTeamName(teamCode) + '」の野手が9人揃っていません';
  if (!starter) return '「' + seasonTeamName(teamCode) + '」の先発投手がいません';
  G.setup[side].batters = ordered.map(o => o.p);
  G.setup[side].batterPos = ordered.map(o => o.pos);
  // リリーフは「役割」を保持する (中継/SU/抑え/MU)。試合中の継投ロジック(pickReliever)が役割別に起用するため必須。
  //   ★疲労管理: 中継・MU は スタミナ残率が25%未満なら今試合は休養(ロースターから外す)。
  //     外した数だけ「控え2,3,4」(控え1=先発予備は除く) から残率の高い順に一時的に補充する。
  const FATIGUE = 0.25;
  const middleArr = (tb.pitchers.middle || []).filter(Boolean);
  const mopArr    = (tb.pitchers.mop    || []).filter(Boolean);
  const relief = [];
  middleArr.forEach(p => { if (seasonStaminaRatio(p) >= FATIGUE) relief.push({ p, role: 'middle' }); });
  (tb.pitchers.setup  || []).forEach(p => { if (p) relief.push({ p, role: 'setup' }); });
  (tb.pitchers.closer || []).forEach(p => { if (p) relief.push({ p, role: 'closer' }); });
  mopArr.forEach(p => { if (seasonStaminaRatio(p) >= FATIGUE) relief.push({ p, role: 'mop' }); });
  // 休養させた中継/MUの数だけ、控え2,3,4 から残率の高い順に補充 (それぞれ元の役割で代替)
  const needMiddle = middleArr.filter(p => seasonStaminaRatio(p) < FATIGUE).length;
  const needMop    = mopArr.filter(p => seasonStaminaRatio(p) < FATIGUE).length;
  const usedKeys = new Set(relief.map(r => playerKey(r.p))); usedKeys.add(playerKey(starter));
  // 控え疲労スワップ適用後の実効控え (控え1〜4全員30%未満なら控え3,4を余剰投手へ一時入替、60%回復で復帰)
  const effBench = seasonEffectiveBenchPitchers(teamCode, tb, orderIdx);
  const benchPool = effBench.slice(1).filter(Boolean)   // slice(1) = 控え2,3,4 (控え1=先発予備は除外)
    .filter(p => seasonStaminaRatio(p) >= FATIGUE && !usedKeys.has(playerKey(p)))
    .sort((a, b) => seasonStaminaRatio(b) - seasonStaminaRatio(a));
  let bi = 0;
  for (let k = 0; k < needMiddle && bi < benchPool.length; k++, bi++) relief.push({ p: benchPool[bi], role: 'middle' });
  for (let k = 0; k < needMop    && bi < benchPool.length; k++, bi++) relief.push({ p: benchPool[bi], role: 'mop' });
  const pitchers = [starter], roles = ['starter'];
  const seen = new Set([playerKey(starter)]);
  relief.forEach(({ p, role }) => { const k = playerKey(p); if (!seen.has(k)) { seen.add(k); pitchers.push(p); roles.push(role); } });
  G.setup[side].pitchers = pitchers;
  G.setup[side].pitcherRoles = roles;
  // 控え: 同名の新しい年度版が編成内にいる古い版(例:2023吉田)は除外する
  const bench = (tb.pinchHitters || []).filter(b => b && !seasonHasNewerSameName(b, dedupYear, teamPoolAll));
  const benchLabels = ['PH', 'PH', 'PH', '代走', '守備', '守備'];
  // ===== 休養日 =====
  //   この試合「休養」する選手の集合 restToday を作る。
  //     手動(customRest指定): すでにロール済みの休養セットを使用。
  //     自動: チームの全野手を seasonRestProb() でロール。
  //   休養選手は (1)スタメンから外し控えと交代(自動のみ。手動は編集済みなので入替なし)、
  //            (2)控え(代打/代走/守備固め)からも除外する。
  const restToday = (customRest instanceof Set) ? customRest
    : (Array.isArray(customRest) ? new Set(customRest) : new Set());
  if (!(customRest) && SEASON && Array.isArray(SEASON.schedule) && SEASON.schedule.length) {
    const perTeam = (SEASON.schedule.length * 2) / SEASON_TEAMS.length;   // 各チームの試合数 (=162)
    applyTeamFilter(getBatters(), teamCode).forEach(p => {
      if (Math.random() < seasonRestProb(p, perTeam)) restToday.add(playerKey(p));
    });
  }
  const setYear = seasonTeamYear(teamCode);
  // (1) 自動モード: 休養スタメンを控えから抜擢して交代。
  //     抜擢候補は「守れる・オーダー(パターン1,2,3)登録可・年度可・非休養・未使用」の控えのみ。
  if (!customPicks && restToday.size) {
    const usedSubs = new Set();   // 今試合で投入したフィラー(控え/余剰)のキー
    let swapped = false;
    // フィラー候補プール: 登録控え(優先) → 未登録の余剰選手(総合力順)。
    const lvPoolAll = applyTeamFilter(getBatters(), teamCode);
    const subPool = bench.concat(
      lvPoolAll.filter(p => p && !bench.some(b => playerKey(b) === playerKey(p)))
               .sort((a, b) => (overallOf(b) || 0) - (overallOf(a) || 0))
    );
    const bat = G.setup[side].batters, bpos = G.setup[side].batterPos;
    // フィラーが pos を守れて起用可能か (未使用・スタメン未登録・非休養・年度可・同名重複なし)
    const fillerOk = (b, pos) => b && !usedSubs.has(playerKey(b))
      && !bat.some(x => playerKey(x) === playerKey(b))
      && canPlay(b, pos) && playerAllowedInOrder(b, orderIdx) && seasonYearAllowed(b, setYear, orderIdx)
      && !restToday.has(playerKey(b)) && !seasonHasNewerSameName(b, setYear, lvPoolAll);
    for (let i = 0; i < 9; i++) {
      if (!restToday.has(playerKey(bat[i]))) continue;          // 休養でない → そのまま出場
      const pos = bpos[i];
      // 1) 直接補完: 控え/余剰で休養者の守備位置を埋める
      const direct = subPool.find(b => fillerOk(b, pos));
      if (direct) { usedSubs.add(playerKey(direct)); bat[i] = direct; swapped = true; continue; }
      // 2) 連鎖補完: その位置を守れる「他の非休養スタメン」を pos へ移し、空いた元位置を控え/余剰で埋める。
      //    (例: Moベッツ(SS)が休養 → SSを守れる他スタメンをSSへ移動し、その元位置を控えで補完する)
      let done = false;
      for (let j = 0; j < 9 && !done; j++) {
        if (j === i) continue;
        const mover = bat[j];
        if (!mover || restToday.has(playerKey(mover))) continue;   // 休養者は動かさない(外す対象)
        if (!canPlay(mover, pos)) continue;                        // mover が休養者の位置を守れること
        const backfill = subPool.find(b => fillerOk(b, bpos[j]));  // mover の元位置を控え/余剰で補完
        if (!backfill) continue;
        bat[i] = mover; usedSubs.add(playerKey(backfill)); bat[j] = backfill;   // mover→空き位置, backfill→moverの元位置
        swapped = true; done = true;
      }
      // 3) どうしても埋められない → 休養者をそのまま残す(安全フォールバック・稀)
    }
    // 入替があれば打順を規定ルールでAIにて組み直す (9人揃っている場合のみ)
    if (swapped && !(Array.isArray(customLineOrder) && customLineOrder.length)) {
      const cur = bat.map((p, k) => ({ pos: bpos[k], p }));
      if (cur.length === 9 && cur.every(e => e.p)) {
        const reordered = seasonAiBattingOrder(cur);
        G.setup[side].batters = reordered.map(e => e.p);
        G.setup[side].batterPos = reordered.map(e => e.pos);
      }
    }
  }
  // (2) 控え再編成: 先発入り(抜擢含む)・休養 を除いた控えをベースに、抜擢で減った分を
  //     「余っている適格選手」(年度可・非休養・未使用、総合力順)で既存控え枠数まで補充する。
  let benchList = bench.filter(b => !G.setup[side].batters.includes(b) && !restToday.has(playerKey(b)));
  const TARGET_BENCH = (tb.pinchHitters || []).filter(Boolean).length || 6;
  if (benchList.length < TARGET_BENCH) {
    const usedK = new Set(G.setup[side].batters.map(playerKey).concat(benchList.map(playerKey)));
    const lvPool = applyTeamFilter(getBatters(), teamCode);
    const leftovers = lvPool
      .filter(p => !usedK.has(playerKey(p)) && !restToday.has(playerKey(p)) && seasonYearAllowed(p, setYear, orderIdx)
        && !seasonHasNewerSameName(p, setYear, lvPool))
      .sort((a, b) => (overallOf(b) || 0) - (overallOf(a) || 0));
    for (const p of leftovers) {
      if (benchList.length >= TARGET_BENCH) break;
      benchList.push(p); usedK.add(playerKey(p));
    }
  }
  G.setup[side].bench = benchList;
  G.setup[side].benchRole = new Map();
  // 役割ラベル: 元の控え選手は元の枠の役割を維持、補充選手は控えリスト順の役割を割当
  benchList.forEach((b, i) => {
    const origIdx = (tb.pinchHitters || []).findIndex(x => x && playerKey(x) === playerKey(b));
    G.setup[side].benchRole.set(b, benchLabels[origIdx >= 0 ? origIdx : i] || 'PH');
  });
  return null;
}

// 投手のスタミナ残率 (現在残量 / 上限)
function seasonStaminaRatio(p) {
  if (!p) return 0;
  const k = playerKey(p), max = (p.stats && p.stats['スタミナ']) || 70;
  const cur = (SEASON && SEASON.stamina[k] != null) ? SEASON.stamina[k] : max;
  return max > 0 ? cur / max : 0;
}
// ローテーション候補一覧 (手動の先発選択UI用): 先発1..5 / 控え1..4 / MU1..2
function seasonRotationList(teamCode) {
  const tb = getSavedTeamBuild(teamCode, 0);
  if (!tb) return [];
  const out = [];
  (tb.pitchers.starter || []).forEach((p, i) => { if (p) out.push({ p, label: '先発' + (i + 1) }); });
  (tb.pitchers.bench || []).forEach((p, i) => { if (p) out.push({ p, label: '控え' + (i + 1) }); });
  (tb.pitchers.mop || []).forEach((p, i) => { if (p) out.push({ p, label: 'MU' + (i + 1) }); });
  return out;
}
// チームのローテーション位置 / 先発人数 / 進行
function seasonRotIdx(team) { return (SEASON.rotIdx && SEASON.rotIdx[team] != null) ? SEASON.rotIdx[team] : 0; }
function seasonStarterCount(team) { const tb = getSavedTeamBuild(team, 0); const n = tb ? (tb.pitchers.starter || []).filter(Boolean).length : 0; return n || 5; }
function seasonAdvanceRot(team) {
  if (!SEASON.rotIdx) SEASON.rotIdx = {};
  SEASON.rotIdx[team] = (seasonRotIdx(team) + 1) % seasonStarterCount(team);
}
function seasonEnsureRot() { if (!SEASON.rotIdx) SEASON.rotIdx = {}; SEASON_TEAMS.forEach(t => { if (SEASON.rotIdx[t] == null) SEASON.rotIdx[t] = 0; }); }

// 自動モードの先発選定 (ローテーション順を基本に維持。rotN = 今回のローテーション該当スロット)
//   先発1→2→3→4→5→… の順を守り、該当先発が回復していなければ
//   控え1/MUのスポット先発・谷間先発でつなぎ、最後の手段でのみローテを崩す。
function seasonAutoStarter(teamCode, rotN) {
  const tb = getSavedTeamBuild(teamCode, 0);
  if (!tb) return null;
  const st = (tb.pitchers.starter || []).filter(Boolean);   // 先発1..5
  const mu = (tb.pitchers.mop || []).filter(Boolean);       // MU1, MU2
  const bn = (tb.pitchers.bench || []).filter(Boolean);     // 控え1..4
  const B1 = bn[0] || null, M1 = mu[0] || null, M2 = mu[1] || null;
  const R = seasonStaminaRatio;
  if (st.length === 0) return B1 || M1 || bn[1] || null;
  const n = (((rotN || 0) % st.length) + st.length) % st.length;
  const SN = st[n];   // 今回のローテーション該当先発

  if (R(SN) >= 0.90) return SN;                              // 1) 該当先発が残率90%以上 → ローテ通り
  if (B1 && R(B1) > 0.90) return B1;                         // 2) 控え1 残率>90% → スポット先発
  if (M1 && R(M1) >= 1.0) return M1;                         //    or MU1/MU2 残量100% → スポット先発
  if (M2 && R(M2) >= 1.0) return M2;
  if (R(SN) >= 0.85) return SN;                              // 3) 該当先発が残率85%以上 → ローテ通り
  if (M1 && R(M1) > 0.90) return M1;                         // 4) MU1,MU2 残率>90% → 谷間先発
  if (M2 && R(M2) > 0.90) return M2;
  for (const p of [B1, M1, M2]) if (p && R(p) > 0.80) return p;   // 5) 控え1/MU1/MU2 残率>80% → 谷間
  for (const p of st) if (R(p) >= 1.0) return p;             // 6) 先発で残率100%が居れば → 初めてローテを崩す
  for (let k = 0; k < st.length; k++) { const p = st[(n + k) % st.length]; if (R(p) > 0.85) return p; }  // 7) 先発 残率>85% をローテ順(該当先発起点)で
  for (let i = 1; i < bn.length; i++) if (bn[i]) return bn[i];    // 8) 控え2,3,4 の順に緊急先発
  return SN || st[0] || B1 || M1;                            // 最終保険
}
// オーダー制限抽選: 60%=全選択可(自動は1), 30%=2のみ, 10%=3のみ
function seasonRollOrders() {
  const r = Math.random();
  if (r < 0.60) return [0, 1, 2];
  if (r < 0.90) return [1];
  return [2];
}
function seasonEnsureRoll(g) {
  let rolled = false;
  if (!g.allowAway) { g.allowAway = seasonRollOrders(); rolled = true; }
  if (!g.allowHome) { g.allowHome = seasonRollOrders(); rolled = true; }
  if (rolled && SEASON) { try { saveSeason(); } catch (e) {} }   // 抽選結果を永続化(リロードでも安定)
}
// 試合開始時、持ち越しスタミナで現役投手の残量を上書き
function applySeasonStamina() {
  if (!SEASON) return;
  for (const side of ['away', 'home']) {
    (G.setup[side].pitchers || []).forEach((p, idx) => {
      const k = playerKey(p), max = G.setup[side].pitcherMax[idx];
      const cur = (SEASON.stamina[k] != null) ? Math.min(max, SEASON.stamina[k]) : max;
      G.setup[side].pitcherStamina[idx] = cur;
    });
  }
}

// 現在の試合を開始 (auto=自動で最後まで進める / sel=手動の選択{awayOrder,homeOrder,awayStarter,homeStarter})
// 未保存チームのチーム編成を自動生成して保存する (保存済みは絶対に触れない)。
//   30球団の自動進行で、ユーザーが手動編成していないチームをその場で自動編成する。
//   TB_STATE と描画関数を一時退避し、autoFillTeamBuild をヘッドレス実行 → saveTeamBuild。
function seasonEnsureBuild(teamCode) {
  const t = normalizeTeam(teamCode);
  if (!t || t === 'original') return;
  if (localStorage.getItem(TB_STORAGE_PREFIX + t)) return;   // 既に保存済み → 触れない
  const _tb = TB_STATE, _render = renderTeamBuild, _last = localStorage.getItem(TB_LAST_TEAM_KEY);
  try {
    renderTeamBuild = function () {};         // ヘッドレス化 (DOM非依存)
    TB_STATE = blankTeamState(t);
    TB_STATE.year = null;                     // 年度フィルタ無し=チームの全カードを候補に
    TB_STATE.currentOrder = 0;
    autoFillTeamBuild({ batters: true, pitchers: true });
    saveTeamBuild();
  } catch (e) { console.error('seasonEnsureBuild', t, e); }
  finally {
    renderTeamBuild = _render; TB_STATE = _tb;
    if (_last != null) localStorage.setItem(TB_LAST_TEAM_KEY, _last);
  }
}
function seasonStartCurrentGame(auto, sel) {
  const g = SEASON.schedule[SEASON.cursor];
  if (!g) return;
  seasonEnsureBuild(g.away); seasonEnsureBuild(g.home);   // 未保存チームは自動編成で補う
  seasonEnsureRoll(g);
  seasonEnsureRot();
  const aN = seasonRotIdx(g.away), hN = seasonRotIdx(g.home);   // 今回のローテーション該当スロット
  let awayOrder, homeOrder, awayStarter, homeStarter;
  if (auto) {
    awayOrder = g.allowAway[0]; homeOrder = g.allowHome[0];
    awayStarter = seasonAutoStarter(g.away, aN); homeStarter = seasonAutoStarter(g.home, hN);
  } else {
    awayOrder = (sel && g.allowAway.includes(sel.awayOrder)) ? sel.awayOrder : g.allowAway[0];
    homeOrder = (sel && g.allowHome.includes(sel.homeOrder)) ? sel.homeOrder : g.allowHome[0];
    awayStarter = (sel && sel.awayStarter) || seasonAutoStarter(g.away, aN);
    homeStarter = (sel && sel.homeStarter) || seasonAutoStarter(g.home, hN);
  }
  seasonAdvanceRot(g.away); seasonAdvanceRot(g.home);   // 次の試合は次の先発の番へ
  const e1 = buildSeasonSideSetup('away', g.away, awayOrder, awayStarter, !!auto, sel && sel.awayPicks, sel && sel.awayLineOrder, sel && sel.awayRest);
  const e2 = e1 ? null : buildSeasonSideSetup('home', g.home, homeOrder, homeStarter, !!auto, sel && sel.homePicks, sel && sel.homeLineOrder, sel && sel.homeRest);
  if (e1 || e2) { alert('試合を開始できません:\n' + (e1 || e2) + '\n(そのチームのカードが不足しています。9人＋投手分のカードを追加してください)'); return false; }
  G.innings = 9;
  G.seasonMode = true;
  G.seasonCtx = { away: g.away, home: g.home, awayOrder, homeOrder, auto: !!auto };
  beginGame();
  if (auto) {
    let guard = 0;
    while (!G.ended && guard++ < 4000) { const p = autoPick(); if (!p) break; pitchOne(p, true); }
    if (!G.ended) { G.ended = true; G.awaitingResult = true; }   // 安全装置: 万一終わらない試合は確実に終了扱いにして記録 (シーズンを止めない)
  }
  return true;
}

// 直前に終了した試合の結果を集計・保存し、cursorを進める
function seasonRecordCurrentGame() {
  const g = SEASON.schedule[SEASON.cursor];
  if (!g) return;
  const dec = computePitcherDecisions();
  const stA = SEASON.standings[g.away], stH = SEASON.standings[g.home];
  stA.rs += dec.sA; stA.ra += dec.sH; stH.rs += dec.sH; stH.ra += dec.sA;
  seasonEnsureH2H();
  const hA = SEASON.h2h[g.away], hH = SEASON.h2h[g.home];
  if (dec.winSide === 'away') { stA.w++; stH.l++; if (hA && hA[g.home]) hA[g.home].w++; if (hH && hH[g.away]) hH[g.away].l++; }
  else if (dec.winSide === 'home') { stH.w++; stA.l++; if (hH && hH[g.away]) hH[g.away].w++; if (hA && hA[g.home]) hA[g.home].l++; }
  else { stA.d++; stH.d++; if (hA && hA[g.home]) hA[g.home].d++; if (hH && hH[g.away]) hH[g.away].d++; }
  seasonAccumBatting(g.away, 'away', stA);
  seasonAccumBatting(g.home, 'home', stH);
  seasonAccumPitching(g.away, 'away', dec.pitcherRoles);
  seasonAccumPitching(g.home, 'home', dec.pitcherRoles);
  seasonUpdateStamina(g.away, 'away');
  seasonUpdateStamina(g.home, 'home');
  seasonFlushFieldInn();   // この試合の守備イニングをシーズン累計へ合算
  SEASON.cursor++;
  // 自動進行中は毎試合の保存をスキップ (シーズン全体のJSON化が重いため、自動側で定期保存する)
  if (!SEASON_AUTORUN) saveSeason();
}
function seasonStatBat(key, p, team, store) {
  store = store || SEASON.bat;
  if (!store[key]) store[key] = { name: p.fullNameTop, team: normalizeTeam(team), year: p.year || '',
    G: 0, PA: 0, AB: 0, H: 0, dbl: 0, tpl: 0, HR: 0, RBI: 0, R: 0, SO: 0, BB: 0, HBP: 0, SAC: 0, SB: 0, CS: 0, E: 0 };
  return store[key];
}
// standObj が null のときは順位表加算をスキップ。store 省略時は SEASON.bat (=レギュラー)。
function seasonAccumBatting(team, side, standObj, store) {
  const all = (G.batterStats[side] || []).concat(G.subLog[side] || []);
  const counted = new Set();
  all.forEach(bs => {
    if (!bs || !bs.player) return;
    const p = bs.player, key = playerKey(p);
    const appeared = ((bs.AB || 0) + (bs.BB || 0) + (bs.HBP || 0) + (bs.SAC || 0)) > 0 || bs.fielded || (bs.SB || 0) > 0 || (bs.R || 0) > 0;
    if (!appeared) return;
    const s = seasonStatBat(key, p, team, store);
    if (!counted.has(key)) { s.G += 1; counted.add(key); }
    s.AB += bs.AB || 0; s.H += bs.H || 0; s.dbl += bs.doubles || 0; s.tpl += bs.triples || 0;
    s.HR += bs.HR || 0; s.RBI += bs.RBI || 0; s.R += bs.R || 0; s.SO += bs.K || 0;
    s.BB += bs.BB || 0; s.HBP += bs.HBP || 0; s.SAC += bs.SAC || 0; s.SB += bs.SB || 0; s.CS += bs.CS || 0; s.E += bs.E || 0;
    s.PA += (bs.AB || 0) + (bs.BB || 0) + (bs.HBP || 0) + (bs.SAC || 0);
    if (standObj) { standObj.hr += bs.HR || 0; standObj.sb += bs.SB || 0; standObj.e += bs.E || 0; }
  });
}
function seasonStatPit(key, p, team, store) {
  store = store || SEASON.pit;
  if (!store[key]) store[key] = { name: p.fullNameTop, team: normalizeTeam(team), year: p.year || '',
    G: 0, GS: 0, W: 0, L: 0, S: 0, HLD: 0, outs: 0, H: 0, HR: 0, K: 0, BB: 0, HBP: 0, ER: 0, R: 0, balk: 0 };
  return store[key];
}
function seasonAccumPitching(team, side, roles, store) {
  (G.pitcherLog[side] || []).forEach((lg, idx) => {
    if (!lg || !lg.pitcher) return;
    if (((lg.battersFaced || 0) <= 0) && ((lg.outs || 0) <= 0)) return;
    const s = seasonStatPit(playerKey(lg.pitcher), lg.pitcher, team, store);
    s.G += 1;
    if (idx === 0) s.GS += 1;
    s.outs += lg.outs || 0; s.H += lg.hits || 0; s.HR += lg.HR || 0; s.K += lg.K || 0;
    s.BB += lg.BB || 0; s.HBP += lg.HBP || 0; s.ER += lg.earnedRuns || 0; s.R += lg.runsAllowed || 0; s.balk += lg.balks || 0;
    const role = roles.get(lg);
    if (role === 'W') s.W += 1; else if (role === 'L') s.L += 1; else if (role === 'S') s.S += 1; else if (role === 'H') s.HLD += 1;
  });
}
// 直前の試合(G.fieldInn)の守備イニングをシーズン累計(SEASON.fieldInn)へ合算
function seasonFlushFieldInn() {
  if (!G.fieldInn) return;
  if (!SEASON.fieldInn) SEASON.fieldInn = {};
  for (const k in G.fieldInn) {
    if (!SEASON.fieldInn[k]) SEASON.fieldInn[k] = {};
    const src = G.fieldInn[k], dst = SEASON.fieldInn[k];
    for (const pos in src) dst[pos] = (dst[pos] || 0) + src[pos];
  }
}
// チーム編成(全オーダー)から playerKey → 守備位置コード のマップを作る (遡及推定の補完に使用)
function seasonBuildPosMap() {
  const map = {};
  (SEASON && SEASON.teams ? SEASON.teams : SEASON_TEAMS).forEach(team => {
    for (let oi = 0; oi < 3; oi++) {
      const tb = getSavedTeamBuild(team, oi);
      if (!tb || !tb.batters) continue;
      for (const pos of ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH']) {
        const bp = tb.batters[pos];
        if (bp) { const k = playerKey(bp); if (!(k in map)) map[k] = pos; }
      }
    }
  });
  return map;
}
function seasonUpdateStamina(team, side) {
  (G.setup[side].pitchers || []).forEach((p, idx) => { SEASON.stamina[playerKey(p)] = G.setup[side].pitcherStamina[idx]; });
  const tb = getSavedTeamBuild(team, 0);
  if (!tb) return;
  const allP = [];
  for (const r of ['starter', 'mop', 'middle', 'setup', 'closer', 'bench']) (tb.pitchers[r] || []).forEach(p => { if (p) allP.push(p); });
  allP.forEach(p => {
    const k = playerKey(p), max = (p.stats && p.stats['スタミナ']) || 70;
    const cur = (SEASON.stamina[k] != null) ? SEASON.stamina[k] : max;
    SEASON.stamina[k] = Math.min(max, cur + getRecoveryOf(p));
    SEASON.staminaMax[k] = max;
  });
}

// ============== シーズン画面 描画 ==============
function openSeason() {
  loadSeason();
  if (!SEASON) { SEASON = seasonNewState(); saveSeason(); }
  // スケジュール欠損時の保険
  if (!Array.isArray(SEASON.schedule) || SEASON.schedule.length === 0) { SEASON.schedule = seasonGenerateSchedule(); saveSeason(); }
  // 未開始(cursor=0)なら最新スケジュール(各チーム162試合)へ自動更新。進行済みは保持。
  if ((SEASON.cursor || 0) === 0) {
    const fresh = seasonGenerateSchedule();
    if (SEASON.schedule.length !== fresh.length) { SEASON.schedule = fresh; saveSeason(); }
  }
  seasonEnsureH2H();   // 対戦成績の構造を保証
  seasonEnsureRot();   // 先発ローテーション位置の構造を保証
  if (!SEASON.fieldInn) SEASON.fieldInn = {};   // 守備イニング集計の構造を保証 (旧シーズン対応)
  seasonDedupeStats(); // 同名(同チーム)の古い年度版の成績を整理 (例: 2023吉田を削除し2025のみ残す)
  SEASON_VIEW = 'menu';
  showScreen('season');
  renderSeason();
}
// 既存シーズンの成績から、同名(同チーム)の古い年度版エントリを削除する。
//   例: 「吉田正尚 2023」と「吉田正尚 2025」が両方ある場合、新しい2025のみ残す。
function seasonDedupeStats() {
  if (!SEASON) return;
  let changed = false;
  ['bat', 'pit'].forEach(tbl => {
    const entries = SEASON[tbl]; if (!entries) return;
    const groups = {};
    Object.keys(entries).forEach(k => {
      const e = entries[k]; const nk = (e.name || '') + '|' + (e.team || '');
      (groups[nk] = groups[nk] || []).push({ key: k, year: Number(e.year) || 0 });
    });
    Object.keys(groups).forEach(nk => {
      const arr = groups[nk]; if (arr.length <= 1) return;
      const maxY = arr.reduce((m, x) => Math.max(m, x.year), -Infinity);
      arr.forEach(x => { if (x.year < maxY) { delete entries[x.key]; changed = true; } });  // 古い同名版を削除
    });
  });
  if (changed) saveSeason();
}
function seasonDone() { return SEASON && SEASON.cursor >= SEASON.schedule.length; }
function ipText(outs) { return Math.floor(outs / 3) + '.' + (outs % 3); }
function eraOf(er, outs) { return outs > 0 ? (er * 27 / outs) : 0; }
function avgOf(h, ab) { return ab > 0 ? h / ab : 0; }
function fmt3(x) { const s = (Math.round(x * 1000) / 1000).toFixed(3); return s.replace(/^0\./, '.'); }
function fmt2(x) { return (Math.round(x * 100) / 100).toFixed(2); }
// チームの消化試合数
function seasonTeamGames(team) { const st = SEASON.standings[normalizeTeam(team)]; return st ? ((st.w || 0) + (st.l || 0) + (st.d || 0)) : 0; }
// 規定打席 = チーム試合数 × 3.1 / 規定投球回(アウト) = チーム試合数 × 1.0回(=3アウト)
function seasonRegPA(team) { return Math.ceil(3.1 * seasonTeamGames(team)); }
function seasonRegOuts(team) { return seasonTeamGames(team) * 3; }

function renderSeason() {
  const body = document.querySelector('#season-body');
  if (!body) return;
  if (SEASON_VIEW === 'menu') { body.innerHTML = seasonMenuHtml(); return; }
  const nav = `<div class="season-nav"><button class="btn btn-sub" data-sv="menu">← メニュー</button></div>`;
  if (SEASON_VIEW === 'manual') body.innerHTML = nav + seasonManualHtml();
  else if (SEASON_VIEW === 'auto') body.innerHTML = nav + seasonAutoHtml();
  else if (SEASON_VIEW === 'stats') body.innerHTML = nav + seasonStatsHtml();
  else if (SEASON_VIEW === 'postseason') body.innerHTML = nav + seasonPostseasonHtml();
  else if (SEASON_VIEW === 'postseason-manual') body.innerHTML = nav + seasonPostseasonManualHtml();
}
function seasonMenuHtml() {
  const total = SEASON.schedule.length;
  return `<div class="season-menu">
    <p class="season-progress">進行状況: <b>${SEASON.cursor} / ${total}</b> 試合 ${seasonDone() ? '（シーズン終了）' : ''}</p>
    <div class="season-modebtns">
      <button class="mode-btn" data-sv="manual"><span class="mode-icon">🎮</span><span class="mode-name">手動モード</span><span class="mode-desc">オーダー・投手を選んで1試合ずつ進める</span></button>
      <button class="mode-btn" data-sv="auto"><span class="mode-icon">⏩</span><span class="mode-name">自動モード</span><span class="mode-desc">試合数を指定して高速で自動進行</span></button>
      <button class="mode-btn" data-sv="stats"><span class="mode-icon">📊</span><span class="mode-name">成績モード</span><span class="mode-desc">勝敗表・チーム別成績・ベスト20</span></button>
      ${seasonDone() ? '<button class="mode-btn mode-btn-ps" data-sv="postseason"><span class="mode-icon">🏆</span><span class="mode-name">ポストシーズン</span><span class="mode-desc">地区優勝＋ワイルドカードで世界一を争う</span></button>' : ''}
    </div>
  </div>`;
}
function seasonGameHeaderHtml() {
  const g = SEASON.schedule[SEASON.cursor];
  return `<div class="season-gamehdr">第 <b>${SEASON.cursor + 1}</b> 試合　${seasonTeamName(g.away)} (AWAY) <span class="vs">vs</span> ${seasonTeamName(g.home)} (HOME)</div>`;
}
// シーズン累積成績の表示テキスト (セットアップのプルダウン用)
// 表示用のアクティブな成績ストア。ポストシーズン文脈なら postseason 成績、それ以外はレギュラー。
function seasonActiveStores() {
  const inPS = (SEASON_VIEW === 'postseason-manual' || (G.seasonCtx && G.seasonCtx.postseason)) && SEASON && SEASON.postseason;
  if (inPS) return { bat: SEASON.postseason.bat || {}, pit: SEASON.postseason.pit || {} };
  return { bat: (SEASON && SEASON.bat) || {}, pit: (SEASON && SEASON.pit) || {} };
}
function seasonBatStatLine(p) {
  const s = seasonActiveStores().bat[playerKey(p)];
  if (!s || !s.AB) return '.--- 0本 0打点 0盗';
  return `${fmt3(avgOf(s.H, s.AB))} ${s.HR || 0}本 ${s.RBI || 0}打点 ${s.SB || 0}盗`;
}
function seasonPitStatLine(p) {
  const k = playerKey(p);
  const s = seasonActiveStores().pit[k];
  const max = (p.stats && p.stats['スタミナ']) || 70;
  const cur = (SEASON && SEASON.stamina[k] != null) ? SEASON.stamina[k] : max;
  const sta = `スタ ${Math.round(cur)}/${max}(${max > 0 ? Math.round(cur / max * 100) : 0}%)`;
  if (!s || !s.outs) return `--.-- 0勝0敗0H0S ${sta}`;
  return `${fmt2(eraOf(s.ER, s.outs))} ${s.W || 0}勝${s.L || 0}敗${s.HLD || 0}H${s.S || 0}S ${sta}`;
}
function seasonPlayerLink(p, label) {
  if (!p) return label || '';
  return `<span class="player-link" data-player-name="${p.fullNameTop}" data-player-year="${p.year || ''}" data-player-type="${playerType(p)}" data-player-team="${p.team || ''}" title="クリックで選手カードを表示">${label != null ? label : p.fullNameTop}</span>`;
}
// 自動モードでAIが選ぶ先発の、seasonRotationList 内インデックス (手動プルダウンの初期選択用)
function seasonAutoStarterIdx(teamCode) {
  seasonEnsureRot();
  const aiP = seasonAutoStarter(teamCode, seasonRotIdx(teamCode));
  if (!aiP) return 0;
  const k = playerKey(aiP);
  const idx = seasonRotationList(teamCode).findIndex(it => playerKey(it.p) === k);
  return idx >= 0 ? idx : 0;
}
// 手動モードの1試合分の選択状態 (打順・先発・スタメン・休養セット)。元の保存編成は変更しない。
//   cursor または対戦カードが変わったら作り直す。
function seasonManualInit(g) {
  seasonEnsureRoll(g);   // allowAway/allowHome を必ず保証 (描画前に必須)
  const valid = SEASON_MANUAL_SEL && SEASON_MANUAL_SEL.cursor === SEASON.cursor
      && SEASON_MANUAL_SEL.away.team === g.away && SEASON_MANUAL_SEL.home.team === g.home
      && g.allowAway.includes(SEASON_MANUAL_SEL.away.order) && g.allowHome.includes(SEASON_MANUAL_SEL.home.order);
  if (valid) return;
  SEASON_OPEN_PICK = null; SEASON_OPEN_STARTER = null;
  const perTeam = (SEASON.schedule.length * 2) / SEASON_TEAMS.length;
  const mkSide = (teamCode, allow) => {
    const rest = new Set();
    applyTeamFilter(getBatters(), teamCode).forEach(p => {
      if (Math.random() < seasonRestProb(p, perTeam)) rest.add(playerKey(p));
    });
    const order = (allow && allow[0] != null) ? allow[0] : 0;
    const picks = seasonDefaultPicks(teamCode, order, rest);
    // 休養入替後は規定ルールでAI打順を組み直して初期表示する
    const lineOrder = seasonManualLineOrder(teamCode, order, picks);
    return { team: teamCode, order, starterIdx: seasonAutoStarterIdx(teamCode), rest, picks, lineOrder };
  };
  SEASON_MANUAL_SEL = { cursor: SEASON.cursor, away: mkSide(g.away, g.allowAway), home: mkSide(g.home, g.allowHome) };
}
// 保存チームのセット年度を取得 (mlb_team_build_v1_<TEAM> の year)
function seasonTeamYear(teamCode) {
  try { const raw = localStorage.getItem('mlb_team_build_v1_' + teamCode); if (!raw) return null; const d = JSON.parse(raw); return (d && d.year != null) ? d.year : null; }
  catch (e) { return null; }
}
// 年度フィルタ (チーム編成と同じ既定ルール): セット年度 Y に対し
//   ・同年(py===Y)は可  ・低総合力(オーダー1は70以下/オーダー2・3は65以下)は Y-2 ≤ py < Y まで可
//   ・未来(py>Y)・年度不明・3年以上前 は不可
function seasonYearAllowed(p, setYear, orderIdx) {
  if (setYear == null) return true;
  const py = p && p.year;
  if (!Number.isFinite(py)) return false;
  if (py === setYear) return true;
  const thresh = (orderIdx === 0) ? 70 : 65;
  return ((overallOf(p) || 0) <= thresh && py >= setYear - 2 && py < setYear);
}
// 同名(同チーム・同種別)で「セット年度以下のより新しい年度版」が存在するなら true (= このpは古い重複版)。
//   例) 2025セットで 2025吉田 と 2023吉田 がいる場合、2023吉田は true → 候補から除外し2025のみ使う。
function seasonHasNewerSameName(p, setYear, pool) {
  if (setYear == null || !p || !Number.isFinite(p.year)) return false;
  const nm = p.fullNameTop, ty = playerType(p);
  return pool.some(q => q !== p && q.fullNameTop === nm && playerType(q) === ty
    && Number.isFinite(q.year) && q.year <= setYear && q.year > p.year);
}
// 指定オーダーのスタメン候補 (チーム + 守備可 + オーダー制約 + 年度制約 + 同名重複除外 + 使用済み除外)
function seasonLineupCandidates(teamCode, pos, orderIdx, excludeKeys) {
  const setYear = seasonTeamYear(teamCode);
  const pool = applyTeamFilter(getBatters(), teamCode);
  return pool.filter(p =>
    canPlay(p, pos) && playerAllowedInOrder(p, orderIdx) && seasonYearAllowed(p, setYear, orderIdx)
    && !seasonHasNewerSameName(p, setYear, pool) && !excludeKeys.has(playerKey(p)));
}
// オーダーの既定スタメン: 保存編成の選手を基本に、休養中なら同ポジの非休養候補へ差し替え
function seasonDefaultPicks(teamCode, orderIdx, restSet) {
  const SLOTS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];
  const tb = getSavedTeamBuild(teamCode, orderIdx) || getSavedTeamBuild(teamCode, 0);
  const picks = {}, used = new Set();
  for (const pos of SLOTS) {
    const saved = tb && tb.batters[pos];
    const savedK = saved ? playerKey(saved) : null;
    if (saved && !restSet.has(savedK) && !used.has(savedK)) { picks[pos] = savedK; used.add(savedK); continue; }
    const cand = seasonLineupCandidates(teamCode, pos, orderIdx, used).find(p => !restSet.has(playerKey(p)));
    if (cand) { picks[pos] = playerKey(cand); used.add(playerKey(cand)); }
    else if (saved) { picks[pos] = savedK; used.add(savedK); }   // 候補なし→保存選手のまま(休養でも出す)
  }
  return picks;
}
// 規定ルールによるAI打順決定。entries=[{pos, p}](9人) を受け取り、打順スポット順(1番→9番)に並べた [{pos, p}] を返す。
//   ロジックはチーム編成の自動編成と同一 (aiAssignBattingSpots = 現代MLBの監督采配ロジック)。
function seasonAiBattingOrder(entries) {
  const spotOf = aiAssignBattingSpots(entries);
  return entries.slice().sort((a, b) => (spotOf.get(a) || 99) - (spotOf.get(b) || 99));
}
// 保存編成の打順(batterOrder)から、守備位置を打順スポット順に並べた配列を返す (ドラッグ用の初期打順)
function seasonLineOrderFor(teamCode, orderIdx) {
  const SLOTS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];
  const tb = getSavedTeamBuild(teamCode, orderIdx) || getSavedTeamBuild(teamCode, 0);
  const bo = (tb && tb.batterOrder) || {};
  const seq = [];
  for (let spot = 1; spot <= 9; spot++) { const pos = SLOTS.find(p => bo[p] === spot); if (pos && !seq.includes(pos)) seq.push(pos); }
  for (const pos of SLOTS) if (!seq.includes(pos)) seq.push(pos);
  return seq;
}
// 手動モードの初期打順: 休養入替で選手が変わっていれば規定ルールでAIにて打順を組む。
//   入替が無ければ保存編成の打順をそのまま使う(保存済みもAI編成済みのため)。
function seasonManualLineOrder(teamCode, orderIdx, picks) {
  const SLOTS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];
  const tb = getSavedTeamBuild(teamCode, orderIdx) || getSavedTeamBuild(teamCode, 0);
  let swapped = false;   // 保存スタメンと異なるpick(=休養入替)があるか
  if (tb && tb.batters) {
    for (const pos of SLOTS) {
      const savedK = tb.batters[pos] ? playerKey(tb.batters[pos]) : null;
      if (picks[pos] && savedK && picks[pos] !== savedK) { swapped = true; break; }
    }
  }
  if (!swapped) return seasonLineOrderFor(teamCode, orderIdx);
  const pool = applyTeamFilter(getBatters(), teamCode);
  const byKey = k => pool.find(p => playerKey(p) === k);
  const entries = SLOTS.filter(pos => picks[pos]).map(pos => ({ pos, p: byKey(picks[pos]) })).filter(e => e.p);
  if (entries.length !== 9) return seasonLineOrderFor(teamCode, orderIdx);
  return seasonAiBattingOrder(entries).map(e => e.pos);   // 入替後の選手で規定ルールの打順を再構成
}
// 打順の並べ替え (ドラッグ&ドロップ): side の lineOrder の fromIdx を toIdx へ移動
function seasonReorderLine(side, fromIdx, toIdx) {
  if (!SEASON_MANUAL_SEL || !SEASON_MANUAL_SEL[side]) return;
  const lo = SEASON_MANUAL_SEL[side].lineOrder;
  if (!Array.isArray(lo) || fromIdx < 0 || toIdx < 0 || fromIdx >= lo.length || toIdx >= lo.length || fromIdx === toIdx) return;
  const moved = lo.splice(fromIdx, 1)[0];
  lo.splice(toIdx, 0, moved);
  renderSeason();
}
const SEASON_POS_JP = { C: '捕', '1B': '一', '2B': '二', '3B': '三', SS: '遊', LF: '左', CF: '中', RF: '右', DH: 'DH' };
// 打者のシーズン成績(打率/本/点/盗)を取り出す (セットアップ表示・比較用)
function seasonBatStatObj(p) {
  const s = seasonActiveStores().bat[playerKey(p)];   // ポストシーズンはポストシーズン通算
  if (!s || !s.AB) return { avg: '.---', hr: 0, rbi: 0, sb: 0 };
  return { avg: fmtAvg(s.H, s.AB), hr: s.HR || 0, rbi: s.RBI || 0, sb: s.SB || 0 };
}
// 揃った成績セル (打率/本/点/盗) を右揃え列で出力
function seasonBatStatCells(p) {
  const st = seasonBatStatObj(p);
  return `<span class="slc slc-avg">${st.avg}</span><span class="slc">${st.hr}</span><span class="slc">${st.rbi}</span><span class="slc">${st.sb}</span>`;
}
// 投手のシーズン成績(防御率/勝/敗/H/S/スタ%)を取り出す
function seasonPitStatObj(p) {
  const k = playerKey(p);
  const s = seasonActiveStores().pit[k];   // ポストシーズンはポストシーズン通算 (スタミナは共通)
  const max = (p.stats && p.stats['スタミナ']) || 70;
  const cur = (SEASON && SEASON.stamina && SEASON.stamina[k] != null) ? SEASON.stamina[k] : max;
  const sta = max > 0 ? Math.round(cur / max * 100) : 0;
  if (!s || !s.outs) return { era: '-.--', w: 0, l: 0, hld: 0, sv: 0, sta };
  return { era: fmtERA(s.ER, s.outs), w: s.W || 0, l: s.L || 0, hld: s.HLD || 0, sv: s.S || 0, sta };
}
// 揃った成績セル (防/勝/敗/H/S/スタ%) を右揃え列で出力
function seasonPitStatCells(p) {
  const st = seasonPitStatObj(p);
  return `<span class="slc slc-avg">${st.era}</span><span class="slc">${st.w}</span><span class="slc">${st.l}</span><span class="slc">${st.hld}</span><span class="slc">${st.sv}</span><span class="slc slc-sta">${st.sta}%</span>`;
}
// スタメン9枠。列がそろう自作ドロップダウン (選択=ボタン / 候補=パネル)。
//   休養/怪我の選手は灰色・選択不可・ツールチップ「休養/怪我の為、選択できません。」
function seasonLineupHtml(side, teamCode, orderIdx, picks, restSet, lineOrder) {
  const orderedPos = (Array.isArray(lineOrder) && lineOrder.length === 9) ? lineOrder : seasonLineOrderFor(teamCode, orderIdx);
  const pool = applyTeamFilter(getBatters(), teamCode);
  const setYear = seasonTeamYear(teamCode);
  const byKey = k => pool.find(p => playerKey(p) === k);
  const header = `<div class="season-linehead"><span class="slh-grip"></span><span class="slh-spot">順</span><span class="slh-pos">守</span><div class="season-pickbtn season-pickhead"><span class="slc-name">選手</span><span class="slc">打率</span><span class="slc">本</span><span class="slc">点</span><span class="slc">盗</span><span class="slc-arrow"></span></div><span class="slh-card"></span></div>`;
  const rows = orderedPos.map((pos, idx) => {
    const exclude = new Set();
    orderedPos.forEach(op => { if (op !== pos && picks[op]) exclude.add(picks[op]); });
    let cands = pool.filter(p => canPlay(p, pos) && playerAllowedInOrder(p, orderIdx) && seasonYearAllowed(p, setYear, orderIdx) && !seasonHasNewerSameName(p, setYear, pool) && !exclude.has(playerKey(p)));
    if (picks[pos] && !cands.some(p => playerKey(p) === picks[pos])) { const cur = byKey(picks[pos]); if (cur) cands.push(cur); }
    cands.sort((a, b) => (overallOf(b) || 0) - (overallOf(a) || 0));
    const selP = byKey(picks[pos]);
    const open = !!(SEASON_OPEN_PICK && SEASON_OPEN_PICK.side === side && SEASON_OPEN_PICK.pos === pos);
    const btnInner = selP
      ? `<span class="slc-name">${selP.fullNameTop}</span>${seasonBatStatCells(selP)}`
      : `<span class="slc-name">(候補なし)</span><span class="slc"></span><span class="slc"></span><span class="slc"></span><span class="slc"></span>`;
    const btn = `<div class="season-pickbtn${open ? ' open' : ''}" data-pick-side="${side}" data-pick-pos="${pos}" title="クリックで選手を変更">${btnInner}<span class="slc-arrow">▼</span></div>`;
    let panel = '';
    if (open) {
      const optRows = cands.map(p => {
        const k = playerKey(p), resting = restSet.has(k), selected = (k === picks[pos]);
        const cls = 'season-pickopt' + (resting ? ' resting' : '') + (selected ? ' selected' : '');
        const attrs = resting ? ` title="休養/怪我の為、選択できません。"` : ` data-pick-side="${side}" data-pick-pos="${pos}" data-pick-key="${k}"`;
        return `<div class="${cls}"${attrs}><span class="slc-name">${p.fullNameTop}${resting ? '（休養/怪我）' : ''}</span>${seasonBatStatCells(p)}<span class="slc-arrow"></span></div>`;
      }).join('');
      panel = `<div class="season-pickpanel">${optRows || '<div class="season-pickopt">(候補なし)</div>'}</div>`;
    }
    const card = selP ? `<span class="season-linecard">${seasonPlayerLink(selP, '📋')}</span>` : '';
    return `<div class="season-linerow" data-slin-side="${side}" data-line-idx="${idx}"><span class="season-linedrag" draggable="true" title="ドラッグで打順を入れ替え">⠿</span><span class="season-linespot">${idx + 1}</span><span class="season-linepos">${SEASON_POS_JP[pos] || pos}</span><div class="season-pickwrap">${btn}${panel}</div>${card}</div>`;
  }).join('');
  return `<div class="season-lineup"><div class="season-linehdr">野手オーダー（この試合のみ／元の編成は不変・⠿をドラッグで打順入替）</div>${header}${rows}<div class="season-linenote">灰色（休養/怪我）の選手は選択できません</div></div>`;
}
function seasonSideSetupHtml(side, label, teamCode, allow) {
  if (!Array.isArray(allow) || !allow.length) allow = [0, 1, 2];
  const s = SEASON_MANUAL_SEL[side];
  const oOpts = allow.map(i => `<option value="${i}"${i === s.order ? ' selected' : ''}>オーダー${i + 1}</option>`).join('');
  const oLock = allow.length === 1 ? ' disabled' : '';
  const orderH = `<label class="season-sel">${label} 打順:
    <select data-slin-order="${side}"${oLock}>${oOpts}</select>${allow.length === 1 ? `<span class="season-lock">（オーダー${allow[0] + 1}限定）</span>` : ''}</label>`;
  // 先発投手: 自作ドロップダウン (防/勝/敗/H/S/スタ% を整列表示・比較)
  const list = seasonRotationList(teamCode);
  const selStarter = (list[s.starterIdx] || {}).p;
  const pstOpen = !!(SEASON_OPEN_STARTER && SEASON_OPEN_STARTER.side === side);
  const pstHead = `<div class="season-pstbtn season-psthead"><span class="slc-name">先発投手</span><span class="slc">防</span><span class="slc">勝</span><span class="slc">敗</span><span class="slc">H</span><span class="slc">S</span><span class="slc">スタ</span><span class="slc-arrow"></span></div>`;
  const pstBtnInner = selStarter
    ? `<span class="slc-name">${list[s.starterIdx].label}: ${selStarter.fullNameTop}</span>${seasonPitStatCells(selStarter)}`
    : `<span class="slc-name">(投手なし)</span><span class="slc"></span><span class="slc"></span><span class="slc"></span><span class="slc"></span><span class="slc"></span><span class="slc"></span>`;
  const pstBtn = `<div class="season-pstbtn${pstOpen ? ' open' : ''}" data-pst-side="${side}" title="クリックで先発を変更">${pstBtnInner}<span class="slc-arrow">▼</span></div>`;
  let pstPanel = '';
  if (pstOpen) {
    const pOptRows = list.map((it, i) => `<div class="season-pstopt${i === s.starterIdx ? ' selected' : ''}" data-pst-side="${side}" data-pst-idx="${i}"><span class="slc-name">${it.label}: ${it.p.fullNameTop}</span>${seasonPitStatCells(it.p)}<span class="slc-arrow"></span></div>`).join('');
    pstPanel = `<div class="season-pstpanel">${pOptRows || '<div class="season-pstopt">(投手なし)</div>'}</div>`;
  }
  const starterCard = selStarter ? `<span class="season-linecard">${seasonPlayerLink(selStarter, '📋')}</span>` : '';
  const starterH = `<div class="season-pstblock"><div class="season-pstlabel">${label} 先発</div><div class="season-pstwrap">${pstHead}${pstBtn}${pstPanel}</div>${starterCard}</div>`;
  return `<div class="season-setcol">
    <h4>${seasonTeamName(teamCode)}（${label}）</h4>
    ${orderH}
    ${starterH}
    ${seasonLineupHtml(side, teamCode, s.order, s.picks, s.rest, s.lineOrder)}
  </div>`;
}
function seasonManualHtml() {
  if (seasonDone()) return `<div class="season-end"><h3>🏁 シーズン終了！</h3><p>「成績モード」で結果を確認できます。</p></div>`;
  const g = SEASON.schedule[SEASON.cursor];
  seasonManualInit(g);
  return `<div class="season-manual">
    ${seasonGameHeaderHtml()}
    <div class="season-setrow">
      ${seasonSideSetupHtml('away', 'AWAY', g.away, g.allowAway)}
      ${seasonSideSetupHtml('home', 'HOME', g.home, g.allowHome)}
    </div>
    <div class="season-actions">
      <button class="btn btn-primary" data-sact="play-manual">▶ 試合開始</button>
    </div>
  </div>`;
}
function seasonAutoHtml() {
  if (seasonDone()) return `<div class="season-end"><h3>🏁 シーズン終了！</h3><p>「成績モード」で結果を確認できます。</p></div>`;
  if (SEASON_AUTORUN) {
    const tot = SEASON_AUTORUN.played + SEASON_AUTORUN.remaining;
    return `<div class="season-auto">
      <p class="season-progress">⏳ 自動進行中… <b id="sv-auto-progress">${SEASON_AUTORUN.played} / ${tot}</b> 試合（通算 ${SEASON.cursor} / ${SEASON.schedule.length}）</p>
      <div class="season-actions"><button class="btn btn-danger" data-sact="stop-auto">■ 中断してセーブ</button></div>
      <p class="season-note">進めている間も成績はこまめに保存されます。中断ボタンでいつでも止められます。</p>
      <h4 class="season-livehdr">📊 順位表（リアルタイム）</h4>
      <div class="season-tablewrap" id="sv-auto-standings">${seasonStandingsTable()}</div>
    </div>`;
  }
  const remain = SEASON.schedule.length - SEASON.cursor;
  const opts = [1, 3, 9, 30, 100, 200, 400, remain].filter((v, i, a) => v <= remain && a.indexOf(v) === i)
    .map(v => `<option value="${v}">${v === remain ? '残り全部 (' + remain + ')' : v + ' 試合'}</option>`).join('');
  return `<div class="season-auto">
    ${seasonGameHeaderHtml()}
    <p>次の試合から指定した試合数を自動で進めます（オーダー・先発は自動選定）。</p>
    <div class="season-actions">
      <label class="season-sel">進める試合数: <select id="sv-auto-count">${opts}</select></label>
      <button class="btn btn-primary" data-sact="play-auto">⏩ 自動進行</button>
    </div>
  </div>`;
}
function seasonStatsHtml() {
  const tabs = [['standings', '勝敗表'], ['team', 'チーム別成績'], ['bat', '打撃ベスト20'], ['pit', '投手ベスト20'], ['awards', '各種アワード']];
  const tabH = tabs.map(([k, l]) => `<button class="season-tab${SEASON_STATS_TAB === k ? ' on' : ''}" data-stab="${k}">${l}</button>`).join('');
  let inner = '';
  if (SEASON_STATS_TAB === 'standings') inner = seasonStandingsTable() + seasonH2HTable();
  else if (SEASON_STATS_TAB === 'team') inner = seasonTeamStatsTable();
  else if (SEASON_STATS_TAB === 'bat') inner = seasonBatLeaders();
  else if (SEASON_STATS_TAB === 'pit') inner = seasonPitLeaders();
  else if (SEASON_STATS_TAB === 'awards') inner = seasonAwardsHtml();
  return `<div class="season-stats"><div class="season-tabs">${tabH}</div><div class="season-tablewrap">${inner}</div></div>`;
}
// チームの集計(打率・防御率)
function seasonTeamAgg(team) {
  let H = 0, AB = 0, ER = 0, outs = 0;
  Object.values(SEASON.bat).forEach(s => { if (s.team === team) { H += s.H; AB += s.AB; } });
  Object.values(SEASON.pit).forEach(s => { if (s.team === team) { ER += s.ER; outs += s.outs; } });
  return { avg: avgOf(H, AB), era: eraOf(ER, outs) };
}
function seasonStandingsTable() {
  return SEASON_DIV_ORDER.map(dk => seasonStandingsTableFor(dk)).join('');
}
function seasonStandingsTableFor(divKey) {
  const divInfo = SEASON_DIVISIONS[divKey];
  const rows = divInfo.teams.map(t => {
    const st = SEASON.standings[t] || { w: 0, l: 0, d: 0, rs: 0, ra: 0, hr: 0, sb: 0, e: 0 };
    const gp = st.w + st.l + st.d;
    const pct = (st.w + st.l) > 0 ? st.w / (st.w + st.l) : 0;
    const agg = seasonTeamAgg(t);
    return { t, gp, w: st.w, l: st.l, d: st.d, pct, rs: st.rs, ra: st.ra, hr: st.hr, sb: st.sb, e: st.e, avg: agg.avg, era: agg.era };
  }).sort((a, b) => b.pct - a.pct || b.w - a.w);
  const top = rows[0] || { w: 0, l: 0 };
  const gb = (r) => { const g = ((top.w - r.w) + (r.l - top.l)) / 2; return g <= 0 ? '－' : g.toFixed(1); };
  const body = rows.map((r, i) => `<tr>
    <td>${i + 1}</td><td class="lname">${seasonTeamName(r.t)}</td><td>${r.gp}</td><td>${r.w}</td><td>${r.l}</td><td>${r.d}</td>
    <td>${fmt3(r.pct)}</td><td>${i === 0 ? '－' : gb(r)}</td><td>${r.rs}</td><td>${r.ra}</td><td>${r.hr}</td><td>${r.sb}</td>
    <td>${fmt3(r.avg)}</td><td>${fmt2(r.era)}</td><td>${r.e}</td></tr>`).join('');
  return `<h4>${divInfo.jp}</h4><table class="season-table season-standings"><thead><tr>
    <th>順位</th><th>チーム</th><th>試合</th><th>勝</th><th>敗</th><th>分</th><th>勝率</th><th>勝差</th><th>得点</th><th>失点</th><th>本塁打</th><th>盗塁</th><th>打率</th><th>防御率</th><th>失策</th>
    </tr></thead><tbody>${body}</tbody></table>`;
}
// 対戦成績マトリクス (リーグ別 15×15。各チームの 相手別 勝敗)
function seasonH2HTable() {
  seasonEnsureH2H();
  return ['AL', 'NL'].map(seasonH2HTableFor).join('');
}
function seasonH2HTableFor(lg) {
  // 勝率順にチームを並べる (勝敗表と同じ並び)
  const order = SEASON_LEAGUES[lg].slice().sort((a, b) => {
    const A = SEASON.standings[a], B = SEASON.standings[b];
    const pa = (A.w + A.l) > 0 ? A.w / (A.w + A.l) : 0, pb = (B.w + B.l) > 0 ? B.w / (B.w + B.l) : 0;
    return pb - pa || B.w - A.w;
  });
  const head = '<th class="lname">チーム＼相手</th>' + order.map(t => `<th>${seasonTeamName(t)}</th>`).join('');
  const body = order.map(a => {
    const cells = order.map(b => {
      if (a === b) return '<td class="h2h-self">—</td>';
      const r = (SEASON.h2h[a] && SEASON.h2h[a][b]) || { w: 0, l: 0, d: 0 };
      const txt = r.d > 0 ? `${r.w}-${r.l}-${r.d}` : `${r.w}-${r.l}`;
      const cls = r.w > r.l ? 'h2h-win' : (r.l > r.w ? 'h2h-lose' : '');
      return `<td class="${cls}">${txt}</td>`;
    }).join('');
    return `<tr><th class="lname">${seasonTeamName(a)}</th>${cells}</tr>`;
  }).join('');
  return `<h4>対戦成績（${SEASON_LEAGUE_JP[lg]}・同リーグ内）</h4><p class="season-note">数字は 勝-敗（引分があれば 勝-敗-分）。各行＝そのチームの相手別成績。インターリーグは省略。</p>
    <table class="season-table h2h-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}
// 成績表の選手名セル (クリックでHTMLカード表示)
function seasonNameCell(r, type) {
  const nm = String(r.name || '');
  return `<span class="player-link srow-name" title="${nm}" data-player-name="${nm}" data-player-year="${r.year || ''}" data-player-type="${type}" data-player-team="${r.team || ''}">${nm}</span>`;
}
// 打撃の列定義 (lower=小さいほど良い指標)
function batColumns() {
  return [
    { key: 'name', label: '選手', lname: true, sortable: false, get: r => r.name, fmt: (v, r) => seasonNameCell(r, 'batter') },
    { key: 'team', label: 'チーム', sortable: false, get: r => seasonTeamName(r.team), fmt: (v) => `<span class="srow-team" title="${v}">${v}</span>` },
    { key: 'avg', label: '打率', rate: true, get: r => r.avg, fmt: v => fmt3(v) },
    { key: 'G', label: '試合', get: r => r.G },
    { key: 'AB', label: '打数', get: r => r.AB },
    { key: 'H', label: '安打', get: r => r.H },
    { key: 'dbl', label: '二', get: r => r.dbl },
    { key: 'tpl', label: '三', get: r => r.tpl },
    { key: 'HR', label: '本', get: r => r.HR },
    { key: 'tb', label: '塁打', get: r => r.tb },
    { key: 'RBI', label: '打点', get: r => r.RBI },
    { key: 'R', label: '得点', get: r => r.R },
    { key: 'SO', label: '三振', get: r => r.SO },
    { key: 'BB', label: '四球', get: r => r.BB },
    { key: 'SB', label: '盗塁', get: r => r.SB },
    { key: 'CS', label: '盗塁死', lower: true, get: r => r.CS },
    { key: 'obp', label: '出塁率', rate: true, get: r => r.obp, fmt: v => fmt3(v) },
    { key: 'slg', label: '長打率', rate: true, get: r => r.slg, fmt: v => fmt3(v) },
    { key: 'ops', label: 'OPS', rate: true, get: r => r.ops, fmt: v => fmt3(v) },
    { key: 'E', label: '失策', lower: true, get: r => r.E },
  ];
}
function pitColumns() {
  return [
    { key: 'name', label: '選手', lname: true, sortable: false, get: r => r.name, fmt: (v, r) => seasonNameCell(r, 'pitcher') },
    { key: 'team', label: 'チーム', sortable: false, get: r => seasonTeamName(r.team), fmt: (v) => `<span class="srow-team" title="${v}">${v}</span>` },
    { key: 'era', label: '防御率', rate: true, lower: true, get: r => r.era, fmt: v => fmt2(v) },
    { key: 'G', label: '登板', get: r => r.G },
    { key: 'GS', label: '先発', get: r => r.GS },
    { key: 'W', label: '勝', get: r => r.W },
    { key: 'L', label: '敗', lower: true, get: r => r.L },
    { key: 'S', label: 'S', get: r => r.S },
    { key: 'HLD', label: 'H', get: r => r.HLD },
    { key: 'outs', label: '投球回', get: r => r.outs, fmt: v => ipText(v) },
    { key: 'H', label: '被安打', lower: true, get: r => r.H },
    { key: 'HR', label: '被本', lower: true, get: r => r.HR },
    { key: 'K', label: '奪三振', get: r => r.K },
    { key: 'BB', label: '与四球', lower: true, get: r => r.BB },
    { key: 'R', label: '失点', lower: true, get: r => r.R },
    { key: 'ER', label: '自責', lower: true, get: r => r.ER },
    { key: 'kbb', label: 'K/BB', rate: true, get: r => r.kbb, fmt: v => fmt2(v) },
    { key: 'whip', label: 'WHIP', rate: true, lower: true, get: r => r.whip, fmt: v => fmt2(v) },
  ];
}
// 集計に派生指標を足した行を作る
function seasonBatRows(filterFn) {
  return Object.values(SEASON.bat).filter(filterFn || (() => true)).map(s => {
    const obp = (s.PA > 0) ? (s.H + s.BB + s.HBP) / s.PA : 0;
    const tb = s.H + s.dbl + 2 * s.tpl + 3 * s.HR;
    const slg = (s.AB > 0) ? tb / s.AB : 0;
    return Object.assign({}, s, { avg: avgOf(s.H, s.AB), obp, slg, ops: obp + slg, tb });
  });
}
function seasonPitRows(filterFn) {
  return Object.values(SEASON.pit).filter(filterFn || (() => true)).map(s => {
    const whip = s.outs > 0 ? (s.BB + s.H) / (s.outs / 3) : 0;
    return Object.assign({}, s, { era: eraOf(s.ER, s.outs), whip, kbb: s.BB > 0 ? s.K / s.BB : s.K });
  });
}
// 並び替え可能なテーブルを描画 (列ヘッダクリックでソート / 上位3行に金銀銅 / ソート列を強調)
function renderSortTable(rows, cols, sort, tbl, opts) {
  opts = opts || {};
  const col = cols.find(c => c.key === sort.key && c.sortable !== false) || cols.find(c => c.sortable !== false);
  const data = rows.slice().sort((a, b) => {
    // 率指標(打率/防御率等)を並べる時は「規定到達者を常に上」に固める (チーム別成績用)
    if (opts.qualify && col.rate) {
      const qa = opts.qualify(a) ? 1 : 0, qb = opts.qualify(b) ? 1 : 0;
      if (qa !== qb) return qb - qa;
    }
    const va = Number(col.get(a)) || 0, vb = Number(col.get(b)) || 0;
    return va === vb ? 0 : (va < vb ? -1 : 1) * sort.dir;
  });
  const shown = opts.top ? data.slice(0, opts.top) : data;
  const rankHead = opts.rank ? '<th class="rankcell">順位</th>' : '';
  const head = rankHead + cols.map(c => {
    if (c.sortable === false) return `<th class="${c.lname ? 'lname' : ''}">${c.label}</th>`;
    const active = (c.key === sort.key);
    const arrow = active ? (sort.dir === 1 ? ' ▲' : ' ▼') : '';
    return `<th class="sortable${active ? ' sorted' : ''}" data-sortkey="${c.key}" data-sorttbl="${tbl}" title="クリックで並び替え">${c.label}${arrow}</th>`;
  }).join('');
  const colspan = cols.length + (opts.rank ? 1 : 0);
  const body = shown.map((r, i) => {
    const rc = i < 3 ? ' rank-' + (i + 1) : '';
    const qc = (opts.qualify && opts.qualify(r)) ? ' qualified' : '';
    const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : (i + 1);
    const rankCell = opts.rank ? `<td class="rankcell">${medal}</td>` : '';
    const tds = cols.map(c => {
      const v = c.get(r);
      const disp = c.fmt ? c.fmt(v, r) : v;
      const active = (c.key === sort.key);
      return `<td class="${c.lname ? 'lname' : ''}${active ? ' sortcol' : ''}">${disp}</td>`;
    }).join('');
    return `<tr class="srow${rc}${qc}">${rankCell}${tds}</tr>`;
  }).join('') || `<tr><td colspan="${colspan}">記録なし</td></tr>`;
  const tcls = 'season-table sortable-table' + (opts.qualify ? ' qual-table' : '');
  return `<table class="${tcls}"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}
function seasonTeamStatsTable() {
  const sel = SEASON._teamView || SEASON_TEAMS[0];
  const tabs = SEASON_TEAMS.map(t => `<button class="season-subtab${sel === t ? ' on' : ''}" data-tteam="${t}">${seasonTeamName(t)}</button>`).join('');
  // 規定打席 = 試合数 × 3.1 / 規定投球回 = 試合数 × 1.0回(=3アウト)
  const minPA = seasonRegPA(sel);
  const minOuts = seasonRegOuts(sel);
  const bats = seasonBatRows(s => s.team === sel);
  const pits = seasonPitRows(s => s.team === sel);
  const batOpts = { rank: true, qualify: r => (r.PA || 0) >= minPA };
  const pitOpts = { rank: true, qualify: r => (r.outs || 0) >= minOuts };
  return `<div class="season-subtabs">${tabs}</div>
    <p class="season-note">列クリックで並び替え。<b class="qmark">赤字</b>＝規定到達（打者 ${minPA} 打席 / 投手 ${ipText(minOuts)} 回）。率指標(打率・出塁率・長打率・OPS・防御率など)は規定到達者を上に並べます。</p>
    <h4>打撃</h4>${renderSortTable(bats, batColumns(), SEASON_SORT.tbat, 'tbat', batOpts)}
    <h4>投手</h4>${renderSortTable(pits, pitColumns(), SEASON_SORT.tpit, 'tpit', pitOpts)}`;
}
function seasonBatLeaders() {
  const cols = batColumns();
  const col = cols.find(c => c.key === SEASON_SORT.bat.key) || cols[2];
  const note = col.rate ? `規定打席(チーム試合数×3.1)到達者・${col.label}順` : `${col.label}順（全選手）`;
  const section = lg => {
    let rows = seasonBatRows(s => (s.PA || 0) > 0 && seasonLeagueOf(s.team) === lg);
    if (col.rate) rows = rows.filter(s => (s.PA || 0) >= seasonRegPA(s.team));   // 率指標は規定打席到達者のみ
    return `<h4>${SEASON_LEAGUE_JP[lg]}</h4>` + renderSortTable(rows, cols, SEASON_SORT.bat, 'bat', { rank: true, top: 20 });
  };
  return `<p class="season-note">${note}（列クリックで並び替え・リーグ別ベスト20）</p>` + section('AL') + section('NL');
}
function seasonPitLeaders() {
  const cols = pitColumns();
  const col = cols.find(c => c.key === SEASON_SORT.pit.key) || cols[2];
  const note = col.rate ? `規定投球回(チーム試合数×1.0)到達者・${col.label}順` : `${col.label}順（全投手）`;
  const section = lg => {
    let rows = seasonPitRows(s => (s.outs || 0) > 0 && seasonLeagueOf(s.team) === lg);
    if (col.rate) rows = rows.filter(s => (s.outs || 0) >= seasonRegOuts(s.team));   // 率指標は規定投球回到達者のみ
    return `<h4>${SEASON_LEAGUE_JP[lg]}</h4>` + renderSortTable(rows, cols, SEASON_SORT.pit, 'pit', { rank: true, top: 20 });
  };
  return `<p class="season-note">${note}（列クリックで並び替え・リーグ別ベスト20）</p>` + section('AL') + section('NL');
}
// ===== 各種アワード =====
// 選手の主守備位置 (守備イニング最多の位置 / なければ position から)
function seasonPrimaryPos(p) {
  const FIELD = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF'];
  // カードのポジション表記が明示的にDHならDH
  if (p && (p.position === '指名打者' || /ＤＨ|DH/i.test(String(p.position || '')))) return 'DH';
  const fl = (p && p.drs ? p.drs : []).filter(d => FIELD.includes(d.pos) && (Number(d.innings) || 0) > 0).sort((a, b) => (b.innings || 0) - (a.innings || 0));
  if (fl.length) {
    // DH判定: 出場試合が多いのに守備イニングが極端に少ない(平均2回/試合未満 ≒ ほぼ守備につかない)→ DH専属とみなす
    const totalInn = fl.reduce((s, d) => s + (Number(d.innings) || 0), 0);
    const games = getCardGamesOf(p) || 0;
    if (games >= 40 && totalInn < games * 2) return 'DH';
    return fl[0].pos;
  }
  // 守備DRSが無い → ポジション表記から (該当なしはDH)
  const map = { '捕手': 'C', '一塁手': '1B', '二塁手': '2B', '三塁手': '3B', '遊撃手': 'SS', '左翼手': 'LF', '中堅手': 'CF', '右翼手': 'RF', '指名打者': 'DH' };
  return (p && map[p.position]) || 'DH';
}
function seasonDrsAt(p, pos) {
  if (!p || !p.drs) return null;
  const d = p.drs.find(x => x.pos === pos && (Number(x.innings) || 0) > 0);
  return d ? (Number(d.value) || 0) : null;
}
// アワードカード (顔写真 + 名前 + チーム + 成績)。クリックで選手カード表示。
//   league ('AL'|'NL') を渡すと、カード右上に獲得リーグのチップを表示する。
function seasonAwardCard(label, p, statText, badge, league) {
  const photo = (p && p.photo) ? `<img src="${p.photo}" alt="" onerror="this.style.display='none'">` : '<span class="award-noimg">👤</span>';
  const link = p ? `data-player-name="${p.fullNameTop}" data-player-year="${p.year || ''}" data-player-type="${playerType(p)}" data-player-team="${p.team || ''}"` : '';
  const lgChip = league ? `<span class="award-lg-chip lg-${league}" title="${SEASON_LEAGUE_JP[league] || league}">${league}</span>` : '';
  return `<div class="award-card${p ? ' player-link' : ''}" ${link} title="${p ? 'クリックで選手カードを表示' : ''}">
    ${lgChip}
    <div class="award-label">${badge ? `<span class="award-pos">${badge}</span>` : ''}${label}</div>
    <div class="award-photo">${photo}</div>
    <div class="award-name${p ? ' ' + longNameClass(p.fullNameTop) : ''}">${p ? p.fullNameTop : '—'}</div>
    <div class="award-team">${p ? seasonTeamName(p.team) : ''}</div>
    <div class="award-stat">${statText}</div>
  </div>`;
}
function seasonAwardsHtml() {
  const byKey = {};
  [...getBatters(), ...getPitchers()].forEach(p => { byKey[playerKey(p)] = p; });
  const batAll = Object.keys(SEASON.bat).map(k => {
    const s = SEASON.bat[k];
    const obp = (s.PA > 0) ? (s.H + s.BB + s.HBP) / s.PA : 0;
    const tb = s.H + s.dbl + 2 * s.tpl + 3 * s.HR;
    const slg = (s.AB > 0) ? tb / s.AB : 0;
    return { s, p: byKey[k], avg: avgOf(s.H, s.AB), ops: obp + slg };
  }).filter(e => e.p);
  const pitAll = Object.keys(SEASON.pit).map(k => ({ s: SEASON.pit[k], p: byKey[k], era: eraOf(SEASON.pit[k].ER, SEASON.pit[k].outs) })).filter(e => e.p);
  if (!batAll.length && !pitAll.length) return '<p class="season-note">まだ成績データがありません。シーズンを進めてください。</p>';
  // シルバースラッガー判定用に、各打者の守備イニング(実測+遡及補完)を付与する。
  //   守備イニングは実測(SEASON.fieldInn)。実測が無い消化済み試合分は「編成での守備位置 × 出場試合 × 9回」で補完。
  const REQ_INN = 800;
  const FIELD_POS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];
  const buildPosByKey = seasonBuildPosMap();
  const fieldInn = SEASON.fieldInn || {};
  const effInnOf = (k) => {
    const real = fieldInn[k] || {};
    let realTotal = 0; FIELD_POS.forEach(pp => realTotal += (real[pp] || 0));
    const G = (SEASON.bat[k] && SEASON.bat[k].G) || 0;
    const untracked = Math.max(0, G - realTotal / 9);
    const bpos = buildPosByKey[k] || null;
    const out = {};
    FIELD_POS.forEach(pp => { out[pp] = (real[pp] || 0) + (bpos === pp ? untracked * 9 : 0); });
    out.OF = out.LF + out.CF + out.RF;
    return out;
  };
  batAll.forEach(e => { e.inn = effInnOf(playerKey(e.p)); });
  // リーグごとに選出 (打撃成績・投手成績・各タイトルはリーグ別)
  const leagueHtml = lg => seasonAwardsForLeague(lg,
    batAll.filter(e => seasonLeagueOf(e.s.team) === lg),
    pitAll.filter(e => seasonLeagueOf(e.s.team) === lg), REQ_INN);
  return `<div class="awards">
    ${['AL', 'NL'].map(lg => `<div class="award-league-block"><h3 class="award-league lg-${lg}"><span class="award-lg-chip lg-${lg}">${lg}</span>${SEASON_LEAGUE_JP[lg]}</h3>${leagueHtml(lg)}</div>`).join('')}
    ${seasonDone() ? '' : '<p class="season-note">※シーズン途中の暫定です（確定はシーズン終了後）。</p>'}
  </div>`;
}
// 1リーグ分のアワード(個人タイトル + シルバースラッガー)を描画。batE/pitEは当該リーグの選手のみ(各eに e.inn 付与済)。
//   各アワードカードに獲得リーグ(lg)のチップを表示する。
function seasonAwardsForLeague(lg, batE, pitE, REQ_INN) {
  const pick = (pool, metric, lower) => pool.length ? pool.slice().sort((a, b) => (metric(b) - metric(a)) * (lower ? -1 : 1))[0] : null;
  const qB = pool => { const q = pool.filter(e => (e.s.PA || 0) >= seasonRegPA(e.s.team)); return q.length ? q : pool; };
  const qP = pool => { const q = pool.filter(e => (e.s.outs || 0) >= seasonRegOuts(e.s.team)); return q.length ? q : pool; };
  const c = (label, e, txt, badge) => seasonAwardCard(label, e ? e.p : null, e ? txt(e) : '—', badge, e ? lg : null);
  const titles = [
    c('首位打者', pick(qB(batE), e => e.avg), e => fmt3(e.avg)),
    c('本塁打王', pick(batE, e => e.s.HR || 0), e => (e.s.HR || 0) + '本'),
    c('打点王', pick(batE, e => e.s.RBI || 0), e => (e.s.RBI || 0) + '打点'),
    c('盗塁王', pick(batE, e => e.s.SB || 0), e => (e.s.SB || 0) + '盗'),
    c('最高OPS', pick(qB(batE), e => e.ops), e => fmt3(e.ops)),
    c('最優秀防御率', pick(qP(pitE), e => e.era, true), e => fmt2(e.era)),
    c('最多勝', pick(pitE, e => e.s.W || 0), e => (e.s.W || 0) + '勝'),
    c('最多セーブ', pick(pitE, e => e.s.S || 0), e => (e.s.S || 0) + 'S'),
    c('最多ホールド', pick(pitE, e => e.s.HLD || 0), e => (e.s.HLD || 0) + 'H'),
  ].join('');
  const batQ = batE.filter(e => (e.s.PA || 0) >= seasonRegPA(e.s.team));   // 規定打席到達者のみ
  const bat = '<span class="silver-bat">🏏</span>';
  const bestInn = cond => { const pl = batQ.filter(e => e.inn && cond(e.inn)).sort((a, b) => b.ops - a.ops); return pl[0] || null; };
  const ssLine = e => `${fmt3(e.avg)} ${e.s.HR || 0}本 ${e.s.RBI || 0}点 OPS${fmt3(e.ops)}`;
  const ssDef = [['C', '捕'], ['1B', '一'], ['2B', '二'], ['3B', '三'], ['SS', '遊']];
  let ss = ssDef.map(([pos, jp]) => c('', bestInn(inn => inn[pos] >= REQ_INN), ssLine, bat + jp));
  const ofPl = batQ.filter(e => e.inn && e.inn.OF >= REQ_INN).sort((a, b) => b.ops - a.ops).slice(0, 3);
  for (let i = 0; i < 3; i++) ss.push(seasonAwardCard('', ofPl[i] ? ofPl[i].p : null, ofPl[i] ? ssLine(ofPl[i]) : '—', bat + 'OF', ofPl[i] ? lg : null));
  const dhBest = bestInn(inn => inn.DH >= REQ_INN);
  if (dhBest) ss.push(seasonAwardCard('', dhBest.p, ssLine(dhBest), bat + 'DH', lg));
  return `<h4 class="award-h">🏆 個人タイトル（${SEASON_LEAGUE_JP[lg]}）</h4>
    <div class="award-grid">${titles}</div>
    <h4 class="award-h"><span class="silver-bat">🏏</span> シルバースラッガー賞（${SEASON_LEAGUE_JP[lg]}・打撃ベストナイン・規定打席かつ各守備位置800回以上）</h4>
    <div class="award-grid award-ss">${ss.join('')}</div>`;
}

// シーズン画面のクリック処理 (委譲)
// 手動モードのプルダウン変更 (打順 / 先発 / スタメン枠)
function seasonHandleChange(e) {
  const t = e.target;
  if (!t || t.tagName !== 'SELECT' || !SEASON_MANUAL_SEL) return;
  const g = SEASON.schedule[SEASON.cursor];
  if (!g) return;
  if (t.dataset.slinOrder) {
    const side = t.dataset.slinOrder, s = SEASON_MANUAL_SEL[side];
    s.order = +t.value;
    const tc = side === 'away' ? g.away : g.home;
    s.picks = seasonDefaultPicks(tc, s.order, s.rest);   // 新オーダーの既定スタメンへ
    s.lineOrder = seasonLineOrderFor(tc, s.order);       // 打順も新オーダーの並びへ
    SEASON_OPEN_PICK = null; SEASON_OPEN_STARTER = null;
    renderSeason();
  } else if (t.dataset.slinStarter) {
    SEASON_MANUAL_SEL[t.dataset.slinStarter].starterIdx = +t.value;   // 状態に保持 (再描画不要)
  } else if (t.dataset.slinPos) {
    SEASON_MANUAL_SEL[t.dataset.slinSide].picks[t.dataset.slinPos] = t.value;
    renderSeason();   // 重複候補を更新
  }
}
function seasonHandleClick(e) {
  // スタメン選択ドロップダウン: 候補をクリック → 選択
  const pickOpt = e.target.closest('.season-pickopt[data-pick-key]');
  if (pickOpt) {
    const side = pickOpt.dataset.pickSide, pos = pickOpt.dataset.pickPos, key = pickOpt.dataset.pickKey;
    if (SEASON_MANUAL_SEL && SEASON_MANUAL_SEL[side]) SEASON_MANUAL_SEL[side].picks[pos] = key;
    SEASON_OPEN_PICK = null; renderSeason(); return;
  }
  // 選択ボタンをクリック → 開閉トグル
  const pickBtn = e.target.closest('.season-pickbtn');
  if (pickBtn) {
    const side = pickBtn.dataset.pickSide, pos = pickBtn.dataset.pickPos;
    SEASON_OPEN_PICK = (SEASON_OPEN_PICK && SEASON_OPEN_PICK.side === side && SEASON_OPEN_PICK.pos === pos) ? null : { side, pos };
    SEASON_OPEN_STARTER = null; renderSeason(); return;
  }
  // 先発選択ドロップダウン: 候補クリック → 選択
  const pstOpt = e.target.closest('.season-pstopt[data-pst-idx]');
  if (pstOpt) {
    const side = pstOpt.dataset.pstSide;
    if (SEASON_MANUAL_SEL && SEASON_MANUAL_SEL[side]) SEASON_MANUAL_SEL[side].starterIdx = +pstOpt.dataset.pstIdx;
    SEASON_OPEN_STARTER = null; renderSeason(); return;
  }
  // 先発選択ボタン → 開閉トグル
  const pstBtn = e.target.closest('.season-pstbtn');
  if (pstBtn && pstBtn.dataset.pstSide) {
    const side = pstBtn.dataset.pstSide;
    SEASON_OPEN_STARTER = (SEASON_OPEN_STARTER && SEASON_OPEN_STARTER.side === side) ? null : { side };
    SEASON_OPEN_PICK = null; renderSeason(); return;
  }
  // パネル外をクリック → 開いていれば閉じる (この後の処理は続行)
  let closedPick = false;
  if (SEASON_OPEN_PICK && !e.target.closest('.season-pickpanel')) { SEASON_OPEN_PICK = null; closedPick = true; }
  if (SEASON_OPEN_STARTER && !e.target.closest('.season-pstpanel')) { SEASON_OPEN_STARTER = null; closedPick = true; }
  const sv = e.target.closest('[data-sv]');
  if (sv) { SEASON_VIEW = sv.dataset.sv; if (SEASON_VIEW === 'stats') SEASON_STATS_TAB = 'standings'; renderSeason(); return; }
  const stab = e.target.closest('[data-stab]');
  if (stab) { SEASON_STATS_TAB = stab.dataset.stab; renderSeason(); return; }
  const pstab = e.target.closest('[data-pstab]');
  if (pstab) { SEASON_PS_TAB = pstab.dataset.pstab; renderSeason(); return; }
  const psteam = e.target.closest('[data-psteam]');
  if (psteam) { SEASON_PS_TEAM = psteam.dataset.psteam; renderSeason(); return; }
  const pssort = e.target.closest('[data-pssort]');
  if (pssort) {
    const parts = pssort.dataset.pssort.split(':'), which = parts[0], key = parts[1], cur = SEASON_PS_SORT[which];
    if (cur.key === key) cur.dir = -cur.dir;
    else { cur.key = key; const lower = (which === 'pit') ? ['era', 'L', 'BB', 'H'] : ['SO']; cur.dir = lower.indexOf(key) >= 0 ? 1 : -1; }
    renderSeason();
    return;
  }
  const tteam = e.target.closest('[data-tteam]');
  if (tteam) { SEASON._teamView = tteam.dataset.tteam; renderSeason(); return; }
  // 成績表の列ヘッダクリック → 並び替え
  const sortTh = e.target.closest('[data-sortkey]');
  if (sortTh) {
    const tbl = sortTh.dataset.sorttbl, key = sortTh.dataset.sortkey, cur = SEASON_SORT[tbl];
    if (cur) {
      if (cur.key === key) { cur.dir = -cur.dir; }
      else {
        cur.key = key;
        const cols = (tbl === 'pit' || tbl === 'tpit') ? pitColumns() : batColumns();
        const c = cols.find(x => x.key === key);
        cur.dir = (c && c.lower) ? 1 : -1;   // 小さいほど良い指標は昇順から
      }
      renderSeason();
    }
    return;
  }
  const act = e.target.closest('[data-sact]');
  if (!act) { if (closedPick) renderSeason(); return; }
  if (act.dataset.sact === 'play-manual') {
    const g = SEASON.schedule[SEASON.cursor];
    seasonManualInit(g);
    const a = SEASON_MANUAL_SEL.away, h = SEASON_MANUAL_SEL.home;
    const aList = seasonRotationList(g.away), hList = seasonRotationList(g.home);
    const sel = {
      awayOrder: a.order,
      homeOrder: h.order,
      awayStarter: (aList[a.starterIdx] || {}).p || null,
      homeStarter: (hList[h.starterIdx] || {}).p || null,
      awayPicks: a.picks,
      homePicks: h.picks,
      awayLineOrder: a.lineOrder,
      homeLineOrder: h.lineOrder,
      awayRest: a.rest,
      homeRest: h.rest,
    };
    seasonStartCurrentGame(false, sel);
  } else if (act.dataset.sact === 'play-auto') {
    const cnt = +(document.querySelector('#sv-auto-count') || {}).value || 1;
    seasonStartAutoRun(cnt);
  } else if (act.dataset.sact === 'stop-auto') {
    if (SEASON_AUTORUN) SEASON_AUTORUN.stop = true;
  } else if (act.dataset.sact === 'play-postseason') {
    G.silent = true;
    try { seasonPostseasonPlayAll(); } catch (err) { console.error('postseason', err); }
    G.silent = false;
    SEASON_VIEW = 'postseason';
    showScreen('season');
    renderSeason();
  } else if (act.dataset.sact === 'play-ps-one') {
    G.silent = true;
    let st = null;
    try { st = seasonPostseasonStepAuto(); } catch (err) { console.error('postseason', err); }
    G.silent = false;
    if (st && st.error) alert('試合を組めませんでした（カード不足の可能性）。');
    SEASON_VIEW = 'postseason';
    showScreen('season');
    renderSeason();
  } else if (act.dataset.sact === 'play-ps-manual') {
    const ng = seasonPostseasonNextGame();
    if (!ng) { SEASON_VIEW = 'postseason'; renderSeason(); return; }
    seasonPostseasonManualInit(ng);
    const a = SEASON_MANUAL_SEL.away, h = SEASON_MANUAL_SEL.home;
    const aList = seasonRotationList(ng.away), hList = seasonRotationList(ng.home);
    const sel = {
      awayOrder: a.order, homeOrder: h.order,
      awayStarter: (aList[a.starterIdx] || {}).p || null,
      homeStarter: (hList[h.starterIdx] || {}).p || null,
      awayPicks: a.picks, homePicks: h.picks,
      awayLineOrder: a.lineOrder, homeLineOrder: h.lineOrder,
      awayRest: a.rest, homeRest: h.rest,
    };
    seasonStartPostseasonManual(sel);
  } else if (act.dataset.sact === 'reset-postseason') {
    if (confirm('ポストシーズンの組み合わせと結果をリセットして、シードからやり直しますか？')) {
      seasonPostseasonInit();
      renderSeason();
    }
  }
}

// 自動進行 (1試合ずつ非同期で進める。描画は抑制し、中断ボタンで停止可能)
function seasonStartAutoRun(n) {
  SEASON_AUTORUN = { remaining: n, played: 0, stop: false };
  SEASON_VIEW = 'auto';
  renderSeason();
  setTimeout(seasonAutoStep, 10);
}
function seasonAutoStep() {
  const run = SEASON_AUTORUN;
  if (!run || run.stop || run.remaining <= 0 || SEASON.cursor >= SEASON.schedule.length) { seasonFinishAutoRun(); return; }
  // 時間バジェット方式: 1ティックで約60ms分の試合をまとめて実行してから画面へ制御を返す。
  //   (1試合ごとの setTimeout 往復・順位表再構築・セーブをなくし、大幅に高速化)
  const t0 = performance.now();
  G.silent = true;
  let ok = true;
  try {
    while (run.remaining > 0 && !run.stop && SEASON.cursor < SEASON.schedule.length) {
      ok = seasonStartCurrentGame(true);
      if (ok && G.ended) { seasonRecordCurrentGame(); run.played++; }
      run.remaining--;
      if (!ok) break;                              // 試合を組めない(未編成等) → 中断
      if (performance.now() - t0 > 60) break;      // UI応答性のため一旦譲る
    }
  } catch (e) { console.error('seasonAutoStep', e); ok = false; }
  G.silent = false;
  if (!ok) { seasonFinishAutoRun(); return; }
  if (run.played - (run.lastSave || 0) >= 20) { saveSeason(); run.lastSave = run.played; }   // 20試合ごとに保険セーブ
  const el = document.querySelector('#sv-auto-progress');
  if (el) el.textContent = run.played + ' / ' + (run.played + run.remaining);
  const stEl = document.querySelector('#sv-auto-standings');
  if (stEl) stEl.innerHTML = seasonStandingsTable();   // 順位表をリアルタイム更新 (ティックごと)
  setTimeout(seasonAutoStep, 0);
}
function seasonFinishAutoRun() {
  SEASON_AUTORUN = null;
  G.silent = false;
  saveSeason();
  SEASON_VIEW = 'auto';
  showScreen('season');
  renderSeason();
}
// 結果画面でシーズン用ボタンを押したとき
function seasonAfterManualResult(toHub) {
  // ポストシーズンの手動試合: 成績記録 + ブラケット反映
  if (G.seasonCtx && G.seasonCtx.postseason && G.seasonCtx.psNext) {
    const ng = G.seasonCtx.psNext;
    const dec = computePitcherDecisions();
    const winner = dec.winSide === 'away' ? ng.away : ng.home;   // 引分はホーム勝ち(稀)
    seasonRecordPostseasonGame(ng.away, ng.home, ng.seriesId === 'WS');
    seasonPostseasonApplyResult(ng, winner, dec.sA, dec.sH);
    SEASON_VIEW = 'postseason';
    showScreen('season');
    renderSeason();
    return;
  }
  seasonRecordCurrentGame();
  SEASON_VIEW = toHub ? 'menu' : 'manual';
  showScreen('season');
  renderSeason();
}

// ============== ポストシーズン (MLB方式12球団) ==============
// 各リーグ 地区優勝3 + ワイルドカード3。ワイルドカード(3戦2勝)→地区(5戦3勝)→リーグ優勝(7戦4勝)→ワールドS(7戦4勝)。
// ポストシーズンの試合はレギュラー成績(standings/bat/pit)には一切加算しない。
function seasonTeamPct(t) { const s = SEASON.standings[t] || { w: 0, l: 0 }; return (s.w + s.l) > 0 ? s.w / (s.w + s.l) : 0; }
function seasonStandCmp(a, b) { return seasonTeamPct(b) - seasonTeamPct(a) || (((SEASON.standings[b] || {}).w || 0) - ((SEASON.standings[a] || {}).w || 0)); }
// リーグの6シード [#1..#3=地区優勝(勝率順) / #4..#6=ワイルドカード(勝率順)]
function seasonPostseasonSeeds(lg) {
  const divWinners = SEASON_LEAGUE_DIVS[lg].map(dk => SEASON_DIVISIONS[dk].teams.slice().sort(seasonStandCmp)[0]);
  divWinners.sort(seasonStandCmp);
  const wc = SEASON_LEAGUES[lg].filter(t => !divWinners.includes(t)).sort(seasonStandCmp).slice(0, 3);
  return divWinners.concat(wc);
}
function seasonPostseasonInit() {
  SEASON.postseason = { seeds: { AL: seasonPostseasonSeeds('AL'), NL: seasonPostseasonSeeds('NL') },
    series: {}, champion: null, bat: {}, pit: {}, wsGames: [], mvp: null };
  saveSeason();
}
// 旧セーブ(成績フィールドなし)向けの遅延初期化
function seasonPostseasonEnsure() {
  const ps = SEASON.postseason; if (!ps) return;
  if (!ps.bat) ps.bat = {}; if (!ps.pit) ps.pit = {}; if (!ps.wsGames) ps.wsGames = []; if (!('mvp' in ps)) ps.mvp = null;
}
function psHomeForGame(bo, gi) {
  if (bo === 3) return true;                                    // WC: 上位が全試合ホスト
  if (bo === 5) return [true, true, false, false, true][gi];    // 2-2-1
  return [true, true, false, false, false, true, true][gi];     // 2-3-2
}
// ブラケット定義: 各シリーズの id / bo / hi / lo (依存未確定は hi,lo が null)。ラウンド順に並ぶ。
function seasonPostseasonBracket() {
  const ps = SEASON.postseason, S = ps.series, seeds = ps.seeds;
  const W = id => (S[id] && S[id].done) ? S[id].winner : null;
  const seedIdx = (lg, t) => seeds[lg].indexOf(t);
  const sp = [];
  for (const lg of ['AL', 'NL']) { const sd = seeds[lg]; sp.push({ id: lg + '_WC1', bo: 3, hi: sd[2], lo: sd[5] }); sp.push({ id: lg + '_WC2', bo: 3, hi: sd[3], lo: sd[4] }); }
  for (const lg of ['AL', 'NL']) { const sd = seeds[lg]; sp.push({ id: lg + '_DS1', bo: 5, hi: sd[0], lo: W(lg + '_WC2') }); sp.push({ id: lg + '_DS2', bo: 5, hi: sd[1], lo: W(lg + '_WC1') }); }
  for (const lg of ['AL', 'NL']) { const d1 = W(lg + '_DS1'), d2 = W(lg + '_DS2'); let lh = null, ll = null; if (d1 && d2) { lh = seedIdx(lg, d1) <= seedIdx(lg, d2) ? d1 : d2; ll = lh === d1 ? d2 : d1; } sp.push({ id: lg + '_LCS', bo: 7, hi: lh, lo: ll }); }
  const alC = W('AL_LCS'), nlC = W('NL_LCS'); let wh = null, wl = null; if (alC && nlC) { wh = seasonStandCmp(alC, nlC) <= 0 ? alC : nlC; wl = wh === alC ? nlC : alC; }
  sp.push({ id: 'WS', bo: 7, hi: wh, lo: wl });
  return sp;
}
// 次にプレイすべき1戦 {seriesId,gi,hi,lo,bo,away,home}。全完了なら null。
function seasonPostseasonNextGame() {
  if (!SEASON.postseason) return null;
  for (const spec of seasonPostseasonBracket()) {
    if (!spec.hi || !spec.lo) continue;                 // 依存未確定
    const s = SEASON.postseason.series[spec.id];
    if (s && s.done) continue;
    const gi = s ? s.games.length : 0;
    const hiHome = psHomeForGame(spec.bo, gi);
    return { seriesId: spec.id, gi, hi: spec.hi, lo: spec.lo, bo: spec.bo,
      home: hiHome ? spec.hi : spec.lo, away: hiHome ? spec.lo : spec.hi };
  }
  return null;
}
// 決着した1戦をブラケットへ反映 (勝数加算・先取で done・WS完了で champion+MVP)
function seasonPostseasonApplyResult(ng, winner, sA, sH) {
  const ps = SEASON.postseason, S = ps.series;
  let s = S[ng.seriesId];
  if (!s) s = S[ng.seriesId] = { hi: ng.hi, lo: ng.lo, bo: ng.bo, winsHi: 0, winsLo: 0, winner: null, games: [], done: false };
  s.games.push({ away: ng.away, home: ng.home, sA, sH, winner });
  if (winner === s.hi) s.winsHi++; else s.winsLo++;
  const need = (s.bo + 1) / 2;
  if (s.winsHi >= need || s.winsLo >= need) { s.done = true; s.winner = s.winsHi > s.winsLo ? s.hi : s.lo; }
  if (ng.seriesId === 'WS' && s.done) { ps.champion = s.winner; seasonComputeWSMVP(); }
  saveSeason();
}
// ポストシーズン1試合の成績・スタミナを記録 (レギュラー成績には加算しない)。isWSならボックスも保存。
function seasonRecordPostseasonGame(away, home, isWS) {
  const ps = SEASON.postseason;
  seasonPostseasonEnsure();
  const dec = computePitcherDecisions();
  seasonAccumBatting(away, 'away', null, ps.bat);
  seasonAccumBatting(home, 'home', null, ps.bat);
  seasonAccumPitching(away, 'away', dec.pitcherRoles, ps.pit);
  seasonAccumPitching(home, 'home', dec.pitcherRoles, ps.pit);
  seasonUpdateStamina(away, 'away');   // スタミナ消費/回復を反映 (試合またぎ)
  seasonUpdateStamina(home, 'home');
  if (isWS) ps.wsGames.push(seasonCaptureWSGame(away, home, dec));
}
// WSの1戦をボックススコア+得点経過として抽出
function seasonCaptureWSGame(away, home, dec) {
  const innings = playedInnings();
  const line = { away: [], home: [] };
  for (let i = 0; i < innings; i++) { line.away.push(G.score.away[i] ?? 0); line.home.push(G.score.home[i] ?? 0); }
  const batBox = side => {
    const all = (G.batterStats[side] || []).concat(G.subLog[side] || []), seen = {};
    all.forEach(bs => {
      if (!bs || !bs.player) return;
      const pa = (bs.AB || 0) + (bs.BB || 0) + (bs.HBP || 0) + (bs.SAC || 0);
      if (pa <= 0 && !(bs.R || 0) && !(bs.SB || 0)) return;
      const k = playerKey(bs.player), e = seen[k] || (seen[k] = { key: k, name: bs.player.fullNameTop, AB: 0, H: 0, dbl: 0, tpl: 0, HR: 0, RBI: 0, R: 0, BB: 0, SO: 0, SB: 0 });
      e.AB += bs.AB || 0; e.H += bs.H || 0; e.dbl += bs.doubles || 0; e.tpl += bs.triples || 0; e.HR += bs.HR || 0;
      e.RBI += bs.RBI || 0; e.R += bs.R || 0; e.BB += bs.BB || 0; e.SO += bs.K || 0; e.SB += bs.SB || 0;
    });
    return Object.values(seen);
  };
  const pitBox = side => (G.pitcherLog[side] || []).filter(lg => lg && lg.pitcher && ((lg.battersFaced || 0) > 0 || (lg.outs || 0) > 0)).map(lg =>
    ({ key: playerKey(lg.pitcher), name: lg.pitcher.fullNameTop, outs: lg.outs || 0, H: lg.hits || 0, ER: lg.earnedRuns || 0, K: lg.K || 0, BB: lg.BB || 0, dec: dec.pitcherRoles.get(lg) || '' }));
  // 各HRに「ポストシーズン通算号数」を割り当てる (postseason.bat は本試合分を加算済み → 事前値 + 試合内累積)
  const psbat = (SEASON.postseason && SEASON.postseason.bat) || {};
  const thisGameHR = {}; (G.hrEvents || []).forEach(e => { thisGameHR[e.batterKey] = (thisGameHR[e.batterKey] || 0) + 1; });
  const running = {};
  const hr = (G.hrEvents || []).map(e => {
    const key = e.batterKey, totalAfter = (psbat[key] && psbat[key].HR) || 0, prior = totalAfter - (thisGameHR[key] || 0);
    running[key] = (running[key] || 0) + 1;
    return { name: e.batterName || e.batter, team: e.batterTeam, inning: e.inning, top: e.top, runs: e.runs, num: prior + running[key] };
  });
  return {
    away, home, sA: G.score.away.reduce((a, b) => a + b, 0), sH: G.score.home.reduce((a, b) => a + b, 0),
    line, hitsA: G.hits.away, hitsH: G.hits.home, innings, hr,
    win: dec.winPitcher ? dec.winPitcher.pitcher.fullNameTop : '', lose: dec.losePitcher ? dec.losePitcher.pitcher.fullNameTop : '', save: dec.savePitcher ? dec.savePitcher.pitcher.fullNameTop : '',
    box: { away: { bat: batBox('away'), pit: pitBox('away') }, home: { bat: batBox('home'), pit: pitBox('home') } },
  };
}
// ワールドシリーズMVP (優勝チームのWS成績から合成スコア最大)
function seasonComputeWSMVP() {
  const ps = SEASON.postseason;
  if (!ps.wsGames || !ps.wsGames.length || !ps.champion) { ps.mvp = null; return; }
  const champ = normalizeTeam(ps.champion), bat = {}, pit = {};
  ps.wsGames.forEach(g => ['away', 'home'].forEach(side => {
    const team = side === 'away' ? g.away : g.home;
    if (normalizeTeam(team) !== champ) return;
    g.box[side].bat.forEach(b => { const e = bat[b.name] || (bat[b.name] = { name: b.name, AB: 0, H: 0, HR: 0, RBI: 0, R: 0, BB: 0 }); e.AB += b.AB; e.H += b.H; e.HR += b.HR; e.RBI += b.RBI; e.R += b.R; e.BB += b.BB; });
    g.box[side].pit.forEach(p => { const e = pit[p.name] || (pit[p.name] = { name: p.name, outs: 0, ER: 0, K: 0, W: 0, S: 0 }); e.outs += p.outs; e.ER += p.ER; e.K += p.K; if (p.dec === 'W') e.W++; if (p.dec === 'S') e.S++; });
  }));
  let best = null, bestScore = -Infinity, bestLine = '';
  Object.values(bat).forEach(b => { const sc = b.H + 2 * b.HR + b.RBI + b.R + b.BB / 2; if (sc > bestScore) { bestScore = sc; best = b.name; bestLine = `打率${fmt3(avgOf(b.H, b.AB))}・${b.HR}本・${b.RBI}打点・${b.R}得点`; } });
  Object.values(pit).forEach(p => { const ip = p.outs / 3, sc = 3 * p.W + 2 * p.S + p.K / 3 + ip / 3 - p.ER; if (sc > bestScore) { bestScore = sc; best = p.name; bestLine = `${p.W}勝${p.S}S・${ipText(p.outs)}回・自責${p.ER}・${p.K}奪三振`; } });
  ps.mvp = best ? { name: best, team: ps.champion, line: bestLine } : null;
}
// 自動: 1試合をオートで実行し勝者と得点を返す。引き分けは再試合(決着必須)。決着試合は記録する。
function playPostseasonGame(away, home, isWS) {
  seasonEnsureBuild(away); seasonEnsureBuild(home);   // 未保存チームは自動編成で補う
  for (let attempt = 0; attempt < 6; attempt++) {
    const as = seasonAutoStarter(away, seasonRotIdx(away)), hs = seasonAutoStarter(home, seasonRotIdx(home));
    seasonAdvanceRot(away); seasonAdvanceRot(home);
    const e1 = buildSeasonSideSetup('away', away, 0, as, true);
    const e2 = e1 ? null : buildSeasonSideSetup('home', home, 0, hs, true);
    if (e1 || e2) return null;   // カード不足
    G.innings = 9; G.seasonMode = true; G.seasonCtx = { away, home, awayOrder: 0, homeOrder: 0, auto: true, postseason: true };
    beginGame();
    let guard = 0;
    while (!G.ended && guard++ < 4000) { const p = autoPick(); if (!p) break; pitchOne(p, true); }
    const sA = G.score.away.reduce((a, b) => a + b, 0), sH = G.score.home.reduce((a, b) => a + b, 0);
    if (sA !== sH) { seasonRecordPostseasonGame(away, home, isWS); return { winner: sA > sH ? away : home, sA, sH }; }
  }
  return { winner: home, sA: 0, sH: 0 };
}
// 自動で次の1戦を消化 → ブラケット反映。{played} / {done:true} / {error:true}
function seasonPostseasonStepAuto() {
  const ng = seasonPostseasonNextGame();
  if (!ng) return { done: true };
  const r = playPostseasonGame(ng.away, ng.home, ng.seriesId === 'WS');
  if (!r) return { error: true };
  seasonPostseasonApplyResult(ng, r.winner, r.sA, r.sH);
  return { played: true };
}
// 残りを全部自動進行。成功でtrue。
function seasonPostseasonPlayAll() {
  if (!SEASON.postseason) seasonPostseasonInit();
  seasonPostseasonEnsure();
  for (let guard = 0; guard < 200; guard++) {
    const st = seasonPostseasonStepAuto();
    if (st.done) return true;
    if (st.error) return false;
  }
  return false;
}
// 手動: 次の1戦をセットアップに反映して試合開始 (結果は seasonAfterManualResult で記録)
function seasonStartPostseasonManual(sel) {
  const ng = seasonPostseasonNextGame();
  if (!ng) return false;
  seasonEnsureBuild(ng.away); seasonEnsureBuild(ng.home);
  const awayStarter = (sel && sel.awayStarter) || seasonAutoStarter(ng.away, seasonRotIdx(ng.away));
  const homeStarter = (sel && sel.homeStarter) || seasonAutoStarter(ng.home, seasonRotIdx(ng.home));
  seasonAdvanceRot(ng.away); seasonAdvanceRot(ng.home);
  const awayOrder = (sel && sel.awayOrder) || 0, homeOrder = (sel && sel.homeOrder) || 0;
  const e1 = buildSeasonSideSetup('away', ng.away, awayOrder, awayStarter, false, sel && sel.awayPicks, sel && sel.awayLineOrder, sel && sel.awayRest);
  const e2 = e1 ? null : buildSeasonSideSetup('home', ng.home, homeOrder, homeStarter, false, sel && sel.homePicks, sel && sel.homeLineOrder, sel && sel.homeRest);
  if (e1 || e2) { alert('試合を開始できません:\n' + (e1 || e2) + '\n(そのチームのカードが不足しています)'); return false; }
  G.innings = 9; G.seasonMode = true;
  G.seasonCtx = { away: ng.away, home: ng.home, awayOrder, homeOrder, auto: false, postseason: true, psNext: ng };
  beginGame();
  return true;
}
function seasonPsSeriesTitle(id) {
  if (id === 'WS') return 'ワールドシリーズ';
  const m = { WC1: 'ワイルドカード①', WC2: 'ワイルドカード②', DS1: '地区シリーズ①', DS2: '地区シリーズ②', LCS: 'リーグ優勝決定戦' };
  return SEASON_LEAGUE_JP[id.slice(0, 2)] + ' ' + (m[id.slice(3)] || id);
}
// ===== ポストシーズン 画面 =====
function seasonPostseasonHtml() {
  if (!seasonDone()) return `<div class="season-end"><h3>ポストシーズンはレギュラーシーズン終了後に開催されます。</h3><p>残り <b>${SEASON.schedule.length - SEASON.cursor}</b> 試合。</p></div>`;
  if (!SEASON.postseason) seasonPostseasonInit();
  seasonPostseasonEnsure();
  const tabs = [['bracket', 'トーナメント表'], ['stats', 'ポストシーズン成績'], ['ws', 'ワールドシリーズ記録']];
  const tabH = tabs.map(([k, l]) => `<button class="season-tab${SEASON_PS_TAB === k ? ' on' : ''}" data-pstab="${k}">${l}</button>`).join('');
  let inner = SEASON_PS_TAB === 'stats' ? seasonPsStatsHtml() : (SEASON_PS_TAB === 'ws' ? seasonPsWSHtml() : seasonPsBracketHtml());
  return `<div class="ps-wrap"><div class="season-tabs">${tabH}</div><div class="season-tablewrap">${inner}</div></div>`;
}
// 本物のトーナメント表 (SVG)。AL左/NL右、中央にトロフィー＋ワールドシリーズ。
function seasonPsBracketSvg(ps) {
  const S = ps.series, seeds = ps.seeds, champ = ps.champion;
  const Win = id => (S[id] && S[id].done) ? S[id].winner : null;
  const specs = {}; seasonPostseasonBracket().forEach(sp => specs[sp.id] = sp);
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const nm = t => t ? esc(seasonTeamName(t)) : '—';
  const seedNo = (lg, t) => { const i = t ? seeds[lg].indexOf(t) : -1; return i >= 0 ? (i + 1) : ''; };
  const GOLD = '#f5c518', DIM = '#54627a', BW = 126, BH = 46;
  const conn = (x1, y1, x2, y2, on) => { const mx = (x1 + x2) / 2; return `<path d="M ${x1} ${y1} H ${mx} V ${y2} H ${x2}" fill="none" stroke="${on ? GOLD : DIM}" stroke-width="${on ? 3.5 : 2}"/>`; };
  // マッチ箱。id からシリーズを引き、hi=上段/lo=下段。
  const mbox = (lg, id, x, yc, title) => {
    const sp = specs[id], s = S[id], winner = s && s.done ? s.winner : null;
    const hi = sp ? sp.hi : null, lo = sp ? sp.lo : null;
    const games = (s && s.games || []).map((g, i) => `第${i + 1}戦 ${seasonTeamName(g.away)} ${g.sA}-${g.sH} ${seasonTeamName(g.home)}`).join('\n');
    const y = yc - BH / 2;
    const rowSvg = (team, seed, wins, isWin, ry) => {
      const name = team ? nm(team) : '—';
      const long = team && seasonTeamName(team).length >= 7;   // 長名は箱幅へ自動圧縮
      const nameX = x + (seed ? 19 : 8);
      const cap = (x + BW - 16) - nameX;
      return `${isWin ? `<rect x="${x + 1.5}" y="${ry}" width="${BW - 3}" height="${BH / 2 - 1.5}" fill="#fff6d8"/>` : ''}`
        + (seed ? `<text x="${x + 6}" y="${ry + 15}" font-size="9" font-weight="700" fill="#8a97ad">${seed}</text>` : '')
        + `<text x="${nameX}" y="${ry + 15}" font-size="12" font-weight="${isWin ? '800' : '600'}" fill="#1a2433"${long ? ` textLength="${cap}" lengthAdjust="spacingAndGlyphs"` : ''}>${name}</text>`
        + `<text x="${x + BW - 8}" y="${ry + 15}" font-size="13" font-weight="800" text-anchor="end" fill="${isWin ? '#b07a00' : '#3a4760'}">${wins != null ? wins : ''}</text>`;
    };
    return `<g>${games ? `<title>${esc(games)}</title>` : ''}
      <rect x="${x}" y="${y}" width="${BW}" height="${BH}" rx="6" fill="#ffffff" stroke="${winner ? '#b9912f' : '#33415c'}" stroke-width="1.5"/>
      <line x1="${x}" y1="${yc}" x2="${x + BW}" y2="${yc}" stroke="#c8d2de"/>
      ${title ? `<text x="${x + BW / 2}" y="${y - 4}" font-size="10.5" text-anchor="middle" fill="#b9c4d6">${title}</text>` : ''}
      ${rowSvg(hi, seedNo(lg, hi), s ? s.winsHi : null, winner && winner === hi, y + 1)}
      ${rowSvg(lo, seedNo(lg, lo), s ? s.winsLo : null, winner && winner === lo, yc + 1)}</g>`;
  };
  // 片側 (left=AL寄せ / right=NL寄せ) のボックス座標
  const side = (lg, L) => {
    const X = p => L ? p : (1000 - p - BW);   // 右側はミラー
    const wc2 = X(14), ds1 = X(160), wc1 = X(14), ds2 = X(160), lcs = X(306);
    const rE = x => x + BW;   // 箱の(描画上の)右端
    // コネクタ: 左側は箱の右端→次の左端 / 右側はその逆
    const c = (ax, ay, bx, by, on) => L ? conn(rE(ax), ay, bx, by, on) : conn(ax, ay, rE(bx), by, on);
    const lcsToWs = L ? conn(rE(lcs), 245, 437, 250, Win(lg + '_LCS')) : conn(lcs, 245, 563, 260, Win(lg + '_LCS'));
    return `
      ${c(wc2, 165, ds1, 105, Win(lg + '_WC2'))}
      ${c(wc1, 325, ds2, 385, Win(lg + '_WC1'))}
      ${c(ds1, 105, lcs, 245, Win(lg + '_DS1'))}
      ${c(ds2, 385, lcs, 245, Win(lg + '_DS2'))}
      ${lcsToWs}
      ${mbox(lg, lg + '_WC2', wc2, 165, 'ワイルドカード')}
      ${mbox(lg, lg + '_WC1', wc1, 325, 'ワイルドカード')}
      ${mbox(lg, lg + '_DS1', ds1, 105, '地区シリーズ')}
      ${mbox(lg, lg + '_DS2', ds2, 385, '地区シリーズ')}
      ${mbox(lg, lg + '_LCS', lcs, 245, 'リーグ優勝')}`;
  };
  const lgLabel = (x, txt, color) => `<text x="${x}" y="22" font-size="16" font-weight="900" fill="${color}" text-anchor="middle">${txt}</text>`;
  return `<svg viewBox="0 0 1000 422" class="ps-svg" preserveAspectRatio="xMidYMid meet">
    ${lgLabel(140, 'アメリカンリーグ', '#ff7a7a')}${lgLabel(860, 'ナショナルリーグ', '#7aa6ff')}
    ${side('AL', true)}${side('NL', false)}
    <text x="500" y="148" font-size="50" text-anchor="middle">🏆</text>
    ${mbox('AL', 'WS', 437, 255, 'ワールドシリーズ')}
    <text x="500" y="312" font-size="11" text-anchor="middle" fill="#ffd479">WORLD SERIES</text>
  </svg>`;
}
function seasonPsBracketHtml() {
  const ps = SEASON.postseason, champ = ps.champion;
  const champHtml = champ ? `<div class="ps-champion">🏆 世界一 <b>${seasonTeamName(champ)}</b>${ps.mvp ? `<div class="ps-mvp">WS MVP: <b>${ps.mvp.name}</b>（${seasonTeamName(ps.mvp.team)}）<span>${ps.mvp.line}</span></div>` : ''}</div>` : '';
  const ng = seasonPostseasonNextGame();
  let actions;
  if (champ) actions = `<button class="btn btn-sub" data-sact="reset-postseason">↻ ポストシーズンをやり直す</button>`;
  else if (ng) actions = `<span class="ps-next">次の試合: ${seasonPsSeriesTitle(ng.seriesId)} 第${ng.gi + 1}戦 ${seasonTeamName(ng.away)} vs ${seasonTeamName(ng.home)}</span>
      <button class="btn btn-primary" data-sv="postseason-manual">▶ 手動でプレイ</button>
      <button class="btn btn-sub" data-sact="play-ps-one">⏩ この試合を自動</button>
      <button class="btn btn-sub" data-sact="play-postseason">⏩⏩ 残りを全部自動</button>`;
  else actions = `<button class="btn btn-sub" data-sact="reset-postseason">↻ ポストシーズンをやり直す</button>`;
  return `<div class="season-actions">${actions}</div>
    <p class="season-note">WC(3戦2勝)→地区(5戦3勝)→リーグ優勝(7戦4勝)→ワールドシリーズ(7戦4勝)。各シリーズにマウスを乗せると試合別スコアが出ます。</p>
    ${champHtml}
    <div class="ps-bracket-wrap">${seasonPsBracketSvg(ps)}</div>`;
}
// ポストシーズン累積成績 (列ヘッダクリックで並び替え)
function seasonPsOps(s) {
  const ob = (s.AB + s.BB + s.HBP + s.SAC) > 0 ? (s.H + s.BB + s.HBP) / (s.AB + s.BB + s.HBP + s.SAC) : 0;
  const sl = s.AB > 0 ? (s.H + s.dbl + 2 * s.tpl + 3 * s.HR) / s.AB : 0;
  return ob + sl;
}
function seasonPsStatsHtml() {
  const ps = SEASON.postseason;
  let bats = Object.values(ps.bat || {}).filter(s => (s.PA || 0) > 0);
  let pits = Object.values(ps.pit || {}).filter(s => (s.outs || 0) > 0);
  if (!bats.length && !pits.length) return `<p class="season-note">まだポストシーズンの成績がありません。試合を進めると記録されます。</p>`;
  // 参加チーム (ポストシーズン出場球団) を SEASON_TEAMS 順に並べてタブ表示
  const teamSet = new Set();
  bats.forEach(s => teamSet.add(normalizeTeam(s.team)));
  pits.forEach(s => teamSet.add(normalizeTeam(s.team)));
  const teams = SEASON_TEAMS.filter(t => teamSet.has(t));
  const sel = (SEASON_PS_TEAM && teamSet.has(SEASON_PS_TEAM)) ? SEASON_PS_TEAM : '';
  const teamTabs = `<div class="season-subtabs"><button class="season-subtab${sel === '' ? ' on' : ''}" data-psteam="">全体</button>${teams.map(t => `<button class="season-subtab${sel === t ? ' on' : ''}" data-psteam="${t}">${seasonTeamName(t)}</button>`).join('')}</div>`;
  if (sel) { bats = bats.filter(s => normalizeTeam(s.team) === sel); pits = pits.filter(s => normalizeTeam(s.team) === sel); }
  const batCols = [
    { key: 'name', label: '選手', lname: true, nosort: true, get: r => r.name, fmt: (v, r) => seasonNameCell(r, 'batter') },
    { key: 'team', label: 'チーム', nosort: true, get: r => seasonTeamName(r.team) },
    { key: 'avg', label: '打率', get: r => avgOf(r.H, r.AB), fmt: v => fmt3(v) },
    { key: 'G', label: '試合', get: r => r.G }, { key: 'AB', label: '打数', get: r => r.AB },
    { key: 'H', label: '安打', get: r => r.H }, { key: 'HR', label: '本', get: r => r.HR },
    { key: 'RBI', label: '打点', get: r => r.RBI }, { key: 'R', label: '得点', get: r => r.R },
    { key: 'BB', label: '四球', get: r => r.BB }, { key: 'SO', label: '三振', get: r => r.SO },
    { key: 'SB', label: '盗塁', get: r => r.SB }, { key: 'ops', label: 'OPS', get: r => seasonPsOps(r), fmt: v => fmt3(v) },
  ];
  const pitCols = [
    { key: 'name', label: '選手', lname: true, nosort: true, get: r => r.name, fmt: (v, r) => seasonNameCell(r, 'pitcher') },
    { key: 'team', label: 'チーム', nosort: true, get: r => seasonTeamName(r.team) },
    { key: 'era', label: '防御率', get: r => eraOf(r.ER, r.outs), fmt: v => fmt2(v) },
    { key: 'G', label: '登板', get: r => r.G }, { key: 'outs', label: '投球回', get: r => r.outs, fmt: v => ipText(v) },
    { key: 'W', label: '勝', get: r => r.W }, { key: 'L', label: '敗', get: r => r.L },
    { key: 'S', label: 'S', get: r => r.S }, { key: 'HLD', label: 'H', get: r => r.HLD },
    { key: 'K', label: '奪三振', get: r => r.K }, { key: 'BB', label: '四球', get: r => r.BB }, { key: 'H', label: '被安打', get: r => r.H },
  ];
  const tbl = (which, cols, rows, title) => {
    const srt = SEASON_PS_SORT[which], col = cols.find(c => c.key === srt.key) || cols[2];
    rows = rows.slice().sort((a, b) => { const va = col.get(a), vb = col.get(b); return (va < vb ? -1 : va > vb ? 1 : 0) * srt.dir; });
    const head = cols.map(c => c.nosort
      ? `<th class="${c.lname ? 'lname' : ''}">${c.label}</th>`
      : `<th class="sortable${srt.key === c.key ? ' sorted' : ''}" data-pssort="${which}:${c.key}">${c.label}${srt.key === c.key ? (srt.dir < 0 ? ' ▼' : ' ▲') : ''}</th>`).join('');
    const body = rows.slice(0, 40).map(r => '<tr>' + cols.map(c => `<td class="${c.lname ? 'lname' : ''}${srt.key === c.key ? ' sortcol' : ''}">${c.fmt ? c.fmt(c.get(r), r) : c.get(r)}</td>`).join('') + '</tr>').join('');
    return `<h4>${title}</h4><div class="season-tablewrap"><table class="season-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  };
  const title = sel ? '（' + seasonTeamName(sel) + '）' : '';
  return teamTabs + tbl('bat', batCols, bats, '打撃成績' + title) + tbl('pit', pitCols, pits, '投手成績' + title);
}
// ワールドシリーズ 試合記録 (各戦のラインスコア + ボックス)
function seasonPsWSHtml() {
  const ps = SEASON.postseason;
  if (!ps.wsGames || !ps.wsGames.length) return `<p class="season-note">ワールドシリーズはまだ行われていません。</p>`;
  const lineTable = g => {
    const head = Array.from({ length: g.innings }, (_, i) => `<th>${i + 1}</th>`).join('');
    const row = (name, arr, tot, h) => `<tr><th class="rsb-team">${seasonTeamName(name)}</th>${arr.map(v => `<td>${v}</td>`).join('')}<td class="rsb-total">${tot}</td><td>${h}</td></tr>`;
    return `<table class="ws-line"><thead><tr><th class="rsb-team">チーム</th>${head}<th>計</th><th>H</th></tr></thead><tbody>
      ${row(g.away, g.line.away, g.sA, g.hitsA)}${row(g.home, g.line.home, g.sH, g.hitsH)}</tbody></table>`;
  };
  const hrType = r => ['', 'ソロ', '2ラン', '3ラン', '満塁'][Math.min(4, r || 1)];
  const hrList = g => g.hr && g.hr.length ? `<div class="ws-hr">🏟️ 本塁打: ${g.hr.map(h => `${h.name}${h.num ? h.num + '号' : ''}(${h.inning}回${h.top ? '表' : '裏'}${hrType(h.runs)})`).join('・')}</div>` : '';
  const decLine = g => `<div class="ws-dec">${g.win ? `勝: ${g.win}` : ''}${g.lose ? `　負: ${g.lose}` : ''}${g.save ? `　S: ${g.save}` : ''}</div>`;
  // ポストシーズン通算の打率/防御率 (key で参照)
  // ポストシーズン通算 (赤字「ポ◯」列) を box の key で参照。打率/防御率＋HR/打点・勝敗S/奪三振。
  const psB = key => ps.bat[key] || {};
  const psP = key => ps.pit[key] || {};
  const psAvg = s => s.AB ? fmt3(avgOf(s.H, s.AB)) : '.---';
  const psEra = s => s.outs ? fmt2(eraOf(s.ER, s.outs)) : '-.--';
  const batT = arr => `<table class="ws-box"><thead><tr><th class="lname">打者</th><th>ポ打率</th><th class="ps-n">ポ本</th><th class="ps-n">ポ点</th><th>打数</th><th>安</th><th>本</th><th>点</th><th>得</th><th>四</th><th>三</th></tr></thead><tbody>${arr.map(b => { const c = psB(b.key); return `<tr><td class="lname">${seasonNameCell({ name: b.name, year: c.year, team: c.team }, 'batter')}</td><td class="ps-cum">${psAvg(c)}</td><td class="ps-cum ps-n">${c.HR || 0}</td><td class="ps-cum ps-n">${c.RBI || 0}</td><td>${b.AB}</td><td>${b.H}</td><td>${b.HR}</td><td>${b.RBI}</td><td>${b.R}</td><td>${b.BB}</td><td>${b.SO}</td></tr>`; }).join('')}</tbody></table>`;
  const pitT = arr => `<table class="ws-box"><thead><tr><th class="lname">投手</th><th>ポ防</th><th class="ps-n">ポ勝</th><th class="ps-n">ポ敗</th><th class="ps-n">ポH</th><th class="ps-n">ポS</th><th class="ps-n">ポ奪三</th><th>回</th><th>安</th><th>自</th><th>三</th><th>四</th></tr></thead><tbody>${arr.map(p => { const c = psP(p.key); return `<tr><td class="lname">${seasonNameCell({ name: p.name, year: c.year, team: c.team }, 'pitcher')}${p.dec ? '(' + p.dec + ')' : ''}</td><td class="ps-cum">${psEra(c)}</td><td class="ps-cum ps-n">${c.W || 0}</td><td class="ps-cum ps-n">${c.L || 0}</td><td class="ps-cum ps-n">${c.HLD || 0}</td><td class="ps-cum ps-n">${c.S || 0}</td><td class="ps-cum ps-n">${c.K || 0}</td><td>${ipText(p.outs)}</td><td>${p.H}</td><td>${p.ER}</td><td>${p.K}</td><td>${p.BB}</td></tr>`; }).join('')}</tbody></table>`;
  return ps.wsGames.map((g, i) => `<details class="ws-game"${i === ps.wsGames.length - 1 ? ' open' : ''}>
    <summary>第${i + 1}戦　${seasonTeamName(g.away)} ${g.sA} - ${g.sH} ${seasonTeamName(g.home)}</summary>
    <div class="ws-body">${lineTable(g)}${decLine(g)}${hrList(g)}
      <div class="ws-boxes"><div><h6>${seasonTeamName(g.away)} 打撃</h6>${batT(g.box.away.bat)}<h6>投手</h6>${pitT(g.box.away.pit)}</div>
      <div><h6>${seasonTeamName(g.home)} 打撃</h6>${batT(g.box.home.bat)}<h6>投手</h6>${pitT(g.box.home.pit)}</div></div>
    </div></details>`).join('');
}
// 手動: 次の1戦のセットアップ画面 (レギュラー手動UIの部品を流用)
function seasonPostseasonManualInit(ng) {
  const cursorKey = 'ps:' + ng.seriesId + ':' + ng.gi;
  const valid = SEASON_MANUAL_SEL && SEASON_MANUAL_SEL.cursor === cursorKey
    && SEASON_MANUAL_SEL.away.team === ng.away && SEASON_MANUAL_SEL.home.team === ng.home;
  if (valid) return;
  SEASON_OPEN_PICK = null; SEASON_OPEN_STARTER = null;
  const mkSide = (teamCode) => {
    const rest = new Set();   // ポストシーズンは休養なし
    const picks = seasonDefaultPicks(teamCode, 0, rest);
    return { team: teamCode, order: 0, starterIdx: seasonAutoStarterIdx(teamCode), rest, picks, lineOrder: seasonManualLineOrder(teamCode, 0, picks) };
  };
  SEASON_MANUAL_SEL = { cursor: cursorKey, away: mkSide(ng.away), home: mkSide(ng.home) };
}
function seasonPostseasonManualHtml() {
  const ng = seasonPostseasonNextGame();
  if (!ng) { SEASON_VIEW = 'postseason'; return seasonPostseasonHtml(); }
  seasonPostseasonManualInit(ng);
  return `<div class="season-manual">
    <div class="season-gamehdr">${seasonPsSeriesTitle(ng.seriesId)} 第${ng.gi + 1}戦　${seasonTeamName(ng.away)} (AWAY) <span class="vs">vs</span> ${seasonTeamName(ng.home)} (HOME)</div>
    <div class="season-setrow">
      ${seasonSideSetupHtml('away', 'AWAY', ng.away, [0])}
      ${seasonSideSetupHtml('home', 'HOME', ng.home, [0])}
    </div>
    <div class="season-actions">
      <button class="btn btn-primary" data-sact="play-ps-manual">▶ 試合開始</button>
      <button class="btn btn-sub" data-sv="postseason">← トーナメント表へ戻る</button>
    </div>
  </div>`;
}

// ============== 起動 ==============
// 追加カードを IndexedDB から読み込んでから初期化する (同期参照キャッシュを用意)
document.addEventListener('DOMContentLoaded', async () => {
  if (window.CardStore && window.CardStore.init) {
    try { await window.CardStore.init(); } catch (e) { console.error('CardStore init', e); }
  }
  init();
});

})();
