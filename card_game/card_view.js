// 共通: 選手詳細カードを「本物のカード」風に表示するモーダル
// 使い方: window.CARD_VIEW.show(player)  /  window.CARD_VIEW.hide()
(function(){
'use strict';

let maskEl = null;
let popEl  = null;   // 動画ポップアップ (開いていなければ null)

function ensureMask() {
  if (maskEl) return maskEl;
  maskEl = document.createElement('div');
  maskEl.className = 'modal-mask hidden';
  maskEl.id = 'card-modal';
  document.body.appendChild(maskEl);
  // 背景クリックで閉じる
  maskEl.addEventListener('click', e => {
    if (e.target === maskEl) hide();
  });
  // Esc で閉じる
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape' || maskEl.classList.contains('hidden')) return;
    if (popEl) closeVideoPopup(); else hide();   // 動画ポップアップが開いていればそれだけ閉じる
  });
  return maskEl;
}

function classByVal(v) {
  if (v == null) return '';
  if (v >= 90) return 's-rank';
  if (v >= 80) return 'a-rank';
  if (v < 0)   return 'neg';
  return '';
}
function pitchClass(v) {
  if (v == null || v === 0) return 'neu';
  return v > 0 ? 'pos' : 'neg';
}
function pitchBoxClass(v) {
  if (v == null || v === 0) return 'neu-b';
  return v > 0 ? 'pos-b' : 'neg-b';
}
function formatPitchVal(v) {
  if (v == null) return '±0';
  if (v === 0) return '±0';
  return (v > 0 ? '+' : '') + v;
}

function renderBatterRight(p) {
  const r = p.record || {};
  const s = p.stats || {};
  const m = p.statsMini || {};
  const pp = p.pitchPoints || {};
  // record bar (打者は7項目)
  const recItems = [
    {lbl: '打率',  val: r['打率'],   cls: 'hl'},
    {lbl: '本塁打', val: r['本塁打'], cls: 'pp'},
    {lbl: '打点',  val: r['打点'],   cls: ''},
    {lbl: '盗塁',  val: r['盗塁'],   cls: ''},
    {lbl: '出塁率', val: r['出塁率'], cls: 'hl'},
    {lbl: 'OPS',  val: r['OPS'],   cls: 'hl'},
    {lbl: 'WAR',  val: r['WAR'],   cls: 'hl'},
  ];

  // 4つの主能力 + 選球眼/三振耐性 + ミニ3 (HR能/対左/盗塁能)
  const main4 = ['ミート','パワー','スピード','チャンス'];
  const sub2  = ['選球眼','三振耐性'];
  const miniKeys = ['盗塁能','対左投手','HR能'];

  const pitchOrder = ['FB','2C','CT','SL','CB','CH','SF'];

  return `
    <div class="sec-title">■ 年間成績</div>
    <div class="record-bar">
      ${recItems.map(i => `
        <div class="rec-item">
          <span class="lbl">${i.lbl}</span>
          <span class="val ${i.cls}">${i.val ?? '-'}</span>
        </div>
      `).join('')}
    </div>

    <div class="sec-title">■ ゲームステータス</div>
    <div class="stats-grid">
      ${main4.map(k => `
        <div class="stat-box">
          <span class="name">${k}</span>
          <span class="val ${classByVal(s[k])}">${s[k] ?? '-'}</span>
        </div>
      `).join('')}
    </div>
    <div class="stats-grid" style="grid-template-columns: 1fr 1fr;">
      ${sub2.map(k => `
        <div class="stat-box">
          <span class="name">${k}</span>
          <span class="val ${classByVal(s[k])}">${s[k] ?? '-'}</span>
        </div>
      `).join('')}
    </div>
    <div class="stats-mini-grid">
      ${miniKeys.map(k => `
        <div class="stat-box-mini">
          <span class="name">${k}</span>
          <span class="val ${(m[k] != null && m[k] < 0) ? 'neg' : ''}">${m[k] ?? 0}</span>
        </div>
      `).join('')}
    </div>

    <div class="sec-title">■ 対球種ポイント</div>
    <div class="pitch-grid">
      ${pitchOrder.map(k => `
        <div class="pitch-box ${pitchBoxClass(pp[k])}">
          <span class="lbl">${k}</span>
          <span class="val ${pitchClass(pp[k])}">${formatPitchVal(pp[k])}</span>
        </div>
      `).join('')}
    </div>

    ${renderDrs(p)}
    ${renderCatcher(p)}
  `;
}

function renderPitcherRight(p) {
  const r = p.record || {};
  const s = p.stats || {};
  // 投手のrecord項目
  const recItems = [
    {lbl: '防御率', val: r['防御率'], cls: 'hl'},
    {lbl: '勝',    val: r['勝利'],   cls: ''},
    {lbl: '敗',    val: r['敗北'],   cls: ''},
    {lbl: 'セーブ', val: r['セーブ'], cls: ''},
    {lbl: 'イニング', val: r['イニング'], cls: ''},
    {lbl: '奪三振', val: r['奪三振'], cls: 'pp'},
    {lbl: 'WAR',  val: r['WAR'],   cls: 'hl'},
  ];

  const pitStatKeys = ['スタミナ','制球','緩急','精神','奪三振','重さ','対左','阻止'];

  return `
    <div class="sec-title">■ 年間成績</div>
    <div class="record-bar">
      ${recItems.map(i => `
        <div class="rec-item">
          <span class="lbl">${i.lbl}</span>
          <span class="val ${i.cls}">${i.val ?? '-'}</span>
        </div>
      `).join('')}
    </div>

    <div class="sec-title">■ ゲームステータス</div>
    <div class="stats-grid">
      ${pitStatKeys.slice(0,4).map(k => `
        <div class="stat-box">
          <span class="name">${k}</span>
          <span class="val ${classByVal(s[k])}">${s[k] ?? '-'}</span>
        </div>
      `).join('')}
    </div>
    <div class="stats-grid">
      ${pitStatKeys.slice(4).map(k => `
        <div class="stat-box">
          <span class="name">${k}</span>
          <span class="val ${classByVal(s[k])}">${s[k] ?? '-'}</span>
        </div>
      `).join('')}
    </div>

    <div class="sec-title">■ 球種</div>
    <table class="pitch-table">
      <thead>
        <tr><th>球種</th><th>球速</th><th>球威</th><th>割合</th></tr>
      </thead>
      <tbody>
        ${(p.pitches || []).map(pi => `
          <tr>
            <td class="pn">${pi.name}</td>
            <td>${pi.speed ?? '-'} km/h</td>
            <td>${pi.power ?? '-'}</td>
            <td>${pi.ratio ?? '-'} %</td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${renderDrs(p)}
  `;
}

function renderDrs(p) {
  if (!p.drs || !p.drs.length) return '';
  return `
    <div class="sec-title">■ 守備 DRS</div>
    <div class="drs-bar">
      ${p.drs.map(d => `
        <div class="drs-item">
          <span class="drs-pos">${d.pos}</span>
          <span class="drs-num ${d.value > 0 ? 'pos' : (d.value < 0 ? 'neg' : 'neu')}">${d.value > 0 ? '+' : ''}${d.value ?? 0}</span>
          ${d.innings != null ? `<span class="drs-inn">${d.innings} inn</span>` : ''}
        </div>
      `).join('')}
    </div>
  `;
}

function renderCatcher(p) {
  if (!p.catcher || Object.keys(p.catcher).length === 0) return '';
  const items = Object.entries(p.catcher).map(([k, v]) => {
    const cls = k.includes('リード') ? 'lead' : 'cs';
    return `
      <div class="ca-item">
        <span class="ca-label">${k}</span>
        <span class="ca-val ${cls}">${v ?? '-'}</span>
      </div>
    `;
  }).join('');
  return `
    <div class="sec-title">■ 捕手能力</div>
    <div class="catcher-bar">${items}</div>
  `;
}

// ============== 選手固有の動画 (写真左下の「動画」ボタン → 半透明ポップアップで再生) ==============
// 登録データは player_videos.js (window.PLAYER_VIDEOS)。試合中に流れる動画と同じものを確認できる。
const VIDEO_BTN_TITLE_ON  = 'この選手の試合中動画を確認する';
const VIDEO_BTN_TITLE_OFF = 'この選手の専用動画は登録されていません';

function listVideos(player) {
  try { return (window.PLAYER_VIDEOS && window.PLAYER_VIDEOS.listFor(player)) || []; }
  catch (e) { return []; }
}
function videoButtonHtml(vids) {
  const on = vids.length > 0;
  return `<button type="button" class="cv-video-btn"${on ? '' : ' disabled'} title="${on ? VIDEO_BTN_TITLE_ON : VIDEO_BTN_TITLE_OFF}">🎬 動画${on ? `<span class="cnt">${vids.length}</span>` : ''}</button>`;
}
// ボタンは写真の最下段(左下・下端から6px)に置く。選手名は写真下部に中央寄せで入っている
//   (日本語名36px + 任意の英語名12px) ため、名前の「文字そのもの」と重なる場合だけ表記を段階的に縮める:
//   「🎬 動画 N」→「🎬 N」→「🎬」。それでも重なる(英語名が無く日本語名が横幅いっぱいの場合等)ときだけ
//   最終手段として名前ブロックのすぐ上へ逃がす。文字幅は Range で実測する (名前の要素自体は全幅のため)。
function placeAtBottom(panel, btn, vids) {
  try {
    btn.style.bottom = '6px';
    const doc = panel.ownerDocument;
    const texts = Array.from(panel.querySelectorAll('.player-name-ja, .player-name-en, .player-name, .name-en'));
    const textRects = () => texts.map(el => { const r = doc.createRange(); r.selectNodeContents(el); const b = r.getBoundingClientRect(); return b.width > 0 ? b : null; }).filter(Boolean);
    const overlaps = () => { const b = btn.getBoundingClientRect(); return textRects().some(t => b.left < t.right && b.right > t.left && b.top < t.bottom && b.bottom > t.top); };
    if (!overlaps()) return;
    const n = vids.length;
    const shorter = n ? ['🎬<span class="cnt">' + n + '</span>', '🎬'] : ['🎬'];
    for (const html of shorter) { btn.innerHTML = html; if (!overlaps()) return; }
    const nb = panel.querySelector('.player-name-block');
    if (!nb) return;
    const gap = panel.getBoundingClientRect().bottom - nb.getBoundingClientRect().top;
    if (gap > 0 && gap < 200) btn.style.bottom = Math.round(gap + 6) + 'px';
  } catch (e) { /* 位置調整は見た目だけの処理 */ }
}
// 原本カード(iframe)の写真左下へボタンを差し込む。iframe内のCSSは親と別なので、ボタン用スタイルも一緒に入れる。
//   写真(.left-panel)にはホログラム用の ::after (z-index:6, クリック透過) が重なるため、ボタンは z-index:7 で上に置く。
function injectVideoButton(doc, player, vids) {
  if (!doc) return;
  const panel = doc.querySelector('#card-root .left-panel') || doc.querySelector('.left-panel');
  if (!panel || panel.querySelector('.cv-video-btn')) return;
  const st = doc.createElement('style');
  st.textContent = `
    .cv-video-btn { position:absolute; left:8px; bottom:6px; z-index:7;
      display:inline-flex; align-items:center; gap:4px;
      background:linear-gradient(135deg,#ff4444,#cc0000); color:#fff; border:none; border-radius:6px;
      padding:4px 8px; font-size:12px; font-weight:900; letter-spacing:1px; font-family:inherit;
      cursor:pointer; box-shadow:0 2px 6px rgba(0,0,0,.6); }
    .cv-video-btn:hover:not(:disabled) { transform:translateY(-1px); box-shadow:0 4px 10px rgba(0,0,0,.6); }
    .cv-video-btn:disabled { background:linear-gradient(135deg,#777,#444); color:#ccc; cursor:not-allowed; opacity:.75; }
    .cv-video-btn .cnt { background:rgba(0,0,0,.35); border-radius:10px; padding:0 6px; font-size:11px; }`;
  doc.head.appendChild(st);
  const wrap = doc.createElement('div');
  wrap.innerHTML = videoButtonHtml(vids);
  const btn = wrap.firstElementChild;
  panel.appendChild(btn);
  placeAtBottom(panel, btn, vids);
  // クリック処理は親(このファイル)側で受ける (同一オリジンのため iframe 内の要素へ直接ハンドラを付けられる)
  if (vids.length) btn.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); openVideoPopup(player, vids); });
}

// 半透明のポップアップ。カテゴリ(項目)ボタンを押すとその動画を再生する。
//   同じ項目に複数本ある場合は、押すたびに次の本へ切り替える (全部を順に確認できる)。
function openVideoPopup(player, vids) {
  closeVideoPopup();
  const mask = ensureMask();
  const name = player.fullNameTop || '';
  const both = vids.some(v => v.kind === 'batter') && vids.some(v => v.kind === 'pitcher');
  const total = vids.reduce((s, v) => s + v.files.length, 0);
  // 項目ボタン + (複数本の項目には) 1,2,3… の番号ボタン。
  //   項目名 = その項目を1本目から順番に再生して項目の最後で停止 / 番号 = その本から項目の最後まで再生して停止
  const catHtml = (v, i) => `
    <div class="cv-vcat-wrap">
      <button type="button" class="cv-vcat" data-i="${i}" title="${v.files.length > 1 ? 'この項目を1本目から順番に再生し、項目の最後で停止します' : v.files[0] + '.mp4 を再生して停止します'}">
        <span class="lbl">${v.label}</span>${v.files.length > 1 ? `<span class="cnt">全${v.files.length}本</span>` : ''}
      </button>
      ${v.files.length > 1 ? `<div class="cv-vnums">${v.files.map((f, k) => `<button type="button" class="cv-vnum" data-i="${i}" data-k="${k}" title="${f}.mp4 から項目の最後まで再生して停止します">${k + 1}</button>`).join('')}</div>` : ''}
    </div>`;
  const catsOf = (kind) => vids.map((v, i) => v.kind === kind ? catHtml(v, i) : '').join('');
  const groups = both
    ? `<div class="cv-vgroup">■ 打者として</div>${catsOf('batter')}<div class="cv-vgroup">■ 投手として</div>${catsOf('pitcher')}`
    : catsOf('batter') + catsOf('pitcher');
  popEl = document.createElement('div');
  popEl.className = 'cv-vpop';
  popEl.innerHTML = `
    <div class="cv-vpop-panel" role="dialog" aria-label="${name} の動画">
      <div class="cv-vpop-head">
        <span class="ttl">🎬 ${name} の試合中動画</span>
        <span class="sub">${vids.length}項目 / 計${total}本　　開くと上からエンドレス連続再生　項目名・番号＝その項目内を順番に再生して停止</span>
        <button type="button" class="cv-vpop-all" title="最初の項目の1本目から、全項目を上から順にエンドレスで連続再生します">▶ 上から連続再生</button>
        <button type="button" class="cv-vpop-close" aria-label="閉じる">✕ 閉じる</button>
      </div>
      <div class="cv-vpop-body">
        <div class="cv-vpop-cats">${groups}</div>
        <div class="cv-vpop-stage">
          <video playsinline controls preload="none"></video>
          <div class="cv-vpop-now"></div>
        </div>
      </div>
    </div>`;
  mask.appendChild(popEl);
  const video = popEl.querySelector('video');
  const now   = popEl.querySelector('.cv-vpop-now');
  const state = { i: -1, k: -1, mode: 'all', errStreak: 0 };
  const mark = () => {
    popEl.querySelectorAll('.cv-vcat').forEach(b => b.classList.toggle('active', +b.dataset.i === state.i));
    popEl.querySelectorAll('.cv-vnum').forEach(b => b.classList.toggle('active', +b.dataset.i === state.i && +b.dataset.k === state.k));
  };
  // 通し番号 (全体で何本目か): 連続再生の進み具合を表示する
  const flatNo = (i, k) => vids.slice(0, i).reduce((s, v) => s + v.files.length, 0) + k + 1;
  // 再生が終わった(または再生できなかった)ときに次へ進む先を返す。
  //   mode 'all' = 全項目を上から連続、最後まで行ったら最初に戻る(エンドレス)
  //   mode 'seq' = その項目内だけ順番に進み、項目の最後で停止 (null)
  const nextOf = (i, k, mode) => {
    if (k + 1 < vids[i].files.length) return [i, k + 1];
    if (mode === 'all') return (i + 1 < vids.length) ? [i + 1, 0] : [0, 0];
    return null;
  };
  const play = (i, k, mode) => {
    const v = vids[i];
    if (!v || k < 0 || k >= v.files.length) return;
    state.i = i; state.k = k; state.mode = mode;
    mark();
    const file = v.files[k];
    const nth = v.files.length > 1 ? ` ${k + 1}/${v.files.length}本目` : '';
    const head = mode === 'all' ? `連続再生中 (${flatNo(i, k)}/${total}本目)` : (v.files.length > 1 ? '順番再生中' : '再生中');
    now.textContent = `${head}: ${v.label}${nth}　${file}.mp4`;
    const advance = () => {
      const nx = nextOf(i, k, mode);
      if (nx) { play(nx[0], nx[1], mode); return; }
      now.textContent = `再生終了: ${v.label}${nth}　${file}.mp4　(この項目で停止)`;
    };
    video.onplaying = () => { state.errStreak = 0; };
    video.onended = () => advance();
    // 再生できないファイルがあっても止まらないよう、案内を出して次へ進める。
    //   ただし全部が再生できない場合は(エンドレスで空回りしないよう)止める。
    video.onerror = () => {
      const msg = `再生できませんでした: ${file}.mp4 (MLB/douga/ にファイルがあるか確認してください)`;
      state.errStreak++;
      if (state.errStreak >= total) { now.textContent = msg + '　再生できる動画が無いため停止しました'; return; }
      if (nextOf(i, k, mode)) { now.textContent = msg + '　→ 次へ進みます'; setTimeout(() => { if (popEl && state.i === i && state.k === k) advance(); }, 800); }
      else now.textContent = msg;
    };
    video.src = window.PLAYER_VIDEOS.fileUrl(file);
    video.load();
    const pr = video.play();
    if (pr && pr.catch) pr.catch(() => { /* 自動再生がブロックされた場合は controls から再生できる */ });
  };
  popEl.querySelectorAll('.cv-vcat').forEach(b => b.addEventListener('click', () => play(+b.dataset.i, 0, 'seq')));
  popEl.querySelectorAll('.cv-vnum').forEach(b => b.addEventListener('click', () => play(+b.dataset.i, +b.dataset.k, 'seq')));
  popEl.querySelector('.cv-vpop-all').addEventListener('click', () => { state.errStreak = 0; play(0, 0, 'all'); });
  popEl.querySelector('.cv-vpop-close').addEventListener('click', closeVideoPopup);
  popEl.addEventListener('click', e => { if (e.target === popEl) closeVideoPopup(); });   // 半透明部分のクリックで閉じる
  // 開いた時点で、最初の項目の1本目から全項目を上から順にエンドレスで連続再生する
  play(0, 0, 'all');
}
function closeVideoPopup() {
  if (!popEl) return;
  try { const v = popEl.querySelector('video'); if (v) { v.pause(); v.removeAttribute('src'); v.load(); } } catch (e) {}
  popEl.remove();
  popEl = null;
}

function show(player) {
  if (!player) return;
  const mask = ensureMask();
  closeVideoPopup();
  const vids = listVideos(player);   // この選手の登録動画 (無ければ空 → ボタンは押せない)

  // 原本HTMLがあれば iframe でそのまま表示（フリップ・写真・サイズ全て原本通り）
  if (player.rawHtml) {
    mask.innerHTML = `
      <button class="modal-close" type="button" aria-label="閉じる">✕ 閉じる</button>
      <iframe class="card-iframe" sandbox="allow-same-origin allow-scripts"></iframe>
    `;
    // srcdoc はプロパティ経由で渡す（HTML属性エスケープが不要になり崩れない）
    const ifr = mask.querySelector('.card-iframe');
    // 読み込み完了後に写真左下へ「動画」ボタンを差し込む (srcdoc 設定より先にハンドラを付ける)
    ifr.addEventListener('load', () => { try { injectVideoButton(ifr.contentDocument, player, vids); } catch (e) { /* 動画ボタンは補助機能 */ } });
    ifr.srcdoc = player.rawHtml;
  } else {
    // 原本がない場合のフォールバック表示
    const right = (player.type === 'pitcher') ? renderPitcherRight(player) : renderBatterRight(player);
    mask.innerHTML = `
      <button class="modal-close" type="button" aria-label="閉じる">✕ 閉じる</button>
      <div class="full-card">
        <div class="top-banner">
          <span class="name-ja">${player.fullNameTop}</span>
          <span class="year">${player.seasonLabel || ''}</span>
          <span class="hand">${player.hand || ''}</span>
        </div>
        <div class="left-panel">
          <div class="left-overlay">
            <div class="team-badge">${player.team || '-'}</div>
            <div class="position">${player.position || '-'}</div>
          </div>
          <div class="player-name-block">
            <span class="player-name">${player.fullNameTop}</span>
            ${player.nameEn ? `<span class="name-en">${player.nameEn}</span>` : ''}
          </div>
          ${videoButtonHtml(vids)}
        </div>
        <div class="right-panel">
          ${right}
          ${player.retsuden ? `
            <div class="retsuden-box">
              <span class="lbl">■ 列伝</span>
              <div class="retsuden-text">${player.retsuden}</div>
            </div>
          ` : ''}
        </div>
      </div>
    `;
    const vb = mask.querySelector('.cv-video-btn');
    if (vb) placeAtBottom(mask.querySelector('.full-card .left-panel'), vb, vids);
    if (vb && vids.length) vb.addEventListener('click', () => openVideoPopup(player, vids));
  }
  mask.classList.remove('hidden');
  mask.querySelector('.modal-close').addEventListener('click', hide);
}

function hide() {
  closeVideoPopup();
  if (maskEl) maskEl.classList.add('hidden');
}

window.CARD_VIEW = { show, hide };
})();
