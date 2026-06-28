// 共通: 選手詳細カードを「本物のカード」風に表示するモーダル
// 使い方: window.CARD_VIEW.show(player)  /  window.CARD_VIEW.hide()
(function(){
'use strict';

let maskEl = null;

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
    if (e.key === 'Escape' && !maskEl.classList.contains('hidden')) hide();
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

function show(player) {
  if (!player) return;
  const mask = ensureMask();

  // 原本HTMLがあれば iframe でそのまま表示（フリップ・写真・サイズ全て原本通り）
  if (player.rawHtml) {
    mask.innerHTML = `
      <button class="modal-close" type="button" aria-label="閉じる">✕ 閉じる</button>
      <iframe class="card-iframe" sandbox="allow-same-origin allow-scripts"></iframe>
    `;
    // srcdoc はプロパティ経由で渡す（HTML属性エスケープが不要になり崩れない）
    const ifr = mask.querySelector('.card-iframe');
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
  }
  mask.classList.remove('hidden');
  mask.querySelector('.modal-close').addEventListener('click', hide);
}

function hide() {
  if (maskEl) maskEl.classList.add('hidden');
}

window.CARD_VIEW = { show, hide };
})();
