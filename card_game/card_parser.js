// ブラウザ用 カードHTML パーサー
// _parse_cards.js (Node版) と同じロジック。import.html から呼ばれる
(function(){
'use strict';

function pick(html, re) {
  const m = html.match(re);
  return m ? m[1].trim() : '';
}
function pickAll(html, re) {
  const out = [];
  let m;
  while ((m = re.exec(html)) !== null) out.push(m[1].trim());
  return out;
}
function decode(s) {
  return s
    .replace(/<br\s*\/?>/g, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ').replace(/&#x2606;/g, '☆').replace(/&#x2605;/g, '★')
    .trim();
}
function num(s) {
  if (s === '' || s == null) return null;
  const n = parseFloat(String(s).replace(/[^\-0-9.]/g, ''));
  return Number.isFinite(n) ? n : null;
}

function parseCard(html, filename) {
  const card = { sourceFile: filename };
  // 原本HTMLを保存（写真・フリップ含む全レイアウト再現用）
  card.rawHtml = html;

  card.nameJa = decode(pick(html, /<span class="player-name-ja">([\s\S]*?)<\/span>/));
  card.nameEn = decode(pick(html, /<span class="player-name-en">([\s\S]*?)<\/span>/));
  card.position = decode(pick(html, /<span class="position-badge">([\s\S]*?)<\/span>/));
  card.fullNameTop = decode(pick(html, /<span class="player-full-name"[^>]*>([\s\S]*?)<\/span>/));
  const season = decode(pick(html, /<span class="season-badge">([\s\S]*?)<\/span>/));
  card.seasonLabel = season;
  const ym = season.match(/(\d{4})/);
  card.year = ym ? parseInt(ym[1]) : null;

  card.team = decode(pick(html, /<div class="team-badge"[^>]*>([\s\S]*?)<\/div>/));
  // 写真: photo-frame 内の img src
  const photoMatch = html.match(/<div class="photo-frame">\s*<img[^>]+src="([^"]+)"/);
  if (photoMatch) card.photo = photoMatch[1];
  card.hand = decode(pick(html, /<span class="card-back-banner-hand">([\s\S]*?)<\/span>/));
  card.powerValue = num(pick(html, /<(?:div|span) class="power-value"[^>]*>([\s\S]*?)<\/(?:div|span)>/));

  const recBar = pick(html, /<div class="record-bar">([\s\S]*?)<\/div>\s*<div class="sec-title">/);
  const recLabels = pickAll(recBar, /<span class="rec-label">([\s\S]*?)<\/span>/g).map(decode);
  const recVals = pickAll(recBar, /<span class="rec-val[^"]*">([\s\S]*?)<\/span>/g).map(decode);
  card.record = {};
  for (let i = 0; i < recLabels.length; i++) card.record[recLabels[i]] = recVals[i];

  const statBlock = pick(html, /■ ゲームステータス([\s\S]*?)(?:<table class="pitch-table"|<div class="sec-title")/);
  if (statBlock) {
    const names = pickAll(statBlock, /<span class="stat-name(?!-)[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const vals  = pickAll(statBlock, /<span class="stat-val(?!-)[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const namesMini = pickAll(statBlock, /<span class="stat-name-mini[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const valsMini  = pickAll(statBlock, /<span class="stat-val-mini[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.stats = {};
    for (let i = 0; i < names.length; i++) {
      const key = names[i].replace(/\s+/g, '');
      card.stats[key] = num(vals[i]);
    }
    card.statsMini = {};
    for (let i = 0; i < namesMini.length; i++) {
      const key = namesMini[i].replace(/\s+/g, '');
      card.statsMini[key] = num(valsMini[i]);
    }
  }

  const pitchGrid = pick(html, /■ 対球種ポイント[\s\S]*?<div class="pitch-grid">([\s\S]*?)<\/div>\s*<\/div>/);
  if (pitchGrid) {
    const lbls = pickAll(pitchGrid, /<span class="pitch-lbl">([\s\S]*?)<\/span>/g).map(decode);
    const vs = pickAll(pitchGrid, /<span class="pitch-v[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.pitchPoints = {};
    for (let i = 0; i < lbls.length; i++) card.pitchPoints[lbls[i]] = num(vs[i]);
  }

  const pitchTable = pick(html, /<table class="pitch-table">([\s\S]*?)<\/table>/);
  if (pitchTable) {
    const rows = pickAll(pitchTable, /<tr>([\s\S]*?)<\/tr>/g);
    card.pitches = [];
    for (const row of rows) {
      const cells = pickAll(row, /<td[^>]*>([\s\S]*?)<\/td>/g).map(decode);
      if (cells.length >= 4) {
        card.pitches.push({
          name: cells[0],
          speed: num(cells[1]),
          power: num(cells[2]),
          ratio: num(cells[3]),
        });
      }
    }
  }

  // 守備DRS: 全 drs-pos / drs-num / drs-inn を順番にマッチ (マルチポジ対応)
  const allPoses = pickAll(html, /<span class="drs-pos">([\s\S]*?)<\/span>/g).map(decode);
  const allNums  = pickAll(html, /<span class="drs-num[^"]*">([\s\S]*?)<\/span>/g).map(decode);
  const allInns  = pickAll(html, /<span class="drs-inn">([\s\S]*?)<\/span>/g).map(decode);
  card.drs = [];
  for (let i = 0; i < allPoses.length; i++) {
    card.drs.push({ pos: allPoses[i], value: num(allNums[i]), innings: num(allInns[i]) });
  }
  const seen = new Set();
  card.drs = card.drs.filter(d => {
    if (seen.has(d.pos)) return false;
    seen.add(d.pos); return true;
  });

  const catcherBar = pick(html, /<div class="catcher-bar">([\s\S]*?)<\/div>(?=\s*<\/div>|\s*<div)/);
  if (catcherBar) {
    const labels = pickAll(catcherBar, /<span class="ca-label">([\s\S]*?)<\/span>/g).map(decode);
    const vals = pickAll(catcherBar, /<span class="ca-val[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.catcher = {};
    for (let i = 0; i < labels.length; i++) card.catcher[labels[i]] = num(vals[i]);
  }

  card.retsuden = decode(pick(html, /<div class="retsuden"[^>]*>([\s\S]*?)<\/div>/));

  // 種別判定: ファイル名に _P_ / _B_ のマーカーがあればそれを優先する。
  //  (二刀流の大谷翔平のように、打者カードにも投球データが載っていて
  //   球種の有無だけでは投手/打者を正しく分けられないケースに対応)
  //   例: 大谷翔平_P_2025_card.html → 投手 / 大谷翔平_B_2025_card.html → 打者
  // マーカーが無い通常カードは、従来通り球種の有無で自動判定する。
  let typeFromName = null;
  if (filename) {
    if (/_P_/i.test(filename) || /_P\./i.test(filename)) typeFromName = 'pitcher';
    else if (/_B_/i.test(filename) || /_B\./i.test(filename)) typeFromName = 'batter';
  }
  card.type = typeFromName || ((card.pitches && card.pitches.length > 0) ? 'pitcher' : 'batter');

  return card;
}

window.CARD_PARSER = { parse: parseCard };
})();
