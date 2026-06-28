// 添付カードHTML群を解析して players.js を生成する
// 使い方: node _parse_cards.js
const fs = require('fs');
const path = require('path');

const DOWNLOADS = 'C:\\Users\\unatt\\Downloads';
const FILES = [
  'Deジーター_1999_card.html',
  'Aaジャッジ_2022_card.html',
  'Geコール_2023_card.html',
  'Raジョンソン_2001_card.html',
  'Scシールズ_2005_card.html',
  'Buポージー_2012_card (1).html',
  'Wiメイズ_1965_card.html',
];

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
  // 原本HTMLをそのまま保存（写真・フリップ含む全レイアウト再現用）
  card.rawHtml = html;
  // ban名 / season
  card.nameJa = decode(pick(html, /<span class="player-name-ja">([\s\S]*?)<\/span>/));
  card.nameEn = decode(pick(html, /<span class="player-name-en">([\s\S]*?)<\/span>/));
  card.position = decode(pick(html, /<span class="position-badge">([\s\S]*?)<\/span>/));
  card.fullNameTop = decode(pick(html, /<span class="player-full-name"[^>]*>([\s\S]*?)<\/span>/));
  const season = decode(pick(html, /<span class="season-badge">([\s\S]*?)<\/span>/));
  card.seasonLabel = season; // ex: "1999 PEAK"
  const ym = season.match(/(\d{4})/);
  card.year = ym ? parseInt(ym[1]) : null;

  // チームバッジ
  card.team = decode(pick(html, /<div class="team-badge"[^>]*>([\s\S]*?)<\/div>/));
  // 写真: photo-frame 内の img src (data: URI または通常URL)
  const photoMatch = html.match(/<div class="photo-frame">\s*<img[^>]+src="([^"]+)"/);
  if (photoMatch) card.photo = photoMatch[1];
  // 投打
  const hand = decode(pick(html, /<span class="card-back-banner-hand">([\s\S]*?)<\/span>/));
  card.hand = hand; // ex: "右投右打"
  // パワーディスプレイ(総合値)
  card.powerValue = num(pick(html, /<div class="power-value"[^>]*>([\s\S]*?)<\/div>/));

  // record-bar 中の rec-label/rec-val を順に
  const recBar = pick(html, /<div class="record-bar">([\s\S]*?)<\/div>\s*<div class="sec-title">/);
  const recLabels = pickAll(recBar, /<span class="rec-label">([\s\S]*?)<\/span>/g).map(decode);
  const recVals = pickAll(recBar, /<span class="rec-val[^"]*">([\s\S]*?)<\/span>/g).map(decode);
  card.record = {};
  for (let i = 0; i < recLabels.length; i++) card.record[recLabels[i]] = recVals[i];

  // ゲームステータス stats-grid (4個) + stats-grid-bot (選球眼/三振耐性 + ミニ3個)
  // ゲームステータスから 対球種ポイント or 球種テーブル or 守備セクションまでをひとまとめにパース
  const statBlock = pick(html, /■ ゲームステータス([\s\S]*?)(?:<table class="pitch-table"|<div class="sec-title")/);
  if (statBlock) {
    // stat-name (mini除外) を抽出
    const names = pickAll(statBlock, /<span class="stat-name(?!-)[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const vals  = pickAll(statBlock, /<span class="stat-val(?!-)[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const namesMini = pickAll(statBlock, /<span class="stat-name-mini[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    const valsMini  = pickAll(statBlock, /<span class="stat-val-mini[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.stats = {};
    for (let i = 0; i < names.length; i++) {
      // 全角・半角空白を除去してキー正規化 (例: "制　球" → "制球")
      const key = names[i].replace(/\s+/g, '');
      card.stats[key] = num(vals[i]);
    }
    card.statsMini = {};
    for (let i = 0; i < namesMini.length; i++) {
      const key = namesMini[i].replace(/\s+/g, '');
      card.statsMini[key] = num(valsMini[i]);
    }
  }

  // 対球種ポイント (打者カード) または 球種テーブル (投手カード)
  const pitchGrid = pick(html, /■ 対球種ポイント[\s\S]*?<div class="pitch-grid">([\s\S]*?)<\/div>\s*<\/div>/);
  if (pitchGrid) {
    const lbls = pickAll(pitchGrid, /<span class="pitch-lbl">([\s\S]*?)<\/span>/g).map(decode);
    const vs = pickAll(pitchGrid, /<span class="pitch-v[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.pitchPoints = {};
    for (let i = 0; i < lbls.length; i++) card.pitchPoints[lbls[i]] = num(vs[i]);
  }

  // 投手カードの球種テーブル(pitch-table)
  const pitchTable = pick(html, /<table class="pitch-table">([\s\S]*?)<\/table>/);
  if (pitchTable) {
    const rows = pickAll(pitchTable, /<tr>([\s\S]*?)<\/tr>/g);
    card.pitches = [];
    for (const row of rows) {
      const cells = pickAll(row, /<td[^>]*>([\s\S]*?)<\/td>/g).map(decode);
      if (cells.length >= 4) {
        // pn(球種名), ps(球速), pp2(球威), pr(割合)
        card.pitches.push({
          name: cells[0],
          speed: num(cells[1]),
          power: num(cells[2]),
          ratio: num(cells[3]),
        });
      }
    }
  }

  // 守備DRS (drs-bar) — 複数ポジ対応
  // 守備DRS: drs-bar ブロックの境界に依存せず、全 drs-pos / drs-num / drs-inn を順番にマッチ
  // (旧版は drs-bar 内の最初の drs-item しか拾えず、マルチポジ選手が壊れていた)
  const allPoses = pickAll(html, /<span class="drs-pos">([\s\S]*?)<\/span>/g).map(decode);
  const allNums  = pickAll(html, /<span class="drs-num[^"]*">([\s\S]*?)<\/span>/g).map(decode);
  const allInns  = pickAll(html, /<span class="drs-inn">([\s\S]*?)<\/span>/g).map(decode);
  card.drs = [];
  for (let i = 0; i < allPoses.length; i++) {
    card.drs.push({
      pos: allPoses[i],
      value: num(allNums[i]),
      innings: num(allInns[i]),
    });
  }
  // 同ポジション重複を除去 (combo-row 内などで2回出る場合)
  const seen = new Set();
  card.drs = card.drs.filter(d => {
    if (seen.has(d.pos)) return false;
    seen.add(d.pos);
    return true;
  });

  // 捕手の リード/盗塁阻止
  const catcherBar = pick(html, /<div class="catcher-bar">([\s\S]*?)<\/div>(?=\s*<\/div>|\s*<div)/);
  if (catcherBar) {
    const labels = pickAll(catcherBar, /<span class="ca-label">([\s\S]*?)<\/span>/g).map(decode);
    const vals = pickAll(catcherBar, /<span class="ca-val[^"]*">([\s\S]*?)<\/span>/g).map(decode);
    card.catcher = {};
    for (let i = 0; i < labels.length; i++) card.catcher[labels[i]] = num(vals[i]);
  }

  // 列伝
  card.retsuden = decode(pick(html, /<div class="retsuden"[^>]*>([\s\S]*?)<\/div>/));

  // タイプ判定: pitches 要素があれば投手、なければ打者
  card.type = (card.pitches && card.pitches.length > 0) ? 'pitcher' : 'batter';

  return card;
}

const out = [];
for (const f of FILES) {
  const fp = path.join(DOWNLOADS, f);
  if (!fs.existsSync(fp)) {
    console.error('NOT FOUND:', fp);
    continue;
  }
  const html = fs.readFileSync(fp, 'utf8');
  const c = parseCard(html, f);
  out.push(c);
  console.log('--- ' + f + ' ---');
  console.log(JSON.stringify(c, null, 2));
}

// players.js を生成
const outFile = path.join(__dirname, 'players.js');
const content = '// 自動生成: _parse_cards.js から生成された選手データ\n' +
  '// 新しいカードを追加する場合は、import.html を使ってデータを追加してください\n' +
  'window.PLAYERS = ' + JSON.stringify(out, null, 2) + ';\n';
fs.writeFileSync(outFile, content, 'utf8');
console.log('\n✅ 生成完了: ' + outFile);
console.log('選手数: ' + out.length);
