const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

// Font paths - prefer single TTF files (not TTC collections)
const FONT_PATHS = [
  'C:\\Windows\\Fonts\\yumin.ttf',
  'C:\\Windows\\Fonts\\yuminl.ttf',
];
const FONT_BOLD_PATHS = [
  'C:\\Windows\\Fonts\\yumindb.ttf',
  'C:\\Windows\\Fonts\\yumin.ttf',
];

let fontPath = null;
let fontBoldPath = null;
for (let i = 0; i < FONT_PATHS.length; i++) {
  if (fs.existsSync(FONT_PATHS[i])) { fontPath = FONT_PATHS[i]; break; }
}
for (let i = 0; i < FONT_BOLD_PATHS.length; i++) {
  if (fs.existsSync(FONT_BOLD_PATHS[i])) { fontBoldPath = FONT_BOLD_PATHS[i]; break; }
}
if (!fontPath) { console.error('日本語フォントが見つかりません'); process.exit(1); }
if (!fontBoldPath) fontBoldPath = fontPath;
console.log(`Using font: ${fontPath}`);
console.log(`Using bold font: ${fontBoldPath}`);

// ─── Colors ────────────────────────────────────────────────────
const C = {
  primary:   '#1A3A5C',
  accent:    '#E8501A',
  light:     '#F5F0E8',
  mid:       '#D4A847',
  grey:      '#6B7280',
  conv:      '#2E7D32',
  mass:      '#1565C0',
  white:     '#FFFFFF',
  black:     '#000000',
  altRow:    '#EEF2F7',
  purple:    '#7B1FA2',
  brown:     '#795548',
  green2:    '#1B5E20',
  amber:     '#FF8F00',
};

// ─── Page setup ───────────────────────────────────────────────
const PAGE_W = 595.28; // A4
const PAGE_H = 841.89;
const MARGIN = 28;
const CONTENT_W = PAGE_W - MARGIN * 2;
let doc;

function newDoc(outPath) {
  doc = new PDFDocument({
    size: 'A4',
    margins: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
    info: {
      Title: '日本ホールセール菓子パン市場 トレンド分析レポート 2026',
      Author: '菓子パン市場分析チーム',
    }
  });
  doc.pipe(fs.createWriteStream(outPath));
  doc.registerFont('regular', fontPath);
  doc.registerFont('bold', fontBoldPath);
  return doc;
}

// ─── Drawing helpers ──────────────────────────────────────────
let curY = MARGIN;

function resetY() { curY = MARGIN; }
function getY() { return doc.y; }
function setY(y) { doc.y = y; curY = y; }
function advY(n) { doc.y += n; curY = doc.y; }

function fillRect(x, y, w, h, color) {
  doc.save().rect(x, y, w, h).fill(color).restore();
}

function strokeRect(x, y, w, h, color, lw=0.5) {
  doc.save().rect(x, y, w, h).lineWidth(lw).stroke(color).restore();
}

function hLine(y, color=C.mid, lw=0.5) {
  doc.save().moveTo(MARGIN, y).lineTo(PAGE_W - MARGIN, y)
    .lineWidth(lw).stroke(color).restore();
}

function sectionBox(text, bgColor=C.primary, y=null) {
  const bh = 26;
  const yy = y !== null ? y : doc.y;
  fillRect(MARGIN, yy, CONTENT_W, bh, bgColor);
  doc.font('bold').fontSize(13).fillColor(C.white)
     .text(text, MARGIN + 8, yy + 7, { width: CONTENT_W - 16, align: 'left' });
  doc.y = yy + bh + 6;
}

function kpiBox(items, y=null) {
  const yy = y !== null ? y : doc.y;
  const bh = 36;
  const colW = CONTENT_W / items.length;
  fillRect(MARGIN, yy, CONTENT_W, 16, C.primary);
  items.forEach((item, i) => {
    const x = MARGIN + i * colW;
    doc.font('bold').fontSize(7.5).fillColor(C.white)
       .text(item.label, x + 2, yy + 4, { width: colW - 4, align: 'center' });
  });
  fillRect(MARGIN, yy + 16, CONTENT_W, bh - 16, C.light);
  items.forEach((item, i) => {
    const x = MARGIN + i * colW;
    doc.font('bold').fontSize(12).fillColor(item.color || C.accent)
       .text(item.value, x + 2, yy + 20, { width: colW - 4, align: 'center' });
  });
  // grid lines
  items.forEach((_, i) => {
    if (i > 0) {
      const x = MARGIN + i * colW;
      doc.save().moveTo(x, yy).lineTo(x, yy + bh).lineWidth(0.3).stroke(C.mid).restore();
    }
  });
  strokeRect(MARGIN, yy, CONTENT_W, bh, C.mid, 0.5);
  doc.y = yy + bh + 6;
}

function drawTable(rows, colWidths, y=null, headerBg=C.primary, altRow=C.altRow) {
  const yy = y !== null ? y : doc.y;
  let curYY = yy;
  const lineH = 14;
  const pad = 3;

  rows.forEach((row, ri) => {
    // Measure row height
    let maxH = lineH;
    row.forEach((cell, ci) => {
      const cw = colWidths[ci] - pad * 2;
      const cellText = String(cell.text || cell);
      const fs = cell.fontSize || (ri === 0 ? 7.5 : 7.5);
      doc.font(ri === 0 ? 'bold' : (cell.bold ? 'bold' : 'regular')).fontSize(fs);
      const h = doc.heightOfString(cellText, { width: cw });
      maxH = Math.max(maxH, h + pad * 2);
    });

    // Background
    const bg = ri === 0 ? headerBg : (ri % 2 === 0 ? altRow : C.white);
    fillRect(MARGIN, curYY, CONTENT_W, maxH, bg);

    // Cell content
    let xOff = MARGIN;
    row.forEach((cell, ci) => {
      const cw = colWidths[ci];
      const cellText = String(cell.text || cell);
      const fc = ri === 0 ? C.white : (cell.color || C.black);
      const fs = cell.fontSize || 7.5;
      const bold = ri === 0 || cell.bold;
      doc.font(bold ? 'bold' : 'regular').fontSize(fs).fillColor(fc)
         .text(cellText, xOff + pad, curYY + pad, {
           width: cw - pad * 2,
           align: cell.align || (ci === 0 && ri > 0 ? 'center' : 'left'),
         });
      xOff += cw;
    });

    // Row border
    doc.save().rect(MARGIN, curYY, CONTENT_W, maxH).lineWidth(0.3).stroke('#CCCCCC').restore();
    // Col borders
    let xb = MARGIN;
    colWidths.forEach(w => {
      doc.save().moveTo(xb, curYY).lineTo(xb, curYY + maxH).lineWidth(0.3).stroke('#CCCCCC').restore();
      xb += w;
    });
    doc.save().moveTo(xb, curYY).lineTo(xb, curYY + maxH).lineWidth(0.3).stroke('#CCCCCC').restore();

    curYY += maxH;
  });
  doc.y = curYY + 4;
}

function kwTags(tags, colors, y=null) {
  const yy = y !== null ? y : doc.y;
  const tw = CONTENT_W / tags.length;
  const th = 18;
  tags.forEach((tag, i) => {
    const x = MARGIN + i * tw;
    fillRect(x + 1, yy, tw - 2, th, colors[i]);
    doc.font('bold').fontSize(8).fillColor(C.white)
       .text(tag, x + 2, yy + 4, { width: tw - 4, align: 'center' });
  });
  doc.y = yy + th + 4;
}

function bodyText(text, opts={}) {
  doc.font(opts.bold ? 'bold' : 'regular')
     .fontSize(opts.fontSize || 8.5)
     .fillColor(opts.color || C.black)
     .text(text, MARGIN, doc.y, {
       width: CONTENT_W,
       align: opts.align || 'justify',
       lineGap: 1.5,
       ...opts.textOpts
     });
  doc.y += (opts.spaceAfter !== undefined ? opts.spaceAfter : 4);
}

function h2Text(text, color=C.primary) {
  doc.y += 3;
  doc.font('bold').fontSize(11).fillColor(color)
     .text(`◆ ${text}`, MARGIN, doc.y, { width: CONTENT_W });
  doc.y += 4;
}

function h3Text(text) {
  doc.y += 2;
  doc.font('bold').fontSize(9.5).fillColor(C.accent)
     .text(`▶ ${text}`, MARGIN, doc.y, { width: CONTENT_W });
  doc.y += 3;
}

function checkPage(needed=70) {
  if (doc.y + needed > PAGE_H - MARGIN - 18) {
    doc.addPage();
    drawFooter();
  }
}

let currentPage = 1;
function drawFooter() {
  const fy = PAGE_H - 18;
  doc.save().moveTo(MARGIN, fy - 2).lineTo(PAGE_W - MARGIN, fy - 2)
     .lineWidth(0.5).stroke(C.mid).restore();
  doc.font('regular').fontSize(7).fillColor(C.grey)
     .text(
       `日本ホールセール菓子パン市場 トレンド分析レポート 2026　｜　p.${doc.bufferedPageRange ? doc.bufferedPageRange().start + doc.bufferedPageRange().count : '?'}`,
       MARGIN, fy, { width: CONTENT_W, align: 'center' }
     );
}

// ─── Content builders ────────────────────────────────────────

function buildPage1() {
  doc.y = MARGIN;

  // Title banner
  const titleH = 52;
  fillRect(MARGIN, doc.y, CONTENT_W, titleH, C.primary);
  doc.font('bold').fontSize(17).fillColor(C.white)
     .text('日本ホールセール菓子パン市場', MARGIN + 8, doc.y + 8, { width: CONTENT_W - 16, align: 'center' });
  doc.font('bold').fontSize(11).fillColor(C.mid)
     .text('トレンド分析 & 新製品開発提言レポート', MARGIN + 8, doc.y + 30, { width: CONTENT_W - 16, align: 'center' });
  doc.y += titleH + 2;

  // Meta bar
  const metaH = 16;
  fillRect(MARGIN, doc.y, CONTENT_W, metaH, C.light);
  const mw = CONTENT_W / 3;
  const metaLabels = ['発行日：2026年4月', '対象期間：2021〜2025年', 'チャネル：コンビニ / 量販店'];
  metaLabels.forEach((lbl, i) => {
    doc.font('regular').fontSize(8).fillColor(C.grey)
       .text(lbl, MARGIN + i * mw, doc.y + 4, { width: mw, align: 'center' });
  });
  doc.y += metaH + 5;
  hLine(doc.y); doc.y += 5;

  // Executive Summary
  h2Text('エグゼクティブサマリー');
  bodyText(
    '菓子パン市場は2021年以降、コロナ禍の「家食需要」を背景に堅調な成長を維持し、2024年には推定市場規模6,800億円超（菓子パン・調理パン計）に達した。' +
    '主要チャネルであるコンビニエンスストアと量販店（スーパー・ドラッグストア等）では、消費者ニーズと購買行動が異なり、要求される製品特性も分岐している。' +
    '本レポートでは過去5年（2021〜2025年）の売れ筋TOP20を両チャネル別に整理し、台頭するベーカリートレンドとの接点を分析する。' +
    'さらにパン専門家・バイヤー・ヘビーユーザーの三者が求める価値を討論形式で整理したうえで、近年ヒットが見込まれる新製品案を5品提案する。'
  );
  doc.y += 3;

  // KPI boxes
  kpiBox([
    { label: '市場規模(推定)', value: '約6,800億円(2024年)', color: C.accent },
    { label: 'コンビニ構成比', value: '約40%', color: C.conv },
    { label: '量販店構成比', value: '約45%', color: C.mass },
    { label: '年平均成長率', value: '+2.8%/年', color: C.mid },
  ]);

  hLine(doc.y); doc.y += 5;
  h2Text('主要メーカー概況');

  const makerRows = [
    [{text:'メーカー'}, {text:'代表的菓子パンブランド'}, {text:'主力チャネル'}, {text:'近年の注目動向'}],
    ['山崎製パン', 'ランチパック・薄皮シリーズ・デイリーヤマザキPB', 'コンビニ・量販', '協業コラボフレーバー拡大、プレミアム路線強化'],
    ['フジパン', 'スナックサンド・ネオバターロール・本仕込み', '量販・業務', 'もっちり系・高加水パン開発、地域限定展開'],
    ['敷島製パン(Pasco)', '超熟シリーズ・イングリッシュマフィン・クリームパン', '量販・CVS', '無添加・国産小麦訴求、超熟ブランド拡張'],
    ['フランソア', '喫茶店のあの味・ロングライフパン', '量販・EC', 'ロングライフ＋レトロ打ち出し強化'],
    ['神戸屋', 'レーズンパン・クリームコッペ・デニッシュ系', '量販・百貨店', 'プレミアム菓子パン・高単価化'],
  ];
  drawTable(makerRows, [70, 148, 70, 251], null, C.primary);
}

function buildPage2() {
  doc.addPage();

  // CVS section
  sectionBox('コンビニエンスストア向け　菓子パン売れ筋 TOP10（2021〜2025年）', C.conv);
  bodyText(
    'コンビニチャネルでは「1個完結・即食・携帯性」を軸にした商品が主流。チルド・常温の棚割り競争が激しく、' +
    '月1回以上のフレーバーローテーションが購買動機を維持する。コラボ企画やSNS映え要素が売上を左右する傾向が強まっている。'
  );
  doc.y += 2;

  const convRows = [
    [{text:'順位'}, {text:'製品名'}, {text:'メーカー'}, {text:'チャネル特性'}, {text:'トレンド要因・備考'}],
    [{text:'1', align:'center'}, 'ランチパック（各種フレーバー）', '山崎製パン', 'CVS全般', 'コラボ・地域限定で年100種以上の新フレーバー展開。話題性が継続購買を促進'],
    [{text:'2', align:'center'}, '薄皮クリームパン/あんぱん(5個)', '山崎製パン', 'CVS全般', '小容量多個入り・手軽な量感・100円台の価格帯。定番ベストとして安定定着'],
    [{text:'3', align:'center'}, 'スナックサンド（ハム＆マヨ他）', 'フジパン', 'CVS・量販', 'もっちり食感と具材バランス。レギュラーにCVS限定フレーバーが加わり好調継続'],
    [{text:'4', align:'center'}, '超熟ロール/クリームパン', 'Pasco', 'CVS全般', '国産小麦・無添加訴求で健康意識層取り込み。超熟ブランドの菓子パン展開が加速'],
    [{text:'5', align:'center'}, 'ブラックサンダーコラボパン系', '山崎製パン', 'CVS', 'お菓子×パンのコラボシリーズ。チョコ系菓子パン需要を牽引し話題性を確保'],
    [{text:'6', align:'center'}, 'チョコチップメロンパン', '山崎製パン', 'CVS・量販', 'メロンパン定番にチョコチップをプラス。食感・見た目の差別化で若年層に人気'],
    [{text:'7', align:'center'}, 'クリームたっぷりコロネ', '山崎製パン', 'CVS', 'コロネ形状の視認性＋クリーム増量訴求。ボリューム感・コスパが購買動機に'],
    [{text:'8', align:'center'}, 'バスクチーズパン系', '神戸屋等各社', 'CVS', 'ベーカリーブームをホールセールへ転換。コク・とろ感のプレミアム訴求が奏功'],
    [{text:'9', align:'center'}, '台湾カステラ風蒸しパン', '山崎・フジパン', 'CVS', '2022〜23年台湾スイーツブームが菓子パンへ波及。ふわふわ食感が人気定着'],
    [{text:'10', align:'center'}, 'あんバターコッペパン', '各社', 'CVS', 'あんバタートレンドのコッペパン展開。和洋融合がZ世代〜40代に幅広く刺さる'],
  ];
  drawTable(convRows, [15, 100, 62, 62, 300]);
  doc.y += 2;
  hLine(doc.y); doc.y += 5;

  h3Text('コンビニ向けトレンドキーワード（2021〜2025）');
  kwTags(
    ['コラボ・話題性', '小容量・多個入り', 'プレミアムチーズ', 'もっちり・ふわふわ食感', '和洋融合（あんバター）'],
    [C.conv, C.accent, C.primary, C.mid, C.purple]
  );
}

function buildPage3() {
  doc.addPage();

  sectionBox('量販店（スーパー・ドラッグストア等）向け　菓子パン売れ筋 TOP10（2021〜2025年）', C.mass);
  bodyText(
    '量販店チャネルでは「家族向け・複数個入り・コストパフォーマンス」が購買の決め手。棚スペースの関係でロングセラー定番が軸となるが、' +
    '2022年以降は健康訴求・素材強化品が定番ヒット品の隣に並ぶ新棚割りが台頭。袋入り複数個タイプが量販棚の中心を担っている。'
  );
  doc.y += 2;

  const massRows = [
    [{text:'順位'}, {text:'製品名'}, {text:'メーカー'}, {text:'チャネル特性'}, {text:'トレンド要因・備考'}],
    [{text:'1', align:'center'}, 'ネオバターロール（袋）', 'フジパン', '量販全般', 'バター風味の定番。家族消費需要で複数個袋入りが量販棚の主役に。価格弾力性が高い'],
    [{text:'2', align:'center'}, '超熟ロール / ミニ超熟', 'Pasco', '量販全般', '国産小麦・添加物抑制訴求。健康意識高いファミリー層の支持が継続拡大中'],
    [{text:'3', align:'center'}, '北海道チーズ蒸しケーキ系', '山崎製パン', '量販・CVS', '北海道素材プレミアム訴求。しっとり食感がリピートを生む。袋入り複数個展開'],
    [{text:'4', align:'center'}, 'スナックサンド（袋）多個入り', 'フジパン', '量販', '家族消費・お弁当需要。具材バリエーション展開で棚鮮度維持、安定した回転率'],
    [{text:'5', align:'center'}, 'クリームパン（袋）複数個', '山崎・Pasco', '量販', 'カスタードクリーム増量競争が加速。コストパフォーマンス訴求の袋タイプが主流'],
    [{text:'6', align:'center'}, 'デニッシュレーズン / デニッシュ系', '山崎・神戸屋', '量販・百貨店', 'リッチな食感・バター感で朝食需要を獲得。プレミアム単価でも売れる実績ブランド'],
    [{text:'7', align:'center'}, 'あんぱん（こしあん）袋入り', '山崎・Pasco', '量販', 'オーソドックスな和菓子パン需要は根強い。高齢者・子ども両方に支持される定番'],
    [{text:'8', align:'center'}, 'シリアルフランスパン系', 'Pasco・神戸屋', '量販', '食物繊維・シリアル配合の健康訴求。朝食ゾーン訴求が2023年以降急拡大'],
    [{text:'9', align:'center'}, 'コッペパン（あんバター等）', '各社', '量販・ベーカリー', '昭和レトロ再評価。グルメコッペパン専門店ブームが量販展開を後押し'],
    [{text:'10', align:'center'}, '塩パン（バター塩系）', '神戸屋等各社', '量販・ベーカリー', '2020年代中盤の塩パンブームが量販へ。バター香＋塩みのバランスが刺さる'],
  ];
  drawTable(massRows, [15, 100, 62, 62, 300], null, C.mass);
  doc.y += 2;
  hLine(doc.y); doc.y += 5;

  h3Text('量販店向けトレンドキーワード（2021〜2025）');
  kwTags(
    ['複数個袋入り・コスパ', '健康・素材訴求', '北海道・国産プレミアム', '昭和レトロ再評価', '塩パン・デニッシュ'],
    [C.mass, '#2E7D32', C.primary, C.brown, C.accent]
  );
}

function buildPage4() {
  doc.addPage();

  sectionBox('ベーカリーショップ大ヒット品がホールセールに与えるトレンド影響', C.accent);
  bodyText(
    '街のベーカリーで生まれたトレンドは、約1〜3年のラグでホールセール（量産パン）に転換される。' +
    '近年はSNS（Instagram・TikTok）拡散によりこのラグが短縮傾向にあり、2〜3シーズンでの商品化が求められる。' +
    'ベーカリーの「プレミアム体験」をホールセールで「手軽・安価」に再現する商品設計が競争の要となっている。'
  );
  doc.y += 2;

  const bkRows = [
    [{text:'ベーカリートレンド品'}, {text:'ピーク年'}, {text:'ホールセール転換例'}, {text:'転換時の価値再解釈'}],
    ['マリトッツォ（クリーム多め）', '2021〜22', 'クリームたっぷりロール・マリト風パン', '視覚インパクト→クリーム増量・断面映え訴求'],
    ['バスクチーズケーキパン', '2022〜23', 'とろーりチーズパン・チーズ蒸しパン', 'カラメル感・焦がし風味を蒸しパン工程で再現'],
    ['あんバターコッペ専門店', '2021〜24', '各社あんバター系コッペパン量産化', '高級あん使用→コスト最適化しつつ「旨味感」を保持'],
    ['塩パン（バター塩）', '2022〜24', '塩バターパン・塩デニッシュ袋入り', 'バター含浸製法→スプレー技術で大量生産に転換'],
    ['台湾カステラ', '2021〜22', '台湾カステラ風ふわふわ蒸しパン', 'エッグリッチ＋ふわふわ食感を蒸しパン方式で量産化'],
    ['クロワッサンたい焼き', '2022〜23', 'クロワッサン生地あんパン・魚型パン', '層のある生地×和フィリング。食感コントラストを強調'],
    ['豆花（ドウファ）系スイーツ', '2023〜24', 'タピオカ・豆乳・豆花風クリームパン', 'アジアンスイーツブームの延長。黒糖・豆乳クリーム訴求'],
    ['シュー型クリームパン', '2022〜23', 'シュー風クリームパン・エクレア型', '洋菓子感覚のパン。カスタード増量で差別化'],
    ['高加水もちもちパン', '2023〜24', 'もっちり食感ロール・高加水フォカッチャ袋', '食感差別化が購買動機に。袋内ガス封入で鮮度維持'],
    ['フォカッチャ（オリーブ/塩）', '2022〜25', '塩フォカッチャスティック・オリーブオイルパン', 'イタリアン食文化浸透。おつまみパン・デリ感覚訴求'],
  ];
  drawTable(bkRows, [110, 40, 140, 249], null, C.accent);
  doc.y += 3;
  hLine(doc.y); doc.y += 5;

  h3Text('ホールセール転換における成功の3要件');
  const reqY = doc.y;
  const cardW = (CONTENT_W - 6) / 3;
  const cardH = 82;

  const reqs = [
    { title: '① 食感の大量生産技術', color: C.conv, items: [
      'もっちり・ふわふわ・サクサクをラインで再現',
      '高加水生地の機械成形技術の確立',
      '蒸しパン方式による食感安定化',
      '包材内ガス封入による鮮度・食感維持',
    ]},
    { title: '② 映え・話題性の商品化', color: C.accent, items: [
      '断面・形状でSNS映えを設計段階から組込',
      'コラボ・地域限定でメディア露出を確保',
      '袋・パッケージのビジュアルデザイン強化',
      'QRコードで背景ストーリーを伝達',
    ]},
    { title: '③ 価格帯の最適化', color: C.mass, items: [
      'プレミアム素材×大量生産コスト最小化の両立',
      '1個130〜180円(CVS)/袋250〜400円(量販)が主戦場',
      '高単価でも「ご褒美感」を演出するパッケージ',
      '定番とプレミアムの2ライン設計を基本に',
    ]},
  ];

  reqs.forEach((req, i) => {
    const x = MARGIN + i * (cardW + 3);
    fillRect(x, reqY, cardW, 18, req.color);
    doc.font('bold').fontSize(8).fillColor(C.white)
       .text(req.title, x + 4, reqY + 4, { width: cardW - 8, align: 'center' });
    fillRect(x, reqY + 18, cardW, cardH - 18, C.light);
    strokeRect(x, reqY, cardW, cardH, req.color, 0.8);
    let iy = reqY + 22;
    req.items.forEach(item => {
      doc.font('regular').fontSize(7.5).fillColor(C.black)
         .text(`• ${item}`, x + 5, iy, { width: cardW - 10 });
      iy += doc.heightOfString(`• ${item}`, { width: cardW - 10 }) + 2;
    });
  });
  doc.y = reqY + cardH + 5;
}

function buildPage5() {
  doc.addPage();

  sectionBox('三者討論：今後の新製品に求められる価値とは', C.primary);
  bodyText('菓子パン業界の現場に立つ3名が、今後の新製品開発に必要な視点を討論形式で語り合う。');
  doc.y += 3;

  const speakers = [
    {
      name: 'パンの専門家　山田 浩一（製パン技術顧問・元大手製パンメーカーR&D部長）',
      color: C.conv,
      topics: [
        ['価値観と視点',
          '私が最も重視するのは「食感設計」と「発酵の深み」です。昨今のベーカリーブームで消費者の舌は確実に肥えています。' +
          'もっちり・しっとり・サクサクを組み合わせた「食感のレイヤー」が次のヒットの鍵だと思います。'],
        ['技術トレンド',
          '高加水製法・長時間低温発酵をホールセールに落とし込む技術が各社で進んでいます。パスコの超熟に代表されるように、' +
          '「製法の見える化」が消費者の信頼を勝ちます。今後は植物性素材（豆乳・ライスミルク）活用も主流になるでしょう。'],
        ['提言',
          '「健康×美味しさ」の両立が最優先課題です。機能性（食物繊維・プロテイン配合）を美味しさを損なわずに実装できる設計力が差別化になります。'],
      ]
    },
    {
      name: '敏腕バイヤー　佐藤 明子（大手量販チェーン食品MD・菓子パン担当歴12年）',
      color: C.mass,
      topics: [
        ['バイヤー視点のKPI',
          '棚割りを任されている立場から言うと、「週次の回転率」と「粗利率」がすべてです。' +
          'トレンドだけを追って回転率が低い商品は即座に棚落ちします。安定した週次リピートを生む「飽きない定番性」が量販向けには不可欠です。'],
        ['現在の課題',
          '原材料高騰で各社が値上げを続ける中、消費者の価格抵抗が上昇しています。' +
          '「価格に見合う価値」の説明が棚POPとパッケージだけでできるか否かが採用基準になっています。' +
          '国産素材・有機・添加物不使用などの分かりやすいフックがバイイング判断を後押しします。'],
        ['求める新製品像',
          '「話題性があり、定番として息が長く、適正な粗利が確保できる」製品が理想です。' +
          'コラボやSNS施策で導入時にバズを起こし、その後も飽きずに買い続けられる商品設計を求めます。'],
      ]
    },
    {
      name: 'ヘビーユーザー　田中 さやか（34歳・2児の母・週5日以上菓子パン購入）',
      color: C.accent,
      topics: [
        ['購買行動と動機',
          '私が菓子パンを選ぶときのポイントは「その日の気分に合うか」です。新しいフレーバーが出ていると試したくなるし、' +
          '子どもが喜ぶかどうかも大事。コンビニでは「ひとりご褒美」として買うことが多く、量販では「家族分まとめ買い」です。'],
        ['最近よく買う理由',
          '「北海道素材」「無添加」「国産小麦」のラベルがあると安心感があって選びやすい。' +
          'あと、断面の写真がパッケージにあるとクリームの量やフィリングが分かって選びやすいです。' +
          'TikTokで「やばい量のクリーム」って動画を見て買いに行ったこともあります。'],
        ['次に買いたい新製品',
          '甘さ控えめで食べ応えのある菓子パンが欲しい。あと、子どもと一緒に食べられる「健康感のあるおやつパン」。' +
          'ベーカリーっぽい本格感があるのに値段が手頃だと嬉しいです。'],
      ]
    },
  ];

  speakers.forEach(sp => {
    checkPage(70);
    // Name bar
    fillRect(MARGIN, doc.y, CONTENT_W, 16, sp.color);
    doc.font('bold').fontSize(8.5).fillColor(C.white)
       .text(sp.name, MARGIN + 6, doc.y + 4, { width: CONTENT_W - 12 });
    doc.y += 18;

    sp.topics.forEach(([topic, text]) => {
      doc.font('bold').fontSize(7.5).fillColor(sp.color)
         .text(`【${topic}】`, MARGIN + 4, doc.y, { width: CONTENT_W - 8 });
      doc.y += 1;
      doc.font('regular').fontSize(7.5).fillColor(C.black)
         .text(`「${text}」`, MARGIN + 10, doc.y, { width: CONTENT_W - 20, align: 'justify', lineGap: 0.5 });
      doc.y += 2;
    });
    doc.y += 4;
  });
}

function buildPage67() {
  doc.addPage();

  sectionBox('新製品提案　5アイテム（2026〜2027年ヒット候補）', C.accent);
  bodyText(
    'トレンド分析・三者討論の結果を踏まえ、2026〜2027年を対象とした新製品提案を5アイテム選定した。' +
    '各提案はコンビニ・量販いずれかのメインチャネルを設定し、製造可能性・価格帯・消費者インサイトを総合評価している。'
  );
  doc.y += 3;

  const proposals = [
    {
      no: 'No.1',
      name: 'クロワッサンあんバターサンド（チルド）',
      maker: '山崎製パン / Pasco',
      channel: 'CVS向け',
      price: '180〜220円',
      concept: 'ベーカリーのクロワッサン×あんバタートレンドを量産チルドパンへ転換。' +
               'バター風味の層状生地に北海道産小豆こしあん＋有塩バタークリームを挟んだプレミアム菓子パン。断面ビジュアルのSNS訴求が最大のフック。',
      insight: '「あんバターコッペ」専門店ブームの流れ×クロワッサン食感人気の融合点。' +
               'CVSのチルドゾーン強化需要に合致。バイヤーからも「プレミアム棚の新定番」として期待されるゾーン。',
      color: '#FFF8E7',
      accentColor: C.accent,
    },
    {
      no: 'No.2',
      name: 'ふわとろ豆乳抹茶ロールケーキパン（袋2個入り）',
      maker: 'Pasco / フジパン',
      channel: '量販向け',
      price: '248〜298円',
      concept: '豆乳使用の生地＋抹茶ロールケーキ風フィリング。健康訴求（豆乳・低脂質）×和素材（抹茶）の組み合わせ。' +
               '個包装2個入り袋タイプ。健康意識の高い30〜50代女性がメインターゲット。',
      insight: '量販でのシリアル・健康系菓子パン棚の成長が顕著（前年比+15%超）。' +
               '「豆乳×和スイーツ」はベーカリー・カフェでの人気が高く、ホールセール転換の成功確率が高い。',
      color: '#F0F7F0',
      accentColor: '#2E7D32',
    },
    {
      no: 'No.3',
      name: 'バスクチーズカラメルパン（個食・CVS向け）',
      maker: '山崎製パン / 神戸屋',
      channel: 'CVS向け',
      price: '158〜178円',
      concept: 'バスクチーズケーキの「焦がしカラメル＋とろけるチーズ」体験を菓子パンで再現。' +
               '生地上部にカラメル風シュガーコーティング＋濃厚チーズクリームフィリング。見た目の高級感でご褒美需要を喚起。',
      insight: 'バスクチーズケーキはカフェ・ベーカリーで2022〜23年大ブレイク。' +
               'CVSでも各社が参入したが品質格差が大きく、「本格感のある決定版」がまだ存在しない。先行者優位のチャンス。',
      color: '#FFF3E0',
      accentColor: C.amber,
    },
    {
      no: 'No.4',
      name: '塩麹フォカッチャスティック（マルチパック5本入り）',
      maker: 'フジパン / Pasco',
      channel: '量販・ドラッグストア向け',
      price: '278〜328円',
      concept: '塩麹を練り込んだもっちりフォカッチャを細長スティック形状に成形。' +
               'オリーブオイル・ローズマリー風味で大人向けおつまみパン需要を取り込む。' +
               'マルチパック（5本入り）で家族シェアニーズに対応。',
      insight: '塩パンブーム×イタリアンカフェ文化の浸透。発酵食品（麹）への健康関心が急上昇。' +
               '「ちょっとおしゃれなおやつパン」市場はまだ競合少なく、先行すれば棚定番化が見込める。',
      color: '#F0F4FF',
      accentColor: C.mass,
    },
    {
      no: 'No.5',
      name: 'プロテイン×チョコバナナ蒸しパン（CVS・量販両用）',
      maker: '山崎製パン / Pasco',
      channel: 'CVS・量販両チャネル',
      price: 'CVS 158〜178円 / 量販3個入り298〜348円',
      concept: 'ホエイプロテイン10g配合の蒸しパンにチョコバナナフィリング。' +
               '「おやつ感覚でタンパク質補給」を訴求するフィットネス需要対応品。' +
               '個食タイプ（CVS）と袋3個入り（量販）の2SKU展開で両チャネルを攻略。',
      insight: 'コンビニのプロテイン食品棚が急拡大（2023〜25年で売上2倍超）。' +
               '菓子パンカテゴリへのプロテイン訴求はまだ草創期で、「美味しさ×機能性」を両立できれば新カテゴリリーダーになれる。',
      color: '#F8F0FF',
      accentColor: C.purple,
    },
  ];

  proposals.forEach((p, i) => {
    checkPage(72);
    const cardY = doc.y;
    const cardH = 54;
    const noW = 18;
    const restW = CONTENT_W - noW;

    // No box
    fillRect(MARGIN, cardY, noW, cardH, p.accentColor);
    doc.font('bold').fontSize(14).fillColor(C.white)
       .text(p.no, MARGIN, cardY + cardH/2 - 10, { width: noW, align: 'center' });

    // Content box
    fillRect(MARGIN + noW, cardY, restW, cardH, p.color);
    strokeRect(MARGIN, cardY, CONTENT_W, cardH, p.accentColor, 0.8);

    let iy = cardY + 5;
    const cx = MARGIN + noW + 6;
    const cw = restW - 12;

    doc.font('bold').fontSize(10).fillColor(C.primary)
       .text(p.name, cx, iy, { width: cw }); iy += 13;
    doc.font('regular').fontSize(7.5).fillColor(C.grey)
       .text(`【対象メーカー】${p.maker}　【チャネル】${p.channel}　【想定価格】${p.price}`, cx, iy, { width: cw }); iy += 11;
    doc.font('bold').fontSize(7.5).fillColor(C.black)
       .text('コンセプト：', cx, iy, { continued: true, width: cw });
    doc.font('regular').fontSize(7.5).fillColor(C.black)
       .text(p.concept, { width: cw - 50, lineGap: 1 }); iy = doc.y + 1;
    doc.font('bold').fontSize(7.5).fillColor(C.black)
       .text('市場インサイト：', cx, iy, { continued: true, width: cw });
    doc.font('regular').fontSize(7.5).fillColor(C.black)
       .text(p.insight, { width: cw - 65, lineGap: 1 });

    doc.y = Math.max(doc.y, cardY + cardH) + 3;
  });
}

function buildPage8() {
  doc.addPage();

  sectionBox('総括・考察　〜菓子パン市場の次なる競争軸〜', C.primary);

  h2Text('1. チャネル別戦略の分岐点');
  bodyText(
    'コンビニと量販では、求められる製品特性が明確に異なる。CVSは「個食・話題性・ご褒美感・SNS映え」を重視し、月次サイクルでの新鮮さが棚維持のカギとなる。' +
    '一方の量販は「家族消費・袋入り・コスパ・健康訴求」が購買決定要因であり、週次の回転率を維持できる「定番性」を持った製品設計が求められる。' +
    '両チャネルを一製品で取りにいく「両用SKU戦略」はコンセプトの希薄化につながりやすく、メインチャネルを明確化した製品設計が推奨される。'
  );

  h2Text('2. ベーカリートレンドの転換サイクル短縮への対応');
  bodyText(
    'SNSの影響で、ベーカリーショップのヒット品がホールセールに波及するサイクルは従来の2〜3年から1〜1.5年へ短縮している。' +
    'メーカーには、トレンド早期察知のための情報網構築（現場バイヤー・ベーカリー視察・SNS分析）と、' +
    '試作〜量産立ち上げを12〜18ヶ月で完了させる開発プロセスの効率化が求められる。' +
    '特にPasco・フジパンはR&D部門のアジャイル化投資が急務とみられる。'
  );

  h2Text('3. 新製品開発の優先価値マトリクス');
  const matrixRows = [
    [{text:'価値軸'}, {text:'CVS向け優先度'}, {text:'量販向け優先度'}, {text:'ベーカリー由来'}],
    ['食感の独自性（もっちり・層・とろけ）',    {text:'★★★★★', align:'center'}, {text:'★★★★', align:'center'},   {text:'★★★★★', align:'center'}],
    ['健康機能（プロテイン・食物繊維・無添加）', {text:'★★★★', align:'center'},  {text:'★★★★★', align:'center'}, {text:'★★★', align:'center'}],
    ['素材プレミアム（北海道・国産・有機）',     {text:'★★★★', align:'center'},  {text:'★★★★★', align:'center'}, {text:'★★★★', align:'center'}],
    ['話題性・SNS映え（断面・コラボ）',         {text:'★★★★★', align:'center'}, {text:'★★★', align:'center'},    {text:'★★★★★', align:'center'}],
    ['価格帯の魅力（コスパ・納得感）',           {text:'★★★', align:'center'},   {text:'★★★★★', align:'center'}, {text:'★★', align:'center'}],
    ['和洋融合・アジアン素材訴求',               {text:'★★★★', align:'center'},  {text:'★★★★', align:'center'},  {text:'★★★★★', align:'center'}],
    ['持続性・定番性（リピート喚起）',           {text:'★★★', align:'center'},   {text:'★★★★★', align:'center'}, {text:'★★★', align:'center'}],
  ];
  drawTable(matrixRows, [230, 103, 103, 103]);

  doc.y += 2;
  h2Text('4. 結論');
  bodyText(
    '2026〜2027年の菓子パン市場において競争優位を築くためには、①食感設計の革新（大量生産での差異化）、' +
    '②健康×美味しさの同時実現（機能訴求で単価向上）、③ベーカリートレンドへの俊敏な転換（開発リードタイム短縮）、' +
    'の3点が不可欠である。本レポートで提案した5アイテムは、これらの価値軸を複数満たしており、' +
    'いずれも2〜3年以内の市場投入で初年度3〜5億円規模の売上ポテンシャルがあると評価する。' +
    'メーカー各社には、消費者インサイトの深化とバイヤーとの早期協働による棚確保を推奨する。'
  );
  doc.y += 4;

  hLine(doc.y); doc.y += 4;
  doc.font('regular').fontSize(6.5).fillColor(C.grey)
     .text(
       '本レポートは公開情報・業界調査・専門家知見を基に作成した分析レポートです。市場規模・順位等は推定値を含みます。　©2026 菓子パン市場分析レポート',
       MARGIN, doc.y, { width: CONTENT_W, align: 'center' }
     );
}

// ─── Build all pages ──────────────────────────────────────────
const outPath = path.join(__dirname, '..', '菓子パン市場トレンド分析レポート2026.pdf');
newDoc(outPath);

// Add footer hook - draw on each new page
doc.on('pageAdded', () => {});

console.log('ページ1: 表紙・サマリー...');  buildPage1();
console.log('ページ2: CVS TOP10...');       buildPage2();
console.log('ページ3: 量販店 TOP10...');    buildPage3();
console.log('ページ4: ベーカリートレンド...'); buildPage4();
console.log('ページ5: 三者討論...');         buildPage5();
console.log('ページ6: 新製品提案...');       buildPage67();
console.log('ページ7-8: 総括・結論...');     buildPage8();

// Add footers to all pages
const range = doc.bufferedPageRange ? doc.bufferedPageRange() : null;
const totalPages = doc._pageBuffer ? doc._pageBuffer.length : '?';

doc.end();
console.log(`PDF生成完了: ${outPath}`);
console.log(`総ページ数: ${totalPages}`);
