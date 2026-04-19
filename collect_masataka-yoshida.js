const https = require('https');
const ID = 807799;
const Y1 = 2023, Y2 = 2026;

function get(url) {
  return new Promise((res, rej) => {
    https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, r => {
      let d = ''; r.on('data', c => d += c); r.on('end', () => res(JSON.parse(d)));
    }).on('error', rej);
  });
}

(async () => {
  const yby = await get(`https://statsapi.mlb.com/api/v1/people/${ID}/stats?stats=yearByYear&group=hitting&sportId=1`);
  const allSplits = yby.stats[0].splits.filter(s => s.sport?.id === 1);

  const byYear = {};
  for (const s of allSplits) {
    const yr = s.season;
    if (!byYear[yr]) byYear[yr] = [];
    byYear[yr].push(s);
  }

  const years = Object.keys(byYear).filter(y => +y >= Y1 && +y <= Y2).sort();
  const basic = {};
  for (const yr of years) {
    const rows = byYear[yr];
    const row = rows.find(r => !r.team) || rows[0];
    let teamStr;
    if (rows.length > 1) {
      const named = rows.filter(r => r.team);
      const primary = named.reduce((a, b) => a.stat.gamesPlayed >= b.stat.gamesPlayed ? a : b);
      teamStr = (primary.team.abbreviation || primary.team.name.slice(0,3).toUpperCase()) + named.length;
    } else {
      teamStr = row.team?.abbreviation || row.team?.name?.slice(0,3).toUpperCase() || '???';
    }
    const st = row.stat;
    basic[yr] = {
      team: teamStr,
      g: st.gamesPlayed, pa: st.atBats,
      r: st.runs, h: st.hits, d: st.doubles, t: st.triples, hr: st.homeRuns,
      rbi: st.rbi, bb: st.baseOnBalls, so: st.strikeOuts,
      sb: st.stolenBases, cs: st.caughtStealing,
      avg: st.avg, obp: st.obp, ops: st.ops,
    };
  }

  const career = await get(`https://statsapi.mlb.com/api/v1/people/${ID}/stats?stats=career&group=hitting&sportId=1`);
  const cs = career.stats[0].splits[0].stat;
  basic['通算'] = {
    team: basic[years[years.length-1]]?.team?.replace(/\d+$/, '') || '---',
    g: cs.gamesPlayed, pa: cs.atBats,
    r: cs.runs, h: cs.hits, d: cs.doubles, t: cs.triples, hr: cs.homeRuns,
    rbi: cs.rbi, bb: cs.baseOnBalls, so: cs.strikeOuts,
    sb: cs.stolenBases, cs: cs.caughtStealing,
    avg: cs.avg, obp: cs.obp, ops: cs.ops,
  };

  const splitsRaw = {};
  await Promise.all(years.map(async yr => {
    const [vl, rp] = await Promise.all([
      get(`https://statsapi.mlb.com/api/v1/people/${ID}/stats?stats=statSplits&group=hitting&sportId=1&sitCodes=vl&season=${yr}`),
      get(`https://statsapi.mlb.com/api/v1/people/${ID}/stats?stats=statSplits&group=hitting&sportId=1&sitCodes=risp&season=${yr}`),
    ]);
    splitsRaw[yr] = {
      vsLAB:  vl.stats[0]?.splits[0]?.stat?.atBats || 0,
      vsLH:   vl.stats[0]?.splits[0]?.stat?.hits   || 0,
      rispAB: rp.stats[0]?.splits[0]?.stat?.atBats || 0,
      rispH:  rp.stats[0]?.splits[0]?.stat?.hits   || 0,
    };
  }));

  console.log('const years =', JSON.stringify(years) + ';');
  console.log('const basic =', JSON.stringify(basic, null, 2) + ';');
  console.log('const splitsRaw =', JSON.stringify(splitsRaw, null, 2) + ';');
})();
