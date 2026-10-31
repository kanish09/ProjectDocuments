// liquipedia-match.js
// Usage: getMatchFromUrl("https://liquipedia.net/valorant/Match:ID_CHAMP26GrB_0001")

const WIKI_API = "https://liquipedia.net/valorant/api.php";
const UA       = "MyFantasyApp/1.0 (contact@yoursite.com)"; // ← update me

// ── Fetch & parse a single match by its Liquipedia URL ────────────────────────
export async function getMatchFromUrl(url) {
  const page = decodeURIComponent(new URL(url).pathname.split("/valorant/")[1]);

  const apiUrl = `${WIKI_API}?${new URLSearchParams({
    action: "parse", page, prop: "text", format: "json", formatversion: "2",
  })}`;

  const res = await fetch(apiUrl, {
    headers: { "User-Agent": UA, "Accept-Encoding": "gzip" },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const json = await res.json();
  if (json.error) throw new Error(json.error.info ?? JSON.stringify(json.error));

  return parseMatchHtml(page, json.parse?.text ?? "");
}

// ── Top-level match parser ─────────────────────────────────────────────────────
function parseMatchHtml(page, html) {
  const finished = /match-bm-match-header-result-text[^>]*>\s*finished/i.test(html);

  const scoreM = html.match(
    /class="match-bm-match-header-result">\s*(\d+)\s*&#160;:&#160;\s*(\d+)/
  );
  const score1 = scoreM ? parseInt(scoreM[1]) : null;
  const score2 = scoreM ? parseInt(scoreM[2]) : null;
  const winner = score1 != null && score2 != null && score1 !== score2
    ? (score1 > score2 ? 1 : 2) : null;

  const teams = [...html.matchAll(
    /class="match-bm-match-header-team-long"[^>]*>\s*<a[^>]*>([^<]+)<\/a>/g
  )].map(m => m[1].trim()).slice(0, 2);

  const dateM = html.match(/data-timestamp="(\d+)"/);
  const date  = dateM ? new Date(parseInt(dateM[1]) * 1000).toISOString() : null;

  // Overall player stats (first players-wrapper before any map section)
  const overallPlayers = parsePlayers(html.split("match-bm-lol-game-overview")[0]);

  // Per-map data (each match-bm-lol-game-overview section = one map)
  const maps = html.split("match-bm-lol-game-overview").slice(1).map(section => {
    const mapM    = section.match(/match-bm-lol-game-summary-map">([^<]+)/);
    const scoreM2 = section.match(/match-bm-lol-game-summary-score">(\d+)&#8209;(\d+)/);
    const lengthM = section.match(/match-bm-lol-game-summary-length">([^<]+)/);
    const ms1     = scoreM2 ? parseInt(scoreM2[1]) : null;
    const ms2     = scoreM2 ? parseInt(scoreM2[2]) : null;

    // Half scores (def/atk)
    const halvesT1 = section.match(
      /match-bm-lol-game-summary-team">[\s\S]*?half--def">(\d+)[\s\S]*?half--atk">(\d+)/
    );
    const halvesT2 = section.match(
      /data-highlighting-class="[^"]*"[\s\S]*?half--atk">(\d+)[\s\S]*?half--def">(\d+)/
    );

    return {
      map:     mapM    ? mapM[1].trim()   : null,
      length:  lengthM ? lengthM[1].trim(): null,
      score1:  ms1,
      score2:  ms2,
      winner:  ms1 != null && ms2 != null && ms1 !== ms2 ? (ms1 > ms2 ? 1 : 2) : null,
      players: parsePlayers(section),
    };
  });

  return {
    page, date, finished, winner,
    team1: { name: teams[0] ?? null, score: score1, players: overallPlayers[0] ?? [] },
    team2: { name: teams[1] ?? null, score: score2, players: overallPlayers[1] ?? [] },
    maps,
  };
}

// ── Parse player stats for both teams from an HTML chunk ──────────────────────
function parsePlayers(html) {
  const teamSections = html.split('class="match-bm-players-team"').slice(1);
  return teamSections.slice(0, 2).map(section => {
    const playerBlocks = section
      .split(/class="match-bm-players-player match-bm-players-player--col-\d+"/)
      .slice(1);

    return playerBlocks.map(block => {
      const nameM  = block.match(/match-bm-players-player-name">\s*<a[^>]*>([^<]+)<\/a>/);
      const name   = nameM ? nameM[1].trim() : null;

      const agentSection = block.match(/brkts-champion-icon">([\s\S]*?)<\/div>\s*<\/div>\s*<\/div>/);
      const agents = agentSection
        ? [...agentSection[1].matchAll(/<img alt="([^"]+)"/g)].map(m => m[1])
        : [];

      const statRe =
        /match-bm-players-player-stat-title"><i[^>]*><\/i>([^<]+)<\/div><div class="match-bm-players-player-stat-data">([\s\S]*?)<\/div>/g;

      const stats = {};
      for (const [, title, rawData] of block.matchAll(statRe)) {
        const value = rawData.replace(/<[^>]+>/g, "").trim();
        const key   = title.trim();
        if      (key === "KDA")    { const [k,d,a] = value.split("/").map(Number); stats.kills=k; stats.deaths=d; stats.assists=a; }
        else if (key === "FK / FD") { const [fk,fd] = value.split("/").map(Number); stats.firstKills=fk; stats.firstDeaths=fd; }
        else if (key === "ACS")    stats.acs    = parseFloat(value);
        else if (key === "KAST")   stats.kast   = value;
        else if (key === "ADR")    stats.adr    = parseFloat(value);
        else                       stats[key]   = value;
      }

      return { name, agents, ...stats };
    });
  });
}

// ── Run directly: node test.js <url> ──────────────────────────────────────────
if (import.meta.url === `file://${process.argv[1]}`) {
  const url = process.argv[2] ?? "https://liquipedia.net/valorant/Match:ID_CHAMP26GrB_0001";
  console.log(`Fetching: ${url}\n`);

  const match = await getMatchFromUrl(url);
  const t1 = match.team1.name, t2 = match.team2.name;

  console.log(`${t1}  ${match.team1.score} – ${match.team2.score}  ${t2}  (finished: ${match.finished})\n`);

  // ── Per-map breakdown ──
  for (const [i, map] of match.maps.entries()) {
    console.log(`\n━━ Map ${i+1}: ${map.map}  ${map.score1}–${map.score2}  (${map.length}) ━━`);
    console.log(`  Team 1: ${t1}`);
    console.table(map.players[0]?.map(p => ({
      player: p.name, agents: p.agents.join(", "),
      ACS: p.acs, K: p.kills, D: p.deaths, A: p.assists,
      KAST: p.kast, ADR: p.adr, FK: p.firstKills, FD: p.firstDeaths,
    })));
    console.log(`  Team 2: ${t2}`);
    console.table(map.players[1]?.map(p => ({
      player: p.name, agents: p.agents.join(", "),
      ACS: p.acs, K: p.kills, D: p.deaths, A: p.assists,
      KAST: p.kast, ADR: p.adr, FK: p.firstKills, FD: p.firstDeaths,
    })));
  }

  // ── Overall series stats ──
  console.log(`\n━━ Overall Series Stats ━━`);
  for (const [i, team] of [match.team1, match.team2].entries()) {
    console.log(`  Team ${i+1}: ${team.name}`);
    console.table(team.players.map(p => ({
      player: p.name, agents: p.agents.join(", "),
      ACS: p.acs, K: p.kills, D: p.deaths, A: p.assists,
      KAST: p.kast, ADR: p.adr, FK: p.firstKills, FD: p.firstDeaths,
    })));
  }
}
