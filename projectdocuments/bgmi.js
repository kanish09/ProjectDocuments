// bgmi.js — BGMI / PUBG Mobile Liquipedia scraper (Battle Royale format)
// ─────────────────────────────────────────────────────────────────────────────
// ✏️  CONFIGURE HERE — change these three values and run:  node bgmi.js
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {
  url:     "https://liquipedia.net/pubgmobile/Battlegrounds_Mobile_India_Showdown/2026",
  bracket: "Upper_Bracket",   // e.g. "Upper_Bracket" | "Lower_Bracket" | "Semi_Finals" | "Grand_Finals"
  game:    12,                  // game number to show (1, 2, 3 …) — set null for ALL games + overall
};
// ─────────────────────────────────────────────────────────────────────────────

const WIKI_API = "https://liquipedia.net/pubgmobile/api.php";
const UA       = "MyFantasyApp/1.0 (contact@yoursite.com)";  // ← update me

// ── Main export ───────────────────────────────────────────────────────────────
// page   : Liquipedia page path after /pubgmobile/
//          e.g.  "Battlegrounds_Mobile_India_Showdown/2026/Upper_Bracket"
// gameNo : optional 1-based game number. If omitted → all games + overall.
export async function getBgmiStandings(page, gameNo = null) {
  const apiUrl = `${WIKI_API}?${new URLSearchParams({
    action: "parse", page, prop: "text", format: "json", formatversion: "2",
  })}`;

  const res = await fetch(apiUrl, {
    headers: { "User-Agent": UA, "Accept-Encoding": "gzip" },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const json = await res.json();
  if (json.error) throw new Error(json.error.info ?? JSON.stringify(json.error));

  return parseBgmiHtml(json.parse?.text ?? "", gameNo);
}

// ── HTML parser ───────────────────────────────────────────────────────────────
function parseBgmiHtml(html, gameNo) {
  // NOTE: Liquipedia renders "__" as "&#95;&#95;" in class names
  // e.g.  panel-table__row  →  panel-table&#95;&#95;row

  // ── 1. Game headers ──
  const games = parseGameHeaders(html);

  // ── 2. Split into team rows ──
  // Rows use:  class="panel-table&#95;&#95;row" data-js-battle-royale="row"
  const rowSplit = html.split('data-js-battle-royale="row"').slice(1);

  const teams = rowSplit.map(rowHtml => {
    // Rank  →  data-sort-val="1" on cell--rank
    const rankM  = rowHtml.match(/cell--rank"[^>]*data-sort-val="(\d+)"/);
    const rank   = rankM ? parseInt(rankM[1]) : null;

    // Team full name  →  class="name hidden-xs" ... <a>Name</a>
    const longM  = rowHtml.match(/class="name hidden-xs"[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>/);
    const shortM = rowHtml.match(/class="name visible-xs"[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>/);
    const team   = longM?.[1].trim() ?? shortM?.[1].trim() ?? null;

    // Total points  →  data-sort-val="60" on cell--total-points
    const ptM    = rowHtml.match(/cell--total-points"[^>]*data-sort-val="(\d+)"/);
    const total  = ptM ? parseInt(ptM[1]) : null;

    // Per-game placements  (HTML uses &#95;&#95; for __ in class names)
    const placements = [...rowHtml.matchAll(
      /cell&#95;&#95;game-placement[^>]*><div[^>]*>[\s\S]*?<span[^>]*>([^<]+)<\/span>/g
    )].map(m => m[1].trim());

    // Per-game kills
    const kills = [...rowHtml.matchAll(
      /cell&#95;&#95;game-kills"><div[^>]*>(\d+)<\/div>/g
    )].map(m => parseInt(m[1]));

    // Placement points: 10, 6, 5, 4, 3, 2, 1, 1, then 0 from 9th
    const PLACE_PTS = { 1: 10, 2: 6, 3: 5, 4: 4, 5: 3, 6: 2, 7: 1, 8: 1 };
    const perGame   = placements.map((label, i) => {
      const place    = parseInt(label);
      const k        = kills[i] ?? null;
      const pPts     = PLACE_PTS[place] ?? 0;
      const pts      = k != null ? pPts + k : null;
      return { placement: isNaN(place) ? null : place, placementLabel: label, kills: k,
               placementPts: pPts, killPts: k, points: pts };
    });

    return { rank, team, total, games: perGame };
  }).filter(t => t.team); // skip header rows / empty

  // ── 3. Filter to single game if requested ──
  if (gameNo != null) {
    const idx      = gameNo - 1;
    const gameInfo = games[idx] ?? { number: gameNo, map: null, date: null, finished: false };
    const result   = teams
      .filter(t => t.games[idx] != null)
      .map(t => ({
        team:          t.team,
        placement:     t.games[idx].placement,
        placementLabel:t.games[idx].placementLabel,
        kills:         t.games[idx].kills,
        placementPts:  t.games[idx].placementPts,
        killPts:       t.games[idx].killPts,
        totalPoints:   t.games[idx].points,
      }))
      .sort((a, b) => (a.placement ?? 99) - (b.placement ?? 99));

    return { game: gameInfo, teams: result };
  }

  return { games, standings: teams };
}

// ── Parse game column headers ─────────────────────────────────────────────────
function parseGameHeaders(html) {
  // Each game column header:  panel-table__cell__game-title … Game N … data-timestamp="…"
  const headers = [];

  for (const m of html.matchAll(
    /cell&#95;&#95;game-title[\s\S]*?<span[^>]*>Game (\d+)<\/span>[\s\S]*?data-timestamp="(\d+)"(?:[\s\S]*?data-finished="([^"]*)")?/g
  )) {
    headers.push({
      number:   parseInt(m[1]),
      date:     new Date(parseInt(m[2]) * 1000).toISOString(),
      finished: m[3] === "finished",
      map:      null,   // filled below
    });
  }

  // Map names from carousel  "Game #N on MapName"
  for (const m of html.matchAll(/Game #(\d+) on <a[^>]*title="([^"]+)"/g)) {
    const idx = parseInt(m[1]) - 1;
    if (headers[idx]) headers[idx].map = m[2];
  }

  return headers;
}

// ── CLI ───────────────────────────────────────────────────────────────────────
if (import.meta.url === `file://${process.argv[1]}`) {
  // Build the Liquipedia page path from CONFIG
  // e.g. "https://liquipedia.net/pubgmobile/Battlegrounds_Mobile_India_Showdown/2026"
  //   + "Upper_Bracket"  →  "Battlegrounds_Mobile_India_Showdown/2026/Upper_Bracket"
  const basePath = new URL(CONFIG.url).pathname.replace("/pubgmobile/", "").replace(/\/$/, "");
  const PAGE     = CONFIG.bracket ? `${basePath}/${CONFIG.bracket}` : basePath;
  const gameNo   = CONFIG.game;

  const bracketLabel = CONFIG.bracket ?? "Main";
  console.log(`\nFetching: ${CONFIG.url}`);
  console.log(`Bracket : ${bracketLabel}`);
  console.log(`Game    : ${gameNo ?? "All"}\n`);

  const data = await getBgmiStandings(PAGE, gameNo);

  if (gameNo != null) {
    const g = data.game;
    const dateStr = g.date
      ? new Date(g.date).toLocaleString("en-IN", { timeZone: "Asia/Kolkata", dateStyle: "medium", timeStyle: "short" })
      : "TBD";
    console.log(`━━ ${bracketLabel}  |  Game ${g.number}${g.map ? " on " + g.map : ""}  |  ${dateStr}  |  Finished: ${g.finished} ━━\n`);
    console.table(data.teams.map(t => ({
      "#":          t.placement ?? "-",
      team:         t.team,
      kills:        t.kills ?? "-",
      "place pts":  t.placementPts ?? "-",
      "kill pts":   t.killPts ?? "-",
      "total pts":  t.totalPoints ?? "-",
    })));
  } else {
    console.log(`━━ ${bracketLabel}  |  Overall Standings ━━\n`);
    const colNames = data.games.map((g, i) => `G${i + 1}${g.map ? " " + g.map.slice(0, 4) : ""}`);
    console.table(data.standings.map(t => {
      const row = { "#": t.rank, team: t.team, "TOTAL": t.total };
      t.games.forEach((g, i) => {
        row[`${colNames[i]} pl`]  = g.placementLabel || "-";
        row[`${colNames[i]} k`]   = g.kills ?? "-";
        row[`${colNames[i]} pts`] = g.points ?? "-";
      });
      return row;
    }));

    if (data.games.length) {
      console.log("\n━━ Games ━━");
      console.table(data.games.map(g => ({
        game:     `Game ${g.number}`,
        map:      g.map ?? "?",
        date:     g.date ? new Date(g.date).toLocaleString("en-IN", { timeZone: "Asia/Kolkata", dateStyle: "short", timeStyle: "short" }) : "TBD",
        finished: g.finished,
      })));
    }
  }
}
