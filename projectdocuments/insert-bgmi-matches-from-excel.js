#!/usr/bin/env node
/**
 * Insert BGMI battle-royale matches from Excel into MongoDB (Match collection).
 * Each row is one lobby. Participating squads go in Match.teams (not team1 vs team2).
 *
 * The Valorant importer cannot be used for this: it rejects any game other than
 * Valorant and requires exactly two teams.
 *
 * Usage:
 *   node insert-bgmi-matches-from-excel.js
 *   node insert-bgmi-matches-from-excel.js ./bgmi-matches-template.xlsx
 *   node insert-bgmi-matches-from-excel.js ./bgmi-matches-template.xlsx --dry-run
 *
 * Uses MONGODB_URI from Full_project/backend/.env
 *
 * A row is skipped when a BGMI match already exists with the same tournament name,
 * match number, and start time.
 */
const path = require('path');
const {
  parseArgs,
  connectDB,
  readSheetRows,
  cell,
  parseBoolean,
  parseNumber,
  parseDate,
  parseList,
  normalizeGame,
  exactMatch,
  printSummary,
  mongoose,
} = require('./excel-import-helpers');
const Team = require('../backend/models/Team');
const Tournament = require('../backend/models/Tournament');
const Match = require('../backend/models/Match');

const DEFAULT_FILE = path.join(__dirname, 'bgmi-matches-template.xlsx');
const GAME = 'BGMI';
const STATUSES = ['upcoming', 'live', 'completed', 'cancelled'];
const TIERS = ['S', 'A', 'B', 'C'];
// Match.format only allows series labels. BGMI lobbies are stored as BO1.
const FORMAT = 'BO1';

function normalizeName(value) {
  return String(value || '')
    .replace(/[\u00A0\u2000-\u200B\u202F\u205F\u3000]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function namesMatch(left, right) {
  const a = normalizeName(left).toLowerCase();
  const b = normalizeName(right).toLowerCase();
  if (!a || !b) return false;
  return a === b || a.startsWith(`${b} `) || a.endsWith(` ${b}`) || b.startsWith(`${a} `) || b.endsWith(` ${a}`);
}

function parseTeamResults(raw) {
  if (!raw) return null;
  return String(raw).split('|').map((part) => {
    const [name, kills, placement, points] = part.split(':').map((s) => s.trim());
    const cleanName = normalizeName(name);
    if (!cleanName) throw new Error('teamResults has an empty squad name');
    return {
      name: cleanName,
      kills: Number(kills) || 0,
      placement: Number(placement) || 0,
      points: Number(points) || 0,
    };
  }).filter((t) => t.name);
}

function normalizeStatus(value) {
  const raw = String(value || '').trim().toLowerCase();
  return raw || 'upcoming';
}

function rowToMatch(row, rowNumber) {
  const gameCell = cell(row, 'game');
  const game = gameCell ? normalizeGame(gameCell) : GAME;
  if (game !== GAME) {
    throw new Error(`this script only imports BGMI matches (got "${gameCell}")`);
  }

  const tournamentName = cell(row, 'tournament', 'tournamentName');
  const startTime = parseDate(cell(row, 'startTime', 'start_time'));
  const status = normalizeStatus(cell(row, 'status'));
  const tier = (cell(row, 'tier', 'tournamentTier') || 'B').toUpperCase();
  const teamNames = parseList(cell(row, 'teams')).map(normalizeName);
  const teamResults = parseTeamResults(cell(row, 'teamResults', 'team_results'));

  if (!tournamentName) throw new Error('tournament is required');
  if (!startTime) throw new Error('startTime is required (e.g. 2026-09-22T14:30:00+05:30)');
  if (teamNames.length < 2) {
    throw new Error('teams must list at least 2 squad names, separated by commas');
  }
  const lowered = teamNames.map((n) => n.toLowerCase());
  if (new Set(lowered).size !== teamNames.length) {
    throw new Error('teams contains a duplicate squad name');
  }
  if (!STATUSES.includes(status)) {
    throw new Error(`status must be one of: ${STATUSES.join(', ')}`);
  }
  if (!TIERS.includes(tier)) {
    throw new Error(`tier must be one of: ${TIERS.join(', ')}`);
  }

  const winner = cell(row, 'winner') || null;
  if (winner && !lowered.includes(winner.toLowerCase())) {
    throw new Error('winner must be one of the teams on this row');
  }

  if (teamResults) {
    if (teamResults.length !== teamNames.length || teamResults.some((t, i) => !namesMatch(t.name, teamNames[i]))) {
      throw new Error('teamResults must list the same squads, in the same order, as teams');
    }
  }

  return {
    matchNumber: cell(row, 'matchNumber', 'match_number'),
    game: GAME,
    tournamentName,
    tournamentTier: tier,
    teamNames,
    teamResults,
    startTime,
    status,
    winner,
    map: cell(row, 'map') || 'TBD',
    viewerCount: parseNumber(cell(row, 'viewerCount', 'viewers'), 0),
    featured: parseBoolean(cell(row, 'featured'), false),
    isActive: parseBoolean(cell(row, 'isActive', 'is_active'), true),
    externalId: cell(row, 'externalId', 'external_id') || null,
    liquipediaGame: parseNumber(cell(row, 'liquipediaGame', 'liquipedia_game'), null),
    _row: rowNumber,
  };
}

function teamMatchesQuery(team, query) {
  return namesMatch(team.name, query) || namesMatch(team.nickname, query);
}

async function loadBgmiTeams() {
  return Team.find({ game: GAME }).maxTimeMS(15000);
}

function findBgmiTeam(teams, name) {
  const exact = teams.filter((team) => {
    const query = normalizeName(name).toLowerCase();
    return normalizeName(team.name).toLowerCase() === query
      || normalizeName(team.nickname).toLowerCase() === query;
  });
  const matches = exact.length ? exact : teams.filter((team) => teamMatchesQuery(team, name));
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) {
    throw new Error(`"${normalizeName(name)}" matches more than one BGMI team: ${matches.map((t) => t.name).join(', ')}`);
  }
  return null;
}

async function run() {
  const args = parseArgs(process.argv);
  const filePath = path.resolve(args.file || DEFAULT_FILE);

  console.log(`Reading BGMI matches from: ${filePath}`);
  const { sheetName, rows } = readSheetRows(filePath, 'Matches');
  console.log(`Sheet "${sheetName}": ${rows.length} data row(s)`);

  const stats = { inserted: 0, updated: 0, skipped: 0, errors: [] };
  const docs = [];

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    try {
      docs.push(rowToMatch(row, rowNumber));
    } catch (err) {
      stats.errors.push({ row: rowNumber, message: err.message });
    }
  });

  if (args.dryRun) {
    console.log('\nDry run — no database writes.');
    docs.forEach((d) => {
      console.log(`  ${d.matchNumber || '(no number)'} ${d.tournamentName} @ ${d.startTime.toISOString()} (${d.teamNames.length} teams)`);
    });
    printSummary('Would process', { ...stats, inserted: docs.length });
    return;
  }

  await connectDB();
  const bgmiTeams = await loadBgmiTeams();

  try {
    for (const doc of docs) {
      const rowNumber = doc._row;
      delete doc._row;
      try {
        console.log(`Checking: ${doc.tournamentName} ${doc.matchNumber || ''}`.trim());

        const existing = await Match.findOne({
          game: GAME,
          'tournament.name': exactMatch(doc.tournamentName),
          matchNumber: doc.matchNumber,
          startTime: doc.startTime,
        }).maxTimeMS(15000);

        if (existing) {
          stats.skipped += 1;
          console.log(`Skip (same tournament, match number, and start time): ${doc.matchNumber || existing._id}`);
          continue;
        }

        const resolved = [];
        for (const name of doc.teamNames) {
          const teamDoc = findBgmiTeam(bgmiTeams, name);
          if (!teamDoc) {
            throw new Error(`no BGMI team named "${normalizeName(name)}". Import teams first, and use the exact name from the Teams sheet.`);
          }
          const stats = doc.teamResults
            ? doc.teamResults.find((t) => namesMatch(t.name, teamDoc.name) || namesMatch(t.name, name))
            : null;
          resolved.push({
            name: teamDoc.name,
            logo: teamDoc.logo || 'default-team.png',
            kills: stats ? stats.kills : 0,
            placement: stats ? stats.placement : 0,
            points: stats ? stats.points : 0,
          });
        }

        const tournament = await Tournament.findOne({
          name: exactMatch(doc.tournamentName),
          game: GAME,
        }).maxTimeMS(15000);

        if (!tournament) {
          console.warn(`  Warning: no BGMI Tournament named "${doc.tournamentName}" — snapshot name only`);
        }

        const winnerTeam = doc.winner
          ? resolved.find((t) => namesMatch(t.name, doc.winner))
          : null;
        if (doc.winner && !winnerTeam) {
          throw new Error(`winner "${normalizeName(doc.winner)}" is not one of the resolved squads`);
        }
        const winner = winnerTeam ? winnerTeam.name : null;

        await Match.create({
          matchNumber: doc.matchNumber,
          game: GAME,
          tournament: {
            name: tournament ? tournament.name : doc.tournamentName,
            tier: tournament ? (tournament.tier || doc.tournamentTier) : doc.tournamentTier,
          },
          tournamentId: tournament ? tournament._id : null,
          teams: resolved,
          team1: { name: resolved[0].name, logo: resolved[0].logo, score: 0 },
          team2: { name: resolved[1].name, logo: resolved[1].logo, score: 0 },
          startTime: doc.startTime,
          status: doc.status,
          format: FORMAT,
          winner,
          map: doc.map,
          viewerCount: doc.viewerCount,
          featured: doc.featured,
          isActive: doc.isActive,
          externalId: doc.externalId,
          liquipediaGame: doc.liquipediaGame,
        });

        stats.inserted += 1;
        console.log(`Inserted: ${doc.matchNumber || doc.tournamentName} (${resolved.length} teams)`);
      } catch (err) {
        stats.errors.push({ row: rowNumber, message: err.message });
        console.error(`  Error row ${rowNumber}: ${err.message}`);
      }
    }
  } finally {
    await mongoose.connection.close();
    console.log('Database connection closed');
  }

  printSummary('BGMI match import complete', stats);
}

run().catch(async (err) => {
  console.error(err.message);
  try {
    await mongoose.connection.close();
  } catch {
    // ignore
  }
  process.exit(1);
});
