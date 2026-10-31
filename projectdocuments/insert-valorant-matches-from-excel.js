#!/usr/bin/env node
/**
 * Insert Valorant matches from Excel into MongoDB (Match collection).
 * Game is always Valorant. Rows for other games are rejected.
 *
 * Usage:
 *   node insert-valorant-matches-from-excel.js
 *   node insert-valorant-matches-from-excel.js ./valorant-matches-template.xlsx
 *   node insert-valorant-matches-from-excel.js ./valorant-matches-template.xlsx --dry-run
 *
 * Uses MONGODB_URI from Full_project/backend/.env
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
  normalizeGame,
  printSummary,
  mongoose,
} = require('./excel-import-helpers');
const Team = require('../backend/models/Team');
const Tournament = require('../backend/models/Tournament');
const Match = require('../backend/models/Match');

const DEFAULT_FILE = path.join(__dirname, 'valorant-matches-template.xlsx');
const GAME = 'Valorant';
const STATUSES = ['upcoming', 'live', 'completed', 'cancelled'];
const FORMATS = ['BO1', 'BO2', 'BO3', 'BO5'];
const TIERS = ['S', 'A', 'B', 'C'];

function normalizeFormat(value) {
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) return 'BO3';
  const match = FORMATS.find((f) => f === raw || f === `BO${raw}`);
  return match || raw;
}

function normalizeStatus(value) {
  const raw = String(value || '').trim().toLowerCase();
  return raw || 'upcoming';
}

function rowToMatch(row, rowNumber) {
  const gameCell = cell(row, 'game');
  const game = gameCell ? normalizeGame(gameCell) : GAME;
  if (game !== GAME) {
    throw new Error(`this script only imports Valorant matches (got "${gameCell}")`);
  }

  const team1Name = cell(row, 'team1', 'team1Name');
  const team2Name = cell(row, 'team2', 'team2Name');
  const tournamentName = cell(row, 'tournament', 'tournamentName');
  const startTime = parseDate(cell(row, 'startTime', 'start_time'));
  const status = normalizeStatus(cell(row, 'status'));
  const format = normalizeFormat(cell(row, 'format'));
  const tier = (cell(row, 'tier', 'tournamentTier') || 'B').toUpperCase();

  if (!team1Name) throw new Error('team1 is required');
  if (!team2Name) throw new Error('team2 is required');
  if (team1Name.toLowerCase() === team2Name.toLowerCase()) {
    throw new Error('team1 and team2 must be different');
  }
  if (!tournamentName) throw new Error('tournament is required');
  if (!startTime) throw new Error('startTime is required (e.g. 2026-09-20T18:00:00Z)');
  if (!STATUSES.includes(status)) {
    throw new Error(`status must be one of: ${STATUSES.join(', ')}`);
  }
  if (!FORMATS.includes(format)) {
    throw new Error(`format must be one of: ${FORMATS.join(', ')}`);
  }
  if (!TIERS.includes(tier)) {
    throw new Error(`tier must be one of: ${TIERS.join(', ')}`);
  }

  const winner = cell(row, 'winner') || null;
  if (winner && ![team1Name, team2Name].some((n) => n.toLowerCase() === winner.toLowerCase())) {
    throw new Error('winner must match team1 or team2');
  }

  return {
    matchNumber: cell(row, 'matchNumber', 'match_number'),
    game: GAME,
    tournamentName,
    tournamentTier: tier,
    team1Name,
    team1Logo: cell(row, 'team1Logo', 'team1_logo'),
    team2Name,
    team2Logo: cell(row, 'team2Logo', 'team2_logo'),
    team1Score: parseNumber(cell(row, 'team1Score', 'score1'), 0),
    team2Score: parseNumber(cell(row, 'team2Score', 'score2'), 0),
    startTime,
    status,
    format,
    winner,
    map: cell(row, 'map') || 'TBD',
    viewerCount: parseNumber(cell(row, 'viewerCount', 'viewers'), 0),
    featured: parseBoolean(cell(row, 'featured'), false),
    externalId: cell(row, 'externalId', 'external_id') || null,
    _row: rowNumber,
  };
}

async function findValorantTeam(name) {
  return Team.findOne({ name, game: GAME }).maxTimeMS(15000)
    || Team.findOne({ nickname: name, game: GAME }).maxTimeMS(15000);
}

async function run() {
  const args = parseArgs(process.argv);
  const filePath = path.resolve(args.file || DEFAULT_FILE);

  console.log(`Reading Valorant matches from: ${filePath}`);
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
      console.log(`  ${d.team1Name} vs ${d.team2Name} @ ${d.startTime.toISOString()} (${d.format}, ${d.tournamentName})`);
    });
    printSummary('Would process', { ...stats, inserted: docs.length });
    return;
  }

  await connectDB();

  try {
    for (const doc of docs) {
      const rowNumber = doc._row;
      delete doc._row;
      try {
        console.log(`Inserting: ${doc.team1Name} vs ${doc.team2Name}`);

        const [team1Doc, team2Doc, tournament] = await Promise.all([
          findValorantTeam(doc.team1Name),
          findValorantTeam(doc.team2Name),
          Tournament.findOne({ name: doc.tournamentName, game: GAME }).maxTimeMS(15000),
        ]);

        if (!team1Doc) {
          console.warn(`  Warning: no Valorant Team named "${doc.team1Name}" — using Excel name`);
        }
        if (!team2Doc) {
          console.warn(`  Warning: no Valorant Team named "${doc.team2Name}" — using Excel name`);
        }

        const team1Name = team1Doc ? team1Doc.name : doc.team1Name;
        const team2Name = team2Doc ? team2Doc.name : doc.team2Name;

        const payload = {
          matchNumber: doc.matchNumber,
          game: GAME,
          tournament: {
            name: tournament ? tournament.name : doc.tournamentName,
            tier: tournament ? (tournament.tier || doc.tournamentTier) : doc.tournamentTier,
          },
          tournamentId: tournament ? tournament._id : null,
          team1: {
            name: team1Name,
            logo: doc.team1Logo || (team1Doc && team1Doc.logo) || 'default-team.png',
            score: doc.team1Score,
          },
          team2: {
            name: team2Name,
            logo: doc.team2Logo || (team2Doc && team2Doc.logo) || 'default-team.png',
            score: doc.team2Score,
          },
          startTime: doc.startTime,
          status: doc.status,
          format: doc.format,
          winner: doc.winner,
          map: doc.map,
          viewerCount: doc.viewerCount,
          featured: doc.featured,
          externalId: doc.externalId,
        };

        if (!tournament) {
          console.warn(`  Warning: no Valorant Tournament named "${doc.tournamentName}" — snapshot name only`);
        }

        await Match.create(payload);
        stats.inserted += 1;
        console.log(`Inserted: ${team1Name} vs ${team2Name} (${doc.format})`);
      } catch (err) {
        stats.errors.push({ row: rowNumber, message: err.message });
      }
    }
  } finally {
    await mongoose.connection.close();
    console.log('Database connection closed');
  }

  printSummary('Valorant match import complete', stats);
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
