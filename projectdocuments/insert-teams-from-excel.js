#!/usr/bin/env node
/**
 * Insert esports teams from Excel into MongoDB (Team collection).
 * If a team with the same name and game already exists, it is skipped (not inserted).
 *
 * Usage:
 *   node insert-teams-from-excel.js
 *   node insert-teams-from-excel.js ./teams-template.xlsx
 *   node insert-teams-from-excel.js ./teams-template.xlsx --dry-run
 *
 * Uses MONGODB_URI from Full_project/backend/.env
 */
const path = require('path');
const {
  GAMES,
  parseArgs,
  connectDB,
  readSheetRows,
  cell,
  parseBoolean,
  parseDate,
  parseList,
  normalizeGame,
  printSummary,
  mongoose,
} = require('./excel-import-helpers');
const Team = require('../backend/models/Team');

const DEFAULT_FILE = path.join(__dirname, 'teams-template.xlsx');

function rowToTeam(row, rowNumber) {
  const name = cell(row, 'name');
  const game = normalizeGame(cell(row, 'game'));

  if (!name) {
    throw new Error('name is required');
  }
  if (!GAMES.includes(game)) {
    throw new Error(`game must be one of: ${GAMES.join(', ')}`);
  }

  const achievementsRaw = cell(row, 'achievements');
  let achievements = [];
  if (achievementsRaw) {
    try {
      const parsed = JSON.parse(achievementsRaw);
      achievements = Array.isArray(parsed) ? parsed : [];
    } catch {
      // tournament:position:year | tournament:position:year
      achievements = achievementsRaw.split('|').map((part) => {
        const [tournament, position, year] = part.split(':').map((s) => s.trim());
        return {
          tournament: tournament || '',
          position: position || '',
          year: year ? Number(year) : undefined,
        };
      }).filter((a) => a.tournament);
    }
  }

  return {
    name,
    nickname: cell(row, 'nickname') || '',
    logo: cell(row, 'logo') || 'default-team-logo.png',
    game,
    region: cell(row, 'region') || 'Unknown',
    founded: parseDate(cell(row, 'founded')),
    isActive: parseBoolean(cell(row, 'isActive'), true),
    players: parseList(cell(row, 'players')),
    achievements,
    socialMedia: {
      twitter: cell(row, 'twitter') || '',
      instagram: cell(row, 'instagram') || '',
      website: cell(row, 'website') || '',
    },
    _row: rowNumber,
  };
}

async function run() {
  const args = parseArgs(process.argv);
  const filePath = path.resolve(args.file || DEFAULT_FILE);

  console.log(`Reading teams from: ${filePath}`);
  const { sheetName, rows } = readSheetRows(filePath, 'Teams');
  console.log(`Sheet "${sheetName}": ${rows.length} data row(s)`);

  const stats = { inserted: 0, updated: 0, skipped: 0, errors: [] };
  const docs = [];

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    try {
      docs.push(rowToTeam(row, rowNumber));
    } catch (err) {
      stats.errors.push({ row: rowNumber, message: err.message });
    }
  });

  if (args.dryRun) {
    console.log('\nDry run — no database writes.');
    docs.forEach((d) => console.log(`  ${d.name} (${d.game})`));
    printSummary('Would process', { ...stats, inserted: docs.length });
    return;
  }

  await connectDB();

  try {
    for (const doc of docs) {
      const rowNumber = doc._row;
      delete doc._row;
      try {
        console.log(`Checking: ${doc.name} (${doc.game})`);
        const existing = await Team.findOne({
          name: doc.name,
          game: doc.game,
        }).maxTimeMS(15000);
        if (existing) {
          stats.skipped += 1;
          console.log(`Skip (same name + game): ${doc.name} (${doc.game})`);
          continue;
        }
        await Team.create(doc);
        stats.inserted += 1;
        console.log(`Inserted: ${doc.name} (${doc.game})`);
      } catch (err) {
        stats.errors.push({ row: rowNumber, message: err.message });
      }
    }
  } finally {
    await mongoose.connection.close();
    console.log('Database connection closed');
  }

  printSummary('Team import complete', stats);
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
