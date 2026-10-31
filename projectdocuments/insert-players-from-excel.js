#!/usr/bin/env node
/**
 * Insert players from Excel into MongoDB (Player collection).
 * If a player already exists for the same game + nickname (or real name),
 * a new document is not created; only team / teamId are updated.
 *
 * Usage:
 *   node insert-players-from-excel.js
 *   node insert-players-from-excel.js ./players-template.xlsx
 *   node insert-players-from-excel.js ./players-template.xlsx --dry-run
 *
 * Uses MONGODB_URI from Full_project/backend/.env
 */
const path = require('path');
const {
  GAMES,
  FORMS,
  parseArgs,
  connectDB,
  readSheetRows,
  cell,
  parseBoolean,
  parseNumber,
  normalizeGame,
  printSummary,
  mongoose,
} = require('./excel-import-helpers');
const Team = require('../backend/models/Team');
const Player = require('../backend/models/Player');

const DEFAULT_FILE = path.join(__dirname, 'players-template.xlsx');

function rowToPlayer(row, rowNumber) {
  const nickname = cell(row, 'nickname');
  const team = cell(row, 'team');
  const game = normalizeGame(cell(row, 'game'));
  const price = parseNumber(cell(row, 'price'), 10);
  const form = (cell(row, 'form') || 'average').toLowerCase();
  const role = cell(row, 'role');

  if (!nickname) {
    throw new Error('nickname is required');
  }
  if (!team) {
    throw new Error('team is required');
  }
  if (!GAMES.includes(game)) {
    throw new Error(`game must be one of: ${GAMES.join(', ')}`);
  }
  if (price < 5 || price > 15) {
    throw new Error('price must be between 5 and 15');
  }
  if (!FORMS.includes(form)) {
    throw new Error(`form must be one of: ${FORMS.join(', ')}`);
  }

  return {
    name: cell(row, 'name') || '',
    nickname,
    game,
    team,
    role: role || undefined,
    country: cell(row, 'country') || 'Unknown',
    image: cell(row, 'image') || 'default-player.png',
    price,
    stats: {
      kills: parseNumber(cell(row, 'kills'), 0),
      deaths: parseNumber(cell(row, 'deaths'), 0),
      assists: parseNumber(cell(row, 'assists'), 0),
      rating: parseNumber(cell(row, 'rating'), 1.0),
      headshot_percentage: parseNumber(cell(row, 'headshot_percentage'), 0),
      adr: parseNumber(cell(row, 'adr'), 0),
      kast: parseNumber(cell(row, 'kast'), 0),
      acs: parseNumber(cell(row, 'acs'), 0),
      first_bloods: parseNumber(cell(row, 'first_bloods'), 0),
      clutches: parseNumber(cell(row, 'clutches'), 0),
    },
    form,
    isActive: parseBoolean(cell(row, 'isActive'), true),
    lastUpdated: new Date(),
    _row: rowNumber,
  };
}

async function run() {
  const args = parseArgs(process.argv);
  const filePath = path.resolve(args.file || DEFAULT_FILE);

  console.log(`Reading players from: ${filePath}`);
  const { sheetName, rows } = readSheetRows(filePath, 'Players');
  console.log(`Sheet "${sheetName}": ${rows.length} data row(s)`);

  const stats = { inserted: 0, updated: 0, skipped: 0, errors: [] };
  const docs = [];

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    try {
      docs.push(rowToPlayer(row, rowNumber));
    } catch (err) {
      stats.errors.push({ row: rowNumber, message: err.message });
    }
  });

  if (args.dryRun) {
    console.log('\nDry run — no database writes.');
    docs.forEach((d) => console.log(`  ${d.nickname} @ ${d.team} (${d.game}) price=${d.price}`));
    printSummary('Would process', { ...stats, inserted: docs.length });
    return;
  }

  await connectDB();

  try {
    for (const doc of docs) {
      const rowNumber = doc._row;
      delete doc._row;
      try {
        console.log(`Checking: ${doc.nickname} (${doc.game})`);
        const teamDoc = await Team.findOne({ name: doc.team, game: doc.game }).maxTimeMS(15000)
          || await Team.findOne({ name: doc.team }).maxTimeMS(15000);
        if (teamDoc) {
          doc.teamId = teamDoc._id;
          doc.team = teamDoc.name;
        } else {
          doc.teamId = null;
          console.warn(`  Warning: no Team named "${doc.team}" — player ${doc.nickname} saved with team name only`);
        }

        const identity = [{ nickname: doc.nickname }];
        if (doc.name) {
          identity.push({ name: doc.name });
        }

        const existing = await Player.findOne({
          game: doc.game,
          $or: identity,
        }).maxTimeMS(15000);

        if (existing) {
          const teamChanged = String(existing.team) !== String(doc.team)
            || String(existing.teamId || '') !== String(doc.teamId || '');
          if (!teamChanged) {
            stats.skipped += 1;
            console.log(`Skip (same game + player, same team): ${doc.nickname} (${doc.game})`);
            continue;
          }
          await Player.updateOne(
            { _id: existing._id },
            {
              $set: {
                team: doc.team,
                teamId: doc.teamId,
                lastUpdated: new Date(),
              },
            }
          );
          stats.updated += 1;
          console.log(`Updated team only: ${doc.nickname} (${doc.game}) → ${doc.team}`);
        } else {
          await Player.create(doc);
          stats.inserted += 1;
          console.log(`Inserted: ${doc.nickname} (${doc.team})`);
        }
      } catch (err) {
        stats.errors.push({ row: rowNumber, message: err.message });
      }
    }
  } finally {
    await mongoose.connection.close();
    console.log('Database connection closed');
  }

  printSummary('Player import complete', stats);
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
