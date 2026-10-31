const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
// Use backend's mongoose so models (Team, Player, …) share this connection.
// projectdocuments/node_modules/mongoose is a separate copy; connecting it
// leaves backend models disconnected and findOne() buffer-times out.
const mongoose = require(require.resolve('mongoose', {
  paths: [path.join(__dirname, '../backend'), __dirname],
}));
require('dotenv').config({ path: path.join(__dirname, '../backend/.env') });

const GAMES = ['CS2', 'Valorant', 'BGMI'];
const FORMS = ['excellent', 'good', 'average', 'poor'];

function parseArgs(argv) {
  const args = { file: null, dryRun: false, upsert: false };
  for (const arg of argv.slice(2)) {
    if (arg === '--dry-run') args.dryRun = true;
    else if (arg === '--upsert') args.upsert = true;
    else if (!arg.startsWith('-')) args.file = arg;
  }
  return args;
}

async function connectDB() {
  const uri = process.env.MONGODB_URI;
  if (!uri) {
    throw new Error('MONGODB_URI is missing. Set it in Full_project/backend/.env');
  }
  mongoose.set('bufferCommands', false);
  mongoose.set('autoIndex', false);
  await mongoose.connect(uri, {
    autoIndex: false,
    serverSelectionTimeoutMS: 15000,
    socketTimeoutMS: 45000,
  });
  console.log(`Connected to MongoDB: ${mongoose.connection.host}`);
}

function readSheetRows(filePath, sheetName) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }
  const workbook = XLSX.readFile(filePath);
  const name = sheetName && workbook.SheetNames.includes(sheetName)
    ? sheetName
    : workbook.SheetNames.find((n) => n.toLowerCase() !== 'instructions') || workbook.SheetNames[0];
  const sheet = workbook.Sheets[name];
  if (!sheet) {
    throw new Error(`No worksheet found in ${filePath}`);
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  return { sheetName: name, rows };
}

function cell(row, ...keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
      return String(row[key]).trim();
    }
    const match = Object.keys(row).find((k) => k.toLowerCase() === key.toLowerCase());
    if (match && row[match] !== undefined && row[match] !== null && String(row[match]).trim() !== '') {
      return String(row[match]).trim();
    }
  }
  return '';
}

function parseBoolean(value, defaultValue = true) {
  if (value === undefined || value === null || String(value).trim() === '') {
    return defaultValue;
  }
  const v = String(value).trim().toLowerCase();
  if (['true', 'yes', 'y', '1'].includes(v)) return true;
  if (['false', 'no', 'n', '0'].includes(v)) return false;
  return defaultValue;
}

function parseNumber(value, defaultValue) {
  if (value === undefined || value === null || String(value).trim() === '') {
    return defaultValue;
  }
  const n = Number(String(value).trim());
  return Number.isFinite(n) ? n : defaultValue;
}

function parseDate(value) {
  if (!value || String(value).trim() === '') return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

function parseList(value) {
  if (!value || String(value).trim() === '') return [];
  return String(value)
    .split(/[,;|]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function normalizeGame(value) {
  const g = String(value || '').trim();
  const match = GAMES.find((x) => x.toLowerCase() === g.toLowerCase());
  return match || g;
}

function exactMatch(value) {
  const escaped = String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`^${escaped}$`, 'i');
}

function printSummary(title, { inserted, updated, skipped, errors }) {
  console.log(`\n${title}`);
  console.log(`  Inserted: ${inserted}`);
  console.log(`  Updated:  ${updated}`);
  console.log(`  Skipped:  ${skipped}`);
  console.log(`  Errors:   ${errors.length}`);
  errors.forEach((err, i) => {
    console.log(`    ${i + 1}. Row ${err.row}: ${err.message}`);
  });
}

module.exports = {
  GAMES,
  FORMS,
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
};
