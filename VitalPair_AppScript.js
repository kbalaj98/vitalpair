// ═══════════════════════════════════════════════════════════════
// VitalPair — Google Apps Script Backend
// Deploy as: Extensions > Apps Script > Deploy > Web App
// Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

const SHEET_LOG     = 'DailyLog';
const SHEET_CONFIG  = 'Config';
const SHEET_MEALS   = 'Meals';
const SHEET_HABITS  = 'Habits';

// ── CORS wrapper ────────────────────────────────────────────────
function doGet(e)  { return respond(route(e)); }
function doPost(e) { return respond(route(e)); }

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function route(e) {
  try {
    const p = e.parameter || {};
    const body = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = p.action || body.action;

    switch (action) {
      case 'getConfig':   return getConfig();
      case 'saveConfig':  return saveConfig(body);
      case 'getMeals':    return getMeals();
      case 'saveMeals':   return saveMeals(body);
      case 'getHabits':   return getHabits();
      case 'saveHabits':  return saveHabits(body);
      case 'getLog':      return getLog(p.date);
      case 'getLogs':     return getLogs(p.from, p.to);
      case 'saveLog':     return saveLog(body);
      case 'getAnalytics':return getAnalytics(p.days);
      default:            return { ok: false, error: 'Unknown action: ' + action };
    }
  } catch(err) {
    return { ok: false, error: err.toString() };
  }
}

// ── Sheet helpers ────────────────────────────────────────────────
function getOrCreate(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length)
        .setBackground('#1a2035').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  return sh;
}

function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ── CONFIG ───────────────────────────────────────────────────────
function getConfig() {
  const sh = getOrCreate(SHEET_CONFIG, ['key','value']);
  const rows = sheetToObjects(sh);
  const cfg = {};
  rows.forEach(r => cfg[r.key] = r.value);
  return { ok: true, data: cfg };
}

function saveConfig(body) {
  const sh = getOrCreate(SHEET_CONFIG, ['key','value']);
  const cfg = body.config || {};
  // Upsert each key
  Object.entries(cfg).forEach(([key, value]) => {
    const data = sh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i + 1, 2).setValue(value);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, value]);
  });
  return { ok: true };
}

// ── MEALS ────────────────────────────────────────────────────────
function getMeals() {
  const sh = getOrCreate(SHEET_MEALS, ['id','name','time','foods','active']);
  const rows = sheetToObjects(sh);
  return { ok: true, data: rows.filter(r => r.active !== 'false') };
}

function saveMeals(body) {
  const sh = getOrCreate(SHEET_MEALS, ['id','name','time','foods','active']);
  const meals = body.meals || [];
  // Clear content (keep header)
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 5).clearContent();
  meals.forEach(m => sh.appendRow([m.id, m.name, m.time, m.foods, m.active !== false ? 'true' : 'false']));
  return { ok: true };
}

// ── HABITS ───────────────────────────────────────────────────────
function getHabits() {
  const sh = getOrCreate(SHEET_HABITS, ['id','name','icon','active']);
  const rows = sheetToObjects(sh);
  return { ok: true, data: rows.filter(r => r.active !== 'false') };
}

function saveHabits(body) {
  const sh = getOrCreate(SHEET_HABITS, ['id','name','icon','active']);
  const habits = body.habits || [];
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  habits.forEach(h => sh.appendRow([h.id, h.name, h.icon, h.active !== false ? 'true' : 'false']));
  return { ok: true };
}

// ── DAILY LOG ────────────────────────────────────────────────────
// Columns: date | person | type | itemId | itemName | done | value | notes
function getLog(date) {
  if (!date) date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sh = getOrCreate(SHEET_LOG, ['date','person','type','itemId','itemName','done','value','notes']);
  const rows = sheetToObjects(sh);
  const filtered = rows.filter(r => r.date === date);
  return { ok: true, date, data: filtered };
}

function getLogs(from, to) {
  const sh = getOrCreate(SHEET_LOG, ['date','person','type','itemId','itemName','done','value','notes']);
  const rows = sheetToObjects(sh);
  const filtered = rows.filter(r => {
    if (from && r.date < from) return false;
    if (to   && r.date > to)   return false;
    return true;
  });
  return { ok: true, data: filtered };
}

function saveLog(body) {
  // body: { date, person, entries: [{type, itemId, itemName, done, value, notes}] }
  const sh = getOrCreate(SHEET_LOG, ['date','person','type','itemId','itemName','done','value','notes']);
  const { date, person, entries } = body;
  if (!date || !person || !entries) return { ok: false, error: 'Missing date/person/entries' };

  // Remove existing rows for this date+person
  const data = sh.getDataRange().getValues();
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === date && data[i][1] === person) toDelete.push(i + 1);
  }
  toDelete.forEach(r => sh.deleteRow(r));

  // Append new entries
  entries.forEach(e => {
    sh.appendRow([date, person, e.type, e.itemId, e.itemName, e.done ? 'true' : 'false', e.value || '', e.notes || '']);
  });

  return { ok: true };
}

// ── ANALYTICS ────────────────────────────────────────────────────
function getAnalytics(days) {
  days = parseInt(days) || 30;
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const from = new Date(now); from.setDate(now.getDate() - days + 1);
  const fromStr = Utilities.formatDate(from, tz, 'yyyy-MM-dd');
  const toStr   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const logsRes = getLogs(fromStr, toStr);
  const rows = logsRes.data;

  // Build date range
  const dates = [];
  for (let d = new Date(from); d <= now; d.setDate(d.getDate() + 1)) {
    dates.push(Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd'));
  }

  const people = [...new Set(rows.map(r => r.person))];
  const summary = {};

  people.forEach(p => {
    summary[p] = {
      dailyScores: {},
      habitCompliance: {},
      mealCompliance: {},
      waterAvg: 0,
      waterDays: 0,
    };
    dates.forEach(date => {
      const dayRows = rows.filter(r => r.date === date && r.person === p);
      const meals   = dayRows.filter(r => r.type === 'meal');
      const habits  = dayRows.filter(r => r.type === 'habit');
      const water   = dayRows.find(r => r.type === 'water');
      const done    = [...meals, ...habits].filter(r => r.done === 'true').length;
      const total   = meals.length + habits.length;
      summary[p].dailyScores[date] = total ? Math.round(done / total * 100) : null;
      if (water) { summary[p].waterAvg += parseFloat(water.value) || 0; summary[p].waterDays++; }
      habits.forEach(h => {
        if (!summary[p].habitCompliance[h.itemName]) summary[p].habitCompliance[h.itemName] = { done: 0, total: 0 };
        summary[p].habitCompliance[h.itemName].total++;
        if (h.done === 'true') summary[p].habitCompliance[h.itemName].done++;
      });
    });
    if (summary[p].waterDays) summary[p].waterAvg = +(summary[p].waterAvg / summary[p].waterDays).toFixed(1);
  });

  return { ok: true, dates, people, summary };
}
