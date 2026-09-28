// ═══════════════════════════════════════════════════════
// VitalPair — Google Apps Script Backend v3
// Paste into Code.gs — bound to your Google Sheet
// ═══════════════════════════════════════════════════════

var SHEET_LOG    = 'DailyLog';
var SHEET_CONFIG = 'Config';
var SHEET_MEALS  = 'Meals';
var SHEET_HABITS = 'Habits';

// ── CORS + Entry points ──────────────────────────────────
function doGet(e) {
  var result;
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    result = route(params, null);
  } catch(err) {
    result = {ok: false, error: err.toString()};
  }
  return buildResponse(result);
}

function doPost(e) {
  var result;
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    result = route(params, body);
  } catch(err) {
    result = {ok: false, error: err.toString()};
  }
  return buildResponse(result);
}

function buildResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Router ───────────────────────────────────────────────
function route(params, body) {
  // Support payload param from no-cors GET fallback
  if (params.payload && !body) {
    try { body = JSON.parse(decodeURIComponent(params.payload)); } catch(e) {}
  }
  var action = params.action || (body ? body.action : '') || '';
  if (action === 'getConfig')    return getConfig();
  if (action === 'saveConfig')   return saveConfig(body);
  if (action === 'getMeals')     return getMeals();
  if (action === 'saveMeals')    return saveMeals(body);
  if (action === 'getHabits')    return getHabits();
  if (action === 'saveHabits')   return saveHabits(body);
  if (action === 'getLog')       return getLog(params.date || (body && body.date));
  if (action === 'getLogs')      return getLogs(params.from, params.to);
  if (action === 'saveLog')      return saveLog(body || params);
  if (action === 'getAnalytics') return getAnalytics(params.days);
  return {ok: true, message: 'VitalPair API v3 running'};
}

// ── Sheet helpers ────────────────────────────────────────
function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length)
        .setBackground('#1a2035')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  }
  return sh;
}

function toObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue; // skip empty rows
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j])] = data[i][j] !== undefined ? String(data[i][j]) : '';
    }
    out.push(obj);
  }
  return out;
}

// ── CONFIG ───────────────────────────────────────────────
function getConfig() {
  var sh = getSheet(SHEET_CONFIG, ['key','value']);
  var rows = toObjects(sh);
  var cfg = {};
  for (var i = 0; i < rows.length; i++) {
    cfg[rows[i].key] = rows[i].value;
  }
  return {ok: true, data: cfg};
}

function saveConfig(body) {
  var sh = getSheet(SHEET_CONFIG, ['key','value']);
  var cfg = (body && body.config) ? body.config : {};
  var keys = Object.keys(cfg);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var val = cfg[key];
    var data = sh.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        sh.getRange(i + 1, 2).setValue(val);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, val]);
  }
  return {ok: true};
}

// ── MEALS ────────────────────────────────────────────────
function getMeals() {
  var sh = getSheet(SHEET_MEALS, ['id','name','time','tags','active']);
  var rows = toObjects(sh);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].active !== 'false') out.push(rows[i]);
  }
  return {ok: true, data: out};
}

function saveMeals(body) {
  var sh = getSheet(SHEET_MEALS, ['id','name','time','tags','active']);
  var meals = (body && body.meals) ? body.meals : [];
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 5).clearContent();
  for (var i = 0; i < meals.length; i++) {
    var m = meals[i];
    sh.appendRow([m.id||'', m.name||'', m.time||'', m.tags||'', 'true']);
  }
  return {ok: true};
}

// ── HABITS ───────────────────────────────────────────────
function getHabits() {
  var sh = getSheet(SHEET_HABITS, ['id','name','icon','active']);
  var rows = toObjects(sh);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].active !== 'false') out.push(rows[i]);
  }
  return {ok: true, data: out};
}

function saveHabits(body) {
  var sh = getSheet(SHEET_HABITS, ['id','name','icon','active']);
  var habits = (body && body.habits) ? body.habits : [];
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  for (var i = 0; i < habits.length; i++) {
    var h = habits[i];
    sh.appendRow([h.id||'', h.name||'', h.icon||'', 'true']);
  }
  return {ok: true};
}

// ── DAILY LOG ────────────────────────────────────────────
var LOG_HEADERS = ['date','person','type','itemId','itemName','done','value','foods','cal','protein','notes'];

function getLog(date) {
  if (!date) {
    date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var sh = getSheet(SHEET_LOG, LOG_HEADERS);
  var rows = toObjects(sh);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].date === date) out.push(rows[i]);
  }
  return {ok: true, date: date, data: out};
}

function getLogs(from, to) {
  var sh = getSheet(SHEET_LOG, LOG_HEADERS);
  var rows = toObjects(sh);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (from && r.date < from) continue;
    if (to   && r.date > to)   continue;
    out.push(r);
  }
  return {ok: true, data: out};
}

function saveLog(body) {
  var sh      = getSheet(SHEET_LOG, LOG_HEADERS);
  var date    = (body && body.date)    ? body.date    : '';
  var person  = (body && body.person)  ? body.person  : '';
  var entries = (body && body.entries) ? body.entries : [];

  if (!date || !person) {
    return {ok: false, error: 'Missing date or person'};
  }

  // Remove existing rows for date+person (iterate backwards)
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === date && String(data[i][1]) === person) {
      sh.deleteRow(i + 1);
    }
  }

  // Append fresh entries
  for (var e = 0; e < entries.length; e++) {
    var en = entries[e];
    sh.appendRow([
      date,
      person,
      en.type     || '',
      en.itemId   || '',
      en.itemName || '',
      en.done ? 'true' : 'false',
      en.value    || '',
      en.foods    || '',
      en.cal      || '',
      en.protein  || '',
      en.notes    || ''
    ]);
  }
  return {ok: true, saved: entries.length};
}

// ── ANALYTICS ────────────────────────────────────────────
function getAnalytics(days) {
  days = parseInt(days) || 30;
  var tz   = Session.getScriptTimeZone();
  var now  = new Date();
  var from = new Date(now);
  from.setDate(now.getDate() - days + 1);
  var fromStr = Utilities.formatDate(from, tz, 'yyyy-MM-dd');
  var toStr   = Utilities.formatDate(now,  tz, 'yyyy-MM-dd');

  var rows = getLogs(fromStr, toStr).data;

  var dates = [];
  var d = new Date(from);
  while (d <= now) {
    dates.push(Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }

  var peopleMap = {};
  for (var i = 0; i < rows.length; i++) peopleMap[rows[i].person] = true;
  var people = Object.keys(peopleMap);

  var summary = {};
  for (var pi = 0; pi < people.length; pi++) {
    var p = people[pi];
    summary[p] = {dailyScores:{}, habitCompliance:{}, waterAvg:0, waterDays:0};
    for (var di = 0; di < dates.length; di++) {
      var date  = dates[di];
      var meals = [], habits = [], water = null;
      for (var ri = 0; ri < rows.length; ri++) {
        if (rows[ri].date !== date || rows[ri].person !== p) continue;
        if (rows[ri].type === 'meal')  meals.push(rows[ri]);
        if (rows[ri].type === 'habit') habits.push(rows[ri]);
        if (rows[ri].type === 'water') water = rows[ri];
      }
      var done  = 0;
      var total = meals.length + habits.length;
      for (var m = 0; m < meals.length;  m++) { if (meals[m].done  === 'true') done++; }
      for (var h = 0; h < habits.length; h++) {
        if (habits[h].done === 'true') done++;
        var hn = habits[h].itemName;
        if (!summary[p].habitCompliance[hn]) summary[p].habitCompliance[hn] = {done:0, total:0};
        summary[p].habitCompliance[hn].total++;
        if (habits[h].done === 'true') summary[p].habitCompliance[hn].done++;
      }
      summary[p].dailyScores[date] = total ? Math.round(done / total * 100) : null;
      if (water) {
        summary[p].waterAvg  += parseFloat(water.value) || 0;
        summary[p].waterDays += 1;
      }
    }
    if (summary[p].waterDays > 0) {
      summary[p].waterAvg = Math.round(summary[p].waterAvg / summary[p].waterDays * 10) / 10;
    }
  }
  return {ok: true, dates: dates, people: people, summary: summary};
}
