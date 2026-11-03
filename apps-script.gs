const MANAGERS_SHEET = 'Managers';
const COMPETITIONS_SHEET = 'Competitions';
const STAGE_DATE_COL = { flagged: 11, sent: 7, responded: 8, pilot: 9, passed: 10, nudge1: 17, nudge2: 18 }; // 1-indexed
const MANAGER_EXTRA_COL = { warmth: 12, intro: 13, followUp: 14, fit: 15, category: 16 }; // 1-indexed: L, M, N, O, P
const COMP_COL = { name: 2, fee: 3, deadline: 4, status: 5, notes: 6, url: 7 }; // 1-indexed: B, C, D, E, F, G

function doGet(e) {
  const type = (e.parameter.type || 'managers');
  if (type === 'competitions') return getCompetitions();
  return getManagers();
}

function getManagers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANAGERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).map(r => ({
    id: r[0],
    company: r[1],
    contact: r[2],
    email: r[3],
    stage: r[4],
    notes: r[5],
    sentDate: formatDate(r[6]),
    respondedDate: formatDate(r[7]),
    pilotDate: formatDate(r[8]),
    passedDate: formatDate(r[9]),
    flaggedDate: formatDate(r[10]),
    warmth: r[11] || '',
    intro: r[12] || '',
    followUp: formatIsoDate(r[13]),
    fit: r[14] || '',
    category: r[15] || 'manager',
    nudge1Date: formatDate(r[16]),
    nudge2Date: formatDate(r[17])
  })).filter(r => r.id !== '');
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function getCompetitions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPETITIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).map(r => ({
    id: r[0],
    name: r[1],
    fee: r[2],
    deadline: formatIsoDate(r[3]),
    status: r[4],
    notes: r[5],
    url: r[6] || ''
  })).filter(r => r.id !== '');
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
}

// The page's date picker needs yyyy-mm-dd; Sheets turns those into Date values, so format them back.
function formatIsoDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Parse yyyy-mm-dd as a local date so it doesn't shift a day with time zones.
function parseIsoDate(s) {
  if (!s) return '';
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const type = body.type || 'managers';
  if (type === 'competitions') return postCompetitions(body);
  return postManagers(body);
}

function postManagers(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANAGERS_SHEET);
  const action = body.action;

  if (action === 'add') {
    const id = Date.now();
    const today = new Date();
    const row = [id, body.company, body.contact || '', body.email || '', body.stage, body.notes || '', '', '', '', '', '',
      body.warmth || '', body.intro || '', parseIsoDate(body.followUp), body.fit || '', body.category || 'manager', '', ''];
    const dateCol = STAGE_DATE_COL[body.stage];
    if (dateCol) row[dateCol - 1] = today;
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, id }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'update') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        const rowNum = i + 1;
        if (body.stage !== undefined) {
          sheet.getRange(rowNum, 5).setValue(body.stage);
          const dateCol = STAGE_DATE_COL[body.stage];
          // Stamp today every time a stage is picked, even if it had an older date.
          if (dateCol) sheet.getRange(rowNum, dateCol).setValue(new Date());
        }
        if (body.notes !== undefined) {
          sheet.getRange(rowNum, 6).setValue(body.notes);
        }
        if (body.warmth !== undefined) sheet.getRange(rowNum, MANAGER_EXTRA_COL.warmth).setValue(body.warmth);
        if (body.intro !== undefined) sheet.getRange(rowNum, MANAGER_EXTRA_COL.intro).setValue(body.intro);
        if (body.followUp !== undefined) sheet.getRange(rowNum, MANAGER_EXTRA_COL.followUp).setValue(parseIsoDate(body.followUp));
        if (body.fit !== undefined) sheet.getRange(rowNum, MANAGER_EXTRA_COL.fit).setValue(body.fit);
        if (body.category !== undefined) sheet.getRange(rowNum, MANAGER_EXTRA_COL.category).setValue(body.category);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function postCompetitions(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPETITIONS_SHEET);
  const action = body.action;

  if (action === 'add') {
    const id = Date.now();
    sheet.appendRow([id, body.name, body.fee || '', parseIsoDate(body.deadline), body.status || '', body.notes || '', body.url || '']);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, id }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'update') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        const rowNum = i + 1;
        if (body.name !== undefined) sheet.getRange(rowNum, COMP_COL.name).setValue(body.name);
        if (body.fee !== undefined) sheet.getRange(rowNum, COMP_COL.fee).setValue(body.fee);
        if (body.deadline !== undefined) sheet.getRange(rowNum, COMP_COL.deadline).setValue(parseIsoDate(body.deadline));
        if (body.status !== undefined) sheet.getRange(rowNum, COMP_COL.status).setValue(body.status);
        if (body.notes !== undefined) sheet.getRange(rowNum, COMP_COL.notes).setValue(body.notes);
        if (body.url !== undefined) sheet.getRange(rowNum, COMP_COL.url).setValue(body.url);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}
