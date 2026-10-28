/**
 * IGCSE Study Hub - Google Sheets database backend
 *
 * IMPORTANT:
 * 1. Create a Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Paste this entire file into Code.gs.
 * 4. Run setupDatabase() once and authorize it.
 * 5. Deploy as Web app: Execute as Me, Who has access: Anyone.
 * 6. Copy the /exec URL into google-config.js in the website.
 *
 * This is designed for a school/student project. Passwords are never stored
 * as plain text: the website sends a SHA-256 password hash and this script
 * stores that hash. Use HTTPS (the Apps Script /exec URL is HTTPS).
 */

const USERS_SHEET = 'Users';
const PROGRESS_SHEET = 'Progress';

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let users = ss.getSheetByName(USERS_SHEET);
  let progress = ss.getSheetByName(PROGRESS_SHEET);

  if (!users) users = ss.insertSheet(USERS_SHEET);
  if (!progress) progress = ss.insertSheet(PROGRESS_SHEET);

  if (users.getLastRow() === 0) {
    users.appendRow(['username', 'password_hash', 'token', 'created_at', 'last_login']);
  }
  if (progress.getLastRow() === 0) {
    progress.appendRow([
      'username', 'subject', 'question_index', 'selected_index',
      'correct', 'skipped', 'timestamp'
    ]);
  }

  users.setFrozenRows(1);
  progress.setFrozenRows(1);
  users.autoResizeColumns(1, 5);
  progress.autoResizeColumns(1, 7);

  return 'Database ready.';
}

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || '').trim();
    const result = handleAction(action, p);
    return respond_(result, p.callback);
  } catch (err) {
    return respond_({ ok: false, error: safeMessage_(err) }, e && e.parameter ? e.parameter.callback : '');
  }
}

function handleAction(action, p) {
  switch (action) {
    case 'signup': return signup_(p);
    case 'login': return login_(p);
    case 'session': return session_(p);
    case 'logout': return logout_(p);
    case 'allProgress': return allProgress_(p);
    case 'saveAnswer': return saveAnswer_(p);
    case 'resetSubject': return resetSubject_(p);
    default: return { ok: false, error: 'Unknown action.' };
  }
}

function signup_(p) {
  const username = cleanUsername_(p.username);
  const passwordHash = String(p.passwordHash || '').trim();

  if (!username) return { ok: false, error: 'Username is required.' };
  if (!/^[A-Za-z0-9_.-]{3,30}$/.test(username)) {
    return { ok: false, error: 'Username must be 3-30 characters and use letters, numbers, _, ., or -.' };
  }
  if (!/^[a-f0-9]{64}$/i.test(passwordHash)) {
    return { ok: false, error: 'Invalid password data.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet_(USERS_SHEET);
    const rows = getRows_(sheet);
    if (rows.some(r => String(r[0]).toLowerCase() === username.toLowerCase())) {
      return { ok: false, error: 'That username is already taken.' };
    }

    const token = makeToken_();
    const now = new Date().toISOString();
    sheet.appendRow([username, passwordHash.toLowerCase(), token, now, now]);
    return { ok: true, username, token };
  } finally {
    lock.releaseLock();
  }
}

function login_(p) {
  const username = cleanUsername_(p.username);
  const passwordHash = String(p.passwordHash || '').trim().toLowerCase();
  if (!username || !/^[a-f0-9]{64}$/.test(passwordHash)) {
    return { ok: false, error: 'Please enter a valid username and password.' };
  }

  const sheet = getSheet_(USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowUsername = String(values[i][0] || '');
    const rowHash = String(values[i][1] || '').toLowerCase();
    if (rowUsername.toLowerCase() === username.toLowerCase() && rowHash === passwordHash) {
      const token = makeToken_();
      sheet.getRange(i + 1, 3).setValue(token);
      sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      return { ok: true, username: rowUsername, token };
    }
  }
  return { ok: false, error: 'Incorrect username or password.' };
}

function session_(p) {
  const user = findUserByToken_(String(p.token || ''));
  if (!user) return { ok: false, error: 'Session expired.' };
  return { ok: true, username: user.username };
}

function logout_(p) {
  const token = String(p.token || '');
  const sheet = getSheet_(USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2] || '') === token) {
      sheet.getRange(i + 1, 3).setValue('');
      break;
    }
  }
  return { ok: true };
}

function allProgress_(p) {
  const user = requireUser_(p);
  if (!user.ok) return user;

  const subjectKeys = ['bio', 'chem', 'phy', 'ict'];
  const result = {};
  subjectKeys.forEach(k => {
    result[k] = { idx: 0, correct: 0, attempted: 0, answers: {} };
  });

  const sheet = getSheet_(PROGRESS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[0]) !== user.username) continue;
    const subject = String(row[1]);
    if (!result[subject]) continue;
    const index = Number(row[2]);
    const selectedRaw = row[3];
    const selected = selectedRaw === '' ? null : Number(selectedRaw);
    const correct = String(row[4]).toUpperCase() === 'TRUE' || row[4] === true;
    const skipped = String(row[5]).toUpperCase() === 'TRUE' || row[5] === true;
    result[subject].answers[String(index)] = {
      selected,
      correct,
      skipped,
      timestamp: row[6] ? new Date(row[6]).getTime() : Date.now()
    };
    result[subject].attempted += 1;
    if (correct) result[subject].correct += 1;
    result[subject].idx = Math.max(result[subject].idx, index + 1);
  }

  return { ok: true, progress: result };
}

function saveAnswer_(p) {
  const user = requireUser_(p);
  if (!user.ok) return user;

  const subject = String(p.subject || '');
  const questionIndex = Number(p.questionIndex);
  const selectedIndex = p.selectedIndex === '' || p.selectedIndex === 'null' ? null : Number(p.selectedIndex);
  const isCorrect = String(p.isCorrect).toLowerCase() === 'true';
  const skipped = String(p.skipped).toLowerCase() === 'true';

  if (!['bio', 'chem', 'phy', 'ict'].includes(subject) || !Number.isInteger(questionIndex) || questionIndex < 0) {
    return { ok: false, error: 'Invalid progress data.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet_(PROGRESS_SHEET);
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === user.username && String(values[i][1]) === subject && Number(values[i][2]) === questionIndex) {
        targetRow = i + 1;
        break;
      }
    }

    const row = [user.username, subject, questionIndex, selectedIndex === null ? '' : selectedIndex, isCorrect, skipped, new Date().toISOString()];
    if (targetRow === -1) sheet.appendRow(row);
    else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function resetSubject_(p) {
  const user = requireUser_(p);
  if (!user.ok) return user;
  const subject = String(p.subject || '');
  if (!['bio', 'chem', 'phy', 'ict'].includes(subject)) return { ok: false, error: 'Invalid subject.' };

  const sheet = getSheet_(PROGRESS_SHEET);
  const values = sheet.getDataRange().getValues();
  const keep = [values[0]];
  for (let i = 1; i < values.length; i++) {
    if (!(String(values[i][0]) === user.username && String(values[i][1]) === subject)) {
      keep.push(values[i]);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  return { ok: true };
}

function requireUser_(p) {
  const user = findUserByToken_(String(p.token || ''));
  return user ? { ok: true, username: user.username } : { ok: false, error: 'Please log in again.' };
}

function findUserByToken_(token) {
  if (!token || token.length < 20) return null;
  const sheet = getSheet_(USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2] || '') === token) {
      return { username: String(values[i][0]) };
    }
  }
  return null;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setupDatabase() first.');
  return sheet;
}

function getRows_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function cleanUsername_(value) {
  return String(value || '').trim();
}

function makeToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function safeMessage_(err) {
  return err && err.message ? err.message : String(err);
}

function respond_(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
