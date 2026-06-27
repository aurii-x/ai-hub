
// =============================================================================
// ClickUp Careerflow Sync  v3
//
// Reads Tab 1 of the Job Search Time Reconciliation CSV, fuzzy-matches each
// job against existing ClickUp tasks, then:
//   - UPDATES matched tasks' Start Date / Due Date (from Careerflow columns)
//   - CREATES a new task (status = NEW_TASK_STATUS) when no match is found
//   - Adds a tracked time entry sourced from Rize (falls back to Timely, then
//     Careerflow, if a job has no Rize match)
//   - Checks ClickUp LIVE before adding a time entry — skips if one already
//     exists on that task, so re-running never creates duplicates
//   - Logs 3 verification columns: what Rize's source data showed vs. what
//     ClickUp actually has stored, so you can audit the sync
//
// CSV COLUMNS EXPECTED (Tab 1 of Job_Search_Time_Reconciliation.xlsx):
//   0  Job Title
//   1  Company
//   2  Description
//   3  Careerflow Start Date     (yyyy-mm-dd or m/d/yy)
//   4  Careerflow End Date
//   5  Careerflow Start Time     (HH:MM:SS)
//   6  Careerflow End Time
//   7  Timely Start Date
//   8  Timely End Date
//   9  Timely Start Time
//   10 Timely End Time
//   11 Rize Start Date
//   12 Rize End Date
//   13 Rize Start Time
//   14 Rize End Time
//
// ClickUp task name pattern: "Apply to [Job Title] at [Company]"
// initialization is defined in "Clickup_API_Setup.gs"

// =============================================================================

/** Run once to find the exact status names available in your list. */
function getListStatuses() {
  const p      = PropertiesService.getScriptProperties();
  const token  = p.getProperty('CLICKUP_TOKEN');
  const listId = p.getProperty('CLICKUP_LIST_ID');
  const r = apiGet(`https://api.clickup.com/api/v2/list/${listId}`, token);
  if (!r) { Logger.log('Failed — check token/list ID.'); return; }
  Logger.log('Statuses available in this list:');
  r.statuses.forEach(s => Logger.log(`  "${s.status}"  (type: ${s.type})`));
}

// ─── MAIN ────────────────────────────────────────────────────────────────────
function syncAll() {
  const t0     = Date.now();
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('CLICKUP_TOKEN');
  const teamId = props.getProperty('CLICKUP_TEAM_ID');
  const listId = props.getProperty('CLICKUP_LIST_ID');
  if (!token || !teamId || !listId) { notify('❌ Run setup() first.'); return; }

  let startRow = parseInt(props.getProperty('SYNC_START_ROW') || '0', 10);
  Logger.log('▶ Resuming from row ' + startRow);

  const rows = loadCsv();
  if (!rows) { notify('❌ CSV not found. Check CSV_FILE_NAME.'); return; }
  Logger.log('CSV rows: ' + rows.length);

  if (startRow >= rows.length) {
    props.deleteProperty('SYNC_START_ROW');
    notify('✅ Sync complete — all rows processed.');
    return;
  }

  const existingTasks = fetchAllTasksFromList(listId, token);
  if (!existingTasks) { notify('❌ Could not fetch ClickUp tasks.'); return; }
  Logger.log('Existing ClickUp tasks: ' + existingTasks.length);

  const log = getOrCreateLogSheet();
  const logIndex = loadLogIndex(log); // key -> {rowNum, datesOk, timeOk}
  Logger.log('Existing log entries: ' + Object.keys(logIndex).length);

  let processed = 0, updated = 0, created = 0, verifiedOnly = 0, errors = 0;

  for (let i = startRow; i < rows.length; i++) {
    if (Date.now() - t0 > MAX_RUNTIME_MS) {
      props.setProperty('SYNC_START_ROW', String(i));
      notify('⏸ Paused at row ' + i + ' of ' + rows.length + '. Run syncAll() again to continue.');
      return;
    }

    const row = rows[i];
    const jobTitle  = (row[0]  || '').trim();
    const company   = (row[1]  || '').trim();
    const descNote  = (row[2]  || '').trim();

    const cfStartDate = (row[3]  || '').trim();
    const cfEndDate   = (row[4]  || '').trim();
    const cfStartTime = (row[5]  || '').trim();
    const cfEndTime   = (row[6]  || '').trim();

    const tlyStartDate = (row[7]  || '').trim();
    const tlyEndDate   = (row[8]  || '').trim();
    const tlyStartTime = (row[9]  || '').trim();
    const tlyEndTime   = (row[10] || '').trim();

    const rizeStartDate = (row[11] || '').trim();
    const rizeEndDate   = (row[12] || '').trim();
    const rizeStartTime = (row[13] || '').trim();
    const rizeEndTime   = (row[14] || '').trim();

    if (!jobTitle && !company) { processed++; continue; }

    const rowKey = `${jobTitle}|${company}`;
    const existing = logIndex[rowKey];
    const taskName = buildTaskName(jobTitle, company);

    // ── Row already fully done: verify-only pass, no task/time-entry writes ──
    if (existing && existing.datesOk && existing.timeOk) {
      const m = findBestMatch(taskName, existingTasks); // CPU-only, no API call
      let vClickupStartMs = null, vClickupEndMs = null;
      if (m && m.score >= MATCH_THRESHOLD) {
        const entries = getTimeEntriesForTask(m.task.id, teamId, token); // 1 API call
        if (entries.length > 0) {
          vClickupStartMs = Number(entries[0].start);
          vClickupEndMs   = Number(entries[0].end);
        }
        Utilities.sleep(150);
      }
      const vRizeStartStr = formatTimeForDisplay(parseDateTimeToMs(rizeStartDate, rizeStartTime));
      const vRizeEndStr   = formatTimeForDisplay(parseDateTimeToMs(rizeEndDate,   rizeEndTime));
      const vClickupStr   = (vClickupStartMs && vClickupEndMs)
        ? `${formatTimeForDisplay(vClickupStartMs)} → ${formatTimeForDisplay(vClickupEndMs)}`
        : '(no time entry found on task)';
      updateLogVerificationCols(log, existing.rowNum, vRizeStartStr, vRizeEndStr, vClickupStr);
      verifiedOnly++;
      processed++;
      continue;
    }

    // Task Start/Due dates always come from Careerflow
    const cfStartMs = parseDateTimeToMs(cfStartDate, cfStartTime);
    const cfEndMs   = parseDateTimeToMs(cfEndDate,   cfEndTime);
    const cfDatesValid = isValidMs(cfStartMs) && isValidMs(cfEndMs);

    // Time-entry source: prefer Rize, fall back to Timely, then Careerflow
    let teStartMs = parseDateTimeToMs(rizeStartDate, rizeStartTime);
    let teEndMs   = parseDateTimeToMs(rizeEndDate,   rizeEndTime);
    let teSource  = 'Rize';
    if (!(isValidMs(teStartMs) && isValidMs(teEndMs) && teEndMs > teStartMs)) {
      teStartMs = parseDateTimeToMs(tlyStartDate, tlyStartTime);
      teEndMs   = parseDateTimeToMs(tlyEndDate,   tlyEndTime);
      teSource  = 'Timely';
      if (!(isValidMs(teStartMs) && isValidMs(teEndMs) && teEndMs > teStartMs)) {
        teStartMs = cfStartMs;
        teEndMs   = cfEndMs;
        teSource  = 'Careerflow';
      }
    }
    const teValid = isValidMs(teStartMs) && isValidMs(teEndMs) && teEndMs > teStartMs;

    const match   = findBestMatch(taskName, existingTasks);
    const isMatch = match && match.score >= MATCH_THRESHOLD;

    let taskId = null, action = '', datesOk = false, notes = [];

    // ── Match existing task: update its dates ──────────────────────────────
    if (isMatch) {
      taskId = match.task.id;
      action = 'UPDATED';
      if (cfDatesValid) {
        const r = updateTaskDates(taskId, cfStartMs, cfEndMs, token);
        datesOk = r.ok;
        if (!r.ok) notes.push('Date update failed: ' + r.error);
        Utilities.sleep(200);
      } else {
        notes.push('Dates skipped (unparseable Careerflow date/time)');
      }

    // ── No match: create a new task ─────────────────────────────────────────
    } else {
      action = 'CREATED';
      const cr = createTask({
        listId, name: taskName, status: NEW_TASK_STATUS, priority: TASK_PRIORITY,
        startMs: cfDatesValid ? cfStartMs : null,
        endMs:   cfDatesValid ? cfEndMs   : null,
        description: descNote,
        token,
      });
      if (cr.ok) {
        taskId  = cr.taskId;
        datesOk = cfDatesValid;
        existingTasks.push({ id: taskId, name: taskName }); // avoid dup creation later in this run
        created++;
      } else {
        notes.push('Task creation failed: ' + cr.error);
        errors++;
      }
      Utilities.sleep(300);
    }

    // ── Time entry: check ClickUp live first, skip if already tracked ──────
    let timeEntryOk = false;
    let clickupStartMs = null, clickupEndMs = null;

    if (taskId && teValid) {
      const existingEntries = getTimeEntriesForTask(taskId, teamId, token);
      if (existingEntries.length > 0) {
        const e = existingEntries[0];
        clickupStartMs = Number(e.start);
        clickupEndMs   = Number(e.end);
        timeEntryOk = true;
        notes.push(`Time entry already existed (not duplicated)`);
      } else {
        const r = createTimeEntry(taskId, teamId, teStartMs, teEndMs, token);
        if (r.ok) {
          timeEntryOk    = true;
          clickupStartMs = r.start;
          clickupEndMs   = r.end;
          notes.push(`Time entry created from ${teSource}`);
        } else {
          notes.push('Time entry failed: ' + r.error);
        }
      }
      Utilities.sleep(200);
    } else if (taskId) {
      notes.push('Time entry skipped: no valid date/time in Rize, Timely, or Careerflow');
    }

    if (isMatch && datesOk) updated++;

    const scoreStr  = match ? (match.score * 100).toFixed(1) + '%' : '—';
    const matchName = isMatch ? match.task.name : '(new task)';

    // Verification columns — what Rize's source data showed vs. what
    // ClickUp now actually has on record for the time entry
    const rizeStartStr = formatTimeForDisplay(parseDateTimeToMs(rizeStartDate, rizeStartTime));
    const rizeEndStr   = formatTimeForDisplay(parseDateTimeToMs(rizeEndDate,   rizeEndTime));
    const clickupStr   = (clickupStartMs && clickupEndMs)
      ? `${formatTimeForDisplay(clickupStartMs)} → ${formatTimeForDisplay(clickupEndMs)}`
      : '';

    upsertLogRow(log, existing, [
      new Date(), jobTitle, company, taskName,
      action, matchName, scoreStr,
      datesOk     ? '✅' : '❌',
      timeEntryOk ? '✅' : (taskId ? '❌' : '—'),
      notes.join(' | '),
      rizeStartStr, rizeEndStr, clickupStr,
    ]);

    processed++;
    if (processed % BATCH_SIZE === 0) {
      props.setProperty('SYNC_START_ROW', String(i + 1));
    }
  }

  props.deleteProperty('SYNC_START_ROW');
  const summary = `✅ Complete — Processed: ${processed}, Updated: ${updated}, Created: ${created}, `
                 + `Verified only (already done): ${verifiedOnly}, Errors: ${errors}`;
  Logger.log(summary);
  notify(summary);
}

function resetSync() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_START_ROW');
  notify('Reset. Next syncAll() will start from row 0.');
}

function openLogSheet() {
  Logger.log('Log: ' + getOrCreateLogSheet().getParent().getUrl());
}

// ─── CLICKUP API ─────────────────────────────────────────────────────────────

function fetchAllTasksFromList(listId, token) {
  const tasks = [];
  let page = 0;
  while (true) {
    const url = `https://api.clickup.com/api/v2/list/${listId}/task`
      + `?page=${page}&include_closed=true&subtasks=true&order_by=created`;
    const r = apiGet(url, token);
    if (!r || !r.tasks) return null;
    r.tasks.forEach(t => tasks.push({ id: t.id, name: t.name || '' }));
    if (r.last_page || r.tasks.length === 0) break;
    page++;
    Utilities.sleep(200);
  }
  return tasks;
}

function createTask({ listId, name, status, priority, startMs, endMs, description, token }) {
  const url  = `https://api.clickup.com/api/v2/list/${listId}/task`;
  const body = { name, status, priority, notify_all: false };
  if (description) body.description = description;        // task-level description — fine on all plans
  if (isValidMs(startMs)) { body.start_date = startMs; body.start_date_time = true; }
  if (isValidMs(endMs))   { body.due_date   = endMs;   body.due_date_time   = true; }

  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const code = r.getResponseCode();
    if (code < 200 || code >= 300) {
      return { ok: false, error: `HTTP ${code}: ${r.getContentText().substring(0, 200)}` };
    }
    const data = JSON.parse(r.getContentText());
    return { ok: true, taskId: data.id };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function updateTaskDates(taskId, startMs, dueMs, token) {
  const body = {};
  if (isValidMs(startMs)) { body.start_date = startMs; body.start_date_time = true; }
  if (isValidMs(dueMs))   { body.due_date   = dueMs;   body.due_date_time   = true; }
  return apiPut(`https://api.clickup.com/api/v2/task/${taskId}`, body, token);
}

/**
 * Get existing time entries for a specific task.
 * Returns an array (empty if none / on failure).
 */
function getTimeEntriesForTask(taskId, teamId, token) {
  const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries?task_id=${taskId}`;
  const r = apiGet(url, token);
  if (!r || !r.data) return [];
  return r.data;
}

/**
 * Create a manual time entry. Unlimited-plan safe: only tid/start/stop are
 * sent — "description", "tags", and "billable" are advanced features that
 * require the Business plan or higher and will error/be ignored on Unlimited.
 */
function createTimeEntry(taskId, teamId, startMs, stopMs, token) {
  const url  = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`;
  const body = { tid: taskId, start: startMs, stop: stopMs };
  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const code = r.getResponseCode();
    if (code < 200 || code >= 300) {
      return { ok: false, error: `HTTP ${code}: ${r.getContentText().substring(0, 200)}` };
    }
    const data  = JSON.parse(r.getContentText());
    const entry = data.data || data;
    return {
      ok: true,
      start: entry.start ? Number(entry.start) : startMs,
      end:   entry.end   ? Number(entry.end)   : stopMs,
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// ─── FUZZY MATCHING ──────────────────────────────────────────────────────────

function buildTaskName(jobTitle, company) {
  if (jobTitle && company) return `Apply to ${jobTitle} at ${company}`;
  if (jobTitle) return `Apply to ${jobTitle}`;
  return company;
}

function findBestMatch(searchName, tasks) {
  const sNorm = normalize(searchName);
  let best = null;
  for (const task of tasks) {
    const tNorm = normalize(task.name);
    if (!tNorm) continue;
    if (sNorm === tNorm) return { task, score: 1.0 };
    let score = jaroWinkler(sNorm, tNorm);
    const titleOnly = extractTitlePart(sNorm);
    if (titleOnly) score = Math.max(score, jaroWinkler(titleOnly, tNorm));
    if (sNorm.includes(tNorm) || tNorm.includes(sNorm)) {
      const ratio = Math.min(sNorm.length, tNorm.length) / Math.max(sNorm.length, tNorm.length);
      score = Math.max(score, 0.85 + ratio * 0.10);
    }
    if (!best || score > best.score) best = { task, score };
  }
  return best;
}

function normalize(str) {
  return (str || '').toLowerCase().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim();
}

function extractTitlePart(normalized) {
  const m = normalized.match(/^apply to (.+?) at /);
  return m ? m[1] : null;
}

function jaroWinkler(s1, s2) {
  if (s1 === s2) return 1.0;
  if (!s1 || !s2) return 0.0;
  const l1 = s1.length, l2 = s2.length;
  const dist = Math.max(Math.floor(Math.max(l1, l2) / 2) - 1, 0);
  const m1 = new Array(l1).fill(false), m2 = new Array(l2).fill(false);
  let matches = 0, trans = 0;
  for (let i = 0; i < l1; i++) {
    for (let j = Math.max(0, i - dist); j < Math.min(i + dist + 1, l2); j++) {
      if (m2[j] || s1[i] !== s2[j]) continue;
      m1[i] = m2[j] = true; matches++; break;
    }
  }
  if (!matches) return 0.0;
  let k = 0;
  for (let i = 0; i < l1; i++) {
    if (!m1[i]) continue;
    while (!m2[k]) k++;
    if (s1[i] !== s2[k]) trans++;
    k++;
  }
  const jaro = (matches/l1 + matches/l2 + (matches - trans/2)/matches) / 3;
  let prefix = 0;
  for (let i = 0; i < Math.min(4, l1, l2); i++) {
    if (s1[i] === s2[i]) prefix++; else break;
  }
  return jaro + prefix * 0.1 * (1 - jaro);
}

// ─── DATE / TIME PARSING ─────────────────────────────────────────────────────

function isValidMs(ms) {
  return ms !== null && ms !== undefined && !isNaN(ms) && ms > 0;
}

/**
 * Parse "yyyy-mm-dd" OR "m/d/yy" / "m/d/yyyy" + "HH:MM:SS" → Unix ms (UTC),
 * adjusted from Eastern Time (Raleigh, NC).
 * 2-digit years are expanded to 20xx (fixes the m/dd/yy "year 26" bug).
 */
function parseDateTimeToMs(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  dateStr = dateStr.trim(); timeStr = timeStr.trim();
  let year, month, day;
  if (dateStr.includes('-')) {
    [year, month, day] = dateStr.split('-').map(Number);
  } else if (dateStr.includes('/')) {
    const p = dateStr.split('/').map(Number);
    if (p[0] > 31) [year, month, day] = p; else [month, day, year] = p;
    if (year < 100) year += 2000;          // ← fixes m/dd/yy (e.g. 6/15/26 → 2026)
  } else return null;

  const [hour = 0, min = 0, sec = 0] = timeStr.split(':').map(Number);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;

  const offsetH = isEasternDst(year, month, day) ? 4 : 5;
  return Date.UTC(year, month - 1, day, hour + offsetH, min, sec);
}

function isEasternDst(y, m, d) {
  if (m < 3 || m > 11) return false;
  if (m > 3 && m < 11) return true;
  if (m === 3)  return d >= nthWeekday(y, 3, 0, 2);
  if (m === 11) return d < nthWeekday(y, 11, 0, 1);
}

function nthWeekday(year, month, dow, nth) {
  let count = 0;
  for (let d = 1; d <= 31; d++) {
    const dt = new Date(year, month - 1, d);
    if (dt.getMonth() !== month - 1) break;
    if (dt.getDay() === dow && ++count === nth) return d;
  }
}

/** Format a ms epoch timestamp as a human-readable Eastern Time string. */
function formatTimeForDisplay(ms) {
  if (!isValidMs(ms)) return '';
  return Utilities.formatDate(new Date(ms), 'America/New_York', 'M/d/yyyy h:mm a');
}

// ─── CSV LOADING ─────────────────────────────────────────────────────────────

function loadCsv() {
  try {
    let content = null;
    try {
      const f = getFolder().getFilesByName(CSV_FILE_NAME);
      if (f.hasNext()) content = f.next().getBlob().getDataAsString('UTF-8');
    } catch (_) {}
    if (!content) {
      const f = DriveApp.getFilesByName(CSV_FILE_NAME);
      if (!f.hasNext()) return null;
      content = f.next().getBlob().getDataAsString('UTF-8');
    }
    return parseCsv(content).slice(1);
  } catch (e) { Logger.log('CSV load error: ' + e); return null; }
}

function parseCsv(text) {
  const rows = []; let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], n = text[i+1];
    if (inQ) {
      if (c === '"' && n === '"') { field += '"'; i++; }
      else if (c === '"')         { inQ = false; }
      else                        { field += c; }
    } else {
      if      (c === '"')                          { inQ = true; }
      else if (c === ',')                          { row.push(field); field = ''; }
      else if (c === '\n' || (c==='\r'&&n==='\n')) {
        if (c==='\r') i++;
        row.push(field); field = ''; rows.push(row); row = [];
      } else { field += c; }
    }
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

// ─── DRIVE / LOG HELPERS ─────────────────────────────────────────────────────

function getFolder() {
  let f = DriveApp.getRootFolder();
  for (const n of DRIVE_FOLDER_PATH) {
    const s = f.getFoldersByName(n);
    f = s.hasNext() ? s.next() : f.createFolder(n);
  }
  return f;
}

const LOG_HEADERS = [
  'Timestamp', 'Job Title', 'Company', 'Task Name Built',
  'Action', 'Matched / Created Task', 'Similarity',
  'Dates Updated', 'Time Entry', 'Notes',
  'Rize Start Time', 'Rize End Time', 'ClickUp Time Entry (Start → End)',
];

function getOrCreateLogSheet() {
  const folder = getFolder();
  const files  = folder.getFilesByName(LOG_SHEET_NAME);
  let ss, ws;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
    ws = ss.getSheets()[0];
    const lastCol = ws.getLastColumn();
    if (lastCol < LOG_HEADERS.length) {
      // extend an existing log with the new verification columns
      const range = ws.getRange(1, lastCol + 1, 1, LOG_HEADERS.length - lastCol);
      range.setValues([LOG_HEADERS.slice(lastCol)]);
      range.setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    }
  } else {
    ss = SpreadsheetApp.create(LOG_SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    ws = ss.getActiveSheet();
    ws.setName('Sync Log');
    ws.appendRow(LOG_HEADERS);
    ws.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
  }
  return ws;
}

function logRow(sheet, values) {
  try { sheet.appendRow(values); } catch (e) { Logger.log('Log error: ' + e); }
}

/**
 * Reads the log sheet into an index: "JobTitle|Company" -> { rowNum, datesOk, timeOk }.
 * rowNum is the 1-indexed sheet row, used to update that row in place rather
 * than appending a duplicate. If the same key appears more than once in the
 * log, the LAST occurrence wins.
 */
function loadLogIndex(sheet) {
  const index = {};
  try {
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) { // skip header
      const title   = (data[r][1] || '').toString().trim();
      const company  = (data[r][2] || '').toString().trim();
      if (!title && !company) continue;
      const key = `${title}|${company}`;
      index[key] = {
        rowNum:  r + 1, // sheet rows are 1-indexed
        datesOk: (data[r][7] || '').toString().trim() === '✅',
        timeOk:  (data[r][8] || '').toString().trim() === '✅',
      };
    }
  } catch (e) { Logger.log('Could not load log index: ' + e); }
  return index;
}

/**
 * Writes only the 3 verification columns (K, L, M) into an existing log row,
 * without touching the rest of that row. Used for the "already done, verify
 * only" pass so re-runs refresh the comparison data without disturbing
 * anything else.
 */
function updateLogVerificationCols(sheet, rowNum, rizeStartStr, rizeEndStr, clickupStr) {
  try {
    sheet.getRange(rowNum, 11, 1, 3).setValues([[rizeStartStr, rizeEndStr, clickupStr]]);
  } catch (e) { Logger.log('Verification column update error: ' + e); }
}

/**
 * Writes a full row of values — updates the existing row in place if
 * `existing` (from loadLogIndex) has a rowNum, otherwise appends a new row.
 * Prevents duplicate log lines for the same job across multiple runs.
 */
function upsertLogRow(sheet, existing, values) {
  try {
    if (existing && existing.rowNum) {
      sheet.getRange(existing.rowNum, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  } catch (e) { Logger.log('Log write error: ' + e); }
}

// ─── HTTP HELPERS ────────────────────────────────────────────────────────────

function apiGet(url, token) {
  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: token, 'Content-Type': 'application/json' },
      muteHttpExceptions: true,
    });
    if (r.getResponseCode() !== 200) {
      Logger.log('GET ' + r.getResponseCode() + ': ' + url.split('?')[0]);
      return null;
    }
    return JSON.parse(r.getContentText());
  } catch (e) { Logger.log('GET error: ' + e); return null; }
}

function apiPut(url, body, token) {
  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: { Authorization: token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const c = r.getResponseCode();
    return c >= 200 && c < 300
      ? { ok: true }
      : { ok: false, error: `HTTP ${c}: ${r.getContentText().substring(0, 150)}` };
  } catch (e) { return { ok: false, error: String(e) }; }
}

function apiPost(url, body, token) {
  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const c = r.getResponseCode();
    return c >= 200 && c < 300
      ? { ok: true }
      : { ok: false, error: `HTTP ${c}: ${r.getContentText().substring(0, 150)}` };
  } catch (e) { return { ok: false, error: String(e) }; }
}

function notify(msg) {
  Logger.log(msg);
  if (!SILENT_MODE) { try { SpreadsheetApp.getUi().alert(msg); } catch (_) {} }
}

// ─── DRY RUN ─────────────────────────────────────────────────────────────────

function dryRun() {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('CLICKUP_TOKEN');
  const listId = props.getProperty('CLICKUP_LIST_ID');
  if (!token || !listId) { Logger.log('Run setup() first.'); return; }

  const rows  = loadCsv();
  if (!rows)  { Logger.log('CSV not found.'); return; }
  const tasks = fetchAllTasksFromList(listId, token);
  if (!tasks) { Logger.log('Could not fetch tasks.'); return; }

  Logger.log(`CSV rows: ${rows.length}  |  ClickUp tasks: ${tasks.length}`);
  Logger.log('─'.repeat(90));

  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row    = rows[i];
    const title  = (row[0]||'').trim(), co = (row[1]||'').trim();
    if (!title && !co) continue;
    const name    = buildTaskName(title, co);
    const match   = findBestMatch(name, tasks);
    const isMatch = match && match.score >= MATCH_THRESHOLD;
    const action  = isMatch ? '🔄 UPDATE' : '🆕 CREATE';
    const score   = match ? (match.score * 100).toFixed(1) + '%' : '—';
    Logger.log(`[${String(i+1).padStart(3,'0')}] ${action} | ${score.padStart(6)} | "${name}"`);
    if (isMatch) Logger.log(`         → matched: "${match.task.name}"`);
  }
}
function checkProgress() {
  Logger.log('SYNC_START_ROW = ' + PropertiesService.getScriptProperties().getProperty('SYNC_START_ROW'));
}
