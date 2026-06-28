// =============================================================================
// ClickUp Dashboard Export  v1.5
//
// CHANGELOG
//   v1.5 - Full regeneration consolidating all prior changes. Version number
//           now correct in file header. No functional changes from v1.4.
//
//   v1.4 - Fixed fetchAllTimeEntries and fetchTimeEntriesForTaskIds:
//           ClickUp time entries endpoint ignores the page parameter and
//           returns all entries in the date range in a single call. The
//           prior while(page++) loop fetched the same data repeatedly until
//           the time budget ran out. Fixed with monthly date chunking —
//           one API call per month. Added buildMonthlyChunks() helper.
//
//   v1.3 - runFullExport() switched from clear-and-rewrite to upsert mode.
//           Re-runs now show accurate Added/Updated counts in Sync Log.
//           Row index built once upfront — no data loss on timeout.
//
//   v1.2 - Renamed conflicting functions to db-prefix (getOrCreateWorkbook →
//           dbGetOrCreateWorkbook, fmtTs → dbFmtTs, fmtDur → dbFmtDur,
//           setup → dbSetup, checkSetup → dbCheckSetup) to prevent global
//           namespace collisions when both scripts share an Apps Script project.
//
//   v1.1 - Credential keys renamed CLICKUP_TOKEN / CLICKUP_TEAM_ID to match
//           ClickUp_Rize_PipelineV2.gs. EXPORT_FOLDER_PATH updated.
//           Workbook created FIRST in runFullExport() before any API calls.
//           Tasks written page-by-page as fetched. Added patchTimeEntryColumns()
//           and runRefreshTimeEntries() for recovery from timeouts.
//           Verbose progress logging. Added findWorkbook() and listDriveFolder().
//
// RUN ORDER (first time):
//   1. dbSetup()                → save credentials
//   2. dbCheckSetup()           → confirm they saved
//   3. runFullExport()          → full pull from FULL_EXPORT_START_DATE
//      if it times out during time entries:
//   4. runRefreshTimeEntries()  → patch time data without re-fetching tasks
//   5. setupHourlyTrigger()     → automates incremental syncs going forward
// =============================================================================

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const EXPORT_FOLDER_PATH     = ['AppData', '3.1 clickup-automation', 'dashboard'];
const EXPORT_WORKBOOK_NAME   = 'ClickUp Dashboard Data';
const TASKS_SHEET_NAME       = 'Tasks';
const SYNC_LOG_SHEET_NAME    = 'Sync Log';
const FULL_EXPORT_START_DATE = '2025-01-01';
const INCREMENTAL_BUFFER_HRS = 2;
//const MAX_RUNTIME_MS         = 330000;
const TIME_BUFFER_MS         = 25000;
const TIMEZONE               = 'America/New_York';
const PAGE_SIZE              = 100;

// ─── COLUMN SCHEMA ───────────────────────────────────────────────────────────
// To add a field: append to HEADERS and populate it in buildTaskRow().
const HEADERS = [
  'task_id',             // A
  'task_name',           // B
  'task_start_date',     // C
  'task_end_date',       // D
  'tags',                // E  comma-separated tag names
  'space_id',            // F
  'folder_id',           // G
  'list_id',             // H
  'folder_name',         // I
  'list_name',           // J
  'status',              // K
  'tracked_time_start',  // L  earliest entry start (Eastern)
  'tracked_time_end',    // M  latest entry end (Eastern)
  'tracked_time_ms',     // N  total duration in milliseconds
  'tracked_time',        // O  human-readable e.g. "3h 7m"
  'time_entry_count',    // P
  'task_created',        // Q
  'task_last_modified',  // R
];

// ─── SETUP ───────────────────────────────────────────────────────────────────

/**
 * Save ClickUp credentials to Script Properties.
 * Uses the same key names as ClickUp_Rize_PipelineV2.gs.
 * Run once, then delete the values from the function body.
 */
function dbSetup() {
  PropertiesService.getScriptProperties().setProperties({
    CLICKUP_TOKEN:   'pk_216003478_5N2AE9LICIRWR32210VR9PO1J2OKS4U7',
    CLICKUP_TEAM_ID: '90141302224',
  });
  Logger.log('✅ Credentials saved. Clear the values from dbSetup() now.');
}

function dbCheckSetup() {
  const p = PropertiesService.getScriptProperties();
  const t = p.getProperty('CLICKUP_TOKEN');
  const i = p.getProperty('CLICKUP_TEAM_ID');
  Logger.log('CLICKUP_TOKEN:   ' + (t ? t.substring(0, 8) + '…' : 'NOT SET'));
  Logger.log('CLICKUP_TEAM_ID: ' + (i || 'NOT SET'));
}

// ─── TRIGGERS ────────────────────────────────────────────────────────────────

function setupHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runIncrementalSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runIncrementalSync').timeBased().everyHours(1).create();
  Logger.log('✅ Hourly trigger created for runIncrementalSync().');
}

function removeHourlyTrigger() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runIncrementalSync') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log(`Removed ${n} trigger(s).`);
}

// ─── MAIN ENTRY POINTS ───────────────────────────────────────────────────────

/**
 * Full export — run ONCE on first setup.
 *
 * Phase 1: Creates the workbook immediately — file exists in Drive even if
 *           the script times out later.
 * Phase 2: Fetches tasks page-by-page, upserts each page as it arrives —
 *           a timeout mid-fetch preserves all pages already written.
 * Phase 3: Fetches time entries (monthly chunks) and patches cols L–P.
 *           If Phase 3 times out, run runRefreshTimeEntries() to complete
 *           without re-fetching tasks.
 */
function runFullExport() {
  const t0 = Date.now();
  const { token, teamId } = getCredentials();
  if (!token) return;

  const props   = PropertiesService.getScriptProperties();
  const lastSync = props.getProperty('LAST_SYNC_TS');
  if (lastSync) {
    Logger.log('⚠️ Full export has already run (LAST_SYNC_TS is set).');
    Logger.log('   Use runIncrementalSync() or setupHourlyTrigger() instead.');
    Logger.log('   To force a clean re-export, run clearExportState() first.');
    return;
  }

  Logger.log(`🚀 Full export started — window: ${FULL_EXPORT_START_DATE} → now`);
  const startMs = new Date(FULL_EXPORT_START_DATE + 'T00:00:00Z').getTime();
  const nowMs   = Date.now();

  // ── Phase 1: Create workbook FIRST ────────────────────────────────────────
  Logger.log('📂 Creating / opening workbook...');
  const ss    = dbGetOrCreateWorkbook();
  const sheet = getOrCreateTasksSheet(ss, /*clear=*/false);
  Logger.log(`📂 Workbook ready: "${EXPORT_WORKBOOK_NAME}"`);
  Logger.log(`   URL: ${ss.getUrl()}`);

  // Build row index for upsert
  Logger.log('🗂️ Building row index from existing sheet data...');
  const index = buildTaskIdIndex(sheet);
  Logger.log(`   Existing rows in sheet: ${Object.keys(index).length}`);

  // ── Phase 2: Fetch tasks page-by-page, upsert immediately ────────────────
  Logger.log('📡 Fetching tasks from ClickUp (upserting each page as it arrives)...');
  let page    = 0;
  let total   = 0;
  const stats = { added: 0, updated: 0 };

  while (true) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit at page ${page} — ${total} tasks processed, data is safe.`);
      Logger.log('   Re-run runFullExport() to continue (clears LAST_SYNC_TS guard first).');
      return;
    }

    Logger.log(`  Fetching page ${page}...`);
    const url = `https://api.clickup.com/api/v2/team/${teamId}/task`
      + `?page=${page}&include_closed=true&subtasks=true`
      + `&order_by=updated&reverse=true&date_updated_gt=${startMs}`;

    const res = cuGet(url, token);
    if (!res || !res.tasks) { Logger.log(`⚠️ Task fetch failed on page ${page}.`); break; }

    Logger.log(`  Page ${page}: ${res.tasks.length} tasks`);

    for (const task of res.tasks) {
      const row    = buildTaskRow(task, null);
      const rowNum = index[task.id];
      if (rowNum) {
        sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([row]);
        stats.updated++;
      } else {
        const nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1, 1, HEADERS.length).setValues([row]);
        index[task.id] = nextRow;
        stats.added++;
      }
      total++;
    }

    SpreadsheetApp.flush();
    Logger.log(`  ✅ Page ${page} upserted — added: ${stats.added}, updated: ${stats.updated}`);

    if (res.tasks.length < PAGE_SIZE || res.last_page) { Logger.log('  Last page reached.'); break; }
    page++;
    Utilities.sleep(200);
  }

  Logger.log(`\n✅ Task fetch complete — ${total} tasks (added: ${stats.added}, updated: ${stats.updated}).`);

  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⚠️ No time budget left for time entries. Run runRefreshTimeEntries().');
    return;
  }

  // ── Phase 3: Fetch time entries by monthly chunks and patch sheet ─────────
  Logger.log('\n⏱️ Fetching time entries (monthly chunks — this may take a few minutes)...');
  const timeMap = fetchAllTimeEntries(token, teamId, startMs, nowMs, t0);
  Logger.log(`✅ Time entry fetch complete — ${Object.keys(timeMap).length} tasks have tracked time.`);

  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⚠️ Time budget hit — patching incomplete. Run runRefreshTimeEntries() to finish.');
    return;
  }

  if (Object.keys(timeMap).length > 0) {
    Logger.log('📝 Patching time entry columns into sheet...');
    patchTimeEntryColumns(sheet, timeMap);
  }

  props.setProperty('LAST_SYNC_TS', String(nowMs));
  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  logSyncRun(ss, 'FULL EXPORT', total, elapsed, stats);
  Logger.log(`\n🏁 Full export complete — ${total} tasks in ${elapsed}s.`);
  Logger.log('   Next: run setupHourlyTrigger() to automate incremental syncs.');
}

/**
 * Recovery function — run when runFullExport() timed out during Phase 3.
 * Tasks are already in the sheet. Only fetches and patches time entry cols.
 */
function runRefreshTimeEntries() {
  const t0 = Date.now();
  const { token, teamId } = getCredentials();
  if (!token) return;

  const ss    = dbGetOrCreateWorkbook();
  const sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('❌ Tasks sheet is empty — run runFullExport() first.');
    return;
  }

  const taskCount = sheet.getLastRow() - 1;
  const startMs   = new Date(FULL_EXPORT_START_DATE + 'T00:00:00Z').getTime();
  const nowMs     = Date.now();

  Logger.log(`⏱️ Refreshing time entries for ${taskCount} tasks — ${FULL_EXPORT_START_DATE} → now`);
  const timeMap = fetchAllTimeEntries(token, teamId, startMs, nowMs, t0);
  Logger.log(`✅ Tasks with time entries: ${Object.keys(timeMap).length}`);

  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⚠️ Time budget hit — re-run runRefreshTimeEntries() to continue.');
    return;
  }

  Logger.log('📝 Patching time entry columns into sheet...');
  patchTimeEntryColumns(sheet, timeMap);

  PropertiesService.getScriptProperties().setProperty('LAST_SYNC_TS', String(nowMs));
  Logger.log(`🏁 Time entries refreshed in ${((Date.now() - t0) / 1000).toFixed(1)}s.`);
}

/**
 * Incremental sync — runs hourly via trigger.
 * Only fetches tasks modified since the last sync (with a 2-hour buffer).
 */
function runIncrementalSync() {
  const t0 = Date.now();
  const { token, teamId } = getCredentials();
  if (!token) return;

  const props      = PropertiesService.getScriptProperties();
  const lastSyncTs = Number(props.getProperty('LAST_SYNC_TS') || '0');
  if (!lastSyncTs) {
    Logger.log('❌ No LAST_SYNC_TS found — run runFullExport() first.');
    return;
  }

  const bufferMs    = INCREMENTAL_BUFFER_HRS * 60 * 60 * 1000;
  const fetchFromMs = lastSyncTs - bufferMs;
  const nowMs       = Date.now();

  Logger.log(`🔄 Incremental sync — from: ${dbFmtTs(fetchFromMs)}`);

  Logger.log('📡 Fetching updated tasks...');
  const tasks = fetchAllTasks(token, teamId, fetchFromMs, t0);
  Logger.log(`   Tasks updated since last sync: ${tasks.length}`);

  if (tasks.length === 0) {
    props.setProperty('LAST_SYNC_TS', String(nowMs));
    Logger.log('✅ Nothing new. Anchor advanced.');
    return;
  }

  Logger.log('⏱️ Fetching time entries for updated tasks...');
  const taskIds = tasks.map(t => t.id);
  const timeMap = fetchTimeEntriesForTaskIds(token, teamId, taskIds, fetchFromMs, nowMs);
  Logger.log(`   Tasks with time entries: ${Object.keys(timeMap).length}`);

  Logger.log('📝 Upserting rows in sheet...');
  const ss    = dbGetOrCreateWorkbook();
  const sheet = getOrCreateTasksSheet(ss, /*clear=*/false);
  const stats = writeTaskRows(sheet, tasks, timeMap, /*upsert=*/true);

  props.setProperty('LAST_SYNC_TS', String(nowMs));
  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  logSyncRun(ss, 'INCREMENTAL', tasks.length, elapsed, stats);
  Logger.log(`🏁 Sync complete — updated: ${stats.updated}, new: ${stats.added}, time: ${elapsed}s.`);
}

/**
 * Clears the LAST_SYNC_TS anchor so runFullExport() can be re-run from scratch.
 * Use only when you want a complete reset of the export.
 */
function clearExportState() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_SYNC_TS');
  Logger.log('✅ Export state cleared. runFullExport() will do a clean re-export.');
  Logger.log('   Also consider clearing the Tasks sheet manually for a true clean slate.');
}

// ─── CLICKUP TASK FETCH ───────────────────────────────────────────────────────

function fetchAllTasks(token, teamId, updatedAfterMs, t0) {
  const allTasks = [];
  let page       = 0;

  while (true) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit fetching tasks at page ${page} (${allTasks.length} so far).`);
      break;
    }

    Logger.log(`  Fetching tasks page ${page}...`);
    const url = `https://api.clickup.com/api/v2/team/${teamId}/task`
      + `?page=${page}&include_closed=true&subtasks=true`
      + `&order_by=updated&reverse=true&date_updated_gt=${updatedAfterMs}`;

    const res = cuGet(url, token);
    if (!res || !res.tasks) { Logger.log(`⚠️ Task fetch failed on page ${page}.`); break; }

    Logger.log(`  Page ${page}: ${res.tasks.length} tasks`);
    res.tasks.forEach(t => allTasks.push(t));

    if (res.tasks.length < PAGE_SIZE || res.last_page) break;
    page++;
    Utilities.sleep(200);
  }

  return allTasks;
}

// ─── CLICKUP TIME ENTRY FETCH ────────────────────────────────────────────────

/**
 * Full export: fetch ALL time entries using monthly date chunks.
 * The ClickUp time entries endpoint ignores the page param — it returns all
 * entries in the date range in a single call. Monthly chunking handles large
 * date ranges without hitting response size limits.
 */
function fetchAllTimeEntries(token, teamId, startMs, endMs, t0) {
  const map    = {};
  let total    = 0;
  const chunks = buildMonthlyChunks(startMs, endMs);

  Logger.log(`  Fetching time entries in ${chunks.length} monthly chunks...`);

  for (let i = 0; i < chunks.length; i++) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit at chunk ${i + 1}/${chunks.length} (${total} entries so far).`);
      break;
    }

    const { chunkStart, chunkEnd, label } = chunks[i];
    const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`
      + `?start_date=${chunkStart}&end_date=${chunkEnd}`;

    const res = cuGet(url, token);
    if (!res) { Logger.log(`⚠️ Time entry fetch failed for chunk ${label}.`); continue; }

    const data = Array.isArray(res.data) ? res.data : [];
    Logger.log(`  ${label}: ${data.length} time entries`);
    data.forEach(e => mergeTimeEntry(map, e));
    total += data.length;
    Utilities.sleep(200);
  }

  Logger.log(`  Total: ${total} time entries across ${Object.keys(map).length} tasks`);
  return map;
}

function buildMonthlyChunks(startMs, endMs) {
  const chunks = [];
  let cursor   = new Date(startMs);
  cursor.setUTCDate(1);
  cursor.setUTCHours(0, 0, 0, 0);

  while (cursor.getTime() < endMs) {
    const chunkStart = cursor.getTime();
    const next       = new Date(cursor);
    next.setUTCMonth(next.getUTCMonth() + 1);
    const chunkEnd   = Math.min(next.getTime() - 1, endMs);
    const label      = Utilities.formatDate(new Date(chunkStart), 'UTC', 'yyyy-MM');
    chunks.push({ chunkStart, chunkEnd, label });
    cursor = next;
  }
  return chunks;
}

/**
 * Incremental: fetch time entries only for a specific window.
 * Filters results to only the task IDs we know changed.
 */
function fetchTimeEntriesForTaskIds(token, teamId, taskIds, startMs, endMs) {
  const taskIdSet = new Set(taskIds.map(String));
  const map       = {};

  Logger.log(`  Fetching time entries for ${taskIds.length} updated tasks...`);
  const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`
    + `?start_date=${startMs}&end_date=${endMs}`;

  const res = cuGet(url, token);
  if (!res) { Logger.log('⚠️ Time entry fetch failed.'); return map; }

  const data = Array.isArray(res.data) ? res.data : [];
  Logger.log(`  ${data.length} time entries in window — filtering to updated tasks...`);
  data.forEach(e => {
    if (e.task && taskIdSet.has(String(e.task.id))) mergeTimeEntry(map, e);
  });
  Logger.log(`  ${Object.keys(map).length} tasks matched with time entries.`);
  return map;
}

function mergeTimeEntry(map, entry) {
  if (!entry.task) return;
  const tid     = String(entry.task.id);
  const startMs = Number(entry.start);
  const endMs   = Number(entry.end);
  const durMs   = Number(entry.duration);
  if (!tid || !startMs || !endMs) return;

  if (!map[tid]) map[tid] = { totalMs: 0, minStart: startMs, maxEnd: endMs, count: 0 };
  map[tid].totalMs  += durMs;
  map[tid].minStart  = Math.min(map[tid].minStart, startMs);
  map[tid].maxEnd    = Math.max(map[tid].maxEnd, endMs);
  map[tid].count++;
}

// ─── PATCH TIME ENTRIES INTO EXISTING SHEET ROWS ─────────────────────────────

function patchTimeEntryColumns(sheet, timeMap) {
  const index   = buildTaskIdIndex(sheet);
  const L_COL   = 12;
  const COL_CNT = 5;
  let patched   = 0;
  const taskIds = Object.keys(timeMap);

  Logger.log(`  Patching ${taskIds.length} tasks with time entry data...`);

  for (const taskId of taskIds) {
    const rowNum = index[taskId];
    if (!rowNum) continue;
    const td = timeMap[taskId];
    sheet.getRange(rowNum, L_COL, 1, COL_CNT).setValues([[
      dbFmtTs(td.minStart),
      dbFmtTs(td.maxEnd),
      td.totalMs,
      dbFmtDur(td.totalMs),
      td.count,
    ]]);
    patched++;
    if (patched % 100 === 0) {
      SpreadsheetApp.flush();
      Logger.log(`  Patched ${patched} / ${taskIds.length}...`);
    }
  }

  SpreadsheetApp.flush();
  Logger.log(`✅ Time entries patched into ${patched} rows.`);
}

// ─── ROW BUILDING ────────────────────────────────────────────────────────────

function buildTaskRow(task, timeData) {
  const td = timeData || null;
  return [
    task.id,
    task.name || '',
    dbFmtTs(Number(task.start_date)),
    dbFmtTs(Number(task.due_date)),
    (task.tags || []).map(t => t.name).join(', '),
    task.space  ? task.space.id    : '',
    task.folder ? task.folder.id   : '',
    task.list   ? task.list.id     : '',
    task.folder ? task.folder.name : '',
    task.list   ? task.list.name   : '',
    task.status ? task.status.status : '',
    td ? dbFmtTs(td.minStart)  : '',
    td ? dbFmtTs(td.maxEnd)    : '',
    td ? td.totalMs            : '',
    td ? dbFmtDur(td.totalMs)  : '',
    td ? td.count              : 0,
    dbFmtTs(Number(task.date_created)),
    dbFmtTs(Number(task.date_updated)),
  ];
}

// ─── SHEET WRITE (incremental upsert) ────────────────────────────────────────

function writeTaskRows(sheet, tasks, timeMap, upsert) {
  const stats = { added: 0, updated: 0 };

  if (!upsert) {
    const matrix = tasks.map(t => buildTaskRow(t, timeMap[t.id]));
    if (matrix.length > 0) sheet.getRange(2, 1, matrix.length, HEADERS.length).setValues(matrix);
    stats.added = tasks.length;
    return stats;
  }

  const index = buildTaskIdIndex(sheet);
  for (const task of tasks) {
    const row    = buildTaskRow(task, timeMap[task.id]);
    const rowNum = index[task.id];
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([row]);
      stats.updated++;
    } else {
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, 1, HEADERS.length).setValues([row]);
      index[task.id] = nextRow;
      stats.added++;
    }
    if ((stats.added + stats.updated) % 50 === 0) {
      SpreadsheetApp.flush();
      Logger.log(`  Upserted ${stats.added + stats.updated} rows so far...`);
    }
  }

  SpreadsheetApp.flush();
  return stats;
}

function buildTaskIdIndex(sheet) {
  const index   = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return index;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  ids.forEach((row, i) => { if (row[0]) index[String(row[0])] = i + 2; });
  return index;
}

// ─── SHEET MANAGEMENT ────────────────────────────────────────────────────────

function dbGetOrCreateWorkbook() {
  const folder = resolveExportFolder();
  const files  = folder.getFilesByName(EXPORT_WORKBOOK_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
    Logger.log('📂 Opened existing workbook.');
  } else {
    ss = SpreadsheetApp.create(EXPORT_WORKBOOK_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    Logger.log(`✨ Created new workbook: "${EXPORT_WORKBOOK_NAME}"`);
  }
  return ss;
}

function getOrCreateTasksSheet(ss, clear) {
  let sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET_NAME);
  } else if (clear) {
    sheet.clearContents();
  }
  const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
  hRange.setValues([HEADERS])
        .setFontWeight('bold')
        .setBackground('#2F5597')
        .setFontColor('#FFFFFF')
        .setFontFamily('Arial')
        .setFontSize(10);
  sheet.setFrozenRows(1);
  const widths = [16,40,20,20,24,12,12,12,20,20,14,20,20,16,12,8,20,20];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w * 6));
  return sheet;
}

function logSyncRun(ss, mode, taskCount, elapsed, stats) {
  let log = ss.getSheetByName(SYNC_LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(SYNC_LOG_SHEET_NAME);
    log.appendRow(['Timestamp','Mode','Tasks Processed','Added','Updated','Duration (s)']);
    log.getRange(1,1,1,6).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
  }
  log.appendRow([
    new Date(), mode, taskCount,
    stats ? stats.added   : taskCount,
    stats ? stats.updated : 0,
    elapsed,
  ]);
}

function resolveExportFolder() {
  let f = DriveApp.getRootFolder();
  for (const name of EXPORT_FOLDER_PATH) {
    const s = f.getFoldersByName(name);
    f = s.hasNext() ? s.next() : f.createFolder(name);
  }
  return f;
}

// ─── DIAGNOSTICS ─────────────────────────────────────────────────────────────

function findWorkbook() {
  let folder = DriveApp.getRootFolder();
  for (const name of EXPORT_FOLDER_PATH) {
    const s = folder.getFoldersByName(name);
    if (!s.hasNext()) { Logger.log(`❌ Folder not found: "${name}"`); return; }
    folder = s.next();
  }
  const files = folder.getFilesByName(EXPORT_WORKBOOK_NAME);
  if (files.hasNext()) {
    const f = files.next();
    Logger.log(`✅ Found: ${f.getName()}`);
    Logger.log(`   URL: ${f.getUrl()}`);
  } else {
    Logger.log(`❌ "${EXPORT_WORKBOOK_NAME}" not found in "${folder.getName()}".`);
    Logger.log('   Run runFullExport() to create it.');
  }
}

function listDriveFolder() {
  let folder = DriveApp.getRootFolder();
  for (const name of EXPORT_FOLDER_PATH) {
    const s = folder.getFoldersByName(name);
    if (!s.hasNext()) { Logger.log(`❌ Folder not found: "${name}"`); return; }
    folder = s.next();
  }
  Logger.log(`Contents of "${folder.getName()}" (${folder.getUrl()}):`);
  const subs = folder.getFolders();
  while (subs.hasNext()) Logger.log('  📁 ' + subs.next().getName());
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    Logger.log(`  📄 ${f.getName()} — ${f.getUrl()}`);
  }
}

// ─── UTILITIES ───────────────────────────────────────────────────────────────

function getCredentials() {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('CLICKUP_TOKEN');
  const teamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!token || !teamId) {
    Logger.log('❌ Credentials missing — run dbSetup() first.');
    return {};
  }
  return { token, teamId };
}

function cuGet(url, token) {
  try {
    const res  = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code !== 200) {
      Logger.log(`❌ HTTP ${code}: ${url.split('?')[0]} — ${body.substring(0, 200)}`);
      return null;
    }
    if (!body || body.trimStart().charAt(0) === '<') {
      Logger.log(`⚠️ Non-JSON response from: ${url.split('?')[0]}`);
      return null;
    }
    return JSON.parse(body);
  } catch(e) {
    Logger.log(`❌ Fetch error: ${e.message}`);
    return null;
  }
}

function dbFmtTs(ms) {
  if (!ms || isNaN(ms) || ms <= 0) return '';
  return Utilities.formatDate(new Date(ms), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function dbFmtDur(ms) {
  if (!ms || isNaN(ms) || ms <= 0) return '';
  const totalMin = Math.floor(Math.abs(ms) / 60000);
  return `${Math.floor(totalMin / 60)}h ${totalMin % 60}m`;
}