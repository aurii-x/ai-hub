// =============================================================================
// ClickUp Dashboard Export  v1.7
//
// CHANGELOG
//   v1.7 - Sync Log always written on every run, including zero-task runs.
//           Added column G "Log" to Sync Log with status messages:
//             ✅ Nothing new. Anchor advanced.  (when tasks = 0)
//             ✅ Updated successfully. Anchor advanced.  (on success)
//             ❌ <error message>  (on failure)
//           runIncrementalSync wrapped in try/catch — failures now logged to
//           Sync Log instead of disappearing silently.
//           writeTaskRows performance fix: reads all existing rows once into
//           memory, merges updates in-place, writes back in one batch call
//           instead of per-cell setValue calls. Eliminates the slowness and
//           Sheets API rate limit issue seen at 3:35 AM.
//
//   v1.6 - Fixed writeTaskRows upsert mode: incremental sync was overwriting
//           tracked time columns (L-P) and date columns (C-D) with blanks.
//           Now preserves existing cell values unless fresh data available.
//
//   v1.5 - Full regeneration consolidating all prior changes. Version number
//           corrected in file header.
//
//   v1.4 - Fixed fetchAllTimeEntries: ClickUp time entries endpoint ignores
//           page param. Switched to monthly date chunking (one call per month).
//           Added buildMonthlyChunks() helper.
//
//   v1.3 - runFullExport() switched to upsert mode. Re-runs show accurate
//           Added/Updated counts. Added clearExportState().
//
//   v1.2 - Renamed conflicting functions to db-prefix to prevent namespace
//           collisions when sharing a project with PipelineV2.
//
//   v1.1 - Credential keys renamed to CLICKUP_TOKEN/CLICKUP_TEAM_ID.
//           Workbook created first. Tasks written page-by-page. Added
//           runRefreshTimeEntries(), findWorkbook(), listDriveFolder().
//
// RUN ORDER (first time):
//   1. dbSetup()                → save credentials
//   2. dbCheckSetup()           → confirm they saved
//   3. runFullExport()          → full pull from FULL_EXPORT_START_DATE
//      if it times out during time entries:
//   4. runRefreshTimeEntries()  → patch time cols without re-fetching tasks
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
const HEADERS = [
  'task_id',             // A
  'task_name',           // B
  'task_start_date',     // C
  'task_end_date',       // D
  'tags',                // E
  'space_id',            // F
  'folder_id',           // G
  'list_id',             // H
  'folder_name',         // I
  'list_name',           // J
  'status',              // K
  'tracked_time_start',  // L
  'tracked_time_end',    // M
  'tracked_time_ms',     // N
  'tracked_time',        // O
  'time_entry_count',    // P
  'task_created',        // Q
  'task_last_modified',  // R
];

// Column index constants (0-based, matching HEADERS above)
const COL = {
  TASK_ID:    0,  // A
  TASK_NAME:  1,  // B
  START_DATE: 2,  // C
  END_DATE:   3,  // D
  TAGS:       4,  // E
  SPACE_ID:   5,  // F
  FOLDER_ID:  6,  // G
  LIST_ID:    7,  // H
  FOLDER:     8,  // I
  LIST:       9,  // J
  STATUS:     10, // K
  TT_START:   11, // L
  TT_END:     12, // M
  TT_MS:      13, // N
  TT_HUMAN:   14, // O
  TT_COUNT:   15, // P
  CREATED:    16, // Q
  MODIFIED:   17, // R
};

// Columns to update when fresh time entry data IS available (all columns)
const ALL_COLS = Object.values(COL);

// Columns to update when NO time entry data — excludes L-P and avoids
// blanking C-D if ClickUp returns empty (pipeline may have backfilled them)
const META_COLS_IDX = [0,1,4,5,6,7,8,9,10,16,17]; // A,B,E-K,Q,R

// ─── SETUP ───────────────────────────────────────────────────────────────────

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

function runFullExport() {
  const t0 = Date.now();
  const { token, teamId } = getCredentials();
  if (!token) return;

  const props    = PropertiesService.getScriptProperties();
  const lastSync = props.getProperty('LAST_SYNC_TS');
  if (lastSync) {
    Logger.log('⚠️ Full export already ran (LAST_SYNC_TS is set).');
    Logger.log('   Use runIncrementalSync() or setupHourlyTrigger() instead.');
    Logger.log('   To force a re-export, run clearExportState() first.');
    return;
  }

  Logger.log(`🚀 Full export — ${FULL_EXPORT_START_DATE} → now`);
  const startMs = new Date(FULL_EXPORT_START_DATE + 'T00:00:00Z').getTime();
  const nowMs   = Date.now();

  Logger.log('📂 Opening / creating workbook...');
  const ss    = dbGetOrCreateWorkbook();
  const sheet = getOrCreateTasksSheet(ss, false);
  Logger.log(`   URL: ${ss.getUrl()}`);

  Logger.log('🗂️ Building row index...');
  const index = buildTaskIdIndex(sheet);
  Logger.log(`   Existing rows: ${Object.keys(index).length}`);

  Logger.log('📡 Fetching tasks (writing each page as it arrives)...');
  let page    = 0;
  let total   = 0;
  const stats = { added: 0, updated: 0 };

  while (true) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit at page ${page} — ${total} tasks processed.`);
      Logger.log('   Run clearExportState() then runFullExport() to restart cleanly.');
      return;
    }

    Logger.log(`  Page ${page}...`);
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
    Logger.log(`  ✅ Page ${page} done — added: ${stats.added}, updated: ${stats.updated}`);

    if (res.tasks.length < PAGE_SIZE || res.last_page) { Logger.log('  Last page.'); break; }
    page++;
    Utilities.sleep(200);
  }

  Logger.log(`\n✅ Tasks complete — ${total} tasks (added: ${stats.added}, updated: ${stats.updated}).`);

  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⚠️ No time left for time entries. Run runRefreshTimeEntries().');
    return;
  }

  Logger.log('\n⏱️ Fetching time entries (monthly chunks)...');
  const timeMap = fetchAllTimeEntries(token, teamId, startMs, nowMs, t0);
  Logger.log(`✅ ${Object.keys(timeMap).length} tasks have tracked time.`);

  if (Object.keys(timeMap).length > 0) {
    Logger.log('📝 Patching time entry columns...');
    patchTimeEntryColumns(sheet, timeMap);
  }

  props.setProperty('LAST_SYNC_TS', String(nowMs));
  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  logSyncRun(ss, 'FULL EXPORT', total, elapsed, stats, '✅ Full export complete. Anchor set.');
  Logger.log(`\n🏁 Full export done — ${total} tasks in ${elapsed}s.`);
  Logger.log('   Next: run setupHourlyTrigger() to automate incremental syncs.');
}

/**
 * Recovery function — use when runFullExport timed out during time entry phase.
 * Skips task fetch entirely; only patches time entry columns L–P.
 */
function runRefreshTimeEntries() {
  const t0 = Date.now();
  const { token, teamId } = getCredentials();
  if (!token) return;

  const ss    = dbGetOrCreateWorkbook();
  const sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('❌ Tasks sheet is empty — run runFullExport() first.'); return;
  }

  const startMs = new Date(FULL_EXPORT_START_DATE + 'T00:00:00Z').getTime();
  const nowMs   = Date.now();

  Logger.log(`⏱️ Refreshing time entries — ${FULL_EXPORT_START_DATE} → now`);
  const timeMap = fetchAllTimeEntries(token, teamId, startMs, nowMs, t0);
  Logger.log(`✅ ${Object.keys(timeMap).length} tasks have tracked time.`);

  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⚠️ Time budget hit — re-run runRefreshTimeEntries() to continue.'); return;
  }

  Logger.log('📝 Patching time entry columns...');
  patchTimeEntryColumns(sheet, timeMap);

  PropertiesService.getScriptProperties().setProperty('LAST_SYNC_TS', String(nowMs));
  Logger.log(`🏁 Time entries refreshed in ${((Date.now() - t0) / 1000).toFixed(1)}s.`);
}

/**
 * Incremental sync — runs hourly via trigger.
 * Always writes a row to Sync Log, even when nothing changed.
 * Failures are caught and logged to Sync Log column G.
 */
function runIncrementalSync() {
  const t0  = Date.now();
  const ss  = dbGetOrCreateWorkbook();
  let logMsg = '';

  try {
    const { token, teamId } = getCredentials();
    if (!token) { logSyncRun(ss, 'INCREMENTAL', 0, '0', null, '❌ Credentials missing.'); return; }

    const props      = PropertiesService.getScriptProperties();
    const lastSyncTs = Number(props.getProperty('LAST_SYNC_TS') || '0');
    if (!lastSyncTs) {
      logMsg = '❌ No LAST_SYNC_TS — run runFullExport() first.';
      logSyncRun(ss, 'INCREMENTAL', 0, '0', null, logMsg);
      Logger.log(logMsg);
      return;
    }

    const bufferMs    = INCREMENTAL_BUFFER_HRS * 60 * 60 * 1000;
    const fetchFromMs = lastSyncTs - bufferMs;
    const nowMs       = Date.now();

    Logger.log(`🔄 Incremental sync — from: ${dbFmtTs(fetchFromMs)}`);

    Logger.log('📡 Fetching updated tasks...');
    const tasks = fetchAllTasks(token, teamId, fetchFromMs, t0);
    Logger.log(`   Updated tasks: ${tasks.length}`);

    // ── Always advance anchor and log, even when nothing changed ─────────────
    if (tasks.length === 0) {
      props.setProperty('LAST_SYNC_TS', String(nowMs));
      logMsg = '✅ Nothing new. Anchor advanced.';
      logSyncRun(ss, 'INCREMENTAL', 0, ((Date.now()-t0)/1000).toFixed(1), {added:0,updated:0}, logMsg);
      Logger.log(logMsg);
      return;
    }

    Logger.log('⏱️ Fetching time entries for updated tasks...');
    const taskIds = tasks.map(t => t.id);
    const timeMap = fetchTimeEntriesForTaskIds(token, teamId, taskIds, fetchFromMs, nowMs);
    Logger.log(`   Tasks with time entries: ${Object.keys(timeMap).length}`);

    Logger.log('📝 Upserting rows...');
    const sheet = getOrCreateTasksSheet(ss, false);
    const stats = writeTaskRows(sheet, tasks, timeMap, true);

    props.setProperty('LAST_SYNC_TS', String(nowMs));
    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    logMsg = `✅ Updated successfully. Anchor advanced. (added: ${stats.added}, updated: ${stats.updated})`;
    logSyncRun(ss, 'INCREMENTAL', tasks.length, elapsed, stats, logMsg);
    Logger.log(`🏁 Sync done — ${logMsg}`);

  } catch(err) {
    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    logMsg = `❌ ${err.message}`;
    logSyncRun(ss, 'INCREMENTAL', 0, elapsed, null, logMsg);
    Logger.log(`❌ runIncrementalSync failed: ${err.message}`);
  }
}

function clearExportState() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_SYNC_TS');
  Logger.log('✅ Export state cleared. runFullExport() will do a clean re-export.');
}

// ─── CLICKUP TASK FETCH ───────────────────────────────────────────────────────

function fetchAllTasks(token, teamId, updatedAfterMs, t0) {
  const allTasks = [];
  let page       = 0;

  while (true) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit fetching tasks at page ${page}.`); break;
    }
    Logger.log(`  Tasks page ${page}...`);
    const url = `https://api.clickup.com/api/v2/team/${teamId}/task`
      + `?page=${page}&include_closed=true&subtasks=true`
      + `&order_by=updated&reverse=true&date_updated_gt=${updatedAfterMs}`;

    const res = cuGet(url, token);
    if (!res || !res.tasks) { Logger.log(`⚠️ Failed on page ${page}.`); break; }
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
 * Fetches all time entries using monthly chunks.
 * ClickUp time entries endpoint ignores the page param — monthly chunking
 * is the correct pagination strategy for large date ranges.
 */
function fetchAllTimeEntries(token, teamId, startMs, endMs, t0) {
  const map    = {};
  let total    = 0;
  const chunks = buildMonthlyChunks(startMs, endMs);

  Logger.log(`  Fetching time entries in ${chunks.length} monthly chunks...`);

  for (let i = 0; i < chunks.length; i++) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Time budget hit at chunk ${i+1}/${chunks.length} (${total} entries).`); break;
    }
    const { chunkStart, chunkEnd, label } = chunks[i];
    const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`
      + `?start_date=${chunkStart}&end_date=${chunkEnd}`;
    const res  = cuGet(url, token);
    if (!res) { Logger.log(`⚠️ Failed for chunk ${label}.`); continue; }
    const data = Array.isArray(res.data) ? res.data : [];
    Logger.log(`  ${label}: ${data.length} entries`);
    data.forEach(e => mergeTimeEntry(map, e));
    total += data.length;
    Utilities.sleep(200);
  }

  Logger.log(`  Total: ${total} entries across ${Object.keys(map).length} tasks`);
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
 * Incremental: single call for the recent window, filtered to updated task IDs.
 */
function fetchTimeEntriesForTaskIds(token, teamId, taskIds, startMs, endMs) {
  const taskIdSet = new Set(taskIds.map(String));
  const map       = {};

  Logger.log(`  Fetching time entries for ${taskIds.length} tasks...`);
  const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`
    + `?start_date=${startMs}&end_date=${endMs}`;
  const res  = cuGet(url, token);
  if (!res) { Logger.log('⚠️ Time entry fetch failed.'); return map; }
  const data = Array.isArray(res.data) ? res.data : [];
  Logger.log(`  ${data.length} entries in window — filtering...`);
  data.forEach(e => { if (e.task && taskIdSet.has(String(e.task.id))) mergeTimeEntry(map, e); });
  Logger.log(`  ${Object.keys(map).length} tasks matched.`);
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

// ─── PATCH TIME ENTRIES INTO EXISTING SHEET ──────────────────────────────────

function patchTimeEntryColumns(sheet, timeMap) {
  const index   = buildTaskIdIndex(sheet);
  let patched   = 0;
  const taskIds = Object.keys(timeMap);
  Logger.log(`  Patching ${taskIds.length} tasks...`);

  for (const taskId of taskIds) {
    const rowNum = index[taskId];
    if (!rowNum) continue;
    const td = timeMap[taskId];
    sheet.getRange(rowNum, COL.TT_START + 1, 1, 5).setValues([[
      dbFmtTs(td.minStart), dbFmtTs(td.maxEnd), td.totalMs, dbFmtDur(td.totalMs), td.count,
    ]]);
    patched++;
    if (patched % 100 === 0) { SpreadsheetApp.flush(); Logger.log(`  Patched ${patched}/${taskIds.length}...`); }
  }
  SpreadsheetApp.flush();
  Logger.log(`✅ Patched ${patched} rows.`);
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

// ─── SHEET WRITE (batch upsert) ───────────────────────────────────────────────

/**
 * Writes task rows to the sheet.
 *
 * In upsert mode (incremental sync):
 *   1. Reads ALL existing rows into memory in one call
 *   2. Merges updates in-place in the in-memory array:
 *      - If timeMap has data for the task: update all columns
 *      - If not: update metadata cols only, preserve L-P (time) and
 *        only update C-D (dates) if ClickUp returns a non-empty value
 *   3. Writes the entire updated region back in one batch call
 *   4. Appends new rows in one batch call
 *
 * Result: 1 read + 2 writes regardless of how many rows were updated,
 * vs N individual reads+writes in the previous per-cell approach.
 */
function writeTaskRows(sheet, tasks, timeMap, upsert) {
  const stats = { added: 0, updated: 0 };

  if (!upsert) {
    const matrix = tasks.map(t => buildTaskRow(t, timeMap[t.id]));
    if (matrix.length > 0) sheet.getRange(2, 1, matrix.length, HEADERS.length).setValues(matrix);
    stats.added = tasks.length;
    return stats;
  }

  // ── Build index and read all existing rows into memory in one call ────────
  const index      = buildTaskIdIndex(sheet);
  const lastRow    = sheet.getLastRow();
  const existingMatrix = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues()
    : [];

  const newRows    = [];
  let   hasUpdates = false;

  for (const task of tasks) {
    const row    = buildTaskRow(task, timeMap[task.id] || null);
    const td     = timeMap[task.id] || null;
    const rowNum = index[task.id];

    if (rowNum) {
      // Merge into existing row in memory
      const matrixIdx = rowNum - 2; // adjust for 0-index and header row
      const existing  = existingMatrix[matrixIdx];
      const merged    = [...existing];

      if (td) {
        // Full update — fresh time entry data available
        ALL_COLS.forEach(i => { merged[i] = row[i]; });
      } else {
        // Partial update — update metadata cols, preserve time entry cols L-P
        META_COLS_IDX.forEach(i => { merged[i] = row[i]; });
        // Update date cols only if ClickUp returned a value
        if (row[COL.START_DATE] !== '') merged[COL.START_DATE] = row[COL.START_DATE];
        if (row[COL.END_DATE]   !== '') merged[COL.END_DATE]   = row[COL.END_DATE];
        // COL.TT_START through COL.TT_COUNT (11-15) are left untouched
      }

      existingMatrix[matrixIdx] = merged;
      hasUpdates = true;
      stats.updated++;

    } else {
      // New row — append
      newRows.push(row);
      stats.added++;
    }
  }

  // ── Batch write ───────────────────────────────────────────────────────────
  if (hasUpdates && existingMatrix.length > 0) {
    Logger.log(`  Writing ${stats.updated} updated rows in one batch...`);
    sheet.getRange(2, 1, existingMatrix.length, HEADERS.length).setValues(existingMatrix);
  }
  if (newRows.length > 0) {
    Logger.log(`  Appending ${newRows.length} new rows...`);
    const startRow = lastRow + 1;
    sheet.getRange(startRow, 1, newRows.length, HEADERS.length).setValues(newRows);
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
    Logger.log(`✨ Created workbook: "${EXPORT_WORKBOOK_NAME}"`);
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
  sheet.getRange(1, 1, 1, HEADERS.length)
       .setValues([HEADERS])
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

/**
 * Logs every run to the Sync Log — including zero-task runs.
 * Column G "Log" holds the status message.
 */
function logSyncRun(ss, mode, taskCount, elapsed, stats, logMsg) {
  let log = ss.getSheetByName(SYNC_LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(SYNC_LOG_SHEET_NAME);
    log.appendRow(['Timestamp','Mode','Tasks Processed','Added','Updated','Duration (s)','Log']);
    log.getRange(1,1,1,7).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    log.setColumnWidth(7, 400);
  }
  log.appendRow([
    new Date(),
    mode,
    taskCount,
    stats ? stats.added   : '',
    stats ? stats.updated : '',
    elapsed,
    logMsg || '',
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
    Logger.log(`❌ Not found in "${folder.getName()}". Run runFullExport().`);
  }
}

function listDriveFolder() {
  let folder = DriveApp.getRootFolder();
  for (const name of EXPORT_FOLDER_PATH) {
    const s = folder.getFoldersByName(name);
    if (!s.hasNext()) { Logger.log(`❌ Folder not found: "${name}"`); return; }
    folder = s.next();
  }
  Logger.log(`Contents of "${folder.getName()}":`);
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
  if (!token || !teamId) { Logger.log('❌ Credentials missing — run dbSetup() first.'); return {}; }
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
      Logger.log(`⚠️ Non-JSON response from: ${url.split('?')[0]}`); return null;
    }
    return JSON.parse(body);
  } catch(e) {
    Logger.log(`❌ Fetch error: ${e.message}`); return null;
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