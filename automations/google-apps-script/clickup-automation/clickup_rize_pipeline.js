// =============================================================================
// GLOBAL USER ENVIRONMENT CONFIGURATION
// =============================================================================
const START_DATE_OVERRIDE = "2026-01-01"; 
const END_DATE_OVERRIDE   = "2026-06-29"; 
const CLEAN_TAB_NAME      = 'Rize_Clean_Sync';
const TARGET_FOLDER_PATH  = ['AppData', '3.1 clickup-automation']; 
const BASE_FILENAME       = 'ClickUp-Rize Sync';         
const INCLUDE_CLOSED_TASKS = true;
//const TIME_BUFFER_MS      = 25000;
//const MAX_RUNTIME_MS      = 660000;

// =============================================================================
// STEP 1: Extract Rize entries into staging — IDEMPOTENT VERSION
//
// KEY CHANGES FROM ORIGINAL:
//   • No longer archives and recreates the workbook on every run.
//   • Opens the existing workbook if it already exists (creates once on first run).
//   • Reads all rize_time_entry_ids already in the clean sheet BEFORE fetching
//     from Rize, so duplicate fetching is impossible at the source.
//   • The raw staging tab is still overwritten each run (it's a true staging area
//     and has no permanent value once Step 2 has processed it).
// =============================================================================
function runStep1_ExtractRizeToStaging() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const rizeKey = props.getProperty('RIZE_API_KEY');
  if (!rizeKey) { Logger.log('❌ RIZE_API_KEY missing.'); return; }

  let startISO, endISO, executionMode;
  if (START_DATE_OVERRIDE.trim() !== "" && END_DATE_OVERRIDE.trim() !== "") {
    executionMode = "MANUAL OVERRIDE MODE";
    startISO = new Date(START_DATE_OVERRIDE).toISOString();
    endISO   = new Date(END_DATE_OVERRIDE).toISOString();
  } else {
    executionMode = "AUTOMATED ROLLING MODE";
    endISO   = new Date().toISOString();
    startISO = new Date(Date.now() - LOOKBACK_HOURS * 60 * 60 * 1000).toISOString();
  }
  Logger.log(`⚙️ Running in: ${executionMode}`);
  Logger.log(`📅 Window: ${startISO} ──> ${endISO}`);

  // ── IDEMPOTENCY CHANGE 1: open existing workbook, create only if missing ──
  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const ss = getOrCreateWorkbook(destinationFolder, BASE_FILENAME);
  Logger.log(`📖 Workbook ready: "${BASE_FILENAME}"`);

  // ── IDEMPOTENCY CHANGE 2: load already-processed IDs from the clean sheet
  //    BEFORE hitting the Rize API, so we can skip re-fetching known entries ──
  const alreadyProcessedIds = loadAlreadyProcessedRizeIds(ss);
  Logger.log(`📚 Already staged in clean sheet: ${alreadyProcessedIds.size} unique Rize entry IDs`);

  Logger.log(`📡 Fetching from Rize...`);
  const rawNodes = fetchRizeLookbackDataset(rizeKey, startISO, endISO, t0);
  Logger.log(`📊 Rize returned ${rawNodes.length} entries in this window`);

  // ── IDEMPOTENCY CHANGE 3: filter out entries that are already in the clean
  //    sheet — only stage genuinely new ones ──
  const newNodes = rawNodes.filter(node => !alreadyProcessedIds.has(String(node.id)));
  Logger.log(`✨ New entries not yet staged: ${newNodes.length} (skipped ${rawNodes.length - newNodes.length} already-processed)`);

  // Always rewrite the raw staging tab with only the NEW nodes for this run.
  // The raw tab is a true staging area — it has no permanent value and is safe
  // to overwrite. The clean sheet is the permanent record.
  commitRawDatasetToSheet(ss, 'Rize_Stage_Raw', newNodes);
  Logger.log(`🏁 Step 1 complete in ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

// =============================================================================
// STEP 2: Transform raw staging → clean sync sheet — already mostly correct,
// but now also benefits from Step 1 only staging genuinely new nodes.
// No structural changes needed here — the dedup logic already works correctly
// when the workbook is preserved across runs (which it now is).
// =============================================================================
function runStep2_CleanAndTransformStaging() {
  const t0 = Date.now();

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) {
    Logger.log(`❌ Workbook "${BASE_FILENAME}" not found. Run Step 1 first.`); return;
  }
  const ss = SpreadsheetApp.open(fileSearch.next());

  const rawSheet = ss.getSheetByName('Rize_Stage_Raw');
  if (!rawSheet) { Logger.log('❌ Rize_Stage_Raw tab missing.'); return; }

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length <= 1) {
    Logger.log('🏁 Step 2: No new raw records to process — clean sheet is up to date.');
    return;
  }

  // Check if all non-header rows are "no new entries" placeholder rows
  // Placeholder rows have '—' in col 1 (the id column, after the Run Log col)
  const dataRows = rawData.slice(1).filter(r => String(r[1] || '') !== '—' && String(r[1] || '').trim() !== '');
  if (dataRows.length === 0) {
    Logger.log('🏁 Step 2: Raw sheet has no new entries — nothing to transform.');
    return;
  }

  const schemaHeaders = ['start', 'end', 'duration', 'description', 'rize_time_entry_id', 'rize_task_id', 'task_name_lookup'];
  const schemaTypes   = ['Integer (Unix MS)', 'Integer (Unix MS)', 'Integer (MS)', 'String (Text Note)', 'String (Rize ID)', 'String (Rize Task ID)', 'String (Text Name)'];

  let cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) cleanSheet = ss.insertSheet(CLEAN_TAB_NAME);

  // Read existing clean records to build the dedup set
  const existingCleanData = cleanSheet.getDataRange().getValues();
  const processedEntryIds = new Set();
  let masterRecordsList = [];

  if (existingCleanData.length > 2) {
    for (let i = 2; i < existingCleanData.length; i++) {
      const row = existingCleanData[i];
      const existingId = String(row[4]);
      if (existingId && existingId.trim() !== "" && existingId !== "—") {
        processedEntryIds.add(existingId);
        masterRecordsList.push({
          start: row[0], end: row[1], duration: row[2], description: row[3],
          rizeId: existingId, rizeTaskId: row[5], taskName: row[6],
          // Preserve any diagnostic/sync columns already written by Steps 3a/3b
          extraCols: row.slice(7)
        });
      }
    }
    Logger.log(`📚 Loaded ${masterRecordsList.length} existing clean records`);
  }

  let newEntriesCount = 0;
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    // Col 0 = Run Log (timestamp) — skip it; data starts at col 1
    const rizeId = String(row[1]);
    if (!rizeId || rizeId === '—' || rizeId === '') continue; // skip placeholder rows
    if (processedEntryIds.has(rizeId)) continue;

    const startMs    = row[2] ? new Date(row[2]).getTime() : 0;
    const endMs      = row[3] ? new Date(row[3]).getTime() : 0;
    const durationMs = (startMs > 0 && endMs > 0) ? (endMs - startMs) : 0;
    const title      = row[4];
    const desc       = row[5];
    const taskJson   = row[6];

    let taskName = 'Unassigned Rize Task', rizeTaskId = 'N/A';
    if (taskJson && taskJson.trim() !== "") {
      try {
        const parsed = JSON.parse(taskJson);
        if (parsed.name) taskName  = String(parsed.name);
        if (parsed.id)   rizeTaskId = String(parsed.id);
      } catch(e) {}
    }

    const commentText = desc && desc.trim() !== "" ? desc : title;
    masterRecordsList.push({
      start: startMs, end: endMs, duration: durationMs, description: commentText,
      rizeId, rizeTaskId, taskName,
      extraCols: [] // new entries have no diagnostic cols yet
    });
    processedEntryIds.add(rizeId);
    newEntriesCount++;
  }

  Logger.log(`✨ ${newEntriesCount} new entries added. Total: ${masterRecordsList.length}`);
  if (masterRecordsList.length === 0) { Logger.log('🏁 Step 2: Nothing to write.'); return; }

  // Sort by task name then chronologically
  masterRecordsList.sort((a, b) => {
    const na = a.taskName.toLowerCase(), nb = b.taskName.toLowerCase();
    if (na < nb) return -1; if (na > nb) return 1;
    return a.start - b.start;
  });

  // Rewrite the clean sheet — preserve extraCols (diagnostic/sync status columns
  // written by Steps 3a and 3b) so a Step 2 re-run doesn't wipe sync history
  cleanSheet.clearContents();
  cleanSheet.appendRow(schemaHeaders);
  cleanSheet.appendRow(schemaTypes);

 /* const finalMatrix = masterRecordsList.map(r => [
    r.start, r.end, r.duration, r.description, r.rizeId, r.rizeTaskId, r.taskName,
    ...(r.extraCols || [])
  ]);
  cleanSheet.getRange(3, 1, finalMatrix.length, finalMatrix[0].length).setValues(finalMatrix);
*/
// AFTER — normalize all rows to the same width before writing
const finalMatrix = masterRecordsList.map(r => [
    r.start, r.end, r.duration, r.description, r.rizeId, r.rizeTaskId, r.taskName,
    ...(r.extraCols || [])
]);

// Find the widest row (existing rows with diagnostic cols will be wider than new rows)
const maxCols = Math.max(...finalMatrix.map(row => row.length));

// Pad every row to the same width with empty strings so setValues() doesn't crash
const normalizedMatrix = finalMatrix.map(row => {
    while (row.length < maxCols) row.push('');
    return row;
});

cleanSheet.getRange(3, 1, normalizedMatrix.length, maxCols).setValues(normalizedMatrix);
//==========================================================================================
  // Reapply headers styling for the base 7 columns
  cleanSheet.getRange(1, 1, 1, schemaHeaders.length).setFontWeight('bold').setBackground('#2F5597').setFontColor('#FFFFFF');
  cleanSheet.getRange(2, 1, 1, schemaHeaders.length).setFontStyle('italic').setBackground('#D9E1F2');
  cleanSheet.setFrozenRows(2);
  cleanSheet.autoResizeColumns(1, schemaHeaders.length);

  Logger.log(`✅ Clean sheet updated. ${masterRecordsList.length} total records.`);
  Logger.log(`🏁 Step 2 complete in ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

// =============================================================================
// STEP 3a: Diagnostic matcher — no structural changes needed.
// Now genuinely idempotent: re-running 3a refreshes the overlap analysis in
// columns H–L without touching the sync status in column L that Step 3b wrote,
// because it only overwrites columns 8–12 (H through L) and Step 3b's result
// is in column 12 (L). Wait — that IS column L, which Step 3a also writes to.
//
// FIX: Step 3a now SKIPS rows already marked 'Synced Successfully' by Step 3b
// so a 3a re-run never resets already-synced rows back to 'New Entry'.
// =============================================================================
function runStep3a_DiagnosticClickUpMatcher() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const cuToken  = props.getProperty('CLICKUP_TOKEN');
  const cuTeamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!cuToken || !cuTeamId) { Logger.log('❌ CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.'); return; }

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) { Logger.log('❌ Workbook not found. Run Step 1 first.'); return; }

  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) { Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return; }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 2) { Logger.log('🏁 Step 3a: No records to evaluate.'); return; }

  Logger.log('📡 Fetching ClickUp task inventory...');
  const clickUpTasks = fetchAllClickUpTasksInventory(cuTeamId, cuToken, INCLUDE_CLOSED_TASKS);
  const clickUpTaskNameIndex = {};
  clickUpTasks.forEach(task => {
    if (task.name) clickUpTaskNameIndex[task.name.trim().toLowerCase()] = String(task.id);
  });
  Logger.log(`✨ ClickUp task index built: ${Object.keys(clickUpTaskNameIndex).length} tasks`);

  const diagnosticHeaders = ['clickup_task_id', 'clickup_start_display', 'clickup_end_display', 'clickup_duration_display', 'sync_match_status'];
  const diagnosticTypes   = ['String (ClickUp ID)', 'String (Human Date)', 'String (Human Date)', 'String (Human Duration)', 'String (Match Status)'];
  cleanSheet.getRange(1, 8, 1, diagnosticHeaders.length).setValues([diagnosticHeaders]);
  cleanSheet.getRange(2, 8, 1, diagnosticTypes.length).setValues([diagnosticTypes]);

  const outputMatrix = [];
  const cachedTimeEntries = {};

  for (let i = 2; i < cleanData.length; i++) {
    const row = cleanData[i];

    // ── IDEMPOTENCY FIX: skip rows Step 3b already synced successfully ──────
    const existingStatus = String(row[11] || '');
    if (existingStatus === 'Synced Successfully') {
      outputMatrix.push([row[7], row[8], row[9], row[10], 'Synced Successfully']);
      continue;
    }

    const startMs      = Number(row[0]);
    const endMs        = Number(row[1]);
    const durationMs   = Number(row[2]);
    const taskNameLookup = String(row[6]);

    const matchedTaskId = clickUpTaskNameIndex[taskNameLookup.trim().toLowerCase()];

    const humanStart    = startMs > 0 ? Utilities.formatDate(new Date(startMs), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : '—';
    const humanEnd      = endMs   > 0 ? Utilities.formatDate(new Date(endMs),   Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : '—';
    const humanDuration = durationMs > 0 ? formatMsToHumanReadable(durationMs) : '0h 0m';

    if (!matchedTaskId) {
      outputMatrix.push(['Missing Target Task', humanStart, humanEnd, humanDuration, 'Missing ClickUp Task']);
      continue;
    }

    if (!cachedTimeEntries[matchedTaskId]) {
      cachedTimeEntries[matchedTaskId] = fetchLiveClickUpTimeEntriesArray(matchedTaskId, cuToken, cuTeamId);
    }
    const liveEntries = cachedTimeEntries[matchedTaskId];

    let matchStatus = 'New Entry';
    for (const cuEntry of liveEntries) {
      const isOverlapping = Math.max(startMs, cuEntry.start) < Math.min(endMs, cuEntry.end) + 5000;
      if (isOverlapping) {
        const durationDiff = Math.abs(durationMs - cuEntry.duration);
        const startDiff    = Math.abs(startMs - cuEntry.start);
        matchStatus = (durationDiff < 5000 && startDiff < 5000) ? 'Already Matched' : 'Conflict';
        break;
      }
    }

    outputMatrix.push([matchedTaskId, humanStart, humanEnd, humanDuration, matchStatus]);
  }

  cleanSheet.getRange(3, 8, outputMatrix.length, diagnosticHeaders.length).setValues(outputMatrix);
  cleanSheet.getRange(1, 8, 1, diagnosticHeaders.length).setFontWeight('bold').setBackground('#5B9BD5').setFontColor('#FFFFFF');
  cleanSheet.getRange(2, 8, 1, diagnosticHeaders.length).setFontStyle('italic').setBackground('#E2EFDA');
  cleanSheet.autoResizeColumns(8, diagnosticHeaders.length);

  Logger.log(`✅ Step 3a complete. Evaluated ${outputMatrix.length} rows in ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

// =============================================================================
// STEP 3b: Sync engine — KEY FIX: Unlimited-plan-safe payload (tid/start/stop
// only). Also now writes sync result to column 12 AND updates the task's
// start/due date if currently empty, and posts the description as a comment.
// =============================================================================
function runStep3b_ActiveClickUpSyncLoader() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const cuToken  = props.getProperty('CLICKUP_TOKEN');
  const cuTeamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!cuToken || !cuTeamId) { Logger.log('❌ CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.'); return; }

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) { Logger.log('❌ Workbook not found.'); return; }

  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) { Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return; }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 2) { Logger.log('🏁 Step 3b: No records to sync.'); return; }

  // ── Live entry cache: task id → existing ClickUp time entries ──────────────
  // Fetched lazily on first encounter per task, so we only pay the API cost for
  // tasks that actually have rows to process. This makes Step 3b independently
  // idempotent — it never posts a duplicate even if Step 3a missed it or if a
  // previous run timed out before writing the Synced Successfully status.
  const liveEntryCache = {};

  function getOrFetchLiveEntries(taskId) {
    if (!liveEntryCache[taskId]) {
      liveEntryCache[taskId] = fetchLiveClickUpTimeEntriesArray(taskId, cuToken, cuTeamId);
    }
    return liveEntryCache[taskId];
  }

  function hasOverlap(taskId, startMs, endMs) {
    const existing = getOrFetchLiveEntries(taskId);
    return existing.some(e => {
      const overlapMs = Math.min(endMs, e.end) - Math.max(startMs, e.start);
      return overlapMs > 5000; // more than 5-second overlap = duplicate
    });
  }

  let successCount = 0, skippedCount = 0, liveSkipCount = 0, errorCount = 0;

  for (let i = 2; i < cleanData.length; i++) {
    const row = cleanData[i];

    const startMs       = Number(row[0]);
    const endMs         = Number(row[1]);
    const matchedTaskId = String(row[7]);
    const currentStatus = String(row[11] || '');

    // ── Sheet-level skip: rows already confirmed done ──────────────────────
    if (currentStatus !== 'New Entry' && !currentStatus.startsWith('API Error')) {
      skippedCount++;
      continue;
    }

    // Skip rows with no matched task
    if (!matchedTaskId || matchedTaskId === 'Missing Target Task' || matchedTaskId === '') {
      skippedCount++;
      continue;
    }

    if (!isValidMs(startMs) || !isValidMs(endMs)) {
      skippedCount++;
      continue;
    }

    // Runtime ceiling check
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log('⚠️ Approaching runtime ceiling — halting safely. Re-run Step 3b to continue.');
      break;
    }

    // ── Live duplicate check: verify against ClickUp before posting ────────
    if (hasOverlap(matchedTaskId, startMs, endMs)) {
      liveSkipCount++;
      cleanSheet.getRange(i + 1, 12).setValue('Already Matched').setBackground('#D9EAD3').setFontColor('#274E13');
      SpreadsheetApp.flush();
      Logger.log(`⏭ Row ${i + 1} already exists in ClickUp — skipped`);
      continue;
    }

    // ── POST time entry ────────────────────────────────────────────────────
    const payload = { tid: matchedTaskId, start: startMs, stop: endMs };
    const url = `https://api.clickup.com/api/v2/team/${cuTeamId}/time_entries`;
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: { 'Authorization': cuToken, 'Content-Type': 'application/json' },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code === 200 || code === 201) {
        successCount++;
        // Add to local cache so later rows for the same task don't duplicate
        liveEntryCache[matchedTaskId].push({ start: startMs, end: endMs });
        cleanSheet.getRange(i + 1, 12).setValue('Synced Successfully').setBackground('#E2EFDA').setFontColor('#375623');
        SpreadsheetApp.flush();
        Logger.log(`✅ Row ${i + 1} synced`);
        Utilities.sleep(1000);
      } else {
        errorCount++;
        const errText = response.getContentText().substring(0, 200);
        Logger.log(`🚨 Row ${i + 1} rejected (${code}): ${errText}`);
        cleanSheet.getRange(i + 1, 12).setValue(`API Error (${code})`).setBackground('#FCE4D6').setFontColor('#C65911');
        SpreadsheetApp.flush();
      }
    } catch(e) {
      errorCount++;
      Logger.log(`❌ Row ${i + 1} network error: ${e.message}`);
      cleanSheet.getRange(i + 1, 12).setValue('Network Error').setBackground('#FCE4D6').setFontColor('#C65911');
      SpreadsheetApp.flush();
    }
  }

  Logger.log(`🏁 Step 3b complete. Synced: ${successCount}, Sheet-skipped: ${skippedCount}, Live-deduped: ${liveSkipCount}, Errors: ${errorCount}`);
}

// =============================================================================
// NEW UTILITY: getOrCreateWorkbook
// Opens the existing workbook if found, creates it once if not.
// This replaces archiveExistingFileIfFound + SpreadsheetApp.create pattern.
// =============================================================================
function getOrCreateWorkbook(folder, filename) {
  const fileSearch = folder.getFilesByName(filename);
  if (fileSearch.hasNext()) {
    const existing = SpreadsheetApp.open(fileSearch.next());
    Logger.log(`📂 Opened existing workbook: "${filename}"`);
    return existing;
  }
  // First-time creation only
  const newSs = SpreadsheetApp.create(filename);
  const file  = DriveApp.getFileById(newSs.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  Logger.log(`✨ Created new workbook: "${filename}"`);
  return newSs;
}

// =============================================================================
// NEW UTILITY: loadAlreadyProcessedRizeIds
// Reads the clean sheet and returns a Set of all rize_time_entry_ids already
// staged, so Step 1 can pre-filter before fetching from Rize.
// =============================================================================
function loadAlreadyProcessedRizeIds(ss) {
  const ids = new Set();
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) return ids;
  const data = cleanSheet.getDataRange().getValues();
  // Column E (index 4) = rize_time_entry_id; rows start at index 2 (skip 2 header rows)
  for (let i = 2; i < data.length; i++) {
    const id = String(data[i][4] || '').trim();
    if (id && id !== '—' && id !== '') ids.add(id);
  }
  return ids;
}

// =============================================================================
// SHARED INFRASTRUCTURE & API HELPERS (unchanged from original)
// =============================================================================

function resolveOrCreateFolderPath(pathArray) {
  let currentFolder = DriveApp.getRootFolder();
  for (const folderName of pathArray) {
    const search = currentFolder.getFoldersByName(folderName);
    currentFolder = search.hasNext() ? search.next() : currentFolder.createFolder(folderName);
  }
  return currentFolder;
}

function fetchRizeLookbackDataset(apiKey, startISO, endISO, t0) {
  const rizeUrl = 'https://api.rize.io/api/v1/graphql';
  let capturedNodes = [], hasNextPage = true, cursor = null;

  while (hasNextPage) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log('🚨 Time limit approaching — halting Rize fetch safely.');
      break;
    }

    const query = `query {
      timeEntries(startTime: "${startISO}", endTime: "${endISO}", first: 50${cursor ? `, after: "${cursor}"` : ''}) {
        pageInfo { hasNextPage endCursor }
        nodes { id startTime endTime title description task { id name } }
      }
    }`;

    try {
      const response = UrlFetchApp.fetch(rizeUrl, {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + apiKey.trim(), 'Content-Type': 'application/json' },
        payload: JSON.stringify({ query }),
        muteHttpExceptions: true
      });

      const json = JSON.parse(response.getContentText());
      if (json.errors) { Logger.log('🚨 Rize API error: ' + JSON.stringify(json.errors)); break; }
      if (!json.data || !json.data.timeEntries) { Logger.log('❌ Unexpected Rize response structure.'); break; }

      const nodes    = json.data.timeEntries.nodes;
      const pageInfo = json.data.timeEntries.pageInfo;
      if (nodes && nodes.length > 0) nodes.forEach(n => capturedNodes.push(n));
      hasNextPage = pageInfo.hasNextPage;
      cursor      = pageInfo.endCursor;
      Utilities.sleep(200);
    } catch(e) {
      Logger.log(`❌ Rize fetch error: ${e.message}`); break;
    }
  }
  return capturedNodes;
}

/**
 * Appends new Rize staging entries to the tab rather than overwriting.
 * Column A is always "Run Log" — stamped with the Eastern Time of this run
 * in yyyy-mm-dd_H:mm format so every row is traceable to a specific pipeline
 * execution. Existing rows from prior runs are preserved unchanged.
 *
 * On first run the tab is created fresh. On all subsequent runs new rows are
 * simply appended at the bottom — no data is ever deleted.
 */
function commitRawDatasetToSheet(ss, tabName, nodes) {
  // Remove the default blank Sheet1 that Google creates with new workbooks
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) try { ss.deleteSheet(defaultSheet); } catch(e) {}

  let sheet = ss.getSheetByName(tabName);
  const isNewSheet = !sheet;
  if (isNewSheet) sheet = ss.insertSheet(tabName);

  // Run timestamp in Eastern Time (Raleigh) — format: 2026-06-25_9:34
  const runLabel = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd_H:mm');

  // DATA_HEADERS are the Rize payload columns — Run Log always precedes them in col A
  const DATA_HEADERS = ['id', 'startTime', 'endTime', 'title', 'description', 'task_json'];
  const ALL_HEADERS  = ['Run Log', ...DATA_HEADERS];

  if (isNewSheet) {
    // ── First run: write the header row and freeze it ──────────────────────
    sheet.appendRow(ALL_HEADERS);
    const hRange = sheet.getRange(1, 1, 1, ALL_HEADERS.length);
    hRange.setFontWeight('bold').setBackground('#7030A0').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140); // Run Log column
  } else {
    // ── Subsequent runs: ensure Run Log column exists ─────────────────────
    // (guard for sheets created before this change was applied)
    const existingHeader = sheet.getRange(1, 1).getValue();
    if (existingHeader !== 'Run Log') {
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue('Run Log')
        .setFontWeight('bold').setBackground('#7030A0').setFontColor('#FFFFFF');
      sheet.setColumnWidth(1, 140);
      Logger.log('📋 Inserted missing "Run Log" column into existing tab.');
    }
  }

  if (nodes.length === 0) {
    // Still log the run so there's a visible record even when nothing is new
    sheet.appendRow([runLabel, '—', 'No new entries this run', '—', '—', '—', '—']);
    Logger.log('ℹ️ Raw staging tab: no new entries this run (run logged).');
    return;
  }

  // Append new rows — each stamped with this run's label in col A
  const matrix = nodes.map(n => [
    runLabel,
    n.id,
    n.startTime,
    n.endTime,
    n.title || '',
    n.description || '',
    n.task ? JSON.stringify(n.task) : '',
  ]);

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, matrix.length, ALL_HEADERS.length).setValues(matrix);
  sheet.autoResizeColumns(2, DATA_HEADERS.length); // resize data cols; keep Run Log fixed width

  Logger.log(`📄 Raw staging tab: appended ${nodes.length} new entries (run: ${runLabel})`);
}

function fetchAllClickUpTasksInventory(teamId, token, includeClosed) {
  let allTasks = [], page = 0;
  while (true) {
    const url = `https://api.clickup.com/api/v2/team/${teamId}/task?page=${page}&subtasks=true&include_closed=${includeClosed}`;
    try {
      const res  = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': token }, muteHttpExceptions: true });
      const data = JSON.parse(res.getContentText());
      if (!data || !data.tasks || data.tasks.length === 0) break;
      data.tasks.forEach(t => allTasks.push(t));
      if (data.last_page) break;
      page++;
      Utilities.sleep(150);
    } catch(e) {
      Logger.log(`⚠️ ClickUp task fetch error on page ${page}: ${e.message}`); break;
    }
  }
  return allTasks;
}

function fetchLiveClickUpTimeEntriesArray(taskId, token, teamId) {
  const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries?task_id=${taskId}`;
  try {
    const res     = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': token }, muteHttpExceptions: true });
    const code    = res.getResponseCode();
    const rawText = res.getContentText();
    if (code !== 200) {
      Logger.log(`⚠️ Task ${taskId} time entries — HTTP ${code}: ${rawText.substring(0, 200)}`);
      return [];
    }
    // Guard against non-JSON responses (deleted/inaccessible tasks)
    if (!rawText || rawText.trimStart().charAt(0) !== '{') {
      Logger.log(`⚠️ Task ${taskId} returned non-JSON response — skipping. First 80 chars: ${rawText.substring(0, 80)}`);
      return [];
    }
    let payload;
    try {
      payload = JSON.parse(rawText);
    } catch(parseErr) {
      Logger.log(`⚠️ Task ${taskId} JSON parse failed: ${parseErr.message}. First 80 chars: ${rawText.substring(0, 80)}`);
      return [];
    }
    if (!payload || !Array.isArray(payload.data)) return [];
    return payload.data.map(e => ({
      id: e.id, start: Number(e.start), end: Number(e.end), duration: Number(e.duration)
    }));
  } catch(e) {
    Logger.log(`⚠️ Could not fetch time entries for task ${taskId}: ${e.message}`); return [];
  }
}

function formatMsToHumanReadable(ms) {
  const totalMinutes = Math.floor(ms / 60000);
  return `${Math.floor(totalMinutes / 60)}h ${totalMinutes % 60}m`;
}

// =============================================================================
// STEP 0: Delete all ClickUp tracked time entries between two dates
//
// PURPOSE: Clean slate before a re-sync. Useful when you've already pushed
// entries to ClickUp but need to redo them (wrong timestamps, wrong task
// associations, payload bugs like the 403 you hit, etc.).
//
// WHAT IT DOES:
//   1. Reads the clean sheet to get the list of ClickUp task IDs that were
//      targeted during the sync (so it only touches tasks this pipeline owns,
//      not unrelated manually-entered time elsewhere in your workspace)
//   2. For each unique task ID, fetches all time entries currently in ClickUp
//   3. Deletes any entry whose start timestamp falls within START_DATE_OVERRIDE
//      → END_DATE_OVERRIDE (the same window as the rest of the pipeline)
//   4. Resets the sync status column (col 12) in the clean sheet back to
//      'New Entry' for all affected rows, so Steps 3a → 3b can reprocess them
//
// SAFETY GUARDRAILS:
//   • Dry-run mode (DRY_RUN_STEP0 = true) logs everything it WOULD delete
//     without actually deleting anything — always run this first
//   • Only deletes entries within the configured date window, never outside it
//   • Only touches task IDs that appear in column H of the clean sheet
//     (tasks this pipeline wrote to), never the whole workspace
//   • Confirms count before proceeding so you can abort if the number looks wrong
// =============================================================================

const DRY_RUN_STEP0 = false; // ← SET TO FALSE ONLY AFTER CONFIRMING DRY RUN OUTPUT

function runStep0_DeleteTrackedTimeInWindow() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const cuToken  = props.getProperty('CLICKUP_TOKEN');
  const cuTeamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!cuToken || !cuTeamId) { Logger.log('❌ CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.'); return; }

  // Resolve the date window (same logic as the rest of the pipeline)
  if (!START_DATE_OVERRIDE.trim() || !END_DATE_OVERRIDE.trim()) {
    Logger.log('❌ START_DATE_OVERRIDE and END_DATE_OVERRIDE must both be set for Step 0.');
    Logger.log('   Step 0 refuses to run without explicit date boundaries to prevent accidental mass deletion.');
    return;
  }
  const windowStartMs = new Date(START_DATE_OVERRIDE).getTime();
  const windowEndMs   = new Date(END_DATE_OVERRIDE).getTime();
  Logger.log(`🗑️ Step 0: Delete tracked time window: ${START_DATE_OVERRIDE} ──> ${END_DATE_OVERRIDE}`);
  Logger.log(DRY_RUN_STEP0 ? '🔍 DRY RUN MODE — nothing will be deleted' : '⚠️ LIVE MODE — deletions are permanent');

  // Open the workbook and clean sheet
  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) {
    Logger.log(`❌ Workbook "${BASE_FILENAME}" not found. Run Step 1 first so the pipeline knows which tasks to clean.`);
    return;
  }
  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) {
    Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing. Run Steps 1 and 2 first.`);
    return;
  }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 2) {
    Logger.log('🏁 Step 0: Clean sheet is empty — nothing to clean up.');
    return;
  }

  // Collect unique ClickUp task IDs from column H (index 7) — only tasks
  // the pipeline previously targeted, to scope the deletion safely
  const targetTaskIds = new Set();
  for (let i = 2; i < cleanData.length; i++) {
    const taskId = String(cleanData[i][7] || '').trim();
    if (taskId && taskId !== '' && taskId !== 'Missing Target Task' && taskId !== 'undefined') {
      targetTaskIds.add(taskId);
    }
  }

  Logger.log(`📋 Unique ClickUp task IDs in scope: ${targetTaskIds.size}`);
  if (targetTaskIds.size === 0) {
    Logger.log('🏁 Step 0: No ClickUp task IDs found in clean sheet column H. Run Step 3a first to populate task IDs.');
    return;
  }

  // Phase 1: discover all time entries to delete across all targeted tasks
  const toDelete = []; // [{taskId, entryId, startMs, endMs}]

  for (const taskId of targetTaskIds) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log('⚠️ Approaching runtime limit during discovery phase — halting. Re-run Step 0 to continue.');
      break;
    }

    const entries = fetchLiveClickUpTimeEntriesArray(taskId, cuToken, cuTeamId);
    const inWindow = entries.filter(e => e.start >= windowStartMs && e.start < windowEndMs);

    if (inWindow.length > 0) {
      Logger.log(`📌 Task ${taskId}: ${inWindow.length} entries in window (${entries.length} total on task)`);
      inWindow.forEach(e => toDelete.push({ taskId, entryId: e.id, startMs: e.start, endMs: e.end }));
    }
    Utilities.sleep(150);
  }

  Logger.log(`\n📊 SUMMARY: Found ${toDelete.length} time entries to delete within the window`);
  toDelete.forEach(e => {
    const start = Utilities.formatDate(new Date(e.startMs), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const end   = Utilities.formatDate(new Date(e.endMs),   Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    Logger.log(`  ${DRY_RUN_STEP0 ? '[DRY RUN] Would delete' : 'Deleting'}: task=${e.taskId} entry=${e.entryId} | ${start} → ${end}`);
  });

  if (DRY_RUN_STEP0) {
    Logger.log(`\n✅ DRY RUN complete. ${toDelete.length} entries would be deleted.`);
    Logger.log('   Set DRY_RUN_STEP0 = false at the top of the script and re-run to execute.');
    return;
  }

  // Phase 2: execute deletions
  let deletedCount = 0, errorCount = 0;

  for (const entry of toDelete) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Runtime limit hit mid-deletion. Deleted ${deletedCount} of ${toDelete.length}. Re-run Step 0 to finish.`);
      break;
    }

    const url = `https://api.clickup.com/api/v2/team/${cuTeamId}/time_entries/${entry.entryId}`;
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'delete',
        headers: { 'Authorization': cuToken, 'Content-Type': 'application/json' },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code === 200 || code === 204) {
        deletedCount++;
        Logger.log(`🗑️ Deleted entry ${entry.entryId} from task ${entry.taskId}`);
      } else {
        errorCount++;
        Logger.log(`🚨 Failed to delete entry ${entry.entryId} (HTTP ${code}): ${response.getContentText().substring(0, 150)}`);
      }
    } catch(e) {
      errorCount++;
      Logger.log(`❌ Network error deleting entry ${entry.entryId}: ${e.message}`);
    }
    Utilities.sleep(200);
  }

  // Phase 3: reset sync status in clean sheet back to 'New Entry' so Steps
  // 3a → 3b will reprocess all rows on the next run
  Logger.log('\n🔄 Resetting sync status column in clean sheet back to "New Entry"...');
  let resetCount = 0;
  for (let i = 2; i < cleanData.length; i++) {
    const currentStatus = String(cleanData[i][11] || '').trim();
    if (currentStatus === 'Synced Successfully') {
      cleanSheet.getRange(i + 1, 12).setValue('New Entry').setBackground(null).setFontColor(null);
      resetCount++;
    }
  }
  SpreadsheetApp.flush();

  Logger.log(`\n🏁 Step 0 complete.`);
  Logger.log(`   Deleted: ${deletedCount} | Errors: ${errorCount} | Clean sheet rows reset: ${resetCount}`);
  Logger.log(`   Next: run Step 3a → Step 3b to re-sync cleanly.`);
}

// =============================================================================
// STEP 0b: Delete ONLY duplicate time entries — leaves the first copy intact
//
// Unlike Step 0 which deletes everything in the window, this function:
//   1. Fetches all time entries team-wide for the configured date window
//   2. Groups them by task
//   3. Sorts each group by entry ID ascending (oldest first as a proxy for
//      creation order) — the first entry per overlap group is kept
//   4. Flags later entries whose start/end overlaps an already-kept entry by
//      more than 5 seconds → these are the duplicates to remove
//   5. Deletes ONLY the duplicates — the original entry on each task is
//      preserved exactly as-is
//
// After running this you do NOT need to re-run Steps 1, 2, or 3a.
// The clean sheet already has the correct entries marked Synced Successfully.
// Step 3b's new live duplicate check will prevent this from recurring.
//
// DRY_RUN_STEP0b = true  → logs what would be deleted, no actual deletes
// DRY_RUN_STEP0b = false → performs the deletes
// =============================================================================

function runStep0b_DeleteDuplicatesOnly() {
  const t0 = Date.now();
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('CLICKUP_TOKEN');
  const teamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!token || !teamId) { Logger.log('❌ CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.'); return; }

  if (!START_DATE_OVERRIDE.trim() || !END_DATE_OVERRIDE.trim()) {
    Logger.log('❌ START_DATE_OVERRIDE and END_DATE_OVERRIDE must both be set.'); return;
  }

  // Toggle via setDryRunStep0b_Live() / setDryRunStep0b_DryRun() — no code edits needed
  const dryRun = props.getProperty('DRY_RUN_STEP0b') !== 'false';

  const windowStartMs = new Date(START_DATE_OVERRIDE + 'T00:00:00Z').getTime();
  const windowEndMs   = new Date(END_DATE_OVERRIDE   + 'T23:59:59Z').getTime();

  Logger.log(`🔍 Duplicate-only cleanup: ${START_DATE_OVERRIDE} → ${END_DATE_OVERRIDE}`);
  Logger.log(dryRun ? '🔍 DRY RUN — nothing will be deleted' : '⚠️  LIVE — deletions are permanent');

  const allEntries = fetchAllTimeEntriesTeamWide(teamId, token, windowStartMs, windowEndMs, t0);
  Logger.log(`Total entries in window: ${allEntries.length}`);
  if (allEntries.length === 0) { Logger.log('Nothing to process.'); return; }

  // Group by task
  const byTask = {};
  for (const e of allEntries) {
    const tid = e.task_id || 'no_task';
    if (!byTask[tid]) byTask[tid] = { name: e.task_name || '(no task)', entries: [] };
    byTask[tid].entries.push(e);
  }

  // Identify duplicates within each task group
  const toDelete = [];
  for (const [tid, { name, entries }] of Object.entries(byTask)) {
    // Sort by entry ID ascending — lower ID = created first = keep it
    const sorted = entries.slice().sort((a, b) => String(a.id).localeCompare(String(b.id)));
    const kept   = [];
    for (const e of sorted) {
      const isDup = kept.some(k => Math.min(e.end, k.end) - Math.max(e.start, k.start) > 5000);
      if (isDup) {
        toDelete.push({ entryId: e.id, taskId: tid, taskName: name, startMs: e.start, endMs: e.end });
      } else {
        kept.push(e);
      }
    }
  }

  Logger.log(`\n📊 ${toDelete.length} duplicates found across ${Object.keys(byTask).length} tasks`);
  if (toDelete.length === 0) { Logger.log('✅ No duplicates — nothing to delete.'); return; }

  toDelete.forEach(d =>
    Logger.log(`  ${dryRun ? '[DRY RUN]' : 'DELETE'} entry=${d.entryId} | ${d.taskName} | ${fmtTs(d.startMs)} → ${fmtTs(d.endMs)}`)
  );

  if (dryRun) {
    Logger.log(`\n✅ Dry run complete. Run setDryRunStep0b_Live() then re-run to delete.`);
    return;
  }

  // Delete duplicates only
  let deletedCount = 0, errorCount = 0;
  for (const d of toDelete) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Runtime limit — deleted ${deletedCount}/${toDelete.length}. Re-run to finish.`); break;
    }
    try {
      const res  = UrlFetchApp.fetch(`https://api.clickup.com/api/v2/team/${teamId}/time_entries/${d.entryId}`, {
        method: 'delete', headers: { 'Authorization': token }, muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      if (code === 200 || code === 204) { deletedCount++; Logger.log(`🗑️ Deleted ${d.entryId}`); }
      else { errorCount++; Logger.log(`🚨 Failed ${d.entryId} (HTTP ${code})`); }
    } catch(ex) { errorCount++; Logger.log(`❌ Error deleting ${d.entryId}: ${ex.message}`); }
    Utilities.sleep(300);
  }

  Logger.log(`\n✅ Step 0b complete. Deleted: ${deletedCount}, Errors: ${errorCount}`);
  Logger.log('   Original entries are untouched. No need to re-run Steps 1–3a.');
}

/** Run this to enable LIVE deletion mode for Step 0b, then run runStep0b_DeleteDuplicatesOnly() */
function setDryRunStep0b_Live() {
  PropertiesService.getScriptProperties().setProperty('DRY_RUN_STEP0b', 'false');
  Logger.log('⚠️ Step 0b is now in LIVE mode — next run will DELETE duplicates permanently.');
}

/** Run this to revert Step 0b to safe dry-run mode */
function setDryRunStep0b_DryRun() {
  PropertiesService.getScriptProperties().setProperty('DRY_RUN_STEP0b', 'true');
  Logger.log('✅ Step 0b is back in DRY RUN mode — safe to run without deleting anything.');
}

// =============================================================================
// STEP 4: Write Rize IDs to ClickUp Custom Fields
//
// Reads columns E (rize_time_entry_id) and F (rize_task_id) from the clean
// sheet and writes them into two custom fields you created on each ClickUp task.
//
// SETUP:
//   1. Run discoverCustomFieldIds() first — it logs the field IDs for every
//      custom field on your tasks so you can find the right ones
//   2. Paste the two field IDs into the constants below
//   3. Run runStep4_WriteRizeIdsToCustomFields()
//
// MULTI-ENTRY HANDLING: multiple Rize time entries can map to the same ClickUp
// task. For rize_time_entry_id, all IDs are comma-separated into one field
// value. For rize_task_id, they should all be the same value per task
// (it's the Rize task, not the time entry) — the first non-N/A value is used.
// =============================================================================



/**
 * Diagnostic: logs all custom fields and their IDs for every task in the
 * clean sheet. Run this once to find the field IDs to paste above.
 */
function discoverCustomFieldIds() {
  const props = PropertiesService.getScriptProperties();
  const cuToken = props.getProperty('CLICKUP_TOKEN');
  if (!cuToken) { Logger.log('❌ CLICKUP_TOKEN missing.'); return; }

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) { Logger.log('❌ Workbook not found. Run Step 1 first.'); return; }

  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) { Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return; }

  const cleanData = cleanSheet.getDataRange().getValues();

  // Collect unique task IDs from column H (index 7)
  const taskIds = new Set();
  for (let i = 2; i < cleanData.length; i++) {
    const id = String(cleanData[i][7] || '').trim();
    if (id && id !== '' && id !== 'Missing Target Task' && id !== 'undefined') {
      taskIds.add(id);
    }
  }

  Logger.log(`🔍 Checking custom fields on ${taskIds.size} unique tasks...`);
  const seenFieldIds = new Set();

  for (const taskId of taskIds) {
    const url = `https://api.clickup.com/api/v2/task/${taskId}?custom_fields=true`;
    try {
      const res  = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': cuToken },
        muteHttpExceptions: true
      });
      const data = JSON.parse(res.getContentText());
      if (!data.custom_fields || data.custom_fields.length === 0) continue;

      data.custom_fields.forEach(f => {
        if (!seenFieldIds.has(f.id)) {
          seenFieldIds.add(f.id);
          Logger.log(`📌 Task: ${data.name}`);
          Logger.log(`   Field Name: "${f.name}"`);
          Logger.log(`   Field ID:   ${f.id}`);
          Logger.log(`   Type:       ${f.type}`);
          Logger.log(`   Value:      ${JSON.stringify(f.value)}`);
          Logger.log('   ─────────────────────────────────────');
        }
      });
    } catch(e) {
      Logger.log(`⚠️ Could not fetch fields for task ${taskId}: ${e.message}`);
    }
    Utilities.sleep(150);
  }

  Logger.log(`✅ Discovery complete. Found ${seenFieldIds.size} unique custom fields.`);
  Logger.log('   Copy the Field IDs above, then save them to Script Properties:');
  Logger.log('   PropertiesService.getScriptProperties().setProperties({');
  Logger.log('     CUSTOM_FIELD_ID_RIZE_TIME_ENTRY: "paste-id-here",');
  Logger.log('     CUSTOM_FIELD_ID_RIZE_TASK: "paste-id-here"');
  Logger.log('   });');
}

/**
 * STEP 4: Write rize_time_entry_id and rize_task_id values from the clean
 * sheet into the two ClickUp custom fields you created.
 */
function runStep4_WriteRizeIdsToCustomFields() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const cuToken = props.getProperty('CLICKUP_TOKEN');
  if (!cuToken) { Logger.log('❌ CLICKUP_TOKEN missing.'); return; }

  const fieldIdTimeEntry = props.getProperty('CUSTOM_FIELD_ID_RIZE_TIME_ENTRY') || '';
  const fieldIdTask      = props.getProperty('CUSTOM_FIELD_ID_RIZE_TASK')       || '';

  if (!fieldIdTimeEntry || !fieldIdTask) {
    Logger.log('❌ Custom field IDs not set in Script Properties.');
    Logger.log('   Run discoverCustomFieldIds(), then save the IDs:');
    Logger.log('   PropertiesService.getScriptProperties().setProperties({');
    Logger.log('     CUSTOM_FIELD_ID_RIZE_TIME_ENTRY: "paste-id-here",');
    Logger.log('     CUSTOM_FIELD_ID_RIZE_TASK: "paste-id-here"');
    Logger.log('   });');
    return;
  }

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) { Logger.log('❌ Workbook not found. Run Step 1 first.'); return; }

  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) { Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return; }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 2) { Logger.log('🏁 Step 4: No records in clean sheet.'); return; }

  // Build a map: clickup_task_id → { rizeTimeEntryIds: [], rizeTaskId: '' }
  // Groups all Rize entries that share the same ClickUp task
  const taskMap = {};
  for (let i = 2; i < cleanData.length; i++) {
    const row             = cleanData[i];
    const rizeEntryId     = String(row[4] || '').trim(); // col E
    const rizeTaskId      = String(row[5] || '').trim(); // col F
    const clickupTaskId   = String(row[7] || '').trim(); // col H

    if (!clickupTaskId || clickupTaskId === 'Missing Target Task' || clickupTaskId === '') continue;
    if (!rizeEntryId || rizeEntryId === '—' || rizeEntryId === 'N/A') continue;

    if (!taskMap[clickupTaskId]) {
      taskMap[clickupTaskId] = { rizeTimeEntryIds: [], rizeTaskId: '' };
    }

    taskMap[clickupTaskId].rizeTimeEntryIds.push(rizeEntryId);

    // rize_task_id should be the same for all entries on this task —
    // take the first non-N/A value
    if (!taskMap[clickupTaskId].rizeTaskId && rizeTaskId && rizeTaskId !== 'N/A') {
      taskMap[clickupTaskId].rizeTaskId = rizeTaskId;
    }
  }

  const totalTasks = Object.keys(taskMap).length;
  Logger.log(`📋 Writing Rize IDs to ${totalTasks} unique ClickUp tasks...`);

  let successCount = 0, errorCount = 0;

  for (const [clickupTaskId, data] of Object.entries(taskMap)) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log('⚠️ Approaching runtime limit — halting. Re-run Step 4 to continue.');
      break;
    }

    // Comma-separate multiple Rize time entry IDs for tasks with multiple entries
    const entryIdValue = data.rizeTimeEntryIds.join(', ');
    const taskIdValue  = data.rizeTaskId;

    Logger.log(`📝 Task ${clickupTaskId}: entry_ids="${entryIdValue}" task_id="${taskIdValue}"`);

    // Write rize_time_entry_id custom field
    const entryResult = writeCustomField(clickupTaskId, fieldIdTimeEntry, entryIdValue, cuToken);
    if (!entryResult.ok) {
      Logger.log(`🚨 Failed to write rize_time_entry_id on task ${clickupTaskId}: ${entryResult.error}`);
      errorCount++;
    }
    Utilities.sleep(150);

    // Write rize_task_id custom field
    if (taskIdValue) {
      const taskResult = writeCustomField(clickupTaskId, fieldIdTask, taskIdValue, cuToken);
      if (!taskResult.ok) {
        Logger.log(`🚨 Failed to write rize_task_id on task ${clickupTaskId}: ${taskResult.error}`);
        errorCount++;
      }
      Utilities.sleep(150);
    }

    if (entryResult.ok) successCount++;
  }

  Logger.log(`🏁 Step 4 complete. Tasks updated: ${successCount}, Errors: ${errorCount}`);
}

/**
 * Writes a value to a single ClickUp custom field on a task.
 */
function writeCustomField(taskId, fieldId, value, token) {
  const url = `https://api.clickup.com/api/v2/task/${taskId}/field/${fieldId}`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ value }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 200 || code === 201) return { ok: true };
    return { ok: false, error: `HTTP ${code}: ${res.getContentText().substring(0, 200)}` };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// =============================================================================
// INCREMENTAL MODE — Automated Pipeline Orchestration
//
// HOW THE TRANSITION WORKS:
//   Historical mode:  START_DATE_OVERRIDE + END_DATE_OVERRIDE are set → the
//                     pipeline processes that fixed window each run.
//   Incremental mode: Call switchToIncrementalMode() once → it saves the
//                     current timestamp to Script Properties as the anchor.
//                     From then on, runFullPipeline() uses the stored
//                     "last run end" as the next run's start, and "now" as
//                     the end — so each run picks up exactly where the last
//                     left off with no gaps and no overlaps.
//                     Date override constants in the script are ignored once
//                     incremental mode is active.
//
// RUN ORDER (one-time setup, do this when historical runs are done):
//   1. switchToIncrementalMode()   → marks the transition point
//   2. setupIncrementalTrigger()   → creates the hourly trigger
//
// After that, runFullPipeline() fires automatically every hour and handles
// Steps 1 → 2 → 3a → 3b in sequence. Check the log sheet for results.
//
// TO PAUSE:   removeIncrementalTrigger()
// TO RESUME:  setupIncrementalTrigger()   (picks up from last stored anchor)
// TO RESET:   switchToIncrementalMode()   (resets the anchor to now)
// =============================================================================

const INCREMENTAL_TRIGGER_HOURS = 1;   // how often the pipeline runs (hours)
const INCREMENTAL_LOOKBACK_BUFFER_MINS = 30; // overlap buffer — start each run
                                              // 30 min before last end to catch
                                              // any Rize entries that were still
                                              // being written when last run ended

/**
 * Call this ONCE when your historical backfill is complete.
 * Saves the current timestamp as the incremental anchor and flips the
 * pipeline into rolling-window mode. The date override constants at the
 * top of the script are ignored from this point on.
 */
function switchToIncrementalMode() {
  const props = PropertiesService.getScriptProperties();
  const now   = new Date().toISOString();
  props.setProperty('INCREMENTAL_MODE', 'true');
  props.setProperty('INCREMENTAL_LAST_END', now);
  Logger.log('✅ Switched to INCREMENTAL MODE.');
  Logger.log(`   Anchor set to: ${now}`);
  Logger.log('   Next: run setupIncrementalTrigger() to automate the hourly pipeline.');
}

/**
 * Reverts to manual/historical mode. The date override constants at the top
 * of the script take effect again. Call this if you need to re-process a
 * specific historical window.
 */
function switchToHistoricalMode() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('INCREMENTAL_MODE');
  props.deleteProperty('INCREMENTAL_LAST_END');
  Logger.log('✅ Switched back to HISTORICAL MODE.');
  Logger.log('   Edit START_DATE_OVERRIDE and END_DATE_OVERRIDE at the top of the script, then run steps manually.');
}

/**
 * Creates a time-based trigger that fires runFullPipeline() every
 * INCREMENTAL_TRIGGER_HOURS hours. Safe to call multiple times — removes
 * any existing runFullPipeline triggers before creating a new one.
 */
function setupIncrementalTrigger() {
  // Remove any existing triggers for runFullPipeline to avoid stacking
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFullPipeline') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFullPipeline')
    .timeBased()
    .everyHours(INCREMENTAL_TRIGGER_HOURS)
    .create();
  Logger.log(`✅ Hourly trigger created. runFullPipeline() will fire every ${INCREMENTAL_TRIGGER_HOURS} hour(s).`);
  Logger.log('   To stop: call removeIncrementalTrigger()');
}

/**
 * Removes the automated trigger. The pipeline stops running automatically
 * but all data and state are preserved — setupIncrementalTrigger() resumes
 * from where it left off.
 */
function removeIncrementalTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFullPipeline') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(`✅ Removed ${removed} trigger(s). Pipeline is paused.`);
}

/**
 * Shows current pipeline mode and trigger status.
 */
function checkPipelineStatus() {
  const props       = PropertiesService.getScriptProperties();
  const isIncremental = props.getProperty('INCREMENTAL_MODE') === 'true';
  const lastEnd     = props.getProperty('INCREMENTAL_LAST_END');
  const triggers    = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'runFullPipeline');
  const histStart   = props.getProperty('HIST_SYNC_START_IDX');

  Logger.log('=== Pipeline Status ===');
  Logger.log(`Mode:            ${isIncremental ? '🔄 INCREMENTAL (rolling window)' : '📅 HISTORICAL (fixed date range)'}`);
  if (isIncremental) {
    Logger.log(`Last run ended:  ${lastEnd || 'not set — run switchToIncrementalMode() first'}`);
  } else {
    Logger.log(`Date window:     ${START_DATE_OVERRIDE} → ${END_DATE_OVERRIDE}`);
    Logger.log(`Batch position:  ${histStart || 'not in progress'}`);
  }
  Logger.log(`Trigger active:  ${triggers.length > 0 ? `✅ Yes (every ${INCREMENTAL_TRIGGER_HOURS}h)` : '⛔ No trigger set'}`);
  Logger.log('=======================');
}

/**
 * Master orchestrator — runs all 4 steps in sequence.
 * In incremental mode: resolves the date window from Script Properties
 * (last run end → now) and updates the anchor on completion.
 * In historical mode: uses the date override constants at the top of
 * the script, same as running steps manually.
 *
 * Each step is time-budget-aware and will halt gracefully if the
 * 5.5-minute ceiling is approached — the next trigger invocation picks up.
 */
function runFullPipeline() {
  const t0    = Date.now();
  const props = PropertiesService.getScriptProperties();
  const isIncremental = props.getProperty('INCREMENTAL_MODE') === 'true';

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`🚀 runFullPipeline() — ${isIncremental ? 'INCREMENTAL' : 'HISTORICAL'} MODE`);
  Logger.log(`   Started: ${new Date().toISOString()}`);
  Logger.log('='.repeat(60));

  // ── Resolve the window for this run ────────────────────────────────────────
  if (isIncremental) {
    const rawLastEnd  = props.getProperty('INCREMENTAL_LAST_END');
    if (!rawLastEnd) {
      Logger.log('❌ INCREMENTAL_LAST_END not set. Run switchToIncrementalMode() first.');
      return;
    }
    // Step back by the buffer to catch any late-arriving Rize entries
    const windowStartMs = new Date(rawLastEnd).getTime() - (INCREMENTAL_LOOKBACK_BUFFER_MINS * 60 * 1000);
    const windowEndMs   = Date.now();
    // Temporarily inject these as script properties so the step functions
    // can read them without touching the top-level constants
    props.setProperty('PIPELINE_START_ISO', new Date(windowStartMs).toISOString());
    props.setProperty('PIPELINE_END_ISO',   new Date(windowEndMs).toISOString());
    Logger.log(`📅 Window: ${new Date(windowStartMs).toISOString()} → ${new Date(windowEndMs).toISOString()}`);
  }
  // In historical mode the step functions read START/END_DATE_OVERRIDE directly

  // ── Step 1 ─────────────────────────────────────────────────────────────────
  Logger.log('\n── STEP 1: Extract Rize → Staging ──');
  try {
    runStep1_WithWindow(props, isIncremental, t0);
  } catch(e) {
    Logger.log(`❌ Step 1 crashed: ${e.message}`); return;
  }
  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⏸ Time budget reached after Step 1 — will continue next trigger invocation.'); return;
  }

  // ── Step 2 ─────────────────────────────────────────────────────────────────
  Logger.log('\n── STEP 2: Transform → Clean Sheet ──');
  try {
    runStep2_CleanAndTransformStaging();
  } catch(e) {
    Logger.log(`❌ Step 2 crashed: ${e.message}`); return;
  }
  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⏸ Time budget reached after Step 2 — will continue next trigger invocation.'); return;
  }

  // ── Step 3a ────────────────────────────────────────────────────────────────
  Logger.log('\n── STEP 3a: Diagnostic Match ──');
  try {
    runStep3a_DiagnosticClickUpMatcher();
  } catch(e) {
    Logger.log(`❌ Step 3a crashed: ${e.message}`); return;
  }
  if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
    Logger.log('⏸ Time budget reached after Step 3a — will continue next trigger invocation.'); return;
  }

  // ── Step 3b ────────────────────────────────────────────────────────────────
  Logger.log('\n── STEP 3b: Sync to ClickUp ──');
  try {
    runStep3b_ActiveClickUpSyncLoader();
  } catch(e) {
    Logger.log(`❌ Step 3b crashed: ${e.message}`); return;
  }

  // ── Update the anchor on successful completion ──────────────────────────
  if (isIncremental) {
    const newAnchor = props.getProperty('PIPELINE_END_ISO');
    props.setProperty('INCREMENTAL_LAST_END', newAnchor);
    Logger.log(`\n✅ Pipeline complete. Incremental anchor advanced to: ${newAnchor}`);
    // Clean up temp properties
    props.deleteProperty('PIPELINE_START_ISO');
    props.deleteProperty('PIPELINE_END_ISO');
  } else {
    Logger.log('\n✅ Historical pipeline run complete.');
  }

  Logger.log(`   Total runtime: ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

/**
 * Internal: Step 1 variant that reads the window from Script Properties
 * when in incremental mode, instead of the top-level constants.
 * This avoids having to change constants for each mode switch.
 */
function runStep1_WithWindow(props, isIncremental, t0) {
  const rizeKey = props.getProperty('RIZE_API_KEY');
  if (!rizeKey) { Logger.log('❌ RIZE_API_KEY missing.'); return; }

  let startISO, endISO;
  if (isIncremental) {
    startISO = props.getProperty('PIPELINE_START_ISO');
    endISO   = props.getProperty('PIPELINE_END_ISO');
  } else {
    startISO = new Date(START_DATE_OVERRIDE).toISOString();
    endISO   = new Date(END_DATE_OVERRIDE).toISOString();
  }

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const ss = getOrCreateWorkbook(destinationFolder, BASE_FILENAME);

  const alreadyProcessedIds = loadAlreadyProcessedRizeIds(ss);
  Logger.log(`📚 Already staged: ${alreadyProcessedIds.size} entries`);

  const rawNodes = fetchRizeLookbackDataset(rizeKey, startISO, endISO, t0);
  Logger.log(`📡 Rize returned ${rawNodes.length} entries`);

  const newNodes = rawNodes.filter(n => !alreadyProcessedIds.has(String(n.id)));
  Logger.log(`✨ New entries to stage: ${newNodes.length}`);

  commitRawDatasetToSheet(ss, 'Rize_Stage_Raw', newNodes);
}

// =============================================================================
// STEP 5: Inject Subtask Mapping directly into the clean sheet
//
// PURPOSE: The 134 Rize entries grouped into 21 subtasks are already in the
// clean sheet but marked "Missing ClickUp Task" because their Rize task names
// don't match the new ClickUp subtask names. This function bypasses the normal
// name-matching flow and writes the correct ClickUp subtask IDs directly into
// column H for each Rize entry ID, then resets column L to "New Entry" so
// Step 3b will pick them up and post the time entries.
//
// After running this: run Step 3b ONLY (no need for Steps 1, 2, or 3a).
// Step 3b will process only the rows reset to "New Entry".
//
// Rize is never touched — this only updates the Google Sheet and ClickUp.
// =============================================================================
function runStep5_InjectSubtaskMapping() {
  const t0 = Date.now();

  // ── MAPPING: Rize Entry ID → ClickUp Subtask ID ───────────────────────────
  // Built from the analysis of the 5 groups and 21 subtasks created.
  // Key = rize_time_entry_id (column E), Value = clickup subtask task ID (column H)
  const RIZE_TO_SUBTASK = {
    // ── G1: Game Theory Coursework (parent: 86badude9) ────────────────────
    // Canvas Study, Notes & Reading (86bakmbaw)
    '11766382': '86bakmbaw', '11766255': '86bakmbaw', '11766183': '86bakmbaw',
    '11766383': '86bakmbaw', '11766387': '86bakmbaw', '11766388': '86bakmbaw',
    '11766395': '86bakmbaw', '11766410': '86bakmbaw', '11766412': '86bakmbaw',
    '11766441': '86bakmbaw', '11766585': '86bakmbaw', '11766595': '86bakmbaw',
    '11766608': '86bakmbaw', '11766615': '86bakmbaw', '11766622': '86bakmbaw',
    '11767011': '86bakmbaw',
    // Lectures & Class Attendance (86bakmbc4)
    '11766358': '86bakmbc4', '11766225': '86bakmbc4', '11766271': '86bakmbc4',
    '11766293': '86bakmbc4', '11766327': '86bakmbc4', '11766381': '86bakmbc4',
    '11766407': '86bakmbc4', '11766552': '86bakmbc4', '11766588': '86bakmbc4',
    '11766618': '86bakmbc4', '11766645': '86bakmbc4',
    // Quizzes, Assessments & Finals (86bakmbd6)
    '11766280': '86bakmbd6', '11766443': '86bakmbd6', '11766442': '86bakmbd6',
    '11767183': '86bakmbd6', '11766601': '86bakmbd6',
    // Written Assignments & Payoff Matrix Work (86bakmbe7)
    '11766326': '86bakmbe7', '11766330': '86bakmbe7', '11766386': '86bakmbe7',
    '11766387': '86bakmbe7', '11766395': '86bakmbe7',

    // ── G2: Coursera Work (parent: 86bae94j7) ────────────────────────────
    // IBM PM Specialization (86bakmbf8)
    '11766214': '86bakmbf8', '11766217': '86bakmbf8', '11766219': '86bakmbf8',
    '11766220': '86bakmbf8', '11766234': '86bakmbf8', '11766246': '86bakmbf8',
    '11766247': '86bakmbf8', '11766251': '86bakmbf8', '11766216': '86bakmbf8',
    '11766226': '86bakmbf8', '11766329': '86bakmbf8', '11766348': '86bakmbf8',
    '12414122': '86bakmbf8',
    // IBM AI & Cloud Labs (86bakmbfv)
    '11766272': '86bakmbfv',
    // Scalable Innovation Workshop (86bakmbgp)
    '11766270': '86bakmbgp',
    // Azure ML Module (86bakmbjq)
    '11767210': '86bakmbjq',
    // Databricks & Lakeflow (86bakmbkw)
    '11841656': '86bakmbkw',

    // ── G3: Gmail LabelClassifier (parent: 86badxq62) ────────────────────
    // Workspace & Gmail Setup (86bakmbmt)
    '11766515': '86bakmbmt',
    // Core Classifier Development (86bakmbnw)
    '12067519': '86bakmbnw', '12067592': '86bakmbnw', '12067601': '86bakmbnw',
    '12067606': '86bakmbnw', '12067613': '86bakmbnw', '12067620': '86bakmbnw',
    '12067624': '86bakmbnw', '12067635': '86bakmbnw', '12067642': '86bakmbnw',
    '12088019': '86bakmbnw', '12088035': '86bakmbnw', '12088066': '86bakmbnw',
    '12088085': '86bakmbnw', '12088095': '86bakmbnw',
    // Testing, Refinement & Deployment (86bakmbq5)
    '12173388': '86bakmbq5', '13450618': '86bakmbq5', '13472002': '86bakmbq5',

    // ── G4: Notion Migration (parent: 86baduc6p) ─────────────────────────
    // Notion for Client Work Ascension (86bakmbrr)
    '11766028': '86bakmbrr',
    // Plaud to Notion Integration (86bakmbut)
    '11766103': '86bakmbut', '11766635': '86bakmbut', '11766648': '86bakmbut',
    // Notion Journal Reorganizer Script (86bakmbwc)
    '12173374': '86bakmbwc', '12173378': '86bakmbwc', '12173379': '86bakmbwc',
    '12173381': '86bakmbwc', '12173382': '86bakmbwc', '12173383': '86bakmbwc',
    '12173384': '86bakmbwc', '12173385': '86bakmbwc', '12173386': '86bakmbwc',
    '12173387': '86bakmbwc', '12238538': '86bakmbwc', '12238539': '86bakmbwc',
    '12324537': '86bakmbwc',
    // Mission Control Workspace Setup (86bakmbyf)
    '12164732': '86bakmbyf', '12173374': '86bakmbyf',

    // ── G5: DVPO Documentation (parent: 86badudn6) ───────────────────────
    // Separation Agreement & Legal Strategy (86bakmc7f)
    '11766249': '86bakmc7f', '11766610': '86bakmc7f',
    // DVPO & Evidence Organization (86bakmcdg)
    '11761873': '86bakmcdg', '11766607': '86bakmcdg', '11767009': '86bakmcdg',
    '11766613': '86bakmcdg', '11766619': '86bakmcdg', '11766623': '86bakmcdg',
    '11766624': '86bakmcdg', '11766632': '86bakmcdg', '11767133': '86bakmcdg',
    '11767134': '86bakmcdg', '11767164': '86bakmcdg', '11768199': '86bakmcdg',
    '13161761': '86bakmcdg',
    // Device Backups & Message Exports (86bakmch1)
    '11766613': '86bakmch1', '11766619': '86bakmch1', '11766620': '86bakmch1',
    '11766750': '86bakmch1', '11767066': '86bakmch1', '11767165': '86bakmch1',
    '11767166': '86bakmch1', '11767167': '86bakmch1', '11767181': '86bakmch1',
    '11767182': '86bakmch1', '11767185': '86bakmch1', '11767214': '86bakmch1',
    '11767215': '86bakmch1', '11767216': '86bakmch1', '11767288': '86bakmch1',
    '11767486': '86bakmch1', '11767586': '86bakmch1',
    // Rebuttal to Complaint (86bakmcqh)
    '11765944': '86bakmcqh', '11765945': '86bakmcqh', '11766840': '86bakmcqh',
    '11767166': '86bakmcqh', '11767183': '86bakmcqh', '11767215': '86bakmcqh',
    '11767289': '86bakmcqh', '11767587': '86bakmcqh', '11767712': '86bakmcqh',
    '13161761': '86bakmcqh',
    // Court Hearings & Mediation (86bakmcwz)
    '11767132': '86bakmcwz', '11767826': '86bakmcwz',
  };

  Logger.log(`📋 Mapping table loaded: ${Object.keys(RIZE_TO_SUBTASK).length} Rize entry IDs`);

  // ── Open workbook and clean sheet ─────────────────────────────────────────
  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) {
    Logger.log(`❌ Workbook "${BASE_FILENAME}" not found.`); return;
  }
  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) {
    Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return;
  }

  const data = cleanSheet.getDataRange().getValues();
  if (data.length <= 2) {
    Logger.log('🏁 Step 5: Clean sheet is empty.'); return;
  }

  // ── Scan every row, inject subtask ID + reset status ──────────────────────
  // Columns (0-indexed): E=4 rize_time_entry_id, H=7 clickup_task_id, L=11 sync_match_status
  let injected = 0, alreadySynced = 0, notFound = 0;

  for (let i = 2; i < data.length; i++) {
    const row           = data[i];
    const rizeEntryId   = String(row[4] || '').trim();
    const currentStatus = String(row[11] || '').trim();
    const subtaskId     = RIZE_TO_SUBTASK[rizeEntryId];

    if (!subtaskId) { notFound++; continue; }

    // Don't overwrite rows already successfully synced to a DIFFERENT task —
    // only inject into rows that are Missing or not yet synced
    if (currentStatus === 'Synced Successfully' || currentStatus === 'Already Matched') {
      alreadySynced++;
      Logger.log(`⏭ Row ${i+1} (${rizeEntryId}): already synced — skipping`);
      continue;
    }

    // Write subtask ID into col H (1-indexed = col 8)
    cleanSheet.getRange(i + 1, 8).setValue(subtaskId);
    // Reset status in col L (1-indexed = col 12) so Step 3b picks it up
    cleanSheet.getRange(i + 1, 12).setValue('New Entry')
      .setBackground(null).setFontColor(null);
    injected++;
  }

  SpreadsheetApp.flush();

  Logger.log(`\n✅ Step 5 complete.`);
  Logger.log(`   Injected:     ${injected} rows (ready for Step 3b)`);
  Logger.log(`   Already done: ${alreadySynced} rows (skipped)`);
  Logger.log(`   Not in map:   ${notFound} rows (not part of these 5 groups)`);
  Logger.log(`\n   ▶ Next: run runStep3b_ActiveClickUpSyncLoader() to post time entries`);
  Logger.log(`   Total runtime: ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

// =============================================================================
// STEP 5b: Inject mapping for the 12 additional tasks (new categories)
// Same mechanics as Step 5 — run AFTER Step 5, then run Step 3b once.
// =============================================================================
function runStep5b_InjectAdditionalMapping() {
  const t0 = Date.now();

  const RIZE_TO_TASK = {
    // Resume Tailoring & ResumeOS (86bakmr7p)
    '11766170':'86bakmr7p','11766209':'86bakmr7p','11766210':'86bakmr7p','11766213':'86bakmr7p','11766235':'86bakmr7p','11766258':'86bakmr7p','11766259':'86bakmr7p','11766260':'86bakmr7p','11766266':'86bakmr7p','11766286':'86bakmr7p','11766287':'86bakmr7p','11766289':'86bakmr7p','11766296':'86bakmr7p','11766297':'86bakmr7p','11766298':'86bakmr7p','11766299':'86bakmr7p','11766305':'86bakmr7p','11766306':'86bakmr7p','11766307':'86bakmr7p','11766314':'86bakmr7p','11766335':'86bakmr7p','11766349':'86bakmr7p','11766363':'86bakmr7p','11766369':'86bakmr7p','11766372':'86bakmr7p','11766373':'86bakmr7p','11766377':'86bakmr7p','11766378':'86bakmr7p','11766384':'86bakmr7p','11766385':'86bakmr7p','11766391':'86bakmr7p','11766392':'86bakmr7p','11766401':'86bakmr7p','11766403':'86bakmr7p','11766414':'86bakmr7p','11766417':'86bakmr7p','11766425':'86bakmr7p','11766431':'86bakmr7p','11766432':'86bakmr7p','11766433':'86bakmr7p','11766437':'86bakmr7p','11766438':'86bakmr7p','11766439':'86bakmr7p','11766444':'86bakmr7p','11766507':'86bakmr7p','11766508':'86bakmr7p','11766530':'86bakmr7p','11766531':'86bakmr7p','11766534':'86bakmr7p','11766536':'86bakmr7p','11766537':'86bakmr7p','11766547':'86bakmr7p','11766556':'86bakmr7p','11766566':'86bakmr7p','11766567':'86bakmr7p','11766576':'86bakmr7p','11766577':'86bakmr7p','11766578':'86bakmr7p','11766628':'86bakmr7p','11766629':'86bakmr7p','11766630':'86bakmr7p','11766641':'86bakmr7p','11766925':'86bakmr7p','11766926':'86bakmr7p','11766927':'86bakmr7p','11767003':'86bakmr7p','11767004':'86bakmr7p','11767005':'86bakmr7p','11767064':'86bakmr7p','11767065':'86bakmr7p','11767130':'86bakmr7p','11767131':'86bakmr7p','11767201':'86bakmr7p','11767204':'86bakmr7p','11767206':'86bakmr7p','11767286':'86bakmr7p','11767287':'86bakmr7p','11768129':'86bakmr7p','11761292':'86bakmr7p','11761281':'86bakmr7p','11761295':'86bakmr7p','11761870':'86bakmr7p',
    // Interview Prep & Practice (86bakmr9u)
    '11766190':'86bakmr9u','11766191':'86bakmr9u','11766533':'86bakmr9u','11766548':'86bakmr9u','11766562':'86bakmr9u','11766591':'86bakmr9u',
    // Recruiter Outreach & Job Research (86bakmray)
    '11766120':'86bakmray','11766357':'86bakmray','11766415':'86bakmray','11766450':'86bakmray','11766553':'86bakmray','11766555':'86bakmray','11766639':'86bakmray','11767062':'86bakmray','11767063':'86bakmray','11767203':'86bakmray','11767395':'86bakmray','11767396':'86bakmray','11767709':'86bakmray',
    // DFD / Ascension Client Delivery (86bakmrc3)
    '11765946':'86bakmrc3','11765953':'86bakmrc3','11765958':'86bakmrc3','11765954':'86bakmrc3','11765955':'86bakmrc3','11765957':'86bakmrc3','11765964':'86bakmrc3','11765961':'86bakmrc3','11765963':'86bakmrc3','11765962':'86bakmrc3','11765982':'86bakmrc3','11765984':'86bakmrc3','11765983':'86bakmrc3','11766023':'86bakmrc3','11766015':'86bakmrc3','11766024':'86bakmrc3','11766016':'86bakmrc3','11766025':'86bakmrc3','11766017':'86bakmrc3','11766018':'86bakmrc3','11766026':'86bakmrc3','11766041':'86bakmrc3','11766027':'86bakmrc3','11766042':'86bakmrc3','11766043':'86bakmrc3','11766029':'86bakmrc3','11766044':'86bakmrc3','11766045':'86bakmrc3','11766046':'86bakmrc3','11766047':'86bakmrc3','11766052':'86bakmrc3','11766053':'86bakmrc3','11766054':'86bakmrc3','11766055':'86bakmrc3','11766056':'86bakmrc3','11766066':'86bakmrc3','11766067':'86bakmrc3','11766068':'86bakmrc3','11766069':'86bakmrc3','11766070':'86bakmrc3','11766089':'86bakmrc3','11766096':'86bakmrc3','11766097':'86bakmrc3','11766098':'86bakmrc3','11766110':'86bakmrc3','11766111':'86bakmrc3','11766112':'86bakmrc3','11766115':'86bakmrc3','11766118':'86bakmrc3','11766132':'86bakmrc3','11766133':'86bakmrc3','11766134':'86bakmrc3','11766137':'86bakmrc3','11766153':'86bakmrc3','11766154':'86bakmrc3','11766155':'86bakmrc3','11766156':'86bakmrc3','11766157':'86bakmrc3','11766163':'86bakmrc3','11766164':'86bakmrc3','11766165':'86bakmrc3','11766166':'86bakmrc3','11766168':'86bakmrc3','11766167':'86bakmrc3','11766173':'86bakmrc3','11766176':'86bakmrc3','11766177':'86bakmrc3','11766184':'86bakmrc3','11766188':'86bakmrc3','11766196':'86bakmrc3','11766197':'86bakmrc3','11766233':'86bakmrc3','11766282':'86bakmrc3','11766424':'86bakmrc3','11766546':'86bakmrc3',
    // Straventis Client Delivery (86bakmrdp)
    '11765990':'86bakmrdp','11765992':'86bakmrdp','11766051':'86bakmrdp','11766109':'86bakmrdp','11766125':'86bakmrdp','11766126':'86bakmrdp','11766127':'86bakmrdp',
    // Straventis Internal Ops & Admin (86bakmrf0)
    '11765981':'86bakmrf0','11766136':'86bakmrf0','11766143':'86bakmrf0','11766144':'86bakmrf0','11766162':'86bakmrf0','11766193':'86bakmrf0','11766334':'86bakmrf0','11766376':'86bakmrf0','11767831':'86bakmrf0','11768136':'86bakmrf0',
    // Mac, Device & Account Setup (86bakmrgb)
    '11765976':'86bakmrgb','11766008':'86bakmrgb','11766035':'86bakmrgb','11766036':'86bakmrgb','11766038':'86bakmrgb','11766059':'86bakmrgb','11766060':'86bakmrgb','11766061':'86bakmrgb','11766077':'86bakmrgb','11766078':'86bakmrgb','11766079':'86bakmrgb','11766088':'86bakmrgb','11766093':'86bakmrgb','11766094':'86bakmrgb','11766149':'86bakmrgb','11766207':'86bakmrgb','11766208':'86bakmrgb','11766242':'86bakmrgb','11766301':'86bakmrgb','11766322':'86bakmrgb','11766323':'86bakmrgb','11766325':'86bakmrgb','11766331':'86bakmrgb','11766332':'86bakmrgb','11766333':'86bakmrgb','11766342':'86bakmrgb','11766343':'86bakmrgb','11766344':'86bakmrgb','11766390':'86bakmrgb','11766423':'86bakmrgb','11766561':'86bakmrgb','11766600':'86bakmrgb','11766627':'86bakmrgb','11766638':'86bakmrgb','11766713':'86bakmrgb','11767184':'86bakmrgb','11767217':'86bakmrgb','11767326':'86bakmrgb','11767374':'86bakmrgb','11767534':'86bakmrgb','11767536':'86bakmrgb','11767555':'86bakmrgb','11767590':'86bakmrgb','11767591':'86bakmrgb','11765212':'86bakmrgb','11768218':'86bakmrgb','11768379':'86bakmrgb','11756667':'86bakmrgb','11768585':'86bakmrgb','11934224':'86bakmrgb','12096949':'86bakmrgb',
    // Personal Healthcare & Medical Admin (86bakmrhq)
    '11765940':'86bakmrhq','11765941':'86bakmrhq','11765943':'86bakmrhq','11765956':'86bakmrhq','11766021':'86bakmrhq','11766032':'86bakmrhq','11766072':'86bakmrhq','11766073':'86bakmrhq','11766087':'86bakmrhq','11766090':'86bakmrhq','11766091':'86bakmrhq','11766092':'86bakmrhq','11766099':'86bakmrhq','11766100':'86bakmrhq','11766138':'86bakmrhq','11766141':'86bakmrhq','11766186':'86bakmrhq','11766187':'86bakmrhq','11766194':'86bakmrhq','11766195':'86bakmrhq','11766223':'86bakmrhq','11766224':'86bakmrhq','11766236':'86bakmrhq','11766238':'86bakmrhq','11766263':'86bakmrhq','11766268':'86bakmrhq','11766274':'86bakmrhq','11766275':'86bakmrhq','11766276':'86bakmrhq','11766284':'86bakmrhq','11766285':'86bakmrhq','11766290':'86bakmrhq','11766291':'86bakmrhq','11766292':'86bakmrhq','11766317':'86bakmrhq','11766318':'86bakmrhq','11766351':'86bakmrhq','11766354':'86bakmrhq','11766355':'86bakmrhq','11766361':'86bakmrhq','11766366':'86bakmrhq','11766404':'86bakmrhq','11766405':'86bakmrhq','11766419':'86bakmrhq','11766420':'86bakmrhq','11766502':'86bakmrhq','11766511':'86bakmrhq','11766525':'86bakmrhq','11766550':'86bakmrhq','11766551':'86bakmrhq','11766572':'86bakmrhq','11766579':'86bakmrhq','11766583':'86bakmrhq','11766586':'86bakmrhq','11766587':'86bakmrhq','11766603':'86bakmrhq','11766605':'86bakmrhq','11766634':'86bakmrhq','11766707':'86bakmrhq','11766931':'86bakmrhq','11767211':'86bakmrhq','11767442':'86bakmrhq','11767465':'86bakmrhq','11767585':'86bakmrhq','11767658':'86bakmrhq','11767710':'86bakmrhq','11767711':'86bakmrhq','11767825':'86bakmrhq','11768131':'86bakmrhq','11768174':'86bakmrhq','11768181':'86bakmrhq','11768185':'86bakmrhq','11768186':'86bakmrhq','11768189':'86bakmrhq','11761872':'86bakmrhq',
    // Personal Finance & Budgeting (86bakmrkr)
    '11765971':'86bakmrkr','11765997':'86bakmrkr','11765998':'86bakmrkr','11766006':'86bakmrkr','11766007':'86bakmrkr','11766074':'86bakmrkr','11766147':'86bakmrkr','11766148':'86bakmrkr','11766330':'86bakmrkr','11766338':'86bakmrkr','11766339':'86bakmrkr','11766541':'86bakmrkr','11766569':'86bakmrkr','11766930':'86bakmrkr',
    // Home, Family & Personal Admin (86bakmrnb)
    '11765972':'86bakmrnb','11765973':'86bakmrnb','11765974':'86bakmrnb','11765975':'86bakmrnb','11765986':'86bakmrnb','11765989':'86bakmrnb','11765993':'86bakmrnb','11765999':'86bakmrnb','11766020':'86bakmrnb','11766031':'86bakmrnb','11766033':'86bakmrnb','11766034':'86bakmrnb','11766049':'86bakmrnb','11766050':'86bakmrnb','11766057':'86bakmrnb','11766058':'86bakmrnb','11766075':'86bakmrnb','11766086':'86bakmrnb','11766121':'86bakmrnb','11766123':'86bakmrnb','11766124':'86bakmrnb','11766128':'86bakmrnb','11766140':'86bakmrnb','11766146':'86bakmrnb','11766158':'86bakmrnb','11766159':'86bakmrnb','11766160':'86bakmrnb','11766171':'86bakmrnb','11766179':'86bakmrnb','11766180':'86bakmrnb','11766182':'86bakmrnb','11766192':'86bakmrnb','11766198':'86bakmrnb','11766199':'86bakmrnb','11766200':'86bakmrnb','11766201':'86bakmrnb','11766203':'86bakmrnb','11766204':'86bakmrnb','11766205':'86bakmrnb','11766206':'86bakmrnb','11766222':'86bakmrnb','11766227':'86bakmrnb','11766228':'86bakmrnb','11766237':'86bakmrnb','11766240':'86bakmrnb','11766250':'86bakmrnb','11766252':'86bakmrnb','11766254':'86bakmrnb','11766264':'86bakmrnb','11766267':'86bakmrnb','11766277':'86bakmrnb','11766278':'86bakmrnb','11766279':'86bakmrnb','11766281':'86bakmrnb','11766283':'86bakmrnb','11766294':'86bakmrnb','11766316':'86bakmrnb','11766319':'86bakmrnb','11766320':'86bakmrnb','11766340':'86bakmrnb','11766341':'86bakmrnb','11766350':'86bakmrnb','11766352':'86bakmrnb','11766359':'86bakmrnb','11766365':'86bakmrnb','11766367':'86bakmrnb','11766368':'86bakmrnb','11766374':'86bakmrnb','11766375':'86bakmrnb','11766379':'86bakmrnb','11766380':'86bakmrnb','11766389':'86bakmrnb','11766396':'86bakmrnb','11766408':'86bakmrnb','11766418':'86bakmrnb','11766421':'86bakmrnb','11766428':'86bakmrnb','11766429':'86bakmrnb','11766503':'86bakmrnb','11766504':'86bakmrnb','11766506':'86bakmrnb','11766510':'86bakmrnb','11766528':'86bakmrnb','11766529':'86bakmrnb','11766540':'86bakmrnb','11766542':'86bakmrnb','11766543':'86bakmrnb','11766557':'86bakmrnb','11766558':'86bakmrnb','11766560':'86bakmrnb','11766570':'86bakmrnb','11766573':'86bakmrnb','11766580':'86bakmrnb','11766598':'86bakmrnb','11766602':'86bakmrnb','11766604':'86bakmrnb','11766617':'86bakmrnb','11766625':'86bakmrnb','11766636':'86bakmrnb','11766642':'86bakmrnb','11766643':'86bakmrnb','11766650':'86bakmrnb','11766652':'86bakmrnb','11766749':'86bakmrnb','11766886':'86bakmrnb','11766894':'86bakmrnb','11766929':'86bakmrnb','11767007':'86bakmrnb','11767008':'86bakmrnb','11767010':'86bakmrnb','11767012':'86bakmrnb','11767255':'86bakmrnb','11767325':'86bakmrnb','11767528':'86bakmrnb','11767550':'86bakmrnb','11767588':'86bakmrnb','11767589':'86bakmrnb','11767657':'86bakmrnb','11767659':'86bakmrnb','11767750':'86bakmrnb','11767824':'86bakmrnb','11767827':'86bakmrnb','11767829':'86bakmrnb','11767910':'86bakmrnb','11767952':'86bakmrnb','11767995':'86bakmrnb','11768051':'86bakmrnb','11768079':'86bakmrnb','11768096':'86bakmrnb','11768097':'86bakmrnb','11768226':'86bakmrnb','11761880':'86bakmrnb','11768348':'86bakmrnb','11768352':'86bakmrnb',
    // Personal Reading & Devotional (86bakmrpx)
    '11766661':'86bakmrpx','11766674':'86bakmrpx','11766682':'86bakmrpx','11766694':'86bakmrpx','11766696':'86bakmrpx','11766754':'86bakmrpx','11766760':'86bakmrpx','11766766':'86bakmrpx','11766770':'86bakmrpx','11766831':'86bakmrpx','11766848':'86bakmrpx','11766849':'86bakmrpx','11766854':'86bakmrpx','11766862':'86bakmrpx','11766869':'86bakmrpx','11766868':'86bakmrpx','11766915':'86bakmrpx','11766921':'86bakmrpx','11766956':'86bakmrpx','11766960':'86bakmrpx','11766928':'86bakmrpx','11766963':'86bakmrpx','11766969':'86bakmrpx','11766977':'86bakmrpx','11766980':'86bakmrpx','11766992':'86bakmrpx','11766993':'86bakmrpx','11766994':'86bakmrpx','11767016':'86bakmrpx','11767034':'86bakmrpx','11767042':'86bakmrpx','11767043':'86bakmrpx','11767050':'86bakmrpx','11767057':'86bakmrpx','11767061':'86bakmrpx','11767117':'86bakmrpx','11767118':'86bakmrpx','11767149':'86bakmrpx','11767158':'86bakmrpx','11767179':'86bakmrpx','11767218':'86bakmrpx','11767235':'86bakmrpx','11767243':'86bakmrpx','11767238':'86bakmrpx','11767240':'86bakmrpx','11767242':'86bakmrpx','11767247':'86bakmrpx','11767248':'86bakmrpx','11767251':'86bakmrpx','11767252':'86bakmrpx','11767253':'86bakmrpx','11767254':'86bakmrpx','11767256':'86bakmrpx','11767257':'86bakmrpx','11767258':'86bakmrpx','11767259':'86bakmrpx','11767264':'86bakmrpx','11767266':'86bakmrpx','11767270':'86bakmrpx','11767281':'86bakmrpx','11767277':'86bakmrpx','11767279':'86bakmrpx','11767291':'86bakmrpx','11767292':'86bakmrpx','11767357':'86bakmrpx','11767517':'86bakmrpx','11765190':'86bakmrpx','11765191':'86bakmrpx','11765192':'86bakmrpx','11765301':'86bakmrpx','11765302':'86bakmrpx','11765303':'86bakmrpx','11765327':'86bakmrpx','11765332':'86bakmrpx','11765334':'86bakmrpx','11765335':'86bakmrpx','11767694':'86bakmrpx','11767705':'86bakmrpx','11767714':'86bakmrpx','11767742':'86bakmrpx','11767797':'86bakmrpx','11767836':'86bakmrpx','11768018':'86bakmrpx','11768026':'86bakmrpx','11768028':'86bakmrpx','11768029':'86bakmrpx','11768062':'86bakmrpx','11768390':'86bakmrpx','11768391':'86bakmrpx','12067559':'86bakmrpx',
    // Rest & Entertainment (86bakmrr4)
    '11766229':'86bakmrr4','11766248':'86bakmrr4','11766253':'86bakmrr4','11766544':'86bakmrr4','11766574':'86bakmrr4','11766582':'86bakmrr4','11766846':'86bakmrr4','11767025':'86bakmrr4','11767175':'86bakmrr4','11767176':'86bakmrr4','11767914':'86bakmrr4','11767987':'86bakmrr4','11768001':'86bakmrr4',
  };

  Logger.log(`📋 Step 5b mapping loaded: ${Object.keys(RIZE_TO_TASK).length} entries`);

  const destinationFolder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const fileSearch = destinationFolder.getFilesByName(BASE_FILENAME);
  if (!fileSearch.hasNext()) { Logger.log(`❌ Workbook "${BASE_FILENAME}" not found.`); return; }

  const ss = SpreadsheetApp.open(fileSearch.next());
  const cleanSheet = ss.getSheetByName(CLEAN_TAB_NAME);
  if (!cleanSheet) { Logger.log(`❌ "${CLEAN_TAB_NAME}" tab missing.`); return; }

  const data = cleanSheet.getDataRange().getValues();
  if (data.length <= 2) { Logger.log('🏁 Step 5b: Clean sheet is empty.'); return; }

  let injected = 0, alreadySynced = 0, notFound = 0;

  for (let i = 2; i < data.length; i++) {
    const row           = data[i];
    const rizeEntryId   = String(row[4] || '').trim();
    const currentStatus = String(row[11] || '').trim();
    const taskId        = RIZE_TO_TASK[rizeEntryId];

    if (!taskId) { notFound++; continue; }
    if (currentStatus === 'Synced Successfully' || currentStatus === 'Already Matched') {
      alreadySynced++; continue;
    }

    cleanSheet.getRange(i + 1, 8).setValue(taskId);
    cleanSheet.getRange(i + 1, 12).setValue('New Entry').setBackground(null).setFontColor(null);
    injected++;
  }

  SpreadsheetApp.flush();
  Logger.log(`✅ Step 5b complete. Injected: ${injected}, Already synced: ${alreadySynced}, Not in map: ${notFound}`);
  Logger.log(`▶ Next: run runStep3b_ActiveClickUpSyncLoader()`);
  Logger.log(`Total runtime: ${((Date.now() - t0) / 1000).toFixed(2)}s`);
}

// =============================================================================
// DIAGNOSTIC: Time Entry Audit
//
// Scans ALL time entries across the entire ClickUp workspace and flags:
//   1. DUPLICATE    — two entries on the same task with identical start + end
//   2. OVERLAP      — two entries on the same task whose time windows overlap
//                     by more than 2 minutes (120 seconds)
//   3. CONFLICTING  — an entry where end time is before start time
//   4. TOO LONG     — an entry whose duration exceeds 4 hours
//
// Results are written to a "Time Entry Audit" tab in the existing
// ClickUp-Rize Sync workbook — one row per issue, color-coded by type.
// Re-running clears and rewrites the tab so it always shows the latest state.
// =============================================================================

const OVERLAP_THRESHOLD_MS  = 2 * 60 * 1000;
const TOO_LONG_THRESHOLD_MS = 4 * 60 * 60 * 1000;

const AUDIT_COLORS = {
  DUPLICATE:   { bg: 'FFE599', fg: '7F6000' }, // yellow
  OVERLAP:     { bg: 'F9CB9C', fg: '7F3F00' }, // orange
  CONFLICTING: { bg: 'EA9999', fg: '660000' }, // red
  'TOO LONG':  { bg: 'A4C2F4', fg: '1C4587' }, // blue
};

function runDiagnostic_TimeEntryAudit() {
  const t0 = Date.now();
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('CLICKUP_TOKEN');
  const teamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!token || !teamId) { Logger.log('❌ CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.'); return; }

  const startMs = new Date(START_DATE_OVERRIDE + 'T00:00:00Z').getTime();
  const endMs   = new Date(END_DATE_OVERRIDE   + 'T23:59:59Z').getTime();

  Logger.log(`🔍 Time Entry Audit — ${START_DATE_OVERRIDE} → ${END_DATE_OVERRIDE}`);
  Logger.log('Fetching all time entries...');

  const allEntries = fetchAllTimeEntriesTeamWide(teamId, token, startMs, endMs, t0);
  Logger.log(`Total entries fetched: ${allEntries.length}`);

  if (allEntries.length === 0) {
    Logger.log('No time entries found in this window.'); return;
  }

  // ── Group by task ─────────────────────────────────────────────────────────
  const byTask = {};
  for (const e of allEntries) {
    const tid = e.task_id || 'no_task';
    if (!byTask[tid]) byTask[tid] = { name: e.task_name || '(no task)', entries: [] };
    byTask[tid].entries.push(e);
  }

  // ── Run checks, collect issues ─────────────────────────────────────────────
  const issues = []; // { type, taskId, taskName, entryIds, start, end, detail }

  for (const [tid, { name, entries }] of Object.entries(byTask)) {
    const sorted = entries.slice().sort((a, b) => a.start - b.start);

    for (const e of sorted) {
      // CONFLICTING
      if (e.end < e.start) {
        issues.push({ type: 'CONFLICTING', taskId: tid, taskName: name,
          entryIds: e.id, start: e.start, end: e.end,
          detail: `End (${fmtTs(e.end)}) is before start (${fmtTs(e.start)})` });
      }
      // TOO LONG
      const dur = e.end - e.start;
      if (dur > TOO_LONG_THRESHOLD_MS) {
        issues.push({ type: 'TOO LONG', taskId: tid, taskName: name,
          entryIds: e.id, start: e.start, end: e.end,
          detail: `Duration: ${fmtDur(dur)}` });
      }
    }

    for (let i = 0; i < sorted.length; i++) {
      for (let j = i + 1; j < sorted.length; j++) {
        const a = sorted[i], b = sorted[j];
        if (b.start > a.end + OVERLAP_THRESHOLD_MS) break;

        if (a.start === b.start && a.end === b.end) {
          issues.push({ type: 'DUPLICATE', taskId: tid, taskName: name,
            entryIds: `${a.id}, ${b.id}`, start: a.start, end: a.end,
            detail: `Identical entries: ${fmtTs(a.start)} → ${fmtTs(a.end)}` });
          continue;
        }

        const overlapMs = Math.min(a.end, b.end) - Math.max(a.start, b.start);
        if (overlapMs > OVERLAP_THRESHOLD_MS) {
          issues.push({ type: 'OVERLAP', taskId: tid, taskName: name,
            entryIds: `${a.id}, ${b.id}`, start: a.start, end: b.end,
            detail: `Overlap: ${fmtDur(overlapMs)} | A: ${fmtTs(a.start)}→${fmtTs(a.end)} | B: ${fmtTs(b.start)}→${fmtTs(b.end)}` });
        }
      }
    }
  }

  // ── Write to sheet ────────────────────────────────────────────────────────
  const sheet = getOrCreateAuditSheet();
  sheet.clearContents();
  sheet.clearFormats();

  // Header
  const headers = ['Type', 'Task ID', 'Task Name', 'Entry ID(s)',
                   'Start', 'End', 'Detail', 'Task URL'];
  const widths  = [14, 14, 38, 26, 20, 20, 70, 42];
  sheet.appendRow(headers);
  const hRow = sheet.getRange(1, 1, 1, headers.length);
  hRow.setFontWeight('bold').setBackground('#2F5597').setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(10);
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);

  if (issues.length === 0) {
    sheet.appendRow(['✅ No issues found', '', '', '', '', '', '', '']);
    sheet.getRange(2, 1, 1, 8).setFontWeight('bold').setFontColor('#274E13')
        .setBackground('#B7E1CD');
  } else {
    // Sort: CONFLICTING first, then OVERLAP, DUPLICATE, TOO LONG
    const order = { CONFLICTING: 0, OVERLAP: 1, DUPLICATE: 2, 'TOO LONG': 3 };
    issues.sort((a, b) => (order[a.type] ?? 9) - (order[b.type] ?? 9));

    const matrix = issues.map(iss => [
      iss.type,
      iss.taskId,
      iss.taskName,
      iss.entryIds,
      iss.start ? fmtTs(iss.start) : '',
      iss.end   ? fmtTs(iss.end)   : '',
      iss.detail,
      `https://app.clickup.com/t/${iss.taskId}`,
    ]);
    sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);

    // Color-code each row by type
    for (let r = 0; r < matrix.length; r++) {
      const type   = matrix[r][0];
      const colors = AUDIT_COLORS[type] || { bg: 'F3F3F3', fg: '000000' };
      const range  = sheet.getRange(r + 2, 1, 1, headers.length);
      range.setBackground('#' + colors.bg).setFontColor('#' + colors.fg)
           .setFontFamily('Arial').setFontSize(10)
           .setVerticalAlignment('top');
      // Make URL column a clickable hyperlink
      const urlCell = sheet.getRange(r + 2, 8);
      urlCell.setFormula(`=HYPERLINK("https://app.clickup.com/t/${matrix[r][1]}","Open task")`);
      // Wrap the detail column
      sheet.getRange(r + 2, 7).setWrap(true);
    }
  }

  sheet.autoResizeColumns(1, 7);
  sheet.setColumnWidth(8, 100); // URL col — fixed

  // Summary row at the bottom
  const counts = { DUPLICATE: 0, OVERLAP: 0, CONFLICTING: 0, 'TOO LONG': 0 };
  issues.forEach(iss => { if (counts[iss.type] !== undefined) counts[iss.type]++; });
  sheet.appendRow(['']);
  sheet.appendRow([
    `TOTAL: ${issues.length} issues`,
    `Duplicates: ${counts.DUPLICATE}`,
    `Overlaps: ${counts.OVERLAP}`,
    `Conflicting: ${counts.CONFLICTING}`,
    `Too Long: ${counts['TOO LONG']}`,
    '',
    `Audit run: ${fmtTs(Date.now())}  |  Entries scanned: ${allEntries.length}  |  Tasks: ${Object.keys(byTask).length}`,
    '',
  ]);
  const sumRow = sheet.getLastRow();
  sheet.getRange(sumRow, 1, 1, 8)
      .setFontWeight('bold').setBackground('#EFEFEF')
      .setFontFamily('Arial').setFontSize(10);

  const summary = `✅ Audit complete — ${issues.length} issues found `
    + `(Dup: ${counts.DUPLICATE}, Overlap: ${counts.OVERLAP}, `
    + `Conflict: ${counts.CONFLICTING}, Long: ${counts['TOO LONG']}). `
    + `Runtime: ${((Date.now() - t0) / 1000).toFixed(1)}s`;
  Logger.log(summary);
}

function getOrCreateAuditSheet() {
  const folder = resolveOrCreateFolderPath(TARGET_FOLDER_PATH);
  const files  = folder.getFilesByName(WORKBOOK_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(WORKBOOK_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  const TAB = 'Time Entry Audit';
  return ss.getSheetByName(TAB) || ss.insertSheet(TAB);
}

// ── Fetch all time entries team-wide with pagination ─────────────────────────
function fetchAllTimeEntriesTeamWide(teamId, token, startMs, endMs, t0) {
  const all = [];
  let page  = 0;
  const PAGE_SIZE = 100;

  while (true) {
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log(`⚠️ Runtime limit hit at page ${page} (${all.length} entries). Re-run to continue.`);
      break;
    }
    const url = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`
      + `?start_date=${startMs}&end_date=${endMs}&page=${page}`;
    try {
      const res  = UrlFetchApp.fetch(url, {
        method: 'get', headers: { 'Authorization': token }, muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      const body = res.getContentText();
      if (code !== 200) { Logger.log(`❌ HTTP ${code} page ${page}: ${body.substring(0, 200)}`); break; }
      const data = JSON.parse(body);
      if (!data.data || data.data.length === 0) break;
      for (const e of data.data) {
        all.push({
          id:        e.id,
          start:     Number(e.start),
          end:       Number(e.end),
          duration:  Number(e.duration),
          task_id:   e.task ? String(e.task.id)   : null,
          task_name: e.task ? String(e.task.name) : null,
        });
      }
      if (data.data.length < PAGE_SIZE) break;
      page++;
      Utilities.sleep(150);
    } catch(ex) {
      Logger.log(`❌ Fetch error page ${page}: ${ex.message}`); break;
    }
  }
  return all;
}

// ── Formatters ────────────────────────────────────────────────────────────────
function fmtTs(ms) {
  if (!ms || isNaN(ms)) return '(invalid)';
  return Utilities.formatDate(new Date(ms), 'America/New_York', 'yyyy-MM-dd HH:mm:ss');
}

function fmtDur(ms) {
  if (!ms || isNaN(ms)) return '?';
  const totalMin = Math.floor(Math.abs(ms) / 60000);
  return `${Math.floor(totalMin / 60)}h ${totalMin % 60}m`;
}

/**
 * Resets pipeline sync state so the next run starts completely fresh.
 * Does NOT touch ClickUp or delete any data — only clears the Script
 * Properties that track pipeline position/anchors.
 */
function resetSync() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('RIZE_SYNC_START_IDX');
  props.deleteProperty('INCREMENTAL_LAST_END');
  props.deleteProperty('INCREMENTAL_MODE');
  Logger.log('✅ Sync state cleared.');
  Logger.log('   RIZE_SYNC_START_IDX, INCREMENTAL_LAST_END, INCREMENTAL_MODE removed.');
  Logger.log('   Next runFullPipeline() or manual Step 1 will start fresh from your date override.');
  Logger.log('   ClickUp data is untouched — Step 3a/3b live checks still prevent duplicates.');
}