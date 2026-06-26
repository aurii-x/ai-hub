// =============================================================================
// GLOBAL USER ENVIRONMENT CONFIGURATION
// =============================================================================
//const START_DATE_OVERRIDE = "2026-01-01"; 
//const END_DATE_OVERRIDE   = "2026-01-31"; 
//const CLEAN_TAB_NAME      = 'Rize_Clean_Sync';
//const TARGET_FOLDER_PATH  = ['AppData', '3.1 clickup-automation']; 
//const BASE_FILENAME       = 'ClickUp-Rize Sync';         
//const INCLUDE_CLOSED_TASKS = true;
//const TIME_BUFFER_MS      = 25000;        // 25-second execution safety margin
//const MAX_RUNTIME_MS      = 330000;       // 5.5 minutes — leaves headroom under 6-min hard limit
//const OVERLAP_TOLERANCE_MS = 5000;        // 5-second tolerance for overlap detection
const SYNC_DELAY_SECONDS  = 1;            // seconds to wait between ClickUp POST calls

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

  // Check if the raw sheet contains the "no new data" placeholder row
  if (rawData.length === 2 && String(rawData[1][0]) === '—') {
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
    const rizeId = String(row[0]);
    if (processedEntryIds.has(rizeId)) continue;

    const startMs    = row[1] ? new Date(row[1]).getTime() : 0;
    const endMs      = row[2] ? new Date(row[2]).getTime() : 0;
    const durationMs = (startMs > 0 && endMs > 0) ? (endMs - startMs) : 0;
    const title      = row[3];
    const desc       = row[4];
    const taskJson   = row[5];

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

  const finalMatrix = masterRecordsList.map(r => [
    r.start, r.end, r.duration, r.description, r.rizeId, r.rizeTaskId, r.taskName,
    ...(r.extraCols || [])
  ]);
  cleanSheet.getRange(3, 1, finalMatrix.length, finalMatrix[0].length).setValues(finalMatrix);

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
      cachedTimeEntries[matchedTaskId] = fetchLiveClickUpTimeEntriesArray(matchedTaskId, cuToken);
    }
    const liveEntries = cachedTimeEntries[matchedTaskId];

    let matchStatus = 'New Entry';
    for (const cuEntry of liveEntries) {
      const isOverlapping = Math.max(startMs, cuEntry.start) < Math.min(endMs, cuEntry.end) + OVERLAP_TOLERANCE_MS;
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

  let successCount = 0, skippedCount = 0, errorCount = 0;

  for (let i = 2; i < cleanData.length; i++) {
    const row = cleanData[i];

    const startMs       = Number(row[0]);
    const endMs         = Number(row[1]);
    const description   = String(row[3]);
    const matchedTaskId = String(row[7]);
    const currentStatus = String(row[11] || '');

    // ── IDEMPOTENCY: only process rows 3a confirmed as new, or prior errors ──
    if (currentStatus !== 'New Entry' && !currentStatus.startsWith('API Error')) {
      skippedCount++;
      continue;
    }

    // Skip rows with no matched task
    if (!matchedTaskId || matchedTaskId === 'Missing Target Task' || matchedTaskId === '') {
      skippedCount++;
      continue;
    }

    // Runtime ceiling check
    if (Date.now() - t0 > MAX_RUNTIME_MS - TIME_BUFFER_MS) {
      Logger.log('⚠️ Approaching runtime ceiling — halting safely. Re-run Step 3b to continue.');
      break;
    }

    // ── UNLIMITED-PLAN-SAFE PAYLOAD: tid + start + stop only ────────────────
    const payload = {
      tid:   matchedTaskId,
      start: startMs,
      stop:  endMs,
    };

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
        cleanSheet.getRange(i + 1, 12).setValue('Synced Successfully').setBackground('#E2EFDA').setFontColor('#375623');
        SpreadsheetApp.flush();
        Logger.log(`✅ Row ${i + 1} synced`);
        Utilities.sleep(SYNC_DELAY_SECONDS * 1000);
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

  Logger.log(`🏁 Step 3b complete. Synced: ${successCount}, Skipped: ${skippedCount}, Errors: ${errorCount}`);
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

function commitRawDatasetToSheet(ss, tabName, nodes) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  else {
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  sheet.clearContents();
  const headers = ['id', 'startTime', 'endTime', 'title', 'description', 'task_json'];
  sheet.appendRow(headers);

  if (nodes.length === 0) {
    sheet.appendRow(['—', 'No new entries to stage in this run', '—', '—', '—', '—']);
    Logger.log('ℹ️ Raw staging sheet updated: no new entries this run.');
    return;
  }

  const matrix = nodes.map(n => [
    n.id, n.startTime, n.endTime, n.title || '', n.description || '',
    n.task ? JSON.stringify(n.task) : ''
  ]);
  sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#7030A0').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`📄 Raw staging tab written: ${nodes.length} new entries`);
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

function fetchLiveClickUpTimeEntriesArray(taskId, token) {
  const url = `https://api.clickup.com/api/v2/time_entries?task_id=${taskId}`;
  try {
    const res     = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': token }, muteHttpExceptions: true });
    const payload = JSON.parse(res.getContentText());
    if (!payload || !payload.data) return [];
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