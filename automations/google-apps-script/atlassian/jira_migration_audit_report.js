/**
 * ClickUp ↔ Jira Migration Audit & Reconciliation
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * PURPOSE
 * Independent verification — does NOT trust the migration pipeline's own
 * self-reported Result/Success columns. Pulls fresh data directly from
 * both ClickUp and Jira APIs and cross-checks them.
 *
 * ARCHITECTURE (deliberate choice, not the default)
 * The two extraction functions do ONLY mechanical data pulling — reusing
 * the same API patterns already proven in the migration pipeline. The
 * actual comparison/matching logic is NOT written in this script — it's
 * built as native Google Sheets FORMULAS (COUNTIF, INDEX/MATCH, IF, ABS)
 * placed into the Comparison tab by buildComparisonTab(). This is
 * deliberate: formulas are inspectable — click any cell and see exactly
 * why it says what it says — instead of trusting opaque script logic a
 * second time on something this high-stakes.
 *
 * TABS PRODUCED
 *   ClickUp Extract   — one row per ClickUp task, master list
 *   Jira Extract      — one row per Jira Task/Sub-task issue, ALL enabled
 *                        fields (not just the ones actively used)
 *   Comparison        — formula-driven cross-check, keyed on ClickUp Task ID
 *
 * JOIN KEY
 * ClickUp Task ID is recovered from Jira's URL(s) field (customfield_10080)
 * via regex matching the distinctive app.clickup.com/t/{team}/{id} pattern
 * — this is the only place a ClickUp identifier exists on the Jira side.
 *
 * ONE THING NOT YET CONFIRMED WITH FRESH EVIDENCE TODAY
 * fields=*all on GET /rest/api/3/issue/{key} is long-standing documented
 * Jira behavior, but I don't have a fresh confirming source from today's
 * research. extractJiraData() logs the raw field list from the FIRST
 * issue fetched — check that log before trusting the full run.
 *
 * CREDENTIALS (Script Properties)
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN
 *   CLICKUP_API_TOKEN, CLICKUP_TEAM_ID
 *   SPREADSHEET_ID       (target Google Sheet — you provide the file ID)
 *   MIGRATION_START_DATE, MIGRATION_END_DATE
 *
 * RUN ORDER
 *   checkSetup() → extractClickUpData() → [manually export Jira CSV,
 *   import as "Jira Raw Export" tab] → transformJiraRawExport() →
 *   buildComparisonTab() → buildDuplicateResolutionPlan() → buildCompletenessAudit()
 */

// ==================== CONFIG ====================

function getProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

//const PROJECT_KEY = 'SOS';
//const CLICKUP_URL_FIELD_ID = 'customfield_10080'; // URL(s) — where ClickUp Task URL was stored

// ==================== SHEET HELPERS ====================

function ss_() { return SpreadsheetApp.openById(getProp_('SPREADSHEET_ID')); }
function getOrCreateTab_(name) {
  const ss = ss_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
function clearAndHeader_(sheet, headers) {
  sheet.clear();
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
}

// ==================== CHECK SETUP ====================

function checkSetup() {
  ['JIRA_SITE_URL', 'JIRA_EMAIL', 'JIRA_API_TOKEN', 'CLICKUP_API_TOKEN', 'CLICKUP_TEAM_ID', 'SPREADSHEET_ID', 'MIGRATION_START_DATE', 'MIGRATION_END_DATE'].forEach(function (key) {
    Logger.log(key + ': ' + (getProp_(key) ? '✅ set' : '❌ NOT SET'));
  });
}

// ==================== JIRA HELPERS ====================

function jiraAuthHeader_() {
  return 'Basic ' + Utilities.base64Encode(getProp_('JIRA_EMAIL') + ':' + getProp_('JIRA_API_TOKEN'));
}
function jiraSiteUrl_() { return getProp_('JIRA_SITE_URL').replace(/\/$/, ''); }
function jiraRequest_(method, path, payload) {
  const options = {
    method: method,
    headers: { 'Authorization': jiraAuthHeader_(), 'Accept': 'application/json' },
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);
  const response = UrlFetchApp.fetch(jiraSiteUrl_() + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  Logger.log('❌ Jira ' + method + ' ' + path + ' → HTTP ' + code + ': ' + body);
  return null;
}

// ==================== CLICKUP HELPERS ====================

function clickUpRequest_(path) {
  const options = { method: 'get', headers: { 'Authorization': getProp_('CLICKUP_API_TOKEN') }, muteHttpExceptions: true };
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  Logger.log('❌ ClickUp GET ' + path + ' → HTTP ' + code + ': ' + body);
  return null;
}

// ==================== PART 1: CLICKUP EXTRACT ====================

function extractClickUpData() {
  Logger.log('=== Extracting ClickUp data ===');
  const sheet = getOrCreateTab_('ClickUp Extract');
  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'Task Name', 'Description', 'Status', 'Priority',
    'Start Date', 'Due Date', 'Time Tracked (hrs)', 'Attachment Count',
    'Rize Task ID', 'Rize Time Entry ID', 'Comment Count', 'Comments (combined)', 'Date Filter Used',
  ]);

  const mappingSheet = ss_().getSheetByName(getProp_('MAPPING_TAB', 'Mapping'));
  const data = mappingSheet.getDataRange().getValues();
  const spaceRows = data.filter(function (row) { return row[0] === 'Space'; });

  const startMs = new Date(getProp_('MIGRATION_START_DATE')).getTime();
  const endMs = new Date(getProp_('MIGRATION_END_DATE')).getTime();
  const teamId = getProp_('CLICKUP_TEAM_ID');
  let rowCount = 0;

  spaceRows.forEach(function (row) {
    const listId = findListId_(teamId, row[1], row[2]);
    if (!listId) { Logger.log('⚠️ Could not find list "' + row[2] + '" in space "' + row[1] + '" — skipping'); return; }

    let page = 0;
    let allTasks = [];
    while (true) {
      const res = clickUpRequest_('/list/' + listId + '/task?archived=false&include_closed=true&subtasks=true&page=' + page);
      if (!res || !res.tasks || res.tasks.length === 0) break;
      allTasks = allTasks.concat(res.tasks);
      if (res.last_page) break;
      page++;
    }

    allTasks.filter(function (t) {
      // Original filter: has a Start Date and it falls in range.
      if (t.start_date) return Number(t.start_date) >= startMs && Number(t.start_date) <= endMs;
      // Fallback: no Start Date at all — use Created Date instead of
      // silently excluding the task entirely.
      return t.date_created && Number(t.date_created) >= startMs && Number(t.date_created) <= endMs;
    })
      .forEach(function (task) {
        const fullTask = clickUpRequest_('/task/' + task.id);
        const attachmentCount = fullTask && fullTask.attachments ? fullTask.attachments.length : 0;

        const timeEntries = clickUpRequest_('/team/' + teamId + '/time_entries?task_id=' + task.id);
        const entries = timeEntries && timeEntries.data ? timeEntries.data : [];
        const totalMs = entries.reduce(function (sum, e) { return sum + (Number(e.duration) || 0); }, 0);
        const timeTrackedHrs = (totalMs / 3600000).toFixed(2);

        const rizeTaskId = getCustomFieldValue_(task, 'Rize Task ID');
        const rizeTimeEntryId = getCustomFieldValue_(task, 'Rize Time Entry ID');

        const commentsRes = clickUpRequest_('/task/' + task.id + '/comment');
        const comments = commentsRes && commentsRes.comments ? commentsRes.comments : [];
        const commentCount = comments.length;
        const commentsCombined = comments.map(function (c) { return c.comment_text || ''; }).join(' | ').substring(0, 2000);

        const dateFilterUsed = task.start_date ? 'Start Date' : 'Created Date (fallback)';

        sheet.appendRow([
          task.id, task.name, (task.description || '').substring(0, 2000),
          task.status && task.status.status, task.priority && task.priority.priority,
          task.start_date ? new Date(Number(task.start_date)) : '', task.due_date ? new Date(Number(task.due_date)) : '',
          timeTrackedHrs, attachmentCount, rizeTaskId, rizeTimeEntryId, commentCount, commentsCombined, dateFilterUsed,
        ]);
        rowCount++;
      });
  });

  Logger.log('✅ ClickUp extract complete: ' + rowCount + ' tasks.');
}

function findListId_(teamId, spaceName, listName) {
  const spacesRes = clickUpRequest_('/team/' + teamId + '/space?archived=false');
  const space = (spacesRes && spacesRes.spaces ? spacesRes.spaces : []).find(function (s) { return s.name === spaceName; });
  if (!space) return null;
  const folderless = clickUpRequest_('/space/' + space.id + '/list?archived=false');
  let list = (folderless && folderless.lists ? folderless.lists : []).find(function (l) { return l.name === listName; });
  if (list) return list.id;
  const folders = clickUpRequest_('/space/' + space.id + '/folder?archived=false');
  (folders && folders.folders ? folders.folders : []).forEach(function (folder) {
    (folder.lists || []).forEach(function (l) { if (l.name === listName) list = l; });
  });
  return list ? list.id : null;
}

function getCustomFieldValue_(task, fieldName) {
  const cf = (task.custom_fields || []).find(function (f) { return f.name === fieldName; });
  return cf && cf.value !== undefined && cf.value !== null ? cf.value : '';
}

// ==================== PART 2: JIRA EXTRACT ====================

// ==================== PART 2 (REPLACED): TRANSFORM MANUAL JIRA CSV EXPORT ====================
// Switched away from the live API extraction — /rest/api/3/search/jql is
// too unreliable (confirmed widely broken as of this year, per multiple
// independent reports). Instead, this reads a manually-exported Jira CSV.
//
// SETUP: In Jira, run: project = SOS (includes Epics — per your
// instruction, issue type is ignored for matching purposes, not filtered
// out). Export → Export Excel CSV (all fields). Import that file into
// THIS spreadsheet as a new tab named "Jira Raw Export" (File → Import →
// Upload → Insert new sheet) — use Sheets' own CSV import, not a manual
// paste, so quoted/comma-containing values parse correctly.
//
// OUTPUT COLUMN ORDER IS DELIBERATELY UNCHANGED from the old live-API
// version, so buildComparisonTab()'s formulas keep working without
// modification — only how "Jira Extract" gets populated changed, not
// its shape.
//
// ONE THING I HAVEN'T VERIFIED: the "Time Spent" column's unit. Jira CSV
// exports sometimes give raw seconds, sometimes a formatted duration
// string like "1h 30m", depending on export settings. parseTimeSpent_()
// below handles both, but check the first few rows of "Time Tracked
// (hrs)" in the output against what you know to be true for a couple of
// issues before trusting the whole sheet.

function transformJiraRawExport() {
  Logger.log('=== Transforming Jira Raw Export ===');
  const rawSheet = ss_().getSheetByName('Jira Raw Export');
  if (!rawSheet) { Logger.log('❌ "Jira Raw Export" tab not found — import your Jira CSV first.'); return; }

  const data = rawSheet.getDataRange().getValues();
  const headers = data[0];

  const idx = {
    summary: findColumnIndex_(headers, 'Summary'),
    issueKey: findColumnIndex_(headers, 'Issue key'),
    issueType: findColumnIndex_(headers, 'Issue Type'),
    status: findColumnIndex_(headers, 'Status'),
    priority: findColumnIndex_(headers, 'Priority'),
    dueDate: findColumnIndex_(headers, 'Due date'),
    description: findColumnIndex_(headers, 'Description'),
    timeSpent: findColumnIndex_(headers, 'Time Spent'),
    clickUpTaskId: findColumnIndex_(headers, 'Custom field (ClickUp Task ID)'),
    parentKey: findColumnIndex_(headers, 'Parent key'),
    urlField: findColumnIndex_(headers, 'Custom field (URL(s))'),
    startDate: findColumnIndex_(headers, 'Custom field (Start date)'),
    rizeId: findColumnIndex_(headers, 'Custom field (Rize ID)'),
    rizeTimeEntryId: findColumnIndex_(headers, 'Custom field (Rize Time Entry ID)'),
  };

  const missing = Object.keys(idx).filter(function (k) { return idx[k] === -1; });
  if (missing.length > 0) {
    Logger.log('⚠️ Could not find these expected columns: ' + missing.join(', ') + ' — check your export included them. Continuing with what was found.');
  }

  const attachmentIndices = findAllColumnIndices_(headers, 'Attachment');
  const commentIndices = findAllColumnIndices_(headers, 'Comment');

  const outSheet = getOrCreateTab_('Jira Extract');
  clearAndHeader_(outSheet, [
    'Jira Issue Key', 'ClickUp Task ID', 'Issue Type', 'Summary', 'Description (plain text)',
    'Status', 'Priority', 'Start Date', 'Due Date', 'Time Tracked (hrs, from worklog)',
    'Attachment Count', 'Rize Task ID', 'Rize Time Entry ID', 'Parent Key', 'Comment Count', 'Comments (combined)', 'All Other Fields (JSON)',
  ]);

  let rowCount = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[idx.issueKey]) continue; // skip blank rows

    // ClickUp Task ID: prefer the direct field, fall back to parsing the URL field.
    // BUG FIX: the dedicated "ClickUp Task ID" field can itself contain a full
    // URL (confirmed in real data), not always a bare ID despite the field
    // name — so the same regex extraction must run on BOTH paths, not just
    // the fallback. Previously the primary path used the raw value as-is,
    // which silently broke every match against ClickUp Extract's bare IDs.
    const idPattern = /app\.clickup\.com\/t\/\d+\/([a-zA-Z0-9]+)/i;
    let clickUpTaskId = '';
    if (idx.clickUpTaskId !== -1 && row[idx.clickUpTaskId]) {
      const raw = String(row[idx.clickUpTaskId]).trim();
      const match = raw.match(idPattern);
      clickUpTaskId = match ? match[1] : raw; // already a bare ID if no URL pattern found
    }
    if (!clickUpTaskId && idx.urlField !== -1 && row[idx.urlField]) {
      const match = String(row[idx.urlField]).match(idPattern);
      if (match) clickUpTaskId = match[1];
    }

    const attachmentCount = attachmentIndices.filter(function (i) { return row[i] && String(row[i]).trim() !== ''; }).length;
    const commentValues = commentIndices.map(function (i) { return row[i]; }).filter(function (v) { return v && String(v).trim() !== ''; });
    const commentCount = commentValues.length;
    const commentsCombined = commentValues.join(' | ').substring(0, 2000);
    const timeSpentHrs = idx.timeSpent !== -1 ? parseTimeSpent_(row[idx.timeSpent]) : '';

    outSheet.appendRow([
      row[idx.issueKey], clickUpTaskId, idx.issueType !== -1 ? row[idx.issueType] : '',
      idx.summary !== -1 ? row[idx.summary] : '', idx.description !== -1 ? row[idx.description] : '',
      idx.status !== -1 ? row[idx.status] : '', idx.priority !== -1 ? row[idx.priority] : '',
      idx.startDate !== -1 ? row[idx.startDate] : '', idx.dueDate !== -1 ? row[idx.dueDate] : '',
      timeSpentHrs, attachmentCount,
      idx.rizeId !== -1 ? row[idx.rizeId] : '', idx.rizeTimeEntryId !== -1 ? row[idx.rizeTimeEntryId] : '',
      idx.parentKey !== -1 ? row[idx.parentKey] : '',
      commentCount, commentsCombined,
      '', // "All Other Fields (JSON)" not populated from CSV — not worth reconstructing from a flat row
    ]);
    rowCount++;
  }

  Logger.log('✅ Transformed ' + rowCount + ' issues into "Jira Extract".');
  Logger.log('⚠️ Spot-check the "Time Tracked (hrs)" column against a couple of real issues — the source unit was not confirmed today.');
}

function findColumnIndex_(headers, name) {
  return headers.indexOf(name);
}
function findAllColumnIndices_(headers, name) {
  const result = [];
  headers.forEach(function (h, i) { if (h === name) result.push(i); });
  return result;
}
function parseTimeSpent_(val) {
  if (!val) return '';
  if (typeof val === 'number') return (val / 3600).toFixed(2); // assume seconds
  const str = String(val).trim();
  if (/^\d+$/.test(str)) return (Number(str) / 3600).toFixed(2); // numeric string, assume seconds
  // Duration string like "1h 30m" or "2d 3h"
  let totalHours = 0;
  const weekMatch = str.match(/(\d+)w/); if (weekMatch) totalHours += Number(weekMatch[1]) * 40; // Jira default 5-day week
  const dayMatch = str.match(/(\d+)d/); if (dayMatch) totalHours += Number(dayMatch[1]) * 8; // Jira default 8hr day
  const hourMatch = str.match(/(\d+)h/); if (hourMatch) totalHours += Number(hourMatch[1]);
  const minMatch = str.match(/(\d+)m/); if (minMatch) totalHours += Number(minMatch[1]) / 60;
  return totalHours > 0 ? totalHours.toFixed(2) : '';
}

function adfToPlainText_(adf) {
  if (!adf || !adf.content) return '';
  let text = '';
  function walk(node) {
    if (node.type === 'text') text += node.text;
    if (node.content) node.content.forEach(walk);
    if (node.type === 'paragraph') text += '\n';
  }
  adf.content.forEach(walk);
  return text.trim();
}

// ==================== PART 3: COMPARISON TAB (FORMULA-DRIVEN) ====================

function buildComparisonTab() {
  Logger.log('=== Building Comparison tab (formulas, not script logic) ===');
  const clickUpSheet = ss_().getSheetByName('ClickUp Extract');
  const jiraSheet = ss_().getSheetByName('Jira Extract');
  if (!clickUpSheet || !jiraSheet) { Logger.log('❌ Run extractClickUpData() and extractJiraData() first.'); return; }

  const numRows = clickUpSheet.getLastRow() - 1; // exclude header
  const sheet = getOrCreateTab_('Comparison');
  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'Task Name', 'Jira Duplicate Count (should = 1)', 'Jira Issue Key',
    'ClickUp Attachments', 'Jira Attachments', 'Attachment Match',
    'ClickUp Time Tracked (hrs)', 'Jira Time Tracked (hrs)', 'Time Match (±3min)',
    'Description Note',
  ]);

  for (let i = 0; i < numRows; i++) {
    const r = i + 2; // sheet row number (1 = header)
    const row = [];
    row.push('=\'ClickUp Extract\'!A' + r);                                     // ClickUp Task ID
    row.push('=\'ClickUp Extract\'!B' + r);                                     // Task Name
    row.push('=COUNTIF(\'Jira Extract\'!B:B, A' + r + ')');                     // Duplicate count — should be 1
    row.push('=IFERROR(INDEX(\'Jira Extract\'!A:A, MATCH(A' + r + ',\'Jira Extract\'!B:B,0)), "NOT FOUND")'); // Jira Issue Key
    row.push('=\'ClickUp Extract\'!I' + r);                                     // ClickUp Attachments
    row.push('=IFERROR(INDEX(\'Jira Extract\'!K:K, MATCH(A' + r + ',\'Jira Extract\'!B:B,0)), "")'); // Jira Attachments
    row.push('=IF(E' + r + '=F' + r + ',"Match","MISMATCH")');
    row.push('=\'ClickUp Extract\'!H' + r);                                     // ClickUp Time Tracked
    row.push('=IFERROR(INDEX(\'Jira Extract\'!J:J, MATCH(A' + r + ',\'Jira Extract\'!B:B,0)), "")'); // Jira Time Tracked
    row.push('=IF(H' + r + '="","",IF(ABS(H' + r + '-I' + r + ')<=0.05,"Pass","FAIL"))'); // ±3min = 0.05hrs
    row.push('="Manual check recommended — ADF conversion may cause minor formatting differences even when content matches"');
    sheet.appendRow(row);
  }

  Logger.log('✅ Comparison tab built with ' + numRows + ' rows of formulas.');
  Logger.log('Check "Jira Duplicate Count" column — every value should be exactly 1.');
  Logger.log('0 = not migrated. 2+ = duplicate creation happened, needs investigation.');
}

// ==================== PART 4: DUPLICATE RESOLUTION PLAN ====================
// Replicates the manual analysis: same-type duplicates get scored on
// completeness (has worklog time > more attachments > longer description);
// mixed-type sets (Epic alongside Task/Sub-task) are flagged separately,
// not auto-scored, since that's a structural difference, not a redundant
// copy. Also flags any delete-candidate that is itself a parent of other
// issues — deleting it would cascade-delete real children.

function buildDuplicateResolutionPlan() {
  Logger.log('=== Building Duplicate Resolution Plan ===');
  const jiraSheet = ss_().getSheetByName('Jira Extract');
  if (!jiraSheet) { Logger.log('❌ "Jira Extract" tab not found — run transformJiraRawExport() first.'); return; }

  const data = jiraSheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  const required = ['Jira Issue Key', 'ClickUp Task ID', 'Issue Type', 'Attachment Count', 'Time Tracked (hrs, from worklog)', 'Description (plain text)', 'Parent Key'];
  const missingCols = required.filter(function (r) { return col[r] === undefined; });
  if (missingCols.length > 0) { Logger.log('❌ Missing required columns: ' + missingCols.join(', ') + ' — re-run transformJiraRawExport() with the updated script.'); return; }

  // Group by ClickUp Task ID
  const byClickUpId = {};
  for (let r = 1; r < data.length; r++) {
    const cid = data[r][col['ClickUp Task ID']];
    if (!cid) continue;
    byClickUpId[cid] = byClickUpId[cid] || [];
    byClickUpId[cid].push({
      key: data[r][col['Jira Issue Key']],
      type: data[r][col['Issue Type']],
      attachments: data[r][col['Attachment Count']] || 0,
      timeHrs: parseFloat(data[r][col['Time Tracked (hrs, from worklog)']]) || 0,
      descLen: (data[r][col['Description (plain text)']] || '').length,
      parentKey: data[r][col['Parent Key']],
    });
  }

  // Find which keys are themselves a parent of other issues
  const parentKeys = {};
  for (let r = 1; r < data.length; r++) {
    const pk = data[r][col['Parent Key']];
    if (pk) parentKeys[pk] = true;
  }

  const dupSets = {};
  Object.keys(byClickUpId).forEach(function (cid) {
    if (byClickUpId[cid].length > 1) dupSets[cid] = byClickUpId[cid];
  });

  const sheet = getOrCreateTab_('Duplicate Resolution Plan');
  clearAndHeader_(sheet, ['ClickUp Task ID', 'Issue Type', 'Jira Key to KEEP', 'Jira Key to DELETE', 'Deleted - Attachments', 'Deleted - Time (hrs)', 'Deleted - Has Children (DO NOT auto-delete)']);

  let sameTypeCount = 0, mixedTypeCount = 0, parentConflictCount = 0;

  Object.keys(dupSets).forEach(function (cid) {
    const candidates = dupSets[cid];
    const types = candidates.map(function (c) { return c.type; });
    const uniqueTypes = types.filter(function (t, i) { return types.indexOf(t) === i; });

    if (uniqueTypes.length > 1) {
      // Mixed type — flag separately, recommend keeping the Epic
      mixedTypeCount++;
      const epic = candidates.find(function (c) { return c.type === 'Epic'; });
      candidates.forEach(function (c) {
        const action = c.type === 'Epic' ? 'KEEP (Epic — structural parent)' : 'MIXED TYPE — do not auto-delete, move work data to a Task/Sub-task under the Epic first';
        sheet.appendRow([cid, c.type, epic ? epic.key : '(no Epic found)', c.type === 'Epic' ? '' : c.key, c.attachments, c.timeHrs, action]);
      });
      return;
    }

    // Same type — score by completeness
    const scored = candidates.slice().sort(function (a, b) {
      if ((b.timeHrs > 0) !== (a.timeHrs > 0)) return (b.timeHrs > 0) - (a.timeHrs > 0);
      if (b.attachments !== a.attachments) return b.attachments - a.attachments;
      return b.descLen - a.descLen;
    });
    const keeper = scored[0];
    scored.slice(1).forEach(function (c) {
      sameTypeCount++;
      const hasChildren = parentKeys[c.key] === true;
      if (hasChildren) parentConflictCount++;
      sheet.appendRow([cid, c.type, keeper.key, c.key, c.attachments, c.timeHrs, hasChildren ? 'YES — has children, do not auto-delete' : '']);
    });
  });

  Logger.log('✅ Plan built. Same-type deletions: ' + sameTypeCount + ' (of which ' + parentConflictCount + ' have children and need manual handling). Mixed-type sets flagged: ' + mixedTypeCount + '.');
  Logger.log('Filter the "Deleted - Has Children" column for anything non-blank before generating a delete list — those need children moved or reconsidered first.');
}

// ==================== PART 5: COMPLETENESS AUDIT ====================
// Checks, for every ClickUp task with a Jira match, whether Description
// and Comments actually carried over — not just "does a Jira issue
// exist" (that's buildComparisonTab()'s job), but "does it actually
// have the content it should."

function buildCompletenessAudit() {
  Logger.log('=== Building Completeness Audit ===');
  const clickSheet = ss_().getSheetByName('ClickUp Extract');
  const jiraSheet = ss_().getSheetByName('Jira Extract');
  if (!clickSheet || !jiraSheet) { Logger.log('❌ Run extractClickUpData() and transformJiraRawExport() first.'); return; }

  const jData = jiraSheet.getDataRange().getValues();
  const jHeaders = jData[0];
  const jCol = {};
  jHeaders.forEach(function (h, i) { jCol[h] = i; });

  const jiraByClickUpId = {};
  for (let r = 1; r < jData.length; r++) {
    const cid = jData[r][jCol['ClickUp Task ID']];
    if (cid) jiraByClickUpId[cid] = jData[r]; // last match wins if duplicates — fine for a completeness spot-check
  }

  const cData = clickSheet.getDataRange().getValues();
  const cHeaders = cData[0];
  const cCol = {};
  cHeaders.forEach(function (h, i) { cCol[h] = i; });

  const sheet = getOrCreateTab_('Completeness Audit');
  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'Task Name', 'Jira Issue Key',
    'ClickUp Has Description', 'Jira Has Description', 'Description Status',
    'ClickUp Comment Count', 'Jira Comment Count', 'Comment Status',
    'ClickUp Attachment Count', 'Jira Attachment Count', 'Attachment Status',
  ]);

  let flaggedCount = 0;

  for (let r = 1; r < cData.length; r++) {
    const cid = cData[r][cCol['ClickUp Task ID']];
    if (!cid) continue;
    const jRow = jiraByClickUpId[cid];
    if (!jRow) continue; // not migrated at all — buildComparisonTab() already covers this case, skip here

    const cDesc = cData[r][cCol['Description']];
    const jDesc = jRow[jCol['Description (plain text)']];
    const cHasDesc = cDesc && String(cDesc).trim() !== '';
    const jHasDesc = jDesc && String(jDesc).trim() !== '';
    let descStatus = 'OK';
    if (cHasDesc && !jHasDesc) descStatus = 'MISSING — ClickUp has it, Jira does not';
    else if (!cHasDesc && jHasDesc) descStatus = 'Jira has content ClickUp does not (unexpected but not data loss)';

    const cCommentCount = Number(cData[r][cCol['Comment Count']]) || 0;
    const jCommentCount = Number(jRow[jCol['Comment Count']]) || 0;
    let commentStatus = 'OK';
    if (cCommentCount > jCommentCount) commentStatus = 'MISSING — ClickUp has ' + cCommentCount + ', Jira has ' + jCommentCount;

    const cAttCount = Number(cData[r][cCol['Attachment Count']]) || 0;
    const jAttCount = Number(jRow[jCol['Attachment Count']]) || 0;
    let attStatus = 'OK';
    if (cAttCount > jAttCount) attStatus = 'MISSING — ClickUp has ' + cAttCount + ', Jira has ' + jAttCount;

    if (descStatus.indexOf('MISSING') === 0 || commentStatus.indexOf('MISSING') === 0 || attStatus.indexOf('MISSING') === 0) flaggedCount++;

    sheet.appendRow([
      cid, cData[r][cCol['Task Name']], jRow[jCol['Jira Issue Key']],
      cHasDesc, jHasDesc, descStatus,
      cCommentCount, jCommentCount, commentStatus,
      cAttCount, jAttCount, attStatus,
    ]);
  }

  Logger.log('✅ Completeness Audit built. ' + flaggedCount + ' task(s) flagged with genuinely MISSING content.');
  Logger.log('Sort/filter by the three Status columns for anything containing "MISSING" to see exactly what did not carry over.');
}