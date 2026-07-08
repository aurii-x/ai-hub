/**
 * ClickUp → Jira Migration Pipeline — santhoshOS
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * FUNCTIONS
 *   discoverClickUpSchema()   — Function 1: crawl ClickUp, write schema tab
 *   discoverJiraSchema()      — Function 2: crawl Jira, write schema tab
 *   dryRunExport()            — Function 3a: build export sheet from mapping, no writes to Jira
 *   liveCreateJiraTasks()     — Function 3b: create issues in Jira from export sheet
 *   resetJiraTasksInRange()  — Function 5: delete everything this pipeline created in range
 *   logRun_()                 — Function 4 (internal): all functions write to the Log tab
 *
 * ALL CONFIG IS IN SCRIPT PROPERTIES — Project Settings → Script Properties:
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN     (reused from earlier scripts)
 *   CLICKUP_API_TOKEN, CLICKUP_TEAM_ID
 *   SPREADSHEET_ID                                 (target Google Sheet)
 *   MIGRATION_START_DATE, MIGRATION_END_DATE       (YYYY-MM-DD)
 * Optional (sensible defaults if unset — see getProp_ calls below):
 *   CLICKUP_SCHEMA_TAB, JIRA_SCHEMA_TAB, FIELD_MAPPING_TAB,
 *   STATUS_MAPPING_TAB, SPACE_MAPPING_TAB, EXPORT_TAB, LOG_TAB
 *
 * KNOWN PLATFORM LIMIT
 * Jira subtasks cannot have their own subtasks. ClickUp nesting beyond
 * 2 levels (Task → Subtask → Sub-subtask) cannot be represented — those
 * get written to the export sheet as SKIP – MANUAL with their full
 * ancestor chain, per your instruction to recreate those by hand.
 *
 * ASSUMPTIONS TO VERIFY ON FIRST RUN (flagged because they're the most
 * likely things to need adjustment against your actual data):
 *   - ClickUp's Get Tasks endpoint includes subtasks alongside their
 *     in-range parent even if the subtask's own date is outside range
 *   - Jira issue type names are exactly "Task" and "Sub-task" (hyphenated)
 *     — verify against your actual project if creation fails
 */

// ==================== CONFIG HELPERS ====================

function getProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

function ss_() {
  return SpreadsheetApp.openById(getProp_('SPREADSHEET_ID'));
}

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

// ==================== LOGGING (Function 4) ====================

function logRun_(functionName, status, details) {
  const sheet = getOrCreateTab_(getProp_('LOG_TAB', 'Migration Log'));
  if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp', 'Function', 'Status', 'Details']);
  sheet.appendRow([new Date(), functionName, status, details || '']);
  Logger.log('[' + status + '] ' + functionName + ': ' + (details || ''));
}

// ==================== CHECK SETUP ====================

function checkSetup() {
  const required = ['JIRA_SITE_URL', 'JIRA_EMAIL', 'JIRA_API_TOKEN', 'CLICKUP_API_TOKEN', 'CLICKUP_TEAM_ID', 'SPREADSHEET_ID'];
  let allSet = true;
  required.forEach(function (key) {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    Logger.log(key + ': ' + (v ? '✅ set' : '❌ NOT SET'));
    if (!v) allSet = false;
  });
  Logger.log(allSet ? '✅ All required properties set.' : '⚠️ Missing required properties — set them before running.');
}

// ==================== JIRA HELPERS ====================

function jiraAuthHeader_() {
  return 'Basic ' + Utilities.base64Encode(getProp_('JIRA_EMAIL') + ':' + getProp_('JIRA_API_TOKEN'));
}

function jiraSiteUrl_() {
  return getProp_('JIRA_SITE_URL').replace(/\/$/, '');
}

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
  const options = {
    method: 'get',
    headers: { 'Authorization': getProp_('CLICKUP_API_TOKEN') },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  Logger.log('❌ ClickUp GET ' + path + ' → HTTP ' + code + ': ' + body);
  return null;
}

// ==================== FUNCTION 1: DISCOVER CLICKUP SCHEMA ====================

function discoverClickUpSchema() {
  logRun_('discoverClickUpSchema', 'Info', 'Starting ClickUp discovery');
  const teamId = getProp_('CLICKUP_TEAM_ID');
  const sheet = getOrCreateTab_(getProp_('CLICKUP_SCHEMA_TAB', 'ClickUp Schema'));
  clearAndHeader_(sheet, ['Space', 'Folder', 'List', 'List ID', 'Field Name', 'Field Type', 'Field ID', 'Options / Config']);

  // Standard (built-in) fields every ClickUp task has, regardless of List —
  // written once at the top since these don't vary by Space/Folder/List.
  const STANDARD_FIELDS = [
    ['Task Name', 'text', 'name'],
    ['Description', 'text (rich)', 'description'],
    ['Status', 'status', 'status'],
    ['Assignee(s)', 'user(s)', 'assignees'],
    ['Priority', 'priority (urgent/high/normal/low)', 'priority'],
    ['Due Date', 'date', 'due_date'],
    ['Start Date', 'date', 'start_date'],
    ['Time Estimate', 'duration (ms)', 'time_estimate'],
    ['Time Tracked', 'duration (ms)', 'time_spent'],
    ['Tags', 'label(s)', 'tags'],
    ['Parent Task', 'relationship', 'parent'],
    ['Task ID', 'text', 'id'],
    ['Task URL', 'url', 'url'],
  ];
  STANDARD_FIELDS.forEach(function (f) {
    sheet.appendRow(['(all)', '(all)', '(all)', '(standard)', f[0], f[1], f[2], '']);
  });

  const spacesRes = clickUpRequest_('/team/' + teamId + '/space?archived=false');
  if (!spacesRes || !spacesRes.spaces) {
    logRun_('discoverClickUpSchema', 'Failed', 'Could not fetch spaces — check CLICKUP_TEAM_ID / token');
    return;
  }

  let rowCount = STANDARD_FIELDS.length;
  spacesRes.spaces.forEach(function (space) {
    // Folderless lists directly in the space
    const folderlessLists = clickUpRequest_('/space/' + space.id + '/list?archived=false');
    (folderlessLists && folderlessLists.lists ? folderlessLists.lists : []).forEach(function (list) {
      rowCount += writeListFields_(sheet, space.name, '(none)', list);
    });

    // Folders and their lists
    const foldersRes = clickUpRequest_('/space/' + space.id + '/folder?archived=false');
    (foldersRes && foldersRes.folders ? foldersRes.folders : []).forEach(function (folder) {
      (folder.lists || []).forEach(function (list) {
        rowCount += writeListFields_(sheet, space.name, folder.name, list);
      });
    });
  });

  logRun_('discoverClickUpSchema', 'Success', rowCount + ' field rows written to ' + sheet.getName());
}

function writeListFields_(sheet, spaceName, folderName, list) {
  const fieldsRes = clickUpRequest_('/list/' + list.id + '/field');
  const fields = fieldsRes && fieldsRes.fields ? fieldsRes.fields : [];
  let count = 0;
  if (fields.length === 0) {
    sheet.appendRow([spaceName, folderName, list.name, list.id, '(no custom fields)', '', '', '']);
    return 1;
  }
  fields.forEach(function (f) {
    const options = f.type_config && f.type_config.options
      ? f.type_config.options.map(function (o) { return o.name; }).join(', ')
      : '';
    sheet.appendRow([spaceName, folderName, list.name, list.id, f.name, f.type, f.id, options]);
    count++;
  });
  return count;
}

// ==================== FUNCTION 2: DISCOVER JIRA SCHEMA ====================

function discoverJiraSchema() {
  logRun_('discoverJiraSchema', 'Info', 'Starting Jira discovery');
  const sheet = getOrCreateTab_(getProp_('JIRA_SCHEMA_TAB', 'Jira Schema'));
  clearAndHeader_(sheet, ['Project', 'Issue Type', 'Field ID', 'Field Name', 'Field Type', 'Custom?', 'Required']);

  const projectsRes = jiraRequest_('GET', '/rest/api/3/project/search?maxResults=100');
  if (!projectsRes || !projectsRes.values) {
    logRun_('discoverJiraSchema', 'Failed', 'Could not fetch projects');
    return;
  }

  let rowCount = 0;
  projectsRes.values.forEach(function (project) {
    const statusesRes = jiraRequest_('GET', '/rest/api/3/project/' + project.key + '/statuses');
    const issueTypes = statusesRes || [];
    issueTypes.forEach(function (issueType) {
      // createmeta returns BOTH standard fields (summary, description, labels,
      // components, priority, timetracking, parent, etc.) AND custom fields
      // that are on this issue type's create screen — nothing extra needed
      // to surface standard fields, they're already in this response.
      const meta = jiraRequest_('GET', '/rest/api/3/issue/createmeta/' + project.key + '/issuetypes/' + issueType.id);
      const fields = meta && meta.fields ? meta.fields : [];
      fields.forEach(function (f) {
        const isCustom = f.schema && f.schema.custom ? true : false;
        sheet.appendRow([project.key, issueType.name, f.fieldId, f.name, (f.schema ? f.schema.type : ''), isCustom, f.required]);
        rowCount++;
      });
    });
  });

  logRun_('discoverJiraSchema', 'Success', rowCount + ' field rows written to ' + sheet.getName());
}

// ==================== MAPPING READERS ====================

function readMappingTab_() {
  const sheet = ss_().getSheetByName(getProp_('MAPPING_TAB', 'Mapping'));
  const result = { fieldMap: {}, statusMap: {}, spaceMap: {} };
  if (!sheet) { Logger.log('❌ "Mapping" tab not found — check MAPPING_TAB property / tab name.'); return result; }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // skip header row
    const row = data[i];
    const type = String(row[0] || '').trim();
    if (!type) continue;

    if (type === 'Field') {
      result.fieldMap[row[1]] = row[3];
    } else if (type === 'Status') {
      result.statusMap[row[1]] = row[3];
    } else if (type === 'Space') {
      const key = row[1] + '::' + row[2];
      result.spaceMap[key] = { projectKey: row[3], component: row[4], componentCategory: row[5] || '' };
    }
  }
  return result;
}

// ==================== FUNCTION 3a: DRY RUN EXPORT ====================

function dryRunExport() {
  logRun_('dryRunExport', 'Info', 'Starting dry run export');

  const mapping = readMappingTab_();
  const fieldMap = mapping.fieldMap;
  const statusMap = mapping.statusMap;
  const spaceMap = mapping.spaceMap;

  const startMs = new Date(getProp_('MIGRATION_START_DATE')).getTime();
  const endMs = new Date(getProp_('MIGRATION_END_DATE')).getTime();

  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));

  // Preserve already-created Jira Issue Key/Result before clearing — without
  // this, re-running dryRunExport() after liveCreateJiraTasks() has already
  // created real issues would wipe that linkage, orphaning those issues
  // (resetJiraTasksInRange() would no longer be able to find them).
  const preserved = {}; // ClickUp Task ID -> {jiraIssueKey, result, failureReason}
  const existingData = sheet.getDataRange().getValues();
  if (existingData.length > 1) {
    const existingHeaders = existingData[0];
    const idCol = existingHeaders.indexOf('ClickUp Task ID');
    const keyCol = existingHeaders.indexOf('Jira Issue Key');
    const resultCol = existingHeaders.indexOf('Result');
    const reasonCol = existingHeaders.indexOf('Failure Reason');
    if (idCol !== -1 && keyCol !== -1) {
      for (let i = 1; i < existingData.length; i++) {
        const cid = existingData[i][idCol];
        const key = existingData[i][keyCol];
        if (cid && key) {
          preserved[cid] = { jiraIssueKey: key, result: existingData[i][resultCol], failureReason: existingData[i][reasonCol] };
        }
      }
    }
  }
  if (Object.keys(preserved).length > 0) {
    logRun_('dryRunExport', 'Info', 'Preserving ' + Object.keys(preserved).length + ' already-created Jira issue link(s) from the previous export.');
  }

  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'ClickUp Task URL', 'Task Name', 'Description', 'Depth', 'Parent ClickUp ID',
    'Jira Issue Type', 'Project Key', 'Component', 'Component Category', 'ClickUp Status (raw)', 'Status (Jira)', 'Priority',
    'Start Date', 'Due Date', 'Time Estimate (hrs)', 'Time Tracked (hrs)', 'Tags', 'Rize Task ID', 'Rize Time Entry ID', 'ClickUp Attachments',
    'Migration Action', 'Migration Notes', 'Jira Issue Key', 'Result', 'Failure Reason',
  ]);

  const teamId = getProp_('CLICKUP_TEAM_ID');
  let unmappedStatusWarnings = 0;

  Object.keys(spaceMap).forEach(function (key) {
    const parts = key.split('::');
    const spaceName = parts[0], listName = parts[1];
    const spaceMapping = spaceMap[key];

    const listId = findListId_(teamId, spaceName, listName);
    if (!listId) {
      logRun_('dryRunExport', 'Warning', 'Could not find ClickUp list "' + listName + '" in space "' + spaceName + '" — skipping');
      return;
    }

    const tasks = fetchTasksInRange_(listId, startMs, endMs);
    const byId = {};
    tasks.forEach(function (t) { byId[t.id] = t; });

    const hasChildren = {};
    tasks.forEach(function (t) { if (t.parent) hasChildren[t.parent] = true; });

    tasks.forEach(function (task) {
      const depth = computeDepth_(task, byId);
      // The list-tasks endpoint doesn't reliably include attachments (confirmed
      // via ClickUp's own community feedback) — fetch the single task to get
      // an accurate count.
      const fullTask = clickUpRequest_('/task/' + task.id);
      const attachmentCount = fullTask && fullTask.attachments ? fullTask.attachments.length : 0;
      const row = buildExportRow_(task, depth, hasChildren, spaceMapping, fieldMap, statusMap, attachmentCount);
      if (row.statusUnmapped) unmappedStatusWarnings++;

      // Dynamic header lookup, not hardcoded indices — a hardcoded position
      // here would silently break every time a column gets added/reordered
      // (exactly what just happened when ClickUp Attachments was added).
      const prior = preserved[task.id];
      const headerRow = ['ClickUp Task ID', 'ClickUp Task URL', 'Task Name', 'Description', 'Depth', 'Parent ClickUp ID',
        'Jira Issue Type', 'Project Key', 'Component', 'Component Category', 'ClickUp Status (raw)', 'Status (Jira)', 'Priority',
        'Start Date', 'Due Date', 'Time Estimate (hrs)', 'Time Tracked (hrs)', 'Tags', 'Rize Task ID', 'Rize Time Entry ID', 'ClickUp Attachments',
        'Migration Action', 'Migration Notes', 'Jira Issue Key', 'Result', 'Failure Reason'];
      if (prior) {
        row.values[headerRow.indexOf('Jira Issue Key')] = prior.jiraIssueKey;
        row.values[headerRow.indexOf('Result')] = prior.result;
        row.values[headerRow.indexOf('Failure Reason')] = prior.failureReason;
      }

      sheet.appendRow(row.values);
    });
  });

  if (unmappedStatusWarnings > 0) {
    logRun_('dryRunExport', 'Warning', unmappedStatusWarnings + ' task(s) had a ClickUp status with no match in your Status mapping — defaulted to "New". Check the "ClickUp Status (raw)" column in the export to see the actual text and fix your Mapping tab.');
  }
  logRun_('dryRunExport', 'Success', 'Export sheet built at ' + sheet.getName());
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

function fetchTasksInRange_(listId, startMs, endMs) {
  let allTasks = [];
  let page = 0;
  while (true) {
    // No server-side start_date filter exists in ClickUp's API (confirmed —
    // only due_date/date_created/date_updated are supported) so we fetch
    // everything in the list and filter by Start Date client-side below.
    const res = clickUpRequest_('/list/' + listId + '/task?archived=false&include_closed=true&subtasks=true&page=' + page);
    if (!res || !res.tasks || res.tasks.length === 0) break;
    allTasks = allTasks.concat(res.tasks);
    if (res.last_page) break;
    page++;
  }

  return allTasks.filter(function (task) {
    if (!task.start_date) return false; // no Start Date at all — excluded, can't evaluate against range
    const start = Number(task.start_date);
    return start >= startMs && start <= endMs;
  });
}

// ==================== DIAGNOSTIC: COMPONENT CATEGORY OPTIONS ====================
// Run this once before wiring Component Category into issue creation — it
// shows the ACTUAL parent/child values configured in that field, since it
// was built outside these scripts and there's no record of its structure.

function discoverComponentCategoryOptions() {
  const COMPONENT_CATEGORY_FIELD_ID = 'customfield_10130';
  const contexts = jiraRequest_('GET', '/rest/api/3/field/' + COMPONENT_CATEGORY_FIELD_ID + '/context');
  const contextId = contexts && contexts.values && contexts.values[0] ? contexts.values[0].id : null;
  if (!contextId) { Logger.log('❌ Could not find a context for Component Category.'); return; }

  const options = jiraRequest_('GET', '/rest/api/3/field/' + COMPONENT_CATEGORY_FIELD_ID + '/context/' + contextId + '/option?maxResults=200');
  if (!options || !options.values) { Logger.log('❌ Could not fetch options.'); return; }

  Logger.log('--- Component Category configured options ---');
  const parents = options.values.filter(function (o) { return !o.optionId; });
  const children = options.values.filter(function (o) { return o.optionId; });

  parents.forEach(function (p) {
    Logger.log('PARENT: "' + p.value + '" (id: ' + p.id + ')');
    children.filter(function (c) { return c.optionId === p.id; }).forEach(function (c) {
      Logger.log('   └─ CHILD: "' + c.value + '" (id: ' + c.id + ')');
    });
  });
}

function getCustomFieldValue_(task, fieldName) {
  const cf = (task.custom_fields || []).find(function (f) { return f.name === fieldName; });
  return cf && cf.value !== undefined && cf.value !== null ? cf.value : '';
}

function computeDepth_(task, byId) {
  let depth = 0;
  let current = task;
  while (current.parent) {
    let parentTask = byId[current.parent];
    if (!parentTask) {
      // Parent wasn't in the date-filtered result set — fetch it directly
      // rather than silently treating this task as top-level (the bug:
      // a real subtask whose parent falls outside the migration date
      // range was previously miscategorized as a Task instead of a Sub-task).
      parentTask = clickUpRequest_('/task/' + current.parent);
      if (!parentTask || !parentTask.id) break; // parent genuinely doesn't exist / API error — stop climbing
      byId[current.parent] = parentTask; // cache so repeated lookups don't re-fetch
    }
    depth++;
    current = parentTask;
  }
  return depth;
}

function buildExportRow_(task, depth, hasChildren, mapping, fieldMap, statusMap, attachmentCount) {
  const clickUpUrl = 'https://app.clickup.com/t/' + getProp_('CLICKUP_TEAM_ID') + '/' + task.id;
  // Capped at Jira's actual documented limit (32,767 chars), not an
  // arbitrary smaller number — this is the same value stored in the sheet
  // and sent to Jira, so there's no mismatch between preview and reality.
  const description = cleanDescription_(task.description || '').substring(0, 32767);
  const rawStatus = task.status && task.status.status ? task.status.status : '';
  const statusUnmapped = rawStatus && !statusMap[rawStatus];
  const jiraStatus = statusMap[rawStatus] || 'New';
  const priority = task.priority ? task.priority.priority : '';
  const startDate = task.start_date ? Utilities.formatDate(new Date(Number(task.start_date)), 'UTC', 'yyyy-MM-dd') : '';
  const dueDate = task.due_date ? Utilities.formatDate(new Date(Number(task.due_date)), 'UTC', 'yyyy-MM-dd') : '';
  // ClickUp stores these in milliseconds; convert to hours for the Jira number fields
  const estimateHrs = task.time_estimate ? (Number(task.time_estimate) / 3600000).toFixed(2) : '';
  const trackedHrs = task.time_spent ? (Number(task.time_spent) / 3600000).toFixed(2) : '';
  const tags = (task.tags || []).map(function (t) { return t.name; }).join(', ');
  const rizeTaskId = getCustomFieldValue_(task, 'Rize Task ID');
  const rizeTimeEntryId = getCustomFieldValue_(task, 'Rize Time Entry ID');

  let issueType, action, notes;
  if (depth === 0) {
    issueType = 'Task'; action = 'CREATE'; notes = '';
  } else if (depth === 1) {
    issueType = 'Sub-task'; action = 'CREATE';
    notes = hasChildren[task.id] ? 'Has deeper children not representable in Jira — those flagged separately below' : '';
  } else {
    issueType = 'Sub-task (nested)'; action = 'SKIP - MANUAL';
    notes = 'Depth ' + depth + ' — Jira cannot nest subtasks. Recreate manually under parent ' + task.parent;
  }

  if (statusUnmapped) {
    notes = (notes ? notes + ' | ' : '') + 'STATUS NOT MAPPED — "' + rawStatus + '" has no entry in your Mapping tab, defaulted to New.';
  }

  return {
    statusUnmapped: statusUnmapped,
    values: [
      task.id, clickUpUrl, task.name, description, depth, task.parent || '',
      issueType, mapping.projectKey, mapping.component, mapping.componentCategory || '', rawStatus, jiraStatus, priority,
      startDate, dueDate, estimateHrs, trackedHrs, tags, rizeTaskId, rizeTimeEntryId, attachmentCount || 0,
      action, notes, '', '', '',
    ],
  };
}

// ==================== FUNCTION 3b: LIVE CREATE ====================

// Google Sheets sometimes auto-converts a YYYY-MM-DD-looking string into an
// actual Date object when read back via getValue() — handle both cases so
// the format sent to Jira is reliably 'yyyy-MM-dd', not a full timestamp.
function toJiraDateString_(cellValue) {
  if (!cellValue) return null;
  if (Object.prototype.toString.call(cellValue) === '[object Date]') {
    return Utilities.formatDate(cellValue, 'UTC', 'yyyy-MM-dd');
  }
  return String(cellValue).substring(0, 10); // already a string — take just the date portion defensively
}

function liveCreateJiraTasks() {
  logRun_('liveCreateJiraTasks', 'Info', 'Starting live create');
  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  const fieldMap = readMappingTab_().fieldMap; // read once, not once per task

  if (!validateFieldMapping_(fieldMap)) {
    logRun_('liveCreateJiraTasks', 'Failed', 'Aborting — fix the Mapping tab issues logged above before running live create.');
    return;
  }

  const startRow = Number(getProp_('LAST_PROCESSED_ROW', '1')); // 1 = first data row (0 is header)
  const clickUpIdToJiraKey = {}; // built across this run for parent resolution

  // Pass 1: Tasks
  for (let i = startRow; i < data.length; i++) {
    if (data[i][col['Migration Action']] !== 'CREATE') continue;
    if (data[i][col['Jira Issue Type']] !== 'Task') continue;
    if (data[i][col['Result']] === 'Success') continue;
    createIssueRow_(sheet, i, col, clickUpIdToJiraKey, null, fieldMap);
    if (isNearTimeLimit_()) {
      PropertiesService.getScriptProperties().setProperty('LAST_PROCESSED_ROW', String(i));
      logRun_('liveCreateJiraTasks', 'Warning', 'Paused at row ' + i + ' (Task pass) — re-run to resume');
      return;
    }
  }

  // Pass 2: Subtasks (parents now resolved from pass 1, or from prior runs already in sheet)
  for (let i = 1; i < data.length; i++) {
    if (data[i][col['Migration Action']] !== 'CREATE') continue;
    if (data[i][col['Jira Issue Type']] !== 'Sub-task') continue;
    if (data[i][col['Result']] === 'Success') continue;
    const parentClickUpId = data[i][col['Parent ClickUp ID']];
    const parentKey = clickUpIdToJiraKey[parentClickUpId] || findJiraKeyForClickUpId_(data, col, parentClickUpId);
    if (!parentKey) {
      sheet.getRange(i + 1, col['Result'] + 1).setValue('Failed');
      sheet.getRange(i + 1, col['Failure Reason'] + 1).setValue('Parent not yet created in Jira');
      logRun_('liveCreateJiraTasks', 'Failed', 'Row ' + (i + 1) + ': parent not resolved');
      continue;
    }
    createIssueRow_(sheet, i, col, clickUpIdToJiraKey, parentKey, fieldMap);
    if (isNearTimeLimit_()) {
      logRun_('liveCreateJiraTasks', 'Warning', 'Paused at row ' + i + ' (Subtask pass) — re-run to resume');
      return;
    }
  }

  PropertiesService.getScriptProperties().deleteProperty('LAST_PROCESSED_ROW');
  logRun_('liveCreateJiraTasks', 'Success', 'Live create run complete');
}

function findJiraKeyForClickUpId_(data, col, clickUpId) {
  for (let i = 1; i < data.length; i++) {
    if (data[i][col['ClickUp Task ID']] == clickUpId) return data[i][col['Jira Issue Key']];
  }
  return null;
}

function isNearTimeLimit_() {
  // Apps Script hard limit is 6 min; stop at 5 to leave margin (mirrors your Gmail classifier pattern)
  return (new Date().getTime() - scriptStartTime_) > 5 * 60 * 1000;
}
const scriptStartTime_ = new Date().getTime();

// Confirmed from your actual Jira Schema export — hardcoded rather than
// left to the Mapping tab, since a mix-up there (pointing at the native
// "timetracking" field instead of these) caused a real failure: Jira's
// native Time Tracking field needs a structured value like
// {originalEstimate: "2h 30m"}, not a plain decimal, and rejects it.
// Field IDs for these come from the Mapping tab (Field rows), not hardcoded —
// customfield_XXXXX numbers change on every factory reset + rebuild, so
// hardcoding them here would silently break next time that happens.
// What broke before wasn't the sheet-driven design — it was a bad value
// (the native 'timetracking' field ID) sitting in the Mapping tab for these
// rows. validateFieldMapping_() below catches that specific mistake loudly
// instead of either hardcoding around it or silently sending a bad payload.

function validateFieldMapping_(fieldMap) {
  const problems = [];
  // "Time Tracked" is intentionally not checked here anymore — it's no
  // longer a field lookup at all. Jira's time-spent can only be set via a
  // worklog entry (confirmed via Atlassian docs), not any field, custom or
  // native, so createIssueRow_ handles it as a separate POST after creation.
  ['Time Estimate', 'Start Date'].forEach(function (key) {
    if (fieldMap[key] === 'timetracking' || fieldMap[key] === 'duedate') {
      problems.push('"' + key + '" in your Mapping tab points to "' + fieldMap[key] + '" — that\'s a native Jira field with a different format requirement, not your dedicated custom field. Fix this row in the Mapping tab before running live create.');
    }
  });
  if (problems.length > 0) {
    problems.forEach(function (p) { logRun_('validateFieldMapping_', 'Failed', p); });
    return false;
  }
  return true;
}

// Checks LIVE Jira (not just the local sheet) for an issue already linked
// to this ClickUp task, before creating a new one. This is the actual
// fix for the duplicate problem — duplicates in the real data came from
// entirely separate historical runs across different projects (WORK,
// LERN, CARE, SOS), which local sheet bookkeeping alone can't catch,
// since each run only knows about its own sheet's history, not what
// other runs created. Uses /rest/api/3/search/jql (POST) — the same
// endpoint proven to work in the nested-cleanup and audit scripts.
function findExistingJiraIssueForClickUpTask_(clickUpTaskId) {
  const jql = 'cf[10080] ~ "' + clickUpTaskId + '"';
  const res = jiraRequest_('POST', '/rest/api/3/search/jql', { jql: jql, maxResults: 1, fields: ['summary', 'issuetype'] });
  if (res && res.issues && res.issues.length > 0) return res.issues[0].key;
  return null;
}

function createIssueRow_(sheet, i, col, clickUpIdToJiraKey, parentKey, fieldMap) {
  const rowNum = i + 1;
  const clickUpTaskId = sheet.getRange(rowNum, col['ClickUp Task ID'] + 1).getValue();

  // Live duplicate check — before creating anything, ask Jira directly
  // whether this ClickUp task already has an issue, regardless of which
  // script or run created it.
  const existingKey = findExistingJiraIssueForClickUpTask_(clickUpTaskId);
  if (existingKey) {
    Logger.log('⏭️  ' + clickUpTaskId + ' already exists in Jira as ' + existingKey + ' — skipping creation, not creating a duplicate.');
    sheet.getRange(rowNum, col['Jira Issue Key'] + 1).setValue(existingKey);
    sheet.getRange(rowNum, col['Result'] + 1).setValue('Already existed');
    clickUpIdToJiraKey[clickUpTaskId] = existingKey;
    return;
  }

  const fields = {
    project: { key: sheet.getRange(rowNum, col['Project Key'] + 1).getValue() },
    issuetype: { name: parentKey ? 'Sub-Task' : 'Task' },
    summary: sheet.getRange(rowNum, col['Task Name'] + 1).getValue(),
    description: toADF_(sheet.getRange(rowNum, col['Description'] + 1).getValue()),
  };
  if (parentKey) fields.parent = { key: parentKey };

  const component = sheet.getRange(rowNum, col['Component'] + 1).getValue();
  if (component) fields.components = [{ name: component }];

  // Component Category (customfield_10130) — confirmed via discovery that
  // this field currently has NO parent/child nesting despite being a
  // cascading-select type; every configured option is a flat top-level
  // value, so this is a plain {value: ...} write, no "child" key needed.
  const componentCategory = sheet.getRange(rowNum, col['Component Category'] + 1).getValue();
  if (componentCategory) fields['customfield_10130'] = { value: componentCategory };

  const clickUpUrl = sheet.getRange(rowNum, col['ClickUp Task URL'] + 1).getValue();
  const clickUpFieldId = fieldMap['ClickUp Task URL'];
  if (clickUpFieldId && clickUpUrl) fields[clickUpFieldId] = toADF_(clickUpUrl);

  const estimateFieldId = fieldMap['Time Estimate'];
  const estimateVal = sheet.getRange(rowNum, col['Time Estimate (hrs)'] + 1).getValue();
  if (estimateFieldId && estimateVal !== '') fields[estimateFieldId] = Number(estimateVal);

  const trackedVal = sheet.getRange(rowNum, col['Time Tracked (hrs)'] + 1).getValue();

  // Previously missing entirely — Rize Task ID/Time Entry ID were mapped in
  // the Mapping tab (to Rize ID / Rize Time Entry ID custom fields) but the
  // code never actually read ClickUp's custom_fields to populate them.
  const rizeTaskIdVal = sheet.getRange(rowNum, col['Rize Task ID'] + 1).getValue();
  const rizeTaskIdFieldId = fieldMap['Rize Task ID'];
  if (rizeTaskIdFieldId && rizeTaskIdVal !== '') fields[rizeTaskIdFieldId] = String(rizeTaskIdVal);

  const rizeTimeEntryIdVal = sheet.getRange(rowNum, col['Rize Time Entry ID'] + 1).getValue();
  const rizeTimeEntryIdFieldId = fieldMap['Rize Time Entry ID'];
  if (rizeTimeEntryIdFieldId && rizeTimeEntryIdVal !== '') fields[rizeTimeEntryIdFieldId] = String(rizeTimeEntryIdVal);

  // Start Date field ID comes from the Mapping tab; Due Date uses the
  // native 'duedate' key directly since it's a true system field, not
  // something that varies or needs mapping.
  const startDateFieldId = fieldMap['Start Date'];
  const startDateStr = toJiraDateString_(sheet.getRange(rowNum, col['Start Date'] + 1).getValue());
  if (startDateFieldId && startDateStr) fields[startDateFieldId] = startDateStr;

  const dueDateStr = toJiraDateString_(sheet.getRange(rowNum, col['Due Date'] + 1).getValue());
  if (dueDateStr) fields.duedate = dueDateStr;

  const result = jiraRequest_('POST', '/rest/api/3/issue', { fields: fields });

  if (result && result.key && trackedVal !== '') {
    // Time Spent cannot be set as a field on create/edit — Jira only exposes
    // it via worklog entries. Confirmed via Atlassian docs: the timetracking
    // field only accepts originalEstimate/remainingEstimate, never timeSpent.
    const seconds = Math.round(Number(trackedVal) * 3600);
    const startedRaw = sheet.getRange(rowNum, col['Start Date'] + 1).getValue();
    const startedDate = startedRaw ? (Object.prototype.toString.call(startedRaw) === '[object Date]' ? startedRaw : new Date(startedRaw)) : new Date();
    const startedStr = Utilities.formatDate(startedDate, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
    const worklogResult = jiraRequest_('POST', '/rest/api/3/issue/' + result.key + '/worklog', {
      timeSpentSeconds: seconds,
      started: startedStr,
    });
    if (worklogResult) Logger.log('✅ Logged ' + trackedVal + 'h worklog on ' + result.key);
    else Logger.log('❌ Failed to log worklog on ' + result.key + ' — issue itself was created fine, just the time entry failed');
  }

  if (result && result.key) {
    const attachmentCount = Number(sheet.getRange(rowNum, col['ClickUp Attachments'] + 1).getValue()) || 0;
    if (attachmentCount > 0) {
      const clickUpTaskId = sheet.getRange(rowNum, col['ClickUp Task ID'] + 1).getValue();
      migrateAttachments_(clickUpTaskId, result.key);
    }
  }

  if (result && result.key) {
    // Validate by fetching it back
    const verify = jiraRequest_('GET', '/rest/api/3/issue/' + result.key);
    const ok = verify && verify.fields && verify.fields.summary === fields.summary;
    sheet.getRange(rowNum, col['Jira Issue Key'] + 1).setValue(result.key);
    sheet.getRange(rowNum, col['Result'] + 1).setValue(ok ? 'Success' : 'Warning');
    sheet.getRange(rowNum, col['Failure Reason'] + 1).setValue(ok ? '' : 'Created but validation mismatch — check manually');
    clickUpIdToJiraKey[sheet.getRange(rowNum, col['ClickUp Task ID'] + 1).getValue()] = result.key;
    logRun_('liveCreateJiraTasks', ok ? 'Success' : 'Warning', 'Row ' + rowNum + ' → ' + result.key);
  } else {
    sheet.getRange(rowNum, col['Result'] + 1).setValue('Failed');
    sheet.getRange(rowNum, col['Failure Reason'] + 1).setValue('Jira API rejected creation — see Apps Script log for HTTP error');
    logRun_('liveCreateJiraTasks', 'Failed', 'Row ' + rowNum + ' creation failed');
  }
}

function cleanDescription_(text) {
  if (!text) return '';
  // ClickUp embeds raw markup like [table-embed:1:1 ... |] in descriptions —
  // strip it to a plain note rather than dumping the syntax verbatim into Jira.
  let cleaned = text.replace(/\[table-embed:[^\]]*\]/g, '[table content omitted — view original in ClickUp]');
  cleaned = cleaned.trim();
  // Prevent Sheets from misinterpreting a leading "=" as a formula trigger
  // (this is what caused the === → #ERROR! problem) — a single leading
  // space is enough to stop formula parsing without meaningfully altering
  // the visible text.
  if (cleaned.indexOf('=') === 0) cleaned = ' ' + cleaned;
  return cleaned;
}

// ==================== REPAIR: BROKEN DESCRIPTIONS (=== / #ERROR!) ====================
// Some ClickUp descriptions start with "===" — Sheets interprets a cell
// starting with "=" as a formula, "==text" isn't valid syntax, so it shows
// #ERROR!. The original text isn't actually lost: getFormula() returns the
// literal typed content (including the leading "="s), separately from
// getValue() which only shows the broken computed result. This recovers
// the real text from there and pushes the corrected description to the
// already-created Jira issue.
//
// LIMIT: this only works if the cell is CURRENTLY showing the formula
// error — if the original text was already overwritten some other way,
// there's nothing left to recover it from.

function fixBrokenDescriptions() {
  Logger.log('=== Fixing broken (=== / #ERROR!) descriptions ===');
  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  let fixed = 0, skippedNotBroken = 0, skippedNoKey = 0, failed = 0;

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const descCell = sheet.getRange(rowNum, col['Description'] + 1);
    const displayVal = descCell.getValue();
    const formula = descCell.getFormula();

    const looksBroken = formula !== '' ||
      (typeof displayVal === 'string' && (displayVal.indexOf('#ERROR') !== -1 || displayVal.indexOf('#REF') !== -1));
    if (!looksBroken) { skippedNotBroken++; continue; }

    const rawText = formula !== '' ? formula.replace(/^=+/, '') : String(displayVal);
    const jiraKey = data[i][col['Jira Issue Key']];
    if (!jiraKey) {
      Logger.log('⏭️  Row ' + rowNum + ' has a broken description but no Jira Issue Key — not created yet, nothing to fix on the Jira side.');
      skippedNoKey++;
      continue;
    }

    const result = jiraRequest_('PUT', '/rest/api/3/issue/' + jiraKey, { fields: { description: toADF_(rawText) } });
    if (result !== null) {
      Logger.log('✅ Fixed description on ' + jiraKey);
      fixed++;
      // Force the cell to Plain Text format before writing back, otherwise
      // setValue() would just trigger the same formula misinterpretation again.
      descCell.setNumberFormat('@');
      descCell.setValue(rawText);
    } else {
      Logger.log('❌ Failed to update ' + jiraKey + ' — see error above.');
      failed++;
    }
  }

  Logger.log('=== Done. Fixed: ' + fixed + ', not broken: ' + skippedNotBroken + ', no Jira key yet: ' + skippedNoKey + ', failed: ' + failed + ' ===');
}

// ==================== ATTACHMENT MIGRATION ====================
// NEW / UNTESTED — confirmed against ClickUp and Jira docs but not yet run
// end-to-end. Test on one or two tasks with attachments before trusting on
// a full batch, per the plan going in.
//
// ClickUp: the list-tasks endpoint doesn't reliably return attachments
// (confirmed via ClickUp's own community feedback) — only the single-task
// GET /task/{id} does, which is why dryRunExport() fetches that separately
// just for the attachment count.
//
// Jira: POST /rest/api/3/issue/{key}/attachments requires multipart/form-data
// with the file under a parameter literally named 'file', and the header
// X-Atlassian-Token: no-check — Jira blocks the request without it.

function migrateAttachments_(clickUpTaskId, jiraIssueKey) {
  const fullTask = clickUpRequest_('/task/' + clickUpTaskId);
  const attachments = fullTask && fullTask.attachments ? fullTask.attachments : [];
  if (attachments.length === 0) return;

  attachments.forEach(function (att) {
    try {
      // ClickUp attachment URLs are secured by being an unguessable random
      // string, NOT by requiring auth (confirmed via ClickUp's own Help
      // Center docs on Private Attachment Links) — no Authorization header
      // here. If this still 401s, the Workspace likely has "Private
      // Attachment Links" turned on (Settings → Security & Permissions →
      // Advanced Permissions), which is a different, non-default auth model
      // this script doesn't currently handle.
      const fileResponse = UrlFetchApp.fetch(att.url, { muteHttpExceptions: true });
      if (fileResponse.getResponseCode() >= 300) {
        Logger.log('❌ Failed to download attachment "' + att.title + '" from ClickUp — HTTP ' + fileResponse.getResponseCode());
        return;
      }
      const blob = fileResponse.getBlob().setName(att.title || att.id);

      const uploadResponse = UrlFetchApp.fetch(jiraSiteUrl_() + '/rest/api/3/issue/' + jiraIssueKey + '/attachments', {
        method: 'post',
        headers: { 'Authorization': jiraAuthHeader_(), 'X-Atlassian-Token': 'no-check' },
        payload: { file: blob },
        muteHttpExceptions: true,
      });
      if (uploadResponse.getResponseCode() < 300) {
        Logger.log('✅ Migrated attachment "' + (att.title || att.id) + '" to ' + jiraIssueKey);
      } else {
        Logger.log('❌ Failed to upload attachment "' + (att.title || att.id) + '" to ' + jiraIssueKey + ' — HTTP ' + uploadResponse.getResponseCode() + ': ' + uploadResponse.getContentText());
      }
    } catch (err) {
      Logger.log('❌ Error migrating attachment "' + (att.title || att.id) + '": ' + err.message);
    }
  });
}

function toADF_(text) {
  const cleaned = cleanDescription_(text);
  if (!cleaned) return { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [] }] };

  // Split on line breaks so multi-line descriptions become separate ADF
  // paragraphs instead of one run-on blob with embedded \n characters.
  const paragraphs = cleaned.split(/\n+/).map(function (p) { return p.trim(); }).filter(function (p) { return p.length > 0; });
  if (paragraphs.length === 0) return { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [] }] };

  return {
    type: 'doc', version: 1,
    content: paragraphs.map(function (p) {
      return { type: 'paragraph', content: [{ type: 'text', text: p }] };
    }),
  };
}

// ==================== FUNCTION 5: RESET ====================

function resetJiraTasksInRange() {
  logRun_('resetJiraTasksInRange', 'Info', 'Starting reset');
  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  let deleted = 0;
  for (let i = 1; i < data.length; i++) {
    const issueKey = data[i][col['Jira Issue Key']];
    const issueType = data[i][col['Jira Issue Type']];
    if (!issueKey) continue;
    if (issueType === 'Sub-task') continue; // cascade-deleted with parent below

    const res = UrlFetchApp.fetch(jiraSiteUrl_() + '/rest/api/3/issue/' + issueKey + '?deleteSubtasks=true', {
      method: 'delete', headers: { 'Authorization': jiraAuthHeader_() }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() < 300) {
      sheet.getRange(i + 1, col['Jira Issue Key'] + 1).setValue('');
      sheet.getRange(i + 1, col['Result'] + 1).setValue('');
      deleted++;
      logRun_('resetJiraTasksInRange', 'Success', 'Deleted ' + issueKey + ' (and subtasks)');
    } else {
      logRun_('resetJiraTasksInRange', 'Failed', 'Could not delete ' + issueKey + ': HTTP ' + res.getResponseCode());
    }
  }

  logRun_('resetJiraTasksInRange', 'Success', 'Reset complete — ' + deleted + ' top-level issues deleted');
}