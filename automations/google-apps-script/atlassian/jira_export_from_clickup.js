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

function readKeyValueTab_(tabName) {
  const sheet = ss_().getSheetByName(tabName);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) { // skip header row
    if (data[i][0]) map[String(data[i][0]).trim()] = data[i][1];
  }
  return map;
}

function readSpaceMapping_() {
  const sheet = ss_().getSheetByName(getProp_('SPACE_MAPPING_TAB', 'Space Mapping'));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map = {}; // key: "Space::List" -> {projectKey, component}
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const key = row[0] + '::' + row[1];
    map[key] = { projectKey: row[2], component: row[3] };
  }
  return map;
}

// ==================== FUNCTION 3a: DRY RUN EXPORT ====================

function dryRunExport() {
  logRun_('dryRunExport', 'Info', 'Starting dry run export');

  const fieldMap = readKeyValueTab_(getProp_('FIELD_MAPPING_TAB', 'Field Mapping'));      // logical name -> Jira field ID
  const statusMap = readKeyValueTab_(getProp_('STATUS_MAPPING_TAB', 'Status Mapping'));   // ClickUp status -> Jira status
  const spaceMap = readSpaceMapping_();                                                   // "Space::List" -> {projectKey, component}

  const startMs = new Date(getProp_('MIGRATION_START_DATE')).getTime();
  const endMs = new Date(getProp_('MIGRATION_END_DATE')).getTime();

  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));
  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'ClickUp Task URL', 'Task Name', 'Description', 'Depth', 'Parent ClickUp ID',
    'Jira Issue Type', 'Project Key', 'Component', 'Status (Jira)', 'Priority', 'Start Date', 'Due Date',
    'Tags', 'Migration Action', 'Migration Notes', 'Jira Issue Key', 'Result', 'Failure Reason',
  ]);

  const teamId = getProp_('CLICKUP_TEAM_ID');

  Object.keys(spaceMap).forEach(function (key) {
    const parts = key.split('::');
    const spaceName = parts[0], listName = parts[1];
    const mapping = spaceMap[key];

    // Find the list ID by re-crawling (small — reuses discovery calls)
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
      const row = buildExportRow_(task, depth, hasChildren, mapping, fieldMap, statusMap);
      sheet.appendRow(row);
    });
  });

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
    const res = clickUpRequest_('/list/' + listId + '/task?archived=false&include_closed=true&subtasks=true&page=' + page +
      '&date_created_gt=' + startMs + '&date_created_lt=' + endMs);
    if (!res || !res.tasks || res.tasks.length === 0) break;
    allTasks = allTasks.concat(res.tasks);
    if (res.last_page) break;
    page++;
  }
  return allTasks;
}

function computeDepth_(task, byId) {
  let depth = 0;
  let current = task;
  while (current.parent && byId[current.parent]) {
    depth++;
    current = byId[current.parent];
  }
  return depth;
}

function buildExportRow_(task, depth, hasChildren, mapping, fieldMap, statusMap) {
  const clickUpUrl = 'https://app.clickup.com/t/' + getProp_('CLICKUP_TEAM_ID') + '/' + task.id;
  const jiraStatus = statusMap[task.status && task.status.status] || 'New';
  const priority = task.priority ? task.priority.priority : '';
  const startDate = task.start_date ? new Date(Number(task.start_date)).toISOString() : '';
  const dueDate = task.due_date ? new Date(Number(task.due_date)).toISOString() : '';
  const tags = (task.tags || []).map(function (t) { return t.name; }).join(', ');

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

  return [
    task.id, clickUpUrl, task.name, (task.description || '').substring(0, 500), depth, task.parent || '',
    issueType, mapping.projectKey, mapping.component, jiraStatus, priority, startDate, dueDate, tags,
    action, notes, '', '', '',
  ];
}

// ==================== FUNCTION 3b: LIVE CREATE ====================

function liveCreateJiraTasks() {
  logRun_('liveCreateJiraTasks', 'Info', 'Starting live create');
  const sheet = getOrCreateTab_(getProp_('EXPORT_TAB', 'Migration Export'));
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  const startRow = Number(getProp_('LAST_PROCESSED_ROW', '1')); // 1 = first data row (0 is header)
  const clickUpIdToJiraKey = {}; // built across this run for parent resolution

  // Pass 1: Tasks
  for (let i = startRow; i < data.length; i++) {
    if (data[i][col['Migration Action']] !== 'CREATE') continue;
    if (data[i][col['Jira Issue Type']] !== 'Task') continue;
    if (data[i][col['Result']] === 'Success') continue;
    createIssueRow_(sheet, i, col, clickUpIdToJiraKey, null);
    if (new Date().getTime() % 100 === 0) {} // no-op placeholder to keep loop simple
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
    createIssueRow_(sheet, i, col, clickUpIdToJiraKey, parentKey);
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

function createIssueRow_(sheet, i, col, clickUpIdToJiraKey, parentKey) {
  const rowNum = i + 1;
  const fields = {
    project: { key: sheet.getRange(rowNum, col['Project Key'] + 1).getValue() },
    issuetype: { name: parentKey ? 'Sub-task' : 'Task' },
    summary: sheet.getRange(rowNum, col['Task Name'] + 1).getValue(),
    description: toADF_(sheet.getRange(rowNum, col['Description'] + 1).getValue()),
  };
  if (parentKey) fields.parent = { key: parentKey };

  const component = sheet.getRange(rowNum, col['Component'] + 1).getValue();
  if (component) fields.components = [{ name: component }];

  const clickUpUrl = sheet.getRange(rowNum, col['ClickUp Task URL'] + 1).getValue();
  const clickUpFieldId = readKeyValueTab_(getProp_('FIELD_MAPPING_TAB', 'Field Mapping'))['ClickUp Task URL'];
  if (clickUpFieldId && clickUpUrl) fields[clickUpFieldId] = clickUpUrl;

  const result = jiraRequest_('POST', '/rest/api/3/issue', { fields: fields });

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

function toADF_(text) {
  return {
    type: 'doc', version: 1,
    content: [{ type: 'paragraph', content: [{ type: 'text', text: text || '' }] }],
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