/**
 * ClickUp Deeply-Nested Task Cleanup — Epic/Task/Sub-task
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * PURPOSE
 * Handles ClickUp task trees 3+ levels deep, which the main migration
 * pipeline correctly flagged as SKIP - MANUAL (Jira subtasks cannot have
 * their own subtasks — a genuine platform limit, not a bug).
 *
 * MAPS:
 *   ClickUp depth 0 (root)        → Jira Epic
 *   ClickUp depth 1 (root's kids) → Jira Task, linked to the Epic via
 *                                    Epic Link (customfield_10014)
 *   ClickUp depth 2 (their kids)  → Jira Sub-task, parented under the
 *                                    matching depth-1 Task
 *   ClickUp depth 3+ (if any)     → still cannot be represented — logged
 *                                    as still needing manual handling,
 *                                    same as before
 *
 * IDEMPOTENT: before creating anything, searches Jira by ClickUp Task ID
 * (embedded in the URL(s) field) to avoid duplicates on re-run.
 *
 * CREDENTIALS (Script Properties)
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN
 *   CLICKUP_API_TOKEN, CLICKUP_TEAM_ID
 *   SPREADSHEET_ID         — target Google Sheet for the plan tab
 *   ROOT_CLICKUP_TASK_ID   — the depth-0 task ID for the tree you're
 *                            cleaning up right now (e.g. "86bamp7eq")
 * Optional:
 *   TARGET_COMPONENT       — Component to set on created issues
 *   TARGET_CATEGORY        — Component Category (customfield_10130) value
 *
 * RUN ORDER
 *   checkSetup() → scanForNestedTrees() (finds every nested tree across
 *   all your Lists, writes "Nested Trees Found" tab) → for each root ID
 *   found: set ROOT_CLICKUP_TASK_ID to it → planNestedTreeCleanup()
 *   (dry run, writes "Nested Cleanup Plan" tab, zero Jira writes) →
 *   review the plan, especially "Delete Needed" → executeNestedTreeCleanup()
 */

// ==================== CONFIG ====================

function getProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

const PROJECT_KEY_LERN = 'LERN';
//const CLICKUP_URL_FIELD_ID = 'customfield_10080';  // URL(s)
const EPIC_LINK_FIELD_ID = 'customfield_10014';    // Epic Link
const EPIC_NAME_FIELD_ID = 'customfield_10011';    // Epic Name — required on Epic create

// ==================== HELPERS ====================

function checkSetup() {
  ['JIRA_SITE_URL', 'JIRA_EMAIL', 'JIRA_API_TOKEN', 'CLICKUP_API_TOKEN', 'CLICKUP_TEAM_ID', 'ROOT_CLICKUP_TASK_ID'].forEach(function (key) {
    Logger.log(key + ': ' + (getProp_(key) ? '✅ set' : '❌ NOT SET'));
  });
}

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

function clickUpRequest_(path) {
  const options = { method: 'get', headers: { 'Authorization': getProp_('CLICKUP_API_TOKEN') }, muteHttpExceptions: true };
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  Logger.log('❌ ClickUp GET ' + path + ' → HTTP ' + code + ': ' + body);
  return null;
}

function clickUpUrl_(taskId) {
  return 'https://app.clickup.com/t/' + getProp_('CLICKUP_TEAM_ID') + '/' + taskId;
}

// Idempotency check — searches Jira for an existing issue already linked
// to this ClickUp task, so re-running this script is safe. Returns both
// the key AND issue type, since a match of the WRONG type (e.g. a Task
// that was created before we knew this tree needed to be an Epic) needs
// to be deleted and replaced, not treated as "already done."
function findExistingJiraIssue_(clickUpTaskId) {
  const jql = 'project = ' + PROJECT_KEY_LERN + ' AND "URL(s)[URL Field]" ~ "' + clickUpTaskId + '"';
  const res = jiraRequest_('GET', '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) + '&maxResults=5&fields=issuetype,summary');
  if (!res || !res.issues || res.issues.length === 0) return null;
  return res.issues.map(function (i) { return { key: i.key, issueType: i.fields.issuetype.name, summary: i.fields.summary }; });
}

// Ensures the existing issue (if any) matches the expected type. If it
// matches, returns its key (already correct, nothing to do). If it's the
// WRONG type — the exact scenario of an old Task existing where an Epic
// should now be — deletes it and returns null so the caller creates fresh.
function reconcileExistingIssue_(clickUpTaskId, expectedType) {
  const matches = findExistingJiraIssue_(clickUpTaskId);
  if (!matches) return null;

  const correct = matches.find(function (m) { return m.issueType === expectedType; });
  if (correct) return correct.key;

  matches.forEach(function (wrong) {
    Logger.log('🗑️  Found "' + wrong.summary + '" (' + wrong.key + ', type: ' + wrong.issueType + ') — wrong type, expected ' + expectedType + '. Deleting.');
    const res = UrlFetchApp.fetch(jiraSiteUrl_() + '/rest/api/3/issue/' + wrong.key + '?deleteSubtasks=true', {
      method: 'delete', headers: { 'Authorization': jiraAuthHeader_() }, muteHttpExceptions: true,
    });
    Logger.log((res.getResponseCode() < 300 ? '✅ Deleted ' : '❌ Failed to delete ') + wrong.key);
  });
  return null;
}

const PLAN_TAB_NAME = 'Nested Cleanup Plan';

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

// ==================== PHASE 1: PLAN (dry run — no Jira writes) ====================

function planNestedTreeCleanup() {
  Logger.log('=== Planning nested tree cleanup (dry run — no Jira writes) ===');
  const rootId = getProp_('ROOT_CLICKUP_TASK_ID');
  if (!rootId) { Logger.log('❌ Set ROOT_CLICKUP_TASK_ID in Script Properties first.'); return; }

  const rootTask = clickUpRequest_('/task/' + rootId);
  if (!rootTask || !rootTask.id) { Logger.log('❌ Could not fetch root task ' + rootId); return; }

  const allTasks = fetchAllTasksInList_(rootTask.list.id);
  const depth1Tasks = allTasks.filter(function (t) { return t.parent === rootId; });

  const sheet = getOrCreateTab_(PLAN_TAB_NAME);
  clearAndHeader_(sheet, [
    'ClickUp Task ID', 'Task Name', 'Depth', 'Planned Action', 'Parent ClickUp ID',
    'Existing Match Key', 'Existing Match Type', 'Delete Needed', 'Jira Issue Key', 'Result', 'Failure Reason',
  ]);

  planRow_(sheet, rootTask, 0, 'CREATE EPIC', '');
  let depth3PlusCount = 0;

  depth1Tasks.forEach(function (depth1Task) {
    planRow_(sheet, depth1Task, 1, 'CREATE TASK', rootId);

    const depth2Tasks = allTasks.filter(function (t) { return t.parent === depth1Task.id; });
    depth2Tasks.forEach(function (depth2Task) {
      planRow_(sheet, depth2Task, 2, 'CREATE SUBTASK', depth1Task.id);

      const depth3Tasks = allTasks.filter(function (t) { return t.parent === depth2Task.id; });
      if (depth3Tasks.length > 0) {
        depth3PlusCount += depth3Tasks.length;
        depth3Tasks.forEach(function (depth3Task) {
          planRow_(sheet, depth3Task, 3, 'SKIP - MANUAL (Jira cannot nest this deep)', depth2Task.id);
        });
      }
    });
  });

  Logger.log('✅ Plan written to "' + PLAN_TAB_NAME + '" — review before running executeNestedTreeCleanup().');
  Logger.log('Check the "Delete Needed" column especially — anything marked Y will be DELETED from Jira on execute.');
  if (depth3PlusCount > 0) Logger.log('⚠️ ' + depth3PlusCount + ' task(s) at depth 3+ flagged SKIP - MANUAL — these need manual recreation regardless.');
}

function planRow_(sheet, task, depth, plannedAction, parentClickUpId) {
  let existingKey = '', existingType = '', deleteNeeded = 'N';

  if (plannedAction !== 'SKIP - MANUAL (Jira cannot nest this deep)') {
    const expectedType = depth === 0 ? 'Epic' : (depth === 1 ? 'Task' : 'Sub-Task');
    const matches = findExistingJiraIssue_(task.id);
    if (matches) {
      const correct = matches.find(function (m) { return m.issueType === expectedType; });
      if (correct) {
        existingKey = correct.key;
        existingType = correct.issueType;
        plannedAction = 'ALREADY CORRECT — skip';
      } else {
        existingKey = matches.map(function (m) { return m.key; }).join(', ');
        existingType = matches.map(function (m) { return m.issueType; }).join(', ');
        deleteNeeded = 'Y';
      }
    }
  }

  sheet.appendRow([task.id, task.name, depth, plannedAction, parentClickUpId, existingKey, existingType, deleteNeeded, '', '', '']);
}

// ==================== PHASE 2: EXECUTE (live — reads the plan, does the writes) ====================

function executeNestedTreeCleanup() {
  Logger.log('=== Executing nested tree cleanup (LIVE) ===');
  const sheet = getOrCreateTab_(PLAN_TAB_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  const clickUpIdToJiraKey = {};

  // Pass 1: deletions (wrong-type duplicates), then Epic (depth 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col['Delete Needed']] === 'Y') {
      String(row[col['Existing Match Key']]).split(', ').forEach(function (key) {
        Logger.log('🗑️  Deleting wrong-type duplicate: ' + key);
        const res = UrlFetchApp.fetch(jiraSiteUrl_() + '/rest/api/3/issue/' + key + '?deleteSubtasks=true', {
          method: 'delete', headers: { 'Authorization': jiraAuthHeader_() }, muteHttpExceptions: true,
        });
        Logger.log((res.getResponseCode() < 300 ? '✅ Deleted ' : '❌ Failed to delete ') + key);
      });
    }
  }

  // Pass 2: Epic
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][col['Depth']]) !== 0) continue;
    processPlanRow_(sheet, i, col, clickUpIdToJiraKey, null);
  }

  // Pass 3: Tasks (depth 1)
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][col['Depth']]) !== 1) continue;
    const epicKey = clickUpIdToJiraKey[data[i][col['Parent ClickUp ID']]] || data.find(function (r) { return r[col['ClickUp Task ID']] === data[i][col['Parent ClickUp ID']]; })[col['Jira Issue Key']];
    processPlanRow_(sheet, i, col, clickUpIdToJiraKey, epicKey);
  }

  // Pass 4: Sub-tasks (depth 2)
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][col['Depth']]) !== 2) continue;
    const taskKey = clickUpIdToJiraKey[data[i][col['Parent ClickUp ID']]];
    processPlanRow_(sheet, i, col, clickUpIdToJiraKey, taskKey);
  }

  Logger.log('=== Execute complete. Check the plan sheet for per-row Result. ===');
}

function processPlanRow_(sheet, i, col, clickUpIdToJiraKey, parentKey) {
  const rowNum = i + 1;
  const action = sheet.getRange(rowNum, col['Planned Action'] + 1).getValue();
  const clickUpId = sheet.getRange(rowNum, col['ClickUp Task ID'] + 1).getValue();
  const taskName = sheet.getRange(rowNum, col['Task Name'] + 1).getValue();

  if (String(action).indexOf('SKIP') === 0) return;

  if (String(action).indexOf('ALREADY CORRECT') === 0) {
    const existingKey = sheet.getRange(rowNum, col['Existing Match Key'] + 1).getValue();
    clickUpIdToJiraKey[clickUpId] = existingKey;
    sheet.getRange(rowNum, col['Jira Issue Key'] + 1).setValue(existingKey);
    sheet.getRange(rowNum, col['Result'] + 1).setValue('Already existed');
    return;
  }

  const task = { id: clickUpId, name: taskName, description: '' }; // re-fetch for full description
  const fullTask = clickUpRequest_('/task/' + clickUpId);
  if (fullTask) task.description = fullTask.description || '';

  let resultKey = null;
  if (action === 'CREATE EPIC') resultKey = createEpic_(task);
  else if (action === 'CREATE TASK') resultKey = createTask_(task, parentKey);
  else if (action === 'CREATE SUBTASK') resultKey = createSubtask_(task, parentKey);

  if (resultKey) {
    clickUpIdToJiraKey[clickUpId] = resultKey;
    sheet.getRange(rowNum, col['Jira Issue Key'] + 1).setValue(resultKey);
    sheet.getRange(rowNum, col['Result'] + 1).setValue('Success');
  } else {
    sheet.getRange(rowNum, col['Result'] + 1).setValue('Failed');
    sheet.getRange(rowNum, col['Failure Reason'] + 1).setValue('See Apps Script log for HTTP error detail');
  }
}

function fetchAllTasksInList_(listId) {
  let allTasks = [];
  let page = 0;
  while (true) {
    const res = clickUpRequest_('/list/' + listId + '/task?archived=false&include_closed=true&subtasks=true&page=' + page);
    if (!res || !res.tasks || res.tasks.length === 0) break;
    allTasks = allTasks.concat(res.tasks);
    if (res.last_page) break;
    page++;
  }
  return allTasks;
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

function computeDepthLocal_(task, byId) {
  let depth = 0;
  let current = task;
  while (current.parent && byId[current.parent]) {
    depth++;
    current = byId[current.parent];
  }
  return depth;
}

// ==================== PHASE 0: SCAN FOR NESTED TREES ====================
// Walks every List in your Mapping tab's Space rows (same scope as the
// main migration), finds every task at depth 2+, and traces each one
// back to its depth-0 root — so you get one list of trees needing
// cleanup, instead of hunting for them one at a time.

function scanForNestedTrees() {
  Logger.log('=== Scanning for nested trees (depth >= 2) ===');
  const mappingSheet = ss_().getSheetByName(getProp_('MAPPING_TAB', 'Mapping'));
  if (!mappingSheet) { Logger.log('❌ Mapping tab not found.'); return; }

  const data = mappingSheet.getDataRange().getValues();
  const spaceRows = data.filter(function (row) { return row[0] === 'Space'; });
  const teamId = getProp_('CLICKUP_TEAM_ID');
  const foundRoots = {}; // rootId -> {name, listName, maxDepth, count}

  spaceRows.forEach(function (row) {
    const listId = findListId_(teamId, row[1], row[2]);
    if (!listId) { Logger.log('⚠️ List not found: ' + row[2] + ' — skipping'); return; }

    const allTasks = fetchAllTasksInList_(listId);
    const byId = {};
    allTasks.forEach(function (t) { byId[t.id] = t; });

    allTasks.forEach(function (task) {
      const depth = computeDepthLocal_(task, byId);
      if (depth < 2) return;

      let current = task;
      while (current.parent && byId[current.parent]) current = byId[current.parent];
      const rootId = current.id;

      if (!foundRoots[rootId]) {
        foundRoots[rootId] = { name: current.name, listName: row[2], maxDepth: depth, count: 1 };
      } else {
        foundRoots[rootId].count++;
        if (depth > foundRoots[rootId].maxDepth) foundRoots[rootId].maxDepth = depth;
      }
    });
  });

  const sheet = getOrCreateTab_('Nested Trees Found');
  clearAndHeader_(sheet, ['Root ClickUp Task ID', 'Root Task Name', 'Root Task URL', 'List', 'Max Depth Found', 'Depth 2+ Task Count']);
  Object.keys(foundRoots).forEach(function (rootId) {
    const r = foundRoots[rootId];
    sheet.appendRow([rootId, r.name, clickUpUrl_(rootId), r.listName, r.maxDepth, r.count]);
  });

  Logger.log('✅ Found ' + Object.keys(foundRoots).length + ' root task(s) with nesting Jira cannot fully represent.');
  Logger.log('Next: set ROOT_CLICKUP_TASK_ID to each one in the "Nested Trees Found" tab, in turn, and run planNestedTreeCleanup() → executeNestedTreeCleanup() for each.');
}

// ==================== ISSUE CREATION ====================

function createEpic_(task) {
  const fields = {
    project: { key: PROJECT_KEY_LERN },
    issuetype: { name: 'Epic' },
    summary: task.name,
    description: toADF_(task.description || ''),
  };
  fields[EPIC_NAME_FIELD_ID] = task.name; // required on Epic create
  fields[CLICKUP_URL_FIELD_ID] = toADF_(clickUpUrl_(task.id));
  applyOptionalTargets_(fields);

  const result = jiraRequest_('POST', '/rest/api/3/issue', { fields: fields });
  if (result && result.key) { Logger.log('✅ Created Epic: ' + result.key + ' — ' + task.name); return result.key; }
  Logger.log('❌ Failed to create Epic for ' + task.id);
  return null;
}

function createTask_(task, epicKey) {
  const fields = {
    project: { key: PROJECT_KEY_LERN },
    issuetype: { name: 'Task' },
    summary: task.name,
    description: toADF_(task.description || ''),
  };
  fields[EPIC_LINK_FIELD_ID] = epicKey; // classic Epic Link takes the plain issue key string
  fields[CLICKUP_URL_FIELD_ID] = toADF_(clickUpUrl_(task.id));
  applyOptionalTargets_(fields);

  const result = jiraRequest_('POST', '/rest/api/3/issue', { fields: fields });
  if (result && result.key) { Logger.log('✅ Created Task: ' + result.key + ' — ' + task.name + ' (under Epic ' + epicKey + ')'); return result.key; }
  Logger.log('❌ Failed to create Task for ' + task.id);
  return null;
}

function createSubtask_(task, parentTaskKey) {
  const fields = {
    project: { key: PROJECT_KEY_LERN },
    issuetype: { name: 'Sub-Task' }, // confirmed exact casing from your Jira Schema export
    summary: task.name,
    description: toADF_(task.description || ''),
    parent: { key: parentTaskKey },
  };
  fields[CLICKUP_URL_FIELD_ID] = toADF_(clickUpUrl_(task.id));
  applyOptionalTargets_(fields);

  const result = jiraRequest_('POST', '/rest/api/3/issue', { fields: fields });
  if (result && result.key) { Logger.log('✅ Created Sub-task: ' + result.key + ' — ' + task.name + ' (under Task ' + parentTaskKey + ')'); return result.key; }
  Logger.log('❌ Failed to create Sub-task for ' + task.id);
  return null;
}

function applyOptionalTargets_(fields) {
  const component = getProp_('TARGET_COMPONENT');
  if (component) fields.components = [{ name: component }];
  const category = getProp_('TARGET_CATEGORY');
  if (category) fields['customfield_10130'] = { value: category };
}

function toADF_(text) {
  if (!text) return { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [] }] };
  return { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: text }] }] };
}