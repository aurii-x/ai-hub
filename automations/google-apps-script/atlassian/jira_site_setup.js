/**
 * Jira Full Setup 3.0 — santhoshOS Single Project
 * ─────────────────────────────────────────
 * Version 3.0  |  July 2026
 *
 * FIELD DESIGN (final, confirmed)
 *   Category  — custom single-select field, 14 granular values
 *               (1.1 Deep Work ... 6.3 Client Work) — YOU set this manually
 *   Component — native Jira field, 4 top-level values
 *               (1.x Career Development, 2.x Learning, 3.x Family & Life,
 *               6.x Straventis) — set AUTOMATICALLY by JIRA_update_Scheduled.gs
 *               based on which Category you picked. Not done here —
 *               this script only creates the fields/values; the sync
 *               logic lives in the companion script.
 *
 * FUNCTIONS (in dependency order, not necessarily your original numbering)
 *   populateComponent()        — creates the 4 native Components
 *   populateCategory()         — creates the Category custom field + 14 options
 *   setupStatuses()            — 6 statuses across 3 categories
 *   setupWorkFlows()           — 1 shared workflow, all-to-all transitions
 *   setupWorkflowSchemes()     — scheme wrapping the workflow above
 *   setupWorkTypes()           — confirms Epic/Task/Sub-task exist (they're
 *                                 Jira defaults, nothing to create)
 *   setupWorkTypeSchema()      — restricts the project to Epic/Task/Sub-task,
 *                                 excludes Story/Bug (this IS the hierarchy
 *                                 control — see note in the function)
 *   setupPriorities()          — combined with scheme per your instruction:
 *                                 low/normal/high/urgent, old defaults removed
 *   setupScreens()             — 3-tab screen (Default/Space/Relational)
 *   setupScreenSchemes()       — scheme wrapping the screen above
 *   setupWorkTypeScreenSchemes() — scheme wrapping the screen scheme above
 *   createProject()            — single project, all schemes pre-assigned
 *   runFullJiraSetup()         — orchestrator, calls everything in order
 *   factoryResetJira()         — wipes it all, no dry run, as established
 *
 * NOTE ON WORK TYPE HIERARCHY (function 5 in your original list)
 * There's no separate "hierarchy" API call — hierarchy is just each work
 * type's fixed hierarchyLevel (-1 Sub-task, 0 Task, 1 Epic), which Jira
 * already has. Excluding Story and Bug via the Work Type Scheme
 * (setupWorkTypeSchema) is what actually achieves "Epic top level, Task
 * and Sub-task enabled, Story and Bug disabled" — there's nothing further
 * to configure once that scheme is applied.
 *
 * CREDENTIALS (Script Properties)
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN
 *
 * RUN ORDER
 *   checkSetup() → testConnection() → runFullJiraSetup()
 *
 * KNOWN LIMITATION (carried over, still true)
 * Jira Software Kanban boards default to a locked "Simplified workflow."
 * If the project is still on it after this runs, one manual step:
 * Board Settings → Columns → Workflow type → switch off "Simplified".
 */

// ==================== CONFIG ====================

const PROJECT_KEY = 'OS';
const PROJECT_NAME = 'santhoshOS';
const PROJECT_TYPE_KEY = 'software';
const PROJECT_TEMPLATE_KEY = 'com.pyxis.greenhopper.jira:gh-kanban-template';

const COMPONENTS = [
  '1.x Career Development', '2.x Learning', '3.x Family & Life', '6.x Straventis',
];

const CATEGORY_FIELD_NAME = 'Category';
const CATEGORY_OPTIONS = [
  '1.1 Deep Work', '1.2 Job Search', '1.3 Resume Updates',
  '2.1 Certs & Courses', '2.2 Research', '2.3 Automation',
  '3.1 Kids', '3.2 Health', '3.3 Finance', '3.4 Legal', '3.5 Shopping',
  '6.1 Admin', '6.2 Webspace', '6.3 Client Work',
];

const WORK_TYPE_SCHEME_NAME = 'santhoshOS Work Type Scheme';
const ALLOWED_WORK_TYPES = ['Epic', 'Task', 'Sub-task']; // Story, Bug excluded

const STATUS_CONFIG = [
  { name: 'New', category: 'TODO' },
  { name: 'Backlog', category: 'TODO' },
  { name: 'In Progress', category: 'IN_PROGRESS' },
  { name: 'Deferred', category: 'IN_PROGRESS' },
  { name: 'Done', category: 'DONE' },
  { name: 'Archived', category: 'DONE' },
];
const WORKFLOW_NAME = 'santhoshos default project workflow';
const WORKFLOW_SCHEME_NAME = 'santhoshOS Workflow Scheme';

const PRIORITY_CONFIG = ['Low', 'Normal', 'High', 'Urgent'];
const DEFAULT_PRIORITY_NAME = 'Normal';

const SCREEN_NAME = 'santhoshOS Work Item Screen';
const SCREEN_SCHEME_NAME = 'santhoshOS Screen Scheme';
const ISSUE_TYPE_SCREEN_SCHEME_NAME = 'santhoshOS Issue Type Screen Scheme';

const SYSTEM_FIELD_TABS = [
  { fieldId: 'summary', tab: 'default' },
  { fieldId: 'description', tab: 'default' },
  { fieldId: 'labels', tab: 'default' },
  { fieldId: 'components', tab: 'default' },
  { fieldId: 'timetracking', tab: 'default' },
  { fieldId: 'duedate', tab: 'default' },
  { fieldId: 'parent', tab: 'relational' },
];

// ==================== CREDENTIALS / CORE HELPERS ====================

function checkSetup() {
  const props = PropertiesService.getScriptProperties();
  ['JIRA_SITE_URL', 'JIRA_EMAIL', 'JIRA_API_TOKEN'].forEach(function (key) {
    Logger.log(key + ': ' + (props.getProperty(key) ? '✅ set' : '❌ NOT SET'));
  });
}

function testConnection() {
  const res = jiraRequest_('GET', '/rest/api/3/myself');
  if (res && res.accountId) Logger.log('✅ Connected as: ' + res.displayName);
  else Logger.log('❌ Connection failed. Check checkSetup().');
}

function getAuthHeader_() {
  const props = PropertiesService.getScriptProperties();
  return 'Basic ' + Utilities.base64Encode(props.getProperty('JIRA_EMAIL') + ':' + props.getProperty('JIRA_API_TOKEN'));
}
function getSiteUrl_() {
  const url = PropertiesService.getScriptProperties().getProperty('JIRA_SITE_URL');
  return url ? url.replace(/\/$/, '') : url;
}
function getMyAccountId_() {
  const me = jiraRequest_('GET', '/rest/api/3/myself');
  return me ? me.accountId : null;
}

function jiraRequest_(method, path, payload) {
  const options = {
    method: method,
    headers: { 'Authorization': getAuthHeader_(), 'Accept': 'application/json' },
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);
  Logger.log('→ ' + method + ' ' + path);
  const response = UrlFetchApp.fetch(getSiteUrl_() + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  if (code === 404) { Logger.log('  (404 — not found)'); return null; }
  Logger.log('  ❌ HTTP ' + code + ': ' + body);
  return null;
}

function findByName_(path, listKey, name) {
  const res = jiraRequest_('GET', path);
  const list = res && res[listKey] ? res[listKey] : (Array.isArray(res) ? res : []);
  return list.find(function (item) { return item.name === name; }) || null;
}

// ==================== ORCHESTRATOR ====================

function runFullJiraSetup() {
  Logger.log('========== SANTHOSHOS JIRA SETUP 3.0 ==========');
  setupStatuses();
  const workflowName = setupWorkFlows();
  const workflowSchemeId = setupWorkflowSchemes(workflowName);
  const categoryFieldId = populateCategory();
  const screenInfo = setupScreens(categoryFieldId);
  const screenSchemeId = setupScreenSchemes(screenInfo.screenId);
  const issueTypeScreenSchemeId = setupWorkTypeScreenSchemes(screenSchemeId);
  setupWorkTypes();
  const workTypeSchemeId = setupWorkTypeSchema();
  setupPriorities();

  const project = createProject(workflowSchemeId, issueTypeScreenSchemeId, workTypeSchemeId);
  if (project) {
    populateComponent(project.key);
    wireSchemeToProject_('/rest/api/3/workflowscheme/project', 'workflowSchemeId', workflowSchemeId, project.id);
    wireSchemeToProject_('/rest/api/3/issuetypescreenscheme/project', 'issueTypeScreenSchemeId', issueTypeScreenSchemeId, project.id);
  }

  Logger.log('========== DONE ==========');
  Logger.log('Manual follow-up:');
  Logger.log('  1. Check Board Settings → Columns — if still "Simplified workflow", switch it off manually.');
  Logger.log('  2. Category field created with 14 options — Component sync happens in JIRA_update_Scheduled.gs, not here.');
  Logger.log('  3. Check Settings → Issues → Statuses for any duplicates from workflow creation.');
}

// ==================== COMPONENTS (native, 4 top-level) ====================

function populateComponent(projectKey) {
  Logger.log('--- populateComponent() ---');
  const key = projectKey || PROJECT_KEY;
  const existing = jiraRequest_('GET', '/rest/api/3/project/' + key + '/components') || [];
  COMPONENTS.forEach(function (name) {
    if (existing.find(function (c) { return c.name === name; })) { Logger.log('⏭️  Component "' + name + '" already exists — skipping.'); return; }
    const result = jiraRequest_('POST', '/rest/api/3/component', { name: name, project: key });
    if (result) Logger.log('✅ Created component: ' + name);
    else Logger.log('❌ Failed to create component: ' + name);
  });
}

// ==================== CATEGORY (custom field, 14 granular values) ====================

function populateCategory() {
  Logger.log('--- populateCategory() ---');
  const existingFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  let field = existingFields.find(function (f) { return f.name === CATEGORY_FIELD_NAME; });

  if (!field) {
    field = jiraRequest_('POST', '/rest/api/3/field', {
      name: CATEGORY_FIELD_NAME,
      description: 'Granular sub-area — manually set. Component is auto-derived from this via JIRA_update_Scheduled.gs',
      type: 'com.atlassian.jira.plugin.system.customfieldtypes:select',
      searcherKey: 'com.atlassian.jira.plugin.system.customfieldtypes:selectsearcher',
    });
    if (field) Logger.log('✅ Created field: ' + CATEGORY_FIELD_NAME);
    else { Logger.log('❌ Failed to create Category field.'); return null; }
  } else {
    Logger.log('⏭️  Field "' + CATEGORY_FIELD_NAME + '" already exists — skipping creation.');
  }

  const contexts = jiraRequest_('GET', '/rest/api/3/field/' + field.id + '/context');
  const contextId = contexts && contexts.values && contexts.values[0] ? contexts.values[0].id : null;
  if (!contextId) { Logger.log('❌ No context found for Category field — options not added.'); return field.id; }

  const existingOptions = jiraRequest_('GET', '/rest/api/3/field/' + field.id + '/context/' + contextId + '/option');
  const existingValues = existingOptions && existingOptions.values ? existingOptions.values.map(function (o) { return o.value; }) : [];
  const toAdd = CATEGORY_OPTIONS.filter(function (o) { return existingValues.indexOf(o) === -1; });

  if (toAdd.length > 0) {
    const result = jiraRequest_('POST', '/rest/api/3/field/' + field.id + '/context/' + contextId + '/option', {
      options: toAdd.map(function (v) { return { value: v }; }),
    });
    if (result) Logger.log('✅ Added ' + toAdd.length + ' Category options.');
  } else {
    Logger.log('⏭️  All Category options already present.');
  }

  return field.id;
}

// ==================== STATUSES ====================

function setupStatuses() {
  Logger.log('--- setupStatuses() ---');
  const existing = jiraRequest_('GET', '/rest/api/3/statuses/search?maxResults=200');
  const existingNames = existing && existing.values ? existing.values.map(function (s) { return s.name; }) : [];
  const toCreate = STATUS_CONFIG.filter(function (s) { return existingNames.indexOf(s.name) === -1; });

  if (toCreate.length > 0) {
    const result = jiraRequest_('POST', '/rest/api/3/statuses', {
      scope: { type: 'GLOBAL' },
      statuses: toCreate.map(function (s) { return { name: s.name, statusCategory: s.category, description: '' }; }),
    });
    if (result) result.forEach(function (s) { Logger.log('✅ Created status: ' + s.name); });
    else Logger.log('❌ Status creation failed.');
  }
  STATUS_CONFIG.forEach(function (s) {
    if (existingNames.indexOf(s.name) !== -1) Logger.log('⏭️  Status "' + s.name + '" already exists — skipping.');
  });
}

// ==================== WORKFLOW + SCHEME ====================

function setupWorkFlows() {
  Logger.log('--- setupWorkFlows() ---');
  const existing = jiraRequest_('GET', '/rest/api/3/workflows/search?workflowName=' + encodeURIComponent(WORKFLOW_NAME));
  const exactMatch = existing && existing.values ? existing.values.find(function (w) { return w.id && w.id.name === WORKFLOW_NAME; }) : null;
  if (exactMatch) { Logger.log('⏭️  Workflow already exists — skipping.'); return WORKFLOW_NAME; }

  const statusRefs = STATUS_CONFIG.map(function () { return Utilities.getUuid(); });
  const payload = {
    scope: { type: 'GLOBAL' },
    statuses: STATUS_CONFIG.map(function (s, i) {
      return { name: s.name, statusCategory: s.category, description: '', statusReference: statusRefs[i] };
    }),
    workflows: [{
      name: WORKFLOW_NAME,
      description: 'Single shared workflow across all santhoshOS work types',
      statuses: STATUS_CONFIG.map(function (s, i) { return { statusReference: statusRefs[i], properties: {} }; }),
      transitions: (function () {
        const t = []; let id = 1;
        STATUS_CONFIG.forEach(function (from, i) {
          STATUS_CONFIG.forEach(function (to, j) {
            if (i === j) return;
            t.push({ id: String(id++), name: 'To ' + to.name, to: statusRefs[j], from: [statusRefs[i]], type: 'DIRECTED' });
          });
        });
        return t;
      })(),
    }],
  };

  const result = jiraRequest_('POST', '/rest/api/3/workflows/create', payload);
  if (result) Logger.log('✅ Created workflow: ' + WORKFLOW_NAME);
  else Logger.log('❌ Workflow creation failed.');
  return WORKFLOW_NAME;
}

function setupWorkflowSchemes(workflowName) {
  Logger.log('--- setupWorkflowSchemes() ---');
  const existing = findByName_('/rest/api/3/workflowscheme?maxResults=100', 'values', WORKFLOW_SCHEME_NAME);
  if (existing) { Logger.log('⏭️  Workflow scheme already exists — skipping.'); return existing.id; }
  const result = jiraRequest_('POST', '/rest/api/3/workflowscheme', {
    name: WORKFLOW_SCHEME_NAME, description: 'santhoshOS shared workflow scheme', defaultWorkflow: workflowName,
  });
  if (result) Logger.log('✅ Created workflow scheme: ' + WORKFLOW_SCHEME_NAME);
  return result ? result.id : null;
}

// ==================== WORK TYPES + SCHEMA (hierarchy control) ====================

function setupWorkTypes() {
  Logger.log('--- setupWorkTypes() ---');
  const allTypes = jiraRequest_('GET', '/rest/api/3/issuetype') || [];
  ALLOWED_WORK_TYPES.forEach(function (name) {
    const t = allTypes.find(function (x) { return x.name === name; });
    Logger.log((t ? '⏭️  ' : '⚠️ ') + name + (t ? ' exists (Jira default, hierarchyLevel ' + t.hierarchyLevel + ')' : ' NOT FOUND by name — check spelling'));
  });
}

function setupWorkTypeSchema() {
  Logger.log('--- setupWorkTypeSchema() ---');
  const allTypes = jiraRequest_('GET', '/rest/api/3/issuetype') || [];
  const allowedIds = ALLOWED_WORK_TYPES.map(function (name) {
    const t = allTypes.find(function (x) { return x.name === name; });
    return t ? t.id : null;
  }).filter(function (id) { return id; });
  const taskType = allTypes.find(function (t) { return t.name === 'Task'; });

  let scheme = findByName_('/rest/api/3/issuetypescheme/search?maxResults=100', 'values', WORK_TYPE_SCHEME_NAME);
  if (scheme) { Logger.log('⏭️  Work type scheme already exists — skipping.'); return scheme.id; }

  const result = jiraRequest_('POST', '/rest/api/3/issuetypescheme', {
    name: WORK_TYPE_SCHEME_NAME,
    description: 'Epic, Task, Sub-task only — Story and Bug excluded (this is what disables them)',
    issueTypeIds: allowedIds,
    defaultIssueTypeId: taskType ? taskType.id : undefined,
  });
  if (result) Logger.log('✅ Created work type scheme: ' + WORK_TYPE_SCHEME_NAME);
  return result ? (result.issueTypeSchemeId || result.id) : null;
}

// ==================== PRIORITIES (combined with scheme, per instruction) ====================

function setupPriorities() {
  Logger.log('--- setupPriorities() (combined with scheme) ---');
  const existing = jiraRequest_('GET', '/rest/api/3/priority') || [];
  const existingNames = existing.map(function (p) { return p.name; });

  PRIORITY_CONFIG.forEach(function (name) {
    if (existingNames.indexOf(name) !== -1) { Logger.log('⏭️  Priority "' + name + '" already exists — skipping.'); return; }
    const result = jiraRequest_('POST', '/rest/api/3/priority', { name: name, description: '', iconUrl: 'https://cdn.prod.website-files.com/6258e1e2f74a2d178f31af0c/652890889c6ceb254e6b6da2_icon-priority-medium.svg' });
    if (result) Logger.log('✅ Created priority: ' + name);
    else Logger.log('❌ Failed to create priority: ' + name);
  });

  // Remove Jira's original defaults so only our 4 remain available
  const stillDefault = jiraRequest_('GET', '/rest/api/3/priority') || [];
  const toRemove = stillDefault.filter(function (p) { return PRIORITY_CONFIG.indexOf(p.name) === -1; });
  toRemove.forEach(function (p) {
    const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/priority/' + p.id, {
      method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
    });
    Logger.log((res.getResponseCode() < 300 ? '🗑️  Removed default priority: ' : '❌ Could not remove (may be in use): ') + p.name);
  });

  // Set default priority
  const finalList = jiraRequest_('GET', '/rest/api/3/priority') || [];
  const normal = finalList.find(function (p) { return p.name === DEFAULT_PRIORITY_NAME; });
  if (normal) {
    jiraRequest_('PUT', '/rest/api/3/priority/default', { id: normal.id });
    Logger.log('✅ Set default priority: ' + DEFAULT_PRIORITY_NAME);
  }
}

// ==================== SCREEN / SCREEN SCHEME / WORK TYPE SCREEN SCHEME ====================

function setupScreens(categoryFieldId) {
  Logger.log('--- setupScreens() ---');
  const existingScreens = jiraRequest_('GET', '/rest/api/3/screens?maxResults=100');
  let screen = existingScreens && existingScreens.values ? existingScreens.values.find(function (s) { return s.name === SCREEN_NAME; }) : null;
  if (!screen) {
    screen = jiraRequest_('POST', '/rest/api/3/screens', { name: SCREEN_NAME, description: 'santhoshOS work item fields' });
    if (screen) Logger.log('✅ Created screen: ' + SCREEN_NAME);
  } else {
    Logger.log('⏭️  Screen already exists — skipping.');
  }
  if (!screen || !screen.id) return {};

  const tabNames = { default: 'Default Properties', space: 'Space Properties', relational: 'Relational Properties' };
  const existingTabs = jiraRequest_('GET', '/rest/api/3/screens/' + screen.id + '/tabs') || [];
  const tabIds = {};
  Object.keys(tabNames).forEach(function (key) {
    let tab = existingTabs.find(function (t) { return t.name === tabNames[key]; });
    if (!tab) { tab = jiraRequest_('POST', '/rest/api/3/screens/' + screen.id + '/tabs', { name: tabNames[key] }); if (tab) Logger.log('✅ Created tab: ' + tabNames[key]); }
    else Logger.log('⏭️  Tab "' + tabNames[key] + '" already exists — skipping.');
    if (tab) tabIds[key] = tab.id;
  });

  SYSTEM_FIELD_TABS.forEach(function (sf) { addFieldToTab_(screen.id, tabIds[sf.tab], sf.fieldId); });
  if (categoryFieldId) addFieldToTab_(screen.id, tabIds.default, categoryFieldId);
  wireNativeStartDateField_(screen.id, tabIds);
  wireEpicNameField_(screen.id, tabIds);

  return { screenId: screen.id, tabIds: tabIds };
}

function addFieldToTab_(screenId, tabId, fieldId) {
  if (!screenId || !tabId || !fieldId) return;
  const existingFields = jiraRequest_('GET', '/rest/api/3/screens/' + screenId + '/tabs/' + tabId + '/fields') || [];
  if (existingFields.find(function (f) { return f.id === fieldId; })) { Logger.log('⏭️  Field ' + fieldId + ' already on tab — skipping.'); return; }
  const result = jiraRequest_('POST', '/rest/api/3/screens/' + screenId + '/tabs/' + tabId + '/fields', { fieldId: fieldId });
  if (result) Logger.log('✅ Added ' + fieldId + ' to tab ' + tabId);
}

function wireNativeStartDateField_(screenId, tabIds) {
  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  const f = allFields.find(function (x) { return x.name === 'Start date'; });
  if (f) addFieldToTab_(screenId, tabIds.default, f.id);
  else Logger.log('⚠️ Native "Start date" field not found by name.');
}

function wireEpicNameField_(screenId, tabIds) {
  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  const f = allFields.find(function (x) { return x.name === 'Epic Name'; });
  if (f) { addFieldToTab_(screenId, tabIds.default, f.id); addFieldToTab_(screenId, tabIds.relational, f.id); }
  else Logger.log('⚠️ "Epic Name" not found yet — appears once an Epic exists. Re-run setupScreens() after creating one.');
}

function setupScreenSchemes(screenId) {
  Logger.log('--- setupScreenSchemes() ---');
  if (!screenId) return null;
  let scheme = findByName_('/rest/api/3/screenscheme?maxResults=100', 'values', SCREEN_SCHEME_NAME);
  if (!scheme) {
    scheme = jiraRequest_('POST', '/rest/api/3/screenscheme', { name: SCREEN_SCHEME_NAME, screens: { default: screenId } });
    if (scheme) Logger.log('✅ Created screen scheme: ' + SCREEN_SCHEME_NAME);
  } else {
    Logger.log('⏭️  Screen scheme already exists — skipping.');
  }
  return scheme ? scheme.id : null;
}

function setupWorkTypeScreenSchemes(screenSchemeId) {
  Logger.log('--- setupWorkTypeScreenSchemes() ---');
  if (!screenSchemeId) return null;
  let itss = findByName_('/rest/api/3/issuetypescreenscheme?maxResults=100', 'values', ISSUE_TYPE_SCREEN_SCHEME_NAME);
  if (!itss) {
    itss = jiraRequest_('POST', '/rest/api/3/issuetypescreenscheme', {
      name: ISSUE_TYPE_SCREEN_SCHEME_NAME, issueTypeMappings: [{ issueTypeId: 'default', screenSchemeId: screenSchemeId }],
    });
    if (itss) Logger.log('✅ Created work type screen scheme: ' + ISSUE_TYPE_SCREEN_SCHEME_NAME);
  } else {
    Logger.log('⏭️  Work type screen scheme already exists — skipping.');
  }
  return itss ? itss.id : null;
}

// ==================== PROJECT ====================

function createProject(workflowSchemeId, issueTypeScreenSchemeId, issueTypeSchemeId) {
  Logger.log('--- createProject() ---');
  const accountId = getMyAccountId_();
  let project = jiraRequest_('GET', '/rest/api/3/project/' + PROJECT_KEY);
  if (project) { Logger.log('⏭️  Project ' + PROJECT_KEY + ' already exists — skipping.'); return project; }

  const payload = {
    key: PROJECT_KEY, name: PROJECT_NAME, projectTypeKey: PROJECT_TYPE_KEY,
    projectTemplateKey: PROJECT_TEMPLATE_KEY, leadAccountId: accountId,
  };
  if (workflowSchemeId) payload.workflowSchemeId = workflowSchemeId;
  if (issueTypeScreenSchemeId) payload.issueTypeScreenSchemeId = issueTypeScreenSchemeId;
  if (issueTypeSchemeId) payload.issueTypeSchemeId = issueTypeSchemeId;

  project = jiraRequest_('POST', '/rest/api/3/project', payload);
  if (project) Logger.log('✅ Created project: ' + PROJECT_KEY);
  else Logger.log('❌ Failed to create project.');
  return project;
}

function wireSchemeToProject_(path, bodyKey, schemeId, projectId) {
  if (!schemeId || !projectId) return;
  const body = {}; body[bodyKey] = schemeId; body.projectId = projectId;
  const res = UrlFetchApp.fetch(getSiteUrl_() + path, {
    method: 'put', headers: { 'Authorization': getAuthHeader_(), 'Content-Type': 'application/json' },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  Logger.log((res.getResponseCode() < 300 ? '✅' : '❌') + ' ' + path);
}

// ==================== FACTORY RESET ====================

function factoryResetJira() {
  Logger.log('========== ⚠️ FACTORY RESET ==========');
  const projects = jiraRequest_('GET', '/rest/api/3/project/search?maxResults=200');
  if (projects && projects.values) {
    projects.values.forEach(function (p) {
      const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/project/' + p.key, { method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true });
      Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted project ' : '❌ Failed to delete ') + p.key);
    });
  }

  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  const catField = allFields.find(function (f) { return f.name === CATEGORY_FIELD_NAME; });
  if (catField) {
    const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/field/' + catField.id, { method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true });
    Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted field ' : '❌ Failed to delete field ') + CATEGORY_FIELD_NAME);
  }

  deleteByName_('/rest/api/3/issuetypescheme/search?maxResults=100', 'values', WORK_TYPE_SCHEME_NAME, '/rest/api/3/issuetypescheme/');
  deleteByName_('/rest/api/3/issuetypescreenscheme?maxResults=100', 'values', ISSUE_TYPE_SCREEN_SCHEME_NAME, '/rest/api/3/issuetypescreenscheme/');
  deleteByName_('/rest/api/3/screenscheme?maxResults=100', 'values', SCREEN_SCHEME_NAME, '/rest/api/3/screenscheme/');

  const screens = jiraRequest_('GET', '/rest/api/3/screens?maxResults=100');
  if (screens && screens.values) {
    const s = screens.values.find(function (x) { return x.name === SCREEN_NAME; });
    if (s) {
      const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/screens/' + s.id, { method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true });
      Logger.log(res.getResponseCode() < 300 ? '🗑️  Deleted screen' : '❌ Failed to delete screen: HTTP ' + res.getResponseCode());
    }
  }

  deleteByName_('/rest/api/3/workflowscheme?maxResults=100', 'values', WORKFLOW_SCHEME_NAME, '/rest/api/3/workflowscheme/');

  Logger.log('========== FACTORY RESET COMPLETE ==========');
  Logger.log('Note: priorities and global statuses are not removed by this reset — clean up manually via Settings if needed.');
}

function deleteByName_(searchPath, listKey, name, deletePathPrefix) {
  const res = jiraRequest_('GET', searchPath);
  const list = res && res[listKey] ? res[listKey] : [];
  const item = list.find(function (x) { return x.name === name; });
  if (!item) return;
  const delRes = UrlFetchApp.fetch(getSiteUrl_() + deletePathPrefix + item.id, { method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true });
  Logger.log((delRes.getResponseCode() < 300 ? '🗑️  Deleted ' : '❌ Failed to delete ') + name + (delRes.getResponseCode() >= 300 ? ': HTTP ' + delRes.getResponseCode() : ''));
}