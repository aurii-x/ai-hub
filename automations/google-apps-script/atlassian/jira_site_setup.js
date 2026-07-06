/**
 * Jira Full Setup — santhoshOS Migration
 * ─────────────────────────────────────────
 * Version 2.0  |  July 2026
 *
 * CHANGELOG
 * v2.4 - Removed Start Date/Time and End Date/Time custom fields —
 *        replaced with native "duedate" (true system field) and
 *        "Start date" (locked, auto-provisioned field, resolved by
 *        name at runtime). Added deleteRedundantCustomFields() as
 *        Phase 0, run first in runFullJiraSetup(), to clean up any
 *        already-created copies from earlier runs.
 * v2.3 - Fixed workflow transitions missing required 'id' field.
 *        Renamed Learning Space key LEARN → LRN to avoid a key
 *        conflict with a leftover trashed project from the original
 *        v1 team-managed run (Jira reserves keys of deleted projects
 *        for a retention period even though GET returns 404 for them).
 * v2.2 - Fixed 4 bugs found on first real run against the API:
 *        (1) statuses endpoint was /rest/api/3/statuses/create (404),
 *            correct path is /rest/api/3/statuses
 *        (2) custom field searcherKey used wrong namespace
 *            (customfieldsearchers → customfieldtypes)
 *        (3) project template key was invalid; corrected to
 *            com.atlassian.jira-core-project-templates:jira-core-project-management
 *        (4) workflow-exists check gave a false positive (found
 *            unrelated search results); now requires exact name match
 * v2.1 - Switched to business project type, Epic dropped entirely.
 *        Hierarchy is now Task → Subtask only. Removed Epic Name
 *        field wiring since Epic no longer exists as a work type.
 * v2.0 - Full rebuild: company-managed projects, custom workflow
 *        (5 statuses), 16 properties across 3 screen tabs, 4 Spaces
 *        with Components, and a global factory-reset function.
 * v1.1 - Credentials moved to Script Properties (see Jira_Space_Setup_v1)
 * v1.0 - Initial minimal team-managed Space creation
 *
 * CREDENTIALS
 * Set as Script Properties (Project Settings → Script Properties):
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN
 *
 * RUN ORDER
 *   checkSetup() → testConnection() → runFullJiraSetup()
 *
 * Each phase inside runFullJiraSetup() is also individually callable
 * for debugging: createStatuses(), createWorkflow(), createCustomFields(),
 * createScreenAndTabs(), createProjectsAndComponents(), wireSchemesToProjects()
 *
 * MANUAL STEPS THIS SCRIPT CANNOT DO
 * - "Goals (Atlassian)" and "Project (Atlassian)" fields: these are
 *   built-in Atlassian Home fields, enabled per issue type via the UI
 *   (issue view → ••• → "Add Goals field" or similar), not creatable
 *   via API. Enable once per project after this script runs.
 * - Status colors beyond the native 3 (grey/blue/green): needs a
 *   marketplace app (e.g. "Status Colors for Jira"), not scriptable.
 */

// ==================== CONFIG: SPACES & COMPONENTS ====================

const SPACE_CONFIG = [
  {
    key: 'CARE', name: 'Career Development',
    components: ['1.1 Deep Work', '1.2 Job Search', '1.3 Resume Updates'],
  },
  {
    key: 'LRN', name: 'Learning',
    components: ['2.1 Certs & Courses', '2.2 Research', '2.3 Automation'],
  },
  {
    key: 'FAM', name: 'Family & Life',
    components: ['3.1 Kids', '3.2 Health', '3.3 Finance', '3.4 Legal', '3.5 Shopping'],
  },
  {
    key: 'STRAV', name: 'Straventis',
    components: ['6.1 Admin', '6.2 Webspace', '6.3 Client Work'],
  },
];

// Classic (company-managed) business template. Hierarchy is Task →
// Subtask only — Jira's classic Business templates do not reliably
// support Epic as a work type (the manual workaround is documented as
// flaky and unfixable without Atlassian Support, unavailable on Free).
const PROJECT_TYPE_KEY = 'business';
const PROJECT_TEMPLATE_KEY = 'com.atlassian.jira-core-project-templates:jira-core-project-management';

// ==================== CONFIG: WORKFLOW STATUSES ====================

const STATUS_CONFIG = [
  { name: 'New', category: 'TODO' },
  { name: 'Planned', category: 'TODO' },
  { name: 'In Progress', category: 'IN_PROGRESS' },
  { name: 'Done', category: 'DONE' },
  { name: 'Archived', category: 'DONE' },
];
const WORKFLOW_NAME = 'santhoshOS Standard Workflow';
const WORKFLOW_SCHEME_NAME = 'santhoshOS Workflow Scheme';

// ==================== CONFIG: CUSTOM FIELDS ====================
// tab: 'default' | 'space' | 'relational' — used when wiring to screen tabs

const CUSTOM_FIELD_CONFIG = [
  { name: 'URL(s)', type: 'com.atlassian.jira.plugin.system.customfieldtypes:textarea',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:textsearcher', tab: 'default' },
  { name: 'Estimated Time (hrs)', type: 'com.atlassian.jira.plugin.system.customfieldtypes:float',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:exactnumber', tab: 'default' },
  { name: 'Tracked Time (hrs)', type: 'com.atlassian.jira.plugin.system.customfieldtypes:float',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:exactnumber', tab: 'default' },
  { name: 'Rize Time Entry ID', type: 'com.atlassian.jira.plugin.system.customfieldtypes:textfield',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:textsearcher', tab: 'space' },
  { name: 'Rize ID', type: 'com.atlassian.jira.plugin.system.customfieldtypes:textfield',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:textsearcher', tab: 'space' },
  { name: 'Confluence Page(s)', type: 'com.atlassian.jira.plugin.system.customfieldtypes:textarea',
    searcher: 'com.atlassian.jira.plugin.system.customfieldtypes:textsearcher', tab: 'relational' },
];

// Fields we deliberately do NOT create as custom fields, because a
// native/locked equivalent already exists — kept here so
// deleteRedundantCustomFields() knows what to clean up if it finds
// versions of these already created by an earlier run.
const REDUNDANT_FIELD_NAMES = ['Start Date/Time', 'End Date/Time'];

// Built-in system fields to place on tabs (fieldId is the literal system key)
const SYSTEM_FIELD_TABS = [
  { fieldId: 'summary', tab: 'default' },
  { fieldId: 'description', tab: 'default' },
  { fieldId: 'labels', tab: 'default' },
  { fieldId: 'components', tab: 'default' },
  { fieldId: 'timetracking', tab: 'default' },       // covers estimate; kept alongside custom hrs fields
  { fieldId: 'duedate', tab: 'default' },             // native system field — true built-in, no custom field needed
  { fieldId: 'parent', tab: 'relational' },
  // "Start date" is a locked, auto-provisioned field (not a true system
  // field, but not something to create either) — resolved by name at
  // runtime in wireNativeStartDateField_(), since its ID isn't a fixed
  // literal like the ones above.
];

const SCREEN_NAME = 'santhoshOS Work Item Screen';
const SCREEN_SCHEME_NAME = 'santhoshOS Screen Scheme';
const ISSUE_TYPE_SCREEN_SCHEME_NAME = 'santhoshOS Issue Type Screen Scheme';

// ==================== CREDENTIALS / CORE HELPERS ====================

function checkSetup() {
  const props = PropertiesService.getScriptProperties();
  const siteUrl = props.getProperty('JIRA_SITE_URL');
  const email = props.getProperty('JIRA_EMAIL');
  const token = props.getProperty('JIRA_API_TOKEN');
  Logger.log('JIRA_SITE_URL: ' + (siteUrl || '❌ NOT SET'));
  Logger.log('JIRA_EMAIL: ' + (email || '❌ NOT SET'));
  Logger.log('JIRA_API_TOKEN: ' + (token ? token.substring(0, 8) + '... (set)' : '❌ NOT SET'));
}

function testConnection() {
  const res = jiraRequest_('GET', '/rest/api/3/myself');
  if (res && res.accountId) {
    Logger.log('✅ Connected as: ' + res.displayName + ' (' + res.accountId + ')');
  } else {
    Logger.log('❌ Connection failed. Check checkSetup().');
  }
}

function getAuthHeader_() {
  const props = PropertiesService.getScriptProperties();
  const credentials = Utilities.base64Encode(props.getProperty('JIRA_EMAIL') + ':' + props.getProperty('JIRA_API_TOKEN'));
  return 'Basic ' + credentials;
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
  const url = getSiteUrl_() + path;
  const options = {
    method: method,
    headers: { 'Authorization': getAuthHeader_(), 'Accept': 'application/json' },
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  Logger.log('→ ' + method + ' ' + path);
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code >= 200 && code < 300) {
    return body ? JSON.parse(body) : {};
  } else if (code === 404) {
    Logger.log('  (404 — not found)');
    return null;
  } else {
    Logger.log('  ❌ HTTP ' + code + ': ' + body);
    return null;
  }
}

// ==================== ORCHESTRATOR ====================

function runFullJiraSetup() {
  Logger.log('========== SANTHOSHOS JIRA SETUP — FULL RUN ==========');
  deleteRedundantCustomFields();
  const statuses = createStatuses();
  const workflowName = createWorkflow(statuses);
  const workflowSchemeId = createWorkflowScheme(workflowName);
  const fieldIds = createCustomFields();
  const screenInfo = createScreenAndTabs(fieldIds);
  const issueTypeScreenSchemeId = createIssueTypeScreenScheme(screenInfo.screenId);
  const projects = createProjectsAndComponents();
  wireSchemesToProjects(projects, workflowSchemeId, issueTypeScreenSchemeId);
  Logger.log('========== DONE ==========');
  Logger.log('Manual follow-up needed:');
  Logger.log('  1. Enable "Goals" and "Project" (Atlassian Home) fields per project via UI.');
  Logger.log('  2. Optionally install a status-color marketplace app for 5 distinct colors.');
}

// ==================== PHASE 0: DELETE REDUNDANT CUSTOM FIELDS ====================
// Removes any custom fields this pipeline previously created that turned
// out to duplicate a native/locked Jira field (Start Date/Time, End
// Date/Time — superseded by native "Start date" and "duedate"). Safe to
// re-run: does nothing if they're already gone.

function deleteRedundantCustomFields() {
  Logger.log('--- Phase 0: Delete Redundant Custom Fields ---');
  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  REDUNDANT_FIELD_NAMES.forEach(function (name) {
    const f = allFields.find(function (x) { return x.name === name; });
    if (!f) {
      Logger.log('⏭️  "' + name + '" not present — nothing to delete.');
      return;
    }
    const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/field/' + f.id, {
      method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
    });
    Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted redundant field: ' : '❌ Failed to delete: ') + name);
  });
}

// ==================== PHASE 1: STATUSES ====================

function createStatuses() {
  Logger.log('--- Phase 1: Statuses ---');
  const existing = jiraRequest_('GET', '/rest/api/3/statuses/search?maxResults=200');
  const existingNames = existing && existing.values ? existing.values.map(function (s) { return s.name; }) : [];

  const toCreate = STATUS_CONFIG.filter(function (s) { return existingNames.indexOf(s.name) === -1; });
  const results = [];

  if (toCreate.length > 0) {
    const payload = {
      scope: { type: 'GLOBAL' },
      statuses: toCreate.map(function (s) {
        return { name: s.name, statusCategory: s.category, description: '' };
      }),
    };
    const created = jiraRequest_('POST', '/rest/api/3/statuses', payload);
    if (created) {
      created.forEach(function (s) { Logger.log('✅ Created status: ' + s.name); });
      results.push.apply(results, created);
    } else {
      Logger.log('❌ Status creation failed — see error above. You may need to create these 5 statuses manually (Settings → Issues → Statuses) and re-run createWorkflow() only.');
    }
  }

  STATUS_CONFIG.forEach(function (s) {
    if (existingNames.indexOf(s.name) !== -1) Logger.log('⏭️  Status "' + s.name + '" already exists — skipping.');
  });

  return STATUS_CONFIG; // downstream steps reference by name; workflow API resolves names to IDs
}

// ==================== PHASE 2: WORKFLOW ====================
// NOTE: this endpoint's JSON shape is the least stable part of the Jira
// API. If this call fails, check the logged error body first — Jira's
// message usually tells you exactly which field name it expected.

function createWorkflow(statuses) {
  Logger.log('--- Phase 2: Workflow ---');
  const existing = jiraRequest_('GET', '/rest/api/3/workflows/search?workflowName=' + encodeURIComponent(WORKFLOW_NAME));
  Logger.log('  (debug) workflow search raw: ' + JSON.stringify(existing));
  const exactMatch = existing && existing.values ? existing.values.find(function (w) {
    return w.id && w.id.name === WORKFLOW_NAME;
  }) : null;
  if (exactMatch) {
    Logger.log('⏭️  Workflow "' + WORKFLOW_NAME + '" already exists — skipping.');
    return WORKFLOW_NAME;
  }

  const statusRefs = STATUS_CONFIG.map(function (s, i) { return 's' + i; });
  const payload = {
    scope: { type: 'GLOBAL' },
    workflows: [{
      name: WORKFLOW_NAME,
      description: 'Shared workflow across all santhoshOS Spaces',
      statuses: STATUS_CONFIG.map(function (s, i) {
        return { statusReference: statusRefs[i], name: s.name, statusCategory: s.category };
      }),
      // All-to-all transitions for maximum flexibility (no rigid gating)
      transitions: (function () {
        const t = [];
        let transitionId = 1;
        STATUS_CONFIG.forEach(function (fromStatus, i) {
          STATUS_CONFIG.forEach(function (toStatus, j) {
            if (i === j) return;
            t.push({
              id: String(transitionId++),
              name: 'To ' + toStatus.name,
              to: statusRefs[j],
              from: [statusRefs[i]],
              type: 'DIRECTED',
            });
          });
        });
        return t;
      })(),
    }],
  };

  const result = jiraRequest_('POST', '/rest/api/3/workflows/create', payload);
  if (result) {
    Logger.log('✅ Created workflow: ' + WORKFLOW_NAME);
    return WORKFLOW_NAME;
  } else {
    Logger.log('❌ Workflow creation failed. Fallback: create "' + WORKFLOW_NAME + '" manually in the UI (Settings → Issues → Workflows) with the 5 statuses above, all-to-all transitions, then re-run createWorkflowScheme() and later phases.');
    return WORKFLOW_NAME; // proceed optimistically; downstream calls will fail loudly if it truly doesn't exist
  }
}

function createWorkflowScheme(workflowName) {
  const existing = jiraRequest_('GET', '/rest/api/3/workflowscheme?maxResults=100');
  const found = existing && existing.values ? existing.values.find(function (w) { return w.name === WORKFLOW_SCHEME_NAME; }) : null;
  if (found) {
    Logger.log('⏭️  Workflow scheme already exists — skipping.');
    return found.id;
  }
  const payload = { name: WORKFLOW_SCHEME_NAME, description: 'santhoshOS shared workflow scheme', defaultWorkflow: workflowName };
  const result = jiraRequest_('POST', '/rest/api/3/workflowscheme', payload);
  if (result && result.id) {
    Logger.log('✅ Created workflow scheme: ' + WORKFLOW_SCHEME_NAME);
    return result.id;
  }
  Logger.log('❌ Workflow scheme creation failed.');
  return null;
}

// ==================== PHASE 3: CUSTOM FIELDS ====================

function createCustomFields() {
  Logger.log('--- Phase 3: Custom Fields ---');
  const existingFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  const fieldIds = {}; // name -> id, plus tab grouping

  CUSTOM_FIELD_CONFIG.forEach(function (cfg) {
    const existing = existingFields.find(function (f) { return f.name === cfg.name; });
    if (existing) {
      Logger.log('⏭️  Field "' + cfg.name + '" already exists — skipping.');
      fieldIds[cfg.name] = { id: existing.id, tab: cfg.tab };
      return;
    }
    const payload = { name: cfg.name, description: '', type: cfg.type, searcherKey: cfg.searcher };
    const result = jiraRequest_('POST', '/rest/api/3/field', payload);
    if (result && result.id) {
      Logger.log('✅ Created field: ' + cfg.name + ' (' + result.id + ')');
      fieldIds[cfg.name] = { id: result.id, tab: cfg.tab };
    } else {
      Logger.log('❌ Failed to create field: ' + cfg.name);
    }
  });

  return fieldIds;
}

// ==================== PHASE 4: SCREEN + TABS ====================

function createScreenAndTabs(customFieldIds) {
  Logger.log('--- Phase 4: Screen & Tabs ---');
  const existingScreens = jiraRequest_('GET', '/rest/api/3/screens?maxResults=100');
  let screen = existingScreens && existingScreens.values ? existingScreens.values.find(function (s) { return s.name === SCREEN_NAME; }) : null;

  if (!screen) {
    screen = jiraRequest_('POST', '/rest/api/3/screens', { name: SCREEN_NAME, description: 'santhoshOS work item fields' });
    if (screen) Logger.log('✅ Created screen: ' + SCREEN_NAME);
  } else {
    Logger.log('⏭️  Screen already exists — skipping creation.');
  }
  if (!screen || !screen.id) {
    Logger.log('❌ No screen available — aborting screen phase.');
    return {};
  }

  const tabNames = { default: 'Default Properties', space: 'Space Properties', relational: 'Relational Properties' };
  const existingTabs = jiraRequest_('GET', '/rest/api/3/screens/' + screen.id + '/tabs') || [];
  const tabIds = {};

  Object.keys(tabNames).forEach(function (key) {
    let tab = existingTabs.find(function (t) { return t.name === tabNames[key]; });
    if (!tab) {
      tab = jiraRequest_('POST', '/rest/api/3/screens/' + screen.id + '/tabs', { name: tabNames[key] });
      if (tab) Logger.log('✅ Created tab: ' + tabNames[key]);
    } else {
      Logger.log('⏭️  Tab "' + tabNames[key] + '" already exists — skipping.');
    }
    if (tab) tabIds[key] = tab.id;
  });

  // Add system fields to their tabs
  SYSTEM_FIELD_TABS.forEach(function (sf) {
    addFieldToTab_(screen.id, tabIds[sf.tab], sf.fieldId);
  });

  // Add custom fields to their tabs
  Object.keys(customFieldIds).forEach(function (name) {
    const f = customFieldIds[name];
    addFieldToTab_(screen.id, tabIds[f.tab], f.id);
  });

  // "Start date" — locked, auto-provisioned field, not a fixed system key; resolve by name
  wireNativeStartDateField_(screen.id, tabIds);

  return { screenId: screen.id, tabIds: tabIds };
}

function wireNativeStartDateField_(screenId, tabIds) {
  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  const startDateField = allFields.find(function (f) { return f.name === 'Start date'; });
  if (!startDateField) {
    Logger.log('⚠️ Native "Start date" field not found by name — check spelling/casing on your site, or add it to the tab manually.');
    return;
  }
  addFieldToTab_(screenId, tabIds.default, startDateField.id);
}

function addFieldToTab_(screenId, tabId, fieldId) {
  if (!screenId || !tabId || !fieldId) return;
  const existingFields = jiraRequest_('GET', '/rest/api/3/screens/' + screenId + '/tabs/' + tabId + '/fields') || [];
  if (existingFields.find(function (f) { return f.id === fieldId; })) {
    Logger.log('⏭️  Field ' + fieldId + ' already on tab — skipping.');
    return;
  }
  const result = jiraRequest_('POST', '/rest/api/3/screens/' + screenId + '/tabs/' + tabId + '/fields', { fieldId: fieldId });
  if (result) Logger.log('✅ Added ' + fieldId + ' to tab ' + tabId);
}

// ==================== PHASE 5: SCREEN SCHEME + ISSUE TYPE SCREEN SCHEME ====================

function createIssueTypeScreenScheme(screenId) {
  Logger.log('--- Phase 5: Screen Scheme & Issue Type Screen Scheme ---');
  if (!screenId) { Logger.log('❌ No screenId — skipping.'); return null; }

  let screenScheme = findByName_('/rest/api/3/screenscheme?maxResults=100', 'values', SCREEN_SCHEME_NAME);
  if (!screenScheme) {
    screenScheme = jiraRequest_('POST', '/rest/api/3/screenscheme', {
      name: SCREEN_SCHEME_NAME,
      screens: { default: screenId },
    });
    if (screenScheme) Logger.log('✅ Created screen scheme: ' + SCREEN_SCHEME_NAME);
  } else {
    Logger.log('⏭️  Screen scheme already exists — skipping.');
  }
  if (!screenScheme || !screenScheme.id) { Logger.log('❌ Screen scheme unavailable.'); return null; }

  let itss = findByName_('/rest/api/3/issuetypescreenscheme?maxResults=100', 'values', ISSUE_TYPE_SCREEN_SCHEME_NAME);
  if (!itss) {
    itss = jiraRequest_('POST', '/rest/api/3/issuetypescreenscheme', {
      name: ISSUE_TYPE_SCREEN_SCHEME_NAME,
      issueTypeMappings: [{ issueTypeId: 'default', screenSchemeId: screenScheme.id }],
    });
    if (itss) Logger.log('✅ Created issue type screen scheme: ' + ISSUE_TYPE_SCREEN_SCHEME_NAME);
  } else {
    Logger.log('⏭️  Issue type screen scheme already exists — skipping.');
  }
  return itss ? itss.id : null;
}

function findByName_(path, listKey, name) {
  const res = jiraRequest_('GET', path);
  const list = res && res[listKey] ? res[listKey] : (Array.isArray(res) ? res : []);
  return list.find(function (item) { return item.name === name; }) || null;
}

// ==================== PHASE 6: PROJECTS + COMPONENTS ====================

function createProjectsAndComponents() {
  Logger.log('--- Phase 6: Projects & Components ---');
  const accountId = getMyAccountId_();
  const created = [];

  SPACE_CONFIG.forEach(function (space) {
    let project = jiraRequest_('GET', '/rest/api/3/project/' + space.key);
    if (!project) {
      project = jiraRequest_('POST', '/rest/api/3/project', {
        key: space.key,
        name: space.name,
        projectTypeKey: PROJECT_TYPE_KEY,
        projectTemplateKey: PROJECT_TEMPLATE_KEY,
        leadAccountId: accountId,
      });
      if (project) Logger.log('✅ Created project: ' + space.key + ' — ' + space.name);
    } else {
      Logger.log('⏭️  Project ' + space.key + ' already exists — skipping creation.');
    }
    if (!project) { Logger.log('❌ Failed to create/find ' + space.key); return; }

    const existingComponents = jiraRequest_('GET', '/rest/api/3/project/' + space.key + '/components') || [];
    space.components.forEach(function (compName) {
      if (existingComponents.find(function (c) { return c.name === compName; })) {
        Logger.log('⏭️  Component "' + compName + '" already exists in ' + space.key + ' — skipping.');
        return;
      }
      const comp = jiraRequest_('POST', '/rest/api/3/component', { name: compName, project: space.key });
      if (comp) Logger.log('✅ Created component "' + compName + '" in ' + space.key);
    });

    created.push({ key: space.key, id: project.id });
  });

  return created;
}

// ==================== PHASE 7: WIRE SCHEMES TO PROJECTS ====================

function wireSchemesToProjects(projects, workflowSchemeId, issueTypeScreenSchemeId) {
  Logger.log('--- Phase 7: Wire Schemes to Projects ---');
  projects.forEach(function (p) {
    if (workflowSchemeId) {
      const res = jiraRequest_('PUT', '/rest/api/3/workflowscheme/project', { workflowSchemeId: workflowSchemeId, projectId: p.id });
      Logger.log((res !== null ? '✅' : '❌') + ' Workflow scheme → ' + p.key);
    }
    if (issueTypeScreenSchemeId) {
      const res = jiraRequest_('PUT', '/rest/api/3/issuetypescreenscheme/project', { issueTypeScreenSchemeId: issueTypeScreenSchemeId, projectId: p.id });
      Logger.log((res !== null ? '✅' : '❌') + ' Issue type screen scheme → ' + p.key);
    }
  });
}

// ==================== GLOBAL FACTORY RESET ====================
// Per your instruction: no dry run. Deletes ALL projects on the site,
// plus the custom fields, screen, screen scheme, issue type screen
// scheme, workflow scheme, and workflow this script creates.
// Run once, then delete this function as planned.

function factoryResetJira() {
  Logger.log('========== ⚠️ FACTORY RESET — DELETING EVERYTHING ==========');

  // 1. Delete all projects
  const projects = jiraRequest_('GET', '/rest/api/3/project/search?maxResults=200');
  if (projects && projects.values) {
    projects.values.forEach(function (p) {
      const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/project/' + p.key, {
        method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
      });
      Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted project ' : '❌ Failed to delete ') + p.key);
    });
  }

  // 2. Delete our custom fields
  const allFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  CUSTOM_FIELD_CONFIG.forEach(function (cfg) {
    const f = allFields.find(function (x) { return x.name === cfg.name; });
    if (!f) return;
    const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/field/' + f.id, {
      method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
    });
    Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted field ' : '❌ Failed to delete field ') + cfg.name);
  });

  // 3. Delete issue type screen scheme, screen scheme, screen
  deleteByName_('/rest/api/3/issuetypescreenscheme?maxResults=100', 'values', ISSUE_TYPE_SCREEN_SCHEME_NAME, '/rest/api/3/issuetypescreenscheme/');
  deleteByName_('/rest/api/3/screenscheme?maxResults=100', 'values', SCREEN_SCHEME_NAME, '/rest/api/3/screenscheme/');
  const screens = jiraRequest_('GET', '/rest/api/3/screens?maxResults=100');
  if (screens && screens.values) {
    const s = screens.values.find(function (x) { return x.name === SCREEN_NAME; });
    if (s) {
      const res = UrlFetchApp.fetch(getSiteUrl_() + '/rest/api/3/screens/' + s.id, {
        method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
      });
      Logger.log((res.getResponseCode() < 300 ? '🗑️  Deleted screen' : '❌ Failed to delete screen'));
    }
  }

  // 4. Delete workflow scheme (workflow itself often can't be deleted while referenced; delete scheme first)
  deleteByName_('/rest/api/3/workflowscheme?maxResults=100', 'values', WORKFLOW_SCHEME_NAME, '/rest/api/3/workflowscheme/');

  Logger.log('========== FACTORY RESET COMPLETE ==========');
  Logger.log('Note: statuses created via /rest/api/3/statuses/create are global');
  Logger.log('and may persist even after this reset — Jira does not always allow');
  Logger.log('deleting statuses that were ever used. Check Settings → Issues → Statuses');
  Logger.log('manually if you want those fully gone too.');
}

function deleteByName_(searchPath, listKey, name, deletePathPrefix) {
  const res = jiraRequest_('GET', searchPath);
  const list = res && res[listKey] ? res[listKey] : [];
  const item = list.find(function (x) { return x.name === name; });
  if (!item) return;
  const delRes = UrlFetchApp.fetch(getSiteUrl_() + deletePathPrefix + item.id, {
    method: 'delete', headers: { 'Authorization': getAuthHeader_() }, muteHttpExceptions: true,
  });
  Logger.log((delRes.getResponseCode() < 300 ? '🗑️  Deleted ' : '❌ Failed to delete ') + name);
}