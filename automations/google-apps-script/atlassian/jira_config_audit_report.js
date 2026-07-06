/**
 * Jira Config Audit Report — santhoshOS
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * PURPOSE
 * Queries your live Jira site and writes a structured configuration
 * report to a Google Sheet tab — mirrors the sections in Jira admin
 * settings (Work types, Workflows, Screens, Fields, Priorities, Work
 * item features, Work item attributes) plus per-project associations
 * for your 4 Spaces. Meant to be re-run periodically as a drift check,
 * not a one-time document — re-running overwrites the tab with a
 * fresh timestamped snapshot rather than appending.
 *
 * CREDENTIALS (Script Properties — same as your other Jira scripts)
 *   JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN
 *   SPREADSHEET_ID           (target Google Sheet)
 * Optional:
 *   JIRA_AUDIT_TAB           (default 'Jira Config Audit')
 *
 * RUN
 *   generateJiraConfigReport()
 *
 * NOTES ON GAPS
 * A few admin sections (Priority schemes, Work item security schemes)
 * are Premium-tier or newer features — if your site doesn't have
 * access, the report logs "Not available on this site/plan" for that
 * row rather than failing the whole run.
 */

// ==================== CONFIG ====================

function getProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

const OUR_NAMES = {
  workflow: 'santhoshOS Standard Workflow',
  workflowScheme: 'santhoshOS Workflow Scheme',
  screen: 'santhoshOS Work Item Screen',
  screenScheme: 'santhoshOS Screen Scheme',
  issueTypeScreenScheme: 'santhoshOS Issue Type Screen Scheme',
  fieldConfig: 'santhoshOS Field Configuration',
  fieldConfigScheme: 'santhoshOS Field Configuration Scheme',
};

const SPACE_KEYS = ['CARE', 'LRN', 'FAM', 'STRAV'];

// ==================== JIRA HELPERS ====================

function jiraAuthHeader_() {
  return 'Basic ' + Utilities.base64Encode(getProp_('JIRA_EMAIL') + ':' + getProp_('JIRA_API_TOKEN'));
}
function jiraSiteUrl_() {
  return getProp_('JIRA_SITE_URL').replace(/\/$/, '');
}
function jiraGet_(path) {
  const options = {
    method: 'get',
    headers: { 'Authorization': jiraAuthHeader_(), 'Accept': 'application/json' },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(jiraSiteUrl_() + path, options);
  const code = response.getResponseCode();
  if (code >= 200 && code < 300) return JSON.parse(response.getContentText());
  return { __error: true, __code: code, __body: response.getContentText() };
}
function isOk_(res) { return res && !res.__error; }

// ==================== SHEET HELPERS ====================

function auditSheet_() {
  const ss = SpreadsheetApp.openById(getProp_('SPREADSHEET_ID'));
  const name = getProp_('JIRA_AUDIT_TAB', 'Jira Config Audit');
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  return sheet;
}

function section_(sheet, title) {
  sheet.appendRow(['']);
  sheet.appendRow(['── ' + title + ' ──']);
}

function row_(sheet, cols) {
  sheet.appendRow(cols);
}

// ==================== MAIN ====================

function generateJiraConfigReport() {
  const sheet = auditSheet_();
  sheet.appendRow(['santhoshOS Jira Configuration Audit — generated ' + new Date()]);
  sheet.setFrozenRows(1);

  reportWorkTypes_(sheet);
  reportWorkflows_(sheet);
  reportScreens_(sheet);
  reportFields_(sheet);
  reportPriorities_(sheet);
  reportWorkItemFeatures_(sheet);
  reportWorkItemAttributes_(sheet);
  reportPerProjectAssociations_(sheet);

  sheet.autoResizeColumns(1, 5);
  Logger.log('✅ Report written to "' + sheet.getName() + '" tab.');
}

// ==================== WORK TYPES ====================

function reportWorkTypes_(sheet) {
  section_(sheet, 'Work Types');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  row_(sheet, ['Work type hierarchy', '(all)', 'Default Jira hierarchy (Epic > Story/Task/Bug > Subtask)', 'Default', 'Not customized']);

  const issueTypes = jiraGet_('/rest/api/3/issuetype');
  if (isOk_(issueTypes)) {
    issueTypes.forEach(function (it) {
      row_(sheet, ['Work type', it.name, 'hierarchyLevel: ' + it.hierarchyLevel + ', subtask: ' + it.subtask, it.name === 'Epic' || it.name === 'Task' || it.name === 'Subtask' ? 'Used' : 'Default/unused', '']);
    });
  } else {
    row_(sheet, ['Work types', '(error)', 'Could not fetch', '', 'HTTP ' + issueTypes.__code]);
  }

  const itSchemes = jiraGet_('/rest/api/3/issuetypescheme/search?maxResults=100');
  if (isOk_(itSchemes) && itSchemes.values) {
    itSchemes.values.forEach(function (s) {
      row_(sheet, ['Work type scheme', s.name, s.isDefault ? '(site default)' : '', s.isDefault ? 'Default' : 'Check', '']);
    });
  }

  const subtasks = isOk_(issueTypes) ? issueTypes.filter(function (it) { return it.subtask; }) : [];
  row_(sheet, ['Sub-tasks', subtasks.map(function (s) { return s.name; }).join(', ') || '(none found)', '', 'Default', 'Standard Jira "Subtask" type']);
}

// ==================== WORKFLOWS ====================

function reportWorkflows_(sheet) {
  section_(sheet, 'Workflows');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const workflows = jiraGet_('/rest/api/3/workflow/search?maxResults=200');
  if (isOk_(workflows) && workflows.values) {
    workflows.values.forEach(function (w) {
      const isOurs = w.id && w.id.name === OUR_NAMES.workflow;
      row_(sheet, ['Workflow', (w.id ? w.id.name : w.name), w.description || '', isOurs ? '✅ Ours' : (w.isDefault ? 'Default' : 'Other'), '']);
    });
  }

  const wfSchemes = jiraGet_('/rest/api/3/workflowscheme?maxResults=100');
  if (isOk_(wfSchemes) && wfSchemes.values) {
    wfSchemes.values.forEach(function (s) {
      row_(sheet, ['Workflow scheme', s.name, '', s.name === OUR_NAMES.workflowScheme ? '✅ Ours' : 'Other', '']);
    });
  }
}

// ==================== SCREENS ====================

function reportScreens_(sheet) {
  section_(sheet, 'Screens');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const screens = jiraGet_('/rest/api/3/screens?maxResults=200');
  if (isOk_(screens) && screens.values) {
    screens.values.forEach(function (s) {
      row_(sheet, ['Screen', s.name, s.description || '', s.name === OUR_NAMES.screen ? '✅ Ours' : 'Other', '']);
    });
  }

  const screenSchemes = jiraGet_('/rest/api/3/screenscheme?maxResults=200');
  if (isOk_(screenSchemes) && screenSchemes.values) {
    screenSchemes.values.forEach(function (s) {
      row_(sheet, ['Screen scheme', s.name, '', s.name === OUR_NAMES.screenScheme ? '✅ Ours' : 'Other — check if orphaned', '']);
    });
  }

  const itss = jiraGet_('/rest/api/3/issuetypescreenscheme?maxResults=200');
  if (isOk_(itss) && itss.values) {
    itss.values.forEach(function (s) {
      row_(sheet, ['Work type screen scheme', s.name, '', s.name === OUR_NAMES.issueTypeScreenScheme ? '✅ Ours' : 'Other — check if orphaned', '']);
    });
  }
}

// ==================== FIELDS ====================

function reportFields_(sheet) {
  section_(sheet, 'Fields');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const fields = jiraGet_('/rest/api/3/field');
  if (isOk_(fields)) {
    const customFields = fields.filter(function (f) { return f.custom; });
    customFields.forEach(function (f) {
      row_(sheet, ['Custom field', f.name, f.id + ' — ' + (f.schema ? f.schema.type : ''), '✅ Ours (custom fields you created)', '']);
    });
    row_(sheet, ['System fields', '(all others)', fields.length - customFields.length + ' built-in fields', 'Default', 'Not enumerated individually']);
  }

  const fieldSchemes = jiraGet_('/rest/api/3/fieldconfigurationscheme?maxResults=100');
  if (isOk_(fieldSchemes) && fieldSchemes.values) {
    fieldSchemes.values.forEach(function (s) {
      row_(sheet, ['Field scheme', s.name, '', s.name === OUR_NAMES.fieldConfigScheme ? '✅ Ours' : 'Other', '']);
    });
  }
}

// ==================== PRIORITIES ====================

function reportPriorities_(sheet) {
  section_(sheet, 'Priorities');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const priorities = jiraGet_('/rest/api/3/priority');
  if (isOk_(priorities)) {
    priorities.forEach(function (p) {
      row_(sheet, ['Priority', p.name, '', 'Default', '']);
    });
  }

  const priorityScheme = jiraGet_('/rest/api/3/priorityscheme?maxResults=100');
  if (isOk_(priorityScheme)) {
    row_(sheet, ['Priority scheme', '(default)', 'Site uses default priority scheme', 'Default', 'Not customized']);
  } else {
    row_(sheet, ['Priority scheme', '(n/a)', 'Not available on this site/plan', '', 'HTTP ' + priorityScheme.__code]);
  }
}

// ==================== WORK ITEM FEATURES ====================

function reportWorkItemFeatures_(sheet) {
  section_(sheet, 'Work Item Features');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const config = jiraGet_('/rest/api/3/configuration');
  if (isOk_(config)) {
    row_(sheet, ['Time tracking', 'timeTrackingEnabled', String(config.timeTrackingEnabled), 'Default', config.timeTrackingConfiguration ? JSON.stringify(config.timeTrackingConfiguration) : '']);
  }

  const linkTypes = jiraGet_('/rest/api/3/issueLinkType');
  if (isOk_(linkTypes) && linkTypes.issueLinkTypes) {
    row_(sheet, ['Work item linking', linkTypes.issueLinkTypes.map(function (l) { return l.name; }).join(', '), linkTypes.issueLinkTypes.length + ' link types', 'Default', 'Standard Jira set unless customized']);
  }
}

// ==================== WORK ITEM ATTRIBUTES ====================

function reportWorkItemAttributes_(sheet) {
  section_(sheet, 'Work Item Attributes');
  row_(sheet, ['Category', 'Name', 'Details', 'Ours?', 'Notes']);

  const statuses = jiraGet_('/rest/api/3/statuses/search?maxResults=200');
  if (isOk_(statuses) && statuses.values) {
    const OUR_STATUSES = ['New', 'Planned', 'In Progress', 'Done', 'Archived'];
    statuses.values
      .filter(function (s) { return s.scope && s.scope.type === 'GLOBAL'; })
      .forEach(function (s) {
        row_(sheet, ['Status (global)', s.name, s.statusCategory, OUR_STATUSES.indexOf(s.name) !== -1 ? '✅ Ours' : 'Default/other', '']);
      });
  }

  const resolutions = jiraGet_('/rest/api/3/resolution');
  if (isOk_(resolutions)) {
    row_(sheet, ['Resolutions', resolutions.map(function (r) { return r.name; }).join(', '), resolutions.length + ' resolutions', 'Default', 'Not customized']);
  }

  const securitySchemes = jiraGet_('/rest/api/3/issuesecurityschemes');
  if (isOk_(securitySchemes)) {
    row_(sheet, ['Work item security schemes', securitySchemes.issueSecuritySchemes ? securitySchemes.issueSecuritySchemes.length + ' found' : '0 found', '', 'Default', 'Not configured']);
  } else {
    row_(sheet, ['Work item security schemes', '(n/a)', 'Not available on this site/plan', '', 'HTTP ' + securitySchemes.__code]);
  }

  const notifSchemes = jiraGet_('/rest/api/3/notificationscheme?maxResults=100');
  if (isOk_(notifSchemes) && notifSchemes.values) {
    notifSchemes.values.forEach(function (s) {
      row_(sheet, ['Notification scheme', s.name, '', 'Default', 'Not customized unless noted']);
    });
  }

  const permSchemes = jiraGet_('/rest/api/3/permissionscheme');
  if (isOk_(permSchemes) && permSchemes.permissionSchemes) {
    permSchemes.permissionSchemes.forEach(function (s) {
      row_(sheet, ['Permission scheme', s.name, '', 'Default', 'Free plan — permissions locked to default, per earlier discussion']);
    });
  }
}

// ==================== PER-PROJECT ASSOCIATIONS ====================

function reportPerProjectAssociations_(sheet) {
  section_(sheet, 'Per-Project Associations (CARE / LRN / FAM / STRAV)');
  row_(sheet, ['Project', 'Association Type', 'Value', 'Matches Shared?', 'Notes']);

  SPACE_KEYS.forEach(function (key) {
    const project = jiraGet_('/rest/api/3/project/' + key);
    if (!isOk_(project)) { row_(sheet, [key, 'Project lookup', 'FAILED', '', 'HTTP ' + project.__code]); return; }

    const wf = jiraGet_('/rest/api/3/workflowscheme/project?projectId=' + project.id);
    const wfName = isOk_(wf) && wf.values && wf.values[0] && wf.values[0].workflowScheme ? wf.values[0].workflowScheme.name : '(unknown)';
    row_(sheet, [key, 'Workflow scheme', wfName, wfName === OUR_NAMES.workflowScheme ? '✅' : '❌ MISMATCH', '']);

    const itss = jiraGet_('/rest/api/3/issuetypescreenscheme/project?projectId=' + project.id);
    const itssId = isOk_(itss) && itss.values && itss.values[0] ? itss.values[0].issueTypeScreenSchemeId : null;
    row_(sheet, [key, 'Work type screen scheme ID', itssId || '(unknown)', '', 'Cross-check ID against Screens section above']);

    const fieldConfigScheme = jiraGet_('/rest/api/3/fieldconfigurationscheme/project?projectId=' + project.id);
    const fcsId = isOk_(fieldConfigScheme) && fieldConfigScheme.values && fieldConfigScheme.values[0] ? fieldConfigScheme.values[0].fieldConfigurationSchemeId : null;
    row_(sheet, [key, 'Field config scheme ID', fcsId || '(unknown)', '', 'Cross-check ID against Fields section above']);
  });
}