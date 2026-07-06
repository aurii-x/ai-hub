/**
 * Add Custom Field — santhoshOS Jira Add-on
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * PURPOSE
 * Adds one new custom field, "custom Task ID", to the Space
 * Properties tab of the shared santhoshOS Work Item Screen — since
 * all 4 Spaces (CARE, LRN, FAM, STRAV) share that same screen via
 * the santhoshOS Screen Scheme, this one field addition covers all
 * 4 Spaces' work types in a single run. No per-project changes needed.
 *
 * FIELD TYPE NOTE
 * Built as a URL field: stores the full custom link
 * (e.g. https://app.custom.com/t/90141302224/86bamp9tt) and renders
 * as a clickable link. Jira custom fields can't show short display
 * text (like "86bamp9tt") linking to a different underlying URL —
 * that's a Markdown-style feature Jira's field types don't support.
 * If you'd rather store just the bare ID as plain text instead,
 * change FIELD_TYPE/FIELD_SEARCHER below to the textfield pair
 * (see commented alternative).
 *
 * CREDENTIALS
 * Reuses the same Script Properties as the other santhoshOS Jira
 * scripts: JIRA_SITE_URL, JIRA_EMAIL, JIRA_API_TOKEN. Paste this
 * into the SAME Apps Script project as Jira_Full_Setup_v2 so those
 * are already set — no new setup needed.
 *
 * RUN
 *   addClickUpTaskIdField()
 */

// ==================== CONFIG ====================

const CUSTOM_FIELD_NAME = 'Modified Created Date';

// URL field (recommended — clickable, shows full link as text)
const CUSTOM_FIELD_TYPE = 'com.atlassian.jira.plugin.system.customfieldtypes:datepicker';
const CUSTOM_FIELD_SEARCHER = 'com.atlassian.jira.plugin.system.customfieldtypes:daterange';

// Alternative — plain text field storing just the bare ID, not clickable:
// const CUSTOM_FIELD_TYPE = 'com.atlassian.jira.plugin.system.customfieldtypes:textfield';
// const CUSTOM_FIELD_SEARCHER = 'com.atlassian.jira.plugin.system.customfieldtypes:textsearcher';

const TARGET_SCREEN_NAME = 'santhoshOS Work Item Screen';
const TARGET_TAB_NAME = 'Space Properties';

// ==================== CORE HELPERS (self-contained) ====================

function getAuthHeader_() {
  const props = PropertiesService.getScriptProperties();
  const credentials = Utilities.base64Encode(props.getProperty('JIRA_EMAIL') + ':' + props.getProperty('JIRA_API_TOKEN'));
  return 'Basic ' + credentials;
}

function getSiteUrl_() {
  const url = PropertiesService.getScriptProperties().getProperty('JIRA_SITE_URL');
  return url ? url.replace(/\/$/, '') : url;
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

// ==================== MAIN ====================

function addClickUpTaskIdField() {
  Logger.log('=== Adding "' + CUSTOM_FIELD_NAME + '" field ===');

  // 1. Create field (idempotent)
  const existingFields = jiraRequest_('GET', '/rest/api/3/field') || [];
  let field = existingFields.find(function (f) { return f.name === CUSTOM_FIELD_NAME; });

  if (field) {
    Logger.log('⏭️  Field already exists (' + field.id + ') — skipping creation.');
  } else {
    field = jiraRequest_('POST', '/rest/api/3/field', {
      name: CUSTOM_FIELD_NAME,
      description: 'Direct link to the corresponding custom task',
      type: CUSTOM_FIELD_TYPE,
      searcherKey: CUSTOM_FIELD_SEARCHER,
    });
    if (field) {
      Logger.log('✅ Created field: ' + CUSTOM_FIELD_NAME + ' (' + field.id + ')');
    } else {
      Logger.log('❌ Field creation failed — see error above. Stopping.');
      return;
    }
  }

  // 2. Find the shared screen and its Space Properties tab
  const screens = jiraRequest_('GET', '/rest/api/3/screens?maxResults=100');
  const screen = screens && screens.values ? screens.values.find(function (s) { return s.name === TARGET_SCREEN_NAME; }) : null;
  if (!screen) {
    Logger.log('❌ Screen "' + TARGET_SCREEN_NAME + '" not found. Did you run Jira_Full_Setup_v2 first?');
    return;
  }

  const tabs = jiraRequest_('GET', '/rest/api/3/screens/' + screen.id + '/tabs') || [];
  const tab = tabs.find(function (t) { return t.name === TARGET_TAB_NAME; });
  if (!tab) {
    Logger.log('❌ Tab "' + TARGET_TAB_NAME + '" not found on screen ' + screen.id + '.');
    return;
  }

  // 3. Add field to tab (idempotent) — covers all 4 Spaces since they share this screen
  const tabFields = jiraRequest_('GET', '/rest/api/3/screens/' + screen.id + '/tabs/' + tab.id + '/fields') || [];
  if (tabFields.find(function (f) { return f.id === field.id; })) {
    Logger.log('⏭️  Field already on "' + TARGET_TAB_NAME + '" tab — skipping.');
  } else {
    const result = jiraRequest_('POST', '/rest/api/3/screens/' + screen.id + '/tabs/' + tab.id + '/fields', { fieldId: field.id });
    if (result) {
      Logger.log('✅ Added "' + CUSTOM_FIELD_NAME + '" to "' + TARGET_TAB_NAME + '" tab.');
      Logger.log('This now applies to all 4 Spaces (CARE, LRN, FAM, STRAV) — they share this screen.');
    } else {
      Logger.log('❌ Failed to add field to tab.');
    }
  }

  Logger.log('=== Done ===');
}