/**
 * Notion Schema Extraction — santhoshOS Migration
 * ─────────────────────────────────────────
 * Version 1.0  |  July 2026
 *
 * PURPOSE
 * Enumerates every database and top-level page your Notion integration
 * has access to, and extracts the full property schema for each database,
 * writing it to a Google Sheet tab.
 *
 * IMPORTANT CORRECTION ON SCOPE
 * Notion's public API has no "workspace ID" filter — there's no way to
 * say "give me everything in workspace X." Access is entirely determined
 * by which pages/databases were manually shared with your integration
 * (in Notion: ⋯ menu → Connections → your integration, on each page/DB
 * you want it to see). This script sees whatever's been shared, full stop.
 * NOTION_WORKSPACE_ID in Script Properties is not used by this script for
 * that reason — only NOTION_API_TOKEN is required.
 *
 * CREDENTIALS (Script Properties)
 *   NOTION_API_TOKEN   — your integration's secret token
 *   SPREADSHEET_ID     — target Google Sheet
 * Optional:
 *   NOTION_SCHEMA_TAB  — default 'Notion Schema'
 *
 * RUN
 *   checkSetup() → testConnection() → discoverNotionSchema()
 */

// ==================== CONFIG ====================

function getProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

const NOTION_VERSION = '2022-06-28';

// ==================== NOTION HELPERS ====================

function notionRequest_(method, path, payload) {
  const options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + getProp_('NOTION_API_TOKEN'),
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);
  const response = UrlFetchApp.fetch('https://api.notion.com/v1' + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  Logger.log('❌ Notion ' + method + ' ' + path + ' → HTTP ' + code + ': ' + body);
  return null;
}

// ==================== CHECK SETUP ====================

function checkSetup() {
  ['NOTION_API_TOKEN', 'SPREADSHEET_ID'].forEach(function (key) {
    Logger.log(key + ': ' + (getProp_(key) ? '✅ set' : '❌ NOT SET'));
  });
}

function testConnection() {
  const res = notionRequest_('GET', '/users/me');
  if (res && res.bot) Logger.log('✅ Connected as integration: ' + (res.name || res.bot.owner ? JSON.stringify(res.bot.owner) : res.id));
  else Logger.log('❌ Connection failed — check NOTION_API_TOKEN.');
}

// ==================== SHEET HELPERS ====================

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

// ==================== MAIN ====================

function discoverNotionSchema() {
  Logger.log('=== Discovering Notion schema ===');
  const sheet = getOrCreateTab_(getProp_('NOTION_SCHEMA_TAB', 'Notion Schema'));
  clearAndHeader_(sheet, ['Object Type', 'Database/Page Title', 'ID', 'Property Name', 'Property Type', 'Additional Config']);

  let rowCount = 0;
  let cursor = undefined;
  let hasMore = true;

  while (hasMore) {
    const payload = { page_size: 100 };
    if (cursor) payload.start_cursor = cursor;

    const res = notionRequest_('POST', '/search', payload);
    if (!res || !res.results) {
      Logger.log('❌ Search failed — check token and that pages/databases are shared with the integration.');
      break;
    }

    res.results.forEach(function (item) {
      if (item.object === 'database') {
        rowCount += writeDatabaseSchema_(sheet, item);
      } else if (item.object === 'page') {
        const title = extractPageTitle_(item);
        sheet.appendRow(['Page', title, item.id, '(standalone page — no schema, has body content only)', '', '']);
        rowCount++;
      }
    });

    hasMore = res.has_more;
    cursor = res.next_cursor;
  }

  Logger.log('✅ Discovery complete. ' + rowCount + ' rows written to "' + sheet.getName() + '".');
  Logger.log('If this looks incomplete: check that all relevant pages/databases are shared with your integration in Notion (⋯ → Connections).');
}

function writeDatabaseSchema_(sheet, database) {
  const title = database.title && database.title[0] ? database.title[0].plain_text : '(untitled)';
  const properties = database.properties || {};
  const propNames = Object.keys(properties);

  if (propNames.length === 0) {
    sheet.appendRow(['Database', title, database.id, '(no properties found)', '', '']);
    return 1;
  }

  let count = 0;
  propNames.forEach(function (propName) {
    const prop = properties[propName];
    let config = '';
    if (prop.type === 'select' && prop.select) config = (prop.select.options || []).map(function (o) { return o.name; }).join(', ');
    if (prop.type === 'multi_select' && prop.multi_select) config = (prop.multi_select.options || []).map(function (o) { return o.name; }).join(', ');
    if (prop.type === 'relation' && prop.relation) config = 'related_database_id: ' + prop.relation.database_id;
    if (prop.type === 'formula' && prop.formula) config = 'expression: ' + prop.formula.expression;
    if (prop.type === 'rollup' && prop.rollup) config = 'rollup_property: ' + prop.rollup.rollup_property_name + ', function: ' + prop.rollup.function;

    sheet.appendRow(['Database', title, database.id, propName, prop.type, config]);
    count++;
  });
  return count;
}

function extractPageTitle_(page) {
  const props = page.properties || {};
  const titleProp = Object.keys(props).find(function (k) { return props[k].type === 'title'; });
  if (titleProp && props[titleProp].title && props[titleProp].title[0]) {
    return props[titleProp].title[0].plain_text;
  }
  return '(untitled)';
}