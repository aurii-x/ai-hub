// ============================================================
//  TodoistClickUp_DateSync.gs  —  v3
//  Todoist API v1  |  ClickUp API v2
// ============================================================

// ─────────────────────────────────────────────
//  CONFIGURATION
// ─────────────────────────────────────────────
const CONFIG = {
  TODOIST_API_TOKEN:         PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN'),
  CLICKUP_API_TOKEN:         PropertiesService.getScriptProperties().getProperty('CLICKUP_API_TOKEN'),

  EXPORT_SPREADSHEET_ID:     '1JW-oU87SvouLeTnUzpxXWkhC1qh4zXkrJOpttvjF4Bo',   // ← Sheet ID for Function A output
  EXPORT_SHEET_NAME:         'Todoist_Export',

  MAPPED_CSV_SPREADSHEET_ID: 'YOUR_MAPPED_CSV_SPREADSHEET_ID', // ← Sheet ID for Gemini mapping
  MAPPED_CSV_SHEET_NAME:     'Mapping',

  DRY_RUN: true,
};

const TODOIST_BASE = 'https://api.todoist.com/api/v1';
const CLICKUP_BASE = 'https://api.clickup.com/api/v2';


// ─────────────────────────────────────────────
//  SETUP — paste tokens here, run ONCE, then clear
// ─────────────────────────────────────────────
function setup() {
  PropertiesService.getScriptProperties().setProperties({
    TODOIST_API_TOKEN: 'PASTE_TODOIST_TOKEN_HERE',   // ← replace, run, then delete
    CLICKUP_API_TOKEN: 'PASTE_CLICKUP_TOKEN_HERE',   // ← replace, run, then delete
  });
  Logger.log('✅ Credentials stored. Clear the values from setup() now and save.');
}


// ─────────────────────────────────────────────
//  TOKEN TEST — run this first to verify auth
// ─────────────────────────────────────────────
function testTokens() {
  const props = PropertiesService.getScriptProperties().getProperties();

  // Show what's actually stored (first 8 chars only)
  const td = props['TODOIST_API_TOKEN'] || '';
  const cu = props['CLICKUP_API_TOKEN'] || '';
  Logger.log(`Todoist token stored: ${td ? td.substring(0,8) + '...' : 'NOT SET'} (length: ${td.length})`);
  Logger.log(`ClickUp token stored: ${cu ? cu.substring(0,8) + '...' : 'NOT SET'} (length: ${cu.length})`);

  if (!td) { Logger.log('❌ Run setup() first to store the Todoist token.'); return; }

  // Test Todoist — hit the simplest possible endpoint
  Logger.log('\n── Testing Todoist API v1 ──');
  const res = UrlFetchApp.fetch(`${TODOIST_BASE}/projects`, {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${td}`,
      'Content-Type':  'application/json',
    },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  Logger.log(`HTTP ${code}`);
  if (code === 200) {
    const parsed = JSON.parse(body);
    const list   = parsed.results || parsed;
    Logger.log(`✅ Todoist auth OK — ${list.length} projects returned`);
    list.slice(0,3).forEach(p => Logger.log(`   • ${p.name} (${p.id})`));
  } else {
    Logger.log(`❌ Todoist auth FAILED: ${body}`);
    Logger.log('   → Re-run setup() with a fresh token from:');
    Logger.log('     https://app.todoist.com/app/settings/integrations/developer');
  }

  // Test ClickUp
  if (cu) {
    Logger.log('\n── Testing ClickUp API v2 ──');
    const cuRes = UrlFetchApp.fetch(`${CLICKUP_BASE}/user`, {
      method: 'get',
      headers: { 'Authorization': cu },
      muteHttpExceptions: true,
    });
    const cuCode = cuRes.getResponseCode();
    if (cuCode === 200) {
      const u = JSON.parse(cuRes.getContentText()).user;
      Logger.log(`✅ ClickUp auth OK — logged in as: ${u.username} (${u.email})`);
    } else {
      Logger.log(`❌ ClickUp auth FAILED (HTTP ${cuCode}): ${cuRes.getContentText().substring(0,200)}`);
    }
  }
}


// ================================================================
//  FUNCTION A — Export Todoist → Google Sheet
// ================================================================
function functionA_ExportTodoist() {
  const token = CONFIG.TODOIST_API_TOKEN;
  if (!token) { Logger.log('❌ No token. Run setup() then testTokens() first.'); return; }

  Logger.log('▶ Todoist export starting...');

  const ss    = SpreadsheetApp.openById(CONFIG.EXPORT_SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(CONFIG.EXPORT_SHEET_NAME)
              || ss.insertSheet(CONFIG.EXPORT_SHEET_NAME);
  sheet.clearContents();

  // ── Projects ───────────────────────────────────────────────
  Logger.log('  Loading projects...');
  const projectMap = {};
  _todoistPages(token, '/projects').forEach(p => { projectMap[p.id] = p.name; });
  Logger.log(`  Projects: ${Object.keys(projectMap).length}`);

  // ── Sections ───────────────────────────────────────────────
  Logger.log('  Loading sections...');
  const sectionMap = {};
  _todoistPages(token, '/sections').forEach(s => { sectionMap[s.id] = s.name; });
  Logger.log(`  Sections: ${Object.keys(sectionMap).length}`);

  // ── Active tasks ───────────────────────────────────────────
  Logger.log('  Loading active tasks...');
  const active = _todoistPages(token, '/tasks');
  Logger.log(`  Active tasks: ${active.length}`);

  // ── Completed tasks ────────────────────────────────────────
  Logger.log('  Loading completed tasks...');
  const completed = _fetchCompleted(token);
  Logger.log(`  Completed tasks: ${completed.length}`);

  const all = [...active, ...completed];
  Logger.log(`  Total: ${all.length}`);

  // Build parent name lookup
  const nameMap = {};
  all.forEach(t => { nameMap[t.id] = t.content; });

  // ── Headers ────────────────────────────────────────────────
  const H = [
    'task_id', 'task_name', 'description',
    'project_id', 'project_name',
    'section_id', 'section_name',
    'parent_task_id', 'parent_task_name', 'is_subtask',
    'labels', 'priority', 'priority_label', 'status',
    'created_at',
    'due_date',        // → use as ClickUp END DATE
    'due_datetime', 'due_is_recurring',
    'deadline_date',
    'start_date',      // populated from deadline.date; else blank
    'duration_amount', 'duration_unit',
    'completed_at',
    'task_url',
  ];
  sheet.appendRow(H);

  // ── Rows ───────────────────────────────────────────────────
  const PRI = { 1:'Normal', 2:'High', 3:'Very High', 4:'Urgent' };
  const rows = all.map(t => {
    const due  = t.due  || {};
    const dl   = t.deadline || {};
    const dur  = t.duration || {};
    return [
      t.id, t.content, t.description || '',
      t.project_id || '', projectMap[t.project_id] || '',
      t.section_id || '', sectionMap[t.section_id] || '',
      t.parent_id  || '', t.parent_id ? (nameMap[t.parent_id] || '') : '', t.parent_id ? 'YES' : 'NO',
      (t.labels || []).join(', '),
      t.priority || '', PRI[t.priority] || '',
      t.is_completed ? 'completed' : 'active',
      t.created_at || t.added_at || '',
      due.date     || '',
      due.datetime || '', due.is_recurring ? 'YES' : 'NO',
      dl.date      || '',
      dl.date      || '',   // start_date = deadline.date if set
      dur.amount   || '', dur.unit || '',
      t.completed_at || '',
      `https://app.todoist.com/app/task/${t.id}`,
    ];
  });

  if (rows.length) sheet.getRange(2, 1, rows.length, H.length).setValues(rows);

  // Format
  const hr = sheet.getRange(1, 1, 1, H.length);
  hr.setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, H.length);

  Logger.log(`✅ Done — ${rows.length} rows exported to "${CONFIG.EXPORT_SHEET_NAME}"`);
  Logger.log(`   ${ss.getUrl()}`);
}


// ── Paginated GET for v1 endpoints ────────────────────────────
function _todoistPages(token, path, extra) {
  const results = [];
  let   cursor  = null;

  for (let page = 0; page < 50; page++) {          // max 50 pages safety cap
    const params = Object.assign({ limit: 200 }, extra || {});
    if (cursor) params.cursor = cursor;

    const qs  = Object.entries(params).map(([k,v]) => `${k}=${encodeURIComponent(v)}`).join('&');
    const url = `${TODOIST_BASE}${path}?${qs}`;

    const res  = UrlFetchApp.fetch(url, {
      method:  'get',
      headers: { 'Authorization': `Bearer ${token}` },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code !== 200) {
      Logger.log(`  ⚠ GET ${path} → HTTP ${code}: ${text.substring(0,300)}`);
      break;
    }

    const data = JSON.parse(text);

    if (Array.isArray(data)) {                      // plain array (no pagination)
      results.push(...data);
      break;
    }

    const items = data.results || data.items || [];
    results.push(...items);

    cursor = data.next_cursor || null;
    if (!cursor) break;
  }

  return results;
}


// ── Completed tasks via v1 ────────────────────────────────────
function _fetchCompleted(token) {
  const results = [];
  let   cursor  = null;

  // Try v1 unified endpoint first; fall back gracefully
  const endpoints = [
    `${TODOIST_BASE}/tasks/completed/get_all`,
    'https://api.todoist.com/sync/v9/items/completed/get_all',  // fallback
  ];

  for (const base of endpoints) {
    cursor = null;
    let ok = true;

    for (let page = 0; page < 50; page++) {
      let url = `${base}?limit=200`;
      if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

      const res  = UrlFetchApp.fetch(url, {
        method:  'get',
        headers: { 'Authorization': `Bearer ${token}` },
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code !== 200) {
        Logger.log(`  Completed endpoint ${base} → HTTP ${code} — ${code === 404 ? 'trying fallback' : text.substring(0,200)}`);
        ok = false;
        break;
      }

      const data  = JSON.parse(text);
      const items = data.results || data.items || [];

      items.forEach(item => results.push({
        id:           item.id,
        content:      item.content,
        description:  item.description   || '',
        project_id:   item.project_id,
        section_id:   item.section_id    || '',
        parent_id:    item.parent_id     || '',
        labels:       item.labels        || [],
        priority:     item.priority,
        is_completed: true,
        created_at:   item.added_at      || item.created_at || '',
        added_at:     item.added_at      || '',
        due:          item.due           || {},
        deadline:     item.deadline      || {},
        duration:     item.duration      || {},
        completed_at: item.completed_at  || '',
      }));

      Logger.log(`  Completed batch (${base.includes('sync') ? 'sync' : 'v1'}): ${items.length} items`);
      cursor = data.next_cursor || null;
      if (!cursor || items.length === 0) break;
    }

    if (ok) break;   // succeeded with this endpoint, no need for fallback
  }

  return results;
}


// ================================================================
//  FUNCTION B — Sync Gemini-mapped dates to ClickUp
//
//  Expected CSV columns (case-insensitive, spaces→underscores):
//    todoist_task_id | todoist_task_name
//    clickup_task_id | clickup_task_name
//    match_confidence
//    todoist_created_at | todoist_start_date | todoist_due_date
//    recommended_clickup_start | recommended_clickup_end
//    approved                  ← put YES in rows you want synced
// ================================================================
function functionB_SyncDatesToClickUp() {
  const dryRun = CONFIG.DRY_RUN;
  const token  = CONFIG.CLICKUP_API_TOKEN;
  if (!token) { Logger.log('❌ No ClickUp token. Run setup() first.'); return; }

  Logger.log(`▶ ClickUp date sync — ${dryRun ? 'DRY RUN' : 'LIVE'}`);

  const sheet = SpreadsheetApp
    .openById(CONFIG.MAPPED_CSV_SPREADSHEET_ID)
    .getSheetByName(CONFIG.MAPPED_CSV_SHEET_NAME);
  if (!sheet) { Logger.log(`❌ Sheet "${CONFIG.MAPPED_CSV_SHEET_NAME}" not found.`); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g,'_'));
  const rows    = data.slice(1);

  const col = n => headers.indexOf(n);
  const C = {
    cu_id:      col('clickup_task_id'),
    cu_name:    col('clickup_task_name'),
    confidence: col('match_confidence'),
    created:    col('todoist_created_at'),
    start:      col('todoist_start_date'),
    due:        col('todoist_due_date'),
    rec_start:  col('recommended_clickup_start'),
    rec_end:    col('recommended_clickup_end'),
    approved:   col('approved'),
  };

  const missing = Object.entries(C).filter(([k,v])=>v===-1).map(([k])=>k);
  if (missing.length) {
    Logger.log(`❌ Missing columns: ${missing.join(', ')}`);
    Logger.log(`   Found: ${headers.join(', ')}`);
    return;
  }

  let processed=0, updated=0, skipped=0, noChange=0, errors=0;

  rows.forEach((row, i) => {
    const rowNum   = i + 2;
    const approved = String(row[C.approved]).trim().toUpperCase();
    const cuId     = String(row[C.cu_id]).trim();

    if (approved !== 'YES') { skipped++; return; }
    if (!cuId) { Logger.log(`  Row ${rowNum}: no ClickUp ID — skip`); skipped++; return; }

    const created  = _pd(row[C.created]);
    const tdStart  = _pd(row[C.start]);
    const tdDue    = _pd(row[C.due]);
    const recStart = _pd(row[C.rec_start]);
    const recEnd   = _pd(row[C.rec_end]);

    const effStart = tdStart  || created;   // start_date if set, else created_at
    const effEnd   = tdDue    || recEnd;

    if (!effStart && !effEnd) { Logger.log(`  Row ${rowNum}: no usable dates`); skipped++; return; }

    const cu = _cuGet(cuId, token);
    if (!cu) { errors++; return; }

    const cuStartMs = cu.start_date ? parseInt(cu.start_date) : null;
    const cuDueMs   = cu.due_date   ? parseInt(cu.due_date)   : null;
    const cuStart   = cuStartMs ? new Date(cuStartMs) : null;
    const cuDue     = cuDueMs   ? new Date(cuDueMs)   : null;

    const upd = {};
    const log = [];

    // ── START DATE ──────────────────────────────────────────
    // Set if empty; replace if ClickUp is newer (keep the earlier date)
    if (effStart) {
      if (!cuStart) {
        upd.start_date = effStart.getTime();
        log.push(`SET start=${_f(effStart)} (was empty)`);
      } else if (cuStart > effStart) {
        upd.start_date = effStart.getTime();
        log.push(`REPLACE start ${_f(cuStart)}→${_f(effStart)} (CU newer)`);
      } else {
        log.push(`KEEP start ${_f(cuStart)}`);
      }
    }

    // ── END DATE ────────────────────────────────────────────
    // Replace if ClickUp is newer AND end > start
    if (effEnd) {
      const startRef = upd.start_date ? new Date(upd.start_date) : cuStart || effStart;
      if (!cuDue) {
        if (!startRef || effEnd > startRef) {
          upd.due_date = effEnd.getTime();
          log.push(`SET due=${_f(effEnd)} (was empty)`);
        } else {
          log.push(`SKIP due=${_f(effEnd)} not after start`);
        }
      } else if (cuDue > effEnd) {
        if (!startRef || effEnd > startRef) {
          upd.due_date = effEnd.getTime();
          log.push(`REPLACE due ${_f(cuDue)}→${_f(effEnd)} (CU newer)`);
        } else {
          log.push(`SKIP due=${_f(effEnd)} not after start`);
        }
      } else {
        log.push(`KEEP due ${_f(cuDue)}`);
      }
    }

    Logger.log(`  Row ${rowNum} [${cuId}] ${cu.name}: ${log.join(' | ')}`);

    if (!Object.keys(upd).length) { noChange++; processed++; return; }

    if (dryRun) {
      Logger.log(`    [DRY RUN] would update: ${JSON.stringify(upd)}`);
      updated++;
    } else {
      if (_cuPut(cuId, upd, token)) { Logger.log('    ✅ Updated'); updated++; }
      else                          { Logger.log('    ❌ Failed');  errors++;  }
    }
    processed++;
  });

  Logger.log('');
  Logger.log('══════════════════════════════');
  Logger.log(`  Mode:      ${dryRun ? 'DRY RUN' : 'LIVE'}`);
  Logger.log(`  Processed: ${processed}`);
  Logger.log(`  Updated:   ${updated}`);
  Logger.log(`  No change: ${noChange}`);
  Logger.log(`  Skipped:   ${skipped}`);
  Logger.log(`  Errors:    ${errors}`);
  Logger.log('══════════════════════════════');
  if (dryRun) Logger.log('Set CONFIG.DRY_RUN = false and re-run to apply.');
}

// ── ClickUp helpers ───────────────────────────────────────────
function _cuGet(id, token) {
  const res  = UrlFetchApp.fetch(`${CLICKUP_BASE}/task/${id}`, {
    method: 'get', headers: { Authorization: token }, muteHttpExceptions: true,
  });
  const body = res.getContentText();
  if (res.getResponseCode() !== 200 || !body.trimStart().startsWith('{')) {
    Logger.log(`  ClickUp GET ${id} failed (${res.getResponseCode()}): ${body.substring(0,150)}`);
    return null;
  }
  return JSON.parse(body);
}
function _cuPut(id, payload, token) {
  const res = UrlFetchApp.fetch(`${CLICKUP_BASE}/task/${id}`, {
    method: 'put',
    headers: { Authorization: token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    Logger.log(`  ClickUp PUT ${id} failed (${res.getResponseCode()}): ${res.getContentText().substring(0,200)}`);
    return false;
  }
  return true;
}

// ── Date helpers ──────────────────────────────────────────────
function _pd(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(String(v).trim());
  return isNaN(d.getTime()) ? null : d;
}
function _f(d) {
  try { return d ? d.toISOString().split('T')[0] : 'null'; } catch(e) { return String(d); }
}


// ─────────────────────────────────────────────
//  SHEET MENU
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Todoist ↔ ClickUp')
    .addItem('0. Test tokens',                         'testTokens')
    .addItem('1. Export Todoist → Sheet',              'functionA_ExportTodoist')
    .addSeparator()
    .addItem('3a. DRY RUN — preview date sync',        'runDryRun')
    .addItem('3b. LIVE    — apply date sync',          'runLive')
    .addToUi();
}
function runDryRun() { CONFIG.DRY_RUN = true;  functionB_SyncDatesToClickUp(); }
function runLive() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ Live Update', 'Write to ClickUp now?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    CONFIG.DRY_RUN = false;
    functionB_SyncDatesToClickUp();
    CONFIG.DRY_RUN = true;
  }
}