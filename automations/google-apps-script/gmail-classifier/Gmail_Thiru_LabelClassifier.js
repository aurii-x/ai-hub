// ============================================================
// Gmail Auto-Classifier using Gemini
// ============================================================
// SETUP:
// 1. Go to script.google.com — sign in with your email
// 2. Create a new project and paste this entire script
// 3. Upload label_criteria.json to your Google Drive folder and
//    update DRIVE_FOLDER_PATH = ['level1', 'level2', ...] to match
// 4. Update GEMINI_MODEL with the model name your API key supports
// 5. Update MAX_LABELS — max number of labels per email (recommended: 3)
// 6. Update MIN_CONFIDENCE — minimum confidence before applying a label (recommended: 0.75)
// 7. Update BATCH_SIZE — number of emails to process per run (recommended: 5)
// 8. Run setup() once — enter your Gemini API key and Gmail address when prompted
//    Your credentials are saved securely in Script Properties, not in the code
// 9. Run classifyNewEmails once and authorize Gmail access when prompted
// 10. Set a trigger: classifyNewEmails → Time-driven → Every hour
// ============================================================

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GMAIL_ADDRESS  = PropertiesService.getScriptProperties().getProperty('GMAIL_ADDRESS');
const CRITERIA_FILE_NAME = 'label_criteria.json';  // must be in Google Drive root\App
const GEMINI_MODEL       = 'gemini-2.5-flash'; //choose the Gemini model your api key is generated for
const MAX_LABELS         = 3;      // max labels per email
const MIN_CONFIDENCE     = 0.75;   // only apply labels with confidence >= this (0.0 - 1.0)
const LOG_SHEET_NAME     = 'GmailClassifierLog';
const DRIVE_FOLDER_PATH  = ['AppData','GmailAutomation', 'LabelModifier']; // nested folder path
const BATCH_SIZE         = 5;
const MAX_RUNTIME_MS     = 5 * 60 * 1000; // 5 minutes

//setup your api key and email with your key one for the first time from 
function setup() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('GEMINI_API_KEY', 'your-key-here');
  props.setProperty('GMAIL_ADDRESS',  'your@gmail.com');
  Logger.log('Setup complete.');
}

function checkSetup() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  const email = props.getProperty('GMAIL_ADDRESS');
  Logger.log('API Key: ' + (key ? key.slice(0, 8) + '...' : 'NOT SET'));
  Logger.log('Gmail: ' + (email || 'NOT SET'));
}

// ── DRIVE FOLDER HELPER ───────────────────────────────────────────────────────
// Navigates to GmailAutomation/LabelModifier, creates folders if missing
function getAppFolder() {
  let folder = DriveApp.getRootFolder();
  for (const name of DRIVE_FOLDER_PATH) {
    const found = folder.getFoldersByName(name);
    if (found.hasNext()) {
      folder = found.next();
    } else {
      folder = folder.createFolder(name);
      Logger.log(`Created folder: ${name}`);
    }
  }
  return folder;
}
 
// ── LOAD LABEL CRITERIA FROM DRIVE FOLDER ────────────────────────────────────
function loadLabelCriteria() {
  try {
    const folder = getAppFolder();
    const files = folder.getFilesByName(CRITERIA_FILE_NAME);
    if (!files.hasNext()) throw new Error(`"${CRITERIA_FILE_NAME}" not found in GmailAutomation/LabelModifier`);
    const content = files.next().getBlob().getDataAsString();
    const parsed = JSON.parse(content);
    Logger.log(`✓ Loaded ${parsed.labels.length} labels from ${CRITERIA_FILE_NAME}`);
    return parsed.labels;
  } catch(e) {
    Logger.log('✗ Could not load criteria file: ' + e.message);
    logError('loadLabelCriteria', e.message, 'Loading label criteria', 'Failed at file read');
    return null;
  }
}
 
function buildLabelDefs(labels) {
  return labels.map(l => `"${l.name}": ${l.criteria}`).join('\n\n');
}
 
// ── GOOGLE SHEETS LOG ─────────────────────────────────────────────────────────
// Gets or creates the log spreadsheet in GmailAutomation/LabelModifier
// Sheet 1: Skipped — Sheet 2: Error Log
function getOrCreateLogSheet() {
  const folder = getAppFolder();
  const files = folder.getFilesByName(LOG_SHEET_NAME);
 
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
 
  // Create new spreadsheet
  const ss = SpreadsheetApp.create(LOG_SHEET_NAME);
 
  // Move it to the correct folder
  const ssFile = DriveApp.getFileById(ss.getId());
  folder.addFile(ssFile);
  DriveApp.getRootFolder().removeFile(ssFile); // remove from root
 
  // Set up Skipped sheet
  const skippedSheet = ss.getSheets()[0];
  skippedSheet.setName('Skipped');
  skippedSheet.appendRow(['Timestamp', 'Subject', 'From', 'Skip Reason', 'Gemini Labels', 'Confidence Scores']);
  skippedSheet.setFrozenRows(1);
  skippedSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  skippedSheet.setColumnWidth(1, 160);
  skippedSheet.setColumnWidth(2, 280);
  skippedSheet.setColumnWidth(3, 200);
  skippedSheet.setColumnWidth(4, 200);
  skippedSheet.setColumnWidth(5, 180);
  skippedSheet.setColumnWidth(6, 220);
 
  // Set up Error Log sheet
  const errorSheet = ss.insertSheet('Error Log');
  errorSheet.appendRow(['Timestamp', 'Error Message', 'Error Line', 'Triggered By', 'Last Completed Step', 'Failed At']);
  errorSheet.setFrozenRows(1);
  errorSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
  errorSheet.setColumnWidth(1, 160);
  errorSheet.setColumnWidth(2, 300);
  errorSheet.setColumnWidth(3, 100);
  errorSheet.setColumnWidth(4, 160);
  errorSheet.setColumnWidth(5, 220);
  errorSheet.setColumnWidth(6, 220);
 
  Logger.log(`✓ Created log sheet: ${LOG_SHEET_NAME}`);
  return ss;
}
 
function logSkipped(entries) {
  if (!entries || entries.length === 0) return;
  try {
    const ss = getOrCreateLogSheet();
    const sheet = ss.getSheetByName('Skipped');
    const ts = new Date();
    const rows = entries.map(e => [
      ts,
      (e.subject || '').slice(0, 100),
      (e.from || '').slice(0, 80),
      e.reason || '',
      (e.labels || []).map(l => l.name).join(' | '),
      (e.labels || []).map(l => `${l.name}: ${(l.confidence * 100).toFixed(0)}%`).join(' | ')
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    Logger.log(`📝 Logged ${entries.length} skipped emails`);
  } catch(e) {
    Logger.log('Skip log error: ' + e.message);
  }
}
 
function logError(triggeredBy, errorMessage, lastCompleted, failedAt, errorLine) {
  try {
    const ss = getOrCreateLogSheet();
    const sheet = ss.getSheetByName('Error Log');
    sheet.appendRow([
      new Date(),
      errorMessage || '',
      errorLine || '',
      triggeredBy || '',
      lastCompleted || '',
      failedAt || ''
    ]);
  } catch(e) {
    Logger.log('Error log write failed: ' + e.message);
  }
}
 
// ── FILTER: confidence threshold + max label cap ──────────────────────────────
function filterClassifications(classifications) {
  return classifications.map(cls => {
    const all = (cls.labels || []).sort((a, b) => (b.confidence || 1.0) - (a.confidence || 1.0));
    const passed  = all.filter(l => (l.confidence || 1.0) >= MIN_CONFIDENCE).slice(0, MAX_LABELS);
    const skipped = all.filter(l => (l.confidence || 1.0) < MIN_CONFIDENCE);
    return { id: cls.id, labels: passed.map(l => l.name), passedRaw: passed, skippedRaw: skipped };
  });
}
 
// ── APPLY CLASSIFICATIONS HELPER ──────────────────────────────────────────────
function applyClassifications(classifications, threads, emailData, props) {
  let done = 0, failed = 0;
  const skipEntries = [];
 
  for (const cls of classifications) {
    const idx = parseInt(cls.id);
    if (isNaN(idx) || idx < 0 || idx >= threads.length) continue;
    const email = emailData[idx];
    try {
      if (cls.labels.length === 0) {
        const reason = cls.skippedRaw && cls.skippedRaw.length > 0
          ? `All labels below confidence threshold (${MIN_CONFIDENCE})`
          : 'Gemini returned no labels';
        Logger.log(`⚠ Skipped: ${email.subject.slice(0, 55)}`);
        skipEntries.push({ subject: email.subject, from: email.from, reason, labels: cls.skippedRaw || [] });
        failed++;
      } else {
        for (const labelName of cls.labels) {
          threads[idx].addLabel(getOrCreateLabel(labelName));
        }
        Logger.log(`✓ [${cls.labels.join(', ')}] ${email.subject.slice(0, 55)}`);
        done++;
      }
      props.setProperty('p_' + threads[idx].getId(), '1');
    } catch(e) {
      Logger.log(`✗ Error on ${idx}: ${e.message}`);
      skipEntries.push({ subject: email.subject, from: email.from, reason: 'Label apply error: ' + e.message, labels: [] });
      logError('applyClassifications', e.message, `Applied ${done} labels`, `Applying label to: ${email.subject.slice(0,50)}`);
      failed++;
    }
  }
 
  logSkipped(skipEntries);
  return { done, failed };
}
 
// ── MAIN FUNCTION (hourly trigger) ────────────────────────────────────────────
function classifyNewEmails() {
  let lastStep = 'Starting';
  try {
    const labels = loadLabelCriteria();
    if (!labels) return;
 
    const labelNames = labels.map(l => l.name);
    const labelDefs  = buildLabelDefs(labels);
    const props = PropertiesService.getScriptProperties();
    lastStep = 'Loaded criteria and props';
 
    const threads = GmailApp.search('in:inbox', 0, BATCH_SIZE * 3)
      .filter(t => !props.getProperty('p_' + t.getId()));
    const batch = threads.slice(0, BATCH_SIZE);
 
    if (batch.length === 0) { Logger.log('No new emails to classify.'); return; }
    Logger.log(`Found ${batch.length} unclassified threads.`);
    lastStep = `Fetched ${batch.length} threads`;
 
    const emailData = buildEmailData(batch);
    lastStep = 'Built email data';
 
    const raw = callGemini(emailData, labelDefs, labelNames);
    if (!raw || raw.length === 0) { Logger.log('No classifications returned.'); return; }
    lastStep = 'Received Gemini response';
 
    const classifications = filterClassifications(raw);
    const { done, failed } = applyClassifications(classifications, batch, emailData, props);
    Logger.log(`\n✅ Done: ${done} classified, ${failed} skipped/failed.`);
 
  } catch(e) {
    Logger.log('✗ classifyNewEmails failed: ' + e.message);
    logError('classifyNewEmails (hourly trigger)', e.message, lastStep, e.message, e.lineNumber || '');
  }
}
 
// ── ONE-TIME: Classify ALL existing inbox emails ──────────────────────────────
function classifyAllInbox() {
  const startTime = Date.now();
  let lastStep = 'Starting';
  try {
    const labels = loadLabelCriteria();
    if (!labels) return;
 
    const labelNames = labels.map(l => l.name);
    const labelDefs  = buildLabelDefs(labels);
    const props = PropertiesService.getScriptProperties();
 
    let start = parseInt(props.getProperty('classifyAllInbox_start') || '0');
    let total = parseInt(props.getProperty('classifyAllInbox_total') || '0');
    Logger.log(`Resuming from position ${start}, ${total} emails classified so far.`);
 
    while (true) {
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        props.setProperty('classifyAllInbox_start', String(start));
        props.setProperty('classifyAllInbox_total', String(total));
        Logger.log(`⏸ Paused at position ${start} — run classifyAllInbox again. Total so far: ${total}`);
        return;
      }
 
      lastStep = `Searching inbox at position ${start}`;
      const rawBatch = GmailApp.search('in:inbox', start, BATCH_SIZE * 3);
 
      // No more emails in inbox at all — we are done
      if (rawBatch.length === 0) {
        props.deleteProperty('classifyAllInbox_start');
        props.deleteProperty('classifyAllInbox_total');
        Logger.log(`\n✅ Complete! ${total} emails classified.`);
        return;
      }
 
      // Filter out already processed or already labeled
      const unprocessed = rawBatch.filter(t => {
        if (props.getProperty('p_' + t.getId())) return false;
        if (hasUserLabels(t)) { props.setProperty('p_' + t.getId(), '1'); return false; }
        return true;
      });
      const threads = unprocessed.slice(0, BATCH_SIZE);
 
      // All emails in this window are already labeled — advance and keep going
      if (threads.length === 0) {
        Logger.log(`Position ${start}: all ${rawBatch.length} emails already labeled, advancing...`);
        start += rawBatch.length;
        continue;
      }
 
      lastStep = `Building email data for ${threads.length} threads`;
      const emailData = buildEmailData(threads);
 
      lastStep = 'Calling Gemini';
      const raw = callGemini(emailData, labelDefs, labelNames);
 
      if (raw && raw.length > 0) {
        lastStep = 'Applying classifications';
        const classifications = filterClassifications(raw);
        const { done } = applyClassifications(classifications, threads, emailData, props);
        total += done;
      }
 
      Logger.log(`Processed ${total} emails so far...`);
      Utilities.sleep(1500);
      start += BATCH_SIZE;
    }
  } catch(e) {
    Logger.log('✗ classifyAllInbox failed: ' + e.message);
    logError('classifyAllInbox (manual run)', e.message, lastStep, e.message, e.lineNumber || '');
  }
}
 
// ── RESET FUNCTIONS ───────────────────────────────────────────────────────────
function resetClassifyAll() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('classifyAllInbox_start');
  props.deleteProperty('classifyAllInbox_total');
  Logger.log('Reset position counters — labels and processed markers untouched.');
}
 
function resetEverything() {
  const props = PropertiesService.getScriptProperties();
  // Save credentials before clearing
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const gmail  = props.getProperty('GMAIL_ADDRESS');
  props.deleteAllProperties();
  // Restore credentials so API key is never lost
  if (apiKey) props.setProperty('GEMINI_API_KEY', apiKey);
  if (gmail)  props.setProperty('GMAIL_ADDRESS', gmail);
  Logger.log('Full reset — processed markers cleared, credentials preserved.');
}
 
// ── OPEN LOG SHEET ────────────────────────────────────────────────────────────
function openLogSheet() {
  const folder = getAppFolder();
  const files = folder.getFilesByName(LOG_SHEET_NAME);
  if (files.hasNext()) {
    Logger.log('Log sheet URL: ' + SpreadsheetApp.open(files.next()).getUrl());
  } else {
    Logger.log('No log sheet found yet — it will be created on the next skipped email.');
  }
}
 
// ── CHECK IF THREAD HAS USER LABELS ──────────────────────────────────────────
function hasUserLabels(thread) {
  const systemLabels = new Set([
    'INBOX','UNREAD','STARRED','IMPORTANT','SENT','DRAFT','SPAM','TRASH',
    'CATEGORY_PERSONAL','CATEGORY_SOCIAL','CATEGORY_PROMOTIONS',
    'CATEGORY_UPDATES','CATEGORY_FORUMS'
  ]);
  return thread.getLabels().some(l => !systemLabels.has(l.getName()));
}
 
// ── BUILD EMAIL DATA ──────────────────────────────────────────────────────────
function buildEmailData(threads) {
  return threads.map((thread, i) => {
    const msg = thread.getMessages()[0];
    return {
      idx: String(i),
      from: msg.getFrom(),
      subject: msg.getSubject(),
      snippet: msg.getPlainBody().replace(/\s+/g, ' ').slice(0, 300)
    };
  });
}
 
// ── GEMINI API CALL ───────────────────────────────────────────────────────────
function callGemini(emailData, labelDefs, labelNames) {
  const emailsText = emailData.map(e =>
    `ID: ${e.idx}\nFrom: ${e.from}\nSubject: ${e.subject}\nSnippet: ${e.snippet}`
  ).join('\n\n---\n\n');
 
  const prompt = `You are an email classifier. Classify each email and return ONLY a valid JSON array — no explanation, no markdown, no code fences.
 
Available labels and their criteria:
${labelDefs}
 
Rules:
- Use EXACT label names including emojis: ${labelNames.map(n => `"${n}"`).join(', ')}
- Assign a maximum of ${MAX_LABELS} labels per email
- For each label, include a confidence score between 0.0 and 1.0
- Only assign a label if you are confident it applies — do not guess
- Every email must get at least one label
- IDs must exactly match the IDs provided
- Return format:
[
  {
    "id": "0",
    "labels": [
      {"name": "💼 Career", "confidence": 0.95},
      {"name": "✅ Task", "confidence": 0.88}
    ]
  },
  ...
]
 
Emails:
${emailsText}`;
 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
 
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 16384 }
      }),
      muteHttpExceptions: true
    });
 
    const data = JSON.parse(response.getContentText());
    if (data.error) {
      Logger.log('Gemini API error: ' + JSON.stringify(data.error));
      logError('callGemini', JSON.stringify(data.error), 'Sent request to Gemini', 'Gemini API returned error');
      return null;
    }
 
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
    Logger.log('Gemini output (first 400 chars): ' + text.slice(0, 400));
    let cleaned = text.replace(/```json|```/g, '').trim();
 
    try {
      return JSON.parse(cleaned);
    } catch(e) {
      // Attempt repair: trim to last complete object
      const lastBrace = cleaned.lastIndexOf('}');
      if (lastBrace !== -1) {
        cleaned = cleaned.slice(0, lastBrace + 1) + ']';
        try { return JSON.parse(cleaned); } catch(e2) {
          logError('callGemini', 'JSON repair failed: ' + e2.message, 'Received Gemini response', 'JSON parse after repair', '');
        }
      }
      Logger.log('JSON repair failed: ' + e.message);
      return null;
    }
  } catch(e) {
    Logger.log('Gemini call failed: ' + e.message);
    logError('callGemini', e.message, 'Sending request to Gemini', 'UrlFetchApp.fetch failed', e.lineNumber || '');
    return null;
  }
}
 
// ── LABEL HELPER ─────────────────────────────────────────────────────────────
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) { label = GmailApp.createLabel(name); Logger.log('Created label: ' + name); }
  return label;
}


// reset already processed emails this will remove all the labels
function resetProcessed() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Reset complete.');
}

function resetEverything() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Full reset complete — all processed markers cleared.');
}

//Diagnostics
function diagnose() {
  const props = PropertiesService.getScriptProperties();
  const threads = GmailApp.search('in:inbox', 0, 10);
  
  threads.forEach((t, i) => {
    const msg = t.getMessages()[0];
    const labels = t.getLabels().map(l => l.getName());
    const processed = props.getProperty('p_' + t.getId());
    Logger.log(`[${i}] Subject: ${msg.getSubject().slice(0,40)}`);
    Logger.log(`     Labels: ${labels.join(', ') || 'NONE'}`);
    Logger.log(`     Processed marker: ${processed || 'NOT SET'}`);
  });
}