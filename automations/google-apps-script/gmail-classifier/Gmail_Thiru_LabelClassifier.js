// ============================================================
// Gmail Auto-Classifier using Gemini
// ============================================================
// SETUP:
// 1. Go to script.google.com — sign in with santhosh.thiruchendru@straventis.com
// 2. Create a new project, paste this entire script
// 3. Upload label_criteria.json to Google Drive root
// 4. Replace GEMINI_API_KEY with your key from aistudio.google.com
// 5. Run classifyNewEmails and authorize when prompted
// 6. Set trigger: classifyNewEmails → Time-driven → Every hour
// ============================================================

const GMAIL_ADDRESS  = 'santhosh.thiruchendru@straventis.com';   // Workspace Gmail to classify
const GEMINI_API_KEY     = 'AQ.Ab8RN6KNYeTODluAOAHE3qwTDIcS4878w4ZdZwIHZk--dqM8Zw';
const CRITERIA_FILE_NAME = 'label_criteria.json';  // must be in Google Drive root
const BATCH_SIZE         = 20;
const GEMINI_MODEL       = 'gemini-2.5-flash';
const MAX_LABELS         = 3;      // max labels per email
const MIN_CONFIDENCE     = 0.75;   // only apply labels with confidence >= this (0.0 - 1.0)

// ── LOAD LABEL CRITERIA FROM DRIVE ───────────────────────────────────────────
function loadLabelCriteria() {
  try {
    const files = DriveApp.getFilesByName(CRITERIA_FILE_NAME);
    if (!files.hasNext()) throw new Error(`"${CRITERIA_FILE_NAME}" not found in Google Drive.`);
    const content = files.next().getBlob().getDataAsString();
    const parsed = JSON.parse(content);
    Logger.log(`✓ Loaded ${parsed.labels.length} labels from ${CRITERIA_FILE_NAME}`);
    return parsed.labels;
  } catch(e) {
    Logger.log('✗ Could not load criteria file: ' + e.message);
    return null;
  }
}

function buildLabelDefs(labels) {
  return labels.map(l => `"${l.name}": ${l.criteria}`).join('\n\n');
}

// ── FILTER: apply confidence threshold and max label cap ─────────────────────
function filterClassifications(classifications) {
  return classifications.map(cls => {
    // Sort by confidence descending, filter by threshold, cap at MAX_LABELS
    const filtered = (cls.labels || [])
      .filter(l => (l.confidence || 1.0) >= MIN_CONFIDENCE)
      .sort((a, b) => (b.confidence || 1.0) - (a.confidence || 1.0))
      .slice(0, MAX_LABELS)
      .map(l => l.name);
    return { id: cls.id, labels: filtered };
  });
}

// ── MAIN FUNCTION ─────────────────────────────────────────────────────────────
function classifyNewEmails() {
  const labels = loadLabelCriteria();
  if (!labels) return;

  const labelNames = labels.map(l => l.name);
  const labelDefs  = buildLabelDefs(labels);
  // Track processed threads using Apps Script Properties (invisible — no label created)
  const props = PropertiesService.getScriptProperties();

  const threads = GmailApp.search('in:inbox', 0, BATCH_SIZE * 3).filter(t => !props.getProperty('p_' + t.getId()));
  const batch = threads.slice(0, BATCH_SIZE);
  if (threads.length === 0) { Logger.log('No new emails to classify.'); return; }
  Logger.log(`Found ${threads.length} unclassified threads.`);

  const emailData = buildEmailData(batch);
  const raw = callGemini(emailData, labelDefs, labelNames);
  if (!raw || raw.length === 0) { Logger.log('No classifications returned.'); return; }

  const classifications = filterClassifications(raw);
  Logger.log(`Gemini returned ${classifications.length} classifications.`);

  let done = 0, failed = 0;
  for (const cls of classifications) {
    const idx = parseInt(cls.id);
    if (isNaN(idx) || idx < 0 || idx >= batch.length) continue;
    try {
      for (const labelName of cls.labels) {
        batch[idx].addLabel(getOrCreateLabel(labelName));
      }
      // Mark as processed invisibly
      props.setProperty('p_' + batch[idx].getId(), '1');
      Logger.log(`✓ [${cls.labels.join(', ')}] ${emailData[idx].subject.slice(0, 55)}`);
      done++;
    } catch(e) {
      Logger.log(`✗ Error on ${idx}: ${e.message}`);
      failed++;
    }
  }
  Logger.log(`\n✅ Done: ${done} classified, ${failed} failed.`);
}

// ── ONE-TIME: Classify ALL existing inbox emails ──────────────────────────────
function classifyAllInbox() {
  const labels = loadLabelCriteria();
  if (!labels) return;

  const labelNames = labels.map(l => l.name);
  const labelDefs  = buildLabelDefs(labels);
  let start = 0, total = 0;

  while (true) {
    const props = PropertiesService.getScriptProperties();
    const allThreads = GmailApp.search('in:inbox', start, BATCH_SIZE * 3).filter(t => !props.getProperty('p_' + t.getId()));
    const threads = allThreads.slice(0, BATCH_SIZE);
    if (threads.length === 0) break;

    const emailData = buildEmailData(threads);
    const raw = callGemini(emailData, labelDefs, labelNames);

    if (raw && raw.length > 0) {
      const classifications = filterClassifications(raw);
      for (const cls of classifications) {
        const idx = parseInt(cls.id);
        if (isNaN(idx) || idx >= threads.length) continue;
        try {
          for (const labelName of cls.labels) {
            threads[idx].addLabel(getOrCreateLabel(labelName));
          }
          props.setProperty('p_' + threads[idx].getId(), '1');
          total++;
        } catch(e) { Logger.log('Error: ' + e.message); }
      }
    }

    Logger.log(`Processed ${total} emails so far...`);
    Utilities.sleep(1500);
    start += BATCH_SIZE;
  }
  Logger.log(`\n✅ Complete! ${total} emails classified.`);
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
        generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
      }),
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());
    if (data.error) { Logger.log('Gemini error: ' + JSON.stringify(data.error)); return null; }

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
    Logger.log('Gemini output (first 400 chars): ' + text.slice(0, 400));
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch(e) {
    Logger.log('Gemini call failed: ' + e.message);
    return null;
  }
}

// ── LABEL HELPER ─────────────────────────────────────────────────────────────
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) { label = GmailApp.createLabel(name); Logger.log('Created label: ' + name); }
  return label;
}

//
function resetProcessed() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Reset complete.');
}
