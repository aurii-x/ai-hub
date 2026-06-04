// ============================================================
// Gmail Auto-Classifier using Gemini
// ============================================================
// SETUP:
// 1. Go to script.google.com (sign in with Workspace account)
// 2. Create a new project, paste this entire script
// 3. Replace GEMINI_API_KEY with your key from aistudio.google.com
// 4. Replace GMAIL_ADDRESS with your personal Gmail address
// 5. Click Run → classifyNewEmails (authorize when prompted)
// 6. Set up a trigger: Triggers → Add Trigger → classifyNewEmails → Time-driven → Every hour
// ============================================================
const GEMINI_API_KEY = 'AQ.Ab8RN6KNYeTODluAOAHE3qwTDIcS4878w4ZdZwIHZk--dqM8Zw';
const GMAIL_ADDRESS  = 'santhosh.thiruchendru@straventis.com';   // Workspace Gmail to classify
const BATCH_SIZE     = 20;                            // emails per Gemini call
const PROCESSED_LABEL = 'xAI-Classified';              // tracks what's already been classified

const LABEL_DEFS = `"💼 Career": All Career related emails — parent label
"Career/Applications": Confirmation emails, application status updates, ATS responses (Workday, Ashby, iCIMS, Greenhouse)
"Career/Recruiters": Direct recruiter outreach, LinkedIn InMail forwards, agency emails, staffing agencies
"Career/Interviews": Interview scheduling, prep materials, panel introductions, Zoom links, calendar invites
"Career/Offers & Rejections": Offer letters, rejection emails, salary negotiations, counter-offer threads
"Learning": Course enrollments, certification updates, learning platform emails (Databricks Academy, Snowflake, Coursera, webinar invites)
"Learning/Certs & Exams": Exam confirmations, certification badges, credential emails (Credly badges, exam vouchers, pass/fail notifications)
"🗞️ Newsletters": Tech newsletters, industry digests, data engineering weekly, AI newsletters, tech blogs
"🩺 Health": Medical appointments, insurance, prescriptions, health portals (MyChart, BCBSNC, Dupixent, Azstarys, doctor offices)
"Home & Family": Home services, school emails, kids activities, family logistics (Elijah/Caleb/Ami related, home repair, utilities)
"💰Finance": Banking, bills, subscriptions, receipts, tax documents (bank statements, Setapp invoices, Stripe receipts)
"⚖️ Legal": Attorney correspondence, court documents, separation related emails (Evan emails, legal filings, affidavits, custody documents)
"🔲 Action Required": Needs a response or action today — highest urgency (recruiter follow-ups due, legal deadlines, appointment confirmations needed)
"🟡 Waiting On": Sent emails waiting for a reply — blocked on someone else (job applications submitted, recruiter follow-ups sent)
"🟢 Read Later": Interesting but not urgent — read when you have focused time (tech articles, newsletters, webinar recordings)
"Reference": Keep for records but no action needed — archive after labeling (receipts, confirmations, account notifications)
"Straventis": All Straventis consulting related emails (client emails, SOW, invoices, project updates)
"🗑️ Delete": Mark for deletion — promotional spam, irrelevant newsletters, expired offers, marketing emails, old job alerts`;

// ── MAIN FUNCTION ─────────────────────────────────────────────────────────────
function classifyNewEmails() {
  const processedLabel = getOrCreateLabel(PROCESSED_LABEL);
  const threads = GmailApp.search(`in:inbox -label:${PROCESSED_LABEL}`, 0, BATCH_SIZE);

  if (threads.length === 0) {
    Logger.log('No new emails to classify.');
    return;
  }
  Logger.log(`Found ${threads.length} unclassified threads.`);

  // Build email data using simple index as ID (avoids any ID mismatch)
  const emailData = threads.map((thread, i) => {
    const msg = thread.getMessages()[0];
    return {
      idx: String(i),  // simple numeric index as ID
      from: msg.getFrom(),
      subject: msg.getSubject(),
      snippet: msg.getPlainBody().replace(/\s+/g, ' ').slice(0, 250)
    };
  });

  // Call Gemini
  const classifications = callGemini(emailData);

  if (!classifications || classifications.length === 0) {
    Logger.log('No classifications returned from Gemini.');
    return;
  }

  Logger.log(`Gemini returned ${classifications.length} classifications.`);

  // Apply labels using index to match back to thread
  let done = 0, failed = 0;
  for (const cls of classifications) {
    const idx = parseInt(cls.id);
    if (isNaN(idx) || idx < 0 || idx >= threads.length) {
      Logger.log(`⚠ Invalid index: ${cls.id}`);
      continue;
    }

    const thread = threads[idx];
    const labels = cls.labels || [];

    try {
      for (const labelName of labels) {
        const label = getOrCreateLabel(labelName);
        thread.addLabel(label);
      }
      thread.addLabel(processedLabel);
      Logger.log(`✓ [${labels.join(', ')}] → ${emailData[idx].subject.slice(0, 55)}`);
      done++;
    } catch(e) {
      Logger.log(`✗ Error on thread ${idx}: ${e.message}`);
      failed++;
    }
  }

  Logger.log(`\n✅ Done: ${done} classified, ${failed} failed.`);
}

// ── GEMINI API CALL ───────────────────────────────────────────────────────────
function callGemini(emailData) {
  const emailsText = emailData.map(e =>
    `ID: ${e.idx}\nFrom: ${e.from}\nSubject: ${e.subject}\nSnippet: ${e.snippet}`
  ).join('\n\n---\n\n');

  const prompt = `Classify these Gmail emails into labels. Return ONLY a valid JSON array — no explanation, no markdown, no code fences.

Available labels:
${LABEL_DEFS}

Rules:
- Use EXACT label names (including emojis)
- Assign 1-3 labels per email
- Apply both a category label AND an action label when appropriate
- Every email must get at least one label
- IDs in your response must exactly match the IDs provided
- Return format: [{"id":"0","labels":["Label1"]},{"id":"1","labels":["Label2","Label3"]},...]

Emails:
${emailsText}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const raw = response.getContentText();
    Logger.log('Gemini raw response: ' + raw.slice(0, 500));

    const data = JSON.parse(raw);

    if (data.error) {
      Logger.log('Gemini API error: ' + JSON.stringify(data.error));
      return null;
    }

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
    Logger.log('Gemini text output: ' + text.slice(0, 500));

    const cleaned = text.replace(/```json|```/g, '').trim();
    return JSON.parse(cleaned);
  } catch(e) {
    Logger.log('Gemini call failed: ' + e.message);
    return null;
  }
}

// ── LABEL HELPER ─────────────────────────────────────────────────────────────
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
    Logger.log('Created label: ' + name);
  }
  return label;
}

// ── ONE-TIME: Classify ALL existing inbox emails ───────────────────────────────
function classifyAllInbox() {
  let start = 0;
  let total = 0;

  while (true) {
    const threads = GmailApp.search(`in:inbox -label:${PROCESSED_LABEL}`, start, BATCH_SIZE);
    if (threads.length === 0) break;

    const emailData = threads.map((thread, i) => {
      const msg = thread.getMessages()[0];
      return {
        idx: String(i),
        from: msg.getFrom(),
        subject: msg.getSubject(),
        snippet: msg.getPlainBody().replace(/\s+/g, ' ').slice(0, 250)
      };
    });

    const classifications = callGemini(emailData);
    if (classifications && classifications.length > 0) {
      for (const cls of classifications) {
        const idx = parseInt(cls.id);
        if (isNaN(idx) || idx >= threads.length) continue;
        try {
          for (const labelName of (cls.labels || [])) {
            threads[idx].addLabel(getOrCreateLabel(labelName));
          }
          threads[idx].addLabel(getOrCreateLabel(PROCESSED_LABEL));
          total++;
        } catch(e) {
          Logger.log('Error: ' + e.message);
        }
      }
    }

    Logger.log(`Processed ${total} emails so far...`);
    Utilities.sleep(1500);
    start += BATCH_SIZE;
  }

  Logger.log(`\n✅ Complete! ${total} emails classified.`);
}