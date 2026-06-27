// ============================================================
// CLICKUP TODOIST MIGRATION SCRIPT
// Generated: June 2026 | santhosh.thiru@gmail.com
// ============================================================
// SETUP:
//   1. Go to script.google.com → New Project
//   2. Paste this entire script
//   3. Run setup() once with your ClickUp API token
//   4. Run migrateAll() — it will run in 5-min chunks and resume
//   5. Check the MigrationLog sheet for progress
// ============================================================

// ---- CONFIGURATION ----
const CU_API_BASE = 'https://api.clickup.com/api/v2';
const LOG_SHEET_NAME = 'MigrationLog';
const BATCH_SIZE = 10; // tasks per batch before checking time
const MAX_RUNTIME_MS = 5 * 60 * 1000;

// ---- LIST ID MAP (created in ClickUp) ----
const LIST_IDS = {
  'career_deep_work':       '901417224062',
  'career_job_search':      '901417224063',
  'career_communication':   '901417224064',
  'learning_courses':       '901417224065',
  'learning_research':      '901417224067',
  'personal_health':        '901417224069',
  'personal_family':        '901417224070',
  'personal_finance_legal': '901417224072',
  'personal_rest':          '901417224073',
  'device_tinkering':       '901417224075',
  'device_focus':           '901417224080',
  'web_productive':         '901417224082',
  'web_news':               '901417224083',
  'web_social':             '901417224086',
  'straventis_client':      '901417224087',
  'straventis_coding':      '901417224091',
  'mindless_vortex':        '901417224093',
  'mindless_drift':         '901417224094',
  'inbox_today':            '901417224095',
  'inbox_tomorrow':         '901417224099',
  'inbox_wip':              '901417224102',
};

// ---- HELPER: parse duration string to minutes ----
function parseDurationToMinutes(dur) {
  if (!dur) return null;
  var total = 0;
  var h = dur.match(/(\d+)h/);
  var m = dur.match(/(\d+)m/);
  if (h) total += parseInt(h[1]) * 60;
  if (m) total += parseInt(m[1]);
  return total > 0 ? total : null;
}

// ---- SETUP: store API token ----
function setup() {
  var token = 'CLICKUP_TOKEN'; // replace before running
  PropertiesService.getScriptProperties().setProperty('CU_TOKEN', token);
  Logger.log('✅ Token saved.');
}

function getToken() {
  var t = PropertiesService.getScriptProperties().getProperty('CU_TOKEN');
  if (!t || t === 'YOUR_CLICKUP_API_TOKEN_HERE') throw new Error('Run setup() first with your ClickUp API token.');
  return t;
}

// ---- CLICKUP API CALLS ----
function cuRequest(method, path, payload) {
  var options = {
    method: method,
    headers: {
      'Authorization': getToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  var resp = UrlFetchApp.fetch(CU_API_BASE + path, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code >= 400) {
    throw new Error('ClickUp API error ' + code + ': ' + body);
  }
  return JSON.parse(body);
}

var PRIORITY_MAP = { 'urgent': 1, 'high': 2, 'normal': 3, 'low': 4 };

function createTask(listId, task) {
  var payload = {
    name: task.name,
    description: task.description || '',
    due_date: task.due_date ? new Date(task.due_date).getTime() : null,
    status: task.completed ? 'complete' : null,
    time_estimate: task.duration_minutes ? task.duration_minutes * 60000 : null
  };
  // numeric priority — ClickUp API requires 1/2/3/4 not strings
  if (task.priority && PRIORITY_MAP[task.priority]) {
    payload.priority = PRIORITY_MAP[task.priority];
  }
  // tags must be array of objects {name: '...'}
  if (task.tags && task.tags.length) {
    payload.tags = task.tags.map(function(t) { return { name: t }; });
  }
  // remove null/undefined
  Object.keys(payload).forEach(function(k) {
    if (payload[k] === null || payload[k] === undefined) delete payload[k];
  });
  return cuRequest('POST', '/list/' + listId + '/task', payload);
}

// ---- LOG SHEET ----
function getOrCreateLogSheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('LOG_SHEET_ID');
  var ss;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch(e) {
      ssId = null; // sheet was deleted, recreate
    }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create('ClickUp Migration Log');
    props.setProperty('LOG_SHEET_ID', ss.getId());
    Logger.log('📋 Log sheet created: ' + ss.getUrl());
  }

  var sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(['Timestamp','Task Name','List','Status','CU Task ID','Error']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function logResult(sheet, taskName, listKey, status, cuId, error) {
  sheet.appendRow([
    new Date().toISOString(),
    taskName,
    listKey,
    status,
    cuId || '',
    error || ''
  ]);
}

// ---- TASK DATA ----
// Priority mapping from Todoist: p1=urgent, p2=high, p3=normal, p4=low
// completed=true tasks will be created with status 'complete'

function getAllTasks() {
  return [

    // ================================================================
    // 1.0 CAREER DEVELOPMENT — OPEN TASKS
    // ================================================================

    // 1.1 Deep Work
    { list: 'career_deep_work', name: 'ResumeOS v3.0 — Strategic Positioning Overhaul', description: 'Complete overhaul of resume positioning for senior TPM/Director AI & data roles.', priority: 'high', due_date: '2026-06-15', tags: ['resume-updates'], completed: false },
    { list: 'career_deep_work', name: 'Build OnePipeline Case Study Document', description: 'Write a polished case study of the OnePipeline data platform program for portfolio.', priority: 'normal', tags: ['portfolio'], completed: false },
    { list: 'career_deep_work', name: 'STAR Interview Bank — AI/Data TPM', description: 'Build comprehensive STAR response bank for AI/data TPM interview scenarios.', priority: 'high', due_date: '2026-06-20', tags: [], completed: false },

    // 1.2 Job Search
    { list: 'career_job_search', name: 'Apply for Job-1', description: 'Recurring daily job application task.', priority: 'high', due_date: '2026-06-12', duration_minutes: 30, tags: [], completed: false },
    { list: 'career_job_search', name: 'Apply for Job-2', description: 'Recurring daily job application task.', priority: 'high', due_date: '2026-06-13', duration_minutes: 30, tags: [], completed: false },
    { list: 'career_job_search', name: 'Apply for Job-3', description: 'Recurring daily job application task.', priority: 'high', due_date: '2026-06-16', duration_minutes: 30, tags: [], completed: false },
    { list: 'career_job_search', name: 'Apply for Job-4', description: 'Recurring daily job application task.', priority: 'high', due_date: '2026-06-17', duration_minutes: 30, tags: [], completed: false },
    { list: 'career_job_search', name: 'Apply for Job-5', description: 'Recurring daily job application task.', priority: 'high', due_date: '2026-06-18', duration_minutes: 30, tags: [], completed: false },
    { list: 'career_job_search', name: 'Careerflow — Update Job Tracker', description: 'Sync all recent applications and pipeline status in Careerflow.', priority: 'normal', due_date: '2026-06-14', tags: [], completed: false },
    { list: 'career_job_search', name: 'LinkedIn Profile Optimization', description: 'Update LinkedIn headline, about section, and featured section for AI/data TPM positioning.', priority: 'normal', tags: [], completed: false },

    // 1.3 Communication
    { list: 'career_communication', name: 'Follow up with Casey Sortini at Rayonic', description: 'Follow up on initial outreach to Casey Sortini.', priority: 'normal', due_date: '2026-06-14', tags: [], completed: false },
    { list: 'career_communication', name: 'Send thank-you note — BuildOps PM Interview', description: 'Send follow-up thank-you note after BuildOps Staff PM interview.', priority: 'high', due_date: '2026-06-12', tags: [], completed: false },

    // ================================================================
    // 2.0 LEARNING — OPEN TASKS
    // ================================================================

    // 2.1 Courses & Certs
    { list: 'learning_courses', name: 'Databricks Data Engineer — Complete remaining modules', description: 'Finish remaining Databricks data engineering learning plan modules on customer-academy.databricks.com.', priority: 'high', due_date: '2026-06-20', duration_minutes: 180, tags: ['learning'], completed: false },
    { list: 'learning_courses', name: 'Module 4: The Product Retire and Next Phase', description: 'https://customer-academy.databricks.com', priority: 'high', due_date: '2026-06-11', tags: ['learning'], completed: false },
    { list: 'learning_courses', name: 'G565 — Game Theory Final Exam Prep', description: 'Review signaling games, Nash equilibrium, and payoff matrix problems for final.', priority: 'urgent', due_date: '2026-06-18', tags: ['mba', 'learning'], completed: false },
    { list: 'learning_courses', name: 'MBA — G565 Final Submission', description: 'Submit final deliverable for G565 Economics/Game Theory course.', priority: 'urgent', due_date: '2026-06-22', tags: ['mba'], completed: false },
    { list: 'learning_courses', name: 'Coursera PM Specialization — Course 4', description: 'Begin next course in Coursera Product Management specialization.', priority: 'normal', tags: ['learning', 'pm-cert'], completed: false },

    // 2.2 Reading & Research
    { list: 'learning_research', name: 'Research: Agentic AI Architecture Patterns', description: 'Research agent orchestration patterns (ReAct, plan-and-execute, multi-agent) for TPM knowledge base.', priority: 'normal', tags: ['ai', 'learning'], completed: false },
    { list: 'learning_research', name: 'Read: The Design of Web APIs', description: 'Continue reading and annotating The Design of Web APIs — focus on REST/GraphQL tradeoffs.', priority: 'normal', tags: ['learning'], completed: false },
    { list: 'learning_research', name: 'SOA & Microservices Research Notes', description: 'Compile SOA vs microservices comparison notes for interview prep and portfolio.', priority: 'low', tags: ['learning'], completed: false },

    // ================================================================
    // 3.0 PERSONAL & LIFE — OPEN TASKS
    // ================================================================

    // 3.1 Health & Wellness
    { list: 'personal_health', name: 'Schedule sleep study follow-up', description: 'Call and schedule follow-up appointment after sleep study.', priority: 'high', due_date: '2026-06-15', tags: [], completed: false },
    { list: 'personal_health', name: 'Schedule MRI follow-up', description: 'Call and confirm MRI results and schedule any follow-up.', priority: 'high', due_date: '2026-06-15', tags: [], completed: false },
    { list: 'personal_health', name: 'Refill Azstarys prescription', description: 'Call Walgreens to order/refill Azstarys.', priority: 'high', due_date: '2026-06-14', tags: [], completed: false },
    { list: 'personal_health', name: 'Schedule therapy appointment', description: 'Book next therapy session with Dr. Madison.', priority: 'normal', due_date: '2026-06-20', tags: [], completed: false },
    { list: 'personal_health', name: 'Nia — PCP appointment', description: 'Schedule Nia PCP checkup appointment.', priority: 'normal', due_date: '2026-06-20', tags: [], completed: false },
    { list: 'personal_health', name: 'Nia — Dermatology follow-up', description: 'Follow up on Nia dermatology appointment at BlueRidge.', priority: 'normal', due_date: '2026-06-20', tags: [], completed: false },
    { list: 'personal_health', name: 'Caleb — Dentist cavity filling', description: 'Confirm and attend Caleb cavity filling appointment at Wells Family Dentistry.', priority: 'normal', due_date: '2026-06-18', tags: [], completed: false },

    // 3.2 Home & Family
    { list: 'personal_family', name: 'Plan summer activity schedule for kids', description: 'Plan and organize summer schedule for Caleb, Elijah, Nia, and Ami.', priority: 'high', due_date: '2026-06-20', tags: [], completed: false },
    { list: 'personal_family', name: 'Elijah — LRHS Band prep', description: 'Support Elijah with LRHS band-related prep and materials.', priority: 'normal', tags: [], completed: false },
    { list: 'personal_family', name: 'Ami — School end-of-year items', description: 'Handle any remaining end-of-year school items for Ami.', priority: 'normal', due_date: '2026-06-15', tags: [], completed: false },
    { list: 'personal_family', name: 'Mow lawn', description: '', priority: 'low', tags: [], completed: false },
    { list: 'personal_family', name: 'Home maintenance checklist — Summer', description: 'Go through summer home maintenance items: HVAC filter, gutters, etc.', priority: 'low', tags: [], completed: false },

    // 3.3 Finance & Legal
    { list: 'personal_finance_legal', name: 'Review separation agreement — final version', description: 'Review final separation agreement draft and confirm all terms before signing.', priority: 'urgent', due_date: '2026-06-15', tags: ['legal'], completed: false },
    { list: 'personal_finance_legal', name: 'Compile dating app records for case', description: 'Compile records of dating app activity and children inquiries circa 2020–2022 for legal case.', priority: 'high', due_date: '2026-06-15', tags: ['legal'], completed: false },
    { list: 'personal_finance_legal', name: 'Document Dr. Madison affidavit request', description: 'Follow up with Dr. Madison on formal diagnosis documentation and affidavit for case.', priority: 'high', tags: ['legal'], completed: false },
    { list: 'personal_finance_legal', name: 'Update BCBSNC payment information', description: 'Update insurance payment method with BCBSNC.', priority: 'high', due_date: '2026-06-14', tags: [], completed: false },
    { list: 'personal_finance_legal', name: 'Update Morgan Stanley mortgage payment info', description: 'Update payment info with Morgan Stanley for mortgage.', priority: 'high', due_date: '2026-06-14', tags: [], completed: false },
    { list: 'personal_finance_legal', name: 'Monarch — monthly finance reconciliation', description: 'Reconcile monthly finances in Monarch, check subscriptions and budget.', priority: 'normal', due_date: '2026-06-30', tags: [], completed: false },

    // 3.4 Rest & Entertainment
    { list: 'personal_rest', name: 'Plan weekend activity with kids', description: 'Plan a fun weekend outing or activity with the kids.', priority: 'normal', due_date: '2026-06-14', tags: [], completed: false },

    // ================================================================
    // 4.0 DEVICE & TINKERING — OPEN TASKS
    // ================================================================

    // 4.1 System & Tinkering
    { list: 'device_tinkering', name: 'santhoshOS Phase 2 — Planning', description: 'Plan and scope Phase 2 of santhoshOS personal AI workspace build.', priority: 'high', tags: ['santhoshos'], completed: false },
    { list: 'device_tinkering', name: 'LobeChat — Configure knowledge base with Notion sync', description: 'Set up LobeChat knowledge base connected to Notion pages for context-aware AI queries.', priority: 'normal', tags: ['santhoshos', 'notion'], completed: false },
    { list: 'device_tinkering', name: 'Notion Migration — Complete remaining source databases', description: 'Move remaining legacy Notion databases into the 8-database clean structure.', priority: 'high', due_date: '2026-06-25', tags: ['notion'], completed: false },
    { list: 'device_tinkering', name: 'Notion — Manual Status property updates across 8 DBs', description: 'Manually update Status property options in each of the 8 Notion databases via UI (API DDL unsupported).', priority: 'normal', tags: ['notion'], completed: false },
    { list: 'device_tinkering', name: 'Notion — Property grouping layout (all 8 DBs)', description: 'Set up Auto Properties, DB Properties, Related To, Common Properties section groupings in all 8 DB layouts manually.', priority: 'normal', tags: ['notion'], completed: false },
    { list: 'device_tinkering', name: 'Gmail Classifier — Review label_criteria.json and tune MIN_CONFIDENCE', description: 'Review skipped emails log and tune confidence threshold and criteria descriptions.', priority: 'normal', tags: ['automation'], completed: false },
    { list: 'device_tinkering', name: 'Set up Zapier: Rize → ClickUp time entry sync', description: 'Build Zapier automation that fires when Rize session ends and logs time to matching ClickUp task.', priority: 'low', tags: ['automation', 'rize'], completed: false },

    // 4.2 Focus Tools
    { list: 'device_focus', name: 'Rize — Audit category taxonomy alignment with ClickUp', description: 'Align Rize category names with ClickUp project/list naming convention for better matching.', priority: 'normal', tags: ['rize'], completed: false },
    { list: 'device_focus', name: 'BTT — Update floating web view shortcuts', description: 'Update BetterTouchTool shortcuts and floating web view triggers for current workflow.', priority: 'low', tags: [], completed: false },
    { list: 'device_focus', name: 'Todoist → ClickUp: Update Rize durations after CSV export', description: 'Export Rize CSV, match task names, and update time estimates on ClickUp open tasks.', priority: 'normal', tags: ['rize'], completed: false },

    // ================================================================
    // 5.0 WEB BROWSING — OPEN TASKS
    // ================================================================
    { list: 'web_productive', name: 'Review Arc browser spaces setup', description: 'Audit and reorganize Arc browser spaces to align with 1.0–9.0 naming convention.', priority: 'low', tags: [], completed: false },

    // ================================================================
    // 6.0 STRAVENTIS — OPEN TASKS
    // ================================================================

    // 6.1 Client Work
    { list: 'straventis_client', name: 'ParticleBlack/Ascension — Program status update', description: 'Prepare and send program status update to ParticleBlack/Ascension stakeholders.', priority: 'high', due_date: '2026-06-15', tags: ['ascension'], completed: false },
    { list: 'straventis_client', name: 'Straventis — Update invoicing and hours log', description: 'Reconcile Ascension hours and update invoice and reporting for Straventis.', priority: 'high', due_date: '2026-06-20', tags: [], completed: false },
    { list: 'straventis_client', name: 'Data Governance framework doc — Xperi', description: 'Complete data governance framework document synthesizing IAM, pipeline architecture, and compliance requirements.', priority: 'normal', tags: ['xperi'], completed: false },

    // 6.2 Coding & Scripting
    { list: 'straventis_coding', name: 'santhoshOS — Docker Compose cleanup and documentation', description: 'Clean up Docker Compose config, document all services, and commit to Git.', priority: 'normal', tags: ['santhoshos'], completed: false },
    { list: 'straventis_coding', name: 'Plaud → Notion automation — validate and monitor', description: 'Check Zapier Plaud Notes → Notion automation is running correctly and fix any failures.', priority: 'normal', tags: ['automation', 'notion'], completed: false },
    { list: 'straventis_coding', name: 'React Dashboard UI — next iteration', description: 'Continue Straventis dashboard development with Horizon UI template and multi-AI features.', priority: 'low', tags: [], completed: false },

    // ================================================================
    // INBOX — OPEN TASKS (unassigned/misc)
    // ================================================================
    { list: 'inbox_today', name: 'Submit referral for substitute teacher', description: '', priority: 'normal', due_date: '2026-06-12', tags: [], completed: false },
    { list: 'inbox_today', name: 'Schedule Nia therapy with Kathy Caputo', description: 'At Chronic Hope Counseling.', priority: 'high', due_date: '2026-06-12', tags: [], completed: false },
    { list: 'inbox_wip', name: 'Complete documenting Kathy Caputo incident for 2025', description: 'https://www.notion.so/Could-have-stayed-With-Kathy-Caputo-3626187008a280f5b96becb5541d03e2', priority: 'high', tags: ['legal'], completed: false },

    // ================================================================
    // 1.0 CAREER DEVELOPMENT — COMPLETED TASKS
    // ================================================================

    // 1.2 Job Search — completed applications
    { list: 'career_job_search', name: 'Apply to Lead Portfolio Management Specialist at Duke Energy', description: '', priority: 'low', due_date: '2026-06-10', duration_minutes: 30, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at BuildOps / Labcorp', description: '', priority: 'low', due_date: '2026-06-09', duration_minutes: 46, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at KPMG / NetApp', description: '', priority: 'low', due_date: '2026-06-05', duration_minutes: 22, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Product Manager at EY', description: '', priority: 'low', due_date: '2026-06-05', duration_minutes: 46, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Staff Product Manager - Platform at BuildOps', description: '', priority: 'low', due_date: '2026-05-28', duration_minutes: 251, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Data Strategy Dir / Dir of AI Adoption at PwC / EvenUP', description: '', priority: 'low', due_date: '2026-05-26', duration_minutes: 120, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal TPM / Data Strategy Dir at MaintainX / PwC', description: '', priority: 'low', due_date: '2026-05-21', duration_minutes: 299, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Research Project Manager at UNC', description: '', priority: 'low', due_date: '2026-05-18', duration_minutes: 180, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Senior Program Manager at CGI', description: '', priority: 'low', due_date: '2026-05-18', duration_minutes: 240, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Product Manager at NetApp', description: '', priority: 'low', due_date: '2026-05-14', duration_minutes: 60, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Engineering Program Manager at Cisco', description: '', priority: 'low', due_date: '2026-05-14', duration_minutes: 600, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Director Product Management — Composite AI at MetLife', description: '', priority: 'low', due_date: '2026-05-13', duration_minutes: 180, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sr. Product Manager / TPM at Aspida Financial Services', description: '', priority: 'low', due_date: '2026-05-12', duration_minutes: 180, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Project Management Office Director at Wake County', description: '', priority: 'low', due_date: '2026-05-07', duration_minutes: 60, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sr. Program Manager / Sr Tech PM at Toshiba / GE Vernova', description: '', priority: 'low', due_date: '2026-05-06', duration_minutes: 180, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various PM/Director roles at Epic Games', description: '', priority: 'low', due_date: '2026-05-05', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sr. Product Manager for LIMS at Labcorp / TEKsystems', description: '', priority: 'low', due_date: '2026-05-01', duration_minutes: 210, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sr. Product Manager for LIMS at Labcorp', description: '', priority: 'low', due_date: '2026-04-29', duration_minutes: 150, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sr. Product Manager for LIMS at TEKsystems / Labcorp', description: '', priority: 'low', due_date: '2026-04-29', duration_minutes: 150, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Technical Program Manager at H2O.ai', description: '', priority: 'low', due_date: '2026-04-20', duration_minutes: 330, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Membership Clerk / Inventory Auditor at Costco', description: '', priority: 'low', due_date: '2026-04-18', duration_minutes: 105, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Retail Roles at Walmart', description: '', priority: 'low', due_date: '2026-04-18', duration_minutes: 30, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Program Manager IV at Centene', description: '', priority: 'low', due_date: '2026-04-17', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Consultant / Dir of Product Mgmt at Evalueserve / RELX', description: '', priority: 'low', due_date: '2026-04-16', duration_minutes: 330, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Product Manager 2 / Sr Project Manager at LexisNexis', description: '', priority: 'low', due_date: '2026-04-15', duration_minutes: 330, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Retail Roles at Walmart / Lowes', description: '', priority: 'low', due_date: '2026-04-13', duration_minutes: 105, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at KPMG / Cisco', description: '', priority: 'low', due_date: '2026-04-10', duration_minutes: 210, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Director of Product Management at RELX', description: '', priority: 'low', due_date: '2026-04-10', duration_minutes: 150, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to AI Initiatives Program Manager at Nutanix', description: '', priority: 'low', due_date: '2026-04-07', duration_minutes: 120, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Product Delivery Manager at JPMorgan Chase', description: '', priority: 'low', due_date: '2026-04-07', duration_minutes: 120, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Program Manager at Microsoft', description: '', priority: 'low', due_date: '2026-04-02', duration_minutes: 270, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at Cox / Eliassen Group / Randstad', description: '', priority: 'low', due_date: '2026-04-01', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to AI and Agentic Transformation Consultant at Cox', description: '', priority: 'low', due_date: '2026-04-01', duration_minutes: 15, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Senior Product Manager Data Platforms at Inmar', description: '', priority: 'low', due_date: '2026-03-31', duration_minutes: 240, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Senior Project Manager at CACI', description: '', priority: 'low', due_date: '2026-03-31', duration_minutes: 120, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Product Manager at Circle', description: '', priority: 'low', due_date: '2026-03-30', duration_minutes: 60, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Product Manager AI Platform at Circle', description: '', priority: 'low', due_date: '2026-03-27', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Product Manager at Chartis', description: '', priority: 'low', due_date: '2026-03-27', duration_minutes: 105, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Product Manager at Microsoft', description: '', priority: 'low', due_date: '2026-03-26', duration_minutes: 150, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Portfolio/Program Manager at Chartis', description: '', priority: 'low', due_date: '2026-03-25', duration_minutes: 45, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Program Manager at Microsoft (Mar 22)', description: '', priority: 'low', due_date: '2026-03-22', duration_minutes: 285, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Principal Product Manager / TPM at Microsoft', description: '', priority: 'low', due_date: '2026-03-20', duration_minutes: 225, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at Deloitte / Microsoft / Apple', description: '', priority: 'low', due_date: '2026-03-20', duration_minutes: 420, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Product/Program Management Roles at Deloitte', description: '', priority: 'low', due_date: '2026-03-18', duration_minutes: 330, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Roles at Deloitte / Granicus', description: '', priority: 'low', due_date: '2026-03-09', duration_minutes: 180, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Director - Strategic Programs at Granicus', description: '', priority: 'low', due_date: '2026-03-09', duration_minutes: 120, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Associate Director Roles at KPMG', description: '', priority: 'low', due_date: '2026-03-05', duration_minutes: 45, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Product Manager Roles at Lenovo', description: '', priority: 'low', due_date: '2026-03-04', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Strategy and Execution Supply Chain at EY-Parthenon', description: '', priority: 'low', due_date: '2026-03-04', duration_minutes: 45, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Various Program Manager Roles at Oracle', description: '', priority: 'low', due_date: '2026-03-04', duration_minutes: 30, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Technical Dir / Sr Product Manager at BairesDev / Avalara', description: '', priority: 'low', due_date: '2026-03-04', duration_minutes: 90, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Business Technology Product Manager at Samsara', description: '', priority: 'low', due_date: '2026-02-24', duration_minutes: 540, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Product Manager Edge Technologies at Lenovo', description: '', priority: 'low', due_date: '2026-02-10', duration_minutes: 150, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to AI Technical Program Manager at Deloitte', description: '', priority: 'low', due_date: '2026-01-30', duration_minutes: 105, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Director of Product Management at OnLogic', description: '', priority: 'low', due_date: '2026-01-14', duration_minutes: 255, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Global eCommerce SMB Customer PM at Lenovo', description: '', priority: 'low', due_date: '2026-01-05', duration_minutes: 135, tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Sam\'s Club HSM position', description: '', priority: 'low', tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply to Deloitte Technology Strategy Manager 321001', description: 'https://apply.deloitte.com/en_US/careers/InviteToApply?jobId=321001', priority: 'low', due_date: '2026-03-21', tags: [], completed: true },
    { list: 'career_job_search', name: 'H2O.ai TPM — Job research and prep', description: 'Lane 1 — AI/Data Platform. Drive end-to-end AI/ML program execution (data pipelines → model deployment → production scaling).', priority: 'low', tags: [], completed: true },
    { list: 'career_job_search', name: 'Principal Product Manager - AI Travel at Hopper', description: 'https://jobs.ashbyhq.com/hopper/10e7bddf-60cb-4d04-90fb-fdd350af0314', priority: 'low', tags: [], completed: true },
    { list: 'career_job_search', name: 'Group Product Manager (Platform) at Ontra', description: 'https://www.ontra.ai/jobs/?gh_jid=8422083002', priority: 'low', tags: [], completed: true },
    { list: 'career_job_search', name: 'Senior Product Manager Desktop Platform at Docker', description: 'https://www.linkedin.com/jobs/view/4305146890', priority: 'low', tags: [], completed: true },
    { list: 'career_job_search', name: 'Apply for 1 job (daily)', description: 'Daily recurring job application.', priority: 'low', due_date: '2026-06-02', duration_minutes: 20, tags: [], completed: true },

    // 1.2 Interviews
    { list: 'career_job_search', name: 'BuildOps PM Interview', description: 'Staff Product Manager interview with BuildOps.', priority: 'low', due_date: '2026-06-02', duration_minutes: 21, tags: [], completed: true },
    { list: 'career_job_search', name: 'Samsara PM Interviews', description: 'Prepared for and completed Samsara Product Manager interviews.', priority: 'low', due_date: '2026-02-27', duration_minutes: 360, tags: [], completed: true },
    { list: 'career_job_search', name: 'OnLogic PM Interview', description: 'Director of Product Management interview at OnLogic.', priority: 'low', due_date: '2026-01-21', duration_minutes: 75, tags: [], completed: true },
    { list: 'career_job_search', name: 'Prep for H2O.ai Interview', description: '', priority: 'low', due_date: '2026-04-20', tags: [], completed: true },
    { list: 'career_job_search', name: 'Prepare for Thursday interview (Inmar Intelligence)', description: '', priority: 'low', due_date: '2026-04-15', tags: [], completed: true },
    { list: 'career_job_search', name: 'Prepare for call with Laird (Inmar Intelligence)', description: '', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },

    // 1.1 Deep Work — completed resume work
    { list: 'career_deep_work', name: 'Resume enhancement — BuildOps tailoring', description: 'Enhanced resumes and organized materials in Careerflow/AI tools while tailoring applications for BuildOps.', priority: 'low', due_date: '2026-05-28', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'ResumeOS v2.0 Executive Audit Report', description: 'Developed ResumeOS v2.0 Executive Resume Audit Report and outlined next resume improvement actions.', priority: 'low', due_date: '2026-05-26', duration_minutes: 60, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'RACI-aligned resume content — AI review pass', description: 'Refined RACI-aligned resume content in Careerflow with AI support.', priority: 'low', due_date: '2026-05-26', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'RACI-aligned resume revisions', description: 'Conducted deep RACI-aligned resume revisions across Careerflow/docs.', priority: 'low', due_date: '2026-05-26', duration_minutes: 150, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director resume structure & impact language', description: 'Director-level resume refinement in Careerflow, tightening structure and impact language.', priority: 'low', due_date: '2026-05-25', duration_minutes: 330, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director Resume OS audit — TPM/data governance', description: 'Director-level Resume OS in Careerflow, auditing consistency and aligning bullets with TPM/data governance.', priority: 'low', due_date: '2026-05-25', duration_minutes: 390, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Senior/director framing & export-ready variants', description: 'Revised Careerflow resumes with AI support, strengthening senior/director-level framing.', priority: 'low', due_date: '2026-05-24', duration_minutes: 509, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director resume revision — metrics & materials', description: 'Extensively revised director-level resume and materials in Careerflow.', priority: 'low', due_date: '2026-05-23', duration_minutes: 449, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director resume leadership bullets — ResumeOS', description: 'Refined director-level resume content and leadership impact bullets.', priority: 'low', due_date: '2026-05-22', duration_minutes: 30, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director/TPM resume & ROI narratives', description: 'Deepened director/TPM resume and ADAS/OnePipeline portfolio; aligned metrics and ROI narratives.', priority: 'low', due_date: '2026-05-22', duration_minutes: 269, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & OnePipeline portfolio overhaul — Cisco prep', description: 'Overhauled resume and portfolio (OnePipeline, TPM focus); modeled ROI and prepared for Cisco TPM HR interview.', priority: 'low', due_date: '2026-05-22', duration_minutes: 480, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'TPM/Data Director resume variants — MaintainX/PwC', description: 'Tailored TPM/Data Director resume variants in Careerflow for MaintainX/PwC roles.', priority: 'low', due_date: '2026-05-21', duration_minutes: 299, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Director-level resume refinement sprint', description: 'Refined Director-level resumes in Careerflow with AI feedback, validated PDFs.', priority: 'low', due_date: '2026-05-21', duration_minutes: 570, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume structure & targeting — AI-assisted', description: 'Advanced resume development in Careerflow with AI support, refining structure.', priority: 'low', due_date: '2026-05-20', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume variants — AI achievement extraction', description: 'Refined resumes in Careerflow using AI to extract achievements and tailor variants.', priority: 'low', due_date: '2026-05-20', duration_minutes: 120, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'LinkedIn PM eval & ResumeOS update', description: 'Evaluated LinkedIn PM opportunities and updated Careerflow/ResumeOS.', priority: 'low', due_date: '2026-05-20', duration_minutes: 60, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'AI-assisted resume variants — CGI Sr PM', description: 'Researched CGI Senior PM role and refined AI-assisted resume variants.', priority: 'low', due_date: '2026-05-18', duration_minutes: 240, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume revision — Word & Careerflow', description: 'Revised resume across Word and Careerflow, refining content with ChatGPT.', priority: 'low', due_date: '2026-05-17', duration_minutes: 60, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'ResumeOS content update & tracker', description: 'Updated Careerflow tracker and ResumeOS content, refining resume language.', priority: 'low', due_date: '2026-05-16', duration_minutes: 150, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & portfolio bullets — extended draft', description: 'Extended resume and portfolio drafting, optimized bullets and narratives in Careerflow.', priority: 'low', due_date: '2026-05-15', duration_minutes: 210, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'ResumeOS full refinement sprint', description: 'Refined ResumeOS resumes and portfolio content, organized career materials.', priority: 'low', due_date: '2026-05-15', duration_minutes: 510, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume iteration — Careerflow & ChatGPT', description: 'Iterated resumes and Job Tracker in Careerflow and ChatGPT.', priority: 'low', due_date: '2026-05-14', duration_minutes: 60, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume OS strategy — ChatGPT planning', description: 'Planned Resume OS strategy in ChatGPT, outlining priorities and approach.', priority: 'low', due_date: '2026-05-14', duration_minutes: 30, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & Aspida Sr. TPM application', description: 'Refined resume and Careerflow profile and completed Aspida Sr. TPM application.', priority: 'low', due_date: '2026-05-12', duration_minutes: 180, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume refinement — Careerflow & ChatGPT', description: 'Extensively refined resume and applications in Careerflow and ChatGPT.', priority: 'low', due_date: '2026-05-11', duration_minutes: 180, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume OS strategy planning', description: 'Refined job-search strategy with Resume OS planning.', priority: 'low', due_date: '2026-05-08', duration_minutes: 30, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Intensive Careerflow job search sprint', description: 'Conducted intensive job search with Careerflow—updating resume, profiles, tracker.', priority: 'low', due_date: '2026-05-07', duration_minutes: 180, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume variants & LinkedIn roles', description: 'Updated Careerflow tracker, refined resume variants, and organized targeted LinkedIn roles.', priority: 'low', due_date: '2026-05-06', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Careerflow tracker & resume tailoring', description: 'Advanced targeted job search: updated Careerflow tracker, tailored resume.', priority: 'low', due_date: '2026-05-06', duration_minutes: 180, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Careerflow targeted resume updates', description: 'Advanced job search in Careerflow Job Tracker/Resume Builder with ChatGPT-assisted updates.', priority: 'low', due_date: '2026-05-05', duration_minutes: 30, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Career outreach & Careerflow positioning', description: 'Managed career outreach, updated Careerflow and resume, refined positioning.', priority: 'low', due_date: '2026-04-28', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & LinkedIn realignment', description: 'Revised resume and LinkedIn, aligned materials to target roles.', priority: 'low', due_date: '2026-04-28', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & application materials org', description: 'Refined Careerflow resume and organized application materials.', priority: 'low', due_date: '2026-04-17', duration_minutes: 89, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Job tracker & interview narratives', description: 'Organized Careerflow job tracker, refined applications, and drafted interview narratives.', priority: 'low', due_date: '2026-04-14', duration_minutes: 105, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & interview prep — Careerflow', description: 'Refined resume and Careerflow tracker, managed interview communications.', priority: 'low', due_date: '2026-04-14', duration_minutes: 75, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume research & Careerflow update', description: 'Researched target roles and significantly updated resume in Careerflow.', priority: 'low', due_date: '2026-04-13', duration_minutes: 60, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume tailoring — Deloitte', description: 'Developed Deloitte-targeted resume, updated Careerflow tracker, prepared PDF/Word versions.', priority: 'low', due_date: '2026-04-10', duration_minutes: 75, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume tailoring — KPMG & Cisco', description: 'Tailored resumes and tracked roles in Careerflow for KPMG and Cisco positions.', priority: 'low', due_date: '2026-04-10', duration_minutes: 210, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Base resume drafting — ADAS/infotainment bullets', description: 'Intensive base resume drafting using AI and Careerflow, refining ADAS/infotainment impact bullets.', priority: 'low', due_date: '2026-04-09', duration_minutes: 209, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Base resume expansion — AI-assisted', description: 'Expanded and organized base resume and role summaries, integrating AI-assisted edits.', priority: 'low', due_date: '2026-04-09', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & LinkedIn refinement — PM/TPM', description: 'Refined resume and LinkedIn for target PM/TPM roles, aligning achievements and keywords.', priority: 'low', due_date: '2026-04-08', duration_minutes: 165, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume restructure — JPMC role', description: 'Advanced job search by updating tracker and extensively restructuring resume for JPMC role.', priority: 'low', due_date: '2026-04-07', duration_minutes: 120, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume tailoring — Microsoft applications', description: 'Tailored resume and materials for Microsoft applications and updated Careerflow tracker.', priority: 'low', due_date: '2026-04-06', duration_minutes: 105, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & Job Tracker maintenance', description: 'Refined Careerflow resume, maintained Job Tracker, and documented outreach.', priority: 'low', due_date: '2026-04-06', duration_minutes: 195, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume overhaul & STAR responses', description: 'Overhauled resume and STAR responses, and organized Careerflow applications pipeline.', priority: 'low', due_date: '2026-04-06', duration_minutes: 330, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume bullets & layout refinement', description: 'Refined Careerflow resume bullets and layout, improving impact and clarity.', priority: 'low', due_date: '2026-04-06', duration_minutes: 45, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Careerflow resume restructure & bullets', description: 'Extensively revised Careerflow resume, restructuring sections and refining achievement bullets.', priority: 'low', due_date: '2026-04-05', duration_minutes: 119, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & OnePipeline tailoring — Data PM', description: 'Iteratively tailored resume and portfolio (OnePipeline) to Data PM roles.', priority: 'low', due_date: '2026-04-04', duration_minutes: 269, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'OnePipeline portfolio & PM methods draft', description: 'Drafted and refined OnePipeline portfolio and PM methods with Gemini research.', priority: 'low', due_date: '2026-04-04', duration_minutes: 135, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'STAR narratives & resume/portfolio refinement', description: 'Developed STAR interview narratives and refined resume and portfolio.', priority: 'low', due_date: '2026-04-04', duration_minutes: 120, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'STAR responses & work history overhaul', description: 'Developed STAR interview responses and overhauled work history/projects.', priority: 'low', due_date: '2026-04-03', duration_minutes: 404, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Job search & resume — Microsoft PM', description: 'Conducted extensive job search, refined resume, and advanced Microsoft PM application.', priority: 'low', due_date: '2026-04-02', duration_minutes: 270, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume update — CareerFlow org', description: 'Organized job search activities and updated resume in CareerFlow.', priority: 'low', due_date: '2026-04-02', duration_minutes: 30, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'CareerFlow board & resume update', description: 'Updated CareerFlow board and resume while reviewing Cox, Eliassen, and Randstad roles.', priority: 'low', due_date: '2026-04-01', duration_minutes: 90, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume rewrite — CareerFlow & NCWorks', description: 'Rewrote and optimized resumes in CareerFlow and NCWorks.', priority: 'low', due_date: '2026-04-01', duration_minutes: 195, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Resume & STAR Prep', description: 'Intensively refined resumes, tracked applications, and prepared STAR-based interview responses.', priority: 'low', due_date: '2026-04-03', duration_minutes: 525, tags: ['resume-updates'], completed: true },
    { list: 'career_deep_work', name: 'Build Lane 1 resume — OnePipeline Platform PM', description: 'Lane 1 (OnePipeline Platform PM): data pipeline, ML/AI platform, APIs, SLAs, observability.', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Build Lane 2 resume — Salesforce Q2C PM', description: 'Lane 2 (Salesforce Q2C PM): CRM/CPQ/CLM, quote-to-cash, entitlements, billing, workflow automation.', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Build Lane 3 resume — Program Manager', description: 'Lane 3 (Program Manager): multi-year delivery, acceptance criteria, exec comms, cross-org dependencies.', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Build Lane 4 resume — Transformation', description: 'Lane 4 (Transformation): operating model, org design, process change, KPI systems, GTM packaging.', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Extract past STAR interview notes', description: 'Extract past STAR interview notes and create a database in Notion.', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Create a detailed project summary', description: 'Create a detailed summary of all projects worked in the past.', priority: 'low', due_date: '2026-04-04', tags: [], completed: true },
    { list: 'career_deep_work', name: 'Cleanup Career and Professional folder in ChatGPT', description: '', priority: 'low', due_date: '2026-04-03', tags: [], completed: true },

    // 1.3 Communication — completed
    { list: 'career_communication', name: 'Sent recruiter message about application (Julie Hall)', description: 'https://www.linkedin.com/in/juliehall0', priority: 'low', tags: [], completed: true },
    { list: 'career_communication', name: 'Reach out to Casey Sortini at Rayonic', description: '', priority: 'low', tags: [], completed: true },
    { list: 'career_communication', name: 'Connect with Danielle Knowles at Deloitte', description: 'https://www.linkedin.com/in/daknowlez', priority: 'low', tags: [], completed: true },
    { list: 'career_communication', name: 'Connect with Austin Westbrook at Deloitte', description: 'https://www.linkedin.com/in/austin-westbrook-998640141', priority: 'low', tags: [], completed: true },
    { list: 'career_communication', name: 'Connect with Agatha at McKinsey', description: 'https://www.linkedin.com/in/agatamirowska', priority: 'low', tags: [], completed: true },

    // ================================================================
    // 2.0 LEARNING — COMPLETED TASKS
    // ================================================================

    { list: 'learning_courses', name: 'Platform Roadmap Vision module', description: 'Developed content for Platform Ownership & Roadmap Vision module in Notion.', priority: 'low', due_date: '2026-06-08', duration_minutes: 71, tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Complete Deploy Workloads with Lakeflow Jobs', description: 'https://customer-academy.databricks.com/learn/learning-plans/10/', priority: 'low', due_date: '2026-06-01', duration_minutes: 60, tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Databricks Study — Intelligent Search & Lakeflow', description: 'Studied Databricks intelligent search and Lakeflow, extending notebook work with notes and Q&A.', priority: 'low', due_date: '2026-05-26', duration_minutes: 60, tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Databricks Certification — Fundamentals modules', description: 'Completed Databricks Fundamentals modules and knowledge checks, began Lakeflow data engineering work.', priority: 'low', due_date: '2026-05-26', duration_minutes: 180, tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Game Theory & Separation Notes', description: 'Combined yard work with C565/Game Theory signaling games study and personal separation notes.', priority: 'low', due_date: '2026-04-20', duration_minutes: 300, tags: ['mba'], completed: true },
    { list: 'learning_courses', name: 'Signaling Games Study', description: 'Studied C565/Game Theory Week 07 signaling games and drafted learning plan.', priority: 'low', due_date: '2026-04-20', duration_minutes: 90, tags: ['mba'], completed: true },
    { list: 'learning_courses', name: 'Game Theory Excel Modeling', description: 'Extended Game Theory study, building an Excel payoff-matrix analyzer and drafting solutions.', priority: 'low', due_date: '2026-04-05', duration_minutes: 555, tags: ['mba'], completed: true },
    { list: 'learning_courses', name: 'Game Theory & Pipeline Diagrams', description: 'Studied Game Theory and designed Data Pipeline and Career WBS diagrams.', priority: 'low', due_date: '2026-03-31', duration_minutes: 240, tags: ['mba'], completed: true },
    { list: 'learning_courses', name: 'AI Chatbot Coursework', description: 'Completed AI chatbot coursework quizzes, IBM Cloud setup, and labs.', priority: 'low', due_date: '2026-03-12', duration_minutes: 285, tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Game Theory Study', description: 'Completed focused Game Theory study in Canvas with practice exercises.', priority: 'low', due_date: '2026-03-09', duration_minutes: 90, tags: ['mba'], completed: true },
    { list: 'learning_courses', name: 'Module 5: Final Quiz and Wrap-up (Databricks)', description: '', priority: 'low', tags: ['learning'], completed: true },
    { list: 'learning_courses', name: 'Module 4: Product Plan Phase (Coursera)', description: 'https://www.coursera.org/learn/product-management-initial-product-strategy-and-plan/home/module/4', priority: 'low', tags: ['learning', 'pm-cert'], completed: true },
    { list: 'learning_courses', name: 'Module 3: Product Concept Investigation (Coursera)', description: 'https://www.coursera.org/learn/product-management-initial-product-strategy-and-plan/home/module/3', priority: 'low', tags: ['learning', 'pm-cert'], completed: true },
    { list: 'learning_courses', name: 'Module 2: Product Concept Identification (Coursera)', description: 'https://www.coursera.org/learn/product-management-initial-product-strategy-and-plan/home/module/2', priority: 'low', tags: ['learning', 'pm-cert'], completed: true },
    { list: 'learning_courses', name: 'Post certificate on LinkedIn', description: '', priority: 'low', tags: [], completed: true },
    { list: 'learning_courses', name: 'Prep for G565 Quiz 2', description: '', priority: 'low', due_date: '2026-04-07', tags: ['mba'], completed: true },

    { list: 'learning_research', name: 'SOA & AI TTS Research', description: 'Explored AI text-to-speech platforms (NaturalReaders, Speechify) while researching SOA concepts.', priority: 'low', due_date: '2026-06-09', duration_minutes: 22, tags: ['learning'], completed: true },
    { list: 'learning_research', name: 'API Design Research', description: 'Researched API design best practices using ChatGPT and Claude, focusing on The Design of Web APIs.', priority: 'low', due_date: '2026-06-01', duration_minutes: 34, tags: ['learning'], completed: true },

    // ================================================================
    // 3.0 PERSONAL & LIFE — COMPLETED TASKS
    // ================================================================

    // 3.1 Health
    { list: 'personal_health', name: 'CC Mediation Class', description: 'Attended court-mandated CC Mediation Class and organized related notes and instructions.', priority: 'low', due_date: '2026-05-20', duration_minutes: 180, tags: [], completed: true },
    { list: 'personal_health', name: 'Call to schedule sleep study', description: '', priority: 'low', due_date: '2026-04-15', tags: [], completed: true },
    { list: 'personal_health', name: 'Call to schedule MRI', description: '', priority: 'low', due_date: '2026-04-15', tags: [], completed: true },
    { list: 'personal_health', name: 'Call Walgreens to order Azstarys', description: '', priority: 'low', due_date: '2026-04-17', tags: [], completed: true },
    { list: 'personal_health', name: 'Schedule Nia PCP this week', description: '', priority: 'low', due_date: '2026-04-20', tags: [], completed: true },
    { list: 'personal_health', name: 'Make Nia Dermatology Appt', description: 'Call BlueRidge to make Nia Dermatology Appt.', priority: 'low', due_date: '2026-04-22', tags: [], completed: true },
    { list: 'personal_health', name: 'Call Nia Orthopaedic', description: '', priority: 'low', due_date: '2026-04-20', tags: [], completed: true },
    { list: 'personal_health', name: 'Ask Dr. Madison for diagnosis and affidavit', description: '1. Ask Dr. Madison for formal diagnosis on file (esp. BPD/Lamotrigine). 2. Affidavit for treatment observations.', priority: 'low', due_date: '2026-04-20', tags: ['legal'], completed: true },
    { list: 'personal_health', name: 'Call to schedule for sleep study (initial)', description: '', priority: 'low', due_date: '2026-04-15', tags: [], completed: true },
    { list: 'personal_health', name: 'Confirm Caleb Cavity Filling Appointment', description: 'Call Wells dental and confirm May appointment.', priority: 'low', due_date: '2026-05-12', tags: [], completed: true },
    { list: 'personal_health', name: 'Schedule Caleb and Elijah Cavity', description: 'Wells Family Dentistry.', priority: 'low', due_date: '2026-05-18', tags: [], completed: true },
    { list: 'personal_health', name: 'Schedule Nia Therapy with Kathy Caputo', description: 'At Chronic Hope Counseling.', priority: 'low', due_date: '2026-05-19', tags: [], completed: true },
    { list: 'personal_health', name: 'Personal Admin & Scheduling', description: 'Managed unpaid bills and insurance, scheduled dental and family appointments, organized separation notes.', priority: 'low', due_date: '2026-02-15', duration_minutes: 450, tags: [], completed: true },
    { list: 'personal_health', name: 'Therapy & DVPO Documentation', description: 'Attended therapy session, then documented DVPO-related conversations, photos, and incidents for counsel.', priority: 'low', due_date: '2026-04-27', duration_minutes: 270, tags: ['legal'], completed: true },

    // 3.2 Home & Family
    { list: 'personal_family', name: 'Submit referral for substitute teacher', description: '', priority: 'low', due_date: '2026-05-19', tags: [], completed: true },
    { list: 'personal_family', name: 'Elijah: LRHS Band auditions video', description: '', priority: 'low', due_date: '2026-04-08', tags: [], completed: true },
    { list: 'personal_family', name: 'Find gifts for Ami teacher', description: '', priority: 'low', due_date: '2026-04-25', tags: [], completed: true },
    { list: 'personal_family', name: 'Plan beach trip for 1st May weekend', description: '', priority: 'low', due_date: '2026-04-19', tags: [], completed: true },
    { list: 'personal_family', name: 'Drop off Sleep Supply Return at UPS', description: '877-265-2426', priority: 'low', due_date: '2026-04-20', tags: [], completed: true },
    { list: 'personal_family', name: 'Ami take water bottle', description: '', priority: 'low', due_date: '2026-04-30', tags: [], completed: true },
    { list: 'personal_family', name: 'Nia Unit 7 Takehome Quiz remind', description: '', priority: 'low', due_date: '2026-04-27', tags: [], completed: true },
    { list: 'personal_family', name: 'Now lawn (April)', description: '', priority: 'low', tags: [], completed: true },
    { list: 'personal_family', name: 'Call Apria for return', description: '', priority: 'low', due_date: '2026-04-07', tags: [], completed: true },
    { list: 'personal_family', name: 'Child Placement Docs — WCPSS', description: 'Managed WCPSS child placement transition and documentation, coordinated logistics and school admin.', priority: 'low', due_date: '2026-03-04', duration_minutes: 210, tags: ['legal'], completed: true },

    // 3.3 Finance & Legal
    { list: 'personal_finance_legal', name: 'Settlement Agreement Finalization', description: 'Drafted and finalized rebuttal and settlement agreement, organized legal evidence, submitted signed notarized documents.', priority: 'low', due_date: '2026-05-12', duration_minutes: 270, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Legal Rebuttal Expansion', description: 'Expanded and polished legal rebuttal with organized evidence and G565 finals review.', priority: 'low', due_date: '2026-05-10', duration_minutes: 360, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'CPS Meeting & Evidence', description: 'Documented personal records and evidence in NotebookLM/Notion and held extended CPS meeting.', priority: 'low', due_date: '2026-05-04', duration_minutes: 390, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Mirian Complaint Breakdown', description: 'Expanded and normalized the Mirian Complaint Breakdown sheet, aligning incidents, dates, categories, and citations.', priority: 'low', due_date: '2026-05-03', duration_minutes: 120, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Custody Evidence Prep', description: 'Developed and organized separation/custody evidence, building Notion notes, NotebookLM files, and indexed Google Sheets tracker.', priority: 'low', due_date: '2026-04-25', duration_minutes: 420, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Legal Agreement Revision', description: 'Reviewed and revised personal Separation Agreement with ChatGPT-supported legal phrasing.', priority: 'low', due_date: '2026-03-06', duration_minutes: 360, tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Finance Reconciliation', description: 'Reconciled personal finances in Monarch, tuned rules and extensions, cross-checked subscriptions and receipts.', priority: 'low', due_date: '2026-02-22', duration_minutes: 180, tags: [], completed: true },
    { list: 'personal_finance_legal', name: 'Send modified No-Contact agreement to Mirian', description: '', priority: 'low', due_date: '2026-06-02', tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Send the notarized form to Monica', description: '', priority: 'low', due_date: '2026-06-01', tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Pull text exchange with kids', description: 'Find text messages that prove healthy parenting.', priority: 'low', due_date: '2026-04-16', tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Compile dating app records', description: 'Compile records of dating app activity and children inquiries circa 2020–2022.', priority: 'normal', tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Complete documenting Kathy Caputo incident for 2025', description: 'https://www.notion.so/Could-have-stayed-With-Kathy-Caputo-3626187008a280f5b96becb5541d03e2', priority: 'low', due_date: '2026-05-16', tags: ['legal'], completed: true },
    { list: 'personal_finance_legal', name: 'Update payment information for BCBSNC', description: '', priority: 'low', due_date: '2026-04-06', tags: [], completed: true },
    { list: 'personal_finance_legal', name: 'Update Morgan Stanley mortgage payment info', description: '', priority: 'low', tags: [], completed: true },
    { list: 'personal_finance_legal', name: 'Additional work towards Separation', description: '', priority: 'low', due_date: '2025-08-08', tags: ['legal', '💔 Separation'], completed: true },
    { list: 'personal_finance_legal', name: 'Upload Latest Financial Statement to Separation share drive', description: 'Upload to MS OneDrive Separation shared folder.', priority: 'low', tags: ['legal', '💔 Separation'], completed: true },

    // 3.4 Rest & Entertainment
    { list: 'personal_rest', name: 'Country Guitar free event — Emma Zinck Hits The Roof', description: 'https://www.visitraleigh.com/event/emma-zinck-hits-the-roof/105203/', priority: 'low', tags: [], completed: true },
    { list: 'personal_rest', name: 'Reply to Meg on Bumble', description: '', priority: 'low', due_date: '2026-04-17', tags: [], completed: true },

    // ================================================================
    // 4.0 DEVICE & TINKERING — COMPLETED TASKS
    // ================================================================

    { list: 'device_tinkering', name: 'GitHub Passwordless Auth Setup', description: 'Managed GitHub account security settings and passwordless authentication while testing local AI configurations in LobeHub.', priority: 'low', due_date: '2026-06-05', duration_minutes: 22, tags: [], completed: true },
    { list: 'device_tinkering', name: 'Open-Source AI Frameworks Eval', description: 'Researched and compared open-source AI chat frameworks (Lobe Chat, LibreChat) to evaluate features and deployment options.', priority: 'low', due_date: '2026-06-04', duration_minutes: 26, tags: ['santhoshos'], completed: true },
    { list: 'device_tinkering', name: 'Mac & Cloud Housekeeping', description: 'Mac housekeeping—remote access setup, cloud storage cleanup, Drive sync config, and system preference checks.', priority: 'low', due_date: '2026-05-16', duration_minutes: 179, tags: [], completed: true },
    { list: 'device_tinkering', name: 'iMazing Backup Config', description: 'Installed and troubleshot iMazing/AnyTrans iPhone backups, tuning storage/configuration.', priority: 'low', due_date: '2026-04-25', duration_minutes: 120, tags: [], completed: true },
    { list: 'device_tinkering', name: 'Workspace Policy Config', description: 'Configured Workspace admin, Gemini, Gmail, and Vault policies to align storage, access, and retention.', priority: 'low', due_date: '2026-04-09', duration_minutes: 90, tags: [], completed: true },
    { list: 'device_tinkering', name: 'Browser Workspace Setup', description: 'Configured and evaluated browsers and tab/workspace extensions across Edge, Chrome, Vivaldi, and Firefox.', priority: 'low', due_date: '2026-03-18', duration_minutes: 194, tags: [], completed: true },
    { list: 'device_tinkering', name: 'Parental Controls Setup', description: 'Configured Bark and Qustodio parental controls across devices, tuning monitoring, Screen Time, and permissions.', priority: 'low', due_date: '2026-02-12', duration_minutes: 120, tags: [], completed: true },

    // ================================================================
    // 6.0 STRAVENTIS — COMPLETED TASKS
    // ================================================================

    // 6.1 Client Work
    { list: 'straventis_client', name: 'Data Governance Synthesis', description: 'Synthesized technical documentation on data governance, IAM security, and pipeline architecture.', priority: 'low', due_date: '2026-06-09', duration_minutes: 63, tags: ['xperi'], completed: true },
    { list: 'straventis_client', name: 'Video Metadata Research', description: 'Conducted research on video metadata management principles and industry standards for Xperi data platform planning.', priority: 'low', due_date: '2026-06-08', duration_minutes: 38, tags: ['xperi'], completed: true },
    { list: 'straventis_client', name: 'Manufacturing Escalations Doc', description: 'Documented and organized manufacturing escalation procedures for CPM Ford project in Notion.', priority: 'low', due_date: '2026-06-07', duration_minutes: 24, tags: ['harman'], completed: true },
    { list: 'straventis_client', name: 'Harman & Zoox Documentation', description: 'Updated and organized project documentation for Harman and Zoox in Notion.', priority: 'low', due_date: '2026-06-06', duration_minutes: 125, tags: ['harman'], completed: true },
    { list: 'straventis_client', name: 'Zoox Documentation', description: 'Refined and organized project details, company information, and assets for Zoox in Notion.', priority: 'low', due_date: '2026-06-06', duration_minutes: 60, tags: [], completed: true },
    { list: 'straventis_client', name: 'DFD Offsite Go-Live', description: 'Led IE/DFD standup and offsite final planning; aligned agenda, risks, logistics, and documented go-live decisions.', priority: 'low', due_date: '2026-02-19', duration_minutes: 435, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'Team-Building Hike', description: 'Participated in onsite team-building hike to strengthen alignment for ParticleBlack/Ascension.', priority: 'low', due_date: '2026-02-17', duration_minutes: 120, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'DFD Standup & SoW Drafting', description: 'Prepared for and attended DFD standup, then drafted and refined SoW-2026PB001 in Google Docs.', priority: 'low', due_date: '2026-02-13', duration_minutes: 90, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'ParticleBlack Contract Call', description: 'Prepared and held ParticleBlack scope/rate call, updating Playbook and follow-ups.', priority: 'low', due_date: '2026-02-11', duration_minutes: 89, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'Oracle HCM/API Workshop', description: 'Led DFD standup and Oracle HCM/API workshops, aligning integration flows, endpoints, and next steps.', priority: 'low', due_date: '2026-02-11', duration_minutes: 210, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'Straventis Workspace Admin', description: 'Managed Google Workspace billing and email migration, reconciled Ascension hours, updated invoicing.', priority: 'low', due_date: '2026-02-05', duration_minutes: 150, tags: [], completed: true },
    { list: 'straventis_client', name: 'Ascension SOW Drafting', description: 'Drafted and refined Ascension SOW, created supporting docs, updated Notion, and organized project files.', priority: 'low', due_date: '2026-02-04', duration_minutes: 89, tags: ['ascension'], completed: true },
    { list: 'straventis_client', name: 'DFD Stand-up & Alignment', description: 'Participated in DFD stand-up with Ascension team to review progress, align priorities, and confirm owners.', priority: 'low', due_date: '2026-01-20', duration_minutes: 30, tags: ['ascension'], completed: true },

    // 6.2 Coding & Scripting
    { list: 'straventis_coding', name: 'Notion API Automation', description: 'Developed Google Apps Script to automate Notion journal reorganization project using Claude for coding.', priority: 'low', due_date: '2026-06-06', duration_minutes: 24, tags: ['notion', 'automation'], completed: true },
    { list: 'straventis_coding', name: 'NotionMoveV4 Script', description: 'Developed and refined Google Apps Script for NotionMoveV4 project, automating data migration between Notion and Google Sheets.', priority: 'low', due_date: '2026-06-06', duration_minutes: 39, tags: ['notion', 'automation'], completed: true },
    { list: 'straventis_coding', name: 'Automation Triggers & Version Control', description: 'Developed automation triggers for Gmail Label Classifier script utilizing Claude and GitHub Desktop for version control.', priority: 'low', due_date: '2026-06-04', duration_minutes: 68, tags: ['automation'], completed: true },
    { list: 'straventis_coding', name: 'Docker LobeHub Deployment', description: 'Configured local settings and Docker environments for LobeHub while reviewing documentation.', priority: 'low', due_date: '2026-06-04', duration_minutes: 38, tags: ['santhoshos'], completed: true },
    { list: 'straventis_coding', name: 'LabelClassifier Apps Script', description: 'Developed and refined Gmail_Thiru_LabelClassifier Google Apps Script to automate email organization.', priority: 'low', due_date: '2026-06-04', duration_minutes: 71, tags: ['automation'], completed: true },
    { list: 'straventis_coding', name: 'GitHub Repo Setup', description: 'Initialized the ai-hub repository and pushed initial codebase to GitHub.', priority: 'low', due_date: '2026-06-03', duration_minutes: 28, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Notion Duplicate Script', description: 'Developed Google Apps Script to identify duplicate entries within Notion.', priority: 'low', due_date: '2026-06-03', duration_minutes: 66, tags: ['notion', 'automation'], completed: true },
    { list: 'straventis_coding', name: 'React Dashboard UI', description: 'Developed Straventis dashboard, focusing on UI implementation and multi-AI features using Horizon UI template.', priority: 'low', due_date: '2026-06-03', duration_minutes: 45, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Next.js App Development', description: 'Worked on Straventis Next.js application, utilizing ChatGPT for technical troubleshooting.', priority: 'low', due_date: '2026-06-03', duration_minutes: 48, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Notion Title Queue Script', description: 'Leveraged Claude and Google AI Studio to develop and debug Google Apps Script for Notion Title Queue project.', priority: 'low', due_date: '2026-06-02', duration_minutes: 29, tags: ['notion', 'automation'], completed: true },
    { list: 'straventis_coding', name: 'Notion Queue Script', description: 'Developed and refined the Notion Title Queue script in Google Apps Script.', priority: 'low', due_date: '2026-06-02', duration_minutes: 21, tags: ['notion', 'automation'], completed: true },
    { list: 'straventis_coding', name: 'OAuth 2.0 & Shortcuts Workflow', description: 'Refined personal productivity system in Notion while configuring OAuth 2.0 integrations and automated workflows.', priority: 'low', due_date: '2026-06-01', duration_minutes: 39, tags: ['automation'], completed: true },
    { list: 'straventis_coding', name: 'Gmail Apps Script', description: 'Advanced development of Gmail_Thiru_LabelClassifier Google Apps Script, focusing on Google AI Studio APIs.', priority: 'low', due_date: '2026-06-01', duration_minutes: 24, tags: ['automation'], completed: true },
    { list: 'straventis_coding', name: 'Timesheet Data Engineering', description: 'Processed and cleaned a year-to-date timesheet CSV from Timely, researching ISO8601 formatting and ClickUp workspace limits.', priority: 'low', due_date: '2026-05-31', duration_minutes: 20, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Diagramming Workflows', description: 'Developed detailed diagrams and flowcharts to visualize time tracking, project structures, and client relationships.', priority: 'low', due_date: '2026-05-31', duration_minutes: 93, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Rize Data Structuring', description: 'Utilized Claude to clean and structure task import data across Excel spreadsheets for use in Rize.', priority: 'low', due_date: '2026-05-31', duration_minutes: 32, tags: ['rize'], completed: true },
    { list: 'straventis_coding', name: 'Python Import Script for Rize', description: 'Refined and tested Python-based import script for Rize, utilizing Claude for development assistance.', priority: 'low', due_date: '2026-05-31', duration_minutes: 50, tags: ['rize'], completed: true },
    { list: 'straventis_coding', name: 'Jupyter Chat Parser Script', description: 'Resolved Google Drive access issues and built Jupyter chat-to-transcript parser for legal evidence prep.', priority: 'low', due_date: '2026-05-12', duration_minutes: 150, tags: [], completed: true },
    { list: 'straventis_coding', name: 'Plaud to Notion Automation', description: 'Built and refined Zapier Plaud Notes → Notion automation and organized related Notion library structures.', priority: 'low', due_date: '2026-04-30', duration_minutes: 120, tags: ['notion', 'automation'], completed: true },

    // ================================================================
    // INBOX — COMPLETED
    // ================================================================
    { list: 'inbox_wip', name: 'Apply to this job (ChatGPT fit analysis 1)', description: 'https://chatgpt.com/s/t_69d6644257488191a48902416d41d82c', priority: 'low', tags: [], completed: true },
    { list: 'inbox_wip', name: 'Apply to this job (ChatGPT fit analysis 2)', description: 'https://chatgpt.com/s/t_69d664119f348191983c11753af8ba40', priority: 'low', tags: [], completed: true },
    { list: 'inbox_wip', name: 'BairesDev application activation', description: 'https://app.smartmailcloud.com/web-share/W4MsKxmO7O8CZQsRxOQIWJXmcO0fyOrzp41ZivvT', priority: 'low', due_date: '2026-03-04', tags: [], completed: true },

  ];
}

// ---- MAIN MIGRATION FUNCTION ----
function migrateAll() {
  var startTime = Date.now();
  var props = PropertiesService.getScriptProperties();
  var sheet = getOrCreateLogSheet();
  var tasks = getAllTasks();
  var startIdx = parseInt(props.getProperty('migrate_start') || '0');

  Logger.log('▶ Starting migration from index ' + startIdx + ' of ' + tasks.length);

  var created = 0, skipped = 0, failed = 0;

  for (var i = startIdx; i < tasks.length; i++) {

    // Time check
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      props.setProperty('migrate_start', String(i));
      Logger.log('⏸ Paused at index ' + i + '. Run migrateAll() again to resume.');
      logResult(sheet, '--- PAUSED at index ' + i + ' ---', '', 'PAUSED', '', '');
      return;
    }

    var task = tasks[i];
    var listId = LIST_IDS[task.list];

    if (!listId) {
      Logger.log('⚠ Unknown list key: ' + task.list + ' for task: ' + task.name);
      logResult(sheet, task.name, task.list, 'SKIPPED', '', 'Unknown list key');
      skipped++;
      continue;
    }

    try {
      var result = createTask(listId, task);
      var cuTaskId = result.id || '';
      logResult(sheet, task.name, task.list, task.completed ? 'CREATED_CLOSED' : 'CREATED', cuTaskId, '');
      created++;
      Logger.log('✅ [' + (i+1) + '/' + tasks.length + '] ' + task.name);
      Utilities.sleep(200); // rate limit buffer
    } catch(e) {
      Logger.log('❌ FAILED: ' + task.name + ' — ' + e.message);
      logResult(sheet, task.name, task.list, 'FAILED', '', e.message);
      failed++;
    }
  }

  // All done
  props.deleteProperty('migrate_start');
  Logger.log('✅ Migration complete! Created: ' + created + ' | Skipped: ' + skipped + ' | Failed: ' + failed);
  logResult(sheet, '=== COMPLETE === Created:' + created + ' Skipped:' + skipped + ' Failed:' + failed, '', 'DONE', '', '');
}

// ---- RESET ----
function resetMigration() {
  PropertiesService.getScriptProperties().deleteProperty('migrate_start');
  Logger.log('Reset. migrateAll() will start from the beginning.');
}

// ---- OPEN LOG ----
function openLog() {
  var id = PropertiesService.getScriptProperties().getProperty('LOG_SHEET_ID');
  if (id) {
    Logger.log('Log sheet: https://docs.google.com/spreadsheets/d/' + id);
  } else {
    Logger.log('No log sheet created yet. Run migrateAll() first.');
  }
}

// ---- STATUS CHECK ----
function checkStatus() {
  var props = PropertiesService.getScriptProperties();
  var idx = props.getProperty('migrate_start');
  var tasks = getAllTasks();
  if (idx) {
    Logger.log('Paused at index ' + idx + ' of ' + tasks.length + '. Run migrateAll() to resume.');
  } else {
    Logger.log('No active migration. Total tasks to migrate: ' + tasks.length + '. Run migrateAll() to start.');
  }
}