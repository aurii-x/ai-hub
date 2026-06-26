
// ─── CONFIG ──────────────────────────────────────────────────────────────────
const CSV_FILE_NAME                 = 'Job_Search_Time_Reconciliation_Tab1.csv';
const DRIVE_FOLDER_PATH             = ['AppData', '3.1 clickup-automation'];
const LOG_SHEET_NAME                = 'ClickUpSyncLog';
const MATCH_THRESHOLD               = 0.90;
const NEW_TASK_STATUS               = 'to do';   // ← run getListStatuses() and replace with the exact name
const TASK_PRIORITY                 = 3;       // 1=Urgent, 2=High, 3=Normal, 4=Low
const WORKBOOK_NAME                 = 'ClickUp-Rize Sync';
const LOOKBACK_HOURS                = 24;     // how far back to pull Rize entries each run
const INCLUDE_CLOSED_CLICKUP_TASKS  = true;  // completed tasks need to match too (Coursera modules, finished applications, etc.)
const OVERLAP_TOLERANCE_MS          = 5 * 60 * 1000; // 5 min — entries within this are "the same"
const BATCH_SIZE                    = 50;     // ClickUp tasks processed per checkpoint
const MAX_RUNTIME_MS                = 10 * 60 * 1000;
const SILENT_MODE                   = true;   // true = no UI alerts (required for hourly trigger)
const SYNC_START_DATE               = '2026-06-01'; // fixed anchor — your historical correction start point
const SYNC_END_DATE                 = '';           // blank = "today" (recalculated each run); set a fixed
const CU_API_BASE                   = 'https://api.clickup.com/api/v2';
const FETCH_MAX_PAGES               = 500
SYNC_DELAY_SECONDS                  = 3;
S3B_TIME_BUFFER_MS                  = 30000;
const CUSTOM_FIELD_ID_RIZE_TIME_ENTRY = '87470de1-ad58-4886-a440-55a1936667f5'; // ← paste field ID from discoverCustomFieldIds()
const CUSTOM_FIELD_ID_RIZE_TASK       = '6c8b4acf-36a1-444e-aa9d-daeb23601309'; // ← paste field ID from discoverCustomFieldIds()

// ─── SETUP ───────────────────────────────────────────────────────────────────
function setup() {
  PropertiesService.getScriptProperties().setProperties({
    CLICKUP_TOKEN:   'pk_216003478_5N2AE9LICIRWR32210VR9PO1J2OKS4U7', //stored in 1Password
    CLICKUP_TEAM_ID: '90141302224', // Workspace ID (numeric string)
    RIZE_API_KEY:    'Co4y8lNoO4DVStBkr3Mwbw-Y1JFnxqCFiwwUdn4vucg',
    CLICKUP_LIST_ID: '901417224063',                // List that holds your job tasks
  });
  notify('✅ Setup complete. Clear the values from setup() now.');
}

function checkSetup() {
  const p = PropertiesService.getScriptProperties();
  Logger.log('ClickUp token: ' + (p.getProperty('CLICKUP_TOKEN') ? 'set' : 'NOT SET'));
  Logger.log('ClickUp team:  ' + (p.getProperty('CLICKUP_TEAM_ID') || 'NOT SET'));
  Logger.log('Rize key:      ' + (p.getProperty('RIZE_API_KEY') ? 'set' : 'NOT SET'));
}

/**
 * Forcefully standardizes the status configurations for ALL Lists across the workspace
 * to match standard dashboard workflow buckets: Not Started, Deferred, Active, Waiting, and Closed.
 */
function syncAllWorkspaceListStatuses() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('CLICKUP_TOKEN');
  const teamId = props.getProperty('CLICKUP_TEAM_ID');
  if (!token || !teamId) {
    Logger.log('❌ Missing credentials. Run setup() first.');
    return;
  }

  Logger.log('Initializing Workspace Status Standardization Engine...');

  // 1. Define your strict custom dashboard status payload array
  const targetStatuses = [
    { status: "to do", type: "open", color: "#d3d3d3" },
    { status: "deferred", type: "custom", color: "#A5A5A5" },
    { status: "in progress", type: "custom", color: "#2ea2f8" },
    { status: "waiting", type: "custom", color: "#f39c12" },
    { status: "completed", type: "closed", color: "#6bc950" }
  ];

  // 2. Extract every single Space inside the workspace
  const spacesUrl = `https://api.clickup.com/api/v2/team/${teamId}/space`;
  const spacesResponse = apiGet(spacesUrl, token);
  if (!spacesResponse || !spacesResponse.spaces) {
    Logger.log('❌ Could not fetch space structures from ClickUp.');
    return;
  }

  let listsUpdatedCount = 0;

  // 3. Drill down into the Space layout tree
  spacesResponse.spaces.forEach(space => {
    Logger.log(`Scanning Space: [${space.name}] (ID: ${space.id})...`);

    // Grab Folderless Lists belonging to this space
    const folderlessUrl = `https://api.clickup.com/api/v2/space/${space.id}/list`;
    const folderlessResponse = apiGet(folderlessUrl, token);
    if (folderlessResponse && folderlessResponse.lists) {
      folderlessResponse.lists.forEach(list => {
        updateListStatuses(list.id, list.name, targetStatuses, token);
        listsUpdatedCount++;
      });
    }

    // Grab Folders inside this space to scrape their inner lists
    const foldersUrl = `https://api.clickup.com/api/v2/space/${space.id}/folder`;
    const foldersResponse = apiGet(foldersUrl, token);
    if (foldersResponse && foldersResponse.folders) {
      foldersResponse.folders.forEach(folder => {
        const folderListsUrl = `https://api.clickup.com/api/v2/folder/${folder.id}/list`;
        const folderListsResponse = apiGet(folderListsUrl, token);
        if (folderListsResponse && folderListsResponse.lists) {
          folderListsResponse.lists.forEach(list => {
            updateListStatuses(list.id, list.name, targetStatuses, token);
            listsUpdatedCount++;
          });
        }
      });
    }
  });

  Logger.log(`\n=========================================================`);
  Logger.log(`✅ Success! Forcefully synchronized statuses for ${listsUpdatedCount} Lists workspace-wide.`);
  Logger.log(`=========================================================`);
}

/**
 * Low-level utility to PUT the clean state configuration onto a target list card.
 */
function updateListStatuses(listId, listName, statusesPayload, token) {
  const url = `https://api.clickup.com/api/v2/list/${listId}`;
  const body = {
    statuses: statusesPayload
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: { 
        'Authorization': token, 
        'Content-Type': 'application/json' 
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      Logger.log(`  ➔ Standardized List: "${listName}" (ID: ${listId})`);
    } else {
      Logger.log(`  ⚠ Failed updating List "${listName}" (ID: ${listId}): HTTP ${response.getResponseCode()}`);
    }
    Utilities.sleep(150); // Socket pacing delay to protect against network rate limits
  } catch (e) {
    Logger.log(`  ❌ Exception error on List ${listId}: ${e}`);
  }
}