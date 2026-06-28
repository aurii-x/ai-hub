// =============================================================================
//
/**  BLIND CLICKUP TIME ENTRY CLEANUP TOOL
// use runStep0_DeleteTrackedTimeInWindow () or runStep0b_DeleteDuplicatesOnly(), but first run setDryRunStep0b_Live() or setDryRunStep0b_DryRun()
*/
// =============================================================================
const CLEANUP_START_DATE = "2025-01-01T00:00:00Z"; // Set to when the duplicates started
const CLEANUP_END_DATE   = "2026-06-27T23:59:59Z"; // Set to current time/end of cleanup window
const DELETE_DELAY_MS    = 100;                   // 2-second delay to avoid rate limits

/**
 * Fetches time entries within a window and deletes them.
 * WARNING: This is destructive. Ensure dates are correct.
 */
function runClickUpTimeCleanup() {
  const props = PropertiesService.getScriptProperties();
  const cuToken = props.getProperty('CLICKUP_TOKEN');
  const cuTeamId = props.getProperty('CLICKUP_TEAM_ID');

  if (!cuToken || !cuTeamId) {
    Logger.log('❌ CRITICAL FAILURE: CLICKUP_TOKEN or CLICKUP_TEAM_ID missing.');
    return;
  }

  // Convert ISO dates to Unix Milliseconds for the ClickUp API
  const startMs = new Date(CLEANUP_START_DATE).getTime();
  const endMs = new Date(CLEANUP_END_DATE).getTime();

  Logger.log(`🔍 Fetching time entries from ${CLEANUP_START_DATE} to ${CLEANUP_END_DATE}...`);

  // 1. GET ALL TIME ENTRIES IN DATE RANGE
  const getUrl = `https://api.clickup.com/api/v2/team/${cuTeamId}/time_entries?start_date=${startMs}&end_date=${endMs}`;
  const getOptions = {
    method: 'get',
    headers: {
      'Authorization': cuToken,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  let timeEntries = [];
  try {
    const response = UrlFetchApp.fetch(getUrl, getOptions);
    const statusCode = response.getResponseCode();
    
    if (statusCode === 200) {
      const data = JSON.parse(response.getContentText());
      timeEntries = data.data || [];
      Logger.log(`✅ Found ${timeEntries.length} time entries to clean up.`);
    } else {
      Logger.log(`🚨 API FETCH FAILED: Returned code ${statusCode}. ${response.getContentText()}`);
      return;
    }
  } catch (error) {
    Logger.log(`❌ NETWORK CRASH during fetch: ${error.message}`);
    return;
  }

  if (timeEntries.length === 0) {
    Logger.log('🏁 Cleanup complete: No time entries found in this range.');
    return;
  }

  // 2. DELETE TIME ENTRIES SEQUENTIALLY
  let deletedCount = 0;
  
  for (let i = 0; i < timeEntries.length; i++) {
    const entryId = timeEntries[i].id;
    const taskId = timeEntries[i].task?.id || 'Workspace Level';
    const durationMins = Math.round(timeEntries[i].duration / 60000);
    
    Logger.log(`🗑️ Deleting entry ${entryId} (${durationMins} mins on task ${taskId})...`);
    
    const deleteUrl = `https://api.clickup.com/api/v2/team/${cuTeamId}/time_entries/${entryId}`;
    const deleteOptions = {
      method: 'delete',
      headers: {
        'Authorization': cuToken
      },
      muteHttpExceptions: true
    };

    try {
      const deleteResponse = UrlFetchApp.fetch(deleteUrl, deleteOptions);
      const deleteStatusCode = deleteResponse.getResponseCode();
      
      if (deleteStatusCode === 200 || deleteStatusCode === 204) {
        deletedCount++;
        Utilities.sleep(DELETE_DELAY_MS);
      } else {
        Logger.log(`⚠️ FAILED to delete ${entryId}. Code: ${deleteStatusCode}`);
      }
    } catch (error) {
      Logger.log(`❌ DELETE CRASH on entry ${entryId}: ${error.message}`);
    }
  }

  Logger.log(`🏁 Cleanup run complete. Successfully deleted ${deletedCount} time entries.`);
}