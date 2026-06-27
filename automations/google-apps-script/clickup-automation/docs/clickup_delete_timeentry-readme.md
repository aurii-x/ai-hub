# ClickUp Time Entry Cleanup Tool

A focused Google Apps Script that deletes all ClickUp time entries within a
configured date window. Used to clear out bad data before a clean re-sync —
for example, after a pipeline bug produced duplicate or incorrectly-timestamped
entries.

> ⚠️ **This script is destructive. Deleted time entries cannot be recovered.
> Always confirm the date window is correct before running.**

---

## Contents

- [When to use this](#when-to-use-this)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Configuration](#configuration)
- [Running the script](#running-the-script)
- [What it does (step by step)](#what-it-does-step-by-step)
- [After cleanup](#after-cleanup)
- [Limitations](#limitations)
- [Comparison with Step 0 in the pipeline](#comparison-with-step-0-in-the-pipeline)

---

## When to use this

Use this script when you need to wipe all ClickUp time entries in a date window
regardless of which task they belong to — for example:

- A pipeline bug posted every entry twice (duplicates across all tasks)
- Wrong timestamps were written (e.g. UTC instead of Eastern Time)
- You want a completely clean slate before re-running the full sync pipeline

If you only need to delete entries for tasks that were synced through the
pipeline and want the clean sheet updated at the same time, use
**`runStep0_DeleteTrackedTimeInWindow()`** in `ClickUp_Rize_PipelineV2.gs`
instead — it is scoped to pipeline-owned tasks and has a dry-run mode.
See [Comparison with Step 0](#comparison-with-step-0-in-the-pipeline).

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Google Account | Script runs under your Google identity |
| ClickUp personal API token | Settings → Apps → Generate |
| ClickUp Team ID | Visible in URL: `app.clickup.com/{team_id}/...` |
| Credentials in Script Properties | Run `setup()` first (see below) |

---

## Setup

### 1. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Paste the contents of `ClickUp_TimeEntry_Cleanup.gs`
3. Save with `Ctrl+S`

### 2. Save credentials

If you are already running `ClickUp_Rize_PipelineV2.gs` in a separate project,
credentials are stored per-project — you need to save them here too.

Add a `setup()` function, run it once, then remove the values:

```javascript
function setup() {
  PropertiesService.getScriptProperties().setProperties({
    CLICKUP_TOKEN:   'pk_YOUR_TOKEN_HERE',
    CLICKUP_TEAM_ID: 'YOUR_TEAM_ID_HERE',
  });
}
```

Credentials are stored in Script Properties — not in the script code.

---

## Configuration

Three constants at the top of the script control all behaviour:

```javascript
const CLEANUP_START_DATE = "2026-01-01T00:00:00Z"; // ISO 8601 UTC — window start
const CLEANUP_END_DATE   = "2026-01-31T23:59:59Z"; // ISO 8601 UTC — window end
const DELETE_DELAY_MS    = 2000;                    // milliseconds between deletions
```

### `CLEANUP_START_DATE` / `CLEANUP_END_DATE`

The date range passed to ClickUp's time entries API. **Use UTC (`Z` suffix).**
ClickUp stores and filters timestamps in UTC regardless of your workspace timezone.

Examples:

```javascript
// All of January 2026
const CLEANUP_START_DATE = "2026-01-01T00:00:00Z";
const CLEANUP_END_DATE   = "2026-01-31T23:59:59Z";

// A single day
const CLEANUP_START_DATE = "2026-03-15T00:00:00Z";
const CLEANUP_END_DATE   = "2026-03-15T23:59:59Z";

// Everything from a specific sync run (use the run timestamp from your log sheet)
const CLEANUP_START_DATE = "2026-06-04T05:00:00Z"; // 1am ET = 5am UTC
const CLEANUP_END_DATE   = "2026-06-04T23:59:59Z";
```

### `DELETE_DELAY_MS`

Pause between each DELETE call. The default of 2000ms (2 seconds) is
conservative and avoids ClickUp rate-limit errors. If you are deleting a small
number of entries you can reduce this to 500ms. Do not set below 200ms.

---

## Running the script

1. Set `CLEANUP_START_DATE` and `CLEANUP_END_DATE` to the exact window you want
   to clear
2. **Read the execution log carefully before confirming** — the script logs
   every entry it finds (count, task ID, duration) before deleting anything
3. Select `runClickUpTimeCleanup` from the function dropdown
4. Click **Run**
5. Watch the execution log

The script logs each deletion as it happens:

```
🔍 Fetching time entries from 2026-01-01T00:00:00Z to 2026-01-31T23:59:59Z...
✅ Found 47 time entries to clean up.
🗑️ Deleting entry abc123 (15 mins on task 86badude9)...
🗑️ Deleting entry def456 (30 mins on task 86bakmr7p)...
...
🏁 Cleanup run complete. Successfully deleted 47 time entries.
```

---

## What it does (step by step)

### 1. Authenticate

Reads `CLICKUP_TOKEN` and `CLICKUP_TEAM_ID` from Script Properties. Exits
immediately with an error if either is missing.

### 2. Fetch all time entries in the window

Calls:

```
GET /api/v2/team/{teamId}/time_entries?start_date={startMs}&end_date={endMs}
```

`start_date` and `end_date` are the ISO date strings converted to Unix
milliseconds. The response is the full list of matching entries across all
tasks in the workspace.

Logs the count found. Exits cleanly if the count is zero.

### 3. Delete entries sequentially

For each entry, calls:

```
DELETE /api/v2/team/{teamId}/time_entries/{entryId}
```

Entries are deleted one at a time (not in bulk) with a `DELETE_DELAY_MS`
pause between each call to respect ClickUp's rate limits.

HTTP 200 and 204 are both treated as success. Any other response code is
logged as a warning and the script continues to the next entry — a single
failed deletion does not abort the run.

### 4. Log completion

Logs the final deleted count. Does not update any Google Sheet — this script
has no sheet dependency.

---

## After cleanup

This script only removes entries from ClickUp. It does **not** update the
`Rize_Clean_Sync` Google Sheet that the main pipeline uses to track sync state.

After running this script, if you want the pipeline to re-sync the same
entries, you need to reset the sheet's sync status manually:

1. Open `ClickUp-Rize Sync` → `Rize_Clean_Sync` tab
2. For the rows covering the deleted date range, change column L
   (`sync_match_status`) from `Synced Successfully` back to `New Entry`
3. Run `runStep3b_ActiveClickUpSyncLoader()` in `ClickUp_Rize_PipelineV2.gs`

Alternatively, use `runStep0_DeleteTrackedTimeInWindow()` from the main
pipeline script — it handles the sheet reset automatically.

---

## Limitations

| Limitation | Detail |
|---|---|
| No dry-run mode | The script deletes immediately. There is no preview step. Always double-check the date window in the constants before running. |
| No pagination | The ClickUp API returns up to 100 time entries per request. If your window contains more than 100 entries, only the first 100 will be fetched and deleted. Re-run the script to catch the next batch, or narrow the date window and run multiple passes. |
| No task filter | Deletes ALL time entries in the window across all tasks, not just pipeline-created ones. If you have manually-entered time in the same window, it will also be deleted. |
| No Apps Script time limit handling | The script does not checkpoint. If the Apps Script 6-minute limit is hit mid-run (unlikely unless deleting hundreds of entries with a high `DELETE_DELAY_MS`), entries processed before the cutoff are already deleted and will not be re-attempted. Reduce `DELETE_DELAY_MS` if needed. |
| Deleted entries are permanent | ClickUp has no recycle bin for time entries. There is no undo. |

---

## Comparison with Step 0 in the pipeline

`ClickUp_Rize_PipelineV2.gs` includes a more sophisticated deletion function
(`runStep0_DeleteTrackedTimeInWindow()`) that is safer for routine use.

| Feature | This script | Step 0 in pipeline |
|---|---|---|
| Scope | All time entries in date window | Only entries on tasks in the clean sheet |
| Dry-run mode | ❌ No | ✅ Yes (`DRY_RUN_STEP0 = true`) |
| Sheet reset | ❌ Manual | ✅ Automatic (resets column L to `New Entry`) |
| Pagination | ❌ First 100 only | ✅ Handles all entries |
| Runtime safety | ❌ No checkpoint | ✅ Pauses and resumes within 5.5-min budget |
| Use case | Full workspace wipe | Targeted pipeline re-sync |

**Rule of thumb:** use Step 0 when you can. Use this script only when you need
to delete entries that were not created by the pipeline, or when the clean sheet
no longer reflects reality and you want a complete reset.