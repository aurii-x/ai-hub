# ClickUp_Rize_PipelineV2.gs

A Google Apps Script pipeline that syncs time entries from
[Rize](https://rize.io) into [ClickUp](https://clickup.com). Built for the
**Unlimited plan** — no Business-tier features required.

---

## Table of Contents

- [How it works](#how-it-works)
- [Prerequisites](#prerequisites)
- [First-time setup](#first-time-setup)
- [Configuration](#configuration)
- [Running the pipeline](#running-the-pipeline)
  - [Historical mode (fixed date range)](#historical-mode-fixed-date-range)
  - [Incremental mode (automated, hourly)](#incremental-mode-automated-hourly)
- [Function reference](#function-reference)
- [Google Sheet layout](#google-sheet-layout)
- [Rize_Stage_Raw tab — Run Log column](#rize_stage_raw-tab--run-log-column)
- [Sync status values](#sync-status-values)
- [ClickUp plan constraints](#clickup-plan-constraints)
- [Re-run safety](#re-run-safety)
- [Known limitations](#known-limitations)

---

## How it works

```
Rize GraphQL API  (read-only)
        │
        ▼
Step 1  Extract    → Rize_Stage_Raw tab     (append, never overwrite)
Step 2  Transform  → Rize_Clean_Sync tab    (deduplicated permanent record)
Step 3a Match      → columns H–L            (ClickUp task lookup + overlap check)
Step 3b Sync       → ClickUp Track Time     (POST tid/start/stop)
        │
        ├── Step 4   Custom fields          (write Rize IDs to ClickUp task fields)
        ├── Step 5   Subtask injection      (historical mapping override)
        └── Step 0   Delete                 (rollback / clean slate)

Diagnostic
Step Audit   Time Entry Audit tab           (duplicates, overlaps, conflicts, long entries)
```

**What gets written to ClickUp per matched entry:**
- A time entry (`tid` + `start` + `stop`)
- Rize description posted as a task comment
- Start Date / Due Date backfilled if currently empty
- Task assigned to you if currently unassigned

**What is never touched in Rize:** nothing. The pipeline is read-only against Rize.

**How matching works:** Rize keeps synced task names as read-only mirrors of ClickUp task names. The pipeline matches on exact normalized name — no fuzzy logic needed here.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Google Account | Script runs under your identity |
| ClickUp personal API token | Settings → Apps → Generate |
| ClickUp Team ID (Workspace ID) | In the URL: `app.clickup.com/{team_id}/...` |
| Rize API key | Rize Settings → API |
| ClickUp Unlimited plan or higher | Free plan blocks time entry creation |
| Rize tasks synced to ClickUp | The native Rize↔ClickUp integration must be active so task names mirror correctly |

---

## First-time setup

### 1. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Paste the full contents of `ClickUp_Rize_PipelineV2.gs`
3. Save (`Ctrl+S`)

### 2. Save credentials

Add a temporary `setup()` function, run it once, then delete it:

```javascript
function setup() {
  PropertiesService.getScriptProperties().setProperties({
    CLICKUP_TOKEN:   'pk_YOUR_TOKEN_HERE',
    CLICKUP_TEAM_ID: 'YOUR_TEAM_ID_HERE',
    RIZE_API_KEY:    'YOUR_RIZE_KEY_HERE',
  });
}
```

Credentials live in Script Properties — never in the script code.

### 3. Verify

Run `checkSetup()`. It logs the first 8 characters of each credential to confirm they saved correctly without exposing the full values.

### 4. Optional — ClickUp custom fields

To write Rize IDs back to ClickUp task fields:

1. In ClickUp: Space → Customize → Custom Fields → create two **Text** fields
   named `Rize Task ID` and `Rize Time Entry ID`
2. Run `discoverCustomFieldIds()` — it logs every custom field ID on your tasks
3. Paste the two IDs into the constants near the top of Step 4:

```javascript
const CUSTOM_FIELD_ID_RIZE_TIME_ENTRY = 'd5414fdd-...';
const CUSTOM_FIELD_ID_RIZE_TASK       = '6c8b4acf-...';
```

---

## Configuration

All configuration lives at the top of the script.

### Date window

```javascript
const START_DATE_OVERRIDE = "2026-01-01"; // YYYY-MM-DD
const END_DATE_OVERRIDE   = "2026-01-31"; // YYYY-MM-DD
```

| Both set | Historical mode — processes exactly this date range |
|---|---|
| Both empty `""` | Incremental mode — uses rolling `LOOKBACK_HOURS` window |

### File locations

```javascript
const TARGET_FOLDER_PATH = ['AppData', '3.1 clickup-automation'];
const BASE_FILENAME      = 'ClickUp-Rize Sync';
const CLEAN_TAB_NAME     = 'Rize_Clean_Sync';
```

The workbook is created once inside `TARGET_FOLDER_PATH` from Drive root.
Nested folders are created automatically if missing.

### Runtime

```javascript
const MAX_RUNTIME_MS     = 330000; // 5.5 min  (Apps Script hard limit = 6 min)
const TIME_BUFFER_MS     = 25000;  // 25 sec safety margin before auto-stop
const SYNC_DELAY_SECONDS = 1;      // pause between ClickUp POST calls
```

### Matching and dedup

```javascript
const INCLUDE_CLOSED_TASKS = true;  // include completed/done tasks in name lookup
const OVERLAP_TOLERANCE_MS = 5000;  // 5 sec — entries closer than this = duplicate
```

### Step 0 safety switch

```javascript
const DRY_RUN_STEP0 = true; // MUST set to false to actually delete anything
```

### Incremental mode

```javascript
const INCREMENTAL_TRIGGER_HOURS        = 1;  // hourly
const INCREMENTAL_LOOKBACK_BUFFER_MINS = 30; // overlap buffer between runs
```

### Audit window

```javascript
const START_AUDIT_DATE = '2026-01-01'; // change per audit run
const END_AUDIT_DATE   = '2026-01-31'; // keep narrow — one month at a time
```

### Custom fields

```javascript
const CUSTOM_FIELD_ID_RIZE_TIME_ENTRY = ''; // from discoverCustomFieldIds()
const CUSTOM_FIELD_ID_RIZE_TASK       = ''; // from discoverCustomFieldIds()
```

---

## Running the pipeline

### Historical mode (fixed date range)

Use for first-time backfill or correcting a specific past window.

1. Set `START_DATE_OVERRIDE` and `END_DATE_OVERRIDE`
2. Run steps in order:

```
runStep1_ExtractRizeToStaging()
runStep2_CleanAndTransformStaging()
runStep3a_DiagnosticClickUpMatcher()
runStep3b_ActiveClickUpSyncLoader()
```

If a step hits the 5.5-minute ceiling it saves its position and stops cleanly.
Re-run the same step — it resumes exactly where it left off.

**Processing the next month:** change the date constants and re-run from Step 1.
Entries already in the clean sheet are skipped automatically.

**If something went wrong and you need to redo a window:**

```
runStep0_DeleteTrackedTimeInWindow()   ← deletes ClickUp entries + resets sheet
runStep3b_ActiveClickUpSyncLoader()    ← re-syncs (no need to redo Steps 1–3a)
```

### Incremental mode (automated, hourly)

Use after historical backfill is complete.

```javascript
switchToIncrementalMode()   // call once — sets today as the anchor
setupIncrementalTrigger()   // creates hourly Apps Script trigger
```

`runFullPipeline()` fires automatically every hour. It resolves the window
(last run's end → now, with a 30-minute overlap buffer), runs all four steps,
and advances the anchor on completion.

```javascript
checkPipelineStatus()       // mode, anchor timestamp, trigger active?
removeIncrementalTrigger()  // pause
setupIncrementalTrigger()   // resume (picks up from last anchor)
switchToHistoricalMode()    // revert to manual date control
```

---

## Function reference

### Core steps

| Function | Run | Does |
|---|---|---|
| `runStep1_ExtractRizeToStaging()` | First | Fetches Rize entries for the configured window. Pre-filters IDs already in the clean sheet. Appends new rows to `Rize_Stage_Raw` with a Run Log timestamp. |
| `runStep2_CleanAndTransformStaging()` | After Step 1 | Reads the raw tab, deduplicates against the clean sheet by Rize entry ID, converts ISO timestamps to Unix ms, parses task JSON, sorts by task name then date, writes to `Rize_Clean_Sync`. Preserves existing diagnostic columns (H–L) on rewrite. |
| `runStep3a_DiagnosticClickUpMatcher()` | After Step 2 | Fetches all ClickUp tasks (subtasks + closed included). Exact-name matches each clean-sheet row. Checks live ClickUp time entries for overlap against the 5-second tolerance. Writes `clickup_task_id` and match status to columns H/L. Skips rows already marked `Synced Successfully`. |
| `runStep3b_ActiveClickUpSyncLoader()` | After Step 3a | Processes only `New Entry` and `API Error` rows. POSTs `{tid, start, stop}` to ClickUp. Posts Rize description as task comment. Assigns task to you if unassigned. Backfills Start/Due dates if empty. Marks `Synced Successfully` on completion. |
| `runFullPipeline()` | Automated trigger | Orchestrates Steps 1–3b in sequence. In incremental mode resolves the rolling window. Budget-checks between steps and pauses if the ceiling is near. |

### Reset and maintenance

| Function | Does |
|---|---|
| `runStep0_DeleteTrackedTimeInWindow()` | Deletes ClickUp time entries within the date window, scoped to task IDs from the clean sheet only. Resets column L to `New Entry`. Always run with `DRY_RUN_STEP0 = true` first. |
| `resetSync()` | Clears the `RIZE_SYNC_START_IDX` Script Property so `runFullPipeline()` restarts from task 0. |
| `checkSetup()` | Verifies credentials are saved. Prints first 8 characters of each. |

### Custom fields (Step 4)

| Function | Does |
|---|---|
| `discoverCustomFieldIds()` | Fetches all custom fields from clean-sheet tasks and logs name, ID, and type. Run once to find the IDs to paste into the constants. |
| `runStep4_WriteRizeIdsToCustomFields()` | Writes comma-separated Rize time entry IDs to `CUSTOM_FIELD_ID_RIZE_TIME_ENTRY` and Rize task ID to `CUSTOM_FIELD_ID_RIZE_TASK` on each matched ClickUp task. |

### Subtask injection (Step 5 / 5b)

Used when Rize entries were tagged generically (`1.1 Deep Work`, `Unassigned`)
and can't match ClickUp task names automatically. Contains hardcoded
`rize_time_entry_id → clickup_task_id` mappings built from a one-time manual
analysis pass.

| Function | Does |
|---|---|
| `runStep5_InjectSubtaskMapping()` | Writes ClickUp subtask IDs into column H for 134 manually-mapped historical entries. Resets column L to `New Entry`. Run Step 3b after. |
| `runStep5b_InjectAdditionalMapping()` | Same as Step 5 for an additional 570 entries across 12 task groups. Run after Step 5, before Step 3b. |

### Incremental mode controls

| Function | Does |
|---|---|
| `switchToIncrementalMode()` | Sets `INCREMENTAL_MODE=true` and `INCREMENTAL_LAST_END=now` in Script Properties. Date override constants are ignored while active. |
| `switchToHistoricalMode()` | Reverts to manual date control. Clears incremental properties. |
| `setupIncrementalTrigger()` | Creates an hourly trigger on `runFullPipeline()`. Removes any existing trigger first to prevent stacking. |
| `removeIncrementalTrigger()` | Deletes the trigger. Data and state are preserved. |
| `checkPipelineStatus()` | Logs mode, last run anchor, and trigger status. |

### Diagnostic

| Function | Does |
|---|---|
| `runDiagnostic_TimeEntryAudit()` | Fetches all ClickUp time entries between `START_AUDIT_DATE` and `END_AUDIT_DATE`. Flags duplicates (same start+end), overlaps (> 2 min), conflicting entries (end before start), and entries over 4 hours. Writes colour-coded results to `Time Entry Audit` tab. Run one month at a time — the full workspace history is too large for a single Apps Script execution. |

---

## Google Sheet layout

One workbook (`ClickUp-Rize Sync`) in the configured Drive folder. Tabs:

| Tab | Created by | Purpose |
|---|---|---|
| `Rize_Stage_Raw` | Step 1 | Raw entries fetched from Rize. Append-only — rows accumulate across runs. |
| `Rize_Clean_Sync` | Step 2 | Deduplicated permanent record. The pipeline's source of truth. |
| `Time Entry Audit` | Diagnostic | Audit results. Overwritten on each audit run. |

### `Rize_Stage_Raw` column schema

| Col | Header | Notes |
|---|---|---|
| A | `Run Log` | Eastern Time timestamp of the run that added this row (`2026-06-25_9:34`) |
| B | `id` | Rize time entry ID |
| C | `startTime` | ISO 8601 string |
| D | `endTime` | ISO 8601 string |
| E | `title` | Rize entry title |
| F | `description` | Rize entry description |
| G | `task_json` | Raw JSON of the linked Rize task `{id, name}` |

### `Rize_Clean_Sync` column schema

| Col | Header | Type | Notes |
|---|---|---|---|
| A | `start` | Unix ms | Entry start timestamp |
| B | `end` | Unix ms | Entry end timestamp |
| C | `duration` | Unix ms | `end - start` |
| D | `description` | String | Rize description or title (whichever is non-empty) |
| E | `rize_time_entry_id` | String | The dedup key — used to skip re-staging |
| F | `rize_task_id` | String | Rize internal task ID |
| G | `task_name_lookup` | String | Rize task name — used for ClickUp exact-name matching |
| H | `clickup_task_id` | String | Written by Step 3a after a match is found |
| I | `clickup_start_display` | String | Human-readable start in Eastern Time |
| J | `clickup_end_display` | String | Human-readable end in Eastern Time |
| K | `clickup_duration_display` | String | e.g. `1h 30m` |
| L | `sync_match_status` | String | See [Sync status values](#sync-status-values) |

---

## Rize_Stage_Raw tab — Run Log column

Column A (`Run Log`) is stamped with the Eastern Time of the pipeline run that
added each row, formatted as `YYYY-MM-DD_H:mm` (no leading zero on hour).

On first run the tab is created with this column already in place.
On subsequent runs new rows are **appended** — existing rows from prior runs
are never deleted or modified. This gives you a full history of what was fetched
on each run and when.

If you open a pre-existing `Rize_Stage_Raw` tab created before this column
was introduced, the script inserts the column automatically on the next Step 1
run and logs:
```
📋 Inserted missing "Run Log" column into existing tab.
```

---

## Sync status values

Column L of `Rize_Clean_Sync` after Step 3a / 3b:

| Value | Set by | Meaning |
|---|---|---|
| `New Entry` | Step 3a / Step 5 | No overlapping entry found in ClickUp — ready for Step 3b |
| `Synced Successfully` | Step 3b | Time entry was posted to ClickUp successfully |
| `Already Matched` | Step 3a | An overlapping entry already exists in ClickUp — skipped |
| `Missing ClickUp Task` | Step 3a | No ClickUp task name matched `task_name_lookup` |
| `Conflict` | Step 3a | A partially-overlapping entry exists — needs manual review |
| `API Error (nnn)` | Step 3b | ClickUp rejected the POST — HTTP status code shown |
| `Network Error` | Step 3b | `UrlFetchApp` threw an exception |

---

## ClickUp plan constraints

The pipeline targets the **Unlimited plan**. Time entry payloads contain only
`tid`, `start`, and `stop`. The following fields are intentionally excluded:

| Field | Why excluded |
|---|---|
| `description` on time entries | Business plan only — triggers `TIMEENTRY_064` on Unlimited |
| `tags` on time entries | Business plan only |
| `billable` | Business plan only |
| `created_with` | Not a valid ClickUp API field |

Rize entry descriptions are written as **task comments** instead, which are
available on all plans.

---

## Re-run safety

Every step is designed to be re-run without creating duplicates.

| Step | How duplicates are prevented |
|---|---|
| Step 1 | Pre-loads all Rize entry IDs from the clean sheet before hitting the Rize API. Filters out any ID already present. |
| Step 2 | Reads existing clean-sheet IDs into a `Set` before processing the raw tab. Skips any row whose `rize_time_entry_id` is already in the set. |
| Step 3a | Fetches live ClickUp time entries per task before writing. Compares start/end with a 5-second tolerance. Skips rows already marked `Synced Successfully`. |
| Step 3b | Only processes rows where column L is `New Entry` or `API Error`. Rows marked `Synced Successfully` or `Already Matched` are untouched. |

---

## Known limitations

**Rize tag corrections require manual work.** If a Rize entry is tagged to
the wrong task, the pipeline cannot detect this. Re-tag the entry in Rize's
review panel — the next run will pick it up under the corrected task.

**Unassigned Rize entries cannot be matched.** Entries with no task tag
(`task = null`) land in the unmatched list. Use Step 5 / 5b injection to map
them manually, or re-tag in Rize.

**Task name collisions cause ambiguous matches.** If two ClickUp tasks have
identical names, entries for that name are flagged `ambiguous` and not synced.
Rename one task to disambiguate.

**The Apps Script 6-minute ceiling applies.** Large historical windows may
need several sequential runs. Every step checkpoints and resumes cleanly.

**Audit runs one month at a time.** The full workspace has too many entries
(700k+) to audit in a single execution. Set `START_AUDIT_DATE` /
`END_AUDIT_DATE` to a single calendar month, run, then advance the window.

**All timestamps in the sheet are UTC.** The `start`, `end`, and `duration`
columns store raw Unix milliseconds. Human-readable display columns (I, J, K)
are formatted in Eastern Time (`America/New_York`) via `fmtTs()`.