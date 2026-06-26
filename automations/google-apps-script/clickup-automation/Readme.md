# santhoshOS — ClickUp Automation Scripts

Google Apps Script tools for syncing [Rize](https://rize.io) time tracking
data into [ClickUp](https://clickup.com), maintaining historical time records,
and auditing/cleaning tracked time.


---

## Scripts

| Script | Purpose | Docs |
|---|---|---|
| `ClickUp_Rize_PipelineV2.gs` | Main 4-step pipeline: fetches Rize time entries, matches to ClickUp tasks, and posts tracked time. Supports historical backfill and incremental hourly automation. | [→ docs](docs/ClickUp_Rize_PipelineV2.md) |
| `ClickUp_TimeEntry_Cleanup.gs` | Destructive utility: deletes all ClickUp time entries in a configured date window. Use before a clean re-sync. | [→ docs](docs/ClickUp_TimeEntry_Cleanup.md) |
| `ClickUp_CareerflowSync.gs` | One-time CSV import: reads a Careerflow job-application export, fuzzy-matches rows to ClickUp tasks, and updates task dates and tracked time. | [→ docs](docs/ClickUp_CareerflowSync.md) |
| `ClickUp_Rize_HourlySync.gs` | Rolling hourly sync: reconciles the last 48 hours of Rize activity against ClickUp tasks across all spaces. | [→ docs](docs/ClickUp_Rize_HourlySync.md) |
| `ClickUp_Rize_HistoricalSync.gs` | Weekly historical backfill: processes Rize history from a fixed anchor date (Jan 1 2026) through today. Re-run-safe. | [→ docs](docs/ClickUp_Rize_HistoricalSync.md) |

---

## Quick Start

Every script requires the same two ClickUp credentials stored in
[Script Properties](https://developers.google.com/apps-script/guides/properties).
Paste this into any script, run it once, then delete the values:

```javascript
function setup() {
  PropertiesService.getScriptProperties().setProperties({
    CLICKUP_TOKEN:   'pk_YOUR_PERSONAL_API_TOKEN',
    CLICKUP_TEAM_ID: 'YOUR_WORKSPACE_ID',
  });
}
```

Scripts that connect to Rize also need:

```javascript
RIZE_API_KEY: 'YOUR_RIZE_API_KEY'
```

Credentials are never stored in code — only in Script Properties.

---

## ClickUp Plan

All scripts are tested against the **Unlimited plan**. Time entry payloads
use only `tid`, `start`, and `stop` — the three fields available without a
Business plan upgrade. Fields like `description`, `tags`, and `billable` on
time entries are intentionally excluded to avoid `TIMEENTRY_064` errors.

---

## Architecture Overview

```
Rize (read-only)         ClickUp
      │                      ▲
      │   Google Apps Script  │
      └──► Step 1  Extract   │
           Step 2  Transform │
           Step 3a Match     │
           Step 3b ──────────┘  POST time entries
           Step 4  ──────────►  Write custom fields
           Step 0  ──────────►  DELETE (cleanup)
```

Full data-flow and design documentation: [ClickUp_Rize_PipelineV2.md](docs/ClickUp_Rize_PipelineV2.md)

---

## Repo Structure

```
├── README.md                            ← you are here
├── ClickUp_Rize_PipelineV2.gs          ← main pipeline (use this one)
├── ClickUp_TimeEntry_Cleanup.gs        ← destructive cleanup utility
├── ClickUp_CareerflowSync.gs           ← Careerflow CSV → ClickUp
├── ClickUp_Rize_HourlySync.gs          ← rolling hourly sync
├── ClickUp_Rize_HistoricalSync.gs      ← weekly historical backfill
└── docs/
    ├── ClickUp_Rize_PipelineV2.md
    ├── ClickUp_TimeEntry_Cleanup.md
    ├── ClickUp_CareerflowSync.md
    ├── ClickUp_Rize_HourlySync.md
    └── ClickUp_Rize_HistoricalSync.md
```