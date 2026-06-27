Functional Capabilities

- Automatic Pagination Checkpoints: Processes massive backlogs safely in 5-minute intervals. If the operation approaches Google's 6-minute execution window execution ceiling, it saves its index to PropertiesService and cleanly halts execution without breaking data strings.
- Bi-Directional State Injection: Maps tasks based on completion status. Open tasks are pushed directly to native open list states, while archival entries are generated as instantly completed closed elements (status: 'complete') to retain data history.
- Smart Duration Parsing: Automatically matches string duration expressions (e.g., 30m, 2h) and formats them into strict millisecond parameters required by ClickUp's time_estimate metric tracking.
- Dynamic Logging System: Automatically spins up a companion Google Sheet (ClickUp Migration Log) to provide line-by-line live status outputs (CREATED, CREATED_CLOSED, SKIPPED, FAILED).

Project Settings & Mappings

The script relies on a clean index dictionary mapping specific task categories to targeted ClickUp Space List IDs:

System Category Key
Associated ClickUp List Target ID

career_deep_work
901417224062

career_job_search
901417224063

career_communication
901417224064

learning_courses
901417224065

learning_research
901417224067

personal_health
901417224069

personal_family
901417224070

personal_finance_legal
901417224072

personal_rest
901417224073

device_tinkering
901417224075

device_focus
901417224080

web_productive
901417224082

web_news
901417224083

web_social
901417224086

straventis_client
901417224087

straventis_coding
901417224091

mindless_vortex
901417224093

mindless_drift
901417224094

inbox_today
901417224095

inbox_tomorrow
901417224099

inbox_wip
901417224102

Execution Protocol

Step 1: Secure Token Authorization

Open the script inside the editor, navigate directly to the setup() block, input your system-wide token parameter, and run the function:

    function setup() {
      var token = 'YOUR_CLICKUP_API_TOKEN_HERE'; // Input personal ClickUp API token here
      PropertiesService.getScriptProperties().setProperty('CU_TOKEN', token);
    }