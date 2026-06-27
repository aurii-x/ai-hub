1. Trigger: You mark a task as completed in Todoist.
2. Webhook Event: Todoist fires an item:completed event payload to this script's web app URL.
3. Evaluation: The script checks if the completed task contains the specified target label (e.g., wip).
4. API Execution: If found, the script filters out that label while preserving all other existing labels, and sends a REST API v2 request back to Todoist to update the task.

Features

- Automated Cleanup: Automatically removes status labels like @wip, @next, or @blocked immediately when tasks are checked off.
- Tag Preservation: Safely filters out only the targeted label, leaving your context tags (e.g., @home, @computer, @5min) completely untouched.
- Lightweight & Free: Runs entirely on Google Apps Script with zero server hosting costs or external dependencies.

Setup Instructions

1. Google Apps Script Setup

1. Head over to Google Apps Script.
2. Create a New Project.
3. Clear out any boilerplate code in Code.gs and paste the provided script code.
4. Replace the value of TODOIST_API_TOKEN with your personal Todoist API token.
(You can find this in Todoist under Settings > Integrations > Developer > API token).
5. If desired, change the value of LABEL_TO_REMOVE (default is 'wip'). Note: This is case-sensitive and should not include the @ symbol.

2. Deploying the Script as a Web App

In order for Todoist to communicate with your script, you must deploy it as a public web application:

1. In the Apps Script editor, click the Deploy button in the top right corner and choose New deployment.
2. Click the gear icon (Select type) and select Web app.
3. Configure the deployment settings:

Description: Enter a meaningful description (e.g., "Todoist Label Remover").
Execute as: Set to Me (your-email@gmail.com).
Who has access: Set to Anyone. (This is required so Todoist can securely send events to it without requiring a Google authentication login).
4. Click Deploy.
5. Grant any necessary permissions if prompted by Google.
6. Copy the generated Web app URL (you will need this for the next phase).

3. Configuring the Todoist Webhook

1. Go to the Todoist Developer Console and sign in.
2. Click Create a new app. Give it a name (e.g., "Label Cleanup Automation") and click Create app.
3. Scroll down to the Webhooks section.
4. In the Webhook URL field, paste the Web app URL you copied from Google Apps Script.
5. In the Watched events list, check the box next to item:completed.
6. Activate the webhook by toggling the status switch or saving changes.

Script Constants Configuration

    // The 40-character alphanumeric token from your Todoist account integrations page
    const TODOIST_API_TOKEN = 'YOUR_ACTUAL_API_TOKEN'; 
    
    // The label you want stripped out when tasks are completed (Case-Sensitive, do not include '@')
    const LABEL_TO_REMOVE   = 'wip';