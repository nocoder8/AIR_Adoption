# AI Interview Insights Report - Google Apps Script Setup Guide

This guide will help you set up the AI Interview Insights Report generator directly in your Google Sheet using Google Apps Script.

## Benefits of Using Google Apps Script

- **No external tools required** - runs entirely within Google Sheets
- **No server setup** - hosted by Google
- **Automatic scheduling** - set up daily email reports with a click
- **Easy access control** - works with your Google account permissions

## Setup Instructions

### Step 1: Open Your Google Sheet

Open the Google Sheet that contains your candidate data with the following columns:
- Profile_id
- Name
- Current_company
- Candidate_city
- Candidate_state
- Profile_link
- Application_ts
- Business_unit
- Job_function
- Last_stage (with 'New' as one of the values)
- Last_stage_ts
- Aging
- Title
- Location
- Location_country
- Recruiter name
- Recruiter email
- Hiring manager name
- Position approved date
- Source_name
- Match_stars
- Application_status
- Position_status
- Position_id
- Sub department
- Ai_interview (with 'Y' or 'N' values)

Make sure these column headers are in the **second row** of your sheet (row 2).

### Step 2: Open the Script Editor

1. In your Google Sheet, click on **Extensions** in the top menu
2. Select **Apps Script**
3. This will open the Google Apps Script editor in a new tab

### Step 3: Add the Script

1. Delete any code in the editor
2. Copy the entire code from the `AIR_GSheets.js` file
3. Paste it into the Apps Script editor

### Step 4: Configure the Script

Update the following constants at the top of the script:

```javascript
const EMAIL_RECIPIENT = 'pkumar@eightfold.ai'; // Already set to your email
const SHEET_NAME = 'Sheet1'; // Change this to match your sheet name
```

### Step 5: Save the Script

1. Click the disk icon or press Ctrl+S (Cmd+S on Mac) to save
2. Give your project a name, such as "AI Interview Report Generator"

### Step 6: Authorize the Script

1. Click the "Run" button (play icon) to run the `onOpen` function
2. Google will prompt you to authorize the script
3. Click "Review permissions"
4. Select your Google account
5. Click "Allow" to grant the script permission to access your Google Sheet and send emails

### Step 7: Refresh Your Google Sheet

1. Go back to your Google Sheet and refresh the page
2. You'll see a new menu item called "AI Interview Report" in the top menu

## Using the Report Generator

### Generate a Report On-Demand

1. In your Google Sheet, click on "AI Interview Report" in the top menu
2. Select "Generate & Send Report Now"
3. The script will run and send the report to your email address

### Schedule Daily Reports

1. In your Google Sheet, click on "AI Interview Report" in the top menu
2. Select "Schedule Daily Report"
3. The script will create a trigger to send the report daily at 8 AM

## Troubleshooting

If you encounter any issues:

1. Check the Apps Script logs:
   - In the Apps Script editor, click on "Execution log" to see any error messages

2. Common issues:
   - Column names don't match: Make sure your column headers exactly match the expected names
   - Headers not in second row: The script expects headers in row 2 (the second row)
   - Permission issues: Make sure you've authorized the script to access your sheet and send emails

## Customization

You can customize the report by modifying the following parts of the code:

- The `createHtmlReport` function to change the HTML formatting of the email
- The time of the daily report in the `createDailyTrigger` function
- Add additional metrics or analysis in the `generateReportData` function 