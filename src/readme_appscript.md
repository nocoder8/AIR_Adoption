The Dasthub SQL query

SELECT b.profile_id,b.name,b.current_company,b.location_city AS candidate_city,b.location_state AS candidate_state,CONCAT('https://app.eightfold.ai/profile/',b.profile_id) AS profile_link,application_ts,business_unit,job_function,last_stage,last_stage_ts,DATEDIFF(day,last_stage_ts,CURRENT_DATE) AS aging,d.title,d.location,d.location_country,d.recruiter_name AS "Recruiter Name",d.recruiter AS "Recruiter Email",d.hiring_manager_name AS "Hiring Manager Name",TO_CHAR(d.position_approved_at,'DD-Mon-YYYY') AS "Position Approved Date",source_name,a.match_stars,a.status AS application_status,d.status AS position_status,d.position_id AS position_id,d.org_units_json.sub_department.title::VARCHAR AS "Sub department",CASE WHEN EXISTS (SELECT 1 FROM volkscience_user_calendar_events ue WHERE ue.profile_id=b.profile_id AND ue.interview_slots_json[0].interviewers[0].is_ai_interviewer=True) THEN 'Y' ELSE 'N' END AS "AI_Interview" FROM (SELECT profile_id,application_id,application_ts,last_stage,last_stage_ts,position_id,is_referral,source_name,match_stars,status FROM volkscience_applications WHERE group_id='volkscience.com') AS a LEFT JOIN (SELECT profile_id,first_name||' '||last_name AS name,current_company,location_city,location_state FROM volkscience_profile WHERE group_id='volkscience.com') AS b ON a.profile_id=b.profile_id LEFT JOIN (SELECT position_id,is_open,title,location,location_country,business_unit,job_function,recruiter_name,recruiter,hiring_manager_name,position_approved_at,org_units_json,status FROM volkscience_positions WHERE group_id='volkscience.com') AS d ON a.position_id=d.position_id WHERE d.is_open='True' AND last_stage IN('Final Interview','Recruiter Screen','HM Quick Feedback','Hiring Manager Screen','Assessment','Onsite Interview','Reference Check','Offer','New','Offer Approvals','Offer Extended','Offer Declined','Pending Start') AND application_ts>=CURRENT_DATE-INTERVAL'180 days' AND d.position_approved_at>='2024-06-01';



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