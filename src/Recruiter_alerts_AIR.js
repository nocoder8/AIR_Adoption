// Recruiter Alerts for AI Recommendations - AIR Script v1.0
// To: Recruiters
// When: Tuesdays thru Fridays, around 1 PM
// What: Checks the Log_Enhanced sheet for completed AI interviews
// where feedback is AI_RECOMMENDED and hasn't been reviewed promptly ( > 1 day ago),
// and sends a consolidated alert email to the relevant recruiter.
// Also sends an Admin Digest email to pkumar@eightfold.ai on Tuesdays thru Fridays, around 1 PM
// containing all pending items from the Log_Enhanced sheet.
// --- Configuration ---
// Use the same Log Enhanced sheet URL as the other scripts
const ALERT_LOG_SHEET_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE/edit'; // <<< VERIFY SPREADSHEET URL
const ALERT_LOG_SHEET_NAME = 'Log_Enhanced'; // <<< VERIFY SHEET NAME
// --- Configuration for Application Sheet (from AIR_Gsheets.js) ---
const ALERT_APP_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw/edit?gid=1957093905#gid=1957093905'; // <<< VERIFY APP SHEET URL
const ALERT_APP_SHEET_NAME = 'Active+Rejected'; // <<< VERIFY APP SHEET NAME
// >>> IMPORTANT: Please double-check this LAUNCH_DATE is correct for your needs <<< 
const ALERT_LAUNCH_DATE = new Date('2025-04-17'); // <<< VERIFY LAUNCH DATE (Needs to match AIR_Gsheets.js)

// Looker Studio URL from AIR_Gsheets.js
const ALERT_LOOKER_STUDIO_URL = 'https://lookerstudio.google.com/reporting/b05c1dfb-d808-4eca-b70d-863fe5be0f27'; // <<< VERIFY LOOKER URL

// Status values to check for
const ALERT_STATUS_COMPLETED = 'COMPLETED';
const ALERT_FEEDBACK_AI_RECOMMENDED = 'AI_RECOMMENDED';
const ALERT_DAYS_THRESHOLD = 1; // Alert if Time_since > 1 day
const ALERT_URGENT_DAYS_THRESHOLD = 3; // Highlight urgency if Time_since > 3 days
const ALERT_STOP_DAYS_THRESHOLD = 15;   // Stop alerting if Time_since > 15 days

// Admin Email for Digest
const ALERT_ADMIN_EMAIL = 'pkumar@eightfold.ai'; // <<< VERIFY ADMIN EMAIL

// --- Main Function & Trigger ---

/**
 * Creates a trigger to run the alert check daily on weekdays.
 */
function createAlertTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendRecruiterAlertsForFeedbackSubmission') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create new triggers to run weekdays (e.g., around 1 PM)
  const daysOfWeek = [
      // ScriptApp.WeekDay.MONDAY, // Removed Monday
      ScriptApp.WeekDay.TUESDAY,
      ScriptApp.WeekDay.WEDNESDAY,
      ScriptApp.WeekDay.THURSDAY,
      ScriptApp.WeekDay.FRIDAY
  ];

  daysOfWeek.forEach(day => {
      ScriptApp.newTrigger('sendRecruiterAlertsForFeedbackSubmission')
          .timeBased()
          .onWeekDay(day)
          .atHour(13).nearMinute(0) // Run around 1 PM
          .create();
  });

  Logger.log(`Weekday triggers created for sendRecruiterAlertsForFeedbackSubmission (around 1 PM)`);
  try {
    // Try to show alert, may fail if run from script editor directly
    SpreadsheetApp.getUi().alert(`Weekday triggers created for Recruiter Alerts (Tuesday-Friday, around 1 PM).`);
  } catch (uiError) {
    Logger.log('Could not display UI alert for trigger creation.');
  }
}

/**
 * Main function to identify candidates needing review and send alerts.
 */
function sendRecruiterAlertsForFeedbackSubmission() {
  Logger.log(`--- Starting Recruiter Alert Check ---`);
  try {
    // 1. Get Log Sheet Data
    const logData = getLogDataForAlerts();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log('No data found in the log sheet or required columns missing. Skipping alerts.');
      return;
    }
    Logger.log(`Successfully retrieved ${logData.rows.length} rows from log sheet.`);

    // 1b. Get Application Sheet Data (for new candidate count)
    const appData = getApplicationDataForAlerts();
    let appRows = [];
    if (appData && appData.rows) {
        appRows = appData.rows;
        Logger.log(`Successfully retrieved ${appRows.length} rows from application sheet.`);
    } else {
        Logger.log('Could not retrieve data from application sheet. New candidate count will be unavailable.');
        // Continue without this data if it fails
    }

    // 2a. Filter out specific Position Names before deduplication
    const positionNameIndex = logData.colIndices['Position_name']; // Already required
    const positionToExclude = "AIR Testing";
    let preFilteredData = logData.rows;
    if (positionNameIndex !== -1) {
        const initialCount = preFilteredData.length;
        preFilteredData = preFilteredData.filter(row => {
            return !(row.length > positionNameIndex && row[positionNameIndex] === positionToExclude);
        });
        Logger.log(`Filtered out ${initialCount - preFilteredData.length} rows with Position_name '${positionToExclude}'. Count before dedupe: ${preFilteredData.length}`);
    } else {
        Logger.log("Position_name column index not found, skipping position filter."); // Should not happen if required
    }

    // 2. Deduplicate by Profile_id + Position_id, prioritizing by status rank
    const deduplicatedData = deduplicateLogData(preFilteredData, logData.colIndices);
    Logger.log(`Deduplicated data count: ${deduplicatedData.length}.`);

    // 3. Filter for Alert Conditions
    const candidatesToAlert = filterCandidatesForAlert(deduplicatedData, logData.colIndices);
    Logger.log(`Found ${candidatesToAlert.length} candidates meeting alert criteria.`);

    if (candidatesToAlert.length === 0) {
      Logger.log('No candidates require alerts today.');
      Logger.log(`--- Finished Recruiter Alert Check (No Alerts Sent) ---`);
      return;
    }

    // 4. Group Candidates by Recruiter (Creator_user_id)
    const groupedByRecruiter = groupCandidatesByRecruiter(candidatesToAlert, logData.colIndices);
    const recruiterEmails = Object.keys(groupedByRecruiter);
    Logger.log(`Grouped candidates for ${recruiterEmails.length} recruiters.`);

    // 5. Send Emails
    // 5a. Send Individual Recruiter Emails
    let emailsSent = 0;
    recruiterEmails.forEach(recruiterEmail => {
      const candidatesForRecruiter = groupedByRecruiter[recruiterEmail];
      if (candidatesForRecruiter && candidatesForRecruiter.length > 0) {
        // Calculate count of NEW, >=4*, Active candidates for this recruiter
        let newHighMatchCount = 0;
        if (appData && appData.colIndices) { // Check if app data retrieval was successful
            newHighMatchCount = countNewHighMatchCandidates(recruiterEmail, appRows, appData.colIndices);
        }

        // Calculate count of PENDING candidates needing nudge for this recruiter
        let pendingNudgeCount = 0;
        // Use deduplicatedData here as it contains the necessary status and email sent date
        // and we only want to count each unique pending profile+position once.
        pendingNudgeCount = countPendingCandidatesToNudge(recruiterEmail, deduplicatedData, logData.colIndices);

        try {
          // Pass the new counts to the HTML function
          const htmlBody = createAlertEmailHtml(recruiterEmail, candidatesForRecruiter, logData.colIndices, newHighMatchCount, pendingNudgeCount);
          const subject = "Urgent: Review AI Screening Feedback"; // Restore original subject
          
          // --- TEMPORARY TESTING MODIFICATION REMOVED ---
          // const testRecipient = 'pkumar@eightfold.ai'; // Override recipient for testing
          // const testSubject = `[TEST For ${recruiterEmail}] ${originalSubject}`; // Add original recipient to subject
          // --- END TEMPORARY TESTING MODIFICATION ---

          // --- TESTING Block Removed ---
          // ... (remains removed)

          // --- Restore original sending logic ---
          MailApp.sendEmail(recruiterEmail, subject, "", { // Use original recruiterEmail and subject
              htmlBody: htmlBody, 
              noReply: true,
              cc: ALERT_ADMIN_EMAIL // Keep CCing the admin
          });
          Logger.log(`Sent alert email to ${recruiterEmail} (CC: ${ALERT_ADMIN_EMAIL}) for ${candidatesForRecruiter.length} candidate(s).`); // Restore original log message
          // --- End Restoration ---
          emailsSent++;
        } catch (emailError) {
          Logger.log(`ERROR sending email to ${recruiterEmail}: ${emailError.toString()}`); // Restore original error log
          // Consider sending a notification to an admin about the failure
          sendAlertErrorNotification(`Failed to send alert email to ${recruiterEmail}`, emailError.stack); // Restore original error notification message
        }
      }
    });

    // 5b. Send Admin Digest Email
    if (ALERT_ADMIN_EMAIL) {
      try {
          const adminHtmlBody = createAdminDigestEmailHtml(candidatesToAlert, logData.colIndices);
          const adminSubject = `Admin Digest: Pending AI Feedback Reviews (${candidatesToAlert.length} items)`;
          MailApp.sendEmail(ALERT_ADMIN_EMAIL, adminSubject, "", { htmlBody: adminHtmlBody, noReply: true });
          Logger.log(`Sent Admin Digest email to ${ALERT_ADMIN_EMAIL} for ${candidatesToAlert.length} total pending items.`);
      } catch (adminEmailError) {
          Logger.log(`ERROR sending Admin Digest email to ${ALERT_ADMIN_EMAIL}: ${adminEmailError.toString()}`);
          // Send separate notification about the digest failure
          sendAlertErrorNotification(`Failed to send Admin Digest email to ${ALERT_ADMIN_EMAIL}`, adminEmailError.stack);
      }
    } else {
        Logger.log('Admin email address not configured. Skipping digest email.');
    }

    Logger.log(`--- Finished Recruiter Alert Check (${emailsSent} Emails Sent) ---`);

  } catch (error) {
    Logger.log(`ERROR in sendRecruiterAlertsForFeedbackSubmission: ${error.toString()} Stack: ${error.stack}`);
    // Optional: Send error notification to admin
     sendAlertErrorNotification(`ERROR in Recruiter Alert Script Execution: ${error.toString()}`, error.stack);
  }
}

// --- Data Retrieval and Processing Functions ---

/**
 * Reads data from the Log_Enhanced sheet for alerts.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null.
 */
function getLogDataForAlerts() {
  Logger.log(`Attempting to open log spreadsheet: ${ALERT_LOG_SHEET_SPREADSHEET_URL}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(ALERT_LOG_SHEET_SPREADSHEET_URL);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening log spreadsheet by URL: ${e}`);
    throw new Error(`Could not open the Log Spreadsheet URL: ${ALERT_LOG_SHEET_SPREADSHEET_URL}`);
  }

  let sheet = spreadsheet.getSheetByName(ALERT_LOG_SHEET_NAME);
  if (!sheet) {
      // Fallback sheet finding logic (like in AIR_Gsheets)
      Logger.log(`Log sheet "${ALERT_LOG_SHEET_NAME}" not found by name. Attempting by gid or first sheet.`);
      const gidMatch = ALERT_LOG_SHEET_SPREADSHEET_URL.match(/gid=(\d+)/);
      if (gidMatch && gidMatch[1]) {
          const gid = gidMatch[1];
          const sheets = spreadsheet.getSheets();
          sheet = sheets.find(s => s.getSheetId().toString() === gid);
          if (sheet) Logger.log(`Using log sheet by ID: "${sheet.getName()}"`);
      }
      if (!sheet) {
          sheet = spreadsheet.getSheets()[0]; // Fallback to first sheet
          if (!sheet) throw new Error(`No sheets found in log spreadsheet: ${ALERT_LOG_SHEET_SPREADSHEET_URL}`);
          Logger.log(`Warning: Log sheet "${ALERT_LOG_SHEET_NAME}" not found. Using first sheet: "${sheet.getName()}"`);
      }
  } else {
     Logger.log(`Using specified log sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) {
    Logger.log(`Not enough data in log sheet "${sheet.getName()}".`);
    return null;
  }

  // Assume headers are in Row 1 (index 0) - aligned with SQL query output
  const headers = data[0].map(String);
  const rows = data.slice(1);

  // Define columns needed for filtering, grouping, and email content
  // Match names based on the SQL query output
  const requiredColumns = [
      'Profile_id',
      'Position_id',
      'interview_status', // From SQL alias (used as fallback or primary if 'Interview Status_Real' isn't there)
      'Feedback_status',
      'Creator_user_id', // Recruiter's email
      'Time_since_interview_completion_days', // Added via SQL
      'Candidate_name',
      'Profile_link', // Corrected case based on user confirmation
      'Current_company', // Added via SQL
      'Position_name',
      'Interview_email_sent_at', // <<< Added for nudge calculation
      'Schedule_start_time' // <<< Added for business days calculation
  ];
  const optionalColumns = [
      'Interview Status_Real' // Check if this exists from potentially different sources
  ];

  const colIndices = {};
  const missingCols = [];

  // --- Determine Status Column to Use ---
  const preferredStatusCol = 'Interview_status_real';
  const fallbackStatusCol = 'interview_status'; // From SQL alias
  let statusColIdx = headers.indexOf(preferredStatusCol);
  let statusColNameUsed = preferredStatusCol;

  if (statusColIdx === -1) {
      Logger.log(`Column "${preferredStatusCol}" not found. Trying fallback "${fallbackStatusCol}".`);
      statusColIdx = headers.indexOf(fallbackStatusCol);
      statusColNameUsed = fallbackStatusCol;
      if (statusColIdx === -1) {
          Logger.log(`Fallback status column "${fallbackStatusCol}" also not found.`);
          missingCols.push(preferredStatusCol + ' or ' + fallbackStatusCol); // Indicate missing status column
      } else {
           Logger.log(`Using fallback status column "${statusColNameUsed}" at index ${statusColIdx}.`);
           colIndices['STATUS_COLUMN'] = statusColIdx;
      }
  } else {
       Logger.log(`Using preferred status column "${statusColNameUsed}" at index ${statusColIdx}.`);
       colIndices['STATUS_COLUMN'] = statusColIdx;
  }
  // Store the name of the column actually used for status checks
  if (statusColIdx !== -1) {
      colIndices['STATUS_COLUMN_NAME'] = headers[statusColIdx];
  }
  // --- End Determine Status Column ---

  // Check remaining required columns
  requiredColumns.forEach(colName => {
      // Skip status columns already checked
      if (colName !== preferredStatusCol && colName !== fallbackStatusCol) {
          const index = headers.indexOf(colName);
          if (index === -1) {
              // Only add if it's truly required and not the status column we couldn't find
              if (!missingCols.includes(preferredStatusCol + ' or ' + fallbackStatusCol)) {
                   missingCols.push(colName);
              }
          } else {
              colIndices[colName] = index;
          }
      }
  });

  // Add optional columns if found
   optionalColumns.forEach(colName => {
       // Ensure we don't add the preferred if it was already found as the main status col
       if (colName !== colIndices['STATUS_COLUMN_NAME']) {
           const index = headers.indexOf(colName);
           if (index !== -1) {
               colIndices[colName] = index;
           }
       }
   });

  if (missingCols.length > 0) {
    // Filter out the generic status message if a specific one wasn't found
    const finalMissing = missingCols.filter(c => c !== preferredStatusCol + ' or ' + fallbackStatusCol);
    // Add the specific status error back if relevant
    if (!colIndices.hasOwnProperty('STATUS_COLUMN')) {
        finalMissing.push(preferredStatusCol + ' or ' + fallbackStatusCol);
    }
    if (finalMissing.length > 0) {
      Logger.log(`ERROR: Missing required column(s) in log sheet "${sheet.getName()}": ${finalMissing.join(', ')}`);
      throw new Error(`Required column(s) not found in log sheet headers (Row 1): ${finalMissing.join(', ')}`);
    }
  }

  Logger.log(`Found required columns for alerts. Indices: ${JSON.stringify(colIndices)}`);
  return { rows, headers, colIndices };
}

/**
 * Deduplicates log data based on Profile_id + Position_id, keeping the row with the best status rank.
 * @param {Array<Array>} rows Raw data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {Array<Array>} Deduplicated rows.
 */
function deduplicateLogData(rows, colIndices) {
    const profileIdIndex = colIndices['Profile_id'];
    const positionIdIndex = colIndices['Position_id'];
    const statusIndex = colIndices['STATUS_COLUMN']; // Use the determined status column index
    const groupedData = {}; // Key: "profileId_positionId", Value: { bestRank: rank, row: rowData }
    let skippedRowCount = 0;

    rows.forEach(row => {
        // Basic check for necessary columns
        if (!row || row.length <= profileIdIndex || row.length <= positionIdIndex || row.length <= statusIndex) {
            skippedRowCount++;
            return; // Skip incomplete rows
        }
        const profileId = row[profileIdIndex];
        const positionId = row[positionIdIndex];
        const status = row[statusIndex] ? String(row[statusIndex]).trim() : null;

        if (!profileId || !positionId) { // Check for blank IDs
             skippedRowCount++;
            return; // Skip rows with blank IDs
        }

        const uniqueKey = `${profileId}_${positionId}`;
        const currentRank = vsGetStatusRank(status); // Use the helper function

        if (!groupedData[uniqueKey] || currentRank < groupedData[uniqueKey].bestRank) {
            groupedData[uniqueKey] = { bestRank: currentRank, row: row };
        }
    });

    if (skippedRowCount > 0) {
        Logger.log(`Skipped ${skippedRowCount} rows during deduplication due to missing IDs, status, or incomplete row data.`);
    }

    return Object.values(groupedData).map(entry => entry.row);
}

/**
 * Filters deduplicated candidates based on alert criteria using business days.
 * @param {Array<Array>} deduplicatedRows Deduplicated rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {Array<Array>} Rows meeting alert criteria.
 */
function filterCandidatesForAlert(deduplicatedRows, colIndices) {
  const statusIdx = colIndices['STATUS_COLUMN'];
  const feedbackStatusIdx = colIndices['Feedback_status'];
  const completionTimeIdx = colIndices['Schedule_start_time']; // Use Schedule_start_time
  const candidateNameIdx = colIndices['Candidate_name']; // Get candidate name index
  // Keep timeSinceIdx only as a fallback for logging/display if needed, not for filtering logic
  // const timeSinceIdx = colIndices['Time_since_interview_completion_days']; 

  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to start of day for consistent comparison

  return deduplicatedRows.filter(row => {
    // Ensure all required columns exist in the row
    if (!row || row.length <= statusIdx || row.length <= feedbackStatusIdx || row.length <= completionTimeIdx || row.length <= candidateNameIdx) {
       Logger.log(`Skipping row in filter due to missing required indices. Row length: ${row ? row.length : 'N/A'}`);
       return false;
    }

    const status = String(row[statusIdx] || '').trim();
    const feedbackStatus = String(row[feedbackStatusIdx] || '').trim();
    const candidateName = String(row[candidateNameIdx] || '').trim();
    const completionDate = parseDateSafe(row[completionTimeIdx]); // Use helper

    if (!completionDate) {
        Logger.log(`Skipping row for Profile ID ${row[colIndices['Profile_id']]} due to invalid Schedule_start_time: ${row[completionTimeIdx]}`);
        return false; // Cannot calculate business days without a valid completion date
    }

    // Calculate business days difference
    const businessDaysSinceCompletion = calculateBusinessDaysDifference(completionDate, today);

    const meetsCriteria = 
      status === ALERT_STATUS_COMPLETED &&
      feedbackStatus === ALERT_FEEDBACK_AI_RECOMMENDED &&
      businessDaysSinceCompletion > ALERT_DAYS_THRESHOLD && // Use business days
      businessDaysSinceCompletion <= ALERT_STOP_DAYS_THRESHOLD; // Use business days

    // Log the business day calculation
    if (status === ALERT_STATUS_COMPLETED && feedbackStatus === ALERT_FEEDBACK_AI_RECOMMENDED) {
        Logger.log(`Check: ProfileID=${row[colIndices['Profile_id']]}, Status=${status}, Feedback=${feedbackStatus}, CompletionDate=${completionDate.toISOString()}, BusinessDays=${businessDaysSinceCompletion}, MeetsCriteria=${meetsCriteria} (Thresholds: >${ALERT_DAYS_THRESHOLD}, <=${ALERT_STOP_DAYS_THRESHOLD})`);
    }

    // Exclude specific candidate (moved check here)
    if (meetsCriteria && candidateName.toLowerCase() === 'erica thomas') {
        Logger.log(`Excluding candidate 'Erica Thomas' from alerts.`);
        return false; // Exclude this candidate
    }
    if (meetsCriteria && candidateName.toLowerCase() === 'daniel sheffield') {
        Logger.log(`Excluding candidate 'Daniel Sheffield' from alerts.`);
        return false; // Exclude this candidate
    }

    return meetsCriteria;
  });
}

/**
 * Groups filtered candidates by recruiter email (Creator_user_id).
 * @param {Array<Array>} candidatesToAlert Filtered rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {object} Object where keys are recruiter emails and values are arrays of candidate rows.
 */
function groupCandidatesByRecruiter(candidatesToAlert, colIndices) {
  const grouped = {};
  const recruiterEmailIdx = colIndices['Creator_user_id'];

  candidatesToAlert.forEach(row => {
    if (row && row.length > recruiterEmailIdx) {
        const recruiterEmail = String(row[recruiterEmailIdx] || '').trim().toLowerCase(); // Normalize email
        if (recruiterEmail && recruiterEmail.includes('@')) { // Basic email validation
            if (!grouped[recruiterEmail]) {
                grouped[recruiterEmail] = [];
            }
            grouped[recruiterEmail].push(row);
        } else {
            Logger.log(`Skipping row due to invalid/missing recruiter email: ${row[recruiterEmailIdx]}`);
        }
    }
  });
  return grouped;
}

// --- Email Generation and Sending ---

/**
 * Creates the HTML email body for a recruiter's alert.
 * @param {string} recruiterEmail The recruiter's email address.
 * @param {Array<Array>} candidates The array of candidate rows for this recruiter.
 * @param {object} colIndices Map of column names to indices.
 * @param {number} newHighMatchCount The count of NEW, >=4* candidates.
 * @param {number} pendingNudgeCount The count of PENDING candidates needing nudge.
 * @returns {string} The HTML content for the email body.
 */
function createAlertEmailHtml(recruiterEmail, candidates, colIndices, newHighMatchCount, pendingNudgeCount) {
  // Sort candidates by Business Days Since Completion descending
  const completionTimeIdx = colIndices['Schedule_start_time'];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  candidates.sort((a, b) => {
      const dateA = parseDateSafe(a[completionTimeIdx]);
      const dateB = parseDateSafe(b[completionTimeIdx]);
      const businessDaysA = dateA ? calculateBusinessDaysDifference(dateA, today) : -1; // Handle invalid dates
      const businessDaysB = dateB ? calculateBusinessDaysDifference(dateB, today) : -1;
      return businessDaysB - businessDaysA; // Descending order
  });

  let tableRowsHtml = '';
  candidates.forEach(row => {
      const candidateName = row[colIndices['Candidate_name']] || 'N/A';
      const profileLink = row[colIndices['Profile_link']] || '#'; // Default to '#' if no link
      const currentCompany = row[colIndices['Current_company']] || 'N/A';
      const positionName = row[colIndices['Position_name']] || 'N/A';
      // Calculate business days for display
      const completionDate = parseDateSafe(row[completionTimeIdx]);
      const businessDaysSince = completionDate ? calculateBusinessDaysDifference(completionDate, today) : null;
      const timeSinceDisplay = businessDaysSince !== null ? businessDaysSince : 'N/A'; // Display calculated business days

      // Determine urgency highlighting based on business days
      const isUrgent = businessDaysSince !== null && businessDaysSince > ALERT_URGENT_DAYS_THRESHOLD;
      const rowStyle = isUrgent ? ' style="background-color: #fff0f0; font-weight: bold;"' : ''; // Light red background + bold for urgent

      // Create a clickable link for the profile
      const candidateLinkHtml = profileLink !== '#' ? `<a href="${profileLink}" target="_blank" style="color: #007bff; text-decoration: none;">${candidateName}</a>` : candidateName;

      tableRowsHtml += `
        <tr${rowStyle}>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${candidateLinkHtml}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${currentCompany}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${positionName}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px; text-align: center;">${timeSinceDisplay}</td>
        </tr>
      `;
  });

  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>Urgent: Review AI Screening Feedback</title>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
      .container { padding: 25px; max-width: 750px; margin: 20px auto; background-color: #ffffff; border: 1px solid #cccccc; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      h2 { color: #c9302c; font-size: 20px; margin-top: 0; border-bottom: 2px solid #eeeeee; padding-bottom: 12px; margin-bottom: 20px; } /* Stronger Red */
      p { margin-bottom: 18px; font-size: 15px; }
      ul { margin-top: 5px; margin-bottom: 15px; padding-left: 20px; }
      li { margin-bottom: 8px; font-size: 15px; }
      table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px; }
      th, td { border: 1px solid #ddd; padding: 8px 10px; text-align: left; vertical-align: middle; }
      th { background-color: #f5f5f5; font-weight: bold; text-transform: uppercase; font-size: 11px; color: #555; }
      tr:nth-child(even) { background-color: #fdfdfd; }
      .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #dddddd; font-size: 13px; color: #888; text-align: center; }
      a { color: #007bff; text-decoration: none; }
      a:hover { text-decoration: underline; }
    </style>
  </head>
  <body>
    <div class="container">
      <h2>Urgent: Review the AI generated feedback within 24 business hours</h2>
      <p>Hi,</p>
      <p>The following candidate(s) are waiting for your next steps based on AI Screening feedback. Please review within 24 business hours:</p>

      <table>
        <thead>
          <tr>
            <th>Candidate Name</th>
            <th>Current Company</th>
            <th>Position Name</th>
            <th>Business Days Since Completion</th>
          </tr>
        </thead>
        <tbody>
          ${tableRowsHtml}
        </tbody>
      </table>

      <p><b>Next Steps:</b></p>
      <ul>
        <li>If positive: Inform the candidate using the email template "Positive and will setup next round", as you plan next interviews.</li>
        <li>If negative: Reject and send the standard rejection email from T2.</li>
      </ul>

      <!-- Start Action Items Box -->
      <div style="border: 1px solid #e0e0e0; padding: 15px 20px; margin-top: 20px; background-color: #f9f9f9; border-radius: 5px;">

        ${newHighMatchCount > 0 ? `<p style="font-size: 14px; background-color: #e9f5ff; border: 1px solid #b3daff; padding: 10px 15px; border-radius: 4px; margin-bottom: 15px;"><b>Additionally:</b> You have <b>${newHighMatchCount}</b> candidate(s) in the 'NEW' stage with a Match Score of 4+ applied since ${ALERT_LAUNCH_DATE.toLocaleDateString()} who are awaiting review.</p>` : ''}

        ${pendingNudgeCount > 0 ? `<p style="font-size: 14px; background-color: #fff3cd; border: 1px solid #ffeeba; padding: 10px 15px; border-radius: 4px; color: #856404; margin-bottom: 15px;"><b>Nudge Reminder:</b> You have <b>${pendingNudgeCount}</b> candidate(s) invited more than 2 days ago who are still in 'PENDING' status and may need a follow-up.</p>` : ''}

        <p class="footer" style="margin-top: 10px; padding-top: 0; border-top: none; text-align: center;">
          <a href="${ALERT_LOOKER_STUDIO_URL}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #007bff; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 13px; border-bottom: 2px solid #0056b3; transition: background-color 0.2s ease;">Review all your AIR related action items here</a>
        </p>

      </div>
      <!-- End Action Items Box -->

    </div>
  </body>
  </html>
  `;
  return html;
}

/**
 * Creates the HTML email body for the Admin Digest.
 * @param {Array<Array>} allAlertCandidates The array of all candidate rows meeting alert criteria.
 * @param {object} colIndices Map of column names to indices.
 * @returns {string} The HTML content for the email body.
 */
function createAdminDigestEmailHtml(allAlertCandidates, colIndices) {
  // Sort all candidates by Business Days Since Completion descending
  const completionTimeIdx = colIndices['Schedule_start_time'];
  const recruiterEmailIdx = colIndices['Creator_user_id']; // Added for summary
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // --- Generate Summary Table Data ---
  const recruiterCounts = {};
  allAlertCandidates.forEach(row => {
      if (row && row.length > recruiterEmailIdx) {
          const recruiterEmail = String(row[recruiterEmailIdx] || 'N/A').trim().toLowerCase();
          if (recruiterEmail !== 'n/a' && recruiterEmail.includes('@')) { // Basic check
              recruiterCounts[recruiterEmail] = (recruiterCounts[recruiterEmail] || 0) + 1;
          }
      }
  });

  // Sort recruiters alphabetically for the summary table
  const sortedRecruiters = Object.keys(recruiterCounts).sort();

  let summaryTableRowsHtml = '';
  sortedRecruiters.forEach(email => {
      summaryTableRowsHtml += `
        <tr>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${email}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px; text-align: center;">${recruiterCounts[email]}</td>
        </tr>
      `;
  });
  // --- End Summary Table Data ---

  // --- Generate Detailed Table Data (Sort and Display Business Days) ---
  allAlertCandidates.sort((a, b) => {
      const dateA = parseDateSafe(a[completionTimeIdx]);
      const dateB = parseDateSafe(b[completionTimeIdx]);
      const businessDaysA = dateA ? calculateBusinessDaysDifference(dateA, today) : -1; // Handle invalid dates
      const businessDaysB = dateB ? calculateBusinessDaysDifference(dateB, today) : -1;
      return businessDaysB - businessDaysA; // Descending order (oldest first)
  });

  let tableRowsHtml = '';
  allAlertCandidates.forEach(row => {
      const recruiterEmail = row[recruiterEmailIdx] || 'N/A';
      const candidateName = row[colIndices['Candidate_name']] || 'N/A';
      const profileLink = row[colIndices['Profile_link']] || '#';
      const currentCompany = row[colIndices['Current_company']] || 'N/A';
      const positionName = row[colIndices['Position_name']] || 'N/A';
      // Calculate business days for display
      const completionDate = parseDateSafe(row[completionTimeIdx]);
      const businessDaysSince = completionDate ? calculateBusinessDaysDifference(completionDate, today) : null;
      const timeSinceDisplay = businessDaysSince !== null ? businessDaysSince : 'N/A'; // Display calculated business days

      // Determine urgency highlighting based on business days for admin digest as well
      const isUrgent = businessDaysSince !== null && businessDaysSince > ALERT_URGENT_DAYS_THRESHOLD;
      const rowStyle = isUrgent ? ' style="background-color: #fff0f0; font-weight: bold;"' : ''; // Light red background + bold for urgent

      const candidateLinkHtml = profileLink !== '#' ? `<a href="${profileLink}" target="_blank" style="color: #007bff; text-decoration: none;">${candidateName}</a>` : candidateName;

      tableRowsHtml += `
        <tr${rowStyle}>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${recruiterEmail}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${candidateLinkHtml}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${currentCompany}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px;">${positionName}</td>
          <td style="border: 1px solid #ddd; padding: 6px 8px; font-size: 12px; text-align: center;">${timeSinceDisplay}</td>
        </tr>
      `;
  });

  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>Admin Digest: Pending AI Feedback Reviews</title>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
      .container { padding: 25px; max-width: 900px; margin: 20px auto; background-color: #ffffff; border: 1px solid #cccccc; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      h2 { font-size: 20px; color: #333366; margin-top: 0; border-bottom: 2px solid #eeeeee; padding-bottom: 12px; margin-bottom: 20px; }
      p { margin-bottom: 18px; font-size: 15px; }
      table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px; }
      th, td { border: 1px solid #ddd; padding: 8px 10px; text-align: left; vertical-align: middle; }
      th { background-color: #f5f5f5; font-weight: bold; text-transform: uppercase; font-size: 11px; color: #555; }
      tr:nth-child(even) { background-color: #fdfdfd; }
      .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #dddddd; font-size: 13px; color: #888; text-align: center; }
      a { color: #007bff; text-decoration: none; }
      a:hover { text-decoration: underline; }
    </style>
  </head>
  <body>
    <div class="container">
      <h2>Admin Digest: AI Feedback Awaiting Review (${allAlertCandidates.length} total)</h2>
      <p>This digest lists all candidates whose AI screening feedback is marked '${ALERT_FEEDBACK_AI_RECOMMENDED}' and has been awaiting review for more than ${ALERT_DAYS_THRESHOLD} day(s) but less than or equal to ${ALERT_STOP_DAYS_THRESHOLD} days.</p>

      <h3>Summary by Recruiter</h3>
      <table style="width: auto; margin-bottom: 25px;"> <!-- Adjusted width and margin -->
        <thead>
          <tr>
            <th style="border: 1px solid #ddd; padding: 8px 10px; background-color: #f5f5f5; font-weight: bold; text-transform: uppercase; font-size: 11px; color: #555; text-align: left;">Recruiter Email</th>
            <th style="border: 1px solid #ddd; padding: 8px 10px; background-color: #f5f5f5; font-weight: bold; text-transform: uppercase; font-size: 11px; color: #555; text-align: center;">Pending Items</th>
          </tr>
        </thead>
        <tbody>
          ${summaryTableRowsHtml}
        </tbody>
      </table>

      <h3>Detailed List</h3>
      <table>
        <thead>
          <tr>
            <th>Recruiter Email</th>
            <th>Candidate Name</th>
            <th>Current Company</th>
            <th>Position Name</th>
            <th>Business Days Since Completion</th>
          </tr>
        </thead>
        <tbody>
          ${tableRowsHtml}
        </tbody>
      </table>

      <p class="footer">
        <a href="${ALERT_LOOKER_STUDIO_URL}" target="_blank" style="display: inline-block; margin-bottom: 10px; padding: 10px 20px; background-color: #007bff; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 13px; border-bottom: 2px solid #0056b3; transition: background-color 0.2s ease;">Review all AIR related action items in Looker Studio</a><br>
        Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.
      </p>
    </div>
  </body>
  </html>
  `;
  return html;
}

/**
 * Sends an error notification email TO THE SCRIPT OWNER.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendAlertErrorNotification(errorMessage, stackTrace = '') {
   // Get the script owner's email
   const recipient = Session.getEffectiveUser().getEmail();
   if (!recipient) {
       Logger.log("CRITICAL ERROR: Cannot send error notification because script owner email could not be determined.");
       return;
   }
   try {
       const subject = `ERROR: Recruiter Alert AIR Script Failed - ${new Date().toLocaleString()}`;
       let body = `Error generating/sending Recruiter Alerts:\n\n${errorMessage}\n\n`;
       if (stackTrace) {
           body += `Stack Trace:\n${stackTrace}\n\n`;
       }
       body += `Log Sheet URL: ${ALERT_LOG_SHEET_SPREADSHEET_URL}`;
       MailApp.sendEmail(recipient, subject, body);
       Logger.log(`Error notification email sent to script owner ${recipient}.`);
    } catch (emailError) {
       Logger.log(`CRITICAL: Failed to send error notification email to ${recipient}: ${emailError}`);
    }
}


// --- Helper Functions ---

/**
 * Calculates the number of business days (Mon-Fri) between two dates (exclusive of start, inclusive of end date comparison).
 * @param {Date} startDate The start date.
 * @param {Date} endDate The end date.
 * @returns {number} The number of business days.
 */
function calculateBusinessDaysDifference(startDate, endDate) {
  if (!startDate || !endDate || startDate >= endDate) {
    return 0;
  }

  let count = 0;
  const current = new Date(startDate);
  current.setHours(12, 0, 0, 0); // Normalize time to avoid DST issues

  // Ensure endDate is also normalized for comparison
  const normalizedEndDate = new Date(endDate);
  normalizedEndDate.setHours(12, 0, 0, 0);

  // Iterate day by day
  while (current < normalizedEndDate) {
    current.setDate(current.getDate() + 1);
    const dayOfWeek = current.getDay(); // 0=Sunday, 6=Saturday
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      count++;
    }
  }
  return count;
}

/**
 * Assigns a numerical rank to interview statuses for prioritization during deduplication.
 * Lower rank means higher priority (e.g., Completed is better than Scheduled).
 * Copied from AIR_Volkscience.js
 * @param {string} status The raw interview status string.
 * @returns {number} The rank of the status.
 */
function vsGetStatusRank(status) {
    // Define statuses indicating completion (used for other metrics)
    // IMPORTANT: These statuses should align with the values actually present in the Log_Enhanced sheet
    const COMPLETED_STATUSES_RAW = ['COMPLETED', 'Feedback Provided', 'Pending Feedback', 'No Show'];
    // Define statuses considered "Scheduled"
    const SCHEDULED_STATUS_RAW = 'SCHEDULED';
    // Define statuses considered "Pending"
    const PENDING_STATUSES_RAW = ['PENDING', 'INVITED', 'EMAIL SENT'];

    if (!status) return 99; // Handle null/undefined status
    const trimmedStatus = status.trim();

    if (COMPLETED_STATUSES_RAW.includes(trimmedStatus)) {
        return 1; // Highest priority
    } else if (trimmedStatus === SCHEDULED_STATUS_RAW) {
        return 2;
    } else if (PENDING_STATUSES_RAW.includes(trimmedStatus)) {
        return 3;
    } else {
        return 99; // Lowest priority for anything else (Expired, Cancelled, Unknown etc.)
    }
}

/**
 * Creates menu items in the Google Sheet UI (when script is opened from a Sheet).
 * Note: Having multiple onOpen functions in one project can be problematic.
 * If this script is associated with a sheet, this adds menu items.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Recruiter Alerts (AIR)')
      .addItem('Run Alert Check Now', 'sendRecruiterAlertsForFeedbackSubmission')
      .addItem('Setup/Reset Daily Trigger', 'createAlertTrigger')
      .addToUi();
  } catch (e) {
    // Log error but don't prevent script from running if not opened from a Sheet
    Logger.log("Error creating Recruiter Alerts menu (might happen if not opened from a Sheet): " + e);
  }
}

// Note: The parseDateSafe function from AIR_GSheets isn't strictly necessary here
// as we are reading the pre-calculated 'Time_since_interview_completion_days'.
// Keeping vsGetStatusRank as it's essential for deduplication. 

// --- Application Data Retrieval ---
/**
 * Reads data from the Application sheet (e.g., Active+Rejected) for context.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null.
 */
function getApplicationDataForAlerts() {
  Logger.log(`Attempting to open application spreadsheet: ${ALERT_APP_SHEET_URL}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(ALERT_APP_SHEET_URL);
    Logger.log(`Opened application spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening application spreadsheet by URL: ${e}`);
    // Return null instead of throwing, so main script can continue
    sendAlertErrorNotification(`Could not open the Application Spreadsheet URL: ${ALERT_APP_SHEET_URL}. New candidate counts will be missing.`, e.stack);
    return null;
  }

  let sheet = spreadsheet.getSheetByName(ALERT_APP_SHEET_NAME);
  if (!sheet) {
      Logger.log(`App sheet "${ALERT_APP_SHEET_NAME}" not found. Trying by GID or first sheet.`);
      const gidMatch = ALERT_APP_SHEET_URL.match(/gid=(\d+)/);
      if (gidMatch && gidMatch[1]) {
          const gid = gidMatch[1];
          const sheets = spreadsheet.getSheets();
          sheet = sheets.find(s => s.getSheetId().toString() === gid);
          if (sheet) Logger.log(`Using app sheet by ID: "${sheet.getName()}"`);
      }
      if (!sheet) {
          sheet = spreadsheet.getSheets()[0];
          if (!sheet) {
              Logger.log(`No sheets found in application spreadsheet: ${ALERT_APP_SHEET_URL}`);
              sendAlertErrorNotification(`No sheets found in application spreadsheet: ${ALERT_APP_SHEET_URL}. New candidate counts will be missing.`);
              return null;
          }
          Logger.log(`Warning: App sheet "${ALERT_APP_SHEET_NAME}" not found. Using first sheet: "${sheet.getName()}"`);
      }
  } else {
     Logger.log(`Using specified app sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Expect headers in Row 2, data starts Row 3 (like AIR_Gsheets)
  if (data.length < 3) {
    Logger.log(`Not enough data in app sheet "${sheet.getName()}" (expected headers in row 2).`);
    return null;
  }

  const headers = data[1].map(String); // Headers from Row 2
  const rows = data.slice(2); // Data from Row 3 onwards

  // Define columns needed for the new candidate count
  // IMPORTANT: Using exact names here. Add fuzzy matching if needed later.
  const requiredAppColumns = [
      'Recruiter email',  // <<< Corrected Column Name
      'Match_stars',      // <<< VERIFY THIS EXACT HEADER NAME
      'Last_stage',       // <<< VERIFY THIS EXACT HEADER NAME
      'Application_status',
      'Position_status',
      'Application_ts'
  ];

  const appColIndices = {};
  const missingAppCols = [];

  requiredAppColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index === -1) {
          missingAppCols.push(colName);
      } else {
          appColIndices[colName] = index;
      }
  });

  if (missingAppCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in app sheet "${sheet.getName()}": ${missingAppCols.join(', ')}`);
    sendAlertErrorNotification(`Required column(s) for new candidate count not found in app sheet headers (Row 2): ${missingAppCols.join(', ')}. New candidate counts will be missing.`);
    return null; // Return null so main script can proceed without this data
  }

  Logger.log(`Found required columns for app data context. Indices: ${JSON.stringify(appColIndices)}`);
  return { rows, headers, colIndices: appColIndices };
}

/**
 * Counts candidates from application data meeting specific criteria for a given recruiter.
 * @param {string} recruiterEmail The email of the recruiter to filter by.
 * @param {Array<Array>} appRows Rows from the application sheet.
 * @param {object} appColIndices Column indices for the application sheet.
 * @returns {number} The count of matching candidates.
 */
function countNewHighMatchCandidates(recruiterEmail, appRows, appColIndices) {
  let count = 0;
  const recruiterIdx = appColIndices['Recruiter email']; // <<< Corrected Column Name
  const matchStarsIdx = appColIndices['Match_stars'];
  const lastStageIdx = appColIndices['Last_stage'];
  const appStatusIdx = appColIndices['Application_status'];
  const posStatusIdx = appColIndices['Position_status'];
  const appTsIdx = appColIndices['Application_ts'];

  appRows.forEach(row => {
    // Basic check for row validity and required columns
    if (!row || row.length <= Math.max(recruiterIdx, matchStarsIdx, lastStageIdx, appStatusIdx, posStatusIdx, appTsIdx)) {
      return; // Skip incomplete rows
    }

    const rowRecruiter = String(row[recruiterIdx] || '').trim().toLowerCase();
    const matchStars = parseFloat(row[matchStarsIdx]);
    const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase(); // Normalize to uppercase
    const appStatus = String(row[appStatusIdx] || '').trim().toLowerCase();
    const posStatus = String(row[posStatusIdx] || '').trim().toLowerCase();
    const appTs = parseDateSafe(row[appTsIdx]); // Use helper function

    if (rowRecruiter === recruiterEmail.toLowerCase() &&
        !isNaN(matchStars) && matchStars >= 4 &&
        lastStage === 'NEW' &&
        appStatus === 'active' &&
        posStatus === 'open' &&
        appTs && appTs >= ALERT_LAUNCH_DATE) {
      count++;
    }
  });
  Logger.log(`Found ${count} NEW, >=4*, Active, Post-Launch candidates for ${recruiterEmail}`);
  return count;
}

/**
 * Counts candidates from log data needing a nudge (Pending > 2 days) for a given recruiter.
 * Uses the deduplicated data to avoid multiple counts for the same profile+position.
 * @param {string} recruiterEmail The email of the recruiter to filter by.
 * @param {Array<Array>} deduplicatedLogRows Deduplicated rows from the log sheet.
 * @param {object} logColIndices Column indices for the log sheet.
 * @returns {number} The count of matching candidates.
 */
function countPendingCandidatesToNudge(recruiterEmail, deduplicatedLogRows, logColIndices) {
  let nudgeCount = 0;
  const recruiterIdx = logColIndices['Creator_user_id'];
  const statusIdx = logColIndices['STATUS_COLUMN']; // Use the determined status column index
  const emailSentIdx = logColIndices['Interview_email_sent_at'];
  // Position name filter is already applied before deduplication in the main flow

  const twoDaysInMillis = 2 * 24 * 60 * 60 * 1000;
  const now = new Date().getTime();

  // Status value to check for 'Pending' or similar invite statuses
  // Using 'PENDING' based on user request, adjust if needed
  const PENDING_STATUS_CHECK = 'PENDING'; 

  deduplicatedLogRows.forEach(row => {
    // Basic check for row validity and required columns
    if (!row || row.length <= Math.max(recruiterIdx, statusIdx, emailSentIdx)) {
      return; // Skip incomplete rows
    }

    const rowRecruiter = String(row[recruiterIdx] || '').trim().toLowerCase();
    const status = String(row[statusIdx] || '').trim().toUpperCase(); // Normalize to uppercase
    const emailSentDate = parseDateSafe(row[emailSentIdx]); // Use helper function

    if (rowRecruiter === recruiterEmail.toLowerCase() &&
        status === PENDING_STATUS_CHECK &&
        emailSentDate) {
          const timeSinceSent = now - emailSentDate.getTime();
          if (timeSinceSent > twoDaysInMillis) {
            nudgeCount++;
          }
    }
  });

  if (nudgeCount > 0) {
      Logger.log(`Found ${nudgeCount} PENDING (>2 days) candidates needing nudge for ${recruiterEmail}`);
  }
  return nudgeCount;
}

// Helper function to safely parse dates (reusing from AIR_Gsheets logic if needed)
function parseDateSafe(dateInput) {
    if (dateInput === null || dateInput === undefined || dateInput === '') {
        return null;
    }
    // Handle potential Google Sheet date serial numbers (if applicable)
    if (typeof dateInput === 'number' && dateInput > 10000) { // Heuristic for date serial number
       try {
           // Convert Excel/Sheets serial number (days since Dec 30, 1899) to JS timestamp (ms since Jan 1, 1970)
           const jsTimestamp = (dateInput - 25569) * 86400 * 1000;
           const date = new Date(jsTimestamp);
            return !isNaN(date.getTime()) ? date : null;
       } catch (e) { /* Ignore conversion error, proceed to standard parsing */ }
    }
    // Standard Date parsing
    const date = new Date(dateInput);
    return !isNaN(date.getTime()) ? date : null;
} 