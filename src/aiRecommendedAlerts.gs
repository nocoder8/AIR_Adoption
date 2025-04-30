/**
 * @OnlyCurrentDoc
 *
 * Sends an email alert when a candidate's Feedback_status changes to 'AI_RECOMMENDED'.
 * Designed to be run on a time-driven trigger (e.g., every 15 minutes).
 */

// --- Constants ---
const SPREADSHEET_ID_FBALERTS = '1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE';
const SHEET_NAME_FBALERTS = 'Log_Enhanced';
const TARGET_STATUS_FBALERTS = 'AI_RECOMMENDED';
const RECIPIENT_EMAIL_FBALERTS = 'pkumar@eightfold.ai,akashyap@eightfold.ai';
const ALERT_SENT_PROP_KEY_FBALERTS = 'alertSentProfileIds_AIRecommended'; // New key for storing IDs where alert was sent

// --- Column Names ---
// !! IMPORTANT: Adjust these if the column names in your Sheet change !!
const COL_PROFILE_ID_FBALERTS = 'Profile_id';
const COL_FEEDBACK_STATUS_FBALERTS = 'Feedback_status';
const COL_CANDIDATE_NAME_FBALERTS = 'Candidate_name';
const COL_PROFILE_LINK_FBALERTS = 'Profile_link';
const COL_CURRENT_COMPANY_FBALERTS = 'Current_company';
const COL_POSITION_NAME_FBALERTS = 'Position_name';
const COL_CREATOR_USER_ID_FBALERTS = 'Creator_user_id';
const COL_SCHEDULE_START_TIME_FBALERTS = 'Schedule_start_time';
const COL_DURATION_MINUTES_FBALERTS = 'Duration_minutes';
const COL_LOCATION_COUNTRY_FBALERTS = 'Location_country';
const COL_JOB_FUNCTION_FBALERTS = 'Job_function';
const COL_HIRING_MANAGER_NAME_FBALERTS = 'Hiring_manager_name';
const COL_RECRUITER_NAME_FBALERTS = 'Recruiter_name';
const COL_INVITATION_COMPLETION_DAYS_FBALERTS = 'Invitation_to_completion_days';
const COL_INTERVIEW_STATUS_REAL_FBALERTS = 'Interview_status_real';

// --- Main Function ---

/**
 * Checks the specified sheet for candidates whose Feedback_status has newly
 * changed to AI_RECOMMENDED and sends a summary email.
 */
function checkAiRecommendedStatus() {
  Logger.log('Starting checkAiRecommendedStatus run...');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID_FBALERTS);
    const sheet = ss.getSheetByName(SHEET_NAME_FBALERTS);
    if (!sheet) {
      throw new Error(`Sheet "${SHEET_NAME_FBALERTS}" not found in Spreadsheet ID: ${SPREADSHEET_ID_FBALERTS}`);
    }

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    Logger.log(`Fetched ${allData.length} total rows (including header) from sheet '${SHEET_NAME_FBALERTS}'.`);

    if (allData.length < 2) {
      Logger.log('No data rows found in the sheet. Exiting.');
      return; // No data to process
    }

    const headers = allData[0];
    const colIndices = getColumnIndices(headers); // Map headers to column indices

    // Validate that all required columns exist
    const requiredCols = [
      COL_PROFILE_ID_FBALERTS, COL_FEEDBACK_STATUS_FBALERTS, COL_CANDIDATE_NAME_FBALERTS, COL_PROFILE_LINK_FBALERTS,
      COL_CURRENT_COMPANY_FBALERTS, COL_POSITION_NAME_FBALERTS, COL_CREATOR_USER_ID_FBALERTS, COL_SCHEDULE_START_TIME_FBALERTS,
      COL_DURATION_MINUTES_FBALERTS, COL_LOCATION_COUNTRY_FBALERTS, COL_JOB_FUNCTION_FBALERTS, COL_HIRING_MANAGER_NAME_FBALERTS,
      COL_RECRUITER_NAME_FBALERTS, COL_INVITATION_COMPLETION_DAYS_FBALERTS, COL_INTERVIEW_STATUS_REAL_FBALERTS
    ];
    for (const col of requiredCols) {
      if (colIndices[col] === undefined) {
        // Adjust the error message to show the base name if needed, or keep the suffixed name
        throw new Error(`Required column constant name "${col}" (representing a sheet column) not found mapped in the sheet headers.`);
      }
    }

    // Retrieve IDs for whom the alert has already been sent
    const scriptProperties = PropertiesService.getScriptProperties();
    const alertSentJson = scriptProperties.getProperty(ALERT_SENT_PROP_KEY_FBALERTS);
    // Use a Set for efficient lookup and to automatically handle duplicates if any were ever stored incorrectly
    const alertSentProfileIds = alertSentJson ? new Set(JSON.parse(alertSentJson)) : new Set();
    Logger.log(`Loaded ${alertSentProfileIds.size} profile IDs from alertSentProfileIds property key '${ALERT_SENT_PROP_KEY_FBALERTS}'.`);

    const newlyRecommendedCandidates = [];
    // Use a Set to track IDs flagged *in this specific run* to avoid duplicates from sheet rows
    const profileIdsFlaggedThisRun = new Set();

    // Process data rows (skip header)
    Logger.log('Starting candidate processing loop...');
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const profileId = row[colIndices[COL_PROFILE_ID_FBALERTS]];
      const currentStatus = row[colIndices[COL_FEEDBACK_STATUS_FBALERTS]];

      if (!profileId) {
          // Logger.log(`Row ${i+1}: Skipping row due to missing Profile ID.`); // Optional: Log skipped rows
          continue;
      }

      // Main check: Is status the target AND alert not sent previously AND not already flagged this run?
      const isTargetStatus = currentStatus === TARGET_STATUS_FBALERTS;
      const wasAlertSentPreviously = alertSentProfileIds.has(profileId);
      const alreadyFlaggedThisRun = profileIdsFlaggedThisRun.has(profileId);

      Logger.log(`Row ${i+1}: Profile ID: ${profileId}, Current Status: '${currentStatus}', Target: '${TARGET_STATUS_FBALERTS}'. isTarget=${isTargetStatus}, wasSent=${wasAlertSentPreviously}, flaggedThisRun=${alreadyFlaggedThisRun}`);

      if (isTargetStatus && !wasAlertSentPreviously && !alreadyFlaggedThisRun) {
        Logger.log(`---> Row ${i+1}: Candidate ${profileId} meets criteria. Adding to alert list and marking as flagged for this run.`);
        newlyRecommendedCandidates.push({
            [COL_PROFILE_ID_FBALERTS]: profileId,
            [COL_CANDIDATE_NAME_FBALERTS]: row[colIndices[COL_CANDIDATE_NAME_FBALERTS]],
            [COL_PROFILE_LINK_FBALERTS]: row[colIndices[COL_PROFILE_LINK_FBALERTS]],
            [COL_CURRENT_COMPANY_FBALERTS]: row[colIndices[COL_CURRENT_COMPANY_FBALERTS]],
            [COL_POSITION_NAME_FBALERTS]: row[colIndices[COL_POSITION_NAME_FBALERTS]],
            [COL_CREATOR_USER_ID_FBALERTS]: row[colIndices[COL_CREATOR_USER_ID_FBALERTS]],
            [COL_SCHEDULE_START_TIME_FBALERTS]: row[colIndices[COL_SCHEDULE_START_TIME_FBALERTS]],
            [COL_DURATION_MINUTES_FBALERTS]: row[colIndices[COL_DURATION_MINUTES_FBALERTS]],
            [COL_FEEDBACK_STATUS_FBALERTS]: currentStatus,
            [COL_LOCATION_COUNTRY_FBALERTS]: row[colIndices[COL_LOCATION_COUNTRY_FBALERTS]],
            [COL_JOB_FUNCTION_FBALERTS]: row[colIndices[COL_JOB_FUNCTION_FBALERTS]],
            [COL_HIRING_MANAGER_NAME_FBALERTS]: row[colIndices[COL_HIRING_MANAGER_NAME_FBALERTS]],
            [COL_RECRUITER_NAME_FBALERTS]: row[colIndices[COL_RECRUITER_NAME_FBALERTS]],
            [COL_INVITATION_COMPLETION_DAYS_FBALERTS]: row[colIndices[COL_INVITATION_COMPLETION_DAYS_FBALERTS]],
            [COL_INTERVIEW_STATUS_REAL_FBALERTS]: row[colIndices[COL_INTERVIEW_STATUS_REAL_FBALERTS]],
          });
        profileIdsFlaggedThisRun.add(profileId); // Mark as flagged for *this run*
      }
    }
    Logger.log('Finished candidate processing loop.');

    Logger.log(`Identified ${newlyRecommendedCandidates.length} candidates to alert.`);

    // Send email and update persistent list of sent alerts
    if (newlyRecommendedCandidates.length > 0) {
      sendAlertEmail(newlyRecommendedCandidates);
      Logger.log(`Sent alert for ${newlyRecommendedCandidates.length} newly AI_RECOMMENDED candidates.`);

      // Add the newly alerted IDs to the persistent Set
      Logger.log(`Adding ${profileIdsFlaggedThisRun.size} newly alerted profile IDs to the persistent alertSentProfileIds Set.`);
      profileIdsFlaggedThisRun.forEach(id => alertSentProfileIds.add(id));

      // Store the updated Set back to ScriptProperties
      // Convert Set back to Array for JSON serialization
      const updatedAlertSentArray = Array.from(alertSentProfileIds);
      Logger.log(`Attempting to save ${updatedAlertSentArray.length} total profile IDs to ScriptProperties key '${ALERT_SENT_PROP_KEY_FBALERTS}'...`);
      try {
        scriptProperties.setProperty(ALERT_SENT_PROP_KEY_FBALERTS, JSON.stringify(updatedAlertSentArray));
        Logger.log('Successfully saved updated alertSentProfileIds list.');

         // Optional Verification Step:
         const savedJson = scriptProperties.getProperty(ALERT_SENT_PROP_KEY_FBALERTS);
         const savedSet = savedJson ? new Set(JSON.parse(savedJson)) : new Set();
         Logger.log(`Verification: Saved list contains ${savedSet.size} IDs. Contains ${profileIdsFlaggedThisRun.values().next().value || '(none flagged)'}? ${savedSet.has(profileIdsFlaggedThisRun.values().next().value)}`);

      } catch (propError) {
        Logger.log(`Error saving updated alertSentProfileIds: ${propError}.`);
        // Consider sending an error email here as well, as failure to save will cause future duplicates
         MailApp.sendEmail(RECIPIENT_EMAIL_FBALERTS,
           'AI Recommended Alert - Error Saving Sent IDs',
           `Failed to save the list of profile IDs for whom alerts were sent. Duplicates may occur in the future. Error: ${propError}`
         );
      }
    } else {
      Logger.log('No new candidates met alert criteria. No alert sent. alertSentProfileIds list remains unchanged.');
    }

    Logger.log('checkAiRecommendedStatus run finished successfully.');

  } catch (error) {
    Logger.log(`Error in checkAiRecommendedStatus: ${error.message}\nStack: ${error.stack}`);
    MailApp.sendEmail(RECIPIENT_EMAIL_FBALERTS,
      'Error in AI Recommended Alert Script',
      `The script encountered an error: ${error.message}\n\nPlease check the script logs for details.`
    );
  }
}

// --- Helper Functions ---

/**
 * Maps header names to their column indices (0-based).
 * @param {string[]} headers - The array of header names.
 * @return {Object} An object where keys are header names (from constants) and values are indices.
 */
function getColumnIndices(headers) {
  const headerMap = {};
  Logger.log('Mapping header names to indices...');
  headers.forEach((header, index) => {
    if (header) {
      const trimmedHeader = header.trim();
      headerMap[trimmedHeader] = index;
      // Logger.log(`Mapped header '${trimmedHeader}' to index ${index}`); // Optional: very verbose log
    }
  });

  // Map the constant values (which are the actual header names) to indices
  const indices = {
    [COL_PROFILE_ID_FBALERTS]: headerMap[COL_PROFILE_ID_FBALERTS],
    [COL_FEEDBACK_STATUS_FBALERTS]: headerMap[COL_FEEDBACK_STATUS_FBALERTS],
    [COL_CANDIDATE_NAME_FBALERTS]: headerMap[COL_CANDIDATE_NAME_FBALERTS],
    [COL_PROFILE_LINK_FBALERTS]: headerMap[COL_PROFILE_LINK_FBALERTS],
    [COL_CURRENT_COMPANY_FBALERTS]: headerMap[COL_CURRENT_COMPANY_FBALERTS],
    [COL_POSITION_NAME_FBALERTS]: headerMap[COL_POSITION_NAME_FBALERTS],
    [COL_CREATOR_USER_ID_FBALERTS]: headerMap[COL_CREATOR_USER_ID_FBALERTS],
    [COL_SCHEDULE_START_TIME_FBALERTS]: headerMap[COL_SCHEDULE_START_TIME_FBALERTS],
    [COL_DURATION_MINUTES_FBALERTS]: headerMap[COL_DURATION_MINUTES_FBALERTS],
    [COL_LOCATION_COUNTRY_FBALERTS]: headerMap[COL_LOCATION_COUNTRY_FBALERTS],
    [COL_JOB_FUNCTION_FBALERTS]: headerMap[COL_JOB_FUNCTION_FBALERTS],
    [COL_HIRING_MANAGER_NAME_FBALERTS]: headerMap[COL_HIRING_MANAGER_NAME_FBALERTS],
    [COL_RECRUITER_NAME_FBALERTS]: headerMap[COL_RECRUITER_NAME_FBALERTS],
    [COL_INVITATION_COMPLETION_DAYS_FBALERTS]: headerMap[COL_INVITATION_COMPLETION_DAYS_FBALERTS],
    [COL_INTERVIEW_STATUS_REAL_FBALERTS]: headerMap[COL_INTERVIEW_STATUS_REAL_FBALERTS]
  };

  // Optional: Log if any constant didn't find a matching header
  for (const key in indices) {
    if (indices[key] === undefined) {
        Logger.log(`Warning: Header column corresponding to constant ${key} ('${key.replace('_FBALERTS', '')}') not found in sheet.`);
    }
  }
  Logger.log('Finished mapping header names.');
  return indices;
}

/**
 * Formats and sends the alert email.
 * @param {Object[]} candidates - An array of candidate data objects.
 */
function sendAlertEmail(candidates) {
  // Updated subject line
  const subject = `Alert: ${candidates.length} Candidate(s) newly completed AI Screening`;
  Logger.log(`Preparing to send alert email with subject: "${subject}"`);

  // --- Build HTML Body ---
  let htmlBody = `
    <html>
      <head>
        <style>
          body { font-family: sans-serif; }
          table { border-collapse: collapse; margin-bottom: 15px; width: 95%; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; color: #333; }
          .candidate-header { background-color: #4CAF50; color: white; font-size: 1.1em; padding: 10px; }
          .label { font-weight: bold; color: #555; }
        </style>
      </head>
      <body>
        <h2>AI Screening Completion Alert</h2>
        <p>The following ${candidates.length} candidate(s) have recently had their Feedback_status updated to <strong>${TARGET_STATUS_FBALERTS}</strong>:</p>
  `;

  candidates.forEach((candidate, index) => {
    htmlBody += `
        <hr>
        <h3>Candidate ${index + 1}</h3>
        <table>
          <tr><td class="label">Profile ID:</td><td>${candidate[COL_PROFILE_ID_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Name:</td><td>${candidate[COL_CANDIDATE_NAME_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Profile Link:</td><td>${candidate[COL_PROFILE_LINK_FBALERTS] ? '<a href="' + candidate[COL_PROFILE_LINK_FBALERTS] + '">Link</a>' : 'N/A'}</td></tr>
          <tr><td class="label">Position:</td><td>${candidate[COL_POSITION_NAME_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Current Company:</td><td>${candidate[COL_CURRENT_COMPANY_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Job Function:</td><td>${candidate[COL_JOB_FUNCTION_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Location:</td><td>${candidate[COL_LOCATION_COUNTRY_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Status:</td><td><strong>${candidate[COL_FEEDBACK_STATUS_FBALERTS]}</strong></td></tr>
          <tr><td class="label">Hiring Manager:</td><td>${candidate[COL_HIRING_MANAGER_NAME_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Recruiter:</td><td>${candidate[COL_RECRUITER_NAME_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Duration (min):</td><td>${candidate[COL_DURATION_MINUTES_FBALERTS] || 'N/A'}</td></tr>
          <tr><td class="label">Creator User ID:</td><td>${candidate[COL_CREATOR_USER_ID_FBALERTS] || 'N/A'}</td></tr>
          </table>
      `;
      // Removed: Schedule Start, Invite to Completion Days, Interview Status Real
  });

  htmlBody += `
        <hr>
        <p style="font-size: 0.9em; color: #777;">
          Checked at: ${new Date().toLocaleString()}<br>
          Sheet: ${SHEET_NAME_FBALERTS}
        </p>
      </body>
    </html>
  `;

  // --- Send Email with HTML Body ---
  try {
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL_FBALERTS,
      subject: subject,
      // body: plainTextBody, // Optional: Keep a plain text version for non-HTML clients
      htmlBody: htmlBody, // Use the htmlBody option
    });
    Logger.log(`Successfully sent alert email to ${RECIPIENT_EMAIL_FBALERTS}.`);
  } catch (mailError) {
    Logger.log(`Failed to send email: ${mailError}`);
  }
}

// --- Trigger Setup Reminder ---
// Remember to set up a time-driven trigger for the 'checkAiRecommendedStatus' function
// in the Apps Script editor (Triggers -> Add Trigger). Recommended: Every 15 minutes. 