// AIR Volkscience - Exec Summary - Company-Level AI Interview Analytics Script v1.0
// To: Akhila and Pavan
// When: Weekly, Monday at 8 AM
// This script analyzes data from the Log_Enhanced sheet to provide company-wide insights
// into the AI interview process funnel, timelines, and outcomes.

// --- Configuration ---
const VS_EMAIL_RECIPIENT = 'akashyap@eightfold.ai'; // <<< UPDATE EMAIL RECIPIENT
const VS_EMAIL_CC = 'pkumar@eightfold.ai'; // Optional CC
// Assuming the Log Enhanced sheet is in a separate Spreadsheet
const VS_LOG_SHEET_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE/edit'; // <<< VERIFY SPREADSHEET URL
const VS_LOG_SHEET_NAME = 'Log_Enhanced'; // <<< VERIFY SHEET NAME
const VS_REPORT_TIME_RANGE_DAYS = 99999; // Set large number to effectively include all time
const VS_COMPANY_NAME = "Eightfold"; // Used in report titles etc.


// --- Main Functions ---

/**
 * Creates a trigger to run the report weekly.
 */
function createVolkscienceTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'AIR_ExecSummary_Daily') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create a new trigger to run weekly (e.g., Monday at 8 AM)
  ScriptApp.newTrigger('AIR_ExecSummary_Daily')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  Logger.log(`Weekly trigger created for AIR_ExecSummary_Daily (Monday 10 AM)`);
  SpreadsheetApp.getUi().alert(`Weekly trigger created for ${VS_COMPANY_NAME} AI Interview Report (Monday 10 AM).`);
}

/**
 * Main function to generate and send the company-level AI interview report.
 */
function AIR_ExecSummary_Daily() {
  try {
    Logger.log(`--- Starting ${VS_COMPANY_NAME} AI Interview Report Generation ---`);

    // 1. Get Log Sheet Data
    const logData = getLogSheetData();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log('No data found in the log sheet or required columns missing. Skipping report generation.');
      // Optional: Send an email notification about missing data/columns
      // sendVsErrorNotification("Report Skipped: No data or required columns found in Log_Enhanced sheet.");
      return;
    }
     Logger.log(`Successfully retrieved ${logData.rows.length} rows from log sheet.`);

    // 2. Filter Data by Time Range (using Interview_email_sent_at)
    const filteredData = filterDataByTimeRange(logData.rows, logData.colIndices);
    if (filteredData.length === 0) {
        Logger.log(`No data found within the last ${VS_REPORT_TIME_RANGE_DAYS} days. Skipping report.`);
        return;
    }
    Logger.log(`Filtered data to ${filteredData.length} rows based on the last ${VS_REPORT_TIME_RANGE_DAYS} days.`);

    // 2b. Filter out specific Position Names
    const positionNameIndex = logData.colIndices.hasOwnProperty('Position_name') ? logData.colIndices['Position_name'] : -1;
    const positionToExclude = "AIR Testing";
    let finalFilteredData = filteredData;
    if (positionNameIndex !== -1) {
        const initialCount = finalFilteredData.length;
        finalFilteredData = finalFilteredData.filter(row => {
            return !(row.length > positionNameIndex && row[positionNameIndex] === positionToExclude);
        });
        Logger.log(`Filtered out ${initialCount - finalFilteredData.length} rows with Position_name '${positionToExclude}'. Final count: ${finalFilteredData.length}`);
    } else {
        Logger.log("Skipping Position_name filter as column was not found.");
    }

    // Check if any data remains after all filters
    if (finalFilteredData.length === 0) {
         Logger.log(`No data remaining after position filtering. Skipping report.`);
         return;
    }

    // 2c. Deduplicate by Profile_id + Position_id, prioritizing by status rank
    const profileIdIndex = logData.colIndices['Profile_id'];
    const positionIdIndex = logData.colIndices['Position_id'];
    const statusIndex = logData.colIndices['STATUS_COLUMN']; // Get the index determined earlier
    const groupedData = {}; // Key: "profileId_positionId", Value: { bestRank: rank, row: rowData }

    finalFilteredData.forEach(row => {
        // Ensure row has the necessary columns
        if (row.length <= profileIdIndex || row.length <= positionIdIndex || row.length <= statusIndex) {
            Logger.log(`Skipping row during grouping due to missing ID or Status columns. Row: ${JSON.stringify(row)}`);
            return; // Skip this row
        }
        const profileId = row[profileIdIndex];
        const positionId = row[positionIdIndex];
        const status = row[statusIndex] ? String(row[statusIndex]).trim() : 'Unknown';

        if (!profileId || !positionId) { // Check for blank IDs
            Logger.log(`Skipping row during grouping due to blank Profile_id or Position_id. Row: ${JSON.stringify(row)}`);
            return; // Skip rows with blank IDs
        }

        const uniqueKey = `${profileId}_${positionId}`;
        const currentRank = vsGetStatusRank(status);

        if (!groupedData[uniqueKey] || currentRank < groupedData[uniqueKey].bestRank) {
            // If no entry exists OR current row has a better (lower) rank, store/replace it
            groupedData[uniqueKey] = { bestRank: currentRank, row: row };
        }
        // If an entry exists and current rank is not better, do nothing (keep the existing better row)
    });

    // Extract the best row for each unique key
    const deduplicatedData = Object.values(groupedData).map(entry => entry.row);

    Logger.log(`Deduplicated data based on Profile_id + Position_id (prioritizing status). Count changed from ${finalFilteredData.length} to ${deduplicatedData.length}.`);

    // Check if any data remains after deduplication
    if (deduplicatedData.length === 0) {
         Logger.log(`No data remaining after deduplication. Skipping report.`);
         return;
    }

    // 3. Calculate Metrics
    const metrics = calculateCompanyMetrics(deduplicatedData, logData.colIndices);
    Logger.log('Successfully calculated company metrics.');
    // Logger.log(`Calculated Metrics: ${JSON.stringify(metrics)}`); // Optional: Log detailed metrics

    // 4. Create HTML Report
    const htmlContent = createVolkscienceHtmlReport(metrics);
    Logger.log('Successfully generated HTML report content.');

    // 5. Send Email
    // Set static subject line as requested
    const reportTitle = `AI Recruiter Adoption: Executive Summary (EF4EF)`;
    sendVolkscienceEmail(VS_EMAIL_RECIPIENT, VS_EMAIL_CC, reportTitle, htmlContent);

    Logger.log(`--- AI Recruiter Adoption: Executive Summary generated and sent successfully! ---`);
    return `Report sent to ${VS_EMAIL_RECIPIENT}`;

  } catch (error) {
    Logger.log(`Error in AIR_ExecSummary_Daily: ${error.toString()} Stack: ${error.stack}`);
    // Send error email
    sendVsErrorNotification(`ERROR generating AI Recruiter Adoption: Executive Summary: ${error.toString()}`, error.stack);
    return `Error: ${error.toString()}`;
  }
}

// --- Data Retrieval and Processing Functions ---

/**
 * Reads and processes data from the Log_Enhanced sheet.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null if error/no sheet/missing columns.
 */
function getLogSheetData() {
  Logger.log(`Attempting to open log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(VS_LOG_SHEET_SPREADSHEET_URL);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening log spreadsheet by URL: ${e}`);
    throw new Error(`Could not open the specified Log Spreadsheet URL. Please verify the URL is correct and accessible: ${VS_LOG_SHEET_SPREADSHEET_URL}`);
  }

  let sheet = spreadsheet.getSheetByName(VS_LOG_SHEET_NAME);

  // Fallback sheet finding logic (like in AIR_Gsheets)
  if (!sheet) {
    Logger.log(`Log sheet "${VS_LOG_SHEET_NAME}" not found by name. Attempting to use sheet by gid or first sheet.`);
    const gidMatch = VS_LOG_SHEET_SPREADSHEET_URL.match(/gid=(\d+)/);
    if (gidMatch && gidMatch[1]) {
      const gid = gidMatch[1];
      const sheets = spreadsheet.getSheets();
      sheet = sheets.find(s => s.getSheetId().toString() === gid);
      if (sheet) Logger.log(`Using log sheet by ID: "${sheet.getName()}"`);
    }
    if (!sheet) {
      sheet = spreadsheet.getSheets()[0];
      if (sheet) {
        Logger.log(`Warning: Using first available sheet in log spreadsheet: "${sheet.getName()}"`);
      } else {
        throw new Error(`Could not find any sheets in the log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL}`);
      }
    }
  } else {
     Logger.log(`Using specified log sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) { // Need at least header row and one data row
    Logger.log(`Not enough data in log sheet "${sheet.getName()}". Found ${data.length} rows. Expected headers + data.`);
    return null; // Not enough data
  }

  // Assume headers are in Row 1 (index 0)
  const headers = data[0].map(String);
  const rows = data.slice(1); // Data starts from row 2 (index 1)

  // Define required and optional columns for this script
  // Adjust based on which metrics are finally implemented
  const requiredColumns = [
      'Interview_email_sent_at',
      'Profile_id', // For uniqueness if needed
      'Position_id', // Needed for unique candidate-position invites
      // Status column - prioritize Interview Status_Real
  ];
  const optionalColumns = [
      'Candidate_name', // For debugging time calculation
      'Position_name', // Needed for filtering
      'Interview_status', // Fallback status column
      'Interview Status_Real', // Preferred status column
      'Schedule_start_time', 'Duration_minutes', 'Feedback_status', 'Feedback_json',
      'Match_stars', 'Location_country', 'Job_function', 'Position_id', 'Recruiter_name',
      'Creator_user_id', 'Reviewer_email', 'Hiring_manager_name',
      'Days_pending_invitation', 'Interview Status_Real'
  ];

  const colIndices = {};
  const missingCols = [];

  // --- Find Status Column --- Enforce Interview Status_Real ---
  const statusColName = 'Interview_status_real';
  const statusColIndex = headers.indexOf(statusColName);
  if (statusColIndex !== -1) {
      colIndices['STATUS_COLUMN'] = statusColIndex;
      Logger.log(`Using column "${statusColName}" (index ${statusColIndex}) for interview status analysis.`);
  } else {
      // Add the specific required column name to the missing list
      missingCols.push(statusColName);
  } // Error will be thrown later if missingCols is not empty
  // --- End Find Status Column ---

  requiredColumns.forEach(colName => {
    const index = headers.indexOf(colName);
    if (index === -1) {
      missingCols.push(colName);
    } else {
      colIndices[colName] = index;
    }
  });

  if (missingCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in log sheet "${sheet.getName()}": ${missingCols.join(', ')}`);
    throw new Error(`Required column(s) not found in log sheet headers (Row 1): ${missingCols.join(', ')}`);
  }

  optionalColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index !== -1) {
          colIndices[colName] = index;
      } else {
          Logger.log(`Optional column "${colName}" not found.`);
      }
  });

  Logger.log(`Found required columns. Indices: ${JSON.stringify(colIndices)}`);
  return { rows, headers, colIndices };
}


/**
 * Filters the data based on a time range (e.g., last N days based on Interview_email_sent_at).
 * @param {Array<Array>} rows The data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {Array<Array>} Filtered rows.
 */
function filterDataByTimeRange(rows, colIndices) {
  if (!colIndices.hasOwnProperty('Interview_email_sent_at')) {
      Logger.log("WARNING: Cannot filter by time range - 'Interview_email_sent_at' column index not found.");
      return rows; // Return all rows if column is missing
  }

  const sentAtIndex = colIndices['Interview_email_sent_at'];
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - VS_REPORT_TIME_RANGE_DAYS);
  const cutoffTimestamp = cutoffDate.getTime();

  Logger.log(`Filtering data for interviews sent on or after ${cutoffDate.toLocaleDateString()}`);

  const filteredRows = rows.filter(row => {
    if (row.length <= sentAtIndex) return false; // Skip short rows
    const rawDate = row[sentAtIndex];
    const sentDate = vsParseDateSafe(rawDate);
    return sentDate && sentDate.getTime() >= cutoffTimestamp;
  });

  return filteredRows;
}


/**
 * Calculates company-level metrics from the filtered data.
 * @param {Array<Array>} filteredRows The filtered data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {object} An object containing calculated metrics.
 */
function calculateCompanyMetrics(filteredRows, colIndices) {
  const metrics = {
    reportStartDate: (() => { const d = new Date(); d.setDate(d.getDate() - VS_REPORT_TIME_RANGE_DAYS); return vsFormatDate(d); })(),
    reportEndDate: vsFormatDate(new Date()),
    totalSent: filteredRows.length, // This now reflects rows after time and position filters
    totalScheduled: 0,
    totalCompleted: 0,
    totalFeedbackSubmitted: 0,
    // Funnel Rates
    sentToScheduledRate: 0,
    scheduledToCompletedRate: 0, // Based on those scheduled
    completedToFeedbackRate: 0, // Based on those completed
    // Timelines (sum and count for calculating average later)
    sentToScheduledDaysSum: 0,
    sentToScheduledCount: 0,
    completedToFeedbackDaysSum: 0, // Needs completion and feedback timestamps
    completedToFeedbackCount: 0,
    // Durations
    // durationMinutesSum: 0, // Removed
    // durationCount: 0, // Removed
    // Match Stars (for completed)
    matchStarsSum: 0,
    matchStarsCount: 0,
    // Breakdowns
    completionRateByJobFunction: {}, // { "JobFunc": { completed: 0, totalConsidered: 0 } }
    avgTimeToFeedbackByCountry: {}, // { "Country": { sumDays: 0, count: 0 } } // Placeholder, complex
    interviewStatusDistribution: {}, // { "Status": { count: X, percentage: Y } }
    // Raw data storage for breakdowns
    byJobFunction: {}, // { "JobFunc": { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} } }
    byCountry: {},     // { "Country": { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, statusCounts: {} } }
    // Timeseries data
    dailySentCounts: {} // { "YYYY-MM-DD": count }
  };

  // --- Status Definitions (Using RAW values from sheet) ---
  // Define statuses for interviews included in Sent-to-Scheduled calculation
  const STATUSES_FOR_AVG_TIME_CALC = ['SCHEDULED', 'COMPLETED']; // Check Interview Status_Real directly
  // Define statuses indicating completion (used for other metrics)
  const COMPLETED_STATUSES = ['COMPLETED', 'Feedback Provided', 'Pending Feedback', 'No Show']; // Raw values? Check case sensitivity
  // Define statuses considered "Pending"
  const PENDING_STATUSES = ['PENDING', 'INVITED', 'EMAIL SENT']; // Raw values? Check case sensitivity
  // Define Feedback_status values
  const FEEDBACK_SUBMITTED_STATUS = 'Submitted'; // Raw value from Feedback_status
  const RECRUITER_SUBMISSION_AWAITED_FEEDBACK = 'AI_RECOMMENDED'; // Raw value from Feedback_status

  // --- Column Indices (Check existence) ---
  const statusIdx = colIndices['STATUS_COLUMN'];
  const sentAtIdx = colIndices['Interview_email_sent_at'];
  const scheduledAtIdx = colIndices.hasOwnProperty('Schedule_start_time') ? colIndices['Schedule_start_time'] : -1;
  const candidateNameIdx = colIndices.hasOwnProperty('Candidate_name') ? colIndices['Candidate_name'] : -1;
  const feedbackStatusIdx = colIndices.hasOwnProperty('Feedback_status') ? colIndices['Feedback_status'] : -1;
  const durationIdx = colIndices.hasOwnProperty('Duration_minutes') ? colIndices['Duration_minutes'] : -1;
  const matchStarsIdx = colIndices.hasOwnProperty('Match_stars') ? colIndices['Match_stars'] : -1;
  const jobFuncIdx = colIndices.hasOwnProperty('Job_function') ? colIndices['Job_function'] : -1;
  const countryIdx = colIndices.hasOwnProperty('Location_country') ? colIndices['Location_country'] : -1;

  filteredRows.forEach(row => {
    // --- Get Sent Date for Timeseries ---
    const sentDate = vsParseDateSafe(row[sentAtIdx]);
    if (sentDate) {
        const dateString = vsFormatDate(sentDate); // Format as DD-MMM-YY
        metrics.dailySentCounts[dateString] = (metrics.dailySentCounts[dateString] || 0) + 1;
    }

    // --- Get Core Values ---
    const statusRaw = row[statusIdx] ? String(row[statusIdx]).trim() : 'Unknown';
    const jobFunc = (jobFuncIdx !== -1 && row[jobFuncIdx]) ? String(row[jobFuncIdx]).trim() : 'Unknown';
    const country = (countryIdx !== -1 && row[countryIdx]) ? String(row[countryIdx]).trim() : 'Unknown';
    const feedbackStatusRaw = (feedbackStatusIdx !== -1 && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim() : '';

    // --- Initialize Breakdown Structures if they don't exist ---
    if (!metrics.byJobFunction[jobFunc]) {
        metrics.byJobFunction[jobFunc] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }
    if (!metrics.byCountry[country]) {
        metrics.byCountry[country] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, statusCounts: {} };
    }

    // --- Increment Base Counts ---
    // Every row in filteredRows represents a 'sent' interview
    metrics.byJobFunction[jobFunc].sent++;
    metrics.byCountry[country].sent++;
    // Store raw status counts before calculating percentages later
    metrics.interviewStatusDistribution[statusRaw] = (metrics.interviewStatusDistribution[statusRaw] || 0) + 1;
    metrics.byJobFunction[jobFunc].statusCounts[statusRaw] = (metrics.byJobFunction[jobFunc].statusCounts[statusRaw] || 0) + 1;
    metrics.byCountry[country].statusCounts[statusRaw] = (metrics.byCountry[country].statusCounts[statusRaw] || 0) + 1;

    // --- Calculate Avg Time Sent to Completion (Scheduled) ---
    // Check if status is SCHEDULED or COMPLETED
    if (STATUSES_FOR_AVG_TIME_CALC.includes(statusRaw)) {
        const candidateName = (candidateNameIdx !== -1 && row[candidateNameIdx]) ? row[candidateNameIdx] : 'Unknown Candidate';
        // Ensure both dates are valid before calculating difference
        const scheduleDateForAvg = (scheduledAtIdx !== -1) ? vsParseDateSafe(row[scheduledAtIdx]) : null;
        if (sentDate && scheduleDateForAvg) {
            const daysDiff = vsCalculateDaysDifference(sentDate, scheduleDateForAvg);
            if (daysDiff !== null) { // vsCalculateDaysDifference handles negative check
                metrics.sentToScheduledDaysSum += daysDiff;
                metrics.sentToScheduledCount++;
                // <<< DETAILED LOGGING FOR DEBUGGING >>>
                Logger.log(`DEBUG_AVG_TIME: Candidate=${candidateName}, Status=${statusRaw}, Sent=${sentDate.toISOString()}, Scheduled=${scheduleDateForAvg.toISOString()}, DiffDays=${daysDiff.toFixed(8)}`);
            }
        }
    }

    // --- Check if Scheduled (for breakdown counts) ---
    let isScheduledForCount = (statusRaw === 'SCHEDULED');

    if (isScheduledForCount) {
         metrics.totalScheduled++; // This count might become less meaningful now?
         metrics.byJobFunction[jobFunc].scheduled++;
         metrics.byCountry[country].scheduled++;
    }

    // --- Check if Pending ---
    if (PENDING_STATUSES.includes(statusRaw)) { // Compare raw status
        metrics.byJobFunction[jobFunc].pending++;
        metrics.byCountry[country].pending++;
        // Note: We don't increment an overall pending count here unless needed elsewhere
    }

    // --- Check if Completed ---
    let isCompleted = COMPLETED_STATUSES.includes(statusRaw); // Compare raw status
    if (isCompleted) {
      metrics.totalCompleted++;
      metrics.byJobFunction[jobFunc].completed++;
      metrics.byCountry[country].completed++;

      // --- Calculate Match Stars ---
       if (matchStarsIdx !== -1 && row[matchStarsIdx] !== null && row[matchStarsIdx] !== '') {
           const stars = parseFloat(row[matchStarsIdx]);
           if (!isNaN(stars) && stars >= 0) {
               metrics.matchStarsSum += stars;
               metrics.matchStarsCount++;
           }
       }

       // --- Check for Feedback Submitted (only if completed) ---
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw === FEEDBACK_SUBMITTED_STATUS) { // Compare raw status
         metrics.totalFeedbackSubmitted++;
         metrics.byJobFunction[jobFunc].feedbackSubmitted++; // Renamed for clarity
         metrics.byCountry[country].feedbackSubmitted++; // Track submitted feedback for country
       }

       // --- Check for Recruiter Submission Awaited (AI_RECOMMENDED in Feedback_status)
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw === RECRUITER_SUBMISSION_AWAITED_FEEDBACK) { // Compare raw status
           metrics.byJobFunction[jobFunc].recruiterSubmissionAwaited++;
           // Note: No overall count added unless specifically needed
       }
    }
  });

  // --- Calculate Final Rates and Averages ---
  if (metrics.totalSent > 0) {
      metrics.sentToScheduledRate = parseFloat(((metrics.totalScheduled / metrics.totalSent) * 100).toFixed(1));
      metrics.completionRate = parseFloat(((metrics.totalCompleted / metrics.totalSent) * 100).toFixed(1)); // Overall Completion Rate
      // Calculate percentages for status distribution
      const statusCountsTemp = { ...metrics.interviewStatusDistribution }; // Copy raw counts
      metrics.interviewStatusDistribution = {}; // Reset to store objects
      for (const status in statusCountsTemp) {
          const count = statusCountsTemp[status];
          metrics.interviewStatusDistribution[status] = {
              count: count,
              percentage: parseFloat(((count / metrics.totalSent) * 100).toFixed(1))
          };
      }
  }
  if (metrics.totalScheduled > 0) {
      // Rate based on those who were at least scheduled
      metrics.scheduledToCompletedRate = parseFloat(((metrics.totalCompleted / metrics.totalScheduled) * 100).toFixed(1));
  }
   if (metrics.totalCompleted > 0) {
      metrics.completedToFeedbackRate = parseFloat(((metrics.totalFeedbackSubmitted / metrics.totalCompleted) * 100).toFixed(1));
      if(metrics.matchStarsCount > 0) {
          metrics.avgMatchStars = parseFloat((metrics.matchStarsSum / metrics.matchStarsCount).toFixed(1));
      }
   }
   if (metrics.sentToScheduledCount > 0) {
       metrics.avgTimeToScheduleDays = parseFloat((metrics.sentToScheduledDaysSum / metrics.sentToScheduledCount).toFixed(1));
   } else {
       metrics.avgTimeToScheduleDays = null; // Set to null if no valid data
   }
    if (metrics.completedToFeedbackCount > 0) {
        metrics.avgCompletedToFeedbackDays = parseFloat((metrics.completedToFeedbackDaysSum / metrics.completedToFeedbackCount).toFixed(1)); // Example
    }

  // --- Calculate Breakdown Metrics ---
  // Iterate through Job Functions
  for (const func in metrics.byJobFunction) {
    const data = metrics.byJobFunction[func];
    data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0;
    data.completedNumber = data.completed; // Store raw number
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0;
    data.pendingNumber = data.pending; // Store raw number
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0;
  }

  // Iterate through Countries
  for (const ctry in metrics.byCountry) {
    const data = metrics.byCountry[ctry];
    data.completedNumber = data.completed;
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0;
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    // Add other country-specific metrics here if needed
  }

  Logger.log(`Metrics calculation complete. Total Sent: ${metrics.totalSent}, Scheduled: ${metrics.totalScheduled}, Completed: ${metrics.totalCompleted}`);
  return metrics;
}

// --- Reporting Functions ---

/**
 * Creates the HTML email report for company-level metrics.
 * @param {object} metrics The calculated metrics object.
 * @returns {string} The HTML content for the email body.
 */
function createVolkscienceHtmlReport(metrics) {

  // --- Helper to generate timeseries table ---
  const generateTimeseriesTable = (dailyCounts) => {
      const sortedDates = Object.keys(dailyCounts).sort((a, b) => {
          // Sort DD-MMM-YY requires parsing back to dates
          try {
              const dateA = new Date(a.replace(/(\d{2})-(\w{3})-(\d{2})/, '$2 $1, 20$3'));
              const dateB = new Date(b.replace(/(\d{2})-(\w{3})-(\d{2})/, '$2 $1, 20$3'));
              return dateA - dateB;
          } catch (e) {
              return a.localeCompare(b); // Fallback to string sort if parsing fails
          }
      });
      if (sortedDates.length === 0) {
          return '<p class="note">No interview invitations sent in this period.</p>';
      }
      // Use data-table for styling, but don't force center if it should fill width
      let tableHtml = '<table class="data-table"><thead><tr><th>🗓️ Date (DD-MMM-YY)</th><th>✉️ Invitations Sent</th></tr></thead><tbody>';
      sortedDates.forEach(date => {
          tableHtml += `<tr><td>${date}</td><td>${dailyCounts[date]}</td></tr>`;
      });
      tableHtml += '</tbody></table>';
      return tableHtml;
  };

  let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <title>${VS_COMPANY_NAME} AI Interview Report</title>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 10px; margin: 0; }
      .container { max-width: 850px; margin: 20px auto; padding: 25px; border: 1px solid #ccc; border-radius: 8px; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
      h1, h2, h3 { color: #333; }
      h1 { text-align: center; border-bottom: 2px solid #eee; padding-bottom: 15px; margin-bottom: 25px; font-size: 26px; color: #1a237e; }
      h2 { margin-top: 30px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; font-size: 18px; color: #3f51b5; }
      .metric-block { background-color: #fff; padding: 15px; border: 1px solid #eee; border-radius: 4px; margin-bottom: 15px; }
      .rate { color: #1976d2; } /* Adjusted Blue for rates */
      .time { color: #ef6c00; } /* Updated Orange for time KPI */
      .count { color: #424242; } /* Darker gray for counts */
      .percent-value { color: #0056b3; font-weight: normal; } /* Dark blue for percentages */
      .note { font-size: 0.85em; color: #757575; margin-top: 15px; }
      table.data-table { border-collapse: collapse; width: 100%; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden; }
      table.centered-table { margin-left: auto; margin-right: auto; width: auto; max-width: 98%; }
      th, td { border: 1px solid #e0e0e0; padding: 6px 10px; /* Reduced padding */ text-align: left; font-size: 12px; vertical-align: middle; }
      th { background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase; font-size: 11px; }
      tr:nth-child(even) { background-color: #fafafa; } /* Alternating row color */
      .breakdown-section { margin-top: 25px; }
      /* KPI Box Styling - Using Nested Tables */
      .kpi-nested-table { width: 100%; height: 130px; border: 1px solid #cccccc; border-radius: 8px; border-collapse: collapse; table-layout: fixed; overflow: hidden; /* Clip content if needed */ }
      .kpi-nested-table td { border: none; /* Remove internal cell borders */ vertical-align: middle; text-align: center; }
      .kpi-header-cell { /* background-color: #f5f5f5; */ padding: 6px 10px; font-size: 12px; font-weight: bold; color: #424242; border-bottom: 1px solid #cccccc; /* Divider line */ height: 30px; /* Fixed header height */ }
      .kpi-value-cell { padding: 10px; font-size: 34px; font-weight: bold; height: 100%; /* Fill remaining height */ }
      .kpi-value-cell .unit { font-size: 16px; font-weight: normal; margin-left: 3px; }

      /* Specific KPI Backgrounds/Value Colors */
      .kpi-value-cell.invitations { background-color: #e8f5e9; color: #2e7d32; }
      .kpi-nested-table.invitations .kpi-header-cell { background-color: #e8f5e9; }

      .kpi-value-cell.completion { background-color: #e3f2fd; color: #1976d2; }
      .kpi-nested-table.completion .kpi-header-cell { background-color: #e3f2fd; }

      .kpi-value-cell.time { background-color: #fff3e0; color: #ef6c00; }
      .kpi-nested-table.time .kpi-header-cell { background-color: #fff3e0; }

      .kpi-value-cell.stars { background-color: #f3e5f5; color: #8e24aa; }
      .kpi-nested-table.stars .kpi-header-cell { background-color: #f3e5f5; }

      .top-kpi-table { width: 100%; border-collapse: separate; border-spacing: 15px 0; /* Adjusted spacing */ margin-bottom: 25px; table-layout: fixed; }
      .top-kpi-cell { width: 25%; /* Four columns */ vertical-align: top; padding: 0; }
      .section-container { background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; /* Removed bottom margin */ }
      /* Bold first column in breakdown tables */
      .breakdown-section table.data-table tr td:first-child {
          font-weight: bold;
          text-align: left; /* Keep first column left-aligned */
      }
      /* Center align other breakdown table cells */
      .breakdown-section table.data-table tr td:not(:first-child) {
          text-align: center;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>AI Recruiter Adoption: Executive Summary</h1>

      <!-- Top KPI Boxes - 1x4 Layout -->
      <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="top-kpi-table">
        <tr>
          <td class="top-kpi-cell">
            <table class="kpi-nested-table invitations">
              <tr><td class="kpi-header-cell">✉️ AI Invitations Sent</td></tr>
              <tr><td class="kpi-value-cell invitations">${metrics.totalSent}</td></tr>
            </table>
          </td>
          <td class="top-kpi-cell">
            <table class="kpi-nested-table completion">
              <tr><td class="kpi-header-cell">✅ Completion Rate</td></tr>
              <tr><td class="kpi-value-cell completion">${metrics.completionRate}<span class="unit">%</span></td></tr>
            </table>
          </td>
          <td class="top-kpi-cell">
            <table class="kpi-nested-table time">
              <tr><td class="kpi-header-cell">⏱️ Avg Time Sent to Completion*</td></tr>
              <tr><td class="kpi-value-cell time">${metrics.avgTimeToScheduleDays !== null ? metrics.avgTimeToScheduleDays : 'N/A'}<span class="unit">days</span></td></tr>
            </table>
          </td>
          <td class="top-kpi-cell">
            <table class="kpi-nested-table stars">
              <tr><td class="kpi-header-cell">⭐ Avg Match Stars (Completed)</td></tr>
              <tr><td class="kpi-value-cell stars">${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}</td></tr>
            </table>
          </td>
        </tr>
      </table>

      <!-- Interview Completion Status -->
      <div class="metric-block">
          <div class="section-title">📊 AI Screening Completion Status</div>
          <table class="data-table centered-table">
             <thead><tr><th>Status</th><th>Count</th><th>Percentage</th></tr></thead>
             <tbody>
             ${Object.entries(metrics.interviewStatusDistribution)
                 .sort(([, dataA], [, dataB]) => dataB.count - dataA.count) // Sort by count descending
                 .map(([status, data]) => `
                     <tr>
                         <td>${status}</td>
                         <td>${data.count}</td>
                         <td><span class="percent-value">${data.percentage}%</span></td>
                     </tr>
                 `).join('')}
             </tbody>
         </table>
         <p class="note">Percentage is based on the total number of invitations sent since 17th April 2025 (Launch of AIR).</p>
     </div>

     <!-- Daily Invitations Sent -->
     <div class="section-container">
        <div class="section-title">🗓️ Daily Invitations Sent</div>
        <!-- Apply centered-table class to the table generated by the helper -->
        ${generateTimeseriesTable(metrics.dailySentCounts).replace('<table class="data-table">', '<table class="data-table centered-table">')}
     </div>

     <div class="breakdown-section">
         <h2>💼 Breakdown by Job Function</h2>
         <table class="data-table centered-table">
             <thead>
                <tr>
                   <th>Job Function</th>
                   <th>Sent</th>
                   <th>Completed (# / %)</th>
                   <th>Scheduled</th>
                   <th>Pending (# / %)</th>
                   <th>Feedback Submitted</th>
                   <th>Recruiter Submission Awaited</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byJobFunction)
                     .sort(([funcA], [funcB]) => funcA.localeCompare(funcB)) // Sort alphabetically
                     .map(([func, data]) => `
                         <tr>
                             <td>${func}</td>
                             <td>${data.sent}</td>
                             <td>${data.completedNumber} (<span class="percent-value">${data.completedPercentOfSent}%</span>)</td>
                             <td>${data.scheduled}</td>
                             <td>${data.pendingNumber} (<span class="percent-value">${data.pendingPercentOfSent}%</span>)</td>
                             <td>${data.feedbackSubmitted}</td>
                             <td>${data.recruiterSubmissionAwaited}</td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>

     <div class="breakdown-section">
         <h2>🌍 Breakdown by Location Country</h2>
          <table class="data-table centered-table">
             <thead>
                <tr>
                   <th>Country</th>
                   <th>Sent</th>
                   <th>Completed (# / %)</th>
                   <th>Scheduled</th>
                   <th>Pending (# / %)</th>
                   <th>Feedback Submitted</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byCountry)
                     .sort(([ctryA], [ctryB]) => ctryA.localeCompare(ctryB)) // Sort alphabetically
                     .map(([ctry, data]) => `
                         <tr>
                             <td>${ctry}</td>
                             <td>${data.sent}</td>
                             <td>${data.completedNumber} (<span class="percent-value">${data.completedPercentOfSent}%</span>)</td>
                             <td>${data.scheduled}</td>
                             <td>${data.pendingNumber} (<span class="percent-value">${data.pendingPercentOfSent}%</span>)</td>
                             <td>${data.feedbackSubmitted}</td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>

     <p class="note" style="text-align: center; margin-top: 30px;">
       *Avg Time Sent to Completion calculation currently uses Schedule Start Date as completion proxy.<br>
       Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.
     </p>
   </div>
 </body>
 </html>
 `;
 return html;
}

/**
 * Sends an email with the report.
 * @param {string} recipient The primary email recipient.
 * @param {string} ccRecipient The CC email recipient (can be empty).
 * @param {string} subject The email subject.
 * @param {string} htmlBody The HTML content of the email.
 */
function sendVolkscienceEmail(recipient, ccRecipient, subject, htmlBody) {
  if (!recipient) {
     Logger.log("ERROR: Email recipient is empty. Cannot send email.");
     return;
  }
   if (!subject) {
     Logger.log("WARNING: Email subject is empty. Using default subject.");
     subject = `${VS_COMPANY_NAME} AI Interview Report`;
  }
   if (!htmlBody) {
     Logger.log("ERROR: Email body is empty. Cannot send email.");
     return;
  }

  const options = {
     to: recipient,
     subject: subject,
     htmlBody: htmlBody
  };

  // Add CC only if it's defined, not empty, and different from the recipient
  if (ccRecipient && ccRecipient.trim() !== '' && ccRecipient.trim().toLowerCase() !== recipient.trim().toLowerCase()) {
    options.cc = ccRecipient;
    Logger.log(`Sending email to ${recipient}, CC ${ccRecipient}`);
  } else {
     Logger.log(`Sending email to ${recipient} (No CC or CC is same as recipient)`);
  }

  try {
      MailApp.sendEmail(options);
      Logger.log("Email sent successfully.");
  } catch (e) {
     Logger.log(`ERROR sending email: ${e.toString()}`);
     // Optional: re-throw or handle error further
     // throw e;
     // Send error notification might be better here if main function doesn't catch
     sendVsErrorNotification(`CRITICAL: Failed to send report email to ${recipient}`, `Error: ${e.toString()}`);
  }
}

/**
 * Sends an error notification email.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendVsErrorNotification(errorMessage, stackTrace = '') {
   const recipient = VS_EMAIL_RECIPIENT; // Send errors to the main recipient
   if (!recipient) {
       Logger.log("CRITICAL ERROR: Cannot send error notification because VS_EMAIL_RECIPIENT is not set.");
       return;
   }
   try {
       const subject = `ERROR: ${VS_COMPANY_NAME} AI Report Failed - ${new Date().toLocaleString()}`;
       let body = `Error generating/sending ${VS_COMPANY_NAME} AI Interview report:\n\n${errorMessage}\n\n`;
       if (stackTrace) {
           body += `Stack Trace:\n${stackTrace}\n\n`;
       }
       body += `Log Sheet URL: ${VS_LOG_SHEET_SPREADSHEET_URL}`;
       MailApp.sendEmail(recipient, subject, body);
       Logger.log(`Error notification email sent to ${recipient}.`);
    } catch (emailError) {
       Logger.log(`CRITICAL: Failed to send error notification email to ${recipient}: ${emailError}`);
    }
}


// --- Utility / Setup Functions ---

/**
 * Creates menu items in the Google Sheet UI (when script is opened from a Sheet).
 * Note: Having multiple onOpen functions in one project can be problematic.
 * Consider combining menu logic or using manual triggers.
 */
function setupVolkscienceMenu() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(`${VS_COMPANY_NAME} AI Report`)
      .addItem('Generate & Send Report Now', 'AIR_ExecSummary_Daily')
      .addItem('Schedule Weekly Report', 'createVolkscienceTrigger')
      .addToUi();
  } catch (e) {
    // Log error but don't prevent sheet opening
    Logger.log("Error creating Volkscience menu (might happen if not opened from a Sheet): " + e);
  }
}

// --- Helper Functions ---
/**
 * Parses date strings safely, returning null for invalid dates/inputs.
 * @param {any} dateInput Input value (string, number, Date object).
 * @returns {Date|null} Parsed Date object or null if invalid/empty.
 */
function vsParseDateSafe(dateInput) {
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

/**
 * Calculates time difference in days between two dates.
 * @param {Date|null} date1 Earlier date object.
 * @param {Date|null} date2 Later date object.
 * @returns {number|null} Difference in days (float), or null if inputs invalid or difference is negative.
 */
function vsCalculateDaysDifference(date1, date2) {
    if (!date1 || !date2) return null;
    const diffTime = date2.getTime() - date1.getTime();
    // Allow zero difference, ignore negative
    if (diffTime < 0) return null;
    return diffTime / (1000 * 60 * 60 * 24);
}

/**
 * Formats a Date object into DD-MMM-YY string (e.g., 25-Jul-24).
 * @param {Date|null} dateObject The date to format.
 * @returns {string} Formatted date string or 'N/A' if input is invalid.
 */
function vsFormatDate(dateObject) {
    if (!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) {
        return 'N/A';
    }
    const day = String(dateObject.getDate()).padStart(2, '0');
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[dateObject.getMonth()];
    const year = String(dateObject.getFullYear()).slice(-2);
    return `${day}-${month}-${year}`;
}

/**
 * Assigns a numerical rank to interview statuses for prioritization during deduplication.
 * Lower rank means higher priority (e.g., Completed is better than Scheduled).
 * @param {string} status The raw interview status string.
 * @returns {number} The rank of the status.
 */
function vsGetStatusRank(status) {
    // Define statuses indicating completion (used for other metrics)
    const COMPLETED_STATUSES_RAW = ['COMPLETED', 'Feedback Provided', 'Pending Feedback', 'No Show']; // Keep raw values consistent with calculateCompanyMetrics
    // Define statuses considered "Scheduled"
    const SCHEDULED_STATUS_RAW = 'SCHEDULED'; // Keep raw value consistent
    // Define statuses considered "Pending"
    const PENDING_STATUSES_RAW = ['PENDING', 'INVITED', 'EMAIL SENT']; // Keep raw values consistent

    if (COMPLETED_STATUSES_RAW.includes(status)) {
        return 1; // Highest priority
    } else if (status === SCHEDULED_STATUS_RAW) {
        return 2;
    } else if (PENDING_STATUSES_RAW.includes(status)) {
        return 3;
    } else {
        return 99; // Lowest priority for anything else (Expired, Cancelled, Unknown etc.)
    }
} 