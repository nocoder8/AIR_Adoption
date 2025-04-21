// AIR Volkscience - Company-Level AI Interview Analytics Script v1.0
// This script analyzes data from the Log_Enhanced sheet to provide company-wide insights
// into the AI interview process funnel, timelines, and outcomes.

// --- Configuration ---
const EMAIL_RECIPIENT = 'pkumar@eightfold.ai'; // <<< UPDATE EMAIL RECIPIENT
const EMAIL_CC = ''; // Optional CC
// Assuming the Log Enhanced sheet is in a separate Spreadsheet
const LOG_SHEET_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE/edit'; // <<< VERIFY SPREADSHEET URL
const LOG_SHEET_NAME = 'Log_Enhanced'; // <<< VERIFY SHEET NAME
const REPORT_TIME_RANGE_DAYS = 90; // Default time range (days back) for the report
const COMPANY_NAME = "Eightfold"; // Used in report titles etc.


// --- Main Functions ---

/**
 * Creates a trigger to run the report weekly.
 */
function createVolkscienceTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateAndSendVolkscienceReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create a new trigger to run weekly (e.g., Monday at 8 AM)
  ScriptApp.newTrigger('generateAndSendVolkscienceReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log(`Weekly trigger created for generateAndSendVolkscienceReport (Monday 8 AM)`);
  SpreadsheetApp.getUi().alert(`Weekly trigger created for ${COMPANY_NAME} AI Interview Report (Monday 8 AM).`);
}

/**
 * Main function to generate and send the company-level AI interview report.
 */
function generateAndSendVolkscienceReport() {
  try {
    Logger.log(`--- Starting ${COMPANY_NAME} AI Interview Report Generation ---`);

    // 1. Get Log Sheet Data
    const logData = getLogSheetData();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log('No data found in the log sheet or required columns missing. Skipping report generation.');
      // Optional: Send an email notification about missing data/columns
      // sendErrorNotification("Report Skipped: No data or required columns found in Log_Enhanced sheet.");
      return;
    }
     Logger.log(`Successfully retrieved ${logData.rows.length} rows from log sheet.`);

    // 2. Filter Data by Time Range (using Interview_email_sent_at)
    const filteredData = filterDataByTimeRange(logData.rows, logData.colIndices);
    if (filteredData.length === 0) {
        Logger.log(`No data found within the last ${REPORT_TIME_RANGE_DAYS} days. Skipping report.`);
        return;
    }
    Logger.log(`Filtered data to ${filteredData.length} rows based on the last ${REPORT_TIME_RANGE_DAYS} days.`);

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

    // 3. Calculate Metrics
    const metrics = calculateCompanyMetrics(finalFilteredData, logData.colIndices);
    Logger.log('Successfully calculated company metrics.');
    // Logger.log(`Calculated Metrics: ${JSON.stringify(metrics)}`); // Optional: Log detailed metrics

    // 4. Create HTML Report
    const htmlContent = createVolkscienceHtmlReport(metrics);
    Logger.log('Successfully generated HTML report content.');

    // 5. Send Email
    const reportTitle = `${COMPANY_NAME} AI Interview Report - Last ${REPORT_TIME_RANGE_DAYS} Days (${new Date().toLocaleDateString()})`;
    sendVolkscienceEmail(EMAIL_RECIPIENT, EMAIL_CC, reportTitle, htmlContent);

    Logger.log(`--- ${COMPANY_NAME} AI Interview Report generated and sent successfully! ---`);
    return `Report sent to ${EMAIL_RECIPIENT}`;

  } catch (error) {
    Logger.log(`Error in generateAndSendVolkscienceReport: ${error.toString()} Stack: ${error.stack}`);
    // Send error email
    sendErrorNotification(`ERROR generating ${COMPANY_NAME} AI Report: ${error.toString()}`, error.stack);
    return `Error: ${error.toString()}`;
  }
}

// --- Data Retrieval and Processing Functions ---

/**
 * Reads and processes data from the Log_Enhanced sheet.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null if error/no sheet/missing columns.
 */
function getLogSheetData() {
  Logger.log(`Attempting to open log spreadsheet: ${LOG_SHEET_SPREADSHEET_URL}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(LOG_SHEET_SPREADSHEET_URL);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening log spreadsheet by URL: ${e}`);
    throw new Error(`Could not open the specified Log Spreadsheet URL. Please verify the URL is correct and accessible: ${LOG_SHEET_SPREADSHEET_URL}`);
  }

  let sheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);

  // Fallback sheet finding logic (like in AIR_Gsheets)
  if (!sheet) {
    Logger.log(`Log sheet "${LOG_SHEET_NAME}" not found by name. Attempting to use sheet by gid or first sheet.`);
    const gidMatch = LOG_SHEET_SPREADSHEET_URL.match(/gid=(\d+)/);
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
        throw new Error(`Could not find any sheets in the log spreadsheet: ${LOG_SHEET_SPREADSHEET_URL}`);
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
      'Interview_status', // Crucial for funnel
      'Profile_id', // For uniqueness if needed
      // Add others as needed for core metrics...
  ];
  const optionalColumns = [
      'Position_name', // Needed for filtering
      'Schedule_start_time', 'Duration_minutes', 'Feedback_status', 'Feedback_json',
      'Match_stars', 'Location_country', 'Job_function', 'Position_id',
      'Creator_user_id', 'Reviewer_email', 'Hiring_manager_name', 'Recruiter_name',
      'Days_pending_invitation', 'Interview Status_Real'
  ];

  const colIndices = {};
  const missingCols = [];

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
  cutoffDate.setDate(cutoffDate.getDate() - REPORT_TIME_RANGE_DAYS);
  const cutoffTimestamp = cutoffDate.getTime();

  Logger.log(`Filtering data for interviews sent on or after ${cutoffDate.toLocaleDateString()}`);

  const filteredRows = rows.filter(row => {
    if (row.length <= sentAtIndex) return false; // Skip short rows
    const rawDate = row[sentAtIndex];
    const sentDate = parseDateSafe(rawDate);
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
    reportStartDate: (() => { const d = new Date(); d.setDate(d.getDate() - REPORT_TIME_RANGE_DAYS); return d.toLocaleDateString(); })(),
    reportEndDate: new Date().toLocaleDateString(),
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
    durationMinutesSum: 0,
    durationCount: 0,
    // Match Stars (for completed)
    matchStarsSum: 0,
    matchStarsCount: 0,
    // Breakdowns
    completionRateByJobFunction: {}, // { "JobFunc": { completed: 0, totalConsidered: 0 } }
    avgTimeToFeedbackByCountry: {}, // { "Country": { sumDays: 0, count: 0 } }
    interviewStatusDistribution: {}, // { "Status": count }
    // Raw data storage for breakdowns
    byJobFunction: {}, // { "JobFunc": { sent: 0, scheduled: 0, completed: 0, feedback: 0, durationSum: 0, durationCount: 0, statusCounts: {} } }
    byCountry: {},     // { "Country": { sent: 0, scheduled: 0, completed: 0, feedback: 0, statusCounts: {} } }
    // Timeseries data
    dailySentCounts: {} // { "YYYY-MM-DD": count }
  };

  // --- Status Definitions (Customize based on your data) ---
  // Define statuses that indicate an interview was definitively scheduled
  const SCHEDULED_STATUSES = ['scheduled', 'confirmed', 'rescheduled', 'completed', 'feedback provided', 'pending feedback', 'no show']; // Lowercase
  // Define statuses that indicate an interview was completed
  const COMPLETED_STATUSES = ['completed', 'feedback provided', 'pending feedback', 'no show']; // Lowercase, include no-shows as technically completed appointment slot?
  // Define the status indicating feedback was submitted (from Feedback_status column)
  const FEEDBACK_SUBMITTED_STATUS = 'submitted'; // Lowercase

  // --- Column Indices (Check existence) ---
  const statusIdx = colIndices['Interview_status'];
  const sentAtIdx = colIndices['Interview_email_sent_at'];
  const scheduledAtIdx = colIndices.hasOwnProperty('Schedule_start_time') ? colIndices['Schedule_start_time'] : -1;
  const feedbackStatusIdx = colIndices.hasOwnProperty('Feedback_status') ? colIndices['Feedback_status'] : -1;
  const durationIdx = colIndices.hasOwnProperty('Duration_minutes') ? colIndices['Duration_minutes'] : -1;
  const matchStarsIdx = colIndices.hasOwnProperty('Match_stars') ? colIndices['Match_stars'] : -1;
  const jobFuncIdx = colIndices.hasOwnProperty('Job_function') ? colIndices['Job_function'] : -1;
  const countryIdx = colIndices.hasOwnProperty('Location_country') ? colIndices['Location_country'] : -1;

  // TODO: Define what statuses count as "Scheduled" and "Completed"
  // Example: const scheduledStatuses = ['Scheduled', 'Invited', 'Confirmed']; // Adjust as needed
  const completedStatuses = COMPLETED_STATUSES; // Using defined constant
  const feedbackSubmittedStatus = FEEDBACK_SUBMITTED_STATUS; // Using defined constant

  filteredRows.forEach(row => {
    // --- Get Sent Date for Timeseries ---
    const sentDate = parseDateSafe(row[sentAtIdx]);
    if (sentDate) {
        const dateString = sentDate.toISOString().split('T')[0]; // Format as YYYY-MM-DD
        metrics.dailySentCounts[dateString] = (metrics.dailySentCounts[dateString] || 0) + 1;
    }

    // --- Get Core Values ---
    const statusRaw = row[statusIdx] ? String(row[statusIdx]).trim() : 'Unknown';
    const statusLower = statusRaw.toLowerCase();
    const jobFunc = (jobFuncIdx !== -1 && row[jobFuncIdx]) ? String(row[jobFuncIdx]).trim() : 'Unknown';
    const country = (countryIdx !== -1 && row[countryIdx]) ? String(row[countryIdx]).trim() : 'Unknown';
    const feedbackStatusRaw = (feedbackStatusIdx !== -1 && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim() : '';
    const feedbackStatusLower = feedbackStatusRaw.toLowerCase();

    // --- Initialize Breakdown Structures if they don't exist ---
    if (!metrics.byJobFunction[jobFunc]) {
        metrics.byJobFunction[jobFunc] = { sent: 0, scheduled: 0, completed: 0, feedback: 0, durationSum: 0, durationCount: 0, statusCounts: {} };
    }
    if (!metrics.byCountry[country]) {
        metrics.byCountry[country] = { sent: 0, scheduled: 0, completed: 0, feedback: 0, statusCounts: {} };
    }

    // --- Increment Base Counts ---
    // Every row in filteredRows represents a 'sent' interview
    metrics.byJobFunction[jobFunc].sent++;
    metrics.byCountry[country].sent++;
    metrics.interviewStatusDistribution[statusRaw] = (metrics.interviewStatusDistribution[statusRaw] || 0) + 1;
    metrics.byJobFunction[jobFunc].statusCounts[statusRaw] = (metrics.byJobFunction[jobFunc].statusCounts[statusRaw] || 0) + 1;
    metrics.byCountry[country].statusCounts[statusRaw] = (metrics.byCountry[country].statusCounts[statusRaw] || 0) + 1;

    // --- Check if Scheduled ---
    // Logic: Either the status indicates scheduled OR Schedule_start_time has a valid date.
    let isScheduled = SCHEDULED_STATUSES.includes(statusLower);
    let scheduleDate = null;
    if (scheduledAtIdx !== -1) {
        scheduleDate = parseDateSafe(row[scheduledAtIdx]);
        if (scheduleDate) {
            isScheduled = true; // If date exists, consider it scheduled regardless of status list (safer)
        }
    }

    if (isScheduled) {
       metrics.totalScheduled++;
       metrics.byJobFunction[jobFunc].scheduled++;
       metrics.byCountry[country].scheduled++;

       // --- Calculate Sent to Scheduled Time ---
       const daysDiff = calculateDaysDifference(sentDate, scheduleDate); // Use parsed scheduleDate and sentDate
       if (daysDiff !== null) {
         metrics.sentToScheduledDaysSum += daysDiff;
         metrics.sentToScheduledCount++;
         // TODO: Add breakdown timeline if needed
       }
    }

    // --- Check if Completed ---
    let isCompleted = COMPLETED_STATUSES.includes(statusLower);
    if (isCompleted) {
      metrics.totalCompleted++;
      metrics.byJobFunction[jobFunc].completed++;
      metrics.byCountry[country].completed++;

      // --- Calculate Duration ---
      if (durationIdx !== -1 && row[durationIdx] !== null && row[durationIdx] !== '') {
          const duration = parseFloat(row[durationIdx]);
          if (!isNaN(duration) && duration >= 0) {
              metrics.durationMinutesSum += duration;
              metrics.durationCount++;
              metrics.byJobFunction[jobFunc].durationSum += duration;
              metrics.byJobFunction[jobFunc].durationCount++;
          }
      }

      // --- Calculate Match Stars ---
       if (matchStarsIdx !== -1 && row[matchStarsIdx] !== null && row[matchStarsIdx] !== '') {
           const stars = parseFloat(row[matchStarsIdx]);
           if (!isNaN(stars) && stars >= 0) {
               metrics.matchStarsSum += stars;
               metrics.matchStarsCount++;
           }
       }

       // --- Check for Feedback Submitted (only if completed) ---
       if (feedbackStatusIdx !== -1 && row[feedbackStatusIdx] === feedbackSubmittedStatus) {
         metrics.totalFeedbackSubmitted++;
         metrics.byJobFunction[jobFunc].feedback++;
         metrics.byCountry[country].feedback++;

         // --- Calculate Completed to Feedback Time ---
         // TODO: Needs reliable 'completed' and 'feedback submitted' timestamps.
         // This is complex and might require parsing Feedback_json or another column.
         // Placeholder:
         // const completedDate = ... ; // Infer completion date
         // const feedbackDate = ... ; // Infer feedback submission date
         // const feedbackDaysDiff = calculateDaysDifference(completedDate, feedbackDate);
         // if (feedbackDaysDiff !== null) {
         //    metrics.completedToFeedbackDaysSum += feedbackDaysDiff;
         //    metrics.completedToFeedbackCount++;
         //    // Add to breakdown by country?
         // }
       }
    }
  });

  // --- Calculate Final Rates and Averages ---
  if (metrics.totalSent > 0) {
      metrics.sentToScheduledRate = parseFloat(((metrics.totalScheduled / metrics.totalSent) * 100).toFixed(1));
      // Calculate percentages for status distribution
      for (const status in metrics.interviewStatusDistribution) {
          const count = metrics.interviewStatusDistribution[status];
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
       metrics.avgSentToScheduledDays = parseFloat((metrics.sentToScheduledDaysSum / metrics.sentToScheduledCount).toFixed(1));
   }
    if (metrics.durationCount > 0) {
        metrics.avgInterviewDuration = parseFloat((metrics.durationMinutesSum / metrics.durationCount).toFixed(1));
    }
   // if (metrics.completedToFeedbackCount > 0) {
   //     metrics.avgCompletedToFeedbackDays = parseFloat((metrics.completedToFeedbackDaysSum / metrics.completedToFeedbackCount).toFixed(1));
   // }


  // --- Calculate Breakdown Metrics ---
  // Iterate through Job Functions
  for (const func in metrics.byJobFunction) {
    const data = metrics.byJobFunction[func];
    data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0;
    data.completedRate = data.scheduled > 0 ? parseFloat(((data.completed / data.scheduled) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedback / data.completed) * 100).toFixed(1)) : 0;
    data.avgDuration = data.durationCount > 0 ? parseFloat((data.durationSum / data.durationCount).toFixed(1)) : null;
  }

  // Iterate through Countries
  for (const ctry in metrics.byCountry) {
    const data = metrics.byCountry[ctry];
    data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0;
    data.completedRate = data.scheduled > 0 ? parseFloat(((data.completed / data.scheduled) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedback / data.completed) * 100).toFixed(1)) : 0;
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
  // TODO: Enhance styling, add charts (can be complex in email), improve layout.
  // Use CSS similar to AIR_Gsheets report for consistency if desired.

  // --- Helper to generate timeseries table ---
  const generateTimeseriesTable = (dailyCounts) => {
      const sortedDates = Object.keys(dailyCounts).sort();
      if (sortedDates.length === 0) {
          return '<p class="note">No interview invitations sent in this period.</p>';
      }
      let tableHtml = '<table style="margin-top: 10px;"><thead><tr><th>Date (YYYY-MM-DD)</th><th>Invitations Sent</th></tr></thead><tbody>';
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
    <title>${COMPANY_NAME} AI Interview Report</title>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 10px; }
      .container { max-width: 800px; margin: 20px auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; background-color: #ffffff; }
      h1, h2, h3 { color: #333; }
      h1 { text-align: center; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 20px; font-size: 24px;}
      h2 { margin-top: 30px; border-bottom: 1px solid #eee; padding-bottom: 5px;}
      .metric-block { background-color: #fff; padding: 15px; border: 1px solid #eee; border-radius: 4px; margin-bottom: 15px; }
      .metric { margin-bottom: 8px; font-size: 1.1em; }
      .metric-label { font-weight: bold; color: #555; display: inline-block; width: 250px; }
      .metric-value { font-weight: bold; color: #0056b3; }
      .rate { color: #007bff; } /* Blue for rates */
      .time { color: #dc3545; } /* Red for time */
      .count { color: #28a745; } /* Green for counts */
      .note { font-size: 0.9em; color: #666; margin-top: 15px; }
      table { border-collapse: collapse; width: 100%; margin-top: 15px; }
      th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 13px; }
      th { background-color: #f2f2f2; font-weight: bold; }
      .breakdown-section { margin-top: 20px; }
      .kpi-box { background-color: #e6f4ea; border: 1px solid #c8e6c9; border-radius: 4px; padding: 20px; text-align: center; height: 150px; display: flex; flex-direction: column; justify-content: center; }
      .kpi-label { font-size: 16px; color: #333; margin-bottom: 5px; }
      .kpi-value { font-size: 48px; font-weight: bold; color: #2e7d32; }
      .dashboard-table { width: 100%; border-collapse: collapse; border-spacing: 15px; margin-bottom: 20px; } /* Increased spacing */
      .dashboard-cell { vertical-align: top; padding: 0 7.5px; } /* Add horizontal padding */
      .dashboard-cell-left { width: 30%; }
      .dashboard-cell-right { width: 70%; }
      .section-container { background-color: #fff; padding: 15px; border: 1px solid #eee; border-radius: 4px; margin-bottom: 15px; }
      .section-title { font-size: 18px; color: #333; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>${COMPANY_NAME} AI Interview Report</h1>
      <p style="text-align: center; margin-top: -15px; margin-bottom: 25px; color: #555;">
        Data from ${metrics.reportStartDate} to ${metrics.reportEndDate} (Based on Interview Sent Date)
      </p>

      <!-- Dashboard Layout Table -->
      <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="dashboard-table">
        <tr>
          <td class="dashboard-cell dashboard-cell-left">
            <!-- KPI Box -->
            <div class="kpi-box">
              <div class="kpi-label">AI Invitations</div>
              <div class="kpi-value">${metrics.totalSent}</div>
            </div>
          </td>
          <td class="dashboard-cell dashboard-cell-right">
            <!-- Daily Invitations Table -->
            <div class="section-container">
               <div class="section-title">Daily Invitations Sent</div>
               ${generateTimeseriesTable(metrics.dailySentCounts)}
            </div>
          </td>
        </tr>
      </table>

      <h2>Overall Funnel Performance</h2>
      <div class="metric-block">
        <div class="metric"><span class="metric-label">Total Scheduled:</span> <span class="metric-value count">${metrics.totalScheduled}</span> (<span class="metric-value rate">${metrics.sentToScheduledRate}%</span> of Sent)</div>
        <div class="metric"><span class="metric-label">Total Completed:</span> <span class="metric-value count">${metrics.totalCompleted}</span> (<span class="metric-value rate">${metrics.scheduledToCompletedRate}%</span> of Scheduled)</div>
        <div class="metric"><span class="metric-label">Feedback Submitted:</span> <span class="metric-value count">${metrics.totalFeedbackSubmitted}</span> (<span class="metric-value rate">${metrics.completedToFeedbackRate}%</span> of Completed)</div>
      </div>

      <h2>Key Timelines & Metrics</h2>
      <div class="metric-block">
        <div class="metric"><span class="metric-label">Avg. Time Sent to Scheduled:</span> <span class="metric-value time">${metrics.avgSentToScheduledDays !== null ? metrics.avgSentToScheduledDays + ' days' : 'N/A'}</span></div>
        <div class="metric"><span class="metric-label">Avg. Interview Duration:</span> <span class="metric-value time">${metrics.avgInterviewDuration !== null ? metrics.avgInterviewDuration + ' mins' : 'N/A'}</span></div>
        <div class="metric"><span class="metric-label">Avg. Match Stars (Completed):</span> <span class="metric-value">${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}</span></div>
        <!-- <div class="metric"><span class="metric-label">Avg. Time Completed to Feedback:</span> <span class="metric-value time">${metrics.avgCompletedToFeedbackDays !== null ? metrics.avgCompletedToFeedbackDays + ' days' : 'N/A'}</span></div> -->
        <!-- Uncomment above line when completed-to-feedback time calculation is implemented -->
      </div>

      <h2>Interview Status Distribution</h2>
      <div class="metric-block">
          <div class="section-title">Interview Completion Status</div>
          <table>
             <thead><tr><th>Status</th><th>Count</th><th>Percentage</th></tr></thead>
             <tbody>
             ${Object.entries(metrics.interviewStatusDistribution)
                 .sort(([, dataA], [, dataB]) => dataB.count - dataA.count) // Sort by count descending
                 .map(([status, data]) => `
                     <tr>
                         <td>${status}</td>
                         <td>${data.count}</td>
                         <td>${data.percentage}%</td>
                     </tr>
                 `).join('')}
             </tbody>
          </table>
          <p class="note">Percentage is based on the total number of interviews sent in the period.</p>
      </div>

      <!-- TODO: Add Breakdown Sections (e.g., By Job Function, By Country) -->
      <!-- Example structure:
      <div class="breakdown-section">
          <h2>Breakdown by Job Function</h2>
          <table>
              <thead><tr><th>Job Function</th><th>Sent</th><th>Scheduled (%)</th><th>Completed (%)</th><th>Feedback (%)</th><th>Avg Duration</th></tr></thead>
              <tbody>
                  // Iterate through metrics.byJobFunction and calculate/display values
              </tbody>
          </table>
      </div>
      -->

      <div class="breakdown-section">
          <h2>Breakdown by Job Function</h2>
          <table>
              <thead>
                 <tr>
                    <th>Job Function</th>
                    <th>Sent</th>
                    <th>Scheduled (%)</th>
                    <th>Completed (% of Sched.)</th>
                    <th>Feedback (% of Comp.)</th>
                    <th>Avg. Duration (mins)</th>
                  </tr>
              </thead>
              <tbody>
                  ${Object.entries(metrics.byJobFunction)
                      .sort(([funcA], [funcB]) => funcA.localeCompare(funcB)) // Sort alphabetically
                      .map(([func, data]) => `
                          <tr>
                              <td>${func}</td>
                              <td>${data.sent}</td>
                              <td>${data.scheduledRate}%</td>
                              <td>${data.completedRate}%</td>
                              <td>${data.feedbackRate}%</td>
                              <td>${data.avgDuration !== null ? data.avgDuration : 'N/A'}</td>
                          </tr>
                      `).join('')}
              </tbody>
          </table>
      </div>

      <div class="breakdown-section">
          <h2>Breakdown by Location Country</h2>
           <table>
              <thead>
                 <tr>
                    <th>Country</th>
                    <th>Sent</th>
                    <th>Scheduled (%)</th>
                    <th>Completed (% of Sched.)</th>
                    <th>Feedback (% of Comp.)</th>
                  </tr>
              </thead>
              <tbody>
                  ${Object.entries(metrics.byCountry)
                      .sort(([ctryA], [ctryB]) => ctryA.localeCompare(ctryB)) // Sort alphabetically
                      .map(([ctry, data]) => `
                          <tr>
                              <td>${ctry}</td>
                              <td>${data.sent}</td>
                              <td>${data.scheduledRate}%</td>
                              <td>${data.completedRate}%</td>
                              <td>${data.feedbackRate}%</td>
                          </tr>
                      `).join('')}
              </tbody>
          </table>
      </div>

      <p class="note" style="text-align: center; margin-top: 30px;">Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.</p>
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
     subject = `${COMPANY_NAME} AI Interview Report`;
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
     sendErrorNotification(`CRITICAL: Failed to send report email to ${recipient}`, `Error: ${e.toString()}`);
  }
}

/**
 * Sends an error notification email.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendErrorNotification(errorMessage, stackTrace = '') {
   const recipient = EMAIL_RECIPIENT; // Send errors to the main recipient
   if (!recipient) {
       Logger.log("CRITICAL ERROR: Cannot send error notification because EMAIL_RECIPIENT is not set.");
       return;
   }
   try {
       const subject = `ERROR: ${COMPANY_NAME} AI Report Failed - ${new Date().toLocaleString()}`;
       let body = `Error generating/sending ${COMPANY_NAME} AI Interview report:\n\n${errorMessage}\n\n`;
       if (stackTrace) {
           body += `Stack Trace:\n${stackTrace}\n\n`;
       }
       body += `Log Sheet URL: ${LOG_SHEET_SPREADSHEET_URL}`;
       MailApp.sendEmail(recipient, subject, body);
       Logger.log(`Error notification email sent to ${recipient}.`);
    } catch (emailError) {
       Logger.log(`CRITICAL: Failed to send error notification email to ${recipient}: ${emailError}`);
    }
}


// --- Utility / Setup Functions ---

/**
 * Creates menu items in the Google Sheet UI (when script is opened from a Sheet).
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(`${COMPANY_NAME} AI Report`)
      .addItem('Generate & Send Report Now', 'generateAndSendVolkscienceReport')
      .addItem('Schedule Weekly Report', 'createVolkscienceTrigger')
      .addToUi();
  } catch (e) {
    // Log error but don't prevent sheet opening
    Logger.log("Error creating menu (might happen if not opened from a Sheet): " + e);
  }
}

// --- Helper Functions ---
/**
 * Parses date strings safely, returning null for invalid dates/inputs.
 * @param {any} dateInput Input value (string, number, Date object).
 * @returns {Date|null} Parsed Date object or null if invalid/empty.
 */
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

/**
 * Calculates time difference in days between two dates.
 * @param {Date|null} date1 Earlier date object.
 * @param {Date|null} date2 Later date object.
 * @returns {number|null} Difference in days (float), or null if inputs invalid or difference is negative.
 */
function calculateDaysDifference(date1, date2) {
    if (!date1 || !date2) return null;
    const diffTime = date2.getTime() - date1.getTime();
    // Allow zero difference, ignore negative
    if (diffTime < 0) return null;
    return diffTime / (1000 * 60 * 60 * 24);
} 