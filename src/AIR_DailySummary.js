// AIR Volkscience - Exec Summary - Company-Level AI Interview Analytics Script v1.0 (Recruiter Breakdown)
// To: Akhila and Pavan
// When: Weekly, Monday at 10 AM (Can be adjusted)
// This script analyzes data from the Log_Enhanced sheet to provide company-wide insights
// including a breakdown by recruiter.

// --- Configuration ---
const VS_EMAIL_RECIPIENT_RB = 'pkumar@eightfold.ai'; // <<< UPDATE EMAIL RECIPIENT
const VS_EMAIL_CC_RB = 'pkumar@eightfold.ai'; // Optional CC
// Assuming the Log Enhanced sheet is in a separate Spreadsheet
const VS_LOG_SHEET_SPREADSHEET_URL_RB = 'https://docs.google.com/spreadsheets/d/1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE/edit'; // <<< VERIFY SPREADSHEET URL
const VS_LOG_SHEET_NAME_RB = 'Log_Enhanced'; // <<< VERIFY SHEET NAME
const VS_REPORT_TIME_RANGE_DAYS_RB = 99999; // Set large number to effectively include all time
const VS_COMPANY_NAME_RB = "Eightfold"; // Used in report titles etc.

// --- Configuration for Application Sheet (for Adoption Chart) ---
const APP_SHEET_SPREADSHEET_URL_RB = 'https://docs.google.com/spreadsheets/d/1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw/edit'; // <<< Weekly Report's App Sheet URL
const APP_SHEET_NAME_RB = 'Active+Rejected'; // <<< Weekly Report's App Sheet Name
const APP_LAUNCH_DATE_RB = new Date('2025-04-17'); // <<< Weekly Report's Launch Date (Needs to match)
const APP_MATCH_SCORE_THRESHOLD_RB = 4; // <<< Weekly Report's Score Threshold


// --- Main Functions ---

/**
 * Creates a trigger to run the recruiter breakdown report daily.
 */
function createRecruiterBreakdownTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'AIR_DailySummarytoAP') { // Updated Handler Name
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create a new trigger to run daily at 10 AM
  ScriptApp.newTrigger('AIR_DailySummarytoAP') // Updated Handler Name
    .timeBased()
    .everyDays(1) // Run daily
    .atHour(10) // Keep 10 AM or adjust as needed
    .create();
  Logger.log(`Daily trigger created for AIR_DailySummarytoAP (at 10 AM)`);
  // SpreadsheetApp.getUi().alert(`Daily trigger created for ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report (at 10 AM).`); // Removed: Cannot call getUi from trigger context
}

/**
 * Main function to generate and send the company-level AI interview report with recruiter breakdown.
 * Renamed from AIR_RecruiterBreakdown_Daily.
 */
function AIR_DailySummarytoAP() {
  try {
    Logger.log(`--- Starting ${VS_COMPANY_NAME_RB} AI Interview Daily Summary Report Generation ---`); // Updated log

    // 1. Get Log Sheet Data (Uses RB config)
    const logData = getLogSheetDataRB();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log('No data found in the log sheet or required columns missing. Skipping report generation.');
      // Optional: Send an email notification about missing data/columns
      // sendVsErrorNotificationRB("Report Skipped: No data or required columns found in Log_Enhanced sheet.");
      return;
    }
     Logger.log(`Successfully retrieved ${logData.rows.length} rows from log sheet.`);

    // 1b. Get Application Sheet Data (for Adoption Chart)
    let adoptionChartData = null;
    try {
        const appData = getApplicationDataForChartRB();
        if (appData && appData.rows) {
            Logger.log(`Successfully retrieved ${appData.rows.length} rows from application sheet.`);
            adoptionChartData = calculateAdoptionMetricsForChartRB(appData.rows, appData.colIndices);
            Logger.log(`Successfully calculated adoption chart metrics.`);
            // Logger.log(`Adoption Chart Data: ${JSON.stringify(adoptionChartData)}`); // Optional detail log
        } else {
            Logger.log(`WARNING: No data retrieved from application sheet. Adoption chart will be skipped.`);
        }
    } catch (appError) {
        Logger.log(`ERROR retrieving or processing application data for adoption chart: ${appError.toString()}`);
        // Continue without adoption chart data
        // Optional: Send notification about this specific failure?
        sendVsErrorNotificationRB(`Error getting data for Adoption Chart from ${APP_SHEET_NAME_RB}`, appError.stack);
    }

    // 2. Filter Data by Time Range (using Interview_email_sent_at)
    const filteredData = filterDataByTimeRangeRB(logData.rows, logData.colIndices);
    if (filteredData.length === 0) {
        Logger.log(`No data found within the last ${VS_REPORT_TIME_RANGE_DAYS_RB} days. Skipping report.`);
        return;
    }
    Logger.log(`Filtered data to ${filteredData.length} rows based on the last ${VS_REPORT_TIME_RANGE_DAYS_RB} days.`);

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

    // <<< Calculate Recruiter Last Sent Activity & Daily Trends >>>
    const recruiterLastSentMap = new Map();
    const recruiterDailyCounts = new Map(); // Map<RecruiterName, Map<DateString, Count>>
    const recruiterNameIdx_Log = logData.colIndices.hasOwnProperty('Recruiter_name') ? logData.colIndices['Recruiter_name'] : -1;
    const emailSentIdx_Log = logData.colIndices['Interview_email_sent_at']; // Already required

    if (recruiterNameIdx_Log !== -1) {
        finalFilteredData.forEach(row => {
            if (row.length > Math.max(recruiterNameIdx_Log, emailSentIdx_Log)) {
                const recruiterName = row[recruiterNameIdx_Log]?.trim();
                const rawSentDate = row[emailSentIdx_Log];
                if (recruiterName && recruiterName !== 'Unknown' && rawSentDate) {
                    const sentDate = vsParseDateSafeRB(rawSentDate);
                    if (sentDate) {
                        // Update Last Sent Date
                        if (!recruiterLastSentMap.has(recruiterName) || sentDate > recruiterLastSentMap.get(recruiterName)) {
                            recruiterLastSentMap.set(recruiterName, sentDate);
                        }

                        // Update Daily Count
                        const dateString = vsFormatDateRB(sentDate); // Use consistent format for map key
                        if (!recruiterDailyCounts.has(recruiterName)) {
                             recruiterDailyCounts.set(recruiterName, new Map());
                        }
                        const dailyMap = recruiterDailyCounts.get(recruiterName);
                        dailyMap.set(dateString, (dailyMap.get(dateString) || 0) + 1);
                    }
                }
            }
        });
        Logger.log(`Found last sent dates for ${recruiterLastSentMap.size} recruiters and daily counts.`);
    } else {
        Logger.log(`Recruiter_name column not found in log sheet, cannot calculate last sent activity or trends.`);
    }

    const recruiterActivityData = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Use start of today

    // Generate dates for the last 10 days (yesterday back to 10 days ago)
    const trendDates = [];
    for (let i = 1; i <= 10; i++) {
        const date = new Date(today);
        date.setDate(today.getDate() - i);
        trendDates.push(date);
    }

    recruiterLastSentMap.forEach((lastDate, recruiter) => {
        const timeDiff = today.getTime() - lastDate.getTime();
        const daysAgo = Math.floor(timeDiff / (1000 * 60 * 60 * 24)); // Calculate whole days

        // Build Daily Trend String
        const dailyMap = recruiterDailyCounts.get(recruiter);
        const trendValues = trendDates.map(date => {
            const dayOfWeek = date.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
            if (dayOfWeek === 0) return 'Sun';
            if (dayOfWeek === 6) return 'Sat';
            const dateString = vsFormatDateRB(date);
            return dailyMap?.get(dateString) || 0;
        });
        const dailyTrend = trendValues.join(',');

        recruiterActivityData.push({ recruiter: recruiter, daysAgo: daysAgo, dailyTrend: dailyTrend });
    });

    // Sort by days ago (most recent first), then alphabetically
    recruiterActivityData.sort((a, b) => {
        if (a.daysAgo !== b.daysAgo) {
            return a.daysAgo - b.daysAgo;
        }
        return a.recruiter.localeCompare(b.recruiter);
    });
    // <<< End Recruiter Last Sent Activity Calculation >>>

    // <<< RESTORED: Deduplicate by Profile_id + Position_id, prioritizing by status rank >>>
    const profileIdIndex = logData.colIndices['Profile_id'];
    const positionIdIndex = logData.colIndices['Position_id'];
    const statusIndex = logData.colIndices['STATUS_COLUMN']; // Get the index determined earlier
    const groupedData = {}; // Key: "profileId_positionId", Value: { bestRank: rank, row: rowData }
    let skippedRowCount = 0;

    finalFilteredData.forEach(row => {
        // Ensure row has the necessary columns
        if (!row || row.length <= profileIdIndex || row.length <= positionIdIndex || row.length <= statusIndex) {
            skippedRowCount++;
            // Logger.log(`Skipping row during grouping due to missing ID or Status columns. Row: ${JSON.stringify(row)}`);
            return; // Skip this row
        }
        const profileId = row[profileIdIndex];
        const positionId = row[positionIdIndex];
        const status = row[statusIndex] ? String(row[statusIndex]).trim() : 'Unknown';

        if (!profileId || !positionId) { // Check for blank IDs
             skippedRowCount++;
            // Logger.log(`Skipping row during grouping due to blank Profile_id or Position_id. Row: ${JSON.stringify(row)}`);
            return; // Skip rows with blank IDs
        }

        const uniqueKey = `${profileId}_${positionId}`;
        const currentRank = vsGetStatusRankRB(status); // Use RB helper

        if (!groupedData[uniqueKey] || currentRank < groupedData[uniqueKey].bestRank) {
            // If no entry exists OR current row has a better (lower) rank, store/replace it
            groupedData[uniqueKey] = { bestRank: currentRank, row: row };
        }
        // If an entry exists and current rank is not better, do nothing (keep the existing better row)
    });

    if (skippedRowCount > 0) {
        Logger.log(`Skipped ${skippedRowCount} rows during deduplication due to missing IDs, status, or incomplete row data.`);
    }

    // Extract the best row for each unique key
    const deduplicatedData = Object.values(groupedData).map(entry => entry.row);

    Logger.log(`Deduplicated data based on Profile_id + Position_id (prioritizing status). Count changed from ${finalFilteredData.length} to ${deduplicatedData.length}.`);

    // Check if any data remains after deduplication
    if (deduplicatedData.length === 0) {
         Logger.log(`No data remaining after deduplication. Skipping report.`);
         return;
    }
    // <<< END RESTORED BLOCK >>>

    // 3. Calculate Metrics (Uses RB calculator)
    const metrics = calculateCompanyMetricsRB(deduplicatedData, logData.colIndices);
    Logger.log('Successfully calculated company metrics with recruiter breakdown.');
    // Logger.log(`Calculated Metrics: ${JSON.stringify(metrics)}`); // Optional: Log detailed metrics

    // 4. Create HTML Report (Uses RB creator) - Pass adoption, activity data, and log recruiter index
    const htmlContent = createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, recruiterActivityData, recruiterNameIdx_Log);
    Logger.log('Successfully generated HTML report content.');

    // 5. Send Email (Uses RB functions/config)
    // Set static subject line for this specific report
    const reportTitle = `AI Recruiter Adoption: Daily Summary`; // <<< Renamed Subject
    sendVsEmailRB(VS_EMAIL_RECIPIENT_RB, VS_EMAIL_CC_RB, reportTitle, htmlContent);

    Logger.log(`--- AI Recruiter Adoption: Daily Summary generated and sent successfully! ---`); // Updated log message
    return `Report sent to ${VS_EMAIL_RECIPIENT_RB}`;

  } catch (error) {
    Logger.log(`Error in AIR_DailySummarytoAP: ${error.toString()} Stack: ${error.stack}`); // Updated log
    // Send error email (Uses RB notifier)
    sendVsErrorNotificationRB(`ERROR generating AI Recruiter Adoption: Daily Summary: ${error.toString()}`, error.stack);
    return `Error: ${error.toString()}`;
  }
}

// --- Data Retrieval and Processing Functions ---

/**
 * Reads and processes data from the Log_Enhanced sheet for Recruiter Breakdown.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null if error/no sheet/missing columns.
 */
function getLogSheetDataRB() {
  Logger.log(`Attempting to open log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(VS_LOG_SHEET_SPREADSHEET_URL_RB);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening log spreadsheet by URL: ${e}`);
    throw new Error(`Could not open the specified Log Spreadsheet URL. Please verify the URL is correct and accessible: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
  }

  let sheet = spreadsheet.getSheetByName(VS_LOG_SHEET_NAME_RB);

  // Fallback sheet finding logic
  if (!sheet) {
    Logger.log(`Log sheet "${VS_LOG_SHEET_NAME_RB}" not found by name. Attempting to use sheet by gid or first sheet.`);
    const gidMatch = VS_LOG_SHEET_SPREADSHEET_URL_RB.match(/gid=(\d+)/);
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
        throw new Error(`Could not find any sheets in the log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
      }
    }
  } else {
     Logger.log(`Using specified log sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) {
    Logger.log(`Not enough data in log sheet "${sheet.getName()}". Found ${data.length} rows. Expected headers + data.`);
    return null;
  }

  const headers = data[0].map(String);
  const rows = data.slice(1);

  // <<< DEBUGGING: Log the headers the script actually sees >>>
  Logger.log(`DEBUG: Headers found in sheet: ${JSON.stringify(headers)}`);
  // <<< END DEBUGGING >>>

  const requiredColumns = [
      'Interview_email_sent_at',
      'Profile_id',
      'Position_id',
      // Status column - prioritize Interview Status_Real
  ];
  const optionalColumns = [
      'Candidate_name',
      'Position_name',
      'Interview_status',
      'Interview Status_Real',
      'Schedule_start_time', 'Duration_minutes', 'Feedback_status', 'Feedback_json',
      'Match_stars', 'Location_country', 'Job_function', 'Position_id', 'Recruiter_name', // Ensure Recruiter_name is here
      'Creator_user_id', 'Reviewer_email', 'Hiring_manager_name',
      'Days_pending_invitation', 'Interview Status_Real'
  ];

  const colIndices = {};
  const missingCols = [];

  // --- Find Status Column --- Enforce Interview Status_Real ---
  const statusColName = 'Interview_status_real'; // <<< Updated name
  const statusColIndex = headers.indexOf(statusColName);
  if (statusColIndex !== -1) {
      colIndices['STATUS_COLUMN'] = statusColIndex;
      Logger.log(`Using column "${statusColName}" (index ${statusColIndex}) for interview status analysis.`);
  } else {
      missingCols.push(statusColName);
  }
  // --- End Find Status Column ---

  requiredColumns.forEach(colName => {
    const index = headers.indexOf(colName);
    if (index === -1) {
      missingCols.push(colName);
    } else {
      colIndices[colName] = index;
    }
  });

  // Check for Recruiter_name specifically as it's needed for the breakdown
  if (headers.indexOf('Recruiter_name') === -1) {
      Logger.log(`WARNING: Optional column "Recruiter_name" not found. Recruiter breakdown will show 'Unknown'.`);
  }

  if (missingCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in log sheet "${sheet.getName()}": ${missingCols.join(', ')}`);
    throw new Error(`Required column(s) not found in log sheet headers (Row 1): ${missingCols.join(', ')}`);
  }

  optionalColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index !== -1) {
          colIndices[colName] = index;
      } else if(colName !== 'Recruiter_name') { // Only log missing optional if not Recruiter_name (already warned)
          Logger.log(`Optional column "${colName}" not found.`);
      }
  });

  Logger.log(`Found required columns. Indices: ${JSON.stringify(colIndices)}`);
  return { rows, headers, colIndices };
}

/**
 * Reads and processes data from the Application Sheet (e.g., Active+Rejected) for the Adoption Chart.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null.
 */
function getApplicationDataForChartRB() {
  Logger.log(`--- Starting getApplicationDataForChartRB ---`);
  Logger.log(`Attempting to open application spreadsheet: ${APP_SHEET_SPREADSHEET_URL_RB}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(APP_SHEET_SPREADSHEET_URL_RB);
    Logger.log(`Opened application spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening application spreadsheet by URL: ${e}`);
    // Throw error as this data is essential for the requested chart
    throw new Error(`Could not open the Application Spreadsheet URL: ${APP_SHEET_SPREADSHEET_URL_RB}. Please verify the URL.`);
  }

  let sheet = spreadsheet.getSheetByName(APP_SHEET_NAME_RB);

  // Fallback sheet finding logic (similar to weekly report)
  if (!sheet) {
    Logger.log(`App sheet "${APP_SHEET_NAME_RB}" not found by name. Trying by GID or first sheet.`);
    const gidMatch = APP_SHEET_SPREADSHEET_URL_RB.match(/gid=(\d+)/);
    if (gidMatch && gidMatch[1]) {
      const gid = gidMatch[1];
      const sheets = spreadsheet.getSheets();
      sheet = sheets.find(s => s.getSheetId().toString() === gid);
      if (sheet) Logger.log(`Using app sheet by ID: "${sheet.getName()}"`);
    }
    if (!sheet) {
      sheet = spreadsheet.getSheets()[0];
      if (!sheet) {
        throw new Error(`No sheets found in application spreadsheet: ${APP_SHEET_SPREADSHEET_URL_RB}`);
      }
      Logger.log(`Warning: App sheet "${APP_SHEET_NAME_RB}" not found. Using first sheet: "${sheet.getName()}"`);
    }
  } else {
     Logger.log(`Using specified app sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Expect headers in Row 2, data starts Row 3 (like AIR_Weekly_Recruiter_Report.js)
  if (data.length < 3) {
    Logger.log(`Not enough data in app sheet "${sheet.getName()}" (expected headers in row 2). Cannot generate adoption chart.`);
    return null; // Return null, main function should handle this
  }

  const headers = data[1].map(String); // Headers from Row 2
  const rows = data.slice(2); // Data from Row 3 onwards

  Logger.log(`DEBUG: App Sheet Headers found in row 2: ${JSON.stringify(headers)}`);

  // --- Find Match Stars Column (Copied from weekly report logic) ---
  let matchStarsColIndex = -1;
  const exactMatchCol = 'Match_stars';
  matchStarsColIndex = headers.indexOf(exactMatchCol);
  if (matchStarsColIndex === -1) {
    Logger.log(`"${exactMatchCol}" column not found directly. Searching for alternatives...`);
    const possibleMatchColumns = ['Match_score', 'Match score', 'Match Stars', 'MatchStars', 'Match_Stars', 'Stars', 'Score'];
    for (const columnName of possibleMatchColumns) {
      matchStarsColIndex = headers.indexOf(columnName);
      if (matchStarsColIndex !== -1) {
        Logger.log(`Found match stars column as "${columnName}" at index ${matchStarsColIndex}`);
        break;
      }
    }
    // Add more fuzzy matching if needed here, similar to weekly report
  }
  if (matchStarsColIndex === -1) {
     Logger.log("WARNING: Could not find any suitable column for Match Stars/Score in App sheet. Adoption chart filter (≥4 Match) cannot be applied accurately.");
     // Proceed without it, the calculation function will handle this
  }
  // --- End Find Match Stars Column ---

  // Define columns needed for the adoption calculation
  const requiredAppColumns = [
      'Profile_id', 'Last_stage', 'Ai_interview', 'Recruiter name', 'Application_status', 'Position_status', 'Application_ts'
      // Add other columns if the weekly report's generateSegmentMetrics uses them
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

  // Add match stars index if found
  if (matchStarsColIndex !== -1) {
      appColIndices['Match_stars'] = matchStarsColIndex; // Use a consistent key
  }

  if (missingAppCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in app sheet "${sheet.getName()}" for adoption chart: ${missingAppCols.join(', ')}`);
    throw new Error(`Required column(s) for adoption chart not found in app sheet headers (Row 2): ${missingAppCols.join(', ')}`);
  }

  Logger.log(`Found required columns for app data chart. Indices: ${JSON.stringify(appColIndices)}`);
  return { rows, headers, colIndices: appColIndices };
}

/**
 * Calculates adoption metrics based on application data, mirroring the weekly report logic.
 * Filters for post-launch, >=4 match score (if possible), and calculates adoption based on eligibility.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} An object containing recruiter adoption data { recruiter: string, totalCandidates: number, takenAI: number, adoptionRate: number }.
 */
function calculateAdoptionMetricsForChartRB(appRows, appColIndices) {
  Logger.log(`--- Starting calculateAdoptionMetricsForChartRB ---`);

  const matchStarsColIndex = appColIndices.hasOwnProperty('Match_stars') ? appColIndices['Match_stars'] : -1;
  const launchDate = APP_LAUNCH_DATE_RB; // Use configured launch date
  const scoreThreshold = APP_MATCH_SCORE_THRESHOLD_RB; // Use configured threshold
  const applyMatchFilter = matchStarsColIndex !== -1;

  // 1. Filter for Post-Launch Date
  let postLaunchCandidates = appRows.filter(row => {
    const rawDate = row.length > appColIndices['Application_ts'] ? row[appColIndices['Application_ts']] : null;
    if (rawDate === null || rawDate === undefined || rawDate === '') return false;
    const applicationDate = vsParseDateSafeRB(rawDate); // Use RB helper
    return applicationDate && applicationDate >= launchDate;
  });
  Logger.log(`Total post-launch candidates (valid date): ${postLaunchCandidates.length}`);

  // 2. Filter by Match Score (if possible)
  let filteredCandidates = postLaunchCandidates;
  if (applyMatchFilter) {
    Logger.log(`Filtering segment by Match Score >= ${scoreThreshold}. Initial count: ${postLaunchCandidates.length}. Score Column Index: ${matchStarsColIndex}`);
    filteredCandidates = postLaunchCandidates.filter(row => {
      if (row.length <= matchStarsColIndex) return false;
      const scoreValue = row[matchStarsColIndex];
      const matchScore = parseFloat(scoreValue);
      return !isNaN(matchScore) && matchScore >= scoreThreshold;
    });
    Logger.log(`After match score filter, count: ${filteredCandidates.length}`);
  } else {
    Logger.log(`Match score filter not applied (column index: ${matchStarsColIndex}). Using all ${postLaunchCandidates.length} post-launch candidates for adoption chart.`);
  }

  // 3. Calculate Eligibility and Adoption (based on weekly report logic)
  const recruiterMap = {};
  let totalEligibleForRate = 0;
  let totalTakenForRate = 0;

  filteredCandidates.forEach(row => {
    const aiInterview = row.length > appColIndices['Ai_interview'] ? row[appColIndices['Ai_interview']] : null;
    const appStatus = row.length > appColIndices['Application_status'] ? row[appColIndices['Application_status']]?.toLowerCase() : null;
    const recruiter = (row.length > appColIndices['Recruiter name'] && row[appColIndices['Recruiter name']]) ? row[appColIndices['Recruiter name']] : 'Unassigned';

    let isEligible = false;
    let tookAI = false;

    if (aiInterview === 'Y') {
      isEligible = true;
      tookAI = true;
    } else if (aiInterview === 'N' || aiInterview === null || aiInterview === undefined || aiInterview === '') {
      // Eligible if not 'Y' AND not 'Rejected'
      if (appStatus !== 'rejected') {
        isEligible = true;
        tookAI = false;
      }
    }

    if (isEligible) {
      totalEligibleForRate++;
      if (!recruiterMap[recruiter]) {
        recruiterMap[recruiter] = { totalEligible: 0, taken: 0 };
      }
      recruiterMap[recruiter].totalEligible++;

      if (tookAI) {
        totalTakenForRate++;
        recruiterMap[recruiter].taken++;
      }
    }
  });

  Logger.log(`Adoption Chart Metrics: Total eligible (post-launch, >=${scoreThreshold} match) = ${totalEligibleForRate}. Total taken AI = ${totalTakenForRate}.`);

  // 4. Format Recruiter Data
  const recruiterAdoptionData = Object.keys(recruiterMap).map(recruiter => {
    const data = recruiterMap[recruiter];
    const adoptionRate = data.totalEligible > 0 ? parseFloat(((data.taken / data.totalEligible) * 100).toFixed(1)) : 0;
    return {
      recruiter: recruiter,
      totalCandidates: data.totalEligible, // Eligible candidates for this recruiter
      takenAI: data.taken,
      adoptionRate: adoptionRate
    };
  }).sort((a, b) => a.recruiter.localeCompare(b.recruiter)); // Sort alphabetically

  // Return structure expected by the chart generation code
  return { recruiterAdoptionData, hasMatchStarsColumn: matchStarsColIndex !== -1, matchScoreThreshold: scoreThreshold };
}

/**
 * Filters the data based on a time range (e.g., last N days based on Interview_email_sent_at).
 * @param {Array<Array>} rows The data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {Array<Array>} Filtered rows.
 */
function filterDataByTimeRangeRB(rows, colIndices) {
  if (!colIndices.hasOwnProperty('Interview_email_sent_at')) {
      Logger.log("WARNING: Cannot filter by time range - 'Interview_email_sent_at' column index not found.");
      return rows;
  }

  const sentAtIndex = colIndices['Interview_email_sent_at'];
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - VS_REPORT_TIME_RANGE_DAYS_RB); // Use RB config
  const cutoffTimestamp = cutoffDate.getTime();

  Logger.log(`Filtering data for interviews sent on or after ${cutoffDate.toLocaleDateString()}`);

  const filteredRows = rows.filter(row => {
    if (row.length <= sentAtIndex) return false;
    const rawDate = row[sentAtIndex];
    const sentDate = vsParseDateSafeRB(rawDate); // Use RB helper
    return sentDate && sentDate.getTime() >= cutoffTimestamp;
  });

  return filteredRows;
}


/**
 * Calculates company-level metrics including recruiter breakdown from the filtered data.
 * @param {Array<Array>} filteredRows The filtered data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {object} An object containing calculated metrics.
 */
function calculateCompanyMetricsRB(filteredRows, colIndices) {
  const COMPLETION_RATE_MATURITY_DAYS_RB = 1; // Exclude invites sent < 1 day ago for KPI box calc

  // Calculate the exact timestamp 24 hours ago from now
  const now = new Date();
  const cutoffTimestampForCompletionRate = now.getTime() - (48 * 60 * 60 * 1000); // Use 48 hours
  const cutoffDateForLog = new Date(cutoffTimestampForCompletionRate);
  Logger.log(`Calculating KPI box completion rate only for invites sent before ${cutoffDateForLog.toLocaleString()} (48-hour cutoff)`);


  const metrics = {
    reportStartDate: (() => { const d = new Date(); d.setDate(d.getDate() - VS_REPORT_TIME_RANGE_DAYS_RB); return vsFormatDateRB(d); })(), // Use RB config/helpers
    reportEndDate: vsFormatDateRB(new Date()), // Use RB helper
    totalSent: filteredRows.length, // This remains the absolute total after filtering/deduplication
    totalScheduled: 0,
    totalCompleted: 0, // Absolute total completed
    totalFeedbackSubmitted: 0,
    sentToScheduledRate: 0,
    scheduledToCompletedRate: 0,
    completedToFeedbackRate: 0,
    sentToScheduledDaysSum: 0,
    sentToScheduledCount: 0,
    completedToFeedbackDaysSum: 0,
    completedToFeedbackCount: 0,
    matchStarsSum: 0,
    matchStarsCount: 0,
    completionRateByJobFunction: {}, // Kept for consistency, but maybe removed if not needed
    avgTimeToFeedbackByCountry: {}, // Kept for consistency
    interviewStatusDistribution: {},
    // Raw data storage for breakdowns
    byJobFunction: {},
    byCountry: {},
    byRecruiter: {}, // <<< ADDED Recruiter Breakdown
    // Timeseries data
    dailySentCounts: {},
    // --- Counters and Rates for KPI Box ---
    matureKpiTotalSent: 0,       // Denominator for adjusted KPI rate
    matureKpiTotalCompleted: 0,  // Numerator for adjusted KPI rate
    kpiCompletionRateAdjusted: 0,// The adjusted rate (%) for the KPI box
    completionRateOriginal: 0    // The original rate (%) using all invites (for footnote)
  };

  // --- Status Definitions (Consistent) ---
  const STATUSES_FOR_AVG_TIME_CALC = ['SCHEDULED', 'COMPLETED'];
  const COMPLETED_STATUSES = ['COMPLETED']; // <<< UPDATED: Strict definition for all metrics
  const PENDING_STATUSES = ['PENDING', 'INVITED', 'EMAIL SENT'];
  const FEEDBACK_SUBMITTED_STATUS = 'Submitted';
  const RECRUITER_SUBMISSION_AWAITED_FEEDBACK = 'AI_RECOMMENDED';

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
  const recruiterIdx = colIndices.hasOwnProperty('Recruiter_name') ? colIndices['Recruiter_name'] : -1; // <<< GET Recruiter Index

  filteredRows.forEach(row => {
    // <<< MOVED: Define core values at the beginning of the loop >>>
    const statusRaw = row[statusIdx] ? String(row[statusIdx]).trim() : 'Unknown';
    const jobFunc = (jobFuncIdx !== -1 && row[jobFuncIdx]) ? String(row[jobFuncIdx]).trim() : 'Unknown';
    const country = (countryIdx !== -1 && row[countryIdx]) ? String(row[countryIdx]).trim() : 'Unknown';
    const recruiter = (recruiterIdx !== -1 && row[recruiterIdx]) ? String(row[recruiterIdx]).trim() : 'Unknown'; // <<< GET Recruiter Name
    const feedbackStatusRaw = (feedbackStatusIdx !== -1 && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim() : '';

    // --- Get Sent Date ---
    const sentDate = vsParseDateSafeRB(row[sentAtIdx]); // Use RB helper
    const isMatureForCompletionRate = sentDate && sentDate.getTime() < cutoffTimestampForCompletionRate; // Check if sent *before* the exact 24hr cutoff timestamp

    // --- Increment Mature Sent Count for KPI ---
    if (isMatureForCompletionRate) {
        metrics.matureKpiTotalSent++;
    }

    // --- Daily Sent Counts ---
    if (sentDate) {
        const dateString = vsFormatDateRB(sentDate); // Use RB helper
        metrics.dailySentCounts[dateString] = (metrics.dailySentCounts[dateString] || 0) + 1;
    }

    // --- Initialize Breakdown Structures if they don't exist ---
    if (!metrics.byJobFunction[jobFunc]) {
        metrics.byJobFunction[jobFunc] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }
    if (!metrics.byCountry[country]) {
        metrics.byCountry[country] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, statusCounts: {} };
    }
    if (!metrics.byRecruiter[recruiter]) { // <<< INITIALIZE Recruiter
        metrics.byRecruiter[recruiter] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }

    // --- Increment Base Counts (These always use the total number of records processed) ---
    metrics.byJobFunction[jobFunc].sent++;
    metrics.byCountry[country].sent++;
    metrics.byRecruiter[recruiter].sent++; // <<< INCREMENT Recruiter Sent
    metrics.interviewStatusDistribution[statusRaw] = (metrics.interviewStatusDistribution[statusRaw] || 0) + 1;
    metrics.byJobFunction[jobFunc].statusCounts[statusRaw] = (metrics.byJobFunction[jobFunc].statusCounts[statusRaw] || 0) + 1;
    metrics.byCountry[country].statusCounts[statusRaw] = (metrics.byCountry[country].statusCounts[statusRaw] || 0) + 1;
    metrics.byRecruiter[recruiter].statusCounts[statusRaw] = (metrics.byRecruiter[recruiter].statusCounts[statusRaw] || 0) + 1; // <<< INCREMENT Recruiter Status Count

    // --- Calculate Avg Time Sent to Completion (Scheduled) ---
    // (No change needed here)
    if (STATUSES_FOR_AVG_TIME_CALC.includes(statusRaw)) {
        const candidateName = (candidateNameIdx !== -1 && row[candidateNameIdx]) ? row[candidateNameIdx] : 'Unknown Candidate';
        const scheduleDateForAvg = (scheduledAtIdx !== -1) ? vsParseDateSafeRB(row[scheduledAtIdx]) : null; // Use RB helper
        if (sentDate && scheduleDateForAvg) {
            const daysDiff = vsCalculateDaysDifferenceRB(sentDate, scheduleDateForAvg); // Use RB helper
            if (daysDiff !== null) {
                metrics.sentToScheduledDaysSum += daysDiff;
                metrics.sentToScheduledCount++;
                // Logger.log(`DEBUG_AVG_TIME: Candidate=${candidateName}, Status=${statusRaw}, Sent=${sentDate.toISOString()}, Scheduled=${scheduleDateForAvg.toISOString()}, DiffDays=${daysDiff.toFixed(8)}`);
            }
        }
    }

    // --- Check if Scheduled (for breakdown counts) ---
    let isScheduledForCount = (statusRaw === 'SCHEDULED');
    if (isScheduledForCount) {
         metrics.totalScheduled++;
         metrics.byJobFunction[jobFunc].scheduled++;
         metrics.byCountry[country].scheduled++;
         metrics.byRecruiter[recruiter].scheduled++; // <<< INCREMENT Recruiter Scheduled
    }

    // --- Check if Pending ---
    if (PENDING_STATUSES.includes(statusRaw)) {
        metrics.byJobFunction[jobFunc].pending++;
        metrics.byCountry[country].pending++;
        metrics.byRecruiter[recruiter].pending++; // <<< INCREMENT Recruiter Pending
    }

    // --- Check if Completed ---
    let isCompleted = COMPLETED_STATUSES.includes(statusRaw);
    if (isCompleted) {
      metrics.totalCompleted++; // Increment original total completed
      metrics.byJobFunction[jobFunc].completed++; // Increment original breakdown completed
      metrics.byCountry[country].completed++;     // Increment original breakdown completed
      metrics.byRecruiter[recruiter].completed++; // Increment original breakdown completed

      // Increment Mature Completed Count for KPI (ONLY if mature)
      if (isMatureForCompletionRate) {
          metrics.matureKpiTotalCompleted++;
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
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw === FEEDBACK_SUBMITTED_STATUS) {
         metrics.totalFeedbackSubmitted++;
         metrics.byJobFunction[jobFunc].feedbackSubmitted++;
         metrics.byCountry[country].feedbackSubmitted++;
         metrics.byRecruiter[recruiter].feedbackSubmitted++; // <<< INCREMENT Recruiter Feedback Submitted
       }

       // --- Check for Recruiter Submission Awaited (AI_RECOMMENDED in Feedback_status)
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw === RECRUITER_SUBMISSION_AWAITED_FEEDBACK) {
           metrics.byJobFunction[jobFunc].recruiterSubmissionAwaited++;
           // Note: No country-specific count for this yet
           metrics.byRecruiter[recruiter].recruiterSubmissionAwaited++; // <<< INCREMENT Recruiter Submission Awaited
       }
    }
  });

  // --- Calculate Final Rates and Averages ---

  // Calculate ORIGINAL Completion Rate (for footnote)
  if (metrics.totalSent > 0) {
      metrics.completionRateOriginal = parseFloat(((metrics.totalCompleted / metrics.totalSent) * 100).toFixed(1));
  }

  // Calculate ADJUSTED Completion Rate (for KPI Box)
  if (metrics.matureKpiTotalSent > 0) {
      metrics.kpiCompletionRateAdjusted = parseFloat(((metrics.matureKpiTotalCompleted / metrics.matureKpiTotalSent) * 100).toFixed(1));
  }

  // Calculate other original rates
  if (metrics.totalSent > 0) {
      metrics.sentToScheduledRate = parseFloat(((metrics.totalScheduled / metrics.totalSent) * 100).toFixed(1));
      // Update status distribution calculation (uses totalSent)
      const statusCountsTemp = { ...metrics.interviewStatusDistribution };
      metrics.interviewStatusDistribution = {};
      for (const status in statusCountsTemp) {
          const count = statusCountsTemp[status];
          metrics.interviewStatusDistribution[status] = {
              count: count,
              percentage: parseFloat(((count / metrics.totalSent) * 100).toFixed(1))
          };
      }
  }
  if (metrics.totalScheduled > 0) {
      metrics.scheduledToCompletedRate = parseFloat(((metrics.totalCompleted / metrics.totalScheduled) * 100).toFixed(1));
  }
   if (metrics.totalCompleted > 0) {
      metrics.completedToFeedbackRate = parseFloat(((metrics.totalFeedbackSubmitted / metrics.totalCompleted) * 100).toFixed(1));
      if(metrics.matchStarsCount > 0) {
          metrics.avgMatchStars = parseFloat((metrics.matchStarsSum / metrics.matchStarsCount).toFixed(1));
      } else {
          metrics.avgMatchStars = null; // Ensure null if no stars
      }
   } else {
      metrics.avgMatchStars = null; // Ensure null if no completions
   }
   if (metrics.sentToScheduledCount > 0) {
       metrics.avgTimeToScheduleDays = parseFloat((metrics.sentToScheduledDaysSum / metrics.sentToScheduledCount).toFixed(1));
   } else {
       metrics.avgTimeToScheduleDays = null;
   }
    if (metrics.completedToFeedbackCount > 0) {
        metrics.avgCompletedToFeedbackDays = parseFloat((metrics.completedToFeedbackDaysSum / metrics.completedToFeedbackCount).toFixed(1));
    } else {
         metrics.avgCompletedToFeedbackDays = null; // Example, if calculation added later
    }

    // <<< DEBUG LOGGING for KPI Rate >>>
    Logger.log(`KPI Rate Calculation: Mature Sent (Denominator) = ${metrics.matureKpiTotalSent}`);
    Logger.log(`KPI Rate Calculation: Mature Completed [Strict 'COMPLETED'] (Numerator) = ${metrics.matureKpiTotalCompleted}`);
    if (metrics.matureKpiTotalSent > 0) {
        Logger.log(`KPI Rate Calculation: Adjusted Rate = (${metrics.matureKpiTotalCompleted} / ${metrics.matureKpiTotalSent}) * 100 = ${metrics.kpiCompletionRateAdjusted}%`);
    } else {
        Logger.log(`KPI Rate Calculation: Adjusted Rate = N/A (Mature Sent is 0)`);
    }
    // <<< END DEBUG LOGGING >>>

  // --- Calculate Breakdown Metrics (Using ORIGINAL 'sent' and 'completed' counts for percentages) ---
  // Job Functions
  for (const func in metrics.byJobFunction) {
    const data = metrics.byJobFunction[func];
    data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0;
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0;
  }

  // Countries
  for (const ctry in metrics.byCountry) {
    const data = metrics.byCountry[ctry];
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    // Add other country-specific metrics here if needed
  }

  // Recruiters <<< CALCULATE Recruiter Breakdown Metrics (Using ORIGINAL counts)
  for (const rec in metrics.byRecruiter) {
    const data = metrics.byRecruiter[rec];
    // data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0; // Optional
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0; // Optional
  }


  Logger.log(`Metrics calculation complete (Recruiter). Total Sent: ${metrics.totalSent}, Completed: ${metrics.totalCompleted}`);
  Logger.log(`KPI Completion Rate (Adjusted): ${metrics.kpiCompletionRateAdjusted}%, Original Rate: ${metrics.completionRateOriginal}%`); // Log both rates
  metrics.colIndices = colIndices;
  return metrics;
}

// --- Reporting Functions ---

/**
 * Generates the HTML for the table rows of the Recruiter Breakdown section.
 * Sorts recruiters by 'Sent' count descending and adds medals.
 * @param {object} recruiterData The metrics.byRecruiter object.
 * @returns {string} HTML string for the table body rows.
 */
function generateRecruiterTableRowsHtml(recruiterData) {
    if (!recruiterData || Object.keys(recruiterData).length === 0) {
        return '<tr><td colspan="7" style="text-align:center; padding: 10px; border: 1px solid #e0e0e0; font-size: 12px;">No recruiter data found or Recruiter_name column missing.</td></tr>';
    }

    // Sort recruiters by Sent descending, keeping Unknown last
    const sortedRecruiters = Object.entries(recruiterData)
        .sort(([recA, dataA], [recB, dataB]) => {
            if (recA === 'Unknown') return 1;
            if (recB === 'Unknown') return -1;
            return dataB.sent - dataA.sent;
        });

    // Create a map for medals for the top 3 (excluding Unknown)
    const medals = ['🥇', '🥈', '🥉'];
    const recruiterMedalMap = {};
    sortedRecruiters
        .filter(([rec]) => rec !== 'Unknown')
        .slice(0, 3)
        .forEach(([rec], index) => {
            recruiterMedalMap[rec] = medals[index];
        });

    // Generate table rows HTML
    return sortedRecruiters
        .map(([rec, data], index) => {
            const medal = recruiterMedalMap[rec] || ''; // Get medal or empty string
            const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
            return `
                <tr style="background-color: ${bgColor};">
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold;">${medal}${medal ? ' ' : ''}${rec}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.recruiterSubmissionAwaited}</td>
                </tr>
            `;
        }).join('');
}

/**
 * Creates the HTML email report including recruiter breakdown.
 * Uses inline styles and table layouts for better email client compatibility.
 * @param {object} metrics The calculated metrics object.
 * @param {object} adoptionChartData The calculated adoption chart data object.
 * @param {Array<object>} recruiterActivityData Array of {recruiter: string, daysAgo: number, dailyTrend: string}.
 * @param {number} recruiterNameIdx_Log The index of the Recruiter_name column from the log sheet (-1 if not found).
 * @returns {string} The HTML content for the email body.
 */
function createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, recruiterActivityData, recruiterNameIdx_Log) {

  // Helper to generate timeseries table (limited to last 7 days)
  const generateTimeseriesTable = (dailyCounts) => {
      const sortedDates = Object.keys(dailyCounts).sort((a, b) => {
          try {
              // Parsing DD-MMM-YY format
              const dateA = new Date(a.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              const dateB = new Date(b.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              return dateB - dateA; // Descending
          } catch (e) { return b.localeCompare(a); }
      });

      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      sevenDaysAgo.setHours(0, 0, 0, 0);

      const filteredDates = sortedDates.filter(dateStr => {
          try {
              const date = new Date(dateStr.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              return date >= sevenDaysAgo;
          } catch (e) { return false; }
      });

      if (filteredDates.length === 0) return '<p style="font-size: 0.85em; color: #757575; margin-top: 15px; text-align: center;">No interview invitations sent in the last 7 days.</p>';

      let tableHtml = '<table align="center" border="0" cellpadding="0" cellspacing="0" width="90%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;"><thead><tr><th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">🗓️ Date (DD-MMM-YY)</th><th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">✉️ Invitations Sent</th></tr></thead><tbody>';
      filteredDates.forEach((date, index) => {
          const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
          tableHtml += `<tr style="background-color: ${bgColor};"><td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle;">${date}</td><td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${dailyCounts[date]}</td></tr>`;
      });
      tableHtml += '</tbody></table>';
      return tableHtml;
  };

  const recruiterIdx = metrics.colIndices && metrics.colIndices.hasOwnProperty('Recruiter_name') ? metrics.colIndices['Recruiter_name'] : -1;
  const hasAdoptionData = adoptionChartData && adoptionChartData.recruiterAdoptionData && adoptionChartData.recruiterAdoptionData.length > 0;
  const hasMatchStarsColumnForAdoption = adoptionChartData && adoptionChartData.hasMatchStarsColumn;
  const adoptionScoreThreshold = adoptionChartData ? (adoptionChartData.matchScoreThreshold || APP_MATCH_SCORE_THRESHOLD_RB) : APP_MATCH_SCORE_THRESHOLD_RB;
  const maturityDays = 1; // Define here or get from a constant if needed elsewhere in HTML


  let html = `<!DOCTYPE html>
<html>
<head>
  <title>${VS_COMPANY_NAME_RB} AI Interview Recruiter Report</title>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <!-- No <style> block needed - all styles are inline -->
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 10px; margin: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
  <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 850px;">
    <tr>
      <td align="center">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #ffffff; border: 1px solid #ccc; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); margin: 20px auto; padding: 25px;">
          <!-- Header -->
          <tr>
            <td style="padding-bottom: 15px; margin-bottom: 25px; border-bottom: 2px solid #eee;">
              <h1 style="color: #1a237e; text-align: center; font-size: 26px; margin: 0;">AI Recruiter Adoption: Daily Summary</h1> <!-- <<< Renamed Header -->
            </td>
          </tr>

          <!-- Top KPI Boxes - Table Layout -->
          <tr>
            <td style="padding-top: 25px; padding-bottom: 10px;">
              <table border="0" cellpadding="0" cellspacing="15" width="100%" style="border-collapse: separate; table-layout: fixed;">
                <tr>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" style="height: 130px; border: 1px solid #cccccc; border-radius: 8px; border-collapse: collapse; table-layout: fixed; overflow: hidden; background-color: #e8f5e9;">
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 6px 10px; font-size: 12px; font-weight: bold; color: #424242; border-bottom: 1px solid #cccccc; height: 30px;">✉️ AI Invitations Sent</td></tr>
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 10px; font-size: 34px; font-weight: bold; height: 100%; color: #2e7d32;">
                          ${metrics.totalSent}
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" style="height: 130px; border: 1px solid #cccccc; border-radius: 8px; border-collapse: collapse; table-layout: fixed; overflow: hidden; background-color: #e3f2fd;">
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 6px 10px; font-size: 12px; font-weight: bold; color: #424242; border-bottom: 1px solid #cccccc; height: 30px;">✅ Completion Rate</td></tr>
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 10px; font-size: 34px; font-weight: bold; height: 100%; color: #1976d2;">
                          ${metrics.kpiCompletionRateAdjusted}<span style="font-size: 16px; font-weight: normal; margin-left: 3px;">%</span>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" style="height: 130px; border: 1px solid #cccccc; border-radius: 8px; border-collapse: collapse; table-layout: fixed; overflow: hidden; background-color: #fff3e0;">
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 6px 10px; font-size: 12px; font-weight: bold; color: #424242; border-bottom: 1px solid #cccccc; height: 30px;">⏱️ Avg Time Sent to Completion*</td></tr>
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 10px; font-size: 34px; font-weight: bold; height: 100%; color: #ef6c00;">
                          ${metrics.avgTimeToScheduleDays !== null ? metrics.avgTimeToScheduleDays : 'N/A'}<span style="font-size: 16px; font-weight: normal; margin-left: 3px;">days</span>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" style="height: 130px; border: 1px solid #cccccc; border-radius: 8px; border-collapse: collapse; table-layout: fixed; overflow: hidden; background-color: #f3e5f5;">
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 6px 10px; font-size: 12px; font-weight: bold; color: #424242; border-bottom: 1px solid #cccccc; height: 30px;">⭐ Avg Match Stars (Completed)</td></tr>
                      <tr><td style="border: none; vertical-align: middle; text-align: center; padding: 10px; font-size: 34px; font-weight: bold; height: 100%; color: #8e24aa;">
                          ${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <!-- New row for the footnote -->
                <tr>
                  <td colspan="4" style="text-align: center; padding-top: 5px; font-size: 10px; color: #666;">
                    (Completion Rate KPI Excl. invites sent < ${maturityDays} day)
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Side-by-side Sections - Table Layout -->
          <tr>
            <td style="padding-bottom: 15px;">
              <table border="0" cellpadding="0" cellspacing="15" width="100%" style="border-collapse: separate; table-layout: fixed;">
                <tr>
                  <!-- Left Cell: Completion Status -->
                  <td width="50%" style="vertical-align: top; padding: 0;">
                    <div style="background-color: #fff; padding: 15px; border: 1px solid #eee; border-radius: 4px;">
                      <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📊 AI Screening Completion Status</div>
                      <table align="center" border="0" cellpadding="0" cellspacing="0" width="95%" style="border-collapse: collapse; table-layout: fixed; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;"> <!-- Added table-layout: fixed -->
                         <thead><tr><th style="width: 50%; border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Status</th><th style="width: 60px; border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Count</th><th style="width: 60px; border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">%</th></tr></thead> <!-- Restored Status Header & Added Widths -->
                 <tbody>
                 ${Object.entries(metrics.interviewStatusDistribution)
                             .sort(([, dataA], [, dataB]) => dataB.count - dataA.count) // Sort by count descending
                             .map(([status, data], index) => `<tr style="background-color: ${index % 2 === 0 ? '#fafafa' : '#ffffff'};"> <!-- Removed title attribute -->
                                 <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; word-wrap: break-word;">${status}</td> <!-- Restored Status Cell, Added word-wrap -->
                                 <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.count}</td>
                                 <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle; color: #0056b3;">${data.percentage}%</td>
                             </tr>`).join('')}
                 </tbody>
             </table>
                     <p style="font-size: 0.85em; color: #757575; margin-top: 15px;">Percentage is based on the total number of invitations sent since 17th April 2025 (Launch of AIR).</p>
            </div>
          </td>
                  <!-- Right Cell: Daily Invitations -->
                  <td width="50%" style="vertical-align: top; padding: 0;">
                     <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
                       <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">🗓️ Daily Invitations Sent (Last 7 Days)</div>
                       ${generateTimeseriesTable(metrics.dailySentCounts)}
            </div>
          </td>
        </tr>
      </table>
            </td>
          </tr>

          <!-- Breakdown by Recruiter (Table) -->
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
                 <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">🧑‍💼 Breakdown by Recruiter</div>
                 <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
             <thead>
                <tr>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Name</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Sent</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Completed (# / %)</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Scheduled</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Pending (# / %)</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Feedback Submitted</th>
                           <th style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Submission Awaited</th>
                 </tr>
             </thead>
             <tbody>
                        ${(() => { // Start IIFE to contain logic
                            // Sort recruiters by Sent descending, keeping Unknown last
                            const sortedRecruiters = Object.entries(metrics.byRecruiter)
                                .sort(([recA, dataA], [recB, dataB]) => {
                          if (recA === 'Unknown') return 1;
                          if (recB === 'Unknown') return -1;
                                    return dataB.sent - dataA.sent;
                                });

                            // Create a map for medals for the top 3 (excluding Unknown)
                            const medals = ['🥇', '🥈', '🥉'];
                            const recruiterMedalMap = {};
                            sortedRecruiters
                                .filter(([rec]) => rec !== 'Unknown')
                                .slice(0, 3)
                                .forEach(([rec], index) => {
                                    recruiterMedalMap[rec] = medals[index];
                                });

                            // Generate table rows
                            return sortedRecruiters
                             .map(([rec, data], index) => {
                                const medal = recruiterMedalMap[rec] || ''; // Get medal or empty string
                                return `
                                  <tr style="background-color: ${index % 2 === 0 ? '#fafafa' : '#ffffff'};">
                                     <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold;">${medal}${medal ? ' ' : ''}${rec}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 3px 8px; text-align: center; font-size: 12px; vertical-align: middle;">${data.recruiterSubmissionAwaited}</td>
                                  </tr>
                                `;
                            }).join('');
                        })()}
                          ${Object.keys(metrics.byRecruiter).length === 0 ? '<tr><td colspan="7" style="text-align:center; padding: 10px; border: 1px solid #e0e0e0; font-size: 12px;">No recruiter data found or Recruiter_name column missing.</td></tr>' : ''}
                      </tbody>
                  </table>
                 ${recruiterIdx === -1 ? '<p style="font-size: 0.85em; color: #757575; margin-top: 15px;">Recruiter breakdown is based on the "Recruiter_name" column, which was not found in the sheet.</p>' : ''}
              </div>
            </td>
          </tr>

          <!-- Recruiter Last Activity Table (Moved Up) -->
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
                <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">⏱️ Recruiter Last Invite Activity</div>
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="95%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
                  <thead>
                    <tr>
                      <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Name</th>
                      <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Last Invite Sent</th>
                      <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase; width: 250px;">Daily Trend (Last 10 Days)</th> <!-- Added Header -->
                         </tr>
                  </thead>
                  <tbody>
                    ${recruiterActivityData && recruiterActivityData.length > 0 ?
                        recruiterActivityData.map((activity, index) => {
                            const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
                            const daysAgoText = activity.daysAgo === 0 ? 'Yesterday' : `${activity.daysAgo} calendar ${activity.daysAgo === 1 ? 'day' : 'days'} ago`;
                            return `
                            <tr style="background-color: ${bgColor};">
                              <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold;">${activity.recruiter}</td>
                              <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${daysAgoText}</td>
                              <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; font-family: monospace;">${activity.dailyTrend || 'N/A'}</td> <!-- Added Trend Column -->
                            </tr>`;
                        }).join('')
                        :
                        '<tr><td colspan="3" style="text-align:center; border: 1px solid #e0e0e0; padding: 10px; color: #777; font-size: 12px;">No recruiter activity data found or Recruiter_name column missing.</td></tr>' // Updated colspan
                    }
             </tbody>
         </table>
                 ${recruiterNameIdx_Log === -1 ? '<p style="font-size: 0.85em; color: #757575; margin-top: 15px;">Recruiter activity based on the "Recruiter_name" column, which was not found in the log sheet.</p>' : ''}
     </div>
            </td>
          </tr>

          <!-- Adoption Rate by Recruiter Bar Chart -->
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
                <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📊 AI Adoption Rate by Recruiter (Post-Launch, ≥${adoptionScoreThreshold} Match Score)</div>
                ${!hasMatchStarsColumnForAdoption ?
                  '<div style="padding: 10px; color: #cc3300; text-align: center; font-size: 0.9em;"><b>Warning:</b> Match score column not found in Application sheet. Filter could not be applied.</div>' : ''}
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 15px 0; border: 1px solid #ddd; padding: 15px; background-color: white; border-radius: 5px; box-sizing: border-box; border-collapse: separate; border-spacing: 0 8px;">
                   <tbody>
                   ${hasAdoptionData ?
                       adoptionChartData.recruiterAdoptionData
                       .sort((a, b) => { if (a.recruiter === 'Unassigned') return 1; if (b.recruiter === 'Unassigned') return -1; return a.recruiter.localeCompare(b.recruiter); })
                       .map(data => {
                           const adoptionRate = data.adoptionRate || 0;
                           const displayWidth = Math.max(adoptionRate, 0);
                           const takenCount = data.takenAI || 0;
                           const eligibleCount = data.totalCandidates || 0;
                           return `
                           <tr>
                             <td style="width: 220px; padding-right: 10px; text-align: left; font-weight: bold; font-size: 13px; overflow: hidden; text-overflow: ellipsis; color: #222; box-sizing: border-box; border: none; padding: 0; vertical-align: middle; height: 24px; line-height: 24px;" title="${data.recruiter} (${takenCount}/${eligibleCount})">
                               ${data.recruiter} (${takenCount}/${eligibleCount})<span style="color: #CC5500; font-weight: bold;"> [${adoptionRate}%]</span>
                             </td>
                             <td style="width: auto; border: none; padding: 0; vertical-align: middle; height: 24px; line-height: 24px;">
                               <div style="width: 100%; height: 100%; background-color: #f0f0f0; border-radius: 3px; position: relative; border: 1px solid #cccccc; box-sizing: border-box;">
                                 <div style="width: ${displayWidth}%; height: 100%; background-color: #4CAF50; border-radius: 3px;"></div>
                               </div>
                             </td>
                           </tr>`;
                       }).join('')
                       :
                       '<tr><td colspan="2" style="text-align:center; border: none; padding: 10px; color: #777;">No adoption data found or could not be calculated.</td></tr>'
                   }
                   </tbody>
                </table>
                <table border="0" cellpadding="0" cellspacing="10" align="center" style="margin-top: 15px; font-size: 12px;">
                   <tr>
                     <td style="vertical-align: middle;"><div style="width: 14px; height: 14px; border: 1px solid #ccc; background-color: #4CAF50;"></div></td>
                     <td style="vertical-align: middle;">Invited (Eligible)</td>
                     <td style="padding-left: 15px; vertical-align: middle;"><div style="width: 14px; height: 14px; border: 1px solid #ccc; background-color: #f0f0f0;"></div></td>
                     <td style="vertical-align: middle;">Not Invited (Eligible)</td>
                   </tr>
                </table>
                <p style="font-size: 11px; text-align: center; color: #757575; margin-top: 15px;">Adoption rate calculated as (Invited / Eligible). Eligible = Not Rejected and AI Interview is 'Y', 'N', or blank. Filtered for applications since ${APP_LAUNCH_DATE_RB.toLocaleDateString()}${hasMatchStarsColumnForAdoption ? ` and Match Score >= ${adoptionScoreThreshold}` : ' (Match Score filter not applied)'}.</p>
              </div>
            </td>
          </tr>

     <!-- Looker Studio Link -->
          <tr>
            <td style="text-align: center; padding-top: 10px; padding-bottom: 10px;">
              <a href="https://lookerstudio.google.com/u/0/reporting/b05c1dfb-d808-4eca-b70d-863fe5be0f27/page/p_58go7mgqrd" target="_blank" style="display: inline-block; background-color: #4285F4; color: #ffffff; padding: 10px 20px; text-align: center; text-decoration: none; border-radius: 4px; font-size: 14px; margin-top: 10px;">
         View Detailed Looker Studio Report
       </a>
            </td>
          </tr>

     <!-- Breakdown by Job Function -->
          <tr>
             <td style="padding-top: 10px; padding-bottom: 10px;">
               <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
                  <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">💼 Breakdown by Job Function</div>
                  <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
             <thead>
                <tr>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Job Function</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Sent</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Completed (# / %)</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Scheduled</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Pending (# / %)</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Feedback Submitted</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Submission Awaited</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byJobFunction)
                     .sort(([funcA], [funcB]) => funcA.localeCompare(funcB))
                              .map(([func, data], index) => `
                                  <tr style="background-color: ${index % 2 === 0 ? '#fafafa' : '#ffffff'};">
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold;">${func}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.recruiterSubmissionAwaited}</td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>
             </td>
          </tr>

     <!-- Breakdown by Location Country -->
           <tr>
             <td style="padding-top: 10px; padding-bottom: 10px;">
               <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
                  <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">🌍 Breakdown by Location Country</div>
                   <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
             <thead>
                <tr>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Country</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Sent</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Completed (# / %)</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Scheduled</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Pending (# / %)</th>
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Feedback Submitted</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byCountry)
                     .sort(([ctryA], [ctryB]) => ctryA.localeCompare(ctryB))
                              .map(([ctry, data], index) => `
                                  <tr style="background-color: ${index % 2 === 0 ? '#fafafa' : '#ffffff'};">
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold;">${ctry}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>
             </td>
          </tr>

          <!-- Footer Note -->
          <tr>
            <td style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
              <p style="font-size: 0.85em; color: #757575; margin: 0;">
                *Avg Time Sent to Completion calculation currently uses Schedule Start Date as completion proxy.<br>
                **Completion Rate KPI** excludes invitations sent within the last 48 hours. Breakdown table % includes all sent invites.<br>
                Overall completion rate (including all invites): ${metrics.completionRateOriginal}%.<br>
                Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.
              </p>
              ${recruiterIdx === -1 ? '<p style="font-size: 0.85em; color: red; margin-top: 5px;">Warning: "Recruiter_name" column not found in the sheet. Recruiter breakdown data will be limited or inaccurate.</p>' : ''}
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  return html;
}


/**
 * Sends an email with the recruiter breakdown report.
 * @param {string} recipient The primary email recipient.
 * @param {string} ccRecipient The CC email recipient (can be empty).
 * @param {string} subject The email subject.
 * @param {string} htmlBody The HTML content of the email.
 */
function sendVsEmailRB(recipient, ccRecipient, subject, htmlBody) {
  if (!recipient) {
     Logger.log("ERROR: Email recipient (RB) is empty. Cannot send email.");
     return;
  }
   if (!subject) {
     Logger.log("WARNING: Email subject (RB) is empty. Using default subject.");
     subject = `${VS_COMPANY_NAME_RB} AI Interview Recruiter Report`;
  }
   if (!htmlBody) {
     Logger.log("ERROR: Email body (RB) is empty. Cannot send email.");
     return;
  }

  const options = {
     to: recipient,
     subject: subject,
     htmlBody: htmlBody
  };

  if (ccRecipient && ccRecipient.trim() !== '' && ccRecipient.trim().toLowerCase() !== recipient.trim().toLowerCase()) {
    options.cc = ccRecipient;
    Logger.log(`Sending recruiter report email to ${recipient}, CC ${ccRecipient}`);
  } else {
     Logger.log(`Sending recruiter report email to ${recipient} (No CC or CC is same as recipient)`);
  }

  try {
      MailApp.sendEmail(options);
      Logger.log("Recruiter report email sent successfully.");
  } catch (e) {
     Logger.log(`ERROR sending recruiter report email: ${e.toString()}`);
     sendVsErrorNotificationRB(`CRITICAL: Failed to send recruiter report email to ${recipient}`, `Error: ${e.toString()}`); // Use RB notifier
  }
}

/**
 * Sends an error notification email for the Recruiter Breakdown script.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendVsErrorNotificationRB(errorMessage, stackTrace = '') {
   const recipient = VS_EMAIL_RECIPIENT_RB; // Use RB config
   if (!recipient) {
       Logger.log("CRITICAL ERROR: Cannot send error notification (RB) because VS_EMAIL_RECIPIENT_RB is not set.");
       return;
   }
   try {
       const subject = `ERROR: ${VS_COMPANY_NAME_RB} AI Recruiter Report Failed - ${new Date().toLocaleString()}`;
       let body = `Error generating/sending ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report:

${errorMessage}

`;
       if (stackTrace) {
           body += `Stack Trace:
${stackTrace}

`;
       }
       body += `Log Sheet URL: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`; // Use RB config
       MailApp.sendEmail(recipient, subject, body);
       Logger.log(`Error notification email (RB) sent to ${recipient}.`);
    } catch (emailError) {
       Logger.log(`CRITICAL: Failed to send error notification email (RB) to ${recipient}: ${emailError}`);
    }
}


// --- Utility / Setup Functions ---

/**
 * Creates menu items for the Recruiter Breakdown report.
 */
function setupRecruiterBreakdownMenu() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(`${VS_COMPANY_NAME_RB} AI Daily Summary`) // Menu Name Updated
      .addItem('Generate & Send Summary Now', 'AIR_DailySummarytoAP') // Updated Item Text & Function Name
      .addItem('Schedule Daily Summary (10 AM)', 'createRecruiterBreakdownTrigger') // Updated Item Text
      .addToUi();
  } catch (e) {
    Logger.log("Error creating Daily Summary menu (might happen if not opened from a Sheet): " + e); // Updated log
  }
}

// --- Helper Functions (Renamed with RB suffix for clarity, logic may be identical) ---
/**
 * Parses date strings safely.
 * @param {any} dateInput Input value.
 * @returns {Date|null} Parsed Date object or null.
 */
function vsParseDateSafeRB(dateInput) {
    if (dateInput === null || dateInput === undefined || dateInput === '') return null;
    if (typeof dateInput === 'number' && dateInput > 10000) {
       try {
           const jsTimestamp = (dateInput - 25569) * 86400 * 1000;
           const date = new Date(jsTimestamp);
            return !isNaN(date.getTime()) ? date : null;
       } catch (e) { /* Ignore */ }
    }
    const date = new Date(dateInput);
    return !isNaN(date.getTime()) ? date : null;
}

/**
 * Calculates time difference in days.
 * @param {Date|null} date1 Earlier date.
 * @param {Date|null} date2 Later date.
 * @returns {number|null} Difference in days or null.
 */
function vsCalculateDaysDifferenceRB(date1, date2) {
    if (!date1 || !date2) return null;
    const diffTime = date2.getTime() - date1.getTime();
    if (diffTime < 0) return null;
    return diffTime / (1000 * 60 * 60 * 24);
}

/**
 * Formats a Date object into DD-MMM-YY.
 * @param {Date|null} dateObject Date to format.
 * @returns {string} Formatted date string or 'N/A'.
 */
function vsFormatDateRB(dateObject) {
    if (!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) return 'N/A';
    const day = String(dateObject.getDate()).padStart(2, '0');
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[dateObject.getMonth()];
    const year = String(dateObject.getFullYear()).slice(-2);
    return `${day}-${month}-${year}`;
}

/**
 * Assigns a numerical rank to interview statuses for prioritization.
 * @param {string} status Raw interview status.
 * @returns {number} Rank.
 */
function vsGetStatusRankRB(status) {
    const COMPLETED_STATUSES_RAW = ['COMPLETED', 'Feedback Provided', 'Pending Feedback', 'No Show'];
    const SCHEDULED_STATUS_RAW = 'SCHEDULED';
    const PENDING_STATUSES_RAW = ['PENDING', 'INVITED', 'EMAIL SENT'];

    if (COMPLETED_STATUSES_RAW.includes(status)) return 1;
    if (status === SCHEDULED_STATUS_RAW) return 2;
    if (PENDING_STATUSES_RAW.includes(status)) return 3;
    return 99;
}