// AIR Volkscience - Exec Summary - Company-Level AI Interview Analytics Script v1.0 (Recruiter Breakdown)
// To: Akhila and Pavan
// When: Weekly, Monday at 10 AM (Can be adjusted)
// This script analyzes data from the Log_Enhanced sheet to provide company-wide insights
// including a breakdown by recruiter.

// --- Configuration ---
const VS_EMAIL_RECIPIENT_RB = 'akashyap@eightfold.ai'; // <<< UPDATE EMAIL RECIPIENT
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

    // 1b. Get Application Sheet Data (for Adoption Chart and AI Coverage)
    let adoptionChartData = null;
    let hiringMetrics = null;
    let aiCoverageMetrics = null;
    let validationSheetUrl = null;
    let recruiterValidationSheets = null;
    try {
        const appData = getApplicationDataForChartRB();
        if (appData && appData.rows) {
            Logger.log(`Successfully retrieved ${appData.rows.length} rows from application sheet.`);
            adoptionChartData = calculateAdoptionMetricsForChartRB(appData.rows, appData.colIndices);
            Logger.log(`Successfully calculated adoption chart metrics.`);
            
            // Calculate hiring metrics
            hiringMetrics = calculateHiringMetricsFromAppData(appData.rows, appData.colIndices);
            Logger.log(`Successfully calculated hiring metrics.`);
            
            // Calculate AI coverage metrics
            aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
            if (aiCoverageMetrics) {
                Logger.log(`Successfully calculated AI coverage metrics. Total eligible: ${aiCoverageMetrics.totalEligible}, Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}, Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
            } else {
                Logger.log(`WARNING: AI coverage metrics calculation returned null. This could be due to missing required columns.`);
            }
            
            // Create validation sheet for candidate count comparison
            try {
                validationSheetUrl = createCandidateCountValidationSheet(appData.rows, appData.colIndices);
                Logger.log(`Successfully created validation sheet: ${validationSheetUrl}`);
            } catch (validationError) {
                Logger.log(`Warning: Could not create validation sheet: ${validationError.toString()}`);
            }
            
            // Create detailed validation sheets for each recruiter
            try {
                recruiterValidationSheets = createAllRecruiterValidationSheets(appData.rows, appData.colIndices);
                if (recruiterValidationSheets) {
                    Logger.log(`Successfully created ${recruiterValidationSheets.successfulSheets} recruiter validation sheets`);
                } else {
                    Logger.log(`Warning: Could not create recruiter validation sheets`);
                }
            } catch (validationError) {
                Logger.log(`Warning: Could not create recruiter validation sheets: ${validationError.toString()}`);
            }
            
            // Logger.log(`Adoption Chart Data: ${JSON.stringify(adoptionChartData, null, 2)}`);
        } else {
            Logger.log(`WARNING: No data retrieved from application sheet. Adoption chart, hiring metrics, and AI coverage will be skipped.`);
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
    const htmlContent = createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, recruiterActivityData, recruiterNameIdx_Log, hiringMetrics, validationSheetUrl, aiCoverageMetrics, recruiterValidationSheets);
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
      'Days_pending_invitation', 'Interview Status_Real',
      'Position_approved_date'
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
      'Profile_id', 'Name', 'Last_stage', 'Ai_interview', 'Recruiter name', 'Application_status', 'Position_status', 'Application_ts', 'Position_id', 'Title'
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

  // Add Position_approved_date index if found
  const positionApprovedDateIndex = headers.indexOf('Position approved date');
  if (positionApprovedDateIndex !== -1) {
      appColIndices['Position approved date'] = positionApprovedDateIndex;
      Logger.log(`Found Position approved date column at index ${positionApprovedDateIndex}`);
  } else {
      Logger.log(`Optional column "Position approved date" not found. Candidate count comparison will be unavailable.`);
      // Try to find similar column names
      const possibleNames = ['Position approved date', 'Position_approved_date', 'Position Approved Date', 'Approved_date', 'Approved Date'];
      for (const name of possibleNames) {
          const index = headers.indexOf(name);
          if (index !== -1) {
              Logger.log(`Found similar column "${name}" at index ${index}. Please update the script to use this column name.`);
              break;
          }
      }
      
      // Search for any column containing "position" and "approved" or "date"
      const positionRelatedColumns = headers.filter(header => 
          header.toLowerCase().includes('position') && 
          (header.toLowerCase().includes('approved') || header.toLowerCase().includes('date'))
      );
      if (positionRelatedColumns.length > 0) {
          Logger.log(`Found position-related columns: ${positionRelatedColumns.join(', ')}`);
      }
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
    // <<< UPDATED: Only calculate for strictly COMPLETED interviews >>>
    if (statusRaw === 'COMPLETED') {
        const candidateName = (candidateNameIdx !== -1 && row[candidateNameIdx]) ? row[candidateNameIdx] : 'Unknown Candidate';
        const scheduleDateForAvg = (scheduledAtIdx !== -1) ? vsParseDateSafeRB(row[scheduledAtIdx]) : null; // Use RB helper
        if (sentDate && scheduleDateForAvg) {
            const daysDiff = vsCalculateDaysDifferenceRB(sentDate, scheduleDateForAvg);
            if (daysDiff !== null) {
                metrics.sentToScheduledDaysSum += daysDiff;
                metrics.sentToScheduledCount++;
                // <<< ADDED: Detailed log for avg time calculation >>>
                Logger.log(`AvgTimeCalc_Include: Candidate=[${candidateName}], Status=[${statusRaw}], Sent=[${sentDate.toISOString()}], Scheduled=[${scheduleDateForAvg.toISOString()}], DiffDays=[${daysDiff.toFixed(2)}]`);
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

    // <<< ADDED: Summary log before calculating average time >>>
    Logger.log(`AvgTimeCalc_Summary: Total Days Sum = ${metrics.sentToScheduledDaysSum.toFixed(2)}, Count = ${metrics.sentToScheduledCount}`);
    // <<< END ADDED >>>

    // <<< DEBUG LOGGING for KPI Rate >>>
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
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold; width: 180px;">${medal}${medal ? ' ' : ''}${rec}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">
                      ${data.recruiterSubmissionAwaited > 0 ? 
                          `<span style="color: red; font-weight: bold;">${data.recruiterSubmissionAwaited}</span>` : 
                          data.recruiterSubmissionAwaited
                      }
                    </td>
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
function createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, recruiterActivityData, recruiterNameIdx_Log, hiringMetrics, validationSheetUrl, aiCoverageMetrics, recruiterValidationSheets) {
  Logger.log(`DEBUG: AI Coverage Metrics in HTML report: ${aiCoverageMetrics ? 'Present' : 'Null/Undefined'}`);
  if (aiCoverageMetrics) {
    Logger.log(`DEBUG: AI Coverage Metrics details - Total eligible: ${aiCoverageMetrics.totalEligible}, Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}, Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
  }
  
  // --- Fetch AI Insights ---
  let aiInsights = "Insights generation is pending or encountered an issue."; // Default message
  try {
    // Check if the API key is likely configured by checking the script property directly
    if (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
        Logger.log("GEMINI_API_KEY found in script properties, attempting to fetch insights.");
        aiInsights = fetchInsightsFromGeminiAPI(metrics, adoptionChartData, recruiterActivityData);
        Logger.log(`Received AI Insights (first 100 chars): ${aiInsights.substring(0,100)}...`);
    } else {
        aiInsights = "AI Insights could not be generated: GEMINI_API_KEY not configured in Script Properties.";
        Logger.log(aiInsights);
    }
  } catch (e) {
      Logger.log(`Error during fetchInsightsFromGeminiAPI call or processing from createRecruiterBreakdownHtmlReport: ${e.toString()} Stack: ${e.stack ? e.stack : 'N/A'}`);
      aiInsights = "An error occurred while trying to generate AI insights. Please check the script logs for more details.";
  }
  // --- End Fetch AI Insights ---
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
                    <div style="text-align: center; font-size: 10px; color: #666; margin-top: 3px;">
                      (Excl. invites sent < 48 hours)
                    </div>
                    <div style="text-align: center; font-size: 10px; color: #666; margin-top: 1px;">
                      Overall Completion rate: ${metrics.completionRateOriginal}%
                    </div>
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
              </table>
            </td>
          </tr>

          <!-- AI Powered Insights - MOVED HERE -->
          <tr>
            <td style="padding-top: 15px; padding-bottom: 15px;">
              <div style="background-color: #fff8e1; padding: 20px; border: 1px solid #ffecb3; border-radius: 8px; margin-bottom: 15px;">
                <div style="font-weight: bold; font-size: 16px; color: #c77700; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #ffe082;">💡 AI-Powered Observations</div>
                <div style="font-size: 13px; color: #5f4300; white-space: pre-wrap; line-height: 1.6;">${formatInsightsForHtml(aiInsights, metrics)}</div>
                ${!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ? '<p style="font-size: 0.8em; color: #757575; margin-top: 10px;">(AI insights feature requires GEMINI_API_KEY to be configured in Script Properties)</p>' : ''}
              </div>
            </td>
          </tr>

          <!-- Hiring Success Metrics -->
          ${hiringMetrics ? `
          <tr>
            <td style="padding-top: 15px; padding-bottom: 15px;">
              <div style="background-color: #e8f5e9; padding: 20px; border: 1px solid #4caf50; border-radius: 8px; margin-bottom: 15px;">
                <div style="font-weight: bold; font-size: 16px; color: #2e7d32; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #4caf50;">🎯 Hiring Success After AI Interviews</div>
                <div style="display: flex; justify-content: space-around; text-align: center; margin: 20px 0;">
                  <div style="flex: 1; padding: 15px;">
                    <div style="font-size: 32px; font-weight: bold; color: #2e7d32;">${hiringMetrics.totalCandidates}</div>
                    <div style="font-size: 12px; color: #424242; margin-top: 5px;">Total Candidates<br/>Reached Offer Stage</div>
                  </div>
                  <div style="flex: 1; padding: 15px;">
                    <div style="font-size: 32px; font-weight: bold; color: #8e24aa;">${hiringMetrics.uniquePositionsFilled}</div>
                    <div style="font-size: 12px; color: #424242; margin-top: 5px;">Unique Positions<br/>at Offer</div>
                  </div>
                  <div style="flex: 1; padding: 15px;">
                    <div style="font-size: 32px; font-weight: bold; color: #1976d2;">
                      ${hiringMetrics.aiAvgMatchScore !== null && hiringMetrics.nonAiAvgMatchScore !== null ? 
                        `${hiringMetrics.aiAvgMatchScore} v/s ${hiringMetrics.nonAiAvgMatchScore}` : 
                        'N/A'
                      }
                    </div>
                    <div style="font-size: 12px; color: #424242; margin-top: 5px;">Match Score: AI v/s non-AI</div>
                  </div>
                  <div style="flex: 1; padding: 15px;">
                    <div style="font-size: 32px; font-weight: bold; color: #ff6f00;">
                      ${hiringMetrics.aiAvgCandidatesNeeded !== null && hiringMetrics.nonAiAvgCandidatesNeeded !== null ? 
                        `${hiringMetrics.aiAvgCandidatesNeeded} v/s ${hiringMetrics.nonAiAvgCandidatesNeeded}` : 
                        'N/A'
                      }
                    </div>
                    <div style="font-size: 12px; color: #424242; margin-top: 5px;">Avg # of candidates<br/>needed to reach offer<br/>(AI v/s non-AI)</div>
                  </div>
                </div>
                ${hiringMetrics.stageBreakdown && Object.keys(hiringMetrics.stageBreakdown).length > 0 ? `
                <div style="margin-top: 15px;">
                  <div style="font-weight: bold; font-size: 14px; color: #424242; margin-bottom: 10px;">Breakdown by Stage:</div>
                  <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; border: 1px solid #4caf50; border-radius: 4px; overflow: hidden;">
                    <thead>
                      <tr>
                        <th style="border: 1px solid #4caf50; padding: 8px 12px; text-align: left; font-size: 12px; background-color: #4caf50; color: white; font-weight: bold;">Hiring Stage</th>
                        <th style="border: 1px solid #4caf50; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #4caf50; color: white; font-weight: bold;">Count</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${Object.entries(hiringMetrics.stageBreakdown)
                        .sort(([, a], [, b]) => b - a) // Sort by count descending
                        .map(([stage, count], index) => `
                        <tr style="background-color: ${index % 2 === 0 ? '#f1f8e9' : '#ffffff'};">
                          <td style="border: 1px solid #4caf50; padding: 8px 12px; text-align: left; font-size: 12px;">${stage}</td>
                          <td style="border: 1px solid #4caf50; padding: 8px 12px; text-align: center; font-size: 12px; font-weight: bold;">${count}</td>
                        </tr>`).join('')}
                    </tbody>
                  </table>
                </div>
                ` : ''}
                <p style="font-size: 11px; color: #424242; margin-top: 15px; text-align: center;">
                  <strong>Filters:</strong> Last_stage = Offer Approvals, etc. | Recruiter != Samrudh/Simran
                  <br/><em>Avg. candidates calculation uses only AI candidates for AI-assisted reqs & requires >= 3 candidates.</em>
                  ${!hiringMetrics.hasPositionIdColumn ? '<br/><span style="color: #f57c00;">Note: Position_id column not found.</span>' : ''}
                  ${!hiringMetrics.hasMatchStarsColumn ? '<br/><span style="color: #f57c00;">Note: Match_stars column not found.</span>' : ''}
                  ${validationSheetUrl ? `<br/><a href="${validationSheetUrl}" style="color: #1976d2; text-decoration: underline;">📊 View Detailed Position-Level Validation Data</a>` : ''}
                </p>
              </div>
            </td>
          </tr>
          ` : ''}

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
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase; width: 180px;">Recruiter Name</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Sent</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Completed (# / %)</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Scheduled</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Pending (# / %)</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Feedback Submitted</th>
                           <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Submission Awaited</th>
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
                                     <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold; width: 180px;">${medal}${medal ? ' ' : ''}${rec}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: red; font-weight: bold;">${data.recruiterSubmissionAwaited}</span>` : 
                                            data.recruiterSubmissionAwaited
                                        }
                                      </td>
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

          <!-- AI Interview Coverage Bar Chart -->
          ${aiCoverageMetrics ? `
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              ${generateAICoverageBarChartHtml(aiCoverageMetrics)}
            </td>
          </tr>
          ` : ''}

          <!-- Detailed Validation Sheets -->
          ${recruiterValidationSheets ? `
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              ${generateValidationSheetsHtml(recruiterValidationSheets)}
            </td>
          </tr>
          ` : ''}

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
                            let daysAgoText = '';
                            if (activity.daysAgo === -1) {
                                daysAgoText = 'Today';
                            } else if (activity.daysAgo === 0) {
                                daysAgoText = 'Yesterday';
                            } else if (activity.daysAgo >= 1) {
                                const actualDays = activity.daysAgo + 1; // Adjust because floor(today-yesterday) is 0
                                daysAgoText = `${actualDays} calendar ${actualDays === 1 ? 'day' : 'days'} ago`;
                            } else {
                                daysAgoText = 'Unknown'; // Should not happen
                            }

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
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: red; font-weight: bold;">${data.recruiterSubmissionAwaited}</span>` : 
                                            data.recruiterSubmissionAwaited
                                        }
                                      </td>
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
                            <th style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 11px; vertical-align: middle; background-color: #f5f5f5; font-weight: bold; color: #424242; text-transform: uppercase;">Recruiter Submission Awaited</th>
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
                                      <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: red; font-weight: bold;">${data.recruiterSubmissionAwaited}</span>` : 
                                            data.recruiterSubmissionAwaited
                                        }
                                      </td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>
             </td>
          </tr>

          <!-- New row for the footnote under KPIs -->
          <tr>
            <td colspan="4" style="text-align: center; padding-top: 5px; font-size: 10px; color: #666;">
              *Avg Time Sent to Completion calculation currently uses Schedule Start Date as completion proxy.<br>
              **Completion Rate KPI** excludes invitations sent within the last 48 hours. Breakdown table % includes all sent invites.<br>
              Overall completion rate (including all invites): ${metrics.completionRateOriginal}%.<br>
              Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.
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

//====================================================================================================
// --- Insight Formatting Helper (Revised) ---
//====================================================================================================
/**
 * Formats AI-generated insights text for HTML display by bolding known entities and percentages.
 * @param {string} insightsText The raw text insights from the LLM.
 * @param {object} metrics The main metrics object to extract known entity names.
 * @return {string} HTML formatted insights string.
 */
function formatInsightsForHtml(insightsText, metrics) {
  if (!insightsText || typeof insightsText !== 'string') {
    return 'No insights available or an error occurred.';
  }

  let formattedText = insightsText;

  // 1. Collect known entities (recruiters, job functions, countries)
  let knownEntities = [];
  if (metrics) {
    if (metrics.byRecruiter) {
      knownEntities.push(...Object.keys(metrics.byRecruiter).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    if (metrics.byJobFunction) {
      knownEntities.push(...Object.keys(metrics.byJobFunction).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    if (metrics.byCountry) {
      knownEntities.push(...Object.keys(metrics.byCountry).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    // Add any other specific entities you want to bold if they are reliably named.
    // Example: If you have a list of project names or specific terms stored elsewhere in metrics.
  }
  
  // Remove duplicates and filter out very short strings (e.g., single characters unless specifically desired)
  knownEntities = [...new Set(knownEntities)].filter(e => e.length > 2); // Min length 3 for an entity name
  // Sort by length descending. Crucial for correct replacement of substrings.
  knownEntities.sort((a, b) => b.length - a.length);

  // 2. Bold known entities (done first)
  knownEntities.forEach(entity => {
    // Escape special regex characters in the entity name itself
    const escapedEntity = entity.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try {
      // Case-insensitive ('i') and global ('g') match, with word boundaries (\b)
      const regex = new RegExp(`\\b(${escapedEntity})\\b`, 'gi'); 
      formattedText = formattedText.replace(regex, '<strong>$1</strong>');
    } catch (e) {
      Logger.log(`Error creating or using regex for entity "${entity}": ${e.toString()}`);
      // Continue without bolding this specific entity if regex fails
    }
  });

  // 3. Bold percentages (e.g., 50.7%, 79.4%, 20%) - done after entities
  // This regex looks for digits, optionally with a decimal, followed by a % sign,
  // ensuring it's a whole word/number using word boundaries.
  formattedText = formattedText.replace(/(\b\d+(\.\d+)?%\b)/g, '<strong>$1</strong>');

  // 4. Replace newlines with <br> for HTML display (done last)
  formattedText = formattedText.replace(/\n/g, '<br>');

  return formattedText;
}

//====================================================================================================
// --- Gemini API Integration for Insights ---
//====================================================================================================

/**
 * Fetches insights from the Google Gemini API based on summarized report data.
 * @param {object} metrics The calculated company metrics.
 * @param {object} adoptionChartData The adoption chart data.
 * @param {Array<object>} recruiterActivityData Recruiter activity data.
 * @return {string} Textual insights from the LLM, or an error/status message.
 */
function fetchInsightsFromGeminiAPI(metrics, adoptionChartData, recruiterActivityData) {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!GEMINI_API_KEY) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties. AI Insights will not be generated.");
    return "AI Insights could not be generated: API Key not configured in Script Properties.";
  }
  // Using gemini-1.5-flash-latest as an example for a fast and capable model.
  const GEMINI_API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${GEMINI_API_KEY}`;

  try {
    // 1. Summarize your data
    let dataSummary = `AI Interview Report Summary (Data for ${VS_COMPANY_NAME_RB}):\n\n`; // VS_COMPANY_NAME_RB should be globally accessible
    dataSummary += `Overall Performance (${metrics.reportStartDate} to ${metrics.reportEndDate}):\n`;
    dataSummary += `- Total AI Invitations Sent: ${metrics.totalSent}\n`;
    dataSummary += `- Overall Completion Rate (KPI Adjusted, for invites older than 48 hours): ${metrics.kpiCompletionRateAdjusted}%\n`;
    dataSummary += `- Overall Completion Rate (All Time): ${metrics.completionRateOriginal}%\n`;
    dataSummary += `- Average Time Sent to Completion (Proxy: Schedule Start): ${metrics.avgTimeToScheduleDays !== null ? metrics.avgTimeToScheduleDays + ' days' : 'N/A'}\n`;
    dataSummary += `- Average Match Stars (for Completed Interviews): ${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}\n`;

    // Revised Recruiter Performance Summary to consider all recruiters for insights
    if (metrics.byRecruiter && Object.keys(metrics.byRecruiter).length > 0) {
      const recruiterStats = Object.entries(metrics.byRecruiter)
        .filter(([name]) => name !== 'Unknown') // Exclude 'Unknown' for specific high/low stats
        .map(([name, data]) => ({
          name: name,
          sent: parseInt(data.sent) || 0,
          completed: parseInt(data.completedNumber) || 0,
          completionRate: parseFloat(data.completedPercentOfSent) || 0,
          pending: parseInt(data.pendingNumber) || 0
        }));

      if (recruiterStats.length > 0) {
        dataSummary += `\nRecruiter Performance Analysis (Based on ${recruiterStats.length} recruiters with known names):
`;

        // Sort by completion rate for high/low
        const sortedByCompletionRate = [...recruiterStats].sort((a, b) => b.completionRate - a.completionRate);
        dataSummary += `- Highest Completion Rate: ${sortedByCompletionRate[0].name} (${sortedByCompletionRate[0].completionRate}% from ${sortedByCompletionRate[0].sent} sent invites).
`;

        // For lowest, find someone with a minimum number of sent invites to make it meaningful
        const minSentForLowConsideration = 5; // Adjustable threshold
        const eligibleForLowest = sortedByCompletionRate.filter(r => r.sent >= minSentForLowConsideration);
        if (eligibleForLowest.length > 0) {
          const lowestPerformer = eligibleForLowest[eligibleForLowest.length - 1];
          if (lowestPerformer.name !== sortedByCompletionRate[0].name) { // Avoid repeating if only one eligible
            dataSummary += `- Lowest Completion Rate (among those with >=${minSentForLowConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent invites).
`;
          }
        }

        // Sort by sent for most active
        const sortedBySent = [...recruiterStats].sort((a, b) => b.sent - a.sent);
        dataSummary += `- Most Invites Sent: ${sortedBySent[0].name} (${sortedBySent[0].sent} invites, ${sortedBySent[0].completionRate}% completion rate).
`;

        // Calculate overall average completion rate for this group
        let totalSentByKnownRecruiters = 0;
        let totalCompletedByKnownRecruiters = 0;
        recruiterStats.forEach(r => {
          totalSentByKnownRecruiters += r.sent;
          totalCompletedByKnownRecruiters += r.completed;
        });
        const avgCompletionRateKnown = totalSentByKnownRecruiters > 0 ? parseFloat(((totalCompletedByKnownRecruiters / totalSentByKnownRecruiters) * 100).toFixed(1)) : 0;
        dataSummary += `- Average Completion Rate (for these ${recruiterStats.length} recruiters): ${avgCompletionRateKnown}%\n`;
      }

      if (metrics.byRecruiter['Unknown'] && metrics.byRecruiter['Unknown'].sent > 0) {
        dataSummary += `- Invites Sent by 'Unknown' Recruiters: ${metrics.byRecruiter['Unknown'].sent} (completion rate: ${metrics.byRecruiter['Unknown'].completedPercentOfSent}%).\n`;
      }
    }

    // Enhanced Recruiter Last AI Invite Activity Summary
    if (recruiterActivityData && recruiterActivityData.length > 0) {
        dataSummary += `\nRecruiter Last AI Invite Activity:\n`;
        const recentActivityDisplayCount = 3; // How many most recent to display

        recruiterActivityData.slice(0, recentActivityDisplayCount).forEach(activity => {
            let daysAgoText = "";
            if (activity.daysAgo === 0) daysAgoText = "Today";
            else if (activity.daysAgo === 1) daysAgoText = "Yesterday";
            else daysAgoText = `${activity.daysAgo} days ago`;
            dataSummary += `- Recently Active: ${activity.recruiter} (Last AI invite: ${daysAgoText}, 10-day trend: ${activity.dailyTrend}).\n`;
        });

        const inactivityThresholdDays = 7; // Recruiters with no AI invites for this many days or more
        // Note: activity.daysAgo = 0 is today, 1 is yesterday. So >= 7 means 7 full days have passed (i.e., last activity was 7+ days ago)
        const lessActiveRecruiters = recruiterActivityData.filter(activity => activity.daysAgo >= inactivityThresholdDays);

        if (lessActiveRecruiters.length > 0) {
            dataSummary += `\n- Recruiters with Notably Low Recent AI Invite Activity (Last AI invite >= ${inactivityThresholdDays} days ago):\n`;
            lessActiveRecruiters.forEach(activity => {
                 let daysAgoText = `${activity.daysAgo} days ago`; // Simplified for this section
                 dataSummary += `  - ${activity.recruiter} (Last AI invite: ${daysAgoText}).\n`;
            });
        }
    }



    if (recruiterActivityData && recruiterActivityData.length > 0) {
        dataSummary += `\nRecruiter Last Invite Activity (Most Recent First):\n`;
        recruiterActivityData.slice(0,3).forEach(activity => {
            let daysAgoText = '';
            if (activity.daysAgo === -1) daysAgoText = 'Today';
            else if (activity.daysAgo === 0) daysAgoText = 'Yesterday';
            else if (activity.daysAgo >= 1) daysAgoText = `${activity.daysAgo + 1} calendar days ago`;
            else daysAgoText = 'Unknown';
            dataSummary += `- ${activity.recruiter}: Last invite sent ${daysAgoText}. Trend (last 10 workdays, excl. weekends): ${activity.dailyTrend}\n`;
        });
    }

    // Revised Performance by Job Function Summary
    if (metrics.byJobFunction && Object.keys(metrics.byJobFunction).length > 0) {
        const jobFunctionStats = Object.entries(metrics.byJobFunction)
            .map(([funcName, data]) => ({
                name: funcName,
                sent: parseInt(data.sent) || 0,
                completed: parseInt(data.completedNumber) || 0,
                completionRate: parseFloat(data.completedPercentOfSent) || 0,
                // recruiterSubmissionAwaited: parseInt(data.recruiterSubmissionAwaited) || 0 // If needed later
            }));

        if (jobFunctionStats.length > 0) {
            dataSummary += `\nPerformance by Job Function (Analysis based on ${jobFunctionStats.length} functions):
`;
            const minSentForConsideration = 5; // Job functions with at least this many sent invites considered for low/high
            const relevantJobFunctions = jobFunctionStats.filter(jf => jf.sent >= minSentForConsideration);

            if (relevantJobFunctions.length > 0) {
                // Sort by completion rate
                const sortedByCompletion = [...relevantJobFunctions].sort((a, b) => b.completionRate - a.completionRate);

                dataSummary += `- Highest Completion Rate (among functions with >=${minSentForConsideration} invites): ${sortedByCompletion[0].name} (${sortedByCompletion[0].completionRate}% from ${sortedByCompletion[0].sent} sent).
`;

                if (sortedByCompletion.length > 1) {
                    const lowestPerformer = sortedByCompletion[sortedByCompletion.length - 1];
                     if (lowestPerformer.name !== sortedByCompletion[0].name) { // Avoid repeating if only one eligible
                        dataSummary += `- Lowest Completion Rate (among functions with >=${minSentForConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent).
`;
                    }
                }
                 // Add one or two other notable ones if they exist and are different, e.g., highest volume
                const sortedBySent = [...jobFunctionStats].sort((a,b) => b.sent - a.sent);
                if (sortedBySent.length > 0 && sortedBySent[0].name !== sortedByCompletion[0].name && (sortedByCompletion.length <=1 || sortedBySent[0].name !== sortedByCompletion[sortedByCompletion.length-1].name)) {
                     dataSummary += `- Highest Volume of Invites: ${sortedBySent[0].name} (${sortedBySent[0].sent} sent, ${sortedBySent[0].completionRate}% completion rate).\n`;
                }

                // Specifically add Engineering stats if it exists and meets threshold, or just add it if it exists
                const engineeringStats = jobFunctionStats.find(jf => jf.name.toLowerCase() === 'engineering');
                if (engineeringStats) {
                    dataSummary += `- Specific Stats for Engineering: ${engineeringStats.sent} invites sent, ${engineeringStats.completionRate}% completion rate.\n`;
                }

            } else {
                dataSummary += `- Insufficient data for detailed job function comparison (few functions with >=${minSentForConsideration} invites sent).
`;
            }
        }
    }

    // Performance by Country Summary
    if (metrics.byCountry && Object.keys(metrics.byCountry).length > 0) {
        const countryStats = Object.entries(metrics.byCountry)
            .map(([countryName, data]) => ({
                name: countryName,
                sent: parseInt(data.sent) || 0,
                completed: parseInt(data.completedNumber) || 0,
                completionRate: parseFloat(data.completedPercentOfSent) || 0,
            }));

        if (countryStats.length > 0) {
            dataSummary += `\nPerformance by Country (Analysis based on ${countryStats.length} countries):
`;
            const minSentForCountryConsideration = 5; // Countries with at least this many sent invites
            const relevantCountries = countryStats.filter(c => c.sent >= minSentForCountryConsideration);

            if (relevantCountries.length > 0) {
                const sortedByCompletion = [...relevantCountries].sort((a, b) => b.completionRate - a.completionRate);
                dataSummary += `- Country with Highest Completion Rate (>=${minSentForCountryConsideration} invites): ${sortedByCompletion[0].name} (${sortedByCompletion[0].completionRate}% from ${sortedByCompletion[0].sent} sent).
`;
                if (sortedByCompletion.length > 1) {
                    const lowestPerformer = sortedByCompletion[sortedByCompletion.length - 1];
                    if (lowestPerformer.name !== sortedByCompletion[0].name) {
                        dataSummary += `- Country with Lowest Completion Rate (>=${minSentForCountryConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent).
`;
                    }
                }
                const sortedBySent = [...countryStats].sort((a,b) => b.sent - a.sent);
                if (sortedBySent.length > 0 && sortedBySent[0].name !== sortedByCompletion[0].name && (sortedByCompletion.length <= 1 || sortedBySent[0].name !== sortedByCompletion[sortedByCompletion.length-1].name)) {
                    dataSummary += `- Country with Highest Volume of Invites: ${sortedBySent[0].name} (${sortedBySent[0].sent} sent, ${sortedBySent[0].completionRate}% completion rate).
`;
                }
            } else {
                dataSummary += `- Insufficient data for detailed country comparison (few countries with >=${minSentForCountryConsideration} invites sent).
`;
            }
        }
    }
    
    dataSummary += "\nConsiderations: 'Completion Rate (KPI Adjusted)' excludes invites sent in the last 48 hours. 'Avg Time Sent to Completion' uses schedule start time as a proxy for completion. Adoption metrics are based on a specific cohort post-launch with a match score filter.\n";

    // 2. Construct the Prompt
    const promptText = `You are an expert data analyst for an HR department. Your task is to provide insightful commentary on a report summarizing AI interview adoption and recruiter activity for ${VS_COMPANY_NAME_RB}.
Based *only* on the following data summary, please generate 3 to 5 bullet-point key observations.
For each observation, identify the primary subject (e.g., overall performance, a specific job function, recruiter activity, country performance, AI adoption).
When discussing a specific category from the summary (like Recruiter Performance, Job Functions, Countries, or Adoption):
- Attempt to highlight the most significant variations by contrasting the best and worst performers (e.g., highest vs. lowest completion rates, most vs. least active) using the specific names and figures provided in the summary for that category.
- Call out specific names (recruiters, job functions, countries) when discussing these variations if the data supports it.
Focus on actionable insights, positive trends, areas that might need attention, or interesting patterns. 
Do not refer to data not present in the summary. Be concise, ensuring each bullet point is a distinct observation.
If data for the "Engineering" job function is present in the summary, please include one bullet point observation specifically about Engineering performance.

Data Summary:
---
${dataSummary}
---

Key Observations (3-5 bullet points):
`;

    const payload = {
      contents: [{ role: "user", parts: [{text: promptText}] }],
      generationConfig: {
        "temperature": 0.6,
        "maxOutputTokens": 500,
        "topP": 0.9,
        "topK": 40
      },
       safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_MEDIUM_AND_ABOVE" }
      ]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    Logger.log(`Sending data to Gemini API. Endpoint: ${GEMINI_API_ENDPOINT}. Summary length: ${dataSummary.length} chars.`);
    const response = UrlFetchApp.fetch(GEMINI_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
          jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
          jsonResponse.candidates[0].content.parts.length > 0 && jsonResponse.candidates[0].content.parts[0].text) {
        
        let insights = jsonResponse.candidates[0].content.parts[0].text;
        Logger.log("Successfully received insights from Gemini API.");
        insights = insights.trim();
        // Attempt to format into bullet points if the model didn't already.
        if (!insights.startsWith("* ") && !insights.startsWith("- ") && !insights.startsWith("• ")) {
            insights = insights.split('\n').map(line => line.trim()).filter(line => line.length > 0).map(line => `• ${line}`).join('\n');
        }
        return insights;
      } else if (jsonResponse.candidates && jsonResponse.candidates[0].finishReason) {
         Logger.log(`Gemini API call finished with reason: ${jsonResponse.candidates[0].finishReason}.`);
         let detail = jsonResponse.candidates[0].finishReason;
         if(jsonResponse.candidates[0].safetyRatings) {
           detail += ` Safety Ratings: ${JSON.stringify(jsonResponse.candidates[0].safetyRatings)}`;
         }
         return `AI Insights generation issue: Model finished with reason '${detail}'. Content might be blocked or prompt needs adjustment.`;
      } else {
        Logger.log(`Gemini API response does not contain expected text data. Response: ${responseBody}`);
        return "AI Insights generation failed: Unexpected API response structure. Check logs.";
      }
    } else {
      Logger.log(`Error calling Gemini API: ${responseCode} - ${responseBody}`);
      return `Could not generate AI insights: API Error ${responseCode}. Details: ${responseBody.substring(0, 500)}. Check logs for full error.`;
    }
  } catch (error) {
    Logger.log(`Critical Error in fetchInsightsFromGeminiAPI: ${error.toString()} \nStack: ${error.stack}`);
    return `Could not generate AI insights due to an internal script error: ${error.message}. Check logs.`;
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
      let body = `Error generating/sending ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report:\n\n${errorMessage}\n\n`;
      if (stackTrace) {
          body += `Stack Trace:\n${stackTrace}\n\n`;
      }
      body += `Log Sheet URL: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`; // Use RB config
      MailApp.sendEmail(recipient, subject, body);
      Logger.log(`Error notification email (RB) sent to ${recipient}.`);
   } catch (emailError) {
      Logger.log(`CRITICAL: Failed to send error notification email (RB) to ${recipient}: ${emailError}`);
   }
}

/**
 * Calculates hiring metrics from application data for candidates who took AI interviews.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} An object containing hiring metrics.
 */
function calculateHiringMetricsFromAppData(appRows, appColIndices) {
  Logger.log(`--- Starting calculateHiringMetricsFromAppData ---`);

  const hiringStages = [
    'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
  ];

  // ---- TOP LEVEL METRICS CALCULATION ----
  const topLevelHiringCandidates = appRows.filter(row => {
    const lastStage = row[appColIndices['Last_stage']];
    return lastStage && hiringStages.includes(lastStage);
  });
  Logger.log(`Found ${topLevelHiringCandidates.length} total candidates who reached hiring stages`);

  const recruiterNameIndex = appColIndices.hasOwnProperty('Recruiter name') ? appColIndices['Recruiter name'] : -1;
  const filteredHiringCandidates = topLevelHiringCandidates.filter(row => {
    if (recruiterNameIndex === -1) return true;
    const recruiterName = row[recruiterNameIndex] || '';
    return !recruiterName.toLowerCase().includes('samrudh') && !recruiterName.toLowerCase().includes('simran');
  });
  Logger.log(`After excluding Samrudh/Simran recruiters: ${filteredHiringCandidates.length} candidates at offer stage`);

  const aiCandidates = filteredHiringCandidates.filter(row => (row[appColIndices['Ai_interview']] || '') === 'Y');
  const nonAiCandidates = filteredHiringCandidates.filter(row => (row[appColIndices['Ai_interview']] || '') !== 'Y');
  Logger.log(`AI candidates at offer stage: ${aiCandidates.length}, Non-AI candidates: ${nonAiCandidates.length}`);

  const positionIdIndex = appColIndices.hasOwnProperty('Position_id') ? appColIndices['Position_id'] : -1;
  const uniquePositions = new Set(aiCandidates.map(row => row[positionIdIndex]).filter(id => id));

  const matchStarsIndex = appColIndices.hasOwnProperty('Match_stars') ? appColIndices['Match_stars'] : -1;
  let aiAvgMatchScore = null, nonAiAvgMatchScore = null;
  if (matchStarsIndex !== -1) {
    const getAvgScore = (candidates) => {
      const scores = candidates.map(row => parseFloat(row[matchStarsIndex])).filter(score => !isNaN(score) && score >= 0);
      return scores.length > 0 ? parseFloat((scores.reduce((sum, score) => sum + score, 0) / scores.length).toFixed(1)) : null;
    };
    aiAvgMatchScore = getAvgScore(aiCandidates);
    nonAiAvgMatchScore = getAvgScore(nonAiCandidates);
  }

  // ---- NUANCED CALCULATION for "Average Candidates Needed" ----
  let aiAvgCandidatesNeeded = null, nonAiAvgCandidatesNeeded = null;
  try {
    const candidatesForRatio = appRows.filter(row => {
      if (recruiterNameIndex === -1) return true;
      const recruiterName = row[recruiterNameIndex] || '';
      return !recruiterName.toLowerCase().includes('samrudh') && !recruiterName.toLowerCase().includes('simran');
    });

    const positionStats = {};
    const progressedStages = [
      'Hiring Manager Screen', 'Assessment', 'Onsite Interview', 'Final Interview', 'Candidate Withdrew', 'Candidate Hold',
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    candidatesForRatio.forEach(row => {
      const posId = row[positionIdIndex];
      if (!posId) return;
      if (!positionStats[posId]) {
        positionStats[posId] = { ai_progressed: 0, ai_offered: 0, non_ai_progressed: 0, non_ai_offered: 0, had_ai_interview: false };
      }
      const stats = positionStats[posId];
      const lastStage = row[appColIndices['Last_stage']] || '';
      const aiInterview = row[appColIndices['Ai_interview']] || '';
      
      const hasProgressed = progressedStages.includes(lastStage);
      
      if (aiInterview === 'Y') {
        stats.had_ai_interview = true;
        if (hasProgressed) {
          stats.ai_progressed++;
          if (hiringStages.includes(lastStage)) stats.ai_offered++;
        }
      } else {
        if (hasProgressed) {
          stats.non_ai_progressed++;
          if (hiringStages.includes(lastStage)) stats.non_ai_offered++;
        }
      }
    });

    const aiPositionRatios = [], nonAiPositionRatios = [];
    for (const posId in positionStats) {
      const stats = positionStats[posId];
      if (stats.had_ai_interview) {
        // For AI positions, only require that an offer was made.
        if (stats.ai_offered > 0) {
          aiPositionRatios.push(stats.ai_progressed / stats.ai_offered);
        }
      } else {
        // For Non-AI positions, keep the significance threshold.
        if (stats.non_ai_progressed >= 3 && stats.non_ai_offered > 0) {
          nonAiPositionRatios.push(stats.non_ai_progressed / stats.non_ai_offered);
        }
      }
    }

    if (aiPositionRatios.length > 0) aiAvgCandidatesNeeded = parseFloat((aiPositionRatios.reduce((a, b) => a + b, 0) / aiPositionRatios.length).toFixed(1));
    if (nonAiPositionRatios.length > 0) nonAiAvgCandidatesNeeded = parseFloat((nonAiPositionRatios.reduce((a, b) => a + b, 0) / nonAiPositionRatios.length).toFixed(1));
    Logger.log(`Calculated avg candidates needed. AI Ratios Count: ${aiPositionRatios.length}, Non-AI Ratios Count: ${nonAiPositionRatios.length}`);
  } catch (e) {
    Logger.log(`ERROR during nuanced average candidate calculation: ${e}`);
  }
  
  // ---- Final Metrics Object ----
  const positionApprovedIndex = appColIndices.hasOwnProperty('Position approved date') ? appColIndices['Position approved date'] : -1;
  const metrics = {
    totalCandidates: aiCandidates.length,
    uniquePositionsFilled: uniquePositions.size,
    stageBreakdown: hiringStages.reduce((acc, stage) => { acc[stage] = aiCandidates.filter(row => row[appColIndices['Last_stage']] === stage).length; return acc; }, {}),
    hasPositionIdColumn: positionIdIndex !== -1,
    aiAvgMatchScore, nonAiAvgMatchScore,
    hasMatchStarsColumn: matchStarsIndex !== -1,
    aiCandidatesCount: aiCandidates.length,
    nonAiCandidatesCount: nonAiCandidates.length,
    aiAvgCandidatesNeeded, nonAiAvgCandidatesNeeded,
    hasPositionApprovedDateColumn: positionApprovedIndex !== -1
  };

  Logger.log(`Hiring Metrics: ${metrics.totalCandidates} AI candidates reached offer stage, ${metrics.uniquePositionsFilled} unique positions filled`);
  Logger.log(`Match Score Comparison: AI avg=${aiAvgMatchScore}, Non-AI avg=${nonAiAvgMatchScore}`);
  return metrics;
}

/**
 * Creates a validation Google Sheet with detailed position-level data for candidate count comparison.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {string} The URL of the created Google Sheet.
 */
function createCandidateCountValidationSheet(appRows, appColIndices) {
  Logger.log(`--- Starting createCandidateCountValidationSheet ---`);
  
  try {
    const spreadsheet = SpreadsheetApp.create(`AI Interview Candidate Count Validation - ${new Date().toISOString().split('T')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    const headers = [
      'Position ID', 'Position Title', 'Recruiter Name', 'Position Approved Date', 'AI Interview Used',
      'Total Candidates Progressed', 'AI Candidates Progressed', 'Non-AI Candidates Progressed',
      'Candidates Reached Offer', 'AI Candidates Reached Offer', 'Non-AI Candidates Reached Offer',
      'AI Cands-to-Offer Ratio', 'Non-AI Cands-to-Offer Ratio',
      'Hired Candidate Name', 'Included in Calc (>3 Progressed)'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    const progressedStages = [
      'Hiring Manager Screen', 'Assessment', 'Onsite Interview', 'Final Interview', 'Candidate Withdrew', 'Candidate Hold',
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    const hiringStages = [
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    const { 
      Position_id: positionIdIndex, Title: positionTitleIndex, 'Recruiter name': recruiterNameIndex,
      'Position approved date': positionApprovedIndex, Ai_interview: aiInterviewIndex,
      Last_stage: lastStageIndex, Name: nameIndex 
    } = appColIndices;

    if (positionIdIndex === undefined || positionApprovedIndex === undefined) {
      throw new Error('Required columns Position_id or Position approved date not found');
    }
    
    const currentYear = new Date().getFullYear();
    const thisYearPositions = appRows.filter(row => {
      const approvedDate = row[positionApprovedIndex];
      if (!approvedDate) return false;
      const date = vsParseDateSafeRB(approvedDate);
      return date && date.getFullYear() === currentYear;
    });
    
    const filteredPositions = thisYearPositions.filter(row => {
      if (recruiterNameIndex === undefined) return true;
      const recruiterName = (row[recruiterNameIndex] || '').toLowerCase();
      return !recruiterName.includes('samrudh') && !recruiterName.includes('simran');
    });
    
    const positionData = {};
    
    filteredPositions.forEach(row => {
      const positionId = row[positionIdIndex];
      if (!positionData[positionId]) {
        positionData[positionId] = {
          positionId, positionTitle: row[positionTitleIndex] || 'N/A',
          recruiterName: row[recruiterNameIndex] || 'N/A', positionApprovedDate: row[positionApprovedIndex],
          totalProgressed: 0, aiProgressed: 0, nonAiProgressed: 0,
          totalReachedOffer: 0, aiReachedOffer: 0, nonAiReachedOffer: 0,
          hasAiCandidates: false, hiredCandidates: []
        };
      }
      
      const stats = positionData[positionId];
      const lastStage = row[lastStageIndex];
      const aiInterview = row[aiInterviewIndex];
      
      if (progressedStages.includes(lastStage)) {
        stats.totalProgressed++;
        const wasOffered = hiringStages.includes(lastStage);

        if (wasOffered) {
          stats.totalReachedOffer++;
        }

        if (aiInterview === 'Y') {
          stats.hasAiCandidates = true;
          stats.aiProgressed++;
          if (wasOffered) stats.aiReachedOffer++;
        } else {
          stats.nonAiProgressed++;
          if (wasOffered) stats.nonAiReachedOffer++;
        }
        
        if (wasOffered && lastStage === 'Hired' && nameIndex !== undefined) {
          stats.hiredCandidates.push(row[nameIndex] || '');
        }
      }
    });
    
    const positionRows = Object.values(positionData)
      .filter(pos => pos.totalProgressed > 0)
      .map(pos => {
        const isAiPosition = pos.hasAiCandidates;
        // Asymmetrical inclusion logic based on user feedback
        const inclusionThresholdMet = isAiPosition 
          ? pos.aiReachedOffer > 0 
          : (pos.nonAiProgressed >= 3 && pos.nonAiReachedOffer > 0);
        return [
          pos.positionId, pos.positionTitle, pos.recruiterName, pos.positionApprovedDate,
          isAiPosition ? 'Yes' : 'No',
          pos.totalProgressed, pos.aiProgressed, pos.nonAiProgressed,
          pos.totalReachedOffer, pos.aiReachedOffer, pos.nonAiReachedOffer,
          pos.aiReachedOffer > 0 ? (pos.aiProgressed / pos.aiReachedOffer).toFixed(1) : 'N/A',
          pos.nonAiReachedOffer > 0 ? (pos.nonAiProgressed / pos.nonAiReachedOffer).toFixed(1) : 'N/A',
          pos.hiredCandidates.join(', '),
          inclusionThresholdMet ? 'Yes' : 'No'
        ];
      });
    
    positionRows.sort((a, b) => b[5] - a[5]);
    
    if (positionRows.length > 0) {
      sheet.getRange(2, 1, positionRows.length, headers.length).setValues(positionRows);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`Validation sheet created with ${positionRows.length} positions.`);
    
    return spreadsheet.getUrl();
    
  } catch (error) {
    Logger.log(`Error creating validation sheet: ${error.toString()} Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Calculates AI interview coverage metrics by recruiter.
 * Shows how many eligible candidates (not in "New" or "Added" stages) should have had AI interviews.
 * @param {Array<Array>} appRows Rows from the application sheet.
 * @param {object} appColIndices Column indices for the application sheet.
 * @returns {object} Object containing coverage metrics by recruiter.
 */
function calculateAICoverageMetricsRB(appRows, appColIndices) {
  Logger.log(`--- Starting calculateAICoverageMetricsRB ---`);
  Logger.log(`DEBUG: Available column indices: ${JSON.stringify(appColIndices)}`);
  
  // More robust column name matching
  const findColumnIndex = (possibleNames) => {
    for (const name of possibleNames) {
      if (appColIndices[name] !== undefined && appColIndices[name] !== -1) {
        Logger.log(`DEBUG: Found column "${name}" at index ${appColIndices[name]}`);
        return appColIndices[name];
      }
    }
    return -1;
  };
  
  const recruiterNameIdx = findColumnIndex(['Recruiter name', 'Recruiter_name', 'RecruiterName', 'recruiter name', 'recruiter_name']);
  const lastStageIdx = findColumnIndex(['Last_stage', 'Last stage', 'LastStage', 'last_stage', 'last stage']);
  const aiInterviewIdx = findColumnIndex(['Ai_interview', 'AI_interview', 'AI Interview', 'ai_interview', 'ai interview']);
  const applicationTsIdx = findColumnIndex(['Application_ts', 'Application ts', 'ApplicationTs', 'application_ts']);
  
  Logger.log(`DEBUG: Column indices found - Recruiter name: ${recruiterNameIdx}, Last_stage: ${lastStageIdx}, Ai_interview: ${aiInterviewIdx}, Application_ts: ${applicationTsIdx}`);
  
  if (recruiterNameIdx === -1 || lastStageIdx === -1 || aiInterviewIdx === -1) {
    Logger.log(`ERROR: Required columns not found for AI coverage calculation.`);
    Logger.log(`ERROR: Available column names: ${Object.keys(appColIndices).join(', ')}`);
    Logger.log(`ERROR: Recruiter: ${recruiterNameIdx}, Last_stage: ${lastStageIdx}, Ai_interview: ${aiInterviewIdx}`);
    return null;
  }
  
  const recruiterCoverage = {};
  let totalEligible = 0;
  let totalAIInterviews = 0;
  
  // Debug: Log first few rows to see the data structure
  Logger.log(`DEBUG: Processing ${appRows.length} rows`);
  if (appRows.length > 0) {
    Logger.log(`DEBUG: First row sample - Recruiter: "${appRows[0][recruiterNameIdx]}", Last Stage: "${appRows[0][lastStageIdx]}", AI Interview: "${appRows[0][aiInterviewIdx]}"`);
    
    // Log more sample rows to see the data variety
    for (let i = 0; i < Math.min(5, appRows.length); i++) {
      const row = appRows[i];
      if (row && row.length > Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
        const recruiter = String(row[recruiterNameIdx] || '').trim();
        const stage = String(row[lastStageIdx] || '').trim().toUpperCase();
        const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
        Logger.log(`DEBUG: Row ${i} - Recruiter: "${recruiter}", Stage: "${stage}", AI Interview: "${aiInterview}"`);
      }
    }
  }
  
  // Set May 1st, 2025 as the cutoff date
  const mayFirst2025 = new Date('2025-05-01');
  mayFirst2025.setHours(0, 0, 0, 0);
  
  appRows.forEach((row, index) => {
    // Basic validation
    if (!row || row.length <= Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to incomplete data`);
      return; // Skip incomplete rows
    }
    
    const recruiterName = String(row[recruiterNameIdx] || '').trim();
    const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
    const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
    
    // Skip if no recruiter name
    if (!recruiterName) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to missing recruiter name`);
      return;
    }
    
    // Skip excluded recruiters
    const excludedRecruiters = ['Samrudh J', 'Pavan Kumar', 'Guruprasad Hegde'];
    if (excludedRecruiters.some(excluded => recruiterName.toLowerCase().includes(excluded.toLowerCase()))) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to excluded recruiter: ${recruiterName}`);
      return;
    }
    
    // Check Application_ts filter (May 1st, 2025 or later)
    const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
    if (!applicationTs || applicationTs < mayFirst2025) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to Application_ts before May 1st, 2025: ${applicationTs}`);
      return; // Skip candidates with application timestamp before May 1st, 2025
    }
    
    // Check if candidate is eligible (only specific stages) - CASE INSENSITIVE
    const eligibleStages = [
      'HIRING MANAGER SCREEN',
      'ASSESSMENT', 
      'ONSITE INTERVIEW',
      'FINAL INTERVIEW',
      'OFFER APPROVALS',
      'OFFER EXTENDED',
      'OFFER DECLINED',
      'PENDING START',
      'HIRED'
    ];
    const isEligible = eligibleStages.some(stage => stage.toUpperCase() === lastStage);
    
    if (index < 5) {
      Logger.log(`DEBUG: Row ${index} - Recruiter: "${recruiterName}", Last Stage: "${lastStage}", AI Interview: "${aiInterview}", Eligible: ${isEligible}`);
    }
    
    if (isEligible) {
      totalEligible++;
      
      // Initialize recruiter data if not exists
      if (!recruiterCoverage[recruiterName]) {
        recruiterCoverage[recruiterName] = {
          totalEligible: 0,
          totalAIInterviews: 0,
          percentage: 0
        };
      }
      
      recruiterCoverage[recruiterName].totalEligible++;
      
      // Check if AI interview was conducted
      if (aiInterview === 'Y') {
        totalAIInterviews++;
        recruiterCoverage[recruiterName].totalAIInterviews++;
      }
    }
  });
  
  // Calculate percentages for each recruiter
  Object.keys(recruiterCoverage).forEach(recruiter => {
    const data = recruiterCoverage[recruiter];
    data.percentage = data.totalEligible > 0 ? 
      parseFloat(((data.totalAIInterviews / data.totalEligible) * 100).toFixed(1)) : 0;
  });
  
  // Calculate overall percentage
  const overallPercentage = totalEligible > 0 ? 
    parseFloat(((totalAIInterviews / totalEligible) * 100).toFixed(1)) : 0;
  
  Logger.log(`AI Coverage Metrics: Total eligible candidates = ${totalEligible}, Total AI interviews = ${totalAIInterviews}, Overall percentage = ${overallPercentage}%`);
  Logger.log(`DEBUG: Recruiter coverage data: ${JSON.stringify(recruiterCoverage)}`);
  
  // If no eligible candidates found, return a special object to show the table with a message
  if (totalEligible === 0) {
    Logger.log(`WARNING: No eligible candidates found. This could mean no candidates are in the target stages.`);
    return {
      recruiterCoverage: {},
      totalEligible: 0,
      totalAIInterviews: 0,
      overallPercentage: 0,
      noDataMessage: "No eligible candidates found. No candidates appear to be in the target stages (Hiring Manager Screen, Assessment, Onsite Interview, Final Interview, Offer Approvals, Offer Extended, Offer Declined, Pending Start, Hired), or there may be a data issue."
    };
  }
  
  return {
    recruiterCoverage,
    totalEligible,
    totalAIInterviews,
    overallPercentage
  };
}

/**
 * Test function to debug AI coverage calculation
 */
function testAICoverageCalculation() {
  try {
    Logger.log(`--- Testing AI Coverage Calculation ---`);
    Logger.log(`Filter: Application_ts ≥ May 1st, 2025`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    Logger.log(`Column indices: ${JSON.stringify(appData.colIndices)}`);
    
    // Test AI coverage calculation
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
    
    if (aiCoverageMetrics) {
      Logger.log(`SUCCESS: AI Coverage metrics calculated`);
      Logger.log(`Total eligible: ${aiCoverageMetrics.totalEligible}`);
      Logger.log(`Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}`);
      Logger.log(`Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
      Logger.log(`Recruiter coverage: ${JSON.stringify(aiCoverageMetrics.recruiterCoverage)}`);
    } else {
      Logger.log(`ERROR: AI Coverage metrics returned null`);
    }
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Generates HTML for a bar chart showing AI interview coverage by recruiter.
 * Each bar represents a recruiter with Y (AI interview done) and N (AI interview missing) stacked.
 * @param {object} aiCoverageMetrics The AI coverage metrics object from calculateAICoverageMetricsRB.
 * @returns {string} HTML string for the bar chart.
 */
function generateAICoverageBarChartHtml(aiCoverageMetrics) {
  if (!aiCoverageMetrics || !aiCoverageMetrics.recruiterCoverage || Object.keys(aiCoverageMetrics.recruiterCoverage).length === 0) {
    return `
      <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📊 AI Interview Coverage by Recruiter (Bar Chart)</div>
        <div style="text-align: center; padding: 20px; color: #666; font-size: 14px;">No AI coverage data available or no eligible candidates found.</div>
      </div>
    `;
  }

  // Sort recruiters by total eligible candidates (descending)
  const sortedRecruiters = Object.entries(aiCoverageMetrics.recruiterCoverage)
    .sort(([, a], [, b]) => b.totalEligible - a.totalEligible);

  // Calculate max value for scaling
  const maxEligible = Math.max(...sortedRecruiters.map(([, data]) => data.totalEligible));

  let chartHtml = `
    <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
      <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📊 AI Interview Coverage by Recruiter (Bar Chart)</div>
      <p style="font-size: 12px; color: #666; margin-bottom: 15px;">
        Shows eligible candidates (in specific stages: Hiring Manager Screen, Assessment, Onsite Interview, Final Interview, Offer Approvals, Offer Extended, Offer Declined, Pending Start, Hired; Application_ts ≥ May 1st, 2025) by recruiter. 
        <span style="color: #4CAF50; font-weight: bold;">Green</span> = AI interview done (Y), 
        <span style="color: #F44336; font-weight: bold;">Red</span> = AI interview missing (N).
      </p>
      
      <div style="margin: 20px 0;">
  `;

  // Generate bars
  sortedRecruiters.forEach(([recruiter, data]) => {
    const totalEligible = data.totalEligible;
    const aiInterviewsDone = data.totalAIInterviews;
    const aiInterviewsMissing = totalEligible - aiInterviewsDone;
    
    // Calculate bar widths (max width 300px)
    const maxBarWidth = 300;
    const barWidth = (totalEligible / maxEligible) * maxBarWidth;
    const doneWidth = (aiInterviewsDone / totalEligible) * barWidth;
    const missingWidth = (aiInterviewsMissing / totalEligible) * barWidth;
    
    // Truncate long recruiter names
    const displayName = recruiter.length > 20 ? recruiter.substring(0, 17) + '...' : recruiter;
    
    chartHtml += `
      <div style="margin-bottom: 15px;">
        <div style="display: flex; align-items: center; margin-bottom: 5px;">
          <div style="width: 120px; font-size: 12px; font-weight: bold; text-align: right; padding-right: 10px; overflow: hidden; text-overflow: ellipsis;" title="${recruiter}">
            ${displayName}
          </div>
          <div style="flex: 1; display: flex; align-items: center;">
            <div style="width: ${barWidth}px; height: 25px; display: flex; border: 1px solid #ccc; border-radius: 3px; overflow: hidden;">
              ${aiInterviewsDone > 0 ? `<div style="width: ${doneWidth}px; background-color: #4CAF50; display: flex; align-items: center; justify-content: center; color: white; font-size: 10px; font-weight: bold;" title="AI Interview Done: ${aiInterviewsDone}">${aiInterviewsDone}</div>` : ''}
              ${aiInterviewsMissing > 0 ? `<div style="width: ${missingWidth}px; background-color: #F44336; display: flex; align-items: center; justify-content: center; color: white; font-size: 10px; font-weight: bold;" title="AI Interview Missing: ${aiInterviewsMissing}">${aiInterviewsMissing}</div>` : ''}
            </div>
            <div style="margin-left: 10px; font-size: 11px; color: #666; min-width: 80px;">
              ${data.percentage}% (${aiInterviewsDone}/${totalEligible})
            </div>
          </div>
        </div>
      </div>
    `;
  });

  // Add legend
  chartHtml += `
      </div>
      
      <div style="display: flex; justify-content: center; align-items: center; margin-top: 20px; font-size: 12px;">
        <div style="display: flex; align-items: center; margin-right: 20px;">
          <div style="width: 20px; height: 15px; background-color: #4CAF50; margin-right: 8px; border: 1px solid #ccc;"></div>
          <span>AI Interview Done (Y)</span>
        </div>
        <div style="display: flex; align-items: center;">
          <div style="width: 20px; height: 15px; background-color: #F44336; margin-right: 8px; border: 1px solid #ccc;"></div>
          <span>AI Interview Missing (N)</span>
        </div>
      </div>
      
      <p style="font-size: 11px; color: #666; margin-top: 15px; text-align: center;">
        Total eligible candidates: ${aiCoverageMetrics.totalEligible} | 
        Total AI interviews done: ${aiCoverageMetrics.totalAIInterviews} | 
        Overall coverage: <strong>${aiCoverageMetrics.overallPercentage}%</strong>
      </p>
    </div>
  `;

  return chartHtml;
}

/**
 * Test function to debug AI coverage bar chart
 */
function testAICoverageBarChart() {
  try {
    Logger.log(`--- Testing AI Coverage Bar Chart ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test AI coverage calculation
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
    
    if (aiCoverageMetrics) {
      Logger.log(`SUCCESS: AI Coverage metrics calculated`);
      Logger.log(`Total eligible: ${aiCoverageMetrics.totalEligible}`);
      Logger.log(`Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}`);
      Logger.log(`Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
      
      // Test bar chart generation
      const barChartHtml = generateAICoverageBarChartHtml(aiCoverageMetrics);
      Logger.log(`SUCCESS: Bar chart HTML generated (${barChartHtml.length} characters)`);
      
      // Log sample of the HTML
      Logger.log(`Sample HTML (first 500 chars): ${barChartHtml.substring(0, 500)}...`);
      
      // Log recruiter data for verification
      Object.entries(aiCoverageMetrics.recruiterCoverage).forEach(([recruiter, data]) => {
        const missing = data.totalEligible - data.totalAIInterviews;
        Logger.log(`Recruiter: ${recruiter} | Eligible: ${data.totalEligible} | AI Done: ${data.totalAIInterviews} | Missing: ${missing} | %: ${data.percentage}%`);
      });
      
    } else {
      Logger.log(`ERROR: AI Coverage metrics returned null`);
    }
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Creates a detailed validation spreadsheet for a specific recruiter showing all their candidates.
 * @param {string} recruiterName The name of the recruiter to analyze.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {string} The URL of the created Google Sheet.
 */
function createRecruiterValidationSheet(recruiterName, appRows, appColIndices) {
  Logger.log(`--- Creating validation sheet for recruiter: ${recruiterName} ---`);
  
  try {
    const spreadsheet = SpreadsheetApp.create(`AI Interview Validation - ${recruiterName} - ${new Date().toISOString().split('T')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    // Set May 1st, 2025 as the cutoff date
    const mayFirst2025 = new Date('2025-05-01');
    mayFirst2025.setHours(0, 0, 0, 0);
    
    // Find column indices
    const findColumnIndex = (possibleNames) => {
      for (const name of possibleNames) {
        if (appColIndices[name] !== undefined && appColIndices[name] !== -1) {
          return appColIndices[name];
        }
      }
      return -1;
    };
    
    const recruiterNameIdx = findColumnIndex(['Recruiter name', 'Recruiter_name', 'RecruiterName', 'recruiter name', 'recruiter_name']);
    const lastStageIdx = findColumnIndex(['Last_stage', 'Last stage', 'LastStage', 'last_stage', 'last stage']);
    const aiInterviewIdx = findColumnIndex(['Ai_interview', 'AI_interview', 'AI Interview', 'ai_interview', 'ai interview']);
    const applicationTsIdx = findColumnIndex(['Application_ts', 'Application ts', 'ApplicationTs', 'application_ts']);
    const nameIdx = findColumnIndex(['Name', 'name', 'Candidate_name', 'Candidate name']);
    const positionIdIdx = findColumnIndex(['Position_id', 'Position id', 'Position ID']);
    const titleIdx = findColumnIndex(['Title', 'title', 'Position_title', 'Position title']);
    const currentCompanyIdx = findColumnIndex(['Current_company', 'Current company', 'Company', 'company']);
    const applicationStatusIdx = findColumnIndex(['Application_status', 'Application status', 'Status', 'status']);
    const positionStatusIdx = findColumnIndex(['Position_status', 'Position status']);
    const matchStarsIdx = findColumnIndex(['Match_stars', 'Match score', 'Match Stars', 'MatchStars', 'Match_Stars', 'Stars', 'Score']);
    
    if (recruiterNameIdx === -1 || lastStageIdx === -1 || aiInterviewIdx === -1) {
      throw new Error('Required columns not found for validation sheet');
    }
    
    // Define headers for the validation sheet
    const headers = [
      'Candidate Name',
      'Position ID', 
      'Position Title',
      'Current Company',
      'Application Status',
      'Position Status',
      'Last Stage',
      'Application Timestamp',
      'AI Interview (Y/N)',
      'Match Stars/Score',
      'Eligible for AI Interview',
      'AI Interview Status',
      'Days Since Application'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    // Filter data for the specific recruiter and apply date filter
    const recruiterData = appRows.filter(row => {
      if (!row || row.length <= Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
        return false;
      }
      
      const rowRecruiterName = String(row[recruiterNameIdx] || '').trim();
      if (rowRecruiterName !== recruiterName) {
        return false;
      }
      
      // Check Application_ts filter (May 1st, 2025 or later)
      const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
      if (!applicationTs || applicationTs < mayFirst2025) {
        return false;
      }
      
      return true;
    });
    
    Logger.log(`Found ${recruiterData.length} candidates for ${recruiterName} with Application_ts ≥ May 1st, 2025`);
    
    // Process each candidate
    const validationRows = recruiterData.map(row => {
      const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
      const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
      const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
      
      // Check if candidate is eligible (only specific stages) - CASE INSENSITIVE
      const eligibleStages = [
        'HIRING MANAGER SCREEN',
        'ASSESSMENT', 
        'ONSITE INTERVIEW',
        'FINAL INTERVIEW',
        'OFFER APPROVALS',
        'OFFER EXTENDED',
        'OFFER DECLINED',
        'PENDING START',
        'HIRED'
      ];
      const isEligible = eligibleStages.some(stage => stage.toUpperCase() === lastStage);
      
      // Determine AI interview status
      let aiInterviewStatus = 'N/A';
      if (isEligible) {
        if (aiInterview === 'Y') {
          aiInterviewStatus = '✅ AI Interview Done';
        } else if (aiInterview === 'N') {
          aiInterviewStatus = '❌ AI Interview Missing';
        } else {
          aiInterviewStatus = '❓ Unknown Status';
        }
      } else {
        aiInterviewStatus = '⏭️ Not Eligible (Not in target stages)';
      }
      
      // Calculate days since application
      const daysSinceApplication = applicationTs ? 
        Math.floor((new Date() - applicationTs) / (1000 * 60 * 60 * 24)) : 'N/A';
      
      return [
        nameIdx !== -1 ? row[nameIdx] || 'N/A' : 'N/A',
        positionIdIdx !== -1 ? row[positionIdIdx] || 'N/A' : 'N/A',
        titleIdx !== -1 ? row[titleIdx] || 'N/A' : 'N/A',
        currentCompanyIdx !== -1 ? row[currentCompanyIdx] || 'N/A' : 'N/A',
        applicationStatusIdx !== -1 ? row[applicationStatusIdx] || 'N/A' : 'N/A',
        positionStatusIdx !== -1 ? row[positionStatusIdx] || 'N/A' : 'N/A',
        lastStage,
        applicationTs ? applicationTs.toLocaleDateString() : 'N/A',
        aiInterview,
        matchStarsIdx !== -1 ? row[matchStarsIdx] || 'N/A' : 'N/A',
        isEligible ? 'Yes' : 'No',
        aiInterviewStatus,
        daysSinceApplication
      ];
    });
    
    // Sort by AI interview status (missing first, then done, then not eligible)
    validationRows.sort((a, b) => {
      const statusA = a[11]; // AI Interview Status column
      const statusB = b[11];
      
      if (statusA.includes('Missing') && !statusB.includes('Missing')) return -1;
      if (!statusA.includes('Missing') && statusB.includes('Missing')) return 1;
      if (statusA.includes('Done') && !statusB.includes('Done')) return -1;
      if (!statusA.includes('Done') && statusB.includes('Done')) return 1;
      return 0;
    });
    
    if (validationRows.length > 0) {
      sheet.getRange(2, 1, validationRows.length, headers.length).setValues(validationRows);
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    // Add summary statistics
    const totalCandidates = validationRows.length;
    const eligibleCandidates = validationRows.filter(row => row[10] === 'Yes').length;
    const aiInterviewsDone = validationRows.filter(row => row[11].includes('Done')).length;
    const aiInterviewsMissing = validationRows.filter(row => row[11].includes('Missing')).length;
    const notEligible = validationRows.filter(row => row[11].includes('Not Eligible')).length;
    
    // Add summary at the top
    const summaryRow = [
      `SUMMARY FOR ${recruiterName.toUpperCase()}`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ];
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([summaryRow]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
    
    // Add statistics rows
    const statsRows = [
      [`Total Candidates (Application_ts ≥ May 1st, 2025): ${totalCandidates}`],
      [`Eligible Candidates (in target stages): ${eligibleCandidates}`],
      [`AI Interviews Done: ${aiInterviewsDone}`],
      [`AI Interviews Missing: ${aiInterviewsMissing}`],
      [`Not Eligible (not in target stages): ${notEligible}`],
      [`Coverage Rate: ${eligibleCandidates > 0 ? ((aiInterviewsDone / eligibleCandidates) * 100).toFixed(1) : 0}%`]
    ];
    
    sheet.insertRowsBefore(2, statsRows.length);
    for (let i = 0; i < statsRows.length; i++) {
      sheet.getRange(2 + i, 1, 1, 1).setValues([statsRows[i]]);
      sheet.getRange(2 + i, 1, 1, 1).setFontWeight('bold');
    }
    
    // Add conditional formatting for AI interview status
    const statusRange = sheet.getRange(3 + statsRows.length, 12, validationRows.length, 1); // AI Interview Status column
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✅ AI Interview Done')
      .setBackground('#d4edda')
      .setRanges([statusRange])
      .build();
    
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('❌ AI Interview Missing')
      .setBackground('#f8d7da')
      .setRanges([statusRange])
      .build();
    
    const rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⏭️ Not Eligible (New/Added)')
      .setBackground('#fff3cd')
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([rule1, rule2, rule3]);
    
    Logger.log(`Validation sheet created for ${recruiterName} with ${validationRows.length} candidates`);
    Logger.log(`Summary: ${eligibleCandidates} eligible, ${aiInterviewsDone} AI done, ${aiInterviewsMissing} AI missing`);
    
    return spreadsheet.getUrl();
    
  } catch (error) {
    Logger.log(`Error creating validation sheet for ${recruiterName}: ${error.toString()} Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Test function to create validation sheet for a specific recruiter
 */
function testCreateRecruiterValidationSheet() {
  try {
    Logger.log(`--- Testing Recruiter Validation Sheet Creation ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test with a specific recruiter (you can change this name)
    const testRecruiter = 'Akhila Kashyap';
    Logger.log(`Creating validation sheet for: ${testRecruiter}`);
    
    const sheetUrl = createRecruiterValidationSheet(testRecruiter, appData.rows, appData.colIndices);
    Logger.log(`SUCCESS: Validation sheet created: ${sheetUrl}`);
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Creates validation sheets for all recruiters and returns a summary of URLs.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} Object containing validation sheet URLs and summary.
 */
function createAllRecruiterValidationSheets(appRows, appColIndices) {
  Logger.log(`--- Creating validation sheets for all recruiters ---`);
  
  try {
    // First, get the list of all recruiters from the AI coverage data
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appRows, appColIndices);
    if (!aiCoverageMetrics || !aiCoverageMetrics.recruiterCoverage) {
      Logger.log(`ERROR: Could not calculate AI coverage metrics`);
      return null;
    }
    
    const recruiterNames = Object.keys(aiCoverageMetrics.recruiterCoverage);
    Logger.log(`Found ${recruiterNames.length} recruiters to create validation sheets for`);
    
    const validationSheets = {};
    const failedRecruiters = [];
    
    // Create validation sheet for each recruiter
    recruiterNames.forEach(recruiterName => {
      try {
        Logger.log(`Creating validation sheet for: ${recruiterName}`);
        const sheetUrl = createRecruiterValidationSheet(recruiterName, appRows, appColIndices);
        validationSheets[recruiterName] = {
          url: sheetUrl,
          eligible: aiCoverageMetrics.recruiterCoverage[recruiterName].totalEligible,
          aiDone: aiCoverageMetrics.recruiterCoverage[recruiterName].totalAIInterviews,
          aiMissing: aiCoverageMetrics.recruiterCoverage[recruiterName].totalEligible - aiCoverageMetrics.recruiterCoverage[recruiterName].totalAIInterviews,
          percentage: aiCoverageMetrics.recruiterCoverage[recruiterName].percentage
        };
        Logger.log(`SUCCESS: Validation sheet created for ${recruiterName}`);
      } catch (error) {
        Logger.log(`ERROR creating validation sheet for ${recruiterName}: ${error.toString()}`);
        failedRecruiters.push(recruiterName);
      }
    });
    
    Logger.log(`Created ${Object.keys(validationSheets).length} validation sheets successfully`);
    if (failedRecruiters.length > 0) {
      Logger.log(`Failed to create sheets for: ${failedRecruiters.join(', ')}`);
    }
    
    return {
      validationSheets,
      failedRecruiters,
      totalRecruiters: recruiterNames.length,
      successfulSheets: Object.keys(validationSheets).length
    };
    
  } catch (error) {
    Logger.log(`ERROR in createAllRecruiterValidationSheets: ${error.toString()} Stack: ${error.stack}`);
    return null;
  }
}

/**
 * Generates HTML for validation sheets section in the report.
 * @param {object} validationData The validation data from createAllRecruiterValidationSheets.
 * @returns {string} HTML string for the validation sheets section.
 */
function generateValidationSheetsHtml(validationData) {
  if (!validationData || !validationData.validationSheets || Object.keys(validationData.validationSheets).length === 0) {
    return `
      <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📋 Detailed Validation Sheets</div>
        <div style="text-align: center; padding: 20px; color: #666; font-size: 14px;">No validation sheets available.</div>
      </div>
    `;
  }
  
  // Sort recruiters by AI interviews missing (descending) to highlight those needing attention
  const sortedRecruiters = Object.entries(validationData.validationSheets)
    .sort(([, a], [, b]) => b.aiMissing - a.aiMissing);
  
  let html = `
    <div style="background-color: #fff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 15px;">
      <div style="font-weight: bold; font-size: 16px; color: #3f51b5; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #eee;">📋 Detailed Validation Sheets</div>
      <p style="font-size: 12px; color: #666; margin-bottom: 15px;">
        Click on any recruiter name to view their detailed candidate list with AI interview status. 
        Sheets show all candidates with Application_ts ≥ May 1st, 2025, sorted by AI interview status.
      </p>
      
      <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
        <thead>
          <tr>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">Recruiter Name</th>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">Eligible Candidates</th>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">AI Done</th>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">AI Missing</th>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">Coverage %</th>
            <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; background-color: #f5f5f5; color: #424242; font-weight: bold;">Action</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  sortedRecruiters.forEach(([recruiterName, data], index) => {
    const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
    const coverageColor = data.percentage >= 80 ? '#4CAF50' : data.percentage >= 60 ? '#FF9800' : '#F44336';
    const missingPercentage = 100 - data.percentage;
    const priorityIcon = missingPercentage > 30 ? '🔴' : missingPercentage > 10 ? '🟡' : '🟢';
    
    html += `
      <tr style="background-color: ${bgColor};">
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; font-weight: bold;">
          ${priorityIcon} ${recruiterName}
        </td>
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px;">${data.eligible}</td>
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; color: #4CAF50; font-weight: bold;">${data.aiDone}</td>
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; color: #F44336; font-weight: bold;">${data.aiMissing}</td>
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; color: ${coverageColor}; font-weight: bold;">${data.percentage}%</td>
        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px;">
          <a href="${data.url}" target="_blank" style="color: #1976d2; text-decoration: none; font-weight: bold;">📊 View Details</a>
        </td>
      </tr>
    `;
  });
  
  html += `
        </tbody>
      </table>
      
      <p style="font-size: 11px; color: #666; margin-top: 15px; text-align: center;">
        🔴 High Priority (>30% missing) | 🟡 Medium Priority (10-30% missing) | 🟢 Low Priority (<10% missing)<br>
        Created ${validationData.successfulSheets} validation sheets successfully.
        ${validationData.failedRecruiters.length > 0 ? `Failed to create sheets for: ${validationData.failedRecruiters.join(', ')}` : ''}
      </p>
    </div>
  `;
  
  return html;
}

/**
 * Comprehensive test function for validation sheet functionality
 */
function testValidationSheetFunctionality() {
  try {
    Logger.log(`--- Testing Complete Validation Sheet Functionality ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test 1: Single recruiter validation sheet
    Logger.log(`\n--- Test 1: Single Recruiter Validation Sheet ---`);
    const testRecruiter = 'Akhila Kashyap';
    const singleSheetUrl = createRecruiterValidationSheet(testRecruiter, appData.rows, appData.colIndices);
    Logger.log(`SUCCESS: Single validation sheet created for ${testRecruiter}: ${singleSheetUrl}`);
    
    // Test 2: All recruiters validation sheets
    Logger.log(`\n--- Test 2: All Recruiters Validation Sheets ---`);
    const allSheetsData = createAllRecruiterValidationSheets(appData.rows, appData.colIndices);
    if (allSheetsData) {
      Logger.log(`SUCCESS: Created ${allSheetsData.successfulSheets} validation sheets out of ${allSheetsData.totalRecruiters} recruiters`);
      Logger.log(`Failed recruiters: ${allSheetsData.failedRecruiters.length > 0 ? allSheetsData.failedRecruiters.join(', ') : 'None'}`);
      
      // Log details for each recruiter
      Object.entries(allSheetsData.validationSheets).forEach(([recruiter, data]) => {
        Logger.log(`${recruiter}: ${data.eligible} eligible, ${data.aiDone} AI done, ${data.aiMissing} AI missing, ${data.percentage}% coverage`);
      });
    } else {
      Logger.log(`ERROR: Could not create all recruiter validation sheets`);
    }
    
    // Test 3: HTML generation
    Logger.log(`\n--- Test 3: HTML Generation ---`);
    if (allSheetsData) {
      const htmlContent = generateValidationSheetsHtml(allSheetsData);
      Logger.log(`SUCCESS: Generated HTML content (${htmlContent.length} characters)`);
      Logger.log(`HTML preview (first 500 chars): ${htmlContent.substring(0, 500)}...`);
    }
    
    Logger.log(`\n--- All Validation Sheet Tests Completed Successfully ---`);
    
  } catch (error) {
    Logger.log(`ERROR in comprehensive test: ${error.toString()}`);
  }
}

/**
 * Detailed debug function to analyze Akhila's data and understand the filtering issue
 */
function debugAkhilaData() {
  try {
    Logger.log(`=== DETAILED DEBUG: AKHILA'S DATA ===`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Total rows in dataset: ${appData.rows.length}`);
    Logger.log(`Column indices: ${JSON.stringify(appData.colIndices)}`);
    
    const { rows, colIndices } = appData;
    const { recruiterNameIdx, lastStageIdx, aiInterviewIdx, applicationTsIdx } = colIndices;
    
    // Set May 1st, 2025 as the cutoff date
    const mayFirst2025 = new Date('2025-05-01');
    mayFirst2025.setHours(0, 0, 0, 0);
    
    // Define eligible stages
    const eligibleStages = [
      'HIRING MANAGER SCREEN',
      'ASSESSMENT', 
      'ONSITE INTERVIEW',
      'FINAL INTERVIEW',
      'OFFER APPROVALS',
      'OFFER EXTENDED',
      'OFFER DECLINED',
      'PENDING START',
      'HIRED'
    ];
    
    // Track Akhila's data specifically
    let akhilaTotal = 0;
    let akhilaAfterMay1st = 0;
    let akhilaEligible = 0;
    let akhilaStages = {};
    let akhilaApplicationDates = [];
    
    // Analyze all rows for Akhila
    rows.forEach((row, index) => {
      const recruiterName = String(row[recruiterNameIdx] || '').trim();
      
      // Only process Akhila's data
      if (recruiterName.toLowerCase().includes('akhila')) {
        akhilaTotal++;
        
        const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
        const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
        const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
        
        // Track stages
        if (!akhilaStages[lastStage]) {
          akhilaStages[lastStage] = 0;
        }
        akhilaStages[lastStage]++;
        
        // Check Application_ts filter
        if (applicationTs && applicationTs >= mayFirst2025) {
          akhilaAfterMay1st++;
          akhilaApplicationDates.push(applicationTs.toISOString().split('T')[0]);
          
          // Check if eligible stage (case insensitive)
          if (eligibleStages.some(stage => stage.toUpperCase() === lastStage)) {
            akhilaEligible++;
            Logger.log(`AKHILA ELIGIBLE: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs.toISOString().split('T')[0]}`);
          } else {
            Logger.log(`AKHILA NOT ELIGIBLE: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs.toISOString().split('T')[0]}`);
          }
        } else {
          Logger.log(`AKHILA BEFORE MAY 1ST: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs ? applicationTs.toISOString().split('T')[0] : 'N/A'}`);
        }
      }
    });
    
    Logger.log(`=== AKHILA SUMMARY ===`);
    Logger.log(`Total Akhila candidates: ${akhilaTotal}`);
    Logger.log(`Akhila candidates after May 1st, 2025: ${akhilaAfterMay1st}`);
    Logger.log(`Akhila eligible candidates: ${akhilaEligible}`);
    Logger.log(`Akhila stages breakdown: ${JSON.stringify(akhilaStages)}`);
    Logger.log(`Akhila application dates (after May 1st): ${akhilaApplicationDates.slice(0, 10).join(', ')}...`);
    
    // Also check for variations of Akhila's name
    Logger.log(`=== CHECKING FOR NAME VARIATIONS ===`);
    const nameVariations = {};
    rows.forEach((row, index) => {
      const recruiterName = String(row[recruiterNameIdx] || '').trim();
      if (recruiterName.toLowerCase().includes('akhila') || recruiterName.toLowerCase().includes('kashyap')) {
        if (!nameVariations[recruiterName]) {
          nameVariations[recruiterName] = 0;
        }
        nameVariations[recruiterName]++;
      }
    });
    Logger.log(`Name variations found: ${JSON.stringify(nameVariations)}`);
    
  } catch (error) {
    Logger.log(`ERROR in debugAkhilaData: ${error.toString()}`);
  }
}

/**
 * Function to list all current filters being applied in AI coverage calculation
 */
function listCurrentFilters() {
  Logger.log(`=== CURRENT FILTERS IN AI COVERAGE CALCULATION ===`);
  Logger.log(`1. Application_ts filter: Must be ≥ May 1st, 2025`);
  Logger.log(`2. Last_stage filter: Must be one of the following:`);
  Logger.log(`   - HIRING MANAGER SCREEN`);
  Logger.log(`   - ASSESSMENT`);
  Logger.log(`   - ONSITE INTERVIEW`);
  Logger.log(`   - FINAL INTERVIEW`);
  Logger.log(`   - OFFER APPROVALS`);
  Logger.log(`   - OFFER EXTENDED`);
  Logger.log(`   - OFFER DECLINED`);
  Logger.log(`   - PENDING START`);
  Logger.log(`   - HIRED`);
  Logger.log(`3. Recruiter name must not be empty`);
  Logger.log(`4. All required columns must have data (recruiter, last stage, AI interview, application timestamp)`);
  Logger.log(`5. Excluded recruiters: Samrudh J, Pavan Kumar, Guruprasad Hegde`);
  Logger.log(`=== END OF FILTERS ===`);
}