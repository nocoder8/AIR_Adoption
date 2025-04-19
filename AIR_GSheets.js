/**
 * AI Interview Report Generator - Google Apps Script version : Gemini 2.17 - % Inline with Name
 * * This script analyzes candidate data in a Google Sheet to generate insights about AI interview adoption,
 * focusing on segmenting candidates by application date relative to the AI Recruiter launch date.
 * * To use:
 * 1. Open your Google Sheet with candidate data
 * 2. Go to Extensions > Apps Script
 * 3. DELETE ALL existing code in the editor
 * 4. Copy and paste ALL of this code below
 * 5. Update the EMAIL_RECIPIENT constant if needed
 * 6. Save and run the script
 */

// Configuration - Update these values
const EMAIL_RECIPIENT = 'pkumar@eightfold.ai';
const EMAIL_CC = 'pkumar@eightfold.ai';
const SHEET_NAME = 'Active+Rejected'; // Sheet name corrected
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw/edit?gid=1957093905#gid=1957093905';
// >>> IMPORTANT: Please double-check this LAUNCH_DATE is correct for your needs <<<
const LAUNCH_DATE = new Date('2025-04-17'); // AI Recruiter launch date

// URL for Looker Studio report
const LOOKER_STUDIO_URL = 'https://lookerstudio.google.com/reporting/b05c1dfb-d808-4eca-b70d-863fe5be0f27';

/**
 * Creates a trigger to run the report daily
 */
function createDailyTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateAndSendAIReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new trigger to run daily at 8 AM (Script's timezone)
  ScriptApp.newTrigger('generateAndSendAIReport')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log('Daily trigger set to run at 8 AM');
  SpreadsheetApp.getUi().alert('Daily trigger set to run generateAndSendAIReport function at 8 AM.'); // User feedback
}

/**
 * Main function to generate and send the AI interview report
 */
function generateAndSendAIReport() {
  try {
    // Get the spreadsheet by URL instead of active spreadsheet
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    Logger.log(`Opened spreadsheet: ${spreadsheet.getName()}`);

    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      // If the specified sheet is not found, try to get the sheet by ID
      Logger.log(`Sheet "${SHEET_NAME}" not found. Attempting to use sheet by gid.`);
      // Extract gid from URL
      const gidMatch = SPREADSHEET_URL.match(/gid=(\d+)/);
      if (gidMatch && gidMatch[1]) {
        const gid = gidMatch[1];
        Logger.log(`Looking for sheet with gid: ${gid}`);
        const sheets = spreadsheet.getSheets();
        Logger.log(`Found ${sheets.length} sheets in the spreadsheet.`);

        for (let i = 0; i < sheets.length; i++) {
          Logger.log(`Sheet ${i+1}: ${sheets[i].getName()} (ID: ${sheets[i].getSheetId()})`);
          if (sheets[i].getSheetId().toString() === gid) {
            sheet = sheets[i];
            Logger.log(`Using sheet by ID: "${sheet.getName()}"`);
            break;
          }
        }
      }

      // If still no sheet, use the first sheet
      if (!sheet) {
        const firstSheet = spreadsheet.getSheets()[0];
        if (!firstSheet) {
          throw new Error(`Could not find any sheets in the spreadsheet: ${SPREADSHEET_URL}`);
        }
        sheet = firstSheet;
        Logger.log(`Warning: Sheet "${SHEET_NAME}" not found. Using first available sheet: "${sheet.getName()}"`);
      }
    } else {
       Logger.log(`Using specified sheet: "${sheet.getName()}"`);
    }

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    Logger.log(`Sheet dimensions: ${data.length} rows x ${data[0].length} columns`);

    // Make sure we have enough rows
    if (data.length < 3) { // Need at least headers (row 2) and one data row
      throw new Error(`Not enough data rows in the sheet "${sheet.getName()}". Need at least headers (row 2) and one data row.`);
    }

    // Since we know headers are in row 2 (index 1), extract directly
    const headers = data[1].map(String); // Ensure headers are strings
    const rows = data.slice(2); // Skip first two rows

    Logger.log(`Headers found in row 2: ${JSON.stringify(headers)}`);

    // --- Find Match Stars Column ---
    let matchStarsColIndex = -1;
    let matchStarsColName = null; // <<< Initialize variable for column name
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

      if (matchStarsColIndex === -1) {
        Logger.log("Could not find match column by alternative names. Looking for partial matches:");
        headers.forEach((header, index) => {
          const headerText = header.toLowerCase();
          if (headerText.includes('match') && (headerText.includes('star') || headerText.includes('score'))) {
             Logger.log(`High confidence potential match column: [${index}] "${header}"`);
             if (matchStarsColIndex === -1) { matchStarsColIndex = index; }
          } else if (matchStarsColIndex === -1 && (headerText.includes('match') || headerText.includes('star') || headerText.includes('score'))) {
             Logger.log(`Lower confidence potential match column: [${index}] "${header}"`);
             // Tentatively select first low confidence match if no high confidence found yet
             if (matchStarsColIndex === -1 && index > 0) { // Avoid index 0 if possible unless it's the only match
                // This part might need refinement based on typical sheet structure
             }
          }
        });
      }
    }
    // --- End Find Match Stars Column ---

     if (matchStarsColIndex === -1) {
        Logger.log("WARNING: Could not find any suitable column for Match Stars/Score. The '≥4 Match' filter will not be applied.");
     } else {
        matchStarsColName = headers[matchStarsColIndex]; // <<< Get the column name
        Logger.log(`Using column index ${matchStarsColIndex} ("${matchStarsColName}") for match score.`);
     }


    // Verify required columns exist
    const requiredColumns = ['Last_stage', 'Ai_interview', 'Recruiter name', 'Title', 'Application_status', 'Position_status', 'Application_ts'];
    const missingCols = [];
    const colIndices = {};

    requiredColumns.forEach(colName => {
       const index = headers.indexOf(colName);
       if (index === -1) {
           missingCols.push(colName);
       } else {
           colIndices[colName] = index;
       }
    });

    if (missingCols.length > 0) {
      throw new Error(`Required column(s) not found in the sheet headers (Row 2): ${missingCols.join(', ')}`);
    }
    Logger.log("All required columns found.");


    // First segment all candidates by application date - without any status filters
    const postLaunchAllCandidates = [];
    const preLaunchAllCandidates = [];
    let invalidDateCount = 0;

    rows.forEach((candidate, idx) => {
      // Ensure candidate row has enough columns
      if (candidate.length <= colIndices['Application_ts']) {
         Logger.log(`Warning: Row ${idx + 3} has too few columns. Skipping.`);
         return; // Skip this row
      }
      const rawDate = candidate[colIndices['Application_ts']];
      // Check for null, undefined, or empty string before creating Date
      if (rawDate === null || rawDate === undefined || rawDate === '') {
         Logger.log(`Warning: Invalid or empty application date found in row ${idx + 3}. Skipping.`);
         invalidDateCount++;
         return; // Skip processing this candidate
      }
      const applicationDate = new Date(rawDate);

      if (!isNaN(applicationDate.getTime())) {
        if (applicationDate >= LAUNCH_DATE) {
          postLaunchAllCandidates.push(candidate);
        } else {
          preLaunchAllCandidates.push(candidate);
        }
      } else {
        Logger.log(`Warning: Could not parse application date "${rawDate}" in row ${idx + 3}. Skipping.`);
        invalidDateCount++;
      }
    });

    if (invalidDateCount > 0) {
        Logger.log(`Total warnings for invalid/unparseable dates: ${invalidDateCount}`);
    }

    Logger.log(`Total candidates processed: ${rows.length}`);
    Logger.log(`All post-launch candidates (valid date): ${postLaunchAllCandidates.length}`);
    Logger.log(`All pre-launch candidates (valid date): ${preLaunchAllCandidates.length}`);

    // Now, apply active+open filter only to pre-launch candidates
    const preLaunchActiveCandidates = preLaunchAllCandidates.filter(row => {
        // Check column existence before accessing
        const appStatus = row.length > colIndices['Application_status'] ? row[colIndices['Application_status']] : null;
        const posStatus = row.length > colIndices['Position_status'] ? row[colIndices['Position_status']] : null;
        return appStatus?.toLowerCase() === 'active' && posStatus?.toLowerCase() === 'open';
    });

    Logger.log(`Pre-launch active+open candidates: ${preLaunchActiveCandidates.length}`);

    // For post-launch, we use all candidates regardless of status
    Logger.log(`Post-launch candidates used in report (no status filter): ${postLaunchAllCandidates.length}`);

    // Generate metrics for different segments
    const preLaunchMetrics = generateSegmentMetrics(preLaunchActiveCandidates, colIndices, matchStarsColIndex, false); // No score filter needed for pre-launch view
    const postLaunchMetrics = generateSegmentMetrics(postLaunchAllCandidates, colIndices, matchStarsColIndex, false); // All post-launch
    const postLaunchMatchScore4Metrics = generateSegmentMetrics(postLaunchAllCandidates, colIndices, matchStarsColIndex, true, 4); // Post-launch with score filter >= 4

    // Prepare combined report data
    const reportData = {
      postLaunch: postLaunchMetrics,
      postLaunchMatchScore4: postLaunchMatchScore4Metrics,
      preLaunch: preLaunchMetrics, // Keep pre-launch separate for leaderboard calculation
      launchDate: LAUNCH_DATE.toLocaleDateString(),
      hasMatchStarsColumn: matchStarsColIndex !== -1,
      matchStarsColName: matchStarsColName // <<< Add the column name here
    };

    // Create and send email
    const htmlContent = createHtmlReport(reportData); // <<< Pass the updated reportData
    const reportTitle = `AI Recruiter Adoption Report - ${new Date().toLocaleDateString()}`; // Add date to subject
    sendEmail(EMAIL_RECIPIENT, reportTitle, htmlContent);

    Logger.log('Report generated and sent successfully!');
    return 'Report sent to ' + EMAIL_RECIPIENT;
  } catch (error) {
    Logger.log('Error in generateAndSendAIReport: ' + error.toString() + ' Stack: ' + error.stack);
    // Send error email
    try {
       const subject = `ERROR: AI Recruiter Report Failed - ${new Date().toLocaleString()}`;
       const body = `Error generating AI Recruiter report:\n\n${error.toString()}\n\nStack:\n${error.stack}\n\nSheet: ${SPREADSHEET_URL}`;
       MailApp.sendEmail(EMAIL_RECIPIENT, subject, body);
       Logger.log('Error notification email sent.');
    } catch (emailError) {
       Logger.log('CRITICAL: Failed to send error notification email: ' + emailError);
    }
    return 'Error: ' + error.toString();
  }
}

/**
 * Generate metrics for a specific segment of candidates
 * @param {Array[]} candidates The array of candidate rows for this segment.
 * @param {Object} colIndices Map of required column names to their indices.
 * @param {number} matchStarsColIndex Index of the match score column, or -1 if not found.
 * @param {boolean} applyMatchFilter Whether to apply the match score filter.
 * @param {number} matchScoreThreshold The minimum score for the match filter (e.g., 4).
 * @returns {Object} An object containing metrics for the segment.
 */
function generateSegmentMetrics(candidates, colIndices, matchStarsColIndex, applyMatchFilter, matchScoreThreshold = 0) {
    let filteredCandidates = candidates;
    let filterApplied = false;

    // Apply match score filter if requested and possible
    if (applyMatchFilter && matchStarsColIndex !== -1 && matchScoreThreshold > 0) {
        filterApplied = true;
        Logger.log(`Filtering segment by Match Score >= ${matchScoreThreshold}. Initial count: ${candidates.length}. Score Column Index: ${matchStarsColIndex}`);

        filteredCandidates = candidates.filter(row => {
            // Check if row has enough columns and score column index is valid
            if (row.length <= matchStarsColIndex) {
                // Logger.log(`Skipping row due to insufficient columns for score check.`); // Potentially noisy
                return false;
            }
            const scoreValue = row[matchStarsColIndex];
            // Handle potential non-numeric values gracefully
            const matchScore = parseFloat(scoreValue);
            const include = !isNaN(matchScore) && matchScore >= matchScoreThreshold;
            return include;
        });
        Logger.log(`After match score filter, count: ${filteredCandidates.length}`);
    } else if (applyMatchFilter && matchStarsColIndex === -1) {
         Logger.log(`Match score filter requested but no score column found. Filter not applied.`);
    }


    // Initialize metrics object
    const metrics = {
        recruiterData: [],
        titleData: [], // Can add title grouping back if needed
        totalEligibleCandidates: 0,
        totalTakenAI: 0,
        totalNotTakenAIButEligible: 0,
        adoptionRate: 0,
        filterApplied: filterApplied,
        matchScoreThreshold: matchScoreThreshold // Store threshold used
    };

    // Return early if no candidates after initial filtering
    if (!filteredCandidates || filteredCandidates.length === 0) {
        Logger.log("No candidates in this segment after initial score/date filtering. Returning zero metrics.");
        return metrics;
    }

    // NEW LOGIC: Calculate stats based on eligibility
    const recruiterMap = {}; // Moved recruiter grouping here
    let eligibleCandidateCount = 0;
    let eligibleTakenAICount = 0;
    let eligibleNotTakenAICount = 0;

    filteredCandidates.forEach(row => {
        // Check required columns exist for this row
        const aiInterview = row.length > colIndices['Ai_interview'] ? row[colIndices['Ai_interview']] : null;
        const appStatus = row.length > colIndices['Application_status'] ? row[colIndices['Application_status']]?.toLowerCase() : null;
        const recruiter = (row.length > colIndices['Recruiter name'] && row[colIndices['Recruiter name']]) ? row[colIndices['Recruiter name']] : 'Unassigned';

        let isEligible = false;
        let tookAI = false;

        if (aiInterview === 'Y') {
            isEligible = true;
            tookAI = true;
        } else if (aiInterview === 'N') {
            // Only eligible if NOT rejected
            if (appStatus !== 'rejected') { // Assuming 'rejected' is the exact lowercase string
                isEligible = true;
                tookAI = false;
            }
            // If appStatus is 'rejected', they are NOT eligible, so isEligible remains false
        } else {
            // Handle cases where Ai_interview is neither 'Y' nor 'N' (e.g., blank)
            // We'll treat them as eligible IF they are not rejected, similar to 'N'
             if (appStatus !== 'rejected') {
                isEligible = true;
                tookAI = false; // They haven't taken it if it's blank
             }
        }

        // If eligible, count towards totals and recruiter stats
        if (isEligible) {
            eligibleCandidateCount++;
            if (!recruiterMap[recruiter]) {
                recruiterMap[recruiter] = { totalEligible: 0, taken: 0 };
            }
            recruiterMap[recruiter].totalEligible++;

            if (tookAI) {
                eligibleTakenAICount++;
                recruiterMap[recruiter].taken++;
            } else {
                 eligibleNotTakenAICount++; // Count those eligible but not taken
            }
        }
    });

    metrics.totalEligibleCandidates = eligibleCandidateCount;
    metrics.totalTakenAI = eligibleTakenAICount;
    metrics.totalNotTakenAIButEligible = eligibleNotTakenAICount; // Store this for potential future use

    // Calculate adoption rate based on ELIGIBLE candidates
    metrics.adoptionRate = metrics.totalEligibleCandidates > 0 ? parseFloat(((metrics.totalTakenAI / metrics.totalEligibleCandidates) * 100).toFixed(1)) : 0;

     Logger.log(`Segment Metrics Calculation: Total initial filtered = ${filteredCandidates.length}. Total eligible = ${metrics.totalEligibleCandidates}. Total taken AI = ${metrics.totalTakenAI}. Adoption Rate = ${metrics.adoptionRate}%`);

    // Process recruiter data based on eligibility
    metrics.recruiterData = Object.keys(recruiterMap).map(recruiter => {
        const data = recruiterMap[recruiter];
        // Calculate adoption rate PER RECRUITER based on their eligible candidates
        const adoptionRate = data.totalEligible > 0 ? parseFloat(((data.taken / data.totalEligible) * 100).toFixed(1)) : 0;
        return {
            recruiter: recruiter,
            totalCandidates: data.totalEligible, // This now represents ELIGIBLE candidates for the recruiter
            takenAI: data.taken,
            adoptionRate: adoptionRate
        };
    });

    // Sort recruiter data alphabetically for consistent display in charts
    metrics.recruiterData.sort((a, b) => a.recruiter.localeCompare(b.recruiter));

    return metrics;
}


/**
 * Create the HTML email report - % Inline with Name
 */
function createHtmlReport(reportData) {
  // Create HTML content
  let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Recruiter Adoption Report</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 0; background-color: #f4f4f4; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
      table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
      th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 13px; }
      th { background-color: #f2f2f2; font-weight: bold; }
      tr:nth-child(even) { background-color: #f9f9f9; }
      .post-launch { background-color: #e6ffe6; padding: 15px; margin-bottom: 20px; border-radius: 5px; border: 1px solid #c1e5c1;}
      .note { font-style: italic; font-size: 12px; color: #666; margin-top: 8px; margin-bottom: 5px; line-height: 1.3; }
      h1 { color: #333; margin-bottom: 20px; font-size: 22px; text-align: center; }
      h2 { color: #333366; margin-top: 25px; margin-bottom: 15px; font-size: 18px; border-bottom: 2px solid #d9d9e6; padding-bottom: 6px; }
      .highlight { font-weight: bold; color: #b32d00; }
      p { margin: 8px 0; line-height: 1.5; font-size: 14px; }
      ul { margin-top: 5px; padding-left: 25px; }
      li { margin-bottom: 5px; }
      a { color: #007bff; text-decoration: none; }
      a:hover { text-decoration: underline; }
      .stats-block { margin-bottom: 20px; padding: 12px; background-color: #ffffff; border: 1px solid #ddd; border-radius: 4px; }

      /* Bar chart styles - TABLE BASED */
      .bar-chart-table { /* Container table for bars */
        width: 100%;
        margin: 15px 0;
        border: 1px solid #ddd;
        padding: 15px;
        background-color: white;
        border-radius: 5px;
        box-sizing: border-box;
        border-collapse: separate; /* Use separate for spacing */
        border-spacing: 0 8px; /* Vertical spacing between rows */
      }
      .bar-chart-table td { /* Style cells within the chart table */
         border: none; /* Remove default borders */
         padding: 0; /* Reset padding */
         vertical-align: middle; /* Align items vertically */
         height: 24px; /* Match old bar height */
         line-height: 24px; /* Explicit line height can help alignment */
      }
      .bar-label-cell { /* Cell containing the label text + percentage */
        width: 220px; /* Fixed width */
        padding-right: 10px; /* Space between label and graph */
        text-align: left;
        font-weight: bold; /* Keep label part bold */
        font-size: 13px;
        /* white-space: nowrap; REMOVED - Allow wrapping if needed */
        overflow: hidden;
        text-overflow: ellipsis; /* Still useful if wrapping occurs */
        color: #222;
        box-sizing: border-box; /* Include padding in width */
      }
       .bar-graph-cell { /* Cell containing ONLY the graph div */
         width: auto; /* Takes remaining width */
         /* position: relative; No longer needed */
       }

      .bar-graph { /* The div inside the graph cell */
        width: 100%; /* Fill the container cell */
        height: 100%; /* Fill the container cell height */
        background-color: #f0f0f0; /* The background bar color */
        border-radius: 3px;
        position: relative; /* Still needed for fill positioning */
        border: 1px solid #cccccc;
        box-sizing: border-box;
        /* text-align: center; REMOVED */
        /* line-height: 24px; REMOVED */
      }
      .bar-fill {
        height: 100%;
        background-color: #4CAF50; /* Standard Green */
        position: absolute; /* Keep absolute */
        left: 0;
        top: 0;
        border-radius: 2px 0 0 2px; /* Rounded left edge */
        min-width: 1px; /* Show even tiny fills */
        box-sizing: border-box;
      }
      /* No label styles needed here */

      .legend { display: flex; justify-content: center; flex-wrap: wrap; margin-top: 15px; font-size: 12px; }
      .legend-item { display: flex; align-items: center; margin: 2px 10px; }
      .legend-color { width: 14px; height: 14px; margin-right: 5px; border: 1px solid #ccc; flex-shrink: 0; }
      .legend-color.adopted { background-color: #4CAF50; }
      .legend-color.not-adopted { background-color: #f0f0f0; }

      /* Layout styles */
      .report-container { width: 100%; max-width: 900px; margin: 20px auto; padding: 20px; background-color: #ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.1); box-sizing: border-box; border-radius: 4px; }
      .report-section { margin-bottom: 30px; clear: both; }
      .grid-header { font-weight: bold; font-size: 14px; text-align: center; background: #f8f8f8; padding: 10px 8px; margin: -15px -15px 15px -15px; border-bottom: 1px solid #ddd; border-radius: 5px 5px 0 0; color: #333; }

      /* Responsive Column Layout - Using Classes */
      .layout-table { width: 100%; display: table; table-layout: fixed; margin-bottom: 20px; border-spacing: 15px 0; }
      .layout-row { display: table-row; }
      .layout-cell { display: table-cell; width: 50%; vertical-align: top; background-color: #ffffff; border: 1px solid #ddd; border-radius: 5px; padding: 15px; box-sizing: border-box; }
      .layout-fallback { width: 100%; }
      .layout-fallback .layout-cell { display: inline-block; width: 48%; }


      /* Media query for mobile devices */
      @media only screen and (max-width: 600px) {
         body { width: 100% !important; min-width: 100% !important; }
         .report-container { width: 100% !important; padding: 10px !important; margin: 0 auto !important; max-width: 100% !important; box-sizing: border-box !important; }
         h1 { font-size: 18px !important; }
         h2 { font-size: 16px !important; }

         /* Stack main layout table cells */
         .layout-table, .layout-row, .layout-cell { display: block !important; width: 100% !important; max-width: 100% !important; box-sizing: border-box !important; padding-left: 0 !important; padding-right: 0 !important; }
         .layout-table { margin-left: 0 !important; border-spacing: 0 !important; }
         .layout-cell { margin-bottom: 20px !important; padding: 10px !important; }
         .grid-header { margin: -10px -10px 15px -10px !important; }

         /* Adjust chart elements for mobile */
         .bar-chart-table { padding: 10px !important; border-spacing: 0 6px !important; /* Reduce spacing */ }
         .bar-chart-table td { /* All chart cells */
             height: 20px !important; /* Shorter bars */
             line-height: 20px !important; /* Match height */
         }
         .bar-label-cell { /* Label cell */
           width: 150px !important; /* Increased width */
           font-size: 12px !important;
           padding-right: 5px !important; /* Less padding */
           white-space: normal !important; /* Allow wrapping */
           overflow: visible !important; /* Ensure wrapped text shows */
         }
         .bar-graph-cell { /* Graph cell */
             /* Width is auto */
         }
         /* .bar-graph styling remains the same */


         /* Make leaderboard table responsive */
          table[class="leaderboard-table"] tbody,
          table[class="leaderboard-table"] tr,
          table[class="leaderboard-table"] td {
            display: block !important;
            width: 100% !important;
            box-sizing: border-box !important;
         }
         table[class="leaderboard-table"] thead { display: none !important; } /* Hide header row */
         table[class="leaderboard-table"] td {
             text-align: right !important; /* Align text to right */
             padding-left: 50% !important; /* Create space for label */
             position: relative !important;
             border-bottom: 1px solid #eee !important;
             min-height: 30px; /* Ensure cells have some height */
             padding-top: 8px !important;
             padding-bottom: 8px !important;
         }
          table[class="leaderboard-table"] tr:last-child td:last-child { border-bottom: 0 !important; } /* Remove border on very last cell */
         table[class="leaderboard-table"] td::before {
             content: attr(data-label); /* Use data-label for faux header */
             position: absolute !important;
             left: 10px !important; /* More padding */
             width: calc(50% - 20px) !important; /* Adjust width for padding */
             padding-right: 10px !important;
             white-space: nowrap !important;
             text-align: left !important; /* Align label text left */
             font-weight: bold !important;
             color: #555; /* Dim label color */
         }
      }
    </style>
  </head>
  <body>
    <div class="report-container">
      <h1>AI Recruiter Adoption Report</h1>

      <div class="report-section post-launch">
        <h2>Post-Launch Analytics (Since ${reportData.launchDate})</h2>

        <div class="stats-block">
          <p>Adoption Rate (All Post-Launch Applications): <span class="highlight">${reportData.postLaunch.adoptionRate}%</span> (${reportData.postLaunch.totalTakenAI} of ${reportData.postLaunch.totalEligibleCandidates})</p>
          ${reportData.hasMatchStarsColumn ?
             `<p>Adoption Rate (Post-Launch Applications ≥${reportData.postLaunchMatchScore4.matchScoreThreshold || 4} Match): <span class="highlight">${reportData.postLaunchMatchScore4.adoptionRate}%</span> (${reportData.postLaunchMatchScore4.totalTakenAI} of ${reportData.postLaunchMatchScore4.totalEligibleCandidates})</p>`
             : '<p class="note">Match score column not found or filter not applied for second rate.</p>'
          }
          
          ${!reportData.hasMatchStarsColumn ? '<p class="note" style="color: #cc3300;"><b>Note:</b> Match score column could not be identified. The "≥ Match" chart shows ALL candidates.</p>' : ''}
        </div>

        <div class="layout-table"> <div class="layout-row">
            <div class="layout-cell layout-cell-left">
              <div class="grid-header">AI Adoption: All New Applications</div>
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="bar-chart-table">
                <tbody>`;

        reportData.postLaunch.recruiterData.forEach(row => {
          const displayWidth = Math.max(row.adoptionRate, 0);
          const adoptionValue = row.adoptionRate.toFixed(1); // Keep using this for the bar width

          html += `
                  <tr>
                    <td class="bar-label-cell" title="${row.recruiter} (${row.takenAI}/${row.totalCandidates})">
                      ${row.recruiter} (${row.takenAI}/${row.totalCandidates})<span style="color: #CC5500; font-weight: bold;"> [${row.adoptionRate.toFixed(1)}%]</span>
                    </td>
                    <td class="bar-graph-cell">
                      <div class="bar-graph">
                        <div class="bar-fill" style="width: ${displayWidth}%; position: absolute; left: 0; top: 0; height: 100%; background-color: #4CAF50; border-radius: 2px 0 0 2px; min-width: 1px; box-sizing: border-box;" data-value="${adoptionValue}">
                          </div>
                        </div>
                    </td>
                  </tr>`;
        });

        html += `
                </tbody>
              </table>
              <div class="legend">
                <div class="legend-item"><div class="legend-color adopted"></div><div>Invited</div></div>
                <div class="legend-item"><div class="legend-color not-adopted"></div><div>Not Invited Yet</div></div>
              </div>
            </div>

            <div class="layout-cell layout-cell-right">
              <div class="grid-header">AI Adoption: New Applications (≥${reportData.postLaunchMatchScore4.matchScoreThreshold || 4} Match)</div>
              ${!reportData.hasMatchStarsColumn ?
                '<div style="padding: 10px; color: #cc3300; text-align: center; font-size: 0.9em;"><b>Warning:</b> Match score column not found. Showing all candidates.</div>' : ''}
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="bar-chart-table">
                 <tbody>`;

        reportData.postLaunchMatchScore4.recruiterData.forEach(row => {
           const displayWidth = Math.max(row.adoptionRate, 0);
           const adoptionValue = row.adoptionRate.toFixed(1); // Keep using this for the bar width

           html += `
                  <tr>
                    <td class="bar-label-cell" title="${row.recruiter} (${row.takenAI}/${row.totalCandidates})">
                       ${row.recruiter} (${row.takenAI}/${row.totalCandidates})<span style="color: #CC5500; font-weight: bold;"> [${row.adoptionRate.toFixed(1)}%]</span>
                    </td>
                    <td class="bar-graph-cell">
                       <div class="bar-graph">
                        <div class="bar-fill" style="width: ${displayWidth}%; position: absolute; left: 0; top: 0; height: 100%; background-color: #4CAF50; border-radius: 2px 0 0 2px; min-width: 1px; box-sizing: border-box;" data-value="${adoptionValue}">
                           </div>
                         </div>
                    </td>
                  </tr>`;
        });

        html += `
                </tbody>
              </table>
              <div class="legend">
                 <div class="legend-item"><div class="legend-color adopted"></div><div>Invited</div></div>
                 <div class="legend-item"><div class="legend-color not-adopted"></div><div>Not Invited Yet</div></div>
              </div>
            </div>
          </div>
        </div> </div> <div class="report-section" style="padding: 15px; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px; margin-top: 20px;">
            <h4 style="margin-top: 0; margin-bottom: 10px; color: #333; font-size: 14px;">How Adoption Rate is Calculated:</h4>
            <p class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 0;">This report focuses on AI Interview adoption among candidates considered potentially eligible. The adoption rate is calculated as:</p>
            <p class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 0; padding-left: 15px;"><strong>(Candidates Who Took AI Interview) / (Eligible Candidates)</strong></p>
            <p class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 0;"><strong>Eligible Candidates</strong> include:</p>
            <ul class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 15px; padding-left: 15px; list-style-type: disc;">
                <li style="margin-bottom: 3px;">Everyone who took the AI Interview ('Y').</li>
                <li>Everyone who did <em>not</em> take the AI Interview ('N' or blank) AND whose application status is NOT 'Rejected'.</li>
            </ul>
            <p class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 0;"><strong>Excluded:</strong> Candidates marked as 'Rejected' who did <em>not</em> take the AI interview are removed from the 'Eligible Candidates' pool. This exclusion assumes the recruiter reviewed these applications and decided they were not suitable matches for AI screening, so they are not counted when measuring adoption among the considered pool.</p>
            <p class="note" style="font-size: 12.5px; color: #555; margin-top: 5px;"><em>Example:</em> A recruiter has 12 applications:</p>
             <ul class="note" style="font-size: 12.5px; color: #555; margin: 0 0 8px 15px; padding-left: 15px; list-style-type: disc;">
                 <li style="margin-bottom: 3px;">4 do not match the job at all so much so that an AI Screening is not warranted at all. These applications have been rejected by the Recruiter. <span style="color: #990000;">&rarr; These 4 are excluded from the analysis.</span></li>
                 <li style="margin-bottom: 3px;">3 took the AI Screening.</li>
                 <li>5 are 'Active' without an AI Screening and can be given an AI screening. <span style="color: #006400;">&rarr; Included as Eligible</span></li>
             </ul>
            <p class="note" style="font-size: 12.5px; color: #555; margin: 0;">The calculation uses these 8 Eligible Candidates (3 who took screening + 5 'Active' without screening). The adoption rate is 3 / 8 = 37.5%.</p>
        </div>

        <div class="report-section" style="text-align: center; margin: 25px 0; padding: 20px; background-color: #e9f5ff; border-radius: 5px; border: 1px solid #b3daff;">
        <p style="font-size: 16px; font-weight: bold; margin-bottom: 10px; color: #004085;">Find Strong-Match Applicants to Invite</p>
        <p style="margin-bottom: 15px; color: #0056b3; font-size: 13px;">Click below for a list of recent, strong-match applicants (≥${reportData.postLaunchMatchScore4.matchScoreThreshold || 4} stars) who haven't received an AI screening invitation yet.</p>
        <a href="${LOOKER_STUDIO_URL}" target="_blank" style="display: inline-block; padding: 12px 25px; background-color: #007bff; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px; border-bottom: 3px solid #0056b3; transition: background-color 0.2s ease;">View Candidates in Looker Studio</a>
      </div>

      <div class="report-section">
        <h2 style="text-align: center;">AI Interview Leaderboard (All Time)</h2>
        <p style="text-align: center; font-size: 13px; margin-top: -10px; margin-bottom: 20px; color: #555;">Total number of AI interviews initiated per recruiter (Active Pre-Launch + All Post-Launch)</p>
        <div style="max-width: 500px; margin: 0 auto;">
          <table class="leaderboard-table" style="width: 100%; font-size: 13px; border: 1px solid #ddd; border-radius: 4px; overflow: hidden;">
            <thead style="border-bottom: 2px solid #ccc;">
              <tr>
                <th style="width: 50px; text-align: center;">Rank</th>
                <th>Recruiter</th>
                <th style="width: 100px; text-align: center;">AI Interviews Initiated</th>
              </tr>
            </thead>
            <tbody>`;

        // Combine all candidate data (both pre and post launch) for leaderboard
        const allRecruiters = {};
        reportData.preLaunch.recruiterData.forEach(row => {
          if (!allRecruiters[row.recruiter]) { allRecruiters[row.recruiter] = 0; }
           allRecruiters[row.recruiter] += row.takenAI;
        });
        reportData.postLaunch.recruiterData.forEach(row => {
          if (!allRecruiters[row.recruiter]) { allRecruiters[row.recruiter] = 0; }
           allRecruiters[row.recruiter] += row.takenAI;
        });
        const leaderboardData = Object.keys(allRecruiters).map(recruiter => ({
            recruiter: recruiter,
            interviews: allRecruiters[recruiter] || 0
        })).sort((a, b) => b.interviews - a.interviews);
        const topRecruiters = leaderboardData.filter(entry => entry.interviews > 0);

        if (topRecruiters.length === 0) {
             html += `<tr><td colspan="3" data-label="Status" style="text-align: center; padding: 15px; color: #777;">No AI interviews initiated yet.</td></tr>`;
        } else {
            topRecruiters.forEach((entry, index) => {
              let rankDisplay = `${index + 1}`;
              if (index === 0) rankDisplay = `🥇`; else if (index === 1) rankDisplay = `🥈`; else if (index === 2) rankDisplay = `🥉`;
              html += `
                  <tr style="height: 32px;">
                    <td data-label="Rank" style="text-align: center; font-size: ${index < 3 ? '1.3em' : '1.1em'}; font-weight: ${index < 3 ? 'bold' : 'normal'};">${rankDisplay}</td>
                    <td data-label="Recruiter" style="font-weight: bold;">${entry.recruiter}</td>
                    <td data-label="Interviews" style="text-align: center; font-weight: bold; font-size: 1.1em;">${entry.interviews}</td>
                  </tr>`;
            });
        }
        html += `
            </tbody>
          </table>
        </div>
      </div>

       <div class="report-section" style="border-top: 1px solid #eee; padding-top: 20px; margin-top: 30px;">
        <h3 style="color: #333; margin-bottom: 10px; font-size: 16px;">Key Focus Areas:</h3>
        <ul style="margin-top: 5px; padding-left: 25px; font-size: 14px; list-style-type: disc;">
          <li>Monitor the <span style="background-color: #e6ffe6; padding: 2px 5px; border-radius: 3px; border: 1px solid #b3ffb3;">Post-Launch Adoption Rate</span> as your primary KPI for new applications.</li>
           ${reportData.hasMatchStarsColumn ?
             `<li>Prioritize inviting candidates with <span style="background-color: #e6ffe6; padding: 2px 5px; border-radius: 3px; border: 1px solid #b3ffb3;">Match Scores ≥ ${reportData.postLaunchMatchScore4.matchScoreThreshold || 4}</span> via the Looker Studio link above.</li>`
             : `<li>Prioritize inviting high-match candidates using the Looker Studio link above (score column not identified for filtering).</li>`
           }
          <li>Review the <span style="background-color: #f8f9fa; padding: 2px 5px; border-radius: 3px; border: 1px solid #eee;">Leaderboard</span> to celebrate successes and identify coaching opportunities based on initiated interviews.</li>
        </ul>
        <p class="note" style="margin-top: 20px; text-align: center; color: #888;">Report generated on ${new Date().toLocaleString()}. Timezone: ${Session.getScriptTimeZone()}.</p>
      </div>

    </div>
  </body>
  </html>`;

  return html;
}


/**
 * Send an email with the report
 */
function sendEmail(recipient, subject, htmlBody) {
  // Basic validation
  if (!recipient) {
     Logger.log("ERROR: Email recipient is empty. Cannot send email.");
     return;
  }
   if (!subject) {
     Logger.log("WARNING: Email subject is empty. Using default subject.");
     subject = "AI Recruiter Report";
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

  // Add CC only if it's defined and different from the recipient
  if (EMAIL_CC && EMAIL_CC.trim() !== '' && EMAIL_CC.trim().toLowerCase() !== recipient.trim().toLowerCase()) {
    options.cc = EMAIL_CC;
    Logger.log(`Sending email to ${recipient}, CC ${EMAIL_CC}`);
  } else {
     Logger.log(`Sending email to ${recipient} (No CC or CC is same as recipient)`);
  }

  try {
      MailApp.sendEmail(options);
      Logger.log("Email sent successfully.");
  } catch (e) {
     Logger.log(`ERROR sending email: ${e.toString()}`);
     // Optional: re-throw or handle error further
     // throw e; // Uncomment to make the main function catch this failure
  }

}

/**
 * Debug function to log sheet headers and structure
 */
function logSheetHeaders() {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    Logger.log(`Opened spreadsheet: ${spreadsheet.getName()}`);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
        Logger.log(`Sheet "${SHEET_NAME}" not found by name. Trying by GID or first sheet.`);
        const gidMatch = SPREADSHEET_URL.match(/gid=(\d+)/);
        if (gidMatch && gidMatch[1]) {
            const gid = gidMatch[1];
            const sheets = spreadsheet.getSheets();
            sheet = sheets.find(s => s.getSheetId().toString() === gid);
        }
        if (!sheet) {
            sheet = spreadsheet.getSheets()[0];
        }
    }

    if (!sheet) {
        Logger.log("Could not find any sheet to analyze.");
        SpreadsheetApp.getUi().alert("Error: Could not find the sheet to analyze.");
        return "Error: Could not find sheet.";
    }

    Logger.log(`Analyzing sheet: "${sheet.getName()}" (ID: ${sheet.getSheetId()})`);

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const numRows = data.length;
    const numCols = numRows > 0 ? data[0].length : 0;
    Logger.log(`Sheet dimensions: ${numRows} rows x ${numCols} columns`);

    if (numRows === 0) {
       Logger.log("Sheet is empty.");
       return "Sheet is empty.";
    }

    // Log the first few rows (max 5 data rows + header rows)
    Logger.log("--- First few rows ---");
    for (let i = 0; i < Math.min(7, numRows); i++) {
        // Truncate long cell values for cleaner logs
        const rowData = data[i].map(cell => {
            const cellStr = String(cell);
            return cellStr.length > 100 ? cellStr.substring(0, 97) + '...' : cellStr;
        });
        Logger.log(`Row ${i + 1}: ${JSON.stringify(rowData)}`);
    }
    Logger.log("--- End of first few rows ---");


    // Log the headers specifically from row 2 (index 1)
    if (numRows > 1) {
        const headers = data[1].map(String); // Ensure headers are strings
        Logger.log(`Headers identified in Row 2: ${JSON.stringify(headers)}`);

        // Check for required column headers
        const requiredColumns = ['Last_stage', 'Ai_interview', 'Recruiter name', 'Title', 'Application_status', 'Position_status', 'Application_ts'];
        const missingColumns = [];
        const foundColumns = {};

        requiredColumns.forEach(colName => {
            const index = headers.indexOf(colName);
            if (index === -1) {
                missingColumns.push(colName);
            } else {
                foundColumns[colName] = index;
            }
        });

        if (missingColumns.length > 0) {
            Logger.log(`WARNING: Missing required columns: ${missingColumns.join(', ')}`);
        } else {
            Logger.log(`All required columns found at indices: ${JSON.stringify(foundColumns)}`);
        }

        // Check a sample of Application_ts values (if column exists)
        const appTsIndex = headers.indexOf('Application_ts');
        if (appTsIndex !== -1) {
            Logger.log(`--- Checking 'Application_ts' date format (Column index ${appTsIndex}) ---`);
            const sampleSize = Math.min(5, numRows - 2); // Max 5 data rows
            let validDates = 0;
            let invalidDates = 0;
            for (let i = 0; i < sampleSize; i++) {
                const rowIndex = i + 2; // Data starts from row 3 (index 2)
                if (rowIndex < numRows) { // Check if row exists
                    const row = data[rowIndex];
                    if (row && row.length > appTsIndex) { // Check if cell exists
                        const rawDate = row[appTsIndex];
                        const parsedDate = new Date(rawDate);
                        const isValid = !isNaN(parsedDate.getTime()) && rawDate !== null && rawDate !== ''; // Check validity more strictly

                        if(isValid) validDates++; else invalidDates++;

                        Logger.log(`Row ${rowIndex + 1}: Raw='${rawDate}', Parsed='${isValid ? parsedDate.toISOString() : 'Invalid'}', Valid=${isValid}`);
                    } else {
                        Logger.log(`Row ${rowIndex + 1}: Cell for Application_ts (Index ${appTsIndex}) does not exist.`);
                         invalidDates++;
                    }
                }
            }
             Logger.log(`Date check summary: ${validDates} valid, ${invalidDates} invalid/empty in sample.`);
             Logger.log(`--- End of date check ---`);
        } else {
             Logger.log(`Column 'Application_ts' not found in headers.`);
        }

    } else {
        Logger.log('Sheet has less than 2 rows, cannot identify headers in Row 2.');
    }

    SpreadsheetApp.getUi().alert("Header logging complete. Check Execution logs (View > Logs) for details.");
    return "Log output complete.";
  } catch (error) {
    Logger.log(`Error in logSheetHeaders: ${error.toString()} Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error logging headers: ${error.message}. Check logs.`);
    return `Error: ${error.toString()}`;
  }
}


/**
 * Debug function specifically for investigating a specific recruiter's candidates
 */
function debugAkhilaCandidates() { // Consider renaming or making recruiter dynamic
  const recruiterToDebug = 'Akhila Kashyap'; // Define recruiter name here
  Logger.log(`--- Starting Debug for Recruiter: ${recruiterToDebug} ---`);
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    Logger.log(`Opened spreadsheet: ${spreadsheet.getName()}`);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

     if (!sheet) {
        Logger.log(`Sheet "${SHEET_NAME}" not found by name. Trying by GID or first sheet.`);
        const gidMatch = SPREADSHEET_URL.match(/gid=(\d+)/);
        if (gidMatch && gidMatch[1]) {
            const gid = gidMatch[1];
            const sheets = spreadsheet.getSheets();
            sheet = sheets.find(s => s.getSheetId().toString() === gid);
        }
        if (!sheet) {
            sheet = spreadsheet.getSheets()[0];
        }
    }

    if (!sheet) {
        Logger.log("Could not find any sheet to analyze for debug.");
        return "Error: Could not find sheet.";
    }
    Logger.log(`Using sheet: "${sheet.getName()}" for debug`);

    // Get data
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
     if (data.length < 3) {
       Logger.log("Not enough data rows for headers and data.");
       return "Error: Not enough data rows."
     }
    const headers = data[1].map(String); // Headers in row 2
    const rows = data.slice(2); // Data rows

    // Find column indices - reuse logic from main function
    const requiredColsForDebug = ['Recruiter name', 'Ai_interview', 'Application_ts', 'Application_status', 'Position_status', 'Last_stage', 'Title'];
    const missingDebugCols = [];
    const colIndicesDebug = {};
    requiredColsForDebug.forEach(colName => {
       const index = headers.indexOf(colName);
       if (index === -1) {
           missingDebugCols.push(colName);
       } else {
           colIndicesDebug[colName] = index;
       }
    });

     if (missingDebugCols.length > 0) {
        Logger.log(`ERROR: Missing columns needed for debugging: ${missingDebugCols.join(', ')}`);
        return `Error: Missing debug columns: ${missingDebugCols.join(', ')}`;
     }
      Logger.log(`Debug Column Indices: ${JSON.stringify(colIndicesDebug)}`);


    // Find all candidates for the specified recruiter
    const allRecruiterCandidates = rows.filter(row => row.length > colIndicesDebug['Recruiter name'] && row[colIndicesDebug['Recruiter name']] === recruiterToDebug);
    Logger.log(`Found ${allRecruiterCandidates.length} total rows for ${recruiterToDebug}`);

    // Check how many took AI interview
    const allRecruiterTakenAI = allRecruiterCandidates.filter(row => row.length > colIndicesDebug['Ai_interview'] && row[colIndicesDebug['Ai_interview']] === 'Y');
    Logger.log(`Total ${recruiterToDebug} candidates who took AI interview (Y): ${allRecruiterTakenAI.length}`);

    // Split into pre-launch and post-launch (all statuses)
    const prelaunchAll = [];
    const postlaunchAll = [];
    let debugInvalidDateCount = 0;

    allRecruiterCandidates.forEach((candidate, idx) => {
        if (candidate.length <= colIndicesDebug['Application_ts']) return; // Skip row if too short
        const rawDate = candidate[colIndicesDebug['Application_ts']];
         if (rawDate === null || rawDate === undefined || rawDate === '') { debugInvalidDateCount++; return; } // Skip invalid/empty
        const applicationDate = new Date(rawDate);

        if (!isNaN(applicationDate.getTime())) {
            if (applicationDate >= LAUNCH_DATE) {
                postlaunchAll.push(candidate);
            } else {
                prelaunchAll.push(candidate);
            }
        } else {
           debugInvalidDateCount++;
        }
    });
     if(debugInvalidDateCount > 0) Logger.log(`Debug: ${debugInvalidDateCount} rows skipped due to invalid/empty dates for ${recruiterToDebug}.`);

    Logger.log(`Total pre-launch candidates (all statuses, valid date): ${prelaunchAll.length}`);
    Logger.log(`Total post-launch candidates (all statuses, valid date): ${postlaunchAll.length}`);

    // Check AI interview adoption for both segments (all statuses)
    const prelaunchAllTakenAI = prelaunchAll.filter(row => row.length > colIndicesDebug['Ai_interview'] && row[colIndicesDebug['Ai_interview']] === 'Y');
    const postlaunchAllTakenAI = postlaunchAll.filter(row => row.length > colIndicesDebug['Ai_interview'] && row[colIndicesDebug['Ai_interview']] === 'Y');

    Logger.log(`Pre-launch took AI (all statuses): ${prelaunchAllTakenAI.length} (${prelaunchAll.length > 0 ? Math.round(prelaunchAllTakenAI.length/prelaunchAll.length*100): 0}%)`);
    Logger.log(`Post-launch took AI (all statuses): ${postlaunchAllTakenAI.length} (${postlaunchAll.length > 0 ? Math.round(postlaunchAllTakenAI.length/postlaunchAll.length*100): 0}%)`);

    // Now filter for active+open
    const activeOpenRecruiterCandidates = allRecruiterCandidates.filter(row =>
       (row.length > colIndicesDebug['Application_status'] && row[colIndicesDebug['Application_status']]?.toLowerCase() === 'active') &&
       (row.length > colIndicesDebug['Position_status'] && row[colIndicesDebug['Position_status']]?.toLowerCase() === 'open')
    );
    Logger.log(`Active+Open ${recruiterToDebug} candidates: ${activeOpenRecruiterCandidates.length}`);

    // Split active+open into pre and post launch
    const prelaunchActiveOpen = [];
    const postlaunchActiveOpen = [];
     let debugInvalidDateCountAO = 0;

    activeOpenRecruiterCandidates.forEach(candidate => {
        if (candidate.length <= colIndicesDebug['Application_ts']) return;
        const rawDate = candidate[colIndicesDebug['Application_ts']];
         if (rawDate === null || rawDate === undefined || rawDate === '') { debugInvalidDateCountAO++; return; }
        const applicationDate = new Date(rawDate);

        if (!isNaN(applicationDate.getTime())) {
            if (applicationDate >= LAUNCH_DATE) {
                postlaunchActiveOpen.push(candidate);
            } else {
                prelaunchActiveOpen.push(candidate);
            }
        } else {
            debugInvalidDateCountAO++;
        }
    });
    if(debugInvalidDateCountAO > 0) Logger.log(`Debug: ${debugInvalidDateCountAO} active/open rows skipped due to invalid/empty dates.`);


    Logger.log(`Pre-launch active+open candidates: ${prelaunchActiveOpen.length}`);
    Logger.log(`Post-launch active+open candidates: ${postlaunchActiveOpen.length}`);

    // Check AI interview adoption for active+open
    const prelaunchActiveOpenTakenAI = prelaunchActiveOpen.filter(row => row.length > colIndicesDebug['Ai_interview'] && row[colIndicesDebug['Ai_interview']] === 'Y');
    const postlaunchActiveOpenTakenAI = postlaunchActiveOpen.filter(row => row.length > colIndicesDebug['Ai_interview'] && row[colIndicesDebug['Ai_interview']] === 'Y');

    Logger.log(`Pre-launch active+open who took AI: ${prelaunchActiveOpenTakenAI.length} (${prelaunchActiveOpen.length > 0 ? Math.round(prelaunchActiveOpenTakenAI.length/prelaunchActiveOpen.length*100) : 0}%)`);
    Logger.log(`Post-launch active+open who took AI: ${postlaunchActiveOpenTakenAI.length} (${postlaunchActiveOpen.length > 0 ? Math.round(postlaunchActiveOpenTakenAI.length/postlaunchActiveOpen.length*100) : 0}%)`);

    // Detailed analysis of post-launch candidates who took AI (any status)
    Logger.log(`\n--- DETAILED ANALYSIS OF POST-LAUNCH ${recruiterToDebug} CANDIDATES WHO TOOK AI INTERVIEW (ANY STATUS) ---`);
    if (postlaunchAllTakenAI.length === 0) {
        Logger.log("None found.");
    } else {
        postlaunchAllTakenAI.forEach((candidate, index) => {
           const appDate = new Date(candidate[colIndicesDebug['Application_ts']]);
           const appStatus = candidate.length > colIndicesDebug['Application_status'] ? (candidate[colIndicesDebug['Application_status']]?.toLowerCase() || 'N/A') : 'N/A';
           const posStatus = candidate.length > colIndicesDebug['Position_status'] ? (candidate[colIndicesDebug['Position_status']]?.toLowerCase() || 'N/A') : 'N/A';

           Logger.log(`Candidate ${index + 1}:`);
           Logger.log(`  App Date: ${!isNaN(appDate.getTime()) ? appDate.toISOString().split('T')[0] : 'Invalid'}`);
           Logger.log(`  Title: ${candidate.length > colIndicesDebug['Title'] ? (candidate[colIndicesDebug['Title']] || 'N/A') : 'N/A'}`);
           Logger.log(`  Last Stage: ${candidate.length > colIndicesDebug['Last_stage'] ? (candidate[colIndicesDebug['Last_stage']] || 'N/A') : 'N/A'}`);
           Logger.log(`  App Status: ${appStatus}`);
           Logger.log(`  Pos Status: ${posStatus}`);
           Logger.log(`  Included in POST-LAUNCH report section: YES (all statuses included)`);
        });
    }
     Logger.log(`--- END OF DETAILED ANALYSIS ---`);


    // Create summary reflecting the *NEW* report logic
    const launchDateStr = LAUNCH_DATE.toLocaleDateString();
    const summary = `
      SUMMARY FOR ${recruiterToDebug} (Reflecting Report Logic - Generated ${new Date().toLocaleDateString()}):

      Total candidates assigned: ${allRecruiterCandidates.length}
      Total who ever took AI interview: ${allRecruiterTakenAI.length}

      PRE-LAUNCH PERIOD (Apps before ${launchDateStr}):
      Considered in report (Active+Open only, valid date): ${prelaunchActiveOpen.length} candidates
      Took AI (Active+Open only): ${prelaunchActiveOpenTakenAI.length} candidates (${prelaunchActiveOpen.length > 0 ? Math.round(prelaunchActiveOpenTakenAI.length/prelaunchActiveOpen.length*100) : 0}%)

      POST-LAUNCH PERIOD (Apps on/after ${launchDateStr}):
      Considered in report (All statuses, valid date): ${postlaunchAll.length} candidates
      Took AI (All statuses): ${postlaunchAllTakenAI.length} candidates (${postlaunchAll.length > 0 ? Math.round(postlaunchAllTakenAI.length/postlaunchAll.length*100) : 0}%)

      LEADERBOARD COUNT (All time, Taken AI): ${prelaunchActiveOpenTakenAI.length} (from pre-launch active/open) + ${postlaunchAllTakenAI.length} (from post-launch all) = ${prelaunchActiveOpenTakenAI.length + postlaunchAllTakenAI.length}
    `;

    Logger.log("\n" + summary);
    Logger.log(`--- End Debug for Recruiter: ${recruiterToDebug} ---`);
    SpreadsheetApp.getUi().alert(`Debug for ${recruiterToDebug} complete. Check Execution Logs.`);
    return "Debug complete.";
  } catch (error) {
    Logger.log(`Error in debugAkhilaCandidates: ${error.toString()} Stack: ${error.stack}`);
     SpreadsheetApp.getUi().alert(`Error during debug: ${error.message}. Check logs.`);
    return `Error: ${error.toString()}`;
  }
}


/**
 * Creates a menu item in the Google Sheet UI
 */
function onOpen() {
  try {
      SpreadsheetApp.getUi()
          .createMenu('AI Interview Report')
          .addItem('Generate & Send Report Now', 'generateAndSendAIReport')
          .addItem('Schedule Daily Report (8 AM)', 'createDailyTrigger')
          .addSeparator()
          .addItem('Log Sheet Headers (Debug)', 'logSheetHeaders')
          .addItem('Debug Recruiter Data (Debug)', 'debugAkhilaCandidates') // Made name more generic
          .addToUi();
  } catch (e) {
      Logger.log("Error creating menu: " + e);
      // Don't throw error here as it might prevent sheet from opening properly
  }
}