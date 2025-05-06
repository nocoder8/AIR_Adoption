/**
 * Generates a daily report on AI interview status for candidates
 * based on specified conditions and emails it.
 * It also writes the detailed list of candidates meeting the criteria to a separate sheet.
 */
function sendDailyAIMissedOppsReport() {
  const SPREADSHEET_ID = '1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw'; // Source Spreadsheet ID
  const SHEET_NAME = 'Active+Rejected'; // Source Sheet Name
  const HEADER_ROW = 2; // Headers are on the 2nd row in the source sheet
  const EMAIL_RECIPIENT = 'pkumar@eightfold.ai,akashyap@eightfold.ai'; // Updated to include multiple recipients
  const THRESHOLD_DATE_STR = '2025-04-25'; // YYYY-MM-DD format

  // --- Destination Sheet Configuration ---
  const DESTINATION_SPREADSHEET_ID = '1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE';
  const DESTINATION_SHEET_NAME = 'Missed Opportunities';

  // --- Configuration End ---

  // --- Helper Function for Date Formatting ---
  function formatDateForDisplay(dateInput) {
    if (!dateInput) return ''; // Handle empty input
    try {
      let dateObj;
      if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
        dateObj = dateInput; // Already a valid Date object
      } else {
        dateObj = new Date(dateInput); // Attempt to parse string/number
      }

      if (isNaN(dateObj.getTime())) {
         // Don't log a warning here as it might be noisy for empty/non-date cells
         return dateInput.toString(); // Return original string if parsing fails
      }
      // Format using GMT to avoid potential timezone shifts changing the date
      return Utilities.formatDate(dateObj, "GMT", 'dd-MMM-yyyy');
    } catch (e) {
      Logger.log(`Error formatting date '${dateInput}': ${e}`);
      // Return original string representation on error
      return (dateInput && typeof dateInput.toString === 'function') ? dateInput.toString() : '';
    }
  }
  // --- End Helper Function ---

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // Use openById for robustness
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Error: Sheet "${SHEET_NAME}" not found.`);
    return;
  }

  // Get data starting from header row
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();

  if (allData.length <= HEADER_ROW) {
    Logger.log('Error: No data found below the header row.');
    return; // No data to process
  }

  // Get headers from the specified row (adjusting to 0-based index)
  const headers = allData[HEADER_ROW - 1];
  const data = allData.slice(HEADER_ROW); // Get data rows below headers

  // Find column indices dynamically
  const colIndices = {};
  const requiredCols = [
    'Application_ts', 'Last_stage', 'Ai_interview', 'Recruiter name',
    // Add columns for the detailed candidate list
    'Name', 'Current_company', 'Profile_link', 'Title', 'Hiring manager name', 'Last_stage_ts'
  ];
  headers.forEach((header, index) => {
    // Trim header to handle potential leading/trailing spaces
    const trimmedHeader = header.trim();
    if (requiredCols.includes(trimmedHeader)) {
      // Store the index using the original requiredCol name for consistency
      const originalHeader = requiredCols.find(h => h === trimmedHeader);
       if (originalHeader) {
         colIndices[originalHeader] = index;
       }
    }
  });

  // Check if all required columns were found
  let missingCols = [];
  for (const col of requiredCols) {
    if (colIndices[col] === undefined) {
      missingCols.push(col);
    }
  }

  if (missingCols.length > 0) {
      const errorMessage = `Error: Required column(s) "${missingCols.join(', ')}" not found in header row ${HEADER_ROW} of sheet "${SHEET_NAME}". Please check the sheet configuration.`;
      Logger.log(errorMessage);
      MailApp.sendEmail(EMAIL_RECIPIENT,
                        `Error in Daily AI Report Script`,
                        errorMessage);
      return;
  }


  const thresholdDate = new Date(THRESHOLD_DATE_STR);
  // Add time component to ensure comparison includes the whole day
  thresholdDate.setHours(0, 0, 0, 0);


  let totalRelevantApps = 0;
  let aiInterviewNotTakenCount = 0;
  const recruiterStats = {}; // { recruiterName: { total: 0, aiNotTaken: 0 } }
  const lostCandidatesDetails = []; // Array to store details of candidates with 'N' status

  data.forEach((row, rowIndex) => {
    // Helper function to safely get cell value or return empty string
    const getCell = (colName) => (row[colIndices[colName]] !== undefined && row[colIndices[colName]] !== null) ? row[colIndices[colName]].toString().trim() : '';

    // Check if essential columns have data to avoid errors
    const appTsValue = getCell('Application_ts');
    const lastStageValue = getCell('Last_stage');
    const aiInterviewValue = getCell('Ai_interview');
    const recruiterNameValue = getCell('Recruiter name'); // Get recruiter name early

    // Basic check for non-empty essential values for filtering logic
    if (!appTsValue || !lastStageValue || !aiInterviewValue) {
        // Optionally log skipped rows: Logger.log(`Skipping row ${rowIndex + HEADER_ROW + 1} due to missing essential data for filtering.`);
        return; // Skip row if essential data for filtering is missing
    }

    let applicationDate;
    try {
      applicationDate = new Date(appTsValue);
      if (isNaN(applicationDate.getTime())) {
        throw new Error("Invalid date format");
      }
      applicationDate.setHours(0, 0, 0, 0); // Normalize date for comparison
    } catch (e) {
      Logger.log(`Warning: Could not parse Application_ts "${appTsValue}" on row ${rowIndex + HEADER_ROW + 1}. Skipping row. Error: ${e.message}`);
      return; // Skip row if date is invalid
    }

    const lastStage = lastStageValue; // Already trimmed by getCell
    const aiInterview = aiInterviewValue.toUpperCase(); // Normalize to uppercase 'N'
    const recruiterName = recruiterNameValue ? recruiterNameValue : 'Unknown Recruiter'; // Handle blank recruiter names

    // Define stages to exclude
    const excludedStages = ['New', 'HM Quick Feedback', 'Added', 'Recruiter Quick Feedback'];

    // Apply conditions
    // Exclude rows if Last_stage is in the excludedStages list OR if applicationDate is not after thresholdDate
    if (!excludedStages.includes(lastStage) && applicationDate > thresholdDate) { // Using > as originally intended, change to >= if needed
      totalRelevantApps++;

      // Initialize recruiter stats if not present
      if (!recruiterStats[recruiterName]) {
        recruiterStats[recruiterName] = { total: 0, aiNotTaken: 0 };
      }
      recruiterStats[recruiterName].total++;

      if (aiInterview === 'N') {
        aiInterviewNotTakenCount++;
        recruiterStats[recruiterName].aiNotTaken++;

        // Collect details for the candidate list table
         const candidateData = {
            recruiterName: recruiterName,
            name: getCell('Name'),
            currentCompany: getCell('Current_company'),
            profileLink: getCell('Profile_link'),
            title: getCell('Title'),
            hiringManager: getCell('Hiring manager name'),
            aiInterviewStatus: aiInterview,
            lastStage: lastStage,
            // Format the timestamps for display
            lastStageTsFormatted: formatDateForDisplay(getCell('Last_stage_ts')),
            applicationTsFormatted: formatDateForDisplay(applicationDate) // Pass the Date object
         };
         lostCandidatesDetails.push(candidateData);
      }
    }
  });

  // Sort the candidate details by Recruiter Name
  lostCandidatesDetails.sort((a, b) => a.recruiterName.localeCompare(b.recruiterName));

  // --- Write Detailed List to Destination Sheet ---
  let destinationSheetUrl = '#'; // Default URL in case of error
  try {
    const destSS = SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);
    let destSheet = destSS.getSheetByName(DESTINATION_SHEET_NAME);
    if (!destSheet) {
      destSheet = destSS.insertSheet(DESTINATION_SHEET_NAME);
      Logger.log(`Created destination sheet: ${DESTINATION_SHEET_NAME}`);
    }

    // Define headers for the destination sheet
    const destHeaders = [
        'Recruiter Name', 'Name', 'Current Company', 'Profile Link', 'Title',
        'Hiring Manager', 'AI Interview', 'Last Stage', 'Last Stage Ts', 'Application Ts'
    ];

    // Prepare data for the destination sheet
    const outputData = lostCandidatesDetails.map(candidate => [
        candidate.recruiterName,
        candidate.name || '', // Use empty string if null/undefined
        candidate.currentCompany || '',
        candidate.profileLink || '', // Output the URL directly
        candidate.title || '',
        candidate.hiringManager || '',
        candidate.aiInterviewStatus,
        candidate.lastStage,
        candidate.lastStageTsFormatted,    // Use formatted date
        candidate.applicationTsFormatted   // Use formatted date
    ]);

    // Clear existing data (keeping headers if sheet existed) and write new data
    destSheet.clearContents(); // Clear everything first
    destSheet.appendRow(destHeaders); // Write headers
    if (outputData.length > 0) {
      destSheet.getRange(2, 1, outputData.length, destHeaders.length).setValues(outputData);
      Logger.log(`Wrote ${outputData.length} rows to sheet: ${DESTINATION_SHEET_NAME}`);
    }
    // Auto-resize columns for better readability
     destSheet.autoResizeColumns(1, destHeaders.length);

     destinationSheetUrl = destSS.getUrl() + '#gid=' + destSheet.getSheetId();


  } catch (e) {
    Logger.log(`Error writing to destination sheet ${DESTINATION_SHEET_NAME}: ${e}`);
    // Optionally send an email notification about this specific failure
    MailApp.sendEmail(EMAIL_RECIPIENT,
                      `FAILED: Writing to Destination Sheet in AI Report Script`,
                      `The script failed to write data to sheet "${DESTINATION_SHEET_NAME}" in spreadsheet ID ${DESTINATION_SPREADSHEET_ID}. Error: ${e.message}. Check Logs.`);
    // Continue with email report generation, but mention the failure.
    emailBody += `<p style="color: red;"><b>Error:</b> Failed to write detailed candidate list to the Google Sheet. Please check logs.</p>`;

  }

  // --- Generate Report Email Body ---
  const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let emailSubject = `AI Screening: Missed Opportunities`;

  // Basic CSS Styling
  const styles = `
    <style>
      body {
        font-family: Arial, Helvetica, sans-serif;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; /* Modern font stack */
        line-height: 1.6;
        margin: 0;
        padding: 0;
        background-color: #f4f4f4; /* Light grey background for the email client window */
        color: #333;
      }
      .email-container {
        max-width: 700px; /* Max width for content */
        margin: 20px auto; /* Center the container with space around */
        padding: 20px;
        background-color: #ffffff; /* White background for content */
        border-radius: 8px;
        box-shadow: 0 0 10px rgba(0,0,0,0.1);
      }
      h2 {
        color: #2c3e50; /* Dark blue-grey for main header */
        font-size: 24px;
        margin-top: 0;
        border-bottom: 2px solid #3498db; /* Blue accent border */
        padding-bottom: 10px;
      }
      h3 {
        color: #34495e; /* Slightly lighter blue-grey for sub-headers */
        font-size: 20px;
        border-bottom: 1px solid #bdc3c7; /* Lighter grey border */
        padding-bottom: 8px;
        margin-top: 30px;
      }
      p { margin: 15px 0; font-size: 14px; }
      .sub-header-text { font-weight: bold; color: #555; font-size: 16px; margin-bottom: 5px;}
      .report-description { font-size: 13px; color: #7f8c8d; margin-top: -10px; margin-bottom: 25px; text-align: center;}
      .list-description-text { text-align: center; }
      table {
        border-collapse: collapse;
        width: 100%;
        margin-top: 20px;
        font-size: 13px; /* Slightly smaller font for tables */
        box-shadow: 0 2px 3px rgba(0,0,0,0.05);
      }
      th, td {
        border: 1px solid #e0e0e0; /* Lighter borders for table cells */
        padding: 10px 12px; /* More padding in cells */
        text-align: left;
      }
      th {
        background-color: #3498db; /* Blue header for tables */
        color: #ffffff; /* White text for table headers */
        font-weight: bold;
      }
      tr:nth-child(even) { background-color: #f9f9f9; } /* Zebra striping for rows */
      .number-cell { text-align: center; }
      .highlight-red { color: #e74c3c; font-weight: bold; } /* Brighter red for highlighting */
      .chart-container {
        margin: 25px 0;
        padding: 15px;
        text-align: center;
        background-color: #fdfdfd;
        border: 1px solid #eee;
        border-radius: 5px;
      }
      .section {
        margin-bottom: 30px;
        padding: 20px;
        background-color: #fdfdfd;
        border: 1px solid #ecf0f1;
        border-radius: 5px;
      }
      .footer-note {
        font-size: 12px;
        color: #95a5a6;
        text-align: center;
        margin-top: 30px;
      }
      a { color: #3498db; text-decoration: none; }
      a:hover { text-decoration: underline; }
      .detailed-table-container {
        overflow-x: auto; /* Allow horizontal scroll for wide table */
        width: 100%;
      }
      .detailed-candidate-table {
        font-size: 11px; /* Smaller font for the very detailed table */
        white-space: nowrap; /* Prevent text wrapping in cells that might cause overflow */
      }
      .detailed-candidate-table th, .detailed-candidate-table td {
        padding: 6px 8px; /* Reduce padding slightly for this dense table */
      }
    </style>
  `;

  let emailBody = `<html><head>${styles}</head><body>`;
  emailBody += `<div class="email-container">`; // Start of main email container

  emailBody += `<h2>AI Screening: Missed Opportunities</h2>`;
  emailBody += `<p class="sub-header-text">Candidates eligible for AI screening but overlooked</p>`;
  emailBody += `<p class="report-description">Applications received after ${THRESHOLD_DATE_STR} and have been advanced past NEW stage.</p>`;

  // --- Pie Chart Generation (Overall) ---
  let overallChartBlob = null;
  const inlineImages = {};

  if (totalRelevantApps > 0) {
    try {
      const takenOrOtherCount = totalRelevantApps - aiInterviewNotTakenCount;
      const chartData = Charts.newDataTable()
          .addColumn(Charts.ColumnType.STRING, "Status")
          .addColumn(Charts.ColumnType.NUMBER, "Count")
          .addRow(['Not Extended', aiInterviewNotTakenCount])
          .addRow(['AI Invitation Extended', takenOrOtherCount])
          .build();

      const chart = Charts.newPieChart()
          .setDataTable(chartData)
          .setTitle('Overall AI Interview Status (Relevant Apps)')
          .setOption('pieSliceText', 'percentage') // Show percentages on slices
          .setOption('colors', ['#DC3912', '#109618']) // Set colors: Red for 'Not Taken', Green for 'Taken/Other'
          .setOption('width', 500)
          .setOption('height', 300)
          .build();

      overallChartBlob = chart.getBlob().getAs('image/png'); // Use PNG for wider compatibility
      inlineImages['overallPieChart'] = overallChartBlob;
      Logger.log('Successfully generated overall pie chart.');
    } catch (e) {
      Logger.log(`Error generating overall pie chart: ${e}`);
      // Chart generation failed, proceed without it
    }
  }

  // 1. Overall Metrics
  emailBody += `<div class="section">`; // Start of Overall Summary section
  emailBody += `<h3>Overall Summary</h3>`;
  if (totalRelevantApps > 0) {
    const percentNotTaken = ((aiInterviewNotTakenCount / totalRelevantApps) * 100).toFixed(1);
    emailBody += `<p>Number of applications that were advanced past the NEW stage: ${totalRelevantApps}</p>`;
    emailBody += `<p>AI Invitations not extended: ${aiInterviewNotTakenCount}</p>`;
    emailBody += `<p><b>Missed Opportunities: ${percentNotTaken}%</b></p>`;

    // Embed chart if successfully generated
    if (overallChartBlob) {
        emailBody += `<div class="chart-container"><img src="cid:overallPieChart"></div>`;
    }

  } else {
    emailBody += `<p>No applications met the criteria for this period.</p>`;
  }
  emailBody += `</div>`; // End of Overall Summary section

  // 2. Recruiter Breakdown
  emailBody += `<div class="section">`; // Start of Recruiter Breakdown section
  emailBody += `<h3>Recruiter Breakdown</h3>`;
  if (Object.keys(recruiterStats).length > 0) {
    emailBody += `<table>
                    <thead>
                      <tr>
                        <th>Recruiter Name</th>
                        <th>Total Relevant Apps</th>
                        <th>AI Interviews Not Taken ('N')</th>
                        <th>% Not Taken</th>
                      </tr>
                    </thead>
                    <tbody>`;

    // Sort recruiters alphabetically for consistent reporting
    const sortedRecruiters = Object.keys(recruiterStats).sort();

    sortedRecruiters.forEach(name => {
      const stats = recruiterStats[name];
      const percentRecruiterNotTaken = (stats.total > 0) ? ((stats.aiNotTaken / stats.total) * 100).toFixed(1) : 0.0; // Get number for comparison
      const recruiterNameCell = (parseFloat(percentRecruiterNotTaken) > 20.0) ?
                                 `<td class="highlight-red">${name}</td>` :
                                 `<td>${name}</td>`;

      emailBody += `<tr>
                      ${recruiterNameCell} 
                      <td class="number-cell">${stats.total}</td>
                      <td class="number-cell">${stats.aiNotTaken}</td>
                      <td class="number-cell">${percentRecruiterNotTaken}%</td>
                    </tr>`;
    });

    emailBody += `</tbody></table>`;
  } else {
    emailBody += `<p>No relevant application data found for any recruiter.</p>`;
  }
  emailBody += `</div>`; // End of Recruiter Breakdown section

  // 3. Link to Detailed List in Google Sheet
  emailBody += `<div class="section">`; // Start of Detailed List section
  emailBody += `<h3>Detailed List of Lost Opportunities</h3>`;
  if (lostCandidatesDetails.length > 0) {
    const listDescription = `The following ${lostCandidatesDetails.length} candidates could have been given an AI screening but were missed.`;
    emailBody += `<p class="list-description-text">${listDescription}</p>`;

    // Add link to the sheet
    if (destinationSheetUrl !== '#') {
       emailBody += `<p>This list has also been written to the Google Sheet: <a href="${destinationSheetUrl}" target="_blank">${DESTINATION_SHEET_NAME}</a></p>`;
    } else {
       emailBody += `<p style="color: red;"><b>Note:</b> There was an error writing this list to the Google Sheet. Please check logs.</p>`;
    }

    // Re-add the HTML table to the email body
    emailBody += `<div class="detailed-table-container">`; // Container for horizontal scroll
    emailBody += `<table class="detailed-candidate-table"> <!-- Use styled table with specific class -->
                    <thead>
                      <tr>
                        <th>Recruiter Name</th>
                        <th>Name</th>
                        <th>Current Company</th>
                        <th>Profile Link</th>
                        <th>Title</th>
                        <th>Hiring Manager</th>
                        <th>AI Interview</th>
                        <th>Last Stage</th>
                        <th>Last Stage Ts</th>
                        <th>Application Ts</th>
                      </tr>
                    </thead>
                    <tbody>`;

    lostCandidatesDetails.forEach(candidate => {
      // Format profile link if it exists
      const profileLinkHtml = candidate.profileLink ? `<a href="${candidate.profileLink}" target="_blank">Link</a>` : 'N/A';

      emailBody += `<tr>
                      <td>${candidate.recruiterName}</td>
                      <td>${candidate.name || 'N/A'}</td>
                      <td>${candidate.currentCompany || 'N/A'}</td>
                      <td>${profileLinkHtml}</td>
                      <td>${candidate.title || 'N/A'}</td>
                      <td>${candidate.hiringManager || 'N/A'}</td>
                      <td>${candidate.aiInterviewStatus}</td>
                      <td>${candidate.lastStage}</td>
                      <td>${candidate.lastStageTsFormatted || 'N/A'}</td>
                      <td>${candidate.applicationTsFormatted || 'N/A'}</td>
                    </tr>`;
    });

    emailBody += `</tbody></table>`;
    emailBody += `</div>`; // End of detailed-table-container

   } else {
     emailBody += `<p>No candidates met the criteria to be listed.</p>`;
   }
   emailBody += `</div>`; // End of Detailed List section

   emailBody += `<p class="footer-note">This is an automated report. Please do not reply directly to this email.</p>`;
   emailBody += `</div>`; // End of main email container
   emailBody += `</body></html>`; // Close HTML tags


  // --- Send Email ---
  try {
    // Prepare email options
    const mailOptions = {
        to: EMAIL_RECIPIENT,
        subject: emailSubject,
        htmlBody: emailBody,
        name: 'Automated AI Interview Reporter' // Optional: Set sender name
    };

    // Add inline images if any were generated
    if (Object.keys(inlineImages).length > 0) {
        mailOptions.inlineImages = inlineImages;
    }

    MailApp.sendEmail(mailOptions);
    Logger.log(`Report successfully generated and sent to ${EMAIL_RECIPIENT}`);
  } catch (e) {
    Logger.log(`Error sending email: ${e}`);
    // Optionally send a notification about the failure
     MailApp.sendEmail(EMAIL_RECIPIENT,
                        `FAILED: Daily AI Report Script`,
                        `The script ran but failed to send the report email. Error: ${e.message}. Check Logs in the Apps Script editor.`);
  }
}