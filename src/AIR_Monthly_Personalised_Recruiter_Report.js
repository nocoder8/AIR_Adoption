/**
 * AIR Monthly Personalised Recruiter Reports
 * Sends personalized monthly reports to each recruiter with their AI interview usage data
 * 
 * Last Updated: 2025-08-11 10:21 AM IST
 * Changes: Modified to send reports monthly with 30-day recent data
 */

// --- Configuration ---
const WEEKLY_REPORTS_CONFIG = {
  // Application Sheet Configuration
  APP_SHEET_SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw/edit',
  APP_SHEET_NAME: 'Active+Rejected',
  
  // Interview Log Sheet Configuration (for time-to-send metrics)
  LOG_SHEET_SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE/edit?gid=602557471#gid=602557471',
  LOG_SHEET_NAME: 'Log_Enhanced',
  
  // Email Configuration
  FROM_EMAIL: 'ai-interview-reports@eightfold.ai', // Update with actual sender email
  COMPANY_NAME: 'Eightfold',
  
  // Date Filters
  HISTORICAL_CUTOFF_DATE: new Date('2025-05-01'), // May 1st, 2025 for historical data
  RECENT_DAYS: 30, // Last 30 days for recent data
  
  // Eligible stages for AI interview analysis
  ELIGIBLE_STAGES: [
    'HIRING MANAGER SCREEN',
    'ASSESSMENT', 
    'ONSITE INTERVIEW',
    'FINAL INTERVIEW',
    'OFFER APPROVALS',
    'OFFER EXTENDED',
    'OFFER DECLINED',
    'PENDING START',
    'HIRED'
  ],
  
  // Excluded recruiters (these won't receive reports)
  EXCLUDED_RECRUITERS: [
    'Deepak Sehgal',
    'Pavan Kumar', 
    'Simran Chriselda',
    'Samrudh J',
    'Guruprasad Hegde',
    'Nishant Chaudhary',
    'Greg Schneider',
    'Unknown', 
    'Unassigned'
  ]
};

/**
 * Main function to generate and send monthly personalised recruiter reports
 * This function should be scheduled to run monthly
 */
function generateMonthlyPersonalisedRecruiterReports() {
  try {
    Logger.log('=== Starting Monthly Personalised Recruiter Reports Generation ===');
    
    // Check if this is the 3rd Monday of the month
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    const dayOfMonth = today.getDate();
    
    // Only run on the 3rd Monday of each month (Monday = 1, and day of month between 15-21)
    if (dayOfWeek !== 1 || dayOfMonth < 15 || dayOfMonth > 21) {
      Logger.log(`Skipping report generation - not the 3rd Monday of the month. Today: ${today.toDateString()}, Day of week: ${dayOfWeek}, Day of month: ${dayOfMonth}`);
      return;
    }
    
    Logger.log(`Proceeding with monthly report generation - 3rd Monday of the month confirmed`);
    
    // Get application data
    const appData = getApplicationDataForWeeklyReports();
    if (!appData || !appData.rows || appData.rows.length === 0) {
      Logger.log('ERROR: No application data retrieved. Cannot generate reports.');
      return;
    }
    
    Logger.log(`Retrieved ${appData.rows.length} rows from application sheet`);
    
    // Get interview sent times for time-to-send metrics
    const interviewSentMap = getInterviewSentTimes();
    Logger.log(`Retrieved ${interviewSentMap.size} unique profile IDs with interview sent times.`);
    
    // Calculate recruiter metrics
    const recruiterMetrics = calculateRecruiterMetrics(appData.rows, appData.colIndices, interviewSentMap);
    if (!recruiterMetrics || Object.keys(recruiterMetrics).length === 0) {
      Logger.log('ERROR: No recruiter metrics calculated. Cannot generate reports.');
      return;
    }
    
    Logger.log(`Calculated metrics for ${Object.keys(recruiterMetrics).length} recruiters`);
    
    // Generate and send reports for each recruiter
    let successCount = 0;
    let failureCount = 0;
    
    Object.entries(recruiterMetrics).forEach(([recruiterName, metrics]) => {
      try {
        // Skip excluded recruiters
        if (WEEKLY_REPORTS_CONFIG.EXCLUDED_RECRUITERS.some(excluded => 
          recruiterName.toLowerCase().includes(excluded.toLowerCase()))) {
          Logger.log(`Skipping excluded recruiter: ${recruiterName}`);
          return;
        }
        
        // Generate personalized report
        const reportHtml = generateRecruiterReportHtml(recruiterName, metrics);
        
        // Send email
        const emailSubject = `${WEEKLY_REPORTS_CONFIG.COMPANY_NAME} AI Interview Usage Report - Monthly - ${recruiterName}`;
        const recipientEmail = getRecruiterEmail(recruiterName, metrics); // Pass metrics to get email
        
        if (recipientEmail) {
          sendRecruiterReportEmail(recipientEmail, emailSubject, reportHtml);
          Logger.log(`SUCCESS: Sent report to ${recruiterName} (${recipientEmail})`);
          successCount++;
        } else {
          Logger.log(`WARNING: No email found for recruiter: ${recruiterName}`);
          failureCount++;
        }
        
      } catch (error) {
        Logger.log(`ERROR sending report to ${recruiterName}: ${error.toString()}`);
        failureCount++;
      }
    });
    
    Logger.log(`=== Monthly Personalised Reports Summary ===`);
    Logger.log(`Successfully sent: ${successCount} reports`);
    Logger.log(`Failed to send: ${failureCount} reports`);
    Logger.log(`Total recruiters processed: ${Object.keys(recruiterMetrics).length}`);
    
  } catch (error) {
    Logger.log(`CRITICAL ERROR in generateWeeklyRecruiterReports: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Retrieves application data for weekly reports
 * @returns {object} Object containing rows and column indices
 */
function getApplicationDataForWeeklyReports() {
  Logger.log('--- Getting Application Data for Monthly Reports ---');
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(WEEKLY_REPORTS_CONFIG.APP_SHEET_SPREADSHEET_URL);
    Logger.log(`Opened application spreadsheet: ${spreadsheet.getName()}`);
    
    let sheet = spreadsheet.getSheetByName(WEEKLY_REPORTS_CONFIG.APP_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${WEEKLY_REPORTS_CONFIG.APP_SHEET_NAME}" not found. Trying first sheet.`);
      sheet = spreadsheet.getSheets()[0];
      if (!sheet) {
        throw new Error('No sheets found in application spreadsheet');
      }
      Logger.log(`Using first sheet: "${sheet.getName()}"`);
    }
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length < 3) {
      Logger.log('Not enough data in application sheet (expected headers in row 2)');
      return null;
    }
    
    const headers = data[1].map(String); // Headers from Row 2
    const rows = data.slice(2); // Data from Row 3 onwards
    
    Logger.log(`Found ${headers.length} columns and ${rows.length} data rows`);
    
    // Define required columns
    const requiredColumns = [
      'Recruiter name', 'Recruiter email', 'Last_stage', 'Ai_interview', 'Application_ts', 
      'Name', 'Position_id', 'Title', 'Current_company', 'Application_status', 'Source_name', 'Profile_id', 'Profile_link'
    ];
    
    const colIndices = {};
    const missingColumns = [];
    
    requiredColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index === -1) {
        missingColumns.push(colName);
      } else {
        colIndices[colName] = index;
      }
    });
    
    if (missingColumns.length > 0) {
      Logger.log(`ERROR: Missing required columns: ${missingColumns.join(', ')}`);
      return null;
    }
    
    Logger.log(`Found all required columns. Indices: ${JSON.stringify(colIndices)}`);
    return { rows, headers, colIndices };
    
  } catch (error) {
    Logger.log(`ERROR getting application data: ${error.toString()}`);
    return null;
  }
}

/**
 * Calculates metrics for each recruiter
 * @param {Array<Array>} appRows Application data rows
 * @param {object} colIndices Column indices
 * @param {Map<string, Date>} interviewSentMap Map of Profile ID to interview sent date
 * @returns {object} Object with recruiter metrics
 */
function calculateRecruiterMetrics(appRows, colIndices, interviewSentMap = new Map()) {
  Logger.log('--- Calculating Recruiter Metrics ---');
  Logger.log(`Interview sent map contains ${interviewSentMap.size} profile IDs`);
  
  const recruiterMetrics = {};
  const currentDate = new Date();
  const fourteenDaysAgo = new Date(currentDate.getTime() - (WEEKLY_REPORTS_CONFIG.RECENT_DAYS * 24 * 60 * 60 * 1000));
  
  appRows.forEach((row, index) => {
    // Basic validation
    if (!row || row.length <= Math.max(...Object.values(colIndices))) {
      return;
    }
    
    const recruiterName = String(row[colIndices['Recruiter name']] || '').trim();
    const lastStage = String(row[colIndices['Last_stage']] || '').trim().toUpperCase();
    const aiInterview = String(row[colIndices['Ai_interview']] || '').trim().toUpperCase();
    const applicationTs = vsParseDateSafeRB(row[colIndices['Application_ts']]);
    const profileId = row[colIndices['Profile_id']]?.toString().trim();
    
    // Skip if no recruiter name
    if (!recruiterName) {
      return;
    }
    
    // Skip excluded recruiters
    if (WEEKLY_REPORTS_CONFIG.EXCLUDED_RECRUITERS.some(excluded => 
      recruiterName.toLowerCase().includes(excluded.toLowerCase()))) {
      return;
    }
    
    // Check if candidate is eligible
    const isEligible = WEEKLY_REPORTS_CONFIG.ELIGIBLE_STAGES.some(stage => 
      stage.toUpperCase() === lastStage);
    
    // Initialize recruiter data if not exists
    if (!recruiterMetrics[recruiterName]) {
      recruiterMetrics[recruiterName] = {
        email: String(row[colIndices['Recruiter email']] || '').trim(),
        historical: { eligible: 0, aiDone: 0, percentage: 0 },
        recent: { eligible: 0, aiDone: 0, percentage: 0 },
        detailedData: [],
        timeDiffSumMs: 0, // For calculating average time to send
        timeDiffCount: 0  // Count of valid time differences
      };
    }
    
    const metrics = recruiterMetrics[recruiterName];
    
    if (isEligible) {
      // Historical data (since May 1st, 2025)
      if (applicationTs && applicationTs >= WEEKLY_REPORTS_CONFIG.HISTORICAL_CUTOFF_DATE) {
        metrics.historical.eligible++;
        if (aiInterview === 'Y') {
          metrics.historical.aiDone++;
        }
      }
      
      // Recent data (last 14 days)
      if (applicationTs && applicationTs >= fourteenDaysAgo) {
        metrics.recent.eligible++;
        if (aiInterview === 'Y') {
          metrics.recent.aiDone++;
        }
      }
      
      // Add to detailed data (only eligible candidates since May 1st, 2025, excluding "Not Eligible" status)
      if (applicationTs && applicationTs >= WEEKLY_REPORTS_CONFIG.HISTORICAL_CUTOFF_DATE && 
          (aiInterview !== 'N' || lastStage !== 'NEW' && lastStage !== 'ADDED')) {
        
        // Calculate time to invite sent for this candidate
        let timeToInviteDays = null;
        if (aiInterview === 'Y' && profileId && applicationTs) {
          const sentAtDate = interviewSentMap.get(profileId);
          if (sentAtDate) {
            const timeDiffMs = sentAtDate.getTime() - applicationTs.getTime();
            if (timeDiffMs >= 0) { // Only consider non-negative differences
              timeToInviteDays = parseFloat((timeDiffMs / (1000 * 60 * 60 * 24)).toFixed(1));
            }
          }
        }
        
        metrics.detailedData.push({
          name: row[colIndices['Name']] || 'N/A',
          profileLink: row[colIndices['Profile_link']] || null,
          positionId: row[colIndices['Position_id']] || 'N/A',
          title: row[colIndices['Title']] || 'N/A',
          company: row[colIndices['Current_company']] || 'N/A',
          sourceName: row[colIndices['Source_name']] || 'N/A',
          lastStage: lastStage,
          aiInterview: aiInterview,
          applicationDate: applicationTs ? applicationTs.toLocaleDateString() : 'N/A',
          applicationDateRaw: applicationTs, // Keep raw date for sorting
          applicationStatus: row[colIndices['Application_status']] || 'N/A',
          timeToInviteDays: timeToInviteDays // Add time to invite for this candidate
        });
      }
      
      // Calculate time to invite sent for eligible candidates who received an invite
      if (aiInterview === 'Y' && profileId && applicationTs) {
        const sentAtDate = interviewSentMap.get(profileId);
        if (sentAtDate) {
          const timeDiffMs = sentAtDate.getTime() - applicationTs.getTime();
          if (timeDiffMs >= 0) { // Only consider non-negative differences
            metrics.timeDiffSumMs += timeDiffMs;
            metrics.timeDiffCount++;
          }
        } else {
          // Debug logging for first few unmatched profile IDs
          if (metrics.timeDiffCount === 0 && index < 5) {
            Logger.log(`DEBUG: Profile ID "${profileId}" not found in interview sent map for ${recruiterName}`);
          }
        }
      }
    }
  });
  
      // Calculate percentages, average time, and sort detailed data
    Object.values(recruiterMetrics).forEach(metrics => {
      if (metrics.historical.eligible > 0) {
        metrics.historical.percentage = parseFloat(
          ((metrics.historical.aiDone / metrics.historical.eligible) * 100).toFixed(1)
        );
      }
      
      if (metrics.recent.eligible > 0) {
        metrics.recent.percentage = parseFloat(
          ((metrics.recent.aiDone / metrics.recent.eligible) * 100).toFixed(1)
        );
      }
      
      // Calculate average time to invite sent
      if (metrics.timeDiffCount > 0) {
        const avgTimeMs = metrics.timeDiffSumMs / metrics.timeDiffCount;
        metrics.avgTimeToInviteDays = parseFloat((avgTimeMs / (1000 * 60 * 60 * 24)).toFixed(1));
      } else {
        metrics.avgTimeToInviteDays = null;
      }
      
      // Sort detailed data by application date (most recent first)
      if (metrics.detailedData.length > 0) {
        metrics.detailedData.sort((a, b) => {
          // Handle cases where dates might be 'N/A'
          if (!a.applicationDateRaw && !b.applicationDateRaw) return 0;
          if (!a.applicationDateRaw) return 1; // Put 'N/A' dates at the end
          if (!b.applicationDateRaw) return -1;
          
          // Sort by date (most recent first)
          return b.applicationDateRaw - a.applicationDateRaw;
        });
      }
    });
  
  Logger.log(`Calculated metrics for ${Object.keys(recruiterMetrics).length} recruiters`);
  return recruiterMetrics;
}

/**
 * Generates HTML report for a specific recruiter
 * @param {string} recruiterName Name of the recruiter
 * @param {object} metrics Recruiter metrics
 * @returns {string} HTML content
 */
function generateRecruiterReportHtml(recruiterName, metrics) {
  const currentDate = new Date().toLocaleDateString();
  const reportPeriod = `${WEEKLY_REPORTS_CONFIG.RECENT_DAYS} days ending ${currentDate}`;
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>AI Interview Usage Report - ${recruiterName}</title>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 20px; margin: 0;">
      <div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); padding: 30px;">
        
        <!-- Header -->
        <div style="text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #eee;">
          <h1 style="color: #1a237e; margin: 0; font-size: 28px;">AI Interview Usage Report</h1>
          <p style="color: #666; margin: 10px 0 0 0; font-size: 16px;">Monthly Personalised Report for ${recruiterName}</p>
          <p style="color: #999; margin: 5px 0 0 0; font-size: 14px;">Generated on ${currentDate}</p>
        </div>
        
        <!-- Summary Boxes -->
        <div style="text-align: center; margin-bottom: 40px;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="width: 33.33%; padding-right: 10px; vertical-align: top;">
                <div style="background-color: ${metrics.historical.percentage >= 80 ? '#e8f5e9' : '#ffebee'}; border: 1px solid ${metrics.historical.percentage >= 80 ? '#4caf50' : '#f44336'}; border-radius: 8px; padding: 20px; text-align: center; max-width: 250px; margin: 0 auto;">
                  <h3 style="color: ${metrics.historical.percentage >= 80 ? '#2e7d32' : '#c62828'}; margin: 0 0 10px 0; font-size: 16px;">Historical AI Usage</h3>
                  <div style="font-size: 28px; font-weight: bold; color: ${metrics.historical.percentage >= 80 ? '#2e7d32' : '#c62828'}; margin-bottom: 5px;">
                    ${metrics.historical.percentage}%
                  </div>
                  <div style="font-size: 12px; color: #666;">
                    ${metrics.historical.aiDone} of ${metrics.historical.eligible} eligible candidates
                  </div>
                  <div style="font-size: 11px; color: #999; margin-top: 5px;">
                    Since May 1st, 2025
                  </div>
                </div>
              </td>
              <td style="width: 33.33%; padding-left: 5px; padding-right: 5px; vertical-align: top;">
                <div style="background-color: ${metrics.recent.percentage >= 80 ? '#e8f5e9' : '#ffebee'}; border: 1px solid ${metrics.recent.percentage >= 80 ? '#4caf50' : '#f44336'}; border-radius: 8px; padding: 20px; text-align: center; max-width: 250px; margin: 0 auto;">
                  <h3 style="color: ${metrics.recent.percentage >= 80 ? '#2e7d32' : '#c62828'}; margin: 0 0 10px 0; font-size: 16px;">Recent AI Usage</h3>
                  <div style="font-size: 28px; font-weight: bold; color: ${metrics.recent.percentage >= 80 ? '#2e7d32' : '#c62828'}; margin-bottom: 5px;">
                    ${metrics.recent.percentage}%
                  </div>
                  <div style="font-size: 12px; color: #666;">
                    ${metrics.recent.aiDone} of ${metrics.recent.eligible} eligible candidates
                  </div>
                  <div style="font-size: 11px; color: #999; margin-top: 5px;">
                    Last 30 days
                  </div>
                </div>
              </td>
              <td style="width: 33.33%; padding-left: 10px; vertical-align: top;">
                <div style="background-color: ${metrics.avgTimeToInviteDays !== null && metrics.avgTimeToInviteDays <= 7 ? '#e8f5e9' : '#fff3e0'}; border: 1px solid ${metrics.avgTimeToInviteDays !== null && metrics.avgTimeToInviteDays <= 7 ? '#4caf50' : '#ff9800'}; border-radius: 8px; padding: 20px; text-align: center; max-width: 250px; margin: 0 auto;">
                  <h3 style="color: ${metrics.avgTimeToInviteDays !== null && metrics.avgTimeToInviteDays <= 7 ? '#2e7d32' : '#e65100'}; margin: 0 0 10px 0; font-size: 16px;">Avg. Time to Send</h3>
                  <div style="font-size: 28px; font-weight: bold; color: ${metrics.avgTimeToInviteDays !== null && metrics.avgTimeToInviteDays <= 7 ? '#2e7d32' : '#e65100'}; margin-bottom: 5px;">
                    ${metrics.avgTimeToInviteDays !== null ? `${metrics.avgTimeToInviteDays} days` : 'N/A'}
                  </div>
                  <div style="font-size: 12px; color: #666;">
                    ${metrics.timeDiffCount} candidates with data
                  </div>
                  <div style="font-size: 11px; color: #999; margin-top: 5px;">
                    Application to AI invite
                  </div>
                </div>
              </td>
            </tr>
          </table>
        </div>
        
        <!-- Detailed Data Table -->
        <div style="margin-bottom: 30px;">
          <h2 style="color: #3f51b5; margin-bottom: 20px; font-size: 22px; text-align: center;">📊 Detailed Candidate Data</h2>
          <p style="color: #666; margin-bottom: 15px; font-size: 14px; text-align: center;">
            Showing eligible candidates since May 1st, 2025 (excluding "New" and "Added" stages) with their AI interview status.
          </p>
          <p style="color: #666; margin-bottom: 15px; font-size: 12px; text-align: center; font-style: italic;">
            <strong>Note:</strong> The number of days shown under "✅ Sent" indicates the time between when the candidate applied and when their AI interview invitation was sent. Lower numbers indicate faster response times.
          </p>
          
          ${metrics.detailedData.length > 0 ? `
            <div style="overflow-x: auto;">
              <table style="width: 100%; border-collapse: collapse; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
                <thead>
                  <tr style="background-color: #f5f5f5;">
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; font-weight: bold; color: #424242;">Candidate Name</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; font-weight: bold; color: #424242;">Position Title</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; font-weight: bold; color: #424242;">Company</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px; font-weight: bold; color: #424242;">Source</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; font-weight: bold; color: #424242;">Last Stage</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; font-weight: bold; color: #424242;">AI Interview</th>
                    <th style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px; font-weight: bold; color: #424242;">Application Date</th>
                  </tr>
                </thead>
                <tbody>
                  ${metrics.detailedData.map((candidate, index) => {
                    const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
                    
                    // Create AI status with time difference for sent interviews
                    let aiStatus;
                    if (candidate.aiInterview === 'Y') {
                      const timeText = candidate.timeToInviteDays !== null ? 
                        `<br><span style="color: #999; font-size: 10px; font-weight: normal;">${candidate.timeToInviteDays} Days</span>` : '';
                      aiStatus = `<span style="color: #4CAF50; font-weight: bold;">✅ Sent${timeText}</span>`;
                    } else {
                      aiStatus = '<span style="color: #F44336; font-weight: bold;">❌ Missing</span>';
                    }
                    
                    // Create hyperlinked name if profile link is available
                    const candidateName = candidate.profileLink ? 
                      `<a href="${candidate.profileLink}" target="_blank" style="color: #007bff; text-decoration: none; font-weight: bold;">${candidate.name}</a>` : 
                      candidate.name;
                    
                    return `
                      <tr style="background-color: ${bgColor};">
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px;">${candidateName}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px;">${candidate.title}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px;">${candidate.company}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: left; font-size: 12px;">${candidate.sourceName}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px;">${candidate.lastStage}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px;">${aiStatus}</td>
                        <td style="border: 1px solid #e0e0e0; padding: 8px 12px; text-align: center; font-size: 12px;">${candidate.applicationDate}</td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          ` : `
            <div style="text-align: center; padding: 40px; color: #666; font-size: 16px;">
              No eligible candidates found for this period.
            </div>
          `}
        </div>
        
        <!-- Footer -->
        <div style="text-align: center; padding-top: 20px; border-top: 1px solid #eee; color: #999; font-size: 12px;">
          <p>This personalised report is automatically generated monthly by the AI Interview System.</p>
          <p>Eligible stages: Hiring Manager Screen, Assessment, Onsite Interview, Final Interview, Offer stages, Hired</p>
          <p>For questions or support, please contact the AI Interview team.</p>
        </div>
        
      </div>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Gets email address for a recruiter
 * @param {string} recruiterName Name of the recruiter
 * @param {object} metrics Recruiter metrics containing email
 * @returns {string|null} Email address or null if not found
 */
function getRecruiterEmail(recruiterName, metrics) {
  // Use email from the data if available
  if (metrics && metrics.email && metrics.email.trim() !== '') {
    Logger.log(`Using email from data for ${recruiterName}: ${metrics.email}`);
    return metrics.email;
  }
  
  // Fallback to testing mode if no email in data
  Logger.log(`No email found in data for ${recruiterName}, using testing mode`);
  return 'pkumar@eightfold.ai';
  
  /* 
  // ACTUAL IMPLEMENTATION (uncomment after testing):
  // This is a placeholder function - you'll need to implement this based on your data structure
  // You could:
  // 1. Look up in a separate sheet with recruiter emails
  // 2. Use a naming convention (e.g., firstname.lastname@company.com)
  // 3. Store emails in script properties
  // 4. Use a mapping object
  
  // For now, using a simple mapping - replace with your actual implementation
  const emailMapping = {
    'Akhila Kashyap': 'akashyap@eightfold.ai',
    'Pavan Kumar': 'pkumar@eightfold.ai',
    // Add more mappings as needed
  };
  
  // Try exact match first
  if (emailMapping[recruiterName]) {
    return emailMapping[recruiterName];
  }
  
  // Try partial match
  for (const [name, email] of Object.entries(emailMapping)) {
    if (recruiterName.toLowerCase().includes(name.toLowerCase()) || 
        name.toLowerCase().includes(recruiterName.toLowerCase())) {
      return email;
    }
  }
  
  // If no match found, you could implement a naming convention
  // For example: firstname.lastname@eightfold.ai
  const nameParts = recruiterName.split(' ');
  if (nameParts.length >= 2) {
    const firstName = nameParts[0].toLowerCase();
    const lastName = nameParts[nameParts.length - 1].toLowerCase();
    return `${firstName}.${lastName}@eightfold.ai`;
  }
  
  Logger.log(`WARNING: No email found for recruiter: ${recruiterName}`);
  return null;
  */
}

/**
 * Sends email report to a recruiter
 * @param {string} recipientEmail Email address of the recruiter
 * @param {string} subject Email subject
 * @param {string} htmlBody HTML content of the email
 */
function sendRecruiterReportEmail(recipientEmail, subject, htmlBody) {
  try {
    const options = {
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: WEEKLY_REPORTS_CONFIG.COMPANY_NAME + ' AI Interview System'
    };
    
    MailApp.sendEmail(options);
    Logger.log(`Email sent successfully to: ${recipientEmail}`);
    
  } catch (error) {
    Logger.log(`ERROR sending email to ${recipientEmail}: ${error.toString()}`);
    throw error;
  }
}

/**
 * Creates a trigger to run monthly personalised recruiter reports
 * Runs on the first Monday of each month at 9 AM
 */
function createMonthlyPersonalisedRecruiterReportsTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateMonthlyPersonalisedRecruiterReports') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger for the 3rd Monday of each month at 4 PM
  // Note: Google Apps Script doesn't have monthly triggers, so we use weekly and check the date
  ScriptApp.newTrigger('generateMonthlyPersonalisedRecruiterReports')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(16) // 4 PM (24-hour format)
    .create();
    
  Logger.log('Monthly personalised recruiter reports trigger created (3rd Monday of each month at 4 PM)');
}

/**
 * Test function to generate a single monthly personalised recruiter report
 * @param {string} testRecruiterName Name of recruiter to test with
 */
function testSingleMonthlyPersonalisedRecruiterReport(testRecruiterName = 'Akhila Kashyap') {
  try {
    Logger.log(`=== Testing Single Monthly Personalised Recruiter Report for: ${testRecruiterName} ===`);
    
    // Get application data
    const appData = getApplicationDataForWeeklyReports();
    if (!appData || !appData.rows) {
      Logger.log('ERROR: Could not get application data');
      return;
    }
    
    // Get interview sent times for time-to-send metrics
    const interviewSentMap = getInterviewSentTimes();
    Logger.log(`Retrieved ${interviewSentMap.size} unique profile IDs with interview sent times.`);
    
    // Calculate metrics
    const recruiterMetrics = calculateRecruiterMetrics(appData.rows, appData.colIndices, interviewSentMap);
    if (!recruiterMetrics[testRecruiterName]) {
      Logger.log(`ERROR: No metrics found for recruiter: ${testRecruiterName}`);
      Logger.log(`Available recruiters: ${Object.keys(recruiterMetrics).join(', ')}`);
      return;
    }
    
    // Generate report
    const reportHtml = generateRecruiterReportHtml(testRecruiterName, recruiterMetrics[testRecruiterName]);
    
    // Log metrics for verification
    const metrics = recruiterMetrics[testRecruiterName];
    Logger.log(`=== Metrics for ${testRecruiterName} ===`);
    Logger.log(`Historical: ${metrics.historical.aiDone}/${metrics.historical.eligible} (${metrics.historical.percentage}%)`);
    Logger.log(`Recent: ${metrics.recent.aiDone}/${metrics.recent.eligible} (${metrics.recent.percentage}%)`);
    Logger.log(`Average time to invite: ${metrics.avgTimeToInviteDays !== null ? `${metrics.avgTimeToInviteDays} days` : 'N/A'} (${metrics.timeDiffCount} candidates with data)`);
    Logger.log(`Detailed data rows: ${metrics.detailedData.length}`);
    
    // Test email sending (enabled for testing)
    const testEmail = getRecruiterEmail(testRecruiterName, metrics);
    if (testEmail) {
      sendRecruiterReportEmail(testEmail, `TEST: AI Interview Report - ${testRecruiterName}`, reportHtml);
      Logger.log(`TEST EMAIL SENT to: ${testEmail}`);
    } else {
      Logger.log(`ERROR: No email found for test recruiter: ${testRecruiterName}`);
    }
    
    Logger.log(`SUCCESS: Monthly personalised report generated for ${testRecruiterName}`);
    Logger.log(`HTML length: ${reportHtml.length} characters`);
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Utility function to parse dates safely (copied from main script)
 * @param {any} dateInput Input value
 * @returns {Date|null} Parsed Date object or null
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
 * Reads the interview log sheet and returns a map of Profile ID to the earliest interview sent date.
 * @returns {Map<string, Date>} A map where keys are profile IDs and values are Date objects.
 */
function getInterviewSentTimes() {
  const interviewSentMap = new Map();
  Logger.log(`--- Starting getInterviewSentTimes ---`);
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(WEEKLY_REPORTS_CONFIG.LOG_SHEET_SPREADSHEET_URL);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
    let sheet = spreadsheet.getSheetByName(WEEKLY_REPORTS_CONFIG.LOG_SHEET_NAME);

    // Sheet finding logic (similar to main function)
    if (!sheet) {
      Logger.log(`Log sheet "${WEEKLY_REPORTS_CONFIG.LOG_SHEET_NAME}" not found by name. Attempting to use sheet by gid or first sheet.`);
      const gidMatch = WEEKLY_REPORTS_CONFIG.LOG_SHEET_SPREADSHEET_URL.match(/gid=(\d+)/);
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
          throw new Error(`Could not find any sheets in the log spreadsheet: ${WEEKLY_REPORTS_CONFIG.LOG_SHEET_SPREADSHEET_URL}`);
        }
      }
    }
     else {
        Logger.log(`Using specified log sheet: "${sheet.getName()}"`);
     }

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    if (data.length < 2) { // Need at least header row and one data row
      Logger.log(`Not enough data in log sheet "${sheet.getName()}". Found ${data.length} rows. Skipping time-to-send calculation.`);
      return interviewSentMap; // Return empty map
    }

    // Assume headers are in Row 1 (index 0) for this log sheet
    const headers = data[0].map(String);
    const profileIdColIndex = headers.indexOf('Profile_id');
    const sentAtColIndex = headers.indexOf('Interview_email_sent_at');

    if (profileIdColIndex === -1 || sentAtColIndex === -1) {
      throw new Error(`Required column(s) not found in log sheet headers (Row 1): ${profileIdColIndex === -1 ? 'Profile_id' : ''}${profileIdColIndex === -1 && sentAtColIndex === -1 ? ', ' : ''}${sentAtColIndex === -1 ? 'Interview_email_sent_at' : ''}`);
    }
    Logger.log(`Log sheet headers found: Profile_id at index ${profileIdColIndex}, Interview_email_sent_at at index ${sentAtColIndex}`);

    const rows = data.slice(1); // Data starts from row 2 (index 1)
    let processedCount = 0;
    let validDateCount = 0;

    rows.forEach((row, idx) => {
      if (row.length > Math.max(profileIdColIndex, sentAtColIndex)) {
        const profileId = row[profileIdColIndex]?.toString().trim(); // Ensure string and trim whitespace
        const rawSentAt = row[sentAtColIndex];

        if (profileId && rawSentAt !== null && rawSentAt !== undefined && rawSentAt !== '') {
          const sentAtDate = new Date(rawSentAt);
          if (!isNaN(sentAtDate.getTime())) {
            validDateCount++;
            // If we haven't seen this profileId, or if this date is earlier than the stored one
            if (!interviewSentMap.has(profileId) || sentAtDate < interviewSentMap.get(profileId)) {
              interviewSentMap.set(profileId, sentAtDate);
            }
          } else {
            // Logger.log(`Invalid date format for Interview_email_sent_at in log sheet row ${idx + 2}: ${rawSentAt}`); // Potentially noisy
          }
          processedCount++;
        }
      }
    });

    Logger.log(`Processed ${processedCount} rows from log sheet. Found ${validDateCount} valid send dates. Stored earliest date for ${interviewSentMap.size} unique Profile IDs.`);

  } catch (error) {
    Logger.log(`Error in getInterviewSentTimes: ${error.toString()} Stack: ${error.stack}`);
    // Don't throw error, just return empty map so main report can still run
    SpreadsheetApp.getUi().alert(`Warning: Could not process the Interview Log sheet. Time-to-send metrics will be unavailable. Error: ${error.message}`);
  }
  Logger.log(`--- Finished getInterviewSentTimes ---`);
  return interviewSentMap;
}


