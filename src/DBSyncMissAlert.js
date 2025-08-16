/**
 * Checks the timestamp in cell B1 of the specified sheet.
 * Sends an email alert if the timestamp is older than 1 hour.
 * 
 * Last Updated: 2025-08-11 10:20 AM IST
 * Changes: Fixed ReferenceError for totalOffsetInMinutes variable in alert body construction
 */
function checkTimestampAndAlert() {
  // --- Configuration ---
  const spreadsheetId = '1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE'; // The ID of your spreadsheet
  const sheetName = 'Volkscience Interview Log';
  const cellNotation = 'B1';
  const timeZone = 'GMT+05:30'; // India Standard Time
  const recipientEmail = 'pkumar@eightfold.ai'; // *** REPLACE with your email ID
  const alertSubject = 'ALERT: AIR Data is not syncing';
  const maxAgeMinutes = 60;
  // --- End Configuration ---

  try {
    // Get the sheet and cell value
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Error: Sheet "${sheetName}" not found.`);
      // Optional: Send an email about the sheet not being found
      // MailApp.sendEmail(recipientEmail, `Error: Sheet Not Found`, `Could not find sheet named "${sheetName}" in spreadsheet ID ${spreadsheetId}.`);
      return;
    }
    const cellValue = sheet.getRange(cellNotation).getValue();
    Logger.log(`Value in ${cellNotation}: ${cellValue}`);

    if (!cellValue || typeof cellValue !== 'string') {
        Logger.log(`Error: No valid string value found in ${cellNotation}. Found: ${cellValue}`);
        MailApp.sendEmail(recipientEmail, `ALERT: Invalid Timestamp in Sheet`, `Could not read a valid timestamp string from cell ${cellNotation} in sheet "${sheetName}". The cell value was: "${cellValue}"`);
        return;
    }

    // Parse the date string
    // Google Apps Script's Utilities.parseDate doesn't directly support "GMT+05:30" format well.
    // We'll manually adjust if needed or rely on the sheet's timezone setting interpretation.
    // Let's try parsing directly first, assuming the sheet correctly interprets the timezone offset.
    // A more robust way might involve string manipulation if direct parsing fails.

    let sheetTimestamp;
    try {
      // Attempt parsing directly - relies on underlying Java SimpleDateFormat which might handle offset
       // Format: "dd MMM yyyy HH:mm Z" (e.g., "04 May 2025 11:39 +0530")
       // We need to remove "GMT" first
       const cleanedCellValue = cellValue.replace('GMT', '').trim();
       // Utilities.parseDate expects format like "MMM dd, yyyy HH:mm:ss Z" or similar ISO
       // Let's manually parse
       const parts = cleanedCellValue.match(/(\d{2}) (\w{3}) (\d{4}) (\d{2}):(\d{2}) ([+-]\d{2}):?(\d{2})?/);
       if (!parts) {
         throw new Error(`Could not parse date string parts: ${cleanedCellValue}`);
       }
       // parts[1]=day, parts[2]=month, parts[3]=year, parts[4]=hour, parts[5]=minute, parts[6]=tz_hour, parts[7]=tz_minute
       const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
       const monthIndex = monthNames.indexOf(parts[2]);
       if (monthIndex === -1) {
           throw new Error(`Invalid month name: ${parts[2]}`);
       }

       // Construct date in UTC first to handle timezone correctly
       const year = parseInt(parts[3], 10);
       const day = parseInt(parts[1], 10);
       const hour = parseInt(parts[4], 10);
       const minute = parseInt(parts[5], 10);
       const tzHour = parseInt(parts[6], 10);
       const tzMinute = parts[7] ? parseInt(parts[7], 10) : 0;
       const tzOffsetMinutes = (tzHour * 60 + (tzHour < 0 ? -tzMinute : tzMinute));

       // Create Date object in UTC
       const dateInUTC = new Date(Date.UTC(year, monthIndex, day, hour, minute));

       // Adjust for the timezone offset to get the correct point in time
       const timestampMillis = dateInUTC.getTime() - (tzOffsetMinutes * 60 * 1000);
       sheetTimestamp = new Date(timestampMillis);

      Logger.log(`Parsed Timestamp: ${sheetTimestamp.toString()}`);

    } catch (e) {
      Logger.log(`Error parsing date string "${cellValue}": ${e.message}`);
      MailApp.sendEmail(recipientEmail, `ALERT: Timestamp Parse Error`, `Could not parse the timestamp string "${cellValue}" from cell ${cellNotation} in sheet "${sheetName}". Error: ${e.message}`);
      return;
    }

    // Get current time
    const now = new Date();
    Logger.log(`Current Time: ${now.toString()}`);


    // Calculate the difference in milliseconds
    const diffMillis = now.getTime() - sheetTimestamp.getTime();
    const diffMinutes = diffMillis / (1000 * 60);
    Logger.log(`Time difference: ${diffMinutes.toFixed(2)} minutes.`);

    // Check if the difference is greater than the threshold
    if (diffMinutes > maxAgeMinutes) {
      Logger.log(`Timestamp is older than ${maxAgeMinutes} minutes. Sending alert.`);
      const alertBody = `The timestamp in cell ${cellNotation} of sheet "${sheetName}" is outdated.\n\nCell Content: ${cellValue}\nParsed Timestamp: ${sheetTimestamp.toString()}\nCurrent Time: ${now.toString()}\nDifference: ${diffMinutes.toFixed(2)} minutes.`;
      MailApp.sendEmail(recipientEmail, alertSubject, alertBody);
      Logger.log(`Alert email sent to ${recipientEmail}.`);
    } else {
      Logger.log(`Timestamp is within the allowed ${maxAgeMinutes} minutes.`);
    }

  } catch (error) {
    // Log any unexpected errors during execution
    Logger.log(`An unexpected error occurred: ${error.message}\nStack: ${error.stack}`);
    // Optionally send an error email
    try {
      MailApp.sendEmail(recipientEmail, `Error in AIR Data Sync Checker Script`, `The script encountered an error: ${error.message}\n\nTimestamp Cell: ${cellNotation}\nSheet Name: ${sheetName}\nSpreadsheet ID: ${spreadsheetId}\n\nStack Trace:\n${error.stack}`);
    } catch (mailError) {
      Logger.log(`Failed to send error notification email: ${mailError.message}`);
    }
  }
}

// Global constant for timezone, if applicable for date parsing/display context
// const SCRIPT_TIME_ZONE = 'GMT+05:30'; // India Standard Time - Not directly used by current parsing logic, but good for context.

/**
 * Processes a single timestamp check based on the provided configuration.
 * @param {object} config The configuration object for the alert.
 * @param {string} config.spreadsheetId The ID of the spreadsheet.
 * @param {string} config.sheetName The name of the sheet.
 * @param {string} config.cellNotation The cell notation (e.g., 'B1').
 * @param {string} config.recipientEmail The email address to send alerts to.
 * @param {string} config.alertSubject The subject line for the alert email.
 * @param {number} config.maxAgeMinutes The maximum age of the timestamp in minutes.
 */
function executeSingleTimestampCheck(config) {
  // --- Configuration is now passed via config object ---

  try {
    // Get the sheet and cell value
    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      Logger.log(`Error: Sheet "${config.sheetName}" not found in spreadsheet ${config.spreadsheetId}. Alert: ${config.alertSubject}`);
      // Optional: Send an email about the sheet not being found
      MailApp.sendEmail(config.recipientEmail, `Error: Sheet Not Found - ${config.alertSubject}`, `Could not find sheet named "${config.sheetName}" in spreadsheet ID ${config.spreadsheetId}.`);
      return;
    }
    const cellValue = sheet.getRange(config.cellNotation).getValue();
    Logger.log(`Value in ${config.cellNotation} (Sheet: "${config.sheetName}", Spreadsheet: ${config.spreadsheetId}) for alert "${config.alertSubject}": ${cellValue}`);

    if (!cellValue || typeof cellValue !== 'string' || cellValue.trim() === '') {
        Logger.log(`Error: No valid string value found in ${config.cellNotation} (Sheet: "${config.sheetName}") for alert "${config.alertSubject}". Found: "${cellValue}"`);
        MailApp.sendEmail(config.recipientEmail, `ALERT: Invalid Timestamp in Sheet - ${config.alertSubject}`, `Could not read a valid timestamp string from cell ${config.cellNotation} in sheet "${config.sheetName}" (Spreadsheet: ${config.spreadsheetId}). The cell value was: "${cellValue}"`);
        return;
    }

    let sheetTimestamp;
    let parsedDatePartsForEmail = {};
    let totalOffsetInMinutes = 0;
    let parts = null;
    let cleanedCellValue = '';

    try {
       // Example format: "04 May 2025 11:39 GMT+05:30" or "04 May 2025 11:39 +0530"
       // Remove "GMT" if present, and trim.
       cleanedCellValue = cellValue.replace(/GMT/i, '').trim(); // Case-insensitive GMT removal

       // Regex for "dd MMM yyyy HH:mm [+-]HH:mm" or "dd MMM yyyy HH:mm [+-]HHMM"
       // Supports optional colon in timezone offset.
       parts = cleanedCellValue.match(/(\d{1,2}) (\w{3}) (\d{4}) (\d{1,2}):(\d{2})\s*([+-]\d{2}):?(\d{2})?/);
       if (!parts) {
         throw new Error(`Could not parse date string parts from: "${cleanedCellValue}" (Original: "${cellValue}")`);
       }
       
       parsedDatePartsForEmail = {
           day: parts[1], monthStr: parts[2], year: parts[3], hour: parts[4], minute: parts[5],
           tzSignHour: parts[6], tzMinuteStr: parts[7]
       };

       const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
       const monthIndex = monthNames.findIndex(m => m.toLowerCase() === parts[2].toLowerCase());
       if (monthIndex === -1) {
           throw new Error(`Invalid month name: ${parts[2]} in "${cleanedCellValue}"`);
       }

       const year = parseInt(parts[3], 10);
       const day = parseInt(parts[1], 10);
       const hour = parseInt(parts[4], 10);
       const minute = parseInt(parts[5], 10);
       const tzHourOffset = parseInt(parts[6], 10); // e.g., +5 or -3
       const tzMinuteOffset = parts[7] ? parseInt(parts[7], 10) : 0;
       
       // Calculate total offset in minutes from UTC. Example: +05:30 -> +330 minutes. -04:00 -> -240 minutes.
       // The sign of tzHourOffset determines the sign of the total offset.
       totalOffsetInMinutes = (tzHourOffset * 60) + (tzHourOffset < 0 ? -tzMinuteOffset : tzMinuteOffset);

       // Create a Date object using UTC values for year, month, day, hour, minute.
       // This interprets the parsed hour and minute as if they were at UTC.
       const dateObjectAtParsedTimeAsUTC = new Date(Date.UTC(year, monthIndex, day, hour, minute, 0, 0));
      
       // To get the actual moment in time (true UTC), we adjust this by the timezone offset.
       // If timestamp is "10:00 in +02:00 timezone", it means 10:00 local, which is 08:00 UTC.
       // So, we subtract the offset from the dateObjectAtParsedTimeAsUTC.
       // dateObjectAtParsedTimeAsUTC.getTime() gives millis for "10:00 UTC".
       // totalOffsetInMinutes * 60 * 1000 gives offset in millis.
       // Corrected UTC Millis = (Parsed Time as UTC) - (Offset from UTC)
       const correctedUTCMillis = dateObjectAtParsedTimeAsUTC.getTime() - (totalOffsetInMinutes * 60 * 1000);
       sheetTimestamp = new Date(correctedUTCMillis);

      Logger.log(`Parsed Timestamp for "${config.alertSubject}": ${sheetTimestamp.toUTCString()} (from cell value: "${cellValue}")`);

    } catch (e) {
      Logger.log(`Error parsing date string "${cellValue}" for "${config.alertSubject}": ${e.message}. Stack: ${e.stack}`);
      const errorDetails = `Original: "${cellValue}", Cleaned: "${cleanedCellValue || 'N/A'}", Parts: ${JSON.stringify(parsedDatePartsForEmail)}.`;
      MailApp.sendEmail(config.recipientEmail, `ALERT: Timestamp Parse Error - ${config.alertSubject}`, `Could not parse the timestamp string "${cellValue}" from cell ${config.cellNotation} in sheet "${config.sheetName}" (Spreadsheet: ${config.spreadsheetId}).\n\nError: ${e.message}\n${errorDetails}\n\nPlease check the format.`);
      return;
    }

    const now = new Date();
    // Logger.log(`Current Time (script execution, effectively UTC): ${now.toUTCString()}`);

    const diffMillis = now.getTime() - sheetTimestamp.getTime();
    const diffMinutes = diffMillis / (1000 * 60);
    Logger.log(`Time difference for "${config.alertSubject}": ${diffMinutes.toFixed(2)} minutes.`);

    if (diffMinutes > config.maxAgeMinutes) {
      Logger.log(`Timestamp for "${config.alertSubject}" is older than ${config.maxAgeMinutes} minutes. Sending alert.`);
      const localTimeOfTimestamp = new Date(sheetTimestamp.getTime() + (totalOffsetInMinutes * 60 * 1000));
      const offsetDisplay = parts && parts[6] ? `${parts[6]}:${parts[7] || '00'}` : 'Unknown';
      const alertBody = `ALERT: ${config.alertSubject}\n\nThe timestamp in cell ${config.cellNotation} of sheet "${config.sheetName}" (Spreadsheet ID: ${config.spreadsheetId}) is outdated.\n\nCell Content: "${cellValue}"\nParsed as Local Time (from cell): ${localTimeOfTimestamp.toString()} (Offset: ${offsetDisplay})\nEquivalent UTC: ${sheetTimestamp.toUTCString()}\n\nCurrent Time (UTC): ${now.toUTCString()}\nTime Difference: ${diffMinutes.toFixed(2)} minutes.\nAllowed Max Age: ${config.maxAgeMinutes} minutes.`;
      MailApp.sendEmail(config.recipientEmail, config.alertSubject, alertBody);
      Logger.log(`Alert email sent to ${config.recipientEmail} for "${config.alertSubject}".`);
    } else {
      Logger.log(`Timestamp for "${config.alertSubject}" is within the allowed ${config.maxAgeMinutes} minutes.`);
    }

  } catch (error) {
    Logger.log(`An unexpected error occurred while processing alert "${config.alertSubject}": ${error.message}\nStack: ${error.stack}`);
    try {
      MailApp.sendEmail(config.recipientEmail, `CRITICAL SCRIPT ERROR in: ${config.alertSubject}`, `The script encountered an unexpected error while checking for "${config.alertSubject}":\n\nError: ${error.message}\n\nSpreadsheet ID: ${config.spreadsheetId}\nSheet Name: ${config.sheetName}\nCell: ${config.cellNotation}\n\nStack Trace:\n${error.stack}`);
    } catch (mailError) {
      Logger.log(`Failed to send critical error notification email for "${config.alertSubject}": ${mailError.message}`);
    }
  }
}

/**
 * Main function to check all configured timestamps and send alerts.
 * This function should be selected in the Apps Script editor for triggers.
 */
function checkAllTimestampsAndAlerts() {
  // Configuration for the first (existing) alert
  const alert1Config = {
    spreadsheetId: '1IiI8ppxLSc0DvUbQcEBrDXk2eAExAiaA4iAfsykR8PE',
    sheetName: 'Volkscience Interview Log',
    cellNotation: 'B1',
    recipientEmail: 'pkumar@eightfold.ai', // *** IMPORTANT: Verify or replace with your actual email ID
    alertSubject: 'ALERT: AIR Data is not syncing (Volkscience Log)',
    maxAgeMinutes: 240
  };

  executeSingleTimestampCheck(alert1Config);

  // --- Add configuration for the second alert here WHEN YOU HAVE THE DETAILS ---
  // Example for a new alert (replace with actual values):
  const alert2Config = {
    spreadsheetId: '1g-Sp4_Ic91eXT9LeVwDJjRiMa5Xqf4Oks3aV29fxXRw',
    sheetName: 'Active+Rejected',
    cellNotation: 'B1',
    recipientEmail: 'pkumar@eightfold.ai',
    alertSubject: 'ALERT: Active+Rejected database not syncing!!',
    maxAgeMinutes: 120
  };
  executeSingleTimestampCheck(alert2Config);
  
  // You can add more alert configurations by defining more 'alertXConfig' objects
  // and calling executeSingleTimestampCheck(alertXConfig);
}

// If you had checkTimestampAndAlert() as your triggered function,
// ensure you update your trigger to point to checkAllTimestampsAndAlerts().
// You can also create a simple wrapper if needed, e.g.:
// function previouslyTriggeredFunction() {
//   checkAllTimestampsAndAlerts();
// }
// Make sure 'checkAllTimestampsAndAlerts' or your chosen wrapper 
// is selected in "Select function to run" in your Apps Script project triggers.