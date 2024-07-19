// Function to send the daily activity report
function sendDailyActivityReport() {
  // Email addresses
  var recipientEmails = [
    'PLACEHOLDER_EMAIL_1', // Replace with actual recipient emails
    'PLACEHOLDER_EMAIL_2', 
    // ... add other recipients ...
  ];

  var senderEmail = 'PLACEHOLDER_SENDER_EMAIL'; // Replace with your actual sender email

  // Get today's date and format it
  var today = new Date();
  var day = today.getDate();
  var month = today.getMonth() + 1; // Months are zero-indexed
  var monthName = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][month - 1];
  var todayString = Utilities.formatDate(today, 'GMT', 'dd-MM-yyyy');

  // Access the Google Sheet data 
  var spreadsheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID'); // Replace with your spreadsheet ID
  var sheet = spreadsheet.getSheetByName('YOUR_SHEET_NAME'); // Replace with your sheet name
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // Get data from row 2 onwards, adjust as needed

  // Filter data for today's activities
  var todaysActivities = data.filter(function(row) {
    return row[0].toString() === todayString;
  });

  // Sort the activities: 
  // 1. By Activity (ascending)
  // 2. Then, by Processor Name within each Activity group (ascending)
  todaysActivities.sort(function(a, b) {
    if (a[3] === b[3]) { // If activities are the same
      return a[1].localeCompare(b[1]); // Sort by Processor Name
    }
    return a[3].localeCompare(b[3]); // Sort by Activity
  });

 // Convert 'Processor Name' and 'Industry' to proper case
  todaysActivities.forEach(function(row) {
    row[1] = row[1].toString().replace(/\w\S*/g, function(txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    }); 
    row[2] = row[2].toString().replace(/\w\S*/g, function(txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    }); 
  });


  // Check if columns 'Allocation Count', 'Completed Count', 'Completed Count w/ Variant' have values
  var hasAllocationCount = todaysActivities.some(row => row[4] !== "");
  var hasCompletedCount = todaysActivities.some(row => row[5] !== "");
  var hasCompletedWithVariant = todaysActivities.some(row => row[6] !== "");
  var hasPendingCount = hasAllocationCount || hasCompletedCount || hasCompletedWithVariant; // Only include Pending Count if other count columns have data

  // Create email subject and initialize body
  var subject = 'Activity of ' + monthName + ' ' + day;
  var body = '\n\n'; 

  // Only send an email if there are activities for today
  if (todaysActivities.length > 0) {
    // Format data into an HTML table
    body += '<table style="border-collapse: collapse; width: 80%; font-family: Calibri; font-size: 11pt; border: 1px solid black;">';
    body += '<tr>';
    body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Date</th>';
    body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Processor Name</th>';
    body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Industry</th>';
    body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Activity</th>';
    if (hasAllocationCount) {
      body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Allocation Count</th>';
    }
    if (hasCompletedCount) {
      body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Completed Count</th>';
    }
    if (hasCompletedWithVariant) {
      body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Completed Count w/ Variant</th>';
    }
    if (hasPendingCount) {
      body += '<th style="border: 1px solid black; padding: 3px; text-align: center;">Pending Count</th>';
    }
    body += '</tr>';

    // Add data rows to the table (no merging of cells)
    todaysActivities.forEach(function(row) {
      body += '<tr>';
      body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[0].toString() + '</td>'; // Date
      body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[1].toString() + '</td>'; // Processor Name
      body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[2].toString() + '</td>'; // Industry
      body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[3].toString() + '</td>'; // Activity

      if (hasAllocationCount) {
        body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[4].toString() + '</td>';
      }
      if (hasCompletedCount) {
        body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[5].toString() + '</td>';
      }
      if (hasCompletedWithVariant) {
        body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[6].toString() + '</td>';
      }
      if (hasPendingCount) {
        body += '<td style="border: 1px solid black; padding: 3px; text-align: center; white-space: nowrap;">' + row[7].toString() + '</td>'; 
      }
      body += '</tr>';
    });
    
    body += '</table>';

    // Send the email
    MailApp.sendEmail({
      to: recipientEmails.join(', '),
      subject: subject,
      htmlBody: body,
      from: senderEmail,
    });
  } 
}

// Function to set up a daily trigger to run the report
function runDailyReport() {
  ScriptApp.newTrigger('sendDailyActivityReport')
    .timeBased()
    .atHour(20) // 8:00 PM
    .everyDays(1)
    .create();
}
