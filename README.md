# Automated Daily Activity Tracker Mail

This Google Apps Script automates the process of sending a daily activity report based on data from your Google Sheet. It fetches, formats, and emails the report on a schedule you define.

## Features

- **Automated Reporting:** Eliminates manual report creation by automatically gathering data from your Google Sheet.
- **Customizable Data:**  Tailor the report by specifying the sheet, data range, and email recipients.
- **Dynamic Column Hiding:**  Automatically hides columns in the report if they have no data, creating a cleaner output.
- **Scheduled Delivery:** Sets a daily time for the script to run and send the report, ensuring timely updates.

## Setup and Configuration

1. **Clone the Repository:** `git clone https://github.com/your-username/your-repository-name.git`
2. **Open Google Apps Script:**
   - Open your Google Sheet.
   - Go to "Extensions" > "Apps Script".
3. **Paste and Configure Code:**
   - Copy the code from `sendDailyActivityReport.gs` and paste it into the Apps Script editor.
   - **Essential:** Replace the following placeholders in the code:
      -  `'PLACEHOLDER_EMAIL_1'`, `'PLACEHOLDER_EMAIL_2'`, etc.: Your recipient email addresses.
      -  `'PLACEHOLDER_SENDER_EMAIL'`: Your sender email address.
      -  `'YOUR_SPREADSHEET_ID'`: The ID of your Google Sheet. 
      -  `'YOUR_SHEET_NAME'`: The name of the sheet in your spreadsheet.
4. **Set Up Trigger:**
   - In the Apps Script editor, run the `runDailyReport` function once.
   - This will prompt you to authorize the script to access your spreadsheet and send emails. The function also creates a time-based trigger to run the report daily at a specified time.

## Technologies Used

- Google Apps Script
- Google Sheets API
- MailApp (for sending emails)
