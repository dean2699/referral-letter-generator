Client Referral Automation System
with Realtime Monitoring Spreadsheets for Hospital Budget.
This project is a comprehensive web application designed to automate client referrals and feedback collection across multiple hospitals. It leverages Google Apps Script, Google Sheets, Google Docs, and a JSON file to manage and process data efficiently. The system provides real-time updates and automates report generation, streamlining the workflow for hospital staff and the main office.

Table of Contents
Features
Setup
Usage
File Descriptions
Technical Details
Screenshots
Contributing
License
Features

Automated Data Sorting: Automatically sorts client referral data into specific hospital sheets.
Duplicate Detection: Identifies and handles duplicate entries in the data.
Budget Management: Tracks and adjusts budgets for each hospital based on client referrals.
Referral Letter Generation: Automatically generates referral letters for clients to use at hospitals.
Real-Time Updates: Sends updates to the main office in real-time.

Setup
Prerequisites
A Google account with access to Google Sheets, Google Docs, and Google Drive.
Basic knowledge of Google Apps Script.

Installation
Clone the repository:
```bash git clone https://github.com/yourusername/referral-automation-system.git```

Set up Google Apps Script:

Open Google Sheets and create a new spreadsheet.
Go to Extensions > Apps Script.
Copy and paste the code from the code.gs file into the script editor.
Save and name the project.
Configure JSON File:

Upload the JSON file containing hospital information to Google Drive.
Note the file ID for use in the script.
Set up Properties:

Go to File > Project properties > Script properties.
Add script properties for each hospital with the following format:
hospital_name_id: The ID of the hospital's Google Sheet.
hospital_name_cn: The control number prefix for the hospital.
hospital_name_lastNum: The last used control number.
Usage
Open the Web Application:

Launch the web application from the script editor by clicking Deploy > Test deployments > Select type: Web app.
Fill Out Client Profile:

Input client details, including personal information and referral specifics.
Submit Feedback:

Collect feedback from clients regarding hospital services.
Process Referrals:

Automatically sorts data into the respective hospital sheets and generates referral letters.
Real-Time Updates:

Monitor updates in real-time from the main office.
File Descriptions
code.gs: Contains the main Google Apps Script code for the web application.
2024_R1.json: Contains the links of R1s the Referral Letter Template, Individual Spreadsheet, and Folders for Organization.
2024_R2.json: Contains the links of R2s the Referral Letter Template, Individual Spreadsheet, and Folders for Organization.
2024_R3.json: Contains the links of R3s the Referral Letter Template, Individual Spreadsheet, and Folders for Organization.
index.html: The HTML file for the web application interface.
style.css: The CSS file for styling the web application.
script.js: The JavaScript file for handling client-side logic.
README.md: This file, providing an overview of the project.
Technical Details
Google Apps Script: Used for backend automation and interaction with Google Sheets and Google Docs.
HTML/CSS/JavaScript: Used for creating a user-friendly web interface.
JSON: Used for storing and retrieving hospital information and control numbers.
Screenshots
Include screenshots of the web application interface here.

Contributing
Contributions are welcome! Please fork the repository and submit pull requests for any improvements or bug fixes.

License
This project is licensed under the MIT License - see the LICENSE file for details.