# Client Referral Automation System with Realtime Monitoring Spreadsheets

## Overview

This project is an Automated Spreadsheet Client Referral Automation System designed to streamline and automate the process of client referrals and monitoring across various hospitals. The system includes real-time updates, duplicate entry detection, and automated budget adjustments for each hospital.

## Features

- **Automated Client Referrals**: Automatically processes client referrals and updates hospital-specific spreadsheets.
- **Real-Time Monitoring**: Provides real-time updates to the main office.
- **Duplicate Detection**: Identifies and manages duplicate entries.
- **Budget Management**: Adjusts hospital budgets based on client referrals through a automated monitoring spreadsheet.
- **Data Sorting**: Sorts data by region and transfers relevant information to individual hospital sheets.
- **Referral Letter Generation**: Generates referral letters for clients.
- **Status-Based Row Deletion**: Deletes marked patients and refunds the referred budget taken from the monitoring sheets.

## Technology Stack

- `Google Drive`: Storage for all documents and templates.
- `Google Sheets`: Central and individual monitoring sheets for hospitals.
- `Google Forms`: Initial data collection from clients.
- `Google Docs`: Auto-generated referral letters.
- `Google Apps Script`: Automation of data processing and management.
- `JSON`: Directory of hospitals with their respective monitoring sheets and budgets.

## Usage

1. **Form Submission:**

- Clients submit their data via Google Forms.

2. **Data Processing:**

- Data is imported into the main spreadsheet.
- The system sorts data by region and updates individual hospital sheets.
- Duplicate entries are detected and flagged.
- Status changes trigger row deletion and budget adjustments.

3. **Referral Generation:**

- Upon approval, the system generates a referral letter for the client.
- The referral letter is sent to the client for use at the hospital.

4. **Documentation:**

- All processed data is sent to the main office for documentation and record-keeping.

## Global Variables and Sheets

- **Main Sheets:**

`main_spreadsheet`: The main spreadsheet containing all form submissions.
`monitoring_sheet`: Sheet for monitoring overall progress.

**Region Specific Sheets:**

`r1_spreadsheet`, `r2_spreadsheet`, `r3_spreadsheet`: Individual sheets for each region.
`r1_runningbalance`, `r2_runningbalance`, `r3_runningbalance`: Sheets to track running balances.
`r1_cancelled`, `r2_cancelled`, `r3_cancelled`: Sheets for cancelled entries.

**JSON Data:**

Each region has a JSON file that contains hospital data, such as IDs and control numbers.

**User Input and Data Ranges:**

`user_input`, `region`, `range`, `mainsheet_data`, `mainsheet_data_rb`: Variables to handle user inputs and data ranges for processing.

## Key Functions

- `onOpen()`: Adds a custom menu to the Google Sheets UI to sort data and sheets.
- `createHospitalMap(data, count, regionPrefix)`: Creates a map of hospital data based on the region.
- `onFormSubmit(controlNumber)`: Handles form submissions, generates control numbers, and triggers the document autofill process.
- `autofillDocument(hospitalMap, controlNumber, googleDocLink)`: Autofills a Google Doc template with client data and updates the main spreadsheet.
- `status()`: Sets the status of the submission in the main spreadsheet.
- `dataToRespectiveSheets(controlNumber, googleDocLink, valuesToPopulate_main, valuesToPopulate_rb)`: Populates the data ranges for respective regions and individual spreadsheets.
- `dataToIndividualSpreadsheet(valuesToPopulate_main, valuesToPopulate_rb)`: Populates individual hospital sheets with the data.

## Installation

To set up this project, follow these steps:

1. **Clone the Repository:**

```bash git clone https://github.com/your-username/google-sheet-automation.git```

2. **Set Up Google Apps Script:**

- Open the Google Sheet.
- Navigate to Extensions > Apps Script.
- Copy and paste the provided Google Apps Script code into the script editor.

3. **Configure Triggers:**

- Set up time-driven triggers for periodic data processing.
- Set up form submission triggers for immediate data processing.

4. **Set Up JSON Directory:**

- Ensure the JSON file contains the correct directory of hospitals and their respective monitoring sheets.

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License
This project is licensed under the MIT License.

## Contact
For any inquiries or support, please contact deanferrazzini@gmail.com.