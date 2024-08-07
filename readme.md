Client Referral Automation System with Realtime Monitoring Spreadsheets
Overview
This project is an Automated Spreadsheet Client Referral Automation System designed to streamline and automate the process of client referrals and monitoring across various hospitals. The system includes real-time updates, duplicate entry detection, and automated budget adjustments for each hospital.

Features
Automated Client Referrals: Automatically processes client referrals and updates hospital-specific spreadsheets.
Real-Time Monitoring: Provides real-time updates to the main office.
Duplicate Detection: Identifies and manages duplicate entries.
Budget Management: Adjusts hospital budgets based on client referrals.
Data Sorting: Sorts data by region and transfers relevant information to individual hospital sheets.
Referral Letter Generation: Generates referral letters for clients.
Multi-language Support: Supports English and Filipino.
User-Friendly Interface: Mobile-friendly design with centered and justified text.
Technologies Used
HTML
CSS
JavaScript
Google Apps Script
Google Docs
Google Sheets
Google Drive
Google Forms
Installation
Clone the Repository:
```bash git clone https://github.com/yourusername/your-repo-name.git```
Set Up Google Sheets:
Create a main spreadsheet and individual hospital spreadsheets.
Link the spreadsheets with the appropriate IDs in the script.
Deploy Google Apps Script:
Open the Google Apps Script editor.
Paste the code from the code.gs file.
Deploy the script as a web app.
Usage
Language Selection:
Users can select their preferred language (English or Filipino) on the initial screen.
Client Profile:
Fill in the client profile information.
Feedback Questions:
Complete the feedback questions.
Acknowledgement:
Review the acknowledgement page.
Code Overview
Main Functions
handleGet: Handles GET requests and serves the HTML content.
submitForm: Submits client profile data to the 'CLIENT PROFILE SHEET'.
submitFeedback: Submits feedback data to the 'CSM SHEET'.
sortDataToCentral: Sorts and transfers data from the main spreadsheet to regional hospital sheets.
resetProperty: Resets script properties for hospital control numbers.
Data Processing
Duplicate Detection: Ensures no duplicate entries are recorded.
Budget Management: Updates hospital budgets based on client referrals.
Data Sorting: Transfers relevant data to individual hospital sheets.
Utilities
extractNameParts: Extracts and formats client names.
formatDate: Formats dates to 'MM/dd/yyyy'.
Contributing
Contributions are welcome! Please fork the repository and create a pull request with your changes.

License
This project is licensed under the MIT License.

Contact
For any inquiries or support, please contact deanferrazzini@gmail.com.