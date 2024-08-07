function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TEMPORARY SORT DATA')
      .addItem('SORT DATA & SHEETS', 'dataToRespectiveSheets')
      .addToUi();
}

//MAINSHEETS
const main_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN SPREADSHEET");
const monitoring_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MONITORING SHEET");
//1 : SPREADSHEET, RUNNING BALANCE & CANCELLED SHEETS
const r1_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPREADSHEET 1");
const r1_runningbalance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RUNNING BALANCE 1");
const r1_cancelled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CANCELLED 1");
//2 : SPREADSHEET, RUNNING BALANCE & CANCELLED SHEETS
const r2_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPREADSHEET 2");
const r2_runningbalance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RUNNING BALANCE 2");
const r2_cancelled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CANCELLED 2");
//3 : SPREADSHEET, RUNNING BALANCE & CANCELLED SHEETS
const r3_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGION 3");
const r3_runningbalance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RUNNING BALANCE 3");
const r3_cancelled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CANCELLED 3");

//SHEETS LASTROW
const lastrow_main_spreadsheet = main_spreadsheet.getLastRow();
const lastrow_monitoring_sheet = monitoring_sheet.getLastRow();
//REGION XII
const lastrow_spreadsheet_1 = r1_spreadsheet.getLastRow();
const lastrow_runningbalance_1 = r1_runningbalance.getLastRow();
const lastrow_cancelled_1 = r1_cancelled.getLastRow();
//REGION XI
const lastrow_spreadsheet_2 = r2_spreadsheet.getLastRow();
const lastrow_runningbalance_2 = r2_runningbalance.getLastRow();
const lastrow_cancelled_2 = r2_cancelled.getLastRow();
//REGION X
const lastrow_spreadsheet_3 = r3_spreadsheet.getLastRow();
const lastrow_runningbalance_3 = r3_runningbalance.getLastRow();
const lastrow_cancelled_3 = r3_cancelled.getLastRow();


//hospitalData - Hospital Name, Hospital File ID, Hospital Control Number ID
const r1_fileId = "GOOGLE_SPREADSHEET_LINK";
const r1_file = DriveApp.getFileById(r1_fileId);
const r1_fileContent = r1_file.getBlob().getDataAsString();
const r1_data = JSON.parse(r1_fileContent);

const r2_fileId = "GOOGLE_SPREADSHEET_LINK";
const r2_file = DriveApp.getFileById(r2_fileId);
const r2_fileContent = r2_file.getBlob().getDataAsString();
const r2_data = JSON.parse(r2_fileContent);

const r3_fileId = "GOOGLE_SPREADSHEET_LINK";
const r3_file = DriveApp.getFileById(r3_fileId);
const r3_fileContent = r3_file.getBlob().getDataAsString();
const r3_data = JSON.parse(r3_fileContent);

//GLOBAL VARIABLES
const user_input = main_spreadsheet.getRange(lastrow_main_spreadsheet, 1, 1, 19).getValues()[0];
const region = main_spreadsheet.getRange(lastrow_main_spreadsheet, 2).getValues();
const range = main_spreadsheet.getRange(lastrow_main_spreadsheet, 3).getValues();
const mainsheet_data = main_spreadsheet.getRange(lastrow_main_spreadsheet, 1, 1, 19).getValues();
const mainsheet_data_rb = [user_input[4], user_input[5], user_input[13]];

//REGION X, XI, XII SPREADSHEETS DATA RANGE
const r3_dataRange_main = r3_spreadsheet.getRange(lastrow_spreadsheet_3 + 1, 1, 1, 19)
const r2_dataRange_main = r2_spreadsheet.getRange(lastrow_spreadsheet_2 + 1, 1, 1, 19)
const r1_dataRange_main = r1_spreadsheet.getRange(lastrow_spreadsheet_1 + 1, 1, 1, 19)

const r3_dataRange_rb = r3_runningbalance.getRange(lastrow_spreadsheet_3 + 1, 1, 1, 3)
const r2_dataRange_rb = r2_runningbalance.getRange(lastrow_spreadsheet_2 + 1, 1, 1, 3)
const r1_dataRange_rb = r1_runningbalance.getRange(lastrow_spreadsheet_1 + 1, 1, 1, 3)

const r3_dataRange_cancelled = r3_cancelled.getRange(lastrow_spreadsheet_3 + 1, 1, 1, 3)
const r2_dataRange_cancelled = r2_cancelled.getRange(lastrow_spreadsheet_2 + 1, 1, 1, 3)
const r1_dataRange_cancelled = r1_cancelled.getRange(lastrow_spreadsheet_1 + 1, 1, 1, 3)

function r1_datamap(){
var hospitalMap = {};
  for (var i = 1; i <= 36; i++) {
    var hospitalName = r1_data["r1_hospital" + i];
    var hospitalId = r1_data["r1_hospital" + i + "_id"];
    var controlNumber = r1_data["r1_hospital" + i + "_cn"];
    var folderID = r1_data["r1_hospital" + i + "_folderid"];
    var spreadsheetID = r1_data["r1_hospital" + i + "_wordid"];
    hospitalMap[hospitalName] = { id: hospitalId, cn: controlNumber, fid: folderID, ssid: spreadsheetID};
  }
}

function r2_datamap(){
var hospitalMap = {};
  for (var i = 1; i <= 19; i++) {
    var hospitalName = r2_data["r2_hospital" + i];
    var hospitalId = r2_data["r2_hospital" + i + "_id"];
    var controlNumber = r2_data["r2_hospital" + i + "_cn"];
    var folderID = r2_data["r2_hospital" + i + "_folderid"];
    var spreadsheetID = r2_data["r2_hospital" + i + "_wordid"];
    hospitalMap[hospitalName] = { id: hospitalId, cn: controlNumber, fid: folderID, ssid: spreadsheetID};
  }  
}

function r3_datamap(){
var hospitalMap = {};
  for (var i = 1; i <= 12; i++) {
    var hospitalName = r3_data["r3_hospital" + i];
    var hospitalId = r3_data["r3_hospital" + i + "_id"];
    var controlNumber = r3_data["r3_hospital" + i + "_cn"];
    var folderID = r3_data["r3_hospital" + i + "_folderid"];
    var spreadsheetID = r3_data["r3_hospital" + i + "_wordid"];
    hospitalMap[hospitalName] = { id: hospitalId, cn: controlNumber, fid: folderID, ssid: spreadsheetID};
  }  
}

function test(){
  Logger.log(mainsheet_data_rb)
  Logger.log(mainsheet_data_rb.flat())
}

// Function to pad numbers
function padNumber(number) {
  var paddedNumber = number.toString();
  var padLength = 5 - paddedNumber.length;
  for (var i = 0; i < padLength; i++) {
    paddedNumber = "0" + paddedNumber;
  }
  return paddedNumber;
}

function createHospitalMap(data, count, regionPrefix) {
  var hospitalMap = {};
  for (var i = 1; i <= count; i++) {
    var hospitalName = data[regionPrefix + "_hospital" + i];
    var hospitalId = data[regionPrefix + "_hospital" + i + "_id"];
    var controlNumber = data[regionPrefix + "_hospital" + i + "_cn"];
    var folderID = data[regionPrefix + "_hospital" + i + "_folderid"];
    var wordID = data[regionPrefix + "_hospital" + i + "_wordid"];
    hospitalMap[hospitalName] = { id: hospitalId, cn: controlNumber, fid: folderID, wid: wordID };
  }
  return hospitalMap;
}

function onFormSubmit(controlNumber) {
  const region = user_input[1];
  const trigger = user_input[2];

  Logger.log("Form submitted. Region: " + region + ", Trigger: " + trigger);

  var hospitalMap;
  if (region == "R3") {
    hospitalMap = createHospitalMap(r3_data, 12, "3");
  } else if (region == "R2") {
    hospitalMap = createHospitalMap(r2_data, 19, "2");
  } else if (region == "R1") {
    hospitalMap = createHospitalMap(r1_data, 36, "1");
  }

  var hospitalName = trigger;
  Logger.log("Triggered hospital: " + hospitalName);
  var hospitalData = hospitalMap[hospitalName];
  if (hospitalData) {
    var propertyKey = hospitalData.cn + '_lastNum';
    var lastNum = PropertiesService.getScriptProperties().getProperty(propertyKey) || 0;
    lastNum = parseInt(lastNum) + 1;
    var paddedLastNum = padNumber(lastNum); // Padding the lastNum using padNumber function
    var controlNumber = hospitalData.cn + "-2024-" + paddedLastNum;
    PropertiesService.getScriptProperties().setProperty(propertyKey, lastNum.toString());
    Logger.log("Control number generated: " + controlNumber);
    main_spreadsheet.getRange(main_spreadsheet.getLastRow(), 6).setValue(controlNumber);

    // Call autofillDocument() with a callback to trigger dataToRespectiveSheets() after document generation
    autofillDocument(hospitalMap, controlNumber);
  } else {
    Logger.log("No hospital data found for the triggered hospital: " + hospitalName);
  }
}

function autofillDocument(hospitalMap, controlNumber, googleDocLink) {
  var hospitalName = user_input[2];
  var hospitalData = hospitalMap[hospitalName];
  if (hospitalData) {
    var intValue = user_input[13];
    var formattedValue = Utilities.formatString('₱%s', intValue.toLocaleString('en-PH', { minimumFractionDigits: 2 }));

    var file = DriveApp.getFileById(hospitalData.wid);
    Logger.log(hospitalData.wid);
    Logger.log(hospitalData.fid);

    // Make a copy of the template and place it in the designated folder
    var folder = DriveApp.getFolderById(hospitalData.fid);
    var copy = file.makeCopy(user_input[4], folder);
    var doc = DocumentApp.openById(copy.getId());

    // Get the document body
    var body = doc.getBody();

    // Define a list of placeholders and corresponding values
    var placeholders = ['#CONNUM#', '#TYPEOFREQUEST#', '#NAMEOFCLIENT#', '#AMOUNT#', '#DATEISSUED#', '#VALIDUNTIL#'];

    // Format the date before replacing it
    var dateIssued = new Date(user_input[14]);
    var dateValid = new Date(user_input[15]);
    var formattedDateIssued = Utilities.formatDate(dateIssued, "GMT+8", "MM/dd/yyyy");
    var formattedDateValid = Utilities.formatDate(dateValid, "GMT+8", "MM/dd/yyyy");
    const control_number = main_spreadsheet.getRange(main_spreadsheet.getLastRow(), 6).getValue();
    Logger.log(control_number);
    Logger.log(user_input[3]);
    Logger.log(user_input[4]);
    Logger.log(formattedValue);
    Logger.log(formattedDateIssued);
    Logger.log(formattedDateValid);

    var values = [control_number, user_input[3], user_input[4], formattedValue, formattedDateIssued, formattedDateValid];

    // Replace all the placeholders in one batch operation
    for (var i = 0; i < placeholders.length; i++) {
      body.replaceText(placeholders[i], values[i]);
    }

    // Save and close the document
    doc.saveAndClose();

    // Update the spreadsheet with the Google Doc link
    var googleDocLink = doc.getUrl();
    Logger.log(googleDocLink);
    main_spreadsheet.getRange(lastrow_main_spreadsheet, 18).setValue(googleDocLink);

    status();
    dataToRespectiveSheets(controlNumber, googleDocLink);
  } else {
    Logger.log("No hospital data found for the triggered hospital: " + hospitalName);
  }
}

function status(){
  // Set the data validation for cell A1 to require "Yes" or "No", with no dropdown menu.
 var cell = main_spreadsheet.getRange(lastrow_main_spreadsheet, 19);
 var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Recorded', 'Released', 'Cancelled', 'Cancelled and Moved'], true).build();
 cell.setDataValidation(rule);
 cell.setValue("Recorded");
}

function dataToRespectiveSheets(controlNumber, googleDocLink, valuesToPopulate_main, valuesToPopulate_rb) {
  Logger.log(controlNumber)
  Logger.log(googleDocLink)

  const region = user_input[1];
  Logger.log("Region from user input: " + region); // Log the hospital name from user input

  // Define common values for all regions
  valuesToPopulate_main = mainsheet_data;
  valuesToPopulate_rb = [user_input[4], controlNumber, user_input[13]];

  // Insert the new value into position 18 of the array
  valuesToPopulate_main[0][5] = controlNumber; // Assuming it's a 2D array with one row
  valuesToPopulate_main[0][17] = googleDocLink; // Assuming it's a 2D array with one row
  valuesToPopulate_main[0][18] = "Recorded"; // Assuming it's a 2D array with one row

  Logger.log(valuesToPopulate_main)
  Logger.log(valuesToPopulate_rb)

  // Assign data ranges based on the region
  switch (region) {
    case "R3":
      hospitalMap = createHospitalMap(r3_data, 12);
      dataRange_main = r3_dataRange_main;
      dataRange_rb = r3_dataRange_rb;
      break;
    case "R2":
      hospitalMap = createHospitalMap(r2_data, 19);
      dataRange_main = r2_dataRange_main;
      dataRange_rb = r2_dataRange_rb;
      break;
    case "R1":
      hospitalMap = createHospitalMap(r1_data, 36);
      dataRange_main = r1_dataRange_main;
      dataRange_rb = r1_dataRange_rb;
      valuesToPopulate_rb = valuesToPopulate_rb.flat(); // Flatten the array for Region XII
      break;
    default:
      Logger.log("Invalid region specified.");
      return;
  }

  // Populate the data ranges with the values
  dataRange_main.setValues(valuesToPopulate_main);
  dataRange_rb.setValues([valuesToPopulate_rb]); // Ensure valuesToPopulate_rb is wrapped in an array

  Logger.log("Data populated successfully in the specified range.");

  // Call a function to process individual spreadsheets using the hospitalMap
  dataToIndividualSpreadsheet(valuesToPopulate_main, valuesToPopulate_rb);
}

function dataToIndividualSpreadsheet(valuesToPopulate_main, valuesToPopulate_rb) {
  var region = user_input[1];
  var hospitalMap;

  // Create hospitalMap based on region
  if (region == "R3") {
    hospitalMap = createHospitalMap(r3_data, 12, "3");
  } else if (region == "R2") {
    hospitalMap = createHospitalMap(r2_data, 19, "2");
  } else if (region == "R1") {
    hospitalMap = createHospitalMap(r1_data, 36, "1");
  } else {
    Logger.log("Region not recognized.");
    return; // Exit function if region is not recognized
  }

  var hospitalName = user_input[2];
  var hospitalData = hospitalMap[hospitalName];
  Logger.log("Hospital Name from user input: " + hospitalName); // Log the hospital name from user input

  if (hospitalData) {
    // Fetching sheets
    var hospitalSheet = SpreadsheetApp.openById(hospitalData.id).getSheetByName("HOSPITAL SHEET");
    var runningBalanceSheet = SpreadsheetApp.openById(hospitalData.id).getSheetByName("RUNNING BALANCE");

// Define data ranges
var hospitalLastRow = hospitalSheet.getLastRow() + 1;
var runningBalanceLastRow = runningBalanceSheet.getLastRow() + 1;
var hospitalDataRange = hospitalSheet.getRange(hospitalLastRow, 1, 1, 19);
var runningBalanceDataRange = runningBalanceSheet.getRange(runningBalanceLastRow, 1, 1, 3);

Logger.log(valuesToPopulate_main);
Logger.log(valuesToPopulate_rb);

// Populate the data ranges with the values using batch update
hospitalDataRange.setValues(valuesToPopulate_main);
runningBalanceDataRange.setValues([valuesToPopulate_rb]);
sortData();

    Logger.log("Data populated successfully in the specified range.");
    // sortData(); // Assuming sortData() sorts the data in the spreadsheet
  } else {
    Logger.log("Hospital data not found.");
  }
}

function sortData() {
  var values = main_spreadsheet.getRange(lastrow_main_spreadsheet, 1, 1, 19).getValues()[0];
  
  // Define hospital ranges for each region
  const hospitalRanges = {
    "R3": {
      "R3_HOSPITAL_1": { beneficiaries: "D3", hospital: "E3" },
      "R3_HOSPITAL_2": {beneficiaries: "D4", hospital: "E4"},
      "R3_HOSPITAL_3": {beneficiaries: "D5", hospital: "E5"},
      "R3_HOSPITAL_4": {beneficiaries: "D6", hospital: "E6"},
      "R3_HOSPITAL_5": {beneficiaries: "D7", hospital: "E7"},
      "R3_HOSPITAL_6": {beneficiaries: "D8", hospital: "E8"},
      "R3_HOSPITAL_7": {beneficiaries: "D9", hospital: "E9"},
      "R3_HOSPITAL_8": {beneficiaries: "D10", hospital: "E10"},
      "R3_HOSPITAL_9": {beneficiaries: "D11", hospital: "E11"},
      "R3_HOSPITAL_10": {beneficiaries: "D12", hospital: "E12"},
      "R3_HOSPITAL_11": {beneficiaries: "D13", hospital: "E13"},
      "R3_HOSPITAL_12": {beneficiaries: "D14", hospital: "E14"}
    },

    "R2": {
      "R2_HOSPITAL_13": { beneficiaries: "D17", hospital: "E17" },
      "R2_HOSPITAL_14": {beneficiaries: "D18", hospital: "E18"},
      "R2_HOSPITAL_15": {beneficiaries: "D19", hospital: "E19"},
      "R2_HOSPITAL_16": {beneficiaries: "D20", hospital: "E20"},
      "R2_HOSPITAL_17": {beneficiaries: "D21", hospital: "E21"},
      "R2_HOSPITAL_18": {beneficiaries: "D22", hospital: "E22"},
      "R2_HOSPITAL_19": {beneficiaries: "D23", hospital: "E23"},
      "R2_HOSPITAL_20": {beneficiaries: "D24", hospital: "E24"},
      "R2_HOSPITAL_21": {beneficiaries: "D25", hospital: "E25"},
      "R2_HOSPITAL_22": {beneficiaries: "D26", hospital: "E26"},
      "R2_HOSPITAL_23": {beneficiaries: "D27", hospital: "E27"},
      "R2_HOSPITAL_24": {beneficiaries: "D28", hospital: "E28"},
      "R2_HOSPITAL_25": {beneficiaries: "D29", hospital: "E29"},
      "R2_HOSPITAL_26": {beneficiaries: "D30", hospital: "E30"},
      "R2_HOSPITAL_27": {beneficiaries: "D31", hospital: "E31"},
      "R2_HOSPITAL_28": {beneficiaries: "D32", hospital: "E32"},
      "R2_HOSPITAL_29": {beneficiaries: "D33", hospital: "E33"},
      "R2_HOSPITAL_30": {beneficiaries: "D34", hospital: "E34"},
      "R2_HOSPITAL_31": {beneficiaries: "D35", hospital: "E35"}
    },

    "R1": {
      "R1_HOSPITAL_32": { beneficiaries: "D38", hospital: "E38" },
      "R1_HOSPITAL_33": {beneficiaries: "D39", hospital: "E39"},
      "R1_HOSPITAL_34": {beneficiaries: "D40", hospital: "E40"},
      "R1_HOSPITAL_35": {beneficiaries: "D42", hospital: "E42"},
      "R1_HOSPITAL_36": {beneficiaries: "D43", hospital: "E43"},
      "R1_HOSPITAL_37": {beneficiaries: "D44", hospital: "E44"},
      "R1_HOSPITAL_38": {beneficiaries: "D45", hospital: "E45"},
      "R1_HOSPITAL_39": {beneficiaries: "D46", hospital: "E46"},
      "R1_HOSPITAL_40": {beneficiaries: "D47", hospital: "E47"},
      "R1_HOSPITAL_41": {beneficiaries: "D48", hospital: "E48"},
      "R1_HOSPITAL_42": {beneficiaries: "D49", hospital: "E49"},
      "R1_HOSPITAL_43": {beneficiaries: "D50", hospital: "E50"},
      "R1_HOSPITAL_44": {beneficiaries: "D51", hospital: "E51"},
      "R1_HOSPITAL_45": {beneficiaries: "D52", hospital: "E52"},
      "R1_HOSPITAL_46": {beneficiaries: "D53", hospital: "E53"},
      "R1_HOSPITAL_47": {beneficiaries: "D54", hospital: "E54"},
      "R1_HOSPITAL_48": {beneficiaries: "D55", hospital: "E55"},
      "R1_HOSPITAL_49": {beneficiaries: "D56", hospital: "E56"},
      "R1_HOSPITAL_50": {beneficiaries: "D57", hospital: "E57"},
      "R1_HOSPITAL_51": {beneficiaries: "D58", hospital: "E58"},
      "R1_HOSPITAL_52": {beneficiaries: "D59", hospital: "E59"},
      "R1_HOSPITAL_53": {beneficiaries: "D60", hospital: "E60"},
      "R1_HOSPITAL_54": {beneficiaries: "D61", hospital: "E61"},
      "R1_HOSPITAL_55": {beneficiaries: "D62", hospital: "E62"},
      "R1_HOSPITAL_56": {beneficiaries: "D63", hospital: "E63"},
      "R1_HOSPITAL_57": {beneficiaries: "D64", hospital: "E64"},
      "R1_HOSPITAL_58": {beneficiaries: "D65", hospital: "E65"},
      "R1_HOSPITAL_59": {beneficiaries: "D66", hospital: "E66"},
      "R1_HOSPITAL_60": {beneficiaries: "D67", hospital: "E67"},
      "R1_HOSPITAL_61": {beneficiaries: "D68", hospital: "E68"},
      "R1_HOSPITAL_62": {beneficiaries: "D69", hospital: "E69"},
      "R1_HOSPITAL_63": {beneficiaries: "D70", hospital: "E70"},
      "R1_HOSPITAL_64": {beneficiaries: "D71", hospital: "E71"},
      "R1_HOSPITAL_65": {beneficiaries: "D72", hospital: "E72"},
      "R1_HOSPITAL_66": {beneficiaries: "D73", hospital: "E73"}
    }
  };

  // Get hospital ranges based on the region
  const regionHospitals = hospitalRanges[region];
  if (!regionHospitals) {
    Logger.log("No hospital data found for the region: " + region);
    return;
  }

  var hospitalName = values[2];
  if (hospitalName in regionHospitals) {
    Logger.log("Processing data for hospital: " + hospitalName);

    var beneficiariesRange = monitoring_sheet.getRange(regionHospitals[hospitalName].beneficiaries);
    var beneficiariesAmount = beneficiariesRange.getValue() + 1;
    beneficiariesRange.setValue(beneficiariesAmount);

    var hospitalRange = monitoring_sheet.getRange(regionHospitals[hospitalName].hospital);
    var hospitalAmount = hospitalRange.getValue() + values[13];
    hospitalRange.setValue(hospitalAmount);

  // sortDataToCentral();
  }
}

function DuplicatePatient() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN SPREADSHEET");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("E2:E" + lastRow); // Update the range to B2:B
  var values = range.getDisplayValues();
  var nameCounts = {};
  for (var i = 0; i < values.length; i++) {
    var name = values[i][0];
    if (name in nameCounts) {
      nameCounts[name]++;
    } else {
      nameCounts[name] = 1;
    }
  }
  var ranges = [];
  for (var i = 0; i < values.length; i++) {
    var name = values[i][0];
    if (nameCounts[name] > 1) {
      ranges.push("E" + (i + 2));
    }
  }
  if (ranges.length == 0) return;
  Logger.log(values);
  Logger.log(nameCounts);
  var earliestRow = Number.MAX_VALUE;
  var earliestTimestamp = new Date();
  var rows = ranges.map(function(range) {
    var row = parseInt(range.replace("E", ""), 10);
    var timestamp = sheet.getRange(row, 1).getValue();
    if (timestamp < earliestTimestamp) {
      earliestTimestamp = timestamp;
      earliestRow = row;
    }
    return row;
  });
  var timestampRange = sheet.getRange(earliestRow, 1);
  var note = "⚠️ Duplicates found on rows " + rows.join(", ") + ". Earliest recorded duplicate was in " + timestampRange.getA1Notation() + " on " + earliestTimestamp;
  sheet.getRangeList(ranges).setNote(note);
}

function logMessage(message) {
  console.log(new Date().toISOString(), message);
}


function moveToCancelled() {
  var triggerCancel = "Cancelled";
  var newTriggerCancel = triggerCancel + " and Moved";
  var red = "#FF0000";

  logMessage("moveToCancelled function started.");

  const hospitalRanges = {
    "R3": {
      "R3_HOSPITAL_1": { beneficiaries: "D3", hospital: "E3" },
      "R3_HOSPITAL_2": {beneficiaries: "D4", hospital: "E4"},
      "R3_HOSPITAL_3": {beneficiaries: "D5", hospital: "E5"},
      "R3_HOSPITAL_4": {beneficiaries: "D6", hospital: "E6"},
      "R3_HOSPITAL_5": {beneficiaries: "D7", hospital: "E7"},
      "R3_HOSPITAL_6": {beneficiaries: "D8", hospital: "E8"},
      "R3_HOSPITAL_7": {beneficiaries: "D9", hospital: "E9"},
      "R3_HOSPITAL_8": {beneficiaries: "D10", hospital: "E10"},
      "R3_HOSPITAL_9": {beneficiaries: "D11", hospital: "E11"},
      "R3_HOSPITAL_10": {beneficiaries: "D12", hospital: "E12"},
      "R3_HOSPITAL_11": {beneficiaries: "D13", hospital: "E13"},
      "R3_HOSPITAL_12": {beneficiaries: "D14", hospital: "E14"}
    },

    "R2": {
      "R2_HOSPITAL_13": { beneficiaries: "D17", hospital: "E17" },
      "R2_HOSPITAL_14": {beneficiaries: "D18", hospital: "E18"},
      "R2_HOSPITAL_15": {beneficiaries: "D19", hospital: "E19"},
      "R2_HOSPITAL_16": {beneficiaries: "D20", hospital: "E20"},
      "R2_HOSPITAL_17": {beneficiaries: "D21", hospital: "E21"},
      "R2_HOSPITAL_18": {beneficiaries: "D22", hospital: "E22"},
      "R2_HOSPITAL_19": {beneficiaries: "D23", hospital: "E23"},
      "R2_HOSPITAL_20": {beneficiaries: "D24", hospital: "E24"},
      "R2_HOSPITAL_21": {beneficiaries: "D25", hospital: "E25"},
      "R2_HOSPITAL_22": {beneficiaries: "D26", hospital: "E26"},
      "R2_HOSPITAL_23": {beneficiaries: "D27", hospital: "E27"},
      "R2_HOSPITAL_24": {beneficiaries: "D28", hospital: "E28"},
      "R2_HOSPITAL_25": {beneficiaries: "D29", hospital: "E29"},
      "R2_HOSPITAL_26": {beneficiaries: "D30", hospital: "E30"},
      "R2_HOSPITAL_27": {beneficiaries: "D31", hospital: "E31"},
      "R2_HOSPITAL_28": {beneficiaries: "D32", hospital: "E32"},
      "R2_HOSPITAL_29": {beneficiaries: "D33", hospital: "E33"},
      "R2_HOSPITAL_30": {beneficiaries: "D34", hospital: "E34"},
      "R2_HOSPITAL_31": {beneficiaries: "D35", hospital: "E35"}
    },

    "R1": {
      "R1_HOSPITAL_32": { beneficiaries: "D38", hospital: "E38" },
      "R1_HOSPITAL_33": {beneficiaries: "D39", hospital: "E39"},
      "R1_HOSPITAL_34": {beneficiaries: "D40", hospital: "E40"},
      "R1_HOSPITAL_35": {beneficiaries: "D42", hospital: "E42"},
      "R1_HOSPITAL_36": {beneficiaries: "D43", hospital: "E43"},
      "R1_HOSPITAL_37": {beneficiaries: "D44", hospital: "E44"},
      "R1_HOSPITAL_38": {beneficiaries: "D45", hospital: "E45"},
      "R1_HOSPITAL_39": {beneficiaries: "D46", hospital: "E46"},
      "R1_HOSPITAL_40": {beneficiaries: "D47", hospital: "E47"},
      "R1_HOSPITAL_41": {beneficiaries: "D48", hospital: "E48"},
      "R1_HOSPITAL_42": {beneficiaries: "D49", hospital: "E49"},
      "R1_HOSPITAL_43": {beneficiaries: "D50", hospital: "E50"},
      "R1_HOSPITAL_44": {beneficiaries: "D51", hospital: "E51"},
      "R1_HOSPITAL_45": {beneficiaries: "D52", hospital: "E52"},
      "R1_HOSPITAL_46": {beneficiaries: "D53", hospital: "E53"},
      "R1_HOSPITAL_47": {beneficiaries: "D54", hospital: "E54"},
      "R1_HOSPITAL_48": {beneficiaries: "D55", hospital: "E55"},
      "R1_HOSPITAL_49": {beneficiaries: "D56", hospital: "E56"},
      "R1_HOSPITAL_50": {beneficiaries: "D57", hospital: "E57"},
      "R1_HOSPITAL_51": {beneficiaries: "D58", hospital: "E58"},
      "R1_HOSPITAL_52": {beneficiaries: "D59", hospital: "E59"},
      "R1_HOSPITAL_53": {beneficiaries: "D60", hospital: "E60"},
      "R1_HOSPITAL_54": {beneficiaries: "D61", hospital: "E61"},
      "R1_HOSPITAL_55": {beneficiaries: "D62", hospital: "E62"},
      "R1_HOSPITAL_56": {beneficiaries: "D63", hospital: "E63"},
      "R1_HOSPITAL_57": {beneficiaries: "D64", hospital: "E64"},
      "R1_HOSPITAL_58": {beneficiaries: "D65", hospital: "E65"},
      "R1_HOSPITAL_59": {beneficiaries: "D66", hospital: "E66"},
      "R1_HOSPITAL_60": {beneficiaries: "D67", hospital: "E67"},
      "R1_HOSPITAL_61": {beneficiaries: "D68", hospital: "E68"},
      "R1_HOSPITAL_62": {beneficiaries: "D69", hospital: "E69"},
      "R1_HOSPITAL_63": {beneficiaries: "D70", hospital: "E70"},
      "R1_HOSPITAL_64": {beneficiaries: "D71", hospital: "E71"},
      "R1_HOSPITAL_65": {beneficiaries: "D72", hospital: "E72"},
      "R1_HOSPITAL_66": {beneficiaries: "D73", hospital: "E73"}
    }
  };

  var values = main_spreadsheet.getRange(2, 1, main_spreadsheet.getLastRow() - 1, 19).getValues();

  logMessage(`Fetched ${values.length} rows from MAIN SPREADSHEET.`);

  values.forEach((row, index) => {
    var region = row[1]; // Assuming the region is in column A
    var hospitalName = row[2]; // Assuming the hospital name is in column B

    if (row[18] === triggerCancel && hospitalName in hospitalRanges[region]) {
      var hospitalData = hospitalRanges[region][hospitalName];

      // Get monitoring sheet references
      var monitoringSheet = monitoring_sheet; // Assuming monitoring_sheet is defined globally
      var beneficiariesRange = monitoringSheet.getRange(hospitalData.beneficiaries);
      var hospitalAmountRange = monitoringSheet.getRange(hospitalData.hospital);

      logMessage(`Processing cancellation for ${hospitalName} in ${region}`);

      try {
        // Update beneficiaries count
        var beneficiariesAmount = beneficiariesRange.getValue() - 1;
        beneficiariesRange.setValue(beneficiariesAmount);

        // Update hospital amount
        var hospitalAmount = hospitalAmountRange.getValue() - row[13];
        hospitalAmountRange.setValue(hospitalAmount);

        // Prepare data to move to cancelled sheet
        var patientName = row[3]; // Assuming Patient Name is in column D
        var controlNumber = row[4]; // Assuming Control Number is in column E
        var amount = row[13]; // Assuming Amount is in column N

        // Move canceled patient to the appropriate Cancelled sheet
        var cancelledSheet;
        switch (region) {
          case "R3":
            cancelledSheet = r3_cancelled; // Assuming r3_cancelled is defined globally
            break;
          case "R2":
            cancelledSheet = r2_cancelled; // Assuming r2_cancelled is defined globally
            break;
          case "R1":
            cancelledSheet = r1_cancelled; // Assuming r1_cancelled is defined globally
            break;
          default:
            logMessage(`Region ${region} not recognized for cancellation.`);
            return; // Skip processing if region not recognized
        }

        // Set new trigger and highlight row
        main_spreadsheet.getRange(index + 2, 19).setValue(newTriggerCancel);
        main_spreadsheet.getRange(index + 2, 1, 1, main_spreadsheet.getLastColumn()).setBackground(red);

        // Append the relevant data to cancelled sheet
        cancelledSheet.appendRow([patientName, controlNumber, amount]);

        logMessage(`Cancelled patient data (${patientName}, ${controlNumber}, ${amount}) moved successfully to ${cancelledSheet.getName()}.`);

      } catch (error) {
        logMessage(`Error processing cancellation for ${hospitalName}: ${error}`);
      }
    }
  });

  logMessage("moveToCancelled function completed.");
}


function extractNameParts(fullName) {
  // Trim any leading or trailing whitespace
  fullName = fullName.trim();

  // Split by comma to get LastName and the rest
  var parts = fullName.split(", ");
  
  var lastName = parts[0]; // The part before the comma is the LastName
  
  // The part after the comma contains FirstName, MiddleName, and MiddleInitial
  var firstNameMiddleNameAndInitial = parts[1];
  
  // Split the FirstName, MiddleName, and MiddleInitial by space
  var nameParts = firstNameMiddleNameAndInitial.split(" ");

  var firstName = ""; // Initialize firstName as an empty string
  var middleName = "";
  var middleInitial = "";

  // If there are more than 1 part after LastName
  if (nameParts.length > 1) {
    // Iterate through all parts except the last one (which is MiddleInitial)
    for (var i = 0; i < nameParts.length - 1; i++) {
      // Concatenate parts to form the FirstName
      firstName += nameParts[i] + " ";
    }
    
    // Trim any extra spaces from the end of firstName
    firstName = firstName.trim();
    
    // MiddleInitial is the last part with dot
    middleInitial = nameParts[nameParts.length - 1].replace(".", "");
  } else {
    // If only one part, it's considered as FirstName
    firstName = nameParts[0];
  }

  return {
    firstName: firstName,
    middleName: middleName,
    lastName: lastName,
    middleInitial: middleInitial
  };
}

function sortDataToCentral() {
  // Get the "MAIN SPREADSHEET" sheet
  const main_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN SPREADSHEET");
  Logger.log("Opened MAIN SPREADSHEET.");

  const davao_spreadsheet = SpreadsheetApp.openById("GOOGLE_SPREADSHEET_LINK").getSheetByName("PLACE");

  // Assuming we want the last row from MAIN SPREADSHEET
  const lastRow = main_spreadsheet.getLastRow();
  Logger.log("Last row number in MAIN SPREADSHEET: " + lastRow);

  // Get the range from the last row, all 19 columns
  const values = main_spreadsheet.getRange(lastRow, 1, 1, 19).getValues()[0];
  Logger.log("Values from the last row in MAIN SPREADSHEET: " + values);

  var region = values[1];
  Logger.log("Region: " + region);

  var hospitalName = values[2];
  Logger.log("Hospital name: " + hospitalName);

  var typeOfRequest = values[3];
  Logger.log("Type of Request: " + typeOfRequest);

  var patientName = values[4]; // Assume the name is in the 5th column (index 4)
  Logger.log("Patient name: " + patientName);

  var controlNumber = values[5];
  Logger.log("Control Number: " + controlNumber);

  var gender = values[6];
  Logger.log("Sex: " + gender);

  var age = values[7];
  Logger.log("Age: " + age);

  var contactNumber = (" " + values[9]);
  Logger.log("Contact Number: " + contactNumber);

  var address = values[10];
  Logger.log("Address: " + address);

  var diagnosis = values[11];
  Logger.log("Diagnosis: " + diagnosis);

  var monthlyIncome = values[12];
  Logger.log("Monthly Income: " + monthlyIncome);

  var amountGiven = values[13];
  Logger.log("Amount Given: " + amountGiven);

  var status = values[18];
  Logger.log("Status: " + status);

  // Date processing
  var dateProcessed = new Date(values[0]);
  var dateOfBirth = new Date(values[8]);
  var dateIssued = new Date(values[14]);
  var dateValid = new Date(values[15]);

  // Check if the dates are valid
  if (isNaN(dateProcessed.getTime())) {
    Logger.log("Invalid dateProcessed value: " + values[0]);
  } 
  if (isNaN(dateOfBirth.getTime())) {
    Logger.log("Invalid dateProcessed value: " + values[8]);
  }
  if (isNaN(dateIssued.getTime())) {
    Logger.log("Invalid dateIssued value: " + values[14]);
  }
  if (isNaN(dateValid.getTime())) {
    Logger.log("Invalid dateValid value: " + values[15]);
  }

  var formattedDateProcessed = Utilities.formatDate(dateProcessed, "GMT+8", "MM/dd/yyyy");
  var formattedDateofBirth = Utilities.formatDate(dateOfBirth, "GMT+8", "MM/dd/yyyy");
  var formattedDateIssued = Utilities.formatDate(dateIssued, "GMT+8", "MM/dd/yyyy");
  var formattedDateValid = Utilities.formatDate(dateValid, "GMT+8", "MM/dd/yyyy");

  Logger.log("Formatted Date Processed: " + formattedDateProcessed);
  Logger.log("Formatted Date of Birth: " + formattedDateofBirth);
  Logger.log("Formatted Date Issued: " + formattedDateIssued);
  Logger.log("Formatted Date Valid: " + formattedDateValid);

  // Check if patientName is defined
  if (patientName) {
    // Split the name into components using the extractNameParts function
    var nameParts = extractNameParts(patientName);
      
    var firstName = nameParts.firstName;
    var middleName = nameParts.middleName;
    var lastName = nameParts.lastName;
    var middleInitial = nameParts.middleInitial;

    // Define the column order in the target spreadsheet
    var columnOrder = [
      1,  // Column 1: Number Series (incremented each run)
      2,  // Column 2: Control Number
      6,  // Column 3: Gender
      4,  // Column 4: Formatted Date Processed
      5,  // Column 5: First Name
      6,  // Column 6: Middle Initial
      7,  // Column 7: Last Name
      8,  // Column 8: Age
      9,  // Column 9: Formatted Date of Birth
      10, // Column 10: Address
      11, // Column 11: Contact Number
      12, // Column 12: Diagnosis
      13, // Column 13: Monthly Income
      14, // Column 14: Region
      15, // Column 15: Hospital Name
      16, // Column 16: Type of Request
      17, // Column 17: Amount Given
      18, // Column 18: Formatted Date Issued
      // Column 19: None (leave blank)
      20  // Column 20: Status
    ];

    // Increment Number Series starting from 1
    var numberSeries = lastRow;

    // Prepare data array based on columnOrder
    var rowData = [
      [numberSeries, controlNumber, gender, formattedDateProcessed, firstName, middleInitial, lastName, age, formattedDateofBirth, address, contactNumber, diagnosis, monthlyIncome, region, hospitalName, typeOfRequest, amountGiven, formattedDateIssued, '', status]
    ];

    // Append data to the target sheet
    davao_spreadsheet.getRange(davao_spreadsheet.getLastRow() + 1, 1, 1, rowData[0].length).setValues(rowData);

    Logger.log("Data transferred successfully to target spreadsheet.");
  } else {
    Logger.log("Error: Patient name is undefined.");
  }
}


//devtool to reset properties
function resetProperty() {
  var jsonFile = DriveApp.getFileById('1QLa5jDZsuIO5jwv8PddMw5sB5_SXV29-');
  var jsonString = jsonFile.getBlob().getDataAsString();
  var json = JSON.parse(jsonString);
  Logger.log(jsonString);

  var hospitalMap = {};
  for (var i = 1; i <= 17; i++) {
    var hospitalName = json["rxi_hospital" + i];
    var hospitalId = json["rxi_hospital" + i + "_id"];
    var controlNumber = json["rxi_hospital" + i + "_cn"];
    hospitalMap[hospitalName] = { id: hospitalId, cn: controlNumber, lastNum: 0 };

    var propertyKey = controlNumber + '_lastNum';
    PropertiesService.getScriptProperties().setProperty(propertyKey, '0');
    Logger.log(PropertiesService.getScriptProperties().getProperty(propertyKey));
  }
}
