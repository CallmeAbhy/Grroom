function searchDataAndWriteToFormResponse() {
  Logger.log("Executing searchDataAndWriteToFormResponse");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formReferenceSheet = spreadsheet.getSheetByName('Form');
  var postAugustSheet = spreadsheet.getSheetByName('August');
  var formResponseSheet = spreadsheet.getSheetByName('FormResponse');

  // Get user input from the Form Reference sheet
  var userInput = getUserInput(formReferenceSheet);

  // Search for matching data in the Post August sheet
  var matchingData = searchPostAugust(userInput, postAugustSheet);

  // Write matching data to the Form Responses sheet
  writeToFormResponse(matchingData, formResponseSheet);
  Logger.log("Finished searchDataAndWriteToFormResponse");
}
function getUserInput(sheet) {
  // // Assuming the input parameters are in columns B to F
  // var userInputRange = sheet.getRange(sheet.getLastRow(), 2, 1, 5);
  // var userInputValues = userInputRange.getValues();

  // // Assuming the headers are "Customer", "Quality", "Thickness", "Density", "Track No."
  // var headers = ["Number of Sheets","Customer", "Quality", "Thickness", "Density", "Track No"];
  // var userInputData = userInputValues[0];

  // // Creating an object with parameter names as keys and user input values as values
  // var userInput = {};
  // for (var i = 0; i < headers.length; i++) {
  //   userInput[headers[i]] = userInputData[i];
  // }

  // return userInput;

  // Get the last row with data in the sheet
  try {
  var lastRow = sheet.getLastRow();

  // Assuming the input parameters are in columns B to F
  var userInputRange = sheet.getRange(lastRow, 2, 1, 5);
  var userInputValues = userInputRange.getValues();

  // Assuming the headers are "Number of Sheets", "Customer", "Quality", "Thickness", "Density", "Track No."
  var headers = ["Number of Sheets", "Customer", "Quality", "Thickness", "Density"];
  
  // If there are no responses yet, return an empty object
  if (userInputValues.length === 0) {
    return {};
  }

  // Get the user input data from the last row
  var userInputData = userInputValues[0];

  // Creating an object with parameter names as keys and user input values as values
  var userInput = {};
  for (var i = 0; i < headers.length; i++) {
    userInput[headers[i]] = userInputData[i];
  }

  return userInput;
  } catch(error)
  {
    Logger.log("Error in getUserInput: " + error.toString());
    return {};
  }
}

function searchPostAugust(userInput, sheet) {
  // Assuming headers are in the first row of the Post August sheet
  try {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var matchingData = [];

  // Map abbreviations to full headers
  var headerMappings = {
    "THK": "Thickness",
    "QUAL": "Quality",
    "DEN": "Density",
    "CUST": "Customer",
    "TRACK NO.": "Track No."
  };

  // Column indexes in Post August sheet
  var thicknessColumnIndex = headers.indexOf("THK") + 1;
  var qualityColumnIndex = headers.indexOf("QUAL") + 1;
  var densityColumnIndex = headers.indexOf("DEN") + 1;
  var customerColumnIndex = headers.indexOf("CUST") + 1;
  var trackNoColumnIndex = headers.indexOf("TRACK NO.") + 1;

  // Iterate through the rows in the Post August sheet
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Check if the row matches the user input parameters
    if (
      row[thicknessColumnIndex - 1] == userInput["Thickness"] &&
      row[qualityColumnIndex - 1] == userInput["Quality"] &&
      row[densityColumnIndex - 1] == userInput["Density"] &&
      row[customerColumnIndex - 1] == userInput["Customer"]
    ) {
      // Collect tracking number for each matching row
      var trackingNumber = row[trackNoColumnIndex - 1];

      // Repeat the row based on the "Number of Sheets" parameter
      for (var j = 0; j < userInput["Number of Sheets"]; j++) {
        // Add the whole row along with tracking number to matchingData
        matchingData.push({ row: row, trackingNumber: trackingNumber });
      }
    }
  }

  return matchingData;
  } catch (error) {
    Logger.log("Error in searchPostAugust: " + error.toString());
    // Optionally, you can rethrow the error to stop execution or handle it accordingly
    // throw error;
    return [];
  }
}
function writeToFormResponse(matchingData, sheet) {
  // Assuming headers are in the first row of the FormResponse sheet
  try {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Column indexes in FormResponse sheet
  var customerColumnIndex = headers.indexOf("Customer") + 1;
  var qualityColumnIndex = headers.indexOf("Quality") + 1;
  var thicknessColumnIndex = headers.indexOf("Thickness") + 1;
  var densityColumnIndex = headers.indexOf("Density") + 1;
  var trackNoColumnIndex = headers.indexOf("Track No") + 1;

  // Iterate through the matching data
  for (var i = 0; i < matchingData.length; i++) {
    var data = matchingData[i];

    // Append the row to the FormResponse sheet
    sheet.appendRow([
      data.row[customerColumnIndex - 1], // Customer
      data.row[qualityColumnIndex - 1], // Quality
      data.row[thicknessColumnIndex - 1], // Thickness
      data.row[densityColumnIndex - 1], // Density
      data.trackingNumber // Track No
    ]);
  }
  }
  catch(error){
    Logger.log("Error in writeToFormResponse: " + error.toString());
  }
}

