function searchDataAndWriteToFormResponse() {
  Logger.log("Executing searchDataAndWriteToFormResponse");

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = spreadsheet.getSheetByName('Form');
    var augustSheet = spreadsheet.getSheetByName('August');
    var formResponseSheet = spreadsheet.getSheetByName('FormResponse');

    // Get user input from the Form sheet
    var userInput = getUserInput(formSheet);
    Logger.log("User Input: " + JSON.stringify(userInput));

    // Search for matching data in the August sheet
    var matchingData = searchPostAugust(userInput, augustSheet);
    Logger.log("Matching Data: " + JSON.stringify(matchingData));

    // Write matching data to the Form Response sheet
    writeToFormResponse(matchingData, formResponseSheet);

    Logger.log("Finished searchDataAndWriteToFormResponse");
  } catch (error) {
    Logger.log("Error in searchDataAndWriteToFormResponse: " + error.toString());
  }
}
// Modify getUserInput to work with the "Form" sheet
function getUserInput(sheet) {
  try {
    // Get the last row with data in the Form sheet
    var lastRow = sheet.getLastRow();

    // Assuming the input parameters are in columns B to F
    var userInputRange = sheet.getRange(lastRow, 2, 1, 5);
    var userInputValues = userInputRange.getValues();

    // Assuming the headers are "Number of Sheets", "Customer", "Quality", "Thickness", "Density"
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
  } catch (error) {
    Logger.log("Error in getUserInput: " + error.toString());
    return {};
  }
}

// Modify searchPostAugust to work with the "August" sheet



// function searchPostAugust(userInput, sheet) {
//   try {
//     Logger.log("User Input: " + JSON.stringify(userInput));

//     var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

//     Logger.log("Headers in August Sheet: " + JSON.stringify(headers));

//     var matchingData = [];
//     var data = sheet.getDataRange().getValues();

//     // Column indexes in August sheet
//     var thicknessColumnIndex = headers.indexOf("THK") + 1;
//     var qualityColumnIndex = headers.indexOf("QUAL") + 1;
//     var densityColumnIndex = headers.indexOf("DEN") + 1;
//     var customerColumnIndex = headers.indexOf("CUST") + 1;
//     var trackNoColumnIndex = headers.indexOf("TRACK NO.") + 1;

//     // Counter for the number of sheets
//     var counter = userInput["Number of Sheets"];

//     // Iterate through the rows in the August sheet
//     for (var i = 2; i < data.length; i++) {  // Start from row 3 (index 2)
//       var row = data[i];

//       // Check if the row matches the user input parameters
//       if (
//         row[thicknessColumnIndex - 1] == userInput["Thickness"] &&
//         row[qualityColumnIndex - 1] == userInput["Quality"] &&
//         row[densityColumnIndex - 1] == userInput["Density"] &&
//         row[customerColumnIndex - 1] == userInput["Customer"]
//       ) {
//         // Collect tracking number for each matching row
//         var trackingNumber = row[trackNoColumnIndex - 1];

//         // Create an entry with only the necessary information
//         var matchingEntry = {
//           row: [trackingNumber, userInput["Number of Sheets"], userInput["Quality"], userInput["Density"], userInput["Customer"]],
//           trackingNumber: trackingNumber
//         };

//         Logger.log("Collecting Row: " + JSON.stringify(matchingEntry));

//         // Decrease the counter
//         counter--;

//         // If the counter becomes zero, add the entry to matchingData and break out of the loop
//         if (counter === 0) {
//           matchingData.push(matchingEntry);
//           break;
//         }
//       }
//     }

//     Logger.log("Matching Data: " + JSON.stringify(matchingData));
//     return matchingData;
//   } catch (error) {
//     Logger.log("Error in searchPostAugust: " + error.toString());
//     return [];
//   }
// }

function searchPostAugust(userInput, sheet) {
  try {
    Logger.log("User Input: " + JSON.stringify(userInput));

    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

    Logger.log("Headers in August Sheet: " + JSON.stringify(headers));

    var matchingData = [];
    var data = sheet.getDataRange().getValues();

    // Column indexes in August sheet
    var thicknessColumnIndex = headers.indexOf("THK") + 1;
    var qualityColumnIndex = headers.indexOf("QUAL") + 1;
    var densityColumnIndex = headers.indexOf("DEN") + 1;
    var customerColumnIndex = headers.indexOf("CUST") + 1;
    var trackNoColumnIndex = headers.indexOf("TRACK NO.") + 1;

    // Counter for the number of sheets
    var counter = userInput["Number of Sheets"];

    // Iterate through the rows in the August sheet
    for (var i = 2; i < data.length; i++) {  // Start from row 3 (index 2)
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

        // Create an entry with only the necessary information
        var matchingEntry = {
          row: [userInput["Customer"], userInput["Quality"], userInput["Thickness"], userInput["Density"]],
          trackingNumber: trackingNumber
        };

        Logger.log("Collecting Row: " + JSON.stringify(matchingEntry));

        // Add the entry to matchingData
        matchingData.push(matchingEntry);

        // Decrease the counter
        counter--;

        // If the counter becomes zero, break out of the loop
        if (counter === 0) {
          break;
        }
      }
    }

    Logger.log("Matching Data: " + JSON.stringify(matchingData));
    return matchingData;
  } catch (error) {
    Logger.log("Error in searchPostAugust: " + error.toString());
    return [];
  }
}





// Modify writeToFormResponse to work with the "FormResponse" sheet
function writeToFormResponse(matchingData, sheet) {
  try {
    // Assuming headers are in the first row of the FormResponse sheet
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
  } catch (error) {
    Logger.log("Error in writeToFormResponse: " + error.toString());
  }
}
