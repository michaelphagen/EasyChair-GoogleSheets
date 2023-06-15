//Set up the menu items of the spreadsheet
// Select Spreadsheets: Choose the Spreadsheets to source data from
// Update sheets and overwrite: Update the sheets and overwrite the existing data
// Update sheets and do not overwrite: Update the sheets and do not overwrite the existing data
function onOpen() {
    // Create view page if it doesn't exist
    var page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("View");
    if (page == null) {
        page = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        page.setName("View");
        // Reorder the View page to be the first page
        page.activate();
        page.moveToBeginning();
    }
    // Get the spreadsheets from var page if they exist
    var page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("var");
    if (page != null) {
        // Column 2, row 1 is the id of the submission sheet
        var submissionSheetId = page.getRange(1, 2).getValue();
        // column 2, row 2 is the id of the Conference Data Export sheet
        var conferenceDataExportSheetId = page.getRange(2, 2).getValue();
    }
    else {
        // Create var page if it doesn't exist
        page = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        page.setName("var");
        page.getRange(1, 1).setValue("Submission Sheet");
        page.getRange(2, 1).setValue("Conference Data Export Sheet");
    }
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Update Sheets')
        .addItem('Select Spreadsheets', 'selectSheets')
        .addItem('Update sheets', 'updateSheetsOverwrite')
        //.addItem('Update sheets and do not overwrite', 'updateSheetsNoOverwrite')
        .addToUi();
}

// Path: Code.js
// Select the spreadsheets to source data from
function selectSheets() {
    // Get the spreadsheets from var page if they exist
    var page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("var");
    if (page != null) {
        // Column 2, row 1 is the id of the submission sheet
        var submissionSheetId = page.getRange(1, 2).getValue();
        // column 2, row 2 is the id of the Conference Data Export sheet
        var conferenceDataExportSheetId = page.getRange(2, 2).getValue();
    }
    // Create the UI for selecting the submissionSheetId
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
        'Enter the ID of the submissionSheet to source data from',
        'Enter the ids of the spreadsheet:',
        ui.ButtonSet.OK_CANCEL);
    // Process the user's response.
    if (response.getSelectedButton() == ui.Button.OK) {
        var id = response.getResponseText();
        // If the var page exists, update the id
        if (page != null) {
            page.getRange(1, 2).setValue(id);
        }
    }
// Create the UI for selecting the conferenceDataExportSheetId
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
        'Enter the ID of the conferenceDataExportSheet to source data from',
        'Enter the ids of the spreadsheet:',
        ui.ButtonSet.OK_CANCEL);
    // Process the user's response.
    if (response.getSelectedButton() == ui.Button.OK) {
        var id = response.getResponseText();
        // If the var page exists, update the id
        if (page != null) {
            page.getRange(2, 2).setValue(id);
        }
    }
}

// Path: Code.js
// Update the sheets and overwrite the existing data
// Imports the entirety of the submission sheet into the spreadsheet and adds the conference data export sheet, using the first column for ID
function updateSheetsOverwrite() {
    // Import the submission sheet into the current sheet
    importSubmissionSheet();
    // Import the conference data export sheet into the current sheet
    importConferenceDataExportSheet();
    buildViewPage();
}

function importSubmissionSheet(){
    //get submission sheet id
    var page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("var");
    if (page != null) {
        // Column 2, row 1 is the id of the submission sheet
        var submissionSheetId = page.getRange(1, 2).getValue();
    }
    // Get the submission sheet
    var submissionSheet = SpreadsheetApp.openById(submissionSheetId);
    // Get the submission sheet data
    var submissionSheetData = submissionSheet.getDataRange().getValues();
    // create sheet in current spreadsheet called submissions and import into it
    var submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("main");
    if (submissionsSheet == null) {
        submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        submissionsSheet.setName("main");
    }
    submissionsSheet.getRange(1, 1, submissionSheetData.length, submissionSheetData[0].length).setValues(submissionSheetData);
}

function importConferenceDataExportSheet(){
    // get conference data export sheet id
    var page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("var");
    if (page != null) {
        // Column 2, row 2 is the id of the Conference Data Export sheet
        var conferenceDataExportSheetId = page.getRange(2, 2).getValue();
    }
    // Get the conference data export sheet
    var conferenceDataExportSheet = SpreadsheetApp.openById(conferenceDataExportSheetId);
    // For each sheet in the conference dataspreadsheet, import it into the current spreadsheet
    var sheets = conferenceDataExportSheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
        var sheet = sheets[i];
        var sheetData = sheet.getDataRange().getValues();
        // Title is the name of the sheet with conferenceData at the beginning
        var title = "conferenceData" + sheet.getName();
        // Create the sheet if it doesn't exist
        var conferenceDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(title);
        if (conferenceDataSheet == null) {
            conferenceDataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
            conferenceDataSheet.setName(title);
        }
        // Import the data into the sheet
        conferenceDataSheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
    }
}

// Path: Code.js
// Get Column number from title
function getColumnByName(viewSheetData, title) {
    var titleColumn = 0;
    for (var i = 0; i < viewSheetData[0].length; i++) {
        if (viewSheetData[0][i] == title) {
            titleColumn = i;
            break;
        }
    }
    return titleColumn;
}

// Copy columns
function copyColumns(sourceSheet, destinationSheet, columnName) {
    // Find the column number of the column with the title "Title" in the view sheet
    var viewTitleColumn = getColumnByName(destinationSheet, columnName);
    var submissionsTitleColumn = getColumnByName(sourceSheet, columnName);
    // Copy the title column from the submissions sheet to the view data where the submission id matches
    for (var i = 1; i < sourceSheet.length; i++) {
        for (var j = 1; j < destinationSheet.length; j++) {
            if (destinationSheet[j][0] == sourceSheet[i][0]) {
                destinationSheet[j][viewTitleColumn] = sourceSheet[i][submissionsTitleColumn];
                break;
            }
        }
    }
    return destinationSheet;
}

// Path: Code.js
// Build view page
function buildViewPage() {
    // Get Current View page data as array
    var viewSheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("view").getDataRange().getValues();
    // Get the submission sheet data
    var submissionsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("main").getDataRange().getValues();
    // Get Export submissions data
    var conferenceDataExportData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("conferenceDataSubmissions").getDataRange().getValues();
    // Get the conferenceDataAuthors sheet
    var authorsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("conferenceDataAuthors").getDataRange().getValues();
    // Write column "#" of main to column "Submission ID" of view
    if (viewSheetData[0][0] != "#") {
        viewSheetData[0] = ["#", "Title", "Authors", "Emails","Related Tracks","Duration","A/V","Other equipment","Special Requests","abstract"];
    }
    // Copy the submission id column from the submissions sheet to the view data
    for (var i = 1; i < submissionsData.length; i++) {
        //if not in viewSheetData, add it
        var found = false;
        for (var j = 1; j < viewSheetData.length; j++) {
            if (viewSheetData[j][0] == submissionsData[i][0]) {
                found = true;
                break;
            }
        }
        if (!found) {
            viewSheetData.push([submissionsData[i][0]]);
        }
    }
    // Get viewSheetData[0] except for author, email, and abstract
    var submissionsDataLabels = viewSheetData[0].slice(0, 2).concat(viewSheetData[0].slice(4, 9));
    // For each label in submissionsDataLabels, copy the column from submissionsData to viewSheetData
    for (var i = 0; i < submissionsDataLabels.length; i++) {
        viewSheetData = copyColumns(submissionsData, viewSheetData, submissionsDataLabels[i]);
    }
    var viewAuthorsColumn = getColumnByName(viewSheetData, "Authors");
    var viewEmailsColumn = getColumnByName(viewSheetData, "Emails");
    for (var i = 1; i < viewSheetData.length; i++) {
        // Search the authors sheet for all authors with the same submission id
        var authors = "";
        for (var j = 1; j < authorsData.length; j++) {
            if (authorsData[j][0] == viewSheetData[i][0]) {
                // Combine author first and last name
                var author = authorsData[j][1] + " " + authorsData[j][2];
                if (authors == "") {
                    authors = author;
                }
                else {
                    authors = authors + "\n" + author;
                }

            }
        viewSheetData[i][viewAuthorsColumn] = authors;
        }
    }
    // Copy the emails column from the submissions sheet to the view data. Append if there is already an email
    for (var i = 1; i < viewSheetData.length; i++) {
        // Search the authors sheet for all authors with the same submission id
        var emails = "";
        for (var j = 1; j < authorsData.length; j++) {
            if (authorsData[j][0] == viewSheetData[i][0]) {
                if (emails == "") {
                    emails = authorsData[j][3];
                } else {
                    emails = emails + "\n" + authorsData[j][3];
                }
            }
        }
        viewSheetData[i][viewEmailsColumn] = emails;
    }
    viewSheetData = copyColumns(conferenceDataExportData, viewSheetData, "abstract");



    // write view data to view sheet
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("view").getRange(1, 1, viewSheetData.length, viewSheetData[0].length).setValues(viewSheetData);


}

