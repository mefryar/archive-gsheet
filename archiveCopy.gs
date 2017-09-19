// Program: Archive GSheet with Date Stamp
// Programmer: Michael Fryar
// Date: 19 September 2017
// Google Apps Script to copy Google Sheet to subfolder with date stamp in name

// Add Custom Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Archive')
      .addItem('Archive Copy with Date Stamp', 'archiveCopy')
      .addItem('Open Archive Folder', 'openArchive')
      .addToUi();
}

// Define function to copy sheet to subfolder with date stamp in name
// Building on https://gist.github.com/abhijeetchopra/99a11fb6016a70287112
function archiveCopy() {

  // Get timezone for datestamp
  var timeZone = Session.getScriptTimeZone();

  // Generate the datestamp and store in variable formattedDate as year-month-date
  var formattedDate = Utilities.formatDate(new Date(), timeZone , "yyyy-MM-dd");

  // Add datestamp stored in formattedDate to beginning of original file name
  var name = formattedDate + " your_file_name";

  // Get the destination folder by their ID
  var destination = DriveApp.getFolderById("XXXXXXXXXXXXXXXXXXXXXXXXXXXX");

  // Specify the Google Sheet (This ID is for "Training Team Portfolio Tracker")
  var file = DriveApp.getFileById("1m1_QRvh0ulhU5XWcQXjKx92HvR40CVcCynTWLQaD-Rg")

  // Archive copy of "file" with "name" at the "destination"
  file.makeCopy(name, destination);
}

// Define function to open archive folder in new tab
// Building on https://www.youtube.com/watch?v=2y7Y5hwmPc4
function openArchive() {

  var url = "https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";

  var userInterface = HtmlService.createHtmlOutput(html);  // Output is the HtmlOutput
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening Archive Folder');
}
