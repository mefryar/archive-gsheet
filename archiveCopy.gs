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

  // Replace "spreadsheetId" with the ID of the Google Sheet you wish to copy
  var file = DriveApp.getFileById("spreadsheetId")

  // Replace "folderId" with the ID of the folder where you want the copy saved
  var destination = DriveApp.getFolderById("folderId");

  // Get timezone for datestamp
  var timeZone = Session.getScriptTimeZone();

  // Generate datestamp and store in variable formattedDate as year-month-date
  var formattedDate = Utilities.formatDate(new Date(), timeZone , "yyyy-MM-dd");

  // Replace "file_name" with the name you want to give the copy
  var name = formattedDate + " file_name";

  // Archive copy of "file" with "name" at the "destination"
  file.makeCopy(name, destination);
}

// Define function to open archive folder in new tab
// Building on https://www.youtube.com/watch?v=2y7Y5hwmPc4
function openArchive() {

  // Replace "folderId" with the ID of the folder where you want copies saved
  var url = "https://drive.google.com/drive/folders/folderId"

  // HTML to open folder url in new tab and then close dialogue window in sheet
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";

  // Push HTML into user interface
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening Archive Folder');
}
