//**
//*----------------------------------------------------------------------------------------------------------------------------------------------VELVET MARS v1.0.4------------------------------------------------------------------------------------------------------------------------------------------------------------------
//* Created: 7/9/2018
//* Authors: Gabriel Rosales & Darrell Cheney
//* Purpose: To solve the issue of how the teachers send requests to fix issues with their chromebooks
//**

function focus() {
  // This function focuses the spreadsheet on tickets that are open rather than simply opening at the top and tells the technician how many open tickets they have
  var sheet = SpreadsheetApp.getActiveSheet();
  var statusRange = sheet.getRange(2, 8, 900, 1);
  var statusValues = statusRange.getValues();
  var openTickets = 0;
  var focusRow = 0;

  for (var i = 0; statusValues[i] != ""; ++i) {
    if (statusValues[i] == "Completed") {
    } else if (openTickets == 0) {
      focusRow = i + 28;
      openTickets += 1;
    } else {
      openTickets += 1;
    }
  }

  var range = sheet.getRange(focusRow, 1);
  range.activate();
  Browser.msgBox("You have " + openTickets + " open tickets");
}

function checkEmail() {
  //on edit checks email function
  sendemail();
}

function checkStatus() {
  //on open runs default value function and runs it every minute
  defaultValue();
}

// This constant is written in column C for rows for which an email has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";
var COMPLETED = "Completed";

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendemail() {
  //gets active sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 900; // Number of rows to process

  var dataRange = sheet.getRange(startRow, 3, numRows, 1); //grabbing ranges of values to get from email_sent column
  var data = dataRange.getValues(); //getting values

  var emailRange = sheet.getRange(startRow, 1, numRows, 1); //grabbing ranges of values to get from email column
  var email = emailRange.getValues(); //getting values

  var status = sheet.getRange(startRow, 8, numRows, 1); //grabbing ranges of values to get from status column
  var data_status = status.getValues(); //getting values

  var chromebookNum = sheet.getRange(startRow, 6, numRows, 1); //grabbing ranges of values to get from chromebook number column
  var chromeNum_data = chromebookNum.getValues(); //getting values

  var chromeIssue = sheet.getRange(startRow, 7, numRows, 1); //grabbing ranges of values to get from issue column
  var issue_status = chromeIssue.getValues(); //getting values

  var date = sheet.getRange(startRow, 2, numRows, 1); //grabbing ranges of values to get from date column
  var date_data = date.getValues(); //getting values

  //logic: if a field is populated and both Column C isn't populated, and Status is Completed, populate corresponding row in column C and send email.
  //stays in for loop untill there is data to be read
  for (var i = 0; email[i] != ""; ++i) {
    //if data in array i doesn't have email_sent column and it's status is "completed", it sends a message
    if (data[i] != EMAIL_SENT && data_status[i] == COMPLETED) {
      // Prevents sending duplicates

      var emailAddress = email[i]; // First column

      var dateString = date_data[i].toString();

      dateString = dateString.substring(0, 15);

      var subject = "Chromebook Repair Update";
      //Browser.msgBox(dateString);

      var message =
        "This is an automated message from the Technology department. Your chromebook repair has been completed and should be returned to you within 1 work day. \n \n *** Ticket Details *** \n Chromebook Number: " +
        chromeNum_data[i] +
        "\n Issue: " +
        issue_status[i] +
        " \n Date Submitted: " +
        dateString +
        " \n \n Please do NOT reply to this email. If you need to contact your technician, email them at darrell.cheney@cvisd.org";
      //sets the email address, subject, and message from noreply and the reply to is set as the designated tech

      MailApp.sendEmail(emailAddress, subject, message, {
        replyTo: "darrell.cheney@cvisd.org",
        noReply: true
      });
      sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}

var NEW_ISSUE = "New Issue";

function defaultValue() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;
  var numRows = 900;

  //get data from campus column
  var dataRange = sheet.getRange(startRow, 4, numRows);
  var data = dataRange.getValues();

  //get data from status column
  var statusRange = sheet.getRange(startRow, 8, numRows);
  var status = statusRange.getValues();

  //the for loop is initiated until there is something other than '' in the column. If there is '' then it sets the value as New Issue
  for (var i = 0; data[i] != ""; ++i) {
    if (status[i] == "") {
      sheet.getRange(i + startRow, 8).setValue(NEW_ISSUE);
    }
  }
}

function onColorChange() {
  var ss = SpreadsheetApp.getActiveSheet();
  var startRow = 1; //have to start with one otherwise writing to the cell's don't work as expected
  var range = ss.getRange(startRow, 10, 900); //getting range
  var bgColors = range.getBackgrounds(); //getting cell colors
  var escalated_data = range.getValues(); //getting values
  var escalated = "Escalated";

  for (var i in bgColors) {
    //array 1
    for (var j in bgColors[i]) {
      //array 2 in array 1
      // if color isn't white and if the value isn't set to escalated, set to escalate and send email
      if (bgColors[i][j] != "#ffffff" && escalated_data[i][j] != escalated) {
        MailApp.sendEmail(
          "gabriel.rosales@cvisd.org",
          "colors",
          "something isn't white"
        );

        Browser.msgBox("meets conditions");
        ss.getRange(++i, 10).setValue(escalated);
      }
    }
  }
}
