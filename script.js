/**
 *----------------------------------------------------------------------------------------------------------------------------------------------VELVET MARS v1.1.4------------------------------------------------------------------------------------------------------------------------------------------------------------------
 * Created: 7/9/2018
 * Authors: Gabriel Rosales & Darrell Cheney
 * Purpose: To solve the issue of how the teachers send requests to fix issues with their chromebooks
 */

// GLOBAL  VARIABLES


//Statuses within the status column 
  var COMPLETED = "Completed";
  var IN_PROGRESS = "In progress";
  var OUT_REPAIR = "Out for repair";

//Status to be written inside of the Email Sent column
  var STAT_1 = "COMPLETED";
  var STAT_2 = "IN PROGRESS";
  var STAT_3 = "OUT FOR REPAIR";

//Column numbers
  var ticketCol = 1;
  var emailCol = 2;
  var dateCol = 3;
  var autoEmailCol = 4;
  var campusCol = 5;
  var roomCol = 6;
  var chromeCol = 7;
  var issueCol = 8;
  var statusCol = 9;
  var escalatedCol = 12;

//Getting the variables needed for all functions
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 900; // Number of rows to process

  var dataRange = sheet.getRange(startRow, autoEmailCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from email_sent column
  var data = dataRange.getValues(); //getting values

  var emailRange = sheet.getRange(startRow, emailCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from email column
  var email = emailRange.getValues(); //getting values

  var status = sheet.getRange(startRow, statusCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from status column
  var data_status = status.getValues(); //getting values

  var chromebookNum = sheet.getRange(startRow, chromeCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from chromebook number column
  var chromeNum_data = chromebookNum.getValues(); //getting values

  var chromeIssue = sheet.getRange(startRow, issueCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from issue column
  var issue_status = chromeIssue.getValues(); //getting values

  var date = sheet.getRange(startRow, dateCol, sheet.getLastRow(), 1); //grabbing ranges of values to get from date column
  var date_data = date.getValues(); //getting values

  var ticketNum = sheet.getRange(startRow, ticketCol, sheet.getLastRow());
  var ticket_data = ticketNum.getValues();

  var roomNum = sheet.getRange(startRow, roomCol, sheet.getLastRow());
  var room_data = roomNum.getValues();

function focus() {
  // This function focuses the spreadsheet on tickets that are open rather than simply opening at the top and tells the technician how many open tickets they have
  var sheet = SpreadsheetApp.getActiveSheet();
  var statusRange = sheet.getRange(2, statusCol, 900, 1);
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


/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendemail() {
  
  //logic: if a field is populated and both Column C isn't populated, and Status is Completed, populate corresponding row in column C and send email.
  //stays in for loop untill there is data to be read
  for (var i = 0; email[i] != ""; ++i) 
  {
    
    //if data in array i doesn't have email_sent column and it's status is "completed", it sends a message
    if (data[i] != STAT_1 && data_status[i] == COMPLETED) {
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
          " \n Ticket Number: " + 
            ticket_data[i] +
        " \n Date Submitted: " +
        dateString +
        " \n \n Please do NOT reply to this email. If you need to contact your technician, email them at darrell.cheney@cvisd.org";
      //sets the email address, subject, and message from noreply and the reply to is set as the designated tech

      MailApp.sendEmail(emailAddress, subject, message, {
        replyTo: "darrell.cheney@cvisd.org",
        noReply: true
      });
      sheet.getRange(startRow + i, autoEmailCol).setValue(STAT_1);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
    
    // If email sent column doesn't contain IN PROGRESS and if the status in the status column is In progress, send email and write IN PROGRESS to Email sent column
    if (data[i] != STAT_2 && data_status[i] == IN_PROGRESS) {
      
      var emailAddress = email[i];
      var subject = "Chromebook Repair Update";
      var dateString = date_data[i].toString();

      dateString = dateString.substring(0, 15);
      
      var message =
          "This is an automated message from the Technology department. Your chromebook repair is in progress and we will keep you updated as often as possible. \n \n *** Ticket Details *** \n Chromebook Number: " +
        chromeNum_data[i] +
        "\n Issue: " +
        issue_status[i] +
           " \n Ticket Number: " + 
            ticket_data[i] +
        " \n Date Submitted: " +
        dateString +
        " \n \n Please do NOT reply to this email. If you need to contact your technician, email them at darrell.cheney@cvisd.org";
       MailApp.sendEmail(emailAddress, subject, message, {
        replyTo: "darrell.cheney@cvisd.org",
        noReply: true
      });
      sheet.getRange(startRow + i, autoEmailCol).setValue(STAT_2);
      
      SpreadsheetApp.flush();
      
    }
    
    // If email sent column doesn't contain IN PROGRESS and if the status in the status column is In progress, send email and write IN PROGRESS to Email sent column
    if (data[i] != STAT_3 && data_status[i] == OUT_REPAIR) {
      
      var emailAddress = email[i];
      var subject = "Chromebook Repair Update";
      var dateString = date_data[i].toString();

      dateString = dateString.substring(0, 15);
      
      var message =
          "This is an automated message from the Technology department. Your chromebook repair is out for repair and we will keep you updated as often as possible. \n \n *** Ticket Details *** \n Chromebook Number: " +
        chromeNum_data[i] +
        "\n Issue: " +
        issue_status[i] +
           " \n Ticket Number: " + 
            ticket_data[i] +
        " \n Date Submitted: " +
        dateString +
        " \n \n Please do NOT reply to this email. If you need to contact your technician, email them at darrell.cheney@cvisd.org";
       MailApp.sendEmail(emailAddress, subject, message, {
        replyTo: "darrell.cheney@cvisd.org",
        noReply: true
      });
      sheet.getRange(startRow + i, autoEmailCol).setValue(STAT_3);
      
      SpreadsheetApp.flush();
    }
  }
}

var NEW_ISSUE = "New Issue";

function defaultValue() {
  
  var startRow = 2;
  var first = "";
  
  var roomNum = sheet.getRange(2, roomCol, sheet.getLastRow());
  var room_data = roomNum.getValues();

  //get data from campus column
  var dataRange = sheet.getRange(2, campusCol, sheet.getLastRow());
  var data = dataRange.getValues();

  //get data from status column
  var statusRange = sheet.getRange(2, statusCol, sheet.getLastRow());
  var status = statusRange.getValues();

  //the for loop is initiated until there is something other than '' in the column. If there is '' then it sets the value as New Issue
  for (var i = 0; data[i] != ""; ++i) {
    if (status[i] == "") {
      sheet.getRange(i + startRow, statusCol).setValue(NEW_ISSUE);
      var ticketNum = i+startRow;
      if (ticketNum<10) {
        first = "00"; 
      } else if (ticketNum>9 && ticketNum<100) {
        first = "0"; 
      } else { 
        first = "";
      }
      sheet.getRange(i + startRow, ticketCol ).setValue("C"+ room_data[i] + "-" + first + ticketNum);
      
    }
  }

  var preCart = "Cart ";
  var CHS_Sheet = SpreadsheetApp.openById("1BZJRcB115H5ie-6kKZyXQ48uGUpuQ-yk3yw1W3rNoYU");
  var serial_Data = sheet.getRange(2,10,sheet.getLastRow()).getValues();
  var cLength = chromeNum_data.length;
  cLength -=1;
  
  for (var c=0;c<cLength;++c) {
    if (serial_Data[c] == '') {
    var fullID = preCart + chromeNum_data[c].toString();
    fullID = fullID.substring(0,7);
    var cbookNumber = chromeNum_data[c].toString();
    cbookNumber = cbookNumber.substring(3,5);
    cbookNumber = cbookNumber.valueOf();
    ++cbookNumber;
    var CHS_Tab = CHS_Sheet.getSheetByName(fullID);
    if (CHS_Tab) {
      var CHS_Range = CHS_Tab.getRange(cbookNumber,3);
      var serial = CHS_Range.getValue();
      var serial_Range = sheet.getRange(c+2,10);
      serial_Range.setValue(serial);
    }
    else {
      var serial_Range = sheet.getRange(c+2,10);
      serial_Range.setValue("Cart not found");
    }
    } else {}
    
  }

}

function onColorChange() {
  var ss = SpreadsheetApp.getActiveSheet();
  var startRow = 1; //have to start with one otherwise writing to the cell's don't work as expected
  
  var range = ss.getRange(startRow, escalatedCol, 999); //getting range
  var bgColors = range.getBackgrounds(); //getting cell colors

  var escalated_data = range.getValues(); //getting values
  var escalated = "Escalated";  
  
  for (var i in bgColors) { 
      // if color isn't white and if the value isn't set to escalated, set to escalate and send email
      if (bgColors[i] != "#ffffff" && escalated_data[i] != escalated) {
      
        var emailAddress = email[i]; // First column

      var dateString = date_data[i].toString();

      dateString = dateString.substring(0, 15);
          var subject = "Chromebook Repair Needs Escalation";
          var message =
          "There is an issue that needs to be addressed.  \n \n *** Ticket Details *** \n Chromebook Number: " +
        chromeNum_data[i] +
        "\n Issue: " +
        issue_status[i] +
           " \n Ticket Number: " + 
            ticket_data[i] +
        " \n Date Submitted: " +
        dateString +
        " \n \n Please do NOT reply to this email. If you need to contact your technician, email them at gabriel.rosales@cvisd.org";
       MailApp.sendEmail("gabriel.rosales@cvisd.org", subject, message, {
        noReply: true
      });

        ss.getRange(++i, escalatedCol).setValue(escalated);
      }
   
    // if color is white and if the value is set to escalated, it will delete escalated value and set it to ""
    if(bgColors[i] == "#ffffff" && escalated_data[i] == escalated) {
      ss.getRange(++i, escalatedCol).setValue('');
    
     }
   }
}
