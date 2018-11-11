/*Name: Mark Romero
* Date: 1/22/2016
* Updated: 12/6/2017
* Purpose: This script file contains the code to manage and send Pre Migration Issues
* During check in
*==============================================================================*/

/*Purpose: Opens the side bar, by referencing the dataCorrection.html file
* Parameters: None
* Returns : None
*==============================================================================*/
function dataCorrectionSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataCorrection')
      .setTitle('Pre-Check Issue')
      .setWidth(500);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

/*Purpose: populates the data correction document.
* Parameters: issueType, The Type of issue
*             Description, Description of the issue
*             emergency, If the issue is an emergency
*             otherText, If issue type is other then text will be issue Type
* Returns : None
*==============================================================================*/
function requestDataCorrection(issueType, description, emergency, otherText) {
  //create the workbook class
  var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  var workbookData = getWorkbookData(workbook);  
  
  //get the name and email of those who reported the issue
  var reportedByEmail = Session.getActiveUser().getEmail();
  var reportedByName = convertEmailToName(reportedByEmail);
  
  //Data Fix Sheet.
  var dataFixSheetUrl = "https://docs.google.com/spreadsheets/d/1s_QotFOrH0dWVjfMgy4XOXpbb28kigFRelVmhpANE9I/edit#gid=0";
  var dataFixsheetName = "Data Correction Request";

  //Get the column location of the data fix columns
  var dataFixsheet = SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName);
  var reportedOnColumn = getColumnByName(dataFixsheet, "Reported On");
  var reportedByColumn = getColumnByName(dataFixsheet, "Reported By");
  var coreConsultantColumn = getColumnByName(dataFixsheet, "Core Consultant");
  var acctConsultantColumn = getColumnByName(dataFixsheet, "ACCT Consultant");
  var companyColumn = getColumnByName(dataFixsheet, "Company");
  var siteColumn = getColumnByName(dataFixsheet, "Site");
  var fieldColumn = getColumnByName(dataFixsheet, "Field");
  var descriptionColumn = getColumnByName(dataFixsheet, "Description");
  var requiredColumn = getColumnByName(dataFixsheet, "Required for release to BO");
  var workColumn = getColumnByName(dataFixsheet, "Workbook");
  
  //get last row with data in the issues doc
  var dataFixLastRow = SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getLastRow();
  
  /*****Set the data onto the quality work book ************************************************
  ******DO NOT CHANGE these, any changes made can be changed by changing the vaiable ***********/
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, reportedOnColumn).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, reportedByColumn).setValue(reportedByName);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, coreConsultantColumn).setValue(workbookData.consultantName);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, acctConsultantColumn).setValue(workbookData.accountConsultantName);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, companyColumn).setValue(workbookData.companyName);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, siteColumn).setValue(workbookData.siteName);
  if(issueType == "Other") {
    SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, fieldColumn).setValue(otherText);
  } else {
    SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, fieldColumn).setValue(issueType);
  }
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, descriptionColumn).setValue(description);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, requiredColumn).setValue(emergency);
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, workColumn).setValue(workbookLink);
  
  sendDataCorrectionEmail(reportedByEmail, workbookData, issueType, description, emergency, workbookLink, otherText);
}

function sendDataCorrectionEmail(reportedByEmail, workbookData, issueType, description, emergency, workbookLink, otherText) {
  var issue = "";
  if(issueType == "Other") {
    issue = otherText
  } else {
    issue = issueType;
  }
  
  //Build the email subject, body and ccemail
  var emailSubject = workbookData.companyName + " - " + workbookData.siteName + " Pre-Check Issue";
  

  var emailBody = "While Checking in this site I have come accross the following: "
                   + "<br>" + "<br>"
                   + "<strong>Company:</strong> " + workbookData.companyName
                   + "<br>"
                   + "<strong>Site:</strong> " + workbookData.siteName
                   + "<br>"
                   + "<strong>Issue:</strong> " + issue
                   + "<br>"
                   + "<strong>Description:</strong> " + description
                   + "<br>"
                   + "<strong>Workbook:</strong> " + workbookLink
                   + "<br>"
                   + "<p>Needs to be completed before we can start the migration:" + emergency + "</p>"

  var consultantEmail = convertNameToEmail(workbookData.consultantName);
  var acctEmail = convertNameToEmail(workbookData.accountConsultantName);
  var ccEmail = acctEmail;
  
  //send the email
  if(reportedByEmail != null) {
     GmailApp.sendEmail(consultantEmail, emailSubject, "",{cc: ccEmail, htmlBody:emailBody,from:reportedByEmail});
  }
  else {
     //do not change this, any customization can be done throw the variables.
     MailApp.sendEmail({to: consultantEmail,
                        cc: ccEmail,
                        subject: emailSubject,
                        htmlBody: emailBody});
  }
}
