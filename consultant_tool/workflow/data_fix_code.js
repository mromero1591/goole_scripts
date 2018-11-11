/*Name: Mark Romero
* Date: 1/22/2016
* Updated:
* Purposuse: This file contains functions that are used to request data fixes.
*==============================================================================*/

/*Purpose: Opens the data fix input form as a side bar.
* Parameters: None
* Returns : None
*==============================================================================*/
function dataFixSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataFix')
      .setTitle('Validation issue')
      .setWidth(500);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}


/*Purpose: Sends an email and fills the Data tracker doc when a request is made.
* Parameters: issueType, Text from dropdown that indicates the type of fix
*             description, Text that describes the fix
              emergency, Yes/No stating if the fix is an emergency
              otherText, If type is other then type will be converted to other text value
* Returns : None
*==============================================================================*/
function requestDataFix(issueType, fixType, numOfFixes, description, examples, emergency, otherText) {
  //Sheets
  var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  var workbookData = getWorkbookData(workbook);
  var reportedByEmail = Session.getActiveUser().getEmail();
  var reportedByName = convertEmailToName(reportedByEmail);
  
  if(workbookData.validationMigrationConsultantName == undefined) {
    return "Fail, Not Validated";
  }

  var dataFix = {
    reportedByEmail: reportedByEmail,
    reportedByName: reportedByName,
    issueType: issueType,
    fixType: fixType,
    numOfFixes: numOfFixes,
    description: description,
    examples: examples,
    emergency: emergency,
    otherText: otherText
  };
  
  updateDataFixSheet(workbookData, dataFix,workbookLink);
  updateQueryDoc(workbookData, dataFix,workbookLink);
  
  sendDataFixEmail(workbookData, reportedByEmail, issueType, fixType, description, examples, emergency, workbookLink, otherText);
}
  
function sendDataFixEmail(workbookData, reportedByEmail, issueType, fixType, description, examples, emergency, workbookLink, otherText) {
  
  var issue = "";
  if(issueType == "Other") {
    issue = otherText
  } else {
    issue = issueType;
  }
  
  //Build the email subject, body and ccemail
  var emailSubject = workbookData.companyName + " - " + workbookData.siteName + " Validation Issue";
  

  var emailBody = "While Validating this site I have come accross the following: "
                   + "<br>"
                   + "<strong>Company:</strong> " + workbookData.companyName
                   + "<br>"
                   + "<strong>Site:</strong> " + workbookData.siteName
                   + "<br>"
                   + "<strong>Issue:</strong> " + issue
                   + "<br>"
                   + "<strong>Fix Type:</strong>" + fixType
                   + "<br>"
                   + "<strong>Description:</strong> " + description
                   + "<br>"
                   + "<strong>Examples:</strong>"
                   + "<br>"
                   + examples
                   + "<br>"
                   + "<strong>Workbook:</strong> " + workbookLink
                   + "<br>"
                   + "Needs to be completed before we can send to client: " + emergency + "<br>"
                   + "<p style='color:white'>ValidationIssue</p>"


  var consultantEmail = convertNameToEmail(workbookData.consultantName);
  var acctEmail = convertNameToEmail(workbookData.accountConsultantName);
  var migrationConsultantEmail = convertNameToEmail(workbookData.validationMigrationConsultantName);
  var ccEmail = "";
  if(fixType == "Post-Go-Live issue/change") {
    ccEmail = consultantEmail + "," + acctEmail;
     //send the email
    if(reportedByEmail != null) {
       GmailApp.sendEmail("email here", emailSubject, "",{cc: ccEmail, htmlBody:emailBody,from:reportedByEmail});
    }
    else {
       //do not change this, any customization can be done throw the variables.
       MailApp.sendEmail({to: "email here",
                          cc: ccEmail,
                          subject: emailSubject,
                          htmlBody: emailBody});
    }
    
  }
  else {
    ccEmail = consultantEmail + "," + acctEmail;
      //send the email
    if(reportedByEmail != null) {
  
       GmailApp.sendEmail(migrationConsultantEmail, emailSubject, "",{cc: ccEmail, htmlBody:emailBody,from:reportedByEmail});
    }
    else {
       //do not change this, any customization can be done throw the variables.
       MailApp.sendEmail({to: migrationConsultantEmail,
                          cc: ccEmail,
                          subject: emailSubject,
                          htmlBody: emailBody});
    }
  }
}


