/*Name: Mark Romero
* Date: 12/04/2017
* Updated: 
* Purpose: This script file contains The code that will allow the user to send
* The migration workbook to the migration consultant team. It also creates a trigger
* that will check to see if the migration workbook has been sent to company.
*==============================================================================*/


/*Purpose: Allows user to send the migration workbook to the migration consultants
* Parameters: None
* Returns : None
*==============================================================================*/
function beginMigration() {
  //create the workbook
  var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl(); //https://docs.google.com/spreadsheets/d/14XUxqtJF2TXE7dnmWxlRpD8XJroqn7kSOArSs1kJSMA/edit#gid=1770421006
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  var workbookData = getWorkbookData(workbook);
  
  //get the migration workbook name.
  var migrationWorkbookName = workbook.getName();
  
  //set the emails for validation and migration consultants
  var coreValidationsEmail = "email here";
  var migrationConsultantEmail = "email here";
  
  //get the entrata email address of the person using the tool
  var consultantEntaraEmail = Session.getActiveUser().getEmail();
  
  //get the signature of the person using the tool
  var consultantSig = getEmailSig(consultantEntaraEmail);
  
  //share workbook with coreValidation.
  workbook.addEditor(coreValidationsEmail);
  
  //set sharing to everyone with entrata can edit.
  DriveApp.getFileById(workbook.getId()).setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
 
  //Build the email subject, body and ccemail
  var emailSubject = workbookData.companyName + " - " + workbookData.siteName + " is ready for migration";
  

  var emailBody = "<b>Company Name:</b> " + workbookData.companyName + " - " + workbookData.companyId + "<br />"
                        + "<b>Site Name:</b> " + workbookData.siteName + " - " + workbookData.siteId + "<br />"
                        + "<b>Checklist Link:</b> " + workbookLink + "<br />"
                        + "<b>Special Instructions:</b>" + "<br />"
                        + workbookData.specialInstructions + "<br />" + "<br />";

  var ccEmail = coreValidationsEmail;
  var finalEmail = emailBody + consultantSig;
  
  //send the email
  if(consultantEntaraEmail != null) {
     GmailApp.sendEmail(migrationConsultantEmail, emailSubject, "",{cc: ccEmail, htmlBody:finalEmail,from:consultantEntaraEmail});
  }
  else {
     //do not change this, any customization can be done with the variables.
     MailApp.sendEmail({to: migrationConsultantEmail,
                        cc: ccEmail,
                        subject: emailSubject,
                        htmlBody: finalEmail});
  }
  
  //create clock trigger that will notify migrations if the migration workbook has not been shared with company.
  var now = Utilities.formatDate(new Date(), "GMT-7", "MM/dd/YYYY'.'HH:mm:ss");
  var todaysDate = new Date();
  var today = todaysDate.getDate();
  
  workbookData.preStartSheet.getRange(4, 1).setValue(now).setFontColor("#ffffff");
  workbookData.preStartSheet.setTabColor("#ffff00");

  ScriptApp.newTrigger('checkIfSent').timeBased().atHour(17).onMonthDay(today).create();
  
}

/*Purpose: Check to see if the migration work book has been checkin and sent to company
* Parameters: None
* Returns : None
*==============================================================================*/
function checkIfSent() {
  if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre Start Checklist").getRange(3, 1).getValue() == "") {
    //create the workbook
    var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var workbook = SpreadsheetApp.openByUrl(workbookLink) // getActiveSpreadsheet();
    var workbookData = getWorkbookData(workbook);
  
    //Set needed emails
    var migrationConsultantEmail = "email here";
    var consultantEntaraEmail = Session.getActiveUser().getEmail();
    var consultantSig = getEmailSig(consultantEntaraEmail);
  
    //Build the email
    var emailSubject = "URGENT - Site Not Sent to company name: " + workbookData.companyName + " - " + workbookData.siteName;
    
    var emailBody = "The below site has been submitted to the Migration Consultants but has not been sent to company name" + "<br />" 
                      + "<b>Company Name:</b> " + workbookData.companyName + " - " + workbookData.companyId + "<br />"
                      + "<b>Site Name:</b> " + workbookData.siteName + " - " + workbookData.siteId + "<br />"
                      + "<b>Checklist Link:</b> " + workbookLink + "<br />" 
                      + workbookData.specialInstructions + "<br />" + "<br />";
  
    var finalEmail = emailBody + consultantSig;
  
    //send the email
    if(consultantEntaraEmail != null) {
      GmailApp.sendEmail(migrationConsultantEmail, emailSubject, "",{htmlBody:finalEmail,from:consultantEntaraEmail});
    }
    else {
       //do not change this, any customization can be done throw the variables.
       MailApp.sendEmail({to: migrationConsultantEmail,
                          subject: emailSubject,
                          htmlBody: finalEmail});
    }
  }
}