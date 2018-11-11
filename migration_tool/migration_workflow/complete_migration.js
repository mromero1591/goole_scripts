/*Name: Mark Romero
* Date: 2/23/2016
* Updated: 12/19/2016
* Purpose: This script file contains the code to send a completed migration email.
*==============================================================================*/

/*Purpose: This function sends a completed message email and then calls the update
* Quailty check function.
* Parameters: None
* Returns : None
*==============================================================================*/
function migrationCompletedEmail() {
  //create the workbook class
  var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  var workbookData = getWorkbookData(workbook); 
  
  //Create a variable that will hold the cell spot where the message "validation email sent by" will be place.
  var COMPLETED_MESSAGE_SPOT = "C3";

  //grab the value where the completed message is in order to check if this migration has already been completed.
  var completedMessage =  workbookData.postMigrationSheet.getRange(COMPLETED_MESSAGE_SPOT).getValue();
  
  //If the completed message is present then the site has already been completed, infrom the user with a message.
  if( completedMessage == "Validation email sent by") {
    var alreadySentMessage = "This site has already been Validated";
    SpreadsheetApp.getUi().alert(alreadySentMessage);
    return;
  }
  
  var coreConsultantEmail = getEmailAddress(workbookData.consultantName);

  var secondaryConsultantEmail = getEmailAddress(workbookData.secondaryConsultantName);
  var accountingConsultantEmail = getAccountingConsultantEmail(workbookData.accountConsultantName);

  var secondaryAccountingConsultantEmail = getAccountingConsultantEmail(workbookData.secondaryAccountingConsultantName);
  var accountingQcEmail = getAccountingConsultantEmail(workbookData.accountingConsultantQcName);
  
  var ccEmail = "emails here";
  
  if(workbook.accountConsultantName != "NA" || workbook.accountConsultantName != "")
  {
      ccEmail = ccEmail + "," + accountingConsultantEmail;
  }
  
  if(workbook.secondaryConsultantName != "NA" || workbook.secondaryConsultantName != "")
  {
      ccEmail = ccEmail + "," + secondaryConsultantEmail;
  }
  
  if(workbook.secondaryAccountingConsultantName != "NA" || workbook.secondaryAccountingConsultantName != "")
  {
      ccEmail = ccEmail + "," + secondaryAccountingConsultantEmail;
  }
  
  if(workbook.accountingConsultantQcName != "NA" || workbook.accountingConsultantQcName != "")
  {
      ccEmail = ccEmail + "," + accountingQcEmail;
  }

  //create the subject line for the email using company name and site name.
  var subject = "Entrata Migration Completed for: " + workbookData.companyName + " / " + workbookData.siteName;

  //create the email message using both the reports link and url link.
  var emailBody = "This site has been migrated. If there are any areas for you to review, they will be listed below. "
                   + "Otherwise, the Post Migration checklist link is listed below." + "<br />"
                   + "<br />" + "<b>Reports Link:</b> " + workbookData.reportsLink + "<br />"
                   + "<br />" + "<b>Checklist Link:<b /> "+ workbookLink + "<br />" 
                   + "<br />" + "Thank you, " + "<br />"
                   + "<br />";

  //do not change this, any customization can be done with the variables above.
  //set the entrata email of user as the main email
  var entrataEmail = Session.getActiveUser().getEmail();
  //get the signature of the user.
  var emailSig = getEmailSig(entrataEmail);
  
  //add the signature to the body of the email.
  var emailMessage = emailBody + emailSig;

  //If the user has an entrata email then the email will be sent using the entrata email.
  if(entrataEmail != null) 
  {
        GmailApp.sendEmail(coreConsultantEmail, subject, "",{cc:ccEmail,htmlBody:emailMessage,from:entrataEmail});
  }
  //If the user does not have an entrata email then send with the users propertysolutions email.
  else 
  {
        //do not change this, any customization can be done throw the variables.
        MailApp.sendEmail({to: coreConsultantEmail,
                           cc: ccEmail,
                           subject: subject,
                           htmlBody: emailMessage});
  }

  //Once email has been sent then add the completed migration message, including the migration agent who completed the validation to the post migration checklist.
  var completedMessage = "Validation email sent by";
  var compltedMessageCell = workbookData.postMigrationSheet.getRange(COMPLETED_MESSAGE_SPOT).setValue(completedMessage);
  var migrationAgent = getCompletedConsultant(entrataEmail);
  var completedMigrationAgentCell = workbookData.postMigrationSheet.getRange(3, 4).setValue(migrationAgent);
  
  //update Quailty Check In.
  updateQuailtyCheck(workbookData);
  //Track the queries.
  updateTrackerSheet(workbookData.siteName, workbookData.units, workbookData.queryDocUrl, workbookData.companyName);     
} //end of migrationCompletedEmail function.


