/*Name: Mark Romero
* Date: 1/22/2016
* Updated: 12/19/2016
* Purpose: This script file contains the code That will send the please Migrate
* Email.
*==============================================================================*/

/*Purpose: Send an email to Entrata Migrations requesting a migration, while also calling the update check in and Accounting docs.
* Parameters: None
* Returns : None
*==============================================================================*/
function pleaseMigrate() {
  //create the workbook class
  var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  var workbookData = getWorkbookData(workbook);   

  /*First Check to make sure the migration email has not already been sent, by:
   *1.getting the value in the date sent to company cell.
   *2.If value in that cell is empty then the email has not been sent to company yet, and you can continue with the request.
  */
  var DATE_SENT_TO_company_CELL = "A2";
  
  if(workbookData.preStartSheet.getRange(DATE_SENT_TO_company_CELL).getValue() != "") {
    SpreadsheetApp.getUi().alert("Site has already been Checked In");
    return;
  }

  if(workbookData.migrationType == "Hybrid" || workbookData.migrationType == "Utility" || workbookData.migrationType == "UDS") {
    handleUtilityMigration(workbookData, workbookLink);
  }   
  else if (workbookData.migrationType == "E2E") {
    handleE2EMigration(workbookData, workbookLink);
  }
  else if (workbookData.migrationType == "In-House"){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("You have indicated that this is a In House Migration. Are you sure you want to continue?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) {
      SpreadsheetApp.getUi().alert("Check in has been canceled");
      return;
    }
  }
  else {
    handleAllMigrations(workbookData, workbookLink)
  }

  //Set the check in date
  workbookData.preStartSheet.getRange(DATE_SENT_TO_company_CELL).setFontColor('white').setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
  //change the pre start checklist tab to green, indicating that the email has been sent.
  //workbookLink.preStartSheet.setTabColor("green");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre Start Checklist").setTabColor("green");
  //call the update check in function to update the Data Check In sheet, in the quailty sheet.
  updateCheckIn(workbookData);
  
  //Call the update accounting qc function.
  updateAccountingQC(workbookData, workbookLink);
}

function handleUtilityMigration(workbookData, workbookLink) {
  //create the email to send
  var emailAddress =  "email here"; 
  var entrataEmail = Session.getActiveUser().getEmail();
  var ccEmail = "email here";
  var subject = "(Utility Migration) Please Migrate the following site: " + workbookData.companyName + " - " + workbookData.companyId + " / " + workbookData.siteName + " - " + workbookData.siteId;
  var emailSig = getEmailSig(entrataEmail);
  
  //Create the email for the E2E migration.
  var emailBody = "<b>Company Name: </b> " + workbookData.companyName + " - " + workbookData.companyId + "<br />"
                  + "<b>Site Name: </b> " + workbookData.siteName + " - " + workbookData.siteId + "<br />"
                  + "<b>As of Date: </b> " + workbookData.asOfDate + "<br />"
                  + "<b>Property Type: </b> " + workbookData.propertyType + "<br />"
                  + "<b>Migration Type: </b>" + workbookData.migrationType + "<br />"
                  + "<b>Unit: </b> " + workbookData.units + "<br />"
                  + "<b>Space: </b> " + workbookData.spaces + "<br />"
                  + "<br />" + "<b>Former Property Management Software: </b> " + workbookData.formerPropertySoftware + "<br />"
                  + "<br />" + "<b>Expected Return Date: </b> " + workbookData.expectedReturnDate + "<br />"
                  + "<br />" + "<b>Checklist Link: </b> " + workbookLink + "<br />"
                  + "<br />" + "Please find the reports in the FTP Site" + "<br />" + "<br />";
                  
  //Add the email signature to the email message.
  var emailMessage = emailBody + emailSig;
  
  var email = {
    emailAddress: emailAddress,
    entrataEmail: entrataEmail,
    ccEmail: ccEmail,
    subject: subject,
    emailSig: emailSig,
    emailMessage: emailMessage
  }
  
  sendEmail(email);
  populateUtilityDocument(workbookData, workbookLink);
//  if(workbookData.migrationType == "UDS") {
//    sendUDSEmail(workbookData, workbookLink);
//  }
}

function populateUtilityDocument(workbookData, workbookLink) {
  //Search constants
  var DATE_COL_SEARCH = "Date";
  var COMPANY_NAME_COL_SEARCH = "Company Name";
  var COMPANY_ID_COL_SEARCH = "Company Id";
  var SITE_NAME_COL_SEARCH = "Site Name";
  var SITE_ID_COL_SEARCH = "Site Id";
  var MIGRATION_TYPE_COL_SEARCH = "Migration Type";
  var LOOKUP_CODE_COL_SEARCH = "Look Up Code";
  var WORKBOOK_LINK_COL_SEARCH = "Workbook Link";
  
  //Utility documents values
  var UTILITY_MIGRATION_DOC_LINK = "link here";
  var utilityWorkbook = SpreadsheetApp.openByUrl(UTILITY_MIGRATION_DOC_LINK);
  var utilitySheetName = "Utility Migrations";
  var utilitySheet= utilityWorkbook.getSheetByName(utilitySheetName);

  //get the column locations
  var dateCol = getColumnByName(utilitySheet, DATE_COL_SEARCH);
  var companyNameCol = getColumnByName(utilitySheet, COMPANY_NAME_COL_SEARCH);
  var companyIdCol = getColumnByName(utilitySheet, COMPANY_ID_COL_SEARCH);
  var siteNameCol = getColumnByName(utilitySheet, SITE_NAME_COL_SEARCH);
  var siteIdCol = getColumnByName(utilitySheet, SITE_ID_COL_SEARCH);
  var migrationTypeCol = getColumnByName(utilitySheet, MIGRATION_TYPE_COL_SEARCH);
  var lookupCodeCol = getColumnByName(utilitySheet, LOOKUP_CODE_COL_SEARCH);
  var workbookLinkCol = getColumnByName(utilitySheet, WORKBOOK_LINK_COL_SEARCH);

  //Get the last row in the check in sheet with data.
  var sheetLastRow = utilitySheet.getLastRow();
  
  //set data at then of the migration sheet
  SpreadsheetApp.openByUrl(UTILITY_MIGRATION_DOC_LINK).getSheetByName(utilitySheetName).getRange(sheetLastRow + 1, dateCol).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, companyNameCol, workbookData.companyName);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, companyIdCol, workbookData.companyId);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, siteNameCol, workbookData.siteName);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, siteIdCol, workbookData.siteId);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, migrationTypeCol, workbookData.migrationType);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, lookupCodeCol, workbookData.propertyLookUpCode);
  setData(UTILITY_MIGRATION_DOC_LINK, utilitySheetName, sheetLastRow, workbookLinkCol, workbookLink);
}

function handleE2EMigration(workbookData, workbookLink) {
  //create the email to send
  var emailAddress =  "email here"; 
  var entrataEmail = Session.getActiveUser().getEmail();
  var ccEmail = "email here";
  var subject = "(E2E Migration) Please Migrate the following site: " + workbookData.companyName + " - " + workbookData.companyId + " / " + workbookData.siteName + " - " + workbookData.siteId;
  var emailSig = getEmailSig(entrataEmail);
  
  //Create the email for the E2E migration.
  var emailBody = "<b>Company Name: </b> " + workbookData.companyName + " - " + workbookData.companyId + "<br />"
                  + "<b>Site Name: </b> " + workbookData.siteName + " - " + workbookData.siteId + "<br />"
                  + "<b>As of Date: </b> " + workbookData.asOfDate + "<br />"
                  + "<b>Property Type: </b> " + workbookData.propertyType + "<br />"
                  + "<b>Migration Type: </b>" + workbookData.migrationType + "<br />"
                  + "<b>Unit: </b> " + workbookData.units + "<br />"
                  + "<b>Space: </b> " + workbookData.spaces + "<br />"
                  + "<br />" + "<b>Former Property Management Software: </b> " + workbookData.formerPropertySoftware + "<br />"
                  + "<br />" + "<b>Source Company: </b> " + workbookData.originalCompanyEte + "<br />"
                  + "<br />" + "<b>Source Site: </b> " + workbookData.originalSiteEte + "<br />"
                  + "<br />" + "<b>Expected Return Date: </b> " + workbookData.expectedReturnDate + "<br />"
                  + "<br />" + "<b>Checklist Link: </b> " + workbookLink + "<br />"
                  + "<br />" + "Please find the reports in the FTP Site" + "<br />" + "<br />";
                  
  //Add the email signature to the email message.
  var emailMessage = emailBody + emailSig;
  
  var email = {
    emailAddress: emailAddress,
    entrataEmail: entrataEmail,
    ccEmail: ccEmail,
    subject: subject,
    emailSig: emailSig,
    emailMessage: emailMessage
  }
  
  sendEmail(email);
                  
}

function handleAllMigrations(workbookData, workbookLink) {
  var emailBody = "<b>Company Name:</b> " + workbookData.companyName + " - " + workbookData.companyId + "<br />"
                    + "<b>Site Name:</b> " + workbookData.siteName + " - " + workbookData.siteId + "<br />"
                    + "<b>As of Date:</b> " + workbookData.asOfDate + "<br />"
                    + "<b>Property Type:</b> " + workbookData.propertyType + "<br />"
                    + "<b>Migration Type:</b>" + workbookData.migrationType + "<br />"
                    + "<b>Unit:</b> " + workbookData.units + "<br />"
                    + "<b>Space:</b> " + workbookData.spaces + "<br />"
                    + "<b>Former Property Management Software:</b> " + workbookData.formerPropertySoftware + "<br />"
                    + "<br />" + "<b>Expected Return Date:</b> " + workbookData.expectedReturnDate + "<br />"
                    + "<br />" + "<b>Checklist Link:</b> " + workbookLink + "<br />"
                    + "<br />" + "Please find the reports in the FTP Site" + "<br />";
  var emailAddress =  "email here"; 
  var entrataEmail = Session.getActiveUser().getEmail();
  var ccEmail = "email here";
  var subject = "Please Migrate the following site: " + workbookData.companyName + " - " + workbookData.companyId + " / " + workbookData.siteName + " - " + workbookData.siteId;
  var emailSig = getEmailSig(entrataEmail);
  
  //Add the email signature to the email message.
  var emailMessage = emailBody + emailSig;
    
  var email = {
    emailAddress: emailAddress,
    entrataEmail: entrataEmail,
    ccEmail: ccEmail,
    subject: subject,
    emailSig: emailSig,
    emailMessage: emailMessage
  }
  
  sendEmail(email);
}




