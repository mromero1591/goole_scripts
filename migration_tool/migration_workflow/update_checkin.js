/*Name: Mark Romero
* Date: 2/23/2016
* Updated: 12/20/2016
* Purpose: This script file contains the code to update the check in sheet in the quailty workbook.
*==============================================================================*/

/*Purpose: Update the check in document
* Parameters:None
* Returns : None
*==============================================================================*/
function updateCheckIn(workbookData) {
   //Store the Quality sheet url in a variable.
   var QUALITY_SHEET_URL = "link here";
   
   /*Save the Data Check in sheet into a variable by:
    *1.Creating a variable for the name of the Data check in sheet.
    *2.Creating a variable that holds the Data check in sheet object.
   */
   var SHEET_NAME = "Data Check-Ins";
   var dSheet = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME);
    
   /*Get the columns where we will be inputing data using the get column by name function. Do this for the following columns:
    *1.Migration Type
    *2.Actual check in date
    *3.Company name
    *4.Site name
    *5.Migrated units count
    *6.Property Type
    *7.Previous software
    *8.Consultant, migration consultant, and validation consultant
    *9.Date Site was sent to the migration consultant
    *10.time Site was sent to the migration consultant
   */
   var dMigrationTypeColumn = getColumnByName(dSheet, "Migration type - CSV/Automated");
   var dActualCheckInDateColumn = getColumnByName(dSheet, "Actual CheckIn Date");
   var dCompanyColumn = getColumnByName(dSheet, "Company");
   var dSiteColumn = getColumnByName(dSheet, "Site");
   var dMigratedUnitsColumn = getColumnByName(dSheet, "Migrated Units");
   var dPropertyTypeColumn = getColumnByName(dSheet, "Property Type");
   var dPreviousSoftwareColumn = getColumnByName(dSheet, "Previous Software");
   var dConsultantColumn = getColumnByName(dSheet, "Consultant");
   var dAccConsultantColumn = getColumnByName(dSheet, "Accounting Consultant");
   var dAgentColumn = getColumnByName(dSheet, "Migration Consultant");
   var dValidationAgentColumn = getColumnByName(dSheet, "Validation Consultant");
   var dDateConsultantSentSite = getColumnByName(dSheet, "Date Consultant sent site");
   var dTimeConsultantSentSite = getColumnByName(dSheet, "Time Consultant sent site");
   var dTakeoverColumn = getColumnByName(dSheet, "Takeover");
     
   //Get the last row in the check in sheet with data.
   var qSheetLastRow = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getLastRow();
   
   var dateConsultantSentSite = workbookData.preStartSheet.getRange(4, 1).getValue();
   if (dateConsultantSentSite != "")
   {
     var dateConsultantSentSiteArray = dateConsultantSentSite.toString().split(".");
   }

   //input the data from the pre-start checklist into the data check in sheet, by doing the following for each info.
   //Open the quaility sheet >> get the sheet by its name >> get the range (the column and the first empty row) >> set the value
   //the Date and time the consultant sent the site will be different, you will first need to split the two and then add the info
   SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dActualCheckInDateColumn).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dCompanyColumn, workbookData.companyName);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dSiteColumn, workbookData.siteName);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dMigratedUnitsColumn, workbookData.units);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dPropertyTypeColumn, workbookData.propertyType);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dPreviousSoftwareColumn, workbookData.formerPropertySoftware);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dConsultantColumn, workbookData.consultantName);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dAccConsultantColumn, workbookData.accountConsultantName);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dMigrationTypeColumn, workbookData.migrationType);
   setData(QUALITY_SHEET_URL,SHEET_NAME, qSheetLastRow, dAgentColumn, workbookData.checkInAgent);
   
   if (dateConsultantSentSite != "") 
   {
     SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dDateConsultantSentSite).setValue(dateConsultantSentSiteArray[0]);
     SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dTimeConsultantSentSite).setValue(dateConsultantSentSiteArray[1]);
   }
   else
   {
     SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dDateConsultantSentSite).setValue(dateConsultantSentSite);
   }
   SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dTakeoverColumn).setValue(workbookData.takeover);

   //Save the location of the row we just imported, by adding it into the workbook, this is done so that we can later add the validation agent to the checkin sheet.
   var locationOfValidationCell = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(SHEET_NAME).getRange(qSheetLastRow + 1, dValidationAgentColumn).getA1Notation();
   workbookData.preStartSheet.getRange(3, 1).setFontColor('white').setValue(locationOfValidationCell);
}

function updateQuailtyCheck(workbookData) {    
    
    var entrataEmail = Session.getActiveUser().getEmail();
    
    //Quality sheet location.
    var qualitySheetUrl = "link here";
    var sheetName = "Quality Checks";
      
    /* the data spots for the quality check sheet*/
    var qSheet = SpreadsheetApp.openByUrl(qualitySheetUrl).getSheetByName(sheetName);  
    Logger.log("Test");
    var qMigrationTypeColumn = getColumnByName(qSheet, "Migration type - CSV/Automated");
    var qRetrunDateColumn = getColumnByName(qSheet, "Actual Return Date");
    var qCompanyColumn = getColumnByName(qSheet, "Company");
    var qSiteColumn = getColumnByName(qSheet, "Site");
    var qBilledUnitsColumn = getColumnByName(qSheet, "Billed Units");
    var qMigratedUnitsColumn = getColumnByName(qSheet, "Migrated Units");
    var qPreviousSoftwareColumn = getColumnByName(qSheet, "Previous Software");
    var qPropertyTypeColumn = getColumnByName(qSheet, "Migration Type");
    var qDataDueDateColumn = getColumnByName(qSheet, "When was the data due to Entrata");
    var qDataRecieveDateColumn = getColumnByName(qSheet, "When did we actualy receive the data from the client");
    var qConsultantNameColumn = getColumnByName(qSheet, "Consultant");
    var qAgentColumn = getColumnByName(qSheet, "QC Agent");
    var qAccountingConsultantColumn = getColumnByName(qSheet, "Accounting Consultant");
    var qcompanyMigrationAgent = getColumnByName(qSheet, "Migration Agent Name");
    var qcompanySecondAgent = getColumnByName(qSheet,"2nd Level QA Agent Name");
  Logger.log("Test");
    //get last row with data.
    var qSheetLastRow = SpreadsheetApp.openByUrl(qualitySheetUrl).getSheetByName(sheetName).getLastRow();
    var migrationAgent = getCompletedConsultant(entrataEmail);
    
    var companyAgent = workbookData.companyPostMigrationSheet.getRange(5,2).getValue();
    var secondLevelcompanyAgent = workbookData.companyPostMigrationSheet.getRange(6,2).getValue();
  
    /*****Set the data onto the quality work book ************************************************
    ******DO NOT CHANGE these, any changes made can be changed by changing the vaiable ***********/
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qMigrationTypeColumn, workbookData.migrationType);
    SpreadsheetApp.openByUrl(qualitySheetUrl).getSheetByName(sheetName).getRange(qSheetLastRow + 1, qRetrunDateColumn).setValue(Utilities.formatDate(new Date(), " / ","MM/dd/yyyy")); 
    
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qCompanyColumn, workbookData.companyName);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qSiteColumn, workbookData.siteName);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qBilledUnitsColumn, workbookData.units);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qMigratedUnitsColumn, workbookData.spaces);
    
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qPreviousSoftwareColumn, workbookData.formerPropertySoftware);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qPropertyTypeColumn, workbookData.propertyType);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qDataDueDateColumn, workbookData.asOfDate);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qConsultantNameColumn, workbookData.consultantName);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qAgentColumn, migrationAgent);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qAccountingConsultantColumn, workbookData.accountConsultantName);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qDataRecieveDateColumn, workbookData.dataReviedDate);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qcompanyMigrationAgent, companyAgent);
    setData(qualitySheetUrl,sheetName, qSheetLastRow, qcompanySecondAgent, secondLevelcompanyAgent);
    
    //Check in tab
    var dsheetName = "Data Check-Ins";
    var dSheet = SpreadsheetApp.openByUrl(qualitySheetUrl).getSheetByName(dsheetName);
    var validationAgentCheckIncell = workbookData.preStartSheet.getRange(3,1).getValue();
  
  SpreadsheetApp.openByUrl(qualitySheetUrl).getSheetByName(dsheetName).getRange(validationAgentCheckIncell).setValue(migrationAgent);
       
  }
  
  function updateAccountingQC(workbookData, workbookLink) { 
    //Save the Quality sheet url and name, then create an object using the sheet url and name.
    var QUALITY_SHEET_URL = "link here";
    var Q_SHEET_NAME = "Core & AC QC";
    var dSheet = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME);
    
    /* the data spots for the Data check-in sheet*/
    var dmigrationWeekColumn = getColumnByName(dSheet, "Migration Week");
    var dCompanyColumn = getColumnByName(dSheet, "Company");
    var dPropertyTypeColumn = getColumnByName(dSheet, "Property");
    var dConsultantColumn = getColumnByName(dSheet, "Core Consultant");
    var dAccountinConsultantColumn = getColumnByName(dSheet, "Accounting Consultant");
    var dUsingCoreAccountingColumn = getColumnByName(dSheet, "Using Core Accounting");
    var dPreviousAccountingSystemColumn = getColumnByName(dSheet, "Previous Accounting System");
    var dPropertyStypeColumn = getColumnByName(dSheet, "Property Type");
    var dpostReviewDateColumn = getColumnByName(dSheet, "Post Review Due Date");
    var dQCDocColumn = getColumnByName(dSheet, "QC Review Doc Link");
    var daccQualityControlReviewerColumn = getColumnByName(dSheet, "ACC Quality Control Reviewer");
    var dStatusColumn = getColumnByName(dSheet, "Status");
    var dQcGradedColumn = getColumnByName(dSheet, "QC Graded");
    var dlandrQualityControlReviewerColumn = getColumnByName(dSheet, "Core Quality Control Reviewer");
    var dlandrStautsColumn = 15;//getColumnByName(dSheet, "L&R Quality Control Reviewe Status");
    var dLrQcGradedColumn = 16;//getColumnByName(dSheet, "L&R QC Graded");
    
    //get the last row in the sheet.
    var qSheetLastRow = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getLastRow();
    
    //convert accounting and core names.
    var accountConsultant = convertConsultantName(workbookData.accountConsultantName);
    var consultantName = convertConsultantName(workbookData.consultantName);
    
    if(workbookData.usingCoreAccounting == "Correct") {
      workbookData.usingCoreAccounting = "No";
    }
    else {
      workbookData.usingCoreAccounting = "Yes";
    }
    
    var coreQualityControlReviewer = findQcReviewer(consultantName);
   
    /*****Set the data onto the quality work book ************************************************
    ******DO NOT CHANGE these, any changes made can be changed by changing the vaiable ***********/
    SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dmigrationWeekColumn).setValue(workbookData.migrationDate);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dCompanyColumn, workbookData.companyName);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dPropertyTypeColumn, workbookData.siteName);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dConsultantColumn, consultantName);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dAccountinConsultantColumn, accountConsultant);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dUsingCoreAccountingColumn, workbookData.usingCoreAccounting);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dPreviousAccountingSystemColumn, workbookData.formerPropertySoftware);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dPropertyStypeColumn, workbookData.propertyType);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dQCDocColumn, workbookLink);
    setData(QUALITY_SHEET_URL, Q_SHEET_NAME, qSheetLastRow, dlandrQualityControlReviewerColumn, coreQualityControlReviewer);
 
    var accQuailyControlReviewerCell = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, daccQualityControlReviewerColumn);
    var accQuailyControlReviewerValidationRange = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName('Lists').getRange('F4:F17');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(accQuailyControlReviewerValidationRange).build();
    accQuailyControlReviewerCell.setDataValidation(rule);
    
    var statusCell = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dStatusColumn);
    var statusCellValidationRange = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName('Lists').getRange('A4:A8');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(statusCellValidationRange).build();
    statusCell.setDataValidation(rule);
    
    var landrStautsCell = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dlandrStautsColumn);
    var statusCellValidationRange = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName('Lists').getRange('A4:A8');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(statusCellValidationRange).build();
    landrStautsCell.setDataValidation(rule);
       
    var usingCoreAccountingCell = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dUsingCoreAccountingColumn);
    var usingCoreAccountingRange = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName('Lists').getRange('I4:I5');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(usingCoreAccountingRange).build();
    usingCoreAccountingCell.setDataValidation(rule);
    
    
    var L3 = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dStatusColumn).getA1Notation();
    var A3 = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dmigrationWeekColumn).getA1Notation();
    var J35 = SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dQCDocColumn).getA1Notation();
    
    SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dpostReviewDateColumn).setFormula('=IF('+L3+'="GoLive Review",'+A3+'+30,"")+IF('+L3+'="Post Review",'+A3+'+30,"")+IF('+L3+'="Completed",'+A3+'+30,"")');
    SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dQcGradedColumn).setFormula('=IMPORTRANGE('+J35+',"ACC QC!H37")/IMPORTRANGE('+J35+',"ACC QC!G37")');
    SpreadsheetApp.openByUrl(QUALITY_SHEET_URL).getSheetByName(Q_SHEET_NAME).getRange(qSheetLastRow + 1, dLrQcGradedColumn).setFormula('=IMPORTRANGE('+J35+',"Core QC!G27")');
   
   
    
    
    
 }

