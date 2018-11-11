function updateDataFixSheet(workbookData, dataFix, workbookLink) {
  //Data Fix sheet.
  var dataFixSheetUrl = "link here";
  var dataFixsheetName = "Data Fix Request";
  var dataFixsheet = SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName);
  
  /* the data spots for the data fix sheet*/  
  var reportedOnColumn = getColumnByName(dataFixsheet, "Reported On");
  var reportedByColumn = getColumnByName(dataFixsheet, "Reported By");
  var coreConsultantColumn = getColumnByName(dataFixsheet, "Core Consultant");
  var acctConsultantColumn = getColumnByName(dataFixsheet, "ACCT Consultant");
  var companyColumn = getColumnByName(dataFixsheet, "Company");
  var companyIdColumn = getColumnByName(dataFixsheet, "Company Id");
  var siteColumn = getColumnByName(dataFixsheet, "Site");
  var siteIdColumn = getColumnByName(dataFixsheet, "Site Id");
  var migrationConsultantColumn = getColumnByName(dataFixsheet, "Migration Consultant");
  var fieldColumn = getColumnByName(dataFixsheet, "Field");
  var fixTypeColumn = getColumnByName(dataFixsheet, "Fix Type");
  var descriptionColumn = getColumnByName(dataFixsheet, "Description");
  var examplesColumn = getColumnByName(dataFixsheet, "Examples");
  var requiredColumn = getColumnByName(dataFixsheet, "Required for release");
  var workColumn = getColumnByName(dataFixsheet, "Workbook");
  var SoftwareColumn = getColumnByName(dataFixsheet, "Software"); 
  var migrationTypeColumn = getColumnByName(dataFixsheet, "Migration type");
  var numOfFixesColumn = getColumnByName(dataFixsheet, "Number of Fixes");
  
  //get last row with data.
  var dataFixLastRow = SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getLastRow();

  /*****Set the data onto the quality work book ************************************************
  ******DO NOT CHANGE these, any changes made can be changed by changing the vaiable ***********/
  SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixsheetName).getRange(dataFixLastRow + 1, reportedOnColumn).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, reportedByColumn, dataFix.reportedByName);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, coreConsultantColumn, workbookData.consultantName);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, acctConsultantColumn, workbookData.accountConsultantName);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, companyColumn, workbookData.companyName);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, companyIdColumn, workbookData.companyId);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, siteColumn, workbookData.siteName);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, siteIdColumn, workbookData.siteId);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, migrationConsultantColumn, workbookData.validationMigrationConsultantName);
  if(dataFix.issueType == "Other") {
    setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, fieldColumn, dataFix.otherText);
  } else {
    setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, fieldColumn, dataFix.issueType);
  }
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, fixTypeColumn, dataFix.fixType);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, descriptionColumn, dataFix.description);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, examplesColumn, dataFix.examples);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, requiredColumn, dataFix.emergency);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, workColumn, workbookLink);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, SoftwareColumn, workbookData.formerPropertySoftware);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, migrationTypeColumn, workbookData.migrationType);
  setData(dataFixSheetUrl, dataFixsheetName, dataFixLastRow, numOfFixesColumn, dataFix.numOfFixes);
  
}

function updateQueryDoc(workbookData, dataFix, workbookLink) {
  //Query sheet.
  var workbook = SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openByUrl(workbookLink));
  var dataFixSheetName = "Data Fixes"
  //Open the Query Doc Link
  var queryDoc = SpreadsheetApp.openByUrl(workbookData.queryDocLink);
  
  //check to see if the Data Fix sheet exist
  var dataFixsheet = queryDoc.getSheetByName(dataFixSheetName);
  if(!dataFixsheet) {
    //If the data fix sheet does not exist create it.
    createDataFixSheet(queryDoc, dataFixSheetName);  
    dataFixsheet = queryDoc.getSheetByName(dataFixSheetName);
  }
  
  var dataFixSheet = queryDoc.setActiveSheet(dataFixsheet);
  
  var dataFixSheetUrl = workbookData.queryDocLink;
  Logger.log(dataFixSheetUrl);
  
 //set the data fix
 /* the data spots for the data fix sheet*/  
 var reportedOnColumn = getColumnByName(dataFixsheet, "Reported On");
 var reportedByColumn = getColumnByName(dataFixsheet, "Reported By");
 var coreConsultantColumn = getColumnByName(dataFixsheet, "Core Consultant");
 var acctConsultantColumn = getColumnByName(dataFixsheet, "ACCT Consultant");
 var companyColumn = getColumnByName(dataFixsheet, "Company");
 var companyIdColumn = getColumnByName(dataFixsheet, "Company Id");
 var siteColumn = getColumnByName(dataFixsheet, "Site");
 var siteIdColumn = getColumnByName(dataFixsheet, "Site Id");
 var migrationConsultantColumn = getColumnByName(dataFixsheet, "Migration Consultant");
 var fieldColumn = getColumnByName(dataFixsheet, "Field");
 var fixTypeColumn = getColumnByName(dataFixsheet, "Fix Type");
 var descriptionColumn = getColumnByName(dataFixsheet, "Description");
 var examplesColumn = getColumnByName(dataFixsheet, "Examples");
 var requiredColumn = getColumnByName(dataFixsheet, "Required for release");
 var workColumn = getColumnByName(dataFixsheet, "Workbook");
 var SoftwareColumn = getColumnByName(dataFixsheet, "Software"); 
 var migrationTypeColumn = getColumnByName(dataFixsheet, "Migration type");
 var numOfFixesColumn = getColumnByName(dataFixsheet, "Number of Fixes");
 
 //get last row with data.
 var dataFixLastRow = SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixSheetName).getLastRow();
 /*****Set the data onto the quality work book ************************************************
 ******DO NOT CHANGE these, any changes made can be changed by changing the vaiable ***********/
 SpreadsheetApp.openByUrl(dataFixSheetUrl).getSheetByName(dataFixSheetName).getRange(dataFixLastRow + 1, reportedOnColumn).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, reportedByColumn, dataFix.reportedByName);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, coreConsultantColumn, workbookData.consultantName);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, acctConsultantColumn, workbookData.accountConsultantName);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, companyColumn, workbookData.companyName);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, companyIdColumn, workbookData.companyId);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, siteColumn, workbookData.siteName);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, siteIdColumn, workbookData.siteId);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, migrationConsultantColumn, workbookData.validationMigrationConsultantName);
 if(dataFix.issueType == "Other") {
   setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, fieldColumn, dataFix.otherText);
 } else {
   setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, fieldColumn, dataFix.issueType);
 }
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, fixTypeColumn, dataFix.fixType);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, descriptionColumn, dataFix.description);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, examplesColumn, dataFix.examples);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, requiredColumn, dataFix.emergency);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, workColumn, workbookLink);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, SoftwareColumn, workbookData.formerPropertySoftware);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, migrationTypeColumn, workbookData.migrationType);
 setData(dataFixSheetUrl, dataFixSheetName, dataFixLastRow, numOfFixesColumn, dataFix.numOfFixes);

}


function createDataFixSheet(queryDoc, dataFixSheetName) {
  queryDoc.insertSheet(dataFixSheetName);
  var dataFixSheet = queryDoc.getSheetByName('Data Fixes');
  dataFixSheet.getRange('A1').setValue("Reported On")
  dataFixSheet.getRange('B1').setValue("Reported By")
  dataFixSheet.getRange('C1').setValue("Core Consultant")
  dataFixSheet.getRange('D1').setValue("ACCT Consultant")
  dataFixSheet.getRange('E1').setValue("Company")
  dataFixSheet.getRange('F1').setValue("Company Id")
  dataFixSheet.getRange('G1').setValue("Site")
  dataFixSheet.getRange('H1').setValue("Site Id")
  dataFixSheet.getRange('I1').setValue("Migration Consultant")
  dataFixSheet.getRange('J1').setValue("Field")
  dataFixSheet.getRange('K1').setValue("Fix Type")
  dataFixSheet.getRange('L1').setValue("Description")
  dataFixSheet.getRange('M1').setValue("Examples")
  dataFixSheet.getRange('N1').setValue("Required for release")
  dataFixSheet.getRange('O1').setValue("Workbook")
  dataFixSheet.getRange('P1').setValue("Software")
  dataFixSheet.getRange('Q1').setValue("Migration type")
  dataFixSheet.getRange('R1').setValue("Number of Fixes")
}

