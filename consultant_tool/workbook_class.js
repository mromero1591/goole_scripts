function getWorkbookData(workbook) {
  //VARIABLES
  var IN_FRONT = 1;
  var TWO_IN_FRONT = 2;
  var BEHIND = -1;
  var MAIN_DATA_SEARCH_LOCATION = 1;
  var E_TWO_E_DATA_SEARCH_LOCATION = 3;
  var VALIDATION_MIGRATION_NAME_SEARCH_LOCATION = 2;
  var MIGRATION_DATE_DATA_SERACH_LOCATION = 5;
  
  //values in pre start
  var CONSULTANT_NAME_SEARCH = "Core Consultant Name";
  var ACCOUNTING_CONSULTANT_NAME_SEARCH = "ACC Name";
  var COMPANY_NAME_SEARCH = "Company Name and ID";
  var SITE_NAME_SEARCH = "Site Name and ID";
  var AS_OF_DATE_SEARCH = "As of Date";
  var PROPERTY_TYPE_SEARCH = "Property Type";
  var UNITS_SEARCH = "Units";
  var SPACES_SEARCH = "Spaces";
  var FORMER_PROPERTY_SOFTWARE_SEARCH = "Former Property Management Software";
  var EXPECTED_RETURN_DATE_SEARCH = "Expected Return Date";
  var MIGRATION_TYPE_SEARCH = "Migration Type"
  var MIGRATION_VALIDATION_NAME = "Validation email sent by";
  var SPECIAL_INSTRUCTIONS_SEARCH = "Special Migration Instructions";
  var MIGRATION_TYPE_SEARCH = "Migration Type";
  var QUERY_DOC_SEARCH = 'Queries Doc Link';
  
  var PRE_START_SHEET_NAME = "Pre Start Checklist";
  var POST_MIGRATION_SHEET_NAME = "Post Migration Checklist";
  var ENTRATA_ACCOUNTING_SHEET_NAME = "Entrata Accounting";
  var company_POST_MIGRATION_SHEET_NAME = "Company Post Migration";
  var YARDI_CHECKLIST_NAME = "Yardi Integrated Checklist";
  var E2E_CHECKLIST_NAME = "E2E Checklist";
  
  //Make the migration workbook active and create a variable for the prestart checklist sheet.
  var preStartSheet = workbook.getSheetByName(PRE_START_SHEET_NAME);
  var postMigrationSheet = workbook.getSheetByName(POST_MIGRATION_SHEET_NAME);
  var entrataAccounting = workbook.getSheetByName(ENTRATA_ACCOUNTING_SHEET_NAME);
  var companyPostMigrationSheet = workbook.getSheetByName(company_POST_MIGRATION_SHEET_NAME);
  var yardiChecklistSheet = workbook.getSheetByName(YARDI_CHECKLIST_NAME);
  var e2eSheet = workbook.getSheetByName(E2E_CHECKLIST_NAME);

  //Get the Data from the workbook
  var companyName = searchSheet(COMPANY_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var companyId = searchSheet(COMPANY_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, TWO_IN_FRONT);
  var siteName = searchSheet(SITE_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var siteId = searchSheet(SITE_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, TWO_IN_FRONT);
  var consultantName = searchSheet(CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var accountConsultantName = searchSheet(ACCOUNTING_CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  if (accountConsultantName == "NA")
  accountConsultantName = " ";
  
  if(searchSheet(AS_OF_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT) != ""){
    var asOfDate = Utilities.formatDate(searchSheet(AS_OF_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT), " / ", "MM/dd/yyyy"); 
  }
  else{
    var asOfDate = "";
  }
  var propertyType = searchSheet(PROPERTY_TYPE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var units = searchSheet(UNITS_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var spaces = searchSheet(SPACES_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var formerPropertySoftware = searchSheet(FORMER_PROPERTY_SOFTWARE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  
  if(searchSheet(EXPECTED_RETURN_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT) != ""){
    var expectedReturnDate = Utilities.formatDate(searchSheet(EXPECTED_RETURN_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT), " / ", "MM/dd/yyyy");
  }else {
    var expectedReturnDate = "";
  }
  
  var migrationType = searchSheet(MIGRATION_TYPE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var specialInstructions = searchSheet(SPECIAL_INSTRUCTIONS_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var validationMigrationConsultantName = searchSheet(MIGRATION_VALIDATION_NAME, POST_MIGRATION_SHEET_NAME, VALIDATION_MIGRATION_NAME_SEARCH_LOCATION, IN_FRONT);
  var queryDocLink = searchSheet(QUERY_DOC_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  
  var workbookData = {
    accountConsultantName: accountConsultantName,
    companyName: companyName,
    companyId: companyId,
    consultantName: consultantName,
    entrataAccounting: entrataAccounting,
    e2eSheet: e2eSheet,
    formerPropertySoftware: formerPropertySoftware,
    preStartSheet: preStartSheet,
    postMigrationSheet: postMigrationSheet,
    companyPostMigrationSheet: companyPostMigrationSheet,
    yardiIntegratedSheet: yardiChecklistSheet,
    siteName: siteName,
    siteId: siteId,
    validationMigrationConsultantName: validationMigrationConsultantName,
    specialInstructions: specialInstructions,
    migrationType: migrationType,
    queryDocLink: queryDocLink
  };
  return workbookData;
}
