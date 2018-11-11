function getWorkbookData(workbook) {
  //VARIABLES
  var IN_FRONT = 1;
  var TWO_IN_FRONT = 2;
  var BEHIND = -1;
  var MAIN_DATA_SEARCH_LOCATION = 1;
  var E_TWO_E_DATA_SEARCH_LOCATION = 3;
  var VALIDATION_MIGRATION_NAME_SEARCH_LOCATION = 2;
  var MIGRATION_DATE_DATA_SERACH_LOCATION = 5;
  var GATHER_FROM_CLIENT_SEARCH_LOCATION = 4;

  var ACC_QC_SHEET_NAME = "ACC QC";
  var ACC_QC_NAME_SEARCH = "ACC Performing QC";
  var ACCOUNTING_CONSULTANT_NAME_SEARCH = "ACC Name";
  var AS_OF_DATE_SEARCH = "As of Date";
  var CHECK_IN_AGENT_SEARCH = "Migration Consultant";
  var COMPANY_NAME_SEARCH = "Company Name and ID";
  var CONSULTANT_NAME_SEARCH = "Core Consultant Name";
  var E2E_CHECKLIST_NAME = "E2E Checklist";
  var ENTRATA_ACCOUNTING_SHEET_NAME = "Entrata Accounting";
  var EXPECTED_RETURN_DATE_SEARCH = "Expected Return Date";
  var FORMER_PROPERTY_SOFTWARE_SEARCH = "Former Property Management Software";
  var MIGRATION_DATE_SEARCH = " Migration Date: ";
  var MIGRATION_TYPE_SEARCH = "Migration Type";
  var MIGRATION_VALIDATION_NAME = "Validation email sent by";
  var ORIGNAL_COMPANY_NAME_SEARCH = "SOURCE Company Name and id - ";
  var ORIGNAL_SITE_NAME_SEARCH = "SOURCE Site Name and id - ";
  var POST_MIGRATION_SHEET_NAME = "Post Migration Checklist";
  var PRE_START_SHEET_NAME = "Pre Start Checklist";
  var PROPERTY_TYPE_SEARCH = "Property Type";
  var QUERY_DOC_LINK_SEARCH = "Queries Doc Link";
  var REPORTS_LINK_SEARCH = "Reports are Located at";
  var SECONDARY_ACCOUNTING_CONSULTANT_SEARCH = "Secondary ACC Name";
  var SECONDARY_CORE_CONSULTANT_SEARCH = "Secondary Core Consultant Name";
  var SPECIAL_INSTRUCTIONS_SEARCH = "Special Migration Instructions";
  var SITE_NAME_SEARCH = "Site Name and ID";
  var SPACES_SEARCH = "Spaces";
  var TAKEOVER_SEARCH = "Is this a Takeover";
  var UNITS_SEARCH = "Units";
  var USING_CORE_ACCOUNTING_SEARCH = "Verified by (Client Name):";
  var company_POST_MIGRATION_SHEET_NAME = "company Post Migration";
  var YARDI_CHECKLIST_NAME = "Yardi/UDS";
  var UDS_FILE_LOCATION_SEARCH = "UDS Back File Location";
  var UDS_FILES_NAME_SEARCH = "UDS File Name";
  var UDS_BACKUP_RESTORED_SERACH = "Development - Has Back Up been restored";
  var PROPERTY_LOOK_UP_CODE_SEARCH = "Property Lookup Code";
  
  
  //Make the migration workbook active and create a variable for the prestart checklist sheet.
  var preStartSheet = workbook.getSheetByName(PRE_START_SHEET_NAME);
  var postMigrationSheet = workbook.getSheetByName(POST_MIGRATION_SHEET_NAME);
  var entrataAccounting = workbook.getSheetByName(ENTRATA_ACCOUNTING_SHEET_NAME);
  var companyPostMigrationSheet = workbook.getSheetByName(company_POST_MIGRATION_SHEET_NAME);
  var yardiChecklistSheet = workbook.getSheetByName(YARDI_CHECKLIST_NAME);
  var e2eSheet = workbook.getSheetByName(E2E_CHECKLIST_NAME);
  var accQcSheet = workbook.getSheetByName(ACC_QC_SHEET_NAME);

  //Get the Data from the workbook
  var companyName = searchSheet(COMPANY_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var companyId = searchSheet(COMPANY_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, TWO_IN_FRONT);
  var siteName = searchSheet(SITE_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var siteId = searchSheet(SITE_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, TWO_IN_FRONT);
  var consultantName = searchSheet(CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  
  var secondaryConsultantName = searchSheet(SECONDARY_CORE_CONSULTANT_SEARCH, PRE_START_SHEET_NAME, IN_FRONT);
  if (secondaryConsultantName == undefined) {
    secondaryConsultantName = "NA";
  }
  
  var accountConsultantName = searchSheet(ACCOUNTING_CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  if (accountConsultantName == "NA")
  accountConsultantName = " ";
  
  var secondaryAccountingConsultantName = searchSheet(SECONDARY_ACCOUNTING_CONSULTANT_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  if (secondaryAccountingConsultantName == undefined) {
    secondaryAccountingConsultantName = "NA";
  }
  
  var accountingConsultantQcName = searchSheet(ACC_QC_NAME_SEARCH, ACC_QC_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  if (accountingConsultantQcName == undefined) {
    accountingConsultantQcName = "NA";
  }
  
  if(searchSheet(AS_OF_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT) != ""){
    var asOfDate = Utilities.formatDate(searchSheet(AS_OF_DATE_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT), " / ", "MM/dd/yyyy"); 
  }else{
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
  
  if(searchSheet(MIGRATION_DATE_SEARCH, ENTRATA_ACCOUNTING_SHEET_NAME, MIGRATION_DATE_DATA_SERACH_LOCATION, IN_FRONT) != ""){
    var migrationDate = Utilities.formatDate(searchSheet(MIGRATION_DATE_SEARCH, ENTRATA_ACCOUNTING_SHEET_NAME, MIGRATION_DATE_DATA_SERACH_LOCATION, IN_FRONT), " / ", "MM/dd/yyyy");
  }else {
    var migrationDate = "";
  }
  
  if(searchSheet(CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, BEHIND) != ""){
    var dataReviedDate = Utilities.formatDate(searchSheet(CONSULTANT_NAME_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, BEHIND), " / ", "MM/dd/yyyy");
  }else {
    var dataReviedDate = "";
  }
  
  var queryDocLinkCell = getCell(QUERY_DOC_LINK_SEARCH, PRE_START_SHEET_NAME, IN_FRONT);
  var queryDocUrl = searchSheet(QUERY_DOC_LINK_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var originalCompanyEte = searchSheet(ORIGNAL_COMPANY_NAME_SEARCH, E2E_CHECKLIST_NAME, E_TWO_E_DATA_SEARCH_LOCATION, IN_FRONT);
  var originalSiteEte = searchSheet(ORIGNAL_SITE_NAME_SEARCH, E2E_CHECKLIST_NAME, E_TWO_E_DATA_SEARCH_LOCATION, IN_FRONT);
  var takeover = searchSheet(TAKEOVER_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var checkInAgent = searchSheet(CHECK_IN_AGENT_SEARCH, PRE_START_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);
  var usingCoreAccounting = searchSheet(USING_CORE_ACCOUNTING_SEARCH, ENTRATA_ACCOUNTING_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, TWO_IN_FRONT);
  var reportsLink = searchSheet(REPORTS_LINK_SEARCH, POST_MIGRATION_SHEET_NAME, MAIN_DATA_SEARCH_LOCATION, IN_FRONT);  
  var propertyLookUpCode = "";
  if(migrationType == "UDS" || migrationType == "Hybrid" || migrationType == "Utility") {
    propertyLookUpCode = searchSheet(PROPERTY_LOOK_UP_CODE_SEARCH, YARDI_CHECKLIST_NAME, GATHER_FROM_CLIENT_SEARCH_LOCATION, IN_FRONT);
  }
  
  
  //Data for UDS Migrations
  var udsFileLocation = "";
  var udsFileName = "";
  var udsBackupRestored = "";
  if(migrationType == "UDS") {
    udsFileLocation = searchSheet(UDS_FILE_LOCATION_SEARCH, YARDI_CHECKLIST_NAME, GATHER_FROM_CLIENT_SEARCH_LOCATION, IN_FRONT);
    udsFileName = searchSheet(UDS_FILES_NAME_SEARCH, YARDI_CHECKLIST_NAME, GATHER_FROM_CLIENT_SEARCH_LOCATION, IN_FRONT);
    udsBackupRestored = searchSheet(UDS_BACKUP_RESTORED_SERACH, YARDI_CHECKLIST_NAME, GATHER_FROM_CLIENT_SEARCH_LOCATION, IN_FRONT);
  }
  
  var workbookData = {
    accountConsultantName: accountConsultantName,
    accountingConsultantQcName: accountingConsultantQcName,
    asOfDate: asOfDate,
    checkInAgent: checkInAgent,
    companyId: companyId,
    companyName: companyName,
    consultantName: consultantName,
    dataReviedDate: dataReviedDate,
    e2eSheet: e2eSheet,
    entrataAccounting: entrataAccounting,
    expectedReturnDate: expectedReturnDate,
    formerPropertySoftware: formerPropertySoftware,
    migrationDate: migrationDate,
    migrationType: migrationType,
    originalCompanyEte: originalCompanyEte,
    originalSiteEte: originalSiteEte,
    postMigrationSheet: postMigrationSheet,
    preStartSheet: preStartSheet,
    propertyType: propertyType,
    queryDocLinkCell: queryDocLinkCell,
    queryDocUrl: queryDocUrl,
    reportsLink: reportsLink,
    secondaryAccountingConsultantName: secondaryAccountingConsultantName,
    secondaryConsultantName: secondaryConsultantName,
    siteName: siteName,
    siteId: siteId,
    spaces: spaces,
    takeover: takeover,
    units: units,
    usingCoreAccounting: usingCoreAccounting,
    validationMigrationConsultantName: validationMigrationConsultantName,
    companyPostMigrationSheet: companyPostMigrationSheet,
    yardiIntegratedSheet: yardiChecklistSheet,
    udsBackupRestored: udsBackupRestored,
    udsFileName: udsFileName,
    udsFileLocation: udsFileLocation,
    propertyLookUpCode: propertyLookUpCode
  };
  return workbookData;
}