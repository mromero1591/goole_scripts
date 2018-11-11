/*Name: name Romero
* Date: 1/22/2016
* Updated: 2/23/2016
* Purposuse: This script file contains a menu function, that will create a Add on menu for 
* Query Management.
*==============================================================================*/

/*Purpose: Create a new query doc
* Parameters:string queryDocName, Represents the name of the query doc.
             string siteName, Represents the sites name.
* Returns : string queryDocUrl, the query docs url.
*==============================================================================*/
function createNewQueryDoc(queryDocName, workbookData) {
    //original query template url.
    var originalSheet = "add Link Here";
    
    //Create the Const variables
    var VALIDATION_TAB = "Validation"; //The name of Validation tab
    var TEMPLATE_SHEET_NAME = "Common Query"; //The name of common query tab
    var QueryfolderId = "add  folder id"; //Folder ID
    var originalQueryDocId = 'add doc id' //original query template ID

    //get the location of where the query doc will be stored
    var loctionOfFolder = DriveApp.getFolderById(QueryfolderId);
    
    //create the copy, set the sharing permissions, and store it in the locaiton folder.
    DriveApp.getFileById(originalQueryDocId).makeCopy(queryDocName, loctionOfFolder).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    //get the url or the new query doc.
    var newQueryDocUrl = DriveApp.getFilesByName(queryDocName).next().getUrl();
    var newQueryDocID = DriveApp.getFilesByName(queryDocName).next().getId();
    
    //Add additional editors to validation tab
    var protection = SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).protect();
    protection.addEditor('testemail hereeamil.com');
        
    //add consultants to doc
    DriveApp.getFileById(newQueryDocID).addEditor(convertNameToEmail(workbookData.consultantName));
    
    if(convertNameToEmail(workbookData.secondaryConsultantName) != " ") {
         DriveApp.getFileById(newQueryDocID).addEditor(convertNameToEmail(workbookData.secondaryConsultantName));
    }
    
    if(convertNameToEmail(workbookData.accountingConsultantName) != " ") {
         DriveApp.getFileById(newQueryDocID).addEditor(convertNameToEmail(workbookData.accountingConsultantName));
    }
    
    if(convertNameToEmail(workbookData.secondaryAccountingConsultantName) != " ") {
         DriveApp.getFileById(newQueryDocID).addEditor(convertNameToEmail(workbookData.secondaryAccountingConsultantName));
    }
    
    //set the Core Consultants name and accounting consultant name
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(1,2).setValue(workbookData.consultantName);
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(1,3).setValue(workbookData.secondaryConsultantName);
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(1,4).setValue(workbookData.accountConsultantName);
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(1,5).setValue(workbookData.secondaryAccountingConsultantName);
    
    //set property data.
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(7, 1).setValue(workbookData.siteName);
    SpreadsheetApp.openByUrl(newQueryDocUrl).getSheetByName(VALIDATION_TAB).getRange(7, 2).setValue(workbookData.units);
    
    //Formating the new query doc.............................................................................................
    //set the new querydoc as the active sheet.
    var newQueryId = SpreadsheetApp.openByUrl(newQueryDocUrl).getId(); //get the id
    var ss = SpreadsheetApp.openById(newQueryId); //open by id
    SpreadsheetApp.setActiveSpreadsheet(ss); //set as active sheet
    var templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME); //make the common query sheet the template sheet
    ss.insertSheet(workbookData.siteName, 1, {template:templateSheet}); //enter the new sheet using the name passed in.
    ss.deleteSheet(templateSheet); //delete the common query sheet.
    
    return newQueryDocUrl;
}

/*Purpose: Add new site to queryDoc
* Parameters:string queryDocName, Represents the name of the query doc.
             string siteName, Represents the sites name.
             string queryDocUrl, Represents the url of the query doc
* Returns : None
*==============================================================================*/
function addNewSheetToDoc(queryDocName, workbookData, queryDocURL) {
    //open the query doc
    var queryDoc = SpreadsheetApp.openByUrl(queryDocURL);
    SpreadsheetApp.setActiveSpreadsheet(queryDoc);
    
    var sheets = queryDoc.getSheets();
    var templateSheet = sheets[0];
    if(sheets.length == 2) {
        queryDoc.insertSheet("Common Query", 0, {template:templateSheet}); //enter the new sheet using the name passed in.
    }
    queryDoc.insertSheet(workbookData.siteName, 1, {template:templateSheet}); //enter the new sheet using the name passed in.
    
    //set up property info
    var VALIDATION_TAB = "Validation"; //The name of Validation tab
    var validationSheet = queryDoc.getSheetByName(VALIDATION_TAB)
    
    validationSheet.getRange(validationSheet.getLastRow() + 1, 1).setValue(workbookData.siteName);
    validationSheet.getRange(validationSheet.getLastRow(), 2).setValue(workbookData.units);
}

/*Purpose: will check if a query doc exist if so then a new sheet is added if not
           then a new doc will be created.
* Parameters:None
* Returns : None
*==============================================================================*/
function addQueryDoc() {
    var workbookLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var workbookData = getWorkbookData(workbook);
    
    var QueryfolderId = "folder id here"; //Folder ID
    var locationOfFolder = DriveApp.getFolderById(QueryfolderId);
    var filesInFolder = locationOfFolder.getFiles();
    
    //create the name of the new query doc using the company name.
    var queryDocName = workbookData.companyName + "Querires for Migration of" + workbookData.migrationDate;
    
    //state that the query doc does not exist
    var queryDocExist = false;
    var queryDocURL = "";
    
    //check all the files in the query folder if the doc exist set the queryDocExist to true.
    while(filesInFolder.hasNext()) {
        var file = filesInFolder.next();
        if(file.getName() == queryDocName) {
            queryDocURL = file.getUrl();
            queryDocExist = true; 
        }
    }
    
    //If the Query doc exist add a new sheet to it otherwise create a new doc.
    if(queryDocExist) {
        addNewSheetToDoc(queryDocName, workbookData, queryDocURL);
    }
    else {
        queryDocURL = createNewQueryDoc(queryDocName, workbookData);
    }
    //add the link to the workbook.
    workbookData.preStartSheet.getRange(workbookData.queryDocLinkCell).setValue(queryDocURL);
}

function updateTrackerSheet(siteName, units, queryDocUrl, companyName) 
{
    var queryDoc = SpreadsheetApp.openByUrl(queryDocUrl);
    var sheets = queryDoc.getSheets();
    
    //create variable that will hold the count of each query status.
    //assuming that the person who completes the migration is not the same who answered queries
    //we will create a counter for each migration agent

    for(i = 0; i < sheets.length; ++i)
    {
        if(sheets[i].getName() == "Common Query"){
            processQueryTracker(sheets[i], "Common Query", companyName, 0);
        }
        
        else if(sheets[i].getName() == siteName){
            processQueryTracker(sheets[i], siteName, companyName, units);
        }
    }       
}

function processQueryTracker(sheet, siteName, companyName, units)
{
    var statusAnswerColumnName = "Status Answer"; //the name used to search for status answer Column.
    
    //names
    var nameQueries = [0,0,0,0,0,"name", false];

    //names
    var nameQueries = [0,0,0,0,0, "name", false];
    
    //name
    var nameQueries = [0,0,0,0,0, "name", false];
    
    //name
    var nameQueries = [0,0,0,0,0, "name", false];
    
    //name
    var nameQueries = [0,0,0,0,0, "name", false];
    
    //Barndon
    var nameQueries = [0,0,0,0,0, "name", false];
  
    //name
    var nameQueries = [0,0,0,0,0, "name", false];
    
    //name
    var nameQueries = [0,0,0,0,0, "name", false];
    
    var masterQueries = [nameQueries, nameQueries, nameQueries, nameQueries, nameQueries, nameQueries, nameQueries, nameQueries];
    if(sheet.getRange(1, 26).getValue() == "Status Answer")
    {
        var statusAnswerColumn = getColumnByName(sheet, statusAnswerColumnName);   //store the location of the answers column.
        var values = sheet.getRange(1, statusAnswerColumn, sheet.getLastRow()).getValues();  //get the answer values.
        sheet.getRange(1, statusAnswerColumn).setValue("Tracked");
        
        for(var n = 0; n < values.length; ++n)
        {
          //Step 1.Count
          masterQueries = countResonses(masterQueries, values[n][0]);    
        } //end of moving through the array.
        
        //Step 2.STORE RESAULTS IN TRACKER
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
        if(nameQueries[6]){
          storeAnswers(siteName, companyName, units, nameQueries);
        }
    }
    
}

function storeQueries() {
    var queryDoc = SpreadsheetApp.getActiveSpreadsheet();
    var statusColumnName = "Status";
    var infoColumnName = "Status Answer";
    var migrationConsultant = getCompletedConsultant(GmailApp.getAliases()[0]);
    var sheet = queryDoc.getSheets();
    
    for(var i = 0; i < sheet.length - 1; ++i)
    {
      var currentSheet = sheet[i];
      var statusColumn = getColumnByName(currentSheet, statusColumnName);
      var infoColumn = getColumnByName(currentSheet, infoColumnName);
      var lastCell = currentSheet.getRange(currentSheet.getLastRow(), statusColumn);
      var start = 2;
      var currentCell = currentSheet.getRange(start, statusColumn);
      while(start <= currentSheet.getLastRow())
      {
        var status = currentCell.getValue();
  
        switch(status)
        {
          case "Completed":
          status = migrationConsultant + " - Completed";
          break;
          
          case "Need Help":
          status = migrationConsultant + " - Need Help";
          break;
          
          case "In Progress":
          status = migrationConsultant + " - In Progress";
          break;
          
          case "Sent to Client":
          status = migrationConsultant + " - Sent to Client";
          break;
          
          case "Please Review":
          status = migrationConsultant + " - Please Review";
          break;
          
          default:
          status = " ";
        }
        
        if(status != " " && currentSheet.getRange(currentCell.getRow(), infoColumn).getValue() == "")
        {
          currentSheet.getRange(currentCell.getRow(), infoColumn).setValue(status);
        }
        start += 1;
        currentCell = currentSheet.getRange(start, statusColumn);
  
      }
    }
  }

function storeAnswers(siteName, companyName, units, queryArray)
{
     var trackerDocUrl = "linkhere";
     var trackerSheetName = "Tracker";
     var trackerSheet = SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName);  //get tracker sheet 
     var trackerSheetLastRow = trackerSheet.getLastRow(); // get last row
     
     //find the coulmns in the tracker sheet that need to be updated.
     var companyNameColumn = getColumnByName(trackerSheet, "Company");
     var siteNameColumn = getColumnByName(trackerSheet,"Site");
     var dateColumn = getColumnByName(trackerSheet, "Date");
     var consultantNameColumn = getColumnByName(trackerSheet, "Consultant");
     var unitCountColumn = getColumnByName(trackerSheet, "Unit Count");
     var completedColumn = getColumnByName(trackerSheet, "Completed");
     var needHelpColumn = getColumnByName(trackerSheet, "Need Help");
     var inProgressColumn = getColumnByName(trackerSheet, "In Progress");
     var pleaseReviewColumn = getColumnByName(trackerSheet, "Please Review");
     var sentToClientColumn = getColumnByName(trackerSheet, "Sent To CIlent");
     var consultantName = queryArray[5];

     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, companyNameColumn).setValue(companyName);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, siteNameColumn).setValue(siteName);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, dateColumn).setValue(Utilities.formatDate(new Date(), " / ", "MM/dd/yyyy"));
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, consultantNameColumn).setValue(consultantName);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, unitCountColumn).setValue(units);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, completedColumn).setValue(queryArray[0]);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, inProgressColumn).setValue(queryArray[1]);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, needHelpColumn).setValue(queryArray[2]);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, sentToClientColumn).setValue(queryArray[3]);
     SpreadsheetApp.openByUrl(trackerDocUrl).getSheetByName(trackerSheetName).getRange(trackerSheetLastRow + 1, pleaseReviewColumn).setValue(queryArray[4]);
}

function countResonses(queryArray, value)
{
    if(value == "name - Completed")
    {
        queryArray[0][0] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[0][1] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[0][2] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[0][3] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[0][4] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[1][0] += 1;
        queryArray[1][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[1][1] += 1;
        queryArray[1][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[1][2] += 1;
        queryArray[1][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[1][3] += 1;
        queryArray[1][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[1][4] += 1;
        queryArray[1][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[2][0] += 1;
        queryArray[2][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[2][1] += 1;
        queryArray[2][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[2][2] += 1;
        queryArray[2][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[2][3] += 1;
        queryArray[2][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[2][4] += 1;
        queryArray[2][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[3][0] += 1;
        queryArray[3][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[3][1] += 1;
        queryArray[3][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[3][2] += 1;
        queryArray[3][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[3][3] += 1;
        queryArray[3][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[3][4] += 1;
        queryArray[3][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[4][0] += 1;
        queryArray[4][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[4][1] += 1;
        queryArray[4][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[4][2] += 1;
        queryArray[6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[4][3] += 1;
        queryArray[4][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[4][4] += 1;
        queryArray[4][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[0][0] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[0][1] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[0][2] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[0][3] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[0][4] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[0][0] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[0][1] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[0][2] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[0][3] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[0][4] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Completed")
    {
        queryArray[0][0] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - In Progress")
    {
        queryArray[0][1] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Need Help")
    {
        queryArray[0][2] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Sent to Client")
    {
        queryArray[0][3] += 1;
        queryArray[0][6] = true;
    }
    else if(value == "name - Please Review")
    {
        queryArray[0][4] += 1;
        queryArray[0][6] = true;
    }
    
    return queryArray;
}

function queriesToConsultant() {
    var subject = " ";
    var message = " ";
    var coreConsultantEmail = " ";
    var accountingConsultantEmail = " ";
    var secondaryCoreConsultantEmail = " ";
    var secondaryAccountingConsultantEmail = " ";
    var AdEmail = " ";
    var entrataEmail = " ";
    var validationTabName = "Validation";
    var emailSig = " ";
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet(); 
    var validationTab = sheet.getSheetByName(validationTabName);
      
    var coreConsultantName = validationTab.getRange(1, 2).getValue();
    var secondaryCoreConsultantName = validationTab.getRange(1, 3).getValue();
    var accountingConsultantName = validationTab.getRange(1, 4).getValue();
    var secondaryAccountingConsultantName = validationTab.getRange(1, 5).getValue();
    var adName = getAdName(coreConsultantName);
    
    //get the email addresses
    coreConsultantEmail = getEmailAddress(coreConsultantName);
    secondaryCoreConsultantEmail = getEmailAddress(secondaryCoreConsultantName);
    accountingConsultantEmail = getAccountingConsultantEmail(accountingConsultantName);
    secondaryAccountingConsultantName = getAccountingConsultantEmail(accountingConsultantName);
    AdEmail = getEmailAddress(adName);
    entrataEmail = Session.getActiveUser().getEmail();
    
    var ccEmail = 'testemail hereemail.com, ' + accountingConsultantEmail + ", " + secondaryCoreConsultantEmail + ", " + secondaryAccountingConsultantName + ", " + AdEmail;
  
    var subject = sheet.getName();
    
    var emailBody = "Hey " + coreConsultantName + "," + "<br />"
                     + "<br />" + "There are a few questions we had during the migration, I have attached a google document with those questions. The questions nameed as <b>Need Help</b> " + "<br />"
                     +  "are the ones we need your help with. Can you please look at them and answer on the document. If you have any questions please let me know." + "<br />"
                     + "<br />" + "Google link: " + sheet.getUrl() + "<br />"
                     + "<br />";
  
     var signature = getEmailSig(entrataEmail);
     
     var emailMessage = emailBody + signature;
                     
     //Check to see if an aliases exist and send it with that.
     if(entrataEmail != null){
         //do not change this, any customization can be done with the variables
         GmailApp.sendEmail(coreConsultantEmail, subject, "",{cc: ccEmail, htmlBody:emailMessage, from:entrataEmail});
      }
    
     //if no elias then send with reg email address.
     else{                        
        //do not change this, any customization can be done with the variables.
        MailApp.sendEmail({to: coreConsultantEmail,
                           cc: ccEmail,
                           subject: subject,
                           htmlBody: emailMessage});
  
     }
     
     //storeQueries();
  }
  
  
        
