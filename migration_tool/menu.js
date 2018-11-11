/*Name: Mark Romero
* Date: 1/22/2016
* Updated: 12/6/2017
* Purposuse: This script file contains a menu function, it will add an Add on menu for 
* Please Migrate and Migration Complted Actions.
*==============================================================================*/


/*Purpose: Create an Add on Menu For the Please Migrate and Migration Completed email Actions
* Parameters: None
* Returns : None
*==============================================================================*/
function onOpen(e) {
  //creates add on menu to the UI
  var menu = SpreadsheetApp.getUi().createAddonMenu()
      //Adds the sub menu Please Migrate and Migration Complted to the menu.
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Pre-Migration Tools')
        .addItem('Request Pre-check issue', 'dataCorrectionSidebar'))  
        .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Data Check In Tools')
        .addItem('Add Query Doc', 'addQueryDoc')
        .addItem('Please migrate', 'pleaseMigrate'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Query Tools')
        .addItem('Send queries to Consultant', 'queriesToConsultant'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Post-Migration Tools')
        .addItem('Migration completed', 'migrationCompletedEmail'))
      .addToUi();  
}