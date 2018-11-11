/*Name: Mark Romero
* Date: 1/22/2016
* Updated: 2/23/2016
* Purposuse: This script file contains a menu function, that will create a Add on menu for 
* Consultant migration tool
*==============================================================================*/


/*Purpose: Create an Add on Menu For the consultant migration tool
* Parameters: None
* Returns : None
*==============================================================================*/
//Guide for creating menu: https://developers.google.com/apps-script/guides/menus
function onOpen(e) { //function runs on open.
  //creates add on menu to the UI
  var menu = SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Send to Migration', 'beginMigration')
      .addItem('Request Data Fix', 'dataFixSidebar')
      .addToUi();
}
