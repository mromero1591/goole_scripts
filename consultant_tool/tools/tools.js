/*Name: Mark Romero
* Date: 1/22/2016
* Updated:2/23/2016
* Purposuse: This file contains functions that are used to search the worksheets
* for cells or values.
*==============================================================================*/

/*Purpose: To find the column location by its header name.
* Parameters: The sheet name as a String.
*             The column Name as a string.
* Returns : The columns location as an int.
*==============================================================================*/
function getColumnByName(sheet, columnName) {
    //get the first cell of all columns in the sheet.
    var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
    
    //get all the values from the range.
    var values = range.getValues();
      
    //loope through those cells and if the value in the cell is equal to the column name passed in return the column number as an int.
    for(var row in values) {
       for(var col in values[row]) {
          if(values[row][col] == columnName) {
             return parseInt(col) + 1;
           }
       }
     }
}

/*Purpose: Search a sheet for a cerain value and returns the postion of the cell next to that value.
* Parameters: A string variable that represents the value you are looking for.
*             A string variable that represents the sheet that the value will be in. 
* Returns : Location of the cell next to the value passed in as an array of int.
*==============================================================================*/
function searchSheet(searchString, sheetName, columnLocation, arrayLocation) {
   //create a variable for the sheet you are searching.
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
   //create an array that includes all cells in the sheet.
   var values = sheet.getDataRange().getValues();
   
   //get the length of the array.
   var ilen = values.length;
   
   //move throught the array as long as you are withing the array size.
   for (var i = 0; i < ilen; i++) {
       //test to see if the value is equal to the value you are looking for if so return the cell next to it.
       if(values[i][columnLocation] == searchString) {
            if(values[i][(columnLocation + arrayLocation)] == undefined) {
              return "";
            }
            
            return values[i][(columnLocation + arrayLocation)];
       }
   }
}

function setData(sheetURL, sheetName, lastRow, column, value) {
  SpreadsheetApp.openByUrl(sheetURL).getSheetByName(sheetName).getRange(lastRow + 1, column).setValue(value);
}
