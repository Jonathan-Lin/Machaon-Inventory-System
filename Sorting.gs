/**
* These different Functions will sort by a certain column
* 3, 2, and 10. Easily changeable by changing the colToSort Variable
*
* Author: Jonathan Lin
* Date: 21JUL2017
*/

function sortOrdersBy3(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = ss.getDataRange();
  var numRows = range.getLastRow();
  var numCols = range.getLastColumn();
  var colToSort = 3; //sort by this column
  
  var sortRange = ss.getRange(2,1,numRows, numCols);
  
  sortRange.sort({column: colToSort});
}

function sortOrdersBy2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = ss.getDataRange();
  var numRows = range.getLastRow();
  var numCols = range.getLastColumn();
  var colToSort = 2; //sort by this column
  
  var sortRange = ss.getRange(2,1,numRows, numCols);
  
  sortRange.sort({column: colToSort});
}

function sortOrdersBy10() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = ss.getDataRange();
  var numRows = range.getLastRow();
  var numCols = range.getLastColumn();
  var colToSort = 10; //sort by this column
  
  var sortRange = ss.getRange(2,1,numRows, numCols);
  
  sortRange.sort({column: colToSort});
}
