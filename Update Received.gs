/**
* updateReceived will update the received sheet to show which items
* were received when, by who and ordered when, by who
*
* Author: Jonathan Lin 
* Date: 21JUL2017
*/

function updateReceived() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activatedSheetName = ss.getActiveSheet().getName();
  var ActiveSheet = ss.getSheetByName("Supplies Ordered Pending Delivery"); // source sheet
  var MoveDatatoThisSheet = ss.getSheetByName("Supplies Received"); //    target sheet
  var startRow = 2;
  var getRange = ss.getDataRange();
  var getRow  = getRange.getRow();
  var endRow = getRange.getLastRow();
  
  for (var row = endRow; row >= startRow; row--) {  
    
    var rangeToCheck = ActiveSheet.getRange(row, 16,1,1); //  column N in row ree
    
    if(ActiveSheet.getRange(row,1,1,1).getValue()=="") {
      continue;
    }
    
    if (rangeToCheck.getValue()!="") {   // joining values before checking the expression
     
      MoveDatatoThisSheet.insertRows(2,1);
      var rangeToMove = ActiveSheet.getRange(row, 1,1, ActiveSheet.getMaxColumns());
      rangeToMove.moveTo(MoveDatatoThisSheet.getRange("A2"));
      // add date and time of when approved to target row in column E
      
      var m_names = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");

      var d = new Date();
      var curr_date = d.getDate();
      var curr_month = d.getMonth();
      var curr_year = d.getFullYear();
      MoveDatatoThisSheet.getRange("Q2").setValue(curr_date + "-" + m_names[curr_month] + "-" + curr_year);
    }
  }
}

