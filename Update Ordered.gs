/**
 * Update the Supplies to be Ordered sheet
 *
 * Move any items that are completely filled out (Ordering Issues Optional)
 * to the Supplies Ordered Pending Delivery sheet and Time Stamps with Date
 */

/**
 * Author: Jonathan Lin 
 * Date: 21JUL2017
 * Version: 1.0
 */

function updateOrdered() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activatedSheetName = ss.getActiveSheet().getName();
  var ActiveSheet = ss.getSheetByName("Supplies to be Ordered");  //   source sheet
  var MoveDatatoThisSheet = ss.getSheetByName("Supplies Ordered Pending Delivery");  //   target sheet
  var startRow = 2;
  var getRange = ss.getDataRange();
  var getRow  = getRange.getRow();
  var endRow = getRange.getLastRow();
  var numColsToMove = 12;
  
  for (var row = endRow; row >= startRow; row--) {  
    
    //cells to check are I,J,K,L,M(9,10,11,12,13)
    var totalPrice = ActiveSheet.getRange(row, 9 ,1,1); //total price
    var PO_Number = ActiveSheet.getRange(row, 10,1,1); //Purchase Order Number
    var CCCheck = ActiveSheet.getRange(row, 11 ,1,1); //credit card number
    var confNumber = ActiveSheet.getRange(row, 12 ,1,1); //confirmation number
    var initialsToCheck = ActiveSheet.getRange(row, 13 ,1,1); //intials
    
    if(ActiveSheet.getRange(row,1,1,1).getValue()=="") {
      continue;
    }
    
    //Check if any needed information is missing, if so do not procede to transfer data
    if ( !totalPrice.getValue()=="" && !PO_Number.getValue()=="" && !CCCheck.getValue()=="" && !confNumber.getValue()=="" && !initialsToCheck.getValue()=="" ) {  
      
      MoveDatatoThisSheet.insertRows(2,1); //create new row to enter info
      var orderingIssues = ActiveSheet.getRange(row, 15, 1, 1); // get ordering issue info
      orderingIssues.moveTo(MoveDatatoThisSheet.getRange(2,17,1,1)); //move ordering issue info
      
      var dateRequested = ActiveSheet.getRange(row, 1, 1, 1);
      dateRequested.moveTo(MoveDatatoThisSheet.getRange(2, 13, 1, 1));
      
      var orderedBy = ActiveSheet.getRange(row, 14,1, 1); //get range of data for item to move
      orderedBy.moveTo(MoveDatatoThisSheet.getRange(2,14,1,1)); //move to new row in ordered items sheet
      Logger.log(orderedBy.getValue());
      
      var frontRangeToMove = ActiveSheet.getRange(row, 2, 1, numColsToMove);
      frontRangeToMove.moveTo(MoveDatatoThisSheet.getRange(2,1,1,numColsToMove));
      
      ActiveSheet.deleteRow(row); //delete the row in old ordering sheet
      
      // add date and time of when approved to target row in column M
      var m_names = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");

      var d = new Date();
      var curr_date = d.getDate();
      var curr_month = d.getMonth();
      var curr_year = d.getFullYear();
      MoveDatatoThisSheet.getRange("O2").setValue(curr_date + "-" + m_names[curr_month] + "-" + curr_year);
    }
  }
}

