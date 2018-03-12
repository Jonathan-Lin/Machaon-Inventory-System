/**
* This automatically transfers the info from the Form sheet to the supplies to be ordered sheet
* On opening the sheet, this runs and splits each item on form into individual orders
*
* Author: Jonathan Lin
* Date: 21JUL2017
*/

function onOpen(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get spreadsheet
  var FormSheet = ss.getSheetByName("Form") //get form to get info from
  var moveToSheet = ss.getSheetByName("Supplies To Be Ordered"); //sheet to move data to
  var data = FormSheet.getDataRange(); //get data range
  var startRow = 2;
  var endRow = FormSheet.getLastRow();
  var lastCol = FormSheet.getLastColumn();
  var numColPerItem = 5; //number of columns used for each item
  var colOfLastItem = 26; //column of beginning of last item
  var updateCol = colOfLastItem + numColPerItem; //row where update status is
  
  //check each row from last to first
  for( var row = endRow; row >= startRow; row-- ) { 
    //if empty row continue
    if( FormSheet.getRange(row, 1,1,1).getValue() == "") {
      continue; 
    }
    //if already updated then break
    if( FormSheet.getRange(row, updateCol, 1, 1).getValue() == "Updated" ) {
      break; 
    }

    var numItems = FormSheet.getRange(row, 5, 1, 1).getValue(); //get number of items to expect
    
    //loop through each item and add to sheet
    for( var item = 0; item < numItems; item++ ) {
      
      moveToSheet.insertRows(2,1); //make new row for new info
 
      var name_and_supplier = FormSheet.getRange(row, 2, 1, 2); // copy name and supplier info
      var time_frame = FormSheet.getRange(row, 4, 1, 1); //copy time frame info
      var timestamp = FormSheet.getRange(row, 1, 1, 1);
      
      name_and_supplier.moveTo(moveToSheet.getRange(2,2)); //copy to new sheet at new line
      time_frame.moveTo(moveToSheet.getRange(2,9)); //copy time frame to new sheet at new line
      timestamp.moveTo(moveToSheet.getRange(2,1));
      
      //get range that needs to be moved for each item
      var rangeToBeMoved = FormSheet.getRange(row, colOfLastItem - (numColPerItem * item), 1, numColPerItem)
      
      //insert row, move range, set to updated
      rangeToBeMoved.moveTo(moveToSheet.getRange(2, 4));
      
      FormSheet.getRange(row,updateCol,1,1).setValue("Updated"); 
      
    }
  }
}
