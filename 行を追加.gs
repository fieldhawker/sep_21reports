/**
 * 週を表す行を新規追加する
 * 
 */
function addLine() {

  var ss    = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var today = new Date();
  
  ss.insertRows(result_row);
  ss.getRange(result_row,1).setValue(today);
  
}
