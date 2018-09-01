/**
 * 週を表す行を新規追加する
 * 
 */
function addLineWeekly() {

  var ss    = SpreadsheetApp.getActive().getSheetByName( weekly_sheet_name );
  var today = new Date();
  
  ss.insertRows(result_row);
  ss.getRange(result_row,1).setValue(today);
  
}

/**
 * 月を表す行を新規追加する
 * 
 */
function addLineMonthly() {

  var ss    = SpreadsheetApp.getActive().getSheetByName( monthly_sheet_name );
  var today = new Date();
  
  ss.insertRows(m_result_row);
  ss.getRange(m_result_row,1).setValue(today);
  
}
