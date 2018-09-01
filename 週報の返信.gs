function getWeeklyReportResponseMailText() {

  var mail_template_row = 3; // メールテンプレート内の「週報の返信メールテンプレート」の行番号
  var mail_template_col = 1; // メールテンプレート内の「週報の返信メールテンプレート」の列番号
  
  var tp    = SpreadsheetApp.getActive().getSheetByName( template );
  var range = tp.getRange(mail_template_row,mail_template_col).getValues();
  
  Logger.log(range);
              
  var response_text  = range[0][0];  // 週報の返信メールテンプレート
  
  return response_text;
  
}
