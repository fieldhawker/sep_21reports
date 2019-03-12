

/**
 * 月報が提出されたことを記録する
 *
 *
 */
function onSaveMonthlyReportToSheet() {

  // 対象のメールラベル
  var label = "SEP/SEP二課月報";
  
  // =============================
  // IFTTT連係情報
  // =============================
  
  // 月報の提出者に返信するときに使用するIFTTTのhook名
  var response_hook_event = "submitter_result_mail";
  
  
  // =============================
  // メールタイトル
  // =============================
  
  // 月報の返信で使用するメールタイトル
  var response_mail_title = "★テスト中★[SEP][二課１G] 月報を受領しました";

  var start_rownum = 2;
  var start_colnum = 3;
  var col_size = 4;

  var emails_name_row  = name_rownum - start_rownum;  // スタッフ情報のうちの名前の行要素番号
  var emails_email_row = email_rownum - start_rownum; // スタッフ情報のうちのメルアドの行要素番号
  var emails_id_row = id_rownum - start_rownum; // スタッフ情報のうちの社員番号の行要素番号
  
  // スタッフ情報の抽出
  var ss    = SpreadsheetApp.getActive().getSheetByName( monthly_sheet_name );
  var range = ss.getRange(start_colnum,start_rownum,col_size,staff_count); // ３行目２列目から３行、スタッフ数を抽出する TODO:マジックナンバー
  
  Logger.log('-- 調査対象となるスタッフ情報 ---');
  Logger.log(range.getValues());
  var emails = range.getValues();
  
  // 対象のラベルを持つ未読メールを取得する
  var threads = getMail(label);

  // 取得した新着月報の数だけ繰り返す
  for( var i in threads ) {
  
    var thread = threads[i];
    var msgs = thread.getMessages();
    
    // スレッド内のメッセージの数だけ繰り返す
    for( var j in msgs ) {
    
      var msg = msgs[j];
      
      // スレッド内の未読メッセージのみを対象とする
      if( msg.isUnread() ) {

        // from
        var from = getAddress(msg.getFrom());
        
        // to
        var to = getAddress(msg.getTo());
        
        // body
        var body = msg.getPlainBody();
        var id = body.match(/社員番号 : (\d+)/);
        if (id == null) {
          staff_id = '';
        } else {
          staff_id = id[1];
        }
        Logger.log(staff_id);
        
        var col = 0;
                
        for(var i=0;i<emails[emails_email_row].length;i++){
        
          if (emails[emails_email_row][i] == '') {
            break;
          }
        
          Logger.log(emails[emails_email_row][i]);
          Logger.log(emails[emails_id_row][i]);
      
          if(emails[emails_email_row][i] === from || emails[emails_id_row][i] == staff_id){
          //if(emails[emails_email_row][i] === to){
          
            col = i + 1 + 1; // 列は1から開始 & 1列目は見出し列 TODO:マジックナンバー
            
            if (!/済/.test(ss.getRange(result_row,col,1,1).getValue())) {
            
              // 月報の返信メールテンプレートの取得
              
              var response_text  = getMonthlyReportResponseMailText();
    
              // はてなブックマーク情報の取得
                            
              var bookmarks = getHatenaBookmark();
              
              // IFTTTを経由して返信メールを送付
              
              var tmp   = new Object();
              tmp.mail  = emails[emails_email_row][i];
              tmp.title = response_mail_title;
              tmp.body  = response_text
                            .replace(/__NAME__/,emails[emails_name_row][i])
                            .replace(/__MAIL_ADDRESS__/,emails[emails_email_row][i])
                            .replace(/__HATENA__/,getBookmarkMessage(bookmarks).join('\r\n'));
              
              sendMailHookRequest(response_hook_event, tmp);
              
            }
            
            // 対象のメールアドレス列の今週の行に値を設定
            
            ss.getRange(result_row,col,1,1).setValue('済');
        
            break;
          }
        }
        
        if (col === 0) {
          continue;
        }
        Logger.log(col);
        
      }
    }
    // スレッドを既読にする
    thread.markRead();
    Utilities.sleep(10000);
  }
}

function getMonthlyReportResponseMailText() {

  var mail_template_row = 6; // メールテンプレート内の「月報の返信メールテンプレート」の行番号
  var mail_template_col = 1; // メールテンプレート内の「月報の返信メールテンプレート」の列番号
  
  var tp    = SpreadsheetApp.getActive().getSheetByName( template );
  var range = tp.getRange(mail_template_row,mail_template_col).getValues();
  
  Logger.log(range);
              
  var response_text  = range[0][0];  // 月報の返信メールテンプレート
  
  return response_text;
  
}
