// 対象のメールラベル
var label = "SEP/SEP二課";

// IFTTTのWebHook情報
var hook_key = "bVARe0_iYKvylPAk1e609M";

// 週報の提出結果に関する設定
var remind_hook_event = "not_submitters_mail";
// 社員用のメールタイトル
var staff_mail_title = "★テスト中★[SEP][二課一係] 週報のリマインドメール";
// リーダー用のメールタイトル
var reader_mail_title = "★テスト中★[SEP][二課一係] 週報の提出状況の共有";


// 週報の返信に関する設定
var response_hook_event = "submitter_result_mail";
// 週報の返信で使用するメールタイトル
var response_mail_title = "★テスト中★[SEP][二課一係] 週報を受領しました";

// docomo development の情報
var docomo_api_key = "302e3551613133527864622e352f304a5344674452526e65437a68576766792e66776e3175516e77795736";
var docomo_app_kind = "takasep";
var docomo_app_id = "";


var sheetName         = 'reports';         // 対象のシート名
var template          = 'メールテンプレート'; // メールテンプレート用のシート名
var hatena_sheet_name = 'はてブ';           // はてブ用のシート名

// リーダー用メールアドレス
var reader_mail_to = 'matsushige@se-project.co.jp,izumi@se-project.co.jp,takano@se-project.co.jp,nakamurashinya@se-project.co.jp,machidatakahiro@se-project.co.jp,takenagakota@se-project.co.jp,naruseyuki@se-project.co.jp,takahashi@se-project.co.jp,takayama@se-project.co.jp';
// var reader_mail_to = 'matsushige@se-project.co.jp,izumi@se-project.co.jp,fieldhawker+reader@gmail.com,takahashi@se-project.co.jp,takayama@se-project.co.jp';
// var reader_mail_to = 'fieldhawker+reader@gmail.com';
// リーダー用CC カンマ区切りで複数指定可
var reader_mail_cc = 'fieldhawker+cc@gmail.com,fieldhawker+cc2@gmail.com';

// 最大社員数
var staff_count = 30;


// 未提出者が存在しなかった場合の文言
var not_found_message = "【未提出者は検知されませんでした。】";

var not_groups      = [];
var not_group_addrs = [];
var not_names       = [];
var not_submitters  = [];
var not_tel         = [];

// 行と列の指定 (R1C1) // 1 -

var result_row = 6;  // 済と入力する行番号

// 行と列の指定 (要素番号) 0 -

var group_name_rownum = 0;
var group_addr_rownum = 1;
var name_rownum   = 2;
var tel_rownum    = 3;
var email_rownum  = 4;
var result_rownum = 5;

/**
 * 指定のラベルが付いた未読のメールを取得
 * 
 * @return GmailThread[]
 * 
 */
function getMail() {

  var start = 0;
  var max = 500;
  return GmailApp.search('label:' + label + ' is:unread', start, max);
  
}

/**
 * メールアドレスに含まれる名前や<>を除いた文字列を返す
 *
 * @return String
 */
function getAddress(tmp) {

  tmp = tmp.replace(/.*</,'').replace(/>.*/,'');
  Logger.log(tmp);
  return tmp;
  
}

/**
 * 社員情報の文字列を作成
 *
 * @return 未提出者の詳細情報のリスト
 */
function getStaffString() {

  var tmp = [];
  
  for(var i=0;i<not_submitters.length;i++){
  
    str = '[' + not_groups[i] + '] ' + not_names[i] + ' ' + not_submitters[i] + ' ' + not_tel[i];
    
    tmp.push(str);
    
  }
  
  return tmp;
  
}

/**
 * 週報が提出されたことを記録する
 *
 *
 */
function onSaveReportToSheet() {

  var emails_name_row  = 0; // スタッフ情報のうちの名前の行要素番号
  var emails_email_row = 2; // スタッフ情報のうちのメルアドの行要素番号
  
  // スタッフ情報の抽出
  var ss    = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var range = ss.getRange(3,2,3,staff_count); // ３行目２列目から２行、スタッフ数を抽出する TODO:マジックナンバー
  
  Logger.log('-- 調査対象となるスタッフ情報 ---');
  Logger.log(range.getValues());
  var emails = range.getValues();
  
  // 対象のラベルを持つ未読メールを取得する
  var threads = getMail();

  // 取得した新着週報の数だけ繰り返す
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
        
        var col = 0;
                
        for(var i=0;i<emails[emails_email_row].length;i++){
        
          if (emails[emails_email_row][i] == '') {
            break;
          }
        
          Logger.log(emails[emails_email_row][i]);
      
          if(emails[emails_email_row][i] === from){
          //if(emails[emails_email_row][i] === to){
          
            col = i + 1 + 1; // 列は1から開始 & 1列目は見出し列 TODO:マジックナンバー
            
            if (!/済/.test(ss.getRange(result_row,col,1,1).getValue())) {
            
              // 週報の返信メールテンプレートの取得
              
              var response_text  = getWeeklyReportResponseMailText();
    
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

/**
 * 提出状況を確認し、社員にリマインドメールを送る
 * 提出状況を確認し、リーダーに現状を共有するメールを送る
 *
 *
 */
function checkNotSubmitter() {

  var ss     = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var range  = ss.getRange(1,2,6,staff_count) // TODO:マジックナンバー
  var emails = range.getValues();
  
  Logger.log('-- 調査対象となるメールアドレス ---');
  Logger.log(emails);
  
  for(var i=0;i<emails[email_rownum].length;i++){
       
    if (emails[email_rownum][i] == '') {
      break;
    }
          
    Logger.log(emails[email_rownum][i]);
    
    if (/済/.test(emails[result_rownum][i])) {
      continue;
    }
    
    not_groups.push(emails[group_name_rownum][i]);
    not_names.push(emails[name_rownum][i]);
    not_submitters.push(emails[email_rownum][i]);
    not_tel.push(emails[tel_rownum][i]);
    not_group_addrs.push(emails[group_addr_rownum][i]);
    
    
    
  }
  
  Logger.log('-- 未提出のメールアドレス ---');
  Logger.log(not_submitters);
    
  var tp    = SpreadsheetApp.getActive().getSheetByName( template );
  var range = tp.getRange(1,1,2).getValues(); // TODO:マジックナンバー
  
  Logger.log(range);
  
  var staff_text  = range[0][0];  // 社員用のメールテンプレート
  var reader_text = range[1][0];  // リーダー用のメールテンプレート
  
  var toAdr   = "";//送り先アドレス
  var ccAdr   = "";//Ccアドレス
  var bccAdr  = "";//bccアドレス
  var subject = staff_mail_title;//メールの題目
  var name    = "";//送り主の名前
  var files   = new Array();//添付ファイル，どんな型でも，いくつでも格納できる配列，
  var body    = "";
  
  // ------------------------
  // 各自へリマインドメールの送信
  // ------------------------
  
  for(var i=0;i<not_submitters.length;i++){
  
    // IFTTTを経由してメール送信
    var tmp   = new Object();
    tmp.mail  = not_submitters[i];
    tmp.title = staff_mail_title;
    tmp.body  = staff_text
                .replace(/__NAME__/,not_names[i])
                .replace(/__MAIL_ADDRESS__/,not_submitters[i])
                .replace(/__GROUP_ADDRESS__/,not_group_addrs[i]);
    
    sendMailHookRequest(remind_hook_event, tmp);
//    break;
  
    // GASのメール送信APPを使用する場合　（24時間で100通の制限があるので却下）
//    toAdr = not_submitters[i];
//    body  = staff_text.replace(/__NAME__/,not_names[i]).replace(/__MAIL_ADDRESS__/,not_submitters[i]);
//    MailApp.sendEmail({to:toAdr, cc:ccAdr, bcc:bccAdr, subject:subject, name:name, body:body, attachments:files});

    Utilities.sleep(1000);
    
  }
  
  // ------------------------
  // リーダーへの情報共有
  // ------------------------
  
  // IFTTTを経由してメール送信
  var tmp   = new Object();
  tmp.mail  = reader_mail_to;
  tmp.title = reader_mail_title;
  
  if (not_submitters.length <= 0) {
    tmp.body = reader_text.replace(/__TARGET__/,not_found_message);
  } else {
    tmp.body = reader_text.replace(/__TARGET__/,getStaffString().join('\r\n'));
  }
  
  sendMailHookRequest(remind_hook_event, tmp);
  
  // GASのメール送信APPを使用する場合　（24時間で100通の制限があるので却下）
//  toAdr   = reader_mail_to;
//  ccAdr   = reader_mail_cc;
//  subject = reader_mail_title;
//  
//  if (not_submitters.length <= 0) {
//    body = reader_text.replace(/__TARGET__/,not_found_message);
//  } else {
//    body = reader_text.replace(/__TARGET__/,getStaffString().join('\r\n'));
//  }
//  MailApp.sendEmail({to:toAdr, cc:ccAdr, bcc:bccAdr, subject:subject, name:name, body:body, attachments:files});

  Utilities.sleep(1000);
  
}

/**
 * IFTTTのWebHookをキックする
 *
 *
 */
function sendMailHookRequest(hook_event, tmp) {

  var url = "https://maker.ifttt.com/trigger/"+hook_event+"/with/key/"+hook_key;
  
  var headers = {
    "Accept"      : "application/json",
    "Content-type": "application/json"
  }
  
  var data = {
    "value1": tmp.mail,
    "value2": tmp.title,
    "value3": tmp.body
  }
  
  var options = {
    "method" : "post",
    "payload": JSON.stringify(data),
    "headers": headers
  };
  
  UrlFetchApp.fetch(url, options);

}


