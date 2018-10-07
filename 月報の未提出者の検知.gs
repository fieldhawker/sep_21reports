
/**
 * 提出状況を確認し、社員にリマインドメールを送る
 * 提出状況を確認し、リーダーに現状を共有するメールを送る
 *
 *
 */
function checkNotMonthlySubmitter() {

  // =============================
  // IFTTT連係情報
  // =============================
  
  // 月報の提出結果を通知するときに使用するIFTTTのhook名
  var remind_hook_event = "not_submitters_mail";
  
  
  // =============================
  // メールタイトル
  // =============================
  
  // 社員用のメールタイトル
  var staff_mail_title = "★テスト中★[SEP][二課１G] 月報のリマインドメール";
  // リーダー用のメールタイトル
  var reader_mail_title = "★テスト中★[SEP][二課１G] 月報の提出状況の共有";;
  
  
  // =============================
  // メール本文
  // =============================

  // 未提出者が存在しなかった場合の文言
  var not_found_message = "【未提出者は検知されませんでした。】";

  // =============================


  // =============================
  // シートの1行目2列目から7行分の社員情報を取得
  // =============================

  var ss     = SpreadsheetApp.getActive().getSheetByName( monthly_sheet_name );
  var range  = ss.getRange(1,2,result_row,staff_count) // TODO:マジックナンバー
  var emails = range.getValues();
  
  Logger.log('-- 調査対象となるメールアドレス ---');
  Logger.log(emails);
  
  
  // =============================
  // 未提出の社員の情報のみをグローバル変数に格納
  // =============================
  
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
  
  Logger.log('-- 未提出の社員 ---');
  Logger.log(not_submitters);
  
    
  // =============================
  // テンプレートシートからメール本文のテンプレートを取得
  // =============================
  
  var tp    = SpreadsheetApp.getActive().getSheetByName( template );
  var range = tp.getRange(4,1,2).getValues(); // TODO:マジックナンバー
    
  var staff_text  = range[0][0];  // 社員用のメールテンプレート
  var reader_text = range[1][0];  // リーダー用のメールテンプレート
  
//  var toAdr   = "";//送り先アドレス
//  var ccAdr   = "";//Ccアドレス
//  var bccAdr  = "";//bccアドレス
//  var subject = staff_mail_title;//メールの題目
//  var name    = "";//送り主の名前
//  var files   = new Array();//添付ファイル，どんな型でも，いくつでも格納できる配列，
//  var body    = "";
  
  
  // ------------------------
  // リーダーへ未提出者の通知
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

  Utilities.sleep(5000);
  
  // =============================
  // 各自へリマインドメールの送信
  // =============================
  
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
  
//    // GASのメール送信APPを使用する場合　（24時間で100通の制限があるので却下）
//    toAdr = not_submitters[i];
//    body  = staff_text.replace(/__NAME__/,not_names[i]).replace(/__MAIL_ADDRESS__/,not_submitters[i]);
//    MailApp.sendEmail({to:toAdr, cc:ccAdr, bcc:bccAdr, subject:subject, name:name, body:body, attachments:files});

    Utilities.sleep(15000);
    
  }
  
}
