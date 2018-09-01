
// =============================
// 送信先設定
// =============================

// リーダー用メールアドレス
var reader_mail_to = 'matsushige@se-project.co.jp,izumi@se-project.co.jp,takano@se-project.co.jp,nakamurashinya@se-project.co.jp,machidatakahiro@se-project.co.jp,takenagakota@se-project.co.jp,naruseyuki@se-project.co.jp,takahashi@se-project.co.jp,takayama@se-project.co.jp';
//var reader_mail_to = 'fieldhawker+reader@gmail.com';


// =============================
// シート内の行番号の意味 ( weekly_reports )
// =============================

// 行と列の指定 (R1C1) // 1 -

var result_row = 6;  // 済と入力する行番号

// 行と列の指定 (要素番号) 0 -

var group_name_rownum = 0;
var group_addr_rownum = 1;
var name_rownum   = 2;
var tel_rownum    = 3;
var email_rownum  = 4;
var result_rownum = 5;  // 済と入力する行番号


// =============================
// シート内の行番号の意味 ( monthly_reports )
// =============================

// 行と列の指定 (R1C1) // 1 -

var m_result_row = result_row;  // 済と入力する行番号

// 行と列の指定 (要素番号) 0 -

var m_group_name_rownum = group_name_rownum;
var m_group_addr_rownum = group_addr_rownum;
var m_name_rownum   = name_rownum;
var m_tel_rownum    = tel_rownum;
var m_email_rownum  = email_rownum;
var m_result_rownum = result_rownum;  // 済と入力する行番号


// =============================
// シート名
// =============================

var weekly_sheet_name = 'weekly_reports';         // 対象のシート名
var template          = 'メールテンプレート'; // メールテンプレート用のシート名
var hatena_sheet_name = 'はてブ';           // はてブ用のシート名
var monthly_sheet_name = 'monthly_reports';         // 


// =============================
// 未提出者情報の格納用変数
// =============================

var not_groups      = [];
var not_group_addrs = [];
var not_names       = [];
var not_submitters  = [];
var not_tel         = [];


// =============================
// IFTTT連係情報
// =============================

// IFTTTのWebHook情報
var hook_key = "bVARe0_iYKvylPAk1e609M";


// =============================
// DocomoAI連係情報
// =============================

// docomo development の情報
var docomo_api_key = "302e3551613133527864622e352f304a5344674452526e65437a68576766792e66776e3175516e77795736";
var docomo_app_kind = "takasep";
var docomo_app_id = "";


// =============================
// ループ上限
// =============================

// 最大社員数
var staff_count = 30;



// =============================
// 共通処理
// =============================

/**
 * 指定のラベルが付いた未読のメールを取得
 * 
 * @return GmailThread[]
 * 
 */
function getMail(label) {

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
 * IFTTT経由でメールを送る
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
