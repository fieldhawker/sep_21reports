/**
 * はてなブックマークをシートに保存する
 *
 *
 */
function outputHatenaBookmark() {

  var bookmarks = requestHatenaBookmarkHotEntry();
  var result = putHatenaBookmarkforSheet(bookmarks);

}

/**
 * はてなブックマーク注目記事を取得する
 *
 *
 */
function requestHatenaBookmarkHotEntry() {

  var row = 0;
  var url = "http://b.hatena.ne.jp/hotentry/it.rss?of=1";
  var ret_values = [];
  
  var array = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29];
  array.sort(function(){
    return Math.random()-0.5;
  });
  
  var xml      = UrlFetchApp.fetch(url).getContentText();
  var document = XmlService.parse(xml);
  var root     = document.getRootElement();
  var rss      = XmlService.getNamespace('http://purl.org/rss/1.0/');
  var dc       = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');
  var hatena   = XmlService.getNamespace('hatena', 'http://www.hatena.ne.jp/info/xmlns#');
  var items    = root.getChildren('item', rss);
  
  for (var i = 0; i < array.length; i++) {
  
    var num = array[i];
    
    if (items[num].getChild('bookmarkcount', hatena).getText() <= 100) {
      continue;
    };
    
    var title   = items[num].getChild('title', rss).getText();
    var link    = items[num].getChild('link', rss).getText();
    var pubdate = items[num].getChild('date', dc).getText();
    //var description = items[i].getChild('description',rss).getText();
    var cnt     = items[num].getChild('bookmarkcount', hatena).getText();
    
    var tmp        = {};
    tmp['pubdate'] = pubdate;
    tmp['title']   = title;
    tmp['link']    = link;
    tmp['cnt']     = cnt;
    
    ret_values.push(tmp);
    
    if (ret_values.length >= 3) {
      break;
    }
  
  }
  
  Logger.log('-- はてなブックマークへの問い合わせの結果 ---');
  
  for (var i = 0; i < ret_values.length; i++) {
    Logger.log(ret_values[i]['title']);
  }
  
  return ret_values;

}

/**
 * 取得したブックマークをシートに保存
 *
 *
 */
function putHatenaBookmarkforSheet(bookmarks) {
  
  var ss    = SpreadsheetApp.getActive().getSheetByName( hatena_sheet_name );
  
  for (var i = 0; i < bookmarks.length; i++) {
    var row = i + 1;
    ss.getRange(row,1,1,1).setValue(row);
    ss.getRange(row,2,1,1).setValue(bookmarks[i]['title']);
    ss.getRange(row,3,1,1).setValue(bookmarks[i]['link']);
    ss.getRange(row,4,1,1).setValue(bookmarks[i]['pubdate']);
    ss.getRange(row,5,1,1).setValue(bookmarks[i]['cnt']);
  }
  
  return true;
    
}


/**
 * はてなブックマークのリンク文言を生成
 *
 * @return はてなブックマークのリンク文言
 */
function getBookmarkMessage(bookmarks) {

  var tmp       = [];
  var title_col = 1;
  var url_col   = 2;
  
  for(var i=0;i<bookmarks.length;i++){
  
    str = bookmarks[i][title_col] + '\r\n' + bookmarks[i][url_col];
    
    tmp.push(str);
    
  }
  
  return tmp;
  
}


/**
 * はてなブックマークをスプレットシートから取得
 *
 * @return 
 */
function getHatenaBookmark() {

  var row = 1;
  var col = 1;
  var numrows = 3;
  var numcolumns = 4;

  var tp        = SpreadsheetApp.getActive().getSheetByName( hatena_sheet_name );
  var bookmarks = tp.getRange(row,col,numrows,numcolumns).getValues();
  
  return bookmarks;
  
}
