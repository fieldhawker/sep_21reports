/**
 * 週報の検知処理のトリガを設定
 *
 *
 */
function createTriggers() {

  var hours = [10] // 未来時を指定する必要がある
  
  console.log('createTriggers')
  var now = new Date()

  // ---------------------------
  // 残っているトリガを削除する
  // ---------------------------
  
  var triggers = ScriptApp.getProjectTriggers()
  
  if (Array.isArray(triggers)) {
  
    triggers.forEach(function(trigger) {
      // 他のトリガは削除しないように名前で判定
      if(trigger.getHandlerFunction() === 'checkNotSubmitter') {
        ScriptApp.deleteTrigger(trigger);
      }
    })
    
  }
  
  // ---------------------------
  // トリガを設定
  // ---------------------------

  hours.forEach(function(hour) {
  
    var date = new Date()
    date.setHours(hour)
    date.setMinutes(0)
    
    if (now.valueOf() < date.valueOf()) {
      // checkNotSubmitter() のトリガを指定した日時で作成
      ScriptApp.newTrigger("checkNotSubmitter").timeBased().at(date).create()
    }
    
  })
  
}

/**
 * 月報の検知処理のトリガを設定
 *
 *
 */
function createMonthlyTriggers() {

  var hours = [10] // 未来時を指定する必要がある
  
  console.log('createTriggers')
  var now = new Date()

  // ---------------------------
  // 残っているトリガを削除する
  // ---------------------------
  
  var triggers = ScriptApp.getProjectTriggers()
  
  if (Array.isArray(triggers)) {
  
    triggers.forEach(function(trigger) {
      // 他のトリガは削除しないように名前で判定
      if(trigger.getHandlerFunction() === 'checkNotMonthlySubmitter') {
        ScriptApp.deleteTrigger(trigger);
      }
    })
    
  }
  
  // ---------------------------
  // トリガを設定
  // ---------------------------

  hours.forEach(function(hour) {
  
    var date = new Date()
    date.setHours(hour)
    date.setMinutes(0)
    
    if (now.valueOf() < date.valueOf()) {
      // checkNotSubmitter() のトリガを指定した日時で作成
      ScriptApp.newTrigger("checkNotMonthlySubmitter").timeBased().at(date).create()
    }
    
  })
  
}