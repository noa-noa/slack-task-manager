//　メンション用のユーザーIDと担当者名の対応を宣言
// ここはプッシュしない

function doPost(e) {
  var verifyToken = PropertiesService.getScriptProperties().getProperty('POST_VERIFY_TOKEN');
  if (verifyToken != e.parameter.token) {
    throw new Error("invalid token.");
  }
  
  // コマンドをパース
  var args = e.parameter.text.match(/((?!(：|\s|$)).)+/g);
  
  var result = handleSpreadSheet(args);
  postResultToSlack(e.parameter.channel_id, result);
  
  return null;
}

function test() {
  Logger.log(members["誰か"]);
  //var result = Logger.log(handleSpreadSheet(["登録", "テスト", "カテゴリ", "どっか","期限","２０１９/3/1","担当","誰か","優先度","1","進捗","1"]));
//  var result = Logger.log(handleSpreadSheet(["詳細","５８"]));
//  Logger.log(handleSpreadSheet(["担当", "誰か"]));
//  Logger.log(handleSpreadSheet(["使い方"]));
//  Logger.log(handleSpreadSheet(["削除", "24"]));
//  Logger.log(handleSpreadSheet(["編集", "2","内容","テスト", "カテゴリ", "どっか","期限","２０１９/3/1","担当","誰か","優先度","1","進捗","1"]));
  //Logger.log(handleSpreadSheet(["編集", "2","進捗","完了"]));
}

/**
 * SpreadSheetの処理 
 */
function handleSpreadSheet(args){
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  
  if (args[0] == "使い方") {
    var message = '*登録する時*\n`登録：内容 カテゴリ：カテゴリ名 優先度:1~3(１が一番優先度が高い) 期限：日付(西暦からでも良い)　担当：担当者名 進捗:特に規定なし(完了時のみ「完了」)`\n※優先度，担当，期限，進捗はオプション\n例:`登録　タスク管理する カテゴリ　襷　期限　2019/03/01　担当　誰か 優先度 2　進捗 仕様検討`または`登録　タスク管理する　カテゴリ　襷`など\n\n';
   　　message += '*担当しているタスクのIDを調べる時*\n `担当　名前`\n\n'
    message += '*タスクの詳細を調べる時*\n `詳細　管理ID`\n\n'
    message += '*編集する時*\n `編集：id(管理番号，登録した時に割り振られます)`の後に編集したい項目を指定します.\n例:`編集　2 カテゴリ　襷　 期限　2019/03/01　担当　誰か`または`編集 2 内容　タスク管理する　カテゴリ　襷　進捗 完了`など\n\n'
    message += '*削除する時*\n `削除：id(管理番号，登録した時に割り振られます)`\n例:`削除 10`など\n'
    message += '区切は半角スペース，全角スペース';
    return message;
  }
  
  try {
    var colNum = 20;
    var ss = SpreadsheetApp.openById(sheetId); 
    var sheet = ss.getSheetByName(PropertiesService.getScriptProperties().getProperty('SHEET_NAME'));
    
    // デバッグメッセージ
    var message = "";
    // 引数の偶数番目をkey, 奇数番目をvalueとするオブジェクトを作成
    var params = (function () {
      var ps = {};
      args.filter(function(_, idx) { return idx % 2 == 0}).forEach(function(opt, idx) {
        ps[opt] = args[2 * idx + 1];
      })
      return ps;
    })();
    
    var header = sheet.getSheetValues(1, 1, 1, colNum);
    // ヘッダ行の列名から列のインデックスを取得する関数 (Array版)
    var ci = function(colName) {
      var col = header[0].indexOf(colName);
      if (col < 0) throw new Error("Column '"+colName+"' does not exist (within "+colNum+" columns)");
      return col;
    };
    // ヘッダ行の列名から列の番号を取得する関数 (Range版)
    var c = function(colName) {
      return ci(colName) + 1;
    };
    var str_normalize = function(str){
      return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {return String.fromCharCode(s.charCodeAt(0) - 65248);});
    }
    
    // まず最終行にカテゴリと内容を登録する．その後担当者名があれば登録，
    if ("登録" in params) {
      var lastrowId = sheet.getLastRow();    
      var lastrow = sheet.getRange(lastrowId+1,1,1,colNum);
      var ids = sheet.getRange(2, c("id"),lastrowId,c("id"));
      var id = Math.max.apply(null,ids.getValues())+1;
      lastrow.getCell(1,c("id")).setValue(id);
      lastrow.getCell(1,c("内容")).setValue(args[1]);
      lastrow.getCell(1,c("カテゴリ")).setValue(params["カテゴリ"]);
      message += "[管理ID："+id+"] " + params["カテゴリ"]+ "に関するタスクを登録しました:jack_o_lantern:";
      
      if ("担当" in params) {
        member = params["担当"];
        lastrow.getCell(1, c("担当")).setValue(member);
        message += "\n担当 ： " + member;
        if (member in members) {
          message += "("+members[member]+")";
        }
      }
      if ("優先度" in params) {
        var priority = str_normalize(params["優先度"])
        lastrow.getCell(1, c("優先度\n(1が最優先")).setValue(priority);
        message += "\n優先度 ： " + priority;
      }
      if ("進捗" in params) {
        var progress = str_normalize(params["進捗"])
        lastrow.getCell(1, c("進捗")).setValue(progress);
        message += "\n進捗 ： " + progress;
      }
      if ("期限" in params) {
        var limit = str_normalize(params["期限"])
        lastrow.getCell(1, c("期限")).setValue(limit);
        message += "\n期限 ： " + formatDate(new Date(limit));
      }
      return message;
    }
    
    if ("編集" in params) {
      var targetid = str_normalize(args[1]);
      var rownum = findRow(sheet,parseInt(targetid),c("id"));
      var range = sheet.getRange(rownum,1,1,colNum);
      message += "\n[ID："+targetid+"] " + range.getCell(1, c("カテゴリ")).getValue()+ "に関するタスクを編集しました:jack_o_lantern:";
      if ("内容" in params) {
        range.getCell(1, c("内容")).setValue(params["内容"]);
        message += "\内容 ： " + params["内容"];
      }
      if ("カテゴリ" in params) {
        range.getCell(1, c("カテゴリ")).setValue(params["カテゴリ"]);
        message += "\カテゴリ ： " + params["カテゴリ"];
      }
      if ("担当" in params) {
        var member = params["担当"];
        range.getCell(1, c("担当")).setValue(member);
        message += "\n担当 ： " + member;
        if (member in members) {
          message += "("+members[member]+")";
        }
      }
      if ("優先度" in params) {
        var priority = str_normalize(params["優先度"])
        range.getCell(1, c("優先度\n(1が最優先")).setValue(priority);
        message += "\n優先度 ： " + priority;
      }
      if ("進捗" in params) {
        var progress = str_normalize(params["進捗"])
        range.getCell(1, c("進捗")).setValue(progress);
        message += "\n進捗 ： " + progress;
      }
      if ("期限" in params) {
        var limit = str_normalize(params["期限"])
        range.getCell(1, c("期限")).setValue(limit);
        message += "\n期限 ： " + formatDate(new Date(limit));
      }
      return message;
    }
    if ("削除" in params) {
      var targetid = str_normalize(args[1]);
      // idが一致するrow を取得
      var rownum = findRow(sheet,parseInt(targetid),c("id"));
      var range = sheet.getRange(rownum,1,1,colNum);
      message += "下記のタスクを削除しました\n"
      message += "\n管理ID : "+targetid;
      message += "\n担当 : "+range.getCell(1,c("担当")).getValue();
      message += "\n期限 : "+ formatDate(range.getCell(1,c("期限")).getValue());
      message += "\nカテゴリ : "+range.getCell(1,c("カテゴリ")).getValue();
      message += "\n進捗 : "+range.getCell(1,c("進捗")).getValue();
      message += "\n内容 : "+range.getCell(1,c("内容")).getValue();
      sheet.deleteRow(rownum);
      return message;
    }
    if ("担当" in params) {
      var employee = args[1];
      var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
      var ids = [];
      var targetCol = c("担当");
      var idCol = c("id");
      var progressCol = c("進捗");
      for(var i　=　1;　i　<　dat.length;　i++){
        if(dat[i][targetCol-1] != "" && dat[i][targetCol-1].match(employee)){
          if (dat[i][progressCol-1] == "完了") {
            continue;
          }
          ids.push(dat[i][idCol-1]);
        }
      }
      message += "\n"+employee+"さんの担当タスクID一覧　:　\n"+ids;
      return message;
    }
    
    if ("詳細" in params) {
      var taskid = str_normalize(args[1]);
      // idが一致するrow を取得
      var rownum = findRow(sheet,parseInt(taskid),c("id"));
      var range = sheet.getRange(rownum,1,1,colNum);
      message += "\n管理ID : "+taskid;
      message += "\n担当 : "+range.getCell(1,c("担当")).getValue();
      message += "\n期限 : "+ formatDate(range.getCell(1,c("期限")).getValue());
      message += "\nカテゴリ : "+range.getCell(1,c("カテゴリ")).getValue();
      message += "\n進捗 : "+range.getCell(1,c("進捗")).getValue();
      message += "\n内容 : "+range.getCell(1,c("内容")).getValue();
      return message;
    }
  } catch(e) {
    console.error(e);
    return "\n処理に失敗しました。パラメータを確認してください。 :dizzy_face:";
  }
}
function findRowId(sheet,val,col,idCol) {
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      return dat[i][idCol-1];
    }
  }
  return 0; 
}
function findRowIds(sheet,val,col,idCol){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  var index = [];
  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      index.push(dat[i][idCol-1]);
    }
  }
  return index;
}
function findRow(sheet,val,col){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      return i+1;
    }
  }
  return 0;
}
/**
 * Slackに処理結果を投稿する
 */
function postResultToSlack(channelId, message) {
  var slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
  var slackApp = SlackApp.create(slackToken);
  slackApp.chatPostMessage(channelId, message, {
    username : PropertiesService.getScriptProperties().getProperty('SLACK_BOT_NAME'),
    icon_emoji : ":book:" 
  });
}

/**
 * 当日と前日分(日付を超えての勤務を想定)のデータを取得する
 * 1:日付 2:曜日 3:出社予定 4:出社時刻 5:退社時刻 6:勤務時間 7:休憩時間 8:実働時間 9:勤務地
 */
function getRangeOfToday(sheet, numColumns) {
  var startRow = 2;
  var start = sheet.getRange(startRow, 1).getValue();
  var today = new Date();
  var diffMS = today - start;
  var diffDay = diffMS / 86400000; //1日は86400000ミリ秒
  var top = startRow + diffDay - 1;
  return sheet.getRange(top, 1, 2, numColumns);
}

/**
 * 時刻から時:分のStringを取得する
 */
function getHourMinutes(date, offsetHours) {
  return (date.getHours() + offsetHours) + ":" + ('0' + date.getMinutes()).slice(-2);
}

function alert_limited_task() {
  // もしvalue(入力されてる日付)がdate(今日の日付)と違っていたら
  username = "タスク管理botさん"
  
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
  var slackApp = SlackApp.create(slackToken);
      var colNum = 20;
    var ss = SpreadsheetApp.openById(sheetId); 
    var sheet = ss.getSheetByName(PropertiesService.getScriptProperties().getProperty('SHEET_NAME'));
    
    // デバッグメッセージ
    var header = sheet.getSheetValues(1, 1, 1, colNum);
    // ヘッダ行の列名から列のインデックスを取得する関数 (Array版)
    var ci = function(colName) {
      var col = header[0].indexOf(colName);
      if (col < 0) throw new Error("Column '"+colName+"' does not exist (within "+colNum+" columns)");
      return col;
    };
    // ヘッダ行の列名から列の番号を取得する関数 (Range版)
    var c = function(colName) {
      return ci(colName) + 1;
    };
  var limitValues = sheet.getRange(1, c("期限"), sheet.getLastRow(), 1).getValues();
  var progressValues = sheet.getRange(1, c("進捗"), sheet.getLastRow(), 1).getValues();
  var index = [];
  var today = new Date();
  // アラート対象のタスクidを取得
  for(var i=1;i<limitValues.length;i++){
    if(limitValues[i] != "" && (new Date(limitValues[i])-today)/1000 / 60 / 60 / 24 < 2){
      if (progressValues[i] == "完了") {
        continue;
      }
      index.push(i+1);      
    }
  }
  var message = "以下のタスクは覚えてますか?";
  //
  if (index.length == 0) {
    message = ":earth_asia:＜直近が期限のタスクないよ"
  }
   var str_normalize = function(str){
      return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {return String.fromCharCode(s.charCodeAt(0) - 65248);});
    }
  for (i = 0; i < index.length; i++){
    var row = index[i];
    var range = sheet.getRange(row,1,1,colNum);
    Logger.log(range.getCell(1,c("優先度\n(1が最優先")).getValue())
    var member = range.getCell(1,c("担当")).getValue();
    message += "\n管理ID : "+range.getCell(1,c("id")).getValue();
    if (member != "") {
      message += "\n担当 : "+member;
      if (member in members) {
        message += "（"+members[member]+"）"
      }
    }
    message += "\n優先度 : "+range.getCell(1,c("優先度\n(1が最優先")).getValue();
    message += "\n期限 : "+formatDate(new Date(range.getCell(1,c("期限")).getValue()));
    message += "\nカテゴリ : "+range.getCell(1,c("カテゴリ")).getValue();
    message += "\n進捗 : "+range.getCell(1,c("進捗")).getValue();
    message += "\n内容 : "+range.getCell(1,c("内容")).getValue();
    message += "\n------------------\n";
  }
  slackApp.chatPostMessage("test", message, {
    username : username,
    icon_emoji : ":rabbit:" 
  });
}
function formatDate(date) {
  if (date == "") {
    return date;
  }
  var y = date.getFullYear()
  var m = date.getMonth() + 1
  var d = date.getDate();
  var day = '日月火水木金土'.charAt(date.getDay());
  return y+"年 "+m+"月 "+d+"日 ("+day+")";
}