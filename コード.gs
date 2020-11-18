// LINEのアクセストークン等の設定
var scriptProp　 　　= PropertiesService.getScriptProperties().getProperties();
var channel_token   = scriptProp.LINE_CHANNEL_TOKEN;
var url             = "https://api.line.me/v2/bot/message/reply";
var des_drive_id    = "1AC9WuOYm_DAn_Fsp76f29iIX38gRkMVg";

// GASの指定(ユーザー情報のスプレッドシートのセル位置)
var STATE_USER_NAME = 2;
var STATE_USER_STATUS = 3;
var STATE_USER_FULLNAME = 4;
var STATE_USER_FURIGANA = 5;
var STATE_USER_BIRTHDAY = 6;
var STATE_USER_POSTCODE = 7;
var STATE_USER_ADDRESS = 8;
var STATE_USER_PHONENUMBER = 9;
var STATE_USER_EMAIL =10;

// GASの指定(ドキュメント情報のスプレッドシートのセル位置)
var STATE_DOCID_GDRIVE = 2;
 
// LINEボットに投稿されたときに動作
function doPost(e) {
  var json = e.postData.contents;
  var events = JSON.parse(json).events;
  
  events.forEach(function(event){
    var pos_message = "";
    var eventType = event.type;
    if(eventType == "follow"){
      var displayName = follow(event);
      pos_message = displayName + "さん、はじめまして！友だち追加ありがとうございます。「会員登録」「会員情報」「委任状発行」「メニュー」に対応しています。途中で処理を終えたいときは「やめる」を送信してください。";
      reply(event, pos_message);
    }
    if(eventType == "unfollow") {
      unfollow(event);
    }
    if(eventType == "message") {
      var rec_message = event.message.text;
      switch(rec_message) {
        case "メニュー":
          pos_message = "現在「会員登録」「会員情報」「委任者登録」「委任状発行」のメッセージに対応しています。";
          break;
        case "クリア":
        case "やめる":
        case "中断":
          pos_message = "処理を中断しました。";
          SetUserData(event, STATE_USER_STATUS, "default");
          break;
        case "会員登録":
          pos_message = "会員登録を開始します。あなたの名前を設定してください。";
          SetUserData(event, STATE_USER_STATUS, "set-user-fullname");
          break;
        case "会員情報":
          pos_message = GetUserInfo(event);
          break;
        case "委任状発行":
          pos_message = MakeFormat(event,"ininjyo");
          break;
        default:
          var UserStatus = GetUserData(event, STATE_USER_STATUS);
          switch(UserStatus){
            case "set-user-fullname":
              SetUserData(event,STATE_USER_FULLNAME,rec_message);
              pos_message="あなたの名前は"+rec_message+"です。\nあなたのふりがなを設定してください。";
              SetUserData(event,STATE_USER_STATUS,"set-user-furigana");
              break;
            case "set-user-furigana":
              SetUserData(event,STATE_USER_FURIGANA,rec_message);
              pos_message="あなたのふりがなは"+rec_message+"です。\nあなたの誕生日を設定してください。(YYYY/MM/DD)";
              SetUserData(event,STATE_USER_STATUS,"set-user-birthday");
              break;
            case "set-user-birthday":
              SetUserData(event,STATE_USER_BIRTHDAY,rec_message);
              pos_message="あなたの誕生日は"+rec_message+"です。\nあなたの郵便番号を設定してください。";
              SetUserData(event,STATE_USER_STATUS,"set-user-postcode");
              break;
            case "set-user-postcode":
              SetUserData(event,STATE_USER_POSTCODE,rec_message);
              pos_message="あなたの郵便番号は"+rec_message+"です。\nあなたの住所を設定してください。";
              SetUserData(event,STATE_USER_STATUS,"set-user-address");
              break;
            case "set-user-address":
              SetUserData(event,STATE_USER_ADDRESS,rec_message);
              pos_message="あなたの住所は"+rec_message+"です。\nあなたの電話番号を設定してください。";
              SetUserData(event,STATE_USER_STATUS,"set-user-phonenumber");
              break;
            case "set-user-phonenumber":
              SetUserData(event,STATE_USER_PHONENUMBER,rec_message);
              pos_message="あなたの電話番号は"+rec_message+"です。\nあなたのEメールアドレスを設定してください。";
              SetUserData(event,STATE_USER_STATUS,"set-user-email");
              break;
            case "set-user-email":
              SetUserData(event,STATE_USER_EMAIL,rec_message);
              pos_message="あなたのEメールアドレスは"+rec_message+"です。\n会員登録を終了します。ありがとうございました。";
              SetUserData(event,STATE_USER_STATUS,"default");
              break;
            default:
              pos_message = "このメッセージには対応していません。";
          }
      }
      reply(event, pos_message);
    }
  });
}

//委任状発行
function MakeFormat(e, doc_id) {
  var user_fullname = GetUserData(e,STATE_USER_FULLNAME);
  var user_furigana = GetUserData(e,STATE_USER_FURIGANA);
  var user_birthday = GetUserData(e,STATE_USER_BIRTHDAY);
  var user_postcode = GetUserData(e,STATE_USER_POSTCODE);
  var user_address  = GetUserData(e,STATE_USER_ADDRESS);
  var user_phonenumber = GetUserData(e,STATE_USER_PHONENUMBER);
  var user_email = GetUserData(e,STATE_USER_EMAIL);

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('doc');
  var dat = ss.getDataRange().getValues();

  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === doc_id){//[行][列]
      var src_doc = DriveApp.getFileById(dat[i][1]);
      var des_drive = DriveApp.getFolderById(des_drive_id);
      var fileName = doc_id+Moment.moment().format("YYYYMMDD_HHmmss");
      var duplicateDocument   = src_doc.makeCopy(fileName, des_drive);
      var duplicateDocumentId = duplicateDocument.getId();
      var des_doc = DocumentApp.openById(duplicateDocumentId);

      var body = des_doc.getBody();
      
      body.replaceText("{today}", Moment.moment().format("YYYY年M月D日"));
      body.replaceText("{user-fullname}", user_fullname);
      body.replaceText("{user-furigana}", user_furigana);
      body.replaceText("{user-birthday}", user_birthday);
      body.replaceText("{user-postcode}", user_postcode);
      body.replaceText("{user-address}",  user_address);
      body.replaceText("{user-phonenumber}",  user_phonenumber);
      body.replaceText("{user-email}",  user_email);
      
      des_doc.saveAndClose();
      var des_url = 'https://docs.google.com/document/d/'+ duplicateDocumentId +'/export?';
      const opts         = {
        exportFormat: 'pdf',      // ファイル形式の指定 pdf / csv / xls / xlsx
        format:       'pdf',      // ファイル形式の指定 pdf / csv / xls / xlsx
        size:         'A4',       // 用紙サイズの指定 legal / letter / A4
        portrait:     'true',     // true → 縦向き、false → 横向き
        fitw:         'true',     // 幅を用紙に合わせるか
        docNames:     'false',    // シート名を PDF 上部に表示するか
        printtitle:   'false',    // スプレッドシート名を PDF 上部に表示するか
        pagenumbers:  'false',    // ページ番号の有無
        gridlines:    'false',    // グリッドラインの表示有無
        fzr:          'false',    // 固定行の表示有無
        range :       'A1%3AA1',  // 対象範囲「%3A」 = : (コロン)  
        
      };
      const urlExt = [];
      for(optName in opts){
        urlExt.push(optName + '=' + opts[optName]);
      }
      const options  = urlExt.join('&');
      const token    = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(des_url + options, {headers: {'Authorization': 'Bearer ' +  token}});
      const blob = response.getBlob().setName(fileName + '.pdf');
      var pdf_url = des_drive.createFile(blob).getUrl();  //　PDFを指定したフォルダに保存してURLを取得
      return pdf_url;
    }
  }
  return "書類を発行できませんでした";
}

// 会員情報取得
function GetUserData(e, pos) {
  var userId = e.source.userId;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('user');
  var dat = ss.getDataRange().getValues();

  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === userId){//[行][列]
      return ss.getRange(i+1, pos).getValue();
    }
  }
  return "Error";
}

function GetUserInfo(e) {
  var userId = e.source.userId;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('user');
  var dat = ss.getDataRange().getValues();

  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === userId){//[行][列]      
      return ss.getRange(i+1, STATE_USER_NAME).getValue() +"さんの会員情報です。\n" +
        "名前："+ss.getRange(i+1, STATE_USER_FULLNAME).getValue()+"\n" +
        "ふりがな："+ss.getRange(i+1, STATE_USER_FURIGANA).getValue()+"\n" + 
        "誕生日："+Moment.moment(ss.getRange(i+1,STATE_USER_BIRTHDAY).getValue()).format("YYYY/M/D")+"\n" +
        "郵便番号："+ss.getRange(i+1, STATE_USER_POSTCODE).getValue()+"\n" +
        "住所："+ss.getRange(i+1, STATE_USER_ADDRESS).getValue()+"\n" +
        "電話番号："+ss.getRange(i+1, STATE_USER_PHONENUMBER).getValue()+"\n" +
        "メールアドレス："+ss.getRange(i+1, STATE_USER_EMAIL).getValue();
    }
  }
  return "Error";
}

function SetUserData(e, pos,UserData) {
  var userId = e.source.userId;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('user');
  var dat = ss.getDataRange().getValues();
  
  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === userId){//[行][列]
      ss.getRange(i+1,pos).setNumberFormat("@");
      ss.getRange(i+1,pos).setValue(UserData);
    }
  }
}
 
function follow(e) {
  var userId = e.source.userId;
  var options = {"headers" : {"Authorization" : "Bearer " + channel_token}};
  var json = UrlFetchApp.fetch("https://api.line.me/v2/bot/profile/" + userId , options);
  var displayName = JSON.parse(json).displayName;
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('user');
  ss.appendRow([userId, displayName]); 
  return displayName;
}

function unfollow(e) {
  var userId = e.source.userId;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName('user');
  var dat = ss.getDataRange().getValues();
  
  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === userId){//[行][列]
      ss.deleteRow(i+1);
    }
  }
}
 
function reply(e,message) {
  var message = {
    "replyToken" : e.replyToken,
    "messages" : [{"type": "text","text" : message}]
  };
  var options = {
    "method" : "post",
    "headers" : {
      "Content-Type" : "application/json",
      "Authorization" : "Bearer " + channel_token
    },
    "payload" : JSON.stringify(message)
  };
  UrlFetchApp.fetch(url, options);
}