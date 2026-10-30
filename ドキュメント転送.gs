function exportToTxtFile() {
  // 1. スプレッドシートの情報を取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 
  var data = sheet.getDataRange().getValues();
  
  var textBuffer = [];
  
  // 2行目からループ開始
  for (var i = 1; i < data.length; i++) { 
    var id = data[i][0];         // A列: 管理番号
    var questioner = data[i][1]; // B列: 質問者
    var question = data[i][2];   // C列: 質問内容
    var answer = data[i][3];     // D列: 回答
    var answerer = data[i][4];   // E列: 回答者
    var refUrl = data[i][5];     // F列: 参考URL
    
    if (question && answer) {
      var block = "【管理番号】 " + id + "\n" +
                  "【質問者】 " + questioner + "\n" +
                  "【質問内容】\n" + question + "\n" +
                  "【回答者】 " + answerer + "\n" +
                  "【回答】\n" + answer + "\n" +
                  "【参考URL】 " + (refUrl ? refUrl : "なし") + "\n" +
                  "================================\n";
      textBuffer.push(block);
    }
  }
  
  var fullText = textBuffer.join("\n");
  var fileName = "CSナレッジベース_検索用.txt";
  
  // マイドライブから同名のファイルを探す
  var files = DriveApp.getFilesByName(fileName);
  var file;
  
  if (files.hasNext()) {
    // 既存のファイルがあれば中身を丸ごと最新に上書き
    file = files.next();
    file.setContent(fullText);
    Logger.log("既存のテキストファイルに " + textBuffer.length + " 件のデータを上書き保存しました。");
  } else {
    // なければ新規にテキストファイルを作成
    file = DriveApp.createFile(fileName, fullText, MimeType.PLAIN_TEXT);
    Logger.log("新規にテキストファイルを作成し、 " + textBuffer.length + " 件のデータを保存しました。");
  }
  
  // Gemの指示文に設定するためのURLをログに出力
  Logger.log("★Gemの指示文に貼り付けるURLはコチラです↓");
  Logger.log(file.getUrl());
}