/**
 * ============================================================================
 * ドキュメント転送.gs
 * （※本番の exportToTxtFile は『コード.gs』側に統合済みのため、名前衝突防止済）
 * ============================================================================
 */

function exportToTxtFile_Legacy_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var textBuffer = [];

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var questioner = data[i][1];
    var question = data[i][2];
    var answer = data[i][3];
    var answerer = data[i][4];
    var refUrl = data[i][5];

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

  var files = DriveApp.getFilesByName(fileName);
  var file;

  if (files.hasNext()) {
    file = files.next();
    file.setContent(fullText);
    Logger.log("既存のテキストファイルに " + textBuffer.length + " 件のデータを上書き保存しました。");
  } else {
    file = DriveApp.createFile(fileName, fullText, MimeType.PLAIN_TEXT);
    Logger.log("新規にテキストファイルを作成し、 " + textBuffer.length + " 件のデータを保存しました。");
  }

  Logger.log("★Gemの指示文に貼り付けるURLはコチラです↓");
  Logger.log(file.getUrl());
}