// ===================================================================
// ★★★ 1日1回、この「まとめ役」関数を実行するトリガーを1つだけ設定します ★★★
// ===================================================================
function runDailyTransferAndMatching() {
  const SPREADSHEET_ID = '1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM';

  getQuestionEmailsFromGmail_Final(SPREADSHEET_ID);
  getAnswerEmailsFromGmail_Final(SPREADSHEET_ID);
  matchAndCreateKnowledgeBase_Final(SPREADSHEET_ID);
  exportToTxtFile(SPREADSHEET_ID);
}

/**
 * 医師名のスペース（全角・半角）を除去して比較しやすくするヘルパー
 */
function normalizeDocName_(name) {
  return String(name || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '').trim();
}

/**
 * ステージング（テスト環境）や自動配送テストのメールかを判定するヘルパー
 */
function isStagingOrTestMail_(message, bodyText, doctorName) {
  const fromStr = String(message.getFrom() || '').toLowerCase();
  const body = String(bodyText || '');
  const doc = String(doctorName || '');
  if (fromStr.includes('staging')) return true;
  if (body.includes('staging.capsdoctor365.jp')) return true;
  if (doc.includes('自動確認医師') || doc.includes('テスト')) return true;
  if (body.includes('【自動確認】')) return true;
  return false;
}

/**
 * 「2026年11月15日」などの日本語日付を「2026/11/15」に正規化するヘルパー
 */
function normalizeCancelTargetDateYmd_(rawDateStr) {
  const m = String(rawDateStr || '').match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
  if (!m) return '';
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 『シフトキャンセル履歴』シートを取得（なければ自動作成）するヘルパー
 */
function ensureShiftCancelHistorySheet_(ss) {
  const sheetName = 'シフトキャンセル履歴';
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    const headers = ['管理番号', 'キャンセル承認日時', '医師名', 'キャンセル対象日', '拠点名', '診療科', '時間帯', 'メールID'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#c53929').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * 1. Gmailから「【要確認】医師からのお問い合わせ」を検索し、『医師質問』シートに自動転記
 * ※ステージング環境や自動確認テストメールは除外します
 */
function getQuestionEmailsFromGmail_Final(spreadsheetId) {
  const sheetName = '医師質問';
  const timezone = 'Asia/Tokyo';
  const sourceEmail = 'doctor-support@caps365.jp';

  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.error('シート「' + sheetName + '」が見つかりません。');
    return;
  }

  const searchQuery = 'from:' + sourceEmail + ' subject:("【要確認】医師からのお問い合わせ") newer_than:7d';
  const threads = GmailApp.search(searchQuery, 0, 80);

  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? new Set(sheet.getRange(2, 5, lastRow - 1, 1).getDisplayValues().flat().filter(Boolean))
    : new Set();
  let lastNumber = (lastRow > 1 && !isNaN(Number(sheet.getRange(lastRow, 1).getValue())))
    ? Number(sheet.getRange(lastRow, 1).getValue())
    : 0;

  const newItems = [];

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      const messageId = message.getId();
      if (existingIds.has(messageId)) return;

      const body = message.getPlainBody();
      const doctorMatch = body.match(/【医師名】[：:]\s*(.+)/);
      const contentMatch = body.match(/【内容】[：:]\s*([\s\S]+)/);

      if (doctorMatch && contentMatch) {
        const doctorName = doctorMatch[1].trim();
        const content = contentMatch[1].trim();
        if (isStagingOrTestMail_(message, body, doctorName)) return;

        const msgDate = message.getDate();
        const receivedTimestamp = Utilities.formatDate(msgDate, timezone, 'yyyy/MM/dd HH:mm:ss');

        existingIds.add(messageId);
        newItems.push({
          timeMs: msgDate.getTime(),
          row: [0, receivedTimestamp, doctorName, content, messageId]
        });
      }
    });
  });

  newItems.sort(function(a, b) { return a.timeMs - b.timeMs; });
  for (let i = 0; i < newItems.length; i++) {
    lastNumber++;
    newItems[i].row[0] = lastNumber;
    sheet.appendRow(newItems[i].row);
  }

  console.log(newItems.length + '件の質問・返信メールを転記しました。');
}

/**
 * 2. Gmailから「お問い合わせスタッフ送信」を検索し、
 *   - 通常メールは『DS回答』シートへ自動転記
 *   - 「のキャンセル申請が承認されました。」メールは『シフトキャンセル履歴』シートへ自動蓄積（設計1）
 */
function getAnswerEmailsFromGmail_Final(spreadsheetId) {
  const sheetName = 'DS回答';
  const timezone = 'Asia/Tokyo';
  const sourceEmail = 'doctor-support@caps365.jp';

  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.error('シート「' + sheetName + '」が見つかりません。');
    return;
  }

  const cancelSheet = ensureShiftCancelHistorySheet_(ss);
  const cancelLastRow = cancelSheet.getLastRow();
  const existingCancelIds = cancelLastRow > 1
    ? new Set(cancelSheet.getRange(2, 8, cancelLastRow - 1, 1).getDisplayValues().flat().filter(Boolean))
    : new Set();
  let lastCancelNumber = (cancelLastRow > 1 && !isNaN(Number(cancelSheet.getRange(cancelLastRow, 1).getValue())))
    ? Number(cancelSheet.getRange(cancelLastRow, 1).getValue())
    : 0;

  // 初回（キャンセル履歴が空のとき）は直近30日分をスキャンして過去のキャンセルも一括回収する
  const daysWindow = cancelLastRow <= 1 ? '30d' : '7d';
  const searchQuery = 'from:' + sourceEmail + ' subject:("お問い合わせスタッフ送信") newer_than:' + daysWindow;
  const threads = GmailApp.search(searchQuery, 0, 100);

  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? new Set(sheet.getRange(2, 7, lastRow - 1, 1).getDisplayValues().flat().filter(Boolean))
    : new Set();
  let lastNumber = (lastRow > 1 && !isNaN(Number(sheet.getRange(lastRow, 1).getValue())))
    ? Number(sheet.getRange(lastRow, 1).getValue())
    : 0;

  const newItems = [];
  const newCancelItems = [];

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      const messageId = message.getId();
      const body = message.getPlainBody();
      const doctorMatch = body.match(/【医師名】[：:]\s*(.+)/);
      const contentMatch = body.match(/【内容】[：:]\s*([\s\S]+)/);
      if (!doctorMatch || !contentMatch) return;

      const doctorName = doctorMatch[1].trim();
      const content = contentMatch[1].trim();
      if (isStagingOrTestMail_(message, body, doctorName)) return;

      const msgDate = message.getDate();
      const sentTimestamp = Utilities.formatDate(msgDate, timezone, 'yyyy/MM/dd HH:mm:ss');

      // ★設計1：シフトキャンセル承認メールの自動回収
      if (body.includes('のキャンセル申請が承認されました。')) {
        if (existingCancelIds.has(messageId)) return;
        const cancelMatch = content.match(/(\d{4}年\d{1,2}月\d{1,2}日)[\s ]+([^\s ]+)[\s ]+([^\s ]+)[\s ]+(\d{1,2}:\d{2}[~～\-]\d{1,2}:\d{2})※?のキャンセル申請が承認されました/);
        const targetDateYmd = cancelMatch ? normalizeCancelTargetDateYmd_(cancelMatch[1]) : '';
        const clinicName = cancelMatch ? cancelMatch[2].trim() : '';
        const deptName = cancelMatch ? cancelMatch[3].trim() : '';
        const timeRange = cancelMatch ? cancelMatch[4].replace(/[~～]/g, '-').trim() : '';

        if (targetDateYmd) {
          existingCancelIds.add(messageId);
          newCancelItems.push({
            timeMs: msgDate.getTime(),
            row: [0, sentTimestamp, doctorName, targetDateYmd, clinicName, deptName, timeRange, messageId]
          });
        }
        return;
      }

      // 通常のスタッフ送信メール（7日以内）
      if (existingIds.has(messageId)) return;
      if ((Date.now() - msgDate.getTime()) > 8 * 24 * 60 * 60 * 1000) return;

      let sender = '（不明）';
      const lines = content.trim().split('\n').map(function(l) { return l.trim(); }).filter(Boolean);
      const lastLine = lines.length > 0 ? lines[lines.length - 1] : '';
      const senderMatch = lastLine.replace(/^(ドクターサポート部|ドクターサポート事務局)[\s\u3000]*/, '');
      if (senderMatch && senderMatch.length <= 15) {
        sender = senderMatch;
      }

      const urlMatch = content.match(/https?:\/\/[^\s]+/);
      const url = urlMatch ? urlMatch[0] : '';

      existingIds.add(messageId);
      newItems.push({
        timeMs: msgDate.getTime(),
        row: [0, sentTimestamp, doctorName, content, sender, url, messageId]
      });
    });
  });

  newItems.sort(function(a, b) { return a.timeMs - b.timeMs; });
  for (let i = 0; i < newItems.length; i++) {
    lastNumber++;
    newItems[i].row[0] = lastNumber;
    sheet.appendRow(newItems[i].row);
  }

  newCancelItems.sort(function(a, b) { return a.timeMs - b.timeMs; });
  for (let j = 0; j < newCancelItems.length; j++) {
    lastCancelNumber++;
    newCancelItems[j].row[0] = lastCancelNumber;
    cancelSheet.appendRow(newCancelItems[j].row);
  }

  console.log(newItems.length + '件の回答・打診メール、および ' + newCancelItems.length + '件のシフトキャンセル履歴を転記しました。');
}

/**
 * 3. 『医師質問』と『DS回答』を双方向突合し、『ナレッジシート』を自動生成
 */
function matchAndCreateKnowledgeBase_Final(spreadsheetId) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  const questionSheet = ss.getSheetByName('医師質問');
  const answerSheet = ss.getSheetByName('DS回答');
  const kbSheet = ss.getSheetByName('ナレッジシート');

  if (!questionSheet || !answerSheet || !kbSheet) {
    console.error('必要なシート（医師質問, DS回答, ナレッジシート）が見つかりません。');
    return;
  }

  if (!kbSheet.getRange(1, 7).getValue()) kbSheet.getRange(1, 7).setValue('日時');
  if (!kbSheet.getRange(1, 8).getValue()) kbSheet.getRange(1, 8).setValue('NG候補・種別フラグ');

  if (questionSheet.getLastRow() < 2 || answerSheet.getLastRow() < 2) return;

  const qMaxCols = Math.max(questionSheet.getLastColumn(), 7);
  const aMaxCols = Math.max(answerSheet.getLastColumn(), 8);
  const questions = questionSheet.getRange(2, 1, questionSheet.getLastRow() - 1, qMaxCols).getValues();
  const answers = answerSheet.getRange(2, 1, answerSheet.getLastRow() - 1, aMaxCols).getValues();

  let matchCount = 0;
  const lastKbRow = kbSheet.getLastRow();
  let lastKbNumber = (lastKbRow > 1 && !isNaN(Number(kbSheet.getRange(lastKbRow, 1).getValue())))
    ? Number(kbSheet.getRange(lastKbRow, 1).getValue())
    : 0;

  const MAX_WINDOW_MS = 7 * 24 * 60 * 60 * 1000;

  const isShiftPitch = function(text) {
    return /ご相談|ご変更いただくことは可能でしょうか|延長していただくことは可能でしょうか|前倒ししていただくことは可能でしょうか|医師不在となる可能性|ご勤務につきまして/.test(String(text || ''));
  };

  // ①【スタッフ打診 ➔ 医師返信】の突合（シフト調整パターン）
  for (let j = 0; j < answers.length; j++) {
    const aRow = answers[j];
    if (String(aRow[7] || '').trim() === '処理済み') continue;

    const aContent = String(aRow[3] || '');
    if (!isShiftPitch(aContent)) continue;

    const aTimestamp = new Date(aRow[1]);
    const aDocNorm = normalizeDocName_(aRow[2]);
    if (isNaN(aTimestamp.getTime()) || !aDocNorm) continue;

    let bestQIndex = -1;
    let smallestDiff = Infinity;

    for (let i = 0; i < questions.length; i++) {
      const qRow = questions[i];
      if (String(qRow[6] || '').trim() === '処理済み') continue;

      const qDocNorm = normalizeDocName_(qRow[2]);
      const qTimestamp = new Date(qRow[1]);
      if (qDocNorm !== aDocNorm || isNaN(qTimestamp.getTime())) continue;

      const diff = qTimestamp.getTime() - aTimestamp.getTime();
      if (diff > 0 && diff <= MAX_WINDOW_MS && diff < smallestDiff) {
        smallestDiff = diff;
        bestQIndex = i;
      }
    }

    if (bestQIndex !== -1) {
      const qRow = questions[bestQIndex];
      const docName = qRow[2];
      const docReply = String(qRow[3] || '');
      const aSender = aRow[4];
      const aUrl = aRow[5];
      const dtStr = Utilities.formatDate(new Date(qRow[1]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const flag = detectNgOrConditionFlag_(docReply, aContent);

      lastKbNumber++;
      kbSheet.appendRow([
        lastKbNumber,
        docName,
        '【医師返信】\n' + docReply,
        '【スタッフ打診】\n' + aContent,
        aSender,
        aUrl,
        dtStr,
        flag
      ]);

      questionSheet.getRange(bestQIndex + 2, 7).setValue('処理済み');
      answerSheet.getRange(j + 2, 8).setValue('処理済み');
      questions[bestQIndex][6] = '処理済み';
      answers[j][7] = '処理済み';
      matchCount++;
    }
  }

  // ②【医師質問 ➔ スタッフ回答】の突合（従来のCS・医師発信パターン）
  for (let i = 0; i < questions.length; i++) {
    const qRow = questions[i];
    if (String(qRow[6] || '').trim() === '処理済み') continue;

    const qTimestamp = new Date(qRow[1]);
    const qAsker = qRow[2];
    const qDocNorm = normalizeDocName_(qAsker);
    const qContent = String(qRow[3] || '');
    if (isNaN(qTimestamp.getTime()) || !qDocNorm) continue;

    let bestMatch = null;
    let bestMatchIndex = -1;
    let smallestTimeDiff = Infinity;

    for (let j = 0; j < answers.length; j++) {
      const aRow = answers[j];
      if (String(aRow[7] || '').trim() === '処理済み') continue;

      const aTimestamp = new Date(aRow[1]);
      const aDocNorm = normalizeDocName_(aRow[2]);

      if (qDocNorm === aDocNorm && !isNaN(aTimestamp.getTime()) && aTimestamp > qTimestamp) {
        const timeDiff = aTimestamp.getTime() - qTimestamp.getTime();
        if (timeDiff <= MAX_WINDOW_MS && timeDiff < smallestTimeDiff) {
          smallestTimeDiff = timeDiff;
          bestMatch = aRow;
          bestMatchIndex = j;
        }
      }
    }

    if (bestMatch) {
      const aContent = String(bestMatch[3] || '');
      const aSender = bestMatch[4];
      const aUrl = bestMatch[5];
      const dtStr = Utilities.formatDate(qTimestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const flag = detectNgOrConditionFlag_(qContent, aContent);

      lastKbNumber++;
      kbSheet.appendRow([lastKbNumber, qAsker, qContent, aContent, aSender, aUrl, dtStr, flag]);

      questionSheet.getRange(i + 2, 7).setValue('処理済み');
      answerSheet.getRange(bestMatchIndex + 2, 8).setValue('処理済み');
      questions[i][6] = '処理済み';
      answers[bestMatchIndex][7] = '処理済み';
      matchCount++;
    }
  }

  console.log(matchCount + '件のナレッジを生成しました。');
}

/**
 * ★強化版：H列に入れる「NG候補・種別フラグ」の自動判定
 * （「18時までなら可能」「延長不可」などの時間条件付き返信も逃さずフラグ付け）
 */
function detectNgOrConditionFlag_(docText, staffText) {
  const d = String(docText || '');
  const s = String(staffText || '');

  if (/(\d{1,2})\s*(?:時|:00)\s*(?:までなら|まででしたら|までであれば|までしか|で上が|以降は難|以降は不可|以降は厳)|延長は(?:難|不可|厳)|早出は(?:難|不可|厳)/.test(d)) {
    return 'シフト調整（時間条件・上限あり）';
  }
  if (/からの勤務のため|遠方|以降は難しい|曜日は.*不可能|曜日は.*難しい|より早く着くことは不可能|2診体制でしたら|2診なら|応援に行く事は可能|変更いただいても大丈夫|働けるので声をかけて/.test(d)) {
    return '★要確認（交渉注意医師シート記載候補）';
  }
  if (/ご快諾いただき|にてご勤務確定|勤務変更.*承知|変更の旨承知|延長問題ありません|どちらも可能です/.test(d + '\n' + s)) {
    return 'シフト調整（承諾・成立）';
  }
  if (/難しいです|困難です|予定を入れてしまいました|お力になれず|都合がつかず|可能な日はありません/.test(d)) {
    return 'シフト調整（単発辞退）';
  }
  return '一般連絡';
}

/**
 * 4. 『ナレッジシート』から Chatwork Bot ＆ AI総務用のテキストファイルを出力
 */
function exportToTxtFile(spreadsheetId) {
  var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ナレッジシート') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var textBuffer = [];

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var questioner = data[i][1];
    var question = data[i][2];
    var answer = data[i][3];
    var answerer = data[i][4];
    var refUrl = data[i][5];
    var dt = data[i][6] || '';
    var flag = data[i][7] || '';

    if (question && answer) {
      var block = "【管理番号】 " + id + "\n" +
                  "【日時】 " + (dt ? dt : "不明") + "\n" +
                  "【判定フラグ】 " + (flag ? flag : "一般連絡") + "\n" +
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

  Logger.log("★ファイルURL: " + file.getUrl());
}