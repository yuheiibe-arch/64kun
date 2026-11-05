/**
 * ============================================================================
 * AI会議_3_3d_シフト応募検知と結果反映.gs
 * 1) Gmailの「【要対応】★医師シフト応募通知★」（スレッド形式）を最小API消費で差分検知
 * 2) 『シフト応募履歴』シートへ蓄積し、重複読み込みを完全防止（拠点名の表記ゆれ完全正規化対応）
 * 3) 交渉していた枠に応募が入った場合（本人応募／別医師応募による充足の両方）、
 *    Googleドキュメント・Slack（C0AJTMLF0UD 既存スレッド編集）・『交渉履歴』へ自動反映
 * ============================================================================
 */

const SHIFT_APPLY_SHEET_NAME = 'シフト応募履歴';
const PROP_KEY_LAST_APPLY_MS = 'LAST_SHIFT_APPLY_SCAN_MS_V1';

/**
 * ★単体テスト・即時実行用：
 * Gmailから最新のシフト応募通知を取り込み、ドキュメント・Slack・交渉履歴へ即座に反映する関数
 */
function runScanShiftApplicationsAndSyncNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const newCount = scanGmailShiftApplicationsIncremental3_3d_(ss, false);
  Logger.log('📩 新規シフト応募検知数: ' + newCount + ' 件');
  if (typeof syncAdjustmentResultsToDocAndSlack3_3_ === 'function') {
    syncAdjustmentResultsToDocAndSlack3_3_(ss);
  }
}

/**
 * 🔍 デバッグ確認用（ドライラン・シートやSlackへの書き込みなし）：
 * 直近7日間の「【要対応】★医師シフト応募通知★」スレッドから何が抽出され、
 * 現在の不在枠・Slackスレッドとどうマッチするかをログ出力する関数
 */
function debugDryRunShiftApplications3_3d() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const query = 'subject:("医師シフト応募通知" OR "シフト応募" OR "応募がありました") newer_than:7d';
  const threads = GmailApp.search(query, 0, 30);
  Logger.log('🔎 検索ヒットスレッド数: ' + threads.length + ' 件');

  if (threads.length === 0) return;

  const allMessages = GmailApp.getMessagesForThreads(threads);
  const extractedList = [];

  for (let i = 0; i < threads.length; i++) {
    const msgs = allMessages[i];
    for (let j = msgs.length - 1; j >= 0; j--) {
      const msg = msgs[j];
      const parsedEntries = parseShiftApplicationMessage3_3d_(msg);
      parsedEntries.forEach(function(e) {
        extractedList.push(e);
      });
    }
  }

  Logger.log('📋 抽出されたシフト応募データ (' + extractedList.length + ' 件):\n' + JSON.stringify(extractedList.slice(0, 25), null, 2));
}

/**
 * 1. Gmailから「【要対応】★医師シフト応募通知★」をスレッド対応・差分のみで高速スキャンし、
 *    『シフト応募履歴』シートへ記録する関数（API消費最小化設計）
 */
function scanGmailShiftApplicationsIncremental3_3d_(ssOpt, forceLookback7Days) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  let sheet = ss.getSheetByName(SHIFT_APPLY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHIFT_APPLY_SHEET_NAME);
  }

  const headers = ['応募受信日時', '勤務日(YYYY/MM/DD)', '応募拠点', '応募時間帯', '応募医師名', '掲載日給', 'ユニークキー', 'MessageID'];
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#0f766e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const existingKeys = new Set();
  const existingMsgIds = new Set();
  if (lastRow >= 2) {
    const keyVals = sheet.getRange(2, 7, lastRow - 1, 2).getDisplayValues();
    keyVals.forEach(function(r) {
      if (r[0]) existingKeys.add(String(r[0]).trim());
      if (r[1]) existingMsgIds.add(String(r[1]).trim());
    });
  }

  const props = PropertiesService.getScriptProperties();
  const savedLastMs = Number(props.getProperty(PROP_KEY_LAST_APPLY_MS) || 0);
  const cutoffMs = (!forceLookback7Days && savedLastMs > 0)
    ? (savedLastMs - 2 * 60 * 60 * 1000)
    : (Date.now() - 7 * 24 * 60 * 60 * 1000);

  const query = 'from:caps365.jp subject:("医師シフト応募通知") newer_than:7d';
  const threads = GmailApp.search(query, 0, 25);
  if (threads.length === 0) return 0;

  const allMessages = GmailApp.getMessagesForThreads(threads);
  const newRows = [];
  let maxSeenMs = savedLastMs;

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];
    for (let j = messages.length - 1; j >= 0; j--) {
      const msg = messages[j];
      const msgMs = msg.getDate().getTime();
      if (msgMs > maxSeenMs) maxSeenMs = msgMs;

      if (msgMs < cutoffMs) {
        break;
      }

      const msgId = msg.getId();
      if (existingMsgIds.has(msgId)) {
        continue;
      }

      const entries = parseShiftApplicationMessage3_3d_(msg);
      existingMsgIds.add(msgId);

      entries.forEach(function(ent) {
        if (existingKeys.has(ent.uniqueKey)) return;
        existingKeys.add(ent.uniqueKey);
        newRows.push([
          ent.receivedAtStr,
          ent.workDateYmd,
          ent.clinic,
          ent.workTime,
          ent.doctorName,
          ent.dailyWage,
          ent.uniqueKey,
          msgId
        ]);
      });
    }
  }

  if (newRows.length > 0) {
    newRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    SpreadsheetApp.flush();
  }

  if (maxSeenMs > savedLastMs) {
    props.setProperty(PROP_KEY_LAST_APPLY_MS, String(maxSeenMs));
  }

  return newRows.length;
}

/**
 * 2. 1通の応募通知メールから応募情報（CAPSドクター365標準／共通フォーマット／MRT）を抽出する関数
 */
function parseShiftApplicationMessage3_3d_(message) {
  const receivedDateObj = message.getDate();
  const receivedAtStr = Utilities.formatDate(receivedDateObj, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  const rawHtml = message.getBody() || message.getPlainBody() || '';
  const body = rawHtml
    .replace(/<head[^>]*>[\s\S]*?<\/head>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>|<\/div>|<p>|<div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/^[ \t]+/gm, '')
    .replace(/[\n\r]+/g, '\n');

  let defaultDocName = '';
  const docMatch2 = body.match(/(?:勤務医師名|応募医師名|医師氏名|医師名|氏名)[\s ]*[：:]?[\s ]*([^\n\r]+)/);
  const docMatch1 = body.match(/^\s*([^\n\r]+?)\s*先生/m);
  if (docMatch2) {
    defaultDocName = docMatch2[1].replace(/先生/g, '').replace(/[\s\u3000]+/g, '').trim();
  } else if (docMatch1) {
    defaultDocName = docMatch1[1].replace(/[\s\u3000]+/g, '').trim();
  }

  const wageMatch = body.match(/掲載日給額[\s ]*[：:]?[\s ]*([0-9,]+円)/);
  const defaultWage = wageMatch ? wageMatch[1] : '';

  const results = [];
  const addEntry = function(rawDateStr, rawTimeStr, rawClinicStr, docNameOpt) {
    const dMatch = String(rawDateStr || '').match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
    if (!dMatch) return;
    const workDateYmd = dMatch[1] + '/' + ('0' + dMatch[2]).slice(-2) + '/' + ('0' + dMatch[3]).slice(-2);

    const tMatch = String(rawTimeStr || '').match(/(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/);
    const workTime = tMatch ? (('0' + tMatch[1]).slice(-5) + '-' + ('0' + tMatch[2]).slice(-5)) : '';

    // ★修正：cleanClinicBaseName_ を使用して拠点名の表記ゆれ（新鎌ヶ谷➔新鎌ケ谷、セブンパーク天美➔天美 等）を完全正規化
    const rawTrimmed = String(rawClinicStr || '')
      .replace(/[【】\[\]（）\(\)]/g, ' ')
      .replace(/小児科|内科|耳鼻咽喉科|皮膚科|整形外科/g, '')
      .replace(/[\s\u3000]+/g, '')
      .trim();
    const clinic = (typeof cleanClinicBaseName_ === 'function')
      ? cleanClinicBaseName_(rawTrimmed)
      : rawTrimmed;

    const docClean = String(docNameOpt || defaultDocName || '').replace(/先生/g, '').replace(/[\s\u3000]+/g, '').trim();
    if (!clinic || !docClean || docClean.includes('テスト') || clinic.toLowerCase().includes('test')) return;

    const uniqueKey = [workDateYmd, clinic, docClean, workTime].join('_');
    results.push({
      receivedAtStr: receivedAtStr,
      receivedDateShort: receivedAtStr.slice(5, 10),
      workDateYmd: workDateYmd,
      clinic: clinic,
      workTime: workTime,
      doctorName: docClean,
      dailyWage: defaultWage,
      uniqueKey: uniqueKey
    });
  };

  const regexCaps = /シフト応募日[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日\s*[\(（][月火水木金土日][\)）])(?:\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2}))?[^\n\r]*[\n\r]+応募拠点名[：:]\s*([^\n\r]+)/g;
  let m;
  while ((m = regexCaps.exec(body)) !== null) {
    const afterText = body.slice(m.index, m.index + 250);
    const localDocMatch = afterText.match(/(?:勤務医師名|応募医師名)[：:]\s*([^\n\r]+)/);
    const localDoc = localDocMatch ? localDocMatch[1] : defaultDocName;
    addEntry(m[1], m[2] || '', m[3], localDoc);
  }

  if (results.length === 0) {
    const regexCommon = /(\d{4}年\d{1,2}月\d{1,2}日\s*[\(（][月火水木金土日][\)）])\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})\s*[\(（]([^\)）]+?)[\)）]/g;
    while ((m = regexCommon.exec(body)) !== null) {
      addEntry(m[1], m[2], m[3], defaultDocName);
    }
  }

  if (results.length === 0) {
    const mrtDate = body.match(/日時\s*:\s*(\d{4}\/\d{1,2}\/\d{1,2})\s*[\(（][月火水木金土日][\)）]\s*(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/);
    const mrtClinic = body.match(/クリニック\s*:\s*(?:\d+\s+)?([^\n\r]+)/);
    if (mrtDate && mrtClinic) {
      addEntry(mrtDate[1], mrtDate[2], mrtClinic[1], defaultDocName || 'MRT応募医師');
    }
  }

  return results;
}

/**
 * 3. 『シフト応募履歴』シートから応募インデックスを構築する関数
 *    ・bySlotKey: "YYYY/MM/DD|拠点名" ➔ [{ doctorName, workTime, receivedDateShort, receivedAtStr, dailyWage }]
 *    ・byDocAndDate: "医師名|YYYY/MM/DD" ➔ [{ clinic, workTime, receivedDateShort }]
 */
function buildShiftApplicationIndex3_3d_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const index = {
    bySlotKey: new Map(),
    byDocAndDate: new Map()
  };

  const sheet = ss.getSheetByName(SHIFT_APPLY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return index;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
  rows.forEach(function(r) {
    const receivedAtStr = String(r[0] || '').trim();
    const workDateYmd = (typeof normalizeDateYmd1b3_ === 'function')
      ? normalizeDateYmd1b3_(r[1])
      : String(r[1] || '').trim();
    const rawClinic = String(r[2] || '').trim();
    const clinic = (typeof cleanClinicBaseName_ === 'function')
      ? cleanClinicBaseName_(rawClinic)
      : rawClinic;
    const workTime = String(r[3] || '').trim();
    const doctorName = String(r[4] || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '').trim();
    const dailyWage = String(r[5] || '').trim();
    if (!workDateYmd || !clinic || !doctorName) return;

    const entry = {
      receivedAtStr: receivedAtStr,
      receivedDateShort: receivedAtStr.slice(5, 10),
      workDateYmd: workDateYmd,
      clinic: clinic,
      workTime: workTime,
      doctorName: doctorName,
      dailyWage: dailyWage
    };

    const slotKey = workDateYmd + '|' + clinic;
    if (!index.bySlotKey.has(slotKey)) index.bySlotKey.set(slotKey, []);
    index.bySlotKey.get(slotKey).push(entry);

    const docDateKey = doctorName + '|' + workDateYmd;
    if (!index.byDocAndDate.has(docDateKey)) index.byDocAndDate.set(docDateKey, []);
    index.byDocAndDate.get(docDateKey).push(entry);
  });

  return index;
}

/**
 * 4. 2つの時間帯（例: "18:00-21:00" と "18:00-21:00" または "15:00-21:00"）が重なっているか判定するヘルパー
 */
function isShiftTimeOverlapping3_3d_(timeRangeA, timeRangeB) {
  if (!timeRangeA || !timeRangeB) return true;
  const parseMin = function(s) {
    const m = String(s).match(/(\d{1,2}):(\d{2})\s*[\-~～〜]\s*(\d{1,2}):(\d{2})/);
    if (!m) return null;
    return {
      start: Number(m[1]) * 60 + Number(m[2]),
      end: Number(m[3]) * 60 + Number(m[4])
    };
  };
  const a = parseMin(timeRangeA);
  const b = parseMin(timeRangeB);
  if (!a || !b) return true;
  const overlapStart = Math.max(a.start, b.start);
  const overlapEnd = Math.min(a.end, b.end);
  return (overlapEnd - overlapStart) >= 60;
}

/**
 * 5. 特定の枠（日付・拠点・時間帯）に対して入ったシフト応募を取得する関数
 */
function findMatchingShiftApplicationsForSlot3_3d_(workDateYmd, cleanClinic, targetTimeRange, applyIndex) {
  if (!applyIndex || !applyIndex.bySlotKey) return [];
  const normClinic = (typeof cleanClinicBaseName_ === 'function')
    ? cleanClinicBaseName_(cleanClinic)
    : String(cleanClinic || '').trim();
  const slotKey = workDateYmd + '|' + normClinic;
  const candidates = applyIndex.bySlotKey.get(slotKey) || [];
  return candidates.filter(function(ent) {
    return isShiftTimeOverlapping3_3d_(targetTimeRange, ent.workTime);
  });
}