/**
 * ============================================================================
 * AI会議_3_3d_シフト応募検知と結果反映.gs
 * 1) Gmailの「【要対応】★医師シフト応募通知★」（スレッド形式）を最小API消費で差分検知
 * 2) 現在Googleドキュメントに記載されている日付範囲（最小日〜最大日）の応募のみを記録・維持
 * 3) 交渉枠に応募が入った場合、ドキュメント・Slack（C0AJTMLF0UD）・交渉履歴へ自動反映
 * ============================================================================
 */

const SHIFT_APPLY_SHEET_NAME = 'シフト応募履歴';
const PROP_KEY_LAST_APPLY_MS = 'LAST_SHIFT_APPLY_SCAN_MS_V1';

/**
 * 🔍 【作動確認デバッグ（書き込みゼロ）】
 * 1) 現在のGoogleドキュメントの日付範囲（最小日〜最大日）を自動取得
 * 2) Gmailの「【要対応】★医師シフト応募通知★」から、その日付範囲内の応募のみを抽出
 * 3) ドキュメント上の不在枠と照合し、どう反映されるかをログ出力
 */
function debugVerifyShiftApplicationsWithinDocRange() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    Logger.log('⚠️ SINGLE_MTG_DOC_ID が設定されていません。');
    return;
  }

  const docInfo = getActiveDocDateRangeAndSlots3_3d_(docId);
  Logger.log('=== ① 現在のGoogleドキュメント対象範囲 ===');
  Logger.log('・日付範囲: ' + docInfo.minDate + ' ～ ' + docInfo.maxDate + ' （計 ' + docInfo.dates.length + ' 日分）');
  Logger.log('・ドキュメント内の現在枠数: ' + docInfo.slots.length + ' 枠');

  if (!docInfo.minDate || !docInfo.maxDate) {
    Logger.log('⚠️ ドキュメント内に日付見出し（◆ YYYY/MM/DD）が見つかりません。');
    return;
  }

  const query = 'from:caps365.jp subject:("医師シフト応募通知") newer_than:7d';
  const threads = GmailApp.search(query, 0, 5);
  const allMessages = threads.length > 0 ? GmailApp.getMessagesForThreads(threads) : [];

  const inRangeApps = [];
  const seenKeys = new Set();
  const tempApplyIndex = { bySlotKey: new Map(), byDocAndDate: new Map() };

  for (let i = 0; i < threads.length; i++) {
    const msgs = allMessages[i];
    for (let j = msgs.length - 1; j >= 0; j--) {
      const entries = parseShiftApplicationMessage3_3d_(msgs[j]);
      entries.forEach(function(ent) {
        if (ent.workDateYmd < docInfo.minDate || ent.workDateYmd > docInfo.maxDate) return;
        if (seenKeys.has(ent.uniqueKey)) return;
        seenKeys.add(ent.uniqueKey);
        inRangeApps.push(ent);

        const sKey = ent.workDateYmd + '|' + ent.clinic;
        if (!tempApplyIndex.bySlotKey.has(sKey)) tempApplyIndex.bySlotKey.set(sKey, []);
        tempApplyIndex.bySlotKey.get(sKey).push(ent);
      });
    }
  }

  Logger.log('\n=== ② ドキュメント日付範囲内 (' + docInfo.minDate + '～' + docInfo.maxDate + ') のシフト応募: ' + inRangeApps.length + ' 件 ===');
  Logger.log(JSON.stringify(inRangeApps, null, 2));

  const matchedResults = [];
  docInfo.slots.forEach(function(sl) {
    const hits = findMatchingShiftApplicationsForSlot3_3d_(sl.date, sl.cleanClinic, sl.timeRange, tempApplyIndex);
    if (hits.length > 0) {
      matchedResults.push({
        docSlot: sl.date + ' 【' + sl.cleanClinic + '】 ' + sl.timeRange,
        currentDocCol2Plans: sl.col1Snippet,
        currentDocCol3Result: sl.col2Text,
        matchedApplications: hits.map(function(h) {
          return h.doctorName + '先生 (' + h.workTime + ' / 応募日時:' + h.receivedAtStr + ')';
        }),
        willAddDocLine: hits.map(function(h) {
          return '→ ⭕' + h.doctorName + '：応募あり充足(' + h.receivedDateShort + ')';
        })
      });
    }
  });

  Logger.log('\n=== ③ 現在ドキュメントに残っている枠とシフト応募のマッチ結果: ' + matchedResults.length + ' 枠 ===');
  if (matchedResults.length === 0) {
    Logger.log('👉 現在ドキュメントに残っている未充足枠と直接重なる応募はありません（既に解消済か、別枠への応募です）。');
  } else {
    Logger.log(JSON.stringify(matchedResults, null, 2));
  }
}

/**
 * 手動即時反映用関数（通常時は既存の15分トリガー runDailyTransferAndMatching から自動で呼ばれます）
 */
function runScanShiftApplicationsAndSyncNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const newCount = scanGmailShiftApplicationsIncremental3_3d_(ss, false);
  Logger.log('📩 ドキュメント範囲内の新規シフト応募検知数: ' + newCount + ' 件');
  if (typeof syncAdjustmentResultsToDocAndSlack3_3_ === 'function') {
    syncAdjustmentResultsToDocAndSlack3_3_(ss);
  }
}

/**
 * ドキュメント内の日付範囲（minDate, maxDate）と各枠情報を読み取るヘルパー
 */
function getActiveDocDateRangeAndSlots3_3d_(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const dates = [];
  const slots = [];
  let curDate = '';

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const dm = el.asParagraph().getText().trim().match(/^◆\s*(\d{4}\/\d{2}\/\d{2})/);
      if (dm) {
        curDate = dm[1];
        if (!dates.includes(curDate)) dates.push(curDate);
      }
    } else if (el.getType() === DocumentApp.ElementType.TABLE && curDate) {
      const tbl = el.asTable();
      if (tbl.getNumRows() < 2 || !tbl.getRow(0).getCell(0).getText().includes('拠点・不在時間帯')) continue;
      for (let r = 1; r < tbl.getNumRows(); r++) {
        const row = tbl.getRow(r);
        if (row.getNumCells() < 4) continue;
        const c0Lines = row.getCell(0).getText().trim().split('\n');
        const c1 = row.getCell(1).getText().trim();
        const c2 = row.getCell(2).getText().trim();
        const cleanC = (typeof cleanClinicBaseName_ === 'function')
          ? cleanClinicBaseName_(c0Lines[0] || '')
          : String(c0Lines[0] || '').replace(/[【】\[\]（）\(\)]/g, '').trim();
        const tmMatch = c0Lines.slice(1).join(' ').match(/(\d{1,2}:\d{2}\s*[\-~～]\s*\d{1,2}:\d{2})/);
        slots.push({
          date: curDate,
          cleanClinic: cleanC,
          timeRange: tmMatch ? tmMatch[1].replace(/\s+/g, '') : '',
          col1Snippet: c1.replace(/\n+/g, ' / ').slice(0, 80),
          col2Text: c2.replace(/\n+/g, ' / ')
        });
      }
    }
  }

  dates.sort();
  return {
    minDate: dates.length > 0 ? dates[0] : '',
    maxDate: dates.length > 0 ? dates[dates.length - 1] : '',
    dates: dates,
    slots: slots
  };
}

/**
 * Gmailから「【要対応】★医師シフト応募通知★」を最新5スレッド・差分のみで高速スキャンし、
 * 現在のGoogleドキュメント日付範囲内の応募のみを『シフト応募履歴』シートへ記録・維持する関数
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

  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  let docRange = { minDate: '', maxDate: '' };
  if (docId) {
    try {
      docRange = getActiveDocDateRangeAndSlots3_3d_(docId);
    } catch (e) {}
  }

  const minWorkDateYmd = docRange.minDate || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const maxWorkDateYmd = docRange.maxDate || Utilities.formatDate(new Date(Date.now() + 21 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  // 既存シートの読み込み＆ドキュメント日付範囲外の古い行・長期先行の自動クリーンアップ
  const existingKeys = new Set();
  const existingMsgIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const allVals = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    const keptRows = [];
    let removedOutOfRange = false;

    allVals.forEach(function(r) {
      const wDate = (typeof normalizeDateYmd1b3_ === 'function') ? normalizeDateYmd1b3_(r[1]) : String(r[1] || '').trim();
      if (r[7]) existingMsgIds.add(String(r[7]).trim());
      if (wDate >= minWorkDateYmd && wDate <= maxWorkDateYmd) {
        keptRows.push(r);
        if (r[6]) existingKeys.add(String(r[6]).trim());
      } else {
        removedOutOfRange = true;
      }
    });

    if (removedOutOfRange) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      if (keptRows.length > 0) {
        sheet.getRange(2, 1, keptRows.length, headers.length).setValues(keptRows);
      }
    }
  }

  const savedLastMs = Number(props.getProperty(PROP_KEY_LAST_APPLY_MS) || 0);
  const cutoffMs = (!forceLookback7Days && savedLastMs > 0)
    ? (savedLastMs - 2 * 60 * 60 * 1000)
    : (Date.now() - 7 * 24 * 60 * 60 * 1000);

  const query = 'from:caps365.jp subject:("医師シフト応募通知") newer_than:7d';
  const threads = GmailApp.search(query, 0, 5);
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

      if (msgMs < cutoffMs) break;

      const msgId = msg.getId();
      if (existingMsgIds.has(msgId)) continue;

      const entries = parseShiftApplicationMessage3_3d_(msg);
      existingMsgIds.add(msgId);

      entries.forEach(function(ent) {
        if (ent.workDateYmd < minWorkDateYmd || ent.workDateYmd > maxWorkDateYmd) return;
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