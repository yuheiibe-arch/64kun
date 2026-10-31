/**
 * ============================================================================
 * AI連携_復号受信.gs（caps365.jp 側：マスタ動的参照 ＋ シンプルメニュー統合版）
 * ============================================================================
 */

const AI_RECEIVER_CONFIG = {
  SUBJECT_QUERY: 'subject:"【AI連携暗号化JSON】医師不在・シフト状況パック"',
  DEFAULT_SECRET_KEY: 'CAPS_MNYS_AI_AGENT_SECRET_2026_KEY_#9876',
  LAST_PROCESSED_MSG_KEY: 'AI_BRIDGE_LAST_PROCESSED_MSG_ID',

  // マスタ参照シート名（DSナレッジシート内）
  SHEET_SBU_MASTER: 'SBU',
  SHEET_FULLTIME_ALLOW_MASTER: '常勤依頼手当',
  SHEET_NEGOTIABLE_MASTER: '交渉しやすい医師',
  SHEET_CAUTION_MASTER: '交渉注意医師',
  SHEET_KNOWLEDGE_HISTORY: 'ナレッジシート',

  // 自動生成・更新されるシート名
  SHEET_UNFILLED: '最新_不在枠',
  SHEET_ACTIVE_SHIFTS: '最新_当日シフト',
  SHEET_OUBO_SHIFTS: '最新_応募シフト',
  SHEET_STREAK_WARN: '最新_連勤注意リスト',
  SHEET_SPOT_CANDIDATES: '最新_個別打診候補',
  SHEET_BUNDLED_CANDIDATES: '最新_おまとめ打診候補'
};

/**
 * スプレッドシート起動時のシンプルメニュー登録
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🤖 AI超募集MTG')
    .addItem('🎮 ① 編集くん 64 を開く（シフト確定・mnys提出）', 'openHenshukunModal')
    .addSeparator()
    .addItem('📄 ② AI事前会議ドキュメントを生成', 'runDemoAndSubmitMeetingDoc')
    .addItem('📥 ③ 最新の暗号化シフトデータを取り込む', 'importLatestEncryptedShiftContextManual')
    .addToUi();
}

function installReceiverTrigger() {
  const funcName = 'autoCheckEncryptedMailTrigger';
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(funcName).timeBased().everyMinutes(15).create();
  Logger.log('✅ 受け取り側の自動取込トリガー（15分間隔）を設定しました。');
}

function autoCheckEncryptedMailTrigger() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    importLatestEncryptedShiftContext_(false);
  } catch (e) {
    Logger.log('自動取込エラー（次回トリガーまで待機します）: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function importLatestEncryptedShiftContextManual() {
  importLatestEncryptedShiftContext_(true);
}

function normalizeDocKeyReceive_(name) {
  return String(name || '')
    .replace(/[\(（\[].*?[\)）\]]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

/**
 * 『DSナレッジシート』内の各マスタ・履歴タブを動的に読み込む
 */
function loadDoctorRoleMastersFromSheets_(ss) {
  const sbuMap = new Map();
  const ftAllowMap = new Map();
  const negotiableMap = new Map();
  const cautionMap = new Map();
  const recentKbMap = new Map();

  const sbuSheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_SBU_MASTER);
  if (sbuSheet && sbuSheet.getLastRow() >= 2) {
    const vals = sbuSheet.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const nameKey = normalizeDocKeyReceive_(vals[r][0]);
      if (!nameKey) continue;
      sbuMap.set(nameKey, {
        licenseNo: String(vals[r][1] || '').trim(),
        memo: [vals[r][2], vals[r][3]].map(function(v) { return String(v || '').trim(); }).filter(Boolean).join(' / ')
      });
    }
  }

  const ftSheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_FULLTIME_ALLOW_MASTER);
  if (ftSheet && ftSheet.getLastRow() >= 2) {
    const vals = ftSheet.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const nameKey = normalizeDocKeyReceive_(vals[r][0]);
      if (!nameKey) continue;
      ftAllowMap.set(nameKey, {
        licenseNo: String(vals[r][1] || '').trim(),
        memo: [vals[r][2], vals[r][3]].map(function(v) { return String(v || '').trim(); }).filter(Boolean).join(' / ')
      });
    }
  }

  const negSheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_NEGOTIABLE_MASTER);
  if (negSheet && negSheet.getLastRow() >= 2) {
    const vals = negSheet.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const nameKey = normalizeDocKeyReceive_(vals[r][0]);
      if (!nameKey) continue;
      negotiableMap.set(nameKey, {
        licenseNo: String(vals[r][1] || '').trim(),
        route: String(vals[r][2] || '').trim(),
        humanMemo: String(vals[r][3] || '').trim(),
        aiMemo: String(vals[r][4] || '').trim()
      });
    }
  }

  const cautionSheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_CAUTION_MASTER);
  if (cautionSheet && cautionSheet.getLastRow() >= 2) {
    const vals = cautionSheet.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const nameKey = normalizeDocKeyReceive_(vals[r][0]);
      if (!nameKey) continue;
      const ngText = vals[r].slice(1).map(function(v) { return String(v || '').trim(); }).filter(Boolean).join(' / ');
      if (ngText) cautionMap.set(nameKey, ngText);
    }
  }

  const kbSheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_KNOWLEDGE_HISTORY);
  if (kbSheet) {
    const kbLast = kbSheet.getLastRow();
    if (kbLast >= 2) {
      const startRow = Math.max(2, kbLast - 199);
      const vals = kbSheet.getRange(startRow, 1, kbLast - startRow + 1, 8).getDisplayValues();
      for (let r = 0; r < vals.length; r++) {
        const nameKey = normalizeDocKeyReceive_(vals[r][1]);
        if (!nameKey) continue;
        const reply = String(vals[r][2] || '').replace(/\n+/g, ' ').slice(0, 40);
        const dt = String(vals[r][6] || '').slice(5, 10);
        const flag = String(vals[r][7] || '').trim();
        if (flag.includes('シフト調整') || flag) {
          recentKbMap.set(nameKey, '直近' + (dt || '履歴') + '打診(' + flag + ':' + reply + ')');
        }
      }
    }
  }

  return {
    sbuMap: sbuMap,
    ftAllowMap: ftAllowMap,
    negotiableMap: negotiableMap,
    cautionMap: cautionMap,
    recentKbMap: recentKbMap
  };
}

/**
 * 医師名文字列に対して『SBU』『常勤依頼手当』『交渉しやすい医師』『交渉注意医師』シートの情報を動的に付与する
 */
function enrichDoctorTextWithSheetMasters_(rawText, roleMasters) {
  if (!rawText) return '';
  const sbuMap = roleMasters.sbuMap;
  const ftAllowMap = roleMasters.ftAllowMap;
  const negotiableMap = roleMasters.negotiableMap;
  const cautionMap = roleMasters.cautionMap;

  return String(rawText).split(/([,、，\n]+)/).map(function(token) {
    if (/^[,、，\n]+$/.test(token) || !token.trim()) return token;

    const parts = token.split(/([：:])/);
    const nameWithTag = parts[0].trim();
    const restPart = parts.slice(1).join('');

    const isShotei = nameWithTag.includes('所定休出');
    const isFullTime = nameWithTag.includes('(常勤');
    const baseName = nameWithTag.replace(/[\(（].*?[\)）]/g, '').trim();
    const docKey = normalizeDocKeyReceive_(baseName);
    if (!docKey || docKey === '未充足') return token;

    const cautionSuffix = (cautionMap && cautionMap.has(docKey))
      ? '[⚠️注意:' + cautionMap.get(docKey) + ']'
      : '';

    if (sbuMap.has(docKey)) {
      const sbuInfo = sbuMap.get(docKey);
      const memoSuffix = sbuInfo.memo ? '[' + sbuInfo.memo + ']' : '';
      return baseName + '(★SBU' + memoSuffix + cautionSuffix + ')' + restPart;
    }

    if (ftAllowMap.has(docKey)) {
      const ftInfo = ftAllowMap.get(docKey);
      const memoSuffix = ftInfo.memo ? '[' + ftInfo.memo + ']' : '';
      if (isShotei) {
        return baseName + '(常勤・所定休出★移動相談可' + memoSuffix + cautionSuffix + ')' + restPart;
      } else if (isFullTime) {
        return baseName + '(常勤※通常勤務のため移動不可' + memoSuffix + cautionSuffix + ')' + restPart;
      }
    }

    if (negotiableMap.has(docKey)) {
      const negInfo = negotiableMap.get(docKey);
      const details = [negInfo.route, negInfo.humanMemo, negInfo.aiMemo].filter(Boolean).join(' / ');
      const detailSuffix = details ? '[' + details + ']' : '';
      return baseName + '(★交渉しやすい医師' + detailSuffix + cautionSuffix + ')' + restPart;
    }

    if (cautionSuffix) {
      return nameWithTag + cautionSuffix + restPart;
    }

    return token;
  }).join('');
}

function importLatestEncryptedShiftContext_(isForceManual) {
  const latestMsg = findLatestEncryptedMessage_();
  if (!latestMsg) {
    Logger.log('対象となる【AI連携暗号化JSON】メールがまだ届いていません。');
    return;
  }

  const msgId = latestMsg.getId();
  const props = PropertiesService.getScriptProperties();
  const lastProcessedId = props.getProperty(AI_RECEIVER_CONFIG.LAST_PROCESSED_MSG_KEY);

  if (!isForceManual && lastProcessedId === msgId) {
    Logger.log('最新メールはすでに取込済みのためスキップしました。');
    return;
  }

  props.setProperty(AI_RECEIVER_CONFIG.LAST_PROCESSED_MSG_KEY, msgId);

  const body = latestMsg.getPlainBody();
  const match = body.match(/---BEGIN_ENCRYPTED_AI_PAYLOAD---\s*([\s\S]*?)\s*---END_ENCRYPTED_AI_PAYLOAD---/);
  if (!match || !match[1]) {
    throw new Error('メール本文から暗号化データを抽出できませんでした。');
  }

  const secretKey = props.getProperty('AI_BRIDGE_SECRET_KEY')
    || AI_RECEIVER_CONFIG.DEFAULT_SECRET_KEY;

  const encryptedPackage = JSON.parse(match[1].trim());
  const payloadObj = verifyAndDecryptPayload_(encryptedPackage, secretKey);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roleMasters = loadDoctorRoleMastersFromSheets_(ss);

  writeUnfilledSheet_(ss, payloadObj.unfilledList || [], payloadObj.generatedAt, roleMasters);
  writeActiveShiftsSheet_(ss, payloadObj.activeShiftsByDate || {}, roleMasters);
  writeOuboShiftsSheet_(ss, payloadObj.ouboShiftsByDate || {}, roleMasters);
  writeStreakWarningsSheet_(ss, payloadObj.streakWarnings || [], roleMasters);

  writeKinmukunCandidateSheets_(ss, payloadObj, roleMasters);

  Logger.log(
    '✅ 復号・全6シート更新完了！（生成日時: ' + payloadObj.generatedAt + ' / ' +
    '不在枠: ' + (payloadObj.unfilledList || []).length + '件 / ' +
    '個別打診候補: ' + Math.max(0, (payloadObj.spotCandidateRows || []).length - 1) + '件 / ' +
    '複数日程おまとめ: ' + Math.max(0, (payloadObj.bundledCandidateRows || []).length - 1) + '名）'
  );
}

function findLatestEncryptedMessage_() {
  const threads = GmailApp.search(AI_RECEIVER_CONFIG.SUBJECT_QUERY, 0, 5);
  if (!threads || threads.length === 0) return null;

  let latestMsg = null;
  let latestTime = 0;

  for (const th of threads) {
    for (const m of th.getMessages()) {
      const t = m.getDate().getTime();
      if (t > latestTime) {
        latestTime = t;
        latestMsg = m;
      }
    }
  }
  return latestMsg;
}

function verifyAndDecryptPayload_(pkg, secretKey) {
  if (!pkg || !pkg.iv || !pkg.data || !pkg.sig) {
    throw new Error('暗号化パッケージの形式が不正です。');
  }

  const expectedSigBytes = Utilities.computeHmacSha256Signature(pkg.iv + ':' + pkg.data, secretKey);
  const expectedSigBase64 = Utilities.base64Encode(expectedSigBytes);
  if (expectedSigBase64 !== pkg.sig) {
    throw new Error('署名検証に失敗しました。暗号化キーが一致していません。');
  }

  const cipherBytes = Utilities.base64Decode(pkg.data);
  const gzipBytes = xorStreamCipherReceive_(cipherBytes, secretKey, pkg.iv);

  const gzipBlob = Utilities.newBlob(gzipBytes, 'application/x-gzip', 'payload.json.gz');
  const plainJsonStr = Utilities.ungzip(gzipBlob).getDataAsString('UTF-8');

  return JSON.parse(plainJsonStr);
}

function xorStreamCipherReceive_(inputBytes, secretKey, ivStr) {
  const output = new Array(inputBytes.length);
  let keyStream = [];
  let blockIndex = 0;

  for (let i = 0; i < inputBytes.length; i++) {
    if (i % 32 === 0) {
      keyStream = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        secretKey + '|' + ivStr + '|' + blockIndex,
        Utilities.Charset.UTF_8
      );
      blockIndex++;
    }
    output[i] = (inputBytes[i] ^ keyStream[i % 32]) & 0xff;
    if (output[i] > 127) output[i] -= 256;
  }
  return output;
}

function applyLeftAlignToSheet_(sheet, totalRows, totalCols) {
  if (!sheet || totalRows < 1 || totalCols < 1) return;
  sheet.getRange(1, 1, totalRows, totalCols)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
}

function writeUnfilledSheet_(ss, unfilledList, generatedAt, roleMasters) {
  let sheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_UNFILLED);
  if (!sheet) sheet = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_UNFILLED);

  const headers = [
    '日付', '拠点名', '不在時間', '掲載時給', '前の時間枠の医師', '後ろの時間枠の医師',
    'Phaze', '前医延長', '後医早出', '担当', '対応状況', '特別時給', '作業メモ', 'データ同期日時'
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fce8e6');
  sheet.setFrozenRows(1);

  if (unfilledList.length === 0) {
    applyLeftAlignToSheet_(sheet, 1, headers.length);
    return;
  }

  const rows = unfilledList.map(function(item) {
    return [
      item.date || '',
      item.clinic || '',
      item.absenceTime || '',
      item.listedWage || '',
      enrichDoctorTextWithSheetMasters_(item.prevDocs || '', roleMasters),
      enrichDoctorTextWithSheetMasters_(item.nextDocs || '', roleMasters),
      item.phase || '',
      item.prevExtStatus || '',
      item.nextEarlyStatus || '',
      item.assignee || '',
      item.status || '',
      item.specialWage || '',
      item.memo || '',
      generatedAt || ''
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  applyLeftAlignToSheet_(sheet, rows.length + 1, headers.length);
}

function writeActiveShiftsSheet_(ss, activeShiftsByDate, roleMasters) {
  let sheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_ACTIVE_SHIFTS);
  if (!sheet) sheet = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_ACTIVE_SHIFTS);

  const headers = [
    '日付', 'クリニックNo', '拠点名', '診療科', '午前人数(A)', '午後人数(B)', '夜間人数(C)',
    'Aシフト医師(09:00-13:00)', 'Bシフト医師(15:00-18:00)', 'Cシフト医師(18:00-21:00/20:00)'
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);

  const rows = [];
  const dates = Object.keys(activeShiftsByDate).sort();
  for (const d of dates) {
    const list = activeShiftsByDate[d] || [];
    for (const item of list) {
      rows.push([
        d,
        item.clinicNo || '',
        item.clinic || '',
        item.dept || '',
        item.countA || 0,
        item.countB || 0,
        item.countC || 0,
        enrichDoctorTextWithSheetMasters_(item.docsA || '', roleMasters),
        enrichDoctorTextWithSheetMasters_(item.docsB || '', roleMasters),
        enrichDoctorTextWithSheetMasters_(item.docsC || '', roleMasters)
      ]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  applyLeftAlignToSheet_(sheet, rows.length + 1, headers.length);
}

function writeOuboShiftsSheet_(ss, ouboShiftsByDate, roleMasters) {
  let sheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_OUBO_SHIFTS);
  if (!sheet) sheet = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_OUBO_SHIFTS);

  const headers = ['日付', '医籍番号', '応募医師名', 'クリニックNo', '応募拠点名', '診療科', '応募時間帯'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e6f4ea');
  sheet.setFrozenRows(1);

  const rows = [];
  const dates = Object.keys(ouboShiftsByDate).sort();
  for (const d of dates) {
    const list = ouboShiftsByDate[d] || [];
    for (const item of list) {
      rows.push([
        d,
        item.licenseNo || '',
        enrichDoctorTextWithSheetMasters_(item.doctor || '', roleMasters),
        item.clinicNo || '',
        item.clinic || '',
        item.dept || '',
        item.time || ''
      ]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  applyLeftAlignToSheet_(sheet, rows.length + 1, headers.length);
}

function writeStreakWarningsSheet_(ss, streakWarnings, roleMasters) {
  let sheet = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_STREAK_WARN);
  if (!sheet) sheet = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_STREAK_WARN);

  const headers = ['医籍番号', '医師名・区分', '連続勤務日数', '対象連勤期間（月またぎ含む）', '追加打診NG日（12連勤到達日）', 'シート登録メモ・AI注意事項'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fff2cc');
  sheet.setFrozenRows(1);

  const ftAllowMap = roleMasters.ftAllowMap;
  const sbuMap = roleMasters.sbuMap;

  const filtered = (streakWarnings || []).filter(function(w) {
    const docKey = normalizeDocKeyReceive_(w.doctor);
    const isFtAllowDoc = ftAllowMap.has(docKey);
    return isFtAllowDoc || w.streakDays >= 10 || w.type === 'BRIDGE';
  });

  if (filtered.length === 0) {
    applyLeftAlignToSheet_(sheet, 1, headers.length);
    return;
  }

  const rows = filtered.map(function(w) {
    const docKey = normalizeDocKeyReceive_(w.doctor);
    let label = w.doctor;
    let memoNote = w.note || '';

    if (ftAllowMap.has(docKey)) {
      const m = ftAllowMap.get(docKey).memo;
      label += '(常勤依頼手当対象)';
      if (m) memoNote = '【条件:' + m + '】 ' + memoNote;
    } else if (sbuMap.has(docKey)) {
      label += '(★SBU)';
    } else if (w.isFullTime) {
      label += '(常勤)';
    }

    return [
      w.licenseNo || '',
      label,
      w.streakDays || 0,
      w.period || '',
      w.ngDates || '',
      memoNote
    ];
  });

  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  applyLeftAlignToSheet_(sheet, rows.length + 1, headers.length);
}

/**
 * 受信したペイロード内の「勤務くん個別打診候補」と「複数日程おまとめ候補」に
 * 『交渉しやすい医師』『交渉注意医師』『ナレッジシート（直近打診履歴）』の情報を付加して展開する
 */
function writeKinmukunCandidateSheets_(ss, payload, roleMasters) {
  const negotiableMap = roleMasters.negotiableMap;
  const cautionMap = roleMasters.cautionMap;
  const recentKbMap = roleMasters.recentKbMap;

  const buildMasterNote = function(rawDocName) {
    const docKey = normalizeDocKeyReceive_(rawDocName);
    if (!docKey) return '';
    const notes = [];
    if (negotiableMap && negotiableMap.has(docKey)) {
      const n = negotiableMap.get(docKey);
      const detail = [n.route, n.humanMemo, n.aiMemo].filter(Boolean).join('/');
      notes.push('★交渉しやすい' + (detail ? '[' + detail + ']' : ''));
    }
    if (cautionMap && cautionMap.has(docKey)) {
      notes.push('⚠️交渉注意[' + cautionMap.get(docKey) + ']');
    }
    if (recentKbMap && recentKbMap.has(docKey)) {
      notes.push('📩' + recentKbMap.get(docKey));
    }
    return notes.join(' / ');
  };

  // 1. 『最新_個別打診候補』シート
  const spotRaw = payload.spotCandidateRows || [];
  if (spotRaw.length > 0) {
    let sh = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_SPOT_CANDIDATES);
    if (!sh) sh = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_SPOT_CANDIDATES);
    sh.clearContents();

    const headers = spotRaw[0].concat(['マスタ・NG条件・直近打診履歴']);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
    sh.setFrozenRows(1);

    if (spotRaw.length > 1) {
      const bodyRows = spotRaw.slice(1).map(function(r) {
        return r.concat([buildMasterNote(r[5])]);
      });
      sh.getRange(2, 7, bodyRows.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, bodyRows.length, headers.length).setValues(bodyRows);
      applyLeftAlignToSheet_(sh, bodyRows.length + 1, headers.length);
    } else {
      applyLeftAlignToSheet_(sh, 1, headers.length);
    }
  }

  // 2. 『最新_おまとめ打診候補』シート
  const bundleRaw = payload.bundledCandidateRows || [];
  if (bundleRaw.length > 0) {
    let sh = ss.getSheetByName(AI_RECEIVER_CONFIG.SHEET_BUNDLED_CANDIDATES);
    if (!sh) sh = ss.insertSheet(AI_RECEIVER_CONFIG.SHEET_BUNDLED_CANDIDATES);
    sh.clearContents();

    const headers = bundleRaw[0].concat(['マスタ・NG条件・直近打診履歴']);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#cfe2f3');
    sh.setFrozenRows(1);

    if (bundleRaw.length > 1) {
      const bodyRows = bundleRaw.slice(1).map(function(r) {
        return r.concat([buildMasterNote(r[0])]);
      });
      sh.getRange(2, 2, bodyRows.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, bodyRows.length, headers.length).setValues(bodyRows);
      applyLeftAlignToSheet_(sh, bodyRows.length + 1, headers.length);
    } else {
      applyLeftAlignToSheet_(sh, 1, headers.length);
    }
  }
}