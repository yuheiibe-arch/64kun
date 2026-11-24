/**
 * ============================================================================
 * 6c_1_編集くん64_一時保存管理.gs（6c系 3分割の1）
 * ・シート非アクセスの超高速ドラフト取得（loadLatestDraftFastNoSheet64_）
 * ・毎回全消去・毎回書式設定を廃止した軽量バッチ書込（スプレッドシートタイムアウト防止）
 * ============================================================================
 */

const HENSHUKUN_DRAFT_CACHE_PREFIX_64_ = 'HENSHUKUN_DRAFT_CACHE_V1_';

function saveDraftToPropsSafely64_(props, key, draftObj, docIdOpt) {
  const jsonStr = JSON.stringify(draftObj);
  const docId = docIdOpt || String(key || '').replace(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX, '');

  if (docId) {
    const cache = CacheService.getScriptCache();
    const cacheKey = HENSHUKUN_DRAFT_CACHE_PREFIX_64_ + docId;
    if (jsonStr.length <= 95000) {
      try { cache.put(cacheKey, jsonStr, 21600); } catch (e) {}
    } else {
      try { cache.remove(cacheKey); } catch (e2) {}
    }
  }

  if (jsonStr.length <= 8000) {
    try {
      props.setProperty(key, jsonStr);
      return;
    } catch (e) {}
  }
  try {
    props.deleteProperty(key);
  } catch (e2) {}
}

/**
 * SpreadsheetApp を一切開かずに CacheService / ScriptProperties のみから高速取得する（ハートビート用）
 */
function loadLatestDraftFastNoSheet64_(docId) {
  if (!docId) return null;
  const cache = CacheService.getScriptCache();
  const cachedRaw = cache.get(HENSHUKUN_DRAFT_CACHE_PREFIX_64_ + docId);
  if (cachedRaw) {
    try { return JSON.parse(cachedRaw); } catch (e) {}
  }
  const propRaw = PropertiesService.getScriptProperties().getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId);
  if (propRaw) {
    try { return JSON.parse(propRaw); } catch (e) {}
  }
  return null;
}

function loadLatestDraft64_(ss, docId) {
  if (!docId) return { slots: {}, bundles: {} };

  const fastHit = loadLatestDraftFastNoSheet64_(docId);
  if (fastHit) return fastHit;

  const targetSs = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = targetSs ? targetSs.getSheetByName(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME) : null;
  if (sh && sh.getLastRow() >= 2) {
    const maxCol = Math.max(5, sh.getLastColumn());
    const row2 = sh.getRange(2, 1, 1, maxCol).getDisplayValues()[0];
    const savedDocId = String(row2[1] || '').trim();
    const rawJson = row2.slice(4).join('').trim();
    if (savedDocId === docId && rawJson) {
      try {
        const parsedSheet = JSON.parse(rawJson);
        if (rawJson.length <= 95000) {
          try { CacheService.getScriptCache().put(HENSHUKUN_DRAFT_CACHE_PREFIX_64_ + docId, rawJson, 21600); } catch (e2) {}
        }
        return parsedSheet;
      } catch (e) {}
    }
  }
  return { slots: {}, bundles: {} };
}

/**
 * メモリ上で payload の変更分を draft オブジェクトへマージする共通ヘルパー
 */
function mergePayloadIntoDraftInMemory64_(draft, payload, timeStr) {
  if (!draft.slots) draft.slots = {};
  if (!draft.bundles) draft.bundles = {};

  Object.entries((payload && payload.changedSlots) || {}).forEach(function([slotKey, val]) {
    const cleanDec = String(val.decision || '')
      .split(/\r?\n/)
      .filter(function(l) { return !l.trim().startsWith('→'); })
      .join(' ')
      .trim();

    draft.slots[slotKey] = {
      decision: cleanDec,
      assignee: val.assignee,
      customComment: String(val.customComment || '').trim(),
      customDocs: Array.isArray(val.customDocs) ? val.customDocs : [],
      dateLabel: val.dateLabel || '',
      clinicTitle: val.clinicTitle || '',
      timeAndWage: val.timeAndWage || '',
      updatedAt: Date.now(),
      updatedTimeStr: timeStr
    };
  });

  Object.entries((payload && payload.changedBundles) || {}).forEach(function([bKey, val]) {
    draft.bundles[bKey] = val;
  });

  return draft;
}

/**
 * ★高速化版：重い書式API（setBackground等）は初回作成時のみ実行し、普段は最小の setValues のみで高速書き込みする
 */
function writeDraftToSheet64_(ss, docId, draftObj, playerStaff) {
  let sh = ss.getSheetByName(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  let isNewSheet = false;
  if (!sh) {
    sh = ss.insertSheet(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
    isNewSheet = true;
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const jsonStr = JSON.stringify(draftObj);
  const slotEntries = Object.entries(draftObj.slots || {});

  const CHUNK_SIZE = 40000;
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    chunks.push(jsonStr.slice(i, i + CHUNK_SIZE));
  }
  if (chunks.length === 0) chunks.push('{}');

  const totalCols = Math.max(8, 4 + chunks.length);
  if (sh.getMaxColumns() < totalCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), totalCols - sh.getMaxColumns());
  }

  const padRow = function(arr) {
    const copy = arr.slice();
    while (copy.length < totalCols) copy.push('');
    return copy;
  };

  const metaHeaders = ['最終一時保存日時', '対象ドキュメントID', '最終保存プレイヤー', '一時保存枠数', '一時保存JSONデータ（システム同期・医籍番号連携用）'];
  for (let c = 1; c < chunks.length; c++) {
    metaHeaders.push('一時保存JSON（分割' + (c + 1) + '）');
  }
  const metaRow = [nowStr, docId, playerStaff || '未選択', slotEntries.length].concat(chunks);
  const listHeaders = ['スロットキー', '日付', '拠点', '確定案（一時保存）', '特定済カスタム医師（医籍番号）', '担当者名（一時保存）', 'カスタムコメント・メモ', '更新時刻'];

  const grid = [
    padRow(metaHeaders),
    padRow(metaRow),
    padRow([]),
    padRow(listHeaders)
  ];

  slotEntries.forEach(function([k, v]) {
    const cDocsText = Array.isArray(v.customDocs)
      ? v.customDocs.map(function(d) { return d.name + '(' + (d.licenseNo || '番号なし') + '):' + (d.actionType || ''); }).join(' / ')
      : '';
    grid.push(padRow([
      k,
      v.dateLabel || '',
      v.clinicTitle || '',
      String(v.decision || '').slice(0, 45000),
      cDocsText.slice(0, 45000),
      v.assignee || '',
      String(v.customComment || '').slice(0, 45000),
      v.updatedTimeStr || nowStr
    ]));
  });

  const prevLastRow = sh.getLastRow();
  const prevLastCol = sh.getLastColumn();

  // たった1回の setValues で一括書き込み（APIコール激減）
  sh.getRange(1, 1, grid.length, totalCols).setValues(grid);

  // 前回より行数や列数が減った場合のみ余剰範囲の値だけクリア
  if (prevLastRow > grid.length) {
    sh.getRange(grid.length + 1, 1, prevLastRow - grid.length, Math.max(prevLastCol, totalCols)).clearContent();
  }
  if (prevLastCol > totalCols && prevLastRow > 0) {
    sh.getRange(1, totalCols + 1, Math.max(prevLastRow, grid.length), prevLastCol - totalCols).clearContent();
  }

  if (isNewSheet || prevLastRow === 0) {
    sh.getRange(1, 1, 1, metaHeaders.length).setBackground('#ffc72c').setFontWeight('bold');
    sh.getRange(4, 1, 1, 8).setBackground('#0046ad').setFontColor('#ffffff').setFontWeight('bold');
  }
}

function saveHenshukunDraft(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const docId = props.getProperty('SINGLE_MTG_DOC_ID');
    if (!docId) throw new Error('ドキュメントIDが見つかりません。');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const key = HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId;
    const draft = loadLatestDraft64_(ss, docId);
    const timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');

    mergePayloadIntoDraftInMemory64_(draft, payload, timeStr);

    // まず高速な Cache / Properties に即保存してから、軽量1回バッチでシートへ保存
    saveDraftToPropsSafely64_(props, key, draft, docId);
    writeDraftToSheet64_(ss, docId, draft, (payload && payload.playerStaff) || '');

    return { ok: true, savedAt: timeStr };
  } finally {
    lock.releaseLock();
  }
}