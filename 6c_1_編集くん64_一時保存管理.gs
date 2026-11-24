/**
 * ============================================================================
 * 6c_1_編集くん64_一時保存管理.gs（6c系 3分割の1）
 * ・CacheService（100KB高速同期）＋ ScriptProperties（8KB軽量時）＋ 分割シート保存
 * ・『編集くん_一時保存』シートへの40,000文字チャンク分割書込／読込・saveHenshukunDraft
 * ============================================================================
 */

const HENSHUKUN_DRAFT_CACHE_PREFIX_64_ = 'HENSHUKUN_DRAFT_CACHE_V1_';

/**
 * CacheService（上限100KB・同時編集の高速同期用）および ScriptProperties（8KB以下の軽量時用）へ安全に保存する
 */
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
 * CacheService（100KB）➔ ScriptProperties（8KB）➔ 『編集くん_一時保存』分割シートの順で高速＆確実に読み込む
 */
function loadLatestDraft64_(ss, docId) {
  if (!docId) return { slots: {}, bundles: {} };

  const cache = CacheService.getScriptCache();
  const cacheKey = HENSHUKUN_DRAFT_CACHE_PREFIX_64_ + docId;
  const cachedRaw = cache.get(cacheKey);
  if (cachedRaw) {
    try { return JSON.parse(cachedRaw); } catch (e) {}
  }

  const props = PropertiesService.getScriptProperties();
  const propRaw = props.getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId);
  if (propRaw) {
    try {
      const parsedProp = JSON.parse(propRaw);
      if (propRaw.length <= 95000) {
        try { cache.put(cacheKey, propRaw, 21600); } catch (e2) {}
      }
      return parsedProp;
    } catch (e) {}
  }

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
          try { cache.put(cacheKey, rawJson, 21600); } catch (e2) {}
        }
        return parsedSheet;
      } catch (e) {}
    }
  }
  return { slots: {}, bundles: {} };
}

/**
 * JSON文字列が1セル上限（50,000文字）を超えないよう、40,000文字単位で E2, F2, G2... へ自動分割して書き込む
 */
function writeDraftToSheet64_(ss, docId, draftObj, playerStaff) {
  let sh = ss.getSheetByName(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  sh.clearContents();

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const jsonStr = JSON.stringify(draftObj);
  const slotEntries = Object.entries(draftObj.slots || {});

  const CHUNK_SIZE = 40000;
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    chunks.push(jsonStr.slice(i, i + CHUNK_SIZE));
  }
  if (chunks.length === 0) chunks.push('{}');

  const metaHeaders = ['最終一時保存日時', '対象ドキュメントID', '最終保存プレイヤー', '一時保存枠数', '一時保存JSONデータ（システム同期・医籍番号連携用）'];
  for (let c = 1; c < chunks.length; c++) {
    metaHeaders.push('一時保存JSON（分割' + (c + 1) + '）');
  }

  const metaRow = [nowStr, docId, playerStaff || '未選択', slotEntries.length].concat(chunks);
  if (sh.getMaxColumns() < metaRow.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), metaRow.length - sh.getMaxColumns());
  }

  sh.getRange(1, 1, 1, metaHeaders.length).setValues([metaHeaders]).setBackground('#ffc72c').setFontWeight('bold');
  sh.getRange(2, 1, 1, metaRow.length).setValues([metaRow]);

  const listHeaders = ['スロットキー', '日付', '拠点', '確定案（一時保存）', '特定済カスタム医師（医籍番号）', '担当者名（一時保存）', 'カスタムコメント・メモ', '更新時刻'];
  if (sh.getMaxColumns() < 8) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());
  }
  sh.getRange(4, 1, 1, 8).setValues([listHeaders]).setBackground('#0046ad').setFontColor('#ffffff').setFontWeight('bold');

  if (slotEntries.length > 0) {
    const rows = slotEntries.map(function([k, v]) {
      const cDocsText = Array.isArray(v.customDocs)
        ? v.customDocs.map(function(d) { return d.name + '(' + (d.licenseNo || '番号なし') + '):' + (d.actionType || ''); }).join(' / ')
        : '';
      return [
        k,
        v.dateLabel || '',
        v.clinicTitle || '',
        String(v.decision || '').slice(0, 45000),
        cDocsText.slice(0, 45000),
        v.assignee || '',
        String(v.customComment || '').slice(0, 45000),
        v.updatedTimeStr || nowStr
      ];
    });
    sh.getRange(5, 1, rows.length, 8).setValues(rows);
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

    Object.entries(payload.changedSlots || {}).forEach(function([slotKey, val]) {
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

    Object.entries(payload.changedBundles || {}).forEach(function([bKey, val]) {
      draft.bundles[bKey] = val;
    });

    writeDraftToSheet64_(ss, docId, draft, payload.playerStaff || '');
    saveDraftToPropsSafely64_(props, key, draft, docId);

    return { ok: true, savedAt: timeStr };
  } finally {
    lock.releaseLock();
  }
}