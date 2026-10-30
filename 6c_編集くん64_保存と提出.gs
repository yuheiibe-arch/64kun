/**
 * ============================================================================
 * 6c_編集くん64_保存と提出.gs
 * （『編集くん_一時保存』シートへのJSON保存・ドキュメント反映・mnys提出連携）
 * ============================================================================
 */

function loadLatestDraft64_(ss, docId) {
  const props = PropertiesService.getScriptProperties();
  const propRaw = props.getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId);
  if (propRaw) {
    try { return JSON.parse(propRaw); } catch (e) {}
  }

  const sh = ss.getSheetByName(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  if (sh && sh.getLastRow() >= 2) {
    const savedDocId = String(sh.getRange(2, 2).getDisplayValue() || '').trim();
    const rawJson = String(sh.getRange(2, 5).getDisplayValue() || '').trim();
    if (savedDocId === docId && rawJson) {
      try { return JSON.parse(rawJson); } catch (e) {}
    }
  }
  return { slots: {}, bundles: {} };
}

function writeDraftToSheet64_(ss, docId, draftObj, playerStaff) {
  let sh = ss.getSheetByName(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(HENSHUKUN_CONFIG.DRAFT_SHEET_NAME);
  sh.clearContents();

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const jsonStr = JSON.stringify(draftObj);
  const slotEntries = Object.entries(draftObj.slots || {});

  const metaHeaders = ['最終一時保存日時', '対象ドキュメントID', '最終保存プレイヤー', '一時保存枠数', '一時保存JSONデータ（システム同期・医籍番号連携用）'];
  sh.getRange(1, 1, 1, 5).setValues([metaHeaders]).setBackground('#ffc72c').setFontWeight('bold');
  sh.getRange(2, 1, 1, 5).setValues([[nowStr, docId, playerStaff || '未選択', slotEntries.length, jsonStr]]);

  const listHeaders = ['スロットキー', '日付', '拠点', '確定案（一時保存）', '特定済カスタム医師（医籍番号）', '担当者名（一時保存）', '更新時刻'];
  sh.getRange(4, 1, 1, 7).setValues([listHeaders]).setBackground('#0046ad').setFontColor('#ffffff').setFontWeight('bold');

  if (slotEntries.length > 0) {
    const rows = slotEntries.map(([k, v]) => {
      const cDocsText = Array.isArray(v.customDocs)
        ? v.customDocs.map(d => `${d.name}(${d.licenseNo || '番号なし'}):${d.actionType || ''}`).join(' / ')
        : '';
      return [
        k,
        v.dateLabel || '',
        v.clinicTitle || '',
        v.decision || '',
        cDocsText,
        v.assignee || '',
        v.updatedTimeStr || nowStr
      ];
    });
    sh.getRange(5, 1, rows.length, 7).setValues(rows);
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

    Object.entries(payload.changedSlots || {}).forEach(([slotKey, val]) => {
      draft.slots[slotKey] = {
        decision: val.decision,
        assignee: val.assignee,
        customDocs: Array.isArray(val.customDocs) ? val.customDocs : [],
        dateLabel: val.dateLabel || '',
        clinicTitle: val.clinicTitle || '',
        updatedAt: Date.now(),
        updatedTimeStr: timeStr
      };
    });

    Object.entries(payload.changedBundles || {}).forEach(([bKey, val]) => {
      draft.bundles[bKey] = val;
    });

    props.setProperty(key, JSON.stringify(draft));
    writeDraftToSheet64_(ss, docId, draft, payload.playerStaff || '');

    return { ok: true, savedAt: timeStr };
  } finally {
    lock.releaseLock();
  }
}

function commitHenshukunToDocAndSubmit(payload, shouldSubmitToMnys) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    saveHenshukunDraft(payload);

    const props = PropertiesService.getScriptProperties();
    const docId = props.getProperty('SINGLE_MTG_DOC_ID');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const draft = loadLatestDraft64_(ss, docId);

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    Object.entries(draft.slots || {}).forEach(([slotKey, item]) => {
      const [cIdx, rIdx] = slotKey.split('_').map(Number);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            const rawC0 = cleanSidebarDisplayText_(row.getCell(0).getText());
            const c0Lines = rawC0.split('\n');
            c0Lines[0] = formatClinicTitle64_(c0Lines[0]);
            const c1 = cleanSidebarDisplayText_(row.getCell(1).getText());

            row.getCell(0).setText(c0Lines.join('\n'));
            row.getCell(1).setText(c1);
            row.getCell(2).setText(String(item.decision || '').trim());
            row.getCell(3).setText(String(item.assignee || '').trim());
            row.getCell(0).editAsText().setFontSize(8.5);
            row.getCell(1).editAsText().setFontSize(8.5);
            row.getCell(2).editAsText().setFontSize(8.5).setBold(true);
            row.getCell(3).editAsText().setFontSize(8.5);
          }
        }
      }
    });

    Object.entries(draft.bundles || {}).forEach(([bKey, val]) => {
      const parts = bKey.split('_');
      const cIdx = Number(parts[1]), rIdx = Number(parts[2]);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            row.getCell(0).setText(cleanSidebarDisplayText_(row.getCell(0).getText()));
            row.getCell(3).setText(String(val || '').trim());
            row.getCell(0).editAsText().setFontSize(8.5);
            row.getCell(3).editAsText().setFontSize(8.5).setBold(true);
          }
        }
      }
    });

    doc.saveAndClose();

    if (shouldSubmitToMnys) {
      submitEditedDocToMnys();
    }

    return {
      ok: true,
      committedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss')
    };
  } finally {
    lock.releaseLock();
  }
}