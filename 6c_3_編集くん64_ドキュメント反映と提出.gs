/**
 * ============================================================================
 * 6c_3_編集くん64_ドキュメント反映と提出.gs（6c系 3分割の3）
 * ・時間帯一致を含めた行ズレ完全防止ルックアップ
 * ・Googleドキュメント本体への一括反映 ＆ mnys（超募集くん）への提出トリガー
 * ============================================================================
 */

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
    const numChildren = body.getNumChildren();

    const liveRowLookup = new Map();
    let curSortDate = '';
    for (let i = 0; i < numChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const dm = el.asParagraph().getText().trim().match(/^◆\s*(\d{4}[\/\-年]\d{1,2}[\/\-月]\d{1,2})/);
        if (dm) curSortDate = toSortDateKey64_(dm[1]);
      } else if (el.getType() === DocumentApp.ElementType.TABLE && curSortDate) {
        const tbl = el.asTable();
        if (tbl.getNumRows() < 2 || !tbl.getRow(0).getCell(0).getText().includes('拠点・不在時間帯')) continue;
        for (let r = 1; r < tbl.getNumRows(); r++) {
          const row = tbl.getRow(r);
          if (row.getNumCells() < 4) continue;
          const c0Lines = cleanSidebarDisplayText_(row.getCell(0).getText()).split('\n');
          const cTitle = formatClinicTitle64_(c0Lines[0] || '');
          const cTimeRange = extractSlotTimeRange64_(c0Lines.slice(1).join(' '));
          const lookupKey = curSortDate + '|' + cTitle;
          if (!liveRowLookup.has(lookupKey)) liveRowLookup.set(lookupKey, []);
          liveRowLookup.get(lookupKey).push({ cIdx: i, rIdx: r, row: row, timeRange: cTimeRange });
        }
      }
    }

    const updatedSlotMap = {};

    Object.entries(draft.slots || {}).forEach(function([slotKey, item]) {
      const [cIdx, rIdx] = slotKey.split('_').map(Number);
      let targetRow = null;
      let actualSlotKey = slotKey;
      const itemTimeRange = extractSlotTimeRange64_(item.timeAndWage || '');

      if (cIdx < numChildren) {
        const el = body.getChild(cIdx);
        if (el && el.getType() === DocumentApp.ElementType.TABLE) {
          const tbl = el.asTable();
          if (rIdx < tbl.getNumRows()) {
            const candRow = tbl.getRow(rIdx);
            if (candRow.getNumCells() >= 4) {
              const candLines = cleanSidebarDisplayText_(candRow.getCell(0).getText()).split('\n');
              const candClinic = formatClinicTitle64_(candLines[0] || '');
              const candTimeRange = extractSlotTimeRange64_(candLines.slice(1).join(' '));
              const clinicOk = !item.clinicTitle || candClinic === item.clinicTitle;
              const timeOk = !itemTimeRange || !candTimeRange || candTimeRange === itemTimeRange;
              if (clinicOk && timeOk) {
                targetRow = candRow;
              }
            }
          }
        }
      }

      if (!targetRow && item.dateLabel && item.clinicTitle) {
        const lKey = toSortDateKey64_(item.dateLabel) + '|' + item.clinicTitle;
        const cands = liveRowLookup.get(lKey) || [];
        const timeMatched = itemTimeRange
          ? cands.filter(function(c) { return c.timeRange === itemTimeRange; })
          : cands;
        const resolvedCand = timeMatched.length === 1 ? timeMatched[0] : (cands.length === 1 ? cands[0] : null);
        if (resolvedCand) {
          targetRow = resolvedCand.row;
          actualSlotKey = resolvedCand.cIdx + '_' + resolvedCand.rIdx;
        }
      }

      if (!targetRow) return;
      updatedSlotMap[actualSlotKey] = item;

      const rawC0 = cleanSidebarDisplayText_(targetRow.getCell(0).getText());
      const c0Lines = rawC0.split('\n');
      c0Lines[0] = formatClinicTitle64_(c0Lines[0]);

      let c1 = dedupePlanCellLines64_(cleanSidebarDisplayText_(targetRow.getCell(1).getText()))
        .replace(/\n*📝コメント[：:].*$/g, '')
        .trim();
      const cmnt = String(item.customComment || '').trim();
      if (cmnt) {
        c1 += '\n📝コメント：' + cmnt;
      }

      const existingC2Lines = targetRow.getCell(2).getText().split(/\r?\n/);
      const arrowLines = existingC2Lines.filter(function(l) { return l.trim().startsWith('→'); });

      let baseDec = String(item.decision || '')
        .split(/\r?\n/)
        .filter(function(l) { return !l.trim().startsWith('→'); })
        .join(' ')
        .trim();
      if (baseDec === 'カスタム入力(医師未選択)') {
        baseDec = '';
      }

      const headerDecText = (cmnt && baseDec) ? (baseDec + '（メモ：' + cmnt + '）') : baseDec;
      const fullDecText = [headerDecText].concat(arrowLines).filter(Boolean).join('\n');

      setDocCellTextSafely64_(targetRow.getCell(0), c0Lines.join('\n'));
      setDocCellTextSafely64_(targetRow.getCell(2), fullDecText);
      setDocCellTextSafely64_(targetRow.getCell(3), String(item.assignee || '').trim());

      targetRow.getCell(0).editAsText().setFontSize(8.5);
      targetRow.getCell(2).editAsText().setFontSize(8.5).setBold(true);
      targetRow.getCell(3).editAsText().setFontSize(8.5);

      applySelectedPlanMarkerToDocCell64_(targetRow.getCell(1), c1, fullDecText);
    });

    for (let i = 0; i < numChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
      const table = el.asTable();
      if (table.getNumRows() < 2 || !table.getRow(0).getCell(0).getText().includes('拠点・不在時間帯')) continue;

      for (let r = 1; r < table.getNumRows(); r++) {
        const sKey = i + '_' + r;
        if (updatedSlotMap[sKey]) continue;
        const row = table.getRow(r);
        if (row.getNumCells() < 4) continue;
        const c1Text = dedupePlanCellLines64_(cleanSidebarDisplayText_(row.getCell(1).getText().trim()));
        const decText = row.getCell(2).getText().trim();
        applySelectedPlanMarkerToDocCell64_(row.getCell(1), c1Text, decText);
      }
    }

    draft.slots = updatedSlotMap;
    writeDraftToSheet64_(ss, docId, draft, payload.playerStaff || '');
    saveDraftToPropsSafely64_(props, HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId, draft, docId);

    Object.entries(draft.bundles || {}).forEach(function([bKey, val]) {
      const parts = bKey.split('_');
      const cIdx = Number(parts[1]), rIdx = Number(parts[2]);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            setDocCellTextSafely64_(row.getCell(0), cleanSidebarDisplayText_(row.getCell(0).getText()));
            setDocCellTextSafely64_(row.getCell(3), String(val || '').trim());
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