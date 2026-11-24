/**
 * ============================================================================
 * 6c_3_編集くん64_ドキュメント反映と提出.gs（6c系 3分割の3）
 * ・実際に変更があったセルのみを更新する「差分更新方式」（反映時間を約1/10に短縮）
 * ・ドキュメントを saveAndClose() した後にシート保存を1回だけ実行しタイムアウトを根絶
 * ============================================================================
 */

function commitHenshukunToDocAndSubmit(payload, shouldSubmitToMnys) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const docId = props.getProperty('SINGLE_MTG_DOC_ID');
    if (!docId) throw new Error('ドキュメントIDが見つかりません。');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const key = HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId;
    const timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');

    // 1. メモリ上でドラフトを統合し、まず高速キャッシュへ先行保存（シート書込はまだ行わない）
    const draft = loadLatestDraft64_(ss, docId);
    mergePayloadIntoDraftInMemory64_(draft, payload, timeStr);
    saveDraftToPropsSafely64_(props, key, draft, docId);

    // 2. Googleドキュメントを開き、行インデックスを構築
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
    let modifiedRowCount = 0;

    // 3. 既存のセル内容と実際に差分がある行のみ DOM 更新＆マーカー描画を実行（劇的高速化）
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
              const timeOk = !itemTimeRange || !candTimeRange || isTimeRangeOverlapping64_(candTimeRange, itemTimeRange);
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
          ? cands.filter(function(c) { return isTimeRangeOverlapping64_(c.timeRange, itemTimeRange); })
          : cands;
        const resolvedCand = timeMatched.length === 1 ? timeMatched[0] : (cands.length === 1 ? cands[0] : null);
        if (resolvedCand) {
          targetRow = resolvedCand.row;
          actualSlotKey = resolvedCand.cIdx + '_' + resolvedCand.rIdx;
        }
      }

      if (!targetRow) return;
      updatedSlotMap[actualSlotKey] = item;

      const curCell1Raw = targetRow.getCell(1).getText().replace(/\r\n|\r/g, '\n').trim();
      const curCell2Raw = targetRow.getCell(2).getText().replace(/\r\n|\r/g, '\n').trim();
      const curCell3Raw = targetRow.getCell(3).getText().replace(/\r\n|\r/g, '\n').trim();

      let desiredC1 = dedupePlanCellLines64_(cleanSidebarDisplayText_(curCell1Raw))
        .replace(/\n*📝コメント[：:].*$/g, '')
        .trim();
      const cmnt = String(item.customComment || '').trim();
      if (cmnt) {
        desiredC1 += '\n📝コメント：' + cmnt;
      }

      const arrowLines = curCell2Raw.split('\n').filter(function(l) { return l.trim().startsWith('→'); });
      let baseDec = String(item.decision || '')
        .split(/\r?\n/)
        .filter(function(l) { return !l.trim().startsWith('→'); })
        .join(' ')
        .trim();
      if (baseDec === 'カスタム入力(医師未選択)') {
        baseDec = '';
      }

      const headerDecText = (cmnt && baseDec) ? (baseDec + '（メモ：' + cmnt + '）') : baseDec;
      const desiredC2 = [headerDecText].concat(arrowLines).filter(Boolean).join('\n');
      const desiredC3 = String(item.assignee || '').trim();

      const wasExplicitlyEditedNow = Boolean(payload && payload.changedSlots && payload.changedSlots[slotKey]);
      const isSameAsDoc = (curCell2Raw === desiredC2) && (curCell3Raw === desiredC3) && (curCell1Raw === desiredC1);

      // 既にドキュメント上のテキストと完全一致しており今回の変更対象でもない行は重いDOM操作をスキップ
      if (isSameAsDoc && !wasExplicitlyEditedNow) {
        return;
      }

      modifiedRowCount++;
      const rawC0 = cleanSidebarDisplayText_(targetRow.getCell(0).getText());
      const c0Lines = rawC0.split('\n');
      c0Lines[0] = formatClinicTitle64_(c0Lines[0]);

      setDocCellTextSafely64_(targetRow.getCell(0), c0Lines.join('\n'));
      setDocCellTextSafely64_(targetRow.getCell(2), desiredC2);
      setDocCellTextSafely64_(targetRow.getCell(3), desiredC3);

      targetRow.getCell(0).editAsText().setFontSize(8.5);
      targetRow.getCell(2).editAsText().setFontSize(8.5).setBold(true);
      targetRow.getCell(3).editAsText().setFontSize(8.5);

      applySelectedPlanMarkerToDocCell64_(targetRow.getCell(1), desiredC1, desiredC2);
    });

    Object.entries(draft.bundles || {}).forEach(function([bKey, val]) {
      const parts = bKey.split('_');
      const cIdx = Number(parts[1]), rIdx = Number(parts[2]);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            const desiredVal = String(val || '').trim();
            if (row.getCell(3).getText().trim() !== desiredVal) {
              modifiedRowCount++;
              setDocCellTextSafely64_(row.getCell(0), cleanSidebarDisplayText_(row.getCell(0).getText()));
              setDocCellTextSafely64_(row.getCell(3), desiredVal);
              row.getCell(0).editAsText().setFontSize(8.5);
              row.getCell(3).editAsText().setFontSize(8.5).setBold(true);
            }
          }
        }
      }
    });

    // 4. 先にドキュメントを保存してクローズ（DocumentAppセッションを完全に解放）
    doc.saveAndClose();

    // 5. ドキュメント解放後に1回だけ一時保存シートを更新
    draft.slots = updatedSlotMap;
    saveDraftToPropsSafely64_(props, key, draft, docId);
    writeDraftToSheet64_(ss, docId, draft, (payload && payload.playerStaff) || '');

    if (shouldSubmitToMnys) {
      submitEditedDocToMnys();
    }

    return {
      ok: true,
      modifiedRowCount: modifiedRowCount,
      committedAt: timeStr
    };
  } finally {
    lock.releaseLock();
  }
}