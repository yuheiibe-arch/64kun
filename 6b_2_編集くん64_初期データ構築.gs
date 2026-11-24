/**
 * ============================================================================
 * 6b_2_編集くん64_初期データ構築.gs（6b系 3分割の2）
 * ・Googleドキュメント走査・募集枠時間反映・提出済み台帳照合・一時保存ドラフト突合
 * ・地域順ソート（東京➔神奈川➔千葉➔茨城➔埼玉➔大阪）＆初期データ返却（getHenshukunInitialData）
 * ============================================================================
 */

function getHenshukunInitialData() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    throw new Error('対象ドキュメント（SINGLE_MTG_DOC_ID）が見つかりません。先にドキュメントを生成してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masters = loadClinicAndStaffMasters64_(ss);
  const bkAndRecruit = extractBackupAndRecruitSlots64_(ss, masters.doctorSearchList);
  const backupDocsByDate = bkAndRecruit.backupDocsByDate;
  const recruitSlotsMap = bkAndRecruit.recruitSlotsMap;

  const submittedLedgerMap = (typeof loadSubmittedLedgerMap4_3_ === 'function')
    ? loadSubmittedLedgerMap4_3_(ss)
    : new Map();

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  const draftMap = loadLatestDraft64_(ss, docId);
  const allDraftSlotEntries = Object.values(draftMap.slots || {});

  let currentDateRaw = '';
  const slotRows = [];
  const bundleRows = [];

  for (let i = 0; i < numChildren; i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const txt = el.asParagraph().getText().trim();
      const m = txt.match(/^◆\s*(\d{4}[\/\-年]\d{1,2}[\/\-月]\d{1,2}.*)$/);
      if (m) currentDateRaw = m[1].trim();
    } else if (el.getType() === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      if (table.getNumRows() < 2) continue;
      const h0 = table.getRow(0).getCell(0).getText().trim();

      if (h0.includes('拠点・不在時間帯') && currentDateRaw) {
        const sortDate = toSortDateKey64_(currentDateRaw);
        const jpDate = toJapaneseDateLabel64_(currentDateRaw);

        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;

          const rawCell0 = row.getCell(0).getText().replace(/\r\n|\r/g, '\n').trim();
          const rawCell1 = row.getCell(1).getText().replace(/\r\n|\r/g, '\n')
            .replace(/([^\n])\s*(?=【案[123１２３]】)/g, '$1\n')
            .replace(/([^\n])\s*(?=💡)/g, '$1\n')
            .replace(/([^\n])\s*(?=・【進行中メモ】|・⚠️)/g, '$1\n')
            .trim();

          const col1 = cleanSidebarDisplayText_(rawCell0);
          const rawCol2Cleaned = cleanSidebarDisplayText_(rawCell1);
          let col2 = (typeof dedupePlanCellLines64_ === 'function')
            ? dedupePlanCellLines64_(rawCol2Cleaned)
            : rawCol2Cleaned;

          const rawDocDecFull = row.getCell(2).getText().replace(/\r\n|\r/g, '\n').trim();
          const docStaff = row.getCell(3).getText().replace(/\r\n|\r/g, '\n').trim() || '';

          const decHeaderLines = [];
          const decArrowLines = [];
          rawDocDecFull.split('\n').forEach(function(line) {
            const s = line.trim();
            if (!s) return;
            if (s.startsWith('→')) decArrowLines.push(s);
            else decHeaderLines.push(s);
          });
          const rawDocDec = dedupeRepeatedDecisionHeader64_(decHeaderLines.join(' ').trim());

          if (decArrowLines.length > 0) {
            const statusDisplay = decArrowLines.map(function(a) {
              return '・【調整結果】' + a.replace(/^→\s*/, '');
            }).join('\n');
            col2 = col2 ? (col2 + '\n' + statusDisplay) : statusDisplay;
          }

          let parsedDec = (!docStaff && (rawDocDec === '案1' || rawDocDec === '要検討')) ? '' : rawDocDec;
          let parsedComment = '';
          const cmMatch = parsedDec.match(/^(.*?)\s*[（\(]メモ[：:](.+?)[）\)]$/);
          if (cmMatch) {
            parsedDec = dedupeRepeatedDecisionHeader64_(cmMatch[1].trim());
            parsedComment = cmMatch[2].trim();
          }

          const lines1 = Array.from(new Set(col1.split('\n').map(function(s) { return s.trim(); }).filter(Boolean)));
          const rawClinicTitle = lines1[0] || '';
          const clinicTitle = formatClinicTitle64_(rawClinicTitle);
          const origTimeAndWage = lines1.slice(1).join(' ') || '';
          const rowTimeRange = extractSlotTimeRange64_(origTimeAndWage);
          const meta = resolveClinicMeta64_(rawClinicTitle, masters.clinicMap);

          const recruitResolved = resolveRecruitDisplayTimeAndWage64_(
            sortDate, meta.name, rawClinicTitle, origTimeAndWage, recruitSlotsMap
          );
          const timeAndWage = recruitResolved.displayTimeAndWage;
          if (recruitResolved.rawAbsenceNote && !col2.includes('📌募集枠:')) {
            col2 = col2 ? (col2 + '\n' + recruitResolved.rawAbsenceNote) : recruitResolved.rawAbsenceNote;
          }

          const carryKey = getSlotCarryOverKey3_(sortDate, rawClinicTitle, timeAndWage);
          const subStatus = (typeof evaluateRowSubmittedStatusFor64_ === 'function')
            ? evaluateRowSubmittedStatusFor64_(carryKey, parsedDec, docStaff, decArrowLines, submittedLedgerMap)
            : { isPreviouslySubmitted: false, allDeclinedRevived: false, submittedAt: '' };

          if (subStatus.allDeclinedRevived) {
            col2 += '\n・⚠️【要再選定】前回打診した医師が全員辞退のため、次案を選択してください';
          }

          const slotKey = i + '_' + r;
          let draftItem = (draftMap.slots && draftMap.slots[slotKey]) ? draftMap.slots[slotKey] : null;
          if (draftItem) {
            const clinicMismatch = draftItem.clinicTitle && draftItem.clinicTitle !== clinicTitle;
            const draftRange = extractSlotTimeRange64_(draftItem.timeAndWage);
            const timeMismatch = draftRange && rowTimeRange && !isTimeRangeOverlapping64_(draftRange, rowTimeRange);
            if (clinicMismatch || timeMismatch) {
              draftItem = null;
            }
          }

          if (!draftItem && allDraftSlotEntries.length > 0) {
            const fallbackHits = allDraftSlotEntries.filter(function(d) {
              if (!d.dateLabel || toSortDateKey64_(d.dateLabel) !== sortDate) return false;
              if (d.clinicTitle !== clinicTitle) return false;
              const dRange = extractSlotTimeRange64_(d.timeAndWage);
              if (dRange && rowTimeRange) {
                return isTimeRangeOverlapping64_(dRange, rowTimeRange);
              }
              return true;
            });
            if (fallbackHits.length === 1) {
              draftItem = fallbackHits[0];
            }
          }

          const cleanDraftDec = draftItem
            ? dedupeRepeatedDecisionHeader64_(String(draftItem.decision || '').split(/\r?\n/).filter(function(l) { return !l.trim().startsWith('→'); }).join(' ').trim())
            : parsedDec;
          const cleanDraftAsg = draftItem ? String(draftItem.assignee || '') : docStaff;
          const cleanDraftCmnt = (draftItem && typeof draftItem.customComment === 'string') ? draftItem.customComment : parsedComment;

          slotRows.push({
            slotKey: slotKey,
            carryKey: carryKey,
            childIdx: i,
            rowIdx: r,
            sortDate: sortDate,
            dateLabel: jpDate,
            clinicTitle: clinicTitle,
            clinicNameClean: meta.name,
            clinicGroup: meta.group,
            macroArea: meta.area,
            clinicNo: meta.no,
            timeAndWage: timeAndWage,
            col2Plans: col2,
            origDecision: parsedDec,
            origAssignee: docStaff,
            origComment: parsedComment,
            draftDecision: cleanDraftDec,
            draftAssignee: cleanDraftAsg,
            draftComment: cleanDraftCmnt,
            curDecision: cleanDraftDec,
            curAssignee: cleanDraftAsg,
            customComment: cleanDraftCmnt,
            customDocs: (draftItem && Array.isArray(draftItem.customDocs)) ? draftItem.customDocs : [],
            hasDraft: Boolean(draftItem),
            isPreviouslySubmitted: Boolean(subStatus.isPreviouslySubmitted),
            submittedAt: subStatus.submittedAt || ''
          });
        }
      } else if (h0.includes('対象医師')) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;
          const bKey = 'b_' + i + '_' + r;
          const docVal = row.getCell(3).getText().trim() || '採用';
          const draftVal = (draftMap.bundles && draftMap.bundles[bKey]) ? draftMap.bundles[bKey] : null;

          bundleRows.push({
            bundleKey: bKey,
            childIdx: i,
            rowIdx: r,
            doctor: cleanSidebarDisplayText_(row.getCell(0).getText().replace(/\r\n|\r/g, '\n').trim()),
            slots: row.getCell(1).getText().replace(/\r\n|\r/g, '\n').trim(),
            note: row.getCell(2).getText().replace(/\r\n|\r/g, '\n').trim(),
            origVal: docVal,
            curVal: draftVal || docVal
          });
        }
      }
    }
  }

  slotRows.sort(function(a, b) {
    if (a.sortDate !== b.sortDate) return a.sortDate.localeCompare(b.sortDate);
    const rankA = (typeof getRegionSortRank64_ === 'function') ? getRegionSortRank64_(a.clinicGroup) : 99;
    const rankB = (typeof getRegionSortRank64_ === 'function') ? getRegionSortRank64_(b.clinicGroup) : 99;
    if (rankA !== rankB) return rankA - rankB;
    if (a.clinicNo !== b.clinicNo) return a.clinicNo - b.clinicNo;
    return String(a.timeAndWage || '').localeCompare(String(b.timeAndWage || ''));
  });

  const clientSessionId = Utilities.getUuid().slice(0, 8);
  const activeLocks = getActivePlayerLocks64_();

  return {
    clientSessionId: clientSessionId,
    docId: docId,
    docName: doc.getName(),
    docUrl: doc.getUrl(),
    staffList: masters.staffList,
    doctorSearchList: masters.doctorSearchList,
    backupDocsByDate: backupDocsByDate,
    groups: masters.groups,
    areaToGroups: masters.areaToGroups,
    slotRows: slotRows,
    bundleRows: bundleRows,
    activeLocks: activeLocks
  };
}