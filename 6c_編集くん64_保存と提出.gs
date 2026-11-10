/**
 * ============================================================================
 * 6c_編集くん64_保存と提出.gs
 * （段落残りによる選択肢2重化バグ完全解消・行削除時のインデックスずれ防止・緑/赤マーカー版）
 * ============================================================================
 */

/**
 * ★重要：GoogleドキュメントのTableCellにテキストをセットする際、
 * 2段落目以降が残って選択肢が2重表示されるのを防ぐ安全なセル書き込みヘルパー
 */
/**
 * ★修正版：空文字("")のときは p.clear() で安全にクリアし、
 * 「空のテキスト要素は挿入できません」エラーと段落ダブりの両方を防ぐセル書き込みヘルパー
 */
function setDocCellTextSafely64_(cell, text) {
  const clean = String(text || '');
  while (cell.getNumChildren() > 1) {
    cell.removeChild(cell.getChild(cell.getNumChildren() - 1));
  }
  const first = cell.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const p = first.asParagraph();
    if (clean.length === 0) {
      p.clear();
    } else {
      p.setText(clean);
    }
  } else {
    const p = cell.appendParagraph('');
    if (clean.length > 0) {
      p.setText(clean);
    }
    cell.removeChild(first);
  }
}

/**
 * すでに2重化してしまった【案1】〜【案3】や注意書きを1回ずつに自動修復するヘルパー
 */
function dedupePlanCellLines64_(rawText) {
  const normalized = String(rawText || '').replace(/(【案[123１２３]】[^\n]+?)\1+/g, '$1');
  const lines = normalized.split(/\r?\n/);
  const seenPlanNums = new Set();
  const seenExactLines = new Set();
  const out = [];

  lines.forEach(function(line) {
    const trimmed = line.trim();
    if (!trimmed) return;
    const pm = trimmed.match(/^【案([123１２３])】/);
    if (pm) {
      const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
      if (seenPlanNums.has(pNum)) return;
      seenPlanNums.add(pNum);
      out.push(trimmed);
      return;
    }
    if (seenExactLines.has(trimmed)) return;
    seenExactLines.add(trimmed);
    out.push(trimmed);
  });

  return out.join('\n');
}

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

  const listHeaders = ['スロットキー', '日付', '拠点', '確定案（一時保存）', '特定済カスタム医師（医籍番号）', '担当者名（一時保存）', 'カスタムコメント・メモ', '更新時刻'];
  sh.getRange(4, 1, 1, 8).setValues([listHeaders]).setBackground('#0046ad').setFontColor('#ffffff').setFontWeight('bold');

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
        v.customComment || '',
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

    Object.entries(payload.changedSlots || {}).forEach(([slotKey, val]) => {
      // 誤って「→ 調整結果」行が含まれていた場合も確実に除去して保存
      const cleanDec = String(val.decision || '')
        .split(/\r?\n/)
        .filter(l => !l.trim().startsWith('→'))
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

/**
 * 確定案および調整結果（rawDecisionText）を解析し、
 * - ⭕承諾（OK・応募充足）が出ている案 ➔ 「爽やかな緑色マーカー（#bbf7d0）」
 * - それ以外の選択案（打診中・返信待ち等） ➔ 「薄い赤色マーカー（#ffd6d6）」
 * を2列目の該当する【案X】に付与する関数（2段落目以降の残留バグ完全対策済）
 */
function applySelectedPlanMarkerToDocCell64_(cell1, rawC1Text, rawDecisionText) {
  const c1Text = dedupePlanCellLines64_(rawC1Text);
  setDocCellTextSafely64_(cell1, c1Text);
  if (!c1Text) return;

  const textObj = cell1.editAsText();
  const len = c1Text.length;
  textObj.setFontSize(0, len - 1, 8.5);
  textObj.setBold(0, len - 1, false);
  textObj.setBackgroundColor(0, len - 1, null);

  if (typeof highlightKeywordsInText_ === 'function') {
    highlightKeywordsInText_(textObj);
  }

  const fullDecStr = String(rawDecisionText || '').trim();
  if (!fullDecStr) return;

  const planLineMap = {};
  c1Text.split('\n').forEach(function(line) {
    const pm = line.trim().match(/^【案([123１２３])】(.*)$/);
    if (pm) {
      const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
      planLineMap[pNum] = pm[2].trim();
    }
  });

  const acceptedPlanNums = new Set();
  fullDecStr.split('\n').forEach(function(line) {
    const s = line.trim();
    if (s.startsWith('→') && s.includes('⭕')) {
      const m = s.match(/\(案([123１２３])\)/g);
      if (m) {
        m.forEach(function(tok) {
          const num = tok.replace(/[^123１２３]/g, '').replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
          if (num) acceptedPlanNums.add(num);
        });
      } else {
        const docMatch = s.match(/⭕\s*([^\(（：:\s]+)/);
        if (docMatch && typeof normKeyForDoc_ === 'function') {
          const dKey = normKeyForDoc_(docMatch[1]);
          if (dKey) {
            ['1', '2', '3'].forEach(function(pNum) {
              if (planLineMap[pNum] && normKeyForDoc_(planLineMap[pNum]).includes(dKey)) {
                acceptedPlanNums.add(pNum);
              }
            });
          }
        }
      }
    }
  });

  const firstLine = fullDecStr.split('\n')[0]
    .replace(/\s*[（\(]メモ[：:].*?[）\)]$/g, '')
    .replace(/【特別時給.*?】/g, '')
    .trim();

  const selectedPlanNums = new Set();
  const subDocChoiceByPlan = {};

  if (firstLine && !firstLine.startsWith('→') && !/全案|すべて|全部|ステイ|要検討/.test(firstLine)) {
    const tokens = firstLine.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(function(t) { return t.trim(); }).filter(Boolean);
    tokens.forEach(function(tok) {
      const specMatch = tok.match(/^(?:【?案)?([123１２３])(?:】)?(?:\(([^)]+)\))?$/);
      if (specMatch) {
        const pNum = specMatch[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
        selectedPlanNums.add(pNum);
        const subText = specMatch[2] ? specMatch[2].trim() : '';
        if (subText) {
          if (subText.includes('両方')) subDocChoiceByPlan[pNum] = 'BOTH';
          else if (/^[①1１]/.test(subText)) subDocChoiceByPlan[pNum] = '1';
          else if (/^[②2２]/.test(subText)) subDocChoiceByPlan[pNum] = '2';
        }
      } else {
        const multiMatches = tok.match(/案([123１２３])/g);
        if (multiMatches && tok.replace(/案[123１２３]|[\s、,・と]/g, '') === '') {
          multiMatches.forEach(function(mm) {
            const pNum = mm.replace('案', '').replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
            selectedPlanNums.add(pNum);
          });
        }
      }
    });
  }

  acceptedPlanNums.forEach(function(num) { selectedPlanNums.add(num); });
  if (selectedPlanNums.size === 0) return;

  const COLOR_SELECTED_RED = '#ffd6d6';
  const COLOR_ACCEPTED_GREEN = '#bbf7d0';

  const lines = c1Text.split('\n');
  let curOffset = 0;

  lines.forEach(function(line) {
    const lineLen = line.length;
    const lineStart = curOffset;
    const lineEnd = curOffset + lineLen - 1;
    curOffset += lineLen + 1;

    if (lineLen <= 0) return;
    const m = line.match(/^【案([123１２３])】/);
    if (!m) return;

    const pNum = m[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
    if (!selectedPlanNums.has(pNum)) return;

    const targetColor = acceptedPlanNums.has(pNum) ? COLOR_ACCEPTED_GREEN : COLOR_SELECTED_RED;
    textObj.setBackgroundColor(lineStart, lineEnd, targetColor);
    textObj.setBold(lineStart, lineEnd, true);

    const subChoice = subDocChoiceByPlan[pNum];
    if (subChoice === '1' || subChoice === '2') {
      const pairRegex = /(①[^\/／\n]+?)(\s*[\/／]\s*)(②[^）\)\n]+)/;
      const pm = line.match(pairRegex);
      if (pm && typeof pm.index === 'number') {
        const seg1Start = lineStart + pm.index;
        const seg1End = seg1Start + pm[1].length + pm[2].length - 1;
        const seg2Start = lineStart + pm.index + pm[1].length;
        const seg2End = seg2Start + pm[2].length + pm[3].length - 1;

        if (subChoice === '1' && seg2Start <= seg2End) {
          textObj.setBackgroundColor(seg2Start, seg2End, null);
          textObj.setBold(seg2Start, seg2End, false);
        } else if (subChoice === '2' && seg1Start <= seg1End) {
          textObj.setBackgroundColor(seg1Start, seg1End, null);
          textObj.setBold(seg1Start, seg1End, false);
        }
      }
    }
  });
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
    const numChildren = body.getNumChildren();

    // ★15分トリガーで行削除が起きていた場合でも別枠へ誤上書きしないよう現在のドキュメント位置をマップ化
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
          const lookupKey = curSortDate + '|' + cTitle;
          if (!liveRowLookup.has(lookupKey)) liveRowLookup.set(lookupKey, []);
          liveRowLookup.get(lookupKey).push({ cIdx: i, rIdx: r, row: row });
        }
      }
    }

    const updatedSlotMap = {};

    Object.entries(draft.slots || {}).forEach(([slotKey, item]) => {
      const [cIdx, rIdx] = slotKey.split('_').map(Number);
      let targetRow = null;
      let actualSlotKey = slotKey;

      // まず元の cIdx_rIdx の拠点名が一致するか確認
      if (cIdx < numChildren) {
        const el = body.getChild(cIdx);
        if (el && el.getType() === DocumentApp.ElementType.TABLE) {
          const tbl = el.asTable();
          if (rIdx < tbl.getNumRows()) {
            const candRow = tbl.getRow(rIdx);
            if (candRow.getNumCells() >= 4) {
              const candClinic = formatClinicTitle64_(cleanSidebarDisplayText_(candRow.getCell(0).getText()).split('\n')[0] || '');
              if (!item.clinicTitle || candClinic === item.clinicTitle) {
                targetRow = candRow;
              }
            }
          }
        }
      }

      // もし行削除等でズレていた場合、日付＋拠点名から一意に特定できる行のみ安全にフォールバック
      if (!targetRow && item.dateLabel && item.clinicTitle) {
        const lKey = toSortDateKey64_(item.dateLabel) + '|' + item.clinicTitle;
        const cands = liveRowLookup.get(lKey) || [];
        if (cands.length === 1) {
          targetRow = cands[0].row;
          actualSlotKey = cands[0].cIdx + '_' + cands[0].rIdx;
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

      const baseDec = String(item.decision || '')
        .split(/\r?\n/)
        .filter(function(l) { return !l.trim().startsWith('→'); })
        .join(' ')
        .trim();
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

    // 未編集の行も含め、既存の2重化セルがあれば一括クリーンアップ＆マーカー再適用
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

    // ズレ補正後のスロットキーでドラフトを更新保存
    draft.slots = updatedSlotMap;
    props.setProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId, JSON.stringify(draft));

    Object.entries(draft.bundles || {}).forEach(([bKey, val]) => {
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