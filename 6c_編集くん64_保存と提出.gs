/**
 * ============================================================================
 * 6c_編集くん64_保存と提出.gs
 * （OK［⭕承諾・応募充足］の案は「薄い緑色マーカー #bbf7d0」、通常の選択案は「薄い赤色マーカー #ffd6d6」で色分けする版）
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
      draft.slots[slotKey] = {
        decision: val.decision,
        assignee: val.assignee,
        customComment: String(val.customComment || '').trim(),
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

/**
 * ★改善版：確定案および調整結果（rawDecisionText）を解析し、
 * - ⭕承諾（OK・応募充足）が出ている案 ➔ 「爽やかな緑色マーカー（#bbf7d0）」
 * - それ以外の選択案（打診中・返信待ち等） ➔ 「薄い赤色マーカー（#ffd6d6）」
 * を2列目（AI推奨アプローチ列）の該当する【案X】の行に付与する関数。
 * また、【案1】...【案1】...のような重複文字列があれば自動で1つにクリーニングする。
 */
function applySelectedPlanMarkerToDocCell64_(cell1, rawC1Text, rawDecisionText) {
  const c1Text = String(rawC1Text || '').replace(/(【案[123１２３]】[^\n]+?)\1+/g, '$1');
  cell1.setText(c1Text);
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

  // 各【案X】のテキストをマップ化（医師名での逆引き照合用）
  const planLineMap = {};
  c1Text.split('\n').forEach(function(line) {
    const pm = line.trim().match(/^【案([123１２３])】(.*)$/);
    if (pm) {
      const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
      planLineMap[pNum] = pm[2].trim();
    }
  });

  // 1. 「→ ⭕...(案X)...：承諾 / 応募あり充足」の行から、承諾（OK）済みの案番号を抽出
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
        // もし「→ ⭕宗像恵美子：応募あり充足」のように(案X)が書かれていない場合でも、医師名が案1〜案3にあれば自動で緑マーカー対象にする
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

  // 2. 1行目の確定案本体から選択されている案番号を抽出
  const firstLine = fullDecStr.split('\n')[0]
    .replace(/\s*[（\(]メモ[：:].*?[）\)]$/g, '')
    .replace(/【特別時給.*?】/g, '')
    .trim();

  const selectedPlanNums = new Set();
  const subDocChoiceByPlan = {};

  if (firstLine && !/全案|すべて|全部|ステイ|要検討/.test(firstLine)) {
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

  const COLOR_SELECTED_RED = '#ffd6d6'; // 通常選択・交渉中（薄い赤色）
  const COLOR_ACCEPTED_GREEN = '#bbf7d0'; // ⭕承諾・応募充足確定（爽やかな緑色）

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

            let c1 = cleanSidebarDisplayText_(row.getCell(1).getText())
              .replace(/(【案[123１２３]】[^\n]+?)\1+/g, '$1')
              .replace(/\n*📝コメント[：:].*$/g, '')
              .trim();
            const cmnt = String(item.customComment || '').trim();
            if (cmnt) {
              c1 += '\n📝コメント：' + cmnt;
            }

            const existingC2Lines = row.getCell(2).getText().split('\n');
            const arrowLines = existingC2Lines.filter(function(l) { return l.trim().startsWith('→'); });

            const baseDec = String(item.decision || '').trim();
            const headerDecText = (cmnt && baseDec) ? (baseDec + '（メモ：' + cmnt + '）') : baseDec;
            const fullDecText = [headerDecText].concat(arrowLines).filter(Boolean).join('\n');

            row.getCell(0).setText(c0Lines.join('\n'));
            row.getCell(2).setText(fullDecText);
            row.getCell(3).setText(String(item.assignee || '').trim());
            row.getCell(0).editAsText().setFontSize(8.5);
            row.getCell(2).editAsText().setFontSize(8.5).setBold(true);
            row.getCell(3).editAsText().setFontSize(8.5);

            applySelectedPlanMarkerToDocCell64_(row.getCell(1), c1, fullDecText);
          }
        }
      }
    });

    const numChildren = body.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (table.getNumRows() < 2) continue;
        const h0 = table.getRow(0).getCell(0).getText().trim();
        if (h0.includes('拠点・不在時間帯')) {
          for (let r = 1; r < table.getNumRows(); r++) {
            const row = table.getRow(r);
            if (row.getNumCells() < 4) continue;
            const c1Text = cleanSidebarDisplayText_(row.getCell(1).getText().trim());
            const decText = row.getCell(2).getText().trim();
            applySelectedPlanMarkerToDocCell64_(row.getCell(1), c1Text, decText);
          }
        }
      }
    }

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