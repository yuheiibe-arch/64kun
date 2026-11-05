/**
 * ============================================================================
 * AI会議_3_3a_定期実行とドキュメント更新.gs
 * 1) 手動・定期トリガーからの実行窓口（シフト応募通知メールの差分スキャンも自動実行）
 * 2) Googleドキュメントの「→ 調整結果（シフト応募による充足含む）」＆緑/赤マーカー更新
 * 3) 1日2回のデータ更新で『最新_不在枠』から消えた（確定シフト反映済の）枠をドキュメントから自動削除
 * ============================================================================
 */

function runSyncDocAndSlackResultsNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  if (typeof syncLatestMailsBeforeMeeting1a_ === 'function') {
    syncLatestMailsBeforeMeeting1a_(ss);
  }
  syncAdjustmentResultsToDocAndSlack3_3_(ss);
}

function syncAdjustmentResultsToDocAndSlack3_3_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');

  // 0. 「【要対応】★医師シフト応募通知★」メールの差分スキャン＆応募インデックス構築（3_3d）
  let applyIndex = null;
  if (typeof scanGmailShiftApplicationsIncremental3_3d_ === 'function') {
    try {
      scanGmailShiftApplicationsIncremental3_3d_(ss, false);
      applyIndex = buildShiftApplicationIndex3_3d_(ss);
    } catch (e) {
      Logger.log('⚠️ シフト応募通知スキャンエラー: ' + e.message);
    }
  }

  const mailHistoryIndex = (typeof buildDoctorMailHistoryIndexForResults1b2_ === 'function')
    ? buildDoctorMailHistoryIndexForResults1b2_(ss)
    : null;
  if (!mailHistoryIndex) return;

  // 応募データを mailHistoryIndex に添付して各モジュールから参照可能にする
  mailHistoryIndex._applyIndex = applyIndex;

  // 1. 『交渉履歴』シート（相談日順）＆『交渉注意医師』シートの更新（3_3c）
  if (typeof syncNegotiationHistoryAndCautionSheets3_3_ === 'function') {
    syncNegotiationHistoryAndCautionSheets3_3_(ss, mailHistoryIndex);
  }

  // 2. C0AJTMLF0UD の「超募集MTGログ（前半・後半）」スレッド内の既存 dskun メッセージを上書き編集（3_3b）
  if (typeof updateExistingSlackThreadMessagesInC0AJTMLF0UD_ === 'function') {
    updateExistingSlackThreadMessagesInC0AJTMLF0UD_(mailHistoryIndex);
  }

  // 3. Googleドキュメントの更新（『最新_不在枠』に残っている間は緑マーカー表示、消えたら行削除）
  if (docId) {
    updateGoogleDocAndRemoveResolvedRows3_3_(ss, docId, mailHistoryIndex);
  }
}

function updateGoogleDocAndRemoveResolvedRows3_3_(ss, docId, mailHistoryIndex) {
  const unfilledSh = ss.getSheetByName('最新_不在枠');
  const rawUnfilled = (unfilledSh && unfilledSh.getLastRow() >= 2)
    ? unfilledSh.getDataRange().getDisplayValues().slice(1)
    : [];

  const officialNames = [];
  if (typeof loadDoctorSearchList64_ === 'function') {
    loadDoctorSearchList64_(ss).forEach(function(d) {
      if (d.name) officialNames.push(d.name);
    });
  }

  const activeUnfilledKeys = new Set();
  const memoBySlotKey = new Map();

  rawUnfilled.forEach(function(r) {
    const d = normalizeDateYmd1b3_(r[0]);
    const rawC = String(r[1] || '').trim();
    const rawT = String(r[2] || '').trim();
    if (!d || !rawC) return;

    const key = getSlotCarryOverKey3_(d, rawC, rawT);
    activeUnfilledKeys.add(key);

    const memo = String(r[12] || '').trim();
    if (memo && memo !== '-') {
      const corrected = (typeof correctDoctorNamesInMemoText_ === 'function')
        ? correctDoctorNamesInMemoText_(memo, officialNames)
        : memo;
      memoBySlotKey.set(key, corrected);
    }
  });

  const applyIndex = mailHistoryIndex ? mailHistoryIndex._applyIndex : null;
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  let curDate = '';
  let lastDateHeadingParagraph = null;
  const elementsToRemove = [];
  let removedSlotCount = 0;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const txt = p.getText().trim();
      const dm = txt.match(/^◆\s*(\d{4}\/\d{2}\/\d{2})/);
      if (dm) {
        curDate = dm[1];
        lastDateHeadingParagraph = p;
      }
    } else if (el.getType() === DocumentApp.ElementType.TABLE && curDate) {
      const tbl = el.asTable();
      if (tbl.getNumRows() < 2) continue;
      const h0 = tbl.getRow(0).getCell(0).getText().trim();
      if (!h0.includes('拠点・不在時間帯')) continue;

      const rowsToRemove = [];

      for (let r = 1; r < tbl.getNumRows(); r++) {
        const row = tbl.getRow(r);
        if (row.getNumCells() < 4) continue;

        const c0 = row.getCell(0).getText().trim();
        const c1 = row.getCell(1).getText().trim();
        const c2Raw = row.getCell(2).getText().trim();
        const c2Header = c2Raw.split('\n').filter(function(l) { return !l.trim().startsWith('→'); }).join(' ').trim();

        const c0Lines = c0.split('\n');
        const clinicTitle = c0Lines[0] || '';
        const timeAndWage = c0Lines.slice(1).join(' ');
        const carryKey = getSlotCarryOverKey3_(curDate, clinicTitle, timeAndWage);

        if (activeUnfilledKeys.size > 0 && !activeUnfilledKeys.has(carryKey)) {
          rowsToRemove.push(r);
          removedSlotCount++;
          continue;
        }

        const planObj = { date: curDate, clinicAndDept: clinicTitle, plan1: '', plan2: '', plan3: '' };
        const memoLinesFromCol2 = [];

        c1.split('\n').forEach(function(line) {
          const s = line.trim();
          const pm = s.match(/^【案([123１２３])】(.*)$/);
          if (pm) {
            const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
            planObj['plan' + pNum] = pm[2].trim();
          } else if (s.includes('【進行中メモ】')) {
            memoLinesFromCol2.push(s.replace(/^・?\s*【進行中メモ】\s*/, '').trim());
          }
        });

        const latestMemoText = memoBySlotKey.get(carryKey) || memoLinesFromCol2.join('\n');
        const prevEntry = {
          rawDecision: c2Header,
          assignee: row.getCell(3).getText().trim(),
          prevPlanMap: { '1': planObj.plan1, '2': planObj.plan2, '3': planObj.plan3 }
        };
        const resolved = resolveCarriedOverDecisionForSlot3_(prevEntry, planObj, { currentMemo: latestMemoText }, mailHistoryIndex);
        let finalDecisionText = resolved.decision;

        // ★シフト応募通知メールで該当枠への応募が入っている場合、3列目に「→ ⭕〇〇：応募あり充足」を追記
        if (applyIndex && typeof findMatchingShiftApplicationsForSlot3_3d_ === 'function') {
          const cleanC = cleanClinicBaseName_(clinicTitle);
          const tmMatch = timeAndWage.match(/(\d{1,2}:\d{2}\s*[\-~～]\s*\d{1,2}:\d{2})/);
          const matchedApps = findMatchingShiftApplicationsForSlot3_3d_(curDate, cleanC, tmMatch ? tmMatch[1] : '', applyIndex);
          matchedApps.forEach(function(app) {
            if (!finalDecisionText.includes(app.doctorName)) {
              const appLine = '→ ⭕' + app.doctorName + '：応募あり充足(' + app.receivedDateShort + ')';
              finalDecisionText = finalDecisionText ? (finalDecisionText + '\n' + appLine) : appLine;
            }
          });
        }

        if (finalDecisionText !== c2Raw) {
          row.getCell(2).setText(finalDecisionText);
          row.getCell(2).editAsText().setFontSize(8.5).setBold(true);
        }
        if (typeof applySelectedPlanMarkerToDocCell64_ === 'function' && finalDecisionText) {
          applySelectedPlanMarkerToDocCell64_(row.getCell(1), c1, finalDecisionText);
        }
      }

      for (let idx = rowsToRemove.length - 1; idx >= 0; idx--) {
        tbl.removeRow(rowsToRemove[idx]);
      }

      if (tbl.getNumRows() <= 1) {
        elementsToRemove.push(tbl);
        if (lastDateHeadingParagraph) {
          elementsToRemove.push(lastDateHeadingParagraph);
          lastDateHeadingParagraph = null;
        }
      }
    }
  }

  elementsToRemove.forEach(function(el) {
    try {
      body.removeChild(el);
    } catch (e) {}
  });

  doc.saveAndClose();
  Logger.log('✅ ドキュメント更新完了（最新_不在枠から解消された枠の削除数: ' + removedSlotCount + ' 枠）');
}