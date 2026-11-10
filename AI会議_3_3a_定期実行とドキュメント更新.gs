/**
 * ============================================================================
 * AI会議_3_3a_定期実行とドキュメント更新.gs
 * 1) 手動・定期トリガーからの実行窓口（シフト応募通知メールの差分スキャンも自動実行）
 * 2) Googleドキュメントの「→ 調整結果（シフト応募による充足含む）」＆緑/赤マーカー更新
 * 3) 【本来の仕様準拠】「決定済みプランかつ未充足」および「BQ募集中枠」は確実に継続表示し、
 *    本当に充足・募集終了して解消された枠のみを安全に自動削除する版
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

  mailHistoryIndex._applyIndex = applyIndex;

  if (typeof syncNegotiationHistoryAndCautionSheets3_3_ === 'function') {
    syncNegotiationHistoryAndCautionSheets3_3_(ss, mailHistoryIndex);
  }

  if (typeof updateExistingSlackThreadMessagesInC0AJTMLF0UD_ === 'function') {
    updateExistingSlackThreadMessagesInC0AJTMLF0UD_(mailHistoryIndex);
  }

  if (docId) {
    updateGoogleDocAndRemoveResolvedRows3_3_(ss, docId, mailHistoryIndex);
  }
}

/**
 * BQパック（bosyu / mijyusoku）上で現在まだ「公開募集中」の枠キー集合を取得するヘルパー
 * （2診目募集枠や『最新_不在枠』シート期間外の未来枠を「充足済み」と誤認して消さないため）
 */
function buildOpenBqRecruitKeySet3_3_() {
  const openKeys = new Set();
  if (typeof fetchAndDecryptBqMailForResolver_ !== 'function') return openKeys;

  try {
    const bqData = fetchAndDecryptBqMailForResolver_();
    if (!bqData) return openKeys;

    const list = (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []);
    list.forEach(function(m) {
      const st = String(m.st || '').trim();
      if (st && (st.includes('非掲載') || st.includes('終了') || st.includes('取消') || st.includes('キャンセル') || st.includes('充足'))) {
        return;
      }
      const d = normalizeDateYmd1b3_(m.d || m.date);
      const rawC = String(m.c || m.clinic || '').trim();
      const rawDp = String(m.dp || m.dept || '').trim();
      const rawT = String(m.t || m.time || '').trim();
      if (!d || !rawC || !rawT) return;

      const clinicWithDept = (rawDp.includes('内科') && !rawC.includes('内科')) ? (rawC + '（内科）') : rawC;
      const expanded = (typeof normalizeAndExpandAbsenceTime1b3_ === 'function')
        ? normalizeAndExpandAbsenceTime1b3_(rawT, rawC)
        : [rawT];

      expanded.forEach(function(t) {
        if (t) openKeys.add(getSlotCarryOverKey3_(d, clinicWithDept, t));
      });
    });
  } catch (e) {
    Logger.log('⚠️ BQ公開募集枠チェックの読込スキップ: ' + e.message);
  }
  return openKeys;
}

function updateGoogleDocAndRemoveResolvedRows3_3_(ss, docId, mailHistoryIndex) {
  const unfilledSh = ss.getSheetByName('最新_不在枠');
  const rawUnfilled = (unfilledSh && unfilledSh.getLastRow() >= 2)
    ? unfilledSh.getDataRange().getDisplayValues().slice(1)
    : [];

  const shiftSh = ss.getSheetByName('最新_当日シフト');
  const activeShifts = (shiftSh && shiftSh.getLastRow() >= 2)
    ? shiftSh.getDataRange().getDisplayValues().slice(1)
    : [];

  const todayYmd = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const maxCheckYmd = Utilities.formatDate(new Date(Date.now() + 35 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const effectiveUnfilled = (typeof recoverMissingFullDayOrUnfilledRows1b3_ === 'function')
    ? recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilled, activeShifts, [], todayYmd, maxCheckYmd)
    : rawUnfilled;

  const openBqRecruitKeys = buildOpenBqRecruitKeySet3_3_();

  const officialNames = [];
  if (typeof loadDoctorSearchList64_ === 'function') {
    loadDoctorSearchList64_(ss).forEach(function(d) {
      if (d.name) officialNames.push(d.name);
    });
  }

  const activeUnfilledKeys = new Set();
  const memoBySlotKey = new Map();
  let maxPrimaryUnfilledDate = '';

  rawUnfilled.forEach(function(r) {
    const d = normalizeDateYmd1b3_(r[0]);
    if (d && d > maxPrimaryUnfilledDate) maxPrimaryUnfilledDate = d;
  });

  effectiveUnfilled.forEach(function(r) {
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

  const setCellSafelyLocal = function(cell, txt) {
    if (typeof setDocCellTextSafely64_ === 'function') {
      setDocCellTextSafely64_(cell, txt);
      return;
    }
    const clean = String(txt || '');
    while (cell.getNumChildren() > 1) {
      cell.removeChild(cell.getChild(cell.getNumChildren() - 1));
    }
    const p = cell.getChild(0).asParagraph();
    if (clean.length === 0) p.clear();
    else p.setText(clean);
  };

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

        // 1. セル内改行(\r)を \n に正規化
        const c0 = row.getCell(0).getText().replace(/\r\n|\r/g, '\n').trim();
        const rawC1 = row.getCell(1).getText().replace(/\r\n|\r/g, '\n')
          .replace(/([^\n])\s*(?=【案[123１２３]】)/g, '$1\n')
          .replace(/([^\n])\s*(?=💡|📝|・【進行中メモ】|・⚠️)/g, '$1\n')
          .trim();
        const c1 = (typeof dedupePlanCellLines64_ === 'function')
          ? dedupePlanCellLines64_(rawC1)
          : rawC1.replace(/(【案[123１２３]】[^\n]+?)\1+/g, '$1');

        const c2Raw = row.getCell(2).getText().replace(/\r\n|\r/g, '\n').trim();
        const c3Assignee = row.getCell(3).getText().replace(/\r\n|\r/g, ' ').trim();

        // 2. 過去の「→」行を綺麗に剥がして確定案ヘッダーのみを取得
        const c2HeaderLines = [];
        let inArrowBlock = false;
        c2Raw.split('\n').forEach(function(l) {
          const trimmed = l.trim();
          if (!trimmed) return;
          if (/^(?:→|⇒|⇨)/.test(trimmed)) {
            inArrowBlock = true;
            return;
          }
          if (!inArrowBlock) {
            const strippedInlineArrow = trimmed.replace(/(?:→|⇒|⇨).*$/, '').trim();
            if (strippedInlineArrow) c2HeaderLines.push(strippedInlineArrow);
            if (/(?:→|⇒|⇨)/.test(trimmed)) inArrowBlock = true;
          }
        });
        const c2Header = c2HeaderLines.join(' ').trim();

        const c0Lines = Array.from(new Set(c0.split('\n').map(function(s) { return s.trim(); }).filter(Boolean)));
        const clinicTitle = c0Lines[0] || '';
        const timeAndWage = c0Lines.slice(1).join(' ');
        const carryKey = getSlotCarryOverKey3_(curDate, clinicTitle, timeAndWage);

        // 3. 先に最新の調整結果（承諾・応募充足・返信待ち・辞退）を計算する
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
          assignee: c3Assignee,
          prevPlanMap: { '1': planObj.plan1, '2': planObj.plan2, '3': planObj.plan3 }
        };
        const resolved = resolveCarriedOverDecisionForSlot3_(prevEntry, planObj, { currentMemo: latestMemoText }, mailHistoryIndex);
        let finalDecisionText = String(resolved.decision || '').replace(/\r\n|\r/g, '\n');

        if (applyIndex && typeof findMatchingShiftApplicationsForSlot3_3d_ === 'function') {
          const cleanC = cleanClinicBaseName_(clinicTitle);
          const tmMatch = timeAndWage.match(/(\d{1,2}:\d{2}\s*[\-~～]\s*\d{1,2}:\d{2})/);
          const slotTimeRange = tmMatch ? tmMatch[1].replace(/\s+/g, '').replace(/[～~]/g, '-') : '';
          const matchedApps = findMatchingShiftApplicationsForSlot3_3d_(curDate, cleanC, slotTimeRange, applyIndex);

          if (matchedApps.length > 0) {
            const appsByDoc = new Map();
            matchedApps.forEach(function(app) {
              const dKey = normKeyForDoc_(app.doctorName);
              if (!dKey) return;
              if (!appsByDoc.has(dKey)) {
                appsByDoc.set(dKey, {
                  doctorName: app.doctorName,
                  receivedDateShort: app.receivedDateShort,
                  times: app.workTime ? [app.workTime] : []
                });
              } else {
                const ex = appsByDoc.get(dKey);
                if (app.workTime && !ex.times.includes(app.workTime)) ex.times.push(app.workTime);
              }
            });

            let decLines = finalDecisionText ? finalDecisionText.split('\n') : [];

            appsByDoc.forEach(function(appGroup, dKey) {
              let matchedPlanTag = '';
              ['1', '2', '3'].forEach(function(pNum) {
                if (!matchedPlanTag && planObj['plan' + pNum] && normKeyForDoc_(planObj['plan' + pNum]).includes(dKey)) {
                  matchedPlanTag = '(案' + pNum + ')';
                }
              });

              decLines = decLines.filter(function(line) {
                const trimmed = line.trim();
                if (!trimmed.startsWith('→')) return true;
                if (normKeyForDoc_(trimmed).includes(dKey)) return false;
                return true;
              });

              const sortedTimes = appGroup.times.sort().join('・');
              const timeNote = (sortedTimes && sortedTimes !== slotTimeRange) ? ('[' + sortedTimes + ']') : '';
              const appLine = '→ ⭕' + appGroup.doctorName + matchedPlanTag + '：応募あり充足' + timeNote + '(' + appGroup.receivedDateShort + ')';
              decLines.push(appLine);
            });

            finalDecisionText = decLines.filter(Boolean).join('\n');
          }
        }

        // 重複した「→」行を一意化
        if (finalDecisionText) {
          const seenArrow = new Set();
          finalDecisionText = finalDecisionText.split('\n').filter(function(l) {
            const t = l.trim();
            if (!t.startsWith('→')) return true;
            if (seenArrow.has(t)) return false;
            seenArrow.add(t);
            return true;
          }).join('\n');
        }

        // =====================================================================
        // ★4. 【本来の仕様】継続表示 vs 解消済み削除の厳密判定
        // =====================================================================
        const isPastDate = (curDate < todayYmd);
        const isStillUnfilled =
          activeUnfilledKeys.has(carryKey) ||
          openBqRecruitKeys.has(carryKey) ||
          (maxPrimaryUnfilledDate && curDate > maxPrimaryUnfilledDate && openBqRecruitKeys.size === 0);

        const isFilledOrAccepted =
          /^充足/.test(c2Header) ||
          (finalDecisionText.includes('⭕') && (finalDecisionText.includes('承諾') || finalDecisionText.includes('充足')));

        const hasDecidedPlan =
          Boolean(c2Header && c2Header !== '要検討' && !/^充足/.test(c2Header)) ||
          Boolean(c3Assignee && c3Assignee !== '担当者未定' && c3Assignee !== '(未定)' && c3Assignee !== '未定');

        // 削除するのは「過去日」または「最新_不在枠・BQ募集枠の両方から消滅しており、かつ（未充足の決定済みプランではない or ⭕充足確定済み）」のときのみ！
        const shouldRemoveRow = isPastDate || (
          activeUnfilledKeys.size > 0 &&
          !isStillUnfilled &&
          (!hasDecidedPlan || isFilledOrAccepted)
        );

        if (shouldRemoveRow) {
          rowsToRemove.push(r);
          removedSlotCount++;
          continue;
        }

        // 5. 継続表示する枠のテキスト・マーカーを更新
        if (finalDecisionText !== c2Raw || row.getCell(2).getNumChildren() > 1) {
          setCellSafelyLocal(row.getCell(2), finalDecisionText);
          row.getCell(2).editAsText().setFontSize(8.5).setBold(true);
        }
        if (typeof applySelectedPlanMarkerToDocCell64_ === 'function' && (finalDecisionText || rawC1 !== c1 || row.getCell(1).getNumChildren() > 1)) {
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
  Logger.log('✅ ドキュメント更新完了（本当に充足・解消された枠の削除数: ' + removedSlotCount + ' 枠）');
}