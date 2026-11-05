/**
 * ============================================================================
 * AI会議_3_1_ドキュメント引継と管理.gs
 * （前半：前回MTG引継 ＋ 進行中メモからの案自動特定＆コンパクト「→ 調整結果」記載版）
 * ============================================================================
 */

const DOC_OUTPUT_FOLDER_ID = '1syiypDralKBqHwCL-0R2XeJXbhxXVcLl';

function formatDateWithDow_(dateStr) {
  const s = String(dateStr || '').trim();
  if (!s || /[\(（][日月火水木金土][\)）]/.test(s)) return s;
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return s;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (isNaN(d.getTime())) return s;
  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) + '（' + dows[d.getDay()] + '）';
}

function getSlotCarryOverKey3_(dateStr, rawClinicText, rawTimeText) {
  const dMatch = String(dateStr || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  const d = dMatch ? (dMatch[1] + '/' + ('0' + dMatch[2]).slice(-2) + '/' + ('0' + dMatch[3]).slice(-2)) : String(dateStr || '').trim();
  const cleanC = cleanClinicBaseName_(rawClinicText);
  const dept = String(rawClinicText || '').includes('内科') ? '内科' : '小児科';
  const tm = String(rawTimeText || '').match(/(\d{1,2}):(\d{2})/);
  const sh = tm ? Number(tm[1]) : 9;
  const band = sh < 13 ? 'A' : (sh < 18 ? 'B' : 'C');
  return [d, cleanC, dept, band].join('|');
}

function loadPreviousMeetingDecisionsMap3_() {
  const map = new Map();
  const props = PropertiesService.getScriptProperties();
  const docIdsToScan = [];

  const lastSavedDocId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (lastSavedDocId) docIdsToScan.push(lastSavedDocId);

  try {
    const folder = DriveApp.getFolderById(DOC_OUTPUT_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    const mtgFiles = [];
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      if (f.getName().includes('超募集対策MTG_')) {
        mtgFiles.push(f);
      }
    }
    mtgFiles.sort(function(a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
    mtgFiles.slice(0, 3).reverse().forEach(function(f) {
      const id = f.getId();
      if (!docIdsToScan.includes(id)) docIdsToScan.push(id);
    });
  } catch (e) {}

  docIdsToScan.forEach(function(docId) {
    try {
      const d = DocumentApp.openById(docId);
      const b = d.getBody();
      const numCh = b.getNumChildren();
      let curDate = '';

      for (let i = 0; i < numCh; i++) {
        const el = b.getChild(i);
        if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const txt = el.asParagraph().getText().trim();
          const dm = txt.match(/^◆\s*(\d{4}\/\d{2}\/\d{2})/);
          if (dm) curDate = dm[1];
        } else if (el.getType() === DocumentApp.ElementType.TABLE && curDate) {
          const tbl = el.asTable();
          if (tbl.getNumRows() < 2) continue;
          const h0 = tbl.getRow(0).getCell(0).getText().trim();
          if (!h0.includes('拠点・不在時間帯')) continue;

          for (let r = 1; r < tbl.getNumRows(); r++) {
            const row = tbl.getRow(r);
            if (row.getNumCells() < 4) continue;
            const c0 = row.getCell(0).getText().trim();
            const c1 = row.getCell(1).getText().trim();
            const c2Raw = row.getCell(2).getText().trim();
            const c2Main = c2Raw.split('\n').filter(function(l) { return !l.trim().startsWith('→'); }).join(' ').trim();
            const c3 = row.getCell(3).getText().trim();

            if (!c2Main && !c3) continue;

            const c0Lines = c0.split('\n');
            const key = getSlotCarryOverKey3_(curDate, c0Lines[0] || '', c0Lines.slice(1).join(' '));

            const prevPlanMap = {};
            c1.split('\n').forEach(function(line) {
              const pm = line.trim().match(/^【案([123１２３])】(.*)$/);
              if (pm) {
                const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
                prevPlanMap[pNum] = pm[2].trim();
              }
            });

            map.set(key, {
              rawDecision: c2Main,
              assignee: c3,
              prevPlanMap: prevPlanMap
            });
          }
        }
      }
    } catch (e) {}
  });

  return map;
}

function resolveCarriedOverDecisionForSlot3_(prevEntry, currentPlanObj, origSlotOpt, mailHistoryIndexOpt) {
  let dec = prevEntry ? String(prevEntry.rawDecision || '').trim() : '';
  const assignee = prevEntry ? String(prevEntry.assignee || '').trim() : '';
  let comment = '';

  const memoMatch = dec.match(/^(.*?)\s*[（\(]メモ[：:](.+?)[）\)]$/);
  if (memoMatch) {
    dec = memoMatch[1].trim();
    comment = memoMatch[2].trim();
  }

  let spTag = '';
  const spMatch = dec.match(/(【特別時給.*?】)/);
  if (spMatch) {
    spTag = spMatch[1];
    dec = dec.replace(/【特別時給.*?】/g, '').trim();
  }

  const newPlans = {
    '1': String(currentPlanObj.plan1 || ''),
    '2': String(currentPlanObj.plan2 || ''),
    '3': String(currentPlanObj.plan3 || '')
  };

  const planDocsByNum = {
    '1': (typeof extractCandidateDoctorsFromPlanText2a_ === 'function') ? extractCandidateDoctorsFromPlanText2a_(newPlans['1']) : [],
    '2': (typeof extractCandidateDoctorsFromPlanText2a_ === 'function') ? extractCandidateDoctorsFromPlanText2a_(newPlans['2']) : [],
    '3': (typeof extractCandidateDoctorsFromPlanText2a_ === 'function') ? extractCandidateDoctorsFromPlanText2a_(newPlans['3']) : []
  };

  let mappedDecisionBase = '';

  if (dec) {
    if (/^(ステイ|特別時給.*ステイ|全案|バックアップ医師に勤務依頼)$/.test(dec) || !/案[123１２３]/.test(dec)) {
      mappedDecisionBase = dec;
    } else {
      const tokens = dec.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(function(t) { return t.trim(); }).filter(Boolean);
      const mappedTokens = [];

      tokens.forEach(function(tok) {
        const m = tok.match(/^(?:【?案)?([123１２３])(?:】)?(?:\(([^)]+)\))?$/);
        if (!m) {
          mappedTokens.push(tok);
          return;
        }
        const oldNum = m[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
        const subSpec = m[2] ? m[2].trim() : '';
        const oldPlanText = (prevEntry && prevEntry.prevPlanMap && prevEntry.prevPlanMap[oldNum]) ? prevEntry.prevPlanMap[oldNum] : '';

        if (!oldPlanText) {
          if (newPlans[oldNum]) mappedTokens.push(subSpec ? ('案' + oldNum + '(' + subSpec + ')') : ('案' + oldNum));
          return;
        }

        const oldDocs = (typeof extractCandidateDoctorsFromPlanText2a_ === 'function')
          ? extractCandidateDoctorsFromPlanText2a_(oldPlanText)
          : [];

        let matchedNewNum = '';
        if (oldDocs.length > 0) {
          ['1', '2', '3'].forEach(function(nNum) {
            if (matchedNewNum || !newPlans[nNum]) return;
            const newDocs = planDocsByNum[nNum] || [];
            if (oldDocs.some(function(od) { return newDocs.includes(od); })) {
              matchedNewNum = nNum;
            }
          });
        }

        if (matchedNewNum) {
          const mappedLabel = subSpec ? ('案' + matchedNewNum + '(' + subSpec + ')') : ('案' + matchedNewNum);
          if (!mappedTokens.includes(mappedLabel)) mappedTokens.push(mappedLabel);
        }
      });

      mappedDecisionBase = mappedTokens.join(', ');
    }
  }

  const rawCurrentMemo = origSlotOpt ? String(origSlotOpt.currentMemo || '').trim() : '';
  const officialNames = (typeof buildMasterNameCorrectionMap2a_ === 'function' && origSlotOpt)
    ? buildMasterNameCorrectionMap2a_([origSlotOpt])
    : [];
  const correctedCurrentMemo = (typeof correctDoctorNamesInMemoText_ === 'function')
    ? correctDoctorNamesInMemoText_(rawCurrentMemo, officialNames)
    : rawCurrentMemo;

  // ★修正：「中﨑寿隆先生」の「﨑」や「髙」などCJK拡張漢字も途切れずフルネーム抽出する正規表現
  const memoDocs = [];
  if (correctedCurrentMemo && correctedCurrentMemo !== '-') {
    const memoRe = /([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})先生/g;
    let mm;
    while ((mm = memoRe.exec(correctedCurrentMemo)) !== null) {
      if (!memoDocs.includes(mm[1])) memoDocs.push(mm[1]);
    }
  }

  const matchedPlanNumsFromMemo = [];
  const planLabelByMemoDoc = {};

  memoDocs.forEach(function(mDoc) {
    const mKey = normKeyForDoc_(mDoc);
    ['1', '2', '3'].forEach(function(pNum) {
      const pDocs = planDocsByNum[pNum] || [];
      if (pDocs.some(function(pd) { return normKeyForDoc_(pd) === mKey; })) {
        planLabelByMemoDoc[mDoc] = '案' + pNum;
        if (!matchedPlanNumsFromMemo.includes(pNum)) {
          matchedPlanNumsFromMemo.push(pNum);
        }
      }
    });
  });

  if ((!mappedDecisionBase || mappedDecisionBase === '全案') && matchedPlanNumsFromMemo.length > 0) {
    matchedPlanNumsFromMemo.sort();
    mappedDecisionBase = matchedPlanNumsFromMemo.map(function(n) { return '案' + n; }).join(', ');
  }

  const targetDocsForResult = [];
  memoDocs.forEach(function(d) {
    if (!targetDocsForResult.includes(d)) targetDocsForResult.push(d);
  });

  if (mappedDecisionBase && /案[123]/.test(mappedDecisionBase)) {
    const chosenNums = mappedDecisionBase.match(/[123]/g) || [];
    chosenNums.forEach(function(n) {
      // 「案1(②灘理絵先生)」のように特定の1名が指定されている場合はその医師のみを追加
      const subMatch = mappedDecisionBase.match(new RegExp('案' + n + '\\(([^)]+)\\)'));
      const subSpec = subMatch ? subMatch[1].replace(/^[①②12]\s*/, '').replace(/先生$/g, '').trim() : '';

      (planDocsByNum[n] || []).forEach(function(d) {
        if (subSpec && !d.includes(subSpec) && !subSpec.includes(d)) return;
        if (!targetDocsForResult.includes(d)) {
          targetDocsForResult.push(d);
          planLabelByMemoDoc[d] = '案' + n;
        }
      });
    });
  }

  let resultArrowLines = [];
  if (targetDocsForResult.length > 0 && typeof evaluateDoctorAdjustmentStatusObj1b2_ === 'function' && mailHistoryIndexOpt) {
    const cleanC = cleanClinicBaseName_(currentPlanObj.clinicAndDept || '');
    const statusObjs = targetDocsForResult.map(function(docName) {
      return evaluateDoctorAdjustmentStatusObj1b2_(
        docName,
        planLabelByMemoDoc[docName] || '',
        currentPlanObj.date,
        cleanC,
        mailHistoryIndexOpt
      );
    });
    resultArrowLines = formatCompactAdjustmentLines1b2_(statusObjs);
  } else if (correctedCurrentMemo && correctedCurrentMemo !== '-' && targetDocsForResult.length === 0) {
    resultArrowLines.push('→ 📋' + correctedCurrentMemo.replace(/\r?\n/g, ' / '));
  }

  let finalHeaderDec = mappedDecisionBase
    ? (mappedDecisionBase + spTag + (comment ? ('（メモ：' + comment + '）') : ''))
    : (comment ? ('（メモ：' + comment + '）') : '');

  const fullDecisionCellText = [finalHeaderDec].concat(resultArrowLines).filter(Boolean).join('\n');

  return {
    decision: fullDecisionCellText,
    assignee: assignee,
    comment: comment
  };
}

function getOrResetDailyMeetingDocInFolder_(title) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SINGLE_MTG_DOC_ID';
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  let folder = null;
  try {
    folder = DriveApp.getFolderById(DOC_OUTPUT_FOLDER_ID);
  } catch (e) {
    Logger.log('⚠️ 指定フォルダの取得に失敗したためマイドライブ直下を使用します: ' + e.message);
  }

  let targetDoc = null;
  let targetDocId = '';
  const sameDayFiles = [];

  if (folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      const name = f.getName();
      if (name.includes('超募集対策MTG_' + todayStr)) {
        sameDayFiles.push(f);
      }
    }
  }

  if (sameDayFiles.length > 0) {
    sameDayFiles.sort(function(a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
    const keepFile = sameDayFiles[0];
    targetDocId = keepFile.getId();
    targetDoc = DocumentApp.openById(targetDocId);
    targetDoc.setName(title);

    for (let i = 1; i < sameDayFiles.length; i++) {
      try { sameDayFiles[i].setTrashed(true); } catch (e) {}
    }
  }

  if (!targetDoc) {
    const savedId = props.getProperty(propKey);
    if (savedId) {
      try {
        const savedFile = DriveApp.getFileById(savedId);
        if (!savedFile.isTrashed() && savedFile.getName().includes('超募集対策MTG_' + todayStr)) {
          targetDocId = savedId;
          targetDoc = DocumentApp.openById(targetDocId);
          targetDoc.setName(title);
          if (folder) moveFileToFolderSafely_(savedFile, folder);
        }
      } catch (e) {}
    }
  }

  if (!targetDoc) {
    targetDoc = DocumentApp.create(title);
    targetDocId = targetDoc.getId();
    const newFile = DriveApp.getFileById(targetDocId);
    if (folder) {
      moveFileToFolderSafely_(newFile, folder);
    }
  }

  props.setProperty(propKey, targetDocId);
  return targetDoc;
}

function moveFileToFolderSafely_(file, targetFolder) {
  try {
    file.moveTo(targetFolder);
  } catch (e) {
    try {
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e2) {}
  }
}