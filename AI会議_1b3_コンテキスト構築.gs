/**
 * ============================================================================
 * AI会議_1b3_コンテキスト構築.gs
 * （確定シフト充足済み枠の完全除外・終日0人枠のみ補完・本日分除外・玉突き不在防止版）
 * ============================================================================
 */

function normalizeDateYmd1b3_(rawDateStr) {
  const m = String(rawDateStr || '').trim().match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return String(rawDateStr || '').trim();
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 「終日」や「9:00～21:00」などの表記を標準の「09:00-13:00」形式へ正規化・分解するヘルパー
 */
function normalizeAndExpandAbsenceTime1b3_(rawTimeStr, clinicName) {
  const s = String(rawTimeStr || '').normalize('NFKC').trim();
  const isKitakasai = String(clinicName || '').includes('北葛西');
  const nightSlot = isKitakasai ? '18:00-20:00' : '18:00-21:00';

  if (s.includes('終日')) {
    return ['09:00-13:00', '15:00-18:00', nightSlot];
  }

  const m = s.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
  if (!m) return s ? [s] : [];

  const sh = Number(m[1]);
  const sm = ('0' + m[2]).slice(-2);
  const eh = Number(m[3]);
  const em = ('0' + m[4]).slice(-2);
  const formatted = ('0' + sh).slice(-2) + ':' + sm + '-' + ('0' + eh).slice(-2) + ':' + em;

  if (sh <= 10 && eh >= 20) {
    return ['09:00-13:00', '15:00-18:00', ('0' + eh).slice(-2) + ':' + em === '20:00' ? '18:00-20:00' : '18:00-21:00'];
  }
  if (sh <= 10 && eh === 18) {
    return ['09:00-13:00', '15:00-18:00'];
  }
  if (sh === 15 && eh >= 20) {
    return ['15:00-18:00', '18:00-' + ('0' + eh).slice(-2) + ':' + em];
  }

  return [formatted];
}

/**
 * ★重要ガード：『最新_当日シフト』（確定シフト）を照合し、
 * 指定された日付・拠点・診療科・時間帯（absTimeStr）において「すでに確定医師がいて不在が解消しているか」を判定する関数
 * 戻り値: true = 確定医師がおり不在解消済み（書き出し不要） / false = 本当に医師不在（0人または部分不在）
 */
function isSlotAlreadyCoveredByActiveShift1b3_(dateStr, rawClinic, absTimeStr, activeShifts) {
  const d = normalizeDateYmd1b3_(dateStr);
  const cleanC = cleanClinicBaseName_(rawClinic);
  const isInternal = String(rawClinic || '').includes('内科');

  // 当該日・当該拠点・当該診療科の確定シフト行を検索
  const matchingRows = (activeShifts || []).filter(function(s) {
    if (normalizeDateYmd1b3_(s[0]) !== d) return false;
    if (cleanClinicBaseName_(s[2]) !== cleanC) return false;
    const sIsInternal = String(s[2] || '').includes('内科') || String(s[3] || '').includes('内科');
    return sIsInternal === isInternal;
  });

  // 『最新_当日シフト』に行自体が存在しない場合は「終日0人（未充足）」なのでカバーされていない(false)
  if (matchingRows.length === 0) return false;

  const timeInfo = parseAbsenceTimeline_(absTimeStr);
  const targetBands = timeInfo.targetBands || ['A'];
  const row = matchingRows[0];

  // 対象となるすべての時間帯（A/B/C）で確定医師がいて、かつ不在時間帯がカバーされているかチェック
  for (let i = 0; i < targetBands.length; i++) {
    const band = targetBands[i];
    const countCol = band === 'A' ? 4 : (band === 'B' ? 5 : 6);
    const docsCol = band === 'A' ? 7 : (band === 'B' ? 8 : 9);

    const cntVal = parseInt(String(row[countCol] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const docsStr = String(row[docsCol] || '').trim();

    // 人数が0、または医師欄が空欄・「未充足」なら、本当に医師不在なので false（カバーされていない）
    if (cntVal === 0 || !docsStr || docsStr === '未充足' || docsStr === '-' || docsStr === 'なし') {
      return false;
    }

    // 医師欄に時刻指定（例: "〇〇(15:00-17:15)" のみの早上がり等）があり、absTimeStr（例: 17:15-18:00）をカバーできていない場合もチェック
    const tokens = splitDoctorTokensSafely_(docsStr);
    let bandFullyCovered = false;

    for (let j = 0; j < tokens.length; j++) {
      const tok = tokens[j];
      if (!tok || tok.includes('未充足')) continue;
      const tm = tok.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
      if (!tm) {
        // 時刻のカッコ書きがない通常の確定医師がいれば、その帯はフルタイムでカバーされている
        bandFullyCovered = true;
        break;
      } else {
        const docStart = Number(tm[1]) * 60 + Number(tm[2]);
        const docEnd = Number(tm[3]) * 60 + Number(tm[4]);
        // 確定医師の勤務時間が、今回の募集枠（timeInfo.startMin〜endMin）を包含していればカバー済み
        if (docStart <= timeInfo.startMin && docEnd >= timeInfo.endMin) {
          bandFullyCovered = true;
          break;
        }
      }
    }

    if (!bandFullyCovered) {
      return false; // 早上がり等で一部時間帯が空いている場合は不在未解消
    }
  }

  // すべての対象時間帯に確定医師が配置済み＝「医師不在は解消している」
  return true;
}

/**
 * 『最新_不在枠』をベースにしつつ、
 * 1) 確定シフトですでに医師が配置され不在が解消している枠（東品川17:15-18:00等）を除外し、
 * 2) 終日不在等で『最新_不在枠』から漏れている本当に0人の枠（11/8東品川・錦糸町等）だけを補完する関数
 */
function recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilledRows, activeShifts, spotCandidates, tomorrowYmd, targetEndYmd) {
  const combined = [];
  const seenKeys = new Set();

  const registerRowIfTrulyUnfilled = function(rowArr, isFromPrimaryUnfilledSheet) {
    const d = normalizeDateYmd1b3_(rowArr[0]);
    if (!d || d < tomorrowYmd) return;
    const rawClinic = String(rowArr[1] || '').trim();
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = rawClinic.includes('内科') ? '内科' : '小児科';
    const expandedTimes = normalizeAndExpandAbsenceTime1b3_(rowArr[2], cleanC);

    expandedTimes.forEach(function(tStr) {
      const key = d + '_' + cleanC + '_' + deptTag + '_' + tStr;
      if (seenKeys.has(key)) return;

      // ★『最新_不在枠』以外のサブソース（mijyusoku募集リスト等）から補完する場合は、
      // 『最新_当日シフト』の確定シフトにすでに医師がいれば「不在解消済み」なので絶対に書き出さない！
      // また、『最新_不在枠』自体に残っている枠でも、確定シフトでフルカバーされていれば除外する
      if (isSlotAlreadyCoveredByActiveShift1b3_(d, rawClinic, tStr, activeShifts)) {
        return;
      }

      seenKeys.add(key);
      const copy = rowArr.slice();
      copy[0] = d;
      copy[2] = tStr;
      combined.push(copy);
    });
  };

  // 1. 既存の『最新_不在枠』を登録（確定シフトで解消済みの枠は除外）
  (rawUnfilledRows || []).forEach(function(r) {
    registerRowIfTrulyUnfilled(r, true);
  });

  // 2. 『最新_当日シフト』で医師が空欄・「未充足」・0人になっている枠を復元
  (activeShifts || []).forEach(function(s) {
    const d = normalizeDateYmd1b3_(s[0]);
    if (!d || d < tomorrowYmd || d > targetEndYmd) return;
    const rawC = String(s[2] || '').trim();
    const dept = String(s[3] || '').trim();
    if (!rawC || rawC.includes('バックアップ')) return;
    const cleanC = cleanClinicBaseName_(rawC);
    const clinicLabel = dept.includes('内科') ? (cleanC + '（内科）') : cleanC;

    const parseDocCount = function(cntVal, docsVal) {
      const dStr = String(docsVal || '').trim();
      if (!dStr || dStr === '未充足' || dStr === '-' || dStr === 'なし') return 0;
      const n = parseInt(String(cntVal || '').replace(/[^0-9]/g, ''), 10);
      return isNaN(n) ? 1 : n;
    };

    const cntA = parseDocCount(s[4], s[7]);
    const cntB = parseDocCount(s[5], s[8]);
    const cntC = parseDocCount(s[6], s[9]);
    const docsA = cntA > 0 ? String(s[7] || '') : 'なし';
    const docsB = cntB > 0 ? String(s[8] || '') : 'なし';
    const docsC = cntC > 0 ? String(s[9] || '') : 'なし';

    if (cntA === 0) {
      registerRowIfTrulyUnfilled([d, clinicLabel, '09:00-13:00', '14,000円', 'なし', docsB, '', '', '', '', '未充足', '', '', ''], false);
    }
    if (cntB === 0) {
      registerRowIfTrulyUnfilled([d, clinicLabel, '15:00-18:00', '13,000円', docsA, docsC, '', '', '', '', '未充足', '', '', ''], false);
    }
    if (cntC === 0 && !dept.includes('内科')) {
      const nightTime = cleanC.includes('北葛西') ? '18:00-20:00' : '18:00-21:00';
      registerRowIfTrulyUnfilled([d, clinicLabel, nightTime, '15,000円', docsB, 'なし', '', '', '', '', '未充足', '', '', ''], false);
    }
  });

  // 3. 『BQシフト状況パック(mijyusoku)』や『最新_個別打診候補』からは、
  // 「『最新_当日シフト』でも本当に医師がいない（未解消の）枠」のみを補完（確定シフトで解消済みの募集枠は100%スキップ）
  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData && Array.isArray(bqData.mijyusoku)) {
        bqData.mijyusoku.forEach(function(m) {
          const d = normalizeDateYmd1b3_(m.d || m.date);
          if (!d || d < tomorrowYmd || d > targetEndYmd) return;
          const rawC = String(m.c || m.clinic || '').trim();
          const rawT = String(m.t || m.time || '').trim();
          if (!rawC || !rawT || rawC.includes('バックアップ')) return;
          const wage = String(m.w || m.wage || m.listedWage || '13,000円');
          const spWage = String(m.sp || m.specialWage || '');

          registerRowIfTrulyUnfilled([
            d, rawC, rawT, wage, m.prev || 'なし', m.next || 'なし', '', '', '', '', '未充足', spWage, '', ''
          ], false);
        });
      }
    } catch (e) {}
  }

  combined.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    if (a[1] !== b[1]) return a[1] < b[1] ? -1 : 1;
    return String(a[2] || '') < String(b[2] || '') ? -1 : 1;
  });

  return combined;
}

function evaluateValidAdjacentDocsForSlot_(docsStr, actionType, dateStr, dow, targetClinic, targetArea, timeInfo, masters) {
  const tokens = splitDoctorTokensSafely_(docsStr);
  const validTokens = [];
  const validDocKeys = new Set();

  tokens.forEach(function(tok) {
    const k = normKeyForDoc_(tok);
    if (!k || masters.absoluteNgDocSet.has(k)) return;
    if (masters.sbuSet.has(k)) return;
    if (masters.streakNgMap.has(k) && masters.streakNgMap.get(k).includes(dateStr)) return;

    const tm = tok.match(/(\d{1,2}):(\d{2})\s*[\-〜～]\s*(\d{1,2}):(\d{2})/);
    if (tm) {
      const dStart = Number(tm[1]) * 60 + Number(tm[2]);
      const dEnd = Number(tm[3]) * 60 + Number(tm[4]);
      if (actionType === '延長' && dEnd < timeInfo.startMin) {
        if (!(dEnd === 780 && timeInfo.startMin === 900)) return;
      }
      if (actionType === '早出' && dStart > timeInfo.endMin) {
        if (!(timeInfo.endMin === 780 && dStart === 900)) return;
      }
    }

    if (!isDoctorCompatibleWithSlot_(k, dow, targetClinic, targetArea, timeInfo, actionType, masters)) return;

    const headPart = String(tok).replace(/\[.*$/g, '');
    const isFt = !headPart.includes('非常勤') && headPart.includes('常勤');

    if (isFt) {
      const condStr = masters.ftAllowCondMap.get(k) || '';
      if (masters.ftAllowDocSet.has(k) && headPart.includes('所定休出') && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
        validTokens.push(k + '(常勤・所定休出)');
        validDocKeys.add(k);
      }
      return;
    }

    validTokens.push(k + '(非常勤)');
    validDocKeys.add(k);
  });

  return {
    docs: validTokens.length > 0 ? validTokens.join(' / ') : 'なし',
    status: validTokens.length > 0 ? (actionType + '打診可能') : '',
    docKeySet: validDocKeys
  };
}

function makeNormalizedSpotKey_(dateStr, rawClinicStr, absTimeStr) {
  const cleanD = normalizeDateYmd1b3_(dateStr);
  const cleanC = cleanClinicBaseName_(rawClinicStr);
  const dept = String(rawClinicStr || '').includes('内科') ? '内科' : '小児科';
  return cleanD + '_' + cleanC + '_' + dept + '_' + String(absTimeStr || '').trim();
}

function canDoctorLeaveSourceWithoutCreatingAbsence_(docKey, srcRow, targetBands, pairEval) {
  const countByBand = {
    'A': Number(srcRow[4] || 0),
    'B': Number(srcRow[5] || 0),
    'C': Number(srcRow[6] || 0)
  };
  const docsByBand = {
    'A': splitDoctorTokensSafely_(srcRow[7]).map(normKeyForDoc_),
    'B': splitDoctorTokensSafely_(srcRow[8]).map(normKeyForDoc_),
    'C': splitDoctorTokensSafely_(srcRow[9]).map(normKeyForDoc_)
  };

  for (let i = 0; i < targetBands.length; i++) {
    const tb = targetBands[i];
    if (countByBand[tb] < 2 || !docsByBand[tb].includes(docKey)) {
      return false;
    }
  }

  const allBands = ['A', 'B', 'C'];
  for (let i = 0; i < allBands.length; i++) {
    const b = allBands[i];
    if (targetBands.includes(b)) continue;
    if (!docsByBand[b].includes(docKey)) continue;

    if ((b === 'C' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('C'))) {
      if (countByBand[b] < 2) {
        return false;
      }
    }

    if ((b === 'A' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('A'))) {
      if (pairEval.estMinutes > 60 || (pairEval.rank !== 'S' && pairEval.rank !== 'A')) {
        if (countByBand[b] < 2) return false;
      }
    }
  }

  return true;
}

function collectCompactContextForAi_(ss) {
  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrowDt = new Date(todayMidnight.getTime() + 1 * 86400000);

  const tomorrowYmd = Utilities.formatDate(tomorrowDt, 'Asia/Tokyo', 'yyyy/MM/dd');
  const twoWeeksYmd = Utilities.formatDate(new Date(tomorrowDt.getTime() + 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const oneMonthYmd = Utilities.formatDate(new Date(tomorrowDt.getTime() + 31 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  const getRows = function(name) {
    const sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() >= 2) ? sh.getDataRange().getDisplayValues().slice(1) : [];
  };

  const getRecentRows = function(name, maxCount) {
    const sh = ss.getSheetByName(name);
    if (!sh) return [];
    const lastR = sh.getLastRow();
    if (lastR < 2) return [];
    const startR = Math.max(2, lastR - maxCount + 1);
    const numCols = Math.min(8, sh.getLastColumn());
    return sh.getRange(startR, 1, lastR - startR + 1, numCols).getDisplayValues();
  };

  const activeShifts = getRows('最新_当日シフト').map(function(r) {
    const copy = r.slice();
    copy[0] = normalizeDateYmd1b3_(r[0]);
    return copy;
  });
  const streakWarnings = getRows('最新_連勤注意リスト');
  const spotCandidates = getRows('最新_個別打診候補');
  const kbRows = getRecentRows('ナレッジシート', 120);

  const rawUnfilledRows = getRows('最新_不在枠').map(function(r) {
    const copy = r.slice();
    copy[0] = normalizeDateYmd1b3_(r[0]);
    return copy;
  });

  // ★確定シフトで医師がいる枠は「解消済み」として除外し、本当に0人の不在枠のみを抽出・復元
  const recoveredUnfilledAll = recoverMissingFullDayOrUnfilledRows1b3_(
    rawUnfilledRows,
    activeShifts,
    spotCandidates,
    tomorrowYmd,
    twoWeeksYmd
  );

  const unfilledAll = recoveredUnfilledAll.filter(function(r) { return r[0] && r[0] >= tomorrowYmd; });
  const within2W = unfilledAll.filter(function(r) { return r[0] <= twoWeeksYmd; });
  const targetEndYmd = (within2W.length < 10 && unfilledAll.length > within2W.length) ? oneMonthYmd : twoWeeksYmd;
  const targetUnfilled = unfilledAll.filter(function(r) { return r[0] <= targetEndYmd; });

  const periodStart = targetUnfilled.length > 0 ? targetUnfilled[0][0] : tomorrowYmd;
  const periodEnd = targetUnfilled.length > 0 ? targetUnfilled[targetUnfilled.length - 1][0] : targetEndYmd;

  const dominantMatrix = loadDominantMatrixFromDoc_();
  const masters = buildMasterAndConfirmedShiftIndex_(ss, activeShifts, streakWarnings);

  const latestReplyMap = new Map();
  kbRows.forEach(function(r) {
    const dKey = normKeyForDoc_(r[1]);
    if (!dKey) return;
    const rawFlag = String(r[7] || '').trim();
    const rawPitch = String(r[3] || '').trim();

    if (rawFlag.includes('一般連絡')) return;
    if (!rawPitch.includes('【スタッフ打診】') && !rawFlag.includes('シフト調整')) return;

    const dtStr = String(r[6] || '').slice(5, 10);
    if (rawFlag.includes('承諾') || rawFlag.includes('成立')) {
      latestReplyMap.set(dKey, (dtStr ? dtStr + 'の別枠打診：承諾実績あり' : '直近打診：承諾実績あり'));
    } else if (rawFlag.includes('辞退')) {
      latestReplyMap.set(dKey, (dtStr ? dtStr + 'の別枠打診：予定合わず辞退' : '直近打診：辞退履歴あり'));
    }
  });

  const spotMap = new Map();
  spotCandidates.forEach(function(r) {
    const cleanC = cleanClinicBaseName_(r[2]);
    const expandedTimes = normalizeAndExpandAbsenceTime1b3_(r[3], cleanC);
    const dKey = normKeyForDoc_(r[5]);
    if (!dKey || masters.absoluteNgDocSet.has(dKey) || masters.sbuSet.has(dKey)) return;

    const sameDowSlotNum = Number(String(r[8] || '0').replace(/[^0-9]/g, '')) || 0;
    const subCountNum = Number(String(r[9] || '0').replace(/[^0-9]/g, '')) || 0;
    const totalCountNum = Number(String(r[10] || '0').replace(/[^0-9]/g, '')) || 0;

    if (sameDowSlotNum === 0 && subCountNum === 0 && totalCountNum === 0) return;

    let trackRecordLabel = '';
    if (sameDowSlotNum > 0) {
      trackRecordLabel = '同曜日同枠' + sameDowSlotNum + '回実績';
    } else if (totalCountNum > 0) {
      trackRecordLabel = '同拠点通算' + totalCountNum + '回実績';
    } else {
      trackRecordLabel = '同拠点別枠' + subCountNum + '回実績';
    }

    expandedTimes.forEach(function(tStr) {
      const key = makeNormalizedSpotKey_(r[0], r[2], tStr);
      if (!spotMap.has(key)) spotMap.set(key, []);
      spotMap.get(key).push({
        rank: r[4],
        doctor: dKey,
        sameDowSlotNum: sameDowSlotNum,
        totalCountNum: totalCountNum,
        trackRecordLabel: trackRecordLabel,
        rawNote: String(r[13] || '').trim(),
        latestReplyDetail: latestReplyMap.get(dKey) || ''
      });
    });
  });

  const targetDatesSet = new Set();

  const slotsContext = targetUnfilled.map(function(u, idx) {
    const dateStr = String(u[0] || '').trim();
    const clinic = String(u[1] || '').trim();
    const cleanTargetClinic = cleanClinicBaseName_(clinic);
    const absTime = String(u[2] || '').trim();
    const isInternalMed = clinic.includes('内科');
    const area = inferAreaForDoc_(clinic);
    const dow = getDowForDoc_(dateStr);
    const timeInfo = parseAbsenceTimeline_(absTime);
    targetDatesSet.add(dateStr);

    const prevChecked = evaluateValidAdjacentDocsForSlot_(u[4] || 'なし', '延長', dateStr, dow, cleanTargetClinic, area, timeInfo, masters);
    const nextChecked = evaluateValidAdjacentDocsForSlot_(u[5] || 'なし', '早出', dateStr, dow, cleanTargetClinic, area, timeInfo, masters);

    const dayShifts = activeShifts.filter(function(s) { return String(s[0] || '').trim() === dateStr; });
    const twoDocScoredList = [];
    const oneDocScoredList = [];
    const seenTwoDoc = new Set();
    const seenOneDoc = new Set();
    let backupDocs = '';

    if (!isInternalMed) {
      dayShifts.forEach(function(s) {
        const sClinic = String(s[2] || '').trim();
        const cleanSrcClinic = cleanClinicBaseName_(sClinic);
        if (!cleanSrcClinic) return;

        if (sClinic.includes('バックアップ')) {
          const bDocs = [s[7], s[8], s[9]].filter(Boolean).join(' / ');
          if (bDocs && bDocs !== '未充足') backupDocs = bDocs;
          return;
        }
        if (cleanSrcClinic === cleanTargetClinic || String(s[3] || '').includes('内科')) return;

        const pairEval = evaluateClinicPairMobility_(cleanTargetClinic, cleanSrcClinic, dominantMatrix);
        if (pairEval.rank === 'NG' || pairEval.score <= 0) return;

        const sArea = inferAreaForDoc_(sClinic);
        const primaryBand = timeInfo.targetBands[0];
        const countCol = primaryBand === 'A' ? 4 : (primaryBand === 'B' ? 5 : 6);
        const docsCol = primaryBand === 'A' ? 7 : (primaryBand === 'B' ? 8 : 9);
        const docCount = Number(s[countCol] || 0);
        const rawDocsStr = String(s[docsCol] || '');
        if (!rawDocsStr || rawDocsStr === '未充足' || docCount < 1) return;

        const tokens = splitDoctorTokensSafely_(rawDocsStr);
        const movableForTwoDoc = [];
        const movableForOneDoc = [];
        let bestScore = 0;

        tokens.forEach(function(tok) {
          const k = normKeyForDoc_(tok);
          if (!k || k === '未充足' || masters.absoluteNgDocSet.has(k)) return;
          if (masters.streakNgMap.has(k) && masters.streakNgMap.get(k).includes(dateStr)) return;

          const headPart = String(tok).replace(/\[.*$/g, '');
          const isFt = !headPart.includes('非常勤') && headPart.includes('常勤');
          if (isFt) {
            const condStr = masters.ftAllowCondMap.get(k) || '';
            if (!masters.ftAllowDocSet.has(k) || !headPart.includes('所定休出') || !isFtConditionCompatibleWithTarget_(condStr, cleanTargetClinic, area)) {
              return;
            }
          }

          if (!isDoctorCompatibleWithSlot_(k, dow, cleanTargetClinic, area, timeInfo, '移動', masters)) return;

          const isSbu = masters.sbuSet.has(k);
          const isNeg = masters.negotiableMap.has(k);
          const isSpecialMovable = isSbu || isNeg || isFt;

          if (pairEval.rank === 'B' && !isSpecialMovable) return;

          const docScore = pairEval.score + (isSbu ? 15 : (isNeg ? 10 : 0));
          const label = isSbu ? (k + '[★SBU]') : (isFt ? (k + '[常勤・所定休出]') : k);

          if (docCount >= 2) {
            if (canDoctorLeaveSourceWithoutCreatingAbsence_(k, s, timeInfo.targetBands, pairEval)) {
              movableForTwoDoc.push(label);
              if (docScore > bestScore) bestScore = docScore;
            }
          } else if (docCount === 1 && isSpecialMovable) {
            movableForOneDoc.push(label);
            if (docScore > bestScore) bestScore = docScore;
          }
        });

        if (movableForTwoDoc.length > 0 && !seenTwoDoc.has(cleanSrcClinic)) {
          seenTwoDoc.add(cleanSrcClinic);
          twoDocScoredList.push({
            score: bestScore,
            item: {
              clinic: cleanSrcClinic,
              area: sArea,
              priorityRank: pairEval.rank,
              estMinutes: pairEval.estMinutes,
              movableDoctors: movableForTwoDoc
            }
          });
        } else if (movableForOneDoc.length > 0 && !seenOneDoc.has(cleanSrcClinic)) {
          seenOneDoc.add(cleanSrcClinic);
          oneDocScoredList.push({
            score: bestScore,
            item: {
              clinic: cleanSrcClinic,
              area: sArea,
              priorityRank: pairEval.rank,
              estMinutes: pairEval.estMinutes,
              movableDoctors: movableForOneDoc
            }
          });
        }
      });
    }

    twoDocScoredList.sort(function(a, b) { return b.score - a.score; });
    oneDocScoredList.sort(function(a, b) { return b.score - a.score; });

    const rawSpots = (spotMap.get(makeNormalizedSpotKey_(dateStr, clinic, absTime)) || []).slice();
    rawSpots.sort(function(a, b) {
      if (b.sameDowSlotNum !== a.sameDowSlotNum) return b.sameDowSlotNum - a.sameDowSlotNum;
      return b.totalCountNum - a.totalCountNum;
    });

    const verifiedSpots = [];
    const dayConfirmedMap = masters.confirmedDocScheduleByDate.get(dateStr) || new Map();

    rawSpots.forEach(function(sp) {
      const dKey = sp.doctor;
      if (prevChecked.docKeySet.has(dKey) || nextChecked.docKeySet.has(dKey)) return;
      if (masters.streakNgMap.has(dKey) && masters.streakNgMap.get(dKey).includes(dateStr)) return;
      if (!isDoctorCompatibleWithSlot_(dKey, dow, cleanTargetClinic, area, timeInfo, '個別', masters)) return;

      const confirmedList = dayConfirmedMap.get(dKey) || [];
      let isImpossibleSchedule = false;
      const otherConfirmedLabels = [];

      for (let i = 0; i < confirmedList.length; i++) {
        const cs = confirmedList[i];
        if (Math.max(cs.startMin, timeInfo.startMin) < Math.min(cs.endMin, timeInfo.endMin)) {
          isImpossibleSchedule = true;
          break;
        }
        if (cs.clinic !== cleanTargetClinic) {
          const gapMin = (cs.endMin <= timeInfo.startMin)
            ? (timeInfo.startMin - cs.endMin)
            : (cs.startMin - timeInfo.endMin);
          const pairMobility = evaluateClinicPairMobility_(cleanTargetClinic, cs.clinic, dominantMatrix);

          if ((pairMobility.rank !== 'S' && pairMobility.rank !== 'A') ||
              pairMobility.estMinutes > 60 ||
              (gapMin - pairMobility.estMinutes) < 60) {
            isImpossibleSchedule = true;
            break;
          }
        }
        otherConfirmedLabels.push(cs.clinic + ' ' + cs.timeLabel);
      }

      if (isImpossibleSchedule) return;

      verifiedSpots.push({
        rank: sp.rank,
        doctor: dKey,
        trackRecordLabel: sp.trackRecordLabel,
        daySchedule: otherConfirmedLabels.length > 0 ? ('同日確定枠:' + otherConfirmedLabels.join('、')) : '当日空き',
        masterNote: sp.rawNote,
        latestReplyDetail: sp.latestReplyDetail
      });
    });

    return {
      slotId: idx + 1,
      date: dateStr,
      area: area,
      clinic: clinic,
      dept: isInternalMed ? '内科' : '小児科',
      absenceTime: absTime,
      listedWage: u[3] || '13,000円',
      specialWage: u[11] || '',
      assignee: String(u[9] || '').trim(),
      prevDocs: prevChecked.docs,
      prevExtStatus: prevChecked.status,
      nextDocs: nextChecked.docs,
      nextEarlyStatus: nextChecked.status,
      currentMemo: u[12] || '',
      backupDocs: backupDocs,
      twoDocBreakdownCandidates: twoDocScoredList.slice(0, 4).map(function(x) { return x.item; }),
      oneDocParallelCandidates: oneDocScoredList.slice(0, 3).map(function(x) { return x.item; }),
      kinmukunCandidates: verifiedSpots.slice(0, 4)
    };
  });

  const streakList = streakWarnings.filter(function(w) {
    const docLabel = String(w[1] || '');
    const periodText = String(w[3] || '');
    const ngDateText = String(w[4] || '');
    if (!ngDateText || ngDateText === '-') return false;
    if (periodText.includes('中1日')) return false;
    if (docLabel.includes('(常勤)') && !docLabel.includes('手当対象')) return false;
    return Array.from(targetDatesSet).some(function(d) { return ngDateText.includes(d) || periodText.includes(d); });
  }).slice(0, 10).map(function(w) {
    return {
      licenseNo: w[0],
      doctor: w[1],
      streakDays: w[2],
      period: w[3],
      ngDates: w[4],
      note: w[5]
    };
  });

  const recentFlaggedKb = kbRows.filter(function(r) {
    const flag = String(r[7] || '');
    const body = String(r[2] || '');
    if (!flag.includes('★要確認')) return false;
    if (flag.includes('AIメモ反映済') || flag.includes('自動判定済')) return false;
    if (body.includes('退職することになりました')) return false;
    return true;
  }).slice(-5).map(function(r) {
    return {
      doctor: r[1],
      replySummary: String(r[2] || '').replace(/\n+/g, ' ').slice(0, 80),
      date: r[6],
      flag: r[7]
    };
  });

  return {
    periodStart: periodStart,
    periodEnd: periodEnd,
    unfilledCount: targetUnfilled.length,
    slotsContext: slotsContext,
    targetBundles: [],
    streakList: streakList,
    recentFlaggedKb: recentFlaggedKb
  };
}