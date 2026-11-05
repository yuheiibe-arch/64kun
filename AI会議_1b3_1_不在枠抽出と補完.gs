/**
 * ============================================================================
 * AI会議_1b3_1_不在枠抽出と補完.gs
 * （BQシフト状況パックの「bosyu」1,945件完全連動・複数重複時は「長い方の募集時間」を採用する版）
 * ============================================================================
 */

function normalizeDateYmd1b3_(rawDateStr) {
  const m = String(rawDateStr || '').trim().match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return String(rawDateStr || '').trim();
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

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

function isTimeRangeOverlapping1b3_(timeStr1, timeStr2) {
  if (timeStr1 === timeStr2) return true;
  const p1 = parseAbsenceTimeline_(timeStr1);
  const p2 = parseAbsenceTimeline_(timeStr2);
  return Math.max(p1.startMin, p2.startMin) < Math.min(p1.endMin, p2.endMin);
}

function formatMinToHHMM1b3_1_(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

/**
 * ★BQパックの「bosyu」（募集シフト）に同帯でより長い募集時間（例：15:00-18:00）がある場合は、
 * 不在時間（例：15:00-16:00）をその募集シフトの時間へ自動拡張する関数
 */
function expandToLongestRecruitSlotTime1b3_(dateStr, rawClinic, absTimeStr, recruitSlotsIndex, activeShifts) {
  const cleanC = cleanClinicBaseName_(rawClinic);
  const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';
  const isKitakasai = cleanC.includes('北葛西');
  const p = parseAbsenceTimeline_(absTimeStr);
  let bestStart = p.startMin;
  let bestEnd = p.endMin;

  // 1. 募集シフトインデックス（bqData.bosyu 等）に重なる枠があれば、最も長い範囲を採用
  const key = dateStr + '_' + cleanC + '_' + deptTag;
  const candidates = (recruitSlotsIndex && recruitSlotsIndex.get(key)) || [];
  candidates.forEach(function(candTime) {
    const cp = parseAbsenceTimeline_(candTime);
    if (Math.max(bestStart, cp.startMin) < Math.min(bestEnd, cp.endMin) &&
        cp.targetBands[0] === p.targetBands[0]) {
      bestStart = Math.min(bestStart, cp.startMin);
      bestEnd = Math.max(bestEnd, cp.endMin);
    }
  });

  // 2. まだ帯の途中（例：15:00-16:00）になっている場合、当日シフトを確認して空いている募集枠の上限まで広げる
  const primaryBand = p.targetBands[0] || 'A';
  const bandStart = primaryBand === 'A' ? 540 : (primaryBand === 'B' ? (isKitakasai ? 840 : 900) : 1080);
  const bandEnd = primaryBand === 'A' ? (isKitakasai ? 750 : 780) : (primaryBand === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

  if (bestStart > bandStart || bestEnd < bandEnd) {
    const matchingShiftRows = (activeShifts || []).filter(function(s) {
      if (normalizeDateYmd1b3_(s[0]) !== dateStr) return false;
      if (cleanClinicBaseName_(s[2]) !== cleanC) return false;
      const sDept = (String(s[2] || '').includes('内科') || String(s[3] || '').includes('内科')) ? '内科' : '小児科';
      return sDept === deptTag;
    });

    if (matchingShiftRows.length > 0) {
      const sRow = matchingShiftRows[0];
      const docsCol = primaryBand === 'A' ? 7 : (primaryBand === 'B' ? 8 : 9);
      const tokens = splitDoctorTokensSafely_(sRow[docsCol]);
      const allRowHint = [sRow[7], sRow[8], sRow[9]].join(' / ');

      const shiftIntervals = [];
      tokens.forEach(function(tok) {
        if (!tok || tok.includes('未充足')) return;
        if (typeof parseDocTokenWorkingMinutes1b3_ === 'function') {
          const hrs = parseDocTokenWorkingMinutes1b3_(tok, primaryBand, isKitakasai, allRowHint);
          shiftIntervals.push({ startMin: hrs.startMin, endMin: hrs.endMin });
        }
      });

      if (bestEnd < bandEnd) {
        let nextBoundary = bandEnd;
        shiftIntervals.forEach(function(iv) {
          if (iv.startMin > bestEnd && iv.startMin < nextBoundary) {
            nextBoundary = iv.startMin;
          }
        });
        bestEnd = nextBoundary;
      }

      if (bestStart > bandStart) {
        let prevBoundary = bandStart;
        shiftIntervals.forEach(function(iv) {
          if (iv.endMin < bestStart && iv.endMin > prevBoundary) {
            prevBoundary = iv.endMin;
          }
        });
        bestStart = prevBoundary;
      }
    } else {
      bestStart = Math.min(bestStart, bandStart);
      bestEnd = Math.max(bestEnd, bandEnd);
    }
  }

  return formatMinToHHMM1b3_1_(bestStart) + '-' + formatMinToHHMM1b3_1_(bestEnd);
}

function isSlotAlreadyCoveredByActiveShift1b3_(dateStr, rawClinic, absTimeStr, activeShifts) {
  const d = normalizeDateYmd1b3_(dateStr);
  const cleanC = cleanClinicBaseName_(rawClinic);
  const isInternal = String(rawClinic || '').includes('内科');

  const matchingRows = (activeShifts || []).filter(function(s) {
    if (normalizeDateYmd1b3_(s[0]) !== d) return false;
    if (cleanClinicBaseName_(s[2]) !== cleanC) return false;
    const sIsInternal = String(s[2] || '').includes('内科') || String(s[3] || '').includes('内科');
    return sIsInternal === isInternal;
  });

  if (matchingRows.length === 0) return false;

  const timeInfo = parseAbsenceTimeline_(absTimeStr);
  const targetBands = timeInfo.targetBands || ['A'];
  const row = matchingRows[0];

  for (let i = 0; i < targetBands.length; i++) {
    const band = targetBands[i];
    const countCol = band === 'A' ? 4 : (band === 'B' ? 5 : 6);
    const docsCol = band === 'A' ? 7 : (band === 'B' ? 8 : 9);

    const cntVal = parseInt(String(row[countCol] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const docsStr = String(row[docsCol] || '').trim();

    if (cntVal === 0 || !docsStr || docsStr === '未充足' || docsStr === '-' || docsStr === 'なし') {
      return false;
    }

    const tokens = splitDoctorTokensSafely_(docsStr);
    let bandFullyCovered = false;

    for (let j = 0; j < tokens.length; j++) {
      const tok = tokens[j];
      if (!tok || tok.includes('未充足')) continue;
      const tm = tok.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
      if (!tm) {
        bandFullyCovered = true;
        break;
      } else {
        const docStart = Number(tm[1]) * 60 + Number(tm[2]);
        const docEnd = Number(tm[3]) * 60 + Number(tm[4]);
        if (docStart <= timeInfo.startMin && docEnd >= timeInfo.endMin) {
          bandFullyCovered = true;
          break;
        }
      }
    }

    if (!bandFullyCovered) {
      return false;
    }
  }

  return true;
}

function recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilledRows, activeShifts, spotCandidates, tomorrowYmd, targetEndYmd) {
  const combined = [];
  const recruitSlotsIndex = new Map();

  const addRecruitIndex = function(dRaw, rawClinic, rawTime, deptHint) {
    const d = normalizeDateYmd1b3_(dRaw);
    if (!d) return;
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = (String(rawClinic || '').includes('内科') || String(deptHint || '').includes('内科')) ? '内科' : '小児科';
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!recruitSlotsIndex.has(key)) recruitSlotsIndex.set(key, []);
    const list = recruitSlotsIndex.get(key);
    normalizeAndExpandAbsenceTime1b3_(rawTime, cleanC).forEach(function(t) {
      if (t && !list.includes(t)) list.push(t);
    });
  };

  (rawUnfilledRows || []).forEach(function(r) { addRecruitIndex(r[0], r[1], r[2], ''); });
  (spotCandidates || []).forEach(function(r) { addRecruitIndex(r[0], r[2], r[3], ''); });

  // ★デバッグで判明した bqData.bosyu（1,945件）と m.dp（診療科）を確実にインデックス化！
  let bqBosyuList = [];
  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        bqBosyuList = (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []);
        bqBosyuList.forEach(function(m) {
          addRecruitIndex(m.d || m.date, m.c || m.clinic, m.t || m.time, m.dp || m.dept || '');
        });
      }
    } catch (e) {}
  }

  const registerRowIfTrulyUnfilled = function(rowArr, isFromPrimarySheet) {
    const d = normalizeDateYmd1b3_(rowArr[0]);
    if (!d || d < tomorrowYmd) return;
    const rawClinic = String(rowArr[1] || '').trim();
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = rawClinic.includes('内科') ? '内科' : '小児科';
    const expandedTimes = normalizeAndExpandAbsenceTime1b3_(rowArr[2], cleanC);

    expandedTimes.forEach(function(rawTStr) {
      const tStr = expandToLongestRecruitSlotTime1b3_(d, rawClinic, rawTStr, recruitSlotsIndex, activeShifts);
      const newTl = parseAbsenceTimeline_(tStr);

      const existingIdx = combined.findIndex(function(ex) {
        if (ex[0] !== d) return false;
        if (cleanClinicBaseName_(ex[1]) !== cleanC) return false;
        const exDept = String(ex[1] || '').includes('内科') ? '内科' : '小児科';
        if (exDept !== deptTag) return false;
        return isTimeRangeOverlapping1b3_(ex[2], tStr);
      });

      if (existingIdx !== -1) {
        const exTl = parseAbsenceTimeline_(combined[existingIdx][2]);
        if (newTl.durationHours > exTl.durationHours) {
          const mergedStart = Math.min(exTl.startMin, newTl.startMin);
          const mergedEnd = Math.max(exTl.endMin, newTl.endMin);
          combined[existingIdx][2] = formatMinToHHMM1b3_1_(mergedStart) + '-' + formatMinToHHMM1b3_1_(mergedEnd);
        }
        return;
      }

      // ★注意：『最新_不在枠』に「15:00-16:00」とあり、募集シフトに合わせて「15:00-18:00」へ拡張した場合、
      //   元の rawTStr（15:00-16:00）が未充足なのでそのまま登録する
      if (!isFromPrimarySheet && isSlotAlreadyCoveredByActiveShift1b3_(d, rawClinic, rawTStr, activeShifts)) {
        return;
      }

      const copy = rowArr.slice();
      copy[0] = d;
      copy[2] = tStr;
      combined.push(copy);
    });
  };

  // 1. 最優先：『最新_不在枠』シートにある枠を登録（募集シフトの長い時間で登録）
  (rawUnfilledRows || []).forEach(function(r) {
    registerRowIfTrulyUnfilled(r, true);
  });

  // 2. 次優先：『最新_当日シフト』から完全0人の帯のみ補完
  (activeShifts || []).forEach(function(s) {
    const d = normalizeDateYmd1b3_(s[0]);
    if (!d || d < tomorrowYmd || d > targetEndYmd) return;
    const rawC = String(s[2] || '').trim();
    const dept = String(s[3] || '').trim();
    if (!rawC || rawC.includes('バックアップ')) return;
    const cleanC = cleanClinicBaseName_(s[2]);
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

  combined.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    if (a[1] !== b[1]) return a[1] < b[1] ? -1 : 1;
    return String(a[2] || '') < String(b[2] || '') ? -1 : 1;
  });

  return combined;
}