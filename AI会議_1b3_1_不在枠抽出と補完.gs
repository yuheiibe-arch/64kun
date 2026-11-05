/**
 * ============================================================================
 * AI会議_1b3_1_不在枠抽出と補完.gs
 * （募集シフトの募集時間への自動合わせ込み・複数重複時は「長い方の枠」を採用する版）
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
 * 2つの時間帯文字列（例: "15:00-16:00" と "15:00-18:00"）が重なっているかを判定
 */
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
 * ★新設：不在時間（例：15:00-16:00）に対して、募集シフト一覧や当日シフトの状況から
 * 実際に募集が出ている長い方の募集時間（例：15:00-18:00 や 15:00-17:00）へ拡張する関数
 */
function expandToLongestRecruitSlotTime1b3_(dateStr, rawClinic, absTimeStr, recruitSlotsIndex, activeShifts) {
  const cleanC = cleanClinicBaseName_(rawClinic);
  const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';
  const isKitakasai = cleanC.includes('北葛西');
  const p = parseAbsenceTimeline_(absTimeStr);
  let bestStart = p.startMin;
  let bestEnd = p.endMin;

  // 1. 募集シフトインデックス（BQ未充足・募集シート・不在枠の別行）に重なる枠があれば、最も長い範囲（和集合）を採用
  const key = dateStr + '_' + cleanC + '_' + deptTag;
  const candidates = (recruitSlotsIndex && recruitSlotsIndex.get(key)) || [];
  candidates.forEach(function(candTime) {
    const cp = parseAbsenceTimeline_(candTime);
    // 同じ帯（午前・午後・夜間）で重なっている募集枠なら、長い方に広げる
    if (Math.max(bestStart, cp.startMin) < Math.min(bestEnd, cp.endMin) &&
        cp.targetBands[0] === p.targetBands[0]) {
      bestStart = Math.min(bestStart, cp.startMin);
      bestEnd = Math.max(bestEnd, cp.endMin);
    }
  });

  // 2. まだ帯の途中（例：15:00-16:00 や 12:15-13:00）で1時間以下などの細切れになっている場合、
  //    当日シフト（activeShifts）を確認し、募集枠として帯の終わり（18:00や17:00等）まで空いていればその募集枠時間に合わせる
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

      // その帯にいる医師の開始・終了時刻を取得
      const shiftIntervals = [];
      tokens.forEach(function(tok) {
        if (!tok || tok.includes('未充足')) return;
        if (typeof parseDocTokenWorkingMinutes1b3_ === 'function') {
          const hrs = parseDocTokenWorkingMinutes1b3_(tok, primaryBand, isKitakasai, allRowHint);
          shiftIntervals.push({ startMin: hrs.startMin, endMin: hrs.endMin });
        }
      });

      // 後方の拡張（例：15:00-16:00 で16:00から1人だけ医師が入る場合、次の医師が入る時刻または18:00まで募集枠として広げる）
      if (bestEnd < bandEnd) {
        let nextBoundary = bandEnd;
        shiftIntervals.forEach(function(iv) {
          if (iv.startMin > bestEnd && iv.startMin < nextBoundary) {
            nextBoundary = iv.startMin;
          }
        });
        bestEnd = nextBoundary;
      }

      // 前方の拡張（例：17:00-18:00 が不在で、15:00から募集枠が空いている場合）
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
      // 当日シフト行自体がない（完全空き）なら標準の帯全体（09:00-13:00 / 15:00-18:00 / 18:00-21:00）を募集枠とする
      bestStart = Math.min(bestStart, bandStart);
      bestEnd = Math.max(bestEnd, bandEnd);
    }
  }

  return formatMinToHHMM1b3_1_(bestStart) + '-' + formatMinToHHMM1b3_1_(bestEnd);
}

/**
 * 『最新_当日シフト』（確定シフト）を照合し、すでに確定医師がいて不在が解消しているかを判定
 */
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

/**
 * ★募集シフトの募集時間（重なる枠が複数ある場合は長い方）に統一して不在枠リストを構築する関数
 */
function recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilledRows, activeShifts, spotCandidates, tomorrowYmd, targetEndYmd) {
  const combined = [];
  const recruitSlotsIndex = new Map(); // key: "date_clinic_dept" -> [timeStr, ...]

  const addRecruitIndex = function(dRaw, rawClinic, rawTime) {
    const d = normalizeDateYmd1b3_(dRaw);
    if (!d) return;
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!recruitSlotsIndex.has(key)) recruitSlotsIndex.set(key, []);
    const list = recruitSlotsIndex.get(key);
    normalizeAndExpandAbsenceTime1b3_(rawTime, cleanC).forEach(function(t) {
      if (t && !list.includes(t)) list.push(t);
    });
  };

  // 募集シフト候補の時間帯をすべてインデックス化（不在枠シート・個別打診候補シート・BQシフト状況パック）
  (rawUnfilledRows || []).forEach(function(r) { addRecruitIndex(r[0], r[1], r[2]); });
  (spotCandidates || []).forEach(function(r) { addRecruitIndex(r[0], r[2], r[3]); });

  let bqMijyusokuList = [];
  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        bqMijyusokuList = (bqData.mijyusoku || []).concat(bqData.boshu || []);
        bqMijyusokuList.forEach(function(m) {
          addRecruitIndex(m.d || m.date, m.c || m.clinic, m.t || m.time);
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
      // ★募集シフトの長い方の募集時間（例：15:00-16:00 ➔ 15:00-18:00）へ自動拡張
      const tStr = expandToLongestRecruitSlotTime1b3_(d, rawClinic, rawTStr, recruitSlotsIndex, activeShifts);
      const newTl = parseAbsenceTimeline_(tStr);

      // すでに同日・同拠点・同診療科で時間帯が重なる枠があるかチェック
      const existingIdx = combined.findIndex(function(ex) {
        if (ex[0] !== d) return false;
        if (cleanClinicBaseName_(ex[1]) !== cleanC) return false;
        const exDept = String(ex[1] || '').includes('内科') ? '内科' : '小児科';
        if (exDept !== deptTag) return false;
        return isTimeRangeOverlapping1b3_(ex[2], tStr);
      });

      if (existingIdx !== -1) {
        // ★もし重なる枠が既にある場合、比較して「より長い方の募集時間」に更新する！
        const exTl = parseAbsenceTimeline_(combined[existingIdx][2]);
        if (newTl.durationHours > exTl.durationHours) {
          const mergedStart = Math.min(exTl.startMin, newTl.startMin);
          const mergedEnd = Math.max(exTl.endMin, newTl.endMin);
          combined[existingIdx][2] = formatMinToHHMM1b3_1_(mergedStart) + '-' + formatMinToHHMM1b3_1_(mergedEnd);
        }
        return;
      }

      if (!isFromPrimarySheet && isSlotAlreadyCoveredByActiveShift1b3_(d, rawClinic, tStr, activeShifts)) {
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

  // 2. 次優先：『BQシフト状況パック(mijyusoku)』から未登録の枠またはより長い募集枠を反映
  bqMijyusokuList.forEach(function(m) {
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

  // 3. 最後：『最新_当日シフト』から完全0人の帯のみ補完
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