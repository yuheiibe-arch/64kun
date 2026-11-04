/**
 * ============================================================================
 * AI会議_1b3_1_不在枠抽出と補完.gs
 * （日付・時間帯の正規化、重複時間帯の排除、確定シフト充足済み除外、終日不在補完）
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
 * 2つの時間帯文字列（例: "19:00-21:00" と "18:00-21:00"）が重なっているかを判定
 */
function isTimeRangeOverlapping1b3_(timeStr1, timeStr2) {
  if (timeStr1 === timeStr2) return true;
  const p1 = parseAbsenceTimeline_(timeStr1);
  const p2 = parseAbsenceTimeline_(timeStr2);
  return Math.max(p1.startMin, p2.startMin) < Math.min(p1.endMin, p2.endMin);
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
 * 『最新_不在枠』をベースにしつつ、重複時間帯の二重生成を防ぎながら終日不在枠等を補完する関数
 */
function recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilledRows, activeShifts, spotCandidates, tomorrowYmd, targetEndYmd) {
  const combined = [];

  const registerRowIfTrulyUnfilled = function(rowArr, isFromPrimarySheet) {
    const d = normalizeDateYmd1b3_(rowArr[0]);
    if (!d || d < tomorrowYmd) return;
    const rawClinic = String(rowArr[1] || '').trim();
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = rawClinic.includes('内科') ? '内科' : '小児科';
    const expandedTimes = normalizeAndExpandAbsenceTime1b3_(rowArr[2], cleanC);

    expandedTimes.forEach(function(tStr) {
      // すでに同日・同拠点・同診療科で時間帯が重なる枠（例：19:00-21:00に対する18:00-21:00）があれば追加しない
      const alreadyHasOverlappingSlot = combined.some(function(ex) {
        if (ex[0] !== d) return false;
        if (cleanClinicBaseName_(ex[1]) !== cleanC) return false;
        const exDept = String(ex[1] || '').includes('内科') ? '内科' : '小児科';
        if (exDept !== deptTag) return false;
        return isTimeRangeOverlapping1b3_(ex[2], tStr);
      });
      if (alreadyHasOverlappingSlot) return;

      if (!isFromPrimarySheet && isSlotAlreadyCoveredByActiveShift1b3_(d, rawClinic, tStr, activeShifts)) {
        return;
      }

      const copy = rowArr.slice();
      copy[0] = d;
      copy[2] = tStr;
      combined.push(copy);
    });
  };

  // 1. 最優先：『最新_不在枠』シートにある不在枠を登録
  (rawUnfilledRows || []).forEach(function(r) {
    registerRowIfTrulyUnfilled(r, true);
  });

  // 2. 次優先：『BQシフト状況パック(mijyusoku)』から未登録の時間帯のみ補完
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

  // 3. 最後：『最新_当日シフト』から完全0人の帯（終日不在で漏れた拠点など）のみ補完
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