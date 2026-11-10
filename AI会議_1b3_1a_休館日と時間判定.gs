/**
 * ============================================================================
 * AI会議_1b3_1a_休館日と時間判定.gs（分割前半）
 * 1) 『休館日』シート（全拠点・個別拠点／全日・午前・午後・午後夜間・夜間）の除外判定
 * 2) 時間帯の正規化・帯またぎ募集枠（17:00-21:00等）の照合
 * 3) 『最新_当日シフト』＋『BQ確定シフト(kakutei)』統合による2診目募集枠の完全除外判定
 * ============================================================================
 */

function normalizeDateYmd1b3_(rawDateStr) {
  const m = String(rawDateStr || '').trim().match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return String(rawDateStr || '').trim();
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 『休館日』シートを読み込み、指定枠が休館かどうかを判定・調整するチェッカーを構築
 */
function buildClosedClinicChecker1b3_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const byDate = new Map();

  const sh = ss ? ss.getSheetByName('休館日') : null;
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();
    vals.forEach(function(r) {
      const d = normalizeDateYmd1b3_(r[0]);
      const deptRaw = String(r[2] || '').trim();
      const clinicRaw = String(r[3] || '').trim();
      const timeBandRaw = String(r[4] || '').replace(/[\s\u3000]+/g, '').trim();
      if (!d || !clinicRaw) return;

      const isAllClinics = (clinicRaw === '全拠点' || clinicRaw === '全院' || clinicRaw === 'すべて');
      const cleanC = isAllClinics ? 'ALL' : cleanClinicBaseName_(clinicRaw);

      let dept = 'ALL';
      if (deptRaw.includes('内科') && !deptRaw.includes('小児')) dept = '内科';
      else if (deptRaw.includes('小児') && !deptRaw.includes('内科')) dept = '小児科';

      let closedStart = 0;
      let closedEnd = 1440;
      if (!timeBandRaw || timeBandRaw === '全日' || timeBandRaw === '終日' || timeBandRaw === '前日') {
        closedStart = 0;
        closedEnd = 1440;
      } else if (timeBandRaw === '午前') {
        closedStart = 0;
        closedEnd = 840;
      } else if (timeBandRaw === '午後夜間' || timeBandRaw === '午後・夜間') {
        closedStart = 780;
        closedEnd = 1440;
      } else if (timeBandRaw === '午後') {
        closedStart = 780;
        closedEnd = 1080;
      } else if (timeBandRaw === '夜間') {
        closedStart = 1080;
        closedEnd = 1440;
      } else {
        const tm = timeBandRaw.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
        if (tm) {
          closedStart = Number(tm[1]) * 60 + Number(tm[2]);
          closedEnd = Number(tm[3]) * 60 + Number(tm[4]);
        }
      }

      if (!byDate.has(d)) byDate.set(d, []);
      byDate.get(d).push({
        isAllClinics: isAllClinics,
        cleanClinic: cleanC,
        dept: dept,
        closedStart: closedStart,
        closedEnd: closedEnd
      });
    });
  }

  return {
    resolveOpenSlotTime: function(dateStr, rawClinic, absTimeStr, deptHint) {
      const d = normalizeDateYmd1b3_(dateStr);
      const rules = byDate.get(d);
      if (!rules || rules.length === 0) return absTimeStr;

      const cleanC = cleanClinicBaseName_(rawClinic);
      const slotDept = (String(rawClinic || '').includes('内科') || String(deptHint || '').includes('内科')) ? '内科' : '小児科';
      const tl = parseAbsenceTimeline_(absTimeStr);
      let openStart = tl.startMin;
      let openEnd = tl.endMin;

      for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        if (!rule.isAllClinics && rule.cleanClinic !== cleanC) continue;
        if (rule.dept !== 'ALL' && rule.dept !== slotDept) continue;

        if (rule.closedStart <= openStart && rule.closedEnd >= openEnd) {
          return '';
        }
        if (rule.closedStart <= openStart && rule.closedEnd > openStart && rule.closedEnd < openEnd) {
          openStart = rule.closedEnd;
        }
        if (rule.closedStart > openStart && rule.closedStart < openEnd && rule.closedEnd >= openEnd) {
          openEnd = rule.closedStart;
        }
      }

      if (openEnd - openStart < 60) return '';
      return formatMinToHHMM1b3_1_(openStart) + '-' + formatMinToHHMM1b3_1_(openEnd);
    }
  };
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
 * 2診目専用のピーク短縮枠（10:00-13:00、15:00-17:00 等）を除外するため、
 * 基本の1診目フル帯（09:00-13:00 / 15:00-18:00 / 18:00-21:00 等）かどうかを判定する関数
 */
function isStandardFullBand1ShinSlot1b3_(absTimeStr, clinicName) {
  const isKitakasai = String(clinicName || '').includes('北葛西');
  const tl = parseAbsenceTimeline_(absTimeStr);
  const s = tl.startMin;
  const e = tl.endMin;

  // 午前フル帯（09:00開始〜12:30/13:00終了）※10:00-13:00等は除外
  if (s <= 540 && e >= (isKitakasai ? 750 : 780) && e <= 840) return true;
  // 午後フル帯（14:00/15:00開始〜18:00終了）※15:00-17:00等は除外
  if ((s === 840 || s === 870 || s === 900) && e === 1080) return true;
  // 夜間フル帯（18:00開始〜20:00/21:00終了）
  if (s === 1080 && e >= (isKitakasai ? 1200 : 1260)) return true;

  return false;
}

function expandToLongestRecruitSlotTime1b3_(dateStr, rawClinic, absTimeStr, recruitSlotsIndex, activeShifts) {
  const cleanC = cleanClinicBaseName_(rawClinic);
  const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';
  const p = parseAbsenceTimeline_(absTimeStr);
  const exactOriginal = formatMinToHHMM1b3_1_(p.startMin) + '-' + formatMinToHHMM1b3_1_(p.endMin);

  const key = dateStr + '_' + cleanC + '_' + deptTag;
  const candidates = (recruitSlotsIndex && recruitSlotsIndex.get(key)) || [];
  if (candidates.length === 0) return exactOriginal;

  let matchedBosyuTime = '';
  let matchedDuration = p.durationHours;

  candidates.forEach(function(candTime) {
    const cp = parseAbsenceTimeline_(candTime);
    if (cp.durationHours <= 5 &&
        cp.startMin <= p.startMin &&
        cp.endMin >= p.endMin &&
        cp.durationHours >= matchedDuration) {
      matchedDuration = cp.durationHours;
      matchedBosyuTime = formatMinToHHMM1b3_1_(cp.startMin) + '-' + formatMinToHHMM1b3_1_(cp.endMin);
    }
  });

  return matchedBosyuTime || exactOriginal;
}

/**
 * ★新設：『最新_当日シフト』と『BQ確定シフト(kakutei)』の両方を統合し、
 * 1診目がすでに在席している枠（＝2診目募集）を全日程で確実に除外するカバレッジ判定チェッカー
 */
function buildConfirmedShiftCoverageChecker1b3_(activeShifts, bqKakuteiList) {
  const intervalsByKey = new Map();      // key: d_cleanC_dept -> [{start, end}]
  const datesWithBandRoster = new Set(); // 時間帯付きの確定シフトが存在する日付
  const datesWithAnyBqDoc = new Set();   // BQ確定シフトに医師名が存在する日付
  const clinicsWithAnyDocOnDate = new Set(); // key: d_cleanC_dept（その日に1人でも確定医がいる拠点）

  const addInterval = function(d, cleanC, deptTag, startMin, endMin) {
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!intervalsByKey.has(key)) intervalsByKey.set(key, []);
    intervalsByKey.get(key).push({ start: startMin, end: endMin });
    clinicsWithAnyDocOnDate.add(key);
    datesWithBandRoster.add(d);
  };

  // 1. 『最新_当日シフト』シートを登録
  (activeShifts || []).forEach(function(row) {
    const d = normalizeDateYmd1b3_(row[0]);
    const cleanC = cleanClinicBaseName_(row[2]);
    if (!d || !cleanC || cleanC.includes('バックアップ')) return;
    const isKitakasai = cleanC.includes('北葛西');
    const deptTag = (String(row[2] || '').includes('内科') || String(row[3] || '').includes('内科')) ? '内科' : '小児科';
    datesWithBandRoster.add(d);

    const bands = [
      { cnt: row[4], docs: row[7], start: 540, end: isKitakasai ? 750 : 780 },
      { cnt: row[5], docs: row[8], start: isKitakasai ? 840 : 900, end: 1080 },
      { cnt: row[6], docs: row[9], start: 1080, end: isKitakasai ? 1200 : 1260 }
    ];

    bands.forEach(function(b) {
      const cntVal = parseInt(String(b.cnt || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const docsStr = String(b.docs || '').trim();
      if (cntVal === 0 || !docsStr || docsStr === '未充足' || docsStr === '-' || docsStr === 'なし') return;

      splitDoctorTokensSafely_(docsStr).forEach(function(tok) {
        if (!tok || tok.includes('未充足')) return;
        const tm = tok.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
        if (!tm) {
          addInterval(d, cleanC, deptTag, b.start, b.end);
        } else {
          addInterval(d, cleanC, deptTag, Number(tm[1]) * 60 + Number(tm[2]), Number(tm[3]) * 60 + Number(tm[4]));
        }
      });
    });
  });

  // 2. 『BQ確定シフト（kakutei）』を登録（『最新_当日シフト』期間外の日程でも2診目募集をブロック）
  (bqKakuteiList || []).forEach(function(s) {
    const st = String(s.st || s.status || '').trim();
    if (st && (st.includes('キャンセル') || st.includes('取消') || st.includes('却下'))) return;

    const d = normalizeDateYmd1b3_(s.d || s.date);
    const rawC = String(s.c || s.clinic || '').trim();
    const cleanC = cleanClinicBaseName_(rawC);
    const docName = String(s.n || s.name || '').trim();
    if (!d || !cleanC || cleanC.includes('バックアップ') || !docName || docName === '未充足') return;

    const deptTag = (rawC.includes('内科') || String(s.dp || s.dept || '').includes('内科')) ? '内科' : '小児科';
    const key = d + '_' + cleanC + '_' + deptTag;
    clinicsWithAnyDocOnDate.add(key);
    datesWithAnyBqDoc.add(d);

    const isKitakasai = cleanC.includes('北葛西');
    const rawT = String(s.t || s.time || s.tm || s.workTime || '').normalize('NFKC').trim();
    if (!rawT) return;

    const tm = rawT.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
    if (tm) {
      addInterval(d, cleanC, deptTag, Number(tm[1]) * 60 + Number(tm[2]), Number(tm[3]) * 60 + Number(tm[4]));
    } else if (rawT.includes('終日') || rawT.includes('全日')) {
      addInterval(d, cleanC, deptTag, 540, isKitakasai ? 1200 : 1260);
    } else if (rawT.includes('午前')) {
      addInterval(d, cleanC, deptTag, 540, isKitakasai ? 750 : 780);
    } else if (rawT.includes('午後')) {
      addInterval(d, cleanC, deptTag, isKitakasai ? 840 : 900, 1080);
    } else if (rawT.includes('夜間')) {
      addInterval(d, cleanC, deptTag, 1080, isKitakasai ? 1200 : 1260);
    }
  });

  const isCoveredBy15MinSlices = function(d, cleanC, deptTag, startMin, endMin) {
    const list = intervalsByKey.get(d + '_' + cleanC + '_' + deptTag) || [];
    if (list.length === 0) return false;
    for (let tMin = startMin; tMin < endMin; tMin += 15) {
      const ok = list.some(function(iv) { return iv.start <= tMin && iv.end >= (tMin + 15); });
      if (!ok) return false;
    }
    return true;
  };

  const hasAnyOverlappingConfirmedDoc = function(d, cleanC, deptTag, startMin, endMin) {
    const list = intervalsByKey.get(d + '_' + cleanC + '_' + deptTag) || [];
    return list.some(function(iv) {
      return Math.max(iv.start, startMin) < Math.min(iv.end, endMin);
    });
  };

  return {
    isSlotAlreadyCovered: function(dateStr, rawClinic, absTimeStr) {
      const d = normalizeDateYmd1b3_(dateStr);
      const cleanC = cleanClinicBaseName_(rawClinic);
      const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';
      const tl = parseAbsenceTimeline_(absTimeStr);
      return isCoveredBy15MinSlices(d, cleanC, deptTag, tl.startMin, tl.endMin);
    },

    /**
     * 『最新_不在枠』に載っていない bosyu 枠を「終日0人等の真の医師不在枠」として補完してよいか厳密判定
     * （2診目募集枠［10:00-13:00 / 15:00-17:00 や1診目在席枠］はすべて false で除外）
     */
    canSupplementFromBosyu: function(dateStr, rawClinic, absTimeStr) {
      const d = normalizeDateYmd1b3_(dateStr);
      const cleanC = cleanClinicBaseName_(rawClinic);
      const deptTag = String(rawClinic || '').includes('内科') ? '内科' : '小児科';

      // 1) 10:00-13:00 や 15:00-17:00 などの2診目ピーク短縮枠は無条件で除外
      if (!isStandardFullBand1ShinSlot1b3_(absTimeStr, cleanC)) return false;

      const tl = parseAbsenceTimeline_(absTimeStr);
      const key = d + '_' + cleanC + '_' + deptTag;

      // 2) 時間帯付きの確定シフトデータが存在する日：対象帯に1人でも確定医がいれば2診目募集なので除外
      if (datesWithBandRoster.has(d)) {
        return !hasAnyOverlappingConfirmedDoc(d, cleanC, deptTag, tl.startMin, tl.endMin);
      }

      // 3) 時間なしの確定医リストのみ存在する日：その拠点に終日1人も確定医がいない場合（終日0人拠点）のみ補完許可
      if (datesWithAnyBqDoc.has(d)) {
        return !clinicsWithAnyDocOnDate.has(key);
      }

      // 4) 確定シフトデータが一切存在しない日は2診目判別不能のため補完しない（『最新_不在枠』のみ使用）
      return false;
    }
  };
}

function isSlotAlreadyCoveredByActiveShift1b3_(dateStr, rawClinic, absTimeStr, activeShifts) {
  const checker = buildConfirmedShiftCoverageChecker1b3_(activeShifts, []);
  return checker.isSlotAlreadyCovered(dateStr, rawClinic, absTimeStr);
}