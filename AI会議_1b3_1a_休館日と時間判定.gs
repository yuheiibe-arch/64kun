/**
 * ============================================================================
 * AI会議_1b3_1a_休館日と時間判定.gs（分割前半）
 * 1) 『休館日』シート（全拠点・個別拠点／全日・午前・午後・午後夜間・夜間）の除外判定
 * 2) 時間帯の正規化・帯またぎ募集枠（17:00-21:00等）の照合
 * 3) 『最新_当日シフト』との15分刻み在席カバー判定
 * ============================================================================
 */

function normalizeDateYmd1b3_(rawDateStr) {
  const m = String(rawDateStr || '').trim().match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return String(rawDateStr || '').trim();
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 『休館日』シートを読み込み、指定枠が休館かどうかを判定・調整するチェッカーを構築
 * シート構造：A列=日付, B列=曜日, C列=診療科, D列=拠点名(「全拠点」あり), E列=時間(午前/午後/全日/午後夜間/夜間), F列=備考
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
        closedEnd = 840; // 〜14:00
      } else if (timeBandRaw === '午後夜間' || timeBandRaw === '午後・夜間') {
        closedStart = 780; // 13:00〜24:00
        closedEnd = 1440;
      } else if (timeBandRaw === '午後') {
        closedStart = 780; // 13:00〜18:00
        closedEnd = 1080;
      } else if (timeBandRaw === '夜間') {
        closedStart = 1080; // 18:00〜24:00
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
 * BQパック（bosyu＝実際の掲載募集枠）に元の不在時間を完全に内包する募集枠（5時間以内、例：17:00-21:00）が
 * 実在する場合のみその募集枠を採用し、該当がなければ『最新_不在枠』の実時間をそのまま返す関数
 */
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
 * 『最新_当日シフト』上で対象時間帯に既に1名以上の医師が勤務しているかを15分刻みで厳密判定する関数
 */
function isSlotAlreadyCoveredByActiveShift1b3_(dateStr, rawClinic, absTimeStr, activeShifts) {
  const d = normalizeDateYmd1b3_(dateStr);
  const cleanC = cleanClinicBaseName_(rawClinic);
  const isKitakasai = cleanC.includes('北葛西');
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

    const bandFullStart = band === 'A' ? 540 : (band === 'B' ? 900 : 1080);
    const bandFullEnd = band === 'A' ? 780 : (band === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));
    const reqStart = Math.max(timeInfo.startMin, bandFullStart);
    const reqEnd = Math.min(timeInfo.endMin, bandFullEnd);

    const tokens = splitDoctorTokensSafely_(docsStr);
    const intervals = [];
    tokens.forEach(function(tok) {
      if (!tok || tok.includes('未充足')) return;
      const tm = tok.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
      if (!tm) {
        intervals.push({ start: bandFullStart, end: bandFullEnd });
      } else {
        intervals.push({
          start: Number(tm[1]) * 60 + Number(tm[2]),
          end: Number(tm[3]) * 60 + Number(tm[4])
        });
      }
    });

    if (intervals.length === 0) return false;

    for (let tMin = reqStart; tMin < reqEnd; tMin += 15) {
      const sliceCovered = intervals.some(function(iv) {
        return iv.start <= tMin && iv.end >= (tMin + 15);
      });
      if (!sliceCovered) return false;
    }
  }

  return true;
}