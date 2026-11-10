/**
 * ============================================================================
 * AI会議_1b3_1b_不在枠抽出とBQ補完.gs（分割後半）
 * 1) 『最新_不在枠』・『最新_当日シフト』0人帯・『BQ終日0人不在拠点』の統合（★2診目募集枠は完全除外）
 * 2) 補完枠の時給表記を「特別時給／通常時給」に統一（数値むき出し表記の解消）
 * 3) 地域順ソート（東京 ➔ 神奈川 ➔ 千葉・茨城 ➔ 埼玉 ➔ 大阪）
 * ============================================================================
 */

function getRegionOrderRank1b3_(clinicName) {
  const area = inferAreaForDoc_(cleanClinicBaseName_(clinicName));
  if (area === '東京') return 1;
  if (area === '神奈川') return 2;
  if (area === '千葉') return 3;
  if (area === '茨城') return 4;
  if (area === '埼玉') return 5;
  if (area === '関西') return 6;
  return 99;
}

/**
 * BQパックから ①募集時間帯インデックス、②終日0人拠点の補完候補、③確定シフト一覧（kakutei）を取得
 */
function loadBqBosyuIndexAndSupplementRows1b3_(rawUnfilledRows, tomorrowYmd, effectiveEndYmd) {
  const recruitSlotsIndex = new Map();
  const supplementRows = [];
  let kakuteiList = [];

  if (typeof fetchAndDecryptBqMailForResolver_ !== 'function') {
    return { recruitSlotsIndex: recruitSlotsIndex, supplementRows: supplementRows, kakuteiList: kakuteiList };
  }

  let bqBosyuList = [];
  try {
    const bqData = fetchAndDecryptBqMailForResolver_();
    if (bqData) {
      bqBosyuList = (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []);
      kakuteiList = (bqData.kakutei || []).concat(bqData.obo || []);
    }
  } catch (e) {
    return { recruitSlotsIndex: recruitSlotsIndex, supplementRows: supplementRows, kakuteiList: kakuteiList };
  }

  bqBosyuList.forEach(function(m) {
    const st = String(m.st || '').trim();
    if (st && (st.includes('非掲載') || st.includes('終了') || st.includes('取消') || st.includes('キャンセル'))) return;

    const d = normalizeDateYmd1b3_(m.d || m.date);
    if (!d) return;

    const rawC = String(m.c || m.clinic || '').trim();
    const rawDp = String(m.dp || m.dept || '').trim();
    const cleanC = cleanClinicBaseName_(rawC);
    if (!cleanC || cleanC.includes('バックアップ')) return;

    const isInternal = rawC.includes('内科') || rawDp.includes('内科');
    const deptTag = isInternal ? '内科' : '小児科';
    const rawT = String(m.t || m.time || '').trim();
    if (!rawT) return;

    // 1. 募集時間帯インデックスへ登録
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!recruitSlotsIndex.has(key)) recruitSlotsIndex.set(key, []);
    const list = recruitSlotsIndex.get(key);
    normalizeAndExpandAbsenceTime1b3_(rawT, cleanC).forEach(function(t) {
      if (t && !list.includes(t)) list.push(t);
    });

    // 2. 対象期間内であれば終日0人拠点用の補完候補行を生成（※後段で2診目募集は完全除外）
    if (d < tomorrowYmd || (effectiveEndYmd && d > effectiveEndYmd)) return;

    const hasMultiDept = cleanC.includes('亀有') || cleanC.includes('北葛西');
    const clinicLabel = isInternal ? (cleanC + '（内科）') : (hasMultiDept ? (cleanC + '（小児科）') : cleanC);

    let hourlyStr = '';
    if (m.uMax && Number(m.uMax) >= 8000) {
      hourlyStr = String(m.uMax);
    } else if (m.w) {
      const wNum = Number(String(m.w).replace(/[^0-9]/g, '')) || 0;
      const tl = parseAbsenceTimeline_(rawT);
      if (wNum >= 25000 && tl.durationHours > 0) {
        hourlyStr = String(Math.round(wNum / tl.durationHours));
      } else if (wNum > 0) {
        hourlyStr = String(wNum);
      }
    }
    if (!hourlyStr) hourlyStr = '14000';

    // ★修正：特別時給でない場合は必ず「通常時給」を入れ、(14000) のような数値むき出し表記を防ぐ
    const spRaw = String(m.sp || m.specialWage || '').trim();
    const spLabel = (spRaw && spRaw !== '-' && spRaw !== 'なし' && spRaw !== 'FALSE' && !spRaw.includes('通常'))
      ? (spRaw.includes('特別時給') ? '特別時給' : spRaw)
      : '通常時給';

    const relatedMemos = [];
    (rawUnfilledRows || []).forEach(function(ur) {
      if (normalizeDateYmd1b3_(ur[0]) !== d) return;
      const memoTxt = String(ur[12] || '').trim();
      if (!memoTxt || memoTxt === '-' || !memoTxt.includes(cleanC)) return;
      if (!isTimeRangeOverlapping1b3_(String(ur[2] || ''), rawT)) return;
      memoTxt.split(/\r?\n/).forEach(function(line) {
        const trimmed = line.trim();
        if (trimmed && trimmed.includes(cleanC) && !relatedMemos.includes(trimmed)) {
          relatedMemos.push(trimmed);
        }
      });
    });

    supplementRows.push([
      d,
      clinicLabel,
      rawT,
      hourlyStr,
      'なし',
      'なし',
      'Phaze.4',
      '',
      '',
      '',
      '未充足',
      spLabel,
      relatedMemos.join('\n'),
      ''
    ]);
  });

  return {
    recruitSlotsIndex: recruitSlotsIndex,
    supplementRows: supplementRows,
    kakuteiList: kakuteiList
  };
}

function recoverMissingFullDayOrUnfilledRows1b3_(rawUnfilledRows, activeShifts, spotCandidates, tomorrowYmd, targetEndYmd) {
  const combined = [];
  const closedChecker = buildClosedClinicChecker1b3_();

  let effectiveEndYmd = targetEndYmd || '';
  if (!effectiveEndYmd) {
    (rawUnfilledRows || []).forEach(function(r) {
      const d = normalizeDateYmd1b3_(r[0]);
      if (d && d > effectiveEndYmd) effectiveEndYmd = d;
    });
    (activeShifts || []).forEach(function(s) {
      const d = normalizeDateYmd1b3_(s[0]);
      if (d && d > effectiveEndYmd) effectiveEndYmd = d;
    });
  }

  const bqLoaded = loadBqBosyuIndexAndSupplementRows1b3_(rawUnfilledRows, tomorrowYmd, effectiveEndYmd);
  const recruitSlotsIndex = bqLoaded.recruitSlotsIndex;
  const coverageChecker = buildConfirmedShiftCoverageChecker1b3_(activeShifts, bqLoaded.kakuteiList);

  const registerRowIfTrulyUnfilled = function(rowArr, sourceType) {
    const isFromPrimarySheet = (sourceType === 'PRIMARY');
    const isFromBosyuSupplement = (sourceType === 'BOSYU_SUP');

    const d = normalizeDateYmd1b3_(rowArr[0]);
    if (!d || d < tomorrowYmd) return;
    if (!isFromPrimarySheet && effectiveEndYmd && d > effectiveEndYmd) return;

    const rawClinic = String(rowArr[1] || '').trim();
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = rawClinic.includes('内科') ? '内科' : '小児科';
    const expandedTimes = normalizeAndExpandAbsenceTime1b3_(rowArr[2], cleanC);

    expandedTimes.forEach(function(rawTStr) {
      const candidateTime = isFromPrimarySheet
        ? expandToLongestRecruitSlotTime1b3_(d, rawClinic, rawTStr, recruitSlotsIndex, activeShifts)
        : rawTStr;

      // 『休館日』シート判定
      const tStr = closedChecker.resolveOpenSlotTime(d, rawClinic, candidateTime, deptTag);
      if (!tStr) return;

      // ★2診目募集枠の完全ブロック判定
      if (isFromBosyuSupplement) {
        if (!coverageChecker.canSupplementFromBosyu(d, rawClinic, tStr)) return;
      } else if (!isFromPrimarySheet) {
        if (coverageChecker.isSlotAlreadyCovered(d, rawClinic, tStr)) return;
      }

      const newTl = parseAbsenceTimeline_(tStr);

      const existingIdx = combined.findIndex(function(ex) {
        if (ex[0] !== d) return false;
        if (cleanClinicBaseName_(ex[1]) !== cleanC) return false;
        const exDept = String(ex[1] || '').includes('内科') ? '内科' : '小児科';
        if (exDept !== deptTag) return false;
        return isTimeRangeOverlapping1b3_(ex[2], tStr);
      });

      if (existingIdx !== -1) {
        const exRow = combined[existingIdx];
        const exTl = parseAbsenceTimeline_(exRow[2]);
        if (newTl.startMin <= exTl.startMin && newTl.endMin >= exTl.endMin && newTl.durationHours > exTl.durationHours) {
          exRow[2] = tStr;
        }
        if (rowArr[4] && rowArr[4] !== 'なし' && !String(exRow[4] || '').includes(rowArr[4])) {
          exRow[4] = (!exRow[4] || exRow[4] === 'なし') ? rowArr[4] : (exRow[4] + ',' + rowArr[4]);
        }
        if (rowArr[5] && rowArr[5] !== 'なし' && !String(exRow[5] || '').includes(rowArr[5])) {
          exRow[5] = (!exRow[5] || exRow[5] === 'なし') ? rowArr[5] : (exRow[5] + ',' + rowArr[5]);
        }
        const newMemo = String(rowArr[12] || '').trim();
        if (newMemo && newMemo !== '-') {
          const curMemo = String(exRow[12] || '').trim();
          if (!curMemo || curMemo === '-') {
            exRow[12] = newMemo;
          } else if (!curMemo.includes(newMemo)) {
            exRow[12] = curMemo + '\n' + newMemo;
          }
        }
        if (!exRow[9] && rowArr[9]) exRow[9] = rowArr[9];
        if ((!exRow[11] || exRow[11] === '-') && rowArr[11]) exRow[11] = rowArr[11];
        return;
      }

      const copy = rowArr.slice();
      copy[0] = d;
      copy[2] = tStr;
      combined.push(copy);
    });
  };

  // 1. 最優先：『最新_不在枠』シートにある本当の医師不在枠を登録
  (rawUnfilledRows || []).forEach(function(r) {
    registerRowIfTrulyUnfilled(r, 'PRIMARY');
  });

  // 2. 次優先：『最新_当日シフト』に拠点行があるものの特定帯が0人の枠を補完
  (activeShifts || []).forEach(function(s) {
    const d = normalizeDateYmd1b3_(s[0]);
    if (!d || d < tomorrowYmd || (effectiveEndYmd && d > effectiveEndYmd)) return;
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
      registerRowIfTrulyUnfilled([d, clinicLabel, '09:00-13:00', '14000', 'なし', docsB, '', '', '', '', '未充足', '通常時給', '', ''], 'SHIFT_ZERO');
    }
    if (cntB === 0) {
      registerRowIfTrulyUnfilled([d, clinicLabel, '15:00-18:00', '13000', docsA, docsC, '', '', '', '', '未充足', '通常時給', '', ''], 'SHIFT_ZERO');
    }
    if (cntC === 0 && !dept.includes('内科')) {
      const nightTime = cleanC.includes('北葛西') ? '18:00-20:00' : '18:00-21:00';
      registerRowIfTrulyUnfilled([d, clinicLabel, nightTime, '15000', docsB, 'なし', '', '', '', '', '未充足', '通常時給', '', ''], 'SHIFT_ZERO');
    }
  });

  // 3. 終日0人不在でシートに行自体がない拠点（例：11/12 西新井）のみを厳密補完（★2診目募集は除外）
  bqLoaded.supplementRows.forEach(function(supRow) {
    registerRowIfTrulyUnfilled(supRow, 'BOSYU_SUP');
  });

  // 4. 地域順（東京 ➔ 神奈川 ➔ 千葉・茨城 ➔ 埼玉 ➔ 大阪）でソート
  combined.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    const rA = getRegionOrderRank1b3_(a[1]);
    const rB = getRegionOrderRank1b3_(b[1]);
    if (rA !== rB) return rA - rB;
    if (a[1] !== b[1]) return a[1] < b[1] ? -1 : 1;
    return String(a[2] || '') < String(b[2] || '') ? -1 : 1;
  });

  return combined;
}