/**
 * ============================================================================
 * AI会議_1b3_コンテキスト構築.gs
 * （①前後医延長・早出 ➔ ②ドミナント2診崩し ➔ ③確定シフト＆60分休憩ルール個別打診の構築）
 * ============================================================================
 */

function evaluateValidAdjacentDocsForSlot_(docsStr, actionType, dateStr, dow, targetClinic, targetArea, timeInfo, masters) {
  const tokens = splitDoctorTokensSafely_(docsStr);
  const validTokens = [];
  const validDocKeys = new Set();

  tokens.forEach(function(tok) {
    const k = normKeyForDoc_(tok);
    if (!k || masters.absoluteNgDocSet.has(k)) return;
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

    if (tok.includes('常勤')) {
      const condStr = masters.ftAllowCondMap.get(k) || '';
      if (masters.ftAllowDocSet.has(k) && tok.includes('所定休出') && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
        validTokens.push(k + '(常勤・所定休出)');
        validDocKeys.add(k);
      }
      return;
    }

    const roleBadge = masters.sbuSet.has(k) ? '(★SBU)' : '(非常勤)';
    validTokens.push(k + roleBadge);
    validDocKeys.add(k);
  });

  return {
    docs: validTokens.length > 0 ? validTokens.join(' / ') : 'なし',
    status: validTokens.length > 0 ? (actionType + '打診可能') : '',
    docKeySet: validDocKeys
  };
}

function collectCompactContextForAi_(ss) {
  const today = new Date();
  const todayYmd = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  const twoWeeksYmd = Utilities.formatDate(new Date(today.getTime() + 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const oneMonthYmd = Utilities.formatDate(new Date(today.getTime() + 31 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

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

  const unfilledAll = getRows('最新_不在枠').filter(function(r) { return r[0] && r[0] >= todayYmd; });
  const within2W = unfilledAll.filter(function(r) { return r[0] <= twoWeeksYmd; });
  const targetEndYmd = (within2W.length < 10 && unfilledAll.length > within2W.length) ? oneMonthYmd : twoWeeksYmd;
  const targetUnfilled = unfilledAll.filter(function(r) { return r[0] <= targetEndYmd; });

  const periodStart = targetUnfilled.length > 0 ? targetUnfilled[0][0] : todayYmd;
  const periodEnd = targetUnfilled.length > 0 ? targetUnfilled[targetUnfilled.length - 1][0] : targetEndYmd;

  const activeShifts = getRows('最新_当日シフト');
  const streakWarnings = getRows('最新_連勤注意リスト');
  const spotCandidates = getRows('最新_個別打診候補');
  const bundledCandidates = getRows('最新_おまとめ打診候補');
  const kbRows = getRecentRows('ナレッジシート', 50);

  // Googleドキュメントからドミナント移動マトリクスを読み込み
  const dominantMatrix = loadDominantMatrixFromDoc_();
  const masters = buildMasterAndConfirmedShiftIndex_(ss, activeShifts, streakWarnings);

  const spotMap = new Map();
  spotCandidates.forEach(function(r) {
    const key = r[0] + '_' + r[2] + '_' + r[3];
    if (!spotMap.has(key)) spotMap.set(key, []);
    const dKey = normKeyForDoc_(r[5]);
    if (!dKey || masters.absoluteNgDocSet.has(dKey)) return;
    spotMap.get(key).push({
      rank: r[4],
      doctor: dKey,
      sameDowSlotCount: String(r[8] || '0').replace(/回$/g, '') + '回',
      totalCount: r[10],
      rawNote: String(r[13] || '').trim()
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

    // ① 前後医（延長・早出）
    const prevChecked = evaluateValidAdjacentDocsForSlot_(u[4] || 'なし', '延長', dateStr, dow, cleanTargetClinic, area, timeInfo, masters);
    const nextChecked = evaluateValidAdjacentDocsForSlot_(u[5] || 'なし', '早出', dateStr, dow, cleanTargetClinic, area, timeInfo, masters);

    // ② ドミナント定義書に基づく2診崩し・併行募集の抽出
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

        timeInfo.targetBands.forEach(function(bandCode) {
          const countCol = bandCode === 'A' ? 4 : (bandCode === 'B' ? 5 : 6);
          const docsCol = bandCode === 'A' ? 7 : (bandCode === 'B' ? 8 : 9);
          const docCount = Number(s[countCol] || 0);
          const rawDocsStr = String(s[docsCol] || '');
          if (!rawDocsStr || rawDocsStr === '未充足' || docCount < 1) return;

          const tokens = splitDoctorTokensSafely_(rawDocsStr);
          const movableForThisClinic = [];
          let bestScore = 0;

          tokens.forEach(function(tok) {
            const k = normKeyForDoc_(tok);
            if (!k || k === '未充足' || masters.absoluteNgDocSet.has(k)) return;
            if (masters.streakNgMap.has(k) && masters.streakNgMap.get(k).includes(dateStr)) return;

            const isFt = tok.includes('常勤');
            if (isFt) {
              const condStr = masters.ftAllowCondMap.get(k) || '';
              if (!masters.ftAllowDocSet.has(k) || !tok.includes('所定休出') || !isFtConditionCompatibleWithTarget_(condStr, cleanTargetClinic, area)) {
                return;
              }
            }

            if (!isDoctorCompatibleWithSlot_(k, dow, cleanTargetClinic, area, timeInfo, '移動', masters)) return;

            const isSbu = masters.sbuSet.has(k);
            const isNeg = masters.negotiableMap.has(k);
            const isSpecialMovable = isSbu || isNeg || isFt;

            // 一般非常勤は優先度S・A（移動60分以内）のみ、広域Bは交渉しやすい医師・SBU・所定休出のみ
            if (pairEval.rank === 'B' && !isSpecialMovable) return;

            const docScore = pairEval.score + (isSbu ? 15 : (isNeg ? 10 : 0));
            if (docScore > bestScore) bestScore = docScore;

            const label = isSbu ? (k + '[★SBU]') : (isFt ? (k + '[常勤・所定休出]') : k);
            movableForThisClinic.push({ label: label, isSpecial: isSpecialMovable });
          });

          if (movableForThisClinic.length === 0) return;

          if (docCount >= 2 && !seenTwoDoc.has(cleanSrcClinic)) {
            seenTwoDoc.add(cleanSrcClinic);
            twoDocScoredList.push({
              score: bestScore,
              item: {
                clinic: cleanSrcClinic,
                area: sArea,
                priorityRank: pairEval.rank,
                estMinutes: pairEval.estMinutes,
                movableDoctors: movableForThisClinic.map(function(d) { return d.label; })
              }
            });
          } else if (docCount === 1 && !seenOneDoc.has(cleanSrcClinic)) {
            const specOnly = movableForThisClinic.filter(function(d) { return d.isSpecial; });
            if (specOnly.length > 0) {
              seenOneDoc.add(cleanSrcClinic);
              oneDocScoredList.push({
                score: bestScore,
                item: {
                  clinic: cleanSrcClinic,
                  area: sArea,
                  priorityRank: pairEval.rank,
                  estMinutes: pairEval.estMinutes,
                  movableDoctors: specOnly.map(function(d) { return d.label; })
                }
              });
            }
          }
        });
      });
    }

    twoDocScoredList.sort(function(a, b) { return b.score - a.score; });
    oneDocScoredList.sort(function(a, b) { return b.score - a.score; });

    // ③ 個別打診候補（確定シフト突合 ＆ 移動60分以内・休憩60分確保ルール）
    const cleanK = clinic.includes('北葛西') ? (isInternalMed ? '北葛西（内科）' : '北葛西（小児科）')
      : (clinic.includes('亀有') ? (isInternalMed ? '亀有（内科）' : '亀有（小児科）') : cleanTargetClinic);
    const rawSpots = spotMap.get(dateStr + '_' + cleanK + '_' + absTime) || [];
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

          // 移動時間60分以内（優先度S・A）かつ、空き時間から移動時間を引いて60分以上の休憩が取れる場合のみ許可
          if ((pairMobility.rank !== 'S' && pairMobility.rank !== 'A') ||
              pairMobility.estMinutes > 60 ||
              (gapMin - pairMobility.estMinutes) < 60) {
            isImpossibleSchedule = true;
            break;
          }
        }
        otherConfirmedLabels.push(cs.clinic + '(' + cs.timeLabel + ')');
      }

      if (isImpossibleSchedule) return;

      verifiedSpots.push({
        rank: sp.rank,
        doctor: dKey,
        sameDowSlotCount: sp.sameDowSlotCount,
        totalCount: sp.totalCount,
        daySchedule: otherConfirmedLabels.length > 0 ? ('同日確定枠:' + otherConfirmedLabels.join(', ')) : '当日空き',
        masterNote: sp.rawNote
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

  const targetBundles = bundledCandidates.filter(function(r) {
    const dKey = normKeyForDoc_(r[0]);
    if (masters.absoluteNgDocSet.has(dKey)) return false;
    const text = String(r[4] || '');
    return Array.from(targetDatesSet).some(function(d) { return text.includes(d); });
  }).slice(0, 12).map(function(r) {
    const dKey = normKeyForDoc_(r[0]);
    const extraMemo = masters.docMemoLookup.get(dKey) || '';
    return {
      doctor: dKey,
      licenseNo: r[1],
      slotCount: r[3],
      targetSlotsText: r[4],
      masterNote: [r[5] || '', extraMemo].filter(Boolean).join(' ')
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
    targetBundles: targetBundles,
    streakList: streakList,
    recentFlaggedKb: recentFlaggedKb
  };
}