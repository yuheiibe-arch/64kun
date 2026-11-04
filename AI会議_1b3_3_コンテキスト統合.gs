/**
 * ============================================================================
 * AI会議_1b3_3_コンテキスト統合.gs
 * （各不在枠に対するアプローチ候補の統合・連勤リスト・ナレッジ抽出のメイン構築）
 * ============================================================================
 */

function makeNormalizedSpotKey_(dateStr, rawClinicStr, absTimeStr) {
  const cleanD = normalizeDateYmd1b3_(dateStr);
  const cleanC = cleanClinicBaseName_(rawClinicStr);
  const dept = String(rawClinicStr || '').includes('内科') ? '内科' : '小児科';
  return cleanD + '_' + cleanC + '_' + dept + '_' + String(absTimeStr || '').trim();
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
            if (canDoctorLeaveSourceWithoutCreatingAbsence_(k, s, timeInfo, pairEval)) {
              movableForTwoDoc.push(label);
              if (docScore > bestScore) bestScore = docScore;
            }
          } else if (docCount === 1 && isSpecialMovable) {
            const isKitakasai = cleanSrcClinic.includes('北葛西');
            const docHours = parseDocTokenWorkingMinutes1b3_(tok, primaryBand, isKitakasai);
            if (docHours.rawStartMin <= timeInfo.startMin && docHours.rawEndMin >= timeInfo.endMin) {
              movableForOneDoc.push(label);
              if (docScore > bestScore) bestScore = docScore;
            }
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