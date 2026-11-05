/**
 * ============================================================================
 * AI会議_1b3_2_移動と隣接枠判定.gs
 * （日別メール制約［11/8柿本先生の18時まで等］＆残存医の延長打診＋2診目移動打診セット対応版）
 * ============================================================================
 */

function cleanDisplayDoctorName1b3_(rawName) {
  return String(rawName || '')
    .normalize('NFKC')
    .replace(/[\(（\[【].*$/g, '')
    .replace(/[:：].*$/, '')
    .replace(/[\)）\]】★⚠️📩・]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function parseDocTokenWorkingMinutes1b3_(tok, bandCode, isKitakasai, allRowTextHint) {
  const defaultStart = bandCode === 'A' ? 540 : (bandCode === 'B' ? 900 : 1080);
  const defaultEnd = bandCode === 'A' ? 780 : (bandCode === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

  let targetStr = String(tok || '');
  let tm = targetStr.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);

  if (!tm && allRowTextHint) {
    const cleanName = cleanDisplayDoctorName1b3_(tok);
    if (cleanName) {
      const escaped = cleanName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const hintRegex = new RegExp(escaped + '[^\\n\\/／,、]*?(\\d{1,2}):(\\d{2})\\s*[\\-〜～~]\\s*(\\d{1,2}):(\\d{2})');
      tm = String(allRowTextHint).match(hintRegex);
    }
  }

  if (!tm) {
    return {
      startMin: defaultStart,
      endMin: defaultEnd,
      rawStartMin: defaultStart,
      rawEndMin: defaultEnd,
      hasCustomTime: false
    };
  }

  const rawStart = Number(tm[1]) * 60 + Number(tm[2]);
  const rawEnd = Number(tm[3]) * 60 + Number(tm[4]);

  return {
    startMin: Math.max(defaultStart, rawStart),
    endMin: Math.min(defaultEnd, rawEnd),
    rawStartMin: rawStart,
    rawEndMin: rawEnd,
    hasCustomTime: true
  };
}

function evaluateValidAdjacentDocsForSlot_(docsStr, actionType, dateStr, dow, targetClinic, targetArea, timeInfo, masters) {
  const tokens = splitDoctorTokensSafely_(docsStr);
  const validTokens = [];
  const validDocKeys = new Set();

  tokens.forEach(function(tok) {
    const k = normKeyForDoc_(tok);
    const dispName = cleanDisplayDoctorName1b3_(tok);
    if (!k || !dispName || masters.absoluteNgDocSet.has(k)) return;
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

    // ★dateStr を渡して「11/8の柿本先生は18時までなら可能（＝18時以降の延長不可）」を確実にチェック！
    if (!isDoctorCompatibleWithSlot_(k, dow, targetClinic, targetArea, timeInfo, actionType, masters, dateStr)) return;

    const headPart = String(tok).replace(/\[.*$/g, '');
    const isFt = !headPart.includes('非常勤') && headPart.includes('常勤');

    if (isFt) {
      const condStr = masters.ftAllowCondMap.get(k) || '';
      if (masters.ftAllowDocSet.has(k) && headPart.includes('所定休出') && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
        validTokens.push(dispName + '(常勤・所定休出)');
        validDocKeys.add(k);
      }
      return;
    }

    validTokens.push(dispName + '(非常勤)');
    validDocKeys.add(k);
  });

  return {
    docs: validTokens.length > 0 ? validTokens.join(' / ') : 'なし',
    status: validTokens.length > 0 ? (actionType + '打診可能') : '',
    docKeySet: validDocKeys
  };
}

function formatMinToHm1b3_(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

function evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters) {
  const targetBands = timeInfo.targetBands || ['A'];
  const srcDate = normalizeDateYmd1b3_(srcRow[0]);
  const srcClinic = cleanClinicBaseName_(srcRow[2]);
  const isKitakasai = srcClinic.includes('北葛西');
  const allRowTextHint = [srcRow[7], srcRow[8], srcRow[9]].join(' / ');

  const countByBand = {
    'A': Number(srcRow[4] || 0),
    'B': Number(srcRow[5] || 0),
    'C': Number(srcRow[6] || 0)
  };
  const rawTokensByBand = {
    'A': splitDoctorTokensSafely_(srcRow[7]),
    'B': splitDoctorTokensSafely_(srcRow[8]),
    'C': splitDoctorTokensSafely_(srcRow[9])
  };
  const docsByBand = {
    'A': rawTokensByBand['A'].map(normKeyForDoc_),
    'B': rawTokensByBand['B'].map(normKeyForDoc_),
    'C': rawTokensByBand['C'].map(normKeyForDoc_)
  };

  const inA = docsByBand['A'].includes(docKey);
  const inB = docsByBand['B'].includes(docKey);
  const inC = docsByBand['C'].includes(docKey);
  if (inA && inB && inC) {
    return { canMove: false, conditionNote: '' };
  }

  const conditionNotes = [];

  for (let i = 0; i < targetBands.length; i++) {
    const tb = targetBands[i];
    if (countByBand[tb] < 2 || !docsByBand[tb].includes(docKey)) {
      return { canMove: false, conditionNote: '' };
    }

    const tokens = rawTokensByBand[tb];
    const bandFullStart = tb === 'A' ? 540 : (tb === 'B' ? 900 : 1080);
    const bandFullEnd = tb === 'A' ? 780 : (tb === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

    let movingDocTok = null;
    const remainingTokens = [];
    tokens.forEach(function(t) {
      if (normKeyForDoc_(t) === docKey && !movingDocTok) {
        movingDocTok = t;
      } else {
        remainingTokens.push(t);
      }
    });

    if (!movingDocTok || remainingTokens.length === 0) {
      return { canMove: false, conditionNote: '' };
    }

    const movingHours = parseDocTokenWorkingMinutes1b3_(movingDocTok, tb, isKitakasai, allRowTextHint);
    const reqStart = Math.max(bandFullStart, timeInfo.startMin);
    const reqEnd = Math.min(bandFullEnd, timeInfo.endMin);
    if (movingHours.rawStartMin > reqStart || movingHours.rawEndMin < reqEnd) {
      return { canMove: false, conditionNote: '' };
    }

    if (movingHours.hasCustomTime) {
      if (tb === 'A' && movingHours.rawEndMin > 780) return { canMove: false, conditionNote: '' };
      if (tb === 'B' && (movingHours.rawStartMin < 900 || movingHours.rawEndMin > 1080)) return { canMove: false, conditionNote: '' };
      if (tb === 'C' && movingHours.rawStartMin < 1080) return { canMove: false, conditionNote: '' };
    }

    const validRemainingTokens = remainingTokens.filter(function(remTok) {
      const remKey = normKeyForDoc_(remTok);
      if (tb === 'A' && docsByBand['A'].includes(remKey) && docsByBand['B'].includes(remKey) && docsByBand['C'].includes(remKey)) {
        if (!/09:00|9:00/.test(remTok)) return false;
      }
      return true;
    });

    if (validRemainingTokens.length === 0) {
      return { canMove: false, conditionNote: '' };
    }

    let uncoveredStartMin = null;
    let uncoveredEndMin = null;

    for (let tMin = bandFullStart; tMin < bandFullEnd; tMin += 15) {
      const coveredAtT = validRemainingTokens.some(function(remTok) {
        const remKey = normKeyForDoc_(remTok);
        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai, allRowTextHint);
        if (!remHours.hasCustomTime && tb === 'C' && docsByBand['B'].includes(remKey) && !docsByBand['A'].includes(remKey)) {
          return remHours.startMin <= tMin && 1200 >= (tMin + 15);
        }
        return remHours.startMin <= tMin && remHours.endMin >= (tMin + 15);
      });

      if (!coveredAtT) {
        if (uncoveredStartMin === null) uncoveredStartMin = tMin;
        uncoveredEndMin = tMin + 15;
      }
    }

    if (uncoveredStartMin !== null) {
      const gapDuration = uncoveredEndMin - uncoveredStartMin;
      if (gapDuration > 90) {
        return { canMove: false, conditionNote: '' };
      }

      let bridgeDocFound = null;
      validRemainingTokens.forEach(function(remTok) {
        if (bridgeDocFound) return;
        const remKey = normKeyForDoc_(remTok);
        const remDisp = cleanDisplayDoctorName1b3_(remTok);
        if (!remKey || !remDisp) return;
        if (masters && masters.absoluteNgDocSet && masters.absoluteNgDocSet.has(remKey)) return;

        // 残存医がその日にメール等で延長不可になっていないかも確認
        if (masters && masters.dateMailConstraintsMap && srcDate) {
          const dc = masters.dateMailConstraintsMap.get(srcDate + '_' + remKey);
          if (dc && (dc.noExtend || bandFullEnd > dc.maxEndMin)) return;
        }

        const isFt = !remTok.includes('非常勤') && remTok.includes('常勤') && !remTok.includes('所定休出');
        if (isFt) return;

        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai, allRowTextHint);
        const effectiveEnd = (!remHours.hasCustomTime && tb === 'C' && docsByBand['B'].includes(remKey) && !docsByBand['A'].includes(remKey))
          ? 1200
          : remHours.endMin;

        if (uncoveredEndMin === bandFullEnd && effectiveEnd === uncoveredStartMin) {
          bridgeDocFound = remDisp + '先生へ' + formatMinToHm1b3_(bandFullEnd) + 'まで延長打診を確認後';
        } else if (uncoveredStartMin === bandFullStart && remHours.startMin === uncoveredEndMin && !(docsByBand['A'].includes(remKey) && docsByBand['B'].includes(remKey) && docsByBand['C'].includes(remKey))) {
          bridgeDocFound = remDisp + '先生へ' + formatMinToHm1b3_(bandFullStart) + 'から早出打診を確認後';
        }
      });

      if (!bridgeDocFound) {
        return { canMove: false, conditionNote: '' };
      }
      conditionNotes.push(bridgeDocFound);
    }
  }

  const allBands = ['A', 'B', 'C'];
  for (let i = 0; i < allBands.length; i++) {
    const b = allBands[i];
    if (targetBands.includes(b)) continue;
    if (!docsByBand[b].includes(docKey)) continue;

    if ((b === 'C' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('C'))) {
      return { canMove: false, conditionNote: '' };
    }

    if ((b === 'A' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('A'))) {
      if (countByBand[b] < 2) return { canMove: false, conditionNote: '' };
      if (pairEval.estMinutes > 60 || (pairEval.rank !== 'S' && pairEval.rank !== 'A')) {
        return { canMove: false, conditionNote: '' };
      }
    }
  }

  return {
    canMove: true,
    conditionNote: conditionNotes.join('・')
  };
}

function canDoctorLeaveSourceWithoutCreatingAbsence_(docKey, srcRow, timeInfo, pairEval, masters) {
  const res = evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters);
  return res.canMove && !res.conditionNote;
}