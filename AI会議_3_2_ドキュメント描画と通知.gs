/**
 * ============================================================================
 * AI会議_1b3_2_移動と隣接枠判定.gs
 * 1) BQ確定シフト(kakutei)の実勤務時刻を統合した正確な在席時間判定
 * 2) 移動元・移動先のどちらかに15分でも「部分的医師不在」が生じる移動案を完全ブロック
 * 3) 他帯と跨がる通し勤務（09:00-21:00 / 09:00-18:00 / 15:00-21:00等）の分断移動をブロック
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

/**
 * BQパックの確定シフト（kakutei / obo）から「日付|拠点|医師キー -> [{startMin, endMin}]」の正確な時刻辞書を構築する
 */
function buildBqExactShiftIntervalsMap1b3_() {
  const map = new Map();
  if (typeof fetchAndDecryptBqMailForResolver_ !== 'function') return map;

  try {
    const bqData = fetchAndDecryptBqMailForResolver_();
    if (!bqData) return map;

    const list = (bqData.kakutei || []).concat(bqData.obo || []);
    list.forEach(function(s) {
      const st = String(s.st || s.status || '').trim();
      if (st && (st.includes('キャンセル') || st.includes('取消') || st.includes('却下'))) return;

      const d = normalizeDateYmd1b3_(s.d || s.date);
      const rawC = String(s.c || s.clinic || '').trim();
      const cleanC = cleanClinicBaseName_(rawC);
      const docKey = normKeyForDoc_(s.n || s.name || '');
      if (!d || !cleanC || cleanC.includes('バックアップ') || !docKey || docKey === '未充足') return;

      const rawT = String(s.t || s.time || s.tm || s.workTime || '').normalize('NFKC').trim();
      const tm = rawT.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
      if (!tm) return;

      const startMin = Number(tm[1]) * 60 + Number(tm[2]);
      const endMin = Number(tm[3]) * 60 + Number(tm[4]);
      if (endMin <= startMin) return;

      const key = d + '|' + cleanC + '|' + docKey;
      if (!map.has(key)) map.set(key, []);
      const arr = map.get(key);
      if (!arr.some(function(x) { return x.startMin === startMin && x.endMin === endMin; })) {
        arr.push({ startMin: startMin, endMin: endMin });
      }
    });
  } catch (e) {
    Logger.log('⚠️ BQ確定シフト時刻マップ読込スキップ: ' + e.message);
  }
  return map;
}

/**
 * セル文字列の時刻表記 ＋ BQ確定シフトの実時刻を統合して、対象バンド内での正確な勤務開始・終了分を返す
 */
function parseDocTokenWorkingMinutes1b3_(tok, bandCode, isKitakasai, allRowTextHint, bqIntervalsOpt) {
  const defaultStart = bandCode === 'A' ? 540 : (bandCode === 'B' ? 900 : 1080);
  const defaultEnd = bandCode === 'A' ? (isKitakasai ? 750 : 780) : (bandCode === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

  const targetStr = String(tok || '');
  let tm = targetStr.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);

  if (!tm && allRowTextHint) {
    const cleanName = cleanDisplayDoctorName1b3_(tok);
    if (cleanName) {
      const escaped = cleanName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const hintRegex = new RegExp(escaped + '[^\\n\\/／,、]*?(\\d{1,2}):(\\d{2})\\s*[\\-〜～~]\\s*(\\d{1,2}):(\\d{2})');
      tm = String(allRowTextHint).match(hintRegex);
    }
  }

  if (tm) {
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

  // ★セルに時刻記載がない場合でも、BQ確定シフト（kakutei）に実時刻があればそれを使用する
  if (Array.isArray(bqIntervalsOpt) && bqIntervalsOpt.length > 0) {
    const overlapping = bqIntervalsOpt.filter(function(iv) {
      return Math.max(iv.startMin, defaultStart) < Math.min(iv.endMin, defaultEnd);
    });
    if (overlapping.length > 0) {
      const minRawStart = Math.min.apply(null, overlapping.map(function(iv) { return iv.startMin; }));
      const maxRawEnd = Math.max.apply(null, overlapping.map(function(iv) { return iv.endMin; }));
      return {
        startMin: Math.max(defaultStart, minRawStart),
        endMin: Math.min(defaultEnd, maxRawEnd),
        rawStartMin: minRawStart,
        rawEndMin: maxRawEnd,
        hasCustomTime: (minRawStart > defaultStart || maxRawEnd < defaultEnd || minRawStart < defaultStart || maxRawEnd > defaultEnd)
      };
    }
  }

  return {
    startMin: defaultStart,
    endMin: defaultEnd,
    rawStartMin: defaultStart,
    rawEndMin: defaultEnd,
    hasCustomTime: false
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

/**
 * ★厳格版：移動元・移動先のどちらにも「15分たりとも部分的医師不在（空白時間）」を作らない場合のみ canMove: true を返す
 */
function evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters) {
  const targetBands = timeInfo.targetBands || ['A'];
  const srcDate = normalizeDateYmd1b3_(srcRow[0]);
  const srcClinic = cleanClinicBaseName_(srcRow[2]);
  const isKitakasai = srcClinic.includes('北葛西');
  const allRowTextHint = [srcRow[7], srcRow[8], srcRow[9]].join(' / ');
  const bqMap = (masters && masters.bqExactIntervalsMap) ? masters.bqExactIntervalsMap : null;

  const getBqIntervals = function(dKey) {
    if (!bqMap || !srcDate || !srcClinic || !dKey) return null;
    return bqMap.get(srcDate + '|' + srcClinic + '|' + dKey) || null;
  };

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

  // 1. 終日通し勤務（A・B・Cすべて在席）、または対象外の他帯にも勤務がある医師を分断する移動は原則ブロック
  const inA = docsByBand['A'].includes(docKey);
  const inB = docsByBand['B'].includes(docKey);
  const inC = docsByBand['C'].includes(docKey);
  if (inA && inB && inC) {
    return { canMove: false, conditionNote: '' };
  }

  const movingBqIvs = getBqIntervals(docKey);
  if (Array.isArray(movingBqIvs)) {
    const spansFullDayInBq = movingBqIvs.some(function(iv) { return iv.startMin <= 600 && iv.endMin >= 1200; });
    if (spansFullDayInBq) {
      return { canMove: false, conditionNote: '' };
    }
  }

  // 2. 対象バンド外（例：午前移動なのに午後・夜間も入っている等）のチェック
  const allBands = ['A', 'B', 'C'];
  for (let i = 0; i < allBands.length; i++) {
    const b = allBands[i];
    if (targetBands.includes(b)) continue;
    if (!docsByBand[b].includes(docKey)) continue;

    // B(午後)とC(夜間)が連続しているシフトの片方引き抜きは不可
    if ((b === 'C' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('C'))) {
      return { canMove: false, conditionNote: '' };
    }
    // A(午前)とC(夜間)の両方に入っている場合（通し・中抜け）の片方引き抜きも不可
    if ((b === 'C' && targetBands.includes('A')) || (b === 'A' && targetBands.includes('C'))) {
      return { canMove: false, conditionNote: '' };
    }
    // A(午前)とB(午後)の両方に入っている場合は、昼休みに移動可能な近隣（SまたはAかつ移動45分以内）のみ許可
    if ((b === 'A' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('A'))) {
      if (pairEval.estMinutes > 45 || (pairEval.rank !== 'S' && pairEval.rank !== 'A')) {
        return { canMove: false, conditionNote: '' };
      }
    }
  }

  // 3. 対象バンドごとの「移動先のフルカバー」＆「移動元の部分不在ゼロ」厳密チェック
  for (let i = 0; i < targetBands.length; i++) {
    const tb = targetBands[i];
    if (countByBand[tb] < 2 || !docsByBand[tb].includes(docKey)) {
      return { canMove: false, conditionNote: '' };
    }

    const tokens = rawTokensByBand[tb];
    const bandFullStart = tb === 'A' ? 540 : (tb === 'B' ? 900 : 1080);
    const bandFullEnd = tb === 'A' ? (isKitakasai ? 750 : 780) : (tb === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

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

    // ★チェックA：移動する医師本人が、移動先の要求時間帯（例：板橋 09:00-13:00）を1分も欠けずフルカバーしているか？
    // （例：上春先生が11:00-13:00勤務なら rawStartMin=660 > reqStart=540 となり、09:00-11:00が不在になるためここで即ブロック！）
    const movingHours = parseDocTokenWorkingMinutes1b3_(movingDocTok, tb, isKitakasai, allRowTextHint, movingBqIvs);
    const reqStart = Math.max(bandFullStart, timeInfo.startMin);
    const reqEnd = Math.min(bandFullEnd, timeInfo.endMin);
    if (movingHours.rawStartMin > reqStart || movingHours.rawEndMin < reqEnd) {
      return { canMove: false, conditionNote: '' };
    }

    // 移動する医師が帯を跨ぐ変則シフト（例：11:00-15:00、14:00-19:00 等）の場合も分断不可
    if (movingHours.hasCustomTime) {
      if (tb === 'A' && movingHours.rawEndMin > 780 && !targetBands.includes('B')) return { canMove: false, conditionNote: '' };
      if (tb === 'B' && ((movingHours.rawStartMin < 900 && !targetBands.includes('A')) || (movingHours.rawEndMin > 1080 && !targetBands.includes('C')))) {
        return { canMove: false, conditionNote: '' };
      }
      if (tb === 'C' && movingHours.rawStartMin < 1080 && !targetBands.includes('B')) return { canMove: false, conditionNote: '' };
    }

    const validRemainingTokens = remainingTokens.filter(function(remTok) {
      const remKey = normKeyForDoc_(remTok);
      if (tb === 'A' && docsByBand['A'].includes(remKey) && docsByBand['B'].includes(remKey) && docsByBand['C'].includes(remKey)) {
        const remIvs = getBqIntervals(remKey);
        const startsAtNineInBq = Array.isArray(remIvs) && remIvs.some(function(iv) { return iv.startMin <= 540; });
        if (!/09:00|9:00/.test(remTok) && !startsAtNineInBq) return false;
      }
      return true;
    });

    if (validRemainingTokens.length === 0) {
      return { canMove: false, conditionNote: '' };
    }

    // ★チェックB：移動元の拠点（例：南浦和）で、残る医師だけでバンド全域（09:00-13:00等）を15分刻みすべて隙間なくカバーできているか？
    // 15分でも医師0人の時間帯（部分的医師不在）が生じる場合は、即座に canMove: false で除外する！
    for (let tMin = bandFullStart; tMin < bandFullEnd; tMin += 15) {
      const coveredAtT = validRemainingTokens.some(function(remTok) {
        const remKey = normKeyForDoc_(remTok);
        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai, allRowTextHint, getBqIntervals(remKey));
        if (!remHours.hasCustomTime && tb === 'C' && docsByBand['B'].includes(remKey) && !docsByBand['A'].includes(remKey)) {
          return remHours.startMin <= tMin && 1200 >= (tMin + 15);
        }
        return remHours.startMin <= tMin && remHours.endMin >= (tMin + 15);
      });

      if (!coveredAtT) {
        return { canMove: false, conditionNote: '' };
      }
    }
  }

  return {
    canMove: true,
    conditionNote: ''
  };
}

function canDoctorLeaveSourceWithoutCreatingAbsence_(docKey, srcRow, timeInfo, pairEval, masters) {
  const res = evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters);
  return res.canMove;
}