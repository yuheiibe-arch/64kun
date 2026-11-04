/**
 * ============================================================================
 * AI会議_1b3_2_移動と隣接枠判定.gs
 * （前後医の延長・早出判定、および2診目崩しの分単位フルカバー＆玉突き不在防止判定）
 * ============================================================================
 */

/**
 * 医師トークン（例: "高橋将也(10:00-21:00)" や "太田遼[SBU]"）から、
 * その帯における実際の勤務開始分(startMin)と終了分(endMin)を解析するヘルパー
 */
function parseDocTokenWorkingMinutes1b3_(tok, bandCode, isKitakasai) {
  const defaultStart = bandCode === 'A' ? 540 : (bandCode === 'B' ? 900 : 1080);
  const defaultEnd = bandCode === 'A' ? 780 : (bandCode === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

  const tm = String(tok || '').match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
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

/**
 * 自拠点の前医（延長）・後医（早出）の打診可否を判定する関数
 */
function evaluateValidAdjacentDocsForSlot_(docsStr, actionType, dateStr, dow, targetClinic, targetArea, timeInfo, masters) {
  const tokens = splitDoctorTokensSafely_(docsStr);
  const validTokens = [];
  const validDocKeys = new Set();

  tokens.forEach(function(tok) {
    const k = normKeyForDoc_(tok);
    if (!k || masters.absoluteNgDocSet.has(k)) return;
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

    if (!isDoctorCompatibleWithSlot_(k, dow, targetClinic, targetArea, timeInfo, actionType, masters)) return;

    const headPart = String(tok).replace(/\[.*$/g, '');
    const isFt = !headPart.includes('非常勤') && headPart.includes('常勤');

    if (isFt) {
      const condStr = masters.ftAllowCondMap.get(k) || '';
      if (masters.ftAllowDocSet.has(k) && headPart.includes('所定休出') && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
        validTokens.push(k + '(常勤・所定休出)');
        validDocKeys.add(k);
      }
      return;
    }

    validTokens.push(k + '(非常勤)');
    validDocKeys.add(k);
  });

  return {
    docs: validTokens.length > 0 ? validTokens.join(' / ') : 'なし',
    status: validTokens.length > 0 ? (actionType + '打診可能') : '',
    docKeySet: validDocKeys
  };
}

/**
 * 2診目崩しで医師(docKey)を移動させた際、
 * 1) 移動する医師自身が移動先の不在時間(timeInfo.startMin〜endMin)を100%フルカバーできるか？
 * 2) その医師が抜けた後、移動元の拠点でその帯の全時間を15分刻みで隙間なくカバーする残存医師がいるか？
 * 3) 通し勤務（10:00-21:00等）や別帯直結の医師を無理に引き剥がしていないか？
 * を厳密に検証する関数
 */
function canDoctorLeaveSourceWithoutCreatingAbsence_(docKey, srcRow, timeInfo, pairEval) {
  const targetBands = timeInfo.targetBands || ['A'];
  const srcClinic = cleanClinicBaseName_(srcRow[2]);
  const isKitakasai = srcClinic.includes('北葛西');

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

  for (let i = 0; i < targetBands.length; i++) {
    const tb = targetBands[i];
    if (countByBand[tb] < 2 || !docsByBand[tb].includes(docKey)) {
      return false;
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

    if (!movingDocTok || remainingTokens.length === 0) return false;

    // 1. 移動する医師自身が、移動先の不在枠（timeInfo.startMin〜endMin）を最初から最後までカバーできるか？
    const movingHours = parseDocTokenWorkingMinutes1b3_(movingDocTok, tb, isKitakasai);
    const reqStart = Math.max(bandFullStart, timeInfo.startMin);
    const reqEnd = Math.min(bandFullEnd, timeInfo.endMin);
    if (movingHours.rawStartMin > reqStart || movingHours.rawEndMin < reqEnd) {
      return false;
    }

    // 2. 通し勤務表記（例: 10:00-21:00）になっている医師は一部帯だけ抜くことはできないため除外
    if (movingHours.hasCustomTime) {
      if (tb === 'A' && movingHours.rawEndMin > 780) return false;
      if (tb === 'B' && (movingHours.rawStartMin < 900 || movingHours.rawEndMin > 1080)) return false;
      if (tb === 'C' && movingHours.rawStartMin < 1080) return false;
    }

    // 3. その医師が抜けた後、移動元のクリニックでその帯（bandFullStart〜bandFullEnd）に15分でも0人の穴が空かないか？
    for (let tMin = bandFullStart; tMin < bandFullEnd; tMin += 15) {
      const coveredAtT = remainingTokens.some(function(remTok) {
        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai);
        return remHours.startMin <= tMin && remHours.endMin >= (tMin + 15);
      });
      if (!coveredAtT) {
        return false;
      }
    }
  }

  // 4. 別帯（A・B・C）にも同じ医師が入っている場合のチェック
  const allBands = ['A', 'B', 'C'];
  for (let i = 0; i < allBands.length; i++) {
    const b = allBands[i];
    if (targetBands.includes(b)) continue;
    if (!docsByBand[b].includes(docKey)) continue;

    if ((b === 'C' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('C'))) {
      return false;
    }

    if ((b === 'A' && targetBands.includes('B')) || (b === 'B' && targetBands.includes('A'))) {
      if (pairEval.estMinutes > 60 || (pairEval.rank !== 'S' && pairEval.rank !== 'A')) {
        return false;
      }
    }
  }

  return true;
}