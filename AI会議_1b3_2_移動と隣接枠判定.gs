/**
 * ============================================================================
 * AI会議_1b3_2_移動と隣接枠判定.gs
 * （残存医が20:00上がり等の場合に「残存医の延長打診＋2診目移動打診」のセット案を構築する版）
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
 * 医師トークン（例: "野口優輔(17:00-20:00)" や "井上信明(18:00-21:00)"）から、
 * その帯における実際の勤務開始分(startMin)と終了分(endMin)を解析するヘルパー
 */
function parseDocTokenWorkingMinutes1b3_(tok, bandCode, isKitakasai, allRowTextHint) {
  const defaultStart = bandCode === 'A' ? 540 : (bandCode === 'B' ? 900 : 1080);
  const defaultEnd = bandCode === 'A' ? 780 : (bandCode === 'B' ? 1080 : (isKitakasai ? 1200 : 1260));

  let targetStr = String(tok || '');
  let tm = targetStr.match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);

  // もしC帯のトークン自体に時刻が書かれていなくても、同じ行のB帯（または行全体ヒント）に「野口優輔(17:00-20:00)」のような時刻表記があればそれを引き継ぐ
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

    if (!isDoctorCompatibleWithSlot_(k, dow, targetClinic, targetArea, timeInfo, actionType, masters)) return;

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
 * ★拡張版：2診目崩しで医師(docKey)を移動させた際の検証関数
 * 戻り値:
 *   { canMove: false } -> 移動不可
 *   { canMove: true, conditionNote: "" } -> 無条件で2診目崩し可能
 *   { canMove: true, conditionNote: "野口優輔先生へ21:00まで延長打診とセット" } -> 残存医への延長・早出打診とセットで2診目崩し可能！
 */
function evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters) {
  const targetBands = timeInfo.targetBands || ['A'];
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

  // 移動候補の医師(docKey)自身がA・B・Cの3帯通し勤務の場合は引き抜き不可
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

    // 1. 移動する医師自身が、移動先の不在枠（timeInfo.startMin〜endMin）をフルカバーできるか？
    const movingHours = parseDocTokenWorkingMinutes1b3_(movingDocTok, tb, isKitakasai, allRowTextHint);
    const reqStart = Math.max(bandFullStart, timeInfo.startMin);
    const reqEnd = Math.min(bandFullEnd, timeInfo.endMin);
    if (movingHours.rawStartMin > reqStart || movingHours.rawEndMin < reqEnd) {
      return { canMove: false, conditionNote: '' };
    }

    // 2. 移動する医師自身が別帯またぎの変則通し勤務（例：17:00-20:00や10:00-21:00）なら引き抜き不可
    if (movingHours.hasCustomTime) {
      if (tb === 'A' && movingHours.rawEndMin > 780) return { canMove: false, conditionNote: '' };
      if (tb === 'B' && (movingHours.rawStartMin < 900 || movingHours.rawEndMin > 1080)) return { canMove: false, conditionNote: '' };
      if (tb === 'C' && movingHours.rawStartMin < 1080) return { canMove: false, conditionNote: '' };
    }

    // 3. 残る医師（remainingTokens）による移動元カバー判定
    const validRemainingTokens = remainingTokens.filter(function(remTok) {
      const remKey = normKeyForDoc_(remTok);
      // A帯の2診目崩しにおいて、残る医師がA・B・C全帯に入っており時刻明記がない場合は10:00開始（10:00-21:00）とみなして無条件カバーから外す
      if (tb === 'A' && docsByBand['A'].includes(remKey) && docsByBand['B'].includes(remKey) && docsByBand['C'].includes(remKey)) {
        if (!/09:00|9:00/.test(remTok)) return false;
      }
      return true;
    });

    if (validRemainingTokens.length === 0) {
      return { canMove: false, conditionNote: '' };
    }

    // 15分刻みで移動元に穴（未カバー時間帯）が空かないかチェック
    let uncoveredStartMin = null;
    let uncoveredEndMin = null;

    for (let tMin = bandFullStart; tMin < bandFullEnd; tMin += 15) {
      const coveredAtT = validRemainingTokens.some(function(remTok) {
        const remKey = normKeyForDoc_(remTok);
        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai, allRowTextHint);
        // ★もしC帯(18:00-21:00)で、残る医師(例:野口優輔先生)がB帯とC帯の両方に入っており、かつ時刻明記がない場合、
        // 「17:00-20:00」等の変則シフトで20:00(1200分)に上がる可能性を考慮する
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

    // ★穴が空く場合：残る医師の「延長」または「早出」（90分以内）で埋められるかを検証！
    // （例：17:00-20:00の野口優輔先生に20:00-21:00の延長を打診すれば、18:00-21:00の井上信明先生を移動できる！）
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

        // 通常常勤は延長・早出不可のため除外（非常勤または所定休出のみ可）
        const isFt = !remTok.includes('非常勤') && remTok.includes('常勤') && !remTok.includes('所定休出');
        if (isFt) return;

        const remHours = parseDocTokenWorkingMinutes1b3_(remTok, tb, isKitakasai, allRowTextHint);
        const effectiveEnd = (!remHours.hasCustomTime && tb === 'C' && docsByBand['B'].includes(remKey) && !docsByBand['A'].includes(remKey))
          ? 1200
          : remHours.endMin;

        // パターンA：帯の末尾（例: 20:00-21:00）が空く場合 ➔ 残存医が uncoveredStartMin(20:00) まで勤務していれば「延長打診」でカバー可能！
        if (uncoveredEndMin === bandFullEnd && effectiveEnd === uncoveredStartMin) {
          bridgeDocFound = remDisp + '先生へ' + formatMinToHm1b3_(bandFullEnd) + 'まで延長打診を確認後';
        }
        // パターンB：帯の先頭（例: 09:00-10:00）が空く場合 ➔ 残存医が uncoveredEndMin(10:00) から勤務していれば（ただし通し勤務でない場合）「早出打診」でカバー可能！
        else if (uncoveredStartMin === bandFullStart && remHours.startMin === uncoveredEndMin && !(docsByBand['A'].includes(remKey) && docsByBand['B'].includes(remKey) && docsByBand['C'].includes(remKey))) {
          bridgeDocFound = remDisp + '先生へ' + formatMinToHm1b3_(bandFullStart) + 'から早出打診を確認後';
        }
      });

      if (!bridgeDocFound) {
        return { canMove: false, conditionNote: '' };
      }
      conditionNotes.push(bridgeDocFound);
    }
  }

  // 別帯（A・B・C）にも移動候補の医師が入っている場合のチェック
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

// 既存互換用ラッパー
function canDoctorLeaveSourceWithoutCreatingAbsence_(docKey, srcRow, timeInfo, pairEval, masters) {
  const res = evaluateDoctorLeaveSourceCondition1b3_(docKey, srcRow, timeInfo, pairEval, masters);
  return res.canMove && !res.conditionNote;
}