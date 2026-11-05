/**
 * ============================================================================
 * AI会議_1b2_医師条件とタイムライン.gs
 * （タイムライン解析・シフトキャンセル日ブロック［設計1］・日別メール時間制約［設計2］統合版）
 * ============================================================================
 */

function splitDoctorTokensSafely_(rawStr) {
  const s = String(rawStr || '');
  if (!s || s === 'なし' || s === '未充足') return [];
  const results = [];
  let current = '';
  let depth = 0;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '(' || ch === '（' || ch === '[' || ch === '［' || ch === '{' || ch === '｛') {
      depth++;
      current += ch;
    } else if (ch === ')' || ch === '）' || ch === ']' || ch === '］' || ch === '}' || ch === '｝') {
      depth = Math.max(0, depth - 1);
      current += ch;
    } else if (depth === 0 && (ch === ',' || ch === '、' || ch === '，' || ch === '/' || ch === '／' || ch === '\n')) {
      if (current.trim()) results.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim()) results.push(current.trim());
  return results;
}

function parseAbsenceTimeline_(absTimeStr) {
  const m = String(absTimeStr || '').match(/(\d{1,2}):(\d{2})\s*[\-〜～]\s*(\d{1,2}):(\d{2})/);
  if (!m) {
    return { startMin: 540, endMin: 780, durationHours: 4, targetBands: ['A'] };
  }
  const startMin = Number(m[1]) * 60 + Number(m[2]);
  const endMin = Number(m[3]) * 60 + Number(m[4]);
  const durationHours = Math.round(((endMin - startMin) / 60) * 100) / 100;

  const targetBands = [];
  if (Math.max(startMin, 540) < Math.min(endMin, 780)) targetBands.push('A');
  if (Math.max(startMin, 900) < Math.min(endMin, 1080)) targetBands.push('B');
  if (Math.max(startMin, 1080) < Math.min(endMin, 1260)) targetBands.push('C');
  if (targetBands.length === 0) targetBands.push('B');

  return {
    startMin: startMin,
    endMin: endMin,
    durationHours: durationHours,
    targetBands: targetBands
  };
}

function getDowForDoc_(dateStr) {
  const m = String(dateStr || '').match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return '';
  const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return ['日', '月', '火', '水', '木', '金', '土'][dt.getDay()] || '';
}

function isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea) {
  const c = String(condStr || '').trim();
  if (!c) return true;
  if (c.includes('関西') && targetArea !== '関西') return false;
  if (targetArea === '関西' && (c.includes('関東') || c.includes('東京') || c.includes('埼玉') || c.includes('千葉') || c.includes('神奈川'))) return false;

  const onlyMatch = c.match(/([^、,，\s\/]+?)のみ/);
  if (onlyMatch) {
    const allowedKeyword = onlyMatch[1].replace(/エリア|拠点/g, '').trim();
    if (allowedKeyword && allowedKeyword !== '関東' && allowedKeyword !== '関西') {
      const cleanTarget = cleanClinicBaseName_(targetClinic);
      if (!cleanTarget.includes(allowedKeyword) && !targetArea.includes(allowedKeyword)) {
        return false;
      }
    }
  }
  return true;
}

/**
 * 医師マスタ条件 ＋ 【設計1：シフトキャンセル履歴】 ＋ 【設計2：メール返信の日別時間制約（18時まで等）】の統合適合チェック
 */
function isDoctorCompatibleWithSlot_(docKey, dow, targetClinic, targetArea, timeInfo, actionType, masters, dateStrOpt) {
  const cleanTarget = cleanClinicBaseName_(targetClinic);
  const cautionStructMap = masters.cautionStructMap;
  const negotiableMap = masters.negotiableMap;

  // ★設計1＆設計2：対象日（dateStrOpt）が指定されている場合、キャンセル日ブロックとメール時間制約をチェック
  if (dateStrOpt) {
    // 1) その日にシフトキャンセルを出している医師は終日提案不可（例：11/15・11/26の加藤雅崇先生）
    if (masters.canceledDatesByDoc && masters.canceledDatesByDoc.has(docKey)) {
      if (masters.canceledDatesByDoc.get(docKey).has(dateStrOpt)) {
        return false;
      }
    }
    // 2) メール返信でその日の時間上限（「18時までなら可能」等）や辞退がある医師をブロック（例：11/8の柿本優先生）
    if (masters.dateMailConstraintsMap) {
      const dc = masters.dateMailConstraintsMap.get(dateStrOpt + '_' + docKey);
      if (dc) {
        if (dc.fullDayDeclined && actionType === '個別') return false;
        if (timeInfo && timeInfo.endMin > dc.maxEndMin) return false;
        if (timeInfo && timeInfo.startMin < dc.minStartMin) return false;
        if (actionType === '延長' && dc.noExtend) return false;
        if (actionType === '早出' && dc.noEarly) return false;
        if (actionType === '移動' && dc.noMove) return false;
      }
    }
  }

  if (cautionStructMap.has(docKey)) {
    const c = cautionStructMap.get(docKey);
    if (c.condType === '絶対不可') return false;
    if (c.ngClinics.some(function(ngc) { return cleanTarget.includes(ngc) || ngc.includes(cleanTarget); })) return false;
    if (dow && c.ngDows.some(function(ngd) { return ngd.includes(dow); })) return false;
    const allCaution = c.ngTime + ' ' + c.memo;
    if (actionType === '延長' && /延長不可|延長.*難|以降.*不可|以降.*難/.test(allCaution)) return false;
    if (actionType === '早出' && /早出不可|前倒.*不可|より早く.*不可/.test(allCaution)) return false;
    if (actionType === '移動' && /移動不可|変更.*難/.test(allCaution)) return false;
    if (timeInfo.targetBands.includes('C') && /夜間不可|遅い勤務.*相談しない|18時以降不可|日中.*希望/.test(allCaution)) return false;
  }

  if (negotiableMap.has(docKey)) {
    const memo = negotiableMap.get(docKey);
    if (actionType === '移動' && /拠点移動は相談しない|移動不可|既存勤務のある箇所のみ/.test(memo)) {
      if (!memo.includes(cleanTarget)) return false;
    }
    if (timeInfo.targetBands.includes('C') && /遅い勤務.*相談しない/.test(memo)) return false;
    if (memo.includes('東京、埼玉エリアのみ') && targetArea !== '東京' && targetArea !== '埼玉') return false;
    if (memo.includes('神奈川県、東京のみ') && targetArea !== '神奈川' && targetArea !== '東京') return false;
    if (memo.includes('神奈川県の拠点であれば') && targetArea !== '神奈川') return false;
    if (memo.includes('埼玉県拠点であれば') && targetArea !== '埼玉') return false;
    if (memo.includes('関西勤務') && !memo.includes('関東') && targetArea !== '関西') return false;
  }

  return true;
}

function buildMasterAndConfirmedShiftIndex_(ss, activeShifts, streakWarnings) {
  const getRows = function(name) {
    const sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() >= 2) ? sh.getDataRange().getDisplayValues().slice(1) : [];
  };

  const docMemoLookup = new Map();
  const absoluteNgDocSet = new Set();
  const cautionStructMap = new Map();

  // 1. 『交渉注意医師』
  getRows('交渉注意医師').forEach(function(r) {
    const k = normKeyForDoc_(r[0]);
    if (!k) return;
    const cond = String(r[3] || '').trim();
    if (cond === '絶対不可') absoluteNgDocSet.add(k);

    const ngClinicsArr = [r[5], r[6], r[7], r[8], r[9]].map(function(v) { return String(v || '').trim(); }).filter(Boolean);
    const ngDowsArr = [r[11], r[12], r[13]].map(function(v) { return String(v || '').trim(); }).filter(Boolean);
    const ngTime = String(r[14] || '').trim();
    const noteK = String(r[10] || '').trim();
    const noteP = String(r[15] || '').trim();

    cautionStructMap.set(k, {
      condType: cond,
      ngClinics: ngClinicsArr,
      ngDows: ngDowsArr,
      ngTime: ngTime,
      memo: [noteK, noteP].filter(Boolean).join(' / ')
    });

    const parts = [];
    if (cond) parts.push('【' + cond + '】');
    if (ngClinicsArr.length > 0) parts.push('NG拠点:' + ngClinicsArr.join('・'));
    if (ngDowsArr.length > 0) parts.push('NG曜日:' + ngDowsArr.join('・'));
    if (ngTime) parts.push('不可時間:' + ngTime);
    if (noteK) parts.push(noteK);
    if (noteP) parts.push(noteP);
    if (parts.length > 0) docMemoLookup.set(k, parts.join(' / '));
  });

  // 2. 『交渉しやすい医師』
  const negotiableMap = new Map();
  getRows('交渉しやすい医師').forEach(function(r) {
    const rawA = String(r[0] || '').trim();
    const rawC = String(r[2] || '').trim();
    if (/^\d+$/.test(rawA) || rawC.includes('AI自動抽出')) return;
    const k = normKeyForDoc_(rawA);
    if (!k) return;
    const combined = [String(r[3] || '').trim(), r[4] ? ('(AIメモ:' + String(r[4]).trim() + ')') : ''].filter(Boolean).join(' ');
    negotiableMap.set(k, combined);
    if (combined) docMemoLookup.set(k, (docMemoLookup.get(k) ? docMemoLookup.get(k) + ' / ' : '') + combined);
  });

  // 3. 『SBU』・『常勤依頼手当』
  const sbuSet = new Set();
  getRows('SBU').forEach(function(r) {
    const k = normKeyForDoc_(r[0]);
    if (k) sbuSet.add(k);
  });

  const ftAllowDocSet = new Set();
  const ftAllowCondMap = new Map();
  getRows('常勤依頼手当').forEach(function(r) {
    const k = normKeyForDoc_(r[0]);
    if (!k) return;
    ftAllowDocSet.add(k);
    const cond = String(r[2] || '').trim();
    if (cond) {
      ftAllowCondMap.set(k, cond);
      docMemoLookup.set(k, (docMemoLookup.get(k) ? docMemoLookup.get(k) + ' / ' : '') + '[常勤手当条件:' + cond + ']');
    }
  });

  // 4. 『最新_連勤注意リスト』 ＋ 【設計1：シフトキャンセル履歴の日付統合】
  const streakNgMap = new Map();
  streakWarnings.forEach(function(w) {
    const k = normKeyForDoc_(w[1]);
    const ngStr = String(w[4] || '');
    if (k && ngStr && ngStr !== '-') {
      streakNgMap.set(k, (streakNgMap.get(k) ? streakNgMap.get(k) + ',' : '') + ngStr);
    }
  });

  const canceledDatesByDoc = (typeof loadShiftCancelDatesByDoctor1b2_ === 'function')
    ? loadShiftCancelDatesByDoctor1b2_(ss)
    : new Map();

  // キャンセル日も streakNgMap に合流させることで、既存のすべての NG 日チェックでも自動的に弾かれるようにする
  canceledDatesByDoc.forEach(function(dateSet, docKey) {
    const joinedDates = Array.from(dateSet).join(',');
    if (joinedDates) {
      streakNgMap.set(docKey, (streakNgMap.get(docKey) ? streakNgMap.get(docKey) + ',' : '') + joinedDates);
    }
  });

  // 5. 【設計2】医師返信メールからの日別・時間上限（例：11/8 柿本先生の「18時までなら可能」）マップ構築
  const dateMailConstraintsMap = (typeof buildDateSpecificMailConstraintsMap1b2_ === 'function')
    ? buildDateSpecificMailConstraintsMap1b2_(ss)
    : new Map();

  // 6. 『最新_当日シフト』（確定シフト）の日付×医師インデックス
  const confirmedDocScheduleByDate = new Map();
  activeShifts.forEach(function(r) {
    const d = String(r[0] || '').trim();
    const c = cleanClinicBaseName_(r[2]);
    if (!d || !c || c.includes('バックアップ')) return;
    if (!confirmedDocScheduleByDate.has(d)) confirmedDocScheduleByDate.set(d, new Map());
    const dayMap = confirmedDocScheduleByDate.get(d);

    const bands = [
      { code: 'A', label: '09:00-13:00', startMin: 540, endMin: 780, text: r[7] },
      { code: 'B', label: '15:00-18:00', startMin: 900, endMin: 1080, text: r[8] },
      { code: 'C', label: c.includes('北葛西') ? '18:00-20:00' : '18:00-21:00', startMin: 1080, endMin: c.includes('北葛西') ? 1200 : 1260, text: r[9] }
    ];

    bands.forEach(function(b) {
      splitDoctorTokensSafely_(b.text).forEach(function(tok) {
        const docKey = normKeyForDoc_(tok);
        if (!docKey || docKey === '未充足' || docKey === 'なし') return;
        if (!dayMap.has(docKey)) dayMap.set(docKey, []);
        dayMap.get(docKey).push({
          clinic: c,
          band: b.code,
          timeLabel: b.label,
          startMin: b.startMin,
          endMin: b.endMin
        });
      });
    });
  });

  return {
    docMemoLookup: docMemoLookup,
    absoluteNgDocSet: absoluteNgDocSet,
    cautionStructMap: cautionStructMap,
    negotiableMap: negotiableMap,
    sbuSet: sbuSet,
    ftAllowDocSet: ftAllowDocSet,
    ftAllowCondMap: ftAllowCondMap,
    streakNgMap: streakNgMap,
    canceledDatesByDoc: canceledDatesByDoc,
    dateMailConstraintsMap: dateMailConstraintsMap,
    confirmedDocScheduleByDate: confirmedDocScheduleByDate
  };
}