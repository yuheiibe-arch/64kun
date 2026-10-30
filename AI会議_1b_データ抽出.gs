/**
 * ============================================================================
 * 1b_AI会議_データ抽出.gs
 * （不在枠・当日シフト・常勤依頼手当・『交渉注意医師』A〜P列の正確な抽出ロジック）
 * ============================================================================
 */

function normKeyForDoc_(name) {
  return String(name || '')
    .replace(/[\(（\[].*?[\)）\]]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function inferAreaForDoc_(clinicName) {
  const c = String(clinicName || '');
  if (/関西|大阪|兵庫|京都|奈良|滋賀|和歌山|鶴見緑地|西宮|千里|高槻|あべの|天美|東岸和田|今里|門真|守口|枚方|吹田|豊中|茨木|八尾|松原|堺|尼崎|伊丹|宝塚|川西|明石|神戸|姫路|生駒|大和|橿原|草津|大津/.test(c)) {
    return '関西';
  }
  if (/海老名|高田|武蔵小杉|相模大野|相模原|新百合ヶ丘|新百合ケ丘|茅ヶ崎|茅ケ崎|天王町|小田栄|東戸塚|宮前平|川崎|横浜|藤沢|平塚|厚木|大和|橋本|溝の口|日吉|綱島|たまプラーザ|センター北|センター南|戸塚|上大岡/.test(c)) {
    return '神奈川';
  }
  if (/越谷|川口|川越|志木|所沢|南浦和|与野|草加松原|草加|八潮|浦和|大宮|さいたま|戸田|蕨|和光|朝霞|新座|ふじみ野|富士見|春日部|三郷|吉川|久喜|上尾|桶川|熊谷/.test(c)) {
    return '埼玉';
  }
  if (/二和向台|柏の葉|柏|村上|稲毛海岸|稲毛|流山おおたかの森|流山|新鎌ケ谷|新鎌ヶ谷|幕張ベイパーク|幕張|八千代緑が丘|八千代緑ヶ丘|八千代|妙典|市川妙典|千葉ニュータウン|千葉NT|新浦安|浦安|市川|船橋|千葉|松戸|津田沼|習志野|我孫子|印西|成田|木更津/.test(c)) {
    return '千葉';
  }
  if (/つくば|守谷|取手|土浦|牛久|水戸|茨城/.test(c)) {
    return '茨城';
  }
  return '東京';
}

function isAreaCompatible_(area1, area2) {
  if (area1 === '関西' || area2 === '関西') return area1 === area2;
  if (area1 === area2) return true;
  if (area1 === '東京' && (area2 === '千葉' || area2 === '埼玉' || area2 === '神奈川')) return true;
  if (area2 === '東京' && (area1 === '千葉' || area1 === '埼玉' || area1 === '神奈川')) return true;
  if ((area1 === '千葉' && area2 === '茨城') || (area1 === '茨城' && area2 === '千葉')) return true;
  return false;
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
      const cleanTarget = String(targetClinic || '').replace(/[\(（].*?[\)）]/g, '').trim();
      if (!cleanTarget.includes(allowedKeyword) && !targetArea.includes(allowedKeyword)) {
        return false;
      }
    }
  }
  return true;
}

function collectCompactContextForAi_(ss) {
  const today = new Date();
  const todayYmd = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  const twoWeeksYmd = Utilities.formatDate(new Date(today.getTime() + 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const oneMonthYmd = Utilities.formatDate(new Date(today.getTime() + 31 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  const getRows = (name) => {
    const sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() >= 2) ? sh.getDataRange().getDisplayValues().slice(1) : [];
  };

  const docMemoLookup = new Map();
  const absoluteNgDocSet = new Set();

  // 1. 『交渉注意医師』シート（A列:医師名, D列:交渉条件, F〜J列:NG拠点, K列:注意事項, L〜N列:NG曜日, O列:不可時間, P列:詳細メモ）
  getRows('交渉注意医師').forEach(r => {
    const k = normKeyForDoc_(r[0]);
    if (!k) return;
    const cond = String(r[3] || '').trim();
    if (cond === '絶対不可') absoluteNgDocSet.add(k);

    const ngClinics = [r[5], r[6], r[7], r[8], r[9]].map(v => String(v || '').trim()).filter(Boolean).join('・');
    const ngDows = [r[11], r[12], r[13]].map(v => String(v || '').trim()).filter(Boolean).join('・');
    const ngTime = String(r[14] || '').trim();
    const noteK = String(r[10] || '').trim();
    const noteP = String(r[15] || '').trim();

    const parts = [];
    if (cond) parts.push(`【${cond}】`);
    if (ngClinics) parts.push(`NG拠点:${ngClinics}`);
    if (ngDows) parts.push(`NG曜日:${ngDows}`);
    if (ngTime) parts.push(`不可時間:${ngTime}`);
    if (noteK) parts.push(noteK);
    if (noteP) parts.push(noteP);

    if (parts.length > 0) {
      docMemoLookup.set(k, parts.join(' / '));
    }
  });

  // 2. 『交渉しやすい医師』シート（A列:医師名, D列:人間メモ, E列:AIメモ）
  getRows('交渉しやすい医師').forEach(r => {
    const k = normKeyForDoc_(r[0]) || normKeyForDoc_(r[1]);
    if (!k) return;
    const humanMemo = String(r[3] || '').trim();
    const aiMemo = String(r[4] || '').trim();
    const combined = [humanMemo, aiMemo ? `(AIメモ:${aiMemo})` : ''].filter(Boolean).join(' ');
    if (combined) {
      docMemoLookup.set(k, (docMemoLookup.get(k) ? docMemoLookup.get(k) + ' / ' : '') + combined);
    }
  });

  // 3. 『常勤依頼手当』シート
  const ftAllowDocSet = new Set();
  const ftAllowCondMap = new Map();
  getRows('常勤依頼手当').forEach(r => {
    const k = normKeyForDoc_(r[0]);
    if (!k) return;
    ftAllowDocSet.add(k);
    const cond = String(r[2] || '').trim();
    if (cond) {
      ftAllowCondMap.set(k, cond);
      docMemoLookup.set(k, (docMemoLookup.get(k) ? docMemoLookup.get(k) + ' / ' : '') + `[常勤手当条件:${cond}]`);
    }
  });

  const sanitizeAdjDocsForAi = (docsStr, statusStr, actionType, targetClinic, targetArea) => {
    if (!docsStr || docsStr === 'なし') return { docs: 'なし', status: '' };
    const tokens = String(docsStr).split(/[,、，\/]+/).map(t => t.trim()).filter(Boolean);
    const sanitizedTokens = [];
    let hasMovableFt = false;
    let hasPartTime = false;

    tokens.forEach(tok => {
      const k = normKeyForDoc_(tok);
      if (absoluteNgDocSet.has(k)) return;
      if (tok.includes('常勤')) {
        const condStr = ftAllowCondMap.get(k) || '';
        if (ftAllowDocSet.has(k) && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
          const condLabel = condStr ? `・条件:${condStr}` : '';
          sanitizedTokens.push(`${k}(常勤依頼手当対象${condLabel})`);
          hasMovableFt = true;
        } else {
          sanitizedTokens.push(`${k}(通常常勤・${actionType}不可)`);
        }
      } else {
        sanitizedTokens.push(tok);
        hasPartTime = true;
      }
    });

    let cleanStatus = String(statusStr || '');
    if (!hasMovableFt && !hasPartTime) {
      cleanStatus = `通常常勤のため${actionType}不可（提案およびauditNoteへの記載不要）`;
    }
    return {
      docs: sanitizedTokens.join(' / '),
      status: cleanStatus
    };
  };

  const unfilledAll = getRows('最新_不在枠').filter(r => r[0] && r[0] >= todayYmd);
  const within2W = unfilledAll.filter(r => r[0] <= twoWeeksYmd);
  const targetEndYmd = (within2W.length < 10 && unfilledAll.length > within2W.length) ? oneMonthYmd : twoWeeksYmd;
  const targetUnfilled = unfilledAll.filter(r => r[0] <= targetEndYmd);

  const periodStart = targetUnfilled.length > 0 ? targetUnfilled[0][0] : todayYmd;
  const periodEnd = targetUnfilled.length > 0 ? targetUnfilled[targetUnfilled.length - 1][0] : targetEndYmd;

  const activeShifts = getRows('最新_当日シフト');
  const streakWarnings = getRows('最新_連勤注意リスト');
  const spotCandidates = getRows('最新_個別打診候補');
  const bundledCandidates = getRows('最新_おまとめ打診候補');
  const kbRows = getRows('ナレッジシート');

  const streakNgMap = new Map();
  streakWarnings.forEach(w => {
    const k = normKeyForDoc_(w[1]);
    const ngStr = String(w[4] || '');
    if (k && ngStr && ngStr !== '-') {
      streakNgMap.set(k, (streakNgMap.get(k) ? streakNgMap.get(k) + ',' : '') + ngStr);
    }
  });

  const spotMap = new Map();
  spotCandidates.forEach(r => {
    const key = `${r[0]}_${r[2]}_${r[3]}`;
    if (!spotMap.has(key)) spotMap.set(key, []);
    const dKey = normKeyForDoc_(r[5]);
    if (absoluteNgDocSet.has(dKey)) return;
    const extraMemo = docMemoLookup.get(dKey) || '';
    spotMap.get(key).push({
      rank: r[4],
      doctor: r[5],
      sameDowSlotCount: r[8],
      totalCount: r[10],
      daySchedule: r[12],
      masterNote: [r[13] || '', extraMemo].filter(Boolean).join(' ')
    });
  });

  const filterMovableDoctorsFromText = (docsStr, dateStr, targetClinic, targetArea) => {
    if (!docsStr || docsStr === '未充足') return [];
    const tokens = String(docsStr).split(/[,、，\n]+/).map(t => t.trim()).filter(Boolean);
    const movable = [];
    tokens.forEach(tok => {
      const k = normKeyForDoc_(tok);
      if (!k || absoluteNgDocSet.has(k)) return;
      if (streakNgMap.has(k) && streakNgMap.get(k).includes(dateStr)) return;
      const memoInfo = docMemoLookup.get(k) ? `{${docMemoLookup.get(k)}}` : '';

      if (!isFtConditionCompatibleWithTarget_(memoInfo, targetClinic, targetArea)) return;

      if (tok.includes('常勤')) {
        const condStr = ftAllowCondMap.get(k) || '';
        if (ftAllowDocSet.has(k) && tok.includes('所定休出') && isFtConditionCompatibleWithTarget_(condStr, targetClinic, targetArea)) {
          movable.push(`${k}(常勤依頼手当対象・所定休出)${memoInfo}`);
        }
        return;
      }
      movable.push(`${tok}${memoInfo}`);
    });
    return movable;
  };

  const targetDatesSet = new Set();
  const slotsContext = targetUnfilled.map((u, idx) => {
    const dateStr = u[0];
    const clinic = u[1];
    const absTime = u[2];
    const isInternalMed = clinic.includes('内科');
    const area = inferAreaForDoc_(clinic);
    targetDatesSet.add(dateStr);

    const slotCode = (absTime.startsWith('09') || absTime.startsWith('9')) ? 'A'
      : ((absTime.startsWith('18') || absTime.startsWith('17')) ? 'C' : 'B');
    const countCol = slotCode === 'A' ? 4 : (slotCode === 'B' ? 5 : 6);
    const docsCol = slotCode === 'A' ? 7 : (slotCode === 'B' ? 8 : 9);

    const dayShifts = activeShifts.filter(s => s[0] === dateStr);
    const twoDocClinics = [];
    const oneDocParallelClinics = [];
    let backupDocs = '';

    if (!isInternalMed) {
      dayShifts.forEach(s => {
        const sClinic = s[2];
        const sDocs = String(s[docsCol] || '');
        if (!sDocs || sDocs === '未充足') return;

        if (sClinic.includes('バックアップ')) {
          backupDocs = sDocs;
          return;
        }
        if (sClinic === clinic || s[3].includes('内科')) return;

        const sArea = inferAreaForDoc_(sClinic);
        if (!isAreaCompatible_(area, sArea)) return;

        const docCount = Number(s[countCol] || 0);
        const movableDocs = filterMovableDoctorsFromText(sDocs, dateStr, clinic, area);
        if (movableDocs.length === 0) return;

        if (docCount >= 2) {
          twoDocClinics.push({
            clinic: sClinic,
            area: sArea,
            movableDoctors: movableDocs
          });
        } else if (docCount === 1) {
          const specDocs = movableDocs.filter(d => d.includes('★SBU') || d.includes('★交渉しやすい') || d.includes('★移動相談可'));
          if (specDocs.length > 0) {
            oneDocParallelClinics.push({
              clinic: sClinic,
              area: sArea,
              movableDoctors: specDocs
            });
          }
        }
      });
    }

    const cleanK = clinic.includes('北葛西') ? (isInternalMed ? '北葛西（内科）' : '北葛西（小児科）')
      : (clinic.includes('亀有') ? (isInternalMed ? '亀有（内科）' : '亀有（小児科）') : clinic.replace(/小児科|内科|[\(（].*?[\)）]/g, '').trim());
    const kinmukunTop = spotMap.get(`${dateStr}_${cleanK}_${absTime}`) || [];

    const prevChecked = sanitizeAdjDocsForAi(u[4] || 'なし', u[7] || '', '延長', clinic, area);
    const nextChecked = sanitizeAdjDocsForAi(u[5] || 'なし', u[8] || '', '早出', clinic, area);

    return {
      slotId: idx + 1,
      date: dateStr,
      area: area,
      clinic: clinic,
      dept: isInternalMed ? '内科' : '小児科',
      absenceTime: absTime,
      listedWage: u[3] || '13,000円',
      specialWage: u[11] || '',
      assignee: String(u[9] || '').trim(), // ★『最新_不在枠』J列の担当を引き継ぐ
      prevDocs: prevChecked.docs,
      prevExtStatus: prevChecked.status,
      nextDocs: nextChecked.docs,
      nextEarlyStatus: nextChecked.status,
      currentMemo: u[12] || '',
      backupDocs: backupDocs,
      twoDocBreakdownCandidates: twoDocClinics.slice(0, 4),
      oneDocParallelCandidates: oneDocParallelClinics.slice(0, 4),
      kinmukunCandidates: kinmukunTop.slice(0, 4)
    };
  });

  const targetBundles = bundledCandidates.filter(r => {
    const dKey = normKeyForDoc_(r[0]);
    if (absoluteNgDocSet.has(dKey)) return false;
    const text = String(r[4] || '');
    return Array.from(targetDatesSet).some(d => text.includes(d));
  }).slice(0, 12).map(r => {
    const dKey = normKeyForDoc_(r[0]);
    const extraMemo = docMemoLookup.get(dKey) || '';
    return {
      doctor: r[0],
      licenseNo: r[1],
      slotCount: r[3],
      targetSlotsText: r[4],
      masterNote: [r[5] || '', extraMemo].filter(Boolean).join(' ')
    };
  });

  const streakList = streakWarnings.filter(w => {
    const docLabel = String(w[1] || '');
    const periodText = String(w[3] || '');
    const ngDateText = String(w[4] || '');
    if (!ngDateText || ngDateText === '-') return false;
    if (periodText.includes('中1日')) return false;
    if (docLabel.includes('(常勤)') && !docLabel.includes('手当対象')) return false;
    return Array.from(targetDatesSet).some(d => ngDateText.includes(d) || periodText.includes(d));
  }).slice(0, 10).map(w => ({
    licenseNo: w[0],
    doctor: w[1],
    streakDays: w[2],
    period: w[3],
    ngDates: w[4],
    note: w[5]
  }));

  const recentFlaggedKb = kbRows.filter(r => {
    const flag = String(r[7] || '');
    const body = String(r[2] || '');
    if (!flag.includes('★要確認')) return false;
    if (flag.includes('AIメモ反映済') || flag.includes('自動判定済')) return false;
    if (body.includes('退職することになりました')) return false;
    return true;
  }).slice(-5).map(r => ({
    doctor: r[1],
    replySummary: String(r[2] || '').replace(/\n+/g, ' ').slice(0, 80),
    date: r[6],
    flag: r[7]
  }));

  return {
    periodStart,
    periodEnd,
    unfilledCount: targetUnfilled.length,
    slotsContext,
    targetBundles,
    streakList,
    recentFlaggedKb
  };
}