/**
 * ============================================================================
 * AI会議_4_2_ドキュメント提出_行解析.gs
 * （提出時の募集シフト「bosyu」長時間枠への自動合わせ込み・特別時給交渉・要延長セット・指定医師抽出版）
 * ============================================================================
 */

let SUBMIT_RECRUIT_CACHE_4_2_ = null;

function getSubmitRecruitContext4_2_() {
  if (SUBMIT_RECRUIT_CACHE_4_2_) return SUBMIT_RECRUIT_CACHE_4_2_;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getRows = function(name) {
    const sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() >= 2) ? sh.getDataRange().getDisplayValues().slice(1) : [];
  };

  const activeShifts = getRows('最新_当日シフト');
  const rawUnfilled = getRows('最新_不在枠');
  const spotCands = getRows('最新_個別打診候補');

  const recruitSlotsIndex = new Map();
  const addIdx = function(dRaw, rawClinic, rawTime, deptHint) {
    if (typeof normalizeDateYmd1b3_ !== 'function') return;
    const d = normalizeDateYmd1b3_(dRaw);
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!d || !cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = (String(rawClinic || '').includes('内科') || String(deptHint || '').includes('内科')) ? '内科' : '小児科';
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!recruitSlotsIndex.has(key)) recruitSlotsIndex.set(key, []);
    const list = recruitSlotsIndex.get(key);
    normalizeAndExpandAbsenceTime1b3_(rawTime, cleanC).forEach(function(t) {
      if (t && !list.includes(t)) list.push(t);
    });
  };

  rawUnfilled.forEach(function(r) { addIdx(r[0], r[1], r[2], ''); });
  spotCands.forEach(function(r) { addIdx(r[0], r[2], r[3], ''); });

  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []).forEach(function(m) {
          addIdx(m.d || m.date, m.c || m.clinic, m.t || m.time, m.dp || m.dept || '');
        });
      }
    } catch (e) {}
  }

  SUBMIT_RECRUIT_CACHE_4_2_ = {
    recruitSlotsIndex: recruitSlotsIndex,
    activeShifts: activeShifts
  };
  return SUBMIT_RECRUIT_CACHE_4_2_;
}

function mapCustomActionToApproachType_(actionTypeStr, fullText) {
  const s = String(actionTypeStr || '') + ' ' + String(fullText || '');
  if (s.includes('バックアップ')) return 'バックアップ勤務依頼';
  if (s.includes('移動') || s.includes('2診目崩し')) return '勤務地変更相談';
  if (s.includes('延長')) return '延長相談';
  if (s.includes('早出') || s.includes('前倒')) return '早出相談';
  return '個別アプローチ';
}

function detectBandFromTimeStr4_(timeStr) {
  const m = String(timeStr || '').match(/(\d{1,2}):(\d{2})/);
  if (!m) return 'A';
  const sh = Number(m[1]);
  if (sh < 13) return 'A';
  if (sh < 18) return 'B';
  return 'C';
}

function parseEditedDocTableRow_(dateStr, col1Text, col2Text, decisionText, assigneeText, resolver, customDocs, customComment, isFullTimeChecker, backupDocsByDate, backupUsedIdxByDateBand) {
  let dec = String(decisionText || '').trim();
  let commentText = String(customComment || '').trim();

  const inlineMemoMatch = dec.match(/^(.*?)\s*[（\(]メモ[：:](.+?)[）\)]$/);
  if (inlineMemoMatch) {
    dec = inlineMemoMatch[1].trim();
    if (!commentText) commentText = inlineMemoMatch[2].trim();
  }

  let spWageAddonLabel = '';
  const negSpMatch = dec.match(/【(特別時給(?:さらに)?(?:\+\d,\d{3}円|\+\d{4}円)(?:加算)?交渉)】/);
  if (negSpMatch) {
    spWageAddonLabel = negSpMatch[1];
    dec = dec.replace(/【特別時給(?:さらに)?(?:\+\d,\d{3}円|\+\d{4}円)(?:加算)?交渉】/g, '').trim();
    commentText = commentText ? (commentText + ' ／ 💰' + spWageAddonLabel) : ('💰' + spWageAddonLabel);
  }

  if (!dec || /^(ステイ|特別時給.*ステイ|見送り|なし|不要|ー|\-|保留|要検討|併行募集)$/.test(dec)) return [];

  const firstLine = col1Text.split('\n')[0].trim();
  const clinicMatch = firstLine.match(/【([^（\(]+)(?:[（\(](小児科|内科)[）\)])?】/);
  const clinicName = clinicMatch ? clinicMatch[1].trim() : firstLine.replace(/[【】]/g, '').replace(/[（\(].*?[）\)]/g, '').trim();
  const dept = (clinicMatch && clinicMatch[2]) ? clinicMatch[2].trim() : (firstLine.includes('内科') ? '内科' : '小児科');
  const timeMatch = col1Text.match(/(\d{1,2}:\d{2}-\d{1,2}:\d{2})/);
  const rawTargetTime = timeMatch ? timeMatch[1] : '';

  // ★ドキュメント上の表記が「15:00-16:00」等の短い不在時間になっていても、
  //   募集シフト（bosyu）の長い募集時間（例：「15:00-18:00」や「15:00-17:00」）へ自動で合わせてから提出する！
  let targetTime = rawTargetTime;
  if (rawTargetTime && typeof expandToLongestRecruitSlotTime1b3_ === 'function') {
    const ctx = getSubmitRecruitContext4_2_();
    const rawClinicWithDept = dept === '内科' ? (clinicName + '（内科）') : clinicName;
    targetTime = expandToLongestRecruitSlotTime1b3_(dateStr, rawClinicWithDept, rawTargetTime, ctx.recruitSlotsIndex, ctx.activeShifts);
  }

  const isSpecialWage = col1Text.includes('特別時給') || dec.includes('特別時給') || Boolean(spWageAddonLabel);

  const rawAsg = String(assigneeText || '').trim();
  const cleanAssignee = (!rawAsg || rawAsg === '担当者未定' || rawAsg === 'あとで決定' || rawAsg === '(未定)' || rawAsg === '未定')
    ? '未定'
    : (resolver.resolveStaff(rawAsg) || '未定');

  const memoSuffix = commentText ? ('（メモ: ' + commentText + '）') : '';

  if (dec.indexOf('バックアップ医師に勤務依頼') === 0) {
    const band = detectBandFromTimeStr4_(targetTime);
    const dayObj = (backupDocsByDate && backupDocsByDate[dateStr]) ? backupDocsByDate[dateStr] : { A: [], B: [], C: [] };
    const bandBackups = dayObj[band] || [];
    const bandKey = dateStr + '_' + band;
    const usedIdx = (backupUsedIdxByDateBand && backupUsedIdxByDateBand[bandKey]) ? backupUsedIdxByDateBand[bandKey] : 0;
    const pickedBackup = bandBackups[usedIdx] || bandBackups[0] || { name: 'バックアップ医師', licenseNo: '' };
    if (backupUsedIdxByDateBand) backupUsedIdxByDateBand[bandKey] = usedIdx + 1;

    const resolvedBk = resolver.resolveDoctor(pickedBackup.name, dateStr, 'バックアップ');
    const bkName = resolvedBk.name || pickedBackup.name;
    const bkLic = resolvedBk.licenseNo || pickedBackup.licenseNo || '';
    const isFt = isFullTimeChecker ? isFullTimeChecker(bkName, bkLic, 'バックアップ') : false;

    return [{
      date: dateStr,
      clinic_name: clinicName,
      raw_clinic_name: clinicName,
      department: dept,
      target_time: targetTime,
      time_category: '',
      approach_type: 'バックアップ勤務依頼',
      target_doctor: bkName,
      license_no: bkLic,
      is_full_time: isFt,
      from_clinic: '関東バックアップ',
      is_special_wage: isSpecialWage,
      special_wage_addon: spWageAddonLabel,
      assignee: cleanAssignee,
      memo: commentText,
      comment: commentText,
      evidence: '[バックアップ医師に勤務依頼] ' + bkName + '先生' + memoSuffix
    }];
  }

  const strippedDec = dec.replace(/\((?:①|②|医師未選択)[^)]*\)/g, '').trim();
  const isPresetDec = /^(案[123１２３]|全案)/.test(strippedDec);
  if (!isPresetDec && Array.isArray(customDocs) && customDocs.length > 0) {
    return customDocs.map(function(cd) {
      const appType = mapCustomActionToApproachType_(cd.actionType, dec);
      const resolved = resolver.resolveDoctor(cd.name, dateStr, clinicName);
      const finalName = cd.name || resolved.name;
      const finalLic = cd.licenseNo || resolved.licenseNo || '';
      const isFt = isFullTimeChecker ? isFullTimeChecker(finalName, finalLic, dec) : false;
      const fromClinic = (appType === '勤務地変更相談' && resolved.dayClinic && resolved.dayClinic !== clinicName)
        ? resolved.dayClinic
        : '';
      return {
        date: dateStr,
        clinic_name: clinicName,
        raw_clinic_name: clinicName,
        department: dept,
        target_time: targetTime,
        time_category: '',
        approach_type: appType,
        target_doctor: finalName,
        license_no: finalLic,
        is_full_time: isFt,
        from_clinic: fromClinic,
        is_special_wage: isSpecialWage,
        special_wage_addon: spWageAddonLabel,
        assignee: cleanAssignee,
        memo: commentText,
        comment: commentText,
        evidence: '[編集くんカスタム確定] ' + dec + memoSuffix
      };
    });
  }

  const planMap = {};
  col2Text.split('\n').forEach(function(line) {
    const m = line.match(/^【案([123１２３])】(.*)$/);
    if (m) {
      const num = m[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
      planMap[num] = m[2].trim();
    }
  });

  const chosenActionStrings = [];
  if (/全案|すべて|全部/.test(strippedDec)) {
    ['1', '2', '3'].forEach(function(k) {
      if (planMap[k]) {
        const subMatch = dec.match(new RegExp('案' + k + '\\(([^)]+)\\)'));
        const rawSpec = subMatch ? subMatch[1].trim() : '';
        const specifiedDoc = (rawSpec && !rawSpec.includes('両方') && !rawSpec.includes('未選択'))
          ? rawSpec.replace(/^[①②12]\s*/, '').replace(/先生$/g, '').trim()
          : '';
        chosenActionStrings.push({
          label: rawSpec ? ('案' + k + '(' + rawSpec + ')') : ('案' + k),
          text: planMap[k],
          specifiedDoc: specifiedDoc
        });
      }
    });
  } else {
    const tokens = dec.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(function(t) { return t.trim(); }).filter(Boolean);
    tokens.forEach(function(tok) {
      const specMatch = tok.match(/^(?:【?案)?([123１２３])(?:】)?(?:\(([^)]+)\))?$/);
      if (specMatch) {
        const num = specMatch[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
        const rawSpec = specMatch[2] ? specMatch[2].trim() : '';
        const specifiedDoc = (rawSpec && !rawSpec.includes('両方') && !rawSpec.includes('未選択'))
          ? rawSpec.replace(/^[①②12]\s*/, '').replace(/先生$/g, '').trim()
          : '';
        if (planMap[num]) {
          chosenActionStrings.push({
            label: rawSpec ? ('案' + num + '(' + rawSpec + ')') : ('案' + num),
            text: planMap[num],
            specifiedDoc: specifiedDoc
          });
        }
      } else {
        const multiNums = tok.match(/案([123１２３])/g);
        if (multiNums && tok.replace(/案[123１２３]|[\s、,・と]/g, '') === '') {
          multiNums.forEach(function(mn) {
            const num = mn.replace('案', '').replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
            if (planMap[num]) chosenActionStrings.push({ label: '案' + num, text: planMap[num], specifiedDoc: '' });
          });
        } else {
          chosenActionStrings.push({ label: 'カスタム指示', text: tok, specifiedDoc: '' });
        }
      }
    });
  }

  const results = [];
  chosenActionStrings.forEach(function(actionObj) {
    let autoCondMemo = commentText;
    const condMatch = actionObj.text.match(/※(残る[^\n、，）\)]+?確認後)/);
    if (condMatch) {
      const condNote = '【要セット確認】' + condMatch[1];
      autoCondMemo = autoCondMemo ? (autoCondMemo + ' ／ ' + condNote) : condNote;
    }

    let parsedList = parseSingleActionTextToItems_(actionObj.text, dateStr, clinicName, resolver);
    if (actionObj.specifiedDoc && parsedList.length > 1) {
      const filtered = parsedList.filter(function(p) {
        return p.doctorName.includes(actionObj.specifiedDoc) || actionObj.specifiedDoc.includes(p.doctorName);
      });
      if (filtered.length > 0) parsedList = filtered;
    }

    parsedList.forEach(function(p) {
      const isFt = isFullTimeChecker ? isFullTimeChecker(p.doctorName, p.licenseNo, actionObj.text) : false;
      results.push({
        date: dateStr,
        clinic_name: clinicName,
        raw_clinic_name: clinicName,
        department: dept,
        target_time: targetTime,
        time_category: '',
        approach_type: p.approachType,
        target_doctor: p.doctorName,
        license_no: p.licenseNo || '',
        is_full_time: isFt,
        from_clinic: p.fromClinic,
        is_special_wage: isSpecialWage,
        special_wage_addon: spWageAddonLabel,
        assignee: cleanAssignee,
        memo: autoCondMemo,
        comment: autoCondMemo,
        evidence: '[ドキュメント確定:' + actionObj.label + '] ' + actionObj.text + (autoCondMemo ? ('（' + autoCondMemo + '）') : '')
      });
    });
  });

  return results;
}

function parseBundleTableRow_(defaultYear, docColText, slotsColText, decColText, resolver, isFullTimeChecker) {
  const dec = String(decColText || '').trim();
  if (!dec || /^(ステイ|特別時給.*ステイ|見送り|なし|不要|ー|\-|保留|不採用)$/.test(dec)) return [];

  const rawDoc = String(docColText || '').replace(/[\(（].*?[\)）]/g, '').replace(/先生/g, '').trim();
  const resolvedDoc = resolver.resolveDoctor(rawDoc, '', '');
  const isFt = isFullTimeChecker ? isFullTimeChecker(resolvedDoc.name, resolvedDoc.licenseNo, docColText) : false;
  const assigneeFromDec = dec.replace(/採用|OK|◯|○|[\/／・,\s ]/g, '').trim();
  const cleanAssignee = (!assigneeFromDec || assigneeFromDec === '担当者未定' || assigneeFromDec === 'あとで決定')
    ? '未定'
    : (resolver.resolveStaff(assigneeFromDec) || '未定');

  const items = [];
  let currentClinic = '';
  let currentDept = '小児科';

  String(slotsColText || '').split('\n').map(function(l) { return l.trim(); }).filter(Boolean).forEach(function(line) {
    if (line.startsWith('■')) {
      const rawC = line.replace(/^■\s*/, '').trim();
      currentDept = rawC.includes('内科') ? '内科' : '小児科';
      currentClinic = rawC.replace(/[\(（].*?[\)）]/g, '').trim();
      return;
    }
    const dm = line.match(/(\d{1,2})\/(\d{1,2})[^\d]*(\d{1,2}:\d{2}-\d{1,2}:\d{2}(?:\s*\/\s*\d{1,2}:\d{2}-\d{1,2}:\d{2})*)/);
    if (dm && currentClinic) {
      const dateStr = defaultYear + '/' + ('0' + dm[1]).slice(-2) + '/' + ('0' + dm[2]).slice(-2);
      dm[3].split('/').map(function(t) { return t.trim(); }).filter(Boolean).forEach(function(tRange) {
        items.push({
          date: dateStr,
          clinic_name: currentClinic,
          raw_clinic_name: currentClinic,
          department: currentDept,
          target_time: tRange,
          time_category: '',
          approach_type: '個別アプローチ',
          target_doctor: resolvedDoc.name,
          license_no: resolvedDoc.licenseNo || '',
          is_full_time: isFt,
          from_clinic: '',
          is_special_wage: false,
          special_wage_addon: '',
          assignee: cleanAssignee,
          memo: '',
          comment: '',
          evidence: '[ドキュメント確定:おまとめ打診] ' + resolvedDoc.name + '先生 (' + currentClinic + ' ' + dateStr + ' ' + tRange + ')'
        });
      });
    }
  });

  return items;
}

function parseSingleActionTextToItems_(text, dateStr, targetClinic, resolver) {
  const s = String(text || '').replace(/※残る.*?に、/g, '').replace(/\[※.*?\]/g, '').trim();
  if (!s) return [];

  if (s.includes('1診・併行募集') || s.includes('併行募集') || s.includes('埋まり次第')) {
    return [];
  }

  let approachType = '個別アプローチ';
  let fromClinic = '';
  let rawDoctors = [];

  if (s.includes('2診目崩し') || s.includes('移動')) {
    approachType = '勤務地変更相談';
    const fromMatch = s.match(/】([^\(（\s]+?)[（\(](?:※|[12]診)/) || s.match(/([^\s【】]+?)(?:より|から)/);
    if (fromMatch) fromClinic = fromMatch[1].trim();
    const insideParen = s.match(/[12]診(?:目)?[：:]([^）\)]+)[）\)]/);
    rawDoctors = insideParen ? extractDoctorNamesFromSegment_(insideParen[1]) : extractDoctorNamesFromSegment_(s.replace(/^.*?から/, ''));
  } else if (s.includes('延長')) {
    approachType = '延長相談';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  } else if (s.includes('早出') || s.includes('前倒')) {
    approachType = '早出相談';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  } else {
    approachType = '個別アプローチ';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  }

  if (rawDoctors.length === 0) return [{ approachType: approachType, doctorName: '', licenseNo: '', fromClinic: fromClinic }];

  return rawDoctors.map(function(rawDoc) {
    const resolved = resolver.resolveDoctor(rawDoc, dateStr, fromClinic || targetClinic);
    const finalFromClinic = (approachType === '勤務地変更相談' && !fromClinic && resolved.dayClinic && resolved.dayClinic !== targetClinic)
      ? resolved.dayClinic
      : fromClinic;
    return { approachType: approachType, doctorName: resolved.name, licenseNo: resolved.licenseNo, fromClinic: finalFromClinic };
  });
}

function extractDoctorNamesFromSegment_(seg) {
  const cleaned = String(seg || '')
    .replace(/\[.*?\]|［.*?］/g, '')
    .replace(/[（\(].*?[）\)]/g, '')
    .replace(/前医の|後医の|どちらか1名.*$|へ.*$|に.*$|を.*$/g, '')
    .trim();

  return cleaned.split(/[\/／]|または|・|,|、/)
    .map(function(p) { return p.replace(/[①②③④123１２３:\s ]|先生/g, '').trim(); })
    .filter(function(p) { return p.length >= 2 && !/^(スポット|紹介会社|常勤|非常勤|移動|延長|早出|個別)$/.test(p); });
}