/**
 * ============================================================================
 * AI会議_4_2b_提出_メイン行解析.gs（3分割の2）
 * ・不在枠テーブル1行分の提出データ生成（parseEditedDocTableRow_）
 * ・承諾/充足枠スキップ・辞退医除外・バックアップ医割当・案1〜3＆カスタム結合対応
 * ============================================================================
 */

function buildBackupSubmitItem4_2_(params) {
  const {
    dateStr, clinicName, dept, targetTime, isSpecialWage, spWageAddonLabel,
    cleanAssignee, commentText, memoSuffix, extraNote,
    resolver, isFullTimeChecker, backupDocsByDate, backupUsedIdxByDateBand
  } = params;

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

  return {
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
    evidence: '[バックアップ医師に勤務依頼] ' + bkName + '先生' + (extraNote ? ' ' + extraNote : '') + memoSuffix
  };
}

function resolveChosenActionStrings4_2_(dec, strippedDec, planMap) {
  const chosenActionStrings = [];

  if (/^(?:全案|すべて|全部)$/.test(strippedDec)) {
    ['1', '2', '3'].forEach(function(k) {
      if (!planMap[k]) return;
      const subMatch = dec.match(new RegExp('案\\s*' + k + '\\s*[（\\(]([^）\\)]+)[）\\)]'));
      const rawSpec = subMatch ? subMatch[1].trim() : '';
      const specifiedDoc = (rawSpec && !rawSpec.includes('両方') && !rawSpec.includes('未選択'))
        ? rawSpec.replace(/^[①②③123]\s*/, '').replace(/先生$/g, '').trim()
        : '';
      chosenActionStrings.push({
        label: rawSpec ? ('案' + k + '(' + rawSpec + ')') : ('案' + k),
        text: planMap[k],
        specifiedDoc: specifiedDoc
      });
    });
    return chosenActionStrings;
  }

  const tokens = splitTopLevelDecisionTokens4_2_(dec);
  tokens.forEach(function(tok) {
    if (isStayOrSkipDecision4_2_(tok)) return;

    // 1) 「案1」「案2(① / ②両方)」「案2（②灘理絵先生）」の厳密一致
    const specMatch = tok.match(/^(?:【?案)?\s*([123１２３])(?:】)?(?:\s*[（\(]([^）\)]+)[）\)])?$/);
    if (specMatch) {
      const num = specMatch[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
      const rawSpec = specMatch[2] ? specMatch[2].trim() : '';
      const specifiedDoc = (rawSpec && !rawSpec.includes('両方') && !rawSpec.includes('未選択'))
        ? rawSpec.replace(/^[①②③123]\s*/, '').replace(/先生$/g, '').trim()
        : '';
      if (planMap[num]) {
        chosenActionStrings.push({
          label: rawSpec ? ('案' + num + '(' + rawSpec + ')') : ('案' + num),
          text: planMap[num],
          specifiedDoc: specifiedDoc
        });
      }
      return;
    }

    // 2) 「案1案2」のように連続表記されたケース
    const multiNums = tok.match(/案\s*([123１２３])/g);
    if (multiNums && tok.replace(/案\s*[123１２３]|[\s、,・と]/g, '') === '') {
      multiNums.forEach(function(mn) {
        const num = mn.replace(/案\s*/, '').replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
        if (planMap[num]) chosenActionStrings.push({ label: '案' + num, text: planMap[num], specifiedDoc: '' });
      });
      return;
    }

    // 3) 「案3 花輪和先生へ相談」「案 2向瀬先生へ延長打診」のように案番号＋テキストが書かれたケース
    const planWithCustom = tok.match(/^(?:【?案\s*([123１２３])】?)\s*(.+)$/);
    if (planWithCustom) {
      const pNum = planWithCustom[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
      chosenActionStrings.push({
        label: '案' + pNum + '・カスタム',
        text: planWithCustom[2].trim(),
        planFallbackText: planMap[pNum] || '',
        specifiedDoc: ''
      });
      return;
    }

    // 4) 純粋なカスタム手入力
    chosenActionStrings.push({ label: 'カスタム指示', text: tok, specifiedDoc: '' });
  });

  return chosenActionStrings;
}

function parseEditedDocTableRow_(dateStr, col1Text, col2Text, decisionText, assigneeText, resolver, customDocs, customComment, isFullTimeChecker, backupDocsByDate, backupUsedIdxByDateBand) {
  const normCol1 = String(col1Text || '').replace(/\r\n|\r/g, '\n').trim();
  const normCol2 = String(col2Text || '')
    .replace(/\r\n|\r/g, '\n')
    .replace(/([^\n])\s*(?=【案[123１２３]】)/g, '$1\n')
    .replace(/([^\n])\s*(?=💡|📝|・【進行中メモ】|・⚠️)/g, '$1\n')
    .trim();

  const separated = extractDecisionHeaderAndArrows4_2_(decisionText);
  let dec = separated.header;
  const arrowLines = separated.arrowLines;
  let commentText = String(customComment || '').replace(/\r\n|\r/g, ' ').trim();

  // 既に「⭕承諾」または「⭕応募あり充足」となっている枠は再打診不要のためスキップ
  const isSlotAlreadyAccepted = arrowLines.some(function(l) {
    return l.includes('⭕') && (l.includes('承諾') || l.includes('充足'));
  });
  if (isSlotAlreadyAccepted && (!Array.isArray(customDocs) || customDocs.length === 0)) {
    return [];
  }

  // この枠で「❌辞退」となっている医師キーを収集
  const declinedDocKeys = new Set();
  arrowLines.forEach(function(l) {
    if (!l.includes('❌') && !l.includes('辞退')) return;
    const m1 = l.match(/❌\s*(?:辞退[：:])?\s*([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})/);
    if (m1 && typeof normKeyForDoc_ === 'function') {
      declinedDocKeys.add(normKeyForDoc_(m1[1]));
    }
  });

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

  if (isStayOrSkipDecision4_2_(dec)) return [];

  const firstLine = normCol1.split('\n')[0].trim();
  const clinicMatch = firstLine.match(/【([^（\(]+)(?:[（\(](小児科|内科)[）\)])?】/);
  const clinicName = clinicMatch ? clinicMatch[1].trim() : firstLine.replace(/[【】]/g, '').replace(/[（\(].*?[）\)]/g, '').trim();
  const dept = (clinicMatch && clinicMatch[2]) ? clinicMatch[2].trim() : (firstLine.includes('内科') ? '内科' : '小児科');
  const timeMatch = normCol1.match(/(\d{1,2}:\d{2}-\d{1,2}:\d{2})/);
  const rawTargetTime = timeMatch ? timeMatch[1] : '';

  let targetTime = rawTargetTime;
  if (rawTargetTime && typeof expandToLongestRecruitSlotTime1b3_ === 'function') {
    const ctx = getSubmitRecruitContext4_2_();
    const rawClinicWithDept = dept === '内科' ? (clinicName + '（内科）') : clinicName;
    targetTime = expandToLongestRecruitSlotTime1b3_(dateStr, rawClinicWithDept, rawTargetTime, ctx.recruitSlotsIndex, ctx.activeShifts);
  }

  const isSpecialWage = normCol1.includes('特別時給') || dec.includes('特別時給') || Boolean(spWageAddonLabel);
  const cleanAssignee = cleanAssigneeText4_2_(assigneeText, resolver);
  const memoSuffix = commentText ? ('（メモ: ' + commentText + '）') : '';

  const backupBaseParams = {
    dateStr: dateStr, clinicName: clinicName, dept: dept, targetTime: targetTime,
    isSpecialWage: isSpecialWage, spWageAddonLabel: spWageAddonLabel,
    cleanAssignee: cleanAssignee, commentText: commentText, memoSuffix: memoSuffix,
    resolver: resolver, isFullTimeChecker: isFullTimeChecker,
    backupDocsByDate: backupDocsByDate, backupUsedIdxByDateBand: backupUsedIdxByDateBand
  };

  if (dec.indexOf('バックアップ医師に勤務依頼') === 0) {
    return [buildBackupSubmitItem4_2_(Object.assign({}, backupBaseParams, { extraNote: '' }))];
  }

  const strippedDec = dec.replace(/[（\(][^）\)]*[）\)]/g, '').trim();
  const isPresetDec = /^(?:【?案\s*[123１２３]】?|全案)(?:[\s,、，\/／・\+＋と]+(?:【?案\s*[123１２３]】?|全案))*$/.test(strippedDec);

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
        date: dateStr, clinic_name: clinicName, raw_clinic_name: clinicName,
        department: dept, target_time: targetTime, time_category: '',
        approach_type: appType, target_doctor: finalName, license_no: finalLic,
        is_full_time: isFt, from_clinic: fromClinic,
        is_special_wage: isSpecialWage, special_wage_addon: spWageAddonLabel,
        assignee: cleanAssignee, memo: commentText, comment: commentText,
        evidence: '[編集くんカスタム確定] ' + dec + memoSuffix
      };
    });
  }

  const planMap = {};
  normCol2.split('\n').forEach(function(line) {
    const m = line.trim().match(/^【案([123１２３])】(.*)$/);
    if (m) {
      const num = m[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
      planMap[num] = m[2].trim();
    }
  });

  const chosenActionStrings = resolveChosenActionStrings4_2_(dec, strippedDec, planMap);
  const results = [];

  chosenActionStrings.forEach(function(actionObj) {
    let workingText = actionObj.text;
    let autoCondMemo = commentText;

    const trailingNoteMatch = workingText.match(/^(.*?先生[^\uff08\(]*?)\s*[（\(](?:メモ[：:])?([^）\)]+)[）\)]$/);
    if (trailingNoteMatch && !workingText.includes('2診目') && !workingText.includes('1診')) {
      workingText = trailingNoteMatch[1].trim();
      const extraMemo = trailingNoteMatch[2].trim();
      if (extraMemo && !/実績|非常勤|常勤|移動依頼手当|移動手当/.test(extraMemo)) {
        autoCondMemo = autoCondMemo ? (autoCondMemo + ' ／ ' + extraMemo) : extraMemo;
      }
    }

    const condMatch = workingText.match(/※(残る[^\n、，）\)]+?確認後)/);
    if (condMatch) {
      const condNote = '【要セット確認】' + condMatch[1];
      autoCondMemo = autoCondMemo ? (autoCondMemo + ' ／ ' + condNote) : condNote;
    }

    if (workingText.includes('バックアップ') && extractDoctorNamesFromSegment_(workingText).length === 0) {
      results.push(buildBackupSubmitItem4_2_(Object.assign({}, backupBaseParams, { extraNote: '(' + workingText + ')' })));
      return;
    }

    let parsedList = parseSingleActionTextToItems_(workingText, dateStr, clinicName, resolver);

    if (actionObj.planFallbackText && parsedList.length > 0) {
      const fallbackItems = parseSingleActionTextToItems_(actionObj.planFallbackText, dateStr, clinicName, resolver);
      parsedList = parsedList.map(function(p) {
        if (!p.doctorName) return p;
        const fbHit = fallbackItems.find(function(fb) {
          return fb.doctorName && (fb.doctorName.indexOf(p.doctorName) === 0 || p.doctorName.indexOf(fb.doctorName) === 0);
        });
        return fbHit ? {
          approachType: p.approachType || fbHit.approachType,
          doctorName: fbHit.doctorName,
          licenseNo: p.licenseNo || fbHit.licenseNo,
          fromClinic: p.fromClinic || fbHit.fromClinic
        } : p;
      });
    }

    if (actionObj.specifiedDoc && parsedList.length > 1) {
      const filtered = parsedList.filter(function(p) {
        return p.doctorName.includes(actionObj.specifiedDoc) || actionObj.specifiedDoc.includes(p.doctorName);
      });
      if (filtered.length > 0) parsedList = filtered;
    }

    parsedList.forEach(function(p) {
      if (!p.doctorName && (actionObj.label === 'カスタム指示' || actionObj.label.includes('カスタム'))) return;
      if (p.doctorName && typeof normKeyForDoc_ === 'function' && declinedDocKeys.has(normKeyForDoc_(p.doctorName))) return;

      const isFt = isFullTimeChecker ? isFullTimeChecker(p.doctorName, p.licenseNo, workingText) : false;
      results.push({
        date: dateStr, clinic_name: clinicName, raw_clinic_name: clinicName,
        department: dept, target_time: targetTime, time_category: '',
        approach_type: p.approachType, target_doctor: p.doctorName, license_no: p.licenseNo || '',
        is_full_time: isFt, from_clinic: p.fromClinic,
        is_special_wage: isSpecialWage, special_wage_addon: spWageAddonLabel,
        assignee: cleanAssignee, memo: autoCondMemo, comment: autoCondMemo,
        evidence: '[ドキュメント確定:' + actionObj.label + '] ' + workingText + (autoCondMemo ? ('（' + autoCondMemo + '）') : '')
      });
    });
  });

  return results;
}