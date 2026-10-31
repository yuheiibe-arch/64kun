/**
 * ============================================================================
 * AI会議_4_ドキュメント提出.gs
 * （『編集くん 64』のカスタム医師・医籍番号・常勤フラグ・特別時給ステイ除外完全対応版）
 * ============================================================================
 */

function buildFullTimeDoctorChecker_(ss) {
  const ftNames = new Set();
  const ftLicenses = new Set();

  const norm = function(str) {
    return String(str || '')
      .normalize('NFKC')
      .replace(/[\(（\[【].*$/g, '')
      .replace(/[\)）\]】★⚠️📩・]/g, '')
      .replace(/先生$/g, '')
      .replace(/[\s\u3000]+/g, '')
      .replace(/[﨑嵜]/g, '崎').replace(/髙/g, '高').replace(/[瀧泷]/g, '滝')
      .replace(/[邊邉]/g, '辺').replace(/[齋齊斉]/g, '斎').replace(/澤/g, '沢')
      .replace(/廣/g, '広').replace(/國/g, '国').replace(/[嶋嶌]/g, '島')
      .replace(/[栁柳]/g, '柳').replace(/[濵濱]/g, '浜').replace(/櫻/g, '桜')
      .replace(/眞/g, '真').replace(/惠/g, '恵')
      .trim();
  };

  // 1. 『常勤依頼手当』シートの医師は全員常勤
  const ftSh = ss.getSheetByName('常勤依頼手当');
  if (ftSh && ftSh.getLastRow() >= 2) {
    const vals = ftSh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const n = norm(vals[r][0]);
      const l = String(vals[r][1] || '').trim();
      if (n) ftNames.add(n);
      if (l && l !== '-') ftLicenses.add(l);
    }
  }

  // 2. 『最新_連勤注意リスト』に「(常勤」表記がある医師
  const stSh = ss.getSheetByName('最新_連勤注意リスト');
  if (stSh && stSh.getLastRow() >= 2) {
    const vals = stSh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const rawLabel = String(vals[r][1] || '');
      if (rawLabel.includes('常勤')) {
        const n = norm(rawLabel);
        const l = String(vals[r][0] || '').trim();
        if (n) ftNames.add(n);
        if (l && l !== '-') ftLicenses.add(l);
      }
    }
  }

  // 3. 『最新_当日シフト』に「(常勤」表記がある医師
  const actSh = ss.getSheetByName('最新_当日シフト');
  if (actSh && actSh.getLastRow() >= 2) {
    const vals = actSh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      [vals[r][7], vals[r][8], vals[r][9]].forEach(function(cellTxt) {
        String(cellTxt || '').split(/[,、，\n]+/).forEach(function(tok) {
          if (tok.includes('常勤')) {
            const n = norm(tok);
            if (n) ftNames.add(n);
          }
        });
      });
    }
  }

  return function(docName, licenseNo, actionTextHint) {
    const cleanLic = String(licenseNo || '').trim();
    if (cleanLic && ftLicenses.has(cleanLic)) return true;
    const n = norm(docName);
    if (n && ftNames.has(n)) return true;
    if (actionTextHint && String(actionTextHint).includes('常勤')) return true;
    return false;
  };
}

function submitEditedDocToMnys() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    SpreadsheetApp.getUi().alert('SINGLE_MTG_DOC_ID が見つかりません。先にAI会議ドキュメントを生成してください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resolver = buildCapsDoctorAndStaffResolver_(ss);
  const isFullTimeChecker = buildFullTimeDoctorChecker_(ss);

  const draftMap = (typeof loadLatestDraft64_ === 'function')
    ? loadLatestDraft64_(ss, docId)
    : { slots: {}, bundles: {} };

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  let currentDate = '';
  let defaultYear = new Date().getFullYear();
  const extractedItems = [];
  const seenKeys = new Set();

  const addUniqueItem = function(item) {
    const key = [item.date, item.clinic_name, item.department, item.target_time, item.approach_type, item.target_doctor].join('|');
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      extractedItems.push(item);
    }
  };

  for (let i = 0; i < numChildren; i++) {
    const el = body.getChild(i);
    const type = el.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const txt = el.asParagraph().getText().trim();
      const dateMatch = txt.match(/^◆\s*(\d{4})\/(\d{2})\/(\d{2})/);
      if (dateMatch) {
        currentDate = dateMatch[1] + '/' + dateMatch[2] + '/' + dateMatch[3];
        defaultYear = Number(dateMatch[1]);
      }
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      if (table.getNumRows() < 2) continue;
      const header0 = table.getRow(0).getCell(0).getText().trim();

      if (header0.includes('拠点・不在時間帯') && currentDate) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;

          const slotKey = i + '_' + r;
          const slotDraft = (draftMap.slots && draftMap.slots[slotKey]) ? draftMap.slots[slotKey] : null;
          const customDocs = (slotDraft && Array.isArray(slotDraft.customDocs)) ? slotDraft.customDocs : [];

          const items = parseEditedDocTableRow_(
            currentDate,
            row.getCell(0).getText().trim(),
            row.getCell(1).getText().trim(),
            row.getCell(2).getText().trim(),
            row.getCell(3).getText().trim(),
            resolver,
            customDocs,
            isFullTimeChecker
          );
          items.forEach(addUniqueItem);
        }
      }

      if (header0.includes('対象医師')) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;
          const bItems = parseBundleTableRow_(
            defaultYear,
            row.getCell(0).getText().trim(),
            row.getCell(1).getText().trim(),
            row.getCell(3).getText().trim(),
            resolver,
            isFullTimeChecker
          );
          bItems.forEach(addUniqueItem);
        }
      }
    }
  }

  if (extractedItems.length === 0) {
    SpreadsheetApp.getUi().alert('提出対象の行が見つかりませんでした（※1診・併行募集やステイ、特別時給＋ステイの行は自動的に除外されます）。');
    return;
  }

  let outSheet = ss.getSheetByName('ドキュメント確定出力JSON');
  if (!outSheet) outSheet = ss.insertSheet('ドキュメント確定出力JSON');
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 12).setValues([[
    '日付', '拠点名', '診療科', '不在時間', 'アプローチ種別', '対象医師(正式名補正済)', '医籍番号', '移動元拠点', '特別時給', '担当者', '根拠・採用案', 'JSON出力'
  ]]);

  const rowsForSheet = extractedItems.map(function(it, idx) {
    return [
      it.date, it.clinic_name, it.department, it.target_time, it.approach_type,
      it.target_doctor + (it.is_full_time ? '(常勤)' : ''), it.license_no || '', it.from_clinic, it.is_special_wage ? '特別時給' : '通常時給',
      it.assignee, it.evidence, idx === 0 ? JSON.stringify(extractedItems, null, 2) : ''
    ];
  });
  outSheet.getRange(2, 7, rowsForSheet.length, 1).setNumberFormat('@');
  outSheet.getRange(2, 1, rowsForSheet.length, 12).setValues(rowsForSheet);

  const targetEmail = props.getProperty('MNYS_RECEIVER_EMAIL');
  if (targetEmail) {
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const chunkSize = 30;
    const totalParts = Math.ceil(extractedItems.length / chunkSize);

    for (let p = 0; p < totalParts; p++) {
      const slice = extractedItems.slice(p * chunkSize, (p + 1) * chunkSize);
      const partLabel = totalParts > 1 ? ' [Part ' + (p + 1) + '/' + totalParts + ']' : '';
      GmailApp.sendEmail(
        targetEmail,
        '【Agent解析JSON】ドキュメント確定分 (' + todayStr + ')' + partLabel,
        JSON.stringify(slice, null, 2)
      );
    }
    SpreadsheetApp.getUi().alert('✅ 全 ' + extractedItems.length + ' 件（医籍番号・常勤判定連携済）を ' + targetEmail + ' へ送信しました！');
  } else {
    SpreadsheetApp.getUi().alert(
      '✅ 全 ' + extractedItems.length + ' 件を『ドキュメント確定出力JSON』シートに出力しました！\n（※スクリプトプロパティに MNYS_RECEIVER_EMAIL を設定すると mnys.jp へ自動メール送信されます）'
    );
  }
}

function mapCustomActionToApproachType_(actionTypeStr, fullText) {
  const s = String(actionTypeStr || '') + ' ' + String(fullText || '');
  if (s.includes('移動') || s.includes('2診目崩し')) return '勤務地変更相談';
  if (s.includes('延長')) return '延長相談';
  if (s.includes('早出') || s.includes('前倒')) return '早出相談';
  return '個別アプローチ';
}

function parseEditedDocTableRow_(dateStr, col1Text, col2Text, decisionText, assigneeText, resolver, customDocs, isFullTimeChecker) {
  const dec = String(decisionText || '').trim();
  // ★「ステイ」および「特別時給(+2,000円〜+5,000円)＋ステイ」の行は個別打診対象から除外
  if (!dec || /^(ステイ|特別時給.*ステイ|見送り|なし|不要|ー|\-|保留|要検討|併行募集)$/.test(dec)) return [];

  const firstLine = col1Text.split('\n')[0].trim();
  const clinicMatch = firstLine.match(/【([^（\(]+)(?:[（\(](小児科|内科)[）\)])?】/);
  const clinicName = clinicMatch ? clinicMatch[1].trim() : firstLine.replace(/[【】]/g, '').replace(/[（\(].*?[）\)]/g, '').trim();
  const dept = (clinicMatch && clinicMatch[2]) ? clinicMatch[2].trim() : (firstLine.includes('内科') ? '内科' : '小児科');
  const timeMatch = col1Text.match(/(\d{1,2}:\d{2}-\d{1,2}:\d{2})/);
  const targetTime = timeMatch ? timeMatch[1] : '';
  const isSpecialWage = col1Text.includes('特別時給') || dec.includes('特別時給');
  const cleanAssignee = resolver.resolveStaff(assigneeText);

  const isPresetDec = /^(案[123１２３]|全案)/.test(dec);
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
        assignee: cleanAssignee,
        evidence: '[編集くんカスタム確定] ' + dec
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
  if (/全案|すべて|全部/.test(dec)) {
    ['1', '2', '3'].forEach(function(k) {
      if (planMap[k]) chosenActionStrings.push({ label: '案' + k, text: planMap[k] });
    });
  } else {
    const tokens = dec.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(function(t) { return t.trim(); }).filter(Boolean);
    tokens.forEach(function(tok) {
      const planMatch = tok.match(/^(?:【?案)?([123１２３])(?:】)?$/);
      if (planMatch) {
        const num = planMatch[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
        if (planMap[num]) {
          chosenActionStrings.push({ label: '案' + num, text: planMap[num] });
        }
      } else {
        const multiNums = tok.match(/案([123１２３])/g);
        if (multiNums && tok.replace(/案[123１２３]|[\s、,・と]/g, '') === '') {
          multiNums.forEach(function(mn) {
            const num = mn.replace('案', '').replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
            if (planMap[num]) chosenActionStrings.push({ label: '案' + num, text: planMap[num] });
          });
        } else {
          chosenActionStrings.push({ label: 'カスタム指示', text: tok });
        }
      }
    });
  }

  const results = [];
  chosenActionStrings.forEach(function(actionObj) {
    const parsedList = parseSingleActionTextToItems_(actionObj.text, dateStr, clinicName, resolver);
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
        assignee: cleanAssignee,
        evidence: '[ドキュメント確定:' + actionObj.label + '] ' + actionObj.text
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
  const cleanAssignee = resolver.resolveStaff(assigneeFromDec);

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
          assignee: cleanAssignee,
          evidence: '[ドキュメント確定:おまとめ打診] ' + resolvedDoc.name + '先生 (' + currentClinic + ' ' + dateStr + ' ' + tRange + ')'
        });
      });
    }
  });

  return items;
}

function parseSingleActionTextToItems_(text, dateStr, targetClinic, resolver) {
  const s = String(text || '').trim();
  if (!s) return [];

  if (s.includes('1診・併行募集') || s.includes('併行募集') || s.includes('埋まり次第')) {
    return [];
  }

  let approachType = '個別アプローチ';
  let fromClinic = '';
  let rawDoctors = [];

  if (s.includes('2診目崩し') || s.includes('移動')) {
    approachType = '勤務地変更相談';
    const fromMatch = s.match(/】([^\(（\s]+?)[（\(][12]診/) || s.match(/([^\s【】]+?)(?:より|から)/);
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