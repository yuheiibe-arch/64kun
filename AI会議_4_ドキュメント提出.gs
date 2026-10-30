/**
 * ============================================================================
 * 4_AI会議_ドキュメント提出.gs
 * （人間が編集したGoogleドキュメントを読み取り、1診併行募集はスキップし、有効な打診のみmnysへ送信）
 * ============================================================================
 */

function submitEditedDocToMnys() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    SpreadsheetApp.getUi().alert('SINGLE_MTG_DOC_ID が見つかりません。先にAI会議ドキュメントを生成してください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resolver = buildCapsDoctorAndStaffResolver_(ss);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  let currentDate = '';
  let defaultYear = new Date().getFullYear();
  const extractedItems = [];
  const seenKeys = new Set();

  const addUniqueItem = (item) => {
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
        currentDate = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`;
        defaultYear = Number(dateMatch[1]);
      }
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      if (table.getNumRows() < 2) continue;
      const header0 = table.getRow(0).getCell(0).getText().trim();

      // A. 第1部：日付別アプローチ表
      if (header0.includes('拠点・不在時間帯') && currentDate) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;
          const items = parseEditedDocTableRow_(
            currentDate,
            row.getCell(0).getText().trim(),
            row.getCell(1).getText().trim(),
            row.getCell(2).getText().trim(),
            row.getCell(3).getText().trim(),
            resolver
          );
          items.forEach(addUniqueItem);
        }
      }

      // B. ★複数日程おまとめ打診リスト
      if (header0.includes('対象医師')) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;
          const bItems = parseBundleTableRow_(
            defaultYear,
            row.getCell(0).getText().trim(),
            row.getCell(1).getText().trim(),
            row.getCell(3).getText().trim(),
            resolver
          );
          bItems.forEach(addUniqueItem);
        }
      }
    }
  }

  if (extractedItems.length === 0) {
    SpreadsheetApp.getUi().alert('提出対象の行が見つかりませんでした（※1診・併行募集やステイの行は自動的に除外されます）。');
    return;
  }

  let outSheet = ss.getSheetByName('ドキュメント確定出力JSON');
  if (!outSheet) outSheet = ss.insertSheet('ドキュメント確定出力JSON');
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 12).setValues([[
    '日付', '拠点名', '診療科', '不在時間', 'アプローチ種別', '対象医師(正式名補正済)', '医籍番号', '移動元拠点', '特別時給', '担当者', '根拠・採用案', 'JSON出力'
  ]]);

  const rowsForSheet = extractedItems.map((it, idx) => [
    it.date, it.clinic_name, it.department, it.target_time, it.approach_type,
    it.target_doctor, it.license_no || '', it.from_clinic, it.is_special_wage ? '特別時給' : '通常時給',
    it.assignee, it.evidence, idx === 0 ? JSON.stringify(extractedItems, null, 2) : ''
  ]);
  outSheet.getRange(2, 1, rowsForSheet.length, 12).setValues(rowsForSheet);

  const targetEmail = props.getProperty('MNYS_RECEIVER_EMAIL');
  if (targetEmail) {
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const chunkSize = 30;
    const totalParts = Math.ceil(extractedItems.length / chunkSize);

    for (let p = 0; p < totalParts; p++) {
      const slice = extractedItems.slice(p * chunkSize, (p + 1) * chunkSize);
      const partLabel = totalParts > 1 ? ` [Part ${p + 1}/${totalParts}]` : '';
      GmailApp.sendEmail(
        targetEmail,
        `【Agent解析JSON】ドキュメント確定分 (${todayStr})${partLabel}`,
        JSON.stringify(slice, null, 2)
      );
    }
    SpreadsheetApp.getUi().alert(`✅ 全 ${extractedItems.length} 件（医師名補正済・1診併行募集は除外）を ${targetEmail} へ送信しました！`);
  } else {
    SpreadsheetApp.getUi().alert(
      `✅ 全 ${extractedItems.length} 件（医師名補正済・1診併行募集は除外）を『ドキュメント確定出力JSON』シートに出力しました！\n（※スクリプトプロパティに MNYS_RECEIVER_EMAIL を設定すると mnys.jp へ自動メール送信されます）`
    );
  }
}

function parseEditedDocTableRow_(dateStr, col1Text, col2Text, decisionText, assigneeText, resolver) {
  const dec = String(decisionText || '').trim();
  if (!dec || /^(ステイ|見送り|なし|不要|ー|\-|保留|要検討|併行募集)$/.test(dec)) return [];

  const clinicMatch = col1Text.match(/【([^（\(]+)[（\(](小児科|内科)[）\)]】/);
  const clinicName = clinicMatch ? clinicMatch[1].trim() : col1Text.split('\n')[0].replace(/[【】]/g, '').trim();
  const dept = clinicMatch ? clinicMatch[2].trim() : (col1Text.includes('内科') ? '内科' : '小児科');
  const timeMatch = col1Text.match(/(\d{1,2}:\d{2}-\d{1,2}:\d{2})/);
  const targetTime = timeMatch ? timeMatch[1] : '';
  const isSpecialWage = col1Text.includes('特別時給');
  const cleanAssignee = resolver.resolveStaff(assigneeText);

  const planMap = {};
  col2Text.split('\n').forEach(line => {
    const m = line.match(/^【案([123１２３])】(.*)$/);
    if (m) {
      const num = m[1].replace(/[１２３]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
      planMap[num] = m[2].trim();
    }
  });

  const chosenActionStrings = [];
  if (/全案|すべて|全部/.test(dec)) {
    ['1', '2', '3'].forEach(k => {
      if (planMap[k]) chosenActionStrings.push({ label: `案${k}`, text: planMap[k] });
    });
  } else {
    const tokens = dec.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(t => t.trim()).filter(Boolean);
    tokens.forEach(tok => {
      const planMatch = tok.match(/^(?:【?案)?([123１２３])(?:】)?$/);
      if (planMatch) {
        const num = planMatch[1].replace(/[１２３]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        if (planMap[num]) {
          chosenActionStrings.push({ label: `案${num}`, text: planMap[num] });
        }
      } else {
        const multiNums = tok.match(/案([123１２３])/g);
        if (multiNums && tok.replace(/案[123１２３]|[\s、,・と]/g, '') === '') {
          multiNums.forEach(mn => {
            const num = mn.replace('案', '').replace(/[１２３]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
            if (planMap[num]) chosenActionStrings.push({ label: `案${num}`, text: planMap[num] });
          });
        } else {
          chosenActionStrings.push({ label: 'カスタム指示', text: tok });
        }
      }
    });
  }

  const results = [];
  chosenActionStrings.forEach(actionObj => {
    const parsedList = parseSingleActionTextToItems_(actionObj.text, dateStr, clinicName, resolver);
    parsedList.forEach(p => {
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
        from_clinic: p.fromClinic,
        is_special_wage: isSpecialWage,
        assignee: cleanAssignee,
        evidence: `[ドキュメント確定:${actionObj.label}] ${actionObj.text}`
      });
    });
  });

  return results;
}

function parseBundleTableRow_(defaultYear, docColText, slotsColText, decColText, resolver) {
  const dec = String(decColText || '').trim();
  if (!dec || /^(ステイ|見送り|なし|不要|ー|\-|保留|不採用)$/.test(dec)) return [];

  const rawDoc = String(docColText || '').replace(/[\(（].*?[\)）]/g, '').replace(/先生/g, '').trim();
  const resolvedDoc = resolver.resolveDoctor(rawDoc, '', '');
  const assigneeFromDec = dec.replace(/採用|OK|◯|○|[\/／・,\s ]/g, '').trim();
  const cleanAssignee = resolver.resolveStaff(assigneeFromDec);

  const items = [];
  let currentClinic = '';
  let currentDept = '小児科';

  String(slotsColText || '').split('\n').map(l => l.trim()).filter(Boolean).forEach(line => {
    if (line.startsWith('■')) {
      const rawC = line.replace(/^■\s*/, '').trim();
      currentDept = rawC.includes('内科') ? '内科' : '小児科';
      currentClinic = rawC.replace(/[\(（].*?[\)）]/g, '').trim();
      return;
    }
    const dm = line.match(/(\d{1,2})\/(\d{1,2})[^\d]*(\d{1,2}:\d{2}-\d{1,2}:\d{2}(?:\s*\/\s*\d{1,2}:\d{2}-\d{1,2}:\d{2})*)/);
    if (dm && currentClinic) {
      const dateStr = `${defaultYear}/${('0' + dm[1]).slice(-2)}/${('0' + dm[2]).slice(-2)}`;
      dm[3].split('/').map(t => t.trim()).filter(Boolean).forEach(tRange => {
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
          from_clinic: '',
          is_special_wage: false,
          assignee: cleanAssignee,
          evidence: `[ドキュメント確定:おまとめ打診] ${resolvedDoc.name}先生 (${currentClinic} ${dateStr} ${tRange})`
        });
      });
    }
  });

  return items;
}

/**
 * ★修正：「1診・併行募集」（まだ移動依頼をしない案）はスプレッドシート書き出し不要のため空配列 [] を返してスキップ
 */
function parseSingleActionTextToItems_(text, dateStr, targetClinic, resolver) {
  const s = String(text || '').trim();
  if (!s) return [];

  // 1診・併行募集は移動打診をまだ行わないため書き出し対象外
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

  if (rawDoctors.length === 0) return [{ approachType, doctorName: '', licenseNo: '', fromClinic }];

  return rawDoctors.map(rawDoc => {
    const resolved = resolver.resolveDoctor(rawDoc, dateStr, fromClinic || targetClinic);
    const finalFromClinic = (approachType === '勤務地変更相談' && !fromClinic && resolved.dayClinic && resolved.dayClinic !== targetClinic)
      ? resolved.dayClinic
      : fromClinic;
    return { approachType, doctorName: resolved.name, licenseNo: resolved.licenseNo, fromClinic: finalFromClinic };
  });
}

function extractDoctorNamesFromSegment_(seg) {
  const cleaned = String(seg || '')
    .replace(/\[.*?\]|［.*?］/g, '')
    .replace(/[（\(].*?[）\)]/g, '')
    .replace(/前医の|後医の|どちらか1名.*$|へ.*$|に.*$|を.*$/g, '')
    .trim();

  return cleaned.split(/[\/／]|または|・|,|、/)
    .map(p => p.replace(/[①②③④123１２３:\s ]|先生/g, '').trim())
    .filter(p => p.length >= 2 && !/^(スポット|紹介会社|常勤|非常勤|移動|延長|早出|個別)$/.test(p));
}