/**
 * ============================================================================
 * AI会議_4_1_ドキュメント提出_メインと集計.gs
 * （50,000文字制限完全回避・常勤判定・「AまたはB選択的打診」と「複数枠おまとめ打診」自動判別・mnys送信メイン）
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

  const actSh = ss.getSheetByName('最新_当日シフト');
  if (actSh && actSh.getLastRow() >= 2) {
    const vals = actSh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      [vals[r][7], vals[r][8], vals[r][9]].forEach(function(cellTxt) {
        String(cellTxt || '').split(/[,、，\n]+/).forEach(function(tok) {
          if (!tok.includes('非常勤') && tok.includes('常勤')) {
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
    if (actionTextHint && !String(actionTextHint).includes('非常勤') && String(actionTextHint).includes('常勤')) return true;
    return false;
  };
}

/**
 * 時間帯文字列（例: "15:00-18:00"）から開始分・終了分を返すヘルパー
 */
function parseTimeMinutesForBundle4_(timeStr) {
  const m = String(timeStr || '').match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
  if (!m) return { startMin: 540, endMin: 780 };
  return {
    startMin: Number(m[1]) * 60 + Number(m[2]),
    endMin: Number(m[3]) * 60 + Number(m[4])
  };
}

/**
 * ★同一医師に複数枠を打診する際、
 * 1) 同時間帯（または別拠点直結）のロックを解除して両方選んだ場合 ➔ 「🔀【AまたはBの選択的打診】」
 * 2) 時間が被っておらず両方勤務可能な場合 ➔ 「🔗【複数枠おまとめ打診】」
 * を自動判別して mnys 提出データに付与する関数
 */
function annotateBundledDoctorItemsForMnys_(items) {
  const byDoctor = new Map();
  items.forEach(function(it) {
    const dKey = String(it.license_no || it.target_doctor || '').trim();
    if (!dKey) return;
    if (!byDoctor.has(dKey)) byDoctor.set(dKey, []);
    byDoctor.get(dKey).push(it);
  });

  byDoctor.forEach(function(docItems) {
    if (docItems.length < 2) {
      docItems[0].is_bundled = false;
      docItems[0].bundle_mode = 'SINGLE';
      docItems[0].bundle_slots_summary = '';
      return;
    }

    let isEitherOrChoice = false;
    for (let i = 0; i < docItems.length; i++) {
      for (let j = i + 1; j < docItems.length; j++) {
        if (docItems[i].date !== docItems[j].date) continue;
        const t1 = parseTimeMinutesForBundle4_(docItems[i].target_time);
        const t2 = parseTimeMinutesForBundle4_(docItems[j].target_time);
        const isOverlap = Math.max(t1.startMin, t2.startMin) < Math.min(t1.endMin, t2.endMin);
        const gapMin = (t1.endMin <= t2.startMin) ? (t2.startMin - t1.endMin) : (t1.startMin - t2.endMin);
        const isTightDiffClinic = (docItems[i].clinic_name !== docItems[j].clinic_name && gapMin < 60);

        if (isOverlap || isTightDiffClinic) {
          isEitherOrChoice = true;
          break;
        }
      }
      if (isEitherOrChoice) break;
    }

    const slotLabels = docItems.map(function(x) {
      return x.date.slice(5) + ' ' + x.clinic_name + '(' + x.target_time + ')';
    });

    const summaryText = isEitherOrChoice
      ? ('【選択的打診（AまたはB）】「' + slotLabels.join('」または「') + '」のどちらかで可能か打診')
      : ('【複数枠おまとめ打診（計' + docItems.length + '枠）】' + slotLabels.join(' ／ '));

    docItems.forEach(function(x) {
      x.is_bundled = true;
      x.bundle_mode = isEitherOrChoice ? 'EITHER_OR' : 'BUNDLE_ALL';
      x.bundle_count = docItems.length;
      x.bundle_slots = slotLabels;
      x.bundle_slots_summary = summaryText;
      x.memo = x.memo ? (x.memo + ' ｜ ' + summaryText) : summaryText;
      x.comment = x.memo;
      if (!String(x.evidence || '').includes(summaryText)) {
        x.evidence += ' ' + summaryText;
      }
    });
  });
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
  const backupDocsByDate = (typeof extractBackupDoctorsByDate64_ === 'function')
    ? extractBackupDoctorsByDate64_(ss, [])
    : {};
  const backupUsedIdxByDateBand = {};

  const draftMap = (typeof loadLatestDraft64_ === 'function')
    ? loadLatestDraft64_(ss, docId)
    : { slots: {}, bundles: {} };
  const allDraftSlotEntries = Object.values(draftMap.slots || {});

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  let currentDate = '';
  let defaultYear = new Date().getFullYear();
  const extractedItems = [];
  const seenKeys = new Set();

  const addUniqueItem = function(item) {
    const key = [item.date, item.clinic_name, item.department, item.target_time, item.approach_type, item.target_doctor, item.memo || ''].join('|');
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

          const rawC0 = row.getCell(0).getText().trim();
          const rowClinicTitle = (typeof formatClinicTitle64_ === 'function')
            ? formatClinicTitle64_(rawC0.split(/\r?\n/)[0] || '')
            : (rawC0.split(/\r?\n/)[0] || '').trim();

          const slotKey = i + '_' + r;
          let slotDraft = (draftMap.slots && draftMap.slots[slotKey]) ? draftMap.slots[slotKey] : null;
          if (slotDraft && slotDraft.clinicTitle && slotDraft.clinicTitle !== rowClinicTitle) {
            slotDraft = null;
          }
          if (!slotDraft && allDraftSlotEntries.length > 0 && typeof toSortDateKey64_ === 'function') {
            const matchedDrafts = allDraftSlotEntries.filter(function(d) {
              return toSortDateKey64_(d.dateLabel) === currentDate && d.clinicTitle === rowClinicTitle;
            });
            if (matchedDrafts.length === 1) slotDraft = matchedDrafts[0];
          }

          const customDocs = (slotDraft && Array.isArray(slotDraft.customDocs)) ? slotDraft.customDocs : [];
          const customComment = (slotDraft && slotDraft.customComment) ? String(slotDraft.customComment).trim() : '';

          const items = parseEditedDocTableRow_(
            currentDate,
            rawC0,
            row.getCell(1).getText().trim(),
            row.getCell(2).getText().trim(),
            row.getCell(3).getText().trim(),
            resolver,
            customDocs,
            customComment,
            isFullTimeChecker,
            backupDocsByDate,
            backupUsedIdxByDateBand
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

  annotateBundledDoctorItemsForMnys_(extractedItems);

  let outSheet = ss.getSheetByName('ドキュメント確定出力JSON');
  if (!outSheet) outSheet = ss.insertSheet('ドキュメント確定出力JSON');
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 13).setValues([[
    '日付', '拠点名', '診療科', '不在時間', 'アプローチ種別', '対象医師(正式名補正済)', '医籍番号', '移動元拠点', '特別時給・交渉加算', '担当者', 'コメント・おまとめ/選択打診メモ', '根拠・採用案', 'JSON出力(1行単位)'
  ]]);

  // ★1セル50,000文字制限を確実に回避するため、M列には各行ごとのJSONを出力
  const rowsForSheet = extractedItems.map(function(it) {
    let docTag = it.target_doctor + (it.is_full_time ? '(常勤)' : '');
    if (it.is_bundled) {
      docTag += (it.bundle_mode === 'EITHER_OR') ? '[★AまたはB選択打診]' : '[★おまとめ]';
    }
    const wageCellLabel = it.special_wage_addon
      ? ('特別時給(' + it.special_wage_addon + ')')
      : (it.is_special_wage ? '特別時給' : '通常時給');

    return [
      it.date,
      it.clinic_name,
      it.department,
      it.target_time,
      it.approach_type,
      docTag,
      it.license_no || '',
      it.from_clinic,
      wageCellLabel,
      it.assignee,
      String(it.memo || '').slice(0, 45000),
      String(it.evidence || '').slice(0, 45000),
      JSON.stringify(it).slice(0, 45000)
    ];
  });
  outSheet.getRange(2, 7, rowsForSheet.length, 1).setNumberFormat('@');
  outSheet.getRange(2, 1, rowsForSheet.length, 13).setValues(rowsForSheet);

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
    SpreadsheetApp.getUi().alert('✅ 全 ' + extractedItems.length + ' 件（AまたはB選択打診・特別時給交渉・要延長セット・おまとめ統合済）を ' + targetEmail + ' へ送信しました！');
  } else {
    SpreadsheetApp.getUi().alert(
      '✅ 全 ' + extractedItems.length + ' 件を『ドキュメント確定出力JSON』シートに出力しました！\n（※スクリプトプロパティに MNYS_RECEIVER_EMAIL を設定すると mnys.jp へ自動メール送信されます）'
    );
  }
}