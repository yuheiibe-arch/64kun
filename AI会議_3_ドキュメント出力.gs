/**
 * ============================================================================
 * 3_AI会議_ドキュメント描画と通知.gs
 * （初期値「案1」固定解除・第2部のお知らせ配信1ページ完全収容・高速描画版）
 * ============================================================================
 */

const DOC_OUTPUT_FOLDER_ID = '1syiypDralKBqHwCL-0R2XeJXbhxXVcLl';

function formatDateWithDow_(dateStr) {
  const s = String(dateStr || '').trim();
  if (!s || /[\(（][日月火水木金土][\)）]/.test(s)) return s;
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return s;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (isNaN(d.getTime())) return s;
  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) + '（' + dows[d.getDay()] + '）';
}

function getOrResetDailyMeetingDocInFolder_(title) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SINGLE_MTG_DOC_ID';
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  let folder = null;
  try {
    folder = DriveApp.getFolderById(DOC_OUTPUT_FOLDER_ID);
  } catch (e) {
    Logger.log('⚠️ 指定フォルダの取得に失敗したためマイドライブ直下を使用します: ' + e.message);
  }

  let targetDoc = null;
  let targetDocId = '';
  const sameDayFiles = [];

  if (folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      const name = f.getName();
      if (name.includes('超募集対策MTG_' + todayStr)) {
        sameDayFiles.push(f);
      }
    }
  }

  if (sameDayFiles.length > 0) {
    sameDayFiles.sort(function(a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
    const keepFile = sameDayFiles[0];
    targetDocId = keepFile.getId();
    targetDoc = DocumentApp.openById(targetDocId);
    targetDoc.setName(title);

    for (let i = 1; i < sameDayFiles.length; i++) {
      try { sameDayFiles[i].setTrashed(true); } catch (e) {}
    }
  }

  if (!targetDoc) {
    const savedId = props.getProperty(propKey);
    if (savedId) {
      try {
        const savedFile = DriveApp.getFileById(savedId);
        if (!savedFile.isTrashed() && savedFile.getName().includes('超募集対策MTG_' + todayStr)) {
          targetDocId = savedId;
          targetDoc = DocumentApp.openById(targetDocId);
          targetDoc.setName(title);
          if (folder) moveFileToFolderSafely_(savedFile, folder);
        }
      } catch (e) {}
    }
  }

  if (!targetDoc) {
    targetDoc = DocumentApp.create(title);
    targetDocId = targetDoc.getId();
    const newFile = DriveApp.getFileById(targetDocId);
    if (folder) {
      moveFileToFolderSafely_(newFile, folder);
    }
  }

  props.setProperty(propKey, targetDocId);
  return targetDoc;
}

function moveFileToFolderSafely_(file, targetFolder) {
  try {
    file.moveTo(targetFolder);
  } catch (e) {
    try {
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e2) {}
  }
}

function renderStructuredMeetingDoc_(title, rawContext, aiData) {
  const doc = getOrResetDailyMeetingDocInFolder_(title);
  const body = doc.getBody();
  body.setMarginTop(20).setMarginBottom(20).setMarginLeft(24).setMarginRight(24);

  const titlePara = body.appendParagraph(title);
  while (body.getNumChildren() > 1) {
    body.removeChild(body.getChild(0));
  }
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2).setSpacingAfter(4);

  const metaText = [
    '📅 作成日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' ｜ 対象期間：' + formatDateWithDow_(rawContext.periodStart) + ' ～ ' + formatDateWithDow_(rawContext.periodEnd) + '（対象不在枠：' + rawContext.unfilledCount + '件）',
    '💡 手当ルール：常勤追加勤務手当（1,500円/h）＝『常勤依頼手当』シート記載医師のみ ／ 拠点移動手当（10,000円）＝一律支給（SBUのみ0円）',
    '✍️ 確定案の書き方：初期値は未選択（空欄）です。採用する案（「案1」「案1, 案2」「全案」等）または「ステイ」を選択してください。担当者を後から決める場合は「あとで決定」を選ぶと決定済み扱いとなり、mnys側へは「未定」で送信されます。'
  ].join('\n');
  body.appendParagraph(metaText).setFontSize(8.5).setForegroundColor('#444444').setLineSpacing(1.25).setSpacingAfter(6);

  if (aiData.executiveSummary) {
    const sumTable = body.appendTable([['🤖 AIエージェント総括：' + aiData.executiveSummary]]);
    sumTable.getCell(0, 0).setBackgroundColor('#eef4ff').setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
    sumTable.getCell(0, 0).editAsText().setFontSize(9).setBold(true);
  }

  // 第1部：日付ごとのアプローチ表
  body.appendParagraph('第1部：医師不在枠 アプローチプラン（日付別一覧）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3).setSpacingBefore(8).setSpacingAfter(4);

  const plansByDate = {};
  (aiData.slotPlans || []).forEach(function(p) {
    const d = p.date || '日付未定';
    if (!plansByDate[d]) plansByDate[d] = [];
    plansByDate[d].push(p);
  });

  Object.keys(plansByDate).sort().forEach(function(dateLabel) {
    body.appendParagraph('◆ ' + formatDateWithDow_(dateLabel))
      .setFontSize(10).setBold(true).setSpacingBefore(8).setSpacingAfter(3);

    const tableRows = [
      ['拠点・不在時間帯', 'AIエージェント推奨アプローチ（優先順・現状メモ）', '確定案（複数可・カスタム可）', '担当者名']
    ];

    plansByDate[dateLabel].forEach(function(p) {
      let cleanClinic = String(p.clinicAndDept || '')
        .replace(' (小児科)（小児科）', '（小児科）')
        .replace(' (内科)（内科）', '（内科）');
      if (!cleanClinic.includes('亀有') && !cleanClinic.includes('北葛西')) {
        cleanClinic = cleanClinic.replace(/[（\(]小児科[）\)]/g, '').trim();
      }
      const col1 = cleanClinic + '\n' + (p.absenceTime || '');

      const col2Lines = [];
      if (p.plan1) col2Lines.push('【案1】' + p.plan1);
      if (p.plan2) col2Lines.push('【案2】' + p.plan2);
      if (p.plan3) col2Lines.push('【案3】' + p.plan3);
      if (p.auditNote && p.auditNote !== '-') col2Lines.push('💡現状・注意：' + p.auditNote);

      const col2 = col2Lines.join('\n') || '要個別検討';

      // ★初期値で「案1」に固定せず、未選択（空欄）にする
      const defaultDec = '';

      tableRows.push([col1, col2, defaultDec, p.assignee || '']);
    });

    const table = body.appendTable(tableRows);
    styleReadableDocTable_(table, '#1a73e8', [115, 265, 95, 65]);
  });

  // 第2部：急募のお知らせ配信（改ページで新ページの先頭から開始し、フォント6.5pt・行間1.0・段落余白0で7日分を確実に1ページ内に収める）
  body.appendPageBreak();
  body.appendParagraph('第2部：急募のお知らせ配信（1週間分・1ページ完全収容テンプレート）')
    .setFontSize(9).setBold(true).setSpacingBefore(0).setSpacingAfter(1);

  if (aiData.announcementProposal) {
    body.appendParagraph('📌 ' + (aiData.announcementProposal.recommendedDateReason || ''))
      .setFontSize(7.5).setBold(true).setSpacingBefore(0).setSpacingAfter(2);

    const tplTable = body.appendTable([[aiData.announcementProposal.templateText || '']]);
    const cell = tplTable.getCell(0, 0);
    cell.setBackgroundColor('#fff8e1').setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(6).setPaddingRight(6);
    cell.editAsText().setFontSize(6.5);
    for (let p = 0; p < cell.getNumChildren(); p++) {
      const ch = cell.getChild(p);
      if (ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
        ch.asParagraph().setLineSpacing(1.0).setSpacingBefore(0).setSpacingAfter(0);
      }
    }
  }

  // 第3部は別ページにして第2部の1ページ収容を邪魔しないようにする
  body.appendPageBreak();
  body.appendParagraph('第3部：【参考・サブ情報】対象期間内の12連勤NG日')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3).setSpacingBefore(4).setSpacingAfter(4);

  if (rawContext.streakList && rawContext.streakList.length > 0) {
    const sRows = [['医師名', '連勤数', '対象期間（実働拠点のみ）', '追加打診NG日', '理由']];
    rawContext.streakList.forEach(function(s) {
      sRows.push([s.doctor + '(' + s.licenseNo + ')', s.streakDays + '日', s.period || '', s.ngDates || '-', s.note || '']);
    });
    styleReadableDocTable_(body.appendTable(sRows), '#c53929', [100, 45, 150, 75, 170]);
  } else {
    body.appendParagraph('・対象期間内に12連勤NG日が発生する医師はいません。').setFontSize(8.5);
  }

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT); } catch (e) {}
  return file;
}

function styleReadableDocTable_(table, headerHexColor, colWidths) {
  table.setBorderColor('#cccccc').setBorderWidth(0.5);
  const numRows = table.getNumRows();
  const numCols = table.getRow(0).getNumCells();

  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const isHeader = (r === 0);
    const bgHex = isHeader ? headerHexColor : (r % 2 === 1 ? '#ffffff' : '#f8f9fa');

    for (let c = 0; c < numCols; c++) {
      const cell = row.getCell(c);
      if (colWidths && colWidths[c]) cell.setWidth(colWidths[c]);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
      cell.setBackgroundColor(bgHex);

      const textObj = cell.editAsText();
      if (isHeader) {
        textObj.setForegroundColor('#ffffff').setBold(true).setFontSize(8.5);
      } else {
        textObj.setForegroundColor('#222222').setBold(false).setFontSize(8.5);
        if (c === 1) {
          highlightKeywordsInText_(textObj);
        }
      }
    }
  }
}

function highlightKeywordsInText_(textObj) {
  ['【案1】', '【案2】', '【案3】', '【2診目崩し】', '【1診・併行募集】', '💡現状・注意：'].forEach(function(pat) {
    let found = textObj.findText(pat);
    while (found) {
      found.getElement().asText().setBold(found.getStartOffset(), found.getEndOffsetInclusive(), true);
      found = textObj.findText(pat, found);
    }
  });
}

function replaceAiDocLinkInDescription_(originalDesc, newDocUrl) {
  const cleaned = String(originalDesc || '').replace(/\n*📄 【AI事前会議ドキュメント】\nhttps?:\/\/[^\s\n]+/g, '').trim();
  return (cleaned ? cleaned + '\n\n' : '') + '📄 【AI事前会議ドキュメント】\n' + newDocUrl;
}

function attachDocToMtgCalendar_(docFile) {
  const now = new Date();
  const startWindow = new Date(now.getTime() - 12 * 60 * 60 * 1000);
  const lookAhead = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  if (typeof Calendar !== 'undefined' && Calendar.Events) {
    try {
      const events = Calendar.Events.list(MTG_DOC_CONFIG.CALENDAR_ID, {
        q: MTG_DOC_CONFIG.EVENT_TITLE_KEYWORD,
        timeMin: startWindow.toISOString(),
        timeMax: lookAhead.toISOString(),
        singleEvents: true,
        orderBy: 'startTime'
      });

      if (events.items && events.items.length > 0) {
        const ev = events.items[0];
        const otherAtts = (ev.attachments || []).filter(function(a) {
          return !String(a.title || '').includes('超募集対策MTG_') && a.fileUrl !== docFile.getUrl();
        });
        otherAtts.push({ fileUrl: docFile.getUrl(), title: docFile.getName(), mimeType: 'application/vnd.google-apps.document' });
        const patched = Calendar.Events.patch(
          { attachments: otherAtts, description: replaceAiDocLinkInDescription_(ev.description, docFile.getUrl()) },
          MTG_DOC_CONFIG.CALENDAR_ID, ev.id, { supportsAttachments: true }
        );
        return patched.htmlLink || '';
      }
    } catch (e) {}
  }

  try {
    const cal = CalendarApp.getCalendarById(MTG_DOC_CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
    if (!cal) return '';
    const calEvents = cal.getEvents(startWindow, lookAhead, { search: MTG_DOC_CONFIG.EVENT_TITLE_KEYWORD });
    if (!calEvents || calEvents.length === 0) return '';
    calEvents[0].setDescription(replaceAiDocLinkInDescription_(calEvents[0].getDescription(), docFile.getUrl()));
    return 'https://' + 'calendar.google.com/calendar/r/week';
  } catch (e2) {
    return '';
  }
}

function sendMeetingDocToSlack_(info) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const channel = props.getProperty('SLACK_CHANNEL_ID') || MTG_DOC_CONFIG.DEFAULT_SLACK_CHANNEL_ID;
  if (!token) return;

  const mentionText = resolveSubteamMentions_(token, ['dsshift', 'dspart']);
  const fullText = mentionText + '\n```\n' + info.todayStr + '\n対象期間：' + info.periodStart + '～' + info.periodEnd + '\nAI自動会議終了いたしました。提出済みです。\n```';
  const buttons = [
    { type: 'button', text: { type: 'plain_text', text: '📄 AI会議ドキュメントを開く', emoji: true }, style: 'primary', url: info.docUrl },
    { type: 'button', text: { type: 'plain_text', text: '📓 AIルールブック(Notebook)を開く', emoji: true }, url: MTG_DOC_CONFIG.NOTEBOOK_URL }
  ];
  if (info.eventUrl) {
    buttons.push({ type: 'button', text: { type: 'plain_text', text: '📅 カレンダー予定を開く', emoji: true }, url: info.eventUrl });
  }

  UrlFetchApp.fetch('https://' + 'slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      channel: channel,
      text: fullText,
      blocks: [
        { type: 'section', text: { type: 'mrkdwn', text: fullText } },
        { type: 'actions', elements: buttons }
      ]
    }),
    muteHttpExceptions: true
  });
}

function resolveSubteamMentions_(token, handles) {
  try {
    const res = UrlFetchApp.fetch('https://' + 'slack.com/api/usergroups.list', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.ok && Array.isArray(data.usergroups)) {
      const map = new Map();
      data.usergroups.forEach(function(ug) {
        if (ug.handle) map.set(ug.handle.toLowerCase(), ug.id);
        if (ug.name) map.set(ug.name.toLowerCase(), ug.id);
      });
      return handles.map(function(h) {
        return map.has(h.toLowerCase()) ? ('<!subteam^' + map.get(h.toLowerCase()) + '|@' + h + '>') : ('@' + h);
      }).join(' ');
    }
  } catch (e) {}
  return handles.map(function(h) { return '@' + h; }).join(' ');
}