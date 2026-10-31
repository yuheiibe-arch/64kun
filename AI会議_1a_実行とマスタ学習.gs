/**
 * ============================================================================
 * AI会議_1a_実行とマスタ学習.gs
 * （『交渉しやすい医師』無断行追加禁止・列ズレ自動修復・『交渉注意医師』統合版）
 * ============================================================================
 */

const MTG_DOC_CONFIG = {
  CALENDAR_ID: 'capsdoctor365@caps365.jp',
  EVENT_TITLE_KEYWORD: '超募集対策MTG',
  DEFAULT_SLACK_CHANNEL_ID: 'C09TRHGU64D',
  NOTEBOOK_URL: 'https://' + 'gemini.google.com/notebook/570c0b20-03f4-4ced-aee3-cf07a2cbeddb',
  GEMINI_MODELS: [
    'gemini-3.7-flash',
    'gemini-3.6-flash',
    'gemini-3.5-flash',
    'gemini-flash-lite-latest',
    'gemini-3.5-flash-lite'
  ],
  RULE_SHEET_NAME: 'AI会議ルール',
  BATCH_SIZE: 25
};

/**
 * 成果物プレビュー用（Slack通知なし・カレンダー添付なし）
 */
function runPreviewMeetingDocOnly() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 先にデバッグ用シートが残っていれば自動削除
  const oldDbg = ss.getSheetByName('デバッグ_不在枠解析');
  if (oldDbg) {
    try { ss.deleteSheet(oldDbg); } catch (e) {}
  }

  const customRulesText = loadOrInitCustomRuleSheet_(ss);

  Logger.log('🧠 0/3: マスタの列ズレ修復 ＆ ナレッジシートの新しい医師返答を確認中...');
  repairNegotiableAndCautionSheetsNow(ss);
  syncNewKnowledgeToAiMemo_(ss);

  Logger.log('🤖 1/3: シートから対象データを抽出し、AIエージェント用のコンテキストを構築中...');
  const rawContext = collectCompactContextForAi_(ss);
  Logger.log('📌 対象不在枠: ' + rawContext.unfilledCount + '件（' + rawContext.periodStart + '〜' + rawContext.periodEnd + '）');

  Logger.log('🤖 2/3: AIエージェントが高速会議・厳選中...');
  const aiStructuredResult = callGeminiAgentInBatches_(rawContext, customRulesText);

  Logger.log('📄 3/3: Googleドキュメントに日付別テーブル形式で描画中...');
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const docTitle = '【プレビュー確認用】超募集対策MTG_' + todayStr + '（対象：' + rawContext.periodStart + '〜' + rawContext.periodEnd + '）';
  const docFile = renderStructuredMeetingDoc_(docTitle, rawContext, aiStructuredResult);

  Logger.log('✅ 成果物ドキュメント作成完了！（所要時間: ' + ((Date.now() - t0) / 1000).toFixed(1) + '秒）');
  Logger.log(docFile.getUrl());
}

/**
 * 本番用：AIメモ自動学習 ＋ ドキュメント作成 ＋ カレンダー添付 ＋ Slack通知
 */
function runDemoAndSubmitMeetingDoc() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const oldDbg = ss.getSheetByName('デバッグ_不在枠解析');
  if (oldDbg) {
    try { ss.deleteSheet(oldDbg); } catch (e) {}
  }

  const customRulesText = loadOrInitCustomRuleSheet_(ss);
  repairNegotiableAndCautionSheetsNow(ss);
  syncNewKnowledgeToAiMemo_(ss);

  const rawContext = collectCompactContextForAi_(ss);
  const aiStructuredResult = callGeminiAgentInBatches_(rawContext, customRulesText);

  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const docTitle = '【AI事前会議】超募集対策MTG_' + todayStr + '（対象：' + rawContext.periodStart + '〜' + rawContext.periodEnd + '）';
  const docFile = renderStructuredMeetingDoc_(docTitle, rawContext, aiStructuredResult);

  const eventUrl = attachDocToMtgCalendar_(docFile);

  sendMeetingDocToSlack_({
    todayStr: todayStr,
    periodStart: rawContext.periodStart,
    periodEnd: rawContext.periodEnd,
    docUrl: docFile.getUrl(),
    eventUrl: eventUrl
  });

  Logger.log('✅ 会議ドキュメント生成＆Slack/カレンダー提出完了！（所要時間: ' + ((Date.now() - t0) / 1000).toFixed(1) + '秒）');
  Logger.log(docFile.getUrl());
}

/**
 * 毎週月曜・水曜の朝に自動起動するトリガー設定
 */
function installWeeklyMtgDocTrigger() {
  const funcName = 'runDemoAndSubmitMeetingDoc';
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(funcName).timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(10).nearMinute(0).create();
  ScriptApp.newTrigger(funcName).timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(10).nearMinute(0).create();
  Logger.log('✅ 毎週月曜・水曜の自動ドキュメント提出トリガーを設定しました。');
}

/**
 * ナレッジシートの最新返信（末尾30行のみ高速取得）をAI判定し、『交渉注意医師』および既存の『交渉しやすい医師』に反映する
 */
function syncNewKnowledgeToAiMemo_(ss) {
  try {
    const kbSheet = ss.getSheetByName('ナレッジシート');
    if (!kbSheet) return;
    const lastRow = kbSheet.getLastRow();
    if (lastRow < 2) return;

    const startRow = Math.max(2, lastRow - 29);
    const numRows = lastRow - startRow + 1;
    const kbData = kbSheet.getRange(startRow, 1, numRows, 8).getDisplayValues();
    const targetItems = [];

    for (let i = 0; i < kbData.length; i++) {
      const r = kbData[i];
      const actualRowNumber = startRow + i;
      const docName = String(r[1] || '').trim();
      const bodyText = String(r[2] || '').trim();
      const flagText = String(r[7] || '').trim();

      if (!docName || !bodyText) continue;
      if (flagText.includes('AIメモ反映済') || flagText.includes('自動判定済')) continue;

      const rawColA = String(r[0] || '').trim();
      const validLic = (/^\d{5,7}$/.test(rawColA) && Number(rawColA) >= 100000) ? rawColA : '';

      targetItems.push({
        rowNumber: actualRowNumber,
        licenseNo: validLic,
        doctor: docName,
        replyText: bodyText.slice(0, 300),
        date: String(r[6] || '').trim(),
        currentFlag: flagText
      });
    }

    if (targetItems.length === 0) return;

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return;

    const systemPrompt = [
      'あなたはCAPSクリニックの医師マスタ管理AIです。',
      '医師からの返信メール一覧を読み、今後の非常勤シフト打診で継続して覚えておくべきNG条件や曜日・時間帯・拠点の制約（例：「東品川は遠いためNG」「火曜・金曜の日中のみ希望」「以降の時間帯は不可」など）が含まれているものだけを抽出してください。',
      '🚫除外対象：単なるその日限りの出欠連絡、来年度の常勤契約日数の相談、他病院の退職挨拶などはすべて shouldWrite: false にしてください。',
      '以下のJSON形式のみを出力してください:',
      '{',
      '  "decisions": [',
      '    {',
      '      "rowNumber": 10,',
      '      "shouldWrite": true,',
      '      "targetSheet": "交渉注意医師",',
      '      "conditionType": "部分不可 または 絶対不可",',
      '      "ngClinic": "特定のNG拠点名があれば記載（例: 東品川）、なければ空文字",',
      '      "aiMemoSummary": "メモ列に追記する35文字以内の簡潔な条件要約"',
      '    }',
      '  ]',
      '}'
    ].join('\n');

    const activeModel = (typeof selectFastestWorkingModel_ === 'function')
      ? selectFastestWorkingModel_(apiKey)
      : 'gemini-flash-lite-latest';
    const url = 'https://' + 'generativelanguage.googleapis.com/v1beta/models/' + activeModel + ':generateContent?key=' + encodeURIComponent(apiKey);
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: systemPrompt }] },
        contents: [{ role: 'user', parts: [{ text: JSON.stringify({ items: targetItems }) }] }],
        generationConfig: { temperature: 0.0, maxOutputTokens: 4096, responseMimeType: 'application/json' }
      }),
      muteHttpExceptions: true
    });

    if (!res || res.getResponseCode() !== 200) return;

    const parsed = JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('').trim());
    const decisions = (parsed && Array.isArray(parsed.decisions)) ? parsed.decisions : [];
    const itemMap = new Map(targetItems.map(function(it) { return [it.rowNumber, it]; }));

    decisions.forEach(function(d) {
      const orig = itemMap.get(Number(d.rowNumber));
      if (!orig) return;

      if (d.shouldWrite && d.aiMemoSummary) {
        const stamp = orig.date ? String(orig.date).slice(5, 10) : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd');
        const entryText = '[' + stamp + '] ' + d.aiMemoSummary;

        // すでに『交渉しやすい医師』に登録されている医師ならE列に追記、未登録なら必ず『交渉注意医師』へ記録（新規行の無断追加を防止）
        const updatedInNeg = upsertNegotiableSheetRowIfExist_(ss, orig.doctor, entryText);
        if (!updatedInNeg) {
          const cSheet = ss.getSheetByName('交渉注意医師');
          if (cSheet) {
            upsertCautionDoctorRowProperly_(cSheet, null, {
              doctorName: orig.doctor,
              licenseNo: orig.licenseNo,
              dept: '小児科',
              conditionType: d.conditionType || '部分不可',
              ngArea: '関東',
              ngClinic: d.ngClinic || '',
              memoText: entryText
            });
          }
        }
        const newFlag = (orig.currentFlag ? orig.currentFlag + ' / ' : '') + 'AIメモ反映済';
        kbSheet.getRange(orig.rowNumber, 8).setValue(newFlag);
      } else {
        const newFlag = (orig.currentFlag ? orig.currentFlag + ' / ' : '') + '自動判定済';
        kbSheet.getRange(orig.rowNumber, 8).setValue(newFlag);
      }
    });
  } catch (e) {
    Logger.log('AIメモ自動同期スキップ: ' + e.message);
  }
}

/**
 * 『交渉しやすい医師』および『交渉注意医師』シートで列ズレ・誤分類になっている行を自動修復・移管する関数
 */
function repairNegotiableAndCautionSheetsNow(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const negSheet = ss.getSheetByName('交渉しやすい医師');
  const cautionSheet = ss.getSheetByName('交渉注意医師');

  // 1. 『交渉しやすい医師』の列ズレ・AI自動抽出行を検出し、『交渉注意医師』へ移管して削除
  if (negSheet && negSheet.getLastRow() >= 2) {
    const negVals = negSheet.getDataRange().getDisplayValues();
    const negRowsToDelete = [];

    for (let r = 1; r < negVals.length; r++) {
      const row = negVals[r];
      const colA = String(row[0] || '').trim();
      const colB = String(row[1] || '').trim();
      const colC = String(row[2] || '').trim();
      const colE = String(row[4] || '').trim();

      // A列が管理番号（数字）でB列が医師名、またはC列が「AI自動抽出」の行を検出
      if ((/^\d+$/.test(colA) && colB && !/^\d+$/.test(colB)) || colC.includes('AI自動抽出')) {
        const actualDocName = (/^\d+$/.test(colA) && colB) ? colB : colA;
        const memoTxt = colE || String(row[3] || '').trim();

        // 来年度の契約日数希望（常勤連絡）以外の有効な条件メモは『交渉注意医師』へ移管
        if (cautionSheet && actualDocName && memoTxt && !memoTxt.includes('来年度は勤務日数')) {
          upsertCautionDoctorRowProperly_(cautionSheet, null, {
            doctorName: actualDocName,
            licenseNo: '',
            dept: '小児科',
            conditionType: '部分不可',
            ngArea: '関東',
            ngClinic: '',
            memoText: memoTxt
          });
        }
        negRowsToDelete.push(r + 1);
      }
    }

    negRowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(rNum) {
      negSheet.deleteRow(rNum);
    });
    if (negRowsToDelete.length > 0) {
      Logger.log('✅ 『交渉しやすい医師』の誤分類・列ズレ行を修復・移管しました（' + negRowsToDelete.length + '件）');
    }
  }

  // 2. 『交渉注意医師』自体の列ズレ修復
  repairCautionSheetRowsNow(ss);
}

function repairCautionSheetRowsNow(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('交渉注意医師');
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastRow = sheet.getLastRow();
  const maxCols = Math.max(16, sheet.getLastColumn());
  const vals = sheet.getRange(1, 1, lastRow, maxCols).getDisplayValues();

  const validRowsMap = new Map();
  const rowsToDelete = [];
  const brokenEntries = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const colA = String(row[0] || '').trim();
    const colB = String(row[1] || '').trim();
    const colC = String(row[2] || '').trim();
    const colE = String(row[4] || '').trim();

    if ((/^\d+$/.test(colA) && colB) || colC.includes('AI自動抽出')) {
      brokenEntries.push({
        doctorName: colB || colA,
        memoText: colE || String(row[3] || '').trim()
      });
      rowsToDelete.push(r + 1);
    } else if (colA) {
      validRowsMap.set(normKeyForDoc_(colA), r + 1);
    }
  }

  if (brokenEntries.length === 0) return;

  brokenEntries.forEach(function(entry) {
    upsertCautionDoctorRowProperly_(sheet, validRowsMap, {
      doctorName: entry.doctorName,
      licenseNo: '',
      dept: '小児科',
      conditionType: '部分不可',
      ngArea: '関東',
      memoText: entry.memoText
    });
  });

  rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(rNum) {
    sheet.deleteRow(rNum);
  });

  Logger.log('✅ 『交渉注意医師』シートの列ズレ行を自動修復・統合しました（' + brokenEntries.length + '件）');
}

function upsertCautionDoctorRowProperly_(sheet, existingMapOpt, info) {
  const docKey = normKeyForDoc_(info.doctorName);
  if (!docKey) return;

  let existingMap = existingMapOpt;
  if (!existingMap) {
    existingMap = new Map();
    const lastR = sheet.getLastRow();
    if (lastR >= 2) {
      const aVals = sheet.getRange(2, 1, lastR - 1, 1).getDisplayValues();
      aVals.forEach(function(r, idx) {
        const k = normKeyForDoc_(r[0]);
        if (k) existingMap.set(k, idx + 2);
      });
    }
  }

  const knownClinics = [
    '東品川', '北葛西', '亀有', '代官山', '西日暮里', '錦糸町', '池袋', '紀尾井町', '吉祥寺',
    '新百合ヶ丘', '武蔵小杉', '横浜', '相模大野', '相模原', '天王町', '川崎', 'トレッサ横浜',
    '川口', '所沢', '草加松原', '越谷レイクタウン', '浦和美園', '朝霞台', '南浦和', '与野', '志木', '川越',
    '柏の葉', '流山おおたかの森', '新鎌ケ谷', '市川妙典', '八千代緑が丘', '千葉ニュータウン'
  ];
  let detectedClinic = info.ngClinic || '';
  if (!detectedClinic && info.memoText) {
    detectedClinic = knownClinics.find(function(c) { return info.memoText.includes(c); }) || '';
  }

  if (existingMap.has(docKey)) {
    const targetRow = existingMap.get(docKey);
    const currentP = String(sheet.getRange(targetRow, 16).getDisplayValue() || '').trim();
    if (info.memoText && !currentP.includes(info.memoText)) {
      const newP = currentP ? (currentP + ' / ' + info.memoText) : info.memoText;
      sheet.getRange(targetRow, 16).setValue(newP);
    }
    if (detectedClinic) {
      const fVal = String(sheet.getRange(targetRow, 6).getDisplayValue() || '').trim();
      if (!fVal) {
        try { sheet.getRange(targetRow, 6).setValue(detectedClinic); } catch (e) {}
      }
    }
    return;
  }

  const newRowIdx = sheet.getLastRow() + 1;
  const templateRowIdx = Math.min(5, Math.max(2, sheet.getLastRow()));

  if (templateRowIdx >= 2) {
    const templateRange = sheet.getRange(templateRowIdx, 1, 1, 16);
    const targetRange = sheet.getRange(newRowIdx, 1, 1, 16);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const validCond = (info.conditionType && info.conditionType.includes('絶対')) ? '絶対不可' : '部分不可';
  const validDept = (info.dept && info.dept.includes('内科')) ? '内科' : '小児科';
  const validArea = (info.ngArea && info.ngArea.includes('関西')) ? '関西' : '関東';
  const cleanLic = (info.licenseNo && /^\d{5,7}$/.test(String(info.licenseNo).trim())) ? String(info.licenseNo).trim() : '';

  const rowValues = [
    info.doctorName,
    cleanLic,
    validDept,
    validCond,
    validArea,
    detectedClinic,
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    info.memoText || ''
  ];

  sheet.getRange(newRowIdx, 1, 1, 16).setValues([rowValues]);
  existingMap.set(docKey, newRowIdx);
}

/**
 * 『交渉しやすい医師』にすでに登録されている医師の場合のみE列（AIメモ）を更新し、新規行追加は一切行わない関数
 */
function upsertNegotiableSheetRowIfExist_(ss, doctorName, entryText) {
  const sh = ss.getSheetByName('交渉しやすい医師');
  if (!sh) return false;
  const targetKey = normKeyForDoc_(doctorName);
  if (!targetKey) return false;

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const vals = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      const colA = String(vals[i][0] || '').trim();
      if (/^\d+$/.test(colA)) continue;
      const k = normKeyForDoc_(colA);
      if (k === targetKey) {
        const curE = String(vals[i][4] || '').trim();
        if (!curE.includes(entryText)) {
          sh.getRange(i + 2, 5).setValue(curE ? (curE + ' / ' + entryText) : entryText);
        }
        return true;
      }
    }
  }
  return false;
}

function loadOrInitCustomRuleSheet_(ss) {
  let sh = ss.getSheetByName(MTG_DOC_CONFIG.RULE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MTG_DOC_CONFIG.RULE_SHEET_NAME);
    const defaultRules = [
      ['カテゴリ', 'AIエージェントへの指示・ルール（ここを書き換えるだけで次回の会議から即反映されます）'],
      ['常勤医師の扱い', '『常勤依頼手当』シートに記載がない常勤医師は、延長・早出・休日追加・2診崩し移動のすべてにおいて打診提案しないこと（人間が判断するためAIは提案不要）。'],
      ['不可案の記載禁止（重要）', '【案1】【案2】【案3】には、実際に打診可能な有効アクションのみを優先順に記載すること。「前医の〇〇先生は通常常勤のため延長不可」等の不可理由や自明な文言は一切記載しないこと。'],
      ['手当の支給区分', '常勤勤務依頼手当（1,500円/h）は『常勤依頼手当』シート記載の医師のみ。他拠点からの拠点移動時は一律で「移動依頼手当10,000円」を支給（※SBUのみ0円）。'],
      ['2診目崩しと併行募集の表記', '他拠点からの移動案は、2診体制からの移動なら「【2診目崩し】拠点名（2診目：〇〇先生）」、2人とも移動可能なら「【2診目崩し】拠点名（2診目：①〇〇先生 / ②△△先生）」と両方書く。1診体制からの移動なら「【1診・併行募集】拠点名（1診：〇〇先生）」と明記する。'],
      ['属性表記の省略', '個別打診の候補医師に「（スポット）」「（紹介会社）」という記載は不要。シンプルに「〇〇先生へ個別打診（同曜日同枠〇回実績）」と記載する。'],
      ['エリア制限', '2診目崩しや午前・午後のシフト間移動は、定義書のドミナント優先度（S・A＝60分以内、特例B）に準拠し、関西⇔関東や移動不可領域（NG）の遠距離移動は絶対に提案しない。']
    ];
    sh.getRange(1, 1, defaultRules.length, 2).setValues(defaultRules);
    sh.getRange(1, 1, 1, 2).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 700);
    sh.setFrozenRows(1);
  }

  const vals = sh.getDataRange().getDisplayValues();
  const lines = [];
  for (let i = 1; i < vals.length; i++) {
    const cat = String(vals[i][0] || '').trim();
    const rule = String(vals[i][1] || '').trim();
    if (rule) lines.push('・【' + (cat || '追加ルール') + '】' + rule);
  }
  return lines.join('\n');
}