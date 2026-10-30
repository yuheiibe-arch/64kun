/**
 * ============================================================================
 * AI会議_2_AI分析.gs
 * （スポット/紹介会社表記カット・常勤延長不可の不要メモカット・時給表記修正版）
 * ============================================================================
 */

const CANDIDATE_GEMINI_MODELS = [
  'gemini-3.7-flash',
  'gemini-3.6-flash',
  'gemini-3.5-flash',
  'gemini-flash-lite-latest',
  'gemini-3.5-flash-lite',
  'gemini-3.1-flash-lite'
];

let CACHED_WORKING_MODEL_ = null;

function selectFastestWorkingModel_(apiKey) {
  if (CACHED_WORKING_MODEL_) return CACHED_WORKING_MODEL_;

  const baseUrl = 'https://' + 'generativelanguage.googleapis.com/v1beta/models/';
  const pingReqs = CANDIDATE_GEMINI_MODELS.map(m => ({
    url: baseUrl + m + ':generateContent?key=' + encodeURIComponent(apiKey),
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: '1' }] }],
      generationConfig: { maxOutputTokens: 2 }
    }),
    muteHttpExceptions: true
  }));

  const responses = UrlFetchApp.fetchAll(pingReqs);
  for (let i = 0; i < responses.length; i++) {
    if (responses[i].getResponseCode() === 200) {
      CACHED_WORKING_MODEL_ = CANDIDATE_GEMINI_MODELS[i];
      Logger.log('  🎯 事前疎通テスト完了！最優秀の稼働モデルを特定: 【 ' + CACHED_WORKING_MODEL_ + ' 】');
      return CACHED_WORKING_MODEL_;
    }
  }

  CACHED_WORKING_MODEL_ = 'gemini-flash-lite-latest';
  Logger.log('  ⚠️ フォールバックモデルを選択: 【 ' + CACHED_WORKING_MODEL_ + ' 】');
  return CACHED_WORKING_MODEL_;
}

/**
 * 「特別時給通常時給」や「特別時給特別時給」の重複表示を防ぐ整形関数
 */
function formatCleanWageLabel_(specialWageRaw, listedWageRaw) {
  const sp = String(specialWageRaw || '').trim();
  const li = String(listedWageRaw || '').trim();

  if (sp && sp !== '-' && sp !== 'なし' && sp !== 'FALSE' && sp !== '0') {
    if (sp === '通常時給' || sp.includes('通常')) {
      return '通常時給';
    }
    const cleanSp = sp.replace(/^特別時給[\s:：]*/g, '').replace(/通常時給/g, '').trim();
    return cleanSp ? `特別時給 ${cleanSp}` : '特別時給';
  }
  return li || '通常時給';
}

/**
 * 提案文から不要な「（スポット）」「（紹介会社）」表記や誤字を除去してスッキリさせる関数
 */
function cleanPlanTextFormatting_(text) {
  return String(text || '')
    .replace(/亀原本院/g, '亀有')
    .replace(/亀原/g, '亀有')
    .replace(/[（\(]\s*(?:スポット\s*[\/／・]\s*紹介会社|紹介会社\s*[\/／・]\s*スポット|スポット|紹介会社|スポット医師|紹介会社医師)\s*[）\)]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 現状・注意メモ（auditNote）から「前医の〇〇先生は通常常勤のため延長不可」等の自明な文言を除去する関数
 */
function cleanAuditNoteFormatting_(noteText) {
  const raw = String(noteText || '')
    .replace(/亀原本院/g, '亀有')
    .replace(/亀原/g, '亀有')
    .trim();
  if (!raw || raw === '-' || raw === 'なし' || raw === '特になし') return '';

  // 句点「。」やスラッシュ「/」で分割し、自明な文（「通常常勤のため延長不可」「調整余地なし」等）を除外
  const sentences = raw.split(/[。\/／\n]+/).map(s => s.trim()).filter(Boolean);
  const meaningful = sentences.filter(s => {
    if (/(通常常勤|通常勤務).*?(延長不可|早出不可|移動不可|打診不可|不可)/.test(s)) return false;
    if (/^(前医|後医|前後医).*?(延長不可|早出不可|調整余地なし|不在|なし|不可)$/.test(s)) return false;
    if (/^(延長不可|早出不可|調整余地なし|特になし|該当なし|問題なし|注意点なし)$/.test(s)) return false;
    return true;
  });

  return meaningful.join(' / ');
}

/**
 * plan1〜plan3 から不可理由を排除し、有効な打診案のみを plan1 から順に繰り上げる関数
 */
function cleanAndPromoteValidPlans_(aiPlan) {
  const rawPlans = [aiPlan.plan1, aiPlan.plan2, aiPlan.plan3]
    .map(p => cleanPlanTextFormatting_(p))
    .filter(Boolean);

  const validPlans = [];
  const nonActionRegex = /(延長不可|早出不可|移動不可|打診不可|調整不可|相談不可|対象外|余地なし|該当なし|候補なし|不在のため|12連勤NG)/;

  rawPlans.forEach(p => {
    const hasPositiveAction = /(へ延長打診|へ早出打診|へ移動打診|へ個別打診|時給UP募集を立て)/.test(p);
    if (nonActionRegex.test(p) && !hasPositiveAction) {
      // 不可理由はplanに入れない
    } else if (p === '-' || p === 'なし' || p === '特になし') {
      // スキップ
    } else {
      validPlans.push(p);
    }
  });

  const cleanNote = cleanAuditNoteFormatting_(aiPlan.auditNote);

  return {
    plan1: validPlans[0] || '',
    plan2: validPlans[1] || '',
    plan3: validPlans[2] || '',
    auditNote: cleanNote || '-'
  };
}

function callGeminiAgentInBatches_(rawContext, customRulesText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('スクリプトプロパティに GEMINI_API_KEY が見つかりません。');

  const activeModel = selectFastestWorkingModel_(apiKey);

  const allSlots = rawContext.slotsContext || [];
  const batchSize = Math.max(35, Math.ceil(allSlots.length / 3));
  const chunks = [];

  for (let i = 0; i < allSlots.length; i += batchSize) {
    chunks.push(allSlots.slice(i, i + batchSize));
  }

  Logger.log('  ↳ 確定モデル【' + activeModel + '】で全 ' + chunks.length + ' バッチ＋総括（計' + (chunks.length + 1) + '通信）を一斉並列分析中...');

  const slotSystemPrompt = buildSlotChunkPrompt_(customRulesText);
  const summarySystemPrompt = buildSummaryMetaPrompt_();

  const requestsPayloadList = chunks.map(chunk => ({
    systemPrompt: slotSystemPrompt,
    userObj: { slots: chunk }
  }));

  requestsPayloadList.push({
    systemPrompt: summarySystemPrompt,
    userObj: {
      periodStart: rawContext.periodStart,
      periodEnd: rawContext.periodEnd,
      unfilledCount: rawContext.unfilledCount,
      targetBundles: rawContext.targetBundles || [],
      recentFlaggedKb: rawContext.recentFlaggedKb || [],
      customRules: customRulesText
    }
  });

  const results = invokeVerifiedModelParallel_(apiKey, activeModel, requestsPayloadList);

  const aiPlanMap = new Map();
  for (let i = 0; i < chunks.length; i++) {
    const parsed = results[i];
    if (parsed && Array.isArray(parsed.slotPlans)) {
      parsed.slotPlans.forEach(p => {
        if (p && p.slotId != null) {
          aiPlanMap.set(Number(p.slotId), p);
        }
      });
    }
  }

  const allSlotPlans = allSlots.map((origSlot) => {
    const aiPlan = aiPlanMap.get(Number(origSlot.slotId)) || {};

    const baseClinicName = String(origSlot.clinic || '')
      .replace(/[\(（]小児科[\)）]|[\(（]内科[\)）]|小児科|内科/g, '')
      .trim();
    const exactClinicAndDept = `【${baseClinicName}（${origSlot.dept}）】`;

    const wageLabel = formatCleanWageLabel_(origSlot.specialWage, origSlot.listedWage);
    const exactAbsenceTime = `${origSlot.absenceTime}（${wageLabel}）`;

    const cleaned = cleanAndPromoteValidPlans_(aiPlan);

    return {
      slotId: origSlot.slotId,
      date: origSlot.date,
      clinicAndDept: exactClinicAndDept,
      absenceTime: exactAbsenceTime,
      assignee: origSlot.assignee || '',
      plan1: cleaned.plan1,
      plan2: cleaned.plan2,
      plan3: cleaned.plan3,
      auditNote: cleaned.auditNote
    };
  });

  const summaryPart = results[chunks.length] || {};
  const cleanBundles = (summaryPart.bundleRecommendations || []).map(b => ({
    doctor: cleanPlanTextFormatting_(b.doctor || ''),
    targetSlots: b.targetSlots || '',
    reasonAndNote: b.reasonAndNote || ''
  }));

  return {
    executiveSummary: summaryPart.executiveSummary || '',
    slotPlans: allSlotPlans,
    bundleRecommendations: cleanBundles,
    announcementProposal: buildOfficialAnnouncementProposal_(rawContext.slotsContext),
    masterUpdateQuestions: summaryPart.masterUpdateQuestions || []
  };
}

function invokeVerifiedModelParallel_(apiKey, primaryModel, requestsPayloadList) {
  const baseUrl = 'https://' + 'generativelanguage.googleapis.com/v1beta/models/';
  const results = new Array(requestsPayloadList.length).fill(null);

  const safeParse = (res) => {
    const resJson = JSON.parse(res.getContentText());
    const rawText = resJson.candidates[0].content.parts.map(p => p.text || '').join('').trim();
    try {
      return JSON.parse(rawText);
    } catch (e) {
      const lastCloseBrace = rawText.lastIndexOf('}');
      if (lastCloseBrace > 0) {
        try {
          return JSON.parse(rawText.slice(0, lastCloseBrace + 1) + ']}');
        } catch (e2) {}
      }
      return { slotPlans: [] };
    }
  };

  const makeReqs = (modelName, indices) => {
    const url = baseUrl + modelName + ':generateContent?key=' + encodeURIComponent(apiKey);
    return indices.map(idx => {
      const item = requestsPayloadList[idx];
      return {
        url: url,
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          systemInstruction: { parts: [{ text: item.systemPrompt }] },
          contents: [{ role: 'user', parts: [{ text: JSON.stringify(item.userObj) }] }],
          generationConfig: {
            temperature: 0.0,
            maxOutputTokens: 16384,
            responseMimeType: 'application/json'
          }
        }),
        muteHttpExceptions: true
      };
    });
  };

  const allIndices = requestsPayloadList.map((_, idx) => idx);
  const responses = UrlFetchApp.fetchAll(makeReqs(primaryModel, allIndices));
  const failedIndices = [];

  for (let i = 0; i < responses.length; i++) {
    const code = responses[i].getResponseCode();
    if (code === 200) {
      Logger.log('    ✅ バッチ ' + (i + 1) + '/' + requestsPayloadList.length + ' 完了 (' + primaryModel + ')');
      results[i] = safeParse(responses[i]);
    } else {
      Logger.log('    ⚠️ バッチ ' + (i + 1) + ' (' + code + ') -> 高速バックアップ(gemini-flash-lite-latest)で即回収します');
      failedIndices.push(i);
    }
  }

  if (failedIndices.length > 0) {
    Utilities.sleep(1500);
    const retryModel = 'gemini-flash-lite-latest';
    const retryResponses = UrlFetchApp.fetchAll(makeReqs(retryModel, failedIndices));
    for (let k = 0; k < failedIndices.length; k++) {
      const origIdx = failedIndices[k];
      if (retryResponses[k].getResponseCode() === 200) {
        Logger.log('    ✅ バッチ ' + (origIdx + 1) + '/' + requestsPayloadList.length + ' 回収完了 (' + retryModel + ')');
        results[origIdx] = safeParse(retryResponses[k]);
      } else {
        results[origIdx] = { slotPlans: [] };
      }
    }
  }

  return results;
}

function buildSlotChunkPrompt_(customRulesText) {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（攻め・守り統合）」です。',
    '思考や前置きを省き、直ちに以下のルールで各不在枠の調整案（JSON）のみを簡潔に出力してください。',
    '🚫重要1：入力データにあるクリニック名（拠点名）や医師名は1文字たりとも勝手に書き換えたり創作しないでください。',
    '🚫重要2：plan1 / plan2 / plan3 には「実際に打診できる前向きなアクション」のみを入れてください。不可理由は絶対に入れないこと。',
    '🚫重要3：医師名の後ろに「（スポット）」「（紹介会社）」という属性表記は一切付けないでください（シンプルに「〇〇先生へ個別打診（同曜日同枠〇回実績）」と書くこと）。',
    '',
    '【最優先カスタムルール】',
    customRulesText,
    '',
    '【具体的な表記ルール（厳守）】',
    '1. 前後枠の延長・早出:',
    '   - prevDocs / nextDocs が非常勤医師の場合は「前医の〇〇先生へ延長打診」と提案。',
    '   - 常勤医師の場合は、『常勤依頼手当』対象かつ「所定休出（★移動相談可）」と明記されている場合のみ「前医の〇〇先生（常勤・所定休出）へ延長打診（常勤追加勤務手当1,500円/h）」と提案する。通常の常勤医師には絶対に延長・早出を提案しないこと。',
    '2. 他拠点からの移動提案（twoDocBreakdownCandidates / oneDocParallelCandidates）:',
    '   - 2診体制からの移動（twoDocBreakdownCandidates）は必ず頭に「【2診目崩し】」をつける。',
    '     ・1名の場合: 「【2診目崩し】所沢（2診目：井上信明先生）へ移動打診（移動依頼手当10,000円）」',
    '     ・2名いる場合: 「【2診目崩し】所沢（2診目：①井上信明先生 / ②小池卓也先生）のどちらか1名へ移動打診（移動依頼手当10,000円）」',
    '   - 1診体制からの併行募集移動（oneDocParallelCandidates）は必ず頭に「【1診・併行募集】」をつける。',
    '     ・例: 「【1診・併行募集】八千代緑が丘（1診：伊藤直樹先生）に時給UP募集を立て、埋まり次第移動打診（移動依頼手当10,000円）」',
    '3. 個別打診（kinmukunCandidates）:',
    '   - 「〇〇先生へ個別打診（同曜日同枠〇回実績）」と簡潔に書く（※「（スポット）」「（紹介会社）」は記載不要）。',
    '4. 現状・注意メモ（auditNote）:',
    '   - 🚫禁止事項：「前医の〇〇先生は通常常勤のため延長不可」「後医は早出不可」「調整余地なし」といった当たり前の不可理由は一切書かないでください！',
    '   - currentMemo（作業メモ）の交渉中ステータスや、提案した候補医師に特別な注意点（NG条件・直近断り履歴など）がある場合のみ簡潔に書き、特になければ "-" を出力してください。',
    '',
    '以下のJSON形式のみを出力してください（slotIdは入力と必ず一致させること）:',
    '{',
    '  "slotPlans": [',
    '    {',
    '      "slotId": 1,',
    '      "plan1": "第1優先アクション（スポット/紹介会社の表記なし）",',
    '      "plan2": "第2優先アクション",',
    '      "plan3": "第3優先アクション",',
    '      "auditNote": "作業メモや候補医師の特別な注意点がある場合のみ記載（なければ - ）"',
    '    }',
    '  ]',
    '}'
  ].join('\n');
}

function buildSummaryMetaPrompt_() {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（総務・全体統括）」です。',
    '以下のルールに従って、①全体総括、②複数日程おまとめ打診リスト（上位12名まで）、③マスタ更新確認を作成してください。',
    '🚫重要：医師名の後ろに「（スポット）」「（紹介会社）」の表記は付けないでください。',
    '',
    '【重要：複数日程おまとめ打診リスト（targetSlots）の整形ルール】',
    '必ず「■ 拠点名」ごとに分類し、改行（\\n）を使って箇条書きでスッキリ見やすく整形してください。',
    '',
    '以下のJSON形式のみを出力してください:',
    '{',
    '  "executiveSummary": "対象期間の全体傾向と優先対応ポイント（2文以内）",',
    '  "bundleRecommendations": [',
    '    {',
    '      "doctor": "〇〇 先生",',
    '      "targetSlots": "■ 拠点名\\n・MM/DD(曜) 時間帯",',
    '      "reasonAndNote": "実績が多い拠点や交渉時の注意点（簡潔に）"',
    '    }',
    '  ],',
    '  "masterUpdateQuestions": []',
    '}'
  ].join('\n');
}

function buildOfficialAnnouncementProposal_(slotsContext) {
  const pedSlots = (slotsContext || []).filter(s => s.dept === '小児科');
  if (pedSlots.length === 0) {
    return { recommendedDateReason: '対象の小児科不在枠がありません。', templateText: '' };
  }

  const byDate = {};
  pedSlots.forEach(s => {
    if (!byDate[s.date]) byDate[s.date] = [];
    byDate[s.date].push(s);
  });

  const sortedDates = Object.keys(byDate).sort((a, b) => byDate[b].length - byDate[a].length || a.localeCompare(b));
  const bestDate = sortedDates[0];
  const targetList = byDate[bestDate];

  const areaGroups = { '東京': [], '神奈川': [], '埼玉': [], '千葉': [], '茨城': [], '関西': [] };
  targetList.forEach(u => {
    const cleanClinic = String(u.clinic || '').replace(/小児科|[\(（].*?[\)）]/g, '').trim();
    const absTime = String(u.absenceTime || '').trim().replace('-', '～');
    const baseWage = Number(String(u.listedWage || '13000').replace(/[^0-9]/g, '')) || 13000;
    const specialHourly = baseWage + 2000;

    let hours = 3;
    const tm = absTime.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
    if (tm) {
      hours = (Number(tm[3]) + Number(tm[4]) / 60) - (Number(tm[1]) + Number(tm[2]) / 60);
    }
    const dailyPay = Math.round(specialHourly * hours).toLocaleString();
    const area = u.area || '東京';
    if (!areaGroups[area]) areaGroups[area] = [];
    areaGroups[area].push('・【' + cleanClinic + '】' + absTime + ' 日給' + dailyPay + '円');
  });

  const areaLines = [];
  Object.keys(areaGroups).forEach(a => {
    if (areaGroups[a].length > 0) areaLines.push(a + '\n' + areaGroups[a].join('\n'));
  });

  const templateText = [
    'タイトル：※急募※ ' + bestDate + ' 【小児科】にてご勤務いただける先生を探しております！',
    '',
    '（こちらはキャップスドクター365へご登録いただいている先生方へお送りしております）',
    '先生方',
    '平素より大変お世話になっております。ドクターサポート事務局です。',
    '表題の件につきまして、' + bestDate + ' 複数拠点での医師不在の状況となり、緊急募集を行っております。',
    'ご都合の良い小児科の先生がいらっしゃいましたら、ご勤務のご検討をいただきたく存じます。',
    '============================================',
    '＜急募＞特別時給（基本時給＋2,000円/h 換算済）',
    '◆' + bestDate,
    areaLines.join('\n\n'),
    '============================================',
    '※紹介会社経由でご勤務いただいております先生は紹介会社様経由でのご応募をお願い申し上げます。',
    'ドクターサポート事務局（050-3187-6647 / doctor-support@caps365.jp）'
  ].join('\n');

  return {
    recommendedDateReason: '週1回以下の配信ルール（オオカミ少年防止）に基づき、小児科未充足が最も集中している ' + bestDate + '（計' + targetList.length + '枠）を今週唯一の配信対象日として選定しました。',
    templateText: templateText
  };
}