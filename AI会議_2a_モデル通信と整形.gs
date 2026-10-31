/**
 * ============================================================================
 * AI会議_2a_モデル通信と整形.gs
 * （モデル自動選定・並列バッチ通信・不要な内部思考メモや属性表記の完全除去）
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
  const pingReqs = CANDIDATE_GEMINI_MODELS.map(function(m) {
    return {
      url: baseUrl + m + ':generateContent?key=' + encodeURIComponent(apiKey),
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: '1' }] }],
        generationConfig: { maxOutputTokens: 2 }
      }),
      muteHttpExceptions: true
    };
  });

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

function formatCleanWageLabel_(specialWageRaw, listedWageRaw) {
  const sp = String(specialWageRaw || '').trim();
  const li = String(listedWageRaw || '').trim();

  if (sp && sp !== '-' && sp !== 'なし' && sp !== 'FALSE' && sp !== '0') {
    if (sp === '通常時給' || sp.includes('通常')) {
      return '通常時給';
    }
    const cleanSp = sp.replace(/^特別時給[\s:：]*/g, '').replace(/通常時給/g, '').trim();
    return cleanSp ? ('特別時給 ' + cleanSp) : '特別時給';
  }
  return li || '通常時給';
}

function cleanPlanTextFormatting_(text) {
  return String(text || '')
    .replace(/亀原本院/g, '亀有')
    .replace(/亀原/g, '亀有')
    .replace(/[（\(]\s*(?:スポット\s*[\/／・]\s*紹介会社|紹介会社\s*[\/／・]\s*スポット|スポット|紹介会社|スポット医師|紹介会社医師)\s*[）\)]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 現状・注意メモ（auditNote）から内部思考（不可理由・除外理由・自明な文言）を完全除去する関数
 */
function cleanAuditNoteFormatting_(noteText) {
  const raw = String(noteText || '')
    .replace(/亀原本院/g, '亀有')
    .replace(/亀原/g, '亀有')
    .trim();
  if (!raw || raw === '-' || raw === 'なし' || raw === '特になし') return '';

  const sentences = raw.split(/(?:。|\s+[\/／]\s+|(?<!\d)[\/／](?!\d)|\n)+/).map(function(s) { return s.trim(); }).filter(Boolean);
  const meaningful = sentences.filter(function(s) {
    if (/(通常常勤|通常勤務).*?(延長不可|早出不可|移動不可|打診不可|不可)/.test(s)) return false;
    if (/(中抜け|空き時間).*?(調整困難|不可|対象外|除外)/.test(s)) return false;
    if (/(除外|対象外|相談しない|候補外)/.test(s)) return false;
    if (/^(前医|後医|前後医|前後枠).*?(延長不可|早出不可|調整余地なし|医師なし|不在|なし|不可)/.test(s)) return false;
    if (/^(延長不可|早出不可|調整余地なし|前後枠なし|特になし|該当なし|問題なし|注意点なし|他拠点からの移動候補なし)$/.test(s)) return false;
    return true;
  });

  return meaningful.join(' / ');
}

function cleanAndPromoteValidPlans_(aiPlan) {
  const rawPlans = [aiPlan.plan1, aiPlan.plan2, aiPlan.plan3]
    .map(function(p) { return cleanPlanTextFormatting_(p); })
    .filter(Boolean);

  const validPlans = [];
  const nonActionRegex = /(延長不可|早出不可|移動不可|打診不可|調整不可|相談不可|対象外|余地なし|該当なし|候補なし|不在のため|12連勤NG)/;

  rawPlans.forEach(function(p) {
    const hasPositiveAction = /(へ延長打診|へ早出打診|へ移動打診|へ個別打診|時給UP募集を立て)/.test(p);
    if (nonActionRegex.test(p) && !hasPositiveAction) {
      return;
    }
    if (p === '-' || p === 'なし' || p === '特になし') {
      return;
    }
    validPlans.push(p);
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
  const batchSize = Math.max(30, Math.ceil(allSlots.length / 3));
  const chunks = [];

  for (let i = 0; i < allSlots.length; i += batchSize) {
    chunks.push(allSlots.slice(i, i + batchSize));
  }

  Logger.log('  ↳ 確定モデル【' + activeModel + '】で全 ' + chunks.length + ' バッチ＋総括（計' + (chunks.length + 1) + '通信）を一斉並列分析中...');

  const slotSystemPrompt = buildSlotChunkPrompt_(customRulesText);
  const summarySystemPrompt = buildSummaryMetaPrompt_();

  const requestsPayloadList = chunks.map(function(chunk) {
    return {
      systemPrompt: slotSystemPrompt,
      userObj: { slots: chunk }
    };
  });

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
      parsed.slotPlans.forEach(function(p) {
        if (p && p.slotId != null) {
          aiPlanMap.set(Number(p.slotId), p);
        }
      });
    }
  }

  const allSlotPlans = allSlots.map(function(origSlot) {
    const aiPlan = aiPlanMap.get(Number(origSlot.slotId)) || {};

    const baseClinicName = String(origSlot.clinic || '')
      .replace(/[\(（]小児科[\)）]|[\(（]内科[\)）]|小児科|内科/g, '')
      .trim();
    const showDept = (baseClinicName.includes('亀有') || baseClinicName.includes('北葛西') || origSlot.dept === '内科');
    const exactClinicAndDept = showDept ? ('【' + baseClinicName + '（' + origSlot.dept + '）】') : ('【' + baseClinicName + '】');

    const wageLabel = formatCleanWageLabel_(origSlot.specialWage, origSlot.listedWage);
    const exactAbsenceTime = origSlot.absenceTime + '（' + wageLabel + '）';

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
  const cleanBundles = (summaryPart.bundleRecommendations || []).map(function(b) {
    return {
      doctor: cleanPlanTextFormatting_(b.doctor || ''),
      targetSlots: b.targetSlots || '',
      reasonAndNote: b.reasonAndNote || ''
    };
  });

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

  const safeParse = function(res) {
    const resJson = JSON.parse(res.getContentText());
    const rawText = resJson.candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('').trim();
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

  const makeReqs = function(modelName, indices) {
    const url = baseUrl + modelName + ':generateContent?key=' + encodeURIComponent(apiKey);
    return indices.map(function(idx) {
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

  const allIndices = requestsPayloadList.map(function(_, idx) { return idx; });
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