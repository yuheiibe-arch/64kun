/**
 * ============================================================================
 * AI会議_2a_モデル通信と整形.gs
 * （一般連絡ノイズ100%遮断・シフト調整の承諾/辞退のみ表示・超高速版）
 * ============================================================================
 */

const CANDIDATE_GEMINI_MODELS = [
  'gemini-flash-lite-latest',
  'gemini-2.5-flash-lite',
  'gemini-2.0-flash-lite'
];

let CACHED_WORKING_MODEL_ = null;

function selectFastestWorkingModel_(apiKey) {
  CACHED_WORKING_MODEL_ = 'gemini-flash-lite-latest';
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

function cleanAuditNoteFormatting_(noteText) {
  const raw = String(noteText || '')
    .replace(/亀原本院/g, '亀有')
    .replace(/亀原/g, '亀有')
    .trim();
  if (!raw || raw === '-' || raw === 'なし' || raw === '特になし') return '';

  const sentences = raw.split(/(?:。|\s+[\/／]\s+|\n)+/).map(function(s) { return s.trim(); }).filter(Boolean);
  const meaningful = sentences.filter(function(s) {
    if (s.includes('一般連絡')) return false; // 一般連絡は絶対に出さない
    if (/(通常常勤|通常勤務).*?(延長不可|早出不可|移動不可|打診不可|不可)/.test(s)) return false;
    if (/(中抜け|空き時間).*?(調整困難|不可|対象外|除外)/.test(s)) return false;
    if (/(除外|対象外|相談しない|候補外)/.test(s)) return false;
    if (/^(前医|後医|前後医|前後枠).*?(延長不可|早出不可|調整余地なし|医師なし|不在|なし|不可)/.test(s)) return false;
    if (/^(延長不可|早出不可|調整余地なし|前後枠なし|特になし|該当なし|問題なし|注意点なし|他拠点からの移動候補なし)$/.test(s)) return false;
    return true;
  });

  return meaningful.join(' / ');
}

function buildDeterministicSlotPlanFast_(slot) {
  const actions = [];
  const auditNotes = [];

  if (slot.currentMemo && String(slot.currentMemo).trim() && slot.currentMemo !== '-') {
    auditNotes.push('【進行中メモ】' + String(slot.currentMemo).trim());
  }

  // 1. 第1優先：自拠点の前後医による「延長・早出」（※SBUは対象外）
  const formatAdjAction = function(docsStr, actionLabel) {
    if (!docsStr || docsStr === 'なし') return;
    docsStr.split(' / ').map(function(s) { return s.trim(); }).filter(Boolean).forEach(function(dTok) {
      if (dTok.includes('SBU')) return;
      const cleanName = dTok.replace(/[\(（].*?[\)）]/g, '').trim();
      const isFtHoliday = !dTok.includes('非常勤') && dTok.includes('所定休出');
      if (isFtHoliday) {
        actions.push((actionLabel === '延長' ? '前医の' : '後医の') + cleanName + '先生（常勤・所定休出）へ' + actionLabel + '打診（常勤追加勤務手当1,500円/h）');
      } else {
        actions.push((actionLabel === '延長' ? '前医の' : '後医の') + cleanName + '先生（非常勤）へ' + actionLabel + '打診');
      }
    });
  };
  formatAdjAction(slot.prevDocs, '延長');
  formatAdjAction(slot.nextDocs, '早出');

  // 2. 第2優先：ドミナント定義書（優先度S・A・特例B）に基づく他拠点からの「2診目崩し」
  (slot.twoDocBreakdownCandidates || []).forEach(function(cand) {
    if (actions.length >= 3) return;
    const docs = cand.movableDoctors || [];
    if (docs.length === 0) return;

    const allSbu = docs.every(function(d) { return d.includes('SBU'); });
    const allowLabel = allSbu ? '（移動手当0円）' : '（移動依頼手当10,000円）';

    const formatMovDoc = function(d) {
      const base = d.replace(/\[.*?\]/g, '').trim();
      if (d.includes('SBU')) return base + '先生[★SBU]';
      if (!d.includes('非常勤') && d.includes('所定休出')) return base + '先生[常勤・所定休出]';
      return base + '先生';
    };

    if (docs.length === 1) {
      actions.push('【2診目崩し】' + cand.clinic + '（2診目：' + formatMovDoc(docs[0]) + '）へ移動打診' + allowLabel);
    } else {
      const twoNames = '①' + formatMovDoc(docs[0]) + ' / ②' + formatMovDoc(docs[1]);
      actions.push('【2診目崩し】' + cand.clinic + '（2診目：' + twoNames + '）のどちらか1名へ移動打診' + allowLabel);
    }
  });

  // 3. 第3優先：確定シフト突合済みの「個別打診」
  (slot.kinmukunCandidates || []).forEach(function(sp) {
    if (actions.length >= 3) return;

    let recLabel = sp.trackRecordLabel || '';
    if (!recLabel) {
      const sCount = Number(String(sp.sameDowSlotCount || '0').replace(/[^0-9]/g, '')) || 0;
      const tCount = Number(String(sp.totalCount || '0').replace(/[^0-9]/g, '')) || 0;
      if (sCount === 0 && tCount === 0) return;
      recLabel = sCount > 0 ? ('同曜日同枠' + sCount + '回実績') : ('同拠点通算' + tCount + '回実績');
    }

    actions.push(sp.doctor + '先生へ個別打診（' + recLabel + '）');

    if (sp.daySchedule && sp.daySchedule.startsWith('同日確定枠:')) {
      auditNotes.push(sp.doctor + '先生は' + sp.daySchedule.replace('同日確定枠:', '同日別枠（') + '）あり');
    }

    // ★シフト調整の履歴（承諾・辞退）のみを端的に表示し、一般連絡（12/31の空き連絡や年末調整など）は一切出さない
    if (sp.latestReplyDetail) {
      auditNotes.push(sp.doctor + '先生は' + sp.latestReplyDetail);
    } else if (sp.masterNote && !sp.masterNote.includes('一般連絡')) {
      const dateMatch = sp.masterNote.match(/直近(\d{1,2}\/\d{1,2})打診\(([^:：)]+)/);
      if (dateMatch && dateMatch[2].includes('シフト調整')) {
        const kindLabel = dateMatch[2].includes('辞退') ? 'の別枠打診：辞退履歴あり' : 'の別枠打診：承諾実績あり';
        auditNotes.push(sp.doctor + '先生は' + dateMatch[1] + kindLabel);
      }
    }
  });

  // 4. まだ3案に満たない場合は「1診・併行募集」を追加
  (slot.oneDocParallelCandidates || []).forEach(function(cand) {
    if (actions.length >= 3) return;
    const docs = cand.movableDoctors || [];
    if (docs.length === 0) return;
    const base = docs[0].replace(/\[.*?\]/g, '').trim();
    actions.push('【1診・併行募集】' + cand.clinic + '（1診：' + base + '先生）に時給UP募集を立て、埋まり次第移動打診（移動依頼手当10,000円）');
  });

  const baseClinicName = String(slot.clinic || '')
    .replace(/[\(（]小児科[\)）]|[\(（]内科[\)）]|小児科|内科/g, '')
    .trim();
  const showDept = (baseClinicName.includes('亀有') || baseClinicName.includes('北葛西') || slot.dept === '内科');
  const exactClinicAndDept = showDept ? ('【' + baseClinicName + '（' + slot.dept + '）】') : ('【' + baseClinicName + '】');

  const wageLabel = formatCleanWageLabel_(slot.specialWage, slot.listedWage);
  const exactAbsenceTime = slot.absenceTime + '（' + wageLabel + '）';

  return {
    slotId: slot.slotId,
    date: slot.date,
    clinicAndDept: exactClinicAndDept,
    absenceTime: exactAbsenceTime,
    assignee: slot.assignee || '',
    plan1: cleanPlanTextFormatting_(actions[0] || ''),
    plan2: cleanPlanTextFormatting_(actions[1] || ''),
    plan3: cleanPlanTextFormatting_(actions[2] || ''),
    auditNote: cleanAuditNoteFormatting_(auditNotes.join(' / ')) || '-'
  };
}

function callGeminiAgentInBatches_(rawContext, customRulesText) {
  const t0 = Date.now();
  const allSlots = rawContext.slotsContext || [];

  const allSlotPlans = allSlots.map(function(origSlot) {
    return buildDeterministicSlotPlanFast_(origSlot);
  });
  Logger.log('  ⚡ 全 ' + allSlotPlans.length + ' 枠のプラン構築完了（' + (Date.now() - t0) + 'ms）');

  const executiveSummary = '対象期間（' + rawContext.periodStart + '〜' + rawContext.periodEnd + '）の医師不在枠は計' + rawContext.unfilledCount + '件です。自拠点の前後医による延長・早出、およびドミナント近隣拠点からの2診目崩しを最優先で打診してください。';

  return {
    executiveSummary: executiveSummary,
    slotPlans: allSlotPlans,
    bundleRecommendations: [],
    announcementProposal: buildOfficialAnnouncementProposal_(rawContext.slotsContext),
    masterUpdateQuestions: []
  };
}