/**
 * ============================================================================
 * AI会議_2a_モデル通信と整形.gs
 * （「💡現状・注意：」の改行箇条書き整形・異体字［﨑・髙・齋等］抽出漏れ防止版）
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
    .replace(/（複数枠で確定時は自動でおまとめ打診化）/g, '')
    .replace(/\(複数枠で確定時は自動でおまとめ打診化\)/g, '')
    .trim();
  if (!raw || raw === '-' || raw === 'なし' || raw === '特になし') return '';

  const sentences = raw
    .split(/(?:。|\s+[\/／]\s+|(?<!\d)[\/／](?!\d)|\n)+/)
    .map(function(s) { return s.replace(/^[・\s]+/, '').trim(); })
    .filter(Boolean);

  const meaningful = sentences.filter(function(s) {
    if (s.includes('一般連絡')) return false;
    if (/(通常常勤|通常勤務).*?(延長不可|早出不可|移動不可|打診不可|不可)/.test(s)) return false;
    if (/(中抜け|空き時間).*?(調整困難|不可|対象外|除外)/.test(s)) return false;
    if (/(除外|対象外|相談しない|候補外)/.test(s)) return false;
    if (/^(前医|後医|前後医|前後枠).*?(延長不可|早出不可|調整余地なし|医師なし|不在|なし|不可)/.test(s)) return false;
    if (/^(延長不可|早出不可|調整余地なし|前後枠なし|特になし|該当なし|問題なし|注意点なし|他拠点からの移動候補なし)$/.test(s)) return false;
    return true;
  });

  if (meaningful.length === 0) return '';
  return meaningful.map(function(s) { return '・' + s; }).join('\n');
}

function buildMasterNameCorrectionMap2a_(allSlots) {
  const officialNames = new Set();

  (allSlots || []).forEach(function(slot) {
    [slot.prevDocs, slot.nextDocs, slot.backupDocs].forEach(function(str) {
      String(str || '').split(/[\/／,、\s]+/).forEach(function(tok) {
        const clean = tok.replace(/[\(（\[【｛].*$/g, '').replace(/先生$/g, '').trim();
        if (clean && clean.length >= 2 && clean !== 'なし' && clean !== '未充足') officialNames.add(clean);
      });
    });
    (slot.twoDocBreakdownCandidates || []).concat(slot.oneDocParallelCandidates || []).forEach(function(c) {
      (c.movableDoctors || []).forEach(function(d) {
        const clean = String(d || '').replace(/[\(（\[【｛].*$/g, '').replace(/先生$/g, '').trim();
        if (clean && clean.length >= 2) officialNames.add(clean);
      });
    });
    (slot.kinmukunCandidates || []).forEach(function(sp) {
      const clean = String(sp.doctor || '').replace(/[\(（\[【｛].*$/g, '').replace(/先生$/g, '').trim();
      if (clean && clean.length >= 2) officialNames.add(clean);
    });
  });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof loadDoctorSearchList64_ === 'function') {
      loadDoctorSearchList64_(ss).forEach(function(d) {
        if (d.name && d.name.length >= 2) officialNames.add(d.name);
      });
    }
  } catch (e) {}

  return Array.from(officialNames);
}

function toPhoneticHomophoneKey2a_(nameStr) {
  return String(nameStr || '')
    .normalize('NFKC')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[巧匠]/g, '匠')
    .replace(/[滝瀧泷]/g, '瀧')
    .replace(/[高髙]/g, '高')
    .replace(/[崎﨑嵜]/g, '崎')
    .replace(/[辺邊邉]/g, '辺')
    .replace(/[斎齋齊斉]/g, '斎')
    .replace(/[沢澤]/g, '沢')
    .replace(/[広廣]/g, '広')
    .replace(/[浜濵濱]/g, '浜')
    .replace(/[真眞]/g, '真')
    .replace(/[恵惠]/g, '恵')
    .replace(/[島嶋嶌]/g, '島')
    .replace(/[柳栁]/g, '柳')
    .replace(/[桜櫻]/g, '桜')
    .replace(/[国國]/g, '国')
    .replace(/[徳德]/g, '徳')
    .replace(/[朗郎]/g, '郎')
    .replace(/[介助]/g, '介')
    .replace(/[悠祐裕]/g, '祐')
    .replace(/[佳圭]/g, '佳')
    .replace(/[智知]/g, '智');
}

function correctDoctorNamesInMemoText_(memoText, officialNameList) {
  let s = String(memoText || '').trim();
  if (!s || !officialNameList || officialNameList.length === 0) return s;

  return s.replace(/([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})先生/g, function(fullMatch, rawDocName) {
    if (officialNameList.includes(rawDocName)) return fullMatch;

    const targetKey = toPhoneticHomophoneKey2a_(rawDocName);
    const exactKeyHit = officialNameList.find(function(offName) {
      return toPhoneticHomophoneKey2a_(offName) === targetKey;
    });
    if (exactKeyHit) {
      return exactKeyHit + '先生';
    }

    if (rawDocName.length >= 3) {
      const closeHits = officialNameList.filter(function(offName) {
        if (offName.length !== rawDocName.length) return false;
        if (toPhoneticHomophoneKey2a_(offName.slice(0, 2)) !== toPhoneticHomophoneKey2a_(rawDocName.slice(0, 2))) return false;
        let diff = 0;
        for (let i = 0; i < offName.length; i++) {
          if (offName[i] !== rawDocName[i]) diff++;
        }
        return diff === 1;
      });
      if (closeHits.length === 1) {
        return closeHits[0] + '先生';
      }
    }

    return fullMatch;
  });
}

function extractCandidateDoctorsFromPlanText2a_(planText) {
  const s = String(planText || '').replace(/｛※.*?｝/g, '').replace(/※.*?確認後に/g, '').trim();
  if (!s) return [];
  const names = [];
  const regex = /(?:①|②|③|前医の|後医の|2診目[：:]|1診[：:]|^|\s)([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})先生/g;
  let m;
  while ((m = regex.exec(s)) !== null) {
    const n = m[1].trim();
    if (n && !names.includes(n)) names.push(n);
  }
  return names;
}

function buildDeterministicSlotPlanFast_(slot, officialNameList) {
  const actions = [];
  const auditNotes = [];

  if (slot.currentMemo && String(slot.currentMemo).trim() && slot.currentMemo !== '-') {
    const correctedMemo = correctDoctorNamesInMemoText_(slot.currentMemo, officialNameList);
    correctedMemo.split(/(?:\r?\n|\s+[\/／]\s+)/).map(function(m) { return m.replace(/^[・\s]+/, '').trim(); }).filter(Boolean).forEach(function(mLine) {
      auditNotes.push('【進行中メモ】' + mLine);
    });
  }

  // 1. 第1優先：自拠点の前後医による「延長・早出」
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

  // 2. 第2優先：ドミナント定義書に基づく他拠点からの「2診目崩し」（条件付き延長セット含む）
  (slot.twoDocBreakdownCandidates || []).forEach(function(cand) {
    if (actions.length >= 3) return;
    const docs = cand.movableDoctors || [];
    if (docs.length === 0) return;

    const allSbu = docs.every(function(d) { return d.includes('SBU'); });
    const allowLabel = allSbu ? '（移動手当0円）' : '（移動依頼手当10,000円）';

    const parseMovDocObj = function(d) {
      const condMatch = String(d).match(/｛※(.*?)｝/);
      const condText = condMatch ? condMatch[1].trim() : '';
      const withoutCond = String(d).replace(/｛※.*?｝/g, '').trim();
      const base = withoutCond.replace(/\[.*?\]/g, '').trim();
      let label = base + '先生';
      if (withoutCond.includes('SBU')) label = base + '先生[★SBU]';
      else if (!withoutCond.includes('非常勤') && withoutCond.includes('所定休出')) label = base + '先生[常勤・所定休出]';
      return { label: label, condText: condText };
    };

    if (docs.length === 1) {
      const mObj = parseMovDocObj(docs[0]);
      if (mObj.condText) {
        actions.push('【2診目崩し・要延長セット】' + cand.clinic + '（※残る' + mObj.condText + 'に、2診目：' + mObj.label + '）へ移動打診' + allowLabel);
      } else {
        actions.push('【2診目崩し】' + cand.clinic + '（2診目：' + mObj.label + '）へ移動打診' + allowLabel);
      }
    } else {
      const m1 = parseMovDocObj(docs[0]);
      const m2 = parseMovDocObj(docs[1]);
      const twoNames = '①' + m1.label + (m1.condText ? '[※' + m1.condText + ']' : '') + ' / ②' + m2.label + (m2.condText ? '[※' + m2.condText + ']' : '');
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
    const base = docs[0].replace(/\[.*?\]|｛.*?｝/g, '').trim();
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
    shortClinicLabel: baseClinicName + ' ' + slot.absenceTime,
    clinicAndDept: exactClinicAndDept,
    absenceTime: exactAbsenceTime,
    assignee: slot.assignee || '',
    plan1: cleanPlanTextFormatting_(actions[0] || ''),
    plan2: cleanPlanTextFormatting_(actions[1] || ''),
    plan3: cleanPlanTextFormatting_(actions[2] || ''),
    rawAuditNotes: auditNotes
  };
}

function callGeminiAgentInBatches_(rawContext, customRulesText) {
  const t0 = Date.now();
  const allSlots = rawContext.slotsContext || [];
  const officialNameList = buildMasterNameCorrectionMap2a_(allSlots);

  const allSlotPlans = allSlots.map(function(origSlot) {
    return buildDeterministicSlotPlanFast_(origSlot, officialNameList);
  });

  const dayDocSlotMap = new Map();
  allSlotPlans.forEach(function(sp) {
    [
      { num: '案1', text: sp.plan1 },
      { num: '案2', text: sp.plan2 },
      { num: '案3', text: sp.plan3 }
    ].forEach(function(pObj) {
      const docNames = extractCandidateDoctorsFromPlanText2a_(pObj.text);
      docNames.forEach(function(dName) {
        const key = sp.date + '_' + dName;
        if (!dayDocSlotMap.has(key)) dayDocSlotMap.set(key, []);
        const list = dayDocSlotMap.get(key);
        if (!list.some(function(x) { return x.slotId === sp.slotId; })) {
          list.push({ slotId: sp.slotId, label: sp.shortClinicLabel, planNum: pObj.num });
        }
      });
    });
  });

  allSlotPlans.forEach(function(sp) {
    const notes = (sp.rawAuditNotes || []).slice();
    const myDocs = new Set();
    [sp.plan1, sp.plan2, sp.plan3].forEach(function(pt) {
      extractCandidateDoctorsFromPlanText2a_(pt).forEach(function(dn) { myDocs.add(dn); });
    });

    myDocs.forEach(function(dName) {
      const key = sp.date + '_' + dName;
      const occurrences = (dayDocSlotMap.get(key) || []).filter(function(x) { return x.slotId !== sp.slotId; });
      if (occurrences.length > 0) {
        const otherLabels = occurrences.map(function(x) { return x.label + '[' + x.planNum + ']'; }).join('、');
        notes.push('⚠️' + dName + '先生は同日他枠（' + otherLabels + '）でも候補');
      }
    });

    sp.auditNote = cleanAuditNoteFormatting_(notes.join('\n')) || '-';
  });

  Logger.log('  ⚡ 全 ' + allSlotPlans.length + ' 枠のプラン＆同日重複候補チェック完了（' + (Date.now() - t0) + 'ms）');

  const executiveSummary = '対象期間（' + rawContext.periodStart + '〜' + rawContext.periodEnd + '）の医師不在枠は計' + rawContext.unfilledCount + '件です。自拠点の前後医による延長・早出、およびドミナント近隣拠点からの2診目崩しを最優先で打診してください。';

  return {
    executiveSummary: executiveSummary,
    slotPlans: allSlotPlans,
    bundleRecommendations: [],
    announcementProposal: buildOfficialAnnouncementProposal_(rawContext.slotsContext),
    masterUpdateQuestions: []
  };
}