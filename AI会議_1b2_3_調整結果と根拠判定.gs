/**
 * ============================================================================
 * AI会議_1b2_3_調整結果と根拠判定.gs
 * （完成版：文・節単位の精密解析 ＋ 返答要約の欠落防止版）
 * ============================================================================
 */

function buildDoctorMailHistoryIndexForResults1b2_(ss) {
  const getRecentRows = function(shName, count) {
    const sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return [];
    const lastR = sh.getLastRow();
    const startR = Math.max(2, lastR - count + 1);
    return sh.getRange(startR, 1, lastR - startR + 1, Math.min(8, sh.getLastColumn())).getDisplayValues();
  };

  const kbRows = getRecentRows('ナレッジシート', 400);
  const qRows = getRecentRows('医師質問', 350);
  const aRows = getRecentRows('DS回答', 350);

  const byDoc = new Map();
  const ensureDoc = function(docKey) {
    if (!byDoc.has(docKey)) byDoc.set(docKey, { kbPairs: [], timeline: [] });
    return byDoc.get(docKey);
  };

  kbRows.forEach(function(r) {
    const dKey = normKeyForDoc_(r[1]);
    if (!dKey) return;
    const qText = String(r[2] || '').replace(/^【医師返信】\s*/, '').replace(/\s+/g, ' ').trim();
    const aText = String(r[3] || '').replace(/^【スタッフ打診】\s*/, '').replace(/\s+/g, ' ').trim();
    const dtStr = String(r[6] || '').trim();
    ensureDoc(dKey).kbPairs.push({
      dtStr: dtStr,
      docReply: qText,
      staffText: aText,
      flag: String(r[7] || '')
    });
  });

  aRows.forEach(function(r) {
    const dKey = normKeyForDoc_(r[2]);
    if (!dKey) return;
    const dt = new Date(r[1]);
    const tsMs = isNaN(dt.getTime()) ? 0 : dt.getTime();
    const dtStr = tsMs > 0 ? Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(r[1] || '');
    ensureDoc(dKey).timeline.push({
      role: 'STAFF',
      tsMs: tsMs,
      dtStr: dtStr,
      text: String(r[3] || '').replace(/\s+/g, ' ').trim()
    });
  });

  qRows.forEach(function(r) {
    const dKey = normKeyForDoc_(r[2]);
    if (!dKey) return;
    const dt = new Date(r[1]);
    const tsMs = isNaN(dt.getTime()) ? 0 : dt.getTime();
    const dtStr = tsMs > 0 ? Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(r[1] || '');
    ensureDoc(dKey).timeline.push({
      role: 'DOCTOR',
      tsMs: tsMs,
      dtStr: tsMs > 0 ? dtStr : String(r[1] || ''),
      text: String(r[3] || '').replace(/\s+/g, ' ').trim()
    });
  });

  byDoc.forEach(function(obj) {
    obj.timeline.sort(function(a, b) { return a.tsMs - b.tsMs; });
  });

  return byDoc;
}

function stripPoliteBoilerplate1b2_(rawText, docNameOpt) {
  let s = String(rawText || '')
    .replace(/[^\s]+@[^\s]+/g, '')
    .replace(/(?:ドクターサポート事務局|事務局)?\s*[^\s、。]{1,8}(?:様|さま|先生|御中)/g, '')
    .replace(/いつも大変お世話になっております[。、\s]*/g, '')
    .replace(/平素より大変お世話になっております[。、\s]*/g, '')
    .replace(/いつもお世話になっております[。、\s]*/g, '')
    .replace(/こちらこそお世話になっております[。、\s]*/g, '')
    .replace(/お世話になっております[。、\s]*/g, '')
    .replace(/ご調整いただきありがとうございます[。、\s]*/g, '')
    .replace(/お返事いただきありがとうございました[。、\s]*/g, '')
    .replace(/お知らせいただいた件につきまして了解いたしました[。、\s]*/g, '')
    .replace(/お声かけありがとうございます[。、\s]*/g, '')
    .replace(/お声がけいただきありがとうございます[。、\s]*/g, '')
    .replace(/ご連絡ありがとうございます[。、\s]*/g, '')
    .replace(/ご相談いただき、?誠にありがとうございます[。、\s]*/g, '')
    .replace(/お返事が遅くなりました[。、\s]*/g, '')
    .replace(/お力になれず申し訳(?:ございません|ありません)が?[。、\s]*/g, '')
    .replace(/お役に立てず申し訳(?:ございません|ありません)が?[。、\s]*/g, '')
    .replace(/ご提案にお応えすることができず申し訳(?:ございません|ありません)[。、\s]*/g, '')
    .replace(/誠に申し訳ございませんが?[。、\s]*/g, '')
    .replace(/誠に申し訳ありませんが?[。、\s]*/g, '')
    .replace(/大変申し訳ございませんが?[。、\s]*/g, '')
    .replace(/大変申し訳ありませんが?[。、\s]*/g, '')
    .replace(/申し訳ございませんが?[。、\s]*/g, '')
    .replace(/申し訳ありませんが?[。、\s]*/g, '')
    .replace(/ご検討いただければ幸いです[。、\s]*/g, '')
    .replace(/またの機会に.*?[。！!\s]+/g, '')
    .replace(/またお声をかけて.*?[。！!\s]+/g, '')
    .replace(/畏れ入りますが、?/g, '')
    .replace(/今後とも.*?よろしくお願いいたします[。！!\s]*/g, '')
    .replace(/引き続き.*?よろしくお願い(?:いたします|申し上げます)[。！!\s]*/g, '')
    .replace(/どうぞよろしくお願いいたします[。！!\s]*/g, '')
    .replace(/よろしくお願いいたします[。！!\s]*/g, '')
    .replace(/よろしくお願いします[。！!\s]*/g, '')
    .replace(/よろしくお願い申し上げます[。！!\s]*/g, '');

  if (docNameOpt) {
    const cleanDoc = String(docNameOpt).replace(/[\s\u3000]+/g, '');
    if (cleanDoc.length >= 2) {
      const surname = cleanDoc.slice(0, 2);
      s = s.replace(new RegExp(cleanDoc + '\\s*拝?$', 'g'), '').replace(new RegExp(surname + '\\s*拝?$', 'g'), '');
    }
  }

  return s.replace(/\s+/g, ' ').trim();
}

function splitIntoDateClauses1b2_(text) {
  const s = String(text || '')
    .replace(/(できません|難しいです|厳しいです|不可です|応需できません|お受けできません|入力しました)[がけ]?(?:れど)?[、，\s]*/g, '$1。')
    .replace(/([①②③④])/g, '。$1');

  return s.split(/[。！？!\n／]+/).map(function(c) { return c.trim(); }).filter(Boolean);
}

function extractCoreReasonSnippet1b2_(rawReplyText, slotDate, docNameOpt) {
  const mdShort = Number(slotDate.slice(5, 7)) + '/' + Number(slotDate.slice(8, 10));
  const mdJp = Number(slotDate.slice(5, 7)) + '月' + Number(slotDate.slice(8, 10)) + '日';
  const dayJp = Number(slotDate.slice(8, 10)) + '日';

  const s = stripPoliteBoilerplate1b2_(rawReplyText, docNameOpt);
  const clauses = splitIntoDateClauses1b2_(s);
  const targetClause = clauses.find(function(c) {
    return c.includes(mdJp) || c.includes(mdShort) || c.includes(dayJp);
  }) || clauses.find(function(c) {
    return /時から|までなら|何診|検討|可能|難|不可|厳|誰もいなかったら|お願い|予定があり|都合/.test(c);
  }) || clauses[0] || s;

  return targetClause.slice(0, 28);
}

function evaluateDoctorAdjustmentStatusObj1b2_(docName, planLabelOpt, slotDate, cleanClinic, mailHistoryIndex) {
  const dKey = normKeyForDoc_(docName);
  const docTag = planLabelOpt ? (docName + '(' + planLabelOpt + ')') : docName;
  if (!dKey || !mailHistoryIndex || !mailHistoryIndex.has(dKey)) {
    return { code: 'WAIT', docTag: docTag, dateShort: '', reason: '' };
  }

  const docHistory = mailHistoryIndex.get(dKey);
  const timeline = (docHistory.timeline || []).filter(function(ev) { return ev.dtStr >= '2026/10/28'; });

  const isIrrelevantAdminReply = function(txt) {
    return /パターン[ABＡＢ]|交通費|駐車場|パーキング|有休申請|jinjer|入力内容が消えた|画像が送信されました|2027年2月と3月/.test(txt);
  };

  const isShiftPitchMail = function(txt) {
    return /ご勤務につきましてご相談|ご変更いただくことは可能|延長していただくことは可能|前倒ししていただくことは可能|いずれかの枠にてご協力/.test(txt);
  };

  const mdShort = Number(slotDate.slice(5, 7)) + '/' + Number(slotDate.slice(8, 10));
  const mdJp = Number(slotDate.slice(5, 7)) + '月' + Number(slotDate.slice(8, 10)) + '日';
  const dayOnlyJp = Number(slotDate.slice(8, 10)) + '日';
  const clinicShort = cleanClinic.slice(0, 2);

  const knownOtherClinics = ['新百合ヶ丘', '柏の葉', '市川妙典', '武蔵小杉', '西新井', '東品川', '越谷', '川口', '流山', '海老名', '天王町', '二和向台', '北葛西', '亀有', '志木', '与野', '南浦和', '川越', '所沢', '高田', '用賀', '板橋'];

  const pitchEvents = [];
  for (let i = 0; i < timeline.length; i++) {
    const ev = timeline[i];
    if (ev.role === 'STAFF' && isShiftPitchMail(ev.text)) {
      const pitchDates = extractTargetDatesFromMailText1b2_(ev.text, ev.dtStr);
      const mentionsThisSlotDate = pitchDates.includes(slotDate) || ev.text.includes(mdJp) || ev.text.includes(mdShort);
      if (!mentionsThisSlotDate) continue;

      const docReplies = [];
      const staffFollowUps = [];
      for (let j = i + 1; j < timeline.length; j++) {
        const nextEv = timeline[j];
        if (nextEv.role === 'STAFF' && isShiftPitchMail(nextEv.text)) break;
        if (nextEv.role === 'DOCTOR' && !isIrrelevantAdminReply(nextEv.text)) {
          docReplies.push(nextEv);
        } else if (nextEv.role === 'STAFF' && !isIrrelevantAdminReply(nextEv.text)) {
          staffFollowUps.push(nextEv);
        }
      }

      pitchEvents.push({
        pitchEv: ev,
        pitchDates: pitchDates,
        isMultiDatePitch: pitchDates.length > 1 || /ご相談①|ご相談②|案が多く/.test(ev.text),
        docReplies: docReplies,
        staffFollowUps: staffFollowUps
      });
    }
  }

  (docHistory.kbPairs || []).forEach(function(kb) {
    if (kb.dtStr < '2026/10/28' || isIrrelevantAdminReply(kb.docReply)) return;
    const pDates = extractTargetDatesFromMailText1b2_(kb.staffText, kb.dtStr);
    const rDates = extractTargetDatesFromMailText1b2_(kb.docReply, kb.dtStr);
    if (pDates.includes(slotDate) || rDates.includes(slotDate)) {
      if (pitchEvents.length === 0) {
        pitchEvents.push({
          pitchEv: { dtStr: kb.dtStr, text: kb.staffText },
          pitchDates: pDates,
          isMultiDatePitch: pDates.length > 1 || /ご相談①|ご相談②|案が多く/.test(kb.staffText),
          docReplies: [{ dtStr: kb.dtStr, text: kb.docReply }],
          staffFollowUps: []
        });
      }
    }
  });

  if (pitchEvents.length === 0) {
    for (let k = timeline.length - 1; k >= 0; k--) {
      const ev = timeline[k];
      if (ev.role === 'STAFF' && (ev.text.includes(mdJp) || ev.text.includes(mdShort))) {
        if (/ご協力いただき.*ありがとうございます|ご快諾いただき|確定させていただきました|お引き受けくださり/.test(ev.text)) {
          if (ev.text.includes(clinicShort)) {
            return { code: 'ACCEPT', docTag: docTag, dateShort: ev.dtStr.slice(5, 10), reason: '' };
          }
          const confirmedOther = knownOtherClinics.find(function(oc) { return !cleanClinic.includes(oc) && ev.text.includes(oc); });
          if (confirmedOther) {
            return { code: 'DECLINE', docTag: docTag, dateShort: ev.dtStr.slice(5, 10), reason: '別拠点(' + confirmedOther + ')で確定' };
          }
        }
      }
    }
    return { code: 'WAIT', docTag: docTag, dateShort: '', reason: '' };
  }

  const latestChain = pitchEvents[pitchEvents.length - 1];
  if (latestChain.docReplies.length === 0) {
    return { code: 'WAIT', docTag: docTag, dateShort: latestChain.pitchEv.dtStr.slice(5, 10), reason: '' };
  }

  const substantiveReplies = latestChain.docReplies.filter(function(r) {
    const cleaned = stripPoliteBoilerplate1b2_(r.text, docName);
    return cleaned.length > 3;
  });
  const targetReplies = substantiveReplies.length > 0 ? substantiveReplies : latestChain.docReplies;
  const combinedRepliesText = targetReplies.map(function(r) { return r.text; }).join('。');
  const primaryReply = targetReplies[0];
  const rDateShort = primaryReply.dtStr.slice(5, 10);
  const coreReason = extractCoreReasonSnippet1b2_(combinedRepliesText, slotDate, docName);

  const staffConfirmedDecline = latestChain.staffFollowUps.some(function(sf) {
    return (sf.text.includes(mdJp) || !latestChain.isMultiDatePitch) &&
           /難しい旨承知|ご放念くださいませ|お気になさらないで|またご相談させて/.test(sf.text);
  });
  const staffConfirmedAcceptHere = latestChain.staffFollowUps.some(function(sf) {
    return (sf.text.includes(mdJp) || !latestChain.isMultiDatePitch) &&
           sf.text.includes(clinicShort) &&
           /確定させていただきました|ご快諾いただき|ご協力重ねて感謝|変更させていただきます/.test(sf.text);
  });

  for (let fIdx = 0; fIdx < latestChain.staffFollowUps.length; fIdx++) {
    const sfText = latestChain.staffFollowUps[fIdx].text;
    if (/お引き受けくださり|確定させていただきました|ご快諾いただき/.test(sfText) && !sfText.includes(clinicShort)) {
      const otherConfirmed = knownOtherClinics.find(function(oc) { return !cleanClinic.includes(oc) && sfText.includes(oc); });
      if (otherConfirmed) {
        return { code: 'DECLINE', docTag: docTag, dateShort: rDateShort, reason: '別拠点(' + otherConfirmed + ')で確定' };
      }
    }
  }

  const DECLINE_WORDS_RE = /勤務(?:すること)?が?できません|勤務不可|対応できません|応需できません|お受けできません|協力できません|入ることができません|延長(?:すること)?が?できません|難しい|厳しい|先約があり|予定があり|都合がつかず|都合が悪く|変更は不可|移動は不可|そのまま.*でお願いしたい|調整出来ませんでした|調整できませんでした/;
  const CONDITIONAL_WORDS_RE = /時からであれば|時からなら|までなら|まででしたら|何診体制|検討させて|検討は出来ます|明日.*確認し|可能性があり|誰もいなかったら可能|できれば他の方|時間外申請不可と記載|拘束時間鑑みて|夕から予定があるため/;
  const ACCEPT_WORDS_RE = /勤務可能|変更可能|移動は可能|どちらも可能|で可能です|大丈夫|承知いたしました|承りました|承ります|お受けいたします|勤務させて頂きます|引き受け|お願いできれば/;

  const clauses = splitIntoDateClauses1b2_(combinedRepliesText);
  const dateRegexForSlot = new RegExp('(?:^|[^\\d])(?:' + mdJp + '|' + mdShort + '|' + dayOnlyJp + ')(?:[^\\d]|$)');

  for (let sIdx = 0; sIdx < clauses.length; sIdx++) {
    const cl = clauses[sIdx];
    const hasMyDate = dateRegexForSlot.test(cl);
    const hasMyClinic = cl.includes(clinicShort);

    if (hasMyDate || hasMyClinic) {
      if (CONDITIONAL_WORDS_RE.test(cl)) {
        return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
      }
      if (DECLINE_WORDS_RE.test(cl)) {
        return { code: 'DECLINE', docTag: docTag, dateShort: rDateShort, reason: '' };
      }
      if (ACCEPT_WORDS_RE.test(cl)) {
        const otherInClause = knownOtherClinics.find(function(oc) { return !cleanClinic.includes(oc) && cl.includes(oc); });
        if (otherInClause && !hasMyClinic) {
          return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
        }
        return { code: 'ACCEPT', docTag: docTag, dateShort: rDateShort, reason: '' };
      }
    }
  }

  if (/いずれも.*(?:難|でき|不可|厳)|すべて.*(?:難|でき|不可)|可能な日はありません/.test(combinedRepliesText)) {
    return { code: 'DECLINE', docTag: docTag, dateShort: rDateShort, reason: '' };
  }

  if (CONDITIONAL_WORDS_RE.test(combinedRepliesText)) {
    return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
  }

  if (DECLINE_WORDS_RE.test(combinedRepliesText) || staffConfirmedDecline) {
    const allMentionedDates = extractTargetDatesFromMailText1b2_(combinedRepliesText, primaryReply.dtStr);
    if (allMentionedDates.length === 0 || allMentionedDates.includes(slotDate) || staffConfirmedDecline) {
      return { code: 'DECLINE', docTag: docTag, dateShort: rDateShort, reason: '' };
    }
  }

  if (latestChain.isMultiDatePitch) {
    const rDates = extractTargetDatesFromMailText1b2_(combinedRepliesText, primaryReply.dtStr);
    if (rDates.length > 0 && ACCEPT_WORDS_RE.test(combinedRepliesText) && !DECLINE_WORDS_RE.test(combinedRepliesText)) {
      if (rDates.includes(slotDate)) {
        return { code: 'ACCEPT', docTag: docTag, dateShort: rDateShort, reason: '' };
      } else {
        return { code: 'DECLINE', docTag: docTag, dateShort: rDateShort, reason: '他日程のみ承諾' };
      }
    }
    return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
  }

  if (ACCEPT_WORDS_RE.test(combinedRepliesText) || staffConfirmedAcceptHere) {
    const otherMentioned = knownOtherClinics.find(function(oc) { return !cleanClinic.includes(oc) && combinedRepliesText.includes(oc); });
    if (otherMentioned && !combinedRepliesText.includes(clinicShort)) {
      return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
    }
    return { code: 'ACCEPT', docTag: docTag, dateShort: rDateShort, reason: '' };
  }

  return { code: 'CHECK', docTag: docTag, dateShort: rDateShort, reason: coreReason };
}

function formatCompactAdjustmentLines1b2_(statusObjs) {
  if (!statusObjs || statusObjs.length === 0) return [];

  if (statusObjs.length <= 2) {
    return statusObjs.map(function(st) {
      if (st.code === 'ACCEPT') {
        return '→ ⭕' + st.docTag + '：承諾' + (st.dateShort ? '(' + st.dateShort + ')' : '');
      }
      if (st.code === 'DECLINE') {
        const sub = st.reason ? ('[' + st.reason + ']') : (st.dateShort ? '(' + st.dateShort + ')' : '');
        return '→ ❌' + st.docTag + '：辞退' + sub;
      }
      if (st.code === 'CHECK') {
        return '→ ⚠️' + st.docTag + '：要確認(' + (st.reason || '返信内容確認') + ')';
      }
      return '→ 📩' + st.docTag + '：返信待ち';
    });
  }

  const accepts = statusObjs.filter(function(s) { return s.code === 'ACCEPT'; });
  const checks = statusObjs.filter(function(s) { return s.code === 'CHECK'; });
  const declines = statusObjs.filter(function(s) { return s.code === 'DECLINE'; });
  const waits = statusObjs.filter(function(s) { return s.code === 'WAIT'; });

  const lines = [];
  if (accepts.length > 0) {
    lines.push('→ ⭕承諾：' + accepts.map(function(s) { return s.docTag; }).join('・'));
  }
  checks.forEach(function(s) {
    lines.push('→ ⚠️' + s.docTag + '：要確認(' + (s.reason || '返信確認') + ')');
  });
  if (declines.length > 0) {
    lines.push('→ ❌辞退：' + declines.map(function(s) { return s.docTag; }).join('・'));
  }
  if (waits.length > 0) {
    if (waits.length <= 3) {
      lines.push('→ 📩返信待ち：' + waits.map(function(s) { return s.docTag; }).join('・'));
    } else {
      const planWaits = waits.filter(function(s) { return s.docTag.includes('('); });
      const headList = planWaits.length > 0 ? planWaits : waits.slice(0, 2);
      const remainCount = waits.length - headList.length;
      lines.push('→ 📩返信待ち：' + headList.map(function(s) { return s.docTag; }).join('・') + (remainCount > 0 ? (' 他' + remainCount + '名') : ''));
    }
  }

  return lines;
}