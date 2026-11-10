/**
 * ============================================================================
 * AI会議_4_2a_提出_前処理と分割.gs（3分割の1）
 * ・募集シフト（bosyu）長時間枠キャッシュ構築
 * ・3列目の「確定案ヘッダー」と「→ 調整結果」の安全分離（\r 正規化対応）
 * ・カッコ階層を壊さないトークン分割・ステイ判定・担当者名クリーニング
 * ============================================================================
 */

let SUBMIT_RECRUIT_CACHE_4_2_ = null;

function getSubmitRecruitContext4_2_() {
  if (SUBMIT_RECRUIT_CACHE_4_2_) return SUBMIT_RECRUIT_CACHE_4_2_;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getRows = function(name) {
    const sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() >= 2) ? sh.getDataRange().getDisplayValues().slice(1) : [];
  };

  const activeShifts = getRows('最新_当日シフト');
  const rawUnfilled = getRows('最新_不在枠');
  const spotCands = getRows('最新_個別打診候補');

  const recruitSlotsIndex = new Map();
  const addIdx = function(dRaw, rawClinic, rawTime, deptHint) {
    if (typeof normalizeDateYmd1b3_ !== 'function') return;
    const d = normalizeDateYmd1b3_(dRaw);
    const cleanC = cleanClinicBaseName_(rawClinic);
    if (!d || !cleanC || cleanC.includes('バックアップ')) return;
    const deptTag = (String(rawClinic || '').includes('内科') || String(deptHint || '').includes('内科')) ? '内科' : '小児科';
    const key = d + '_' + cleanC + '_' + deptTag;
    if (!recruitSlotsIndex.has(key)) recruitSlotsIndex.set(key, []);
    const list = recruitSlotsIndex.get(key);
    normalizeAndExpandAbsenceTime1b3_(rawTime, cleanC).forEach(function(t) {
      if (t && !list.includes(t)) list.push(t);
    });
  };

  rawUnfilled.forEach(function(r) { addIdx(r[0], r[1], r[2], ''); });
  spotCands.forEach(function(r) { addIdx(r[0], r[2], r[3], ''); });

  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []).forEach(function(m) {
          addIdx(m.d || m.date, m.c || m.clinic, m.t || m.time, m.dp || m.dept || '');
        });
      }
    } catch (e) {}
  }

  SUBMIT_RECRUIT_CACHE_4_2_ = {
    recruitSlotsIndex: recruitSlotsIndex,
    activeShifts: activeShifts
  };
  return SUBMIT_RECRUIT_CACHE_4_2_;
}

function mapCustomActionToApproachType_(actionTypeStr, fullText) {
  const s = String(actionTypeStr || '') + ' ' + String(fullText || '');
  if (s.includes('バックアップ')) return 'バックアップ勤務依頼';
  if (s.includes('移動') || s.includes('2診目崩し')) return '勤務地変更相談';
  if (s.includes('延長')) return '延長相談';
  if (s.includes('早出') || s.includes('前倒')) return '早出相談';
  return '個別アプローチ';
}

function detectBandFromTimeStr4_(timeStr) {
  const m = String(timeStr || '').match(/(\d{1,2}):(\d{2})/);
  if (!m) return 'A';
  const sh = Number(m[1]);
  if (sh < 13) return 'A';
  if (sh < 18) return 'B';
  return 'C';
}

/**
 * 3列目（確定案・調整結果）から「確定案ヘッダー」と「→ 調整結果行（折り返し引用含む）」を安全に分離する
 */
function extractDecisionHeaderAndArrows4_2_(rawDecisionText) {
  const norm = String(rawDecisionText || '').replace(/\r\n|\r/g, '\n');
  const lines = norm.split('\n');
  const headerLines = [];
  const arrowLines = [];
  let inArrowSection = false;

  for (let i = 0; i < lines.length; i++) {
    const s = lines[i].trim();
    if (!s) continue;
    if (/^(?:→|⇒|⇨|・?【調整結果】)/.test(s)) {
      inArrowSection = true;
      arrowLines.push(s);
      continue;
    }
    if (inArrowSection) {
      if (/^(?:【?案\s*[123１２３]|全案|ステイ|特別時給|バックアップ医師)/.test(s)) {
        inArrowSection = false;
        headerLines.push(s);
      } else {
        arrowLines.push(s);
      }
    } else {
      headerLines.push(s);
    }
  }

  return {
    header: headerLines.join(' ').trim(),
    arrowLines: arrowLines
  };
}

/**
 * カッコ（() / （） / 【】 / [] / 「」）の階層を壊さずに確定案トークンを分割する
 */
function splitTopLevelDecisionTokens4_2_(decStr) {
  const preSplit = String(decStr || '').replace(/\s+(?=【?案\s*[123１２３]】?(?:[\s\(（:：]|$))/g, '\n');
  const results = [];
  let cur = '';
  let depth = 0;

  for (let i = 0; i < preSplit.length; i++) {
    const ch = preSplit[i];
    if ('(（[［{｛【「'.indexOf(ch) !== -1) {
      depth++;
      cur += ch;
      continue;
    }
    if (')）]］}｝】」'.indexOf(ch) !== -1) {
      depth = Math.max(0, depth - 1);
      cur += ch;
      continue;
    }

    if (depth === 0) {
      if (',、，/／\n・+＋'.indexOf(ch) !== -1) {
        if (cur.trim()) results.push(cur.trim());
        cur = '';
        continue;
      }
      if (preSplit.slice(i, i + 3) === ' と ') {
        if (cur.trim()) results.push(cur.trim());
        cur = '';
        i += 2;
        continue;
      }
    }
    cur += ch;
  }
  if (cur.trim()) results.push(cur.trim());
  return results;
}

/**
 * ステイ・充足・未選択など提出不要な確定案かどうかを判定する
 */
function isStayOrSkipDecision4_2_(decStr) {
  const s = String(decStr || '').trim();
  if (!s) return true;
  if (/^(?:ステイ|募集継続・ステイ|見送り|なし|不要|対応不要|充足|応募あり充足|ー|―|\-|保留|要検討|併行募集)(?:[\s（\(].*)?$/.test(s)) {
    return true;
  }
  if (/^特別時給.*ステイ$/.test(s) || /[＋\+]\s*ステイ$/.test(s)) {
    return true;
  }
  if (s.includes('医師未選択')) {
    return true;
  }
  return false;
}

/**
 * 担当者セルからメモ混入（「案1三崎先生ご連絡済北村」等）を除去して担当者名を解決する
 */
function cleanAssigneeText4_2_(assigneeText, resolver) {
  const raw = String(assigneeText || '').replace(/\r\n|\r/g, '\n').trim();
  if (!raw || raw === '担当者未定' || raw === 'あとで決定' || raw === '(未定)' || raw === '未定' || raw === '(担当不要)') {
    return '未定';
  }
  const stripped = raw
    .replace(/[（\(].*?[）\)]/g, ' ')
    .replace(/.*?(?:連絡済[み]?|打診済[み]?|確認済[み]?)/g, ' ')
    .replace(/【?案\s*[123１２３]】?/g, ' ')
    .replace(/[一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{1,6}先生/g, ' ')
    .trim();

  const firstToken = (stripped || raw).split(/[\s\u3000\/／・,、]+/).filter(Boolean).pop() || '';
  if (!firstToken) return '未定';
  const resolved = resolver ? resolver.resolveStaff(firstToken) : firstToken;
  return resolved || '未定';
}