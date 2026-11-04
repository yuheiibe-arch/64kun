/**
 * ============================================================================
 * 6a_編集くん64_マスタと整形.gs
 * （モーダル大画面化1440x920・異体字/漢字ブレ強化・括弧内メモ保護分割版）
 * ============================================================================
 */

const HENSHUKUN_CONFIG = {
  DRAFT_PROP_KEY_PREFIX: 'HENSHUKUN_DRAFT_',
  DRAFT_SHEET_NAME: '編集くん_一時保存',
  ACTIVE_USERS_CACHE_KEY: 'HENSHUKUN_ACTIVE_PLAYERS_V1',
  ACTIVE_USER_TTL_SEC: 45
};

function openHenshukunModal() {
  const html = HtmlService.createHtmlOutput(buildHenshukun64Html_())
    .setWidth(1440)
    .setHeight(920);
  SpreadsheetApp.getUi().showModalDialog(html, '🎮 編集くん 64 ― AI超募集MTG シフト確定スタジオ');
}

function formatClinicTitle64_(rawTitle) {
  const s = String(rawTitle || '').trim();
  if (s.includes('亀有') || s.includes('北葛西')) return s;
  return s.replace(/[（\(]小児科[）\)]/g, '').replace(/小児科/g, '').trim();
}

function cleanSidebarDisplayText_(text) {
  const lines = String(text || '')
    .replace(/（特別時給通常時給）/g, '（通常時給）')
    .replace(/（特別時給特別時給）/g, '（特別時給）')
    .replace(/[（\(]\s*(?:スポット\s*[\/／・]\s*紹介会社|紹介会社\s*[\/／・]\s*スポット|スポット|紹介会社|スポット医師|紹介会社医師)\s*[）\)]/g, '')
    .split('\n');

  const cleanedLines = [];
  for (let i = 0; i < lines.length; i++) {
    let s = lines[i].trim();
    if (!s) continue;
    if (s.startsWith('💡')) {
      const body = s.replace(/^💡(?:現状・注意|注意)[：:]\s*/, '');
      const parts = body
        .split(/(?:。|\s+[\/／]\s+|(?<!\d)[\/／](?!\d))+/)
        .map(function(p) { return p.replace(/^[・\s]+/, '').trim(); })
        .filter(Boolean)
        .filter(function(p) {
          if (/(通常常勤|通常勤務).*?(延長不可|早出不可|移動不可|打診不可|不可)/.test(p)) return false;
          if (/^(前医|後医|前後医|前後枠).*?(延長不可|早出不可|調整余地なし|医師なし|不在|なし|不可)$/.test(p)) return false;
          if (/^(延長不可|早出不可|調整余地なし|前後枠の医師なし|前後枠なし|特になし|該当なし|問題なし)$/.test(p)) return false;
          return true;
        });
      if (parts.length > 0) {
        cleanedLines.push('💡注意：\n' + parts.map(function(p) { return '・' + p; }).join('\n'));
      }
    } else if (s.startsWith('・')) {
      const itemBody = s.replace(/^[・\s]+/, '').trim();
      if (/^(前医|後医|前後医|前後枠).*?(延長不可|早出不可|調整余地なし|医師なし|不在|なし|不可)$/.test(itemBody)) continue;
      if (/^(延長不可|早出不可|調整余地なし|前後枠の医師なし|前後枠なし|特になし|該当なし|問題なし)$/.test(itemBody)) continue;
      cleanedLines.push('・' + itemBody);
    } else {
      cleanedLines.push(s);
    }
  }
  return cleanedLines.join('\n');
}

function toJapaneseDateLabel64_(rawDateStr) {
  const s = String(rawDateStr || '').trim();
  const m = s.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return s;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(y, mo - 1, d);
  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = isNaN(dt.getTime()) ? '' : ' (' + dows[dt.getDay()] + ')';
  return y + '年' + mo + '月' + d + '日' + dow;
}

function toSortDateKey64_(rawDateStr) {
  const m = String(rawDateStr || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return '9999/99/99';
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

function toSearchNorm64_(str) {
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
    .replace(/眞/g, '真').replace(/惠/g, '恵').replace(/[壽寿]/g, '寿')
    .replace(/[靜静]/g, '静').replace(/[榮栄]/g, '栄').replace(/[彌弥]/g, '弥')
    .replace(/[祐裕]/g, '祐').replace(/[佳圭]/g, '佳').replace(/[智知]/g, '智')
    .replace(/ヶ/g, 'ケ')
    .trim();
}

function loadDoctorSearchList64_(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('HENSHUKUN_DOC_SEARCH_LIST_V4');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const docMap = new Map();
  const addDoc = function(rawName, rawLic, tagLabel) {
    const cleanName = String(rawName || '')
      .replace(/[\(（\[【].*$/g, '')
      .replace(/[:：].*$/, '')
      .replace(/[\)）\]】★⚠️📩・]/g, '')
      .replace(/先生$/g, '')
      .replace(/[\s\u3000]+/g, '')
      .trim();
    if (!cleanName || cleanName === 'なし' || cleanName === '未充足' || /^\d+$/.test(cleanName)) return;

    const norm = toSearchNorm64_(cleanName);
    if (!norm) return;

    const licStr = String(rawLic || '').trim();
    const validLic = (/^\d{4,7}$/.test(licStr) || (licStr.length >= 4 && licStr !== '-' && !/^\d{1,3}$/.test(licStr))) ? licStr : '';

    const ex = docMap.get(norm);
    if (!ex) {
      docMap.set(norm, { name: cleanName, norm: norm, licenseNo: validLic, tag: tagLabel || '' });
    } else {
      if (!ex.licenseNo && validLic) ex.licenseNo = validLic;
      if (!ex.tag && tagLabel) ex.tag = tagLabel;
    }
  };

  const scanSheet = function(shName, nameCol, licCol, tagText) {
    const sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      if (shName === '交渉しやすい医師') {
        const rawA = String(vals[r][0] || '').trim();
        const rawC = String(vals[r][2] || '').trim();
        if (/^\d+$/.test(rawA) || rawC.includes('AI自動抽出')) continue;
      }
      addDoc(vals[r][nameCol], licCol >= 0 ? vals[r][licCol] : '', tagText);
    }
  };

  scanSheet('常勤依頼手当', 0, 1, '常勤手当対象');
  scanSheet('SBU', 0, 1, 'SBU');
  scanSheet('交渉しやすい医師', 0, 1, '交渉しやすい');
  scanSheet('交渉注意医師', 0, 1, '交渉注意');
  scanSheet('最新_おまとめ打診候補', 0, 1, '');
  scanSheet('最新_個別打診候補', 5, 6, '');
  scanSheet('最新_応募シフト', 2, 1, '');
  scanSheet('最新_連勤注意リスト', 1, 0, '');

  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        (bqData.kakutei || []).forEach(function(s) { addDoc(s.n, s.id || '', ''); });
        (bqData.obo || []).forEach(function(s) { addDoc(s.n, s.id || '', ''); });
      }
    } catch (e) {}
  }

  const kbSh = ss.getSheetByName('ナレッジシート');
  if (kbSh && kbSh.getLastRow() >= 2) {
    const lastR = kbSh.getLastRow();
    const startR = Math.max(2, lastR - 799);
    const kbVals = kbSh.getRange(startR, 1, lastR - startR + 1, 2).getDisplayValues();
    for (let i = 0; i < kbVals.length; i++) {
      const rawA = String(kbVals[i][0] || '').trim();
      const validKbLic = (/^\d{5,7}$/.test(rawA) && Number(rawA) >= 100000) ? rawA : '';
      addDoc(kbVals[i][1], validKbLic, '');
    }
  }

  const shiftSh = ss.getSheetByName('最新_当日シフト');
  if (shiftSh && shiftSh.getLastRow() >= 2) {
    const sVals = shiftSh.getDataRange().getDisplayValues();
    for (let r = 1; r < sVals.length; r++) {
      [sVals[r][7], sVals[r][8], sVals[r][9]].forEach(function(txt) {
        const tokens = (typeof splitDoctorTokensSafely_ === 'function')
          ? splitDoctorTokensSafely_(txt)
          : String(txt || '').split(/[,、，\n]+/);
        tokens.forEach(function(tok) { addDoc(tok, '', ''); });
      });
    }
  }

  const allDocs = Array.from(docMap.values());
  const withLicList = allDocs.filter(function(d) { return Boolean(d.licenseNo); });

  allDocs.forEach(function(d) {
    if (d.licenseNo) return;
    if (d.norm.length < 2) return;
    const prefix2 = d.norm.slice(0, 2);
    const sameSurnameWithLic = withLicList.filter(function(cand) {
      return cand.norm.slice(0, 2) === prefix2;
    });

    if (sameSurnameWithLic.length === 1) {
      d.licenseNo = sameSurnameWithLic[0].licenseNo;
      if (!d.tag && sameSurnameWithLic[0].tag) d.tag = sameSurnameWithLic[0].tag;
    } else if (sameSurnameWithLic.length > 1) {
      const closeMatch = sameSurnameWithLic.find(function(cand) {
        if (cand.norm.length !== d.norm.length) return false;
        let diff = 0;
        for (let k = 0; k < d.norm.length; k++) {
          if (d.norm[k] !== cand.norm[k]) diff++;
        }
        return diff <= 1;
      });
      if (closeMatch) {
        d.licenseNo = closeMatch.licenseNo;
      }
    }
  });

  const list = allDocs.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });
  try {
    if (list.length > 0 && list.length <= 1500) {
      cache.put('HENSHUKUN_DOC_SEARCH_LIST_V4', JSON.stringify(list), 1800);
    }
  } catch (e) {}
  return list;
}

function loadClinicAndStaffMasters64_(ss) {
  const clinicMap = new Map();
  const groupSet = new Set();
  const areaToGroups = { '関東': [], '関西': [] };

  const clinicSh = ss.getSheetByName('拠点名');
  if (clinicSh && clinicSh.getLastRow() >= 2) {
    const vals = clinicSh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const cName = String(vals[r][0] || '').trim();
      const cGroup = String(vals[r][1] || '').trim() || '東京';
      const cDept = String(vals[r][2] || '').trim() || '小児科';
      const cArea = String(vals[r][3] || '').trim() || '関東';
      const cNo = Number(vals[r][5] || r);
      if (!cName) continue;

      clinicMap.set(cName, { name: cName, group: cGroup, dept: cDept, area: cArea, no: cNo });
      groupSet.add(cGroup);
      if (!areaToGroups[cArea]) areaToGroups[cArea] = [];
      if (!areaToGroups[cArea].includes(cGroup)) areaToGroups[cArea].push(cGroup);
    }
  }

  const staffList = [];
  const seenStaff = new Set();
  const readStaffSheet = function(sh) {
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const surname = String(vals[r][3] || vals[r][0].split(/[\s\u3000]+/)[0] || '').trim();
      if (surname && !seenStaff.has(surname)) {
        seenStaff.add(surname);
        staffList.push(surname);
      }
    }
  };

  readStaffSheet(ss.getSheetByName('担当者'));
  if (staffList.length === 0 && typeof BRIDGE_DECRYPT_CONFIG !== 'undefined') {
    try {
      const extSs = SpreadsheetApp.openById(BRIDGE_DECRYPT_CONFIG.STAFF_SS_ID);
      readStaffSheet(extSs.getSheetByName('担当者') || extSs.getSheets()[0]);
    } catch (e) {}
  }
  staffList.sort(function(a, b) { return a.localeCompare(b, 'ja'); });

  return {
    clinicMap: clinicMap,
    groups: Array.from(groupSet),
    areaToGroups: areaToGroups,
    staffList: staffList,
    doctorSearchList: loadDoctorSearchList64_(ss)
  };
}

function resolveClinicMeta64_(rawClinicText, clinicMap) {
  const clean = String(rawClinicText || '')
    .replace(/【|】/g, '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/小児科|内科/g, '')
    .trim();

  if (clinicMap.has(clean)) return clinicMap.get(clean);
  for (const [k, v] of clinicMap.entries()) {
    if (clean.includes(k) || k.includes(clean)) return v;
  }

  const area = (typeof inferAreaForDoc_ === 'function') ? inferAreaForDoc_(clean) : '東京';
  const macroArea = (area === '関西') ? '関西' : '関東';
  const group = (area === '関西') ? '大阪' : area;
  return { name: clean, group: group, area: macroArea, no: 999 };
}