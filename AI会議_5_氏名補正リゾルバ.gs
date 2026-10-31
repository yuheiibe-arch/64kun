/**
 * ============================================================================
 * AI会議_5_氏名補正リゾルバ.gs
 * （DSナレッジシート・BQシフトパック・担当者マスタから正式な医師名と医籍番号を特定する辞書）
 * ※スクリプトプロパティに BQ_SECRET_KEY を設定して使用します
 * ============================================================================
 */

const BRIDGE_DECRYPT_CONFIG = {
  STAFF_SS_ID: '1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM',
  STAFF_SHEET_NAME: '担当者',
  MAIL_QUERY_BQ: 'subject:"【AI連携暗号化JSON】BQシフト状況パック"',
  BQ_SECRET_PROP_KEY: 'BQ_SECRET_KEY'
};

/**
 * 異体字・旧字体をすべて共通の標準漢字に揃えて比較用キーを作る関数
 */
function toStandardKanjiForCaps_(str) {
  return String(str || '')
    .normalize('NFKC')
    .replace(/[\(（\[【].*$/g, '')
    .replace(/[\)）\]】★⚠️📩・]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[﨑嵜]/g, '崎')
    .replace(/髙/g, '高')
    .replace(/[瀧泷]/g, '滝')
    .replace(/[邊邉]/g, '辺')
    .replace(/[齋齊斉]/g, '斎')
    .replace(/澤/g, '沢')
    .replace(/廣/g, '広')
    .replace(/國/g, '国')
    .replace(/[嶋嶌]/g, '島')
    .replace(/[栁柳]/g, '柳')
    .replace(/[濵濱]/g, '浜')
    .replace(/櫻/g, '桜')
    .replace(/眞/g, '真')
    .replace(/惠/g, '恵')
    .replace(/ヶ/g, 'ケ')
    .trim();
}

function buildCapsDoctorAndStaffResolver_(ss) {
  const doctorMap = new Map();
  const dayShiftMap = new Map();

  const registerDoctor = function(rawName, licenseNo, dateStr, clinicName) {
    const cleanName = String(rawName || '')
      .replace(/[\(（\[【].*$/g, '')
      .replace(/[:：].*$/, '')
      .replace(/[\)）\]】★⚠️📩・]/g, '')
      .replace(/先生$/g, '')
      .replace(/[\s\u3000]+/g, '')
      .trim();
    if (!cleanName || cleanName === 'なし' || cleanName === '未充足' || /^\d+$/.test(cleanName)) return;

    const norm = toStandardKanjiForCaps_(cleanName);
    if (!norm) return;

    const rawLic = String(licenseNo || '').trim();
    const lic = (rawLic && rawLic !== '-' && rawLic.length >= 3) ? rawLic : '';

    const existing = doctorMap.get(norm);
    if (!existing) {
      doctorMap.set(norm, { name: cleanName, normName: norm, licenseNo: lic });
    } else if (!existing.licenseNo && lic) {
      existing.licenseNo = lic;
    }

    if (dateStr) {
      if (!dayShiftMap.has(dateStr)) dayShiftMap.set(dateStr, []);
      const list = dayShiftMap.get(dateStr);
      const cleanC = String(clinicName || '').replace(/[\(（].*?[\)）]/g, '').trim();
      if (!list.some(function(d) { return d.normName === norm && d.clinic === cleanC; })) {
        const master = doctorMap.get(norm) || {};
        list.push({ name: master.name || cleanName, normName: norm, licenseNo: master.licenseNo || lic, clinic: cleanC });
      }
    }
  };

  const collectFromSheet = function(sheetName, nameColIdx, licColIdx, dateColIdx, clinicColIdx) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(function(h) { return String(h || '').trim(); });

    let resolvedLicCol = licColIdx;
    if (resolvedLicCol < 0) {
      const foundLicIdx = headers.findIndex(function(h) { return h.includes('医籍') || h.includes('ID'); });
      if (foundLicIdx >= 0) resolvedLicCol = foundLicIdx;
    }

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      if (sheetName === '交渉しやすい医師') {
        const rawA = String(row[0] || '').trim();
        const rawC = String(row[2] || '').trim();
        if (/^\d+$/.test(rawA) || rawC.includes('AI自動抽出')) continue;
      }
      registerDoctor(
        nameColIdx >= 0 ? row[nameColIdx] : '',
        resolvedLicCol >= 0 ? row[resolvedLicCol] : '',
        dateColIdx >= 0 ? row[dateColIdx] : '',
        clinicColIdx >= 0 ? row[clinicColIdx] : ''
      );
    }
  };

  collectFromSheet('最新_個別打診候補', 5, 6, 0, 2);
  collectFromSheet('最新_おまとめ打診候補', 0, 1, -1, -1);
  collectFromSheet('最新_応募シフト', 2, 1, 0, 4);
  collectFromSheet('最新_連勤注意リスト', 1, 0, -1, -1);
  collectFromSheet('交渉しやすい医師', 0, 1, -1, -1);
  collectFromSheet('交渉注意医師', 0, 1, -1, -1);
  collectFromSheet('常勤依頼手当', 0, 1, -1, -1);
  collectFromSheet('SBU', 0, 1, -1, -1);

  const kbSh = ss.getSheetByName('ナレッジシート');
  if (kbSh && kbSh.getLastRow() >= 2) {
    const lastR = kbSh.getLastRow();
    const startR = Math.max(2, lastR - 499);
    const kbVals = kbSh.getRange(startR, 1, lastR - startR + 1, 2).getDisplayValues();
    for (let i = 0; i < kbVals.length; i++) {
      const rawA = String(kbVals[i][0] || '').trim();
      const validLic = (/^\d{5,7}$/.test(rawA) && Number(rawA) >= 100000) ? rawA : '';
      registerDoctor(kbVals[i][1], validLic, '', '');
    }
  }

  const shiftSh = ss.getSheetByName('最新_当日シフト');
  if (shiftSh && shiftSh.getLastRow() >= 2) {
    const sVals = shiftSh.getDataRange().getDisplayValues();
    for (let r = 1; r < sVals.length; r++) {
      [sVals[r][7], sVals[r][8], sVals[r][9]].forEach(function(cellText) {
        const tokens = (typeof splitDoctorTokensSafely_ === 'function')
          ? splitDoctorTokensSafely_(cellText)
          : String(cellText || '').split(/[,、，\/]+/);
        tokens.forEach(function(tok) { registerDoctor(tok, '', sVals[r][0], sVals[r][2]); });
      });
    }
  }

  try {
    const bqData = typeof fetchAndDecryptBqMailFast_ === 'function'
      ? fetchAndDecryptBqMailFast_()
      : fetchAndDecryptBqMailForResolver_();
    if (bqData) {
      (bqData.kakutei || []).forEach(function(s) { registerDoctor(s.n, s.id || '', s.d, s.c); });
      (bqData.obo || []).forEach(function(s) { registerDoctor(s.n, s.id || '', s.d, s.c); });
    }
  } catch (e) {}

  const staffMap = new Map();
  const loadStaffRows = function(sh) {
    if (!sh || sh.getLastRow() < 2) return;
    const stVals = sh.getDataRange().getDisplayValues();
    for (let r = 1; r < stVals.length; r++) {
      const colA = String(stVals[r][0] || '').trim();
      const colB = String(stVals[r][1] || '').trim();
      const colC = String(stVals[r][2] || '').trim();
      const colD = String(stVals[r][3] || '').trim();
      const surname = colD || colA.split(/[\s\u3000]+/)[0];
      if (!surname) continue;
      [colA, colB, colC, colD, surname].filter(Boolean).forEach(function(alias) {
        staffMap.set(alias.replace(/[\s\u3000]+/g, ''), surname);
      });
    }
  };

  loadStaffRows(ss.getSheetByName(BRIDGE_DECRYPT_CONFIG.STAFF_SHEET_NAME));
  if (staffMap.size === 0) {
    try {
      const staffSs = SpreadsheetApp.openById(BRIDGE_DECRYPT_CONFIG.STAFF_SS_ID);
      loadStaffRows(staffSs.getSheetByName(BRIDGE_DECRYPT_CONFIG.STAFF_SHEET_NAME) || staffSs.getSheets()[0]);
    } catch (e) {}
  }

  return {
    resolveDoctor: function(rawInputName, dateStr, preferredClinic) {
      const normTarget = toStandardKanjiForCaps_(rawInputName);
      if (!normTarget) return { name: rawInputName, licenseNo: '', dayClinic: '' };

      if (dateStr && dayShiftMap.has(dateStr)) {
        const dayDocs = dayShiftMap.get(dateStr);
        const exactDay = dayDocs.find(function(d) {
          return d.normName === normTarget && (!preferredClinic || d.clinic.includes(preferredClinic));
        });
        if (exactDay) {
          const master = doctorMap.get(exactDay.normName) || exactDay;
          return { name: master.name, licenseNo: master.licenseNo || exactDay.licenseNo, dayClinic: exactDay.clinic };
        }
        const prefixMatches = dayDocs.filter(function(d) { return d.normName.startsWith(normTarget); });
        if (prefixMatches.length > 0) {
          const chosen = (preferredClinic ? prefixMatches.find(function(d) { return d.clinic.includes(preferredClinic); }) : null) || prefixMatches[0];
          const master = doctorMap.get(chosen.normName) || chosen;
          return { name: master.name, licenseNo: master.licenseNo || chosen.licenseNo, dayClinic: chosen.clinic };
        }
      }

      if (doctorMap.has(normTarget)) {
        const hit = doctorMap.get(normTarget);
        return { name: hit.name, licenseNo: hit.licenseNo, dayClinic: '' };
      }

      for (const [normKey, docObj] of doctorMap.entries()) {
        if (normKey.startsWith(normTarget) && normTarget.length >= 2) {
          return { name: docObj.name, licenseNo: docObj.licenseNo, dayClinic: '' };
        }
      }

      return { name: String(rawInputName || '').replace(/先生$/g, '').trim(), licenseNo: '', dayClinic: '' };
    },

    resolveStaff: function(rawStaff) {
      const s = String(rawStaff || '').replace(/[\s\u3000]+/g, '').trim();
      if (!s) return '';
      if (staffMap.has(s)) return staffMap.get(s);
      for (const [alias, surname] of staffMap.entries()) {
        if (alias.startsWith(s) || s.startsWith(alias)) return surname;
      }
      return s;
    }
  };
}

function fetchAndDecryptBqMailForResolver_() {
  const threads = GmailApp.search(BRIDGE_DECRYPT_CONFIG.MAIL_QUERY_BQ, 0, 1);
  if (!threads || threads.length === 0) return null;
  const msgs = threads[0].getMessages();
  const latestMsg = msgs[msgs.length - 1];

  const attachments = latestMsg.getAttachments();
  const rawText = (attachments && attachments.length > 0)
    ? String(attachments[0].getDataAsString('UTF-8') || '').trim()
    : String(latestMsg.getPlainBody() || '').trim();

  if (rawText.startsWith('GZIP_B64:')) {
    const bytes = Utilities.base64Decode(rawText.slice(9).trim());
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(bytes, 'application/x-gzip')).getDataAsString('UTF-8'));
  }

  const m = rawText.match(/ENC_V1:([A-Za-z0-9+\/=_\-\r\n]+)/);
  if (m) {
    const secretKey = String(PropertiesService.getScriptProperties().getProperty(BRIDGE_DECRYPT_CONFIG.BQ_SECRET_PROP_KEY) || '').trim();
    if (!secretKey) return null;
    const encBytes = Utilities.base64Decode(m[1].replace(/[\r\n\s]+/g, ''));
    const keyBytes = Utilities.newBlob(secretKey).getBytes();
    const out = encBytes.map(function(b, i) { return b ^ keyBytes[i % keyBytes.length]; });
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(out, 'application/x-gzip')).getDataAsString('UTF-8'));
  }

  return JSON.parse(rawText);
}