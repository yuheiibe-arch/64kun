/**
 * ============================================================================
 * 6b_1_編集くん64_時間と募集枠抽出.gs（6b系 3分割の1）
 * ・時間帯パース・重複判定・2重連結テキスト整形
 * ・バックアップ医師一覧 ＋ BQ募集枠（bosyu）時間帯マップの一括構築・募集枠時間への表示変換
 * ============================================================================
 */

function extractSlotTimeRange64_(str) {
  const m = String(str || '').match(/(\d{1,2}:\d{2}\s*[\-〜～~]\s*\d{1,2}:\d{2})/);
  return m ? m[1].replace(/\s+/g, '').replace(/[〜～~]/g, '-') : String(str || '').trim();
}

function parseTimeRangeToMins64_(rangeStr) {
  const m = String(rangeStr || '').match(/(\d{1,2}):(\d{2})\s*[\-〜～~]\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return {
    startMin: Number(m[1]) * 60 + Number(m[2]),
    endMin: Number(m[3]) * 60 + Number(m[4]),
    normLabel: ('0' + Number(m[1])).slice(-2) + ':' + m[2] + '-' + ('0' + Number(m[3])).slice(-2) + ':' + m[4]
  };
}

function isTimeRangeOverlapping64_(rangeA, rangeB) {
  if (!rangeA || !rangeB) return true;
  if (rangeA === rangeB) return true;
  const a = parseTimeRangeToMins64_(rangeA);
  const b = parseTimeRangeToMins64_(rangeB);
  if (!a || !b) return rangeA === rangeB;
  return Math.max(a.startMin, b.startMin) < Math.min(a.endMin, b.endMin);
}

function getBandCodeFromMins64_(startMin) {
  if (startMin < 780) return 'A';
  if (startMin < 1080) return 'B';
  return 'C';
}

function dedupeRepeatedDecisionHeader64_(text) {
  let s = String(text || '').trim();
  if (!s || s.length < 6) return s;
  const half = Math.floor(s.length / 2);
  if (s.length % 2 === 0 && s.slice(0, half) === s.slice(half)) {
    return s.slice(0, half).trim();
  }
  const spaceDup = s.match(/^(.{4,}?)\s+\1$/);
  if (spaceDup) return spaceDup[1].trim();
  return s;
}

/**
 * バックアップ医師一覧 ＋ BQ募集枠（bosyu）の時間帯マップを1回のスキャンでまとめて構築する
 */
function extractBackupAndRecruitSlots64_(ss, doctorSearchList) {
  const backupDocsByDate = {};
  const recruitSlotsMap = new Map(); // key: "yyyy/MM/dd|拠点名|診療科" -> [{startMin, endMin, normLabel}]

  const docLicMap = new Map();
  (doctorSearchList || []).forEach(function(d) {
    if (d.norm && d.licenseNo) docLicMap.set(d.norm, d.licenseNo);
  });

  const registerBackupDoc = function(sortDate, band, rawDocName, explicitLic) {
    if (!sortDate || sortDate === '9999/99/99' || !band) return;
    const cleanName = String(rawDocName || '')
      .replace(/[\(（\[【].*$/g, '')
      .replace(/[:：].*$/, '')
      .replace(/[\)）\]】★⚠️📩・]/g, '')
      .replace(/先生$/g, '')
      .replace(/[\s\u3000]+/g, '')
      .trim();
    if (!cleanName || cleanName === '未充足' || cleanName === 'なし' || cleanName === '-') return;

    const norm = toSearchNorm64_(cleanName);
    if (!norm) return;

    if (!backupDocsByDate[sortDate]) {
      backupDocsByDate[sortDate] = { A: [], B: [], C: [] };
    }
    const bandList = backupDocsByDate[sortDate][band];
    if (bandList && !bandList.some(function(d) { return d.norm === norm; })) {
      bandList.push({
        name: cleanName,
        norm: norm,
        licenseNo: explicitLic || docLicMap.get(norm) || ''
      });
    }
  };

  // 1. 『最新_当日シフト』シートからバックアップ枠を抽出
  const sh = ss.getSheetByName('最新_当日シフト');
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const clinicRaw = String(row[2] || '').trim();
      if (!clinicRaw.includes('バックアップ')) continue;

      const sortDate = toSortDateKey64_(row[0]);
      if (!sortDate || sortDate === '9999/99/99') continue;

      [
        { band: 'A', text: row[7] },
        { band: 'B', text: row[8] },
        { band: 'C', text: row[9] }
      ].forEach(function(bc) {
        const tokens = (typeof splitDoctorTokensSafely_ === 'function')
          ? splitDoctorTokensSafely_(bc.text)
          : String(bc.text || '').split(/[,、，\/]+/);
        tokens.forEach(function(tok) { registerBackupDoc(sortDate, bc.band, tok, ''); });
      });
    }
  }

  // 2. BQシフトパックから ①バックアップ確定医 と ②実際の募集枠時間（bosyu）を同時取得
  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData) {
        (bqData.kakutei || []).forEach(function(s) {
          const st = String(s.st || s.status || '').trim();
          if (st && (st.includes('キャンセル') || st.includes('取消') || st.includes('却下'))) return;
          const cRaw = String(s.c || s.clinic || '').trim();
          if (!cRaw.includes('バックアップ')) return;

          const sortDate = toSortDateKey64_(s.d || s.date);
          const parsedT = parseTimeRangeToMins64_(String(s.t || s.iv || s.time || s.workTime || '').normalize('NFKC'));
          if (!sortDate || sortDate === '9999/99/99' || !parsedT) return;

          const lic = String(s.id || '').trim();
          if (parsedT.startMin < 780 && parsedT.endMin > 540) registerBackupDoc(sortDate, 'A', s.n || s.name, lic);
          if (parsedT.startMin < 1080 && parsedT.endMin > 900) registerBackupDoc(sortDate, 'B', s.n || s.name, lic);
          if (parsedT.startMin < 1260 && parsedT.endMin > 1080) registerBackupDoc(sortDate, 'C', s.n || s.name, lic);
        });

        const rawBosyuList = (bqData.bosyu || []).concat(bqData.mijyusoku || []).concat(bqData.boshu || []);
        rawBosyuList.forEach(function(b) {
          const st = String(b.st || b.status || '').trim();
          if (st && (st.includes('終了') || st.includes('取消') || st.includes('削除') || st.includes('非公開'))) return;

          const sortDate = toSortDateKey64_(b.d || b.date);
          const cClean = String(b.c || b.clinic || '').replace(/[\(（].*?[\)）]/g, '').trim();
          const dept = String(b.dp || b.dept || '').includes('内科') || String(b.c || '').includes('内科') ? '内科' : '小児科';
          const parsedT = parseTimeRangeToMins64_(String(b.t || b.time || '').normalize('NFKC'));
          if (!sortDate || sortDate === '9999/99/99' || !cClean || !parsedT) return;

          const rKey = sortDate + '|' + cClean + '|' + dept;
          if (!recruitSlotsMap.has(rKey)) recruitSlotsMap.set(rKey, []);
          const arr = recruitSlotsMap.get(rKey);
          if (!arr.some(function(x) { return x.normLabel === parsedT.normLabel; })) {
            arr.push(parsedT);
          }
        });
      }
    } catch (e) {}
  }

  return { backupDocsByDate: backupDocsByDate, recruitSlotsMap: recruitSlotsMap };
}

function extractBackupDoctorsByDate64_(ss, doctorSearchList) {
  return extractBackupAndRecruitSlots64_(ss, doctorSearchList).backupDocsByDate;
}

/**
 * 実不在時間（例: 10:00-11:00）を内包する同時間帯バンドの募集枠（例: 09:00-11:30）があれば募集枠表記へ変換する
 */
function resolveRecruitDisplayTimeAndWage64_(sortDate, clinicClean, rawClinicTitle, timeAndWage, recruitSlotsMap) {
  if (!recruitSlotsMap || recruitSlotsMap.size === 0) {
    return { displayTimeAndWage: timeAndWage, rawAbsenceNote: '' };
  }
  const rawRange = extractSlotTimeRange64_(timeAndWage);
  const parsedAbs = parseTimeRangeToMins64_(rawRange);
  if (!parsedAbs) return { displayTimeAndWage: timeAndWage, rawAbsenceNote: '' };

  const dept = String(rawClinicTitle || '').includes('内科') ? '内科' : '小児科';
  const candidates = recruitSlotsMap.get(sortDate + '|' + clinicClean + '|' + dept) || [];
  if (candidates.length === 0) return { displayTimeAndWage: timeAndWage, rawAbsenceNote: '' };

  const absBand = getBandCodeFromMins64_(parsedAbs.startMin);
  const sameBandMatches = candidates.filter(function(c) {
    if (getBandCodeFromMins64_(c.startMin) !== absBand) return false;
    const cEndBand = c.endMin <= 840 ? 'A' : (c.endMin <= 1080 ? 'B' : 'C');
    if (cEndBand !== absBand) return false;
    return c.startMin <= parsedAbs.startMin && c.endMin >= parsedAbs.endMin;
  });

  if (sameBandMatches.length === 1 && sameBandMatches[0].normLabel !== parsedAbs.normLabel) {
    const recruitLabel = sameBandMatches[0].normLabel;
    const updatedTimeAndWage = String(timeAndWage || '').replace(
      /\d{1,2}:\d{2}\s*[\-〜～~]\s*\d{1,2}:\d{2}/,
      recruitLabel
    );
    return {
      displayTimeAndWage: updatedTimeAndWage,
      rawAbsenceNote: '・📌募集枠: ' + recruitLabel + '（実不在時間: ' + parsedAbs.normLabel + '）'
    };
  }

  return { displayTimeAndWage: timeAndWage, rawAbsenceNote: '' };
}