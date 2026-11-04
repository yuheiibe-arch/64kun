/**
 * ============================================================================
 * 6b_編集くん64_データ読込と同期.gs
 * （A:09-13 / B:15-18 / C:18-21 の3時間帯別バックアップ医師抽出・カスタムコメント対応版）
 * ============================================================================
 */

function extractBackupDoctorsByDate64_(ss, doctorSearchList) {
  const backupDocsByDate = {};
  const sh = ss.getSheetByName('最新_当日シフト');
  if (!sh || sh.getLastRow() < 2) return backupDocsByDate;

  const docLicMap = new Map();
  (doctorSearchList || []).forEach(function(d) {
    if (d.norm && d.licenseNo) docLicMap.set(d.norm, d.licenseNo);
  });

  const vals = sh.getDataRange().getDisplayValues();
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const clinicRaw = String(row[2] || '').trim();
    if (!clinicRaw.includes('バックアップ')) continue;

    const sortDate = toSortDateKey64_(row[0]);
    if (!sortDate || sortDate === '9999/99/99') continue;
    if (!backupDocsByDate[sortDate]) {
      backupDocsByDate[sortDate] = { A: [], B: [], C: [] };
    }

    const bandCols = [
      { band: 'A', text: row[7] }, // 09:00-13:00
      { band: 'B', text: row[8] }, // 15:00-18:00
      { band: 'C', text: row[9] }  // 18:00-21:00
    ];

    bandCols.forEach(function(bc) {
      const tokens = (typeof splitDoctorTokensSafely_ === 'function')
        ? splitDoctorTokensSafely_(bc.text)
        : String(bc.text || '').split(/[,、，\/]+/);

      tokens.forEach(function(tok) {
        const cleanName = String(tok || '')
          .replace(/[\(（\[【].*$/g, '')
          .replace(/[:：].*$/, '')
          .replace(/[\)）\]】★⚠️📩・]/g, '')
          .replace(/先生$/g, '')
          .replace(/[\s\u3000]+/g, '')
          .trim();
        if (!cleanName || cleanName === '未充足' || cleanName === 'なし' || cleanName === '-') return;

        const norm = toSearchNorm64_(cleanName);
        if (!norm) return;

        const bandList = backupDocsByDate[sortDate][bc.band];
        const exists = bandList.some(function(d) { return d.norm === norm; });
        if (!exists) {
          bandList.push({
            name: cleanName,
            norm: norm,
            licenseNo: docLicMap.get(norm) || ''
          });
        }
      });
    });
  }
  return backupDocsByDate;
}

function getHenshukunInitialData() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    throw new Error('対象ドキュメント（SINGLE_MTG_DOC_ID）が見つかりません。先にドキュメントを生成してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masters = loadClinicAndStaffMasters64_(ss);
  const backupDocsByDate = extractBackupDoctorsByDate64_(ss, masters.doctorSearchList);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  const draftMap = loadLatestDraft64_(ss, docId);

  let currentDateRaw = '';
  const slotRows = [];
  const bundleRows = [];

  for (let i = 0; i < numChildren; i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const txt = el.asParagraph().getText().trim();
      const m = txt.match(/^◆\s*(\d{4}[\/\-年]\d{1,2}[\/\-月]\d{1,2}.*)$/);
      if (m) currentDateRaw = m[1].trim();
    } else if (el.getType() === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      if (table.getNumRows() < 2) continue;
      const h0 = table.getRow(0).getCell(0).getText().trim();

      if (h0.includes('拠点・不在時間帯') && currentDateRaw) {
        const sortDate = toSortDateKey64_(currentDateRaw);
        const jpDate = toJapaneseDateLabel64_(currentDateRaw);

        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;

          const col1 = cleanSidebarDisplayText_(row.getCell(0).getText().trim());
          const col2 = cleanSidebarDisplayText_(row.getCell(1).getText().trim());
          const rawDocDec = row.getCell(2).getText().trim();
          const docStaff = row.getCell(3).getText().trim() || '';

          let parsedDec = (!docStaff && (rawDocDec === '案1' || rawDocDec === '要検討')) ? '' : rawDocDec;
          let parsedComment = '';
          const cmMatch = parsedDec.match(/^(.*?)\s*[（\(]メモ[：:](.+?)[）\)]$/);
          if (cmMatch) {
            parsedDec = cmMatch[1].trim();
            parsedComment = cmMatch[2].trim();
          }

          const lines1 = col1.split('\n');
          const rawClinicTitle = lines1[0] || '';
          const clinicTitle = formatClinicTitle64_(rawClinicTitle);
          const timeAndWage = lines1.slice(1).join(' ') || '';
          const meta = resolveClinicMeta64_(rawClinicTitle, masters.clinicMap);

          const slotKey = `${i}_${r}`;
          const draftItem = (draftMap.slots && draftMap.slots[slotKey]) ? draftMap.slots[slotKey] : null;

          slotRows.push({
            slotKey: slotKey,
            childIdx: i,
            rowIdx: r,
            sortDate: sortDate,
            dateLabel: jpDate,
            clinicTitle: clinicTitle,
            clinicNameClean: meta.name,
            clinicGroup: meta.group,
            macroArea: meta.area,
            clinicNo: meta.no,
            timeAndWage: timeAndWage,
            col2Plans: col2,
            origDecision: parsedDec,
            origAssignee: docStaff,
            origComment: parsedComment,
            curDecision: draftItem ? draftItem.decision : parsedDec,
            curAssignee: draftItem ? draftItem.assignee : docStaff,
            customComment: (draftItem && typeof draftItem.customComment === 'string') ? draftItem.customComment : parsedComment,
            customDocs: (draftItem && Array.isArray(draftItem.customDocs)) ? draftItem.customDocs : [],
            hasDraft: Boolean(draftItem)
          });
        }
      } else if (h0.includes('対象医師')) {
        for (let r = 1; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          if (row.getNumCells() < 4) continue;
          const bKey = `b_${i}_${r}`;
          const docVal = row.getCell(3).getText().trim() || '採用';
          const draftVal = (draftMap.bundles && draftMap.bundles[bKey]) ? draftMap.bundles[bKey] : null;

          bundleRows.push({
            bundleKey: bKey,
            childIdx: i,
            rowIdx: r,
            doctor: cleanSidebarDisplayText_(row.getCell(0).getText().trim()),
            slots: row.getCell(1).getText().trim(),
            note: row.getCell(2).getText().trim(),
            origVal: docVal,
            curVal: draftVal || docVal
          });
        }
      }
    }
  }

  slotRows.sort((a, b) => a.sortDate.localeCompare(b.sortDate) || a.clinicNo - b.clinicNo);

  const clientSessionId = Utilities.getUuid().slice(0, 8);
  const activeLocks = getActivePlayerLocks64_();

  return {
    clientSessionId: clientSessionId,
    docId: docId,
    docName: doc.getName(),
    docUrl: doc.getUrl(),
    staffList: masters.staffList,
    doctorSearchList: masters.doctorSearchList,
    backupDocsByDate: backupDocsByDate,
    groups: masters.groups,
    areaToGroups: masters.areaToGroups,
    slotRows: slotRows,
    bundleRows: bundleRows,
    activeLocks: activeLocks
  };
}

function getActivePlayerLocks64_() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(HENSHUKUN_CONFIG.ACTIVE_USERS_CACHE_KEY);
  if (!raw) return {};
  try {
    const obj = JSON.parse(raw);
    const now = Date.now();
    const valid = {};
    Object.keys(obj).forEach(sessionId => {
      if (now - obj[sessionId].ts < HENSHUKUN_CONFIG.ACTIVE_USER_TTL_SEC * 1000) {
        valid[sessionId] = obj[sessionId];
      }
    });
    return valid;
  } catch (e) {
    return {};
  }
}

function heartbeatHenshukunPlayer(sessionId, staffName) {
  const cache = CacheService.getScriptCache();
  const locks = getActivePlayerLocks64_();

  if (staffName) {
    for (const [sid, info] of Object.entries(locks)) {
      if (sid !== sessionId && info.staff === staffName) {
        return { ok: false, conflictStaff: staffName, activeLocks: locks };
      }
    }
    locks[sessionId] = { staff: staffName, ts: Date.now() };
  } else {
    delete locks[sessionId];
  }

  cache.put(HENSHUKUN_CONFIG.ACTIVE_USERS_CACHE_KEY, JSON.stringify(locks), 120);

  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  const draftJson = docId ? props.getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId) : null;

  return {
    ok: true,
    activeLocks: locks,
    latestDraft: draftJson ? JSON.parse(draftJson) : null
  };
}