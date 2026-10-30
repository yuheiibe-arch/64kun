/**
 * ============================================================================
 * 6_編集くん64_バックエンド.gs
 * （NINTENDO64風大画面UI「編集くん」用バックエンド：
 *   拠点名マスタ連動・同時編集ロック・一時保存・差分ドキュメント反映）
 * ============================================================================
 */

const HENSHUKUN_CONFIG = {
  DRAFT_PROP_KEY_PREFIX: 'HENSHUKUN_DRAFT_',
  ACTIVE_USERS_CACHE_KEY: 'HENSHUKUN_ACTIVE_PLAYERS_V1',
  ACTIVE_USER_TTL_SEC: 45
};

/**
 * メニューから『🎮 編集くん 64』大画面モーダルを起動する
 */
function openHenshukunModal() {
  const html = HtmlService.createHtmlOutput(buildHenshukun64Html_())
    .setWidth(1240)
    .setHeight(840);
  SpreadsheetApp.getUi().showModalDialog(html, '🎮 編集くん 64 ― AI超募集MTG シフト確定スタジオ');
}

/**
 * 「2026/11/01（日）」や「2026/11/01」を「2026年11月1日 (日)」形式に変換する
 */
function toJapaneseDateLabel64_(rawDateStr) {
  const s = String(rawDateStr || '').trim();
  const m = s.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return s;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(y, mo - 1, d);
  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = isNaN(dt.getTime()) ? '' : ` (${dows[dt.getDay()]})`;
  return `${y}年${mo}月${d}日${dow}`;
}

function toSortDateKey64_(rawDateStr) {
  const m = String(rawDateStr || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return '9999/99/99';
  return `${m[1]}/${('0' + m[2]).slice(-2)}/${('0' + m[3]).slice(-2)}`;
}

/**
 * 『DSナレッジシート』内の『拠点名』シートおよび『担当者』シートを読み込む
 */
function loadClinicAndStaffMasters64_(ss) {
  const clinicMap = new Map(); // 拠点名 -> { group: '東京', area: '関東', dept: '小児科', no: 1 }
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

  // 『担当者』シート（DSナレッジシート内優先、なければ STAFF_SS_ID）から担当者リストを取得・ソート
  const staffList = [];
  const seenStaff = new Set();
  const readStaffSheet = (sh) => {
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

  staffList.sort((a, b) => a.localeCompare(b, 'ja'));

  return {
    clinicMap: clinicMap,
    groups: Array.from(groupSet),
    areaToGroups: areaToGroups,
    staffList: staffList
  };
}

/**
 * 拠点名からエリア（関東/関西）と地域グループ（東京/千葉/埼玉/神奈川/大阪...）を判定
 */
function resolveClinicMeta64_(rawClinicText, clinicMap) {
  const clean = String(rawClinicText || '')
    .replace(/【|】/g, '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/小児科|内科/g, '')
    .trim();

  if (clinicMap.has(clean)) {
    return clinicMap.get(clean);
  }
  for (const [k, v] of clinicMap.entries()) {
    if (clean.includes(k) || k.includes(clean)) return v;
  }

  const area = (typeof inferAreaForDoc_ === 'function') ? inferAreaForDoc_(clean) : '東京';
  const macroArea = (area === '関西') ? '関西' : '関東';
  const group = (area === '関西') ? '大阪' : area;
  return { name: clean, group: group, area: macroArea, no: 999 };
}

/**
 * 初期データ取得（ドキュメント全行＋一時保存ドラフト＋拠点マスタ＋担当者ロック状況）
 */
function getHenshukunInitialData() {
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  if (!docId) {
    throw new Error('対象ドキュメント（SINGLE_MTG_DOC_ID）が見つかりません。先にドキュメントを生成してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masters = loadClinicAndStaffMasters64_(ss);
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();

  // 一時保存ドラフトの読み込み
  const draftJson = props.getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId);
  const draftMap = draftJson ? JSON.parse(draftJson) : { slots: {}, bundles: {} };

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
          const docDec = row.getCell(2).getText().trim() || '案1';
          const docStaff = row.getCell(3).getText().trim() || '';

          const lines1 = col1.split('\n');
          const clinicTitle = lines1[0] || '';
          const timeAndWage = lines1.slice(1).join(' ') || '';
          const meta = resolveClinicMeta64_(clinicTitle, masters.clinicMap);

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
            clinicGroup: meta.group, // 東京・千葉・埼玉・神奈川・大阪...
            macroArea: meta.area,    // 関東・関西
            clinicNo: meta.no,
            timeAndWage: timeAndWage,
            col2Plans: col2,
            origDecision: docDec,
            origAssignee: docStaff,
            curDecision: draftItem ? draftItem.decision : docDec,
            curAssignee: draftItem ? draftItem.assignee : docStaff,
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

  // 日付順 -> クリニックNo順でソート
  slotRows.sort((a, b) => a.sortDate.localeCompare(b.sortDate) || a.clinicNo - b.clinicNo);

  const clientSessionId = Utilities.getUuid().slice(0, 8);
  const activeLocks = getActivePlayerLocks64_();

  return {
    clientSessionId: clientSessionId,
    docId: docId,
    docName: doc.getName(),
    docUrl: doc.getUrl(),
    staffList: masters.staffList,
    groups: masters.groups,
    areaToGroups: masters.areaToGroups,
    slotRows: slotRows,
    bundleRows: bundleRows,
    activeLocks: activeLocks
  };
}

/**
 * 同時編集中のプレイヤー（担当者）ロック管理（ハートビート）
 */
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
    // 他のセッションがすでに同じ担当者名をロックしていないかチェック
    for (const [sid, info] of Object.entries(locks)) {
      if (sid !== sessionId && info.staff === staffName) {
        return {
          ok: false,
          conflictStaff: staffName,
          activeLocks: locks
        };
      }
    }
    locks[sessionId] = { staff: staffName, ts: Date.now() };
  } else {
    delete locks[sessionId];
  }

  cache.put(HENSHUKUN_CONFIG.ACTIVE_USERS_CACHE_KEY, JSON.stringify(locks), 120);

  // 他プレイヤーが更新した一時保存ドラフトがあれば返す（リアルタイム同期用）
  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  const draftJson = docId ? props.getProperty(HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId) : null;

  return {
    ok: true,
    activeLocks: locks,
    latestDraft: draftJson ? JSON.parse(draftJson) : null
  };
}

/**
 * 一時保存（ドラフト保存：複数人の変更をセル単位で安全にマージ）
 */
function saveHenshukunDraft(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const docId = props.getProperty('SINGLE_MTG_DOC_ID');
    if (!docId) throw new Error('ドキュメントIDが見つかりません。');

    const key = HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId;
    const existingRaw = props.getProperty(key);
    const draft = existingRaw ? JSON.parse(existingRaw) : { slots: {}, bundles: {} };

    Object.entries(payload.changedSlots || {}).forEach(([slotKey, val]) => {
      draft.slots[slotKey] = {
        decision: val.decision,
        assignee: val.assignee,
        updatedAt: Date.now()
      };
    });

    Object.entries(payload.changedBundles || {}).forEach(([bKey, val]) => {
      draft.bundles[bKey] = val;
    });

    props.setProperty(key, JSON.stringify(draft));
    return { ok: true, savedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss') };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ドキュメント本体への反映（同時編集で他人の箇所を潰さないよう、変更された行＋ドラフト分のみをマージ反映）
 */
function commitHenshukunToDocAndSubmit(payload, shouldSubmitToMnys) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // まずドラフトにもマージ保存
    saveHenshukunDraft(payload);

    const props = PropertiesService.getScriptProperties();
    const docId = props.getProperty('SINGLE_MTG_DOC_ID');
    const draftKey = HENSHUKUN_CONFIG.DRAFT_PROP_KEY_PREFIX + docId;
    const draftRaw = props.getProperty(draftKey);
    const draft = draftRaw ? JSON.parse(draftRaw) : { slots: {}, bundles: {} };

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 全テーブルの不要表記クレンジング＋ドラフト変更分の反映
    Object.entries(draft.slots || {}).forEach(([slotKey, item]) => {
      const [cIdx, rIdx] = slotKey.split('_').map(Number);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            row.getCell(0).setText(cleanSidebarDisplayText_(row.getCell(0).getText()));
            row.getCell(1).setText(cleanSidebarDisplayText_(row.getCell(1).getText()));
            row.getCell(2).setText(String(item.decision || '').trim());
            row.getCell(3).setText(String(item.assignee || '').trim());
            row.getCell(0).editAsText().setFontSize(8.5);
            row.getCell(1).editAsText().setFontSize(8.5);
            row.getCell(2).editAsText().setFontSize(8.5).setBold(true);
            row.getCell(3).editAsText().setFontSize(8.5);
          }
        }
      }
    });

    Object.entries(draft.bundles || {}).forEach(([bKey, val]) => {
      const parts = bKey.split('_');
      const cIdx = Number(parts[1]);
      const rIdx = Number(parts[2]);
      const el = body.getChild(cIdx);
      if (el && el.getType() === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        if (rIdx < table.getNumRows()) {
          const row = table.getRow(rIdx);
          if (row.getNumCells() >= 4) {
            row.getCell(0).setText(cleanSidebarDisplayText_(row.getCell(0).getText()));
            row.getCell(3).setText(String(val || '').trim());
            row.getCell(0).editAsText().setFontSize(8.5);
            row.getCell(3).editAsText().setFontSize(8.5).setBold(true);
          }
        }
      }
    });

    doc.saveAndClose();
    // ドキュメント本体に反映完了したらドラフトをクリア
    props.deleteProperty(draftKey);

    if (shouldSubmitToMnys) {
      submitEditedDocToMnys();
    }

    return {
      ok: true,
      committedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss')
    };
  } finally {
    lock.releaseLock();
  }
}