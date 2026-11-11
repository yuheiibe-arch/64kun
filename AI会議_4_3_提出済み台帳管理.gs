/**
 * ============================================================================
 * AI会議_4_3_提出済み台帳管理.gs（新規追加）
 * ・mnysへ提出済みのアプローチプランを『提出済みプラン台帳』シートに永続記録
 * ・次回MTG生成時の「未送信プランのゼロリセット防止（100%引継）」
 * ・編集くん64での「前回指示済（進行中）ロック」と「差分のみ提出」を制御
 * ※「ステイ／特別時給＋ステイ」は次回見直し対象とするため台帳ロックしない設計
 * ============================================================================
 */

const SUBMITTED_LEDGER_SHEET_NAME = '提出済みプラン台帳';

function makeSubmittedSlotKey4_3_(dateStr, rawClinic, deptHint, rawTime) {
  const clinicWithDept = (String(deptHint || '').includes('内科') && !String(rawClinic || '').includes('内科'))
    ? (rawClinic + '（内科）')
    : rawClinic;
  return getSlotCarryOverKey3_(dateStr, clinicWithDept, rawTime);
}

function buildFallbackCustomPhraseFromItems4_3_(items) {
  return (items || []).map(function(it) {
    const doc = it.target_doctor ? (it.target_doctor + '先生') : '';
    if (it.approach_type === 'バックアップ勤務依頼') return 'バックアップ医師に勤務依頼';
    if (it.approach_type === '勤務地変更相談') {
      return (it.from_clinic ? (it.from_clinic + 'から') : '') + (doc ? (doc + 'へ') : '') + '移動打診';
    }
    if (it.approach_type === '延長相談') return (doc ? (doc + 'へ') : '') + '延長打診';
    if (it.approach_type === '早出相談') return (doc ? (doc + 'へ') : '') + '早出打診';
    return (doc ? (doc + 'へ') : '') + '個別打診';
  }).filter(Boolean).join(' / ');
}

/**
 * 『提出済みプラン台帳』を読み込み、今日以降の有効レコードを carryKey 単位の Map で返す
 * （初回導入時、台帳が空なら既存の『ドキュメント確定出力JSON』シートから自動復元する）
 */
function loadSubmittedLedgerMap4_3_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const map = new Map();
  if (!ss) return map;

  const todayYmd = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const sh = ss.getSheetByName(SUBMITTED_LEDGER_SHEET_NAME);

  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getDataRange().getDisplayValues();
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const key = String(row[0] || '').trim();
      const dateStr = normalizeDateYmd1b3_(row[1]);
      if (!key || !dateStr || dateStr < todayYmd) continue;

      let items = [];
      try { items = JSON.parse(row[9] || '[]'); } catch (e) { items = []; }
      const docs = String(row[6] || '').split(/[・,、\/]+/).map(function(s) { return s.trim(); }).filter(Boolean);

      map.set(key, {
        carryKey: key,
        date: dateStr,
        clinicName: String(row[2] || '').trim(),
        dept: String(row[3] || '小児科').trim(),
        targetTime: String(row[4] || '').trim(),
        decisionHeader: String(row[5] || '').trim(),
        doctors: docs,
        assignee: String(row[7] || '').trim(),
        comment: String(row[8] || '').trim(),
        items: items,
        submittedAt: String(row[10] || '').trim()
      });
    }
    return map;
  }

  // 初回フォールバック：『ドキュメント確定出力JSON』が残っていればそこから初期構築
  const outSh = ss.getSheetByName('ドキュメント確定出力JSON');
  if (outSh && outSh.getLastRow() >= 2) {
    const outVals = outSh.getDataRange().getDisplayValues();
    for (let r = 1; r < outVals.length; r++) {
      const orow = outVals[r];
      const d = normalizeDateYmd1b3_(orow[0]);
      if (!d || d < todayYmd) continue;
      const cName = String(orow[1] || '').trim();
      const dept = String(orow[2] || '小児科').trim();
      const tRange = String(orow[3] || '').trim();
      const appType = String(orow[4] || '').trim();
      if (!cName || appType.includes('ステイ')) continue;

      const rawDoc = String(orow[5] || '').replace(/\(常勤\)|\[★.*?\]/g, '').trim();
      const lic = String(orow[6] || '').trim();
      const fromC = String(orow[7] || '').trim();
      const asg = String(orow[9] || '').trim();
      const key = makeSubmittedSlotKey4_3_(d, cName, dept, tRange);

      if (!map.has(key)) {
        map.set(key, {
          carryKey: key,
          date: d,
          clinicName: cName,
          dept: dept,
          targetTime: tRange,
          decisionHeader: '',
          doctors: [],
          assignee: asg,
          comment: '',
          items: [],
          submittedAt: '前回提出分'
        });
      }
      const entry = map.get(key);
      if (rawDoc && !entry.doctors.includes(rawDoc)) entry.doctors.push(rawDoc);
      entry.items.push({
        date: d,
        clinic_name: cName,
        department: dept,
        target_time: tRange,
        approach_type: appType,
        target_doctor: rawDoc,
        license_no: lic,
        from_clinic: fromC,
        assignee: asg
      });
    }
    map.forEach(function(entry) {
      if (!entry.decisionHeader) {
        entry.decisionHeader = buildFallbackCustomPhraseFromItems4_3_(entry.items);
      }
    });
  }

  return map;
}

/**
 * mnys提出時に、今回新たに提出した枠を『提出済みプラン台帳』へアップサートし、
 * 過去日および既に解消された枠をクリーンアップする
 */
function upsertSubmittedLedgerRecords4_3_(ss, newSlotRecords, activeSlotKeysOpt) {
  let sh = ss.getSheetByName(SUBMITTED_LEDGER_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SUBMITTED_LEDGER_SHEET_NAME);

  const existingMap = loadSubmittedLedgerMap4_3_(ss);
  const todayYmd = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  // ドキュメント上に存在しない枠（＝充足等で既に消滅した枠）が判明している場合は整理
  if (activeSlotKeysOpt && activeSlotKeysOpt.size > 0) {
    Array.from(existingMap.keys()).forEach(function(k) {
      if (!activeSlotKeysOpt.has(k)) {
        existingMap.delete(k);
      }
    });
  }

  (newSlotRecords || []).forEach(function(rec) {
    if (!rec || !rec.carryKey || !rec.items || rec.items.length === 0) return;
    if (isStayOrSkipDecision4_2_(rec.decisionHeader)) return; // ステイは記録しない（次回見直し）

    const docs = Array.from(new Set(rec.items.map(function(x) { return x.target_doctor; }).filter(Boolean)));
    existingMap.set(rec.carryKey, {
      carryKey: rec.carryKey,
      date: rec.date,
      clinicName: rec.clinicName,
      dept: rec.dept,
      targetTime: rec.targetTime,
      decisionHeader: rec.decisionHeader || buildFallbackCustomPhraseFromItems4_3_(rec.items),
      doctors: docs,
      assignee: rec.assignee || '未定',
      comment: rec.comment || '',
      items: rec.items,
      submittedAt: nowStr
    });
  });

  const sorted = Array.from(existingMap.values())
    .filter(function(e) { return e.date >= todayYmd; })
    .sort(function(a, b) {
      return a.date.localeCompare(b.date) || a.carryKey.localeCompare(b.carryKey);
    });

  sh.clearContents();
  const headers = [
    '枠キー', '日付', '拠点名', '診療科', '不在時間',
    '確定案ヘッダー', '対象医師一覧', '担当者', 'コメント', '提出アイテムJSON', '提出日時'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  if (sorted.length > 0) {
    const rows = sorted.map(function(e) {
      return [
        e.carryKey,
        e.date,
        e.clinicName,
        e.dept,
        e.targetTime,
        e.decisionHeader,
        (e.doctors || []).join('・'),
        e.assignee,
        String(e.comment || '').slice(0, 10000),
        JSON.stringify(e.items || []).slice(0, 45000),
        e.submittedAt
      ];
    });
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

/**
 * 今回ドキュメント上の入力内容が、前回提出済みの台帳レコードから実質変更されていないかを判定する
 * （true ＝ 前回提出済みと同一のため今回の mnys 再提出から除外する）
 */
function isSameAsSubmittedLedger4_3_(ledgerEntry, currentHeaderDec, currentAssignee, currentItems, arrowLines) {
  if (!ledgerEntry || !currentItems || currentItems.length === 0) return false;

  // 1) 前回打診した医師全員が「❌辞退」になっている場合は再検討対象なので同一扱いしない
  const declinedKeys = new Set();
  (arrowLines || []).forEach(function(l) {
    if (!l.includes('❌') && !l.includes('辞退')) return;
    const m = l.match(/❌\s*(?:辞退[：:])?\s*([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})/);
    if (m) declinedKeys.add(normKeyForDoc_(m[1]));
  });

  const ledgerDocKeys = (ledgerEntry.doctors || []).map(normKeyForDoc_).filter(Boolean);
  if (ledgerDocKeys.length > 0 && ledgerDocKeys.every(function(dk) { return declinedKeys.has(dk); })) {
    return false;
  }

  // 2) 担当者が変更されていたら再提出する
  const normAsg = function(a) {
    const s = String(a || '').trim();
    return (!s || s === '担当者未定' || s === '(未定)') ? '未定' : s;
  };
  if (normAsg(ledgerEntry.assignee) !== normAsg(currentAssignee)) {
    return false;
  }

  // 3) 抽出された対象医師・アプローチ種別の集合が一致しているか比較
  const sig = function(list) {
    return (list || []).map(function(x) {
      return normKeyForDoc_(x.target_doctor || '') + ':' + String(x.approach_type || '');
    }).sort().join('|');
  };

  return sig(ledgerEntry.items) === sig(currentItems);
}

/**
 * 編集くん64の初期ロード用：その行が「✅前回指示済（対応不要）」か「🔄全員辞退で要再選定」かを判定する
 */
function evaluateRowSubmittedStatusFor64_(carryKey, rawDocDec, docStaff, arrowLines, ledgerMap) {
  const entry = ledgerMap ? ledgerMap.get(carryKey) : null;
  if (!entry) return { isPreviouslySubmitted: false, allDeclinedRevived: false, submittedAt: '' };

  if (!rawDocDec || isStayOrSkipDecision4_2_(rawDocDec)) {
    return { isPreviouslySubmitted: false, allDeclinedRevived: false, submittedAt: '' };
  }

  const isAccepted = (arrowLines || []).some(function(l) {
    return l.includes('⭕') && (l.includes('承諾') || l.includes('充足'));
  });
  if (isAccepted) {
    return { isPreviouslySubmitted: true, allDeclinedRevived: false, submittedAt: entry.submittedAt || '' };
  }

  const declinedKeys = new Set();
  (arrowLines || []).forEach(function(l) {
    if (!l.includes('❌') && !l.includes('辞退')) return;
    const m = l.match(/❌\s*(?:辞退[：:])?\s*([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})/);
    if (m) declinedKeys.add(normKeyForDoc_(m[1]));
  });

  const ledgerDocKeys = (entry.doctors || []).map(normKeyForDoc_).filter(Boolean);
  if (ledgerDocKeys.length > 0 && ledgerDocKeys.every(function(dk) { return declinedKeys.has(dk); })) {
    return { isPreviouslySubmitted: false, allDeclinedRevived: true, submittedAt: entry.submittedAt || '' };
  }

  return {
    isPreviouslySubmitted: true,
    allDeclinedRevived: false,
    submittedAt: entry.submittedAt || ''
  };
}