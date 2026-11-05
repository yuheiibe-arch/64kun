/**
 * ============================================================================
 * AI会議_1b2_2_キャンセルとメール制約.gs
 * 1) 【設計1】シフトキャンセル承認メールの自動同期＆当日全枠の提案ブロック
 * 2) 【設計2】医師返信メール（ナレッジシート＆未マッチ医師質問）からの
 *    日別・時間上限（例: 11/8 柿本先生の「18時までなら可能」＝18時以降の延長不可）＆単発辞退の自動抽出
 * ============================================================================
 */

/**
 * 設計1：『シフトキャンセル履歴』シート（未同期分があればGmailからも補完）を読み込み、
 * 医師キーごとのキャンセル対象日セット Map(docKey -> Set("2026/11/15", ...)) を返す
 */
function loadShiftCancelDatesByDoctor1b2_(ss) {
  const cancelMap = new Map();

  // もし『シフトキャンセル履歴』シートがまだ無いか空なら、Gmailから直近30日分を即時同期
  let cancelSh = ss.getSheetByName('シフトキャンセル履歴');
  if (!cancelSh || cancelSh.getLastRow() <= 1) {
    try {
      if (typeof getAnswerEmailsFromGmail_Final === 'function') {
        getAnswerEmailsFromGmail_Final(ss.getId());
        cancelSh = ss.getSheetByName('シフトキャンセル履歴');
      }
    } catch (e) {}
  }

  if (cancelSh && cancelSh.getLastRow() >= 2) {
    const vals = cancelSh.getRange(2, 1, cancelSh.getLastRow() - 1, 7).getDisplayValues();
    vals.forEach(function(r) {
      const docKey = normKeyForDoc_(r[2]);
      const dateStr = normalizeDateYmd1b3_(r[3]);
      if (!docKey || !dateStr) return;
      if (!cancelMap.has(docKey)) cancelMap.set(docKey, new Set());
      cancelMap.get(docKey).add(dateStr);
    });
  }

  return cancelMap;
}

/**
 * テキスト内から対象日付（YYYY/MM/DD）をすべて抽出するヘルパー
 * （年が省略されている「11月8日」「11/8」などは msgTimestamp の年から補完）
 */
function extractTargetDatesFromMailText1b2_(text, msgTimestampStr) {
  const s = String(text || '');
  if (!s) return [];

  let baseYear = new Date().getFullYear();
  let baseMonth = new Date().getMonth() + 1;
  const tsMatch = String(msgTimestampStr || '').match(/(\d{4})[\/\-](\d{1,2})/);
  if (tsMatch) {
    baseYear = Number(tsMatch[1]);
    baseMonth = Number(tsMatch[2]);
  }

  const dates = [];
  const addYmd = function(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > 31) return;
    const ymd = y + '/' + ('0' + m).slice(-2) + '/' + ('0' + d).slice(-2);
    if (!dates.includes(ymd)) dates.push(ymd);
  };

  // 1. 年付き日付（例：2026年11月8日、2026/11/08）
  const fullRe = /(\d{4})\s*[年\/\-]\s*(\d{1,2})\s*[月\/\-]\s*(\d{1,2})\s*日?/g;
  let fm;
  while ((fm = fullRe.exec(s)) !== null) {
    addYmd(Number(fm[1]), Number(fm[2]), Number(fm[3]));
  }

  // 2. 月日のみ（例：11月8日(日)、8/29）※時刻（15:00等）や管理番号と誤認しないようフィルタ
  const mdRe = /(?:^|[^\d\/年])(\d{1,2})\s*(?:月\s*(\d{1,2})\s*日|\/\s*(\d{1,2})(?=[\s\(（日月火水木金土・、,のをにへ]|$))/g;
  let mm;
  while ((mm = mdRe.exec(s)) !== null) {
    const m = Number(mm[1]);
    const d = Number(mm[2] || mm[3]);
    let y = baseYear;
    if (baseMonth >= 10 && m <= 3) y = baseYear + 1;
    addYmd(y, m, d);
  }

  // 3. 「明日」と書かれている場合、メール送信日の翌日を追加
  if (s.includes('明日') && msgTimestampStr) {
    const dtMatch = String(msgTimestampStr).match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (dtMatch) {
      const nextDt = new Date(Number(dtMatch[1]), Number(dtMatch[2]) - 1, Number(dtMatch[3]) + 1);
      addYmd(nextDt.getFullYear(), nextDt.getMonth() + 1, nextDt.getDate());
    }
  }

  return dates;
}

/**
 * 設計2：『ナレッジシート』および未マッチの『医師質問』『DS回答』を解析し、
 * 日別・医師別の時間上限（例：11/8の柿本先生は18:00まで＝18:00以降の延長・夜間NG）や単発辞退を抽出する関数
 * 戻り値: Map( "YYYY/MM/DD_docKey" -> { maxEndMin: 1080, minStartMin: 0, noExtend: true, noEarly: false, noMove: false, fullDayDeclined: false, note: "..." } )
 */
function buildDateSpecificMailConstraintsMap1b2_(ss) {
  const constraintMap = new Map();

  const getRecentRows = function(shName, count) {
    const sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return [];
    const lastR = sh.getLastRow();
    const startR = Math.max(2, lastR - count + 1);
    const cols = Math.min(8, sh.getLastColumn());
    return sh.getRange(startR, 1, lastR - startR + 1, cols).getDisplayValues();
  };

  const kbRows = getRecentRows('ナレッジシート', 350);
  const qRows = getRecentRows('医師質問', 250);
  const aRows = getRecentRows('DS回答', 250);

  // 同じ医師に対してスタッフが打診した対象日付の履歴（複数往復ラリーで2通目以降に日付が書かれていない場合の引き継ぎ用）
  const recentPitchDatesByDoc = new Map(); // docKey -> [{ tsMs, dates: [...] }]

  const registerPitchDates = function(docKey, pitchText, tsStr) {
    if (!docKey || !pitchText) return;
    const extracted = extractTargetDatesFromMailText1b2_(pitchText, tsStr);
    if (extracted.length === 0) return;
    const tsMs = tsStr ? new Date(tsStr).getTime() : Date.now();
    if (!recentPitchDatesByDoc.has(docKey)) recentPitchDatesByDoc.set(docKey, []);
    recentPitchDatesByDoc.get(docKey).push({ tsMs: isNaN(tsMs) ? 0 : tsMs, dates: extracted });
  };

  aRows.forEach(function(r) {
    registerPitchDates(normKeyForDoc_(r[2]), r[3], r[1]);
  });
  kbRows.forEach(function(r) {
    registerPitchDates(normKeyForDoc_(r[1]), r[3], r[6]);
  });

  const applyDoctorReplyConstraint = function(docKey, docReplyText, staffPitchText, tsStr) {
    if (!docKey || !docReplyText) return;
    const reply = String(docReplyText).replace(/^【医師返信】\s*/, '').trim();
    const pitch = String(staffPitchText || '').replace(/^【スタッフ打診】\s*/, '').trim();

    // 1. 時間上限（「18時までなら勤務可能」「18時まででしたら」「18時以降は難しい」）の検出
    let maxEndMin = null;
    const untilMatch = reply.match(/(\d{1,2})\s*(?:時|:00)\s*(?:までなら|まででしたら|までであれば|までしか|までと|で上が|で退勤)/);
    const afterNgMatch = reply.match(/(\d{1,2})\s*(?:時|:00)\s*以降は\s*(?:難|不可|厳|無理|できかね)/);
    if (untilMatch) {
      const h = Number(untilMatch[1]);
      if (h >= 12 && h <= 21) maxEndMin = h * 60;
    } else if (afterNgMatch) {
      const h = Number(afterNgMatch[1]);
      if (h >= 12 && h <= 21) maxEndMin = h * 60;
    }

    // 2. 開始下限（「10時からなら」「19時からの勤務になります」「早く入ることが難しく」）の検出
    let minStartMin = null;
    const fromMatch = reply.match(/(\d{1,2})\s*(?:時|:00)\s*(?:からなら|からでしたら|からであれば|からの勤務になります)/);
    if (fromMatch) {
      const h = Number(fromMatch[1]);
      if (h >= 9 && h <= 20) minStartMin = h * 60;
    }

    const noExtendExplicit = /延長は(?:難|不可|厳|無理)|延長できかね/.test(reply);
    const noEarlyExplicit = /早出は(?:難|不可|厳|無理)|前倒しは(?:難|不可|厳|無理)|早く入ることが難/.test(reply);
    const noMoveExplicit = /移動は(?:難|不可|厳|無理)|拠点変更は(?:難|不可|厳)/.test(reply);
    const isDecline = /勤務困難|難しいです|お役に立てず|都合がつかず|お力になれず|可能な日はありません/.test(reply) &&
                      !/までなら|まででしたら|からなら|どちらも可能/.test(reply);

    if (maxEndMin === null && minStartMin === null && !noExtendExplicit && !noEarlyExplicit && !noMoveExplicit && !isDecline) {
      return;
    }

    // 3. この返信が対象としている日付（YYYY/MM/DD）を特定
    //    ①医師返信文内の日付 ➔ ②ペアのスタッフ打診文内の日付 ➔ ③同医師への直近（前後48時間以内）の打診文内の日付 の順で継承！
    let targetDates = extractTargetDatesFromMailText1b2_(reply, tsStr);
    if (targetDates.length === 0 && pitch) {
      targetDates = extractTargetDatesFromMailText1b2_(pitch, tsStr);
    }
    if (targetDates.length === 0 && recentPitchDatesByDoc.has(docKey)) {
      const replyMs = tsStr ? new Date(tsStr).getTime() : 0;
      const history = recentPitchDatesByDoc.get(docKey);
      for (let i = history.length - 1; i >= 0; i--) {
        if (!replyMs || !history[i].tsMs || Math.abs(replyMs - history[i].tsMs) <= 48 * 3600 * 1000) {
          targetDates = history[i].dates;
          break;
        }
      }
    }

    targetDates.forEach(function(dStr) {
      const key = dStr + '_' + docKey;
      const cur = constraintMap.get(key) || {
        maxEndMin: 9999,
        minStartMin: 0,
        noExtend: false,
        noEarly: false,
        noMove: false,
        fullDayDeclined: false,
        note: ''
      };

      if (maxEndMin !== null && maxEndMin < cur.maxEndMin) {
        cur.maxEndMin = maxEndMin;
        cur.noExtend = true;
        cur.note = (maxEndMin / 60) + '時までなら可能（' + (maxEndMin / 60) + '時以降の延長・夜間不可）';
      }
      if (minStartMin !== null && minStartMin > cur.minStartMin) {
        cur.minStartMin = minStartMin;
        cur.noEarly = true;
        cur.note = (minStartMin / 60) + '時から可能（早出不可）';
      }
      if (noExtendExplicit) cur.noExtend = true;
      if (noEarlyExplicit) cur.noEarly = true;
      if (noMoveExplicit) cur.noMove = true;
      if (isDecline && maxEndMin === null && minStartMin === null) {
        cur.fullDayDeclined = true;
        cur.note = 'メールにて当日辞退済';
      }

      constraintMap.set(key, cur);
    });
  };

  // ナレッジシート（ペア成立分）と医師質問シート（未マッチ単独分含む）の両方をスキャンしてロストを防止
  kbRows.forEach(function(r) {
    applyDoctorReplyConstraint(normKeyForDoc_(r[1]), r[2], r[3], r[6]);
  });
  qRows.forEach(function(r) {
    applyDoctorReplyConstraint(normKeyForDoc_(r[2]), r[3], '', r[1]);
  });

  return constraintMap;
}

/**
 * 🔍 動作確認用テスト関数（実行すると柿本先生や加藤雅崇先生のブロック判定結果をログ出力します）
 */
function testVerifyDesign1And2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cancelMap = loadShiftCancelDatesByDoctor1b2_(ss);
  const mailConstraintMap = buildDateSpecificMailConstraintsMap1b2_(ss);

  Logger.log("=== ✅ 【設計1】シフトキャンセルによる日別ブロック確認 ===");
  Logger.log("・加藤雅崇先生のキャンセル日一覧: " + JSON.stringify(Array.from(cancelMap.get(normKeyForDoc_("加藤雅崇")) || [])));

  Logger.log("\n=== ✅ 【設計2】メール返信による日別・時間上限ブロック確認 ===");
  const kakimotoKey = "2026/11/08_" + normKeyForDoc_("柿本優");
  Logger.log("・2026/11/08 柿本優先生のメール制約判定: " + JSON.stringify(mailConstraintMap.get(kakimotoKey) || "なし"));
}