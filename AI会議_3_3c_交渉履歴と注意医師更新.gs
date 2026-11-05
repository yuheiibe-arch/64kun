/**
 * ============================================================================
 * AI会議_3_3c_交渉履歴と注意医師更新.gs
 * 1) 『交渉履歴』シートを相談日順で自動生成・更新（成功/失敗/条件・時給区分・依頼項目・簡略返答）
 * 2) 『交渉注意医師』シートは2行目のプルダウンをコピーして整理し、K列は空欄維持・P列(16列目)に簡潔な要約を記載
 * ============================================================================
 */

/**
 * ★今すぐ単体で『交渉履歴』シートと『交渉注意医師』シートを整理・更新する実行関数
 */
function runUpdateNegotiationHistorySheetNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const mailHistoryIndex = buildDoctorMailHistoryIndexForResults1b2_(ss);
  syncNegotiationHistoryAndCautionSheets3_3_(ss, mailHistoryIndex);
}

function syncNegotiationHistoryAndCautionSheets3_3_(ssOpt, mailHistoryIndexOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM');
  const mailHistoryIndex = mailHistoryIndexOpt || buildDoctorMailHistoryIndexForResults1b2_(ss);
  if (!mailHistoryIndex) return;

  const parsed = extractStructuredNegotiationRecords3_3_(ss, mailHistoryIndex);

  // =========================================================================
  // A. 『交渉履歴』シートの作成・全件相談日順書き込み
  // =========================================================================
  let histSh = ss.getSheetByName('交渉履歴');
  if (!histSh) {
    histSh = ss.insertSheet('交渉履歴');
  }
  const headers = [
    '交渉日(相談日)',
    '対象勤務日',
    '曜日',
    '対象拠点',
    '依頼勤務時間',
    '時給区分',
    '依頼項目',
    '医師名',
    '交渉結果',
    '返答内容(簡略まとめ)',
    '恒久NG・制約要約'
  ];

  histSh.clearContents();
  histSh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#1e3a8a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  histSh.setFrozenRows(1);

  if (parsed.records.length > 0) {
    const rows = parsed.records.map(function(r) {
      return [
        r.negotiationDate,
        r.targetDate,
        r.dow,
        r.clinic,
        r.requestedTime,
        r.wageType,
        r.requestItem,
        r.doctor,
        r.result,
        r.replySummary,
        r.permanentCaution
      ];
    });
    histSh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ 『交渉履歴』シートを作成・更新しました（全 ' + parsed.records.length + ' 件・相談日順）');

  // =========================================================================
  // B. 『交渉注意医師』シートの整理・更新
  //    ・2行目のプルダウン（入力規則・書式）をコピーして適用
  //    ・K列(11列目)の長い自動抽出文は消去し、P列(16列目)に簡潔な要約のみを記載
  // =========================================================================
  const cautionSh = ss.getSheetByName('交渉注意医師');
  if (!cautionSh) return;

  // P列(16列目)まで列が存在することを保証
  if (cautionSh.getMaxColumns() < 16) {
    cautionSh.insertColumnsAfter(cautionSh.getMaxColumns(), 16 - cautionSh.getMaxColumns());
  }

  // 1. 前回入った不要行（柿本優・大沼奈央やB列崩れ行）の削除と、K列に入ってしまった「【11/...】」長文の消去
  const initLastR = cautionSh.getLastRow();
  if (initLastR >= 2) {
    const checkVals = cautionSh.getRange(2, 1, initLastR - 1, 16).getDisplayValues();
    for (let rIdx = checkVals.length - 1; rIdx >= 0; rIdx--) {
      const rowNum = rIdx + 2;
      const docNameCell = String(checkVals[rIdx][0] || '').replace(/[\s\u3000]+/g, '');
      const colB = String(checkVals[rIdx][1] || '').trim();
      const colK = String(checkVals[rIdx][10] || '').trim();

      // 誤抽出だった「柿本優」「大沼奈央」「中﨑寿隆」またはB列崩れ行は行ごと削除
      if (colB.includes('自動抽出') || docNameCell === '柿本優' || docNameCell === '大沼奈央' || docNameCell === '中﨑寿隆') {
        cautionSh.deleteRow(rowNum);
        continue;
      }

      // K列(11列目)に前回の「【11/05:...】」等の自動抽出長文が入っていたら綺麗に消す
      if (/^【\d{2}\/\d{2}/.test(colK)) {
        cautionSh.getRange(rowNum, 11).clearContent();
      } else if (colK.includes('\n【')) {
        const cleanedK = colK.split('\n').filter(function(l) { return !/^【\d{2}\/\d{2}/.test(l.trim()); }).join('\n');
        cautionSh.getRange(rowNum, 11).setValue(cleanedK);
      }
    }
  }

  if (parsed.cautionUpdates.length === 0) return;

  const docMasterMap = buildDoctorMasterInfoMapForCaution3_3_(ss);
  const lastR = cautionSh.getLastRow();
  const existingData = lastR >= 2
    ? cautionSh.getRange(2, 1, lastR - 1, 16).getDisplayValues()
    : [];

  let updatedCautionCount = 0;

  parsed.cautionUpdates.forEach(function(cu) {
    const docKey = String(cu.doctor || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '');
    const masterInfo = docMasterMap.get(docKey) || { licenseNo: '', dept: '小児科', area: cu.inferredArea || '関東' };

    let matchedRowNum = -1;
    for (let i = 0; i < existingData.length; i++) {
      const rk = String(existingData[i][0] || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '');
      if (rk && rk === docKey) {
        matchedRowNum = i + 2;
        break;
      }
    }

    if (matchedRowNum !== -1) {
      // 既存行がある場合：P列(16列目)に簡潔な要約をセット（既に手書き等で同じ趣旨があれば上書きしない）
      const pCell = cautionSh.getRange(matchedRowNum, 16);
      const curP = String(pCell.getDisplayValue() || '').trim();
      if (!curP || /^【\d{2}\/\d{2}/.test(curP)) {
        pCell.setValue(cu.conciseSummary);
        updatedCautionCount++;
      }
      // NG拠点・NG曜日があれば2行目のプルダウン規則に沿ってセット
      if (cu.ngClinic) trySetNgClinicInRow3_3_(cautionSh, matchedRowNum, cu.ngClinic);
      if (cu.ngDows && cu.ngDows.length > 0) trySetNgDowsInRow3_3_(cautionSh, matchedRowNum, cu.ngDows);
    } else {
      // 新規行を追加する場合：必ず「2行目」からプルダウン（入力規則）と書式をコピーして作成
      const targetRowNum = cautionSh.getLastRow() + 1;
      const templateRow = cautionSh.getRange(2, 1, 1, 15);
      const targetRowRange = cautionSh.getRange(targetRowNum, 1, 1, 15);

      templateRow.copyTo(targetRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      templateRow.copyTo(targetRowRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      targetRowRange.clearContent();

      cautionSh.getRange(targetRowNum, 1).setValue(cu.doctor);              // A列: 医師名
      cautionSh.getRange(targetRowNum, 2).setValue(masterInfo.licenseNo);   // B列: 医籍番号
      try {
        cautionSh.getRange(targetRowNum, 3).setValue(masterInfo.dept || '小児科'); // C列: 診療科目
      } catch (e) {}
      try {
        cautionSh.getRange(targetRowNum, 4).setValue('部分不可');           // D列: 交渉条件
      } catch (e) {}
      try {
        cautionSh.getRange(targetRowNum, 5).setValue(masterInfo.area || '関東'); // E列: NGエリア
      } catch (e) {}

      // K列(11列目)は空欄のままにし、P列(16列目)に簡潔な要約メモを記載
      cautionSh.getRange(targetRowNum, 16).setValue(cu.conciseSummary);

      if (cu.ngClinic) trySetNgClinicInRow3_3_(cautionSh, targetRowNum, cu.ngClinic);
      if (cu.ngDows && cu.ngDows.length > 0) trySetNgDowsInRow3_3_(cautionSh, targetRowNum, cu.ngDows);

      existingData.push([cu.doctor, masterInfo.licenseNo, masterInfo.dept, '部分不可', masterInfo.area]);
      updatedCautionCount++;
    }
  });

  Logger.log('✅ 『交渉注意医師』シートを整理し、P列（16列目）へ簡潔な要約を反映しました（更新: ' + updatedCautionCount + ' 件）');
}

/**
 * F〜J列（6〜10列目：NG拠点）の空きセルにNG拠点名をセットするヘルパー
 */
function trySetNgClinicInRow3_3_(cautionSh, rowNum, ngClinicName) {
  if (!ngClinicName) return;
  const rowVals = cautionSh.getRange(rowNum, 6, 1, 5).getDisplayValues()[0];
  if (rowVals.some(function(v) { return String(v).trim() === ngClinicName; })) return;

  for (let c = 6; c <= 10; c++) {
    if (!String(rowVals[c - 6] || '').trim()) {
      try {
        cautionSh.getRange(rowNum, c).setValue(ngClinicName);
      } catch (e) {
        cautionSh.getRange(rowNum, c).clearContent();
      }
      break;
    }
  }
}

/**
 * L〜O列（12〜15列目：NG曜日1〜4）の空きセルにNG曜日（例: 水曜日）をセットするヘルパー
 */
function trySetNgDowsInRow3_3_(cautionSh, rowNum, ngDows) {
  if (!ngDows || ngDows.length === 0) return;
  const rowVals = cautionSh.getRange(rowNum, 12, 1, 4).getDisplayValues()[0];

  ngDows.forEach(function(dowStr) {
    if (rowVals.some(function(v) { return String(v).trim() === dowStr; })) return;
    for (let c = 12; c <= 15; c++) {
      if (!String(rowVals[c - 12] || '').trim()) {
        try {
          cautionSh.getRange(rowNum, c).setValue(dowStr);
          rowVals[c - 12] = dowStr;
        } catch (e) {
          cautionSh.getRange(rowNum, c).clearContent();
        }
        break;
      }
    }
  });
}

/**
 * 長い返信文から、P列に記載するための「一目でわかる簡潔な1行要約」を生成する関数
 */
function summarizePermanentCautionText3_3_(rawReplyText) {
  const r = String(rawReplyText || '').replace(/\s+/g, ' ').trim();

  // 1. 常勤先・平日日中不可パターン（例: 清水信隆先生）
  if (/平日は常勤/.test(r) && /平日日中の勤務は出来ません/.test(r)) {
    if (/18時からの勤務につきましては検討は出来ます/.test(r)) {
      return '平日日中の勤務不可、18時からの勤務は検討できるが基本不可';
    }
    return '平日は常勤勤務があるため平日日中の勤務不可';
  }

  // 2. 特定の拠点・時間帯のみ可能パターン（例: 野村望先生）
  const onlyMatch = r.match(/(?:もともとの)?([^\s、。]{2,10}?\s*\d{1,2}:\d{2}[~～\-]\d{1,2}:\d{2})以外だと難/);
  if (onlyMatch) {
    if (/子供|子ども|通勤/.test(r)) {
      return '子供の預け先・通勤時間の関係で' + onlyMatch[1].replace(/\s+/g, '') + '以外は勤務不可';
    }
    return onlyMatch[1].replace(/\s+/g, '') + '以外の枠は勤務不可';
  }

  // 3. 育児・実家預けによる特定拠点移動不可パターン（例: 灘理絵先生）
  const childClinicMatch = r.match(/子供を実家に預けてから.*?([^\s、。]{2,8}?)(?:は|への移動は)難/);
  if (childClinicMatch) {
    return '子供を実家に預けて向かう関係で' + childClinicMatch[1] + 'への移動不可';
  }

  // 4. 何年も前から特定拠点不可パターン（例: 玉木愛乃先生）
  const permClinicMatch = r.match(/([^\s、。]{2,10}?)につきましては、?何年も前から勤務不可/);
  if (permClinicMatch) {
    return permClinicMatch[1] + 'は勤務不可';
  }

  // 5. エリア限定パターン（例: 神奈川近辺でないと勤務不可）
  const areaOnlyMatch = r.match(/([^\s、。]{2,8}?近辺)でないと勤務不可/);
  if (areaOnlyMatch) {
    return areaOnlyMatch[1] + '以外は勤務不可';
  }

  // 6. 汎用短縮（挨拶や日付を削ぎ落として要点のみ35文字以内に要約）
  let s = stripPoliteBoilerplate1b2_(r)
    .replace(/^↑.*?(?:件です|承知しました|検討させて頂きます)[。．\s\/]*/g, '')
    .replace(/申し訳(?:ないのですが|ございませんが)[、。\s]*/g, '')
    .replace(/まず、前提として、[、。\s]*/g, '')
    .trim();

  const firstSentence = s.split(/[。！!]/)[0] || s;
  return firstSentence.slice(0, 38);
}

/**
 * 医師マスタ等から「医籍番号・診療科目・エリア（関東/関西）」を取得するヘルパー
 */
function buildDoctorMasterInfoMapForCaution3_3_(ss) {
  const map = new Map();
  if (typeof loadDoctorSearchList64_ === 'function') {
    try {
      loadDoctorSearchList64_(ss).forEach(function(d) {
        const k = String(d.name || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '');
        if (!k) return;
        map.set(k, {
          licenseNo: d.id || d.licenseNo || '',
          dept: d.dept || '小児科',
          area: (d.area && d.area.includes('関西')) ? '関西' : '関東'
        });
      });
    } catch (e) {}
  }

  ss.getSheets().forEach(function(sh) {
    const sName = sh.getName();
    if (sName === '交渉注意医師' || sName === '交渉履歴' || sh.getLastRow() < 2) return;
    const headers = sh.getRange(1, 1, 1, Math.min(15, sh.getLastColumn())).getDisplayValues()[0];
    let nameIdx = -1;
    let licIdx = -1;
    let deptIdx = -1;
    headers.forEach(function(h, idx) {
      const hs = String(h || '').trim();
      if (hs === '医師名' || hs === '氏名' || hs === '医師') nameIdx = idx;
      if (hs.includes('医籍番号') || hs.includes('医師ID')) licIdx = idx;
      if (hs.includes('診療科')) deptIdx = idx;
    });
    if (nameIdx !== -1 && licIdx !== -1) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(nameIdx, licIdx, deptIdx) + 1).getDisplayValues();
      rows.forEach(function(r) {
        const k = String(r[nameIdx] || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '');
        const lic = String(r[licIdx] || '').trim();
        if (k && lic && (!map.has(k) || !map.get(k).licenseNo)) {
          const rawDept = deptIdx !== -1 ? String(r[deptIdx] || '').trim() : '小児科';
          map.set(k, {
            licenseNo: lic,
            dept: rawDept.includes('内科') ? '内科' : '小児科',
            area: '関東'
          });
        }
      });
    }
  });

  return map;
}

function extractStructuredNegotiationRecords3_3_(ss, mailHistoryIndex) {
  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  const getDow = function(ymd) {
    const m = String(ymd || '').match(/(\d{4})\/(\d{2})\/(\d{2})/);
    if (!m) return '';
    return dows[new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay()] + '曜';
  };

  const cleanClinicLocal = function(rawName) {
    return String(rawName || '')
      .replace(/[【】\[\]（）\(\)]/g, ' ')
      .replace(/小児科|内科|耳鼻咽喉科|皮膚科|整形外科/g, '')
      .replace(/[\s\u3000]+/g, '')
      .trim();
  };

  const extractDatesLocal = function(text, baseDtStr) {
    const s = String(text || '');
    const baseYear = String(baseDtStr || '2026').slice(0, 4) || '2026';
    const found = [];
    const reJp = /(?:(\d{4})年\s*)?(\d{1,2})月\s*(\d{1,2})日/g;
    let m;
    while ((m = reJp.exec(s)) !== null) {
      const y = m[1] || baseYear;
      const ymd = y + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
      if (!found.includes(ymd)) found.push(ymd);
    }
    const reSlash = /(?:(\d{4})\/)?(\d{1,2})\/(\d{1,2})(?![0-9])/g;
    while ((m = reSlash.exec(s)) !== null) {
      const mon = Number(m[2]);
      const day = Number(m[3]);
      if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) {
        const y = m[1] || baseYear;
        const ymd = y + '/' + ('0' + mon).slice(-2) + '/' + ('0' + day).slice(-2);
        if (!found.includes(ymd)) found.push(ymd);
      }
    }
    return found;
  };

  const inferRequestType = function(pitchText) {
    const s = String(pitchText || '');
    if (/ご変更|勤務地変更|移動/.test(s)) return '移動依頼';
    if (/延長/.test(s)) return '延長依頼';
    if (/前倒し|早出/.test(s)) return '早出依頼';
    return '個別アプローチ';
  };

  const inferWageType = function(pitchText, targetTimeStr) {
    const s = String(pitchText || '');
    if (/特別時給|特別日給|時給\s*1[45678],?000|加算/.test(s)) return '特別時給';
    const wageMatch = s.match(/日給[：:]\s*(\d{2,3}),?(\d{3})円/);
    const tmMatch = String(targetTimeStr || '').match(/(\d{1,2}):\d{2}[~～\-](\d{1,2}):\d{2}/);
    if (wageMatch && tmMatch) {
      const dailyYen = Number(wageMatch[1] + wageMatch[2]);
      const hours = Number(tmMatch[2]) - Number(tmMatch[1]);
      if (hours > 0 && (dailyYen / hours) >= 13500) return '特別時給';
    }
    return '通常時給';
  };

  // ★恒久的な制約のみを厳選して判定し、簡潔な要約・NG拠点・NG曜日を返す
  const detectPermanentCautionRule = function(replyText) {
    const r = String(replyText || '');

    // 単発の予定都合や承諾メールは恒久制約から確実に除外
    if (/当日は|その日は|両日ともに|相談①②|別勤務があり|夜勤ですので|日勤なので|予定があり|都合が悪く|どちらも可能です/.test(r) &&
        !/平日は常勤|子供|子ども|もともとの|何年も前から/.test(r)) {
      return null;
    }

    const isPerm = /平日は常勤|平日日中の勤務は出来ません|子供.*預け|子ども.*預け|もともとの.*以外だと難|何年も前から勤務不可|近辺でないと勤務不可|曜日は.*(?:難|不可)/.test(r);
    if (!isPerm) return null;

    let ngClinic = '';
    const ngClinicMatch = r.match(/(海老名|越谷レイクタウン|天王町|武蔵小杉|東品川|川口|流山おおたかの森|北葛西|亀有|西新井|新鎌ケ谷)(?:は|につきましては|への移動は)[^。、]*?(?:難|不可|厳)/);
    if (ngClinicMatch) ngClinic = ngClinicMatch[1];

    const ngDows = [];
    const dowMatch = r.match(/([月火水木金土日])曜(?:日)?は[^。、]*?(?:難|不可|不可能)/);
    if (dowMatch) ngDows.push(dowMatch[1] + '曜日');

    return {
      conciseSummary: summarizePermanentCautionText3_3_(r),
      ngClinic: ngClinic,
      ngDows: ngDows
    };
  };

  const records = [];
  const cautionUpdates = [];
  const seenRecordKey = new Set();
  const seenCautionKey = new Set();

  const qSh = ss.getSheetByName('医師質問');
  const aSh = ss.getSheetByName('DS回答');
  const dispNameMap = new Map();
  [qSh, aSh].forEach(function(sh) {
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 3, sh.getLastRow() - 1, 1).getDisplayValues().forEach(function(r) {
      const k = String(r[0] || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '').trim();
      if (k && !dispNameMap.has(k)) dispNameMap.set(k, String(r[0]).trim());
    });
  });

  mailHistoryIndex.forEach(function(docHistory, dKey) {
    const docName = dispNameMap.get(dKey) || dKey;
    const timeline = (docHistory.timeline || []).filter(function(ev) { return ev.dtStr >= '2026/10/28'; });

    for (let i = 0; i < timeline.length; i++) {
      const ev = timeline[i];
      if (ev.role !== 'STAFF') continue;
      if (!/ご勤務につきましてご相談|ご変更いただくことは可能|延長していただくことは可能|前倒ししていただくことは可能|いずれかの枠にてご協力/.test(ev.text)) continue;

      const pitchDate = ev.dtStr.slice(0, 10);
      const reqType = inferRequestType(ev.text);

      const cleanedPitchForSlot = ev.text
        .replace(/現在[）\)][\s\S]*?(?=ご相談[）\)]|変更後[）\)]|$)/g, '')
        .replace(/変更前[）\)][\s\S]*?(?=変更後[）\)]|ご相談[）\)]|$)/g, '');

      const pitchedSlots = [];
      const proposalRe = /ご相談[）\)]\s*【([^】]+)】\s*(\d{1,2}:\d{2}[~～\-]\d{1,2}:\d{2})/g;
      let pm;
      const baseDates = extractDatesLocal(ev.text, ev.dtStr);
      while ((pm = proposalRe.exec(ev.text)) !== null) {
        const tDate = baseDates[0] || '';
        if (tDate) {
          pitchedSlots.push({
            targetDate: tDate,
            clinic: cleanClinicLocal(pm[1]),
            timeRange: pm[2].replace(/[~～]/g, '-')
          });
        }
      }

      if (pitchedSlots.length === 0) {
        const slotRegex = /(\d{1,2}月\d{1,2}日|\d{4}年\d{1,2}月\d{1,2}日)[^\n【]*【([^】]+)】\s*(\d{1,2}:\d{2}[~～\-]\d{1,2}:\d{2})/g;
        let sm;
        while ((sm = slotRegex.exec(cleanedPitchForSlot)) !== null) {
          const dList = extractDatesLocal(sm[1], ev.dtStr);
          if (dList.length > 0) {
            pitchedSlots.push({
              targetDate: dList[0],
              clinic: cleanClinicLocal(sm[2]),
              timeRange: sm[3].replace(/[~～]/g, '-')
            });
          }
        }
      }

      if (pitchedSlots.length === 0) {
        const cm = cleanedPitchForSlot.match(/【([^】]+)】/);
        const tm = cleanedPitchForSlot.match(/(\d{1,2}:\d{2}[~～\-]\d{1,2}:\d{2})/);
        baseDates.forEach(function(fd) {
          pitchedSlots.push({
            targetDate: fd,
            clinic: cm ? cleanClinicLocal(cm[1]) : '複数/不明',
            timeRange: tm ? tm[1].replace(/[~～]/g, '-') : '時間記載なし'
          });
        });
      }

      let hasDocReply = false;
      let combinedDocReply = '';
      for (let j = i + 1; j < timeline.length; j++) {
        if (timeline[j].role === 'STAFF' && /ご勤務につきましてご相談|ご変更いただくことは可能|延長していただくことは可能/.test(timeline[j].text)) break;
        if (timeline[j].role === 'DOCTOR' && !/パターン[ABＡＢ]|交通費|駐車場|有休申請|jinjer/.test(timeline[j].text)) {
          hasDocReply = true;
          combinedDocReply += (combinedDocReply ? ' / ' : '') + timeline[j].text;
        }
      }

      if (!hasDocReply) continue;

      const permObj = detectPermanentCautionRule(combinedDocReply);
      const conciseSummary = permObj ? permObj.conciseSummary : '';

      if (permObj && conciseSummary) {
        const cKey = dKey + '_' + conciseSummary;
        if (!seenCautionKey.has(cKey)) {
          seenCautionKey.add(cKey);
          cautionUpdates.push({
            doctor: docName,
            detectedDate: pitchDate,
            conciseSummary: conciseSummary,
            ngClinic: permObj.ngClinic || '',
            ngDows: permObj.ngDows || [],
            inferredArea: /天美|今里|長吉長原|堺鉄砲町|東岸和田|豊中|阿波座|鶴見緑地/.test(ev.text) ? '関西' : '関東'
          });
        }
      }

      pitchedSlots.forEach(function(sl) {
        const uniqKey = [pitchDate, sl.targetDate, sl.clinic, sl.timeRange, dKey].join('|');
        if (seenRecordKey.has(uniqKey)) return;
        seenRecordKey.add(uniqKey);

        const st = evaluateDoctorAdjustmentStatusObj1b2_(docName, '', sl.targetDate, sl.clinic, mailHistoryIndex);
        if (st.code === 'WAIT') return;

        const resultLabel = st.code === 'ACCEPT' ? '⭕承諾(成功)' : (st.code === 'DECLINE' ? '❌辞退(失敗)' : '⚠️条件付/要確認');
        const shortSummary = st.reason || extractCoreReasonSnippet1b2_(combinedDocReply, sl.targetDate, docName);

        records.push({
          negotiationDate: pitchDate,
          targetDate: sl.targetDate,
          dow: getDow(sl.targetDate),
          clinic: sl.clinic,
          requestedTime: sl.timeRange,
          wageType: inferWageType(ev.text, sl.timeRange),
          requestItem: reqType,
          doctor: docName,
          result: resultLabel,
          replySummary: shortSummary,
          permanentCaution: conciseSummary || '-'
        });
      });
    }
  });

  records.sort(function(a, b) {
    if (a.negotiationDate !== b.negotiationDate) return a.negotiationDate.localeCompare(b.negotiationDate);
    return a.targetDate.localeCompare(b.targetDate);
  });

  return { records: records, cautionUpdates: cautionUpdates };
}