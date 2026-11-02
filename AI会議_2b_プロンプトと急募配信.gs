/**
 * ============================================================================
 * AI会議_2b_プロンプトと急募配信.gs
 * （本日分完全除外・明日起算1週間分の未充足枠自動補完・1P収容版）
 * ============================================================================
 */

function buildSlotChunkPrompt_(customRulesText) {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（攻め・守り統合）」です。',
    '思考や前置きを省き、直ちに以下の優先順位と表記ルールで各不在枠の調整案（JSON）のみを出力してください。',
    '{"slotPlans":[{"slotId":1,"plan1":"","plan2":"","plan3":"","auditNote":"-"}]}'
  ].join('\n');
}

function buildSummaryMetaPrompt_() {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（総務・全体統括）」です。',
    '対象期間の全体傾向と優先対応ポイントを2文以内で簡潔に出力してください。',
    '{"executiveSummary":"対象期間の全体傾向と優先対応ポイント（2文以内）"}'
  ].join('\n');
}

function resolveActualHourlyWage_(listedWageRaw, specialWageRaw) {
  const baseWage = Number(String(listedWageRaw || '13000').replace(/[^0-9]/g, '')) || 13000;
  const spStr = String(specialWageRaw || '').trim();

  if (!spStr || spStr === '-' || spStr === 'なし' || spStr === 'FALSE' || spStr === '0' || spStr.includes('通常')) {
    return baseWage;
  }

  const addMatch = spStr.match(/[\+＋]\s*([0-9,]+)/);
  if (addMatch) {
    const addAmt = Number(addMatch[1].replace(/,/g, '')) || 0;
    return baseWage + addAmt;
  }

  const directNum = Number(spStr.replace(/[^0-9]/g, '')) || 0;
  if (directNum >= 10000) {
    return directNum;
  } else if (directNum >= 1000 && directNum <= 9000) {
    return baseWage + directNum;
  }

  return baseWage;
}

function normalizeYmdKey_(rawDateStr) {
  const m = String(rawDateStr || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return String(rawDateStr || '').trim();
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 『最新_不在枠』に直近数日分（例：明日11/3〜11/6）が含まれていない場合、
 * 『BQシフト状況パック（mijyusoku）』および『最新_当日シフト』シートから明日以降の未充足枠を自動補完する関数
 */
function supplementNearTermUnfilledSlots_(existingSlots, windowStartYmd, windowEndYmd) {
  // ★既存リストからも本日以前（windowStartYmd 未満）は確実に除外する
  const combined = (existingSlots || []).filter(function(s) {
    const d = normalizeYmdKey_(s.date);
    return d >= windowStartYmd && d <= windowEndYmd;
  });

  const seenKeys = new Set();
  combined.forEach(function(s) {
    const d = normalizeYmdKey_(s.date);
    const c = cleanClinicBaseName_(s.clinic);
    const t = String(s.absenceTime || '').replace(/～/g, '-').trim();
    seenKeys.add(d + '_' + c + '_' + t);
  });

  // 1. 『BQシフト状況パック』メールの mijyusoku（未充足リスト）から明日以降の枠を補完
  if (typeof fetchAndDecryptBqMailForResolver_ === 'function') {
    try {
      const bqData = fetchAndDecryptBqMailForResolver_();
      if (bqData && Array.isArray(bqData.mijyusoku)) {
        bqData.mijyusoku.forEach(function(m) {
          const d = normalizeYmdKey_(m.d || m.date);
          if (!d || d < windowStartYmd || d > windowEndYmd) return; // 本日分は除外、明日以降のみ
          const rawC = String(m.c || m.clinic || '').trim();
          if (!rawC || rawC.includes('内科') || String(m.dept || '').includes('内科') || rawC.includes('バックアップ')) return;

          const cleanC = cleanClinicBaseName_(rawC);
          const absTime = String(m.t || m.time || '').replace(/～/g, '-').trim();
          if (!cleanC || !absTime) return;

          const key = d + '_' + cleanC + '_' + absTime;
          if (seenKeys.has(key)) return;
          seenKeys.add(key);

          const wageVal = m.w || m.wage || m.listedWage || '13000';
          combined.push({
            date: d,
            area: inferAreaForDoc_(cleanC),
            clinic: cleanC,
            dept: '小児科',
            absenceTime: absTime,
            listedWage: String(wageVal),
            specialWage: m.sp || m.specialWage || ''
          });
        });
      }
    } catch (e) {}
  }

  // 2. 『最新_当日シフト』シートで「未充足」（または0人）になっている枠からも明日以降の枠を補完
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('最新_当日シフト');
    if (sh && sh.getLastRow() >= 2) {
      const vals = sh.getDataRange().getDisplayValues();
      for (let r = 1; r < vals.length; r++) {
        const row = vals[r];
        const d = normalizeYmdKey_(row[0]);
        if (!d || d < windowStartYmd || d > windowEndYmd) continue; // 本日分は除外、明日以降のみ

        const rawC = String(row[2] || '').trim();
        const dept = String(row[3] || '').trim();
        if (!rawC || rawC.includes('バックアップ') || rawC.includes('内科') || dept.includes('内科')) continue;
        const cleanC = cleanClinicBaseName_(rawC);

        const bands = [
          { count: Number(row[4] || 0), docs: String(row[7] || ''), time: '09:00-13:00', wage: '13000' },
          { count: Number(row[5] || 0), docs: String(row[8] || ''), time: '15:00-18:00', wage: '13000' },
          { count: Number(row[6] || 0), docs: String(row[9] || ''), time: cleanC.includes('北葛西') ? '18:00-20:00' : '18:00-21:00', wage: '15000' }
        ];

        bands.forEach(function(b) {
          if (b.docs.includes('未充足') || b.count === 0) {
            const key = d + '_' + cleanC + '_' + b.time;
            if (seenKeys.has(key)) return;
            seenKeys.add(key);

            combined.push({
              date: d,
              area: inferAreaForDoc_(cleanC),
              clinic: cleanC,
              dept: '小児科',
              absenceTime: b.time,
              listedWage: b.wage,
              specialWage: ''
            });
          }
        });
      }
    }
  } catch (e2) {}

  return combined;
}

/**
 * 急募お知らせ配信テンプレート生成（本日を除き、必ず「明日〜7日間」を抽出して1ページに収める）
 */
function buildOfficialAnnouncementProposal_(slotsContext) {
  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startDt = new Date(todayMidnight.getTime() + 1 * 86400000); // 明日（本日分は除外）
  const endDt = new Date(startDt.getTime() + 6 * 86400000);         // 明日から7日間（1週間分）

  const windowStartYmd = Utilities.formatDate(startDt, 'Asia/Tokyo', 'yyyy/MM/dd');
  const windowEndYmd = Utilities.formatDate(endDt, 'Asia/Tokyo', 'yyyy/MM/dd');

  const basePedSlots = (slotsContext || []).filter(function(s) { return s.dept === '小児科'; });
  const allPedSlots = supplementNearTermUnfilledSlots_(basePedSlots, windowStartYmd, windowEndYmd);

  if (allPedSlots.length === 0) {
    return { recommendedDateReason: '対象の小児科不在枠がありません。', templateText: '' };
  }

  const byDate = {};
  allPedSlots.forEach(function(s) {
    const normDate = normalizeYmdKey_(s.date);
    if (normDate < windowStartYmd) return; // 本日以前は確実に除外
    if (!byDate[normDate]) byDate[normDate] = [];
    byDate[normDate].push(s);
  });

  const sortedAllDates = Object.keys(byDate).sort();
  let targetDates = sortedAllDates.filter(function(d) {
    return d >= windowStartYmd && d <= windowEndYmd;
  });

  if (targetDates.length === 0) {
    targetDates = sortedAllDates.slice(0, 7);
  }

  const dows = ['日', '月', '火', '水', '木', '金', '土'];
  const withDow = function(ymd) {
    const m = String(ymd || '').match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (!m) return ymd;
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) + '（' + dows[dt.getDay()] + '）';
  };

  let totalTargetSlotsCount = 0;

  const dateSections = targetDates.map(function(dateStr) {
    const daySlots = byDate[dateStr] || [];
    totalTargetSlotsCount += daySlots.length;

    const areaOrder = ['東京', '神奈川', '埼玉', '千葉', '茨城', '関西'];
    const areaClinicMap = {};
    areaOrder.forEach(function(a) { areaClinicMap[a] = {}; });

    daySlots.forEach(function(u) {
      const area = u.area || '東京';
      if (!areaClinicMap[area]) areaClinicMap[area] = {};

      const cleanClinic = String(u.clinic || '').replace(/小児科|[\(（].*?[\)）]/g, '').trim();
      const absTime = String(u.absenceTime || '').trim().replace('-', '～');
      const hourlyWage = resolveActualHourlyWage_(u.listedWage, u.specialWage);

      let hours = 3;
      const tm = absTime.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
      if (tm) {
        hours = (Number(tm[3]) + Number(tm[4]) / 60) - (Number(tm[1]) + Number(tm[2]) / 60);
      }
      const dailyPay = Math.round(hourlyWage * hours).toLocaleString();
      const hourlyStr = hourlyWage.toLocaleString();

      if (!areaClinicMap[area][cleanClinic]) {
        areaClinicMap[area][cleanClinic] = [];
      }
      areaClinicMap[area][cleanClinic].push(absTime + ' 日給' + dailyPay + '円(時給' + hourlyStr + '円)');
    });

    const clinicLines = [];
    Object.keys(areaClinicMap).forEach(function(area) {
      const clinicsObj = areaClinicMap[area];
      const clinicNames = Object.keys(clinicsObj);
      if (clinicNames.length === 0) return;

      clinicNames.forEach(function(cName) {
        clinicLines.push('[' + area + ']【' + cName + '】' + clinicsObj[cName].join(' ／ '));
      });
    });

    return '◆' + withDow(dateStr) + '\n' + clinicLines.join('\n');
  });

  const periodStartLabel = withDow(windowStartYmd);
  const periodEndLabel = withDow(windowEndYmd);
  const periodTitle = periodStartLabel + '～' + periodEndLabel.slice(5);

  const templateText = [
    'タイトル：※急募※ ' + periodTitle + ' 【小児科】にてご勤務いただける先生を探しております！',
    '（こちらはキャップスドクター365へご登録いただいている先生方へお送りしております）',
    '先生方 平素より大変お世話になっております。ドクターサポート事務局です。',
    '表題の件につきまして、' + periodTitle + ' において複数拠点での医師不在の状況となり、緊急募集を行っております。',
    'ご都合の良い小児科の先生がいらっしゃいましたら、ご勤務のご検討をいただきたく存じます。',
    '============================================',
    '＜急募対象シフト（日付別・募集時給換算）＞',
    dateSections.join('\n'),
    '============================================',
    '※紹介会社経由でご勤務いただいております先生は紹介会社様経由でのご応募をお願い申し上げます。',
    'ドクターサポート事務局（050-3187-6647 / doctor-support@caps365.jp）'
  ].join('\n');

  return {
    recommendedDateReason: '月曜・水曜の週2回配信ルール（本日を除いた明日以降1週間分）に基づき、' + periodStartLabel + ' ～ ' + periodEndLabel + '（計' + totalTargetSlotsCount + '枠）を1ページ内にまとめました。',
    templateText: templateText
  };
}