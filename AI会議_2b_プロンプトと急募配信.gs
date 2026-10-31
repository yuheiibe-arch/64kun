/**
 * ============================================================================
 * AI会議_2b_プロンプトと急募配信.gs
 * （①前後医 ➔ ②ドミナント2診崩し ➔ ③個別打診の優先順位プロンプト ＆ 急募配信生成）
 * ============================================================================
 */

function buildSlotChunkPrompt_(customRulesText) {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（攻め・守り統合）」です。',
    '思考や前置きを省き、直ちに以下の優先順位と表記ルールで各不在枠の調整案（JSON）のみを出力してください。',
    '',
    '🚫絶対ルール1：入力データにあるクリニック名（拠点名）や医師名は1文字たりとも勝手に書き換えたり創作しないでください。',
    '🚫絶対ルール2：plan1 / plan2 / plan3 には「実際に打診できる前向きなアクション」のみを入れてください。不可理由や除外理由は絶対に入れないこと。',
    '🚫絶対ルール3：医師名の後ろに「（スポット）」「（紹介会社）」という属性表記は一切付けないでください。',
    '',
    '【最優先カスタムルール】',
    customRulesText,
    '',
    '【plan1〜plan3 の優先順位（この順番を厳守して枠を埋めること）】',
    '■ 第1優先：自拠点の前後医による「延長・早出」（prevDocs / nextDocs）',
    '  - prevDocs または nextDocs が「なし」以外の場合、必ず最優先（plan1）として提案してください。',
    '  - 非常勤医師の場合: 「前医の〇〇先生（非常勤）へ延長打診」「後医の〇〇先生（非常勤）へ早出打診」',
    '  - 常勤・所定休出の場合: 「前医の〇〇先生（常勤・所定休出）へ延長打診（常勤追加勤務手当1,500円/h）」',
    '',
    '■ 第2優先：近隣・ドミナント他拠点からの「2診目崩し」（twoDocBreakdownCandidates）',
    '  - twoDocBreakdownCandidates に候補がある場合、単なる個別打診（kinmukunCandidates）よりも優先して plan1（前後医がいる場合は plan2）に必ず入れてください。',
    '  - 配列の先頭ほどドミナント優先度（S・A）が高く移動時間が短い拠点です。上位から順に採用してください。',
    '  - 表記（1名の場合）: 「【2診目崩し】川口（2診目：品川穣先生）へ移動打診（移動依頼手当10,000円）」',
    '  - 表記（2名いる場合）: 「【2診目崩し】越谷レイクタウン（2診目：①高橋あずさ先生 / ②原田匠先生）のどちらか1名へ移動打診（移動依頼手当10,000円）」',
    '  - ※SBU医師のみの移動の場合は「（移動手当0円）」と表記してください。',
    '',
    '■ 第3優先：実績のある医師への「個別打診」（kinmukunCandidates）および「1診・併行募集」（oneDocParallelCandidates）',
    '  - 残りの枠（plan2〜plan3、または前後医・2診崩しがない場合のplan1〜plan3）は、kinmukunCandidates の上位医師から順に「〇〇先生へ個別打診（同曜日同枠〇回実績）」で埋めてください。',
    '  - 有効な候補が少ない場合は無理に埋めず、oneDocParallelCandidates があれば「【1診・併行募集】拠点名（1診：〇〇先生）に時給UP募集を立て、埋まり次第移動打診」を入れてください。',
    '',
    '【現状・注意メモ（auditNote）の記載ルール（厳守）】',
    '  - 🚫禁止事項：「前医は通常常勤のため不可」「後医なし」「中抜けのため調整困難」「〇〇先生はNG条件のため除外」といった不可理由・除外理由・内部思考プロセスは一切書かないでください！',
    '  - 記載してよいのは以下の3点のみです（該当がなければ必ず "-" を出力すること）：',
    '    ① currentMemo（進行中の交渉メモ）の内容',
    '    ② plan1〜plan3に採用した候補医師が同日に他拠点確定枠を持っている場合の事実（例：「〇〇先生は同日別枠（海老名 09:00-13:00）あり」）',
    '    ③ plan1〜plan3に採用した候補医師の直近打診・辞退履歴（例：「〇〇先生は10/27辞退履歴あり」）',
    '',
    '以下のJSON形式のみを出力してください（slotIdは入力と必ず一致させること）:',
    '{',
    '  "slotPlans": [',
    '    {',
    '      "slotId": 1,',
    '      "plan1": "第1優先アクション",',
    '      "plan2": "第2優先アクション",',
    '      "plan3": "第3優先アクション",',
    '      "auditNote": "進行中メモや採用候補の同日確定別枠・直近履歴のみ（なければ - ）"',
    '    }',
    '  ]',
    '}'
  ].join('\n');
}

function buildSummaryMetaPrompt_() {
  return [
    'あなたはCAPSクリニックの「医師シフト調整AIエージェント（総務・全体統括）」です。',
    '以下のルールに従って、①全体総括、②複数日程おまとめ打診リスト（上位12名まで）、③マスタ更新確認を作成してください。',
    '🚫重要：医師名の後ろに「（スポット）」「（紹介会社）」の表記は付けないでください。',
    '',
    '【重要：複数日程おまとめ打診リスト（targetSlots）の整形ルール】',
    '必ず「■ 拠点名」ごとに分類し、改行（\\n）を使って箇条書きでスッキリ見やすく整形してください。',
    '',
    '以下のJSON形式のみを出力してください:',
    '{',
    '  "executiveSummary": "対象期間の全体傾向と優先対応ポイント（2文以内）",',
    '  "bundleRecommendations": [',
    '    {',
    '      "doctor": "〇〇 先生",',
    '      "targetSlots": "■ 拠点名\\n・MM/DD(曜) 時間帯",',
    '      "reasonAndNote": "実績が多い拠点や交渉時の注意点（簡潔に）"',
    '    }',
    '  ],',
    '  "masterUpdateQuestions": []',
    '}'
  ].join('\n');
}

function buildOfficialAnnouncementProposal_(slotsContext) {
  const pedSlots = (slotsContext || []).filter(function(s) { return s.dept === '小児科'; });
  if (pedSlots.length === 0) {
    return { recommendedDateReason: '対象の小児科不在枠がありません。', templateText: '' };
  }

  const byDate = {};
  pedSlots.forEach(function(s) {
    if (!byDate[s.date]) byDate[s.date] = [];
    byDate[s.date].push(s);
  });

  const sortedDates = Object.keys(byDate).sort(function(a, b) {
    return byDate[b].length - byDate[a].length || a.localeCompare(b);
  });
  const bestDate = sortedDates[0];
  const targetList = byDate[bestDate];

  const areaGroups = { '東京': [], '神奈川': [], '埼玉': [], '千葉': [], '茨城': [], '関西': [] };
  targetList.forEach(function(u) {
    const cleanClinic = String(u.clinic || '').replace(/小児科|[\(（].*?[\)）]/g, '').trim();
    const absTime = String(u.absenceTime || '').trim().replace('-', '～');
    const baseWage = Number(String(u.listedWage || '13000').replace(/[^0-9]/g, '')) || 13000;
    const specialHourly = baseWage + 2000;

    let hours = 3;
    const tm = absTime.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
    if (tm) {
      hours = (Number(tm[3]) + Number(tm[4]) / 60) - (Number(tm[1]) + Number(tm[2]) / 60);
    }
    const dailyPay = Math.round(specialHourly * hours).toLocaleString();
    const area = u.area || '東京';
    if (!areaGroups[area]) areaGroups[area] = [];
    areaGroups[area].push('・【' + cleanClinic + '】' + absTime + ' 日給' + dailyPay + '円');
  });

  const areaLines = [];
  Object.keys(areaGroups).forEach(function(a) {
    if (areaGroups[a].length > 0) areaLines.push(a + '\n' + areaGroups[a].join('\n'));
  });

  const templateText = [
    'タイトル：※急募※ ' + bestDate + ' 【小児科】にてご勤務いただける先生を探しております！',
    '',
    '（こちらはキャップスドクター365へご登録いただいている先生方へお送りしております）',
    '先生方',
    '平素より大変お世話になっております。ドクターサポート事務局です。',
    '表題の件につきまして、' + bestDate + ' 複数拠点での医師不在の状況となり、緊急募集を行っております。',
    'ご都合の良い小児科の先生がいらっしゃいましたら、ご勤務のご検討をいただきたく存じます。',
    '============================================',
    '＜急募＞特別時給（基本時給＋2,000円/h 換算済）',
    '◆' + bestDate,
    areaLines.join('\n\n'),
    '============================================',
    '※紹介会社経由でご勤務いただいております先生は紹介会社様経由でのご応募をお願い申し上げます。',
    'ドクターサポート事務局（050-3187-6647 / doctor-support@caps365.jp）'
  ].join('\n');

  return {
    recommendedDateReason: '週1回以下の配信ルール（オオカミ少年防止）に基づき、小児科未充足が最も集中している ' + bestDate + '（計' + targetList.length + '枠）を今週唯一の配信対象日として選定しました。',
    templateText: templateText
  };
}