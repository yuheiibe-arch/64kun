/**
 * ============================================================================
 * AI会議_自動実行_Slackカレンダー連携.gs（APIキー＆クレカ登録不要・完全自立版）
 *
 * ■ 特徴:
 *  - GEMINI_API_KEY が無くても、GASのルールエンジンが3視点（攻め・守り・総務）の
 *    全ロジックを実行し、第1部（案1〜3・FB欄）、第2部（急募配信文面）、第3部（マスタ更新確認）を作成！
 *  - もし後から GEMINI_API_KEY を入れた場合は、自動でAI生成モードにハイグレード化。
 *  - SLACK_CHANNEL_ID (C09TRHGU64D) 標準搭載済み。
 * ============================================================================
 */

const MTG_AUTO_CONFIG = {
  SPREADSHEET_ID: '1vh9hPfoaAGTZOm-GFtJQoSPgFB3GiABsbJ4ETgMqfbM',
  CALENDAR_ID: 'capsdoctor365@caps365.jp',
  EVENT_TITLE_KEYWORD: '超募集対策MTG',
  DEFAULT_SLACK_CHANNEL_ID: 'C09TRHGU64D',
  NOTEBOOK_URL: 'https://gemini.google.com/notebook/570c0b20-03f4-4ced-aee3-cf07a2cbeddb',
  GEMINI_MODEL: 'gemini-2.5-flash',
  LIGHTWEIGHT_TXT_NAME: 'AI会議用_最新コンテキストパック.txt'
};

/**
 * ★初回のみ実行：毎週月曜・水曜の朝 08:45 に自動実行するトリガーを設定
 */
function installWeeklyMtgAiTrigger() {
  const funcName = 'runScheduledAiShiftMeeting';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger(funcName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(45)
    .create();

  ScriptApp.newTrigger(funcName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8)
    .nearMinute(45)
    .create();

  Logger.log('✅ 毎週月曜・水曜 08:45 のAI自動会議＆Slack/カレンダー提出トリガーを設定しました。');
}

/**
 * ★メイン実行関数（手動テストもこの関数を実行してください）
 */
function runScheduledAiShiftMeeting() {
  const ss = SpreadsheetApp.openById(MTG_AUTO_CONFIG.SPREADSHEET_ID);

  // 0. 最新の暗号化メールがあれば先に取り込む
  if (typeof importLatestEncryptedShiftContext_ === 'function') {
    try {
      importLatestEncryptedShiftContext_(false);
    } catch (e) {
      Logger.log('暗号化メール事前取込スキップ: ' + e.message);
    }
  }

  // 1. シートから最新データを収集＆ルールエンジンで第1部〜第3部を自動構築
  const contextData = collectAndAnalyzeShiftMeeting_(ss);

  // 2. NotebookLM添付用の軽量コンテキストファイルも自動保存（「ファイルが大きすぎます」対策）
  exportLightweightContextTxt_(contextData.fullContextText);

  // 3. レポート本文の生成（GEMINI_API_KEYがあればAPI経由、無ければGAS内蔵の3視点エンジンで完全生成）
  const meetingReportText = generateAiMeetingReport_(contextData);

  // 4. Googleドキュメントを作成
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const docTitle = `【AI事前会議】超募集対策MTG_${todayStr}（対象：${contextData.periodStart}〜${contextData.periodEnd}）`;
  const docFile = createMeetingGoogleDoc_(docTitle, meetingReportText);

  // 5. カレンダーの「［実施します］超募集対策MTG」にドキュメントを添付
  const attachedEventUrl = attachDocToUpcomingMtgEvent_(docFile);

  // 6. Slack（C09TRHGU64D）へ @dsshift @dspart ＋ コードブロック ＋ URLボタン付きで送信
  postMeetingResultToSlack_({
    todayStr: todayStr,
    periodStart: contextData.periodStart,
    periodEnd: contextData.periodEnd,
    docUrl: docFile.getUrl(),
    eventUrl: attachedEventUrl
  });

  Logger.log(`✅ AI自動会議＆提出完了！ ドキュメントURL: ${docFile.getUrl()}`);
}

/**
 * NotebookLMにそのまま入れられる軽量テキストファイルをGoogleドライブに出力
 */
function exportLightweightContextTxt_(contextText) {
  const fileName = MTG_AUTO_CONFIG.LIGHTWEIGHT_TXT_NAME;
  const files = DriveApp.getFilesByName(fileName);
  let file;
  if (files.hasNext()) {
    file = files.next();
    file.setContent(contextText);
  } else {
    file = DriveApp.createFile(fileName, contextText, MimeType.PLAIN_TEXT);
  }
  Logger.log(`📄 NotebookLM用軽量ファイル出力完了: ${file.getName()} (${file.getUrl()})`);
}

function normDocNameClean_(name) {
  return String(name || '')
    .replace(/[\(（\[].*?[\)）\]]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

/**
 * 拠点名から都道府県エリアを判定するヘルパー
 */
function inferAreaFromClinic_(clinicName) {
  const c = String(clinicName || '');
  if (/海老名|高田|武蔵小杉|相模大野|相模原|新百合ヶ丘|茅ヶ崎|天王町|小田栄|東戸塚|宮前平|川崎|横浜/.test(c)) return '神奈川';
  if (/越谷|川口|川越|志木|所沢|南浦和|与野|草加松原|八潮|浦和|大宮/.test(c)) return '埼玉';
  if (/二和向台|柏の葉|村上|稲毛海岸|流山おおたかの森|新鎌ケ谷|幕張ベイパーク|八千代緑が丘|妙典|千葉ニュータウン中央|新浦安|市川|船橋|千葉/.test(c)) return '千葉';
  if (/つくば/.test(c)) return '茨城';
  if (/関西|大阪|兵庫|京都|西宮|千里|高槻|あべの/.test(c)) return '関西';
  return '東京';
}

/**
 * 『DSナレッジシート』の全シートを読み込み、APIキー不要の3視点（攻め・守り・総務）エンジンで案を構築
 */
function collectAndAnalyzeShiftMeeting_(ss) {
  const today = new Date();
  const todayYmd = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  const twoWeeksLaterYmd = Utilities.formatDate(new Date(today.getTime() + 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const oneMonthLaterYmd = Utilities.formatDate(new Date(today.getTime() + 31 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  const getRows = (name) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getDisplayValues().slice(1);
  };

  const unfilledAll = getRows('最新_不在枠').filter(r => r[0] && r[0] >= todayYmd);
  const within2W = unfilledAll.filter(r => r[0] <= twoWeeksLaterYmd);
  const targetEndYmd = (within2W.length < 10 && unfilledAll.length > within2W.length) ? oneMonthLaterYmd : twoWeeksLaterYmd;
  const targetUnfilled = unfilledAll.filter(r => r[0] <= targetEndYmd);

  const periodStart = targetUnfilled.length > 0 ? targetUnfilled[0][0] : todayYmd;
  const periodEnd = targetUnfilled.length > 0 ? targetUnfilled[targetUnfilled.length - 1][0] : targetEndYmd;

  const activeShifts = getRows('最新_当日シフト');
  const streakWarnings = getRows('最新_連勤注意リスト');
  const ngMasterRows = getRows('交渉注意医師');
  const kbRows = getRows('ナレッジシート');

  // 1. 直近の打診履歴マップ（医師名 -> 最新の打診日時と結果）
  const recentHistoryMap = new Map();
  kbRows.slice(-150).forEach(r => {
    const docKey = normDocNameClean_(r[1]);
    const docReply = String(r[2] || '').replace(/\n+/g, ' ').slice(0, 45);
    const dt = String(r[6] || '').slice(5, 10);
    const flag = String(r[7] || '');
    if (docKey && flag.includes('シフト調整')) {
      recentHistoryMap.set(docKey, `※直近${dt || '履歴'}に打診あり（${flag}: ${docReply}...）`);
    }
  });

  // 2. 交渉注意医師マップ（医師名 -> NG条件テキスト）
  const ngDocMap = new Map();
  ngMasterRows.forEach(r => {
    const docKey = normDocNameClean_(r[0]);
    if (!docKey) return;
    const ngInfo = r.slice(1).filter(Boolean).join(' / ');
    ngDocMap.set(docKey, ngInfo);
  });

  // 3. 第1部：不在枠ごとのアプローチプラン自動生成
  const part1Lines = [];
  const byDateUnfilled = {};
  targetUnfilled.forEach(r => {
    if (!byDateUnfilled[r[0]]) byDateUnfilled[r[0]] = [];
    byDateUnfilled[r[0]].push(r);
  });

  Object.keys(byDateUnfilled).sort().forEach(dateStr => {
    part1Lines.push(`### ◆ ${dateStr}`);
    const dayShifts = activeShifts.filter(s => s[0] === dateStr);

    byDateUnfilled[dateStr].forEach(u => {
      const clinic = u[1];
      const absTime = u[2];
      const listedWage = u[3];
      const prevDocs = u[4];
      const nextDocs = u[5];
      const prevStatus = u[7];
      const nextStatus = u[8];
      const specialWage = u[11];
      const memo = u[12];
      const isInternalMed = clinic.includes('内科');
      const area = inferAreaFromClinic_(clinic);

      // 時間帯からA(午前)/B(午後)/C(夜間)を判定
      const shiftSlot = absTime.startsWith('09') || absTime.startsWith('9') ? 'A'
        : (absTime.startsWith('18') || absTime.startsWith('17') ? 'C' : 'B');
      const countIdx = shiftSlot === 'A' ? 4 : (shiftSlot === 'B' ? 5 : 6);
      const docsIdx = shiftSlot === 'A' ? 7 : (shiftSlot === 'B' ? 8 : 9);

      part1Lines.push(`#### 【${clinic}】${absTime}（掲載時給：${listedWage || '未設定'} / 特別時給：${specialWage || '検討'}）`);
      if (memo) part1Lines.push(`  - 現状メモ：${memo}`);

      const plans = [];

      // ① 前後枠の非常勤医師への延長・早出（通常勤務の常勤は除外）
      if (prevDocs && !prevDocs.includes('未充足')) {
        if (prevDocs.includes('(常勤)') || prevDocs.includes('移動不可')) {
          plans.push(`【前医延長不可】前枠（${prevDocs}）は通常勤務の常勤医師のため延長対象外。`);
        } else {
          const hist = recentHistoryMap.get(normDocNameClean_(prevDocs)) || '';
          plans.push(`【前医延長】前枠の ${prevDocs} へ延長打診（状況：${prevStatus || '未打診'}） ${hist}`);
        }
      }
      if (nextDocs && !nextDocs.includes('未充足')) {
        if (nextDocs.includes('(常勤)') || nextDocs.includes('移動不可')) {
          plans.push(`【後医早出不可】後枠（${nextDocs}）は通常勤務の常勤医師のため早出対象外。`);
        } else {
          const hist = recentHistoryMap.get(normDocNameClean_(nextDocs)) || '';
          plans.push(`【後医早出】後枠の ${nextDocs} へ早出打診（状況：${nextStatus || '未打診'}） ${hist}`);
        }
      }

      if (isInternalMed) {
        plans.push(`【内科スポット個別打診】過去勤務回数の多い内科スポット・非常勤医師へ個別アプローチ。`);
      } else {
        // ② 関東バックアップシフトの確認
        const backupShift = dayShifts.find(s => s[2].includes('バックアップ'));
        if (backupShift && backupShift[docsIdx] && !backupShift[docsIdx].includes('未充足')) {
          plans.push(`【バックアップ発動】同時間帯のバックアップ枠（${backupShift[docsIdx]}）を配置（勤務開始3時間前までに発動）。`);
        }

        // ③ 2診崩し候補の探索（同エリアまたは近隣で同時間帯に2名以上いる拠点）
        const twoDocClinics = dayShifts.filter(s => {
          const cNum = Number(s[countIdx] || 0);
          return cNum >= 2 && s[2] !== clinic && !s[3].includes('内科');
        });
        twoDocClinics.forEach(tc => {
          const tcArea = inferAreaFromClinic_(tc[2]);
          if (tcArea === area || (area === '千葉' && tcArea === '東京') || (area === '東京' && (tcArea === '千葉' || tcArea === '埼玉'))) {
            plans.push(`【2診崩し移動】同時間帯に2診体制の【${tc[2]}】（配置：${tc[docsIdx]}）から非常勤または所定休出医師1名を移動打診（移動依頼手当+10,000円）。`);
          }
        });

        // ④ 併行募集＋移動依頼（★交渉しやすい医師 / ★所定休出・移動相談可 / ★SBU）
        dayShifts.forEach(s => {
          if (s[2] === clinic) return;
          const slotDocsText = String(s[docsIdx] || '');
          if (slotDocsText.includes('★交渉') || slotDocsText.includes('★移動相談可') || slotDocsText.includes('★SBU')) {
            plans.push(`【併行募集＋移動】同時間帯に【${s[2]}】勤務中の ${slotDocsText} の枠に併行募集（時給UP）を立て、埋まり次第【${clinic}】へ移動依頼。`);
          }
        });
      }

      if (plans.length === 0) {
        plans.push(`【個別打診＋急募配信】過去勤務歴のある定期非常勤・スポット医師への個別打診、および特別時給（+2,000円/h）でのお知らせ配信。`);
      }

      plans.slice(0, 4).forEach((p, idx) => {
        part1Lines.push(`* **案${idx + 1}）** ${p}`);
      });
      part1Lines.push(`* **【FB欄（採用案・修正指示・理由）】**：\n`);
    });
  });

  // 4. 第2部：急募のお知らせ配信文面（特別時給 +2,000円/h × 時間数 で日給自動計算）
  const part2Lines = buildUrgentAnnouncementText_(byDateUnfilled);

  // 5. 第3部：マスタ更新・AIメモ追記確認＆12連勤警告
  const part3Lines = [];
  part3Lines.push('### 1. 『交渉注意医師』・『交渉しやすい医師』シートへの追記確認（承認トリガー）');
  const flaggedKb = kbRows.filter(r => String(r[7] || '').includes('★要確認') || String(r[7] || '').includes('承諾・成立')).slice(-10);
  if (flaggedKb.length === 0) {
    part3Lines.push('* 直近の未確認フラグ付き返信はありません。');
  } else {
    flaggedKb.forEach(r => {
      const docName = r[1];
      const replySummary = String(r[2] || '').replace(/\n+/g, ' ').slice(0, 80);
      const dt = r[6];
      const flag = r[7];
      if (flag.includes('★要確認')) {
        part3Lines.push(`* **${docName} 先生**（${dt}）：「${replySummary}」\n  ➔ **『交渉注意医師』シート（または『交渉しやすい医師』のE列：メモ(AI)）に記載してよろしいですか？ [ OK / NG ]**`);
      } else {
        part3Lines.push(`* **${docName} 先生**（${dt}）：シフト調整の快諾・成立実績あり（「${replySummary}」）\n  ➔ **『交渉しやすい医師』シート（またはE列：メモ(AI)）に追加・記載してよろしいですか？ [ OK / NG ]**`);
      }
    });
  }

  part3Lines.push('\n### 2. 『最新_連勤注意リスト』からの12連続勤務NG警告（月またぎ・所定休出注意）');
  if (streakWarnings.length === 0) {
    part3Lines.push('* 現在、連勤警告対象の医師はいません。');
  } else {
    streakWarnings.forEach(w => {
      part3Lines.push(`* **${w[1]}**（医籍番号:${w[0]}）：**${w[2]}連勤**（期間：${w[3]} / 追加打診NG日：**${w[4]}**）\n  └ ${w[5]}`);
    });
  }

  const structuredReport = [
    `# 【AI事前会議】超募集対策MTG 調整案・議事録`,
    `対象期間：${periodStart} ～ ${periodEnd}（対象不在枠：${targetUnfilled.length}件）`,
    '',
    `## 第1部：医師不在枠 アプローチプラン（結論先行）`,
    part1Lines.join('\n'),
    `---`,
    `## 第2部：急募のお知らせ配信文面（サイト配信用・特別時給換算）`,
    part2Lines,
    `---`,
    `## 第3部：総務AIからのマスタ更新・メモ追記確認（承認トリガー）`,
    part3Lines.join('\n')
  ].join('\n');

  return {
    periodStart: periodStart,
    periodEnd: periodEnd,
    unfilledCount: targetUnfilled.length,
    structuredReport: structuredReport,
    fullContextText: structuredReport
  };
}

/**
 * お知らせ配信文面（特別時給＝基本時給+2,000円/h で日給計算）を自動生成するヘルパー
 */
function buildUrgentAnnouncementText_(byDateUnfilled) {
  const dates = Object.keys(byDateUnfilled).sort().slice(0, 3); // 直近3日分をピックアップ
  if (dates.length === 0) return '現在、対象となる不在枠はありません。';

  const blocks = [];
  dates.forEach(d => {
    const list = byDateUnfilled[d].filter(u => !String(u[1] || '').includes('内科'));
    if (list.length === 0) return;

    const areaGroups = { '東京': [], '神奈川': [], '埼玉': [], '千葉': [], '茨城': [], '関西': [] };
    list.forEach(u => {
      const clinic = String(u[1] || '').replace(/小児科|[\(（].*?[\)）]/g, '').trim();
      const absTime = String(u[2] || '').trim().replace('-', '～');
      const rawWage = Number(String(u[3] || '13000').replace(/[^0-9]/g, '')) || 13000;
      // 特別時給：基本時給 + 2,000円/h
      const specialHourly = rawWage + 2000;

      // 時間数を計算（例: 15:00～18:00 -> 3時間）
      let hours = 3;
      const tm = absTime.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
      if (tm) {
        hours = (Number(tm[3]) + Number(tm[4]) / 60) - (Number(tm[1]) + Number(tm[2]) / 60);
      }
      const dailyPay = Math.round(specialHourly * hours).toLocaleString();
      const area = inferAreaFromClinic_(clinic);
      if (!areaGroups[area]) areaGroups[area] = [];
      areaGroups[area].push(`・【${clinic}】${absTime} 日給${dailyPay}円`);
    });

    const areaTextLines = [];
    Object.keys(areaGroups).forEach(area => {
      if (areaGroups[area].length > 0) {
        areaTextLines.push(`${area}\n${areaGroups[area].join('\n')}`);
      }
    });

    blocks.push([
      `タイトル：※急募※ ${d} 【小児科】にてご勤務いただける先生を探しております！`,
      '',
      '（こちらはキャップスドクター365へご登録いただいている先生方へお送りしております）',
      '',
      '先生方',
      '',
      '平素より大変お世話になっております。ドクターサポート事務局です。',
      `表題の件につきまして、${d} 複数拠点での医師不在の状況となり、緊急募集を行っております。`,
      'ご都合の良い小児科の先生がいらっしゃいましたら、ご勤務のご検討をいただきたく存じます。',
      '恐れ入りますが、ご検討の程何卒よろしくお願い申し上げます。',
      '',
      '============================================',
      '＜急募＞特別時給（基本時給＋2,000円/h 反映済）',
      `◆${d}`,
      areaTextLines.join('\n\n'),
      '============================================',
      '※紹介会社経由でご勤務いただいております先生は紹介会社様経由でのご応募をお願い申し上げます。',
      'ドクターサポート事務局（050-3187-6647 / doctor-support@caps365.jp）'
    ].join('\n'));
  });

  return blocks.join('\n\n--------------------------------------------\n\n');
}

/**
 * GEMINI_API_KEY が設定されていればAIでさらにブラッシュアップし、未設定ならGAS内蔵エンジンのレポートをそのまま返す
 */
function generateAiMeetingReport_(contextData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    // クレジットカードやAPIキーが無くても、上で構築した3視点レポートがそのまま完成品として出力されます！
    return contextData.structuredReport;
  }

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${MTG_AUTO_CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`;
    const payload = {
      contents: [{
        role: 'user',
        parts: [{
          text: `以下の医師不在シフト調整案データをもとに、構成（第1部：結論先行の案1〜3とFB欄、第2部：急募お知らせ配信文面、第3部：マスタ更新確認と12連勤警告）を維持したまま、見やすく整理して出力してください。\n\n${contextData.structuredReport}`
        }]
      }],
      generationConfig: { temperature: 0.2 }
    };

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      return json.candidates[0].content.parts.map(p => p.text || '').join('\n');
    }
  } catch (e) {
    Logger.log('Gemini API呼び出しスキップ（内蔵エンジン出力を使用）: ' + e.message);
  }
  return contextData.structuredReport;
}

function createMeetingGoogleDoc_(title, textBody) {
  const doc = DocumentApp.create(title);
  doc.getBody().setText(textBody);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (e) {
    Logger.log('共有権限設定スキップ: ' + e.message);
  }
  return file;
}

function attachDocToUpcomingMtgEvent_(docFile) {
  const now = new Date();
  const lookAhead = new Date(now.getTime() + 4 * 24 * 60 * 60 * 1000);

  try {
    const events = Calendar.Events.list(MTG_AUTO_CONFIG.CALENDAR_ID, {
      q: MTG_AUTO_CONFIG.EVENT_TITLE_KEYWORD,
      timeMin: new Date(now.getTime() - 2 * 60 * 60 * 1000).toISOString(),
      timeMax: lookAhead.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });

    if (!events.items || events.items.length === 0) return '';

    const targetEvent = events.items[0];
    const existingAttachments = targetEvent.attachments || [];
    existingAttachments.push({
      fileUrl: docFile.getUrl(),
      title: docFile.getName(),
      mimeType: 'application/vnd.google-apps.document'
    });

    const patched = Calendar.Events.patch(
      {
        attachments: existingAttachments,
        description: (targetEvent.description ? targetEvent.description + '\n\n' : '') +
          `📄 【AI事前会議ドキュメント】\n${docFile.getUrl()}`
      },
      MTG_AUTO_CONFIG.CALENDAR_ID,
      targetEvent.id,
      { supportsAttachments: true }
    );

    return patched.htmlLink || '';
  } catch (e) {
    Logger.log('Calendar添付スキップ（左メニュー「サービス」からGoogle Calendar APIを追加してください）: ' + e.message);
    return '';
  }
}

function postMeetingResultToSlack_(info) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const channel = props.getProperty('SLACK_CHANNEL_ID') || MTG_AUTO_CONFIG.DEFAULT_SLACK_CHANNEL_ID;

  if (!token) {
    Logger.log('⚠️ SLACK_BOT_TOKEN が未設定のためSlack送信をスキップしました。');
    return;
  }

  const mentionText = resolveSlackMentions_(token, ['dsshift', 'dspart']);

  const codeBlockBody = [
    '```',
    `${info.todayStr}`,
    `対象期間：${info.periodStart}～${info.periodEnd}`,
    'AI自動会議終了いたしました。提出済みです。',
    '```'
  ].join('\n');

  const fullText = `${mentionText}\n${codeBlockBody}`;

  const actionElements = [
    {
      type: 'button',
      text: { type: 'plain_text', text: '📄 AI会議ドキュメントを開く', emoji: true },
      style: 'primary',
      url: info.docUrl
    },
    {
      type: 'button',
      text: { type: 'plain_text', text: '📓 AIルールブック(Notebook)を開く', emoji: true },
      url: MTG_AUTO_CONFIG.NOTEBOOK_URL
    }
  ];

  if (info.eventUrl) {
    actionElements.push({
      type: 'button',
      text: { type: 'plain_text', text: '📅 カレンダー予定を開く', emoji: true },
      url: info.eventUrl
    });
  }

  const payload = {
    channel: channel,
    text: fullText,
    blocks: [
      {
        type: 'section',
        text: { type: 'mrkdwn', text: fullText }
      },
      {
        type: 'actions',
        elements: actionElements
      }
    ]
  };

  const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const resObj = JSON.parse(res.getContentText());
  if (!resObj.ok) {
    throw new Error(`Slack送信エラー: ${resObj.error}`);
  }
  Logger.log(`✅ Slackチャンネル (${channel}) への通知が完了しました！`);
}

function resolveSlackMentions_(token, handles) {
  try {
    const res = UrlFetchApp.fetch('https://slack.com/api/usergroups.list', {
      method: 'get',
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.ok && Array.isArray(data.usergroups)) {
      const map = new Map();
      data.usergroups.forEach(ug => {
        if (ug.handle) map.set(ug.handle.toLowerCase(), ug.id);
        if (ug.name) map.set(ug.name.toLowerCase(), ug.id);
      });
      return handles.map(h => {
        const id = map.get(h.toLowerCase());
        return id ? `<!subteam^${id}|@${h}>` : `@${h}`;
      }).join(' ');
    }
  } catch (e) {
    Logger.log('Slackユーザーグループ解決スキップ: ' + e.message);
  }
  return handles.map(h => `@${h}`).join(' ');
}