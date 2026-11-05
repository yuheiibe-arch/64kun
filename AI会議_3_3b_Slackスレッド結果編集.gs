/**
 * ============================================================================
 * AI会議_3_3b_Slackスレッド結果編集.gs
 * チャンネル C0AJTMLF0UD の「超募集MTGログ（前半・後半）」スレッド内にある
 * dskun の既存日程メッセージを検出し、引用(>)＋「⇒結果（シフト応募による充足含む）」付きに自動上書き編集（chat.update）
 * ============================================================================
 */

const MTG_THREAD_SLACK_CHANNEL_ID = 'C0AJTMLF0UD';
const DSKUN_BOT_ID = 'B0C441SES4A';

function updateExistingSlackThreadMessagesInC0AJTMLF0UD_(mailHistoryIndex) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  if (!token) return;

  const histUrl = 'https://slack.com/api/conversations.history?channel=' + encodeURIComponent(MTG_THREAD_SLACK_CHANNEL_ID) + '&limit=60';
  const histRes = UrlFetchApp.fetch(histUrl, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const histData = JSON.parse(histRes.getContentText());
  if (!histData.ok || !Array.isArray(histData.messages)) return;

  const now = new Date();
  const curMonth = now.getMonth() + 1;
  const nextMonth = curMonth === 12 ? 1 : (curMonth + 1);
  const todayYmd = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');

  const targetThreads = {};
  histData.messages.forEach(function(m) {
    const txt = String(m.text || '');
    if (txt.includes('超募集MTGログ') && (!m.thread_ts || m.thread_ts === m.ts)) {
      const match = txt.match(/(\d{1,2})月\s*(前半|後半)/);
      if (match) {
        const mNum = Number(match[1]);
        if (mNum === curMonth || mNum === nextMonth) {
          const key = mNum + '月' + match[2];
          if (!targetThreads[key]) targetThreads[key] = m.ts;
        }
      }
    }
  });

  let updatedCount = 0;

  Object.keys(targetThreads).forEach(function(halfKey) {
    const parentTs = targetThreads[halfKey];
    const repUrl = 'https://slack.com/api/conversations.replies?channel=' + encodeURIComponent(MTG_THREAD_SLACK_CHANNEL_ID) + '&ts=' + encodeURIComponent(parentTs) + '&limit=250';
    const repRes = UrlFetchApp.fetch(repUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const repData = JSON.parse(repRes.getContentText());
    if (!repData.ok || !Array.isArray(repData.messages)) return;

    const replies = repData.messages.slice(1);

    replies.forEach(function(msg) {
      if (msg.bot_id !== DSKUN_BOT_ID) return;

      const rawText = String(msg.text || '');
      const dateMatch = rawText.match(/`(\d{1,2})月(\d{1,2})日\s*[\(（][日月火水木金土][\)）]`/);
      if (!dateMatch) return;

      // ★修正：年またぎ（12月に翌1月のスレッドを処理する場合）でも自動で正しい年を計算する
      const targetMonth = Number(dateMatch[1]);
      const targetYear = (curMonth === 12 && targetMonth <= 2) ? (now.getFullYear() + 1) : now.getFullYear();
      const ymd = targetYear + '/' + ('0' + targetMonth).slice(-2) + '/' + ('0' + dateMatch[2]).slice(-2);
      if (ymd < todayYmd) return;

      const built = buildUpdatedThreadMessageText3_3_(rawText, ymd, mailHistoryIndex);
      const normalizedRaw = rawText.replace(/&gt;/g, '>').trim();

      if (built.hasResult && built.newText.trim() !== normalizedRaw) {
        const upRes = UrlFetchApp.fetch('https://slack.com/api/chat.update', {
          method: 'post',
          contentType: 'application/json; charset=utf-8',
          headers: { Authorization: 'Bearer ' + token },
          payload: JSON.stringify({
            channel: MTG_THREAD_SLACK_CHANNEL_ID,
            ts: msg.ts,
            text: built.newText
          }),
          muteHttpExceptions: true
        });
        const upData = JSON.parse(upRes.getContentText());
        if (upData && upData.ok) {
          updatedCount++;
          Utilities.sleep(350);
        }
      }
    });
  });

  Logger.log('✅ C0AJTMLF0UD スレッド内の既存メッセージ上書き編集完了: ' + updatedCount + ' 件更新');
}

function buildUpdatedThreadMessageText3_3_(rawText, ymd, mailHistoryIndex) {
  const lines = String(rawText || '').replace(/&gt;/g, '>').split('\n');
  const coreLines = [];
  lines.forEach(function(l) {
    const trimmed = l.trim();
    if (trimmed.startsWith('⇒') || trimmed.startsWith('⇨') || trimmed.startsWith('→')) return;
    coreLines.push(l.replace(/^>\s?/, ''));
  });

  const cleanCore = coreLines.join('\n').trim();
  const clinicMatch = cleanCore.match(/【([^】]+)】\s*(\d{1,2}:\d{2}-\d{1,2}:\d{2})/);
  if (!clinicMatch) return { hasResult: false, newText: rawText };

  const cleanC = cleanClinicBaseName_(clinicMatch[1]);
  const slotTimeRange = clinicMatch[2];
  const startHour = Number(slotTimeRange.split(':')[0]);
  const timeBand = startHour < 13 ? '午前' : (startHour < 18 ? '午後' : '夜間');

  const docs = [];
  const docRe = /([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠]{2,6})先生/g;
  let dm;
  while ((dm = docRe.exec(cleanCore)) !== null) {
    if (!docs.includes(dm[1])) docs.push(dm[1]);
  }

  const accepts = [];
  const checks = [];
  const declines = [];

  const applyIndex = mailHistoryIndex ? mailHistoryIndex._applyIndex : null;
  const matchedApps = (applyIndex && typeof findMatchingShiftApplicationsForSlot3_3d_ === 'function')
    ? findMatchingShiftApplicationsForSlot3_3d_(ymd, cleanC, slotTimeRange, applyIndex)
    : [];

  // 異体字の揺れも含めて同一医師と判定できるように normKeyForDoc_ ベースでセット化
  const appliedDocKeys = new Set(matchedApps.map(function(a) { return normKeyForDoc_(a.doctorName); }));

  docs.forEach(function(dName) {
    const surname = extractDoctorSurnameForSlack3_3_(dName);
    const dKey = normKeyForDoc_(dName);

    // ① 交渉していた先生本人がサイトからシフト応募してくれた場合
    if (appliedDocKeys.has(dKey)) {
      accepts.push('⇒' + surname + '先生 応募あり・勤務了承 ' + cleanC + timeBand + 'は充足');
      return;
    }

    const st = evaluateDoctorAdjustmentStatusObj1b2_(dName, '', ymd, cleanC, mailHistoryIndex);
    if (st.code === 'ACCEPT') {
      const actWord = cleanCore.includes('移動') ? '変更了承' : (cleanCore.includes('延長') ? '延長了承' : (cleanCore.includes('早出') ? '早出了承' : '勤務了承'));
      accepts.push('⇒' + surname + '先生 ' + actWord + ' ' + cleanC + timeBand + 'は充足');
    } else if (st.code === 'DECLINE') {
      declines.push(surname + '先生');
    } else if (st.code === 'CHECK') {
      checks.push('⇒⚠️' + surname + '先生 要確認（' + st.reason + '）');
    }
  });

  // ② 交渉していた先生以外の別医師から当該枠へシフト応募が入った場合（分割応募での重複出力も防止）
  const seenExternalAppKeys = new Set();
  matchedApps.forEach(function(app) {
    const appKey = normKeyForDoc_(app.doctorName);
    if (!appKey || seenExternalAppKeys.has(appKey)) return;
    seenExternalAppKeys.add(appKey);

    const isAlreadyInDocs = docs.some(function(d) { return normKeyForDoc_(d) === appKey; });
    if (!isAlreadyInDocs) {
      const appSurname = extractDoctorSurnameForSlack3_3_(app.doctorName);
      const pendingDocs = docs.filter(function(d) {
        const st = evaluateDoctorAdjustmentStatusObj1b2_(d, '', ymd, cleanC, mailHistoryIndex);
        return st.code === 'WAIT' || st.code === 'CHECK';
      }).map(function(d) { return extractDoctorSurnameForSlack3_3_(d) + '先生'; });

      const noteSuffix = pendingDocs.length > 0 ? ('（' + pendingDocs.join('・') + 'はご放念連絡要確認）') : '';
      accepts.push('⇒' + appSurname + '先生より応募あり、' + cleanC + timeBand + 'は充足' + noteSuffix);
    }
  });

  const resultLines = [];
  accepts.forEach(function(a) { resultLines.push(a); });
  checks.forEach(function(c) { resultLines.push(c); });
  if (declines.length > 0) {
    if (declines.length <= 2) {
      declines.forEach(function(d) { resultLines.push('⇒' + d + ' 勤務不可'); });
    } else {
      resultLines.push('⇒' + declines.join('・') + ' 勤務不可');
    }
  }

  if (resultLines.length === 0) {
    return { hasResult: false, newText: rawText };
  }

  const quotedCore = cleanCore.split('\n').map(function(l) { return '> ' + l; }).join('\n');
  const newText = quotedCore + '\n' + resultLines.join('\n');
  return { hasResult: true, newText: newText };
}

function extractDoctorSurnameForSlack3_3_(fullName) {
  const clean = String(fullName || '').replace(/[\s\u3000]+/g, '').replace(/先生$/g, '').trim();
  if (clean.length <= 2) return clean;

  const threeCharSurnames = ['宇都宮', '宇治田', '佐々木', '長谷川', '五十嵐', '小野寺', '大久保', '小笠原', '久保田', '加賀谷', '喜多村', '久野木', '勅使河原'];
  for (let i = 0; i < threeCharSurnames.length; i++) {
    if (clean.startsWith(threeCharSurnames[i])) return threeCharSurnames[i];
  }

  const oneCharSurnames = ['馬', '林', '森', '原', '東', '関', '金', '李', '槙', '長', '澤', '瀧', '南', '西', '北', '辻', '堀', '岸', '奥'];
  if (clean.length === 3 && oneCharSurnames.includes(clean.charAt(0)) && !/^(林田|森田|森川|森下|森本|原田|東山|関根|関口|関島|金野|金子|金谷|長尾|長崎|長嶋|澤田|瀧本|西村|西田|西野|西原|北村|北瀬|北川|堀川|堀部|奥山)/.test(clean)) {
    return clean.charAt(0);
  }

  return clean.slice(0, 2);
}