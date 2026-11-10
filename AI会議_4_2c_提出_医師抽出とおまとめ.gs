/**
 * ============================================================================
 * AI会議_4_2c_提出_医師抽出とおまとめ.gs（3分割の3）
 * ・アクション文からの種別・移動元拠点・正式医師名の抽出
 * ・おまとめ打診テーブル行の解析（parseBundleTableRow_）
 * ============================================================================
 */

function parseBundleTableRow_(defaultYear, docColText, slotsColText, decColText, resolver, isFullTimeChecker) {
  const separated = extractDecisionHeaderAndArrows4_2_(decColText);
  const dec = separated.header;
  if (isStayOrSkipDecision4_2_(dec) || dec === '不採用') return [];

  const rawDoc = String(docColText || '').replace(/\r\n|\r/g, ' ').replace(/[\(（].*?[\)）]/g, '').replace(/先生/g, '').trim();
  const resolvedDoc = resolver.resolveDoctor(rawDoc, '', '');
  const isFt = isFullTimeChecker ? isFullTimeChecker(resolvedDoc.name, resolvedDoc.licenseNo, docColText) : false;
  const assigneeFromDec = dec.replace(/採用|OK|◯|○|[\/／・,\s ]/g, '').trim();
  const cleanAssignee = cleanAssigneeText4_2_(assigneeFromDec, resolver);

  const items = [];
  let currentClinic = '';
  let currentDept = '小児科';

  String(slotsColText || '').replace(/\r\n|\r/g, '\n').split('\n').map(function(l) { return l.trim(); }).filter(Boolean).forEach(function(line) {
    if (line.startsWith('■')) {
      const rawC = line.replace(/^■\s*/, '').trim();
      currentDept = rawC.includes('内科') ? '内科' : '小児科';
      currentClinic = rawC.replace(/[\(（].*?[\)）]/g, '').trim();
      return;
    }
    const dm = line.match(/(\d{1,2})\/(\d{1,2})[^\d]*(\d{1,2}:\d{2}-\d{1,2}:\d{2}(?:\s*\/\s*\d{1,2}:\d{2}-\d{1,2}:\d{2})*)/);
    if (dm && currentClinic) {
      const dateStr = defaultYear + '/' + ('0' + dm[1]).slice(-2) + '/' + ('0' + dm[2]).slice(-2);
      dm[3].split('/').map(function(t) { return t.trim(); }).filter(Boolean).forEach(function(tRange) {
        items.push({
          date: dateStr,
          clinic_name: currentClinic,
          raw_clinic_name: currentClinic,
          department: currentDept,
          target_time: tRange,
          time_category: '',
          approach_type: '個別アプローチ',
          target_doctor: resolvedDoc.name,
          license_no: resolvedDoc.licenseNo || '',
          is_full_time: isFt,
          from_clinic: '',
          is_special_wage: false,
          special_wage_addon: '',
          assignee: cleanAssignee,
          memo: '',
          comment: '',
          evidence: '[ドキュメント確定:おまとめ打診] ' + resolvedDoc.name + '先生 (' + currentClinic + ' ' + dateStr + ' ' + tRange + ')'
        });
      });
    }
  });

  return items;
}

function parseSingleActionTextToItems_(text, dateStr, targetClinic, resolver) {
  const s = String(text || '')
    .replace(/\r\n|\r/g, ' ')
    .replace(/※残る.*?に、/g, '')
    .replace(/\[※.*?\]/g, '')
    .trim();
  if (!s) return [];

  if (s.includes('1診・併行募集') || s.includes('併行募集') || s.includes('埋まり次第')) {
    return [];
  }

  let approachType = '個別アプローチ';
  let fromClinic = '';
  let rawDoctors = [];

  if (s.includes('2診目崩し') || s.includes('移動')) {
    approachType = '勤務地変更相談';
    const fromMatch = s.match(/】([^\(（\s]+?)[（\(](?:※|[12]診)/) || s.match(/([^\s【】]+?)(?:より|から)/);
    if (fromMatch) fromClinic = fromMatch[1].trim();
    const insideParen = s.match(/[12]診(?:目)?[：:]([^）\)]+)[）\)]/);
    rawDoctors = insideParen ? extractDoctorNamesFromSegment_(insideParen[1]) : extractDoctorNamesFromSegment_(s.replace(/^.*?から/, ''));
  } else if (s.includes('延長')) {
    approachType = '延長相談';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  } else if (s.includes('早出') || s.includes('前倒')) {
    approachType = '早出相談';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  } else {
    approachType = '個別アプローチ';
    rawDoctors = extractDoctorNamesFromSegment_(s);
  }

  if (rawDoctors.length === 0) return [{ approachType: approachType, doctorName: '', licenseNo: '', fromClinic: fromClinic }];

  return rawDoctors.map(function(rawDoc) {
    const resolved = resolver.resolveDoctor(rawDoc, dateStr, fromClinic || targetClinic);
    const finalFromClinic = (approachType === '勤務地変更相談' && !fromClinic && resolved.dayClinic && resolved.dayClinic !== targetClinic)
      ? resolved.dayClinic
      : fromClinic;
    return { approachType: approachType, doctorName: resolved.name, licenseNo: resolved.licenseNo, fromClinic: finalFromClinic };
  });
}

function extractDoctorNamesFromSegment_(seg) {
  const base = String(seg || '')
    .replace(/\r\n|\r/g, ' ')
    .replace(/\[.*?\]|［.*?］/g, '')
    .replace(/[（\(].*?[）\)]/g, '')
    .replace(/^(?:【?案\s*[123１２３]?】?\s*)+/g, '')
    .trim();

  if (!base) return [];

  return base.split(/[\/／]|または|・|,|、/)
    .map(function(part) {
      let p = part
        .replace(/^[①②③④123１２３:\s ]+/g, '')
        .replace(/^(?:前医の?|後医の?|[12]診(?:目)?[：:]|午前|午後|夜間|終日)+/g, '')
        .replace(/(?:[\s\u3000]*(?:他|ほか|外)\s*\d+\s*名.*$)/g, '')
        .trim();

      // 「〇〇先生...」の形式なら「先生」直前の人名を優先抽出
      const senseiMatch = p.match(/([一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠ァ-ヶー]{2,8})\s*先生/);
      if (senseiMatch) {
        return senseiMatch[1].trim();
      }

      // 「先生」が省略されている場合（「遠山移動」「宇都宮真司」など）は末尾の動詞・助詞を除去
      p = p
        .replace(/どちらか1名.*$/g, '')
        .replace(/(?:先生)?(?:へ|に|を)?(?:個別|追加|延長|早出|前倒し|移動|勤務地変更|バックアップ)?(?:アプローチ|打診|相談|依頼|交渉|確認|変更|延長|早出|移動).*$/g, '')
        .replace(/先生$/g, '')
        .replace(/[\s\u3000]+/g, '')
        .trim();

      return p;
    })
    .filter(function(p) {
      if (!p || p.length < 2 || p.length > 8) return false;
      if (/[\(\)（）\[\]【】→⇒⇨⚠️📩⭕❌\d円:：…\.]/.test(p)) return false;
      if (/^(?:スポット|紹介会社|常勤|非常勤|移動|延長|早出|個別|ステイ|特別時給|通常時給|両方|片方|相談|打診|要確認|返信待ち|辞退|承諾|充足|バックアップ|バックアップ医師)$/.test(p)) {
        return false;
      }
      return /^[一-龠々ヶケ﨑嵜髙瀧泷邊邉齋齊澤廣國嶋嶌栁濱櫻眞惠ァ-ヶー]+$/.test(p);
    });
}