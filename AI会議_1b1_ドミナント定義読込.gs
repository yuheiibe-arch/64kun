/**
 * ============================================================================
 * AI会議_1b1_ドミナント定義読込.gs
 * （Googleドキュメントの拠点間移動・ドミナント定義書を動的パースして参照するモジュール）
 * ============================================================================
 */

const DOMINANT_DOC_CONFIG = {
  DOC_ID: '1XUx2cLqN2P2RhUakpLEOn4xFzv7XnqQ_Jqm376xw1jU',
  CACHE_KEY: 'CAPS_DOMINANT_MATRIX_V3',
  CACHE_TTL_SEC: 21600
};

function normKeyForDoc_(name) {
  return String(name || '')
    .normalize('NFKC')
    .replace(/[\(（\[【].*$/g, '')
    .replace(/[:：].*$/, '')
    .replace(/[\)）\]】★⚠️📩・]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[﨑嵜]/g, '崎').replace(/髙/g, '高').replace(/[瀧泷]/g, '滝')
    .replace(/[邊邉]/g, '辺').replace(/[齋齊斉]/g, '斎').replace(/澤/g, '沢')
    .replace(/廣/g, '広').replace(/國/g, '国').replace(/[嶋嶌]/g, '島')
    .replace(/[栁柳]/g, '柳').replace(/[濵濱]/g, '浜').replace(/櫻/g, '桜')
    .replace(/眞/g, '真').replace(/惠/g, '恵')
    .trim();
}

function cleanClinicBaseName_(raw) {
  return String(raw || '')
    .replace(/^[★☆\s]+/, '')
    .replace(/※.*$/g, '')
    .replace(/【|】/g, '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/小児科|内科/g, '')
    .replace(/越谷LT/g, '越谷レイクタウン')
    .replace(/流山おおたか$/g, '流山おおたかの森')
    .replace(/幕張BP/g, '幕張ベイパーク')
    .replace(/千葉NT中央|千葉NT/g, '千葉ニュータウン')
    .replace(/セブンパーク天美/g, '天美')
    .replace(/新鎌ヶ谷/g, '新鎌ケ谷')
    .replace(/新百合ケ丘/g, '新百合ヶ丘')
    .replace(/八千代緑ヶ丘/g, '八千代緑が丘')
    .replace(/茅ケ崎/g, '茅ヶ崎')
    .trim();
}

function inferAreaForDoc_(clinicName) {
  const c = String(clinicName || '');
  // ★「草加松原」が大阪府松原市（天美）の「松原」に誤爆しないよう埼玉を先に判定
  if (/越谷|川口|川越|志木|所沢|南浦和|与野|草加松原|草加|浦和|吉川/.test(c)) {
    return '埼玉';
  }
  if (/関西|大阪|兵庫|京都|奈良|滋賀|和歌山|鶴見緑地|西宮|千里|高槻|あべの|天美|東岸和田|今里|門真|守口|枚方|吹田|豊中|茨木|八尾|堺|尼崎|伊丹|宝塚|川西|明石|神戸|姫路|生駒|橿原|草津|大津|阿波座|長吉長原/.test(c)) {
    return '関西';
  }
  if (/海老名|高田|武蔵小杉|相模大野|相模原|新百合|茅ヶ崎|茅ケ崎|天王町|小田栄|東戸塚|宮前平|川崎|横浜/.test(c)) {
    return '神奈川';
  }
  if (/二和向台|柏の葉|柏|村上|稲毛|流山|新鎌|幕張|八千代|妙典|千葉|新浦安|浦安|市川|本八幡|船橋/.test(c)) {
    return '千葉';
  }
  if (/つくば|守谷|茨城/.test(c)) {
    return '茨城';
  }
  return '東京';
}

function loadDominantMatrixFromDoc_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(DOMINANT_DOC_CONFIG.CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const matrix = {};
  try {
    const doc = DocumentApp.openById(DOMINANT_DOC_CONFIG.DOC_ID);
    const body = doc.getBody();
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
      const table = el.asTable();
      if (table.getNumRows() < 2) continue;

      const h2 = table.getRow(0).getCell(2).getText().trim();
      if (!h2.includes('【優先度S】')) continue;

      for (let r = 1; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getNumCells() < 5) continue;

        const fromClinic = cleanClinicBaseName_(row.getCell(1).getText().split('\n')[0]);
        if (!fromClinic) continue;
        if (!matrix[fromClinic]) matrix[fromClinic] = {};

        parsePriorityCellIntoMatrix_(matrix[fromClinic], row.getCell(2).getText(), 'S', 100);
        parsePriorityCellIntoMatrix_(matrix[fromClinic], row.getCell(3).getText(), 'A', 80);
        parsePriorityCellIntoMatrix_(matrix[fromClinic], row.getCell(4).getText(), 'B', 60);
      }
    }

    if (Object.keys(matrix).length > 0) {
      cache.put(DOMINANT_DOC_CONFIG.CACHE_KEY, JSON.stringify(matrix), DOMINANT_DOC_CONFIG.CACHE_TTL_SEC);
    }
  } catch (e) {
    Logger.log('⚠️ ドミナント定義書ドキュメントの読み込みエラー: ' + e.message);
  }
  return matrix;
}

function parsePriorityCellIntoMatrix_(targetObj, cellText, rank, baseScore) {
  const s = String(cellText || '').trim();
  if (!s || s.startsWith('なし')) return;

  const itemRegex = /([^,、\n\(\uff08]+?)[\(\uff08]([^)\uff09]+)[\)\uff09]/g;
  let m;
  while ((m = itemRegex.exec(s)) !== null) {
    const destName = cleanClinicBaseName_(m[1]);
    if (!destName || destName.includes('全域')) continue;

    const timePart = m[2].replace(/\[.*?\]/g, '');
    const nums = (timePart.match(/\d+/g) || []).map(Number).filter(function(n) { return n > 0 && n < 300; });
    const minMinutes = nums.length > 0 ? Math.min.apply(null, nums) : (rank === 'S' ? 30 : (rank === 'A' ? 50 : 65));

    const calcScore = Math.round(baseScore - 0.2 * minMinutes);
    targetObj[destName] = {
      rank: rank,
      score: calcScore,
      estMinutes: minMinutes
    };
  }
}

function evaluateClinicPairMobility_(clinic1, clinic2, matrixOpt) {
  const c1 = cleanClinicBaseName_(clinic1);
  const c2 = cleanClinicBaseName_(clinic2);
  const a1 = inferAreaForDoc_(c1);
  const a2 = inferAreaForDoc_(c2);

  if ((a1 === '関西') !== (a2 === '関西')) {
    return { rank: 'NG', score: -9999, estMinutes: 999 };
  }
  if (c1 === c2) {
    return { rank: 'S', score: 120, estMinutes: 0 };
  }

  const matrix = matrixOpt || loadDominantMatrixFromDoc_();
  const hasMatrixLoaded = matrix && Object.keys(matrix).length > 0;

  if (hasMatrixLoaded) {
    if (matrix[c1] && matrix[c1][c2]) return matrix[c1][c2];
    if (matrix[c2] && matrix[c2][c1]) return matrix[c2][c1];

    const findEntry = function(srcMap, destKey) {
      if (!srcMap) return null;
      const keys = Object.keys(srcMap);
      for (let i = 0; i < keys.length; i++) {
        if (keys[i].includes(destKey) || destKey.includes(keys[i])) return srcMap[keys[i]];
      }
      return null;
    };
    const hit1 = findEntry(matrix[c1], c2);
    if (hit1) return hit1;
    const hit2 = findEntry(matrix[c2], c1);
    if (hit2) return hit2;

    return { rank: 'NG', score: 0, estMinutes: 100 };
  }

  if (a1 === a2) {
    return { rank: 'A', score: 75, estMinutes: 50 };
  }
  return { rank: 'NG', score: 0, estMinutes: 100 };
}