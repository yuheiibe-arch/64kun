/**
 * ============================================================================
 * 6c_2_編集くん64_セル整形とマーカー.gs（6c系 3分割の2）
 * ・Googleドキュメントセルの段落二重化防止書込（setDocCellTextSafely64_）
 * ・推奨案テキストの重複行除去（dedupePlanCellLines64_）
 * ・⭕承諾案＝緑マーカー／選択打診案＝赤マーカーの自動ハイライト描画
 * ============================================================================
 */

function setDocCellTextSafely64_(cell, text) {
  const clean = String(text || '');
  while (cell.getNumChildren() > 1) {
    cell.removeChild(cell.getChild(cell.getNumChildren() - 1));
  }
  const first = cell.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const p = first.asParagraph();
    if (clean.length === 0) {
      p.clear();
    } else {
      p.setText(clean);
    }
  } else {
    const p = cell.appendParagraph('');
    if (clean.length > 0) {
      p.setText(clean);
    }
    cell.removeChild(first);
  }
}

function dedupePlanCellLines64_(rawText) {
  const normalized = String(rawText || '').replace(/(【案[123１２３]】[^\n]+?)\1+/g, '$1');
  const lines = normalized.split(/\r?\n/);
  const seenPlanNums = new Set();
  const seenExactLines = new Set();
  const out = [];

  lines.forEach(function(line) {
    const trimmed = line.trim();
    if (!trimmed) return;
    const pm = trimmed.match(/^【案([123１２３])】/);
    if (pm) {
      const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
      if (seenPlanNums.has(pNum)) return;
      seenPlanNums.add(pNum);
      out.push(trimmed);
      return;
    }
    if (seenExactLines.has(trimmed)) return;
    seenExactLines.add(trimmed);
    out.push(trimmed);
  });

  return out.join('\n');
}

function applySelectedPlanMarkerToDocCell64_(cell1, rawC1Text, rawDecisionText) {
  const c1Text = dedupePlanCellLines64_(rawC1Text);
  setDocCellTextSafely64_(cell1, c1Text);
  if (!c1Text) return;

  const textObj = cell1.editAsText();
  const len = c1Text.length;
  textObj.setFontSize(0, len - 1, 8.5);
  textObj.setBold(0, len - 1, false);
  textObj.setBackgroundColor(0, len - 1, null);

  if (typeof highlightKeywordsInText_ === 'function') {
    highlightKeywordsInText_(textObj);
  }

  const fullDecStr = String(rawDecisionText || '').trim();
  if (!fullDecStr) return;

  const planLineMap = {};
  c1Text.split('\n').forEach(function(line) {
    const pm = line.trim().match(/^【案([123１２３])】(.*)$/);
    if (pm) {
      const pNum = pm[1].replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
      planLineMap[pNum] = pm[2].trim();
    }
  });

  const acceptedPlanNums = new Set();
  fullDecStr.split('\n').forEach(function(line) {
    const s = line.trim();
    if (s.startsWith('→') && s.includes('⭕')) {
      const m = s.match(/\(案([123１２３])\)/g);
      if (m) {
        m.forEach(function(tok) {
          const num = tok.replace(/[^123１２３]/g, '').replace(/[１２３]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
          if (num) acceptedPlanNums.add(num);
        });
      } else {
        const docMatch = s.match(/⭕\s*([^\(（：:\s]+)/);
        if (docMatch && typeof normKeyForDoc_ === 'function') {
          const dKey = normKeyForDoc_(docMatch[1]);
          if (dKey) {
            ['1', '2', '3'].forEach(function(pNum) {
              if (planLineMap[pNum] && normKeyForDoc_(planLineMap[pNum]).includes(dKey)) {
                acceptedPlanNums.add(pNum);
              }
            });
          }
        }
      }
    }
  });

  const firstLine = fullDecStr.split('\n')[0]
    .replace(/\s*[（\(]メモ[：:].*?[）\)]$/g, '')
    .replace(/【特別時給.*?】/g, '')
    .trim();

  const selectedPlanNums = new Set();
  const subDocChoiceByPlan = {};

  if (firstLine && !firstLine.startsWith('→') && !/全案|すべて|全部|ステイ|要検討|医師未選択/.test(firstLine)) {
    const tokens = firstLine.split(/[,、，\/／\n・\+＋]|(?:\s+と\s+)/).map(function(t) { return t.trim(); }).filter(Boolean);
    tokens.forEach(function(tok) {
      const specMatch = tok.match(/^(?:【?案)?([123１２３])(?:】)?(?:\(([^)]+)\))?$/);
      if (specMatch) {
        const pNum = specMatch[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
        selectedPlanNums.add(pNum);
        const subText = specMatch[2] ? specMatch[2].trim() : '';
        if (subText) {
          if (subText.includes('両方')) subDocChoiceByPlan[pNum] = 'BOTH';
          else if (/^[①1１]/.test(subText)) subDocChoiceByPlan[pNum] = '1';
          else if (/^[②2２]/.test(subText)) subDocChoiceByPlan[pNum] = '2';
        }
      } else {
        const multiMatches = tok.match(/案([123１２３])/g);
        if (multiMatches && tok.replace(/案[123１２３]|[\s、,・と]/g, '') === '') {
          multiMatches.forEach(function(mm) {
            const pNum = mm.replace('案', '').replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
            selectedPlanNums.add(pNum);
          });
        }
      }
    });
  }

  acceptedPlanNums.forEach(function(num) { selectedPlanNums.add(num); });
  if (selectedPlanNums.size === 0) return;

  const COLOR_SELECTED_RED = '#ffd6d6';
  const COLOR_ACCEPTED_GREEN = '#bbf7d0';

  const lines = c1Text.split('\n');
  let curOffset = 0;

  lines.forEach(function(line) {
    const lineLen = line.length;
    const lineStart = curOffset;
    const lineEnd = curOffset + lineLen - 1;
    curOffset += lineLen + 1;

    if (lineLen <= 0) return;
    const m = line.match(/^【案([123１２３])】/);
    if (!m) return;

    const pNum = m[1].replace(/[１２３]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); });
    if (!selectedPlanNums.has(pNum)) return;

    const targetColor = acceptedPlanNums.has(pNum) ? COLOR_ACCEPTED_GREEN : COLOR_SELECTED_RED;
    textObj.setBackgroundColor(lineStart, lineEnd, targetColor);
    textObj.setBold(lineStart, lineEnd, true);

    const subChoice = subDocChoiceByPlan[pNum];
    if (subChoice === '1' || subChoice === '2') {
      const pairRegex = /(①[^\/／\n]+?)(\s*[\/／]\s*)(②[^）\)\n]+)/;
      const pm = line.match(pairRegex);
      if (pm && typeof pm.index === 'number') {
        const seg1Start = lineStart + pm.index;
        const seg1End = seg1Start + pm[1].length + pm[2].length - 1;
        const seg2Start = lineStart + pm.index + pm[1].length;
        const seg2End = seg2Start + pm[2].length + pm[3].length - 1;

        if (subChoice === '1' && seg2Start <= seg2End) {
          textObj.setBackgroundColor(seg2Start, seg2End, null);
          textObj.setBold(seg2Start, seg2End, false);
        } else if (subChoice === '2' && seg1Start <= seg1End) {
          textObj.setBackgroundColor(seg1Start, seg1End, null);
          textObj.setBold(seg1Start, seg1End, false);
        }
      }
    }
  });
}