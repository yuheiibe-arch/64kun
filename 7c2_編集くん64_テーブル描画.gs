/**
 * ============================================================================
 * 7c2_編集くん64_テーブル描画.gs
 * （メインテーブル・おまとめ打診テーブル・編集後イメージ・医籍番号タグのHTML描画）
 * ============================================================================
 */

function getHenshukun64ScriptPart2_() {
  return `
    function formatPlansHtml64(col2Text) {
      const rawLines = String(col2Text || '').split('\\n');
      const planHtmls = [];
      const noteLines = [];

      rawLines.forEach(line => {
        const s = line.trim();
        if (!s) return;
        if (s.startsWith('💡') || s.startsWith('・')) {
          noteLines.push(s);
        } else {
          const m = s.match(/^(【案[123]】)(.*)$/);
          if (m) planHtmls.push('<div class="plan-line"><b>' + m[1] + '</b>' + m[2] + '</div>');
          else planHtmls.push('<div class="plan-line">' + s + '</div>');
        }
      });

      if (noteLines.length > 0) {
        planHtmls.push('<div class="note-line">' + noteLines.join('\\n') + '</div>');
      }
      return planHtmls.join('');
    }

    function buildAfterDiffHtml64(row) {
      let mainTxt = '【編集後】' + row.curDecision + ' ／ 担当:' + (row.curAssignee || '未定');
      if (Array.isArray(row.customDocs) && row.customDocs.length > 0) {
        const idList = row.customDocs.map(d => d.name + (d.licenseNo ? ':' + d.licenseNo : '')).join(', ');
        mainTxt += '<br><span class="diff-lic-sub">🆔 医籍番号: ' + idList + '</span>';
      }
      return mainTxt;
    }

    function buildPickedDocsTagsHtml64(slotKey, customDocs) {
      if (!Array.isArray(customDocs) || customDocs.length === 0) return '';
      return customDocs.map((d, idx) => {
        const licLabel = d.licenseNo ? ' (医籍:' + d.licenseNo + ')' : ' (番号未登録)';
        return '<span class="picked-doc-tag">' +
          '🆔 ' + d.name + licLabel +
          '<span class="del-btn" title="削除" onclick="removeCustomDoc64(\\'' + slotKey + '\\',' + idx + ')">×</span>' +
        '</span>';
      }).join('');
    }

    function renderStageTable() {
      const lockedByOthers = getLockedStaffSetByOthers();
      const rows = state.slotRows.filter(r => r.dateLabel === state.activeDateTab && isRowMatchRegionFilter(r));

      if (rows.length === 0) {
        document.getElementById('stageTableBody').innerHTML =
          '<tr><td colspan="5" style="text-align:center;padding:24px;font-weight:bold;">該当する不在枠はありません（エリアまたは日付を切り替えてください）</td></tr>';
        return;
      }

      const actTypeOpts = CUSTOM_ACTION_TYPES.map(t => '<option value="' + t + '">' + t + '</option>').join('');

      const html = rows.map(row => {
        const isEdited = (row.curDecision !== row.origDecision) || (row.curAssignee !== row.origAssignee);
        const isStay = (row.curDecision === 'ステイ');
        const isPreset = DECISION_PRESETS.includes(row.curDecision);
        const selVal = isPreset ? row.curDecision : 'カスタム入力';
        const customVal = isPreset ? '' : row.curDecision;

        const decOpts = DECISION_PRESETS.map(opt =>
          '<option value="' + opt + '"' + (opt === selVal ? ' selected' : '') + '>' + opt + '</option>'
        ).join('');

        const staffOpts = ['<option value="">(未定)</option>'].concat(state.staffList.map(s => {
          const locked = lockedByOthers.has(s) && s !== row.curAssignee;
          const label = locked ? '🔒' + s + '(他作業中)' : s;
          return '<option value="' + s + '"' + (s === row.curAssignee ? ' selected' : '') + (locked ? ' disabled' : '') + '>' + label + '</option>';
        })).join('');

        return '<tr class="slot-row' + (isEdited ? ' is-edited' : '') + (isStay ? ' is-stay' : '') + '" id="tr_' + row.slotKey + '">' +
          '<td>' +
            '<div style="font-size:11.5px;font-weight:bold;color:#475569;margin-bottom:4px;">' + row.dateLabel + '</div>' +
            '<div><span class="region-tag' + (row.macroArea === '関西' ? ' kansai' : '') + '">' + row.clinicGroup + '</span>' +
            '<span class="clinic-name">' + row.clinicTitle + '</span></div>' +
            '<div class="time-wage">' + row.timeAndWage + '</div>' +
          '</td>' +
          '<td>' + formatPlansHtml64(row.col2Plans) + '</td>' +
          '<td>' +
            '<div class="diff-box">' +
              '<div class="diff-before">【編集前】' + row.origDecision + ' ／ 担当:' + (row.origAssignee || '未定') + '</div>' +
              '<div class="diff-after' + (isEdited ? ' changed' : '') + '" id="after_' + row.slotKey + '">' +
                buildAfterDiffHtml64(row) +
              '</div>' +
            '</div>' +
          '</td>' +
          '<td>' +
            '<select class="n64-select" style="width:100%;" id="sel_' + row.slotKey + '" onchange="onSlotEdit64(\\'' + row.slotKey + '\\')">' + decOpts + '</select>' +
            '<div class="custom-builder" id="cbox_' + row.slotKey + '" style="display:' + (selVal === 'カスタム入力' ? 'block' : 'none') + ';">' +
              '<div class="search-bar-row">' +
                '<input type="text" class="n64-input" id="srch_' + row.slotKey + '" style="flex:1;padding:4px 7px;font-size:12px;" ' +
                  'placeholder="🔍 医師名検索(例:須貝)" oninput="onSearchDocInput64(\\'' + row.slotKey + '\\')" onfocus="onSearchDocInput64(\\'' + row.slotKey + '\\')">' +
                '<select class="n64-select" id="act_' + row.slotKey + '" style="padding:4px 5px;font-size:12px;">' + actTypeOpts + '</select>' +
              '</div>' +
              '<div class="suggest-dropdown" id="sug_' + row.slotKey + '"></div>' +
              '<input type="text" class="n64-input" id="cus_' + row.slotKey + '" style="width:100%;font-size:12.5px;padding:5px 7px;" ' +
                'placeholder="ドキュメント記載文(医籍番号は含めません)" value="' + customVal.replace(/"/g, '&quot;') + '" oninput="onSlotEdit64(\\'' + row.slotKey + '\\')">' +
              '<div class="picked-tags-box" id="tags_' + row.slotKey + '">' + buildPickedDocsTagsHtml64(row.slotKey, row.customDocs) + '</div>' +
            '</div>' +
          '</td>' +
          '<td>' +
            '<select class="n64-select" style="width:100%;" id="stf_' + row.slotKey + '" onchange="onSlotEdit64(\\'' + row.slotKey + '\\')">' + staffOpts + '</select>' +
          '</td>' +
        '</tr>';
      }).join('');

      document.getElementById('stageTableBody').innerHTML = html;
    }

    function renderBundlesTable() {
      const html = state.bundleRows.map(b => {
        const isEdited = (b.curVal !== b.origVal);
        return '<tr class="slot-row' + (isEdited ? ' is-edited' : '') + '">' +
          '<td style="font-weight:900;color:#0d47a1;">' + b.doctor + '</td>' +
          '<td style="white-space:pre-wrap;font-size:12px;">' + b.slots + '</td>' +
          '<td style="font-size:12px;">' + b.note + '</td>' +
          '<td>' +
            '<div class="diff-box">' +
              '<div class="diff-before">【編集前】' + b.origVal + '</div>' +
              '<div class="diff-after' + (isEdited ? ' changed' : '') + '" id="bafter_' + b.bundleKey + '">【編集後】' + b.curVal + '</div>' +
            '</div>' +
          '</td>' +
          '<td>' +
            '<input type="text" class="n64-input" style="width:100%;" value="' + b.curVal.replace(/"/g, '&quot;') + '" ' +
              'placeholder="例: 採用 佐藤 / ステイ" oninput="onBundleEdit64(\\'' + b.bundleKey + '\\', this.value)">' +
          '</td>' +
        '</tr>';
      }).join('');
      document.getElementById('bundlesTableBody').innerHTML = html || '<tr><td colspan="5">該当なし</td></tr>';
    }
  `;
}