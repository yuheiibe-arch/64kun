/**
 * ============================================================================
 * 7c_編集くん64_スクリプト.gs
 * （『🎮 編集くん 64』のクライアント側JavaScriptロジック）
 * ============================================================================
 */

function getHenshukun64ClientScript_() {
  return `<script>
    let state = {
      sessionId: '',
      myStaff: '',
      macroMode: 'ALL',
      selectedRegions: new Set(),
      activeDateTab: '',
      staffList: [],
      groups: [],
      areaToGroups: {},
      slotRows: [],
      bundleRows: [],
      activeLocks: {},
      changedSlots: {},
      changedBundles: {}
    };

    const DECISION_PRESETS = ['案1', '案2', '案3', '案1, 案2', '案1, 案3', '案2, 案3', '全案', 'ステイ', 'カスタム入力'];

    window.onload = function() {
      reloadAllData();
      setInterval(sendHeartbeat64, 18000);
    };

    function reloadAllData() {
      setStatus('⏳ データを読み込んでいます...', '#ffc72c');
      google.script.run
        .withSuccessHandler(initStudio64)
        .withFailureHandler(err => setStatus('❌ 読込エラー: ' + err.message, '#ff5252'))
        .getHenshukunInitialData();
    }

    function initStudio64(data) {
      if (!state.sessionId) state.sessionId = data.clientSessionId;
      state.staffList = data.staffList || [];
      state.groups = data.groups || [];
      state.areaToGroups = data.areaToGroups || {};
      state.slotRows = data.slotRows || [];
      state.bundleRows = data.bundleRows || [];
      state.activeLocks = data.activeLocks || {};

      document.getElementById('docLinkBox').innerHTML =
        '<a class="doc-link" href="' + data.docUrl + '" target="_blank">📄 ' + data.docName + ' を別タブで開く</a>';

      renderStaffSelectors();
      renderRegionChips();

      const dates = getFilteredUniqueDates();
      if (!state.activeDateTab && dates.length > 0) {
        state.activeDateTab = dates[0];
      }
      renderDateTabs();
      renderStageTable();
      renderBundlesTable();
      updateActivePlayersDisplay();
      updateEditCounter();
      setStatus('✅ 読み込み完了！ステージ（日付タブ）を選んで編集できます。', '#69f0ae');
    }

    function getLockedStaffSetByOthers() {
      const locked = new Set();
      Object.entries(state.activeLocks || {}).forEach(([sid, info]) => {
        if (sid !== state.sessionId && info.staff) {
          locked.add(info.staff);
        }
      });
      return locked;
    }

    function renderStaffSelectors() {
      const lockedByOthers = getLockedStaffSetByOthers();

      const buildOptions = (selectedVal, isMySelector) => {
        const head = isMySelector
          ? '<option value="">-- 自分の名前を選択 --</option>'
          : '<option value="">(担当者未定)</option>';

        return head + state.staffList.map(s => {
          const isLocked = lockedByOthers.has(s);
          const sel = (s === selectedVal) ? ' selected' : '';
          const dis = (isMySelector && isLocked) ? ' disabled' : '';
          const label = isLocked ? '🔒 ' + s + ' (作業中)' : s;
          return '<option value="' + s + '"' + sel + dis + '>' + label + '</option>';
        }).join('');
      };

      document.getElementById('myPlayerSelect').innerHTML = buildOptions(state.myStaff, true);
      document.getElementById('bulkStaffSelect').innerHTML = buildOptions(state.myStaff, false);
    }

    function onMyPlayerChange() {
      const picked = document.getElementById('myPlayerSelect').value;
      state.myStaff = picked;
      document.getElementById('bulkStaffSelect').value = picked;
      sendHeartbeat64();
    }

    function sendHeartbeat64() {
      if (!state.sessionId) return;
      google.script.run
        .withSuccessHandler(res => {
          if (!res.ok && res.conflictStaff) {
            alert('⚠️ 「' + res.conflictStaff + '」さんは現在別画面で作業中のためブロックされています！');
            state.myStaff = '';
            document.getElementById('myPlayerSelect').value = '';
          }
          state.activeLocks = res.activeLocks || {};
          if (res.latestDraft && res.latestDraft.slots) {
            state.slotRows.forEach(r => {
              const d = res.latestDraft.slots[r.slotKey];
              if (d && !state.changedSlots[r.slotKey]) {
                r.curDecision = d.decision;
                r.curAssignee = d.assignee;
              }
            });
          }
          renderStaffSelectors();
          updateActivePlayersDisplay();
        })
        .heartbeatHenshukunPlayer(state.sessionId, state.myStaff);
    }

    function updateActivePlayersDisplay() {
      const names = Object.values(state.activeLocks || {}).map(x => x.staff).filter(Boolean);
      const el = document.getElementById('activePlayersText');
      el.innerText = names.length > 0 ? '🟢 作業中: ' + names.join(' / ') : '🟢 重複作業なし';
    }

    function setMacroMode(mode) {
      state.macroMode = mode;
      ['ALL', '関東', '関西'].forEach(m => {
        const btn = document.getElementById('mode_' + m);
        if (btn) btn.classList.toggle('active-mode', m === mode);
      });
      state.selectedRegions.clear();
      renderRegionChips();

      const dates = getFilteredUniqueDates();
      if (state.activeDateTab !== 'BUNDLES' && !dates.includes(state.activeDateTab)) {
        state.activeDateTab = dates[0] || '';
      }
      renderDateTabs();
      renderStageTable();
    }

    function renderRegionChips() {
      let targetGroups = state.groups;
      if (state.macroMode !== 'ALL' && state.areaToGroups[state.macroMode]) {
        targetGroups = state.areaToGroups[state.macroMode];
      }
      const html = targetGroups.map(g => {
        const isSel = state.selectedRegions.has(g);
        return '<span class="chip' + (isSel ? ' selected' : '') + '" onclick="toggleRegionChip(\\'' + g + '\\')">' +
          (isSel ? '✓ ' : '') + g + '</span>';
      }).join('');
      document.getElementById('regionChipsBox').innerHTML = html;
    }

    function toggleRegionChip(groupName) {
      if (state.selectedRegions.has(groupName)) {
        state.selectedRegions.delete(groupName);
      } else {
        state.selectedRegions.add(groupName);
      }
      renderRegionChips();
      const dates = getFilteredUniqueDates();
      if (state.activeDateTab !== 'BUNDLES' && !dates.includes(state.activeDateTab)) {
        state.activeDateTab = dates[0] || '';
      }
      renderDateTabs();
      renderStageTable();
    }

    function clearRegionSelection() {
      state.selectedRegions.clear();
      renderRegionChips();
      renderDateTabs();
      renderStageTable();
    }

    function isRowMatchRegionFilter(row) {
      if (state.macroMode !== 'ALL' && row.macroArea !== state.macroMode) return false;
      if (state.selectedRegions.size > 0 && !state.selectedRegions.has(row.clinicGroup)) return false;
      return true;
    }

    function getFilteredUniqueDates() {
      const set = new Set();
      state.slotRows.forEach(r => {
        if (isRowMatchRegionFilter(r)) set.add(r.dateLabel);
      });
      return Array.from(set);
    }

    function renderDateTabs() {
      const dates = getFilteredUniqueDates();
      const tabsHtml = dates.map(d => {
        const count = state.slotRows.filter(r => r.dateLabel === d && isRowMatchRegionFilter(r)).length;
        const isAct = (state.activeDateTab === d);
        return '<div class="date-tab' + (isAct ? ' active' : '') + '" onclick="selectStageTab(\\'' + d + '\\')">' +
          '<span>📅 ' + d + '</span><span class="tab-count">' + count + '枠</span></div>';
      });

      const isBundleAct = (state.activeDateTab === 'BUNDLES');
      tabsHtml.push(
        '<div class="date-tab' + (isBundleAct ? ' active' : '') + '" onclick="selectStageTab(\\'BUNDLES\\')">' +
        '<span>🌟 複数日程おまとめ打診</span><span class="tab-count">' + state.bundleRows.length + '名</span></div>'
      );

      document.getElementById('dateTabsBar').innerHTML = tabsHtml.join('');
    }

    function selectStageTab(tabKey) {
      state.activeDateTab = tabKey;
      renderDateTabs();
      if (tabKey === 'BUNDLES') {
        document.getElementById('slotsStageView').style.display = 'none';
        document.getElementById('bundlesStageView').style.display = 'block';
      } else {
        document.getElementById('slotsStageView').style.display = 'block';
        document.getElementById('bundlesStageView').style.display = 'none';
        renderStageTable();
      }
    }

    function formatPlansHtml64(col2Text) {
      return String(col2Text || '').split('\\n').map(line => {
        if (line.startsWith('💡')) {
          return '<div class="note-line">' + line + '</div>';
        }
        const m = line.match(/^(【案[123]】)(.*)$/);
        if (m) return '<div class="plan-line"><b>' + m[1] + '</b>' + m[2] + '</div>';
        return '<div class="plan-line">' + line + '</div>';
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
            '<div style="font-size:11px;font-weight:bold;color:#546e7a;margin-bottom:3px;">' + row.dateLabel + '</div>' +
            '<div><span class="region-tag' + (row.macroArea === '関西' ? ' kansai' : '') + '">' + row.clinicGroup + '</span>' +
            '<span class="clinic-name">' + row.clinicTitle + '</span></div>' +
            '<div class="time-wage">' + row.timeAndWage + '</div>' +
          '</td>' +
          '<td>' + formatPlansHtml64(row.col2Plans) + '</td>' +
          '<td>' +
            '<div class="diff-box">' +
              '<div class="diff-before">【編集前】' + row.origDecision + ' ／ 担当:' + (row.origAssignee || '未定') + '</div>' +
              '<div class="diff-after' + (isEdited ? ' changed' : '') + '" id="after_' + row.slotKey + '">' +
                '【編集後】' + row.curDecision + ' ／ 担当:' + (row.curAssignee || '未定') +
              '</div>' +
            '</div>' +
          '</td>' +
          '<td>' +
            '<select class="n64-select" style="width:100%;" id="sel_' + row.slotKey + '" onchange="onSlotEdit64(\\'' + row.slotKey + '\\')">' + decOpts + '</select>' +
            '<input type="text" class="n64-input" id="cus_' + row.slotKey + '" style="width:100%;margin-top:4px;display:' + (selVal === 'カスタム入力' ? 'block' : 'none') + ';" ' +
              'placeholder="例: 高橋先生へ延長" value="' + customVal.replace(/"/g, '&quot;') + '" oninput="onSlotEdit64(\\'' + row.slotKey + '\\')">' +
          '</td>' +
          '<td>' +
            '<select class="n64-select" style="width:100%;" id="stf_' + row.slotKey + '" onchange="onSlotEdit64(\\'' + row.slotKey + '\\')">' + staffOpts + '</select>' +
          '</td>' +
        '</tr>';
      }).join('');

      document.getElementById('stageTableBody').innerHTML = html;
    }

    function onSlotEdit64(slotKey) {
      const row = state.slotRows.find(r => r.slotKey === slotKey);
      if (!row) return;

      const sel = document.getElementById('sel_' + slotKey).value;
      const cusInput = document.getElementById('cus_' + slotKey);
      cusInput.style.display = (sel === 'カスタム入力') ? 'block' : 'none';

      const finalDec = (sel === 'カスタム入力') ? (cusInput.value.trim() || '要検討') : sel;
      const finalStaff = document.getElementById('stf_' + slotKey).value;

      row.curDecision = finalDec;
      row.curAssignee = finalStaff;
      state.changedSlots[slotKey] = { decision: finalDec, assignee: finalStaff };

      const isEdited = (row.curDecision !== row.origDecision) || (row.curAssignee !== row.origAssignee);
      const tr = document.getElementById('tr_' + slotKey);
      if (tr) {
        tr.classList.toggle('is-edited', isEdited);
        tr.classList.toggle('is-stay', finalDec === 'ステイ');
      }
      const afterEl = document.getElementById('after_' + slotKey);
      if (afterEl) {
        afterEl.className = 'diff-after' + (isEdited ? ' changed' : '');
        afterEl.innerText = '【編集後】' + row.curDecision + ' ／ 担当:' + (row.curAssignee || '未定');
      }
      updateEditCounter();
    }

    function applyBulkStaff64(overwriteAll) {
      const staff = document.getElementById('bulkStaffSelect').value;
      if (!staff) {
        alert('一括セットする担当者名を選択してください！');
        return;
      }
      const lockedByOthers = getLockedStaffSetByOthers();
      if (lockedByOthers.has(staff)) {
        alert('🔒 「' + staff + '」さんは他の方が現在作業中のため一括セットできません。');
        return;
      }

      const scope = document.getElementById('bulkTargetScope').value;
      let count = 0;

      state.slotRows.forEach(row => {
        if (!isRowMatchRegionFilter(row)) return;
        if (scope === 'VISIBLE_DATE' && row.dateLabel !== state.activeDateTab) return;
        if (!overwriteAll && row.curAssignee) return;

        row.curAssignee = staff;
        state.changedSlots[row.slotKey] = { decision: row.curDecision, assignee: row.curAssignee };
        count++;
      });

      renderStageTable();
      updateEditCounter();
      setStatus('⚡ ' + count + ' 枠に担当者「' + staff + '」を一括セットしました！', '#ffc72c');
    }

    function renderBundlesTable() {
      const html = state.bundleRows.map(b => {
        const isEdited = (b.curVal !== b.origVal);
        return '<tr class="slot-row' + (isEdited ? ' is-edited' : '') + '">' +
          '<td style="font-weight:900;color:#0d47a1;">' + b.doctor + '</td>' +
          '<td style="white-space:pre-wrap;font-size:11.5px;">' + b.slots + '</td>' +
          '<td style="font-size:11.5px;">' + b.note + '</td>' +
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

    function onBundleEdit64(bKey, val) {
      const item = state.bundleRows.find(b => b.bundleKey === bKey);
      if (!item) return;
      item.curVal = val.trim();
      state.changedBundles[bKey] = item.curVal;
      const el = document.getElementById('bafter_' + bKey);
      if (el) {
        const isEd = (item.curVal !== item.origVal);
        el.className = 'diff-after' + (isEd ? ' changed' : '');
        el.innerText = '【編集後】' + item.curVal;
      }
      updateEditCounter();
    }

    function updateEditCounter() {
      const total = Object.keys(state.changedSlots).length + Object.keys(state.changedBundles).length;
      document.getElementById('editCounterBadge').innerText = '未保存の変更: ' + total + '件';
    }

    function setStatus(msg, color) {
      const el = document.getElementById('statusBanner');
      el.style.color = color || '#ffc72c';
      el.innerText = msg;
    }

    function saveDraftOnly64() {
      setStatus('💾 一時保存（ドラフト共有）しています...', '#ffc72c');
      google.script.run
        .withSuccessHandler(res => {
          setStatus('✅ 一時保存完了 (' + res.savedAt + ')！同時作業メンバーにも共有されました。', '#69f0ae');
        })
        .withFailureHandler(err => setStatus('❌ 一時保存エラー: ' + err.message, '#ff5252'))
        .saveHenshukunDraft({
          changedSlots: state.changedSlots,
          changedBundles: state.changedBundles
        });
    }

    function commitToDoc64(andSubmit) {
      setStatus(andSubmit ? '🚀 ドキュメントを更新して mnys へ提出中...' : '📄 Googleドキュメント本体へ反映中...', '#ffc72c');
      google.script.run
        .withSuccessHandler(res => {
          state.changedSlots = {};
          state.changedBundles = {};
          state.slotRows.forEach(r => { r.origDecision = r.curDecision; r.origAssignee = r.curAssignee; });
          state.bundleRows.forEach(b => { b.origVal = b.curVal; });
          renderStageTable();
          renderBundlesTable();
          updateEditCounter();
          setStatus(
            andSubmit
              ? '🎉 ドキュメント反映 ＆ mnys（超募集くん）への提出が完了しました！(' + res.committedAt + ')'
              : '✅ Googleドキュメント本体への反映が完了しました！(' + res.committedAt + ')',
            '#69f0ae'
          );
        })
        .withFailureHandler(err => setStatus('❌ 反映エラー: ' + err.message, '#ff5252'))
        .commitHenshukunToDocAndSubmit({
          changedSlots: state.changedSlots,
          changedBundles: state.changedBundles
        }, andSubmit);
    }
  </script>`;
}