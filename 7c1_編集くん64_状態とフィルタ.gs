/**
 * ============================================================================
 * 7c1_編集くん64_状態とフィルタ.gs
 * （状態管理・初期ロード・ハートビート同期・関東/関西モード＆地域複数選択・日付タブ）
 * ============================================================================
 */

function getHenshukun64ScriptPart1_() {
  return `
    let state = {
      sessionId: '',
      myStaff: '',
      macroMode: 'ALL',
      selectedRegions: new Set(),
      activeDateTab: '',
      staffList: [],
      doctorSearchList: [],
      groups: [],
      areaToGroups: {},
      slotRows: [],
      bundleRows: [],
      activeLocks: {},
      changedSlots: {},
      changedBundles: {}
    };

    const DECISION_PRESETS = ['案1', '案2', '案3', '案1, 案2', '案1, 案3', '案2, 案3', '全案', 'ステイ', 'カスタム入力'];
    const CUSTOM_ACTION_TYPES = ['勤務打診', '勤務依頼', '移動打診', '延長打診', '早出打診'];

    window.onload = function() {
      reloadAllData();
      setInterval(sendHeartbeat64, 15000);
      document.addEventListener('click', function(e) {
        if (!e.target.closest('.custom-builder')) {
          document.querySelectorAll('.suggest-dropdown').forEach(el => el.style.display = 'none');
        }
      });
    };

    function normalizeSearchClient64(str) {
      return String(str || '')
        .replace(/[\\s\\u3000]+/g, '')
        .replace(/[﨑嵜]/g, '崎').replace(/髙/g, '高').replace(/[瀧泷]/g, '滝')
        .replace(/[邊邉]/g, '辺').replace(/[齋齊斉]/g, '斎').replace(/澤/g, '沢')
        .replace(/廣/g, '広').replace(/國/g, '国').replace(/嶋/g, '島')
        .replace(/[栁柳]/g, '柳').replace(/[濵濱]/g, '浜').replace(/櫻/g, '桜').replace(/眞/g, '真')
        .trim();
    }

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
      state.doctorSearchList = data.doctorSearchList || [];
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
      setStatus('✅ 読み込み完了！（医師マスタ ' + state.doctorSearchList.length + ' 名検索可能）', '#69f0ae');
    }

    function isSlotDecided64(row) {
      const dec = String(row.curDecision || '').trim();
      const stf = String(row.curAssignee || '').trim();
      if (!dec || dec === '要検討') return false;
      if (dec === 'ステイ') return true;
      return Boolean(stf);
    }

    function getLockedStaffSetByOthers() {
      const locked = new Set();
      Object.entries(state.activeLocks || {}).forEach(([sid, info]) => {
        if (sid !== state.sessionId && info.staff) locked.add(info.staff);
      });
      return locked;
    }

    function renderStaffSelectors() {
      const lockedByOthers = getLockedStaffSetByOthers();
      const buildOptions = (selectedVal, isMySelector) => {
        const head = isMySelector ? '<option value="">-- 作業者を選択 --</option>' : '<option value="">(担当者未定)</option>';
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
          let syncedChanged = false;
          if (res.latestDraft && res.latestDraft.slots) {
            state.slotRows.forEach(r => {
              const d = res.latestDraft.slots[r.slotKey];
              if (d && !state.changedSlots[r.slotKey]) {
                if (r.curDecision !== d.decision || r.curAssignee !== d.assignee) {
                  r.curDecision = d.decision;
                  r.curAssignee = d.assignee;
                  r.customDocs = Array.isArray(d.customDocs) ? d.customDocs : [];
                  syncedChanged = true;
                }
              }
            });
          }
          renderStaffSelectors();
          updateActivePlayersDisplay();
          if (syncedChanged) {
            renderDateTabs();
            renderStageTable();
            updateEditCounter();
          }
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
      updateEditCounter();
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
      if (state.selectedRegions.has(groupName)) state.selectedRegions.delete(groupName);
      else state.selectedRegions.add(groupName);
      renderRegionChips();
      const dates = getFilteredUniqueDates();
      if (state.activeDateTab !== 'BUNDLES' && !dates.includes(state.activeDateTab)) {
        state.activeDateTab = dates[0] || '';
      }
      renderDateTabs();
      renderStageTable();
      updateEditCounter();
    }

    function clearRegionSelection() {
      state.selectedRegions.clear();
      renderRegionChips();
      renderDateTabs();
      renderStageTable();
      updateEditCounter();
    }

    function isRowMatchRegionFilter(row) {
      if (state.macroMode !== 'ALL' && row.macroArea !== state.macroMode) return false;
      if (state.selectedRegions.size > 0 && !state.selectedRegions.has(row.clinicGroup)) return false;
      return true;
    }

    function getFilteredUniqueDates() {
      const set = new Set();
      state.slotRows.forEach(r => { if (isRowMatchRegionFilter(r)) set.add(r.dateLabel); });
      return Array.from(set);
    }

    function renderDateTabs() {
      const dates = getFilteredUniqueDates();
      const tabsHtml = dates.map(d => {
        const dayRows = state.slotRows.filter(r => r.dateLabel === d && isRowMatchRegionFilter(r));
        const decided = dayRows.filter(isSlotDecided64).length;
        const isAct = (state.activeDateTab === d);
        return '<div class="date-tab' + (isAct ? ' active' : '') + '" onclick="selectStageTab(\\'' + d + '\\')">' +
          '<span>📅 ' + d + '</span><span class="tab-count">' + decided + '/' + dayRows.length + '枠</span></div>';
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
  `;
}