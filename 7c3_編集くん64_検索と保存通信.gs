/**
 * ============================================================================
 * 7c3_編集くん64_検索と保存通信.gs
 * （医師サジェスト検索・複数医師追加・一括担当セット・進捗集計・一時保存/提出通信）
 * ============================================================================
 */

function getHenshukun64ScriptPart3_() {
  return `
    function onSearchDocInput64(slotKey) {
      const inp = document.getElementById('srch_' + slotKey);
      const drop = document.getElementById('sug_' + slotKey);
      if (!inp || !drop) return;

      const q = normalizeSearchClient64(inp.value);
      if (!q) {
        drop.style.display = 'none';
        return;
      }

      const matches = state.doctorSearchList.filter(d => d.norm.includes(q) || d.name.includes(inp.value.trim())).slice(0, 12);
      if (matches.length === 0) {
        drop.innerHTML = '<div style="padding:6px 10px;font-size:12px;color:#64748b;">該当する医師が見つかりません（下の欄に直接手入力も可能です）</div>';
        drop.style.display = 'block';
        return;
      }

      drop.innerHTML = matches.map(d => {
        const licBadge = d.licenseNo ? '医籍: ' + d.licenseNo : '番号未登録';
        const tagBadge = d.tag ? ' [' + d.tag + ']' : '';
        return '<div class="suggest-item" onclick="pickDoctorFromSuggest64(\\'' + slotKey + '\\', \\'' + d.name + '\\', \\'' + (d.licenseNo || '') + '\\')">' +
          '<span>👨‍⚕️ ' + d.name + ' 先生' + tagBadge + '</span>' +
          '<span class="suggest-lic">' + licBadge + '</span>' +
        '</div>';
      }).join('');
      drop.style.display = 'block';
    }

    function pickDoctorFromSuggest64(slotKey, docName, licenseNo) {
      const row = state.slotRows.find(r => r.slotKey === slotKey);
      if (!row) return;

      const actSelect = document.getElementById('act_' + slotKey);
      const actionType = actSelect ? actSelect.value : '勤務打診';
      const cusInput = document.getElementById('cus_' + slotKey);
      const srchInput = document.getElementById('srch_' + slotKey);
      const drop = document.getElementById('sug_' + slotKey);

      if (!Array.isArray(row.customDocs)) row.customDocs = [];
      const existingIdx = row.customDocs.findIndex(d => d.name === docName);
      if (existingIdx >= 0) {
        row.customDocs[existingIdx].actionType = actionType;
        if (licenseNo) row.customDocs[existingIdx].licenseNo = licenseNo;
      } else {
        row.customDocs.push({ name: docName, licenseNo: licenseNo || '', actionType: actionType });
      }

      const phraseList = row.customDocs.map(d => d.name + '先生へ' + (d.actionType || '勤務打診'));
      cusInput.value = phraseList.join(' / ');

      if (srchInput) srchInput.value = '';
      if (drop) drop.style.display = 'none';

      const tagsBox = document.getElementById('tags_' + slotKey);
      if (tagsBox) tagsBox.innerHTML = buildPickedDocsTagsHtml64(slotKey, row.customDocs);

      onSlotEdit64(slotKey);
    }

    function removeCustomDoc64(slotKey, idx) {
      const row = state.slotRows.find(r => r.slotKey === slotKey);
      if (!row || !Array.isArray(row.customDocs)) return;

      row.customDocs.splice(idx, 1);
      const cusInput = document.getElementById('cus_' + slotKey);
      if (cusInput) {
        cusInput.value = row.customDocs.map(d => d.name + '先生へ' + (d.actionType || '勤務打診')).join(' / ');
      }
      const tagsBox = document.getElementById('tags_' + slotKey);
      if (tagsBox) tagsBox.innerHTML = buildPickedDocsTagsHtml64(slotKey, row.customDocs);

      onSlotEdit64(slotKey);
    }

    function onSlotEdit64(slotKey) {
      const row = state.slotRows.find(r => r.slotKey === slotKey);
      if (!row) return;

      const sel = document.getElementById('sel_' + slotKey).value;
      const cbox = document.getElementById('cbox_' + slotKey);
      const cusInput = document.getElementById('cus_' + slotKey);
      if (cbox) cbox.style.display = (sel === 'カスタム入力') ? 'block' : 'none';

      if (sel !== 'カスタム入力') {
        row.customDocs = [];
      }

      const finalDec = (sel === 'カスタム入力') ? (cusInput.value.trim() || '要検討') : sel;
      const finalStaff = document.getElementById('stf_' + slotKey).value;

      row.curDecision = finalDec;
      row.curAssignee = finalStaff;
      state.changedSlots[slotKey] = {
        decision: finalDec,
        assignee: finalStaff,
        customDocs: Array.isArray(row.customDocs) ? row.customDocs : [],
        dateLabel: row.dateLabel,
        clinicTitle: row.clinicTitle
      };

      const isEdited = (row.curDecision !== row.origDecision) || (row.curAssignee !== row.origAssignee);
      const tr = document.getElementById('tr_' + slotKey);
      if (tr) {
        tr.classList.toggle('is-edited', isEdited);
        tr.classList.toggle('is-stay', finalDec === 'ステイ');
      }
      const afterEl = document.getElementById('after_' + slotKey);
      if (afterEl) {
        afterEl.className = 'diff-after' + (isEdited ? ' changed' : '');
        afterEl.innerHTML = buildAfterDiffHtml64(row);
      }
      renderDateTabs();
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
        state.changedSlots[row.slotKey] = {
          decision: row.curDecision,
          assignee: row.curAssignee,
          customDocs: Array.isArray(row.customDocs) ? row.customDocs : [],
          dateLabel: row.dateLabel,
          clinicTitle: row.clinicTitle
        };
        count++;
      });

      renderDateTabs();
      renderStageTable();
      updateEditCounter();
      setStatus('⚡ ' + count + ' 枠に担当者「' + staff + '」を一括セットしました！', '#ffc72c');
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
      const totalChanges = Object.keys(state.changedSlots).length + Object.keys(state.changedBundles).length;
      document.getElementById('editCounterBadge').innerText = '未保存の変更: ' + totalChanges + '件';

      const targetRows = state.slotRows.filter(isRowMatchRegionFilter);
      const totalSlots = targetRows.length;
      const decidedSlots = targetRows.filter(isSlotDecided64).length;
      const remainSlots = totalSlots - decidedSlots;

      document.getElementById('progressMainText').innerHTML =
        '📊 決定済み: <span>' + decidedSlots + '</span> / ' + totalSlots + ' 枠';

      const remBadge = document.getElementById('progressRemainBadge');
      if (remainSlots === 0 && totalSlots > 0) {
        remBadge.className = 'prog-remain all-done';
        remBadge.innerText = '🎉 すべて決定済み！';
      } else {
        remBadge.className = 'prog-remain';
        remBadge.innerText = '残り対応(未定): ' + remainSlots + '枠';
      }
    }

    function setStatus(msg, color) {
      const el = document.getElementById('statusBanner');
      el.style.color = color || '#ffc72c';
      el.innerText = msg;
    }

    function saveDraftOnly64() {
      setStatus('💾 『編集くん_一時保存』シートにJSON＆一覧で一時保存しています...', '#ffc72c');
      google.script.run
        .withSuccessHandler(res => {
          setStatus('✅ 一時保存完了 (' + res.savedAt + ')！特定した医籍番号もシートにJSON保存されました。', '#69f0ae');
        })
        .withFailureHandler(err => setStatus('❌ 一時保存エラー: ' + err.message, '#ff5252'))
        .saveHenshukunDraft({
          playerStaff: state.myStaff,
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
          renderDateTabs();
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
          playerStaff: state.myStaff,
          changedSlots: state.changedSlots,
          changedBundles: state.changedBundles
        }, andSubmit);
    }
  `;
}