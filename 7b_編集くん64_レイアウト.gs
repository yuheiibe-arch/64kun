/**
 * ============================================================================
 * 7b_編集くん64_レイアウト.gs
 * （「自分の担当者名:」ラベル削除 ＆ 決定済み/全体数リアルタイム進捗ボード追加版）
 * ============================================================================
 */

function getHenshukun64BodyLayout_() {
  return `
  <!-- 1. 64風トップヘッダー -->
  <div class="n64-header">
    <div class="logo-box">
      <div class="n64-badge">🎮 編集くん 64</div>
      <div>
        <div style="font-weight:900;font-size:15px;">CAPS 超募集対策スタジオ（同時編集＆エリア一括対応）</div>
        <div id="docLinkBox" style="margin-top:2px;">データ読込中...</div>
      </div>
    </div>

    <!-- ラベル文字をなくしスッキリ配置 -->
    <div class="player-box">
      <select id="myPlayerSelect" class="n64-select" onchange="onMyPlayerChange()">
        <option value="">-- 作業者を選択 --</option>
      </select>
      <div id="activePlayersText" class="active-players">🟢 重複作業なし</div>
      <button class="btn-64" onclick="reloadAllData()">🔄 最新同期</button>
    </div>
  </div>

  <!-- 2. コントロールデッキ -->
  <div class="control-deck">
    <!-- 上段：エリアモード ＆ 地域複数選択 -->
    <div class="deck-row">
      <div class="mode-group">
        <span class="deck-label" style="color:#ffeb3b;">🌐 エリアモード:</span>
        <button id="mode_ALL" class="btn-64 active-mode" onclick="setMacroMode('ALL')">全エリア</button>
        <button id="mode_関東" class="btn-64" onclick="setMacroMode('関東')">🗼 関東モード</button>
        <button id="mode_関西" class="btn-64" onclick="setMacroMode('関西')">🐙 関西モード</button>
      </div>

      <div class="mode-group">
        <span class="deck-label" style="color:#84ffff;">📍 地域フィルタ(複数可):</span>
        <div id="regionChipsBox" class="region-chips"></div>
        <button class="btn-64" style="padding:4px 10px;font-size:12px;" onclick="clearRegionSelection()">全解除</button>
      </div>
    </div>

    <!-- 下段：担当者一括セット ＆ 決定済み/全体数（残り対応数）メーター -->
    <div class="deck-row" style="background:rgba(0,0,0,0.35);padding:6px 12px;border-radius:7px;border:1px solid #455a64;">
      <div class="mode-group">
        <span class="deck-label" style="color:#69f0ae;">⚡ 担当者一括セット:</span>
        <select id="bulkTargetScope" class="n64-select">
          <option value="VISIBLE_DATE">表示中の「日付 × 選択地域」のみ</option>
          <option value="ALL_DATES_REGION">全日程の「選択地域」すべて</option>
        </select>
        <select id="bulkStaffSelect" class="n64-select" style="min-width:130px;"></select>
        <button class="btn-64 btn-blue" onclick="applyBulkStaff64(false)">空欄に一括</button>
        <button class="btn-64 btn-yellow" onclick="applyBulkStaff64(true)">上書き一括</button>
      </div>

      <!-- ★新設：決定済み / 全体数（残り対応数）のリアルタイム集計ボード -->
      <div class="progress-board">
        <div class="prog-main" id="progressMainText">📊 決定済み: <span>0</span> / 0 枠</div>
        <div class="prog-remain" id="progressRemainBadge">残り対応(担当未定): 0枠</div>
      </div>
    </div>
  </div>

  <!-- 3. ステージセレクト（日付ごとのタブ＋おまとめ打診タブ） -->
  <div id="dateTabsBar" class="date-tabs-bar"></div>

  <!-- 4. メインステージ -->
  <div class="main-stage">
    <div id="slotsStageView">
      <table class="stage-table">
        <thead>
          <tr>
            <th style="width:215px;">① 日付 ／ 地域・拠点・時間帯</th>
            <th>② AI推奨アプローチ（優先順・なぞってコピー可）</th>
            <th style="width:240px;">編集前 ➔ 編集後イメージ</th>
            <th style="width:190px;">③ 確定案の選択・入力</th>
            <th style="width:140px;">④ 担当者名</th>
          </tr>
        </thead>
        <tbody id="stageTableBody"></tbody>
      </table>
    </div>

    <div id="bundlesStageView" style="display:none;">
      <table class="stage-table">
        <thead>
          <tr>
            <th style="width:190px;">対象医師</th>
            <th>まとめて打診する対象拠点・日程</th>
            <th style="width:230px;">実績・交渉時の注意点</th>
            <th style="width:200px;">編集前 ➔ 編集後</th>
            <th style="width:210px;">確定（採用/ステイ）・担当者</th>
          </tr>
        </thead>
        <tbody id="bundlesTableBody"></tbody>
      </table>
    </div>
  </div>

  <!-- 5. 下部フッター -->
  <div class="n64-footer">
    <div class="status-monitor">
      <span id="statusBanner" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">🕹️ READY! 日付タブを選んで編集してください</span>
      <span id="editCounterBadge" class="badge-counter">未保存の変更: 0件</span>
    </div>

    <div class="footer-actions">
      <button class="btn-action-lg btn-yellow" onclick="saveDraftOnly64()">
        <span class="main-txt">💾 一時保存（シート＆JSON保存）</span>
        <span class="sub-txt">同時編集メンバーへドラフト共有</span>
      </button>
      <button class="btn-action-lg btn-green" style="color:#fff;" onclick="commitToDoc64(false)">
        <span class="main-txt">📄 ドキュメントへ反映</span>
        <span class="sub-txt">Googleドキュメント上書き保存</span>
      </button>
      <button class="btn-action-lg btn-red" style="color:#fff;" onclick="commitToDoc64(true)">
        <span class="main-txt">🚀 ドキュメント反映 ＆ mnysへ提出</span>
        <span class="sub-txt">超募集くんで日給計算・文案生成</span>
      </button>
    </div>
  </div>`;
}