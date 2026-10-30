/**
 * ============================================================================
 * 7b_編集くん64_レイアウト.gs
 * （『🎮 編集くん 64』のHTMLボディ構造：ヘッダー・フィルタ・テーブル・フッター）
 * ============================================================================
 */

function getHenshukun64BodyLayout_() {
  return `
  <!-- 1. 64風トップヘッダー -->
  <div class="n64-header">
    <div class="logo-box">
      <div class="n64-badge">🎮 編集くん 64</div>
      <div>
        <div style="font-weight:900;font-size:14px;">CAPS 超募集対策スタジオ（同時編集＆エリア一括対応）</div>
        <div id="docLinkBox" style="margin-top:2px;">データ読込中...</div>
      </div>
    </div>

    <!-- 自分のプレイヤー選択（選ぶと他の人の画面でその担当者名がブロックされる） -->
    <div class="player-box">
      <span style="font-size:12px;font-weight:bold;">🕹️ 自分の担当者名:</span>
      <select id="myPlayerSelect" class="n64-select" onchange="onMyPlayerChange()">
        <option value="">-- 選択するとロック保護 --</option>
      </select>
      <div id="activePlayersText" class="active-players">オンライン: 確認中</div>
      <button class="btn-64" onclick="reloadAllData()">🔄 最新同期</button>
    </div>
  </div>

  <!-- 2. コントロールデッキ（関東/関西モード・地域複数選択・エリア一括担当入力） -->
  <div class="control-deck">
    <div class="mode-group">
      <span style="font-size:11.5px;font-weight:bold;color:#ffeb3b;">🌐 エリアモード:</span>
      <button id="mode_ALL" class="btn-64 active-mode" onclick="setMacroMode('ALL')">全エリア</button>
      <button id="mode_関東" class="btn-64" onclick="setMacroMode('関東')">🗼 関東モード</button>
      <button id="mode_関西" class="btn-64" onclick="setMacroMode('関西')">🐙 関西モード</button>
    </div>

    <div class="mode-group">
      <span style="font-size:11.5px;font-weight:bold;color:#80d8ff;">📍 地域フィルタ(複数可):</span>
      <div id="regionChipsBox" class="region-chips"></div>
      <button class="btn-64" style="padding:3px 8px;font-size:11px;" onclick="clearRegionSelection()">全解除</button>
    </div>

    <div class="mode-group" style="background:rgba(0,0,0,0.35);padding:4px 10px;border-radius:6px;">
      <span style="font-size:11.5px;font-weight:bold;color:#a5d6a7;">⚡ 担当者一括セット:</span>
      <select id="bulkTargetScope" class="n64-select" style="font-size:11.5px;padding:3px 6px;">
        <option value="VISIBLE_DATE">表示中の日付×選択地域のみ</option>
        <option value="ALL_DATES_REGION">全日程の選択地域すべて</option>
      </select>
      <select id="bulkStaffSelect" class="n64-select" style="font-size:11.5px;padding:3px 6px;"></select>
      <button class="btn-64 btn-blue" style="padding:4px 9px;" onclick="applyBulkStaff64(false)">空欄に一括</button>
      <button class="btn-64 btn-yellow" style="padding:4px 9px;" onclick="applyBulkStaff64(true)">上書き一括</button>
    </div>
  </div>

  <!-- 3. ステージセレクト（日付ごとのタブ＋おまとめ打診タブ） -->
  <div id="dateTabsBar" class="date-tabs-bar"></div>

  <!-- 4. メインステージ（スクロール最小化テーブル） -->
  <div class="main-stage">
    <div id="slotsStageView">
      <table class="stage-table">
        <thead>
          <tr>
            <th style="width:210px;">日付 ／ 地域・拠点・時間帯</th>
            <th>AI推奨アプローチ（優先順）</th>
            <th style="width:220px;">編集前 ➔ 編集後イメージ</th>
            <th style="width:185px;">③ 確定案の選択・入力</th>
            <th style="width:135px;">④ 担当者名</th>
          </tr>
        </thead>
        <tbody id="stageTableBody"></tbody>
      </table>
    </div>

    <!-- おまとめ打診ステージ -->
    <div id="bundlesStageView" style="display:none;">
      <table class="stage-table">
        <thead>
          <tr>
            <th style="width:180px;">対象医師</th>
            <th>まとめて打診する対象拠点・日程</th>
            <th style="width:220px;">実績・交渉時の注意点</th>
            <th style="width:190px;">編集前 ➔ 編集後</th>
            <th style="width:200px;">確定（採用/ステイ）・担当者</th>
          </tr>
        </thead>
        <tbody id="bundlesTableBody"></tbody>
      </table>
    </div>
  </div>

  <!-- 5. 下部フッター（一時保存・ドキュメント反映・mnys提出） -->
  <div class="n64-footer">
    <div class="status-monitor">
      <span id="statusBanner">🕹️ READY! 日付タブを選んで編集してください</span>
      <span id="editCounterBadge" style="background:#37474f;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px;">変更: 0件</span>
    </div>
    <div style="display:flex;gap:10px;">
      <button class="btn-64 btn-yellow" onclick="saveDraftOnly64()">💾 一時保存（ドラフト共有）</button>
      <button class="btn-64 btn-green" onclick="commitToDoc64(false)">📄 ドキュメントへ反映（上書き保存）</button>
      <button class="btn-64 btn-red" style="font-size:13px;padding:7px 16px;" onclick="commitToDoc64(true)">🚀 ドキュメント反映 ＆ mnys（超募集くん）へ提出</button>
    </div>
  </div>`;
}