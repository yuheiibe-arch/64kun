/**
 * ============================================================================
 * 7a_編集くん64_スタイル.gs
 * （『🎮 編集くん 64』のCSSスタイル定義 ＆ HTML全体組み立て関数）
 * ============================================================================
 */

function buildHenshukun64Html_() {
  return [
    '<!DOCTYPE html><html><head><base target="_top">',
    getHenshukun64Style_(),
    '</head><body>',
    getHenshukun64BodyLayout_(),
    getHenshukun64ClientScript_(),
    '</body></html>'
  ].join('\n');
}

function getHenshukun64Style_() {
  return `<style>
    :root {
      --n64-blue: #0046ad;
      --n64-red: #e4002b;
      --n64-yellow: #ffc72c;
      --n64-green: #009b4d;
      --n64-dark: #1e222d;
      --n64-panel: #2b3040;
    }
    * { box-sizing: border-box; }
    body {
      font-family: 'Rounded Mplus 1c', 'Hiragino Maru Gothic ProN', 'Meiryo', sans-serif;
      margin: 0; padding: 0;
      background: var(--n64-dark);
      color: #fff;
      height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      user-select: none;
    }

    /* 上部ヘッダー（64カートリッジ風） */
    .n64-header {
      background: linear-gradient(90deg, #003380 0%, #1a2238 60%, #004d26 100%);
      border-bottom: 4px solid var(--n64-yellow);
      padding: 8px 14px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      box-shadow: 0 4px 0 #000;
    }
    .logo-box { display: flex; align-items: center; gap: 10px; }
    .n64-badge {
      background: var(--n64-red);
      color: var(--n64-yellow);
      font-weight: 900;
      font-size: 18px;
      padding: 4px 10px;
      border-radius: 6px;
      border: 3px solid #fff;
      box-shadow: 3px 3px 0 #000;
      letter-spacing: 1px;
    }
    .doc-link { color: #80d8ff; font-size: 12px; font-weight: bold; text-decoration: underline; }
    .player-box {
      display: flex; align-items: center; gap: 10px;
      background: rgba(0,0,0,0.45);
      padding: 5px 12px;
      border-radius: 8px;
      border: 2px solid #5c6bc0;
    }
    .active-players { font-size: 11px; color: var(--n64-yellow); font-weight: bold; }

    /* コントロールバー（エリアモード・地域複数選択・一括担当） */
    .control-deck {
      background: var(--n64-panel);
      padding: 8px 14px;
      border-bottom: 3px solid #111;
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      justify-content: space-between;
    }
    .mode-group { display: flex; align-items: center; gap: 6px; }
    .btn-64 {
      cursor: pointer;
      font-weight: 900;
      font-size: 12px;
      padding: 6px 12px;
      border-radius: 6px;
      border: 2px solid #fff;
      box-shadow: 2px 3px 0 #000;
      transition: transform 0.05s;
      color: #fff;
      background: #455a64;
    }
    .btn-64:active { transform: translateY(2px); box-shadow: 0 1px 0 #000; }
    .btn-64.active-mode { background: var(--n64-yellow); color: #000; border-color: #fff; }
    .btn-blue { background: var(--n64-blue); }
    .btn-green { background: var(--n64-green); }
    .btn-red { background: var(--n64-red); }
    .btn-yellow { background: var(--n64-yellow); color: #111; }

    /* 地域チップ（複数選択可） */
    .region-chips { display: flex; gap: 5px; flex-wrap: wrap; align-items: center; }
    .chip {
      cursor: pointer;
      font-size: 11.5px;
      font-weight: bold;
      padding: 4px 10px;
      border-radius: 16px;
      border: 2px solid #78909c;
      background: #1c202b;
      color: #cfd8dc;
    }
    .chip.selected {
      background: #00bcd4;
      color: #000;
      border-color: #fff;
      box-shadow: 0 0 6px rgba(0, 188, 212, 0.7);
    }

    /* ステージセレクト（日付タブ） */
    .date-tabs-bar {
      background: #141720;
      padding: 6px 14px 0;
      display: flex;
      gap: 6px;
      overflow-x: auto;
      border-bottom: 3px solid var(--n64-yellow);
    }
    .date-tab {
      cursor: pointer;
      padding: 6px 14px;
      background: #2c3347;
      color: #b0bec5;
      font-weight: 900;
      font-size: 12.5px;
      border-radius: 8px 8px 0 0;
      border: 2px solid #455a64;
      border-bottom: none;
      white-space: nowrap;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .date-tab.active { background: var(--n64-yellow); color: #000; border-color: #fff; }
    .tab-count { background: rgba(0,0,0,0.25); padding: 1px 6px; border-radius: 10px; font-size: 11px; }

    /* メインテーブル領域 */
    .main-stage {
      flex: 1;
      overflow-y: auto;
      padding: 10px 14px;
      background: #e8ecf4;
      color: #111;
    }
    table.stage-table { width: 100%; border-collapse: separate; border-spacing: 0 6px; }
    table.stage-table th {
      background: var(--n64-blue);
      color: #fff;
      padding: 7px 10px;
      font-size: 12px;
      text-align: left;
      position: sticky;
      top: 0;
      z-index: 5;
      border-top: 2px solid #fff;
      border-bottom: 2px solid #000;
    }
    tr.slot-row { background: #fff; box-shadow: 0 2px 0 #b0bec5; transition: background 0.15s; }
    tr.slot-row.is-edited { background: #fffde7; outline: 2px solid #fbc02d; }
    tr.slot-row.is-stay { opacity: 0.65; background: #eceff1; }
    tr.slot-row td {
      padding: 7px 10px;
      vertical-align: middle;
      border-top: 1px solid #cfd8dc;
      border-bottom: 1px solid #cfd8dc;
      font-size: 12px;
    }
    .region-tag {
      display: inline-block;
      padding: 2px 7px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 900;
      color: #fff;
      background: #3949ab;
      margin-right: 5px;
    }
    .region-tag.kansai { background: #d81b60; }
    .clinic-name { font-weight: 900; font-size: 13px; color: #0d47a1; }
    .time-wage { font-size: 11.5px; color: #37474f; font-weight: bold; margin-top: 2px; }

    .plan-line { font-size: 11.5px; line-height: 1.4; margin-bottom: 2px; color: #222; }
    .plan-line b { color: var(--n64-blue); }
    .note-line { font-size: 11px; color: #b06000; background: #fff3e0; padding: 2px 6px; border-radius: 4px; display: inline-block; margin-top: 2px; }

    /* Before -> After 比較ボックス */
    .diff-box {
      display: flex;
      flex-direction: column;
      gap: 3px;
      font-size: 11px;
      background: #f1f5f9;
      padding: 5px 8px;
      border-radius: 6px;
      border: 1px solid #cbd5e1;
    }
    .diff-before { color: #64748b; }
    .diff-after { font-weight: 900; color: #0f172a; }
    .diff-after.changed { color: #d93025; background: #fef08a; padding: 1px 5px; border-radius: 4px; }

    select.n64-select, input.n64-input {
      font-size: 12.5px;
      font-weight: bold;
      padding: 5px 8px;
      border: 2px solid #1e293b;
      border-radius: 6px;
      background: #fff;
      cursor: pointer;
    }
    select.n64-select:focus, input.n64-input:focus {
      outline: none;
      border-color: var(--n64-blue);
      background: #e3f2fd;
    }

    /* 下部フッター */
    .n64-footer {
      background: #141720;
      border-top: 4px solid var(--n64-green);
      padding: 8px 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .status-monitor {
      font-size: 12.5px;
      font-weight: bold;
      color: var(--n64-yellow);
      display: flex;
      align-items: center;
      gap: 10px;
    }
  </style>`;
}