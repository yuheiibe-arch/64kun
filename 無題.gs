/**
 * 【Step 1：読み取り専用】どこに以前の入力プランが残っているかを全箇所スキャンしてログに出す
 */
function Step1_inspectSavedPlans() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties().getProperties();
  const currentDocId = props['SINGLE_MTG_DOC_ID'] || '';

  console.log('📌 現在の SINGLE_MTG_DOC_ID:', currentDocId);

  // 1. ScriptProperties 内の全ドラフトを確認
  Object.keys(props).forEach(k => {
    if (k.startsWith('HENSHUKUN_DRAFT_')) {
      try {
        const obj = JSON.parse(props[k]);
        const slotCount = obj.slots ? Object.keys(obj.slots).length : 0;
        console.log(`💾 プロパティ保存発見: ${k} (保存枠数: ${slotCount}件) ${k.includes(currentDocId) ? '★現在対象' : '⚠️別DocID'}`);
      } catch (e) {
        console.log(`💾 プロパティ保存発見: ${k} (文字数: ${props[k].length})`);
      }
    }
  });

  // 2. 『編集くん_一時保存』シートの確認
  const draftSh = ss.getSheetByName('編集くん_一時保存');
  if (draftSh && draftSh.getLastRow() >= 2) {
    console.log(`\n📄 『編集くん_一時保存』シート: ${draftSh.getLastRow() - 1} 行のデータあり`);
    const preview = draftSh.getRange(1, 1, Math.min(10, draftSh.getLastRow()), draftSh.getLastColumn()).getDisplayValues();
    preview.forEach((r, idx) => console.log(`  [Row ${idx + 1}] ${r.join(' | ').slice(0, 200)}`));
  } else {
    console.log('\n📄 『編集くん_一時保存』シート: データなし');
  }

  // 3. 『ドキュメント確定出力JSON』シートの確認（前回提出時のスナップショット）
  const outSh = ss.getSheetByName('ドキュメント確定出力JSON');
  if (outSh && outSh.getLastRow() >= 2) {
    console.log(`\n📤 『ドキュメント確定出力JSON』シート: ${outSh.getLastRow() - 1} 件残存`);
  }

  // 4. Googleドライブ内の直近MTGドキュメント一覧と「確定案・担当者」埋まり件数をチェック
  const folder = DriveApp.getFolderById('1syiypDralKBqHwCL-0R2XeJXbhxXVcLl');
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  const docList = [];
  while (files.hasNext()) {
    const f = files.next();
    if (!f.isTrashed() && f.getName().includes('超募集対策MTG_')) {
      docList.push(f);
    }
  }
  docList.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());

  console.log('\n📁 直近のMTGドキュメント一覧（最新5件）:');
  docList.slice(0, 5).forEach(f => {
    const id = f.getId();
    const doc = DocumentApp.openById(id);
    const body = doc.getBody();
    let totalRows = 0;
    let decidedRows = 0;
    for (let i = 0; i < body.getNumChildren(); i++) {
      const el = body.getChild(i);
      if (el.getType() === DocumentApp.ElementType.TABLE) {
        const tbl = el.asTable();
        if (tbl.getNumRows() >= 2 && tbl.getRow(0).getCell(0).getText().includes('拠点・不在時間帯')) {
          for (let r = 1; r < tbl.getNumRows(); r++) {
            totalRows++;
            const dec = tbl.getRow(r).getCell(2).getText().replace(/\r\n|\r/g, '\n').split('\n').filter(l => !l.trim().startsWith('→')).join('').trim();
            const asg = tbl.getRow(r).getCell(3).getText().trim();
            if (dec || asg) decidedRows++;
          }
        }
      }
    }
    const updatedStr = Utilities.formatDate(f.getLastUpdated(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    console.log(`  ・${f.getName()} [最終更新: ${updatedStr}] ID=${id} ${id === currentDocId ? '👈現在リンク中' : ''}`);
    console.log(`    ➔ 確定案/担当が入っている枠: ${decidedRows} / ${totalRows} 枠 | URL: ${f.getUrl()}`);
  });
}