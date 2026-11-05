/**
 * ============================================================================
 * 🔍 第2弾ファクト確認用デバッグスクリプト（データ書き換えなし）
 * 1) 「柿本」先生のシステムメール（医師質問・DS回答・ナレッジシート・Gmail）の全履歴調査
 * 2) 直近の医師返信メール（【要確認】医師からのお問い合わせ）とスタッフ打診のペア構造調査
 * ============================================================================
 */
function debugInspectDoctorRepliesAndKakimoto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("=== 【調査A】シート（医師質問 / DS回答 / ナレッジシート）内の「柿本」先生データ ===");
  ["医師質問", "DS回答", "ナレッジシート"].forEach(function(sName) {
    const sh = ss.getSheetByName(sName);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getDisplayValues();
    const hits = [];
    for (let r = 1; r < vals.length; r++) {
      if (vals[r].join(" ").includes("柿本")) {
        hits.push({ row: r + 1, data: vals[r].map(function(v) { return String(v).replace(/\s+/g, " ").slice(0, 120); }) });
      }
    }
    Logger.log("📄 [" + sName + "] 内の「柿本」先生ヒット (" + hits.length + "件):\n" + JSON.stringify(hits.slice(-5), null, 2));
  });

  Logger.log("\n=== 【調査B】Gmail内の「柿本」先生に関する全メール（直近60日） ===");
  const kakimotoThreads = GmailApp.search('"柿本" newer_than:60d', 0, 15);
  const kakimotoMsgs = [];
  kakimotoThreads.forEach(function(th) {
    th.getMessages().forEach(function(m) {
      const body = String(m.getPlainBody() || "");
      if (!body.includes("柿本")) return;
      kakimotoMsgs.push({
        date: Utilities.formatDate(m.getDate(), "Asia/Tokyo", "yyyy/MM/dd HH:mm"),
        subject: m.getSubject(),
        snippet: body.replace(/\s+/g, " ").slice(0, 220)
      });
    });
  });
  Logger.log("📧 Gmail「柿本」先生ヒット (" + kakimotoMsgs.length + "件):\n" + JSON.stringify(kakimotoMsgs.slice(-8), null, 2));

  Logger.log("\n=== 【調査C】直近の『ナレッジシート』最新10件（スタッフ打診 ⇔ 医師返信の突合状況） ===");
  const kbSh = ss.getSheetByName("ナレッジシート");
  if (kbSh && kbSh.getLastRow() >= 2) {
    const lastR = kbSh.getLastRow();
    const startR = Math.max(2, lastR - 9);
    const kbRows = kbSh.getRange(startR, 1, lastR - startR + 1, 8).getDisplayValues();
    const samples = kbRows.map(function(r, idx) {
      return {
        row: startR + idx,
        id: r[0],
        doctor: r[1],
        docReply: String(r[2] || "").replace(/\s+/g, " ").slice(0, 140),
        staffPitch: String(r[3] || "").replace(/\s+/g, " ").slice(0, 140),
        sender: r[4],
        date: r[6],
        flag: r[7]
      };
    });
    Logger.log("📋 ナレッジシート最新10件:\n" + JSON.stringify(samples, null, 2));
  }

  Logger.log("\n=== 【調査D】『医師質問』シートの未処理（スタッフ回答と未マッチの単独メール）最新10件 ===");
  const qSh = ss.getSheetByName("医師質問");
  if (qSh && qSh.getLastRow() >= 2) {
    const qVals = qSh.getDataRange().getDisplayValues();
    const unmatched = [];
    for (let r = Math.max(1, qVals.length - 30); r < qVals.length; r++) {
      if (String(qVals[r][6] || "").trim() !== "処理済み") {
        unmatched.push({
          row: r + 1,
          date: qVals[r][1],
          doctor: qVals[r][2],
          content: String(qVals[r][3] || "").replace(/\s+/g, " ").slice(0, 160),
          status: qVals[r][6]
        });
      }
    }
    Logger.log("📋 医師質問（未マッチ分）最新最大10件:\n" + JSON.stringify(unmatched.slice(-10), null, 2));
  }
}