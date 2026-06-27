// ============================================================
// RAVINA Price API — Google Apps Script
// ============================================================
// 【設置手順】
// 1. スプレッドシートを開く → 拡張機能 → Apps Script
// 2. このコードを貼り付けて保存
// 3. デプロイ → 新しいデプロイ → 種類:ウェブアプリ
//    ・実行ユーザー : 自分
//    ・アクセスできるユーザー : 全員
// 4. 発行されたURLをHTMLの GAS_URL に貼る
// 5. スプレッドシートの共有設定を「非公開」に戻す
// ============================================================

// Price_Num シートのGID（URLの gid=XXXX の値）
const PRICE_NUM_GID = "1469569588";

// ── エントリポイント ──────────────────────────────────────
function doGet(e) {
  const sheet = (e.parameter.sheet || "").toLowerCase();

  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const rows = getSheetRows(ss, PRICE_NUM_GID);
    const csv  = rowsToCsv(rows);

    return ContentService
      .createTextOutput(csv)
      .setMimeType(ContentService.MimeType.TEXT)
      // GitHub Pages (ravinala-oka.github.io) からのアクセスを許可
      // 必要に応じて "*" に変更すれば全オリジン許可
      ;

  } catch (err) {
    return ContentService
      .createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ── GIDでシートを特定して全行取得 ────────────────────────
function getSheetRows(ss, gid) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (String(sh.getSheetId()) === String(gid)) {
      return sh.getDataRange().getValues();
    }
  }
  throw new Error("シートが見つかりません: GID=" + gid);
}

// ── 2次元配列 → CSV文字列 ────────────────────────────────
function rowsToCsv(rows) {
  return rows.map(row =>
    row.map(cell => {
      const s = String(cell === null || cell === undefined ? "" : cell);
      // カンマ・ダブルクォート・改行を含む場合はクォート
      if (s.includes(",") || s.includes('"') || s.includes("\n")) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(",")
  ).join("\n");
}
