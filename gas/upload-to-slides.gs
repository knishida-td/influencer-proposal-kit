/**
 * Google Apps Script: PPTXファイルをGoogle Driveにアップロード → Google Slidesに変換
 *
 * デプロイ手順:
 * 1. https://script.google.com で新規プロジェクト作成
 * 2. このコードを貼り付け
 * 3. FOLDER_ID を出力先フォルダのIDに変更（任意）
 * 4. デプロイ → ウェブアプリ → アクセス: 全員 → デプロイ
 * 5. 表示されたURLをコピー → install.shのGAS_URLに設定
 */

// 出力先Google DriveフォルダID（空ならマイドライブ直下）
const FOLDER_ID = "";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const fileName = data.name || "提案資料.pptx";
    const fileBase64 = data.file;

    if (!fileBase64) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "file (base64) is required" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Base64デコード → Blob作成
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileBase64),
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      fileName
    );

    // Google Driveにアップロード
    let folder;
    if (FOLDER_ID) {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } else {
      folder = DriveApp.getRootFolder();
    }

    const file = folder.createFile(blob);

    // Google Slides形式に変換
    const slidesFile = Drive.Files.copy(
      { title: fileName.replace(/\.pptx$/i, ""), mimeType: "application/vnd.google-apps.presentation" },
      file.getId()
    );

    // 元のPPTXファイルを削除（Slides版のみ残す）
    file.setTrashed(true);

    const slidesUrl = "https://docs.google.com/presentation/d/" + slidesFile.id + "/edit";

    return ContentService.createTextOutput(
      JSON.stringify({
        success: true,
        id: slidesFile.id,
        url: slidesUrl,
        name: slidesFile.title
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// テスト用: GETでヘルスチェック
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "influencer-proposal-kit upload endpoint" })
  ).setMimeType(ContentService.MimeType.JSON);
}
