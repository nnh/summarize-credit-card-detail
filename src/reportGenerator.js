/**
 * 実行日から対象年度を判定し、SummaryシートをPDF出力して指定フォルダに保存する。
 */
function exportSummaryToPDF() {
  const props = PropertiesService.getScriptProperties();
  // スクリプトプロパティからPDF保存先フォルダIDを取得
  const pdfFolderId = props.getProperty('PDF_FOLDER_ID');

  if (!pdfFolderId) {
    uiAlert_(
      'エラー: スクリプトプロパティ「PDF_FOLDER_ID」を設定してください。'
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const now = new Date();

  // ファイル名用：実行年月日 (YYYYMMDD)
  const yyyymmdd = Utilities.formatDate(now, 'JST', 'yyyyMMdd');

  // 実行時の「会計年度」を取得 (2026年3月実行なら2025年度が今年度)
  const currentMonth = now.getMonth() + 1;
  let fiscalYear = now.getFullYear();
  if (currentMonth >= 1 && currentMonth <= 3) fiscalYear--;

  // 対象とする2つの年度 (今年度と前年度)
  const targetYears = [fiscalYear.toString(), (fiscalYear - 1).toString()];

  targetYears.forEach(yyyy => {
    const sheet = ss.getSheetByName('Summary_' + yyyy);
    if (!sheet) {
      console.log(`シート Summary_${yyyy} が見つからないためスキップします。`);
      return;
    }
    const gid = sheet.getSheetId();

    try {
      // --- パターン1: A列〜N列 (項目名〜計) ---
      const name1 = `${yyyymmdd} OSCR理事会用1(${yyyy}).pdf`;
      const blob1 = generatePdfBlob_(ssId, gid, name1, 'A:N');
      DriveApp.getFolderById(pdfFolderId).createFile(blob1);

      // --- パターン2: A列 ＋ N〜S列 (項目名 ＋ 計〜サービス内容) ---
      const name2 = `${yyyymmdd} OSCR理事会用2(${yyyy}).pdf`;
      const blob2 = generatePdfBlobWithHiddenCols_(sheet, ssId, gid, name2);
      DriveApp.getFolderById(pdfFolderId).createFile(blob2);

      console.log(`年度 ${yyyy} のPDF作成に成功しました。`);
    } catch (e) {
      uiAlert_(`年度 ${yyyy} のPDF作成中にエラーが発生しました:\n${e.message}`);
    }
  });

  uiAlert_('PDFの作成処理が終了しました。指定のフォルダを確認してください。');
}

/**
 * 指定範囲のPDF Blobを生成する (連続範囲用)
 * @param {string} ssId - スプレッドシートID
 * @param {number} gid - シートのGID
 * @param {string} fileName - PDFのファイル名
 * @param {string} rangeStr - 出力範囲 (例: "A:N")
 * @returns {GoogleAppsScript.Base.Blob}
 * @private
 */
function generatePdfBlob_(ssId, gid, fileName, rangeStr) {
  const opts = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'A4',
    portrait: 'false', // 横向き
    fitw: 'true', // 幅に合わせる
    sheetnames: 'false',
    printtitle: 'false',
    pagenumbers: 'false',
    gridlines: 'false',
    fzr: 'false', // 固定行を印刷しない
    gid: gid,
    range: rangeStr,
  };

  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  const urlParts = Object.keys(opts).map(key => `${key}=${opts[key]}`);
  const finalUrl = url + urlParts.join('&');

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(finalUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `PDF生成リクエストが失敗しました (Status: ${response.getResponseCode()})`
    );
  }

  return response.getBlob().setName(fileName);
}

/**
 * 不要な列(B-M列)を隠してPDFを生成する。
 * finally句により、エラー時でも必ず列を表示状態に戻します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} ssId - スプレッドシートID
 * @param {number} gid - シートのGID
 * @param {string} fileName - PDFのファイル名
 * @returns {GoogleAppsScript.Base.Blob}
 * @private
 */
function generatePdfBlobWithHiddenCols_(sheet, ssId, gid, fileName) {
  try {
    // B列(2)〜M列(13)を一時的に非表示にする
    sheet.hideColumns(2, 12);
    // スプレッドシートの状態を確定させる
    SpreadsheetApp.flush();

    // 非表示状態のまま、A列〜S列の範囲でPDF作成 (隠れた列は出力されない)
    return generatePdfBlob_(ssId, gid, fileName, 'A:S');
  } finally {
    // 何があっても必ず列を表示状態に戻す
    sheet.showColumns(2, 12);
    SpreadsheetApp.flush();
  }
}
