/**
 * アクティブなスプレッドシートに必要な管理用シートをセットアップする関数。
 * 既存のシートがある場合は、設定のみを更新（上書き）します。
 */
function setupManagementSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. lnk_Software一覧シートのセットアップ ---
  let lnkSheet = ss.getSheetByName('lnk_Software一覧');
  if (!lnkSheet) {
    lnkSheet = ss.insertSheet('lnk_Software一覧');
  }

  /** @type {string} 外部ソースからのインポート式 */
  const importFormula =
    '=importrange("https://docs.google.com/spreadsheets/d/1qL7QSCLeRemHBE8lZx8WWGCTDB_8WasJtkhAZF8O1Zg/edit#gid=1389197276","Software一覧!A:Z")';

  // A1セルに数式を設定
  lnkSheet.getRange('A1').setFormula(importFormula);

  // --- 2. サービス一覧シートのセットアップ ---
  let serviceSheet = ss.getSheetByName('サービス一覧');
  if (!serviceSheet) {
    serviceSheet = ss.insertSheet('サービス一覧');
  }

  /** @type {string[]} 見出し行の定義 */
  const headers = [
    'サービス名',
    '使用者',
    'カテゴリー',
    'サービス内容',
    '使用目的',
    '年間額',
    '月額',
    '1ライセンス月額',
    '1ライセンス（税込）',
    '単位',
    '契約ライセンス数',
  ];

  // 1行目にタイトルを書き込み
  serviceSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // デザイン調整：太字、背景色、ウィンドウ枠の固定
  const headerRange = serviceSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  serviceSheet.setFrozenRows(1);

  // 金額が入る列（F列とG列）をあらかじめカンマ区切りに設定
  serviceSheet.getRange('F2:G').setNumberFormat('#,##0');

  // 完了通知
  uiAlert_(
    'シートのセットアップが完了しました。\n「lnk_Software一覧」シートでインポートのアクセス許可を承認してください。'
  );
}
