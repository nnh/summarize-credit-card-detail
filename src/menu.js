/**
 * スプレッドシートが開かれたときに実行される特別な関数。
 * カスタムメニューをツールバーに追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // メニューの組み立て
  ui.createMenu('🛠️ 業務メニュー') // メニューのメインタイトル
    .addItem('1. CSVインポートと集計実行', 'importCsvWithLogic')
    .addSeparator() // 区切り線
    .addItem('2. 理事会用PDFを作成', 'exportSummaryToPDF')
    .addSeparator()
    .addItem('🔍 データの整合性チェック', 'validateSummaries_')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('⚙️ 設定・管理')
        .addItem('初期シート作成', 'setupManagementSheets')
    )
    .addToUi();
}

/**
 * 注意：スクリプトエディタでこの関数を貼り付けた直後は、
 * スプレッドシートをリロード（再読み込み）するか、
 * エディタの「実行」ボタンで onOpen を一度走らせることでメニューが表示されます。
 */
