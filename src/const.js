const csvSaveFolder = DriveApp.getFolderById(
  PropertiesService.getScriptProperties().getProperty('csvSaveFolderId')
);
const outputSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const sumText = '計';
const validationSumText = 'チェック用：合計差分';
const inputSheetNames = new Map([
  ['csvSheet', 'csv'],
  ['summarySheet', 'summary'],
  ['itemsSheet', '表示名マスタ'],
  ['serviceListSheet', 'サービス一覧'],
  ['displayNameSheet', '表示名と項目名の対応'],
]);
const displayNameSheetIndex = new Map([
  ['displayName', 0],
  ['sortOrder', 1],
  ['name', 2],
]);
const csvSheetIndex = new Map([
  ['targetYear', 0],
  ['name', 2],
  ['price', 6],
  ['displayName', 8],
  ['sortOrder', 9],
  ['key', 10],
]);
const itemSheetIndex = new Map([
  ['name', 0],
  ['sortOrder', 1],
  ['startYear', 2],
  ['endYear', 3],
]);
const serviceSheetIndex = new Map([
  ['name', 0],
  ['category', 2],
  ['service', 3],
]);
const outputSheetIndex = new Map([
  ['itemColumn', 0],
  ['yearStartColumn', 1],
  ['sumColumn', 13],
  ['monthlyAverageColumn', 14],
  ['monthlyAverageLastYearColumn', 15],
  ['differenceColumn', 16],
  ['categoryColumn', 17],
  ['serviceColumn', 18],
  ['headerRow', 0],
  ['bodyStartRow', 1],
]);
function convertIndexToNumber_(targetMap) {
  const result = new Map();
  targetMap.forEach((index, key) => result.set(key, index + 1));
  return result;
}
const outputSheetNumber = convertIndexToNumber_(outputSheetIndex);
const csvSheetNumber = convertIndexToNumber_(csvSheetIndex);
function getSheet_(spreadsheet, sheetName) {
  let targetSheet = spreadsheet.getSheetByName(sheetName);
  if (targetSheet === null) {
    targetSheet = spreadsheet.insertSheet();
    SpreadsheetApp.getActiveSheet().setName(sheetName);
  }
  return targetSheet;
}
function getTargetYears_(targetYear) {
  const result = [4, 5, 6, 7, 8, 9, 10, 11, 12]
    .map(month => `'${targetYear}年${month}月`)
    .concat([1, 2, 3].map(month => `'${Number(targetYear) + 1}年${month}月`));
  return result;
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // メニューに表示されるメニュー項目の作成
  ui.createMenu('クレジットカード明細集計')
    .addItem('CSV保存', 'getCreditCardInfo')
    .addItem('CSV取り込み', 'readCsvFile')
    .addItem('表示名設定', 'getDisplayName')
    .addItem('集計用シート作成', 'execCreateListSheetByMenu')
    .addItem('PDF出力', 'oscrExpenseAggregator')
    .addToUi();
}
