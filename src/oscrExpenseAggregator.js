function oscrExpenseAggregator() {
  // 実行時の年度と一つ前の年度を出力する
  const currentYear = new Date().getFullYear();
  const currentMonth = new Date().getMonth() + 1;
  const targetYears =
    currentMonth > 3
      ? [String(currentYear - 1), String(currentYear)]
      : [String(currentYear - 2), String(currentYear - 1)];
  // 数字4桁のシート名を取得
  const inputSheetNames = outputSpreadSheet
    .getSheets()
    .map(x => x.getName())
    .filter(x => /^\d{4}$/.test(x) && targetYears.includes(x));
  if (inputSheetNames.length === 0) {
    return;
  }
  inputSheetNames.forEach(sheetName => {
    const sheet1 = createSheetForOscrExpenseAggregator_(
      outputSpreadSheet.getSheetByName(sheetName),
      `${sheetName}_1`
    );
    sheet1
      .getRange(
        outputSheetNumber.get('headerRow'),
        outputSheetNumber.get('sumColumn'),
        sheet1.getLastRow(),
        1
      )
      .setBorder(null, true, null, true, null, null);
    sheet1.hideColumns(
      outputSheetNumber.get('monthlyAverageColumn'),
      sheet1.getLastColumn() - outputSheetNumber.get('monthlyAverageColumn') + 1
    );
    const sheet2 = createSheetForOscrExpenseAggregator_(
      outputSpreadSheet.getSheetByName(sheetName),
      `${sheetName}_2`
    );
    sheet2.hideColumns(outputSheetNumber.get('yearStartColumn'), 12);
  });
  exportSheetsToPDF_(outputSpreadSheet, targetYears);
}
function createSheetForOscrExpenseAggregator_(inputSheet, outputSheetName) {
  const targetSheet = getSheet_(outputSpreadSheet, outputSheetName);
  copyValuesAndColumnWidths_(inputSheet, targetSheet);
  targetSheet.protect().setWarningOnly(true);
  return targetSheet;
}
function copyValuesAndColumnWidths_(inputSheet, outputSheet) {
  outputSheet.clear();
  inputSheet.getDataRange().copyTo(outputSheet.getRange(1, 1));
  let columnWidths = [];
  for (let i = 1; i <= inputSheet.getLastColumn(); i++) {
    columnWidths.push(inputSheet.getColumnWidth(i));
  }
  for (let i = 1; i <= outputSheet.getLastColumn(); i++) {
    if (columnWidths[i - 1] !== outputSheet.getColumnWidth[i]) {
      outputSheet.setColumnWidth(i, columnWidths[i - 1]);
    }
  }
  const hideRowIndex = outputSheet
    .getRange(
      1,
      outputSheetNumber.get('itemColumn'),
      outputSheet.getLastRow(),
      1
    )
    .getValues()
    .map((value, idx) => (value[0] === validationSumText ? idx : null))
    .filter(x => x !== null);
  if (hideRowIndex.length > 0) {
    outputSheet.hideRows(hideRowIndex[0] + 1);
  }
}
