function readCsvFile() {
  const outputCsvSheet = getSheet_(outputSpreadSheet, inputSheetNames.get("csvSheet"));
  let [fileNames, startRow] = getTargetfileNamesAndOutputStartRow_(outputCsvSheet);
  fileNames.forEach((fileId, fileName) => {
    const ymText = "'" + fileName.substring(0, 4) + "年" + String(Number(fileName.substring(4, 6))) + "月";
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString("MS932");
    const inputCsvData = Utilities.parseCsv(content);
    const csvData = editCsvData_(inputCsvData).map(x => [ymText, ...x]);
    const csvLength = csvData.length;
    outputCsvSheet.getRange(startRow, 1, csvLength, csvData[0].length).setValues(csvData);
    startRow = startRow + csvLength;
  });
  getDisplayName();
  outputCsvSheet.hideColumns(csvSheetNumber.get("name") + 1, 3);
  outputCsvSheet.hideColumns(csvSheetNumber.get("price") + 1, 1);
  const check = outputSpreadSheet.getSheetByName(inputSheetNames.get("summarySheet"));
  if (check === null){
    const summarySheet = getSheet_(inputSheetNames.get("summarySheet"));
    summarySheet.getRange(1, 1).setValue(`=QUERY(csv!A:K, "SELECT K, A, I, J, SUM(G) WHERE A IS NOT NULL GROUP BY K, A, I, J ORDER BY A, J label J '表示順'", 1)`);
  }
}
function editCsvData_(inputCsvData) {
  const inputCsvItemIdx = 1;
  const inputCsvPriceIdx = 2;
  // Delete the first line as it is unnecessary.
  const tempCsvData = inputCsvData.filter((_, idx) => idx > 0);
  // If the price is not numeric, join the values in the price with the item and add a space in the last column.
  const csvData = tempCsvData.map(values =>{
    if (values[inputCsvItemIdx] === ""){
      values[inputCsvItemIdx] = sumText;
    }
    if (values[inputCsvPriceIdx] === "" || !isNaN(values[inputCsvPriceIdx])){
      return(values);
    }
    values[inputCsvItemIdx] = values[inputCsvItemIdx] + values[inputCsvPriceIdx];
    values.push("");
    const res = [...values.slice(0, inputCsvPriceIdx), ...values.slice(inputCsvPriceIdx + 1)];
    return(res);
  });
  return(csvData);
}
function getTargetfileNamesAndOutputStartRow_(outputCsvSheet) {
  const files = csvSaveFolder.getFiles();
  const fileNames = new Map;
  while (files.hasNext()) {
    const file = files.next()
    fileNames.set(file.getName(), file.getId());
  }
  let lastRow = outputCsvSheet.getLastRow();
  const headerArray = ["対象年月", "年月日", "項目名", "filler1", "filler2", "filler3", "金額", "filler4", "表示名", "表示順", "キー"];
  if (lastRow === 0){
    outputCsvSheet.getRange(1, 1, 1, headerArray.length).setValues(Array(headerArray.map(x => [x])));
    lastRow = lastRow + 2;
  }
  return([fileNames, lastRow]);
}

