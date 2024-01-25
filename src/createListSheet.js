function execCreateListSheet() {
  const targetYear = new Date().getFullYear().toString();
  if (outputSpreadSheet.getSheetByName(targetYear) !== null){
    return;
  }
  createListSheet_(targetYear);
}
function createListSheet_(targetYear) {
  const itemsSheet = outputSpreadSheet.getSheetByName(inputSheetNames.get("itemsSheet"));
  if (itemsSheet === null){
    console.log("No sheet required.");
    return;
  }
  const itemsValue = itemsSheet.getDataRange().getValues();
  const targetItemsValue = itemsValue.filter((values, idx) => idx === 0 || (values[itemSheetIndex.get("startYear")] <= Number(targetYear) && Number(targetYear) <= values[itemSheetIndex.get("endYear")]));
  const sortItemsValue = sortMatrix_(targetItemsValue, itemSheetIndex.get("sortOrder"));
  const outputItems = sortItemsValue.filter((_, idx) => idx > 0).map(x => [x[itemSheetIndex.get("name")]]);
  const outputTargetYearColumns = getTargetYears_(targetYear);
  const outputColumns = [[...outputTargetYearColumns, sumText, "月平均", "前年度月平均", "差額", "", "サービス内容"]];
  const monthCols = Array(12).fill().map((_, idx) => idx + 2);
  const outputFormulas = outputItems.map((_, idx) => {
    const yearsFormulas = monthCols.map(col => 
      `=if(not(isna(match(indirect("R1C${col}", false)&indirect("R${idx + 2}C1", false), summary!A:A,0))), vlookup(indirect("R1C${col}", false)&indirect("R${idx + 2}C1", false),summary!$A:$E,5,false), 0)`
    );
    const sumFormulas = [
      `=SUM(B${idx + 2}:M${idx + 2})`,
      `=round(N${idx + 2}/countif(indirect(ADDRESS(match("計",$A:$A,0),2) & ":" & ADDRESS(match("計",$A:$A,0),13)),">0"))`,
      `=IF(not(ISERR(${Number(targetYear) - 1}!O${idx + 2})), ${Number(targetYear) - 1}!O${idx + 2}, "")`,
      `=O${idx + 2}-P${idx + 2}`,
      `=iferror(vlookup(A${idx + 2},'${inputSheetNames.get("serviceListSheet")}'!A:D,3,false))`, 
      `=iferror(vlookup(A${idx + 2},'${inputSheetNames.get("serviceListSheet")}'!A:D,4,false))`
    ];
    return([...yearsFormulas, ...sumFormulas]);
  });
  const targetSheet = getSheet_(outputSpreadSheet, targetYear);  
  targetSheet.clear();
  targetSheet.getRange(outputSheetNumber.get("bodyStartRow"), outputSheetNumber.get("itemColumn"), outputItems.length, 1).setValues(outputItems);
  const outputValues = [...outputColumns, ...outputFormulas];
  targetSheet.getRange(outputSheetNumber.get("headerRow"), outputSheetNumber.get("yearStartColumn"), outputValues.length, outputValues[0].length).setValues(outputValues);
  const sumValues = monthCols.map(col => `=sum(indirect("R2C${col}", false):indirect("R${outputValues.length - 1}C${col}", false)) - indirect("R${outputValues.length}C${col}", false)`);
  setSheetFormat_(targetSheet);
  targetSheet.getRange(outputValues.length + 1, outputSheetNumber.get("itemColumn"), 1, sumValues.length + 1).setValues([[validationSumText, ...sumValues]]);
  targetSheet.protect().setWarningOnly(true)
}
function setSheetFormat_(targetSheet){
  SpreadsheetApp.flush();
  targetSheet.setColumnWidth(outputSheetNumber.get("itemColumn"), 300);
  targetSheet.setColumnWidths(outputSheetNumber.get("yearStartColumn"), outputSheetNumber.get("categoryColumn") - outputSheetNumber.get("yearStartColumn"), 85);
  targetSheet.setColumnWidth(outputSheetNumber.get("monthlyAverageLastYearColumn"), 95);
  targetSheet.setColumnWidth(outputSheetNumber.get("categoryColumn"), 168);
  targetSheet.setColumnWidth(outputSheetNumber.get("serviceColumn"), 419);
  targetSheet.setFrozenColumns(outputSheetNumber.get("itemColumn"));
  targetSheet.getRange(outputSheetNumber.get("bodyStartRow"), outputSheetNumber.get("itemColumn"), targetSheet.getLastRow() - 2, targetSheet.getLastColumn()).setBorder(true, false, true, false, false, false);
  targetSheet.getRange(outputSheetNumber.get("headerRow"), outputSheetNumber.get("yearStartColumn"), targetSheet.getLastRow() - 1, outputSheetNumber.get("sumColumn") - outputSheetNumber.get("yearStartColumn")).setBorder(null, true, null, true, null, null);
  targetSheet.getRange(outputSheetNumber.get("headerRow"), outputSheetNumber.get("sumColumn"), targetSheet.getLastRow() - 1, 4).setBorder(null, true, null, true, null, null);
  targetSheet.getRange(outputSheetNumber.get("headerRow"), outputSheetNumber.get("itemColumn"), 1, targetSheet.getLastColumn()).setHorizontalAlignment("center");
  targetSheet.getRange(outputSheetNumber.get("bodyStartRow"), outputSheetNumber.get("yearStartColumn"), targetSheet.getLastRow(), outputSheetNumber.get("categoryColumn") - outputSheetNumber.get("yearStartColumn")).setNumberFormat('#,##0');
}
function sortMatrix_(inputMatrix, sortIndex) {
  if (!Array.isArray(inputMatrix) || inputMatrix.length === 0 || !Array.isArray(inputMatrix[0])) {
    console.error('Invalid input. Please provide a non-empty 2D array.');
    return;
  }
  const header = inputMatrix.filter((_, idx) => idx === 0);
  const target = inputMatrix.filter((_, idx) => idx > 0);
  target.sort((a, b) => a[sortIndex] - b[sortIndex]);
  const res = [...header, ...target];
  return res;
}

