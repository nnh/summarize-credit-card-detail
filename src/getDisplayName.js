function getDisplayName() {
  const csvSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    inputSheetNames.get('csvSheet')
  );
  const displayNameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    inputSheetNames.get('displayNameSheet')
  );
  if (csvSheet === null || displayNameSheet === null) {
    return;
  }
  const csvValues = csvSheet.getDataRange().getValues();
  const displayNames = displayNameSheet
    .getDataRange()
    .getValues()
    .filter(x => x[displayNameSheetIndex.get('displayName')] !== '');
  const displayValues = csvValues.map((csvValue, idx) => {
    if (idx === 0) {
      return [
        csvValue[csvSheetIndex.get('displayName')],
        csvValue[csvSheetIndex.get('sortOrder')],
        csvValue[csvSheetIndex.get('key')],
      ];
    }
    const temp = displayNames
      .map(nameAndDisplayName =>
        new RegExp(nameAndDisplayName[displayNameSheetIndex.get('name')]).test(
          csvValue[csvSheetIndex.get('name')]
        )
          ? [
              nameAndDisplayName[displayNameSheetIndex.get('displayName')],
              nameAndDisplayName[displayNameSheetIndex.get('sortOrder')],
              `=A${idx + 1} & I${idx + 1}`,
            ]
          : null
      )
      .filter(x => x !== null);
    if (temp.length === 0) {
      return ['エラー：表示名なし', -1];
    }
    return temp[0];
  });
  csvSheet
    .getRange(
      1,
      csvSheetNumber.get('displayName'),
      displayValues.length,
      displayValues[0].length
    )
    .setValues(displayValues);
  csvSheet.hideColumns(csvSheetNumber.get('key'), 1);
}
