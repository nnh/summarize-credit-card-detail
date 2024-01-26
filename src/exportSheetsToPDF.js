function exportSheetsToPDF_(spreadsheet, outputTargetYears) {
  const outputFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('pdfSaveFolderId')
  );
  const fileId = spreadsheet.getId();
  const baseUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?id=${fileId}`;
  const parameters = new Map([
    ['gridlines', 'true'],
    ['portrait', 'false'],
  ]);
  const options = generateOptionsString_(parameters);
  const url = baseUrl + options;
  const token = ScriptApp.getOAuthToken();
  const fetchOptions = {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  };
  const todayText = getFormattedDate_();
  outputTargetYears.forEach(year => {
    createPdf_(
      spreadsheet,
      `${year}_1`,
      url,
      fetchOptions,
      outputFolder,
      todayText
    );
    createPdf_(
      spreadsheet,
      `${year}_2`,
      url,
      fetchOptions,
      outputFolder,
      todayText
    );
  });
}
function createPdf_(
  spreadsheet,
  targetSheetName,
  url,
  fetchOptions,
  outputFolder,
  todayText
) {
  spreadsheet.getSheets().forEach(sheet => sheet.showSheet());
  spreadsheet.getSheets().forEach(target => {
    if (target.getName() !== targetSheetName) {
      target.hideSheet();
    }
  });
  const [year, seq] = targetSheetName.split('_');
  const newFileName =
    todayText + ' OSCR理事会用' + seq + '(' + year + ')' + '.pdf';
  const blob = UrlFetchApp.fetch(url, fetchOptions)
    .getBlob()
    .setName(newFileName);
  outputFolder.createFile(blob);
  spreadsheet.getSheets().forEach(sheet => sheet.showSheet());
}
function getFormattedDate_() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0'); // Months are zero-based, so we add 1 and zero-pad to two digits
  const day = String(today.getDate()).padStart(2, '0'); // Zero-pad the day to two digits

  const formattedDate = `${year}${month}${day}`;
  return formattedDate;
}
