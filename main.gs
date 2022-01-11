function getCreditCardInfo(){
　let checkDate = new Date();
  checkDate.getDate() > 20 ? checkDate.setMonth(checkDate.getMonth() + 1) : checkDate.setMonth(checkDate.getMonth());
  const checkTargetMonthName = String(checkDate.getMonth() + 1).padStart(2, '0');
  const checkTargetSheetName = checkDate.getFullYear() + checkTargetMonthName;
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checkTargetSheetName).getRange('A2').getValue() != ''){
    return;
  }
  let searchBefore = new Date();
  let searchAfter = new Date();
  searchAfter.setMonth(searchAfter.getMonth() - 1);
  const targetTerm = 'subject:(クレジットカード明細) after:' + searchAfter.toLocaleDateString() + ' before:' + searchBefore.toLocaleDateString();
  const gmailThreads = GmailApp.search(targetTerm, 0, 1);
  gmailThreads.forEach(
    thread => thread.getMessages().forEach(
      message => message.getAttachments().forEach(attachment => {
        const csvname = attachment.getName().toLowerCase();
        const namecheck = checkTargetSheetName + '.csv';
        if (csvname == namecheck){
          const csvtext = attachment.getDataAsString('cp932');
          const splitLf = csvtext.split(/\n/);
          let splitComma = splitLf.map(x => x.split(','));
          splitComma[0][0] = '';
          splitComma[0][1] = '';
          const maxIdx = splitComma.map(x => x.length).reduce((x, y) => Math.max(x, y));
          const setCsvValues = splitComma.map(x => {
            let res;
            if (x.length < maxIdx){
              const pushCount = maxIdx - x.length;
              const temp = new Array(pushCount).fill('');
              res = x.concat(temp);
            } else {
              res = x;
            }
            if (res[1] == 'カブシキガイシヤボツクスジヤパ'){
              res[1] = 'BOX';
            }
            return res;
          });
          const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checkTargetSheetName);
          targetSheet.clearContents();
          targetSheet.getRange(1, 1, setCsvValues.length, setCsvValues[0].length).setValues(setCsvValues);
        };
      })
    )
  );
}
function ssInit(){
  const today = new Date();
  const yyyy = today.getFullYear();
  const nextYyyy = yyyy + 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName('List');
  listSheet.getRange('B1').setValue(yyyy + '/04/01');
  const listLastRow = listSheet.getDataRange().getLastRow();
  if (listSheet.getRange(2, 15, listLastRow, 1).getValues().filter(x => x == '#DIV/0!').length == 0){
    listSheet.getRange(2, 16, listLastRow, 1).setValues(listSheet.getRange(2, 15, listLastRow, 1).getValues());
  }  
  const sheets = ss.getSheets();
  const targetSheetName = /^\d{6}$/;
  const targetSheets = sheets.filter(x => targetSheetName.test(x.getName()));
  targetSheets.forEach(x => {
    const targetMm = x.getName().substring(4, 6);
    const setYear = targetMm < 4 ? nextYyyy : yyyy;
    x.setName(setYear + targetMm);
    x.clearContents();
  });
}