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
  const targetTerm1 = 'subject:(クレジットカード明細) after:' + searchAfter.toLocaleDateString() + ' before:' + searchBefore.toLocaleDateString();
  const targetTerm2 = 'subject:(クレジットカード明細) newer:' + searchBefore.toLocaleDateString();
  const targetTerms = [targetTerm1, targetTerm2];
  let gmailThreads;
  for (let i = 0; i < targetTerms.length; i++){
    gmailThreads = GmailApp.search(targetTerms[i], 0, 1);
    if (gmailThreads.length > 0){
      break;
    }
  }
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
  ss.rename(yyyy + '年度クレジットカード明細');
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
  setTriggers();
}
function setTriggers(){
  ScriptApp.getProjectTriggers().forEach(x => ScriptApp.deleteTrigger(x));
  const today = new Date();
  const yyyy = today.getFullYear();
  const nextYyyy = yyyy + 1;
  const targetyyyyMonths = [...Array(9)].map((_, idx) => idx + 3);
  const targetNextYyyyMonths = [...Array(3)].map((_, idx) => idx);
  const targetYyyyMm = targetyyyyMonths.map(x => [yyyy, x, 1]);
  const targetNextYyyyMm = targetNextYyyyMonths.map(x => [nextYyyy, x, 1]);
  const target = targetYyyyMm.concat(targetNextYyyyMm);
  target.forEach(x => {
    let target = new Date(x[0], x[1], x[2]);
    target.setHours(8);
    target.setMinutes(10); 
    ScriptApp.newTrigger('getCreditCardInfo').timeBased().at(target).create();  
  });
}
