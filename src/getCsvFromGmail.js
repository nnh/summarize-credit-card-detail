function getCreditCardInfo() {
  const existCsvFiles = csvSaveFolder.getFiles();
  const existCsvFileNames = new Set();
  while (existCsvFiles.hasNext()) {
    let file = existCsvFiles.next();
    existCsvFileNames.add(file.getName());
  }
  const targetTerm = 'subject:(クレジットカード明細)';
  GmailApp.search(targetTerm).forEach(thread =>
    thread.getMessages().forEach(message =>
      message.getAttachments().forEach(attachment => {
        if (
          !existCsvFileNames.has(attachment.getName()) &&
          /^\d{6}\.csv/.test(attachment.getName())
        ) {
          if (Number(attachment.getName().substring(0, 4)) >= 2022) {
            console.log(attachment.getName());
            csvSaveFolder.createFile(attachment);
          }
        }
      })
    )
  );
}
