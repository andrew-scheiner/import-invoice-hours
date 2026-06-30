function backupSpreadsheet() {
  const targetFolderId = SS.getRangeByName('settings__driveCopyFolderId').getValue(); //OR hard code string
  const copyLabel = 'Backup'; // Select Archive, Backup
  GASLibrary.copySpreadsheetToDrive(targetFolderId, copyLabel);
}


