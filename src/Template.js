function backupSpreadsheet() {
  GASLibrary.copySpreadsheetToDrive('1nsBZsEA_Sxz6KAPRIrP34NrQlk-HUQdx', "Backup"); // _20nn BMS Forms
}


function hideDoneActions() {
  GASLibrary.hideDoneActions();
}

function resetFilter(){
  GASLibrary.resetFilter();
}

function sortActiveSheet() {
  GASLibrary.sortSheetByConfig(SS, SORT_CONFIGS);
}
