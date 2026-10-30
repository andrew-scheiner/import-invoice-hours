function setPriorityAndStatusDropdownLists() {
  GASLibrary.setGlobalLookupDropdownList('ConfigChange', 'Priority', 'cpi__priority', 1, {
    invalidHandling: 'reject',
  });
  GASLibrary.setGlobalLookupDropdownList('ConfigChange', 'Status', 'cpi__status', 1, {
    invalidHandling: 'reject',
  });
}

// GlobalLookups url: https://docs.google.com/spreadsheets/d/12itBPK7mF78Jvg-gW_00wVNbamwUMqiocUg_fFz9jBg/edit?gid=0#gid=0
