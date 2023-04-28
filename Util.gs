function deleteEmptyRows(sheet){
  var maxRows = sheet.getMaxRows(); 
  var lastRow = sheet.getLastRow();
  sheet.deleteRows(lastRow+1, maxRows-lastRow);
}

//a.sort(sortFunction);
function sortFunction(a, b) {
  var col = 1;
  
  if (a[col] === b[col]) {
    return 0;
  }
  else {
    return (a[col] < b[col]) ? -1 : 1;
  }
}
