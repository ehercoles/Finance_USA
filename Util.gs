function logError(message) {
  MailApp.sendEmail('ehercoles@gmail.com', 'GAS error', message);
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
