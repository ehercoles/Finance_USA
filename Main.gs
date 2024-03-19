var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function addSell(rowData) {
  var sellSheet = spreadsheet.getSheetByName('Sell');
  var row = 2;
  var lastRow = rowData.length + row;
  
  // Remove last column from array (Sell Index)
  var rowData = rowData.map(
    function(val) {
      return val.slice(0, -1);
    });
  
  sellSheet.insertRowsAfter(row, rowData.length);
  sellSheet.getRange('A3:H' + lastRow).setValues(rowData);
  
  var fromRange = sellSheet.getRange('I2:M2');
  fromRange.copyTo(sellSheet.getRange('I3:M' + lastRow), {contentsOnly:false});
}

function addBuy(rowData) {
  var buySheet = spreadsheet.getSheetByName('Buy');
  var lastRow = rowData.length + 1;
  
  // Remove last column from array (Buy Index)
  var rowData = rowData.map(
    function(val) {
      return val.slice(0, -1);
    });
  
  buySheet.insertRowsBefore(2, rowData.length);
  buySheet.getRange('A2:I' + lastRow).setValues(rowData);
}

function sell(usdBrl) {
  var sellRange = spreadsheet.getRangeByName('Sell');
  var positions = spreadsheet.getRangeByName('Position');
  const numRows = sellRange.getNumRows();
  var sellData = [];
  
  for (var i = 1; i <= numRows; i++) {
    var sellQty = sellRange.getCell(i, 1).getValue();
    var sellAp = sellRange.getCell(i, 2).getValue();
      
    if(sellQty > 0 && sellAp > 0) {
      /*
        0: Date
        1: Symbol
        2: Qty
        3: AP
        4: USD AP
        5: Sell qty
        6: Sell AP
        7: USDBRL
        8: Sell Index
      */
      sellData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        positions.getCell(i, 2).getValue(),
        positions.getCell(i, 3).getValue(),
        positions.getCell(i, 4).getValue(),
        sellQty,
        sellAp,
        usdBrl,
        i
      ]);
    }
  }
  
  if (sellData.length > 0) {
    sellData.sort(sortFunction);
    //console.log(sellData);
    
    addSell(sellData);
    
    for (var i = 0; i < sellData.length; i++) {
      var qty = parseInt(sellData[i][2]);
      var sellQty = parseInt(sellData[i][5]);
      var newQty = qty - sellQty;
      var sellIndex = sellData[i][8];
      
      positions.getCell(sellIndex, 2).setValue(newQty);
    }
  }
}

function buy(usdBrl) {
  var buyRange = spreadsheet.getRangeByName('Buy');
  var positions = spreadsheet.getRangeByName('Position');
  const numRows = buyRange.getNumRows();
  var buyData = [];
  
  for (var i = 1; i <= numRows; i++) {
    var buyQty = parseInt(buyRange.getCell(i, 1).getValue());
    var buyAp = parseFloat(buyRange.getCell(i, 2).getValue());
    
    if (buyQty > 0 && buyAp > 0) {
      var qty = parseInt(0 + positions.getCell(i, 2).getValue());
      var ap = parseFloat(0 + positions.getCell(i, 3).getValue());
      var usdAp = parseFloat(0 + positions.getCell(i, 4).getValue());
      
      if (ap == 0) {
        ap = buyAp;
      }
      
      if(usdAp == 0) {
        usdAp = usdBrl;
      }
      
      var newAp = ((qty * ap) + (buyQty * buyAp)) / (qty + buyQty);
      var newUsdAp = ((qty * usdAp) + (buyQty * usdBrl)) / (qty + buyQty);
      
      /*
        0: Date
        1: Symbol
        2: Qty
        3: AP
        4: USD AP
        5: Buy qty
        6: Buy AP
        7: New AP
        8: New USD AP
        9: Buy Index
      */
      buyData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        qty,
        ap,
        usdAp,
        buyQty,
        buyAp,
        newAp,
        newUsdAp,
        i
      ]);
    }
  }
  
  if (buyData.length > 0) {
    buyData.sort(sortFunction);
    //console.log(buyData);
    
    addBuy(buyData);
    
    for (var i = 0; i < buyData.length; i++) {
      var qty = buyData[i][2];
      var buyQty = buyData[i][5];
      var newAp = buyData[i][7];
      var newUsdAp = buyData[i][8];
      var newQty = qty + buyQty;
      var buyIndex = buyData[i][9];
      
      positions.getCell(buyIndex, 2).setValue(newQty);
      positions.getCell(buyIndex, 3).setValue(newAp);
      positions.getCell(buyIndex, 4).setValue(newUsdAp);
    }
  }
}

function clearOrders() {
  spreadsheet.getRangeByName('Sell').setValue('');
  spreadsheet.getRangeByName('Buy').setValue('');
}

function clearPrices() {
  spreadsheet.getRangeByName('SellPrice').setValue('');
  spreadsheet.getRangeByName('BuyPrice').setValue('');
}

function setOrders(mode) {
  try {
    var targetQuantities = spreadsheet.getRangeByName('TargetQuantity');
    var prices = spreadsheet.getRangeByName('Price');
    var sellRange = spreadsheet.getRangeByName('Sell');
    var buyRange = spreadsheet.getRangeByName('Buy');
    const numRows = targetQuantities.getNumRows();
    
    for (var i = 1; i <= numRows; i++) {
      var targetQuantityCell = targetQuantities.getCell(i, 1);
      var qty = targetQuantityCell.getValue();
      var price = prices.getCell(i, 1).getValue();
      
      if ((!mode || mode == 'sell') && targetQuantityCell.getBackgroundColor() == '#ff9900') { // orange
        var quantityCell = sellRange.getCell(i, 1);
        var priceCell = sellRange.getCell(i, 2);

        quantityCell.setValue(qty * -1);
        priceCell.setValue(price);
      }
      else if ((!mode || mode == 'buy') && targetQuantityCell.getBackgroundColor() == '#34a853') { // green
        var quantityCell = buyRange.getCell(i, 1);
        var priceCell = buyRange.getCell(i, 2);

        quantityCell.setValue(qty);
        priceCell.setValue(price);
      }
    }
    
  } catch (err) {
    logError(err.stack);
  }
}

function setPrices() {
  try {
    var prices = spreadsheet.getRangeByName('Price');
    var sellRange = spreadsheet.getRangeByName('Sell');
    var buyRange = spreadsheet.getRangeByName('Buy');
    const numRows = prices.getNumRows();
    
    // Sell range
    for (var i = 1; i <= numRows; i++) {
      var quantityCell = sellRange.getCell(i, 1);
      var qty = quantityCell.getValue();
      
      if (qty > 0) {
        var price = prices.getCell(i, 1).getValue();
        var priceCell = sellRange.getCell(i, 2);
        
        priceCell.setValue(price);
      }
    }

    // Buy range
    for (var i = 1; i <= numRows; i++) {
      var quantityCell = buyRange.getCell(i, 1);
      var qty = quantityCell.getValue();
      
      if (qty > 0) {
        var price = prices.getCell(i, 1).getValue();
        var priceCell = buyRange.getCell(i, 2);
        
        priceCell.setValue(price);
      }
    }
  } catch (err) {
    logError(err.stack);
  }
}

function setSell() {
  setOrders('sell');
}

function setBuy() {
  setOrders('buy');
}

function importCsv(sheetName) {

  try {
    const folder = DriveApp.getRootFolder();
    let file = folder.getFilesByType(MimeType.CSV).next();
    let data = Utilities.parseCsv(file.getBlob().getDataAsString());

    data.splice(0, 1); // Skip header

    var sheet = spreadsheet.insertSheet(sheetName);
    let startRow = sheet.getLastRow() + 1;
    let startCol = 1;
    let numRows = data.length;
    let numColumns = data[0].length;

    sheet.getRange(startRow, startCol, numRows, numColumns).setValues(data);
    file.setTrashed(true);
    return true;

  } catch {
    return false;
  }
}

function importOrders() {

  try {
    const sheetName = 'TMP_CSV';

    if (!importCsv(sheetName)) {

      SpreadsheetApp.getUi().alert('No CSV file found');
      return;
    }

    var sheet = spreadsheet.getSheetByName(sheetName);
    var input = sheet.getRange('B:F');
    var numRows = input.getNumRows();

    clearOrders();

    for (var i = 1; i <= numRows; i++) {

      let symbol_ = input.getCell(i, 2).getValue();
      if (symbol_ == '') { break; }

      let orderType = input.getCell(i, 1).getValue();
      let qty = input.getCell(i, 5).getValue();
      let price = input.getCell(i, 4).getValue();

      let symbols = spreadsheet.getRangeByName('Symbol');
      let rowIndex = symbols.createTextFinder(symbol_).findNext().getRowIndex() - 1;
      let order = spreadsheet.getRangeByName(orderType); // get 'Buy' or 'Sell' named range

      order.getCell(rowIndex, 1).setValue(qty);
      order.getCell(rowIndex, 2).setValue(price);
    }

    spreadsheet.deleteSheet(sheet);

  } catch (err) {
    logError(err.stack);
  }
}

function setBalance() {

  var cash = spreadsheet.getRangeByName('Cash');
  var cashValue = cash.getValue();
  var orderTotalValue = spreadsheet.getRangeByName('OrderTotal').getValue();

  if (cashValue == '') {
    cashValue = 0;
  }

  cash.setValue(cashValue + orderTotalValue);

  // Clear USDBRL
  var usdBrl_buy = spreadsheet.getRangeByName('USDBRL_Buy');
  var usdBrl_sell = spreadsheet.getRangeByName('USDBRL_Sell');

  //usdBrl_buy.setValue('');
  //usdBrl_sell.setValue('');
  usdBrl_buy.clear();
  usdBrl_sell.clear();
  usdBrl_buy.setBorder(true, true, true, true, true, true);
  usdBrl_sell.setBorder(true, true, true, true, true, true);
}

function fillOrders() {
  
  try {
    var usdBrl_buy = parseFloat(spreadsheet.getRangeByName('USDBRL_Buy').getValue().replace(',', '.'));
    var usdBrl_sell = parseFloat(spreadsheet.getRangeByName('USDBRL_Sell').getValue().replace(',', '.'));

    if (!(usdBrl_buy > 0 && usdBrl_sell > 0)) {

      SpreadsheetApp.getUi().alert('USDBRL is required');
      return;
    }

    buy(usdBrl_buy);
    sell(usdBrl_sell);
    setBalance();
    clearOrders();
    
  } catch (err) {
    logError(err.stack);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('*Order')
      .addItem('Set', 'setOrders')
      .addItem('Set Sell', 'setSell')
      .addItem('Set Buy', 'setBuy')
      .addItem('Set Prices', 'setPrices')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearOrders')
      .addItem('Import', 'importOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}
