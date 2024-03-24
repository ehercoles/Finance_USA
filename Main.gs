var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function addOrder(orderType, usdBrl) {

  var orderRange = spreadsheet.getRangeByName(orderType);
  var positions = spreadsheet.getRangeByName('Position');
  var orderData = [];
  var isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  for (let i = 1; i <= numRows; i++) {

    let orderQty = parseInt(orderRange.getCell(i, 1).getValue());
    let orderAp = parseFloat(orderRange.getCell(i, 2).getValue());
    
    if (orderQty > 0 && orderAp > 0) {
      let qty = parseInt(0 + positions.getCell(i, 2).getValue());
      let ap = parseFloat(0 + positions.getCell(i, 3).getValue());
      let usdAp = parseFloat(0 + positions.getCell(i, 4).getValue());
      
      if (ap == 0) {
        ap = orderAp;
      }
      
      if(usdAp == 0) {
        usdAp = usdBrl;
      }

      if (isBuy) {
        var newQty = qty + orderQty;
        var newAp = ((qty * ap) + (orderQty * orderAp)) / (qty + orderQty);
        var newUsdAp = ((qty * usdAp) + (orderQty * usdBrl)) / (qty + orderQty);

      } else {
        var newQty = qty - orderQty;
        var newAp = ap;
        var newUsdAp = usdAp
      }
      
      orderData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        qty,
        ap,
        usdAp,
        orderQty,
        orderAp,
        usdBrl, // Sheet "Sell" only // Sheet "Sell" limit
        newAp,
        newUsdAp, // Sheet "Buy" limit
        newQty,
        i]); // order index
    }
  }
  
  if (orderData.length == 0) return;

  //#region Set position
  let numCol = orderData[0].length;

  for (let i = 0; i < orderData.length; i++) {
    
    let orderIndex = orderData[i][numCol-1];

    positions.getCell(orderIndex, 2).setValue(orderData[i][numCol-2]); // Qty
    positions.getCell(orderIndex, 3).setValue(orderData[i][numCol-4]); // AP
    positions.getCell(orderIndex, 4).setValue(orderData[i][numCol-3]); // USD AP
  }
  //#endregion
  
  //#region Add order
  if (isBuy) {
    Util.spliceColumn(orderData, 7, 1); // rem USDBRL
    orderData = Util.sliceColumn(orderData, 0, -2);
    
  } else {
    orderData = Util.sliceColumn(orderData, 0, -4);
  }

  orderData.sort(Util.sortFunction);
  //Logger.log(orderData);

  let orderSheet = spreadsheet.getSheetByName(orderType);
  const rowStart = 3;
  const rowCount = orderData.length;
  
  numCol = orderData[0].length;
  orderSheet.insertRowsAfter(rowStart-1, rowCount);
  orderSheet.getRange(rowStart, 1, rowCount, numCol).setValues(orderData);
  
  // Copy formula to the new cells
  if (!isBuy) {
    const colStart = numCol + 1;
    const colCount = 5;

    let fromRange = orderSheet.getRange(rowStart-1, colStart, 1, colCount);
    fromRange.copyTo(orderSheet.getRange(rowStart, colStart, rowCount, colCount), {contentsOnly:false});
  }
  //#endregion
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
    Util.logError(err.stack);
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
    Util.logError(err.stack);
  }
}

function setSell() {
  setOrders('sell');
}

function setBuy() {
  setOrders('buy');
}

function importOrders() {

  var data = [];

  try {
    const folder = DriveApp.getRootFolder();
    let file = folder.getFilesByType(MimeType.CSV).next();

    data = Utilities.parseCsv(file.getBlob().getDataAsString());
    file.setTrashed(true);
    //Logger.log(data);

  } catch {
    SpreadsheetApp.getUi().alert('No CSV file found');
  }

  try {
    const numRows = data.length;
    
    for (var i = 1; i < numRows; i++) {

      let symbol_ = data[i][2];
      if (symbol_ == '') { break; }

      let orderType = data[i][1];
      let qty = data[i][5];
      let price = data[i][4];

      let symbols = spreadsheet.getRangeByName('Symbol');
      let rowIndex = symbols.createTextFinder(symbol_).findNext().getRowIndex() - 1;
      let order = spreadsheet.getRangeByName(orderType); // get named range 'Buy' or 'Sell'

      order.getCell(rowIndex, 1).setValue(qty);
      order.getCell(rowIndex, 2).setValue(price);
    }

  } catch (err) {
    Util.logError(err.stack);
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

  usdBrl_buy.setValue('');
  usdBrl_sell.setValue('');
}

function fillOrders() {
  
  try {
    var usdBrl_buy = parseFloat(spreadsheet.getRangeByName('USDBRL_Buy').getValue().replace(',', '.'));
    var usdBrl_sell = parseFloat(spreadsheet.getRangeByName('USDBRL_Sell').getValue().replace(',', '.'));

    if (!(usdBrl_buy > 0 && usdBrl_sell > 0)) {
      SpreadsheetApp.getUi().alert('USDBRL is required');
      return;
    }

    addOrder('Buy', usdBrl_buy);
    addOrder('Sell', usdBrl_sell);
    //setBalance();
    //clearOrders();
    
  } catch (err) {
    Util.logError(err.stack);
  }
}

function incrementThreshold() {

    const range = spreadsheet.getRangeByName('Threshold');
    const value = range.getValue();
    const rule = range.getDataValidation();
    
    if (rule == null) return;

    //const criteria = rule.getCriteriaType();
    const args = rule.getCriteriaValues();
    const validationValues = args[0].getValues().filter(Number);
    const maxValue = validationValues[validationValues.length - 1];

    //Logger.log(validationValues);

    if (value < maxValue) {
      range.setValue(value + 1);
    }
}

function decrementThreshold() {
  
    const range = spreadsheet.getRangeByName('Threshold');
    const value = range.getValue();

    if (value > 1) {
      range.setValue(value - 1);
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
