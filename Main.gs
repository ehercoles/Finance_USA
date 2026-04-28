var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function addOrder(orderType, usdBrl) {

  //#region Set order
  var orderRange = spreadsheet.getRangeByName(orderType);
  var positions = spreadsheet.getRangeByName('Position');
  var orderData = [];
  var isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  for (let i = 1; i <= numRows; i++) {

    let orderQty = parseInt(orderRange.getCell(i, 1).getValue());
    let orderPx = parseFloat(orderRange.getCell(i, 2).getValue());
    
    if (orderQty > 0 && orderPx > 0) {

      let qty = parseInt(0 + positions.getCell(i, 2).getValue());
      let avgCost = parseFloat(0 + positions.getCell(i, 3).getValue());
      let usdAc = parseFloat(0 + positions.getCell(i, 4).getValue());
      
      if (avgCost == 0) { avgCost = orderPx; }
      
      if(usdAc == 0) { usdAc = usdBrl; }

      if (isBuy) {

        var newQty = qty + orderQty;
        var newAvgCost = ((qty * avgCost) + (orderQty * orderPx)) / (qty + orderQty);
        var newUsdAc = ((qty * usdAc) + (orderQty * usdBrl)) / (qty + orderQty);

      } else {

        var newQty = qty - orderQty;
        var newAvgCost = avgCost;
        var newUsdAc = usdAc
      }
      
      orderData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        qty,
        avgCost,
        usdAc,
        orderQty,
        orderPx,
        usdBrl, // Sell only // Sell sheet range limit
        newAvgCost,
        newUsdAc, // Buy sheet range limit
        newQty,
        i+1]); // order index
    }
  }
  
  if (orderData.length == 0) return;
  //#endregion

  //#region Set position
  let numCol = orderData[0].length;

  for (let i = 0; i < orderData.length; i++) {
    
    let order = orderData[i];
    let orderIndex = orderData[i][numCol-1];
    let qty = order[numCol-2];
    let avgCost = order[numCol-4];
    let usdAc = order[numCol-3];
    let values = [[qty, avgCost, usdAc]];
    let rangeStr = Utilities.formatString("B%s:D%s", orderIndex, orderIndex);
    let range = spreadsheet.getRange(rangeStr);

    if (qty == 0) {

      range.setValue("");

    } else {
      
      range.setValues(values);
    }
  }
  //#endregion
  
  //#region Add order
  if (isBuy) {
    
    Util.spliceColumn(orderData, 7, 1); // remove column USDBRL
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
  //#endregion
}

function clearOrders() {

  spreadsheet.getRangeByName('Sell').setValue('');
  spreadsheet.getRangeByName('Buy').setValue('');
  clearPtaxInput();
}

function clearPrices() {

  spreadsheet.getRangeByName('SellPrice').setValue('');
  spreadsheet.getRangeByName('BuyPrice').setValue('');
}

function clearPtaxInput() {

  spreadsheet.getRangeByName('PTAX_Buy').setValue('');
  spreadsheet.getRangeByName('PTAX_Sell').setValue('');
}

function setOrders(mode) {

  try {

    var targetQuantities = spreadsheet.getRangeByName('TargetQuantity');
    var prices = spreadsheet.getRangeByName('Price');
    var sellRange = spreadsheet.getRangeByName('Sell');
    var buyRange = spreadsheet.getRangeByName('Buy');
    const numRows = targetQuantities.getNumRows();

    clearOrders();
    
    for (var i = 1; i <= numRows; i++) {

      var targetQuantityCell = targetQuantities.getCell(i, 1);
      var qty = targetQuantityCell.getValue();
      var price = prices.getCell(i, 1).getValue();
      
      if ((!mode || mode == 'sell') && targetQuantityCell.getBackgroundColor() == '#ff9900') { // orange

        var quantityCell = sellRange.getCell(i, 1);
        var priceCell = sellRange.getCell(i, 2);

        quantityCell.setValue(qty * -1);
        priceCell.setValue(price);

      } else if ((!mode || mode == 'buy') && targetQuantityCell.getBackgroundColor() == '#34a853') { // green

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

function setBuy() {

  setOrders('buy');
}

function setSell() {

  setOrders('sell');
}

// TODO: fix parse result: price with 2 decimal places
function importOrders() {

  var data = [];

  try {

    const folder = DriveApp.getRootFolder();
    let file = folder.getFilesByType(MimeType.CSV).next();

    data = Utilities.parseCsv(file.getBlob().getDataAsString());

    //file.setTrashed(true);
    //Logger.log(data);

  } catch {

    SpreadsheetApp.getUi().alert('No CSV file found');
  }

  try {

    const numRows = data.length;
    const symbolIndex = 0;
    const orderTypeIndex = 4;
    const qtyIndex = 5;
    const priceIndex = 8;
    
    for (var i=1; i<numRows; i++) {

      let symbol_ = data[i][symbolIndex];
      if (symbol_ == '') { break; }

      let orderType = data[i][orderTypeIndex];
      let qty = data[i][qtyIndex];
      let price = data[i][priceIndex];

      qty = qty.split(" ", 1)[0];

      let symbols = spreadsheet.getRangeByName('Symbol');
      let rowIndex = symbols.createTextFinder(symbol_).findNext().getRowIndex()-1;
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

  if (cashValue == '') { cashValue = 0; }

  cash.setValue(cashValue + orderTotalValue);
  clearPtaxInput();
}

function fillOrders() {
  
  try {

    var ptax_buy = parseFloat(spreadsheet.getRangeByName('PTAX_Buy').getValue().replace(',', '.'));
    var ptax_sell = parseFloat(spreadsheet.getRangeByName('PTAX_Sell').getValue().replace(',', '.'));

    if (!(ptax_buy > 0 && ptax_sell > 0)) {
      
      SpreadsheetApp.getUi().alert('PTAX input is required');
      return;
    }

    addOrder('Buy', ptax_buy);
    addOrder('Sell', ptax_sell);
    setBalance();
    clearOrders();
    
  } catch (err) {

    Util.logError(err.stack);
  }
}

function incrementThreshold() {

    const threshold = spreadsheet.getRangeByName('Threshold');
    const value = threshold.getValue();
    const rule = threshold.getDataValidation();
    
    if (rule == null) return;

    //const criteria = rule.getCriteriaType();
    const args = rule.getCriteriaValues();
    const validationValues = args[0].getValues().filter(Number);
    const maxValue = validationValues[validationValues.length - 1];

    //Logger.log(validationValues);

    if (value < maxValue) {

      threshold.setValue(value + 1);
    }
}

function decrementThreshold() {
  
    const threshold = spreadsheet.getRangeByName('Threshold');
    const value = threshold.getValue();

    if (value >= 1) {

      threshold.setValue(value - 1);
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
      //.addItem('Import', 'importOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}
