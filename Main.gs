let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function addOrder(orderType, usdBrl) {

  //#region Set order
  let orderRange = spreadsheet.getRangeByName(orderType);
  let positions = spreadsheet.getRangeByName('Position');
  let orderData = [];
  let isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  for (let i=1; i<=numRows; i++) {

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
        var newUsdAc = usdAc;
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

  for (let i=0; i<orderData.length; i++) {
    
    let orderRow = orderData[i];
    let orderIndex = orderRow[numCol-1];
    let qty = orderRow[numCol-2];
    let avgCost = orderRow[numCol-4];
    let usdAc = orderRow[numCol-3];
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
  spreadsheet.getRangeByName('PTAX_Buy').setValue('');
  spreadsheet.getRangeByName('PTAX_Sell').setValue('');
}

function clearPrices() {

  spreadsheet.getRangeByName('SellPrice').setValue('');
  spreadsheet.getRangeByName('BuyPrice').setValue('');
}

function setOrders() {

  try {

    let targetQuantityRange = spreadsheet.getRangeByName('TargetQuantity');
    let priceRange = spreadsheet.getRangeByName('Price');
    let buyRange = spreadsheet.getRangeByName('Buy');
    let sellRange = spreadsheet.getRangeByName('Sell');
    const numRows = targetQuantityRange.getNumRows();
    const buyOrders = new Array(numRows).fill([undefined,undefined]);
    const sellOrders = new Array(numRows).fill([undefined,undefined]);

    for (let i=1; i<=numRows; i++) {

      let targetQuantityCell = targetQuantityRange.getCell(i, 1);
      let qty = targetQuantityCell.getValue();
      let price = priceRange.getCell(i,1).getValue();
      
      if (targetQuantityCell.getBackgroundColor() == '#ff9900') { // orange

        sellOrders[i-1] = [qty*-1, price];

      } else if (targetQuantityCell.getBackgroundColor() == '#34a853') { // green

        buyOrders[i-1] = [qty, price];
      }
    }

    buyRange.setValues(buyOrders);
    sellRange.setValues(sellOrders);

  } catch (err) {

    Util.logError(err.stack);
  }
}

function setPrices() {

  try {

    let priceRange = spreadsheet.getRangeByName('Price');
    let buyRange = spreadsheet.getRangeByName('Buy');
    let sellRange = spreadsheet.getRangeByName('Sell');
    const numRows = priceRange.getNumRows();

    for (let i=1; i<=numRows; i++) {

      let price = priceRange.getCell(i, 1).getValue();
      let buyQty = buyRange.getCell(i, 1).getValue();
      let sellQty = sellRange.getCell(i, 1).getValue();
      
      if (buyQty > 0) { buyRange.getCell(i, 2).setValue(price); }
      if (sellQty > 0) { sellRange.getCell(i, 2).setValue(price); }
    }

  } catch (err) {

    Util.logError(err.stack);
  }
}

function setBalance() {

  let cash = spreadsheet.getRangeByName('Cash');
  let cashValue = cash.getValue();
  let orderTotalValue = spreadsheet.getRangeByName('OrderTotal').getValue();

  if (cashValue == '') { cashValue = 0; }

  cash.setValue(cashValue + orderTotalValue);
  clearOrders();
}

function fillOrders() {
  
  try {

    let ptax_buy = parseFloat(spreadsheet.getRangeByName('PTAX_Buy').getValue().replace(',', '.'));
    let ptax_sell = parseFloat(spreadsheet.getRangeByName('PTAX_Sell').getValue().replace(',', '.'));

    if (!(ptax_buy > 0 && ptax_sell > 0)) {
      
      SpreadsheetApp.getUi().alert('PTAX input is required');
      return;
    }

    addOrder('Buy', ptax_buy);
    addOrder('Sell', ptax_sell);
    setBalance();
    
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
      .addItem('Set Prices', 'setPrices')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}
