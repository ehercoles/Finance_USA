let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {

  SpreadsheetApp.getUi()
      .createMenu('*Order')
      .addItem('Set', 'setOrders')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}

function addOrder(orderType, usdBrl) {

  let portfolioSheet = spreadsheet.getSheetByName("Portfolio");
  let orderRange = spreadsheet.getRangeByName(orderType);
  let positionRange = spreadsheet.getRangeByName('Position');
  let orderData = [];
  const isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  //#region Set order
  for (let i=1; i<=numRows; i++) {

    let orderSymbol = positionRange.getCell(i, 1).getValue();
    let orderQty = parseInt(orderRange.getCell(i, 1).getValue());
    let orderPrice = parseFloat(orderRange.getCell(i, 2).getValue());
    
    if (orderQty > 0 && orderPrice > 0) {

      let qty = parseInt(0 + positionRange.getCell(i, 2).getValue());
      let avgCost = parseFloat(0 + positionRange.getCell(i, 3).getValue());
      let usdAc = parseFloat(0 + positionRange.getCell(i, 4).getValue());
      
      if (isBuy) {
        
        if (qty == 0) {

          avgCost = orderPrice;
          usdAc = usdBrl;
        }

        var newQty = qty + orderQty;
        var newAvgCost = ((qty * avgCost) + (orderQty * orderPrice)) / (qty + orderQty);
        var newUsdAc = ((qty * usdAc) + (orderQty * usdBrl)) / (qty + orderQty);

      } else {

        var newQty = qty - orderQty;
        var newAvgCost = avgCost;
        var newUsdAc = usdAc;
      }
      
      orderData.push([
        new Date(),
        orderSymbol,
        qty,
        avgCost,
        usdAc,
        orderQty,
        orderPrice,
        usdBrl, // Sheet "Sell" limit
        newQty,
        newAvgCost,
        newUsdAc, // Sheet "Buy" limit
        i+2]); // Order index
    }
  }

  if (orderData.length == 0) return;
  //#endregion

  //#region Set position
  let numCol = orderData[0].length;

  for (let i=0; i<orderData.length; i++) {
    
    let orderRow = orderData[i];
    let orderIndex = orderRow[numCol-1];
    let qty = orderRow[8];
    let avgCost = orderRow[9];
    let usdAc = orderRow[10];
    let values = [[qty, avgCost, usdAc]];
    let positionRow = portfolioSheet.getRange(orderIndex, 2, 1, 2);

    if (qty == 0) {

      positionRow.setValue("");

    } else {
      
      positionRow.setValues(values);
    }
  }
  //#endregion
  
  //#region Add order
  if (isBuy) {
    
    orderData = Util.sliceColumn(orderData, 0, -1);
    
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
  spreadsheet.getRangeByName('UsdBrl_Buy').setValue('');
  spreadsheet.getRangeByName('UsdBrl_Sell').setValue('');
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

//#region USA version: do not replace nor replicate the code below
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

    let usdBrl_buy = parseFloat(spreadsheet.getRangeByName('UsdBrl_Buy').getValue().replace(',', '.'));
    let usdBrl_sell = parseFloat(spreadsheet.getRangeByName('UsdBrl_Sell').getValue().replace(',', '.'));

    if (!(usdBrl_buy > 0 && usdBrl_sell > 0)) {
      
      SpreadsheetApp.getUi().alert('PTAX input is required');
      return;
    }

    addOrder('Buy', usdBrl_buy);
    addOrder('Sell', usdBrl_sell);
    setBalance();
    
  } catch (err) {

    Util.logError(err.stack);
  }
}
//#endregion
