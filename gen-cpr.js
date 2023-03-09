const ExcelJS = require('exceljs')
const { startCase, upperCase } = require('lodash')
const { boldCenter, verticalBold, right, bold, center, rightBold } = require('./styles')

async function main(){
  const cashflow = require('./resources/cashflow.json')
  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'Fenaka';
  workbook.lastModifiedBy = 'Fenaka';
  workbook.created = new Date(2023, 2, 5);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2022, 12, 27);
  workbook.properties.date1904 = true;
  workbook.calcProperties.fullCalcOnLoad = true;

  workbook.views = [{
    x: 0, y: 0, width: 10000, height: 20000,
    firstSheet: 0, activeTab: 1, visibility: 'visible'
  }]

  // FIRST HALF
  workbook.addWorksheet('First Half', {
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });
  worksheet = workbook.getWorksheet('First Half');
  createCashflowExcel('firstHalf', cashflow, worksheet);
  
  //SECOND HALF
  workbook.addWorksheet('Second Half', {
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });
  worksheet = workbook.getWorksheet('Second Half');
  createCashflowExcel('secHalf', cashflow, worksheet);

  await workbook.xlsx.writeFile('./excel-files/cash-position-report.xlsx');
}

main();

// FUNCTIONS
function createCashflowExcel(tableName, cashflow, worksheet){
  let allTotalIndex;
  let amountOfColumns;
  let isFirstHalf;
  let colSpan;
  let span;
  let dayStart;
  let rowCount = 6;

  worksheet.getColumn('A').width = 4;
  worksheet.getColumn('B').width = 15;

  if(tableName == 'firstHalf'){
    allTotalIndex = 0;
    amountOfColumns = 15;
    isFirstHalf = true;
    colSpan = 17;
    dayStart = 1;
  } else {
    allTotalIndex = 1;
    amountOfColumns = cashflow.generalBranchInfo.day - 15;
    isFirstHalf = false;
    colSpan = amountOfColumns + 3
    dayStart = 16;
  }
  span = String.fromCharCode(65 + colSpan);

  // PAGE SETUP
  const days = Array.from({length: amountOfColumns}, (_, i) => i + dayStart);
  const styles = {
    detailStyle: boldCenter,
    valueStyle: boldCenter,
    diff: false,
  }

  let grandTotal = null;
  if(!isFirstHalf){
    grandTotal = 'grandTotal';
  }

  dataFormatter(1, 'Details', days, 'Total', grandTotal, styles);
  rowCount++;

  // CASH IN
  dataFormatter(-1, 'CASH IN');
  rowCount++;

  const openingBalance = cashflow.cashflowObj[allTotalIndex].openingBalance;
  const balanceStyles = {
    detailStyle: boldCenter,
    valueStyle: rightBold,
    diff: false,
  }

  dataFormatter(1, 'openingBalance', openingBalance, null, null, balanceStyles);
  rowCount++;

  const cashInDiff = ['securityDeposit'];
  insertType('cashIn', allTotalIndex, cashInDiff);

  // CASH OUT
  dataFormatter(-1, 'CASH OUT');
  rowCount++;

  const cashOutDiff = ['cashOutOther'];
  insertType('cashOut', allTotalIndex, cashOutDiff);

  const closingBalance = cashflow.cashflowObj[allTotalIndex].closingBalance;
  dataFormatter(1, 'closingBalance', closingBalance, null, null, balanceStyles);
  rowCount++;

  // OTHER BALANCE
  dataFormatter(-1, 'OTHER BALANCE');
  rowCount++;

  const float = cashflow.cashflowObj[allTotalIndex].balance.float;
  const pettyCash = cashflow.cashflowObj[allTotalIndex].balance.pettyCash;
  const balanceTotal = cashflow.cashflowObj[allTotalIndex].allTotal.balance;
  const safeBalance = cashflow.cashflowObj[allTotalIndex].allTotal.safeBalance;
  const otherBalanceStyles = {
    detailStyle: center,
    valueStyle: right,
    diff: true,
  }

  const otherBalanceTotalStyles = {
    detailStyle: boldCenter,
    valueStyle: rightBold,
    diff: true,
  }

  const safeBalanceStyles = {
    detailStyle: boldCenter,
    valueStyle: rightBold,
    diff: false,
  }

  dataFormatter(2, 'float', float, null, null, otherBalanceStyles);
  rowCount++;

  dataFormatter(2, 'pettyCash', pettyCash, null, null, otherBalanceStyles);
  rowCount++;

  dataFormatter(2, 'Total', balanceTotal, null, null, otherBalanceTotalStyles);
  rowCount++;

  dataFormatter(1, 'totalSafeBalance', safeBalance, null, null, safeBalanceStyles);
  rowCount++;

  function insertType(cashflowTypeString, allTotalIndex, diff){
    const cashflowType = cashflow.cashflowObj[allTotalIndex][cashflowTypeString];
  
    for(const type in cashflowType){
      let colLength = 0;
      const found = diff.find(element => element == type);
  
      if(!found){
        for(const category in cashflowType[type]){
          const values = cashflowType[type][category];
          const categoryTotal = cashflow.bothHalfTotal[cashflowTypeString][type][category][allTotalIndex];
          const styles = {
            detailStyle: center,
            valueStyle: right,
            diff: false,
          }
  
          let grandTotal = null;
          if(!isFirstHalf){
            grandTotal = cashflow.grandTotal[cashflowTypeString][type][category][0];
          }
  
          dataFormatter(0, category, values, categoryTotal, grandTotal, styles);
          rowCount++;
          colLength++;
        }
  
        dataFormatter(colLength, type);
  
  
        const typeTotal = cashflow.cashflowObj[allTotalIndex].allTotal[type];
        const typeTotalsTotal = cashflow.bothHalfTotal.allTotal[type][allTotalIndex];
        const styles = {
          detailStyle: boldCenter,
          valueStyle: rightBold,
          diff: true,
        }
  
        let grandTotal = null;
        if(!isFirstHalf){
          grandTotal = cashflow.grandTotal.allTotal[type][0];
        }
  
        dataFormatter(2, 'Total', typeTotal, typeTotalsTotal, grandTotal, styles);
        rowCount++
      }
    }
  
    diff.forEach(type => {
      for(const category in cashflowType[type]){
        const values = cashflowType[type][category];
        const categoryTotal = cashflow.bothHalfTotal[cashflowTypeString][type][category][allTotalIndex];
        const styles = {
          detailStyle: center,
          valueStyle: right,
          diff: true,
        }
  
        let grandTotal = null;
        if(!isFirstHalf){
          grandTotal = cashflow.grandTotal[cashflowTypeString][type][category][0];
        }
  
        dataFormatter(2, category, values, categoryTotal, grandTotal, styles);
        rowCount++;
      }
    });
  
    const cashflowTypeTotal = cashflow.cashflowObj[allTotalIndex].allTotal[cashflowTypeString];
    const cashflowTypeTotalsTotal = cashflow.bothHalfTotal.allTotal[cashflowTypeString][allTotalIndex];
    const styles = {
      detailStyle: boldCenter,
      valueStyle: rightBold,
      diff: false,
    }
  
    let grandTotal = null;
    if(!isFirstHalf){
      grandTotal = cashflow.grandTotal.allTotal[cashflowTypeString][0];
    }
  
    dataFormatter(1, (cashflowTypeString + 'Total'), cashflowTypeTotal, cashflowTypeTotalsTotal, grandTotal, styles);
    rowCount++;
  }
  
  function dataFormatter(insertType, detailName, values, total, grandTotal, styles){
    if(values == null){
      if(insertType > 0){
        const colLength = insertType;
        const typeCell = worksheet.getCell(`A${rowCount - colLength}`);
        if((rowCount - 1) != (rowCount - colLength)){
          worksheet.mergeCells(`A${rowCount - colLength}:A${rowCount - 1}`);
        }
        Object.assign(typeCell, verticalBold, { value: startCase(detailName) });
      } else {
        const value = '  ' + detailName;
        worksheet.mergeCells(`A${rowCount}:${span}${rowCount}`);
        Object.assign(worksheet.getCell(`A${rowCount}`), bold, { value });
      }
    } else {
      let mergedValues = [detailName];
      mergedValues = mergedValues.concat(values);
  
      if(total != null){
        mergedValues.push(total);
      } else {
        mergedValues.push('');
      }
  
      if(!isFirstHalf){
        if(grandTotal != null){
          mergedValues.push(grandTotal);
        } else {
          mergedValues.push('');
        }
      }
  
      if(insertType == 0){
        const { detailStyle, valueStyle, diff } = styles;
        insertRow(mergedValues, detailStyle, valueStyle, diff);
      }
  
      if(insertType == 1){
        const { detailStyle, valueStyle, diff } = styles;
        worksheet.mergeCells(`A${rowCount}:B${rowCount}`);
        insertRow(mergedValues, detailStyle, valueStyle, diff);
      }
  
      if(insertType == 2){
        const { detailStyle, valueStyle, diff } = styles;
        mergedValues = [''].concat(mergedValues);
        insertRow(mergedValues, detailStyle, valueStyle, diff);
      }
    }
  }
  
  function insertRow(values, detailStyle, valueStyle, diff){
    let columnCount = 66;
  
    if(diff == true){
      columnCount -= 1
    }
  
    values.forEach(value => {
      if(typeof value == 'object'){
        if(value.amount == null){
          value = 0;
        } else {
          value = value.amount;
        }
      }
  
      if(typeof value == 'string'){
        value = startCase(value);
      }
  
      if(typeof value == 'number' && values[0] != 'Details'){
        value = (value / 100);
      }
  
      const cellName = String.fromCharCode(columnCount) + rowCount;
      const cell = worksheet.getCell(cellName);
  
      if(columnCount == 66){
        Object.assign(cell, detailStyle, { value });
      } else {
        if(typeof value == 'number' && values[0] != 'Details'){
          if(value == 0){
            Object.assign(cell, valueStyle, { value: '-  ' });
          } else {
            Object.assign(cell, valueStyle, { value }, { numFmt: '#,##0.00' });
          }
        } else {
          Object.assign(cell, valueStyle, { value });
        }
      }
      columnCount++;
    });
  }
}