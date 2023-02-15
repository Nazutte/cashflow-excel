const ExcelJS = require('exceljs')
const cashflow = require('./resources/cashflow.json')
const { startCase } = require('lodash')
const { boldCenter, verticalBold, right, bold, center } = require('./styles')

let worksheet;
let rowCount;
let allTotalIndex;
let amountOfColumns;
let isFirstHalf;
let colSpan;
let span;
let dayStart;

async function main(){
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

  const sheet = workbook.addWorksheet('My Sheet', {
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });

  worksheet = workbook.getWorksheet('My Sheet');

  createCashflowExcel(2022, 2, 'firstHalf', cashflow);

  await workbook.xlsx.writeFile('./excel-files/excel.xlsx');
}

main();

// FUNCTIONS
function createCashflowExcel(year, month, tableName){
  rowCount = 6;

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
    amountOfColumns = new Date(year, month, 0).getDate() - 15;
    isFirstHalf = false;
    colSpan = amountOfColumns + 3
    dayStart = 16;
  }
  span = String.fromCharCode(65 + colSpan);

  // PAGE SETUP
  let details = ['Details'];
  const days = Array.from({length: amountOfColumns}, (_, i) => i + dayStart);
  details = details.concat(days);
  details.push('Total');

  if(!isFirstHalf){
    details.push('Grand Total');
  }

  worksheet.mergeCells(`A${rowCount}:B${rowCount}`);
  insertRow(details, boldCenter, boldCenter, false);
  rowCount++;

  // CASH IN
  worksheet.mergeCells(`A${rowCount}:${span}${rowCount}`);
  Object.assign(worksheet.getCell(`A${rowCount}`), bold, { value: '  CASH IN' });
  rowCount++;

  const cashInDiff = ['securityDeposit'];
  insertType('cashIn', allTotalIndex, cashInDiff);

  // CASH OUT
  worksheet.mergeCells(`A${rowCount}:${span}${rowCount}`);
  Object.assign(worksheet.getCell(`A${rowCount}`), bold, { value: '  CASH OUT' });
  rowCount++;

  const cashOutDiff = ['cashOutOther'];
  insertType('cashOut', allTotalIndex, cashOutDiff);
}

function insertType(cashflowTypeString, allTotalIndex, diff){
  const cashflowType = cashflow.cashflowObj[allTotalIndex][cashflowTypeString];

  for(const type in cashflowType){
    let colLength = 0;
    const found = diff.find(element => element == type);

    if(!found){
      for(const category in cashflowType[type]){
        const categoryTotal = cashflow.bothHalfTotal[cashflowTypeString][type][category][allTotalIndex];
        let arr = [category];
        arr = arr.concat(cashflowType[type][category]);
        arr.push(categoryTotal);
        insertRow(arr, center, right, false);
        rowCount++;
        colLength++;
      }

      const typeCell = worksheet.getCell(`A${rowCount - colLength}`);
      if((rowCount - 1) != (rowCount - colLength)){
        worksheet.mergeCells(`A${rowCount - colLength}:A${rowCount - 1}`);
      }
      Object.assign(typeCell, verticalBold, { value: startCase(type) });

      const typeTotal = cashflow.cashflowObj[allTotalIndex].allTotal[type];
      const typeTotalsTotal = cashflow.bothHalfTotal.allTotal[type][allTotalIndex];
      const arr = ['', 'Total'].concat(typeTotal, typeTotalsTotal);
      insertRow(arr, boldCenter, right, true);
      rowCount++
    }
  }

  diff.forEach(type => {
    for(const category in cashflowType[type]){
      const categoryTotal = cashflow.bothHalfTotal[cashflowTypeString][type][category][allTotalIndex];
      let arr = ['', category];
      arr = arr.concat(cashflowType[type][category]);
      arr.push(categoryTotal);

      insertRow(arr, center, right, true);
      rowCount++;
    }
  });
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
      value = (value / 100).toFixed(2);
      value += ' ';
    }

    const cellName = String.fromCharCode(columnCount) + rowCount;
    const cell = worksheet.getCell(cellName);

    if(columnCount == 66){
      Object.assign(cell, detailStyle, { value });
    } else {
      Object.assign(cell, valueStyle, { value });
    }
    columnCount++;
  });
}