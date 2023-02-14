const ExcelJS = require('exceljs')
const cashflow = require('./resources/cashflow.json');
const { boldCenter } = require('./styles')

let worksheet;
let rowCount;
let allTotalIndex;
let amountOfColumns;
let isFirstHalf;
let colSpan;
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
  rowCount = 5;

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

  let details = ['Details'];
  const days = Array.from({length: amountOfColumns}, (_, i) => i + dayStart);
  details = details.concat(days);
  details.push('Total');

  if(!isFirstHalf){
    details.push('Grand Total');
  }

  worksheet.mergeCells(`A${rowCount}:B${rowCount}`);
  insertRow(details, boldCenter);
  rowCount++;

  const diff = ['securityDeposit'];
  insertType('cashIn', allTotalIndex, diff);
}

function insertType(cashflowTypeString, allTotalIndex, diff){
  const cashflowType = cashflow.cashflowObj[allTotalIndex][cashflowTypeString];

  for(const type in cashflowType){
    let colLength = 0;
    const found = diff.find(element => element == type);

    if(!found){
      for(const category in cashflowType[type]){
        let arr = [category];
        arr = arr.concat(cashflowType[type][category]);
        insertRow(arr, boldCenter);
        rowCount++;
        colLength++;
      }
    }
    
    // console.log(type);
    // console.log('Column Length: ' + colLength);
    // console.log(rowCount - colLength);
    // console.log(rowCount - 1);
    // console.log('\n');

    if(!found){
      if((rowCount - 1) != (rowCount - colLength)){
        worksheet.mergeCells(`A${rowCount - colLength}:A${rowCount - 1}`);
      }
    }
  }
}

function insertRow(values, style){
  let columnCount = 66;

  values.forEach(value => {
    if(typeof value == 'object'){
      if(value.amount == null){
        value = 0;
      } else {
        value = value.amount;
      }
    }
    const cellName = String.fromCharCode(columnCount) + rowCount;
    const cell = worksheet.getCell(cellName);

    Object.assign(cell, style, { value });
    columnCount++;
  });
}