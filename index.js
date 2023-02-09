const ExcelJS = require('exceljs');

async function main(){
  const workbook = new ExcelJS.Workbook();
  const cashflow = require('./resources/cashflow.json');

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

  const worksheet = workbook.getWorksheet('My Sheet');

  createCashflowExcel(2022, 2, 'firstHalf', worksheet);

  await workbook.xlsx.writeFile('./excel-files/excel.xlsx');
}

main();

// FUNCTIONS
function createCashflowExcel(year, month, tableName, worksheet){
  let rowCount = 5;
  let allTotalIndex;
  let amountOfColumns;
  let isFisrtHalf;
  let colSpan;
  let dayStart;
  if(tableName == 'firstHalf'){
    allTotalIndex = 0;
    amountOfColumns = 15;
    isFisrtHalf = true;
    colSpan = 17;
    dayStart = 1;
  } else {
    allTotalIndex = 1;
    amountOfColumns = new Date(year, month, 0).getDate() - 15;
    isFisrtHalf = false;
    colSpan = amountOfColumns + 3
    dayStart = 16;
  }

  let details = ['Details'];
  const days = Array.from({length: amountOfColumns}, (_, i) => i + dayStart);
  details = details.concat(days);
  details.push('Total');

  if(!isFisrtHalf){
    details.push('Grand Total');
  }

  console.log(details);
  const headerRow = worksheet.getRow(rowCount);
  insertRow(headerRow, rowCount, details);
}

function insertRow(row, rowCount, values, detailStyle, valueStyle){
  
}