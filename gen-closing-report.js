const ExcelJS = require('exceljs')
const closingData = require('./resources/closing-balance.json')
const { startCase, upperCase } = require('lodash')
const { boldCenter, verticalBold, right, bold, center, rightBold } = require('./styles')

let worksheet;
let rowCount = 6;

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

  // MAKE CLOSING REPORT
  workbook.addWorksheet('Closing Report', {
    pageSetup: { paperSize: 9, orientation:'landscape' },
  });
  worksheet = workbook.getWorksheet('Closing Report');
  createClosingReport();
  
  await workbook.xlsx.writeFile('./excel-files/closing-report.xlsx');
}

main();

// FUNCTIONS
function createClosingReport(){
  const amountOfDays = closingData.generalInfo.days;
  const days = Array.from({length: amountOfDays}, (_, i) => i + 1);

  insertRow('Branches', days, boldCenter);

  function insertRow(branchName, values, style){
    let columnCount = 65;
    let exceed = 0;
    values = [branchName].concat(values);

    values.forEach(value => {
      let cellName = String.fromCharCode(columnCount) + rowCount;
      console.log(columnCount, cellName);
      // const cell = worksheet.getCell(cellName);
      // Object.assign(cell, style, { value });
      columnCount++;
    });
  }
}