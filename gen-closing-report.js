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

  closingData.branchArr.forEach(branch => {
    const { name, dailyRecord } = branch;
    insertRow(name, dailyRecord, center);
  });

  insertRow('Total', closingData.dailyTotal, boldCenter);

  function insertRow(branchName, values, style){
    let columnCount = [65];
    values = [branchName].concat(values);

    values.forEach(value => {
      let cellName = '';
      columnCount.forEach(charCode => {
        cellName += String.fromCharCode(charCode);
      });
      cellName += rowCount;

      const cell = worksheet.getCell(cellName);

      if(branchName != 'Branches' && typeof value == 'number'){
        value = value/100;
        Object.assign(cell, style, { value }, { numFmt: '#,##0.00' }, {alignment: {horizontal: 'right', vertical: 'middle'}});
      } else {
        Object.assign(cell, style, { value });
      }
      
      for(let j = columnCount.length; j > 0; j--){
        if(columnCount[j - 1] != 90){
          columnCount[j - 1]++;
          break;
        } else {
          columnCount[j - 1] = 65;
          if((j - 1) <= 0){
            columnCount.push(65);
            break;
          }
        }
      }
    });
    rowCount++;
  }
}

// function add() {
//   let num = 0
//   return function B() {
//     num++
//     console.log(num)
//   }
// }

// const increment = add()