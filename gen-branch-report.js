const ExcelJS = require('exceljs')
const branchData = require('./resources/branch-report.json')
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

  // MAKE REGION REPORT
  workbook.addWorksheet('Branch Report', {
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });
  worksheet = workbook.getWorksheet('Branch Report');
  createBranchReport();
  
  await workbook.xlsx.writeFile('./excel-files/branch-report.xlsx');
}

main();

// FUNCTIONS
function createBranchReport(){
  // PAGE SETUP
  const format = branchData.branch[0].record;
  worksheet.getColumn('A').width = 20;
  createFormat(format);
  rowCount += 2;

  branchData.branch.push({ name: 'Total', record: branchData.grandTotal});

  branchData.branch.forEach(branch => {
    insertBranch(branch);
    rowCount++;
  });
}

function insertBranch(branch){
  let columnCount = 65;
  insertCell(branch.name, boldCenter);
  insertCell(branch.record.openingBalance, rightBold);

  const { cashInRecord, cashInRecordTotal, cashOutRecord, cashOutRecordTotal, other } = branch.record;

  for(const type in cashInRecord){
    insertCell(cashInRecord[type], right);
  }
  insertCell(cashInRecordTotal, rightBold);

  for(const type in cashOutRecord){
    insertCell(cashOutRecord[type], right);
  }
  insertCell(cashOutRecordTotal, rightBold);

  insertCell(branch.record.closingBalance, rightBold);

  for(const type in other){
    insertCell(other[type], right);
  }

  function insertCell(value, style){
    let cellName = String.fromCharCode(columnCount) + rowCount;

    if(typeof value == 'number'){
      value = value/100;
      Object.assign(worksheet.getCell(cellName), style, { value }, { numFmt: '#,##0.00' });
    } else {
      Object.assign(worksheet.getCell(cellName), style, { value });
    }

    columnCount++;
  }
}

function createFormat(format){
  let columnCount = 65;
  worksheet.getRow(rowCount + 1).height = 30;

  insert(0, 'Branch')
  insert(0, 'Opening Balance')

  insertType('Cash In', format.cashInRecord);
  insertType('Cash Out', format.cashOutRecord);

  insert(0, 'Closing Balance')

  insertType('Other Balance', format.other, true);

  function insertType(name, type, diff){
    let currentCol = columnCount;

    for(const category in type){
      insert(1, startCase(category));
    }
    if(diff != true){
      insert(1, 'Total');
    }

    worksheet.mergeCells(String.fromCharCode(currentCol) + rowCount + ':' + String.fromCharCode(columnCount - 1) + (rowCount));
    Object.assign(worksheet.getCell(String.fromCharCode(currentCol) + rowCount), boldCenter, {value: name});
  }

  function insert(type, value){
    if(type == 0){
      const col = String.fromCharCode(columnCount);
      worksheet.mergeCells(col + rowCount + ':' + col + (rowCount + 1));
    }

    let cell = worksheet.getCell(String.fromCharCode(columnCount) + (rowCount + 1));
    Object.assign(cell, boldCenter, {value})

    columnCount++;
  }
}