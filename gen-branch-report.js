const ExcelJS = require('exceljs')
const branchData = require('./resources/branch-report.json')
const { startCase, upperCase } = require('lodash')
const { boldCenter, verticalBold, right, bold, center, rightBold } = require('./styles')

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
  const worksheet = workbook.getWorksheet('Branch Report');
  createBranchReport(worksheet);
  
  await workbook.xlsx.writeFile('./excel-files/branch-report.xlsx');
}

main();

// FUNCTIONS
function createBranchReport(worksheet){
  // PAGE SETUP
  const format = branchData.allRegionTotal[0].record;
}