const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();

async function main(){
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
    headerFooter:{ firstHeader: "Hello Exceljs", firstFooter: "Hello World" },
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });

  await workbook.xlsx.writeFile('./excel-files/excel.xlsx');
}

main();