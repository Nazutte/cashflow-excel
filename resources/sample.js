const ExcelJS = require('exceljs');

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
    headerFooter:{ firstHeader: "Hello Exceljs", firstFooter: "Hello World" },
    pageSetup:{ paperSize: 9, orientation:'landscape' },
  });

  const worksheet = workbook.getWorksheet('My Sheet');

  // worksheet.columns = [
  //   { header: 'Id', key: 'id', width: 10 },
  //   { header: 'Name', key: 'name', width: 32 },
  //   { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
  // ];

  // const idCol = worksheet.getColumn('id');
  // const nameCol = worksheet.getColumn('B');
  // const dobCol = worksheet.getColumn(3);

  // dobCol.header = 'Date of Birth';
  // dobCol.header = ['Date of Birth', 'A.K.A. D.O.B.'];
  // dobCol.key = 'dob';
  // dobCol.width = 15;

  worksheet.getRow(5).values = [null,null,null,1,2,3,4,5];

  const row = worksheet.getRow(5);
  const values = [1,2,3,4,5]

  insertRow(row, values, 3)

  worksheet.getCell('A1').border = {
    top: {style:'thick'},
    left: {style:'thick'},
    bottom: {style:'thick'},
    right: {style:'thick'}
  };

  createOuterBorder(worksheet, {startCell: 'C,9', endCell: 'D,10'});

  const colA = worksheet.getColumn('A');
  colA.width = 4;

  worksheet.mergeCells('A1:A4');

  const cellA1 = worksheet.getCell('A1');
  cellA1.value = 'hello world';
  cellA1.alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };
  cellA1.font = { size: 8 }

  await workbook.xlsx.writeFile('./excel-files/excel.xlsx');
}

main();

// FUNCTIONS
function insertRow(row, values, space){
  for(let count = 0; count < space; count++){
    values.unshift(null);
  }
  row.values = values;
}

function createOuterBorder(worksheet, {startCell, endCell}){
  const startCellCol = startCell.split(',')[0];
  const startCellRow = parseInt(startCell.split(',')[1]);
  const endCellCol = endCell.split(',')[0];
  const endCellRow = parseInt(endCell.split(',')[1]);

  // TOP AND BOTTOM BORDER
  console.log('TOP AND BOTTOM BORDER');
  let count = startCellCol.charCodeAt(0);
  for(count; count <= endCellCol.charCodeAt(0); count++){
    console.log('top: ' + String.fromCharCode(count) + startCellRow);
    console.log('bottom: ' + String.fromCharCode(count) + endCellRow);

    worksheet.getCell(String.fromCharCode(count) + startCellRow).border = {
      top: {style:'thin'},
    };

    worksheet.getCell(String.fromCharCode(count) + endCellRow).border = {
      bottom: {style:'thin'},
    };
  }

  // LEFT AND RIGHT BORDER
  console.log('\nLEFT AND RIGHT BORDER');
  count = startCellRow;
  for(count; count <= endCellRow; count++){
    console.log('left: ' + startCellCol + count);
    console.log('right: ' + endCellCol + count);

    if(count == startCellRow){
      worksheet.getCell(startCellCol + count).border = {
        left: {style:'thin'},
        top: {style:'thin'},
      };

      worksheet.getCell(endCellCol + count).border = {
        top: {style:'thin'},
        right: {style:'thin'},
      };
    } else if(count == endCellRow){
      worksheet.getCell(startCellCol + count).border = {
        left: {style:'thin'},
        bottom: {style:'thin'},
      };

      worksheet.getCell(endCellCol + count).border = {
        bottom: {style:'thin'},
        right: {style:'thin'},
      };
    } else {
      worksheet.getCell(startCellCol + count).border = {
        left: {style:'thin'},
      };
  
      worksheet.getCell(endCellCol + count).border = {
        right: {style:'thin'},
      };
    }
  }
}