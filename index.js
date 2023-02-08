function makeCashflowExcel(){
  let allTotalIndex;
  let amountOfColumns;
  let isFisrtHalf;
  let colSpan;
  if(tableName == 'firstHalf'){
    allTotalIndex = 0;
    amountOfColumns = 15;
    isFisrtHalf = true;
    colSpan = 17;
  } else {
    allTotalIndex = 1;
    amountOfColumns = new Date(year, month, 0).getDate() - 15;
    isFisrtHalf = false;
    colSpan = amountOfColumns + 3
  }

  // let details = ['details'];
  // if(isFisrtHalf){
  //   for(let count = ){

  //   }
  // }
}