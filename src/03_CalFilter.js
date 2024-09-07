// Calフィルター
function calFilter(e){
    const filCk = calSheet.getRange('A7').getValue(); // チェック状態
    const eCell = (e['range'].getRow() === 7 && e['range'].getColumn() === 1); // チェック
    if(eCell){
      let filterMBO = calSheet.getFilter();
      if(filterMBO !== null){
          calSheet.getFilter().remove();
      }
      let rule = SpreadsheetApp.newFilterCriteria()
              .setHiddenValues([])
              .build();
      if(filCk){
          rule = SpreadsheetApp.newFilterCriteria()
              .setHiddenValues([0])
              .build();
      }
      calSheet.getRange(beginRow_Cal,1,endRow_Cal-beginRow_Cal+1,1).createFilter()
          .setColumnFilterCriteria(1,rule);
    }
  }