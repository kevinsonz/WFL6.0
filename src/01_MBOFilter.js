// MBOフィルターのモード切替え
function mboFilter(e){
  const eValue1 = statusMBO === '今' && e['value'] === 'TRUE' && e['range'].getRow() === 2 && e['range'].getColumn() === 1;;
  const eValue2 = statusMBO === '全' && e['value'] === '全';
  const eValue3 = statusMBO === '今' && e['value'] === '今' && e['oldValue'] === '全';
  const runFlag = endCol_MBO === mboCol;
  if(runFlag && (eValue1 || eValue2 || eValue3)){
    let filterMBO = mboSheet.getFilter();
    if(filterMBO !== null){
      mboSheet.getFilter().remove();
    }
    let rule = SpreadsheetApp.newFilterCriteria()
          .setHiddenValues([])
          .build();
    if(eValue1 || eValue3){
        rule = SpreadsheetApp.newFilterCriteria()
          .setHiddenValues(["Hidden"])
          .build();
    }
    mboSheet.getRange(beginRow_MBO-1,1,mboRow+1,endCol_MBO).createFilter()
      .setColumnFilterCriteria(hiddenRowNum,rule);
    mboSheet.getRange('A2').setValue(false);
  }
}