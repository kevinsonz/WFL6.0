// MBOフィルターのモード切替え
function mboFilter(e){
  const eValue1 = statusMBO === '今' && e['value'] === 'TRUE';
  const eValue2 = statusMBO === '全' && e['value'] === '全';
  const runFlag = endCol_MBO === mboCol;
  if(runFlag && (eValue1 || eValue2)){
    let filterMBO = mboSheet.getFilter();
    if(filterMBO !== null){
      mboSheet.getFilter().remove();
    }
    let rule = SpreadsheetApp.newFilterCriteria()
          .setHiddenValues([])
          .build();
    if(eValue1){
        rule = SpreadsheetApp.newFilterCriteria()
          .setHiddenValues(["Hidden"])
          .build();
    }
    mboSheet.getRange(beginRow_MBO-1,1,mboRow+1,endCol_MBO).createFilter()
      .setColumnFilterCriteria(2,rule);
    mboSheet.getRange('A2').setValue(false);
  }
}