// MBOフィルターのモード切替え
function mboFilter(e){
  const imaCk = mboSheet.getRange('A3').getValue(); // 今ココ状態
  const nenCk = mboSheet.getRange('E1').getValue(); // 年度モード状態
  const eCell1 = (e['range'].getRow() === 3 && e['range'].getColumn() === 1); // 今ココ
  const eCell2 = (e['range'].getRow() === 1 && e['range'].getColumn() === 5); // 年度
  const eCell3_1 = (e['range'].getRow() === 2 && e['range'].getColumn() === 4); // 開始（数字）
  const eCell3_2 = (e['range'].getRow() === 2 && e['range'].getColumn() === 5); // 開始（単位）
  const eCell4_1 = (e['range'].getRow() === 3 && e['range'].getColumn() === 4); // 終了（数字）
  const eCell4_2 = (e['range'].getRow() === 3 && e['range'].getColumn() === 5); // 終了（単位）
  const eCell = eCell1 || eCell2 || eCell3_1 || eCell3_2 || eCell4_1 || eCell4_2;
  const eCell234 = eCell2 || eCell3_1 || eCell3_2 || eCell4_1 || eCell4_2;
  const eCell34 = eCell3_1 || eCell3_2 || eCell4_1 || eCell4_2;
  const colCheck = endCol_MBO === mboCol;
  const runFlag = eCell && colCheck;
  if(!((!imaCk && eCell234) || (nenCk && eCell34))){
    if(runFlag){
      let filterMBO = mboSheet.getFilter();
      if(filterMBO !== null){
        mboSheet.getFilter().remove();
      }
      let rule = SpreadsheetApp.newFilterCriteria()
            .setHiddenValues([])
            .build();
      if(imaCk){
          rule = SpreadsheetApp.newFilterCriteria()
            .setHiddenValues(["Hidden"])
            .build();
      }
      mboSheet.getRange(beginRow_MBO-1,1,mboRow+1,endCol_MBO).createFilter()
        .setColumnFilterCriteria(hiddenRowNum,rule);
    }
  }
}