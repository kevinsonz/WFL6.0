// MBOフィルターのモード切替え
function mboFilter(e){
  const imaCk = mboSheet.getRange('A3').getValue(); // 今ココ状態
  const nenCk = mboSheet.getRange('E1').getValue(); // 年度モード状態
  const zenCk = mboSheet.getRange('E2').getValue(); // 前月モード状態
  const jigCk = mboSheet.getRange('E3').getValue(); // 次月モード状態
  const eCell1_1 = (e['range'].getRow() === 3 && e['range'].getColumn() === 1); // 今ココ
  const eCell1_2 = (e['range'].getRow() === 518 && e['range'].getColumn() === 5); // 過去ログ
  const eCell2 = (e['range'].getRow() === 1 && e['range'].getColumn() === 5); // 年度
  const eCell3 = (e['range'].getRow() === 2 && e['range'].getColumn() === 5); // 前月
  const eCell4 = (e['range'].getRow() === 3 && e['range'].getColumn() === 5); // 次月
  const userCell = (e['range'].getRow() === 518 && e['range'].getColumn() === 1); // 過去ログ（ユーザー）
  const yearCell = (e['range'].getRow() === 518 && e['range'].getColumn() === 3); // 過去ログ（年度）
  const monthCell = (e['range'].getRow() === 518 && e['range'].getColumn() === 4); // 過去ログ（月）
  const evtValue0 = e.oldValue; // 編集前
  const evtValue1 = e.value; // 編集後
  const eCell5 = (userCell || yearCell || monthCell) && (evtValue0 !== evtValue1); // 過去ログ
  const eCell = eCell1_1 || eCell1_2 || eCell2 || eCell3 || eCell4 || eCell5;
  const eCell234 = eCell2 || eCell3 || eCell4;
  const eCell34 = eCell3 || eCell4;
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
      mboSheet.getRange(beginRow_MBO,1,mboRow,endCol_MBO).createFilter()
        .setColumnFilterCriteria(hiddenColNum,rule);
    }
  }
}