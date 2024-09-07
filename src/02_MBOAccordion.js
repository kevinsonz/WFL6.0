// MBO：数値・グラフのみ列表示

function hideColMBO(e){
  const accCk = mboSheet.getRange('B3').getValue(); // アコーディオン状態
  const eValue = e['value'];
  const eCell = e['range'].getRow() === 3 && e['range'].getColumn() === 2;
  const colCheck = endCol_MBO === mboCol;
  const runFlag = eValue && eCell && colCheck;
    if(runFlag){
        const hideCols = [wStartColNum,fStartColNum,lStartColNum,eStartColNum];
        if(accCk){
          mboSheet.hideColumns(eventColNum,1);
          for(let i=0; i<hideCols.length; i++){
            mboSheet.hideColumns(hideCols[i],kaiheiCols);
          }
        }else if(!accCk){
          mboSheet.showColumns(eventColNum,1);
          for(let i=0; i<hideCols.length; i++){
            mboSheet.showColumns(hideCols[i],kaiheiCols);
          }
        }
    }
}