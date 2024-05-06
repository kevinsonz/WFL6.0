// MBO：数値・グラフのみ列表示（作成中）

function hideColMBO(e){
    const ePosition = e['range'].getRow() === 1 && e['range'].getColumn() === 1;
    const eValue = (e['value'] === '閉' || e['value'] === '開');
    const runFlag = ((endCol_MBO === mboCol) && ePosition && eValue);
    if(runFlag){
        const hideCols = [11,18,25,32];
        if(e['value'] === '閉'){
            for(let i=0; i<hideCols.length; i++){
                mboSheet.hideColumns(hideCols[i],2);
            }
        }else if(e['value'] === '開'){
            for(let i=0; i<hideCols.length; i++){
                mboSheet.showColumns(hideCols[i],2);
            }
        }
        mboSheet.getRange('A1').setValue(e['oldValue']);
    }
}