// Cal範囲指定
function calBatch(){
    const batNumber = calSheet.getRange(beginRow_Cal,batFlagCol_Cal).getValue(); // チェック状態
    const batStartRow = getRange(beginRow_Cal+batNumber,batStartCol_Cal).getValue(); // →列情報
    const batGoalRow = getRange(beginRow_Cal+batNumber,batGoalCol_Cal).getValue(); // ←列情報
    if(batNumber>0){
        let pjRowData = [];
        pjRowData = getRange((beginRow_Cal + batNumber),beginCol_Cal).getValues().flat();
        for(i=0;i<pjRowData.length;i++){
            if((batStartRow-1)<=i && i<=(batGoalRow-1)){
                if(pjRowData[i]=''){
                    pjRowData.splice(i,1,'□');
                }
            }
            if(pjRowData[i]='□'){
                pjRowData.splice(i,1,'');
            }
        }
        getRange((beginRow_Cal + batNumber),beginCol_Cal).setValues(pjRowData);
    }
}