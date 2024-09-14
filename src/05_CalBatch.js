// Cal範囲指定
function calBatch(){
    const batNumber = calSheet.getRange(beginRow_Cal,batFlagCol_Cal).getValue(); // チェック状態
    const batStartRow = calSheet.getRange(beginRow_Cal+batNumber,batStartCol_Cal).getValue(); // →列情報
    const batGoalRow = calSheet.getRange(beginRow_Cal+batNumber,batGoalCol_Cal).getValue(); // ←列情報

    if(batNumber>0){
        let pjRowData = calSheet.getRange(beginRow_Cal+batNumber,beginCol_Cal,1,366).getValues();
        for(i=0;i<pjRowData[0].length;i++){
            if((batStartRow-1)<=i && i<=(batGoalRow-1)){
                if(pjRowData[0][i]===''||pjRowData[0][i]==='→'||pjRowData[0][i]==='←'){
                    pjRowData[0].splice(i,1,'□');
                }
            }else{
                if(pjRowData[0][i]==='□'){
                    pjRowData[0].splice(i,1,'');
                }
            }
        }
        calSheet.getRange(beginRow_Cal+batNumber,beginCol_Cal,1,366).setValues(pjRowData);
    }
}