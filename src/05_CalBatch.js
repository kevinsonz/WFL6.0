// Cal範囲指定
function calBatch(){
    const batNumber = calSheet.getRange(beginRow_Cal,batFlagCol_Cal).getValue(); // 実行チェック
    const batType = calSheet.getRange(beginRow_Cal-1,batFlagCol_Cal).getValue(); // 処理タイプ
    const batStartCol = calSheet.getRange(beginRow_Cal+batNumber,batStartCol_Cal).getValue(); // →列情報
    const batGoalCol = calSheet.getRange(beginRow_Cal+batNumber,batGoalCol_Cal).getValue(); // ←列情報
    let batStartCol2 = batStartCol; // 後で加工するための枠
    let batGoalCol2 = batGoalCol; // 後で加工するための枠
    if(batType==='B'){
        batStartCol2 = batGoalCol; // Bの場合は逆転
        batGoalCol2 = batStartCol; // Bの場合は逆転
    }

    if(batNumber>0){
        let pjRowData = calSheet.getRange(beginRow_Cal+batNumber,beginCol_Cal,1,366).getValues();
        for(i=0;i<pjRowData[0].length;i++){
            if(batType==='A'){
                if((batStartCol2-1)<=i && i<=(batGoalCol2-1)){
                    if(pjRowData[0][i]==='' || pjRowData[0][i]==='→' || pjRowData[0][i]==='←'){
                            pjRowData[0].splice(i,1,'□');
                    }
                }else{
                    if(pjRowData[0][i]==='□'){
                        pjRowData[0].splice(i,1,'');
                    }
                }
            }
            if(batType==='B'){
                if(pjRowData[0][i]==='□' || pjRowData[0][i]==='→' || pjRowData[0][i]==='←'){
                    pjRowData[0].splice(i,1,'');
                }
            }
        }
        calSheet.getRange(beginRow_Cal+batNumber,beginCol_Cal,1,366).setValues(pjRowData);
    }
}