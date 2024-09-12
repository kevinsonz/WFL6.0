// Calフィルター
function calFilter(e){
    const filCk = calSheet.getRange('D5').getValue(); // チェック状態
    const cmpCk = calSheet.getRange('D6').getValue(); // チェック状態
    const eCell1 = (e['range'].getRow() === 5 && e['range'].getColumn() === 4); // チェック
    const eCell2 = (e['range'].getRow() === 6 && e['range'].getColumn() === 4); // チェック
    if(eCell1 || eCell2){
        let filterMBO = calSheet.getFilter();
        if(filterMBO !== null){
            calSheet.getFilter().remove();
        }
        let rule1 = SpreadsheetApp.newFilterCriteria()
                .setHiddenValues([])
                .build();
        if(filCk){
            rule1 = SpreadsheetApp.newFilterCriteria()
                .setHiddenValues([0])
                .build();
        }
        let rule2 = SpreadsheetApp.newFilterCriteria()
                .setHiddenValues([])
                .build();
        if(cmpCk){
            rule2 = SpreadsheetApp.newFilterCriteria()
                .setHiddenValues(['true'])
                .build();
        }
        calSheet.getRange(beginRow_Cal,1,endRow_Cal-beginRow_Cal+1,2).createFilter()
            .setColumnFilterCriteria(1,rule1)
            .setColumnFilterCriteria(2,rule2);
    }
}