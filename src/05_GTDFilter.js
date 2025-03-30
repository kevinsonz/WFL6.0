// GTDフィルタ機能
function gtdFilter(e){
    const gtdCurrentMonthCheckFlag = gtdSheet.getRange(gtdCurrentMonthCheckCell).getValue(); // 当月フラグ
    const gtdDelayCheckFlag = gtdSheet.getRange(gtdDelayCheckCell).getValue(); // 遅延フラグ

    // イベント対象シートがGTDシート上であることの判定
    const eGTDSheetGet = e.source.getActiveSheet();
    const eGTDSheetCk = eGTDSheetGet.getName() === gtdSheet.getName();
  
    // イベント対象セルがマクロ実行セルであることの判定
    const eGTDRangeGet = e.range;
    const eGTDRangeCk1 = eGTDRangeGet.getA1Notation() === gtdCurrentMonthCheckCell; // 当月チェック
    const eGTDRangeCk2 = eGTDRangeGet.getA1Notation() === gtdDelayCheckCell; // 遅延チェック
    const eGTDRangeCk = eGTDRangeCk1 || eGTDRangeCk2;

    // 指定年度のデータに変更が発生していることの判定
    const eGTDValueCk = e.value !== e.oldvalue;
  
    // 実行処理（シート・セル・中身の条件が合致した場合に実行）
    if(eGTDSheetCk && eGTDRangeCk && eGTDValueCk){
        let filterGTD = gtdSheet.getFilter();
        if(filterGTD !== null){
            gtdSheet.getFilter().remove();
        }
        let rule = SpreadsheetApp.newFilterCriteria()
            .whenTextContains('V')
            .build();
        gtdSheet.getRange(beginRow_GTD,gtdStartCol,gtdRow,endCol_GTD).createFilter()
            .setColumnFilterCriteria(gtdFilterColNum,rule);
    }
}