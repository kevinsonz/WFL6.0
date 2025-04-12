// GTDフィルタ機能
function gtdFilter(e){
    // イベント対象シートがGTDシート上であることの判定
    const eGTDSheetGet = e.source.getActiveSheet();
    const eGTDSheetCk = eGTDSheetGet.getName() === gtdSheet.getName();
  
    // イベント対象セルがマクロ実行セルであることの判定
    const eGTDRangeGet = e.range;
    const eGTDRangeCk = eGTDRangeGet.getA1Notation() === gtdFilterCheckCell; // 実行チェック
  
    // 指定年度のデータに変更が発生していることの判定
    const eGTDValueCk = e.value !== e.oldvalue;
  
    // 実行処理（シート・セル・中身の条件が合致した場合に実行）
    if(eGTDSheetCk && eGTDRangeCk && eGTDValueCk){
      let filterGTD = gtdSheet.getFilter();
      if(filterGTD !== null){
          gtdSheet.getFilter().remove();
      }
      if(e.value){
      let rule = SpreadsheetApp.newFilterCriteria()
          .whenTextContains('V')
          .build();
      gtdSheet.getRange(beginRow_GTD,gtdStartCol,gtdRow,endCol_GTD).createFilter()
          .setColumnFilterCriteria(gtdFilterColNum,rule);
      }
    }
    gtdSheet.getRange(gtdFilterCheckCell).setValue(false);
  }