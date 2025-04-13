// GTDフィルタ機能
function gtdFilter(e){
  // イベント対象シートがGTDシート上であることの判定
  const eGTDSheetGet = e.source.getActiveSheet();
  const eGTDSheetCheck = eGTDSheetGet.getName() === gtdSheet.getName();

  // イベント対象セルがマクロ実行セルであることの判定
  const eGTDRangeGet = e.range;
  const eGTDRangeCheck = eGTDRangeGet.getA1Notation() === gtdFilterCheckCell; // 実行チェック
  const runFlug1 = eGTDSheetCheck && eGTDRangeCheck;

  // 実行処理（シート・セル・中身の条件が合致した場合に実行）
  if(runFlug1){
    let filterGTD = gtdSheet.getFilter();
    if(filterGTD !== null){
      gtdSheet.getFilter().remove();
    }

    const nextMonth = gtdSheet.getRange(gtdImakokoMonthNextCell).getValue();
    gtdSheet.getRange(gtdImakokoMonthCurrentCell).setValue(nextMonth);

    const nextShitei = gtdSheet.getRange(gtdShiteiMonthNextCell).getValue();
    gtdSheet.getRange(gtdShiteiMonthCurrentCell).setValue(nextShitei);

    const monthCheck = gtdSheet.getRange(gtdMonthCheckCell).getValue();
    const priorityCheck = gtdSheet.getRange(gtdPriorityCheckCell).getValue();
    const hiddenCheck = gtdSheet.getRange(gtdHiddenCheckCell).getValue() === "なし";
    const runFlug2 = monthCheck || priorityCheck || !hiddenCheck;

    if(runFlug2){
      let rule = SpreadsheetApp.newFilterCriteria()
        .whenTextContains('V')
        .build();
      gtdSheet.getRange(beginRow_GTD,gtdStartCol,gtdRow,endCol_GTD).createFilter()
        .setColumnFilterCriteria(gtdFilterColNum,rule);
    }
  }
  gtdSheet.getRange(gtdFilterCheckCell).setValue(false);
}