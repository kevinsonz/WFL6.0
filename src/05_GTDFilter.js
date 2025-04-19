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
    const nextYYYY = gtdSheet.getRange(gtdShiteiYYYYNextCell).getValue();
    gtdSheet.getRange(gtdShiteiYYYYCurrentCell).setValue(nextYYYY);
    const nextMM = gtdSheet.getRange(gtdShiteiMMNextCell).getValue();
    gtdSheet.getRange(gtdShiteiMMCurrentCell).setValue(nextMM);

    const monthCheck = gtdSheet.getRange(gtdMonthCheckCell).getValue();
    const priorityCheck = gtdSheet.getRange(gtdPriorityCheckCell).getValue();
    const horyuCheck = gtdSheet.getRange(gtdHoryuCheckCell).getValue();
    const shuryoCheck = gtdSheet.getRange(gtdShuryoCheckCell).getValue();
    const onlyHoryuShuryoCheck = gtdSheet.getRange(gtdOnlyHoryuAndShuryoCell).getValue();
    const runFlug2 = !monthCheck && !priorityCheck && horyuCheck && shuryoCheck && !onlyHoryuShuryoCheck;

    if(!runFlug2){
      let rule = SpreadsheetApp.newFilterCriteria()
        .whenTextContains('V')
        .build();
      gtdSheet.getRange(beginRow_GTD,gtdStartCol,gtdRow,endCol_GTD).createFilter()
        .setColumnFilterCriteria(gtdFilterColNum,rule);
    }
  }
  gtdSheet.getRange(gtdFilterCheckCell).setValue(false);
}