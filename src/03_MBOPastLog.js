// 取得した過去ログ（全量）を対象年度のみに絞り込む処理
function filterArrayByColumn(data, columnIndex, targetString) {
    return data.filter(function(row) {
      return row[columnIndex] === targetString;
    });
  }

// 本体の過去ログシートを見出し＋データ行×1のみにしてから行を追加する処理
function deleteAndAddRows(startRow, endRow, numRowsToAdd) {
    // 削除範囲がシートの範囲内にあるか確認
    if (startRow < 1 || endRow > sheet.getMaxRows() || startRow > endRow) {
      Logger.log("削除範囲が無効です。");
      return;
    }
    // 行を削除
    kakoSheetMain.deleteRows(startRow, endRow - startRow + 1);
  
    // 行を追加
    kakoSheetMain.insertRowsBefore(startRow, numRowsToAdd);
  
    Logger.log("行を削除し、" + numRowsToAdd + "行を追加しました。");
  }

// 過去ログ外部読込処理
function getPastLog(e){
    const eSheetGet = e.source.getActiveSheet(); const eSheetCk = eSheetGet === mboSheet;
    const eRangeGet = e.range; const eRangeCk = eRangeGet === 'E518';
    const eValueGet = e.value; const eValueCk = eValueGet;
    if(eSheetCk && eRangeCk && eValueCk){
        const kakoYear = mboSheet.getRange('C518').getValue(); // 指定年度
        let kakoValues = kakoSheetExternal.getRange(2,1,endRow_KakoExternal-1,kakoCol).getValues();
        filterArrayByColumn(kakoValues,2,kakoYear);
        deleteAndAddRows(2+1,endRow_KakoMain,kakoValues.length);
        kakoSheetMain.getRange(2,1,kakoValues.length,kakoCol).setValues(kakoValues);
    }
}