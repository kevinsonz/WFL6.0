// 取得した過去ログ（全量）を対象年度のみに絞り込む処理
function filterArrayByColumn(data, yearColumnIndex, targetYear) {
  data.shift();
  const result = data.filter((row) => Number(row[yearColumnIndex]) === Number(targetYear));
  consoleSheet.getRange(10,1,result.length,35).setValues('result:'+result);
  return result;
  // return data.filter(function(row) {
  //   consoleSheet.getRange('A5').setValue('targetYear:'+targetYear);
  //   consoleSheet.getRange('A6').setValue('row:'+row);
  //   return row[yearColumnIndex] === targetYear;
  // });
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
  // イベント対象シートがMBOシート上であることの判定
  const eSheetGet = e.source.getActiveSheet();
  const eSheetCk = eSheetGet.getName() === mboSheet.getName();

  // イベント対象セルが過去ログ実行フラグであるの判定
  const eRangeGet = e.range;
  const eRangeCk = eRangeGet.getA1Notation() === kakoRunCell;

  // イベント対象セルの中身が実行フラグ（true）であることの判定
  const eValueGet = e.value;
  const eValueCk = eValueGet;

  // 実行処理（シート・セル・中身の条件が合致した場合に実行）
  if(eSheetCk && eRangeCk && eValueCk){
      const kakoYear = mboSheet.getRange(kakoYearCell).getValue(); // 指定年度
      const kakoValues = kakoSheetExternal.getDataRange().getValues();
      const targetYearValues = filterArrayByColumn(kakoValues,2,Number(kakoYear));
      // deleteAndAddRows(2+1,endRow_KakoMain,kakoValues.length);
      kakoSheetMain.getRange(kakoStratRow,kakoStartCol,kakoMainRow,kakoEndCol).clearContent();
      consoleSheet.getRange('A1').setValue('kakoYear:'+kakoYear);
      consoleSheet.getRange('A2').setValue('kakoValues[0][1]:'+kakoValues[0][1]);
      // consoleSheet.getRange('A3').setValue('targetValues[0][1]:'+targetValues[0][1]);
      kakoSheetMain.getRange(kakoStratRow,kakoStartCol,targetYearValues.length,kakoEndCol).setValues(targetYearValues);
  }
}