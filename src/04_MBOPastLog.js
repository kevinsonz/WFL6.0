// 過去ログ外部読込処理
function getPastLog(e){
  // イベント対象シートがMBOシート上であることの判定
  const eSheetGet = e.source.getActiveSheet();
  const eSheetCk = eSheetGet.getName() === mboSheet.getName();

  // イベント対象セルが年度指定セルであることの判定
  const eRangeGet = e.range;
  const eRangeCk = eRangeGet.getA1Notation() === kakoYearCell;

  // 指定年度のデータに変更が発生していることの判定
  const eValueCk = e.value !== e.oldvalue;

  // 実行処理（シート・セル・中身の条件が合致した場合に実行）
  if(eSheetCk && eRangeCk && eValueCk){
      const kakoYear = mboSheet.getRange(kakoYearCell).getValue(); // 指定年度
      const kakoSheetExternal = kakoFile.getSheetByName(kakoYear); // コピー元（外部ファイル）の対象シート
      const kakoSheetExternalValues = kakoSheetExternal.getDataRange().getValues(); // コピー元シートのデータを全て取得
      kakoSheetMain.getDataRange().setValues(kakoSheetExternalValues); // コピー元→コピー先（丸ごとコピー）
  }
}