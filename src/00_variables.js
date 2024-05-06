// 変数まとめ

// ファイル・シート
const wflFile = SpreadsheetApp.getActiveSpreadsheet();
const mboSheet = wflFile.getSheetByName('MBO');
const gtdSheet = wflFile.getSheetByName('GTD');
const todaySheet = wflFile.getSheetByName('Today');

// 行
const beginRow_MBO = 4;
const endRow_MBO = mboSheet.getMaxRows();
const mboRow = 570;
const beginRow_GTD = 3;
const endRow_GTD = gtdSheet.getMaxRows();
const mboRow_DayStart = 208;

// 列(全シート共通)
const endCol_MBO = mboSheet.getMaxColumns();
const mboCol = 97;
const endCol_GTD = gtdSheet.getMaxColumns();

// 行列
const gtdNo = gtdSheet.getRange(beginRow_GTD,1,endRow_GTD,1).getValues();
const gtdData = gtdSheet.getRange(beginRow_GTD,1,endRow_GTD,11).getValues();

// 識別
const statusMBO = mboSheet.getRange('A1').getValue();
const statusGTD = gtdSheet.getRange('A1').getValue();
const filterGTD = gtdSheet.getRange('I1').getValue();
