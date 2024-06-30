// 変数まとめ

// ファイル・シート
const wflFile = SpreadsheetApp.getActiveSpreadsheet();
const mboSheet = wflFile.getSheetByName('MBO');
const gtdSheet = wflFile.getSheetByName('GTD');
const todaySheet = wflFile.getSheetByName('Today');

// 行
const beginRow_MBO = 5;
const endRow_MBO = mboSheet.getMaxRows();
const mboRow = 515;
const mboRow_DayStart = 150;
const beginRow_GTD = 3;
const endRow_GTD = gtdSheet.getMaxRows();

// 列(全シート共通)
const endCol_MBO = mboSheet.getMaxColumns();
const mboCol = 68;
const endCol_GTD = gtdSheet.getMaxColumns();
const hiddenRowNum = 2; // MBO_表示・非表示列（数値）
const wStartColNum = 7; // MBO_Wエリア開始列（数値）
const fStartColNum = 16; // MBO_Fエリア開始列（数値）
const lStartColNum = 25; // MBO_Lエリア開始列（数値）
const eStartColNum = 34; // MBO_Eエリア開始列（数値）
const kaiheiCols = 2; // MBO_アコーディオン非表示列数
// const kishoCol = 'G'; // MBO_起床列（アルファベット）
// const shushinCol = 'H'; // MBO_就寝列（アルファベット）
// const kaminCol = 'K'; // MBO_仮眠列（アルファベット）
const wMokuhyoCol = 'G'; // MBO_W目標列（アルファベット）
const fMokuhyoCol = 'P'; // MBO_F目標列（アルファベット）
const lMokuhyoCol = 'Y'; // MBO_L目標列（アルファベット）
const eMokuhyoCol = 'AH'; // MBO_E目標列（アルファベット）
const wFurikaeriCol = 'H'; // MBO_W振返列（アルファベット）
const fFurikaeriCol = 'Q'; // MBO_F振返列（アルファベット）
const lFurikaeriCol = 'Z'; // MBO_L振返列（アルファベット）
const eFurikaeriCol = 'AI'; // MBO_E振返列（アルファベット）
// const wKousuCol = 'P'; // MBO_W工数列（アルファベット）
// const fKousuCol = 'W'; // MBO_F工数列（アルファベット）
// const lKousuCol = 'AD'; // MBO_L工数列（アルファベット）
// const eKousuCol = 'AK'; // MBO_E工数列（アルファベット）
// const wSousaiCol = 'S'; // MBO_W相殺列（アルファベット）
// const fSousaiCol = 'Z'; // MBO_F相殺列（アルファベット）
// const lSousaiCol = 'AG'; // MBO_L相殺列（アルファベット）
// const eSousaiCol = 'AN'; // MBO_E相殺列（アルファベット）
// const wHyoukaCol = 'R'; // MBO_W評価列（アルファベット）
// const fHyoukaCol = 'Y'; // MBO_F評価列（アルファベット）
// const lHyoukaCol = 'AF'; // MBO_L評価列（アルファベット）
// const eHyoukaCol = 'AM'; // MBO_E評価列（アルファベット）

// 行列
const gtdNo = gtdSheet.getRange(beginRow_GTD,1,endRow_GTD,1).getValues();
const gtdData = gtdSheet.getRange(beginRow_GTD,1,endRow_GTD,11).getValues();

// 識別
const statusMBO = mboSheet.getRange('A2').getValue();
const statusGTD = gtdSheet.getRange('A1').getValue();
const filterGTD = gtdSheet.getRange('I1').getValue();
