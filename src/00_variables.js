// 変数まとめ

// ファイル・シート
const wflFile = SpreadsheetApp.getActiveSpreadsheet();
const mboSheet = wflFile.getSheetByName('MBO');
const calSheet = wflFile.getSheetByName('PJ');
const kakoSheetName = '過去ログ'; // 本体・過去ログファイル（共通）
const kakoSheetMain = wflFile.getSheetByName(kakoSheetName); // 本体ファイル
const kakoId = 'YOUR_EXTERNAL_SPREADSHEET_ID'; // 過去ログファイル
const kakoFile = SpreadsheetApp.openById(kakoId); // 過去ログファイル
const kakoSheetExternal = kakoFile.getSheetByName(kakoSheetName); // 過去ログファイル

// 行
const beginRow_MBO = 4;
const endRow_MBO = mboSheet.getMaxRows();
const mboRow = 893;
const mboRow_DayStart = 150;
const beginRow_Cal = 11;
const endRow_Cal = calSheet.getMaxRows();
const endRow_KakoMain = kakoSheetMain.getMaxRows(); // 本体ファイル
const endRow_KakoExternal = kakoSheetExternal.getMaxRows(); // 過去ログファイル

// 列(全シート共通)
const endCol_MBO = mboSheet.getMaxColumns();
const mboCol = 68;
const hiddenColNum = 2; // MBO_表示・非表示列（数値）
const eventColNum = 6; // MBO_Eventエリア開始列（数値）
const wStartColNum = 7; // MBO_Wエリア開始列（数値）
const fStartColNum = 16; // MBO_Fエリア開始列（数値）
const lStartColNum = 25; // MBO_Lエリア開始列（数値）
const eStartColNum = 34; // MBO_Eエリア開始列（数値）
const kaiheiCols = 2; // MBO_アコーディオン非表示列数
const wMokuhyoCol = 'G'; // MBO_W目標列（アルファベット）
const fMokuhyoCol = 'P'; // MBO_F目標列（アルファベット）
const lMokuhyoCol = 'Y'; // MBO_L目標列（アルファベット）
const eMokuhyoCol = 'AH'; // MBO_E目標列（アルファベット）
const wFurikaeriCol = 'H'; // MBO_W振返列（アルファベット）
const fFurikaeriCol = 'Q'; // MBO_F振返列（アルファベット）
const lFurikaeriCol = 'Z'; // MBO_L振返列（アルファベット）
const eFurikaeriCol = 'AI'; // MBO_E振返列（アルファベット）
const beginCol_Cal = 9;
const workCol_Cal = calSheet.getRange('H11').getValue();
const batStartCol_Cal = workCol_Cal+0;
const batGoalCol_Cal = workCol_Cal+1;
const batFlagCol_Cal = workCol_Cal+2;
const kakoCol = 35; // 本体・過去ログファイル（共通）