// 変数まとめ

// ファイル・シート
const wflFile = SpreadsheetApp.getActiveSpreadsheet();
const mboSheet = wflFile.getSheetByName('MBO');
const consoleSheet = wflFile.getSheetByName('console');
const kakoSheetName = '過去ログ'; // 本体ファイル
const kakoSheetMain = wflFile.getSheetByName(kakoSheetName); // 本体ファイル
const kakoFile = SpreadsheetApp.openById(kakoId); // 過去ログファイル
const gtdSheet = wflFile.getSheetByName('GTD'); // 新GTD

// 行
const beginRow_MBO = 4;
const beginRow_GTD = 11; // 新GTD（開始行）
const endRow_MBO = mboSheet.getMaxRows();
const endRow_GTD = gtdSheet.getMaxRows(); // 新GTD（終了行）
const mboRow = 893;
const gtdRow = endRow_GTD - beginRow_GTD + 1; // 新GTD（行数）※見出し行含む
const mboRow_DayStart = 150;

// 列(全シート共通)
const endCol_MBO = mboSheet.getMaxColumns();
const endCol_GTD = gtdSheet.getMaxColumns(); // 新GTD（列数）
const mboStartCol = 1; // MBOシート開始列
const gtdStartCol = 1; // GTDシート開始列
const mboEndCol = 68; // MBOシート終了列
const hiddenColNum = 2; // MBO_表示・非表示列（数値）
const gtdFilterColNum = 1; // GTD_表示・非表示列（数値）
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

// セル
const imakokoRunCell = 'A3' // 今ココ絞込みを実行するフラグ
const yearViewRunCell = 'E1' // 年度表示モードを実行するフラグ
const prevMonthRunCell = 'E2' // 前月表示モードを実行するフラグ
const nextMonthRunCell = 'E3' // 翌月表示モードを実行するフラグ
const accordionRunCell = 'B3' // アコーディオンを実行するフラグ
const kakoYearCell = 'C518' // 過去ログ対象とする年度
const gtdCurrentMonthCheckCell = 'B6'; // GTD当月チェックボックス
const gtdDelayCheckCell = 'C6'; // GTD遅延チェックボックス
const gtdTargetYearCell = 'E5'; // GTD指定年度プルダウン
const gtdTargetMonthCell = 'E6'; // GTD指定月プルダウン