// MBO：TodayシートのDone等処理を実行

function doneMBO(e){
    const eCheck = e['value'];
    const doneCheck = todaySheet.getRange("G12").getValue();
    const errorCheck = todaySheet.getRange("L12").getValue();
    const kahiCheck = todaySheet.getRange("A2").getValue() === "可";
    const runFlag = eCheck && doneCheck && errorCheck && kahiCheck;
    const mode1 = eCheck === "時間分割モード";
    const mode2 = eCheck === "直接入力モード";
    const todayRow = todaySheet.getRange("O2").getValue();
    let yyyy = todaySheet.getRange("I2").getValue();
    let mm = todaySheet.getRange("J2").getValue()-1;
    let dd = todaySheet.getRange("K2").getValue();
    let yyyymmdd = new Date(yyyy,mm,dd);

    // モード切替え(時間分割<->直接入力)
    if(mode1 || mode2){
        if(mode1){
            todaySheet.showRows(16,3); // W時間分割行表示
            todaySheet.showRows(20,3); // F時間分割行表示
            todaySheet.showRows(24,3); // L時間分割行表示
            todaySheet.showRows(28,3); // E時間分割行表示
            todaySheet.getRange("E15").setBackgroundRGB(255,242,204); // W背景色（編集禁止）
            todaySheet.getRange("E19").setBackgroundRGB(255,242,204); // F背景色（編集禁止）
            todaySheet.getRange("E23").setBackgroundRGB(255,242,204); // L背景色（編集禁止）
            todaySheet.getRange("E27").setBackgroundRGB(255,242,204); // E背景色（編集禁止）
            todaySheet.getRange("E15").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B15,1,1)&"(工数)")'); // W工数集計数式貼付
            todaySheet.getRange("E19").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B19,1,1)&"(工数)")'); // F工数集計数式貼付
            todaySheet.getRange("E23").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B23,1,1)&"(工数)")'); // L工数集計数式貼付
            todaySheet.getRange("E27").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B27,1,1)&"(工数)")'); // E工数集計数式貼付
        }else if(mode2){
            todaySheet.hideRows(16,3); // W時間分割行非表示
            todaySheet.hideRows(20,3); // F時間分割行非表示
            todaySheet.hideRows(24,3); // L時間分割行非表示
            todaySheet.hideRows(28,3); // E時間分割行非表示
            todaySheet.getRange("E15").setBackgroundRGB(255,255,255); // W背景色（編集可能）
            todaySheet.getRange("E19").setBackgroundRGB(255,255,255); // F背景色（編集可能）
            todaySheet.getRange("E23").setBackgroundRGB(255,255,255); // L背景色（編集可能）
            todaySheet.getRange("E27").setBackgroundRGB(255,255,255); // E背景色（編集可能）
            todaySheet.getRange("E15").clearContent(); // W工数集計数式クリア
            todaySheet.getRange("E19").clearContent(); // F工数集計数式クリア
            todaySheet.getRange("E23").clearContent(); // L工数集計数式クリア
            todaySheet.getRange("E27").clearContent(); // E工数集計数式クリア
        }
    }

    // Done処理
    if(runFlag){
        todaySheet.getRange("G13").copyTo(mboSheet.getRange("BK"+todayRow),{contentsOnly:true}); // 起床
        todaySheet.getRange("G14").copyTo(mboSheet.getRange("BL"+todayRow),{contentsOnly:true}); // 就寝
        todaySheet.getRange("E14").copyTo(mboSheet.getRange("H"+todayRow),{contentsOnly:true}); // 仮眠
        todaySheet.getRange("C15").copyTo(mboSheet.getRange("K"+todayRow),{contentsOnly:true}); // W目標
        todaySheet.getRange("C19").copyTo(mboSheet.getRange("R"+todayRow),{contentsOnly:true}); // F目標
        todaySheet.getRange("C23").copyTo(mboSheet.getRange("Y"+todayRow),{contentsOnly:true}); // L目標
        todaySheet.getRange("C27").copyTo(mboSheet.getRange("AF"+todayRow),{contentsOnly:true}); // E目標
        todaySheet.getRange("K16").copyTo(mboSheet.getRange("L"+todayRow),{contentsOnly:true}); // W振返
        todaySheet.getRange("K20").copyTo(mboSheet.getRange("S"+todayRow),{contentsOnly:true}); // F振返
        todaySheet.getRange("K24").copyTo(mboSheet.getRange("Z"+todayRow),{contentsOnly:true}); // L振返
        todaySheet.getRange("K28").copyTo(mboSheet.getRange("AG"+todayRow),{contentsOnly:true}); // E振返
        todaySheet.getRange("E15").copyTo(mboSheet.getRange("M"+todayRow),{contentsOnly:true}); // W工数
        todaySheet.getRange("E19").copyTo(mboSheet.getRange("T"+todayRow),{contentsOnly:true}); // F工数
        todaySheet.getRange("E23").copyTo(mboSheet.getRange("AA"+todayRow),{contentsOnly:true}); // L工数
        todaySheet.getRange("E27").copyTo(mboSheet.getRange("AH"+todayRow),{contentsOnly:true}); // E工数
        todaySheet.getRange("F15").copyTo(mboSheet.getRange("P"+todayRow),{contentsOnly:true}); // W相殺
        todaySheet.getRange("F19").copyTo(mboSheet.getRange("W"+todayRow),{contentsOnly:true}); // F相殺
        todaySheet.getRange("F23").copyTo(mboSheet.getRange("AD"+todayRow),{contentsOnly:true}); // L相殺
        todaySheet.getRange("F27").copyTo(mboSheet.getRange("AK"+todayRow),{contentsOnly:true}); // E相殺
        todaySheet.getRange("G15").copyTo(mboSheet.getRange("O"+todayRow),{contentsOnly:true}); // W評価
        todaySheet.getRange("G19").copyTo(mboSheet.getRange("V"+todayRow),{contentsOnly:true}); // F評価
        todaySheet.getRange("G23").copyTo(mboSheet.getRange("AC"+todayRow),{contentsOnly:true}); // L評価
        todaySheet.getRange("G27").copyTo(mboSheet.getRange("AJ"+todayRow),{contentsOnly:true}); // E評価
        todaySheet.getRange("C15:E15").clearContent(); // W目標・振返・工数クリア
        todaySheet.getRange("C19:E19").clearContent(); // F目標・振返・工数クリア
        todaySheet.getRange("C23:E23").clearContent(); // L目標・振返・工数クリア
        todaySheet.getRange("C27:E27").clearContent(); // E目標・振返・工数クリア
        todaySheet.getRange("E14").clearContent(); // 仮眠クリア
        todaySheet.getRange("F15").setValue(false); // W相殺クリア
        todaySheet.getRange("F19").setValue(false); // W相殺クリア
        todaySheet.getRange("F23").setValue(false); // W相殺クリア
        todaySheet.getRange("F27").setValue(false); // W相殺クリア
        todaySheet.getRange("G13").setValue("起床"); // 起床クリア
        todaySheet.getRange("G14").setValue("就寝"); // 就寝クリア
        todaySheet.getRange("G15").clearContent(); // W評価クリア
        todaySheet.getRange("G19").clearContent(); // F評価クリア
        todaySheet.getRange("G23").clearContent(); // L評価クリア
        todaySheet.getRange("G27").clearContent(); // E評価クリア
        todaySheet.getRange("D16:E18").clearContent(); // W時間分割クリア
        todaySheet.getRange("D20:E22").clearContent(); // F時間分割クリア
        todaySheet.getRange("D24:E26").clearContent(); // L時間分割クリア
        todaySheet.getRange("D28:E30").clearContent(); // E時間分割クリア
        const mode = todaySheet.getRange("H12").getValue(); // 現在のモードを取得
        if(mode === '時間分割モード'){ // 時間分割モードの場合は数式を貼り直す
            todaySheet.getRange("E15").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B15,1,1)&"(工数)")'); // W工数集計数式貼付
            todaySheet.getRange("E19").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B19,1,1)&"(工数)")'); // F工数集計数式貼付
            todaySheet.getRange("E23").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B23,1,1)&"(工数)")'); // L工数集計数式貼付
            todaySheet.getRange("E27").setFormula('=sumifs($J$15:$J$30,$I$15:$I$30,mid($B27,1,1)&"(工数)")'); // E工数集計数式貼付
        }
        yyyymmdd = new Date(yyyymmdd.getFullYear(), yyyymmdd.getMonth(), yyyymmdd.getDate()+1); // 日付スライド
        yyyy = yyyymmdd.getFullYear(); // yyyy再取得
        mm = yyyymmdd.getMonth()+1; // mm再取得
        dd = yyyymmdd.getDate(); // dd再取得
        todaySheet.getRange("I2").setValue(yyyy); // yyyy再貼付
        todaySheet.getRange("J2").setValue(mm); // mm再貼付
        todaySheet.getRange("K2").setValue(dd); // dd再貼付
        todaySheet.getRange("O1").setValue(Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')); // 処理日時
    }
    todaySheet.getRange("G12").setValue(false); // Doneチェックボックス初期化
}