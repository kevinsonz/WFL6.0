// MBO：TodayシートのDone等処理を実行

function doneMBO(e){
    const eCheck = e['value'];
    const doneCheck = todaySheet.getRange("H12").getValue();
    const errorCheck = todaySheet.getRange("N12").getValue();
    const kahiCheck = todaySheet.getRange("A2").getValue() === "可";
    const runFlag = eCheck && doneCheck && errorCheck && kahiCheck;
    const todayRow = todaySheet.getRange("Q2").getValue();
    let yyyy = todaySheet.getRange("K2").getValue();
    let mm = todaySheet.getRange("L2").getValue()-1;
    let dd = todaySheet.getRange("M2").getValue();
    let yyyymmdd = new Date(yyyy,mm,dd);
    if(runFlag){
        todaySheet.getRange("E13").copyTo(mboSheet.getRange("BK"+todayRow),{contentsOnly:true});
        todaySheet.getRange("E14").copyTo(mboSheet.getRange("BL"+todayRow),{contentsOnly:true});
        todaySheet.getRange("H14").copyTo(mboSheet.getRange("H"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C15:F15").copyTo(mboSheet.getRange("K"+todayRow+":N"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C16:F16").copyTo(mboSheet.getRange("R"+todayRow+":U"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C17:F17").copyTo(mboSheet.getRange("Y"+todayRow+":AB"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C18:F18").copyTo(mboSheet.getRange("AF"+todayRow+":AI"+todayRow),{contentsOnly:true});
        todaySheet.getRange("H15").copyTo(mboSheet.getRange("P"+todayRow),{contentsOnly:true});
        todaySheet.getRange("H16").copyTo(mboSheet.getRange("W"+todayRow),{contentsOnly:true});
        todaySheet.getRange("H17").copyTo(mboSheet.getRange("AD"+todayRow),{contentsOnly:true});
        todaySheet.getRange("H18").copyTo(mboSheet.getRange("AK"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C15:E18").clearContent();
        todaySheet.getRange("H14:H18").clearContent();
        todaySheet.getRange("F15:F18").setValue(false);
        todaySheet.getRange("E13").setValue("起床");
        todaySheet.getRange("E14").setValue("就寝");
        yyyymmdd = new Date(yyyymmdd.getFullYear(), yyyymmdd.getMonth(), yyyymmdd.getDate()+1);
        yyyy = yyyymmdd.getFullYear();
        mm = yyyymmdd.getMonth()+1;
        dd = yyyymmdd.getDate();
        todaySheet.getRange("K2").setValue(yyyy);
        todaySheet.getRange("L2").setValue(mm);
        todaySheet.getRange("M2").setValue(dd);
        todaySheet.getRange("Q1").setValue(Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss'));
    }
    todaySheet.getRange("H12").setValue(false);
}