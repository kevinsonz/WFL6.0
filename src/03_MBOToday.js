// MBO：TodayシートのDone等処理を実行

function doneMBO(e){
    const eCheck = e['value'];
    const doneCheck = todaySheet.getRange("G11").getValue();
    const errorCheck = todaySheet.getRange("M11").getValue();
    const kahiCheck = todaySheet.getRange("A2").getValue() === "可";
    const runFlag = eCheck && doneCheck && errorCheck && kahiCheck;
    const todayRow = todaySheet.getRange("P2").getValue();
    let yyyy = todaySheet.getRange("J2").getValue();
    let mm = todaySheet.getRange("K2").getValue()-1;
    let dd = todaySheet.getRange("L2").getValue();
    let yyyymmdd = new Date(yyyy,mm,dd);
    if(eCheck){
        todaySheet.getRange("J4:Q10").clearContent();
        todaySheet.getRange("J4").setValue(eCheck);
        todaySheet.getRange("K4").setValue(doneCheck);
        todaySheet.getRange("L4").setValue(errorCheck);
        todaySheet.getRange("M4").setValue(kahiCheck);
        todaySheet.getRange("N4").setValue(runFlag);
        todaySheet.getRange("O4").setValue(todayRow);
        todaySheet.getRange("J5").setValue(yyyy);
        todaySheet.getRange("K5").setValue(mm);
        todaySheet.getRange("L5").setValue(dd);
        todaySheet.getRange("M5").setValue(yyyymmdd);
    }
    if(runFlag){
        const zValue = todaySheet.getRange("G13").getValue(); todaySheet.getRange("N5").setValue(zValue);
        const wfleValues = todaySheet.getRange("C14:H17").getValues(); todaySheet.getRange("J6:O9").setValues(wfleValues);
        todaySheet.getRange("G13").copyTo(mboSheet.getRange("G"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C14:G14").copyTo(mboSheet.getRange("I"+todayRow+":N"+todayRow),{contentsOnly:true});
        mboSheet.getRange("L"+todayRow).setFormula(mboSheet.getRange("L3").getFormula());
        mboSheet.getRange("N"+todayRow).setFormula(mboSheet.getRange("N3").getFormula());
        todaySheet.getRange("C15:G15").copyTo(mboSheet.getRange("O"+todayRow+":T"+todayRow),{contentsOnly:true});
        mboSheet.getRange("R"+todayRow).setFormula(mboSheet.getRange("R3").getFormula());
        mboSheet.getRange("T"+todayRow).setFormula(mboSheet.getRange("T3").getFormula());
        todaySheet.getRange("C16:G16").copyTo(mboSheet.getRange("U"+todayRow+":Z"+todayRow),{contentsOnly:true});
        mboSheet.getRange("X"+todayRow).setFormula(mboSheet.getRange("X3").getFormula());
        mboSheet.getRange("Z"+todayRow).setFormula(mboSheet.getRange("Z3").getFormula());
        todaySheet.getRange("C17:G17").copyTo(mboSheet.getRange("AA"+todayRow+":AF"+todayRow),{contentsOnly:true});
        mboSheet.getRange("AD"+todayRow).setFormula(mboSheet.getRange("AD3").getFormula());
        mboSheet.getRange("AF"+todayRow).setFormula(mboSheet.getRange("AF3").getFormula());
        todaySheet.getRange("G13").clearContent();
        todaySheet.getRange("C14:H17").clearContent();
        yyyymmdd = new Date(yyyymmdd.getFullYear(), yyyymmdd.getMonth(), yyyymmdd.getDate()+1); todaySheet.getRange("J10").setValue(yyyymmdd);
        yyyy = yyyymmdd.getFullYear(); todaySheet.getRange("K10").setValue(yyyy);
        mm = yyyymmdd.getMonth()+1; todaySheet.getRange("L10").setValue(mm);
        dd = yyyymmdd.getDate(); todaySheet.getRange("M10").setValue(dd);
        todaySheet.getRange("J2").setValue(yyyy);
        todaySheet.getRange("K2").setValue(mm);
        todaySheet.getRange("L2").setValue(dd);
        todaySheet.getRange("P1").setValue(Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss'));
    }
    todaySheet.getRange("G11").setValue(false);
}