// MBO：TodayシートのDone等処理を実行

function doneMBO(e){
    //todaySheet.getRange("L3:P3").clearContent();
    const eCheck = e['value']; //todaySheet.getRange("L3").setValue(eCheck);
    const doneCheck = todaySheet.getRange("G11").getValue(); //todaySheet.getRange("M3").setValue(doneCheck);
    const errorCheck = todaySheet.getRange("M11").getValue(); //todaySheet.getRange("N3").setValue(errorCheck);
    const kahiCheck = todaySheet.getRange("A2").getValue() === "可"; //todaySheet.getRange("O3").setValue(kahiCheck);
    const runFlag = eCheck && doneCheck && errorCheck && kahiCheck; //todaySheet.getRange("P3").setValue(runFlag);
    const todayRow = todaySheet.getRange("P2").getValue();
    let yyyy = todaySheet.getRange("J2").getValue();
    let mm = todaySheet.getRange("K2").getValue()-1;
    let dd = todaySheet.getRange("L2").getValue();
    let yyyymmdd = new Date(yyyy,mm,dd);
    if(runFlag){
        const lFormula = mboSheet.getRange("L"+todayRow).getFormula();
        const nFormula = mboSheet.getRange("N"+todayRow).getFormula();
        const rFormula = mboSheet.getRange("R"+todayRow).getFormula();
        const tFormula = mboSheet.getRange("T"+todayRow).getFormula();
        const xFormula = mboSheet.getRange("X"+todayRow).getFormula();
        const zFormula = mboSheet.getRange("Z"+todayRow).getFormula();
        const adFormula = mboSheet.getRange("AD"+todayRow).getFormula();
        const afFormula = mboSheet.getRange("AF"+todayRow).getFormula();
        todaySheet.getRange("G13").copyTo(mboSheet.getRange("G"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C14:G14").copyTo(mboSheet.getRange("I"+todayRow+":N"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C15:G15").copyTo(mboSheet.getRange("O"+todayRow+":T"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C16:G16").copyTo(mboSheet.getRange("U"+todayRow+":Z"+todayRow),{contentsOnly:true});
        todaySheet.getRange("C17:G17").copyTo(mboSheet.getRange("AA"+todayRow+":AF"+todayRow),{contentsOnly:true});
        mboSheet.getRange("L"+todayRow).setFormula(lFormula);
        mboSheet.getRange("N"+todayRow).setFormula(nFormula);
        mboSheet.getRange("R"+todayRow).setFormula(rFormula);
        mboSheet.getRange("T"+todayRow).setFormula(tFormula);
        mboSheet.getRange("X"+todayRow).setFormula(xFormula);
        mboSheet.getRange("Z"+todayRow).setFormula(zFormula);
        mboSheet.getRange("AD"+todayRow).setFormula(adFormula);
        mboSheet.getRange("AF"+todayRow).setFormula(afFormula);
        todaySheet.getRange("G13").clearContent();
        todaySheet.getRange("C14:H17").clearContent();
        yyyymmdd = new Date(yyyymmdd.getFullYear(), yyyymmdd.getMonth(), yyyymmdd.getDate()+1);
        yyyy = yyyymmdd.getFullYear();
        mm = yyyymmdd.getMonth()+1;
        dd = yyyymmdd.getDate();
        todaySheet.getRange("J2").setValue(yyyy);
        todaySheet.getRange("K2").setValue(mm);
        todaySheet.getRange("L2").setValue(dd);
        todaySheet.getRange("P1").setValue(Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss'));
    }
    todaySheet.getRange("G11").setValue(false);
}