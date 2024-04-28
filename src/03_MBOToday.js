// MBO：TodayシートのDone等処理を実行

function doneMBO(){
    const doneCheck = todaySheet.getRange("G11").getValue();
    const errorCheck = todaySheet.getRange("A1").getValue() === 'OK';
    const runFlag = doneCheck && errorCheck;
    const todayRow = todaySheet.getRange("P2").getValue();
    let yyyy = todaySheet.getRange("J2").getValue();
    let mm = todaySheet.getRange("K2").getValue()-1;
    let dd = todaySheet.getRange("L2").getValue();
    let yyyymmdd = new Date(yyyy,mm,dd);
    if(runFlag){
        const zValue = todaySheet.getRange("G13").getValue();
        const wfleValues = todaySheet.getRange("C14:H17").getValues().flat();
        mboSheet.getRange("G"+todayRow).setValue(zValue);
        mboSheet.getRange("I"+todayRow+":AF"+todayRow).setValues(wfleValues);
        mboSheet.getRange("L"+todayRow).setFormula(mboSheet.getRange("L3").getFormula());
        mboSheet.getRange("N"+todayRow).setFormula(mboSheet.getRange("N3").getFormula());
        mboSheet.getRange("R"+todayRow).setFormula(mboSheet.getRange("R3").getFormula());
        mboSheet.getRange("T"+todayRow).setFormula(mboSheet.getRange("T3").getFormula());
        mboSheet.getRange("X"+todayRow).setFormula(mboSheet.getRange("X3").getFormula());
        mboSheet.getRange("Z"+todayRow).setFormula(mboSheet.getRange("Z3").getFormula());
        mboSheet.getRange("AD"+todayRow).setFormula(mboSheet.getRange("AD3").getFormula());
        mboSheet.getRange("AF"+todayRow).setFormula(mboSheet.getRange("AF3").getFormula());
        yyyymmdd = yyyymmdd+1;
        yyyy = yyyymmdd.getFullYear();
        mm = yyyymmdd.getMonth()+1;
        dd = yyyymmdd.getDate();
        todaySheet.getRange("J2").setValue(yyyy);
        todaySheet.getRange("K2").setValue(mm);
        todaySheet.getRange("L2").setValue(dd);
    }
    todaySheet.getRange("G11").setValue(false);
}