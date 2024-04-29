// GTD行追加

function addRowGTD(){
    if(statusGTD === 'Add' && endCol_GTD === 10){
        gtdSheet.getRange(beginRow_GTD-1,1,endRow_GTD,10).setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        hiddenGTD('call');
        const maxNo = Math.max.apply(null,gtdNo);
        let addNo = 0;
        for(let i=0;i<((endRow_GTD)-(beginRow_GTD)+1);i++){
            if(gtdData[i][9] === ''){
              addNo = addNo + 1 ;
              gtdSheet.getRange(beginRow_GTD+i,1).setValue(maxNo + addNo);
              gtdSheet.getRange(beginRow_GTD+i,6).setValue('－');
              gtdSheet.getRange(beginRow_GTD+i,7).setValue('－');
              gtdSheet.getRange(beginRow_GTD+i,8).setFormula('=iferror(if(or(I'+(beginRow_GTD+i)+'="完了",I'+(beginRow_GTD+i)+'="保留",I'+(beginRow_GTD+i)+'="中止"),9,ifs(and(E'+(beginRow_GTD+i)+'=E$1,F'+(beginRow_GTD+i)+'=F$1),1,and(E'+(beginRow_GTD+i)+'=E$1,F'+(beginRow_GTD+i)+'<>F$1),2,and(E'+(beginRow_GTD+i)+'<>E$1,F'+(beginRow_GTD+i)+'=F$1),3,and(E'+(beginRow_GTD+i)+'<>E$1,F'+(beginRow_GTD+i)+'<>F$1),4)),9)');
              gtdSheet.getRange(beginRow_GTD+i,9).setValue('未着');
              gtdSheet.getRange(beginRow_GTD+i,10).setFormula('=countifs(A$'+beginRow_GTD+':A,A'+(beginRow_GTD+i)+')');
            }
        }
    }
}