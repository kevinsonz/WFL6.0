// Cal数式
function calFormula(){
    const nullFormula = calSheet.getRange(beginRow_Cal,1).getValue() !== 0;
    if(nullFormula){
        calSheet.getRange('A'+(beginRow_Cal+1)+':A').setFormula('=if(counta(E'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')=1,if(G'+(beginRow_Cal+1)+'<>"",counta(K'+(beginRow_Cal+1)+':NL'+(beginRow_Cal+1)+'),1),0)');
        calSheet.getRange('H'+(beginRow_Cal+1)+':H').setFormula('=iferror(if(countifs($K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',H$11)=1,match(H$11,$K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',0),if(countifs($K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',H$11)>1,-1,0)),-1)');
        calSheet.getRange('I'+(beginRow_Cal+1)+':I').setFormula('=iferror(if(countifs($K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',I$11)=1,match(I$11,$K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',0),if(countifs($K'+(beginRow_Cal+1)+':$NL'+(beginRow_Cal+1)+',I$11)>1,-1,0)),-1)');
        calSheet.getRange('J'+(beginRow_Cal+1)+':J').setFormula('=if(and(H'+(beginRow_Cal+1)+'>0,I'+(beginRow_Cal+1)+'>0,H'+(beginRow_Cal+1)+'<I'+(beginRow_Cal+1)+'),"A",if(and(H'+(beginRow_Cal+1)+'>0,I'+(beginRow_Cal+1)+'>0,H'+(beginRow_Cal+1)+'>I'+(beginRow_Cal+1)+'),"B","-"))');
    }
}