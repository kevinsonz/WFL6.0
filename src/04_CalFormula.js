// Cal数式
function calFormula(){
    const nullFormula = calSheet.getRange(beginRow_Cal,1).getValue() !== 0;
    if(nullFormula){
        calSheet.getRange('A'+(beginRow_Cal+1)+':A').setFormula('=counta(F'+(beginRow_Cal+1)+':NF'+(beginRow_Cal+1)+')');
    }
}