// Cal数式
function calFormula(){
    calSheet.getRange('A'+(beginRow_Cal+1)+':A').setFormula('=counta(F'+(beginRow_Cal+1)+':NF'+(beginRow_Cal+1)+')');
}