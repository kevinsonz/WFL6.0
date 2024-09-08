// Cal数式
function calFormula(){
    const nullFormula = calSheet.getRange(beginRow_Cal,1).getValue() !== 0;
    if(nullFormula){
        calSheet.getRange('A'+(beginRow_Cal+1)+':A').setFormula('=if(counta(E'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')=1,if(G'+(beginRow_Cal+1)+'<>"",counta(H'+(beginRow_Cal+1)+':NI'+(beginRow_Cal+1)+'),1),0)');
    }
}