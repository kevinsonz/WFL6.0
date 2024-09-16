// Cal数式
function calFormula(){
    const nullFormula = calSheet.getRange(beginRow_Cal,1).getValue() !== 0;
    if(nullFormula){
        calSheet.getRange('A'+(beginRow_Cal+1)+':A').setFormula('=iferror(if(counta($E'+(beginRow_Cal+1)+':$G'+(beginRow_Cal+1)+')=0,1,if($NM'+(beginRow_Cal+1)+'=1,ifs(and($D$5=false,$D$6=false),1,and($D$5=true,$D$6=false),$NV'+(beginRow_Cal+1)+',and($D$5=false,$D$6=true),$NW'+(beginRow_Cal+1)+',and($D$5=true,$D$6=true),$NV'+(beginRow_Cal+1)+'*$NW'+(beginRow_Cal+1)+'),0)),0)');
        calSheet.getRange('NJ'+(beginRow_Cal+1)+':NJ').setFormula('=iferror(if(countifs($H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',NJ$11)=1,match(NJ$11,$H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',0),if(countifs($H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',NJ$11)>1,-1,0)),-1)');
        calSheet.getRange('NK'+(beginRow_Cal+1)+':NK').setFormula('=iferror(if(countifs($H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',NK$11)=1,match(NK$11,$H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',0),if(countifs($H'+(beginRow_Cal+1)+':$NI'+(beginRow_Cal+1)+',NK$11)>1,-1,0)),-1)');
        calSheet.getRange('NL'+(beginRow_Cal+1)+':NL').setFormula('=if(and(NJ'+(beginRow_Cal+1)+'>0,NK'+(beginRow_Cal+1)+'>0,NJ'+(beginRow_Cal+1)+'<NK'+(beginRow_Cal+1)+'),"A",if(and(NJ'+(beginRow_Cal+1)+'>0,NK'+(beginRow_Cal+1)+'>0,NJ'+(beginRow_Cal+1)+'>NK'+(beginRow_Cal+1)+'),"B","-"))');
        calSheet.getRange('NM'+(beginRow_Cal+1)+':NM').setFormula('=counta(E'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')');
        calSheet.getRange('NN'+(beginRow_Cal+1)+':NN').setFormula('=if(counta(C'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')>0,offset(NN'+(beginRow_Cal+1)+',-1,0)+if(E'+(beginRow_Cal+1)+'<>"",1,0),0)');
        calSheet.getRange('NO'+(beginRow_Cal+1)+':NO').setFormula('=if(E'+(beginRow_Cal+1)+'<>"",countifs(NT:NT,1,NN:NN,NN'+(beginRow_Cal+1)+'),0)');
        calSheet.getRange('NP'+(beginRow_Cal+1)+':NP').setFormula('=if(and(counta(C'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')>0,NN'+(beginRow_Cal+1)+'<>0),if(E'+(beginRow_Cal+1)+'<>"",if(B'+(beginRow_Cal+1)+',1,0),offset(NP'+(beginRow_Cal+1)+',-1,0)),0)');
        calSheet.getRange('NQ'+(beginRow_Cal+1)+':NQ').setFormula('=if(counta(C'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')>0,if(NN'+(beginRow_Cal+1)+'<>offset(NN'+(beginRow_Cal+1)+',-1,0),0,offset(NQ'+(beginRow_Cal+1)+',-1,0)+if(F'+(beginRow_Cal+1)+'<>"",1,0)),0)');
        calSheet.getRange('NR'+(beginRow_Cal+1)+':NR').setFormula('=if(F'+(beginRow_Cal+1)+'<>"",countifs(NT:NT,1,NN:NN,NN'+(beginRow_Cal+1)+',NQ:NQ,NQ'+(beginRow_Cal+1)+'),0)');
        calSheet.getRange('NS'+(beginRow_Cal+1)+':NS').setFormula('=if(counta(C'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')>0,if(NP'+(beginRow_Cal+1)+'=1,1,if(and(NQ'+(beginRow_Cal+1)+'<>0,F'+(beginRow_Cal+1)+'<>""),if(B'+(beginRow_Cal+1)+',1,0),if(NN'+(beginRow_Cal+1)+'=offset(NN'+(beginRow_Cal+1)+',-1,0),offset(NS'+(beginRow_Cal+1)+',-1,0),0))),0)');
        calSheet.getRange('NT'+(beginRow_Cal+1)+':NT').setFormula('=if(counta($E'+(beginRow_Cal+1)+':$G'+(beginRow_Cal+1)+')=1,if($G'+(beginRow_Cal+1)+'<>"",if(index($H'+(beginRow_Cal+1)+':$NT'+(beginRow_Cal+1)+',1,match(today(),$H$3:$NT$3,0))<>"",1,0),0),0)');
        calSheet.getRange('NU'+(beginRow_Cal+1)+':NU').setFormula('=if(counta(C'+(beginRow_Cal+1)+':G'+(beginRow_Cal+1)+')>0,if(or(NP'+(beginRow_Cal+1)+'=1,NS'+(beginRow_Cal+1)+'=1),1,if(G'+(beginRow_Cal+1)+'<>"",if(B'+(beginRow_Cal+1)+',1,0),0)),0)');
        calSheet.getRange('NV'+(beginRow_Cal+1)+':NV').setFormula('=if(or(NO'+(beginRow_Cal+1)+'=1,NR'+(beginRow_Cal+1)+',NT'+(beginRow_Cal+1)+'),1,0)');
        calSheet.getRange('NW'+(beginRow_Cal+1)+':NW').setFormula('=if(or(NP'+(beginRow_Cal+1)+'=1,NS'+(beginRow_Cal+1)+'=1,NU'+(beginRow_Cal+1)+'=1),0,1)');
    }
}