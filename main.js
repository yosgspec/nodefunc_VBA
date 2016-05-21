//モジュール呼び出し
var nf=require("./nodefuncVBA");
//呼び出し関数
eval(nf.func());

//階乗
function fact(x,i){
	if(0<i){
		nf.return_VBA(fact(x*i,i-1));
	}
	else{
		nf.return_VBA(x);
	}
}

//そのまま返す関数
function returnData(x){
	console.log(x);
	nf.return_VBA(x);
}
