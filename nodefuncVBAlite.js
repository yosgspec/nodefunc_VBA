/* nodefunc VBA Lite by.YOS G-spec

【Node.js側で使用できる関数】
FuncStr=func():string
[eval(func()):any]
  VBAから与えられた引数を関数の形に変換する。
  通常はevalと組み合わせて関数を呼び出させる。
  FuncStr:関数の文字列

return_VBA(ReturnArg)
  VBAに値を返す。
  ReturnArg:返り値
*/

module.exports=new function(){
	//関数の定義
	this.func=function(){
		//コマンドライン引数の除去 0:node, 1:ファイル名
		process.argv.shift();
		process.argv.shift();
		//関数名の取得+コマンドライン引数から除去
		var runfunc=process.argv.shift()+"(";
		var args="";

		//関数の引数をコンマ区切り
		if(0<process.argv.length){
			for(var i=0;i<process.argv.length;i++){
				args+=",'"+process.argv[i]+"'";
			}
			//頭のコンマを除去
			args=args.substr(1);
		}
		runfunc+=args+")";
		//関数の"文字列"を返還(eval必須)
		return runfunc;
	};

	//VBAへの返り値
	this.return_VBA=function(arg){
		console.log(arg);
		process.exit();
	};
};
