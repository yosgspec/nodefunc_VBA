/* nodefunc VBA by.YOS G-spec

【予約文字列】文字列として次の単語を使用しないこと
$ArraY0$, $ArraY1$... $ArraY[n]$, $ReturnSectioN$, $ReturnArgS$, $AsTypE$

【Node.js側で使用できる関数】
FuncStr=func():string
[eval(func()):any]
  VBAから与えられた引数を関数の形に変換する。
  通常はevalと組み合わせて関数を呼び出させる。
  FuncStr:関数の文字列

return_VBA([ReturnArgs])
  VBAに値を返す。
  ReturnArgs:返り値[任意の個数]
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

		//関数の引数をキャストしつつコンマ区切り
		if(0<process.argv.length){
			for(var i=0;i<process.argv.length;i++){
				args+=","+typecast(process.argv[i],0);
			}
			//頭のコンマを除去
			args=args.substr(1);
		}
		runfunc+=args+")";
		//関数の"文字列"を返還(eval必須)
		return runfunc;
	};

	//VBAへの返り値
	this.return_VBA=function(/*arguments*/){
		//特殊形式の文字列(型付き)
		var typeargs=[];
		//全ての引数を特殊形式の文字列に変換
		for(var i=0;i<arguments.length;i++){
			typeargs.push(typeRecast(arguments[i],0))
		}
		//返り値領域であることを指示して出力
		var return_end="$ReturnSectioN$"+typeargs.join("$ReturnArgS$");
		console.log(return_end);
		//処理をVBAに返す。
		process.exit();
	};

	//VBAからの値のキャスト
	//arg:引数, arrNo:配列の階層数
	function typecast(arg,arrNo){
		if(arg.indexOf("$ArraY"+arrNo+"$")!=-1){
			//配列であれば分割
			var argArr=arg.split("$ArraY"+arrNo+"$");

			//配列の中身も同様にキャスト
			for(var i=0;i<argArr.length;i++){
				//再帰
				argArr[i]=typecast(argArr[i],arrNo+1);
			}
			return "["+argArr+"]";
		}
		else{
			//各種キャスト
			var args=arg.split("$AsTypE$");
			arg=args[0];			//値
			var jstype=args.pop();	//型情報

			switch(jstype){
				case "Byte":
				case "Integer":
				case "Long":
				case "Decimal":
				case "Currency":
					return "parseInt('"+arg+"',10)";
				case "Single":
				case "Double":
					return "parseFloat('"+arg+"')";
				case "Boolean":
					return arg=="True"?true:false;
				case "Date":
					return "new Date('"+arg+"')";
				case "Empty":
					return "undefined";
				case "Null":
				case "Nothing":
					return "null";
				case "Char":
				case "String":
				case "Variant":
				default:
					return "'"+arg+"'";
			}
		}
	}

	//VBAに返す値を特殊形式の文字列に変換
	//arg:返り値, arrNo:配列の階層数
	function typeRecast(arg,arrNo){
		//型判別
		var toString = Object.prototype.toString;
		var vbatype=toString.call(arg).slice(8,-1);

		//各型に合わせてVBAの型名を追記
		switch(vbatype){
			case "Array":
				for(var i=0;i<arg.length;i++){
					//配列の中身も同様に(再帰)
					arg[i]=typeRecast(arg[i],arrNo+1);
				}
				return arg.join("$ArraY"+arrNo+"$");
			case "Number"://NaN除く
				if(isNaN(arg)){
					console.log("do");
					return arg+"$AsTypE$String";
				}
				return arg+"$AsTypE$Double";
			case "String":
			case "RegExp":
				return arg+"$AsTypE$String";
			case "Boolean":
				return arg+"$AsTypE$Boolean";
			case "Date":
				return [
						arg.getFullYear(),
						arg.getMonth()+1,
						arg.getDate()
					].join("/")+" "+
					arg.toLocaleTimeString()+
					"$AsTypE$Date";
			case "Null":
				return arg+"$AsTypE$Nothing";
			case "Underfines":
				return arg+"$AsTypE$Empty";
			case "Object":
				return "[Object]$AsTypE$String";
			case "Function":
				return "[Function]$AsTypE$String";
			default:
				return arg+"$AsTypE$String";
		}
	}
};
