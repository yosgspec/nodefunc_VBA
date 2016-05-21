# nodefunc VBA
  ExcelVBAからNode.jsを呼び出すためのモジュール軍です。
    
##がいよう
VBAからNode.jsの関数を呼んで返り値を得ることが出来ます。
多分機能としてはそれくらい。
あとはそこそこ遅いから多用厳禁。
細かい使用方法はソース内コメント参照。
    
##ファイル構成
*Main.bas
  VBAの呼び出す方のテスト用ファイル。
標準モジュールとして登録して使用。
*main.js
  VBAから呼び出されるテスト用Node.jsファイル。
*nodefunc.cls
  Node.jsをVBAから呼び出すために必要なVBA用クラスファイル。
*nodefuncVBA.js
  Node.jsをVBAから呼び出すために必要なNode.js用モジュール。
*nodefunclite.cls
*nodefuncVBAlite.js
  上から型変換を除いて機能を簡略化したもの。
  標準出力がそのまま返り値として得られます。
#テスト
##VBA側 Main.bas
	Option Explicit

	Dim nf As New nodefunc
	
	Sub Main()
	    '初期定義
	    'カレントディレクトリ
	    nf.setCd ThisWorkbook.Path
	    '呼び出しjsファイル
	    nf.setJs "main.js"
	
	    '再帰関数
	    Debug.Print nf.nodefn("fact", 1, 10)
	
	    '配列
	    Dim Arr As Variant
	    Arr = nf.nodefn("returnData", Array(0, Array(1, Array(2, Array(3, "ふが")))))
	    Debug.Print Arr(1)(1)(1)(1)
	
	    'Date型
	    Debug.Print nf.nodefn("returnData", Now)
	End Sub

##Node.js側 main.js
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

##実行結果
	 3628800 
	> [ 0, [ 1, [ 2, [Object] ] ] ]
	> 
	ふが
	> Sat May 21 2016 23:54:42 GMT+0900 (東京 (標準時))
	> 
	2016/05/21 23:54:42 

