# nodefunc-VBA
ExcelVBAからNode.jsを呼び出すためのモジュール軍です。

##できること


##ファイル構成
###Main.bas
VBAの呼び出す方のテスト用ファイル。
標準モジュールとして登録して使用。
###main.js
VBAから呼び出されるテスト用Node.jsファイル。
###nodefunc.cls
Node.jsをVBAから呼び出すために必要なVBA用クラスファイル。
###nodefuncVBA.js
Node.jsをVBAから呼び出すために必要なNode.js用モジュール。
###nodefunclite.cls
###nodefuncVBAlite.js
上から型変換を除いて機能を簡略化したもの。
標準出力がそのまま返り値として得られます。
