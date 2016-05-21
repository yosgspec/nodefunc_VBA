# nodefunc VBA
  ExcelVBAからNode.jsを呼び出すためのモジュール軍です。
    
  ##がいよう
    VBAからNode.jsの関数を呼んで返り値を得ることが出来ます。
    多分機能としてはそれくらい。
    あとはそこそこ遅いから多用厳禁。
    細かい使用方法はソース内コメント参照。
    
  ##ファイル構成
    *ain.bas
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
