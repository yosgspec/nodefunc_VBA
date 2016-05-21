Attribute VB_Name = "Main"
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

