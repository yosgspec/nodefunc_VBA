VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "nodefuncVBAlite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'nodefunc VBA Lite by.YOS G-spec

'【VBA側で使用できる関数】
'setCd(CurrentPath As String)
'  カレントディレクトリの設定を行う。
'  CurrentPath:カレントディレクトリのパス(初期値:本ファイルのパス)

'setJs(FileName As String)
'  呼び出すjsファイルの設定。
'  FileName:jsファイル名(フルパス可)(初期値:本ファイルのパス)

'[ReturnArgs] = nodefn(FuncName As String, [Args As Variant]) As Variant
'  Node.jsの関数の呼び出し。
'  FuncName:関数名
'  Args:引数[省略可, 任意個数]
'  ReturnArg:返り値[省略可]

Option Explicit

Private nodeCdPath As Variant   'カレントフォルダ
Private requireJs As Variant    '呼び出すjsファイル

'カレントフォルダの設定
Public Sub setCd(cdpath As String)
    nodeCdPath = cdpath
End Sub

'jsファイルの設定
Public Sub setJs(jsfile As String)
    requireJs = jsfile
End Sub

'nodefunc本体
'func:呼び出す関数名, args():関数の引数(個数不定)
'返り値:関数の返り値
Public Function nodefn(func As String, ParamArray args() As Variant) As Variant
    'カレントフォルダの取得
    nodeCdPath = IIf(nodeCdPath = Empty, ThisWorkbook.path, nodeCdPath)
    'jsファイルの取得
    requireJs = IIf(requireJs = Empty, "main.js", requireJs)
    '一時ファイル(処理結果)
    Const tempfile As String = "Temp$GIy0qPan"

    'コマンドライン初期設定
    Dim WSH As Object
    Set WSH = CreateObject("WScript.Shell")
    WSH.CurrentDirectory = nodeCdPath

    Dim Command As String '実行コマンド

    ' C:\> node main.js (args)
    Command = Join(Array( _
        "node """, requireJs, """ ", func, _
        IIf(0 <= UBound(args), _
            " """ & Join(args, """ """) & """", _
            "") _
    ), "")

    'コマンド実行
    'cmd /c (Command) > (tempfile) 2>&1
    WSH.Run Join(Array( _
        "%ComSpec% /c ", Command, " > ", tempfile, " 2>&1" _
    ), ""), 0, True

    '処理結果の取得
    Dim return_txt As String    '処理結果
    return_txt = ""

    With CreateObject("ADODB.Stream")
        .Open
        .Charset = "utf-8"
        .LoadFromFile nodeCdPath & "\" & tempfile
        return_txt = .ReadText
        .Close
    End With

    '結果が空なら終了
    If Len(return_txt) < 2 Then
        Set nodefn = Nothing
        Exit Function
    End If

    '最後のvbLfを除去
    return_txt = Mid(return_txt, 1, Len(return_txt) - 1)

    '一時ファイルの削除
    Kill nodeCdPath & "\" & tempfile

    nodefn = return_txt
End Function
