VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "nodefunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'nodefunc VBA by.YOS G-spec

'予約文字列:文字列として次の単語を使用しないこと
'$ArraY0$, $ArraY1$... $ArraY[n]$, $ReturnSectioN$, $ReturnArgS$, $AsTypE$

'【VBA側で使用できる関数】
'setCd(CurrentPath As String)
'  カレントディレクトリの設定を行う。
'  CurrentPath:カレントディレクトリのパス(初期値:本ファイルのパス)

'setJs(FileName As String)
'  呼び出すjsファイルの設定。
'  FileName:jsファイル名(フルパス可)(初期値:本ファイルのパス)

'[ReturnArgs] = nodefn(FuncName As String, [Args As Variant]) As Variant
'  Node.jsの関数の呼び出し。
'  Node.jsで行われた標準出力はイミディエイトに出力する。
'  FuncName:関数名
'  Args:引数[省略可, 任意個数]
'  ReturnArgs:返り値[省略可]

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
    
    '引数を特殊形式の文字列に変換
    Dim i As Long
    For i = 0 To UBound(args)
        set_type args(i), 0
    Next
    
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
    
    '一時ファイルの存在チェック
    If Dir(nodeCdPath & "\" & tempfile) = "" Then
        Debug.Print "Error:結果の未取得"
        Exit Function
    End If

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
    
    '返り値として返された値の取得
    Dim return_array As Variant
    return_array = Split(return_txt, "$ReturnSectioN$")

    'Node.jsの標準出力をイミディエイトに出力
    If return_array(0) <> "" Then
        Debug.Print "> " & Replace(return_array(0), vbLf, vbLf & "> ")
    End If

    '返り値として返される値無し
    If UBound(return_array) < 1 Then
        Set nodefn = Nothing
        Exit Function
    End If
    
    Dim return_args As Variant
    return_args = return_array(1)
    '返り値の個数判別
    If InStr(return_args, "$ReturnArgS$") < 1 Then
        'キャストしてreturn
        nodefn = typecast(return_args, 0)
    '複数であれば配列として返す
    Else
        '返り値を分割
        Dim return_arg As Variant
        return_arg = Split(return_args, "$ReturnArgS$")

        For i = 0 To UBound(return_arg)
            return_arg(i) = typecast(return_arg(i), 0)
        Next
        nodefn = return_arg
    End If
End Function

'nodefuncでは実行不可
Public Function nodefunc(func As String, ParamArray args() As Variant) As Variant
    MsgBox "もしかして: nodefn"
    Set nodefunc = Nothing
End Function

'特殊形式の文字列に変換
'args:引数, arrNo:配列の階層数
Private Sub set_type(ByRef arg As Variant, arrNo As Long)
    '配列であれば特殊文字で結合
    If IsArray(arg) Then
        Dim i As Long
        For i = 0 To UBound(arg)
            '配列の中身も同様に(再帰)
            set_type arg(i), arrNo + 1
        Next
        arg = Join(arg, "$ArraY" & arrNo & "$")
    Else
        '文字列に型名を結合して返却
        arg = Join(Array(arg, TypeName(arg)), "$AsTypE$")
    End If
End Sub

'返り値のキャスト
'args:引数, arrNo:配列の階層数
Private Function typecast(arg As Variant, arrNo As Long) As Variant
    If InStr(arg, "$ArraY" & arrNo & "$") Then
        '配列であれば分割
        Dim argArr As Variant
        argArr = Split(arg, "$ArraY" & arrNo & "$")

        'Splitした中身はVariant型でないため新たに定義
        ReDim argArrV(UBound(argArr)) As Variant
        
        '配列の中身もキャスト
        Dim i As Long
        For i = 0 To UBound(argArr)
            '再帰
            argArrV(i) = typecast(argArr(i), arrNo + 1)
        Next
        typecast = argArrV
    Else
        '各種キャスト
        Dim args As Variant
        args = Split(arg, "$AsTypE$")
        arg = args(0)           '値

        Dim vbatype As String   '型情報
        vbatype = args(1)

        Select Case vbatype
            Case "Double":
                typecast = CDbl(arg)
            Case "String":
                typecast = CStr(arg)
            Case "Boolean":
                typecast = CBool(arg)
            Case "Date":
                typecast = CDate(arg)
            Case "Nothing":
                typecast = Nothing
            Case "Empty":
                typecast = Empty
            Case Else
                typecast = CStr(arg & " As ?")
        End Select
    End If
    
End Function

