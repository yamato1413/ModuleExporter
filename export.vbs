Option Explicit

Const DEBUGMODE = 0


Const EXP_STANDARD = 1
Const EXP_CLASS    = 2
Const EXP_USERFORM = 4
Const EXP_OBJ      = 8
Const EXP_ALL      = 15


Const MOD_STANDARD = 1
Const MOD_CLASS    = 2
Const MOD_USERFORM = 3
Const MOD_OBJ      = 100



If DEBUGMODE = 0 Then 
    '// wscriptで起動されていたらcscriptで起動し直す
    If InStr(LCase(WScript.FullName), "wscript") Then
        Dim cmd
        cmd = "cscript /nologo " & WScript.ScriptFullName

        '// 引数のパスに空白が含まれてると意図しないところで切れてしまうので""で囲う
        Dim arg
        For Each arg In WScript.Arguments
            cmd = cmd & " " & Chr(34) & arg & Chr(34)   '// Chr(34)は「"」
        Next
        
        '// 再実行して自身は閉じる
        Const WND_SHOWNORMAL = 1
        Const WAIT_NO = 0
        CreateObject("WScript.Shell").Run cmd, WND_SHOWNORMAL, WAIT_NO
        WScript.Quit
    End If
End If


Dim fso
Set fso = Createobject("Scripting.FileSystemObject")
Dim args
Set args = CreateObject("Scripting.Dictionary")

Dim type_export
type_export = EXP_ALL

For Each arg In WScript.Arguments
    '// オプション引数をチェックして出力するモジュールを選択する
    If Left(arg, 1) = "-" And Len(arg) <> 1 Then
        type_export = 0
        If InStr(arg, "s") Then type_export = type_export Or EXP_STANDARD
        If InStr(arg, "c") Then type_export = type_export Or EXP_CLASS
        If InStr(arg, "u") Then type_export = type_export Or EXP_USERFORM
        If InStr(arg, "o") Then type_export = type_export Or EXP_OBJ
        If InStr(arg, "a") Then type_export = type_export Or EXP_ALL
    End If
    '// 引数にエクセルファイル以外が混じってたら取り除く
    If Left(fso.GetExtensionName(arg), 2) = "xl" Then args.Add args.Count, arg
Next

Dim wsh
Set wsh = CreateObject("WScript.Shell")
Dim val_trust
'// 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」の値を保存
val_trust = wsh.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM")
'// 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」をTrueに書き換え
'// これをしないとVBProjectにアクセスできない
wsh.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 1, "REG_DWORD"

For Each arg In args.Items
    '// 引数のファイルパスを絶対パスに整形
    arg = Replace(arg, "/", "\")
    If InStr(arg, ":\") = 0 And Left(arg, 2) <> "\\" Then 
        If Left(arg, 1) <> "." Then arg = ".\" & arg
        arg = GetAbsolutePathNameEx(CreateObject("WScript.Shell").CurrentDirectory, arg)
    End If

    Dim wb
    Dim Application
    WScript.Echo "起動待機中..."
    Set wb = WScript.GetObject(arg)
    Set Application = wb.Parent

    '// マクロが有効化されておらずVBAプロジェクトの取得に失敗するかどうかをチェックする
    On Error Resume Next
    Dim isErr
    isErr = false
    Dim vbps
    Set vbps = Application.VBE.VBProjects
    If Err Then isErr = True
    On Error Goto 0

    If isErr Then
        MsgBox "このブックはマクロが有効化されていません。" & vbNewLine & _ 
               "手動で開いて「コンテンツを有効化」をクリックしてください。" & vbNewLine & _
               "ブック名:" & wb.Name, vbExclamation,"エラー"
    Else

        '// アドインファイルや個人用マクロを設定していたり，既にエクセルを開いていたりすると
        '// 複数のVBProjectが取得されるので，目的のエクセルファイルのVBProjectを探す
        Dim vbp
        For Each vbp In vbps
            '// 保存していないブックはファイルパスが取れない
            On Error Resume Next
            Dim path_project
            path_project = vbp.FileName
            On Error Goto 0

            '// 目的のファイルならエクスポートを実行
            If path_project = arg Then
                WScript.Echo fso.GetFileName(path_project) & "のモジュールエクスポートを開始..."
                
                '// Sheet1とかThisworkbookとかModule1とか他のブックと名前が被る可能性が非常に高いので
                '// ブック名のフォルダを作ってそこにエクスポートする
                Dim path_folder
                path_folder = fso.GetParentFolderName(path_project) & "\" & fso.GetBaseName(path_project)
                If Not fso.FolderExists(path_folder) Then
                    fso.CreateFolder(path_folder)
                End If

                '// エクスポート
                Dim m
                For Each m In vbp.VBComponents
                    Export path_folder, m, type_export
                Next
            End If
        Next
    End If
    Set Application = Nothing
    Set wb = Nothing
Next

'// 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」をもとの設定値に戻す
wsh.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM", val_trust, "REG_DWORD"
WScript.Echo "Done!"
WScript.Sleep 1000
WScript.Quit


Function GetExtension(m)
    Select Case m.Type
        Case MOD_STANDARD: GetExtension = ".bas"
        Case MOD_CLASS:    GetExtension = ".cls"
        Case MOD_USERFORM: GetExtension = ".frm"
        Case MOD_OBJ:      GetExtension = ".cls"
    End Select
End Function

Sub Export(path_folder, module, type_export)
    '// モジュールの種類分けをした後，論理積をとってエクスポート対象かどうかを判定。
    Select Case module.Type
        Case MOD_STANDARD: If (type_export And EXP_STANDARD) = 0 Then Exit Sub
        Case MOD_CLASS:    If (type_export And EXP_CLASS)    = 0 Then Exit Sub
        Case MOD_USERFORM: If (type_export And EXP_USERFORM) = 0 Then Exit Sub
        Case MOD_OBJ:      If (type_export And EXP_OBJ)      = 0 Then Exit Sub
    End Select
    module.Export path_folder & "\" & module.Name & GetExtension(module)
    WScript.Echo module.Name & "をエクスポート" 
End Sub

'// ちゅん先生のブログより転用（VBS用に一部改変）
'// https://www.excel-chunchun.com/entry/2018/12/30/121243
Function GetAbsolutePathNameEx(ByVal basePath, ByVal RefPath)
    Dim i
    
    basePath = Replace(basePath, "/", "\")
    basePath = Left(basePath, Len(basePath) - IIf(Right(basePath, 1) = "\", 1, 0))
    
    RefPath = Replace(RefPath, "/", "\")
    
    Dim retVal
    Dim rpArr
    rpArr = Split(RefPath, "\")
    
    For i = LBound(rpArr) To UBound(rpArr)
        Select Case rpArr(i)
        Case "", "."
            If retVal = "" Then retVal = basePath
            rpArr(i) = ""
        Case ".."
            If retVal = "" Then retVal = basePath
            If InStrRev(retVal, "\") = 0 Then
                Err.Raise 8888, "GetAbsolutePathNameEx", "到達できないパスを指定しています。"
                GetAbsolutePathNameEx = ""
                Exit Function
            End If
            retVal = Left(retVal, InStrRev(retVal, "\") - 1)
            rpArr(i) = ""
        Case Else
            retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
            rpArr(i) = ""
        End Select
        '相対パス部分が空欄、.\、..\で終わった時、末尾の\が不足するので補完が必要
        If i = UBound(rpArr) Then
            If RefPath <> "" Then
                If Right(RefPath, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '連続\の消去とネットワークパス対策
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    GetAbsolutePathNameEx = retVal
End Function

Function IIf(condition, res1, res2)
    If condition Then IIf = res1 Else IIf = res2
End Function