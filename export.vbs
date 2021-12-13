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
    '// wscript�ŋN������Ă�����cscript�ŋN��������
    If InStr(LCase(WScript.FullName), "wscript") Then
        Dim cmd
        cmd = "cscript /nologo " & WScript.ScriptFullName

        '// �����̃p�X�ɋ󔒂��܂܂�Ă�ƈӐ}���Ȃ��Ƃ���Ő؂�Ă��܂��̂�""�ň͂�
        Dim arg
        For Each arg In WScript.Arguments
            cmd = cmd & " " & Chr(34) & arg & Chr(34)   '// Chr(34)�́u"�v
        Next
        
        '// �Ď��s���Ď��g�͕���
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
    '// �I�v�V�����������`�F�b�N���ďo�͂��郂�W���[����I������
    If Left(arg, 1) = "-" And Len(arg) <> 1 Then
        type_export = 0
        If InStr(arg, "s") Then type_export = type_export Or EXP_STANDARD
        If InStr(arg, "c") Then type_export = type_export Or EXP_CLASS
        If InStr(arg, "u") Then type_export = type_export Or EXP_USERFORM
        If InStr(arg, "o") Then type_export = type_export Or EXP_OBJ
        If InStr(arg, "a") Then type_export = type_export Or EXP_ALL
    End If
    '// �����ɃG�N�Z���t�@�C���ȊO���������Ă����菜��
    If Left(fso.GetExtensionName(arg), 2) = "xl" Then args.Add args.Count, arg
Next

Dim wsh
Set wsh = CreateObject("WScript.Shell")
Dim val_trust
'// �uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�̒l��ۑ�
val_trust = wsh.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM")
'// �uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v��True�ɏ�������
'// ��������Ȃ���VBProject�ɃA�N�Z�X�ł��Ȃ�
wsh.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 1, "REG_DWORD"

For Each arg In args.Items
    '// �����̃t�@�C���p�X���΃p�X�ɐ��`
    arg = Replace(arg, "/", "\")
    If InStr(arg, ":\") = 0 And Left(arg, 2) <> "\\" Then 
        If Left(arg, 1) <> "." Then arg = ".\" & arg
        arg = GetAbsolutePathNameEx(CreateObject("WScript.Shell").CurrentDirectory, arg)
    End If

    Dim wb
    Dim Application
    WScript.Echo "�N���ҋ@��..."
    Set wb = WScript.GetObject(arg)
    Set Application = wb.Parent

    '// �}�N�����L��������Ă��炸VBA�v���W�F�N�g�̎擾�Ɏ��s���邩�ǂ������`�F�b�N����
    On Error Resume Next
    Dim isErr
    isErr = false
    Dim vbps
    Set vbps = Application.VBE.VBProjects
    If Err Then isErr = True
    On Error Goto 0

    If isErr Then
        MsgBox "���̃u�b�N�̓}�N�����L��������Ă��܂���B" & vbNewLine & _ 
               "�蓮�ŊJ���āu�R���e���c��L�����v���N���b�N���Ă��������B" & vbNewLine & _
               "�u�b�N��:" & wb.Name, vbExclamation,"�G���["
    Else

        '// �A�h�C���t�@�C����l�p�}�N����ݒ肵�Ă�����C���ɃG�N�Z�����J���Ă����肷���
        '// ������VBProject���擾�����̂ŁC�ړI�̃G�N�Z���t�@�C����VBProject��T��
        Dim vbp
        For Each vbp In vbps
            '// �ۑ����Ă��Ȃ��u�b�N�̓t�@�C���p�X�����Ȃ�
            On Error Resume Next
            Dim path_project
            path_project = vbp.FileName
            On Error Goto 0

            '// �ړI�̃t�@�C���Ȃ�G�N�X�|�[�g�����s
            If path_project = arg Then
                WScript.Echo fso.GetFileName(path_project) & "�̃��W���[���G�N�X�|�[�g���J�n..."
                
                '// Sheet1�Ƃ�Thisworkbook�Ƃ�Module1�Ƃ����̃u�b�N�Ɩ��O�����\�������ɍ����̂�
                '// �u�b�N���̃t�H���_������Ă����ɃG�N�X�|�[�g����
                Dim path_folder
                path_folder = fso.GetParentFolderName(path_project) & "\" & fso.GetBaseName(path_project)
                If Not fso.FolderExists(path_folder) Then
                    fso.CreateFolder(path_folder)
                End If

                '// �G�N�X�|�[�g
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

'// �uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�����Ƃ̐ݒ�l�ɖ߂�
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
    '// ���W���[���̎�ޕ�����������C�_���ς��Ƃ��ăG�N�X�|�[�g�Ώۂ��ǂ����𔻒�B
    Select Case module.Type
        Case MOD_STANDARD: If (type_export And EXP_STANDARD) = 0 Then Exit Sub
        Case MOD_CLASS:    If (type_export And EXP_CLASS)    = 0 Then Exit Sub
        Case MOD_USERFORM: If (type_export And EXP_USERFORM) = 0 Then Exit Sub
        Case MOD_OBJ:      If (type_export And EXP_OBJ)      = 0 Then Exit Sub
    End Select
    module.Export path_folder & "\" & module.Name & GetExtension(module)
    WScript.Echo module.Name & "���G�N�X�|�[�g" 
End Sub

'// �����搶�̃u���O���]�p�iVBS�p�Ɉꕔ���ρj
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
                Err.Raise 8888, "GetAbsolutePathNameEx", "���B�ł��Ȃ��p�X���w�肵�Ă��܂��B"
                GetAbsolutePathNameEx = ""
                Exit Function
            End If
            retVal = Left(retVal, InStrRev(retVal, "\") - 1)
            rpArr(i) = ""
        Case Else
            retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
            rpArr(i) = ""
        End Select
        '���΃p�X�������󗓁A.\�A..\�ŏI��������A������\���s������̂ŕ⊮���K�v
        If i = UBound(rpArr) Then
            If RefPath <> "" Then
                If Right(RefPath, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '�A��\�̏����ƃl�b�g���[�N�p�X�΍�
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    GetAbsolutePathNameEx = retVal
End Function

Function IIf(condition, res1, res2)
    If condition Then IIf = res1 Else IIf = res2
End Function