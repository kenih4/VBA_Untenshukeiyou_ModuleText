Module: Module12
Option Explicit

Function RunPythonScript(scriptPath As String, workDir As String) As Boolean
    On Error GoTo ErrorHandler
    RunPythonScript = False
'    Dim scriptPath As String
    Dim Command As String
    Dim buttonName As String
'    Const pythonExe As String = "python"
    Dim pythonExe As String
    pythonExe = "python"
    Debug.Print "Debug   " & workDir
    
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        MsgBox "���̃}�N���̓V�[�g��̃{�^������̂ݎ��s���Ă��������B" & vbCrLf & "�I�����܂��B", Buttons:=vbCritical
'        End
'    End If
'
'    If buttonName = "�{�^�� 8" Then
'        scriptPath = "getGunHvOffTime_LOCALTEST.py"
'    ElseIf buttonName = "�{�^�� 9" Then
'        scriptPath = "getBlFaultSummary_LOCALTEST.py bl2"
'    ElseIf buttonName = "�{�^�� 10" Then
'        scriptPath = "getBlFaultSummary_LOCALTEST.py bl3"
'    Else
'        MsgBox "�ُ�ł��B�I�����܂��B" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    
    MsgBox "python " & scriptPath & "��" & vbCrLf & "���s���܂��B", Buttons:=vbInformation
    
    ' �R�}���h��g�ݗ��āF�܂��w��t�H���_�Ɉړ����A���̌�Python�����s
    Command = "cmd.exe /c cd /d " & Chr(34) & workDir & Chr(34) & " && " & pythonExe & " " & workDir & "\" & scriptPath
    
    'Shell command, vbNormalFocus ' Shell�֐���Python�X�N���v�g�����s �I����҂��Ȃ�
    
    Dim shell As Object
    Dim exitCode As Long
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.RUN(Command, vbMaximizedFocus, True)   ' WScript.Shell��Run���\�b�h�ŃR�}���h�����s���A�I����҂�
    If exitCode = 0 Then
        RunPythonScript = True
        Call Fin("Python�X�N���v�g������ɏI�����܂����B " & vbCrLf & "[" & scriptPath & "]", 1)
    Else
        Call Fin("Python�X�N���v�g���G���[�R�[�h " & exitCode & " �ŏI�����܂����B", 3)
    End If
    
    
    Exit Function ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    Call Fin("�G���[�ł��B���e�́@ " & Err.Description, 3)
    Exit Function
    
End Function

