Module: Module16
Option Explicit ' ����`�̕ϐ��͎g�p�ł��Ȃ��悤��


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Click���܂����B"

    Application.VBE.MainWindow.Visible = True
    
'    MsgBox vbInformation

'    MsgBox vbExclamation

'    Call RunGitBashCommands


'    MsgBox Month(ThisWorkbook.sheetS("�菇").Range("E" & UNITROW))
'    Exit Sub

    Dim Command As String
'    Command = "cd /c/Users/kenic/Documents/operation_log_NEW" & ";" & _
'               "./excelgrep_by_XMLparse.sh SACLA/2025_10.xlsm '$|���n' '$|�����n' '$|�g���ύX�˗�' '$|���j�b�g' '$|���p�I��' '$|�^�]�I��'"
    Command = "cd /c/Users/kenic/Documents/operation_log_NEW" & ";" & _
               "./excelgrep_by_XMLparse.sh SACLA/2025_" & Month(ThisWorkbook.sheetS("�菇").Range("E" & UNITROW)) & ".xlsm '$|���n' '$|�����n' '$|�g���ύX�˗�' '$|���j�b�g' '$|���p�I��' '$|�^�]�I��'"
    ExecuteGitBashCommand Command
    
End Sub












