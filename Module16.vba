Module: Module16
Option Explicit ' 未定義の変数は使用できないように


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Clickしました。"

    Application.VBE.MainWindow.Visible = True
    
'    MsgBox vbInformation

'    MsgBox vbExclamation

'    Call RunGitBashCommands


'    MsgBox Month(ThisWorkbook.sheetS("手順").Range("E" & UNITROW))
'    Exit Sub

    Dim Command As String
'    Command = "cd /c/Users/kenic/Documents/operation_log_NEW" & ";" & _
'               "./excelgrep_by_XMLparse.sh SACLA/2025_10.xlsm '$|引渡' '$|引き渡' '$|波長変更依頼' '$|ユニット' '$|利用終了' '$|運転終了'"
    Command = "cd /c/Users/kenic/Documents/operation_log_NEW" & ";" & _
               "./excelgrep_by_XMLparse.sh SACLA/2025_" & Month(ThisWorkbook.sheetS("手順").Range("E" & UNITROW)) & ".xlsm '$|引渡' '$|引き渡' '$|波長変更依頼' '$|ユニット' '$|利用終了' '$|運転終了'"
    ExecuteGitBashCommand Command
    
End Sub












