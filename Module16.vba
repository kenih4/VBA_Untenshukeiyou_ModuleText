Module: Module16
Option Explicit ' 未定義の変数は使用できないように


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Clickしました。"

    Application.VBE.MainWindow.Visible = True
    
    MsgBox vbInformation

    MsgBox vbExclamation


End Sub













