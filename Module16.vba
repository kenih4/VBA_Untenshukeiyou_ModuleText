Module: Module16
Option Explicit ' ����`�̕ϐ��͎g�p�ł��Ȃ��悤��


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Click���܂����B"

    Application.VBE.MainWindow.Visible = True
    
    MsgBox vbInformation

    MsgBox vbExclamation


End Sub













