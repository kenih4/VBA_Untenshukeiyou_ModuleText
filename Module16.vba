Module: Module16
Option Explicit ' 未定義の変数は使用できないように


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Clickしました。"

    Application.VBE.MainWindow.Visible = True
    
    
    Dim BNAME_SHUKEI As String
    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
   
    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Debug.Print "Debug<<<   Before  Function OpenBook(" & BNAME_SHUKEI & ")"
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False) ' フルパスを指定
    Debug.Print "Debug>>>   After  Function OpenBook(" & BNAME_SHUKEI & ")"
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    
    Call CheckFormulaCells(wb_SHUKEI.Worksheets("運転予定時間"))
    
'    Dim pattern As String
'    pattern = "^[1-9][0-9]*-[1-9][0-9]*$" ' パターン: 先頭(^)から、1-9で始まる数字の塊、ハイフン、1-9で始まる数字の塊、末尾($)まで
'    If Not IsValidFormat(ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW), pattern) Then
'        Call CMsg("Err セル [" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW).Value & "] の値が ユニットの形式（例: 2-11）ではありません。終了します。", vbCritical, ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW))
'    Else
'        Call CMsg("OK セル [" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW).Value & "] の値が ユニットの形式（例: 2-11）です", vbInformation, ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW))
'    End If


'    MsgBox vbInformation

'    MsgBox vbExclamation

'    Call RunGitBashCommands


'    MsgBox Month(ThisWorkbook.sheetS("手順").Range(BEGIN_COL & UNITROW))
'    Exit Sub
    
End Sub









