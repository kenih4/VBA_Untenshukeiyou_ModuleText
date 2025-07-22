Module: ThisWorkbook
Option Explicit

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Application.ScreenUpdating = True
End Sub


Private Sub Workbook_Open()
    'MsgBox "ワークブックが開かれました！"
    ThisWorkbook.sheetS("手順").Activate
    ThisWorkbook.sheetS("手順").Cells(1, 1).Select
    
    Call GetWorkbookFolderToCell
    
    Call CheckCircularReference '循環参照の確認
    
    Dim filePath As String
    filePath = BNAME_KEIKAKU
    If Not CheckServerAccess_FSO(BNAME_KEIKAKU) Then 'ネットワークの接続状況を確認
        MsgBox "'" & filePath & "' にアクセスできません。ネットワーク接続に問題があるか、ファイルが存在しないか、アクセス権がありません。", vbCritical
    End If
    
    'MsgBox "TEST@Workbook_Open"
    Call StartMonitoring 'ネットワークの接続状況をモニタリング
    
End Sub
