Module: ThisWorkbook
Option Explicit

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Application.ScreenUpdating = True
End Sub


Private Sub Workbook_Open()
    'MsgBox "���[�N�u�b�N���J����܂����I"
    ThisWorkbook.sheetS("�菇").Activate
    ThisWorkbook.sheetS("�菇").Cells(1, 1).Select
    
    Call GetWorkbookFolderToCell
    
    Call CheckCircularReference '�z�Q�Ƃ̊m�F
    
    Dim filePath As String
    filePath = BNAME_KEIKAKU
    If Not CheckServerAccess_FSO(BNAME_KEIKAKU) Then '�l�b�g���[�N�̐ڑ��󋵂��m�F
        MsgBox "'" & filePath & "' �ɃA�N�Z�X�ł��܂���B�l�b�g���[�N�ڑ��ɖ�肪���邩�A�t�@�C�������݂��Ȃ����A�A�N�Z�X��������܂���B", vbCritical
    End If
    
    'MsgBox "TEST@Workbook_Open"
    Call StartMonitoring '�l�b�g���[�N�̐ڑ��󋵂����j�^�����O
    
End Sub
