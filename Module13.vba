Module: Module13
Option Explicit

'==============================================================================================================================
Sub Button14()
'    Call RunBatchFile("C:\Users\kenichi\Documents\operation_log_NEW\vscode_operation_log.bat")
    Call RunBatchFile("C:\Users\kenic\Dropbox\gitdir\vscode_open\vscode_open.bat C:\Users\kenic\Documents\operation_log_NEW")
End Sub



Sub �{�^��15_Click()
    Dim wb_MATOME As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' �t���p�X���w��
    If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
        
    Dim Sonzai_flg As Boolean: Sonzai_flg = False
    Sonzai_flg = SheetExists(wb_MATOME, "�܂Ƃ� ")
    If Not Sonzai_flg Then
        MsgBox "�V�[�g�����݂��܂���B" & vbCrLf & " �I�����܂��B", Buttons:=vbExclamation
    Else
        Call �K�؂ȉӏ��ɉ��y�[�W������Ver2(wb_MATOME.Worksheets("�܂Ƃ� "))
    End If
End Sub



Sub �{�^��16_Click()
    Dim wb_MATOME As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' �t���p�X���w��
    If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    
    Dim Sonzai_flg As Boolean: Sonzai_flg = False
    Sonzai_flg = SheetExists(wb_MATOME, "Fault�W�v")
    If Not Sonzai_flg Then
        MsgBox "�V�[�g�����݂��܂���B" & vbCrLf & " �I�����܂��B", Buttons:=vbExclamation
    Else
        Call �K�؂ȉӏ��ɉ��y�[�W������Ver2(wb_MATOME.Worksheets("Fault�W�v"))
    End If
End Sub





Sub �{�^��18_Click() ' TEST!!!!!!!!!!!!
'    Dim wb_MATOME As Workbook    '�@�����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
'    Set wb_MATOME = OpenBook("\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�v�܂Ƃ�.xlsm") ' �t���p�X���w��
'    If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B",3)
    
    
'     Dim BL As Integer: BL = 2
'     Select Case BL
'        Case 1
'            Debug.Print "SCSS+"
'        Case 2
'            Debug.Print "BL2"
'            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2.xlsm"
'            SNAME_KEIKAKU_BL = "bl2"
'        Case 3
'            Debug.Print ">>>BL3"
'            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
'            SNAME_KEIKAKU_BL = "bl3"
'        Case Else
'            Debug.Print "Zzz..."
'            Exit Sub
'    End Select
'
'    ' sourceWorkbook���J��
'    Dim sourceWorkbook As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
'    Set sourceWorkbook = OpenBook(BNAME_SOURCE) ' �t���p�X���w��
'    If sourceWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B",3)
'
'
'    ' wb_SHUKEI���J��
'    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
'    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI) ' �t���p�X���w��
'    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B",3)
'
'    Dim result As Boolean
'    Dim macroName As String: macroName = "Cleanup" ' �Ƃ肠�����ASub Cleanup���܂܂�郂�W���[�����폜����B�v�������C
'    result = sourceWorkbook����targetWorkbook��moduleName�𗬂�����(sourceWorkbook, wb_SHUKEI, "Module8", "Cleanup", False)
'    If result Then
'        MsgBox "���� �u�}�N�����낢��xlsm����SACLA�^�]�󋵏W�vBLxlsm�Ƀ}�N���𗬂����ނ����v", Buttons:=vbInformation
'    Else
'        MsgBox "���s �u�}�N�����낢��xlsm����SACLA�^�]�󋵏W�vBLxlsm�Ƀ}�N���𗬂����ނ����v", Buttons:=vbInformation
'    End If
    
    
'    Dim folderPath As String
'    folderPath = GetWorkbookFolder()
'    If folderPath = "" Then
'        MsgBox "���̃u�b�N�͂܂��ۑ�����Ă��܂���B"
'    Else
'        MsgBox "���̃u�b�N���J����Ă���t�H���_�̃p�X: " & folderPath
'    End If
    
    Call ToggleButton
    
End Sub



' ========================================================================================================================
Sub PDF_output_Click()
    Debug.Print "PDF_output_Click"
   
    Dim myArray(2) As Variant
    Dim i As Integer
    Dim pdfPath As String
    Dim OutDir As String
    OutDir = "PDF�쐬"

    Dim wb_MATOME As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' �t���p�X���w��
    If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
        
    Dim sheet As Worksheet
    myArray(0) = "�܂Ƃ� " '�V�[�g���@�܂Ƃ� �V�[�g�ɂ͔��p�X�y�[�X������̂Œ���
    myArray(1) = "Fault�W�v"
    myArray(2) = ThisWorkbook.sheetS("�菇").Range("D" & UNITROW)
    
    For i = LBound(myArray) To UBound(myArray)
'        MsgBox "�v�f " & i & ": " & myArray(i)
        Set sheet = wb_MATOME.Worksheets(myArray(i))
'       sheet.PrintPreview
        pdfPath = CPATH & WHICH & "\" & OutDir & "\" & WHICH & "�^�]�󋵏W�v(" & myArray(i) & ").pdf"
        Debug.Print "pdfPath:   " & pdfPath
        ' �V�[�g��PDF�Ƃ��ăG�N�X�|�[�g
        sheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, _
                              IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                              OpenAfterPublish:=False
                              '  IgnorePrintAreas: False�̏ꍇ�A�ݒ肳�ꂽ����G���A�݂̂�PDF�ɃG�N�X�|�[�g����܂��B
    
        ' PDF���J��
        pdfPath = CPATH & WHICH & "\" & OutDir & "\" & WHICH & "�^�]�󋵏W�v(" & myArray(i) & ").pdf"
        shell """" & edgePath & """ --new-window """ & pdfPath & """", vbNormalFocus
'         shell """" & edgePath & """ --start-maximized """ & pdfPath & """", vbNormalFocus      [--start-maximized]�I�v�V�������Ă��ő剻���ꂸ
        MsgBox "�^�]�󋵏W�v(" & myArray(i) & ").pdf" & vbCrLf & "���o�͂��܂����B", vbInformation
        
    Next i
    
End Sub



