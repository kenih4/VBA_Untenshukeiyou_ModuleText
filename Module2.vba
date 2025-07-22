Module: Module2
Option Explicit

Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(BL As Integer)
    On Error GoTo ErrorHandler

    Dim arr() As String
    Dim BNAME_SHUKEI As String
    Dim SNAME_KEIKAKU_BL As String
    Dim RANGE_GUN_HV_OFF As String
    Dim COL_GUN_HV_OFF As Integer
    Dim tr As Variant
    Dim result As Boolean
    Dim PasteSheet As Worksheet
    Dim PasteRow As Integer
    Debug.Print "============================================================================================================"


'    Dim buttonName As String
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        MsgBox "���̃}�N���̓V�[�g��̃{�^������̂ݎ��s���Ă��������B", Buttons:=vbCritical
'        End
'    End If
'
'    If buttonName = "�{�^�� 6" Then
'        BL = 2
'    ElseIf buttonName = "�{�^�� 7" Then
'        BL = 3
'    Else
'        MsgBox "�ُ�ł��B�I�����܂��B" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    MsgBox "�u�v�掞��.xlsx�v���uSACLA�^�]�󋵏W�vBL" & BL & ".xlsm�v�ɃR�s�[����}�N���ł��B", vbInformation, "BL" & BL

    '    Dim s
    '    s = Application.InputBox("�u�v�掞��.xlsx�v���uSACLA�^�]�󋵏W�vBL" & BL & ".xlsm�v�ɃR�s�[����}�N���ł��B " & vbCrLf & vbCrLf & "BL����͂��ĉ������B", "BL" & BL)
    '    If s = False Then
    '        Exit Sub
    '    ElseIf s = "" Then
    '        MsgBox "�������͂���Ă��܂���"
    '        Exit Sub
    '    Else
    '        BL = s
    '    End If

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2.xlsm"
        SNAME_KEIKAKU_BL = "bl2"
        RANGE_GUN_HV_OFF = "A3:C"
        COL_GUN_HV_OFF = 1
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
        SNAME_KEIKAKU_BL = "bl3"
        RANGE_GUN_HV_OFF = "G3:I"
        COL_GUN_HV_OFF = 7
    Case Else
        MsgBox "BL���s���ł��B�I�����܂��B" & vbCrLf & "�I", Buttons:=vbInformation
        Exit Sub
    End Select



    ' wb_SHUKEI���J��
    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False)    ' �t���p�X���w��
    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)

    ' wb_KEIKAKU���J��
    Dim wb_KEIKAKU As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, False)    ' �t���p�X���w��
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("GUN HV OFF").Select    '�őO�ʂɕ\��


    '�R�s�[���ē\��t��
    Set PasteSheet = wb_SHUKEI.Worksheets("GUN HV OFF���ԋL�^")
    PasteRow = PasteSheet.Range("C5").End(xlDown).ROW + 1
    result = CpPaste(wb_KEIKAKU.Worksheets("GUN HV OFF"), RANGE_GUN_HV_OFF, COL_GUN_HV_OFF, PasteSheet, PasteSheet.Cells(PasteRow, 3), Array(2, 6, 7), 3)    '�u�V�[�g GUN HV OFF�v���R�s�[���ē\��t��

    Set PasteSheet = wb_SHUKEI.Worksheets("�^�]�\�莞��")
    PasteRow = PasteSheet.Range("B3").End(xlDown).ROW + 1
    result = CpPaste(wb_KEIKAKU.Worksheets(SNAME_KEIKAKU_BL), "A2:C", 1, PasteSheet, PasteSheet.Cells(PasteRow, 2), Array(1, 3, 5, 6, 8, 9, 10, 11, 12, 13), 2)    '�u�V�[�g bl*�v���R�s�[���ē\��t��
    result = CpPaste(wb_KEIKAKU.Worksheets(SNAME_KEIKAKU_BL), "D2:D", 1, PasteSheet, PasteSheet.Cells(PasteRow, 7), -1, -1)    '�u�V�[�g bl*�̔��l��v���R�s�[���ē\��t���@' �O�̍s�ŁACheck Array(1, 3, 5, 6, 8, 9, 10, 11, 12, 13), 2  ���Ă邩��{������Ȃ��̂�-1




    '�u�V�������j�b�g�����v�Z�v
    Dim before_unit As String
    Dim latest_unit As Integer
    Dim newunit As String
    PasteSheet.Cells(PasteRow - 1, 1).Select
    before_unit = PasteSheet.Cells(PasteRow - 1, 1)
    Debug.Print "before_unit: " & before_unit
    arr = Split(before_unit, "-")
    If Not IsNumeric(arr(1)) Then
        MsgBox "�V�������j�b�g�������U���Ƃ��܂��������j�b�g�����w���ł��B " & before_unit & vbCrLf & "�I�����܂��B", Buttons:=vbInformation
        Exit Sub
    End If
    latest_unit = Val(arr(1))
    latest_unit = latest_unit + 1
    newunit = arr(0) + "-" + CStr(latest_unit)
    Debug.Print "newunit: " & newunit
    If newunit <> ThisWorkbook.sheetS("�菇").Range("D" & UNITROW) Then
        MsgBox "���j�b�g�����A���ɂȂ�܂��񂯂ǁB������o�͂��悤�Ƃ��Ă��郆�j�b�g���F" & ThisWorkbook.sheetS("�菇").Range("D" & UNITROW) & vbCrLf & "  newunit: " & newunit, Buttons:=vbExclamation
    Else
        MsgBox "OK!" & vbCrLf & "�V�������j�b�g�����v!!!", Buttons:=vbInformation
    End If
    PasteSheet.Activate
    PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).ROW, 1).Activate    ' �Z��B3[�^�]���]�̍ŏI�s��
    If MsgBox("�����ɐV�������j�b�g " & newunit & "�����Ă����ł����H�H", vbYesNo + vbQuestion, "newunit") = vbYes Then
        PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).ROW, 1) = newunit
    End If

    MsgBox "�I�����܂����B" & vbCrLf & "�ۑ����Ă���A" & vbCrLf & "���Afault.txt�o��(getBlFaultSummary.py)�ɐi�݂܂��傤�I", vbInformation, "BL" & BL

    If MsgBox("���̏����ׂ̈ɁA �V�[�g�u���p���ԁi���ԁj�v�̏�̏��ɁA���j�b�g[" & newunit & "]�����Ă����ł����H�H", vbYesNo + vbQuestion, "newunit") = vbYes Then
        wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B2") = newunit
    End If

    Call Fin("����ŏI���ł��B", 1)
    Exit Sub    ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    Call Fin("�G���[�ł��B���e�́@ " & Err.Description, 3)
    Exit Sub

End Sub











Function CpPaste(sheetS As Worksheet, rangeS As String, colS As Integer, sheetT As Worksheet, pasteCELL As Variant, Arr_forCheck As Variant, Col_forCheck As Integer) As Boolean
'   rng1 As Range,
'    MsgBox sheetS.Columns(3).Address ' $C:$C
'    MsgBox sheetS.Range("$A:$A").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' 5
'    MsgBox sheetS.Range("$C:$C").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' 2
'    MsgBox sheetS.Range(Range(HeaderCELL).Columns.Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' �G���[
    Dim tr As Variant
    sheetS.Activate
    Set tr = Range(rangeS & Cells(Rows.Count, colS).End(xlUp).ROW)
    tr.Copy
    tr.Select
    If MsgBox("�I�𕔕����R�s�[���܂����B" & "�s���� " & tr.Rows.Count & vbCrLf & "���ɐi�ނɂ�Yes", vbYesNo + vbQuestion) = vbNo Then Exit Function
    
    If Col_forCheck > 0 Then
        If Check(Arr_forCheck, Col_forCheck, tr.Rows.Count + 10, sheetT) <> 0 Then Call Fin("�\�t����̃V�[�g�ɐ����������Ă��Ȃ��ӏ���������܂����B�I�����܂��B" & vbCrLf & "�����𒼂��Ă���ēx�s���ĉ������B", 3)
    End If
    sheetT.Activate    ' ����K�v�B����Ȃ��ƁA���̍s�ŁA�Z�����A�N�e�B�u�ɂł��Ȃ�
    pasteCELL.Activate

    If MsgBox("�����ɓ\��t���Ă����ł����H", vbYesNo + vbQuestion) = vbYes Then
        pasteCELL.PasteSpecial Paste:=xlPasteValues
        If MsgBox("�\��t���܂�����OK�ł����H�H" & vbCrLf & "���ɐi�ނɂ�Yes", vbYesNo + vbQuestion) = vbNo Then Exit Function
    End If

End Function

