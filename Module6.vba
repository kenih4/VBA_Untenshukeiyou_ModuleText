Module: Module6
Option Explicit

Sub ���p����User�Ɏ蓮����(BL As Integer)
    On Error GoTo ErrorHandler

    Debug.Print "============================================================================================================"
    Debug.Print "============���p����User�Ɏ蓮����================================================================ BL=" & BL
    Debug.Print "============================================================================================================"

    Dim BNAME_SHUKEI As String
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim i As Integer
        
    Dim result As Double
    result = Application.WorksheetFunction.RoundUp(Now - ThisWorkbook.sheetS("�菇").Cells(UNITROW, 5).Value, 0)
    
    MsgBox "python Pickup_from_shiftsummary.py�ŁA" & result & "�����̃V�t�g�T�}���[���擾���܂��B"
    If RunPythonScript("Pickup_from_shiftsummary.py BL" & BL & " " & result * 3, "C:\Users\kenic\Dropbox\gitdir\Pickup_from_shiftsummary") = False Then
        MsgBox "python�ŃG���[�����ł��B�V�t�g�T�}���[����擾�ł��܂���ł����B�蓮�ōs���Ă��������B", Buttons:=vbExclamation
    End If
     
    
    ' �E�B���h�E��W���T�C�Y�ɂ���
    Application.WindowState = xlNormal
    ' �őO�ʂɎ����Ă���
    Application.ActiveWindow.Activate
    
    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
    Case Else
        Debug.Print "Zzz..."
        Exit Sub
    End Select

'    Dim WSH
'    Set WSH = CreateObject("Wscript.Shell")
'    WSH.RUN "http://saclaopr19.spring8.or.jp/~summary/display_ui.html?sort=date%20desc%2Cstart%20desc&limit=0%2C100&search_root=BL" & BL & "#STATUS", 3
'    Set WSH = Nothing

   ' �}�N�����낢����J���@���ɊJ����Ă邪
    Dim sourceWorkbook As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set sourceWorkbook = OpenBook(BNAME_SOURCE, False) ' �t���p�X���w��
    If sourceWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    sourceWorkbook.Worksheets("�ҏW�p_���p����(User)BL" & BL).Activate

    ' wb_SHUKEI���J��
    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False)    ' �t���p�X���w��
    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    wb_SHUKEI.Worksheets("���p����(User)").Activate
    
    LineSta = getLineNum("���j�b�g", 2, wb_SHUKEI.Worksheets("���p����(User)"))
    LineSto = wb_SHUKEI.Worksheets("���p����(User)").Cells(Rows.Count, "B").End(xlUp).ROW
    Debug.Print " LineSto :   " & LineSto
    Dim Kokokara As Long
    Dim Kokomade As Long
    For i = LineSta To LineSto
        Debug.Print "DEBUG �@    i = " & i & "  " & Cells(i, 2).Value
        If wb_SHUKEI.Worksheets("���p����(User)").Cells(i, 2).Value = wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B2") Then
            Debug.Print "���̍s�@i = " & i & " ���A�@�@���j�b�g�F " & Cells(i, 2).Value
            'Cells(i, 15).Select
            Kokokara = i
            Exit For
        End If
    Next
    Debug.Print "���p����(User)�̍ŏI�s = " & wb_SHUKEI.Worksheets("���p����(User)").Range(wb_SHUKEI.Worksheets("���p����(User)").Columns(15).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).ROW '�r���܂܂Ȃ��ŏI�s
    Kokomade = wb_SHUKEI.Worksheets("���p����(User)").Range(wb_SHUKEI.Worksheets("���p����(User)").Columns(15).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).ROW
    Range("O" & Kokokara & ":" & "O" & Kokomade).Select

    Windows.Arrange ArrangeStyle:=xlVertical



    Call Fin("�}�N���͂���ŏI���ł��B" & vbCrLf & "���Ƃ̓V�t�g�T�}���[����G�l���M�[�A�J��Ԃ��A�A�g���A���x���s�b�N�A�b�v���ĉ�����", 1)
    Exit Sub  ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    Call Fin("�G���[�ł��B���e�́@ " & Err.Description, 3)
    Exit Sub

End Sub
