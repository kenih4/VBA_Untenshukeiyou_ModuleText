Module: Module11
Option Explicit

Sub �^�]�W�v_�`������m(BL As Integer)

    '/�ǉ�����----------------------------
    On Error GoTo ErrorHandler
    Dim BNAME_SHUKEI As String
    Dim DOWNTIME_ROW As Integer
    MsgBox "�}�N���u�^�]�W�v_�`������m�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & "�ЂȌ`�V�[�g�u�^�]�󋵁i�Ώۃ��j�b�g�j�v����V�[�g�u24-*(BL" & BL & ")�v���쐬���܂��B", vbInformation, "BL" & BL
    
    Select Case BL
        Case 1
            Debug.Print "SCSS+"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SCSS\SCSS�^�]�󋵏W�vBL1.xlsm"
        Case 2
            Debug.Print "BL2"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2.xlsm"
            DOWNTIME_ROW = 8
        Case 3
            Debug.Print ">>>BL3"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
            DOWNTIME_ROW = 9
        Case Else
            Debug.Print "Zzz..."
            End
    End Select

    ' wb_SHUKEI���J��
    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False) ' �t���p�X���w��
    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
        
    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    wb_SHUKEI.Worksheets("�^�]��(�Ώۃ��j�b�g)").Activate
    
    If MsgBox("�I������Ă郆�j�b�g(�V�[�g�u���p���ԁi���ԁj�v�̃Z��B2)��    " & vbCrLf & wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B2") & "   �ł��B " & vbCrLf & "�ԈႢ�Ȃ��ł����H" & vbCrLf & "�i�ނɂ�YES�������ĉ�����", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
        Call Fin("�uNo�v���I������܂���", 1)
    End If
    '�ǉ�����----------------------------/


    Dim �ŏI�s As Integer
    Dim �V�[�g�� As String

    �ŏI�s = Cells(Rows.Count, 16).End(xlUp).ROW
    �V�[�g�� = (Cells(8, 2).Value & "(BL" & BL & ")")

    Call �����������J�n
    
    '�V�[�g�̏d������'
    Application.DisplayAlerts = False  '--- �m�F���b�Z�[�W���\��
    If SheetDetect(�V�[�g��) Then
            Worksheets(�V�[�g��).Delete
    End If
    Application.DisplayAlerts = True   '--- �m�F���b�Z�[�W��\��

    sheetS("�^�]��(�Ώۃ��j�b�g)").Copy after:=ActiveSheet '�V�[�g�̃R�s�['
    ActiveSheet.Name = �V�[�g�� '�V�[�g���ύX'
    Range("A1:P" & �ŏI�s).Value = Range("A1:P" & �ŏI�s).Value '�����˒l�֕ϊ�'

    If Cells(Range("P1:P500").Find("����_�J�n�s").ROW + 1, 7) = "" Then '���[�U�[�����Ȃ��Ƃ�'
       Rows(Range("P1:P500").Find("�V�t�g��_�J�n�s").ROW + 1 & ":" & Range("P1:P500").Find("�V�t�g��_�I���s").ROW).Delete
       Rows(Range("P1:P500").Find("����_�J�n�s").ROW + 1 & ":" & Range("P1:P500").Find("�V�t�g���[�U�[_�I���s").ROW).Delete
    Else
       Call �󔒍폜(Range("P1:P500").Find("�V�t�g��_�J�n�s").ROW + 1, Range("P1:P500").Find("�V�t�g��_�I���s").ROW - 1, 3)  '�V�t�g��_�󔒍폜'
       Call �󔒍폜(Range("P1:P500").Find("����_�J�n�s").ROW + 1, Range("P1:P500").Find("����_�I���s").ROW - 1, 3) '����_�󔒍폜'
       Call �V�t�g���[�U�[�s�}��
       Call �V�t�g���[�U�[�s_�폜
       Call �����s_�r��
    End If

    Call ����ݒ�

    Columns("O:P").Delete

    Call �����������I��

    '/�ǉ�����----------------------------
    If wb_SHUKEI.Worksheets(�V�[�g��).Cells(DOWNTIME_ROW, 9).Value = 0 Then
        MsgBox "���p�����^�](BL����orBL-study)�͂Ȃ�������ł��ˁB�@" & vbCrLf & "", vbExclamation, "BL" & BL
    End If
    
    If wb_SHUKEI.Worksheets(�V�[�g��).Cells(DOWNTIME_ROW, 11).Value = 0 Then
        MsgBox "���p�^�](���[�U�[)�͂Ȃ�������ł��ˁB�@" & vbCrLf & "" & vbCrLf & "�u���[�U�[�^�]�����v�Ǝ蓮�ŏ������Ȃ��Ƃ����Ȃ�����������܂��B", vbExclamation, "BL" & BL
    Else
        If wb_SHUKEI.Worksheets(�V�[�g��).Cells(DOWNTIME_ROW, 12).Value = 0 Then
            MsgBox "�_�E���^�C���́@" & wb_SHUKEI.Worksheets(�V�[�g��).Cells(DOWNTIME_ROW, 12).Value & " �ł��B�����g���b�v���ĂȂ����Ď��H�m�F���������悢�ł��B" & vbCrLf & "�V�[�g�u�W�v�L�^�v�ɐ����������Ă��Ȃ��\��������܂�", vbExclamation, "BL" & BL
        End If
    End If
    
    If MsgBox("���\�����Ă���V�[�g�u" & �V�[�g�� & "�v���쐬���ꂽ���̂ł��B" & vbCrLf & "������uSACLA�^�]�󋵏W�v�܂Ƃ�.xlsm�v�ɃR�s�[���܂����H", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        ' wb_MATOME���J��
        Dim wb_MATOME As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
        Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' �t���p�X���w��
        If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
        
        wb_SHUKEI.Worksheets(�V�[�g��).Copy after:=wb_MATOME.Worksheets("�܂Ƃ� ")
        wb_MATOME.Worksheets(�V�[�g��).Activate
        MsgBox wb_MATOME.Name & "��" & vbCrLf & "�V�[�g�̃R�s�[���������܂��B" & vbCrLf & "BL2/BL3�����I�������}�[�W���܂��傤�I", Buttons:=vbInformation
    End If
    
'    Call Fin("�}�N���I��" & vbCrLf & "�I", 1) ' �e�uSub �}�N�����낢��xlsm����SACLA�^�]�󋵏W�vBLxlsm�Ƀ}�N���𗬂�����Ŏ��s()�v�ɖ߂肽���̂ŃR�����g�A�E�g����
    Exit Sub  ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    Call Fin("�G���[�ł��B���e�́@ " & Err.Description, 3)
    Exit Sub
    '�ǉ�����----------------------------/
    
    
End Sub

