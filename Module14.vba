Module: Module14
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Initial_Check(BL As Integer)

    On Error GoTo ErrorHandler

    '    Dim BL As Integer
    Dim BNAME_SHUKEI As String
    Dim sname As String
    Dim Cnt As Integer
    Dim result As Boolean

    '    Dim s
    '    s = Application.InputBox("BL����͂��ĉ������B", "�m�F", Type:=1)    '  Type:=1 ���l�̂�
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
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SCSS\SCSS�^�]�󋵏W�vBL1.xlsm"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2TEST.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    '    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2TEST.xlsm"
    MsgBox "�}�N���uInitial_Check()�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & BNAME_SHUKEI & vbCrLf & "�̃`�F�b�N�ł��B" & vbCrLf & "�����������Ă���ׂ��Z���ɐ����������Ă��邩�m�F���܂�", vbInformation, "BL" & BL

    ' wb_SHUKEI���J��
    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' �t���p�X���w��
    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    wb_SHUKEI.Windows(1).WindowState = xlMaximized

    Debug.Print "�V�[�g�S�̂ɃG���[���Ȃ����m�F "
    Dim ws As Worksheet
    For Each ws In wb_SHUKEI.Worksheets
        result = CheckForErrors(ws)
    Next ws
    
    
    If Check_exixt("�^�]�\�莞��", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets("�^�]�\�莞��"))
    If Check_exixt("GUN HV OFF���ԋL�^", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7), 3, 30, wb_SHUKEI.Worksheets("GUN HV OFF���ԋL�^"))
    If Check_exixt("GUN HV OFF���ԋL�^", wb_SHUKEI) = True Then Cnt = Check(Array(9, 10, 11, 12, 13, 14, 15), 9, 30, wb_SHUKEI.Worksheets("GUN HV OFF���ԋL�^"))
    If Check_exixt("�W�v�L�^", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 6, 7, 8, 9), 3, 500, wb_SHUKEI.Worksheets("�W�v�L�^")) ' �Ƃ肠����500�s���炢�`�F�b�N    E��(Fault)���`�F�b�N���������A�����͓���@�ŏI�s��2�s�ڂ���ςȐ����������Ă邪����̂��H
    If Check_exixt("���p���ԁi���ԁj", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14), 2, 30, wb_SHUKEI.Worksheets("���p���ԁi���ԁj")) ' ���p���ԁi���ԁj �̃J�b�R�͑S�p
    If Check_exixt("���p���ԁi���ԁj", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 25, 26, 27, 28, 29), 2, 30, wb_SHUKEI.Worksheets("���p����(User)"))
    If Check_exixt("���p����(�V�t�g)", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), 1, 30, wb_SHUKEI.Worksheets("���p����(�V�t�g)"))
    If Check_exixt("Fault�Ԋu(���j�b�g)", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 2, 30, wb_SHUKEI.Worksheets("Fault�Ԋu(���j�b�g)")) ' �V�[�g���ɔ��p�X�y�[�X�������Ă邱�Ƃ�����̂Œ��Ӂ@���������Ȃ��Ɓu�C���f�b�N�X���L���͈͂ɂ���܂���v�ƃG���[���b�Z�[�W���ł���
    
    '�V�[�g�̑��݂��m�F���鏈����ǉ�����Ƃ����Ȃ邪�A���ɂ����B�B�B�B�B
'    sname = "�^�]�\�莞��"
'    If Not SheetExists(wb_SHUKEI, sname) Then
'        MsgBox "�V�[�g�����݂��܂���B" & vbCrLf & sname & " �I�����܂��B", Buttons:=vbExclamation
'    Else
'        If CheckStringInSheet(wb_SHUKEI.Worksheets(sname), ThisWorkbook.sheetS("�菇").Range("D" & UNITROW)) Then
'            wb_SHUKEI.Worksheets(sname).Activate
'            MsgBox "������o�͂��悤�Ƃ��Ă��郆�j�b�g�����ɃV�[�g��ɑ��݂��܂����ǁA�A�A�@�m�F���ĉ������B�@ ", Buttons:=vbCritical
'        Else
'            Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets(sname))
'        End If
'    End If


'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets("�^�]�\�莞��"))
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7), 3, 30, wb_SHUKEI.Worksheets("GUN HV OFF���ԋL�^"))
'    Cnt = Check(Array(9, 10, 11, 12, 13, 14, 15), 9, 30, wb_SHUKEI.Worksheets("GUN HV OFF���ԋL�^"))
'    Cnt = Check(Array(2, 3, 4, 6, 7, 8, 9), 3, 500, wb_SHUKEI.Worksheets("�W�v�L�^")) ' �Ƃ肠����500�s���炢�`�F�b�N    E��(Fault)���`�F�b�N���������A�����͓���@�ŏI�s��2�s�ڂ���ςȐ����������Ă邪����̂��H
'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14), 2, 30, wb_SHUKEI.Worksheets("���p���ԁi���ԁj")) ' ���p���ԁi���ԁj �̃J�b�R�͑S�p
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 25, 26, 27, 28, 29), 2, 30, wb_SHUKEI.Worksheets("���p����(User)"))
'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), 1, 30, wb_SHUKEI.Worksheets("���p����(�V�t�g)"))
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 2, 30, wb_SHUKEI.Worksheets("Fault�Ԋu(���j�b�g)")) ' �V�[�g���ɔ��p�X�y�[�X�������Ă邱�Ƃ�����̂Œ��Ӂ@���������Ȃ��Ɓu�C���f�b�N�X���L���͈͂ɂ���܂���v�ƃG���[���b�Z�[�W���ł���
        
    Call Fin("�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
    
End Sub





Sub Middle_Check(BL As Integer)

    On Error GoTo ErrorHandler

    Dim BNAME_SHUKEI As String
    Dim sname As String
    Dim Cnt As Integer
    Dim result As Boolean
    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim ws As Worksheet

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SCSS\SCSS�^�]�󋵏W�vBL1.xlsm"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL3.xlsm"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    MsgBox "�}�N���uMiddle_Check()�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & BNAME_SHUKEI & vbCrLf & "�̒��ԃ`�F�b�N�ł��B" & vbCrLf & "���[�U�[�^�]�̊J�n�I�������Ȃǂ̊m�F���܂�", vbInformation, "BL" & BL

    ' wb_SHUKEI���J��
    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' �t���p�X���w��
    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    
    If ThisWorkbook.sheetS("�菇").Range("D" & UNITROW).Value <> wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B2").Value Then
        If MsgBox("�V�[�g�u���p���ԁi���ԁj�v�ɓ��͂���Ă郆�j�b�g�� �� �� �� �� ���ǁA�i�߂܂����H", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
            Exit Sub
        End If
    End If

    wb_SHUKEI.Worksheets("�^�]�\�莞��").Select    '�őO�ʂɕ\��_______________________________________________________________________________
    wb_SHUKEI.Worksheets("�^�]�\�莞��").Activate
    ' �V�[�g[�^�]�\�莞��]��E��ŏI�s��336.0���Ǝ��Ԃŕ\������Ă��邪�A���g�͓��B
    If Int(Cells(GetLastDataRow(wb_SHUKEI.Worksheets("�^�]�\�莞��"), "E"), "E")) <> Int(ThisWorkbook.sheetS("�菇").Range("I" & UNITROW)) Then
        Call CMsg("�V�[�g�u�^�]�\�莞�ԁv��E��ŏI�s�ƃ��j�b�g���v���Ԃ���v���܂���" & vbCrLf & Int(Cells(GetLastDataRow(wb_SHUKEI.Worksheets("�^�]�\�莞��"), "E"), "E")) & " �� " & Int(ThisWorkbook.sheetS("�菇").Range("I" & UNITROW)), vbCritical, Cells(GetLastDataRow(wb_SHUKEI.Worksheets("�^�]�\�莞��"), "E"), "E"))
    Else
        Call CMsg("��v�AOK!!" & vbCrLf & vbCrLf & "���j�b�g���v���Ԃ���v", vbInformation, Cells(GetLastDataRow(wb_SHUKEI.Worksheets("�^�]�\�莞��"), "E"), "E"))
    End If
    
    
    
    wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Select    '�őO�ʂɕ\��_______________________________________________________________________________
    wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Activate
    Set ws = wb_SHUKEI.Worksheets("���p���ԁi���ԁj")
    If MsgBox("���̃��j�b�g�����m�F���܂����H" & vbCrLf & "Yes:�@[" & ThisWorkbook.sheetS("�菇").Range("D" & UNITROW) & "]�����m�F" & vbCrLf & "No:�@�S���j�b�g�m�F", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("�菇").Range("D" & UNITROW), 3, ws)
    Else
        LineSta = 4
    End If
    LineSto = GetLastDataRow(ws, "A")
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("A" & LineSta & ":A" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B" & LineSta & ":B" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("C" & LineSta & ":C" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("D" & LineSta & ":D" & LineSto))
    For i = LineSta To LineSto
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        
        If Not CheckValMatch(ws.Cells(i, "F").Value + ws.Cells(i, "H").Value + ws.Cells(i, "J").Value, ws.Cells(i, "E").Value) Then  ' �u���v���ԁv�̊m�F
            Call CMsg("�u���v���ԁv����v���܂���" & ws.Cells(i, "F").Value & "   " & ws.Cells(i, "H").Value & "   " & ws.Cells(i, "J").Value & "   E=" & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "J"), ws.Cells(i, "M")) Then
            Call CMsg("[���p�^�]�v��]����v���܂���", vbCritical, Cells(i, "J"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "E"), ws.Cells(i, "N")) Then
            Call CMsg("[���^�]����]����v���܂���", vbCritical, Cells(i, "N"))
        End If
        
        If ws.Cells(i, "G").Value > ws.Cells(i, "F").Value Or ws.Cells(i, "G").Value < 0 Then ' �u�{�ݒ����v��v�̊m�F
            Call CMsg("�u�{�ݒ����v��v���u�{�ݒ����v��_�E���^�C���v�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "G"))
        End If

        If ws.Cells(i, "I").Value > ws.Cells(i, "H").Value Or ws.Cells(i, "I").Value < 0 Then ' �u���p�����v��v�̊m�F
            Call CMsg("�u���p�����v��v���u���p�����v��_�E���^�C���v�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "I"))
        End If
                
        If ws.Cells(i, "K").Value > ws.Cells(i, "J").Value Or ws.Cells(i, "K").Value < 0 Then ' �u���p�^�]�v��v�̊m�F
            Call CMsg("�u���p�^�]�v��v���u���p�^�]�v��_�E���^�C���v�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
        
    Next
    
    
    
    wb_SHUKEI.Worksheets("�z��").Select    '�őO�ʂɕ\��_______________________________________________________________________________
    wb_SHUKEI.Worksheets("�z��").Activate
    If GetLastDataRow(wb_SHUKEI.Worksheets("�W�v�L�^"), "C") <> Cells(4, "E").Value Then
        Call CMsg("�V�[�g�u�W�v�L�^�v�̍ŏI�s�ƈ�v���܂���" & vbCrLf & "", vbCritical, Cells(4, "E"))
    Else
        Call CMsg("��v�AOK!!" & vbCrLf & vbCrLf & vbCrLf & "�V�[�g�u�W�v�L�^�v�̍ŏI�s�ƈ�v", vbInformation, Cells(4, "E"))
    End If
 
 
 
 
    wb_SHUKEI.Worksheets("���p����(�V�t�g)").Select    '�őO�ʂɕ\��_______________________________________________________________________________
    wb_SHUKEI.Worksheets("���p����(�V�t�g)").Activate
    Set ws = wb_SHUKEI.Worksheets("���p����(�V�t�g)")
    If MsgBox("���̃��j�b�g�����m�F���܂����H" & vbCrLf & "Yes:�@[" & ThisWorkbook.sheetS("�菇").Range("D" & UNITROW) & "]�����m�F" & vbCrLf & "No:�@�S���j�b�g�m�F", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("�菇").Range("D" & UNITROW), 2, ws)
    Else
        LineSta = 9
    End If
    LineSto = GetLastDataRow(ws, "B")
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p����(�V�t�g)").Range("A" & LineSta & ":A" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p����(�V�t�g)").Range("C" & LineSta & ":C" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("���p����(�V�t�g)").Range("D" & LineSta & ":D" & LineSto))
    
    For i = LineSta To LineSto
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        
        If Not IsDateTimeFormatRegEx(Cells(i, "C")) Or Not IsDateTimeFormatRegEx(Cells(i, "D")) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, "C"))
        End If
        
        If (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) <> ws.Cells(i, "E").Value Then ' �u���v���ԁv�̊m�F
            Call CMsg("�u���v���ԁv����v���܂���   " & vbCrLf & "    �����F" & (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) & "   E��:" & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
        End If
        
        If ws.Cells(i, "F").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "F").Value < 0 Then ' �u���p���ԁv�̊m�F
            Call CMsg("�u���p���ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "F"))
        End If
        
        If ws.Cells(i, "G").Value > 100 Or ws.Cells(i, "G").Value < 0 Then  ' �u���p���v�̊m�F
            Call CMsg("�u���p���v��  0 ~ 100%�͈̔͂łȂ�   " & vbCrLf & "====", 3, Cells(i, "G"))
        End If
        
        If ws.Cells(i, "H").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "H").Value < 0 Then ' �u�������ԁv�̊m�F
            Call CMsg("�u�������ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "H"))
        End If
        
        If ws.Cells(i, "I").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "I").Value < 0 Then ' �uFault���ԁv�̊m�F
            Call CMsg("�uFault���ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "I"))
        End If
        
        If ws.Cells(i, "J").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "J").Value < 0 Then ' �u�_�E���^�C���v�̊m�F
            Call CMsg("�u�_�E���^�C���v���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "J"))
        End If
        
        If ws.Cells(i, "K").Value < 0 Then  ' �uFault���v�v�̊m�F
            Call CMsg("�uFault���v�v��  ��" & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
            
        If ws.Cells(i, "L").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "L").Value < 0 Then ' �uFault�Ԋu�v�̊m�F
            Call CMsg("�uFault�Ԋu�v���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "L"))
        End If
    
        If IsNumeric(ws.Cells(i, "M").Value) Or InStr(Cells(i, "M"), "G") = 0 Then  ' �u���[�U�[�v�̊m�F
            Call CMsg("�u���[�U�[�v�����l�A�܂��́A���[�U�[���Ȃ̂ɁuG�v���Ȃ��A" & vbCrLf & "====", vbExclamation, Cells(i, "M"))
        End If
        
    Next
    
    
    wb_SHUKEI.Worksheets("���p����(User)").Select    '�őO�ʂɕ\��_______________________________________________________________________________
    wb_SHUKEI.Worksheets("���p����(User)").Activate
    Set ws = wb_SHUKEI.Worksheets("���p����(User)")
'    CheckForErrors (ws)
    If MsgBox("���̃��j�b�g�����m�F���܂����H" & vbCrLf & "Yes:�@[" & ThisWorkbook.sheetS("�菇").Range("D" & UNITROW) & "]�����m�F" & vbCrLf & "No:�@�S���j�b�g�m�F", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("�菇").Range("D" & UNITROW), 2, ws)
    Else
        LineSta = 9
    End If
'   LineSto = ws.Cells(wb_SHUKEI.Worksheets("���p����(User)").Rows.Count, "B").End(xlUp).ROW ' ��B�̍ŉ��s���������Ƀf�[�^��T���̂ŁA�󔒂������Ă������ł��܂��B ���ꂾ�Ɛ����������Ă�Ɩ���
    LineSto = GetLastDataRow(ws, "B")
    
    For i = LineSta To LineSto
'       Debug.Print "���̍s�@i = " & i & " ���A" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        If Not IsDateTimeFormatRegEx(Cells(i, "C")) Or Not IsDateTimeFormatRegEx(Cells(i, "D")) Or Not IsDateTimeFormatRegEx(Cells(i, "E")) Or Not IsDateTimeFormatRegEx(Cells(i, "F")) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, "C"))
        Else
            If Not CheckCellsMatch(ws.Cells(i, "C"), ws.Cells(i, "E")) Then
                Call CMsg("��������v���܂���   " & vbCrLf & "" & ws.Cells(i, "C").Value & vbCrLf & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
            End If
            If Not CheckCellsMatch(ws.Cells(i, "D"), ws.Cells(i, "F")) Then
                Call CMsg("��������v���܂���   " & vbCrLf & "" & ws.Cells(i, "D").Value & vbCrLf & ws.Cells(i, "F").Value, vbCritical, Cells(i, "F"))
            End If
        End If
        
        If Not CheckValMatch(ws.Cells(i, "D").Value - ws.Cells(i, "C").Value, ws.Cells(i, "G").Value) Then    ' �u���v���ԁv�̊m�F
            Call CMsg("�u���v���ԁv����v���܂���   " & vbCrLf & "    �����F" & (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) & "   G��:" & ws.Cells(i, "G").Value, vbCritical, Cells(i, "G"))
        End If
        
        If ws.Cells(i, "H").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "H").Value < 0 Then ' �u���p���ԁv�̊m�F
            Call CMsg("�u���p���ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "H"))
        End If
        
        If ws.Cells(i, "I").Value > 100 Or ws.Cells(i, "I").Value < 0 Then  ' �u���p���v�̊m�F
            Call CMsg("�u���p���v��  0 ~ 100%�͈̔͂łȂ�   " & vbCrLf & "====", 3, Cells(i, "I"))
        End If
        
        If ws.Cells(i, "J").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "J").Value < 0 Then ' �u�������ԁv�̊m�F
            Call CMsg("�u�������ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "J"))
        End If
        
        If ws.Cells(i, "K").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "K").Value < 0 Then ' �uFault���ԁv�̊m�F
            Call CMsg("�uFault���ԁv���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
        
        If ws.Cells(i, "L").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "L").Value < 0 Then ' �u�_�E���^�C���v�̊m�F
            Call CMsg("�u�_�E���^�C���v���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "L"))
        End If
        
        If ws.Cells(i, "M").Value < 0 Then  ' �uFault���v�v�̊m�F
            Call CMsg("�uFault���v�v��  ��" & vbCrLf & "====", vbCritical, Cells(i, "M"))
        End If

        If ws.Cells(i, "N").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "N").Value < 0 Then ' �uFault�Ԋu�v�̊m�F
            Call CMsg("�uFault�Ԋu�v���u���v���ԁv�����傫���A�܂��́A��  " & vbCrLf & "====", vbCritical, Cells(i, "N"))
        End If
    
        If IsNumeric(ws.Cells(i, "O").Value) Or InStr(Cells(i, "O"), "G") = 0 Then  ' �u���[�U�[�v�̊m�F
            Call CMsg("�u���[�U�[�v�����l�A�܂��́A���[�U�[���Ȃ̂ɁuG�v���Ȃ��A" & vbCrLf & "====", vbExclamation, Cells(i, "O"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "G"), ws.Cells(i, "W")) Then
            Call CMsg("�u���[�U�[�^�]���ԁi�v��j�v����v���܂���   ", vbCritical, Cells(i, "W"))
        End If
    Next

    Call Fin("�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
    
End Sub











Function Check_exixt(sname As String, wb As Workbook) As Boolean

    Check_exixt = False
    If Not SheetExists(wb, sname) Then
        MsgBox "�V�[�g�����݂��܂���B" & vbCrLf & sname & " �I�����܂��B", Buttons:=vbCritical
    Else
        If CheckStringInSheet(wb.Worksheets(sname), ThisWorkbook.sheetS("�菇").Range("D" & UNITROW)) Then
            wb.Worksheets(sname).Activate
            MsgBox "������o�͂��悤�Ƃ��Ă��郆�j�b�g�����ɃV�[�g��ɑ��݂��܂����ǁA�A�A�@�m�F���ĉ������B�@ ", Buttons:=vbCritical
        Else
            Check_exixt = True
        End If
    End If
    
End Function




'��ŁA�v�m�F�I
'VBA�ł́A�����I�� ByVal �� ByRef ���w�肵�Ȃ��ꍇ�A�f�t�H���g�� ByRef�i�Q�Ɠn���j�ɂȂ�܂��B
'�܂褈����Ƃ��ēn�����ϐ��̒l���ύX�����\�������� �̂Œ��ӂ��K�v�ł��
'Function Check(arr As Variant, ByVal Retsu_for_Find_last_row As Integer, ByVal Check_row_cnt As Integer, ByVal sheet As Worksheet) As Integer
' StartL , EndL�������ɂ������������C������
Function Check(arr As Variant, Retsu_for_Find_last_row As Integer, Check_row_cnt As Integer, sheet As Worksheet) As Integer
' arr:  �`�F�b�N������z��ɃZ�b�g
' Retsu_for_Find_last_row:  �l�̓����Ă���ŏI�s���擾���邽�߂̂��́B�����������Ă��Ȃ�����w�肷��B�����������Ă������w�肷��Ɛ����������Ă��Ȃ��ŏI�s�ɂȂ��Ă��܂�
' Check_row_cnt:    ���s�`�F�b�N���邩�B�Ƃ肠�������߂ɂ��Ƃ�
    Debug.Print "DEBUG  Start Function Check()-------------"
    Dim result As Boolean
    Dim StartL As Integer
    Dim i As Integer
    Dim col As Variant
    Check = 0
    
    sheet.Activate


    '    MsgBox "Columns(Retsu_for_Find_last_row).Address�@=     " & Columns(Retsu_for_Find_last_row).Address

    '    StartL = sheet.Range("B:B").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1  ' �r���͖���
    '    StartL = sheet.Range("A:A").Find(What:="*", LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1 ' ���̕��@���ƌr�����܂񂾍ŏI�s�ɂȂ��Ă��܂�
    '    StartL = sheet.Cells(Rows.Count, Retsu_for_Find_last_row).End(xlUp).Row + 1
    '    StartL = sheet.Range(Columns(Retsu_for_Find_last_row).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1 ' �Ȃ����@�V�[�g�u���p����(User)�v�����A�u�I�u�W�F�N�g�ϐ��܂���With�u���b�N�ϐ����ݒ肳��Ă��܂���v�̃G���[  ���͂����@Columns(Retsu_for_Find_last_row).Address
    StartL = sheet.Range(sheet.Columns(Retsu_for_Find_last_row).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).ROW + 1    ' TEST

    sheet.Cells(StartL, arr(0)).Select
    MsgBox "�V�[�g�u" & sheet.Name & "�v�̂�������A���̍s�ɓ����Ă��鐔�����ȍ~ " & Check_row_cnt & " �s�ɓn���ē����Ă��邩�`�F�b�N���n�߂܂��B", vbInformation

    For Each col In arr
        For i = StartL + 1 To StartL + Check_row_cnt
            sheet.Cells(i, col).Select
            'Sleep 20 ' msec
            result = CheckSameFormulaType(Cells(StartL, col), Cells(i, col))
            If result = True Then
                Debug.Print "OK:    �Z��(" & i & ", " & col & ") �����L  " & Cells(i, col).Formula
                'Cells(i, col).Interior.Color = RGB(0, 255, 0)  �F�t����Ɣ��Ɏ��Ԃ��|����
            Else
                Debug.Print "�v�m�F�I�@�Z��(" & i & ", " & col & ") �����������Ă��Ȃ����A�������قȂ�"
                Cells(i, col).Interior.Color = RGB(255, 0, 0)
                Check = Check + 1
            End If
        Next
    Next col
    If Check <> 0 Then
        MsgBox "�V�[�g�u" & sheet.Name & "�v�ɂāA" & vbCrLf & "�����������Ă��Ȃ����A�������قȂ�Z���� " & Check & " �ӏ��A������܂����I�I�v�m�F�ł�", vbCritical
    End If

End Function




'------------------------------------------------------------






Function CheckSameFormulaType(rng1 As Range, rng2 As Range) As Boolean
    CheckSameFormulaType = (rng1.FormulaR1C1 = rng2.FormulaR1C1)
End Function

'Function CheckSameFormulaType(rng1 As Range, rng2 As Range) As Boolean
'' �Z���ɐ����������Ă��邩�m�F
'    If rng1.HasFormula And rng2.HasFormula Then
'        'Debug.Print "�ǂ��炩�̃Z���ɐ���������"
'        ' R1C1�`���Ŕ�r���āA��v����� True�A�قȂ�� False
'        CheckSameFormulaType = (rng1.FormulaR1C1 = rng2.FormulaR1C1)
'    Else
'        'Debug.Print "�ǂ��炩�̃Z���ɐ���������"
'        CheckSameFormulaType = False
'    End If
'End Function














Sub �v�掞��xlsx_Check(BL As Integer)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim result As Boolean
    Dim pattern As String
'    pattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$"    '       �ʂ̏����i��: YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM�j pattern = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2} - \d{4}-\d{2}-\d{2} \d{2}:\d{2}$"
'    pattern = "^\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}$"    '   ���Ԃ��ꌅ�̏ꍇ������̂ł���ɑΉ�
    pattern = "^\d{4}/\d{1,2}/\d{1,2}[ ]{1,2}\d{1,2}:\d{1,2}:\d{1,2}$"  ' �X�y�[�X�̐���1�A�܂���2�ł��}�b�`����悤�ɂ������ł��B

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
    Case 3
        Debug.Print ">>>BL3"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    '    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL2TEST.xlsm"
    MsgBox "�}�N���u�v�掞��xlsx_Check�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & BNAME_KEIKAKU & vbCrLf & "�̃`�F�b�N�ł��B" & vbCrLf & "�m�F���܂�", vbInformation, "BL" & BL


    ' wb_KEIKAKU���J��
    Dim wb_KEIKAKU As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, True)    ' �t���p�X���w��
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    Debug.Print "�V�[�g�S�̂ɃG���[���Ȃ����m�F "
    Dim ws As Worksheet
    For Each ws In wb_KEIKAKU.Worksheets
        result = CheckForErrors(ws)
    Next ws

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("bl" & BL).Select    '�őO�ʂɕ\��

    wb_KEIKAKU.Worksheets("bl" & BL).Activate    '����厖
    LineSta = 2 ' getLineNum("�^�]���", 1, wb_KEIKAKU.Worksheets("bl" & BL)) + 1
    LineSto = wb_KEIKAKU.Worksheets("bl" & BL).Cells(Rows.Count, "A").End(xlUp).ROW
    
    CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("bl" & BL).Range("B" & LineSta & ":B" & LineSto - 1))
    CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("bl" & BL).Range("C" & LineSta & ":C" & LineSto - 1))
    
    For i = LineSta To LineSto
'       Debug.Print "���̍s�@i = " & i & " ���A" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Rows(i).Interior.Color = RGB(0, 205, 0)
        
        
        
        If Not IsDateTimeFormatRegEx(Cells(i, 2)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, 2))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, 3)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, 3))
        End If

        
'        If Not IsValidFormat(Cells(i, 2), pattern) Then
'            Call CMsg("A�������`���ł͂���܂���B" & vbCrLf & "�������`��: YYYY/MM/DD HH:MM:SS", 3, Cells(i, 2))
'        End If
                    
'        If Not IsValidFormat(Cells(i, 3), pattern) Then
'            Call CMsg("B�������`���ł͂���܂���B" & vbCrLf & "�������`��: YYYY/MM/DD HH:MM:SS", 3, Cells(i, 3))
'        End If
        
        
        If (Cells(i, 3).Value - Cells(i, 2).Value) <= 0 Then
            Call CMsg("���Ԃ������������I�@END�̕����Â�" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
        
        
        If (Cells(i, 3).Value - Cells(LineSta, 2).Value) <= 0 Then
            Call CMsg("���Ԃ������������I�@���j�b�g�J�n�̎��Ԃ��Â������ł��B" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
        
        If InStr(Cells(i, 4).Value, "�v���O����") > 0 Or InStr(Cells(i, 4).Value, "FCBT") > 0 Or InStr(Cells(i, 4).Value, "��w�@") > 0 Or InStr(Cells(i, 4).Value, "���") > 0 Or InStr(Cells(i, 4).Value, "BL") > 0 Then
            Call CMsg("�ς����I�I�I" & vbCrLf & "FCBT�̉^�]��ʂ����[�U�[�^�]�ɂȂ��Ă鎖���m�F�B" & vbCrLf & "��ՊJ���v���O������A��w�@���v���O������BLstudy�ɂȂ�܂��I�I" & vbCrLf & "BL studey�����ꍞ��ł邼�I�I", vbExclamation, Cells(i, 4))
        End If
                
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
                
        If i = LineSto Then
            If (Cells(i, 3).Value - Cells(i, 2).Value) <> 14 Then
                Call CMsg("1���j�b�g�A2�T�Ԃ���Ȃ���ł���" & vbCrLf & "~~~", vbExclamation, Cells(i, 3))
            End If
            
            If Cells(i, 4).Value <> "" Then
                Call CMsg("�󗓂ł���ׂ��Ƃ���ɒl�����͂͂����Ă܂��B" & vbCrLf & "~~~", vbCritical, Cells(i, 4))
            End If
            
        End If

    Next

    
    Call CheckScheduleContinuity(wb_KEIKAKU.Worksheets("bl" & BL))


    Call Fin("�`�F�b�N�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
    
End Sub





Sub �v�掞��xlsx_GUN_HV_OFF_Check(BL As Integer)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim Retsu_GUN_HV_OFF As Integer
    Dim Retsu_GUN_HV_ON As Integer
    Dim result As Boolean
'    Dim pattern As String  �g��Ȃ�
'    pattern = "^\d{4}/\d{1,2}/\d{1,2}[ ]{1,2}\d{1,2}:\d{1,2}:\d{1,2}$"  ' �X�y�[�X�̐���1�A�܂���2�ł��}�b�`����悤�ɂ������ł��B

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
    Case 3
        Debug.Print ">>>BL3"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    MsgBox "�}�N���u�v�掞��xlsx_GUN_HV_OFF_Check�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & BNAME_KEIKAKU & vbCrLf & "�̃`�F�b�N�ł��B" & vbCrLf & "�m�F���܂�", vbInformation, "BL" & BL


    ' wb_KEIKAKU���J��
    Dim wb_KEIKAKU As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, True)    ' �t���p�X���w��
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    Debug.Print "�V�[�g�S�̂ɃG���[���Ȃ����m�F "
    Dim ws As Worksheet
    For Each ws In wb_KEIKAKU.Worksheets
        result = CheckForErrors(ws)
    Next ws

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("bl" & BL).Select    '�őO�ʂɕ\��


    wb_KEIKAKU.Worksheets("GUN HV OFF").Activate    '����厖
    LineSta = 3
    If BL = 2 Then
        LineSto = wb_KEIKAKU.Worksheets("GUN HV OFF").Cells(Rows.Count, "A").End(xlUp).ROW
        Retsu_GUN_HV_OFF = 1
        Retsu_GUN_HV_ON = 2
        CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("GUN HV OFF").Range("A" & LineSta & ":A" & LineSto))
        CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("GUN HV OFF").Range("B" & LineSta & ":B" & LineSto))
    Else
        LineSto = wb_KEIKAKU.Worksheets("GUN HV OFF").Cells(Rows.Count, "G").End(xlUp).ROW
        Retsu_GUN_HV_OFF = 7
        Retsu_GUN_HV_ON = 8
        CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("GUN HV OFF").Range("G" & LineSta & ":G" & LineSto))
        CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("GUN HV OFF").Range("H" & LineSta & ":H" & LineSto))
    End If
    
    
    For i = LineSta To LineSto
        'Debug.Print "���̍s�@i = " & i & " ���A" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        'Application.StatusBar = "Val:    " & i & "   " & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value

        Cells(i, Retsu_GUN_HV_OFF).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_GUN_HV_ON).Interior.Color = RGB(0, 205, 0)
                
        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_GUN_HV_OFF)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, 2))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_GUN_HV_ON)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, 3))
        End If
                
        
        If (Cells(i, Retsu_GUN_HV_ON).Value - Cells(i, Retsu_GUN_HV_OFF).Value) <= 0 Then
            Call CMsg("���Ԃ������������I�@END�̕����Â�" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
               
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
    Next



    Call Fin("�`�F�b�N�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
End Sub









Sub �^�]�W�v�L�^_Check(BL As String, sname As String)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim Retsu_end As Integer
    Dim Retsu_start As Integer
    Dim Retsu_chouseizikan As Integer
    Dim Retsu_total As Integer
    Dim result As Boolean
    Dim wb_name As String

    Select Case BL
    Case "SCSS"
        Debug.Print "SCSS+"
    Case "SACLA"
        Debug.Print "SACLA"
        wb_name = BNAME_UNTENSHUKEIKIROKU_SACLA
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    MsgBox "�}�N���u�^�]�W�v�L�^_Check�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & wb_name & vbCrLf & "�̃`�F�b�N�ł��B" & vbCrLf & "�m�F���܂�", vbInformation, "BL" & BL

    Retsu_end = 2
    Retsu_start = 3
    Retsu_chouseizikan = 4
    Retsu_total = 5

    Dim wb As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb = OpenBook(wb_name, True)    ' �t���p�X���w��
    wb.Activate
    If wb Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    If ActiveWorkbook.Name <> wb.Name Then
        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    'Debug.Print "�V�[�g�S�̂ɃG���[���Ȃ����m�F "
    'Dim ws As Worksheet
    'For Each ws In wb.Worksheets
    '    result = CheckForErrors(ws)
    'Next ws
    
    wb.Windows(1).WindowState = xlMaximized
    wb.Worksheets(sname).Select    '�őO�ʂɕ\��

    wb.Worksheets(sname).Activate    '����厖
    LineSta = 3
    LineSto = GetLastDataRow(wb.Worksheets(sname), "B") ' wb.Worksheets(sname).Cells(Rows.Count, "B").End(xlUp).ROW
    
    CheckAllDuplicatesByRange (wb.Worksheets(sname).Range("B" & LineSta & ":B" & LineSto))
    CheckAllDuplicatesByRange (wb.Worksheets(sname).Range("C" & LineSta & ":C" & LineSto))

    For i = LineSta To LineSto
        'Debug.Print "���̍s�@i = " & i & " ���A" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Cells(i, Retsu_end).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_start).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_total).Interior.Color = RGB(0, 205, 0)
        
        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_end)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, Retsu_end))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_start)) Then
            Call CMsg("�����̌`���ł͂���܂���B��������������t�I�����[��UNIXTIME�����B" & vbCrLf & "�Z���̏����ݒ�𕶎���ɂ���Ɗm�F�ł��܂��B", vbCritical, Cells(i, Retsu_start))
        End If
                
        
        If (Cells(i, Retsu_start).Value - Cells(i, Retsu_end).Value) <= 0 Then
            Call CMsg("���Ԃ������������I�@END�̕����Â�" & vbCrLf & "~~~", vbCritical, Cells(i, Retsu_start))
        End If
        
        If sname = "��~����" Then
            If Cells(i, Retsu_start).Value > ThisWorkbook.sheetS("�菇").Range("E" & UNITROW) Then ' ���j�b�g�J�n�������V�����Ƃ��낾���m�F
                If Cells(i, Retsu_chouseizikan) <> "" Then
                    Call CMsg("��(��������)�ɒ������R��������Ă��邱�Ƃ͂��܂肠��܂��񂪁A�A" & vbCrLf & "�m�F�������������ł�", vbExclamation, Cells(i, Retsu_chouseizikan))
                End If
            End If
        End If
        
        result = CheckSameFormulaType(Cells(LineSta, Retsu_total), Cells(i, Retsu_total))
        If result = False Then
            Debug.Print "�v�m�F�I�@�Z��(" & i & ", " & Retsu_total & ") �����������Ă��Ȃ����A�������قȂ�"
            Call CMsg("�����������Ă��Ȃ����A�������قȂ�I" & vbCrLf & "~~~", vbCritical, Cells(i, Retsu_total))
        End If
                           
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
    Next
        
    Call Fin("�`�F�b�N�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
End Sub





'======  ���t�p�^�[���}�b�`�@=======================================================
'  �Z���̏����ݒ�ŁA������ɂ���ƁA����UNIXTIME�ɂȂ�B
'�@���̊֐��ł́A���t�Ǝ��Ԃ�UNIXTIME��Ture�A���t�݂̂�UNIXTIME��False�ƂȂ�
Function IsDateTimeFormatRegEx(ByVal targetString As String) As Boolean
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 9:0:0")      ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/12/31 23:59:59")  ' True
'    Debug.Print IsDateTimeFormatRegEx("2025/7/9 1:52:43")   ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/01/01 09:00:00") ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/2/29 12:30:00")  ' True (���邤�N�l���Ȃ��A���t�̑Ó����͂��̐��K�\���ł͌����Ƀ`�F�b�N���Ȃ�)
'    Debug.Print IsDateTimeFormatRegEx("2023/13/01 00:00:00") ' False (����13)
'    Debug.Print IsDateTimeFormatRegEx("2023/01/32 00:00:00") ' False (����32)
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 24:0:0")     ' False (����24)
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 9:60:0")     ' False (����60)
'    Debug.Print IsDateTimeFormatRegEx("2023-1-1 9:0:0")     ' False (��؂蕶�����قȂ�)
'    Debug.Print IsDateTimeFormatRegEx("ABCDEFG")            ' False
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp") ' �܂��� New RegExp

    
    With regEx
        .pattern = "^\d{4}/(0?[1-9]|1[0-2])/(0?[1-9]|[12]\d|3[01])\s([01]?\d|2[0-3]):([0-5]?\d):([0-5]?\d)$"
        .IgnoreCase = False ' �啶���E����������ʂ��Ȃ��ꍇ��True
        .Global = False     ' ������S�̂ōŏ��̃}�b�`���O�݂̂���������ꍇ��False
                            ' ��������̂��ׂẴ}�b�`����������ꍇ��True
    End With

    IsDateTimeFormatRegEx = regEx.Test(targetString)

    Set regEx = Nothing
End Function





'======  A��ɗ\���ށAB��ɊJ�n���ԁAC��ɏI�����Ԃ��L�ڂ���Ă���ꍇ�ɁAB��̊J�n���Ԃ��O�̗\���C��̏I�����Ԃƈ�v���Ȃ��ꍇ�Ɍx����\���@=======================================================
Sub CheckScheduleContinuity(sheet As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim prevEndTime As Date
    
    lastRow = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).ROW - 1 ' �ŏI�s�̈�s��O�܂Ń`�F�b�N
    
    For i = 2 To lastRow
        Debug.Print "DEBUG: " & Cells(i, 2).Value
        ' �O�̗\��̏I�����Ԃ��擾
        If i > 2 Then
            If sheet.Cells(i, 2).Value <> prevEndTime Then
                Cells(i, 2).Font.Color = RGB(255, 5, 5)
                MsgBox "�x��: " & sheet.Cells(i, 1).Value & " �̊J�n���Ԃ��O�̗\��̏I�����Ԃƈ�v���܂���B", vbCritical
            End If
        End If
        
        ' ���݂̗\��̏I�����Ԃ�ۑ�
        prevEndTime = sheet.Cells(i, 3).Value
    Next i
End Sub

