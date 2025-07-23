Module: Module5
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Final_Check(BL As Integer)
'
    On Error GoTo ErrorHandler

    Dim tc As Variant
    Dim i, j As Integer
    Dim col As Variant
    Dim BNAME_SHUKEI As String
    Dim DOWNTIME_ROW As Integer
    Dim UNIT As String
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim Check_col_arr As Variant
    Dim result As Boolean
    Dim pattern As String
    Dim CantFindUnit As Integer: CantFindUnit = 0

    MsgBox "�}�N���uFinal_Check()�v�����s���܂��B" & vbCrLf & "���̃}�N���́A" & vbCrLf & BNAME_MATOME & vbCrLf & "�̃`�F�b�N�ł��B" & vbCrLf & "�`�F�b�N���郆�j�b�g���m�F����ׂɈ�U�ASACLA�^�]�󋵏W�vBL*.xlsm���J���܂�", vbInformation, "BL" & BL

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

    UNIT = ThisWorkbook.sheetS("�菇").Range("D" & UNITROW)

'    'wb_SHUKEI���J��  [���j�b�g]���m�F���邽��
'    Dim wb_SHUKEI As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
'    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' �t���p�X���w��
'    If wb_SHUKEI Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
'    wb_SHUKEI.Activate
'    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
'        Call Fin("���݃A�N�e�B�u�ȃu�b�N�����ُ�ł��B�I�����܂��B" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
'    End If
'
'    wb_SHUKEI.Windows(1).WindowState = xlMaximized
'    wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Activate
'
'    Unit = wb_SHUKEI.Worksheets("���p���ԁi���ԁj").Range("B2")
'
'    If MsgBox("�`�F�b�N�Ώۂ̃��j�b�g(�V�[�g�u���p���ԁi���ԁj�v�̃Z��B2)��    " & vbCrLf & "�u " & Unit & " �v" & vbCrLf & "�ł��B " & vbCrLf & "�ԈႢ�Ȃ��ł����H" & vbCrLf & "�i�ނɂ�YES�������ĉ�����", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
'        Call Fin("�uNo�v���I������܂����B�I�����܂��B", 1)
'    End If


    ' wb_MATOME���J��
    Dim wb_MATOME As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set wb_MATOME = OpenBook(BNAME_MATOME, True)    ' �t���p�X���w��
    If wb_MATOME Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    Application.WindowState = xlMaximized

    Debug.Print "�V�[�g�S�̂ɃG���[���Ȃ����m�F "
    Dim ws As Worksheet
    For Each ws In wb_MATOME.Worksheets
        Debug.Print ws.Name
        result = CheckForErrors(ws)
    Next ws



    wb_MATOME.Worksheets("Fault�W�v").Activate    '����厖
    MsgBox "Fault�W�v�V�[�g���`�F�b�N���܂��B" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("SACLA Fault�Ԋu(BL2)", 2, wb_MATOME.Worksheets("Fault�W�v"))
        LineSto = getLineNum("SACLA Fault�Ԋu(BL3)", 2, wb_MATOME.Worksheets("Fault�W�v"))
    Else
        LineSta = getLineNum("SACLA Fault�Ԋu(BL3)", 2, wb_MATOME.Worksheets("Fault�W�v"))
        LineSto = wb_MATOME.Worksheets("Fault�W�v").Cells(Rows.Count, "B").End(xlUp).ROW
    End If

    For i = LineSta To LineSto
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value
        If wb_MATOME.Worksheets("Fault�W�v").Cells(i, 2).Value = UNIT Then
            Debug.Print "���̍s�@i = " & i & " ���A���j�b�g " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("Fault�W�v").Cells(i, 2).MergeArea.Rows.Count - 1

                Check_col_arr = Array(3, 4, 5, 6, 7, 8, 9)  'Check_col_arr = Array(3, 4, 7, 8) ' �`�F�b�N�����̒l��z��ɃZ�b�g  �V�t�g�J�n�A�I���AFaul�Ԋu�AFaul��
                For Each col In Check_col_arr
                    Set tc = wb_MATOME.Worksheets("Fault�W�v").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)
                    'Sleep 100    ' msec
                    If tc.MergeArea.Columns.Count > 1 Or tc.MergeArea.Rows(1).ROW <> j Then
                        Debug.Print "���������Ɍ�������Ă�A�܂��́A���������Ɍ�������Ă��Đ擪�ł��B" & i & "   " & j & "   " & col
                    Else
                        If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("�󗓁A�܂��́A�G���[���������Ă��܂�", 3, tc)
                        Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  �Z����#DIV/0!���� �����A�\���Ŏ��s����̂ŁA���̑O��IsCellError�Ń`�F�b�N����

                        If col = 3 Or col = 4 Then    ' �V�t�g����
                            result = CheckDateTimeFormat(tc)
                        End If

                        If col = 5 And (tc.Value <= 0 Or tc.Value > 8.2 Or Not IsNumeric(tc.Value)) Then  '�G�l���M�[
                            Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                        End If
    
                        If col = 6 And (tc.Value <= 0 Or tc.Value > 25 Or Not IsNumeric(tc.Value)) Then  '�g��
                            Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                        End If

                        If col = 7 Then  'Fault�Ԋu����
                            result = CheckTimeFormat(tc)
                        End If

                        If col = 8 And (tc.Value < 0 Or Not IsNumeric(tc.Value)) Then  'Fault��
                            Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                        End If

                        If col = 9 And (StrComp(Right(tc.Value, 1), "G", vbBinaryCompare) = 0 = False) Then  ' ������1������ "G" ���ǂ����`�F�b�N�i�啶���E����������ʁj
                            Call CMsg("���[�U�[�������锤�Ȃ̂�G������܂����", 3, tc)
                        End If

                    End If
                Next col
            Next
            Exit For
        End If
    Next


    

    wb_MATOME.Worksheets("�܂Ƃ� ").Activate    '����厖======================================================================================

    MsgBox "�܂Ƃ߃V�[�g��(a)�̃`�F�b�N���܂��B" & vbCrLf & "", vbInformation, "BL" & BL
    For i = getLineNum("(a)�^�]���ԁ@���Ԗ�", 2, wb_MATOME.Worksheets("�܂Ƃ� ")) To getLineNum("(b)�^�]���ԁ@�V�t�g��", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, 2).Value = UNIT Then
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            If BL = 2 Then
                DOWNTIME_ROW = i
            Else    'BL3
                DOWNTIME_ROW = i + 1
            End If

            Check_col_arr = Array(3, 5, 6, 7, 9, 10, 11, 12)    ' �`�F�b�N�����̒l��z��ɃZ�b�g
            For Each col In Check_col_arr
                If col >= 9 Then
                    Set tc = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, col)
                Else
                    Set tc = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, col)
                End If
                tc.Select
                tc.Interior.Color = RGB(0, 255, 0)
                'Sleep 100    ' msec
                If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("�󗓁A�܂��́A�G���[���������Ă��܂�", 3, tc)
                Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  �Z����#DIV/0!���� �����A�\���Ŏ��s����̂ŁA���̑O��IsCellError�Ń`�F�b�N����

                If col = 3 Then    ' ���t
                    pattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2} - \d{4}/\d{2}/\d{2} \d{2}:\d{2}$"    '       �ʂ̏����i��: YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM�j pattern = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2} - \d{4}-\d{2}-\d{2} \d{2}:\d{2}$"
                    If Not IsValidFormat(tc, pattern) Then
                        Call CMsg("�Z�� " & tc.Address(False, False) & " �̒l���������`���ł͂���܂���B" & vbCrLf & "�������`��: YYYY/MM/DD HH:MM - YYYY/MM/DD HH:MM", 3, tc)
                    End If
                End If

                If col = 5 Or col = 6 Or col = 7 Or col = 9 Or col = 10 Or col = 11 Or col = 12 Then    '���^�]����(�v��j(�v��, �_�E���^�C��), ���p�����^�](�v��, �_�E���^�C��) , ���p�^�](�v��, �_�E���^�C��)
                    result = CheckTimeFormat(tc)
                End If

            Next col


            If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 9).Value <= 0 Then
                Call CMsg("���p�����^�](BL����orBL-study)�͂Ȃ�������ł��ˁB", 2, wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 9))
            End If

            If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 11).Value <= 0 Then
                Call CMsg("���p�^�](���[�U�[)�͂Ȃ�������ł��ˁB" & vbCrLf & "�u���[�U�[�^�]�����v�Ǝ蓮�ŏ������Ȃ��Ƃ����Ȃ�����������܂��B", 2, wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 11))
            Else
                If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 12).Value <= 0 Then
                    Call CMsg("�����g���b�v���ĂȂ����Ď��H�m�F���������悢�ł��B" & vbCrLf & "�V�[�g�u�W�v�L�^�v�ɐ����������Ă��Ȃ��\��������܂�", 2, wb_MATOME.Worksheets("�܂Ƃ� ").Cells(DOWNTIME_ROW, 12))
                End If
            End If
        End If
    Next






    MsgBox "�܂Ƃ߃V�[�g��(b)�̃`�F�b�N�B" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("(b-1)BL2", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
        LineSto = getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
    Else
        LineSta = getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
        LineSto = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(Rows.Count, "B").End(xlUp).ROW
    End If

    Check_col_arr = Array(3, 4, 5, 6, 7, 8)    ' �`�F�b�N�����̒l��z��ɃZ�b�g  �V�t�g����(�J�n�E�I���E�Ԋu)�A���p���A�r�[���������ԁA�_�E���^�C��

    For i = LineSta To LineSto
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, 2).Value = UNIT Then
            Debug.Print "���̍s�@i = " & i & " ���A���j�b�g " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, 2).MergeArea.Rows.Count - 1
                For Each col In Check_col_arr
                    Set tc = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)
                    'Sleep 100    ' msec
                    If tc.MergeArea.Columns.Count > 1 Then
                        Debug.Print "���������Ɍ�������Ă��܂��B" & i & "   " & j & "   " & col & "     tc.Value = " & tc.Value & "  " & tc.MergeArea.Columns.Count
                    Else

                        If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("�󗓁A�܂��́A�G���[���������Ă��܂�", 3, tc)
                        Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  �Z����#DIV/0!���� �����A�\���Ŏ��s����̂ŁA���̑O��IsCellError�Ń`�F�b�N����

                        If col = 5 And (wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 3).Value = "total" And (StrComp(Right(wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 9).Value, 1), "G", vbBinaryCompare) = 0) = False) Then  '' ������1������ "G" ���ǂ����`�F�b�N�i�啶���E����������ʁj
                            Call CMsg("���[�U�[��������ׂ��ł����B�B�m�F�������������ł��B", 2, wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 9))
                        End If

                        If col = 3 Or col = 4 Then
                            result = CheckDateTimeFormat(tc)
                        End If

                        If col = 5 Or col = 7 Or col = 8 Then
                            result = CheckTimeFormat(tc)
                        End If
    
                        If (col = 5 And wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 3).Value <> "total") And (tc.Value <= 0 Or tc.Value > 0.5 Or Not IsNumeric(tc.Value)) Then
                            Call CMsg("�͈͊O��������Ȃ��ł��B�m�F�������������ł��B", 3, tc)
                        End If

                        If col = 6 And (tc.Value <= 0.8 Or tc.Value > 1 Or Not IsNumeric(tc.Value)) Then  '���p��%
                            Call CMsg("���p���Ⴂ�B�܂��́A�͈͊O or ������   �m�F�������������ł��B", 3, tc)
                        End If

                    End If
                Next col

            Next
            Exit For
        End If
    Next





    MsgBox "�܂Ƃ߃V�[�g��(c)�̃`�F�b�N�B" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("(c-1)BL2", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
        LineSto = getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
    Else
        LineSta = getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets("�܂Ƃ� "))
        LineSto = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(Rows.Count, "B").End(xlUp).ROW
    End If


    For i = LineSta To LineSto
        Debug.Print "DEBUG D    i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, 2).Value = UNIT Then
            Debug.Print "���̍s�@i = " & i & " ���A���j�b�g " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("�܂Ƃ� ").Cells(i, 2).MergeArea.Rows.Count - 1

                For col = 3 To 7
                    Set tc = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)

                    'Sleep 100    ' msec
                    If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("�󗓁A�܂��́A�G���[���������Ă��܂�", 3, tc)
                    Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  �Z����#DIV/0!���� �����A�\���Ŏ��s����̂ŁA���̑O��IsCellError�Ń`�F�b�N����

                    If col = 3 And (tc.Value <= 0 Or tc.Value > 8.2 Or Not IsNumeric(tc.Value)) Then  '�G�l���M�[
                        Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                    End If

                    If col = 4 And (tc.Value <= 0 Or tc.Value > 60 Or Not IsNumeric(tc.Value)) Then  '�J�Ԃ�
                        Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                    End If
    
                    If col = 5 And (tc.Value <= 0 Or tc.Value > 25 Or Not IsNumeric(tc.Value)) Then  '�g��
                        Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
    
                        If InStr(1, tc.Value, "+", vbTextCompare) > 0 Then '�g��
                            Call CMsg("�Z���ɂ́u+�v���܂܂�Ă��܂��B", 2, tc)
                            If MsgBox("���l�Z���ɁA�u�A��F�����v�ƒǂ��������݂܂����H" & vbCrLf & "�����ł��H", vbYesNo + vbQuestion, "�m�F") = vbYes Then
                    '            MsgBox j & "  Cells(j, 7).Value:     " & wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 7).Value, Buttons:=vbInformation
                                wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 7).Value = wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 7).Value + "�A��F����"
                    '            MsgBox "�ǂ��������݂����B     " & wb_MATOME.Worksheets("�܂Ƃ� ").Cells(j, 7).Value & vbCrLf & "���ɐi�݂܂��B", Buttons:=vbInformation
                            End If
                        End If

                    End If

                    If col = 6 And (tc.Value <= 0 Or tc.Value > 2000 Or Not IsNumeric(tc.Value)) Then  '���x
                        Call CMsg("�͈͊O or �񐔒l�ł��B�m�F�������������ł��B", 3, tc)
                    End If

                    If col = 7 And (IsNumeric(tc.Value)) Then  '���l
                        Call CMsg("���l�ł��B�m�F�������������ł��B", 3, tc)
                    End If

                Next

            Next
            Exit For
        End If
    Next




    If CantFindUnit <> 4 Then
        MsgBox "�ُ�ł��B" & vbCrLf & "�`�F�b�N�Ώۂ̃��j�b�g  CantFindUnit : " & CantFindUnit & " ��������܂���ł����B�S����ׂ��ł��B", Buttons:=vbCritical
    End If





    Call Fin("�I�����܂����B" & vbCrLf & "", 1)
    Exit Sub    ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical

End Sub

























Function IsCellErrorType(Target As Variant) As Boolean
    If IsError(Target.Value) Then
        Select Case Target.Value
        Case CVErr(xlErrDiv0)
            IsCellErrorType = False    'IsCellErrorType = "#DIV/0! �G���["
        Case CVErr(xlErrNA)
            IsCellErrorType = False    'IsCellErrorType = "#N/A �G���["
        Case CVErr(xlErrValue)
            IsCellErrorType = False    'IsCellErrorType = "#VALUE! �G���["
        Case Else
            IsCellErrorType = False    'IsCellErrorType = "���̑��̃G���["
        End Select
        Call CMsg("���̃Z���ŃG���[���������Ă��܂��B@IsCellErrorType", 3, Target)
    Else
        IsCellErrorType = True    'IsCellErrorType = "�G���[�Ȃ�"
    End If
End Function







Function CheckDateTimeFormat(Target As Variant) As Boolean
    Dim compareDate As Date
    CheckDateTimeFormat = False
    If IsDate(Target.Value) Then
        If Format(Target.Value, "yyyy/mm/dd hh:mm") <> Target.Text Then
            Call CMsg("�t�H�[�}�b�g������������܂���B@CheckDateTimeFormat" & vbCrLf & "�������`��: 2025/01/28 22:00", 3, Target)
        Else
            CheckDateTimeFormat = True
            compareDate = DateSerial(2025, 1, 1) + TimeSerial(12, 30, 0)
            If Target.Value < compareDate Then
                MsgBox Target.Value & " ���A " & compareDate & " ���O�ł��B�m�F�������������ł��B", vbExclamation
            End If
        End If
    Else
        Call CMsg("�L���ȓ��t�����͂���Ă��܂���B@CheckDateTimeFormat", 3, Target)
    End If
End Function



Function CheckTimeFormat(Target As Variant) As Boolean
    Debug.Print "CheckTimeFormat         target.Value = " & Target.Value
    CheckTimeFormat = False
    If Not IsNumeric(Target.Value) Or Target.Value < 0 Then
        Debug.Print "�L���Ȏ��Ԃ����͂���Ă��܂���B@CheckTimeFormat    target.Value = " & Target.Value
        Call CMsg("�L���Ȏ��Ԃ����͂���Ă��܂���B@CheckTimeFormat", 3, Target)
    Else
        If IsDate(CDate(Target.Value)) Then
            Dim fmt As String
            fmt = Target.NumberFormat
            'Debug.Print "�t�H�[�}�b�g�́@     target.Value = " & target.Value & "  fmt = " & fmt
            If fmt = "h:mm" Or fmt = "hh:mm" Or fmt = "[h]:mm" Or fmt = "h:mm;@" Or fmt = "hh:mm;@" Then    ' [h]:mm�͗ݐώ���
                'Debug.Print "�����f�[�^�Ő������t�H�[�}�b�g�ł��B    target.Value = " & target.Value
                CheckTimeFormat = True
            Else
                Debug.Print "�����f�[�^�ł����A�t�H�[�}�b�g���قȂ�܂��B    target.Value = " & Target.Value
                Call CMsg("�����f�[�^�ł����A�t�H�[�}�b�g���قȂ�܂��B@CheckTimeFormat", 3, Target)
            End If
        End If
    End If
End Function












'--------------------------------------------------------------------------------------------------------------------------------------------
' �Z���̒l���w�肵���p�^�[���Ɉ�v���邩�`�F�b�N����֐�
Function IsValidFormat(cell As Variant, pattern As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    regEx.pattern = pattern
    regEx.IgnoreCase = True
    regEx.Global = False

    ' ���K�\�����}�b�`���邩�𔻒�
    IsValidFormat = regEx.Test(cell.Value)

    ' �I�u�W�F�N�g���
    Set regEx = Nothing
End Function









