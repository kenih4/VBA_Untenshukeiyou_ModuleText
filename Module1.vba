Module: Module1
Option Explicit

Sub �}�N�����낢��xlsm����SACLA�^�]�󋵏W�vBLxlsm�Ƀ}�N���𗬂�����Ŏ��s(BL As Integer, macroName As String)
    On Error GoTo ErrorHandler
    Debug.Print "Debug   �}�N�����낢��xlsm����SACLA�^�]�󋵏W�vBLxlsm�Ƀ}�N���𗬂�����Ŏ��s"
        
    Dim result As Boolean
    Dim sourceModule As Object
    Dim targetModule As Object
       
    'Dim BL As Integer
    Dim BNAME_SHUKEI As String
    'Dim macroName As String
    Dim vbComp As VBIDE.vbComponent
    
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.Add "Fault�W�vm", "Module10"
    dict.Add "�^�]�W�v_�`������m", "Module11"

'    Dim buttonName As String
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        Call Fin("���̃}�N���̓V�[�g��̃{�^������̂ݎ��s���Ă��������B" & vbCrLf & "�I�����܂��B", 3)
'    End If
'
'    If buttonName = "�{�^�� 1" Then
'        BL = 2
'        macroName = "Fault�W�vm"
'    ElseIf buttonName = "�{�^�� 2" Then
'        BL = 2
'        macroName = "�^�]�W�v_�`������m"
'    ElseIf buttonName = "�{�^�� 4" Then
'        BL = 3
'        macroName = "Fault�W�vm"
'    ElseIf buttonName = "�{�^�� 5" Then
'        BL = 3
'        macroName = "�^�]�W�v_�`������m"
'    Else
'        MsgBox "�ُ�ł��B�I�����܂��B" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    
    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\SACLA\SACLA�^�]�󋵏W�vBL" & BL & ".xlsm"

        
    
    ' sourceWorkbook���J��
    Dim sourceWorkbook As Workbook    ' �����Ɛ錾���Ȃ��ƁA�֐�SheetExists�̈������قȂ�Ɠ{����
    Set sourceWorkbook = OpenBook(BNAME_SOURCE, False) ' �t���p�X���w��
    If sourceWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    
'    For Each vbComp In sourceWorkbook.VBProject.VBComponents
'        Debug.Print "Debug   vbComp.name =  " & vbComp.Name & "     vbComp.Type: " & vbComp.Type
'    Next vbComp
'MOTO    Set sourceModule = sourceWorkbook.VBProject.VBComponents(dict(macroName)) ' ���W���[�������m�F       Module10 = Fault�W�vm()
        
    ' targetWorkbook���J��
    Dim targetWorkbook As Workbook
    Set targetWorkbook = OpenBook(BNAME_SHUKEI, False) ' �t���p�X���w��
    If targetWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)

    
    result = sourceWorkbook����targetWorkbook��moduleName�𗬂�����(sourceWorkbook, targetWorkbook, "Module8", "RunBatchFile", False) ' ���ʊ֐�
    If result Then
        Debug.Print "���� �usourceWorkbook����targetWorkbook��moduleName�𗬂����ށv�i���ʊ֐��j"
    Else
        Call Fin("���s �usourceWorkbook����targetWorkbook��moduleName�𗬂����ށv�i���ʊ֐��j", 3)
    End If
    
    result = sourceWorkbook����targetWorkbook��moduleName�𗬂�����(sourceWorkbook, targetWorkbook, dict(macroName), macroName, False)
    If result Then
        Debug.Print "���� �usourceWorkbook����targetWorkbook��moduleName�𗬂����ށv"
    Else
        Call Fin("���s �usourceWorkbook����targetWorkbook��moduleName�𗬂����ށv", 3)
    End If
    
    
'
    If MsgBox("�������񂾃}�N�������s���܂��B" & vbCrLf & "�����ł����H�H", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        Debug.Print "<<<<<<�u�b�N�u" & targetWorkbook.Name & "�v�@�́@�}�N���u" & macroName & "�v �����s���܂�"
        Application.RUN "'" & targetWorkbook.Name & "'!" & macroName, BL
        MsgBox "�}�N���u" & macroName & "�v ���������܂����I", Buttons:=vbInformation
        Debug.Print "�}�N���u" & macroName & "�v ���������܂���>>>>>>>>>>"
    End If


    '�}�N��macroName��ЂÂ���
    result = sourceWorkbook����targetWorkbook��moduleName�𗬂�����(sourceWorkbook, targetWorkbook, "Module8", "RunBatchFile", True) ' ���ʊ֐�
    result = sourceWorkbook����targetWorkbook��moduleName�𗬂�����(sourceWorkbook, targetWorkbook, dict(macroName), macroName, True)
    MsgBox "�������񂾃}�N���̕Еt�����I�����܂����B", Buttons:=vbInformation
    
    ' ���[�N�u�b�N�����
    'sourceWorkbook.Close SaveChanges:=False
    'targetWorkbook.Close SaveChanges:=True
    
   
    Call Fin("����ŏI���ł��B", 1)
    Exit Sub  ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    Call Fin("�G���[�ł��B���e�́@ " & Err.Description, 3)
    Exit Sub
    
    
End Sub





'Not USE    �}�N��macroName���AworkbookName�ɑ��݂��邩�m�F���āu���W���[���v���폜����  �Ԃ�l���~�����̂�Function�ɂ���===========================================================================
Function CheckAndDeleteModuleContainingMacro(WorkBookName As String, macroName As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim vbComponent As VBIDE.vbComponent
    Dim exists As Boolean

    ' �w�肵���u�b�N��ݒ�
    On Error Resume Next
    Set targetWorkbook = Workbooks.Open(WorkBookName) ' �w�肵���u�b�N���ŊJ���Ă��邩�m�F
    targetWorkbook.Windows(1).WindowState = xlMaximized
    On Error GoTo 0

    If targetWorkbook Is Nothing Then
        MsgBox "�w�肵���u�b�N '" & WorkBookName & "' ���J���Ă��܂���B"
        CheckAndDeleteModuleContainingMacro = False
        Exit Function
    End If

    ' ���W���[�������[�v
    exists = False
    For Each vbComponent In targetWorkbook.VBProject.VBComponents
        If vbComponent.Type = vbext_ct_StdModule Or vbComponent.Type = vbext_ct_ClassModule Then
            ' ���W���[������łȂ��ꍇ�̂݊m�F
            If vbComponent.CodeModule.CountOfLines > 0 Then
                ' ���W���[���̃R�[�h���m�F
                If InStr(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "Sub " & macroName & "(") > 0 Then
                    exists = True
                    ' ���W���[�����폜
                    targetWorkbook.VBProject.VBComponents.Remove vbComponent
                    Exit For
                End If
            End If
        End If
    Next vbComponent

    CheckAndDeleteModuleContainingMacro = exists
End Function













Function sourceWorkbook����targetWorkbook��moduleName�𗬂�����(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook, ByVal moduleName As String, ByVal macroName As String, ByVal ONLY_DELETE) As Boolean
    ' moduleName�ɂ́A�ǉ����郂�W���[���Ɋ܂܂��}�N�������w��B
    ' ���W���[����ǉ����邾���iONLY_ADD=TRUE�j�Ȃ�AtargetWorkbook�Ɋ��Ɋ܂܂�Ă��邩�̊m�F�Ɏg�����߁AmoduleName�Ɋ܂܂��}�N�����Ȃ�Ȃ�ł��������ASub�̕��ŁI�I�I�I�I�I�I�I�I�I�I�I�I
        
    On Error GoTo ErrorHandler
    Debug.Print "Debug   Start  sourceWorkbook����targetWorkbook��moduleName�𗬂�����"
    sourceWorkbook����targetWorkbook��moduleName�𗬂����� = True
            
    Dim sourceModule As Object
    Dim targetModule As Object
    Dim vbComp As VBIDE.vbComponent
        
    If sourceWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
    If targetWorkbook Is Nothing Then Call Fin("�u�b�N���J���܂���ł����B�p�X�̈قȂ铯�����O�̃u�b�N�����ɊJ����Ă�\��������܂��B", 3)
     
'    For Each vbComp In sourceWorkbook.VBProject.VBComponents
'        Debug.Print "Debug   vbComp.name =  " & vbComp.Name & "     vbComp.Type: " & vbComp.Type
'    Next vbComp
    Set sourceModule = sourceWorkbook.VBProject.VBComponents(moduleName) ' ���W���[����
    
    targetWorkbook.Windows(1).WindowState = xlMaximized
    '�}�N��macroName���ABNAME_SHUKEI�ɑ��݂�����A�폜����
    ' ���W���[�������[�v
    Dim exists As Boolean: exists = False
    For Each vbComp In targetWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule Then
            ' ���W���[������łȂ��ꍇ�̂݊m�F
            ' Debug.Print "Debug  ���W���[���� vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines) = " & vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            If vbComp.CodeModule.CountOfLines > 0 Then
                ' ���W���[���̃R�[�h���m�F
                If InStr(vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines), "Sub " & macroName & "(") > 0 Then
                    exists = True
                    ' ���W���[�����폜
                    targetWorkbook.VBProject.VBComponents.Remove vbComp
                    Debug.Print "Debug   ���W���[�������ɑ��݂����̂ŁA�폜���܂����I�I�I�I " & moduleName & "  " & macroName
                    Exit For
                End If
            End If
        End If
    Next vbComp
    
    If ONLY_DELETE = True Then Exit Function
        
    If exists Then
        MsgBox "�}�N�� �u" & macroName & "�v ���܂܂�郂�W���[��[" & moduleName & "] �� " & vbCrLf & targetWorkbook.Name & " �ɑ��݂����̂ŁA��U�A���W���[�����폜���āA" & vbCrLf & "�}�N���𗬂����݂܂��B�B", Buttons:=vbInformation
    Else
        MsgBox "�}�N�� �u" & macroName & "�v ���܂܂�郂�W���[��[" & moduleName & "] �� " & vbCrLf & targetWorkbook.Name & vbCrLf & "�ɗ������݂܂��B", Buttons:=vbInformation
    End If
    
    Set targetModule = targetWorkbook.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule = 1  �W�����W���[����ǉ�
    targetModule.CodeModule.AddFromString sourceModule.CodeModule.Lines(1, sourceModule.CodeModule.CountOfLines)
    Debug.Print "Debug   targetWorkbook�ɁA[" & moduleName & "]��ǉ����܂����I"
    
    Debug.Print "Debug   Function sourceWorkbook����targetWorkbook��moduleName�𗬂�����    ��  �I��"

    Exit Function ' �ʏ�̏���������������G���[�n���h�����X�L�b�v
ErrorHandler:
    MsgBox "�G���[�ł��B���e�́@ " & Err.Description, Buttons:=vbCritical
    sourceWorkbook����targetWorkbook��moduleName�𗬂����� = False
    
End Function



