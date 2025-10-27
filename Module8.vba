Module: Module8
Option Explicit ' ����`�̕ϐ��͎g�p�ł��Ȃ��悤��
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const UNITROW As Integer = 18
Public Const CPATH As String = "\\saclaopr18.spring8.or.jp\common\�^�]�󋵏W�v\�ŐV\"
Public Const WHICH As String = "SACLA"
Public Const BNAME_UNTENSHUKEIKIROKU_SACLA As String = CPATH & WHICH & "\SACLA�^�]�W�v�L�^.xlsm"
Public Const BNAME_KEIKAKU As String = CPATH & "�v�掞��.xlsx"
Public Const BNAME_SOURCE As String = "C:\me\unten\�}�N�����낢��.xlsm"
Public Const OperationSummaryDir As String = "C:\me\unten\OperationSummary"
Public Const BNAME_MATOME As String = CPATH & WHICH & "\SACLA�^�]�󋵏W�v�܂Ƃ�.xlsm"

Public Const edgePath As String = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


'�u�b�N���J�� ==============================================================================================================================
'�����[�g�T�[�o�[��̃t�@�C�����J���ہA�J���̂Ɏ��Ԃ��������Ă��邽�߂Ƀ^�C�~���O�̖��ŃG���[���������Ă���\��������܂��B���̏ꍇ�A�ҋ@���Ԃ�݂��čĎ��s���邱�ƂŁA�G���[������ł��邱�Ƃ�����܂��B�ȉ��̕��@�ŁA�w�肳�ꂽ���ԑҋ@���Ȃ���G���[���Ď��s����R�[�h�������ł��܂��B
Function OpenBook(ByVal WorkBookName As String, ByVal RO As Boolean) As Workbook
    
    Debug.Print "Debug---   Start  Function OpenBook(" & WorkBookName & ")"
    Dim OWB As Workbook
    Dim wb As Workbook
    Dim retryCount As Integer
    retryCount = 3  ' �Ď��s�̉�

    ' �J���Ă���u�b�N�̒��Ɏw�肳�ꂽ�p�X�̃u�b�N�����邩���m�F
    For Each wb In Workbooks
        'Debug.Print "Debug   wb.Name =  " & wb.Name & " �͊J����Ă��܂�"
        If wb.FullName = WorkBookName Then
            Set OWB = wb
            Debug.Print "Debug---   OpenBook.Name =  [" & OWB.Name & "] �͊��ɊJ����Ă��܂�"
            Exit For
       End If
    Next wb
    
    On Error Resume Next
    If OWB Is Nothing Then
        Do While retryCount > 0
            Set OWB = Workbooks.Open(WorkBookName, ReadOnly:=RO)
            If Err.Number = 0 Then Exit Do  ' ����ɊJ�����烋�[�v�𔲂���
            Debug.Print "Debug--- �G���[���������܂����B�Ď��s���܂��B�c��Ď��s��: " & retryCount - 1
            Err.Clear
            retryCount = retryCount - 1
            Application.Wait Now + TimeValue("0:00:05")  ' 5�b�ҋ@
        Loop

        ' �Ō�ɃG���[���c���Ă���ꍇ�̑Ή�
        If Err.Number <> 0 Then
            MsgBox "�u�b�N��������Ȃ����A�J���܂���ł����B�G���[�ԍ�: " & Err.Number & vbCrLf & _
                   "�G���[���b�Z�[�W: " & Err.Description & vbCrLf & _
                   "�t�@�C������p�X���m�F���Ă�������: " & WorkBookName, vbExclamation
            Set OWB = Nothing
            Err.Clear
        Else
            Debug.Print "Debug---   OpenBook.Name =  [" & OWB.Name & "] ���J���܂���"
        End If
    End If
    On Error GoTo 0  ' �G���[�n���h�����O����
    
    Set OpenBook = OWB
    
    Debug.Print "Debug---   Finish  Function OpenBook(" & WorkBookName & ")"
End Function









' �g���ĂȂ�
Function OpenBookOLD(ByVal WorkBookName As String) As Workbook
    Debug.Print "Debug   �u�b�N���J���܂��B-----------  " & WorkBookName
    Dim OWB As Workbook
    Dim wb As Workbook

    ' �J���Ă���u�b�N�̒��Ɏw�肳�ꂽ�p�X�̃u�b�N�����邩���m�F
    For Each wb In Workbooks
        'Debug.Print "Debug   wb.Name =  " & wb.Name & " �͊J����Ă��܂�"
        If wb.FullName = WorkBookName Then
            Set OWB = wb
            Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] �͊��ɊJ����Ă��܂�"
            Exit For
        End If
    Next wb

    ' �G���[�n���h�����O�J�n
    On Error Resume Next
    If OWB Is Nothing Then
        ' �w�肵���u�b�N���J����Ă��Ȃ��ꍇ�A�V���ɊJ�����Ƃ���
        Set OWB = Workbooks.Open(WorkBookName, ReadOnly:=False)    ' SACLA�^�]�󋵏W�vBL*.xlsm�@���J�����Ƃ���ƁA�Ȃ����G���[����������̂ňȉ��R�����g�A�E�g����
        If Err.Number <> 0 Then
            ' �G���[�����������ꍇ�A�G���[���b�Z�[�W��\��
            MsgBox "�u�b�N��������Ȃ����A�J���܂���ł����B�G���[�ԍ�: " & Err.Number & vbCrLf & _
                   "�G���[���b�Z�[�W: " & Err.Description & vbCrLf & _
                   "�t�@�C������p�X���m�F���Ă�������: " & WorkBookName, vbExclamation
            Set OWB = Nothing  ' �G���[�������� Nothing ��Ԃ�
            Err.Clear
        Else
            Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] ���J���܂���"
        End If
        Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] ���J���܂���   �J���Ă��Ȃ��\������@�G���[�������p�X���Ă�̂�"
    End If
    On Error GoTo 0  ' �G���[�n���h�����O����

    ' �֐��̖߂�l�Ƃ��Đݒ�
    Set OpenBook = OWB

    Debug.Print "Debug   OpenBook Finish"
End Function





'========================================================================================================
Sub CMsg(ByVal msg As String, ByVal Level As Integer, tc As Variant)

    Debug.Print "_____Msg(" & msg & ")_____"

    tc.Select
    tc.Font.Bold = True
    Select Case Level
    Case vbInformation
        tc.Font.Color = RGB(0, 200, 0)
        tc.Interior.Color = RGB(0, 255, 255)
        MsgBox msg, vbInformation, "���m�点"
    Case vbExclamation
        tc.Interior.Color = RGB(255, 255, 0)
        MsgBox msg, vbExclamation, "����"
    Case vbCritical
        tc.Interior.Color = RGB(255, 0, 0)
        MsgBox msg, vbCritical, "�x��"
    Case Else
        Debug.Print "Zzz..."
    End Select
    
End Sub


'========================================================================================================
Sub Fin(ByVal msg As String, ByVal Level As Integer)

    Debug.Print "_____Fin(" & msg & ")_____"
    Select Case Level
        Case 1
            MsgBox msg, vbInformation, "�I������"
        Case 2
            MsgBox msg, vbExclamation, "�I������"
        Case 3
            MsgBox msg, vbCritical, "�I������"
        Case Else
            Debug.Print "Zzz..."
    End Select
    
'    ActiveWindow.Zoom = 100
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFullScreen = False
    ' �J���Ă��邷�ׂẴu�b�N�����[�v
    Dim wb As Workbook
    For Each wb In Workbooks
        wb.Windows(1).Zoom = 100 ' �e�u�b�N�̃E�B���h�E�ɑ΂��ăY�[����ݒ�
    Next wb
'    End   ���ꂢ��H�H�H
End Sub





'----------------------------------------------------------------------------------------------------------------------
'�V�[�g���̃G���[�Z�������o���A���b�Z�[�W��\������
Function CheckForErrors(ByVal sheet As Worksheet) As Boolean
  Dim cell As Range
  Dim errorRange As Range
  CheckForErrors = False
  
  If sheet Is Nothing Then
    MsgBox "' �̃V�[�g '" & sheet & "' �͑��݂��܂���B", vbOKOnly + vbCritical
    Exit Function
  End If
  sheet.Activate
  
  For Each cell In sheet.UsedRange
    'Debug.Print "Debug  Value =  " & cell.Value & "  Row = " & cell.Row & " Columuns = " & cell.Column
    If IsError(cell.Value) Then
      ' �ŏ��̃G���[�Z���ł���΁AerrorRange�ɐݒ�
      If errorRange Is Nothing Then
        Set errorRange = cell
      Else
        ' 2�ڈȍ~�̃G���[�Z���ł���΁AerrorRange�ɒǉ�
        Set errorRange = Union(errorRange, cell)
        cell.Select
      End If
    End If
  Next cell

  ' �G���[�Z�������������ꍇ�A���b�Z�[�W��\��
  If Not errorRange Is Nothing Then
        MsgBox "�V�[�g '" & sheet.Name & "' �ɃG���[�Z��������܂��B" & vbCrLf & "�G���[�Z��: " & errorRange.Address, vbOKOnly + vbCritical
  Else
'    MsgBox "���S�ł��B�V�[�g '" & sheet.Name & "' �ɃG���[�Z���͂���܂���ł����B", vbOKOnly + vbInformation
        Debug.Print "���S�ł��B�V�[�g '" & sheet.Name & "' �ɃG���[�Z���͂���܂���ł����B"
        CheckForErrors = False
  End If

  Set errorRange = Nothing
End Function




'�w�肳�Ă������񂪑��݂���s���擾  �V�[�g���S��==============================================================================================================================
Function getLineNum(ByVal str As String, ByVal TARGET_COL As Integer, ByVal sheet As Worksheet) As Integer
    getLineNum = getLineNum_RS(str, TARGET_COL, 1, sheet.Cells(Rows.Count, TARGET_COL).End(xlUp).ROW, sheet)
End Function


'�w�肳�Ă������񂪑��݂���s���擾 Range Specification��==============================================================================================================================
Function getLineNum_RS(ByVal str As String, ByVal TARGET_COL As Integer, ByVal beginLine As Integer, ByVal endLine As Integer, ByVal sheet As Worksheet) As Integer
    Dim i As Integer: i = -1
    getLineNum_RS = i
    For i = beginLine To endLine
        'Debug.Print "getLineNum_RS�@�s�ԍ�: " & i & "    Value: " & Cells(i, 2).Value
        If sheet.Cells(i, TARGET_COL).Value = str Then ' #DIV/0!�Ȃǂ̃G���[�Z��������ƁA�������r���Ŏ~�܂�܂��B
            getLineNum_RS = i
            Debug.Print "Hit!!!!!!!!!!!!!!!!!!!!!!!!!   getLineNum_RS�@�s�ԍ�: " & i & "    Value: " & Cells(i, 2).Value
            Exit Function
        End If
    Next
'    Call Fin("@getLineNum_RS    ������u" & str & "�v�ƈ�v����Z���͌�����܂���ł����B", 3)
    MsgBox "@getLineNum_RS    ������u" & str & "�v�ƈ�v����Z���͌�����܂���ł����B", vbExclamation, "�x��"
End Function




'�V�[�g���݂��m�F==============================================================================================================================
Function SheetExists(wb As Workbook, sname As String) As Boolean
    On Error Resume Next ' �G���[���������Ă��������p��
    Dim ws As Worksheet
    Set ws = wb.sheetS(sname) ' �w�肵���V�[�g���Z�b�g
    SheetExists = Not ws Is Nothing ' �V�[�g�����݂����True
    Debug.Print "@SheetExists   Sheetname: [" & sname & "]  " & SheetExists
    On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g
End Function






'ActiveWorkbook�V�[�g���݂��m�F Not Use ==============================================================================================================================
Function SheetExist_ActiveWorkbook(ByVal WorkSheetName As String) As Boolean
  Dim sht As Worksheet
  For Each sht In ActiveWorkbook.Worksheets
    If sht.Name = WorkSheetName Then
        flgExsistSheet = True
        Exit Function
    End If
  Next sht
  flgExsistSheet = False
End Function




'==============================================================================================================================
Sub RunBatchFile(batchFilePath As String)

    ' �o�b�`�t�@�C���̃p�X���w�肳��Ă��邩�m�F
    If batchFilePath = "" Then
        MsgBox "�o�b�`�t�@�C���̃p�X���w�肵�Ă�������", vbExclamation
        Exit Sub
    End If
    
    ' Shell�֐��Ńo�b�`�t�@�C�������s
    shell batchFilePath, vbNormalFocus
End Sub




'�G�N�Z���u�b�N���J���ꂽ�t�H���_���擾==============================================================================================================================
Function GetWorkbookFolder() As String
    Dim folderPath As String
    
    ' �u�b�N���ۑ�����Ă��Ȃ��ꍇ�APath �͋󕶎���ɂȂ�
    folderPath = ThisWorkbook.path
    
    ' �ۑ�����Ă��Ȃ��ꍇ�A�󕶎����Ԃ�
    If folderPath = "" Then
        GetWorkbookFolder = "" ' �󕶎����Ԃ�
    Else
        GetWorkbookFolder = folderPath ' �t�H���_�p�X��Ԃ�
    End If
End Function


Sub GetWorkbookFolderToCell()
' ThisWorkbook.Path �ŃJ�����g�u�b�N�̕ۑ�����Ă���p�X���擾
    Dim folderPath As String
    folderPath = ThisWorkbook.path
'    MsgBox folderPath
    
    If folderPath <> "" Then
        ThisWorkbook.sheetS("�菇").Range("B1").Value = folderPath
        
'        MsgBox folderPath
        If folderPath = "C:\me\unten" Then
            MsgBox "OK: " & vbCrLf & "���[�L���O�t�H���_ = " & folderPath, Buttons:=vbInformation
        Else
            MsgBox "�`�F�b�N: " & vbCrLf & "���[�L���O�t�H���_ = " & folderPath & vbCrLf & "���[�L���O�t�H���_���uC:\me\unten�v�ł���܂���I�I", Buttons:=vbInformation
        End If
        
    Else
        ThisWorkbook.sheetS(1).Range("A1").Value = "���[�L���O�t�H���_���擾�ł��܂���ł���"
        MsgBox "�ُ�: " & vbCrLf & "���[�L���O�t�H���_���擾�ł��܂���ł���", Buttons:=vbCritical
    End If
End Sub






' �z�Q�Ƃ����o
Sub CheckCircularReference()
    Application.Calculate ' ��Ɍv�Z�����s  Application.CircularReference �́A�v�Z��ɒl��Ԃ����߁A�v�Z���܂����s����Ă��Ȃ��ꍇ�ɂ͎g���Ȃ���

    On Error Resume Next ' �G���[�𖳎�
    Dim circRef As Range
    Set circRef = Application.CircularReference
    On Error GoTo 0 ' �G���[������߂�

    If circRef Is Nothing Then
'        MsgBox "�z�Q�Ƃ͌�����܂���ł����B", vbInformation
    Else
        MsgBox "�z�Q�Ƃ�������܂���: " & circRef.Address, vbExclamation
    End If
End Sub







Sub ToggleButton()    '---------------------------------------------------------------------------------
' �{�^���̊O�ς�ύX����
    If ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(255, 255, 255) Then
        ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(0, 0, 0)  ' ���ɕύX
        ActiveSheet.Shapes("Button 18").TextFrame.Characters.Text = "�������ݒ�"
    Else
        ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(255, 255, 255)  ' ���ɖ߂�
        ActiveSheet.Shapes("Button 18").TextFrame.Characters.Text = "�����Ă�������"
    End If
End Sub






' �V�[�g�ɕ����񂪑��݂��邩�m�F����
Function CheckStringInSheet(ByVal ws As Worksheet, ByVal searchString As String) As Boolean
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If foundCell Is Nothing Then
        CheckStringInSheet = False
    Else
        CheckStringInSheet = True
    End If
    
End Function







Function Check_checkbox_status(obj_name) As Boolean
    Dim chk As Shape
    Check_checkbox_status = False
    For Each chk In ActiveSheet.Shapes
'        Debug.Print chk.Name
        If chk.Type = msoFormControl Then
            If chk.FormControlType = xlCheckBox Then
                Debug.Print "Checked checkbox:  " & chk.Name
                If chk.Name = obj_name Then
                    Debug.Print "Checked!  True @Check_checkbox_status"
                    'chk.OLEFormat.Object.Value = xlOff
                    Check_checkbox_status = True
                End If
            Else
                'Debug.Print "Checked checkbox:  " & chk.Name
            End If
        End If
    Next chk
End Function





'=== git�Ǘ��������ׁA���W���[�����ƂɕʁX�̃e�L�X�g�t�@�C���ɃG�N�X�|�[�g���� ===
Sub ExportModulesToSeparateTextFiles()
    Dim vbComp As Object
    Dim filePath As String
    Dim fileNum As Integer
    Dim moduleContent As String
    
    ' �o�̓t�@�C���̃p�X���w��i�����t�H���_�ɕۑ��j
'    filePath = ThisWorkbook.path & "\"
    filePath = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText\"
    
    ' �e���W���[�������[�v���ē��e���擾
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' ���W���[���̍s�����擾
        Dim lineCount As Long
        lineCount = vbComp.CodeModule.CountOfLines
        
        ' ���W���[������łȂ��ꍇ�̂ݓ��e���擾
        If lineCount > 0 Then
            ' ���W���[���̓��e���擾
            moduleContent = vbComp.CodeModule.Lines(1, lineCount)
            
            ' ���W���[�����Ɋ�Â��ăt�@�C�������쐬
            Dim moduleFileName As String
            moduleFileName = filePath & vbComp.Name & ".vba"
            
            ' �V�����e�L�X�g�t�@�C�����쐬
            fileNum = FreeFile
            Open moduleFileName For Output As #fileNum
            
            ' ���W���[�����Ƃ��̓��e���t�@�C���ɏ�������
            Print #fileNum, "Module: " & vbComp.Name
            Print #fileNum, moduleContent
            
            ' �t�@�C�������
            Close #fileNum
        Else
            ' ��̃��W���[���̏ꍇ�̓X�L�b�v
            Debug.Print "���W���[�� " & vbComp.Name & " �͋�ł��B"
        End If
    Next vbComp
     
    
    ThisWorkbook.SaveCopyAs "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText" & "\�}�N�����낢��_" & Format(Date, "yyyymmdd") & ".xlsm"
    
    If MsgBox("���ׂẴ��W���[�����e�L�X�g�ɃG�N�X�|�[�g����܂����B" & vbCrLf & filePath & vbCrLf & "vscode���J���܂����H" & vbCrLf & "git add -A" & vbCrLf & "git commit -m comment", vbYesNo + vbQuestion, "�m�F") = vbNo Then
        MsgBox "No"
    Else
        Dim vscodePath As String
        Dim folderPath As String
        Dim command As String
        vscodePath = "C:\Users\kenic\AppData\Local\Programs\Microsoft VS Code\Code.exe"
        folderPath = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText"
        command = """" & vscodePath & """ """ & folderPath & """"
        shell command, vbNormalFocus
    End If
        
End Sub





'=== �t���p�X�ŏ����ꂽ�t�@�C�����̊g���q��������菜������ ================================
Function RemoveFileExtension(fullPath As String) As String
    Dim fileName As String
    Dim dotPosition As Long
    
    ' �t���p�X����t�@�C�������擾
    'fileName = Dir(fullPath)
    fileName = fullPath
    
    ' �Ō�̃h�b�g�̈ʒu���擾
    dotPosition = InStrRev(fileName, ".")
    
    ' �h�b�g�����������ꍇ�A�g���q����菜��
    If dotPosition > 0 Then
        RemoveFileExtension = Left(fileName, dotPosition - 1)
    Else
        ' �h�b�g��������Ȃ��ꍇ�́A���̃t�@�C������Ԃ�
        RemoveFileExtension = fileName
    End If
End Function





'=== �l�������Ă�ŏI�s���擾������ ================================
'     �ŏI�s�ԍ��i�Ⴆ��1048576�s�ځj��Ԃ��\��������֐��ɂ́ALong �łȂ��� �I�[�o�[�t���[�̊댯������܂��B
Function GetLastDataRow(ws As Worksheet, colName As String) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim colNum As Long
    colNum = ws.Range(colName & "1").Column
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).ROW

    For i = lastRow To 1 Step -1
'        If ws.Cells(i, colNum).HasFormula Then
            If ws.Cells(i, colNum) = "" Then
'                Debug.Print i & "   KARA  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
            Else
'                Debug.Print i & "   Not KARA  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
                GetLastDataRow = i ' �l�������Ă�ꍇ�A�R�R��
                Exit Function
            End If
'        Else
            'Debug.Print i & "   Not HasFormula  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
'        End If
    Next i

    GetLastDataRow = 0 ' �Y���Z���������ꍇ
End Function





'=== 2�̃Z������v����̂����Ȃ��̂��B�����̔�r���s���ۂɂ́A���ɏ������������e������@���L�� ================================
' ���t�̒l�A�Z��A1�ɂ́u45769.66667�v�A�Z��A2�ɂ��������u45769.66667�v�������Ă���̂ł����A��v���܂���B�^��Ɏv���āA�����ɁA�Z��A2 �� �Z��A1�̍������Ƃ�����A�u-7.27695761418343E-12�v�ƂȂ�܂����B
Function CheckCellsMatch(cell1 As Range, cell2 As Range) As Boolean
    Dim tolerance As Double

    ' ���e�덷��ݒ�
    tolerance = 0.0000000001 ' 10��-10���菬�������͖���

    If Abs(cell1.Value - cell2.Value) < tolerance Then
        CheckCellsMatch = True ' ��v���Ă���ꍇ
    Else
        CheckCellsMatch = False ' ��v���Ă��Ȃ��ꍇ
    End If
End Function

'=== 2�̒l����v����̂����Ȃ��̂��B�����̔�r���s���ۂɂ́A���ɏ������������e������@���L�� ================================
Function CheckValMatch(v1 As Variant, v2 As Variant) As Boolean
    Dim tolerance As Double

    ' ���e�덷��ݒ�
    tolerance = 0.0000000001 ' 10��-10���菬�������͖���

    If Abs(v1 - v2) < tolerance Then
        CheckValMatch = True ' ��v���Ă���ꍇ
    Else
        CheckValMatch = False ' ��v���Ă��Ȃ��ꍇ
    End If
End Function



'=== �͈͓��̂��ׂĂ̏d���l�����o���āA�܂Ƃ߂Čx�����b�Z�[�W�ŕ\�� ======================================
Sub CheckAllDuplicatesByRange(targetRange As Range)
    Dim cell As Range
    Dim dict As Object
    Dim duplicates As Collection
    Set dict = CreateObject("Scripting.Dictionary")
    Set duplicates = New Collection
    targetRange.Select
    
    For Each cell In targetRange
        If Not IsEmpty(cell.Value) Then
            If dict.exists(cell.Value) Then
                duplicates.Add cell
            Else
                dict.Add cell.Value, cell
            End If
        End If
    Next cell
    
    If duplicates.Count > 0 Then
        Dim msg As String
        msg = "�ȉ��̏d��������܂�:" & vbCrLf
        
        Dim dupCell As Range
        For Each dupCell In duplicates
            msg = msg & dupCell.Address(False, False) & ": " & dupCell.Value & vbCrLf
            dupCell.Interior.Color = RGB(255, 200, 200)
        Next dupCell
        
        MsgBox msg, vbCritical, "�d�����o����"
    Else
        MsgBox "�d���͂���܂���ł���", vbInformation, "�m�F����"
    End If
End Sub




'=== �V�K�t�H���_�쐬 ======================================
Function CreateFolderWithAutoRename(parentPath As String, newFolderName As String) As String
    Dim fullPath As String
    Dim counter As Long
    Dim candidateName As String

    ' �ŏ��̌��
    candidateName = newFolderName
    fullPath = parentPath & "\" & candidateName

    ' �t�H���_�����݂������A�A�Ԃ�t���Ď���
    Do While Dir(fullPath, vbDirectory) <> ""
        counter = counter + 1
        candidateName = newFolderName & "_(" & counter & ")"
        fullPath = parentPath & "\" & candidateName
    Loop

    On Error GoTo ErrorHandler
    MkDir fullPath
    CreateFolderWithAutoRename = fullPath
    Exit Function

ErrorHandler:
    CreateFolderWithAutoRename = vbNullString  ' �쐬���s���͋󕶎���Ԃ�
End Function



'=== �t�@�C���R�s�[ ======================================
Function CopyFileSafely(sourcePath As String, destinationPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' �R�s�[���s�iTrue�ŏ㏑�����j
    fso.CopyFile sourcePath, destinationPath, True
    CopyFileSafely = True
    Exit Function

ErrorHandler:
    CopyFileSafely = False
End Function


'=== �ǂݎ���p������ݒ� ======================================
Function SetFileReadOnly(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    If Dir(filePath, vbNormal) = "" Then
        SetFileReadOnly = False  ' �t�@�C�������݂��Ȃ�
        Exit Function
    End If

    SetAttr filePath, vbReadOnly
    SetFileReadOnly = True
    Exit Function

ErrorHandler:
    SetFileReadOnly = False
End Function

