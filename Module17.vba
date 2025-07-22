Module: Module17
' ///////////////////////////////////////////////////////////////////////////
' // ����I�ȃt�@�C���A�N�Z�X�m�F�p���W���[��          ����I�ɃT�[�o�[�ɒu���Ă�t�@�C�����A�N�Z�X�\���m�F                    //
' ///////////////////////////////////////////////////////////////////////////

' ���J�ϐ��F�Ď��Ώۂ̃t�@�C���p�X
Public Const TARGET_FILE_PATH As String = BNAME_MATOME ' ������ �������Ď��������t�@�C���̃p�X�ɏ��������Ă������� ������

' ���J�ϐ��F����̎��s�������i�[
Public NextRunTime As Date

' ���J�ϐ��F�O��̃t�@�C���A�N�Z�X��Ԃ��L�� (True:�A�N�Z�X�\, False:�A�N�Z�X�s��)
Private previousAccessStatus As Boolean

' ///////////////////////////////////////////////////////////////////////////
' // �֐��F�t�@�C���̃A�N�Z�X�ۂ��`�F�b�N����  �l�b�g���[�N��̃t�@�C���ɃA�N�Z�X�ł��邩�m�F�@�uMicrosoft Scripting Runtime�v���K�v======================================================
' ///////////////////////////////////////////////////////////////////////////
' VBA�G�f�B�^�Łu�c�[���v>�u�Q�Ɛݒ�v>�uMicrosoft Scripting Runtime�v�Ƀ`�F�b�N�����Ă��������B
Function CheckServerAccess_FSO(ByVal fullNetworkFilePath As String) As Boolean
    Dim fso As Object
    
    CheckServerAccess_FSO = False ' �f�t�H���g�ł͎��s�Ɛݒ�
    
    On Error GoTo ErrorHandler ' �G���[�n���h�����O��ݒ�
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fullNetworkFilePath) Then
        CheckServerAccess_FSO = True ' �A�N�Z�X����
        ThisWorkbook.sheetS("�菇").Range("B2").Value = "Connect"
        Debug.Print "�A�N�Z�X����==="
    End If
    
    Set fso = Nothing
    Exit Function ' ����I�����̓G���[�n���h�����X�L�b�v
    
ErrorHandler:
    ' �G���[�����������ꍇ�i��F�p�X���s���A�������Ȃ��Ȃǁj
    Debug.Print "�G���[���� (CheckServerAccess_FSO): " & Err.Description ' �f�o�b�O�p�ɃG���[���b�Z�[�W��\��
    CheckServerAccess_FSO = False ' �G���[���̓A�N�Z�X���s
    Set fso = Nothing
End Function

' ///////////////////////////////////////////////////////////////////////////
' // �T�u�v���V�[�W���F����I�Ɏ��s�����Ď�����                          //
' ///////////////////////////////////////////////////////////////////////////
Sub MonitorFileAccess()
    Dim currentAccessStatus As Boolean
    
    ' �t�@�C���̌��݂̃A�N�Z�X��Ԃ��m�F
    currentAccessStatus = CheckServerAccess_FSO(TARGET_FILE_PATH)
    
    ' Debug.Print Now & " - �A�N�Z�X���: " & currentAccessStatus & " (�O��: " & previousAccessStatus & ")" ' �f�o�b�O�p
    
    ' ������s���A�܂��̓A�N�Z�X��Ԃ��O�񂩂�ω������ꍇ
    If Not IsEmpty(previousAccessStatus) Then ' ������s���ȊO
        If currentAccessStatus <> previousAccessStatus Then
            If currentAccessStatus = False Then
                ' �A�N�Z�X�\��������Ԃ���A�N�Z�X�s�ɂȂ����ꍇ
                ThisWorkbook.sheetS("�菇").Range("B2").Value = "Not Connect"
                MsgBox "�ڑ����؂ꂽ�\��������܂�: �t�@�C���u" & TARGET_FILE_PATH & "�v�ɃA�N�Z�X�ł��܂���B", vbCritical + vbOKOnly, "�T�[�o�[�t�@�C���A�N�Z�X�x��"
            Else
                ' �A�N�Z�X�s��������Ԃ���A�N�Z�X�\�ɂȂ����ꍇ
                ThisWorkbook.sheetS("�菇").Range("B2").Value = "Connect"
                ' MsgBox "�ʒm: �t�@�C���u" & TARGET_FILE_PATH & "�v�ւ̃A�N�Z�X���񕜂��܂����B", vbInformation + vbOKOnly, "�T�[�o�[�t�@�C���A�N�Z�X��"
            End If
        End If
    End If
    
    ' ���݂̏�Ԃ�O��̏�ԂƂ��ċL��
    previousAccessStatus = currentAccessStatus
    
    ' ����̎��s������ݒ� (��: 5����)
    NextRunTime = Now + TimeValue("00:05:00") ' ������ �Ď��Ԋu�������Œ������Ă������� ������
    Application.OnTime NextRunTime, "MonitorFileAccess"
End Sub

' ///////////////////////////////////////////////////////////////////////////
' // �T�u�v���V�[�W���F�Ď����J�n����                                      //
' ///////////////////////////////////////////////////////////////////////////
Sub StartMonitoring()
    ' ������s���� previousAccessStatus ��ݒ肵�Ȃ� (IsEmpty)
    ' ���Ƀ^�C�}�[���ݒ肳��Ă���ꍇ�́A�����̃^�C�}�[���L�����Z�����Ă���J�n
    Call StopMonitoring ' �O�̂��߁A�����̃^�C�}�[���L�����Z��
    
    ' ����̃t�@�C���A�N�Z�X�`�F�b�N���s���ApreviousAccessStatus ��ݒ�
    previousAccessStatus = CheckServerAccess_FSO(TARGET_FILE_PATH)
    
    MsgBox "�t�@�C���A�N�Z�X�Ď����J�n���܂��B" & vbCrLf & _
           "�Ď��Ώ�: " & TARGET_FILE_PATH & vbCrLf & _
           "�Ď��Ԋu: 5������" & vbCrLf & _
           "����A�N�Z�X���: " & IIf(previousAccessStatus, "�\", "�s��"), vbInformation
    
    ' �ŏ��̊Ď��������Ɏ��s
    Call MonitorFileAccess
End Sub

' ///////////////////////////////////////////////////////////////////////////
' // �T�u�v���V�[�W���F�Ď����~����                                      //
' ///////////////////////////////////////////////////////////////////////////
Sub StopMonitoring()
    On Error Resume Next ' ���s���̃^�C�}�[���Ȃ��ꍇ�̃G���[�𖳎�
    Application.OnTime NextRunTime, "MonitorFileAccess", , False
    On Error GoTo 0 ' �G���[�n���h�����O�����ɖ߂�
    
    ' ��Ԃ����Z�b�g
    previousAccessStatus = Empty ' previousAccessStatus �����Z�b�g
    NextRunTime = 0 ' NextRunTime �����Z�b�g
    
    MsgBox "�t�@�C���A�N�Z�X�Ď����~���܂����B", vbInformation
End Sub

