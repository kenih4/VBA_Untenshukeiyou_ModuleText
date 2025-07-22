Module: Module15




Sub SACLA運転集計記録の確認()
'    If Check_checkbox_status("CheckBox_Untenshukeikiroku") = True Then
'        Debug.Print "チェックが入ってたので終了"
'        End
'    Else
'        Debug.Print "チェックが入なかったよ"
'    End If
'    MsgBox "GO"
    Call 運転集計記録_Check("SACLA", "停止時間")
    Call 運転集計記録_Check("SACLA", "調整時間")
End Sub


Sub check_Initial_Check_BL2_Click()
   Call Initial_Check(2)
End Sub

Sub check_Initial_Check_BL3_Click()
   Call Initial_Check(3)
End Sub







Sub 計画時間xlsxを出力_Click()
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer

    If MsgBox("出力するユニットは" & vbCrLf & "   「 " & ThisWorkbook.sheetS("手順").Range("D" & UNITROW) & " 」" & vbCrLf & ThisWorkbook.sheetS("手順").Range("E" & UNITROW) & "    ～　　" & vbCrLf & ThisWorkbook.sheetS("手順").Range("G" & UNITROW) & vbCrLf & "開始しますか？", vbYesNo) = vbNo Then Exit Sub
    
    fileNum = FreeFile
    Open OperationSummaryDir & "\dt_beg.txt" For Output As #fileNum
    Print #fileNum, Format(ThisWorkbook.sheetS("手順").Cells(UNITROW, 5).Value, "yyyy/mm/dd hh:nn");
    Close #fileNum
    
    fileNum = FreeFile
    Open OperationSummaryDir & "\dt_end.txt" For Output As #fileNum
    Print #fileNum, Format(ThisWorkbook.sheetS("手順").Cells(UNITROW, 7).Value, "yyyy/mm/dd hh:nn");
    Close #fileNum
    If RunPythonScript("getGunHvOffTime_LOCALTEST.py", OperationSummaryDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
    
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    
End Sub

Sub 計画時間xlsx_Check_BL2_Click()
    Call 計画時間xlsx_Check(2)
    Call 計画時間xlsx_GUN_HV_OFF_Check(2)
End Sub

Sub 計画時間xlsx_Check_BL3_Click()
    Call 計画時間xlsx_Check(3)
    Call 計画時間xlsx_GUN_HV_OFF_Check(3)
End Sub








Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI_BL2_Click()
    Call cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(2)
End Sub

Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI_BL3_Click()
    Call cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(3)
End Sub



Sub faulttxtを出力_BL2_Click()
    If RunPythonScript("getBlFaultSummary_LOCALTEST.py bl2", OperationSummaryDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
End Sub

Sub faulttxtを出力_BL3_Click()
    If RunPythonScript("getBlFaultSummary_LOCALTEST.py bl3", OperationSummaryDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
End Sub





Sub 利用時間Userに手動入力_BL2_Click()
    Call 利用時間Userに手動入力(2)
End Sub

Sub 利用時間Userに手動入力_BL3_Click()
    Call 利用時間Userに手動入力(3)
End Sub








Sub Fault集計m_BL2_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(2, "Fault集計m")
End Sub

Sub Fault集計m_BL3_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(3, "Fault集計m")
End Sub



Sub 運転集計_形式処理m_BL2_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(2, "運転集計_形式処理m")
End Sub

Sub 運転集計_形式処理m_BL3_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(3, "運転集計_形式処理m")
End Sub






Sub check_Final_Check_BL2_Click()
    Call Final_Check(2)
End Sub

Sub check_Final_Check_BL3_Click()
    Call Final_Check(3)
End Sub



Sub Clear_Click()
    Dim chk As Shape
    For Each chk In ActiveSheet.Shapes
        Debug.Print chk.Name
        If chk.Type = msoFormControl Then
            If chk.FormControlType = xlCheckBox Then
                chk.OLEFormat.Object.Value = xlOff    ' チェックを外す
            End If
        End If
    Next chk
End Sub
