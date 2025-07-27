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
    '    s = Application.InputBox("BLを入力して下さい。", "確認", Type:=1)    '  Type:=1 数値のみ
    '    If s = False Then
    '        Exit Sub
    '    ElseIf s = "" Then
    '        MsgBox "何も入力されていません"
    '        Exit Sub
    '    Else
    '        BL = s
    '    End If

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SCSS\SCSS運転状況集計BL1.xlsm"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2TEST.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    '    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2TEST.xlsm"
    MsgBox "マクロ「Initial_Check()」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & BNAME_SHUKEI & vbCrLf & "のチェックです。" & vbCrLf & "数式が入っているべきセルに数式が入っているか確認します", vbInformation, "BL" & BL

    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    wb_SHUKEI.Windows(1).WindowState = xlMaximized

    Debug.Print "シート全体にエラーがないか確認 "
    Dim ws As Worksheet
    For Each ws In wb_SHUKEI.Worksheets
        result = CheckForErrors(ws)
    Next ws
    
    
    If Check_exixt("運転予定時間", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets("運転予定時間"))
    If Check_exixt("GUN HV OFF時間記録", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7), 3, 30, wb_SHUKEI.Worksheets("GUN HV OFF時間記録"))
    If Check_exixt("GUN HV OFF時間記録", wb_SHUKEI) = True Then Cnt = Check(Array(9, 10, 11, 12, 13, 14, 15), 9, 30, wb_SHUKEI.Worksheets("GUN HV OFF時間記録"))
    If Check_exixt("集計記録", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 6, 7, 8, 9), 3, 500, wb_SHUKEI.Worksheets("集計記録")) ' とりあえず500行くらいチェック    E列(Fault)もチェックしたいが、ここは特殊　最終行の2行目から変な数式が入ってるがいるのか？
    If Check_exixt("利用時間（期間）", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14), 2, 30, wb_SHUKEI.Worksheets("利用時間（期間）")) ' 利用時間（期間） のカッコは全角
    If Check_exixt("利用時間（期間）", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 25, 26, 27, 28, 29), 2, 30, wb_SHUKEI.Worksheets("利用時間(User)"))
    If Check_exixt("利用時間(シフト)", wb_SHUKEI) = True Then Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), 1, 30, wb_SHUKEI.Worksheets("利用時間(シフト)"))
    If Check_exixt("Fault間隔(ユニット)", wb_SHUKEI) = True Then Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 2, 30, wb_SHUKEI.Worksheets("Fault間隔(ユニット)")) ' シート名に半角スペースが入ってることがあるので注意　正しくしないと「インデックスが有効範囲にありません」とエラーメッセージがでける
    
    'シートの存在を確認する処理を追加するとこうなるが、見にくい。。。。。
'    sname = "運転予定時間"
'    If Not SheetExists(wb_SHUKEI, sname) Then
'        MsgBox "シートが存在しません。" & vbCrLf & sname & " 終了します。", Buttons:=vbExclamation
'    Else
'        If CheckStringInSheet(wb_SHUKEI.Worksheets(sname), ThisWorkbook.sheetS("手順").Range("D" & UNITROW)) Then
'            wb_SHUKEI.Worksheets(sname).Activate
'            MsgBox "今から出力しようとしているユニットが既にシート上に存在しますけど、、、　確認して下さい。　 ", Buttons:=vbCritical
'        Else
'            Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets(sname))
'        End If
'    End If


'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), 2, 30, wb_SHUKEI.Worksheets("運転予定時間"))
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7), 3, 30, wb_SHUKEI.Worksheets("GUN HV OFF時間記録"))
'    Cnt = Check(Array(9, 10, 11, 12, 13, 14, 15), 9, 30, wb_SHUKEI.Worksheets("GUN HV OFF時間記録"))
'    Cnt = Check(Array(2, 3, 4, 6, 7, 8, 9), 3, 500, wb_SHUKEI.Worksheets("集計記録")) ' とりあえず500行くらいチェック    E列(Fault)もチェックしたいが、ここは特殊　最終行の2行目から変な数式が入ってるがいるのか？
'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14), 2, 30, wb_SHUKEI.Worksheets("利用時間（期間）")) ' 利用時間（期間） のカッコは全角
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 25, 26, 27, 28, 29), 2, 30, wb_SHUKEI.Worksheets("利用時間(User)"))
'    Cnt = Check(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), 1, 30, wb_SHUKEI.Worksheets("利用時間(シフト)"))
'    Cnt = Check(Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 2, 30, wb_SHUKEI.Worksheets("Fault間隔(ユニット)")) ' シート名に半角スペースが入ってることがあるので注意　正しくしないと「インデックスが有効範囲にありません」とエラーメッセージがでける
        
    Call Fin("終了しました。" & vbCrLf & "", 1)
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    
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
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SCSS\SCSS運転状況集計BL1.xlsm"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    MsgBox "マクロ「Middle_Check()」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & BNAME_SHUKEI & vbCrLf & "の中間チェックです。" & vbCrLf & "ユーザー運転の開始終了時刻などの確認します", vbInformation, "BL" & BL

    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    
    If ThisWorkbook.sheetS("手順").Range("D" & UNITROW).Value <> wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2").Value Then
        If MsgBox("シート「利用時間（期間）」に入力されてるユニットと 違 い ま す けど、進めますか？", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
            Exit Sub
        End If
    End If

    wb_SHUKEI.Worksheets("運転予定時間").Select    '最前面に表示_______________________________________________________________________________
    wb_SHUKEI.Worksheets("運転予定時間").Activate
    ' シート[運転予定時間]のE列最終行は336.0等と時間で表示されているが、中身は日。
    If Int(Cells(GetLastDataRow(wb_SHUKEI.Worksheets("運転予定時間"), "E"), "E")) <> Int(ThisWorkbook.sheetS("手順").Range("I" & UNITROW)) Then
        Call CMsg("シート「運転予定時間」のE列最終行とユニット合計時間が一致しません" & vbCrLf & Int(Cells(GetLastDataRow(wb_SHUKEI.Worksheets("運転予定時間"), "E"), "E")) & " と " & Int(ThisWorkbook.sheetS("手順").Range("I" & UNITROW)), vbCritical, Cells(GetLastDataRow(wb_SHUKEI.Worksheets("運転予定時間"), "E"), "E"))
    Else
        Call CMsg("一致、OK!!" & vbCrLf & vbCrLf & "ユニット合計時間が一致", vbInformation, Cells(GetLastDataRow(wb_SHUKEI.Worksheets("運転予定時間"), "E"), "E"))
    End If
    
    
    
    wb_SHUKEI.Worksheets("利用時間（期間）").Select    '最前面に表示_______________________________________________________________________________
    wb_SHUKEI.Worksheets("利用時間（期間）").Activate
    Set ws = wb_SHUKEI.Worksheets("利用時間（期間）")
    If MsgBox("今のユニットだけ確認しますか？" & vbCrLf & "Yes:　[" & ThisWorkbook.sheetS("手順").Range("D" & UNITROW) & "]だけ確認" & vbCrLf & "No:　全ユニット確認", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("手順").Range("D" & UNITROW), 3, ws)
    Else
        LineSta = 4
    End If
    LineSto = GetLastDataRow(ws, "A")
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間（期間）").Range("A" & LineSta & ":A" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間（期間）").Range("B" & LineSta & ":B" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間（期間）").Range("C" & LineSta & ":C" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間（期間）").Range("D" & LineSta & ":D" & LineSto))
    For i = LineSta To LineSto
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        
        If Not CheckValMatch(ws.Cells(i, "F").Value + ws.Cells(i, "H").Value + ws.Cells(i, "J").Value, ws.Cells(i, "E").Value) Then  ' 「合計時間」の確認
            Call CMsg("「合計時間」が一致しません" & ws.Cells(i, "F").Value & "   " & ws.Cells(i, "H").Value & "   " & ws.Cells(i, "J").Value & "   E=" & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "J"), ws.Cells(i, "M")) Then
            Call CMsg("[利用運転計画]が一致しません", vbCritical, Cells(i, "J"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "E"), ws.Cells(i, "N")) Then
            Call CMsg("[総運転時間]が一致しません", vbCritical, Cells(i, "N"))
        End If
        
        If ws.Cells(i, "G").Value > ws.Cells(i, "F").Value Or ws.Cells(i, "G").Value < 0 Then ' 「施設調整計画」の確認
            Call CMsg("「施設調整計画」が「施設調整計画ダウンタイム」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "G"))
        End If

        If ws.Cells(i, "I").Value > ws.Cells(i, "H").Value Or ws.Cells(i, "I").Value < 0 Then ' 「利用調整計画」の確認
            Call CMsg("「利用調整計画」が「利用調整計画ダウンタイム」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "I"))
        End If
                
        If ws.Cells(i, "K").Value > ws.Cells(i, "J").Value Or ws.Cells(i, "K").Value < 0 Then ' 「利用運転計画」の確認
            Call CMsg("「利用運転計画」が「利用運転計画ダウンタイム」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
        
    Next
    
    
    
    wb_SHUKEI.Worksheets("配列").Select    '最前面に表示_______________________________________________________________________________
    wb_SHUKEI.Worksheets("配列").Activate
    If GetLastDataRow(wb_SHUKEI.Worksheets("集計記録"), "C") <> Cells(4, "E").Value Then
        Call CMsg("シート「集計記録」の最終行と一致しません" & vbCrLf & "", vbCritical, Cells(4, "E"))
    Else
        Call CMsg("一致、OK!!" & vbCrLf & vbCrLf & vbCrLf & "シート「集計記録」の最終行と一致", vbInformation, Cells(4, "E"))
    End If
 
 
 
 
    wb_SHUKEI.Worksheets("利用時間(シフト)").Select    '最前面に表示_______________________________________________________________________________
    wb_SHUKEI.Worksheets("利用時間(シフト)").Activate
    Set ws = wb_SHUKEI.Worksheets("利用時間(シフト)")
    If MsgBox("今のユニットだけ確認しますか？" & vbCrLf & "Yes:　[" & ThisWorkbook.sheetS("手順").Range("D" & UNITROW) & "]だけ確認" & vbCrLf & "No:　全ユニット確認", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("手順").Range("D" & UNITROW), 2, ws)
    Else
        LineSta = 9
    End If
    LineSto = GetLastDataRow(ws, "B")
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間(シフト)").Range("A" & LineSta & ":A" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間(シフト)").Range("C" & LineSta & ":C" & LineSto))
    CheckAllDuplicatesByRange (wb_SHUKEI.Worksheets("利用時間(シフト)").Range("D" & LineSta & ":D" & LineSto))
    
    For i = LineSta To LineSto
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        
        If Not IsDateTimeFormatRegEx(Cells(i, "C")) Or Not IsDateTimeFormatRegEx(Cells(i, "D")) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, "C"))
        End If
        
        If (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) <> ws.Cells(i, "E").Value Then ' 「合計時間」の確認
            Call CMsg("「合計時間」が一致しません   " & vbCrLf & "    差分：" & (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) & "   E列:" & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
        End If
        
        If ws.Cells(i, "F").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "F").Value < 0 Then ' 「利用時間」の確認
            Call CMsg("「利用時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "F"))
        End If
        
        If ws.Cells(i, "G").Value > 100 Or ws.Cells(i, "G").Value < 0 Then  ' 「利用率」の確認
            Call CMsg("「利用率」が  0 ~ 100%の範囲でない   " & vbCrLf & "====", 3, Cells(i, "G"))
        End If
        
        If ws.Cells(i, "H").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "H").Value < 0 Then ' 「調整時間」の確認
            Call CMsg("「調整時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "H"))
        End If
        
        If ws.Cells(i, "I").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "I").Value < 0 Then ' 「Fault時間」の確認
            Call CMsg("「Fault時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "I"))
        End If
        
        If ws.Cells(i, "J").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "J").Value < 0 Then ' 「ダウンタイム」の確認
            Call CMsg("「ダウンタイム」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "J"))
        End If
        
        If ws.Cells(i, "K").Value < 0 Then  ' 「Fault合計」の確認
            Call CMsg("「Fault合計」が  負" & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
            
        If ws.Cells(i, "L").Value > ws.Cells(i, "E").Value Or ws.Cells(i, "L").Value < 0 Then ' 「Fault間隔」の確認
            Call CMsg("「Fault間隔」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "L"))
        End If
    
        If IsNumeric(ws.Cells(i, "M").Value) Or InStr(Cells(i, "M"), "G") = 0 Then  ' 「ユーザー」の確認
            Call CMsg("「ユーザー」が数値、または、ユーザー名なのに「G」がない、" & vbCrLf & "====", vbExclamation, Cells(i, "M"))
        End If
        
    Next
    
    
    wb_SHUKEI.Worksheets("利用時間(User)").Select    '最前面に表示_______________________________________________________________________________
    wb_SHUKEI.Worksheets("利用時間(User)").Activate
    Set ws = wb_SHUKEI.Worksheets("利用時間(User)")
'    CheckForErrors (ws)
    If MsgBox("今のユニットだけ確認しますか？" & vbCrLf & "Yes:　[" & ThisWorkbook.sheetS("手順").Range("D" & UNITROW) & "]だけ確認" & vbCrLf & "No:　全ユニット確認", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        LineSta = getLineNum(ThisWorkbook.sheetS("手順").Range("D" & UNITROW), 2, ws)
    Else
        LineSta = 9
    End If
'   LineSto = ws.Cells(wb_SHUKEI.Worksheets("利用時間(User)").Rows.Count, "B").End(xlUp).ROW ' 列Bの最下行から上方向にデータを探すので、空白があっても無視できます。 これだと数式が入ってると無理
    LineSto = GetLastDataRow(ws, "B")
    
    For i = LineSta To LineSto
'       Debug.Print "この行　i = " & i & " が、" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Rows(i).Select
        Rows(i).Interior.Color = RGB(0, 255, 0)
        If Not IsDateTimeFormatRegEx(Cells(i, "C")) Or Not IsDateTimeFormatRegEx(Cells(i, "D")) Or Not IsDateTimeFormatRegEx(Cells(i, "E")) Or Not IsDateTimeFormatRegEx(Cells(i, "F")) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, "C"))
        Else
            If Not CheckCellsMatch(ws.Cells(i, "C"), ws.Cells(i, "E")) Then
                Call CMsg("日時が一致しません   " & vbCrLf & "" & ws.Cells(i, "C").Value & vbCrLf & ws.Cells(i, "E").Value, vbCritical, Cells(i, "E"))
            End If
            If Not CheckCellsMatch(ws.Cells(i, "D"), ws.Cells(i, "F")) Then
                Call CMsg("日時が一致しません   " & vbCrLf & "" & ws.Cells(i, "D").Value & vbCrLf & ws.Cells(i, "F").Value, vbCritical, Cells(i, "F"))
            End If
        End If
        
        If Not CheckValMatch(ws.Cells(i, "D").Value - ws.Cells(i, "C").Value, ws.Cells(i, "G").Value) Then    ' 「合計時間」の確認
            Call CMsg("「合計時間」が一致しません   " & vbCrLf & "    差分：" & (ws.Cells(i, "D").Value - ws.Cells(i, "C").Value) & "   G列:" & ws.Cells(i, "G").Value, vbCritical, Cells(i, "G"))
        End If
        
        If ws.Cells(i, "H").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "H").Value < 0 Then ' 「利用時間」の確認
            Call CMsg("「利用時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "H"))
        End If
        
        If ws.Cells(i, "I").Value > 100 Or ws.Cells(i, "I").Value < 0 Then  ' 「利用率」の確認
            Call CMsg("「利用率」が  0 ~ 100%の範囲でない   " & vbCrLf & "====", 3, Cells(i, "I"))
        End If
        
        If ws.Cells(i, "J").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "J").Value < 0 Then ' 「調整時間」の確認
            Call CMsg("「調整時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "J"))
        End If
        
        If ws.Cells(i, "K").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "K").Value < 0 Then ' 「Fault時間」の確認
            Call CMsg("「Fault時間」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "K"))
        End If
        
        If ws.Cells(i, "L").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "L").Value < 0 Then ' 「ダウンタイム」の確認
            Call CMsg("「ダウンタイム」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "L"))
        End If
        
        If ws.Cells(i, "M").Value < 0 Then  ' 「Fault合計」の確認
            Call CMsg("「Fault合計」が  負" & vbCrLf & "====", vbCritical, Cells(i, "M"))
        End If

        If ws.Cells(i, "N").Value > ws.Cells(i, "G").Value Or ws.Cells(i, "N").Value < 0 Then ' 「Fault間隔」の確認
            Call CMsg("「Fault間隔」が「合計時間」よりも大きい、または、負  " & vbCrLf & "====", vbCritical, Cells(i, "N"))
        End If
    
        If IsNumeric(ws.Cells(i, "O").Value) Or InStr(Cells(i, "O"), "G") = 0 Then  ' 「ユーザー」の確認
            Call CMsg("「ユーザー」が数値、または、ユーザー名なのに「G」がない、" & vbCrLf & "====", vbExclamation, Cells(i, "O"))
        End If
        
        If Not CheckCellsMatch(ws.Cells(i, "G"), ws.Cells(i, "W")) Then
            Call CMsg("「ユーザー運転時間（計画）」が一致しません   ", vbCritical, Cells(i, "W"))
        End If
    Next

    Call Fin("終了しました。" & vbCrLf & "", 1)
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    
End Sub











Function Check_exixt(sname As String, wb As Workbook) As Boolean

    Check_exixt = False
    If Not SheetExists(wb, sname) Then
        MsgBox "シートが存在しません。" & vbCrLf & sname & " 終了します。", Buttons:=vbCritical
    Else
        If CheckStringInSheet(wb.Worksheets(sname), ThisWorkbook.sheetS("手順").Range("D" & UNITROW)) Then
            wb.Worksheets(sname).Activate
            MsgBox "今から出力しようとしているユニットが既にシート上に存在しますけど、、、　確認して下さい。　 ", Buttons:=vbCritical
        Else
            Check_exixt = True
        End If
    End If
    
End Function




'後で、要確認！
'VBAでは、明示的に ByVal も ByRef も指定しない場合、デフォルトで ByRef（参照渡し）になります。
'つまり､引数として渡した変数の値が変更される可能性がある ので注意が必要です｡
'Function Check(arr As Variant, ByVal Retsu_for_Find_last_row As Integer, ByVal Check_row_cnt As Integer, ByVal sheet As Worksheet) As Integer
' StartL , EndLを引数にした方がいい気がする
Function Check(arr As Variant, Retsu_for_Find_last_row As Integer, Check_row_cnt As Integer, sheet As Worksheet) As Integer
' arr:  チェックする列を配列にセット
' Retsu_for_Find_last_row:  値の入っている最終行を取得するためのもの。数式が入っていない列を指定する。数式が入っている列を指定すると数式が入っていない最終行になってしまう
' Check_row_cnt:    何行チェックするか。とりあえず多めにしとく
    Debug.Print "DEBUG  Start Function Check()-------------"
    Dim result As Boolean
    Dim StartL As Integer
    Dim i As Integer
    Dim col As Variant
    Check = 0
    
    sheet.Activate


    '    MsgBox "Columns(Retsu_for_Find_last_row).Address　=     " & Columns(Retsu_for_Find_last_row).Address

    '    StartL = sheet.Range("B:B").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1  ' 罫線は無視
    '    StartL = sheet.Range("A:A").Find(What:="*", LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1 ' この方法だと罫線も含んだ最終行になってしまう
    '    StartL = sheet.Cells(Rows.Count, Retsu_for_Find_last_row).End(xlUp).Row + 1
    '    StartL = sheet.Range(Columns(Retsu_for_Find_last_row).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1 ' なぜか　シート「利用時間(User)」だけ、「オブジェクト変数またはWithブロック変数が設定されていません」のエラー  問題はここ　Columns(Retsu_for_Find_last_row).Address
    StartL = sheet.Range(sheet.Columns(Retsu_for_Find_last_row).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).ROW + 1    ' TEST

    sheet.Cells(StartL, arr(0)).Select
    MsgBox "シート「" & sheet.Name & "」のここから、この行に入っている数式が以降 " & Check_row_cnt & " 行に渡って入っているかチェックを始めます。", vbInformation

    For Each col In arr
        For i = StartL + 1 To StartL + Check_row_cnt
            sheet.Cells(i, col).Select
            'Sleep 20 ' msec
            result = CheckSameFormulaType(Cells(StartL, col), Cells(i, col))
            If result = True Then
                Debug.Print "OK:    セル(" & i & ", " & col & ") 数式有  " & Cells(i, col).Formula
                'Cells(i, col).Interior.Color = RGB(0, 255, 0)  色付けると非常に時間が掛かる
            Else
                Debug.Print "要確認！　セル(" & i & ", " & col & ") 数式が入っていないか、数式が異なる"
                Cells(i, col).Interior.Color = RGB(255, 0, 0)
                Check = Check + 1
            End If
        Next
    Next col
    If Check <> 0 Then
        MsgBox "シート「" & sheet.Name & "」にて、" & vbCrLf & "数式が入っていないか、数式が異なるセルが " & Check & " 箇所、見つかりました！！要確認です", vbCritical
    End If

End Function




'------------------------------------------------------------






Function CheckSameFormulaType(rng1 As Range, rng2 As Range) As Boolean
    CheckSameFormulaType = (rng1.FormulaR1C1 = rng2.FormulaR1C1)
End Function

'Function CheckSameFormulaType(rng1 As Range, rng2 As Range) As Boolean
'' セルに数式が入っているか確認
'    If rng1.HasFormula And rng2.HasFormula Then
'        'Debug.Print "どちらかのセルに数式があり"
'        ' R1C1形式で比較して、一致すれば True、異なれば False
'        CheckSameFormulaType = (rng1.FormulaR1C1 = rng2.FormulaR1C1)
'    Else
'        'Debug.Print "どちらかのセルに数式が無し"
'        CheckSameFormulaType = False
'    End If
'End Function














Sub 計画時間xlsx_Check(BL As Integer)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim result As Boolean
    Dim pattern As String
'    pattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$"    '       別の書式（例: YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM） pattern = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2} - \d{4}-\d{2}-\d{2} \d{2}:\d{2}$"
'    pattern = "^\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}$"    '   時間が一桁の場合もあるのでそれに対応
    pattern = "^\d{4}/\d{1,2}/\d{1,2}[ ]{1,2}\d{1,2}:\d{1,2}:\d{1,2}$"  ' スペースの数も1つ、または2つでもマッチするようにしたいです。

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

    '    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2TEST.xlsm"
    MsgBox "マクロ「計画時間xlsx_Check」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & BNAME_KEIKAKU & vbCrLf & "のチェックです。" & vbCrLf & "確認します", vbInformation, "BL" & BL


    ' wb_KEIKAKUを開く
    Dim wb_KEIKAKU As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, True)    ' フルパスを指定
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    Debug.Print "シート全体にエラーがないか確認 "
    Dim ws As Worksheet
    For Each ws In wb_KEIKAKU.Worksheets
        result = CheckForErrors(ws)
    Next ws

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("bl" & BL).Select    '最前面に表示

    wb_KEIKAKU.Worksheets("bl" & BL).Activate    'これ大事
    LineSta = 2 ' getLineNum("運転種別", 1, wb_KEIKAKU.Worksheets("bl" & BL)) + 1
    LineSto = wb_KEIKAKU.Worksheets("bl" & BL).Cells(Rows.Count, "A").End(xlUp).ROW
    
    CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("bl" & BL).Range("B" & LineSta & ":B" & LineSto - 1))
    CheckAllDuplicatesByRange (wb_KEIKAKU.Worksheets("bl" & BL).Range("C" & LineSta & ":C" & LineSto - 1))
    
    For i = LineSta To LineSto
'       Debug.Print "この行　i = " & i & " が、" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Rows(i).Interior.Color = RGB(0, 205, 0)
        
        
        
        If Not IsDateTimeFormatRegEx(Cells(i, 2)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, 2))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, 3)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, 3))
        End If

        
'        If Not IsValidFormat(Cells(i, 2), pattern) Then
'            Call CMsg("A正しい形式ではありません。" & vbCrLf & "正しい形式: YYYY/MM/DD HH:MM:SS", 3, Cells(i, 2))
'        End If
                    
'        If Not IsValidFormat(Cells(i, 3), pattern) Then
'            Call CMsg("B正しい形式ではありません。" & vbCrLf & "正しい形式: YYYY/MM/DD HH:MM:SS", 3, Cells(i, 3))
'        End If
        
        
        If (Cells(i, 3).Value - Cells(i, 2).Value) <= 0 Then
            Call CMsg("時間がおかしいぞ！　ENDの方が古い" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
        
        
        If (Cells(i, 3).Value - Cells(LineSta, 2).Value) <= 0 Then
            Call CMsg("時間がおかしいぞ！　ユニット開始の時間より古い日時です。" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
        
        If InStr(Cells(i, 4).Value, "プログラム") > 0 Or InStr(Cells(i, 4).Value, "FCBT") > 0 Or InStr(Cells(i, 4).Value, "大学院") > 0 Or InStr(Cells(i, 4).Value, "基盤") > 0 Or InStr(Cells(i, 4).Value, "BL") > 0 Then
            Call CMsg("変だぞ！！！" & vbCrLf & "FCBTの運転種別がユーザー運転になってる事を確認。" & vbCrLf & "基盤開発プログラムや、大学院生プログラムはBLstudyになります！！" & vbCrLf & "BL studeyが紛れ込んでるぞ！！", vbExclamation, Cells(i, 4))
        End If
                
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
                
        If i = LineSto Then
            If (Cells(i, 3).Value - Cells(i, 2).Value) <> 14 Then
                Call CMsg("1ユニット、2週間じゃないんですね" & vbCrLf & "~~~", vbExclamation, Cells(i, 3))
            End If
            
            If Cells(i, 4).Value <> "" Then
                Call CMsg("空欄であるべきところに値が入力はいってます。" & vbCrLf & "~~~", vbCritical, Cells(i, 4))
            End If
            
        End If

    Next

    
    Call CheckScheduleContinuity(wb_KEIKAKU.Worksheets("bl" & BL))


    Call Fin("チェック終了しました。" & vbCrLf & "", 1)
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    
End Sub





Sub 計画時間xlsx_GUN_HV_OFF_Check(BL As Integer)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim Retsu_GUN_HV_OFF As Integer
    Dim Retsu_GUN_HV_ON As Integer
    Dim result As Boolean
'    Dim pattern As String  使わない
'    pattern = "^\d{4}/\d{1,2}/\d{1,2}[ ]{1,2}\d{1,2}:\d{1,2}:\d{1,2}$"  ' スペースの数も1つ、または2つでもマッチするようにしたいです。

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

    MsgBox "マクロ「計画時間xlsx_GUN_HV_OFF_Check」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & BNAME_KEIKAKU & vbCrLf & "のチェックです。" & vbCrLf & "確認します", vbInformation, "BL" & BL


    ' wb_KEIKAKUを開く
    Dim wb_KEIKAKU As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, True)    ' フルパスを指定
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    Debug.Print "シート全体にエラーがないか確認 "
    Dim ws As Worksheet
    For Each ws In wb_KEIKAKU.Worksheets
        result = CheckForErrors(ws)
    Next ws

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("bl" & BL).Select    '最前面に表示


    wb_KEIKAKU.Worksheets("GUN HV OFF").Activate    'これ大事
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
        'Debug.Print "この行　i = " & i & " が、" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        'Application.StatusBar = "Val:    " & i & "   " & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value

        Cells(i, Retsu_GUN_HV_OFF).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_GUN_HV_ON).Interior.Color = RGB(0, 205, 0)
                
        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_GUN_HV_OFF)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, 2))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_GUN_HV_ON)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, 3))
        End If
                
        
        If (Cells(i, Retsu_GUN_HV_ON).Value - Cells(i, Retsu_GUN_HV_OFF).Value) <= 0 Then
            Call CMsg("時間がおかしいぞ！　ENDの方が古い" & vbCrLf & "~~~", vbCritical, Cells(i, 3))
        End If
               
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
    Next



    Call Fin("チェック終了しました。" & vbCrLf & "", 1)
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
End Sub









Sub 運転集計記録_Check(BL As String, sname As String)
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

    MsgBox "マクロ「運転集計記録_Check」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & wb_name & vbCrLf & "のチェックです。" & vbCrLf & "確認します", vbInformation, "BL" & BL

    Retsu_end = 2
    Retsu_start = 3
    Retsu_chouseizikan = 4
    Retsu_total = 5

    Dim wb As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb = OpenBook(wb_name, True)    ' フルパスを指定
    wb.Activate
    If wb Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    If ActiveWorkbook.Name <> wb.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_KEIKAKU, 3)
    End If

    'Debug.Print "シート全体にエラーがないか確認 "
    'Dim ws As Worksheet
    'For Each ws In wb.Worksheets
    '    result = CheckForErrors(ws)
    'Next ws
    
    wb.Windows(1).WindowState = xlMaximized
    wb.Worksheets(sname).Select    '最前面に表示

    wb.Worksheets(sname).Activate    'これ大事
    LineSta = 3
    LineSto = GetLastDataRow(wb.Worksheets(sname), "B") ' wb.Worksheets(sname).Cells(Rows.Count, "B").End(xlUp).ROW
    
    CheckAllDuplicatesByRange (wb.Worksheets(sname).Range("B" & LineSta & ":B" & LineSto))
    CheckAllDuplicatesByRange (wb.Worksheets(sname).Range("C" & LineSta & ":C" & LineSto))

    For i = LineSta To LineSto
        'Debug.Print "この行　i = " & i & " が、" & Cells(i, 2).Value & "    " & Cells(i, 3).Value & "   " & Cells(i, 4).Value
        Cells(i, Retsu_end).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_start).Interior.Color = RGB(0, 205, 0)
        Cells(i, Retsu_total).Interior.Color = RGB(0, 205, 0)
        
        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_end)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, Retsu_end))
        End If

        If Not IsDateTimeFormatRegEx(Cells(i, Retsu_start)) Then
            Call CMsg("日時の形式ではありません。もしかしたら日付オンリーのUNIXTIMEかも。" & vbCrLf & "セルの書式設定を文字列にすると確認できます。", vbCritical, Cells(i, Retsu_start))
        End If
                
        
        If (Cells(i, Retsu_start).Value - Cells(i, Retsu_end).Value) <= 0 Then
            Call CMsg("時間がおかしいぞ！　ENDの方が古い" & vbCrLf & "~~~", vbCritical, Cells(i, Retsu_start))
        End If
        
        If sname = "停止時間" Then
            If Cells(i, Retsu_start).Value > ThisWorkbook.sheetS("手順").Range("E" & UNITROW) Then ' ユニット開始時刻より新しいところだけ確認
                If Cells(i, Retsu_chouseizikan) <> "" Then
                    Call CMsg("列(調整時間)に調整理由が書かれていることはあまりありませんが、、" & vbCrLf & "確認した方がいいです", vbExclamation, Cells(i, Retsu_chouseizikan))
                End If
            End If
        End If
        
        result = CheckSameFormulaType(Cells(LineSta, Retsu_total), Cells(i, Retsu_total))
        If result = False Then
            Debug.Print "要確認！　セル(" & i & ", " & Retsu_total & ") 数式が入っていないか、数式が異なる"
            Call CMsg("数式が入っていないか、数式が異なる！" & vbCrLf & "~~~", vbCritical, Cells(i, Retsu_total))
        End If
                           
        'Debug.Print "Debug<<<   Cells(i, 4) [ " & Cells(i, 4) & " ]"
    Next
        
    Call Fin("チェック終了しました。" & vbCrLf & "", 1)
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
End Sub





'======  日付パターンマッチ　=======================================================
'  セルの書式設定で、文字列にすると、生のUNIXTIMEになる。
'　この関数では、日付と時間のUNIXTIMEはTure、日付のみのUNIXTIMEはFalseとなる
Function IsDateTimeFormatRegEx(ByVal targetString As String) As Boolean
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 9:0:0")      ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/12/31 23:59:59")  ' True
'    Debug.Print IsDateTimeFormatRegEx("2025/7/9 1:52:43")   ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/01/01 09:00:00") ' True
'    Debug.Print IsDateTimeFormatRegEx("2023/2/29 12:30:00")  ' True (うるう年考慮なし、日付の妥当性はこの正規表現では厳密にチェックしない)
'    Debug.Print IsDateTimeFormatRegEx("2023/13/01 00:00:00") ' False (月が13)
'    Debug.Print IsDateTimeFormatRegEx("2023/01/32 00:00:00") ' False (日が32)
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 24:0:0")     ' False (時が24)
'    Debug.Print IsDateTimeFormatRegEx("2023/1/1 9:60:0")     ' False (分が60)
'    Debug.Print IsDateTimeFormatRegEx("2023-1-1 9:0:0")     ' False (区切り文字が異なる)
'    Debug.Print IsDateTimeFormatRegEx("ABCDEFG")            ' False
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp") ' または New RegExp

    
    With regEx
        .pattern = "^\d{4}/(0?[1-9]|1[0-2])/(0?[1-9]|[12]\d|3[01])\s([01]?\d|2[0-3]):([0-5]?\d):([0-5]?\d)$"
        .IgnoreCase = False ' 大文字・小文字を区別しない場合はTrue
        .Global = False     ' 文字列全体で最初のマッチングのみを検索する場合はFalse
                            ' 文字列内のすべてのマッチを検索する場合はTrue
    End With

    IsDateTimeFormatRegEx = regEx.Test(targetString)

    Set regEx = Nothing
End Function





'======  A列に予定種類、B列に開始時間、C列に終了時間が記載されている場合に、B列の開始時間が前の予定のC列の終了時間と一致しない場合に警告を表示　=======================================================
Sub CheckScheduleContinuity(sheet As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim prevEndTime As Date
    
    lastRow = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).ROW - 1 ' 最終行の一行手前までチェック
    
    For i = 2 To lastRow
        Debug.Print "DEBUG: " & Cells(i, 2).Value
        ' 前の予定の終了時間を取得
        If i > 2 Then
            If sheet.Cells(i, 2).Value <> prevEndTime Then
                Cells(i, 2).Font.Color = RGB(255, 5, 5)
                MsgBox "警告: " & sheet.Cells(i, 1).Value & " の開始時間が前の予定の終了時間と一致しません。", vbCritical
            End If
        End If
        
        ' 現在の予定の終了時間を保存
        prevEndTime = sheet.Cells(i, 3).Value
    Next i
End Sub

