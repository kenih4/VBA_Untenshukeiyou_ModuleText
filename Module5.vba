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

    MsgBox "マクロ「Final_Check()」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & BNAME_MATOME & vbCrLf & "のチェックです。" & vbCrLf & "チェックするユニットを確認する為に一旦、SACLA運転状況集計BL*.xlsmを開きます", vbInformation, "BL" & BL

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
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
        DOWNTIME_ROW = 8
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
        DOWNTIME_ROW = 9
    Case Else
        Debug.Print "Zzz..."
        End
    End Select

    UNIT = ThisWorkbook.sheetS("手順").Range("D" & UNITROW)

'    'wb_SHUKEIを開く  [ユニット]を確認するため
'    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
'    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, True)    ' フルパスを指定
'    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
'    wb_SHUKEI.Activate
'    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
'        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
'    End If
'
'    wb_SHUKEI.Windows(1).WindowState = xlMaximized
'    wb_SHUKEI.Worksheets("利用時間（期間）").Activate
'
'    Unit = wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2")
'
'    If MsgBox("チェック対象のユニット(シート「利用時間（期間）」のセルB2)は    " & vbCrLf & "「 " & Unit & " 」" & vbCrLf & "です。 " & vbCrLf & "間違いないですか？" & vbCrLf & "進むにはYESを押して下さい", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
'        Call Fin("「No」が選択されました。終了します。", 1)
'    End If


    ' wb_MATOMEを開く
    Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_MATOME = OpenBook(BNAME_MATOME, True)    ' フルパスを指定
    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    Application.WindowState = xlMaximized

    Debug.Print "シート全体にエラーがないか確認 "
    Dim ws As Worksheet
    For Each ws In wb_MATOME.Worksheets
        Debug.Print ws.Name
        result = CheckForErrors(ws)
    Next ws



    wb_MATOME.Worksheets("Fault集計").Activate    'これ大事
    MsgBox "Fault集計シートをチェックします。" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("SACLA Fault間隔(BL2)", 2, wb_MATOME.Worksheets("Fault集計"))
        LineSto = getLineNum("SACLA Fault間隔(BL3)", 2, wb_MATOME.Worksheets("Fault集計"))
    Else
        LineSta = getLineNum("SACLA Fault間隔(BL3)", 2, wb_MATOME.Worksheets("Fault集計"))
        LineSto = wb_MATOME.Worksheets("Fault集計").Cells(Rows.Count, "B").End(xlUp).ROW
    End If

    For i = LineSta To LineSto
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value
        If wb_MATOME.Worksheets("Fault集計").Cells(i, 2).Value = UNIT Then
            Debug.Print "この行　i = " & i & " が、ユニット " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("Fault集計").Cells(i, 2).MergeArea.Rows.Count - 1

                Check_col_arr = Array(3, 4, 5, 6, 7, 8, 9)  'Check_col_arr = Array(3, 4, 7, 8) ' チェックする列の値を配列にセット  シフト開始、終了、Faul間隔、Faul回数
                For Each col In Check_col_arr
                    Set tc = wb_MATOME.Worksheets("Fault集計").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)
                    'Sleep 100    ' msec
                    If tc.MergeArea.Columns.Count > 1 Or tc.MergeArea.Rows(1).ROW <> j Then
                        Debug.Print "水平方向に結合されてる、または、垂直方向に結合されていて先頭です。" & i & "   " & j & "   " & col
                    Else
                        If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("空欄、または、エラーが発生しています", 3, tc)
                        Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  セルが#DIV/0!だと ここ、表示で失敗するので、その前でIsCellErrorでチェックする

                        If col = 3 Or col = 4 Then    ' シフト時間
                            result = CheckDateTimeFormat(tc)
                        End If

                        If col = 5 And (tc.Value <= 0 Or tc.Value > 8.2 Or Not IsNumeric(tc.Value)) Then  'エネルギー
                            Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                        End If
    
                        If col = 6 And (tc.Value <= 0 Or tc.Value > 25 Or Not IsNumeric(tc.Value)) Then  '波長
                            Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                        End If

                        If col = 7 Then  'Fault間隔時間
                            result = CheckTimeFormat(tc)
                        End If

                        If col = 8 And (tc.Value < 0 Or Not IsNumeric(tc.Value)) Then  'Fault回数
                            Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                        End If

                        If col = 9 And (StrComp(Right(tc.Value, 1), "G", vbBinaryCompare) = 0 = False) Then  ' 末尾の1文字が "G" かどうかチェック（大文字・小文字を区別）
                            Call CMsg("ユーザー名が入る筈なのにGがありませんよ", 3, tc)
                        End If

                    End If
                Next col
            Next
            Exit For
        End If
    Next


    

    wb_MATOME.Worksheets("まとめ ").Activate    'これ大事======================================================================================

    MsgBox "まとめシートの(a)のチェックします。" & vbCrLf & "", vbInformation, "BL" & BL
    For i = getLineNum("(a)運転時間　期間毎", 2, wb_MATOME.Worksheets("まとめ ")) To getLineNum("(b)運転時間　シフト毎", 2, wb_MATOME.Worksheets("まとめ "))
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("まとめ ").Cells(i, 2).Value = UNIT Then
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            If BL = 2 Then
                DOWNTIME_ROW = i
            Else    'BL3
                DOWNTIME_ROW = i + 1
            End If

            Check_col_arr = Array(3, 5, 6, 7, 9, 10, 11, 12)    ' チェックする列の値を配列にセット
            For Each col In Check_col_arr
                If col >= 9 Then
                    Set tc = wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, col)
                Else
                    Set tc = wb_MATOME.Worksheets("まとめ ").Cells(i, col)
                End If
                tc.Select
                tc.Interior.Color = RGB(0, 255, 0)
                'Sleep 100    ' msec
                If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("空欄、または、エラーが発生しています", 3, tc)
                Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  セルが#DIV/0!だと ここ、表示で失敗するので、その前でIsCellErrorでチェックする

                If col = 3 Then    ' 日付
                    pattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2} - \d{4}/\d{2}/\d{2} \d{2}:\d{2}$"    '       別の書式（例: YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM） pattern = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2} - \d{4}-\d{2}-\d{2} \d{2}:\d{2}$"
                    If Not IsValidFormat(tc, pattern) Then
                        Call CMsg("セル " & tc.Address(False, False) & " の値が正しい形式ではありません。" & vbCrLf & "正しい形式: YYYY/MM/DD HH:MM - YYYY/MM/DD HH:MM", 3, tc)
                    End If
                End If

                If col = 5 Or col = 6 Or col = 7 Or col = 9 Or col = 10 Or col = 11 Or col = 12 Then    '総運転時間(計画）(計画, ダウンタイム), 利用調整運転(計画, ダウンタイム) , 利用運転(計画, ダウンタイム)
                    result = CheckTimeFormat(tc)
                End If

            Next col


            If wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 9).Value <= 0 Then
                Call CMsg("利用調整運転(BL調整orBL-study)はなかったんですね。", 2, wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 9))
            End If

            If wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 11).Value <= 0 Then
                Call CMsg("利用運転(ユーザー)はなかったんですね。" & vbCrLf & "「ユーザー運転無し」と手動で処理しないといけない部分があります。", 2, wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 11))
            Else
                If wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 12).Value <= 0 Then
                    Call CMsg("一回もトリップしてないって事？確認した方がよいです。" & vbCrLf & "シート「集計記録」に数式が入っていない可能性があります", 2, wb_MATOME.Worksheets("まとめ ").Cells(DOWNTIME_ROW, 12))
                End If
            End If
        End If
    Next






    MsgBox "まとめシートの(b)のチェック。" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("(b-1)BL2", 2, wb_MATOME.Worksheets("まとめ "))
        LineSto = getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets("まとめ "))
    Else
        LineSta = getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets("まとめ "))
        LineSto = wb_MATOME.Worksheets("まとめ ").Cells(Rows.Count, "B").End(xlUp).ROW
    End If

    Check_col_arr = Array(3, 4, 5, 6, 7, 8)    ' チェックする列の値を配列にセット  シフト時間(開始・終了・間隔)、利用率、ビーム調整時間、ダウンタイム

    For i = LineSta To LineSto
        Debug.Print "i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("まとめ ").Cells(i, 2).Value = UNIT Then
            Debug.Print "この行　i = " & i & " が、ユニット " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("まとめ ").Cells(i, 2).MergeArea.Rows.Count - 1
                For Each col In Check_col_arr
                    Set tc = wb_MATOME.Worksheets("まとめ ").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)
                    'Sleep 100    ' msec
                    If tc.MergeArea.Columns.Count > 1 Then
                        Debug.Print "水平方向に結合されています。" & i & "   " & j & "   " & col & "     tc.Value = " & tc.Value & "  " & tc.MergeArea.Columns.Count
                    Else

                        If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("空欄、または、エラーが発生しています", 3, tc)
                        Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  セルが#DIV/0!だと ここ、表示で失敗するので、その前でIsCellErrorでチェックする

                        If col = 5 And (wb_MATOME.Worksheets("まとめ ").Cells(j, 3).Value = "total" And (StrComp(Right(wb_MATOME.Worksheets("まとめ ").Cells(j, 9).Value, 1), "G", vbBinaryCompare) = 0) = False) Then  '' 末尾の1文字が "G" かどうかチェック（大文字・小文字を区別）
                            Call CMsg("ユーザー名が入るべきですが。。確認した方がいいです。", 2, wb_MATOME.Worksheets("まとめ ").Cells(j, 9))
                        End If

                        If col = 3 Or col = 4 Then
                            result = CheckDateTimeFormat(tc)
                        End If

                        If col = 5 Or col = 7 Or col = 8 Then
                            result = CheckTimeFormat(tc)
                        End If
    
                        If (col = 5 And wb_MATOME.Worksheets("まとめ ").Cells(j, 3).Value <> "total") And (tc.Value <= 0 Or tc.Value > 0.5 Or Not IsNumeric(tc.Value)) Then
                            Call CMsg("範囲外かもしれないです。確認した方がいいです。", 3, tc)
                        End If

                        If col = 6 And (tc.Value <= 0.8 Or tc.Value > 1 Or Not IsNumeric(tc.Value)) Then  '利用率%
                            Call CMsg("利用率低い。または、範囲外 or 文字列   確認した方がいいです。", 3, tc)
                        End If

                    End If
                Next col

            Next
            Exit For
        End If
    Next





    MsgBox "まとめシートの(c)のチェック。" & vbCrLf & "", vbInformation, "BL" & BL
    If BL = 2 Then
        LineSta = getLineNum("(c-1)BL2", 2, wb_MATOME.Worksheets("まとめ "))
        LineSto = getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets("まとめ "))
    Else
        LineSta = getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets("まとめ "))
        LineSto = wb_MATOME.Worksheets("まとめ ").Cells(Rows.Count, "B").End(xlUp).ROW
    End If


    For i = LineSta To LineSto
        Debug.Print "DEBUG D    i = " & i & "  " & Cells(i, 2).Value

        If wb_MATOME.Worksheets("まとめ ").Cells(i, 2).Value = UNIT Then
            Debug.Print "この行　i = " & i & " が、ユニット " & Cells(i, 2).Value
            CantFindUnit = CantFindUnit + 1
            Cells(i, 2).Select
            Cells(i, 2).Interior.Color = RGB(0, 255, 0)
            For j = i To i + wb_MATOME.Worksheets("まとめ ").Cells(i, 2).MergeArea.Rows.Count - 1

                For col = 3 To 7
                    Set tc = wb_MATOME.Worksheets("まとめ ").Cells(j, col)
                    tc.Select
                    tc.Interior.Color = RGB(0, 255, 0)

                    'Sleep 100    ' msec
                    If IsCellErrorType(tc) = False Or IsEmpty(tc.Value) Then Call CMsg("空欄、または、エラーが発生しています", 3, tc)
                    Debug.Print i & "   " & j & "   " & col & "     tc.Value = " & tc.Value    '!!!!!!!!!  セルが#DIV/0!だと ここ、表示で失敗するので、その前でIsCellErrorでチェックする

                    If col = 3 And (tc.Value <= 0 Or tc.Value > 8.2 Or Not IsNumeric(tc.Value)) Then  'エネルギー
                        Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                    End If

                    If col = 4 And (tc.Value <= 0 Or tc.Value > 60 Or Not IsNumeric(tc.Value)) Then  '繰返し
                        Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                    End If
    
                    If col = 5 And (tc.Value <= 0 Or tc.Value > 25 Or Not IsNumeric(tc.Value)) Then  '波長
                        Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
    
                        If InStr(1, tc.Value, "+", vbTextCompare) > 0 Then '波長
                            Call CMsg("セルには「+」が含まれています。", 2, tc)
                            If MsgBox("備考セルに、「、二色実験」と追い書き込みますか？" & vbCrLf & "いいです？", vbYesNo + vbQuestion, "確認") = vbYes Then
                    '            MsgBox j & "  Cells(j, 7).Value:     " & wb_MATOME.Worksheets("まとめ ").Cells(j, 7).Value, Buttons:=vbInformation
                                wb_MATOME.Worksheets("まとめ ").Cells(j, 7).Value = wb_MATOME.Worksheets("まとめ ").Cells(j, 7).Value + "、二色実験"
                    '            MsgBox "追い書き込みした。     " & wb_MATOME.Worksheets("まとめ ").Cells(j, 7).Value & vbCrLf & "次に進みます。", Buttons:=vbInformation
                            End If
                        End If

                    End If

                    If col = 6 And (tc.Value <= 0 Or tc.Value > 2000 Or Not IsNumeric(tc.Value)) Then  '強度
                        Call CMsg("範囲外 or 非数値です。確認した方がいいです。", 3, tc)
                    End If

                    If col = 7 And (IsNumeric(tc.Value)) Then  '備考
                        Call CMsg("数値です。確認した方がいいです。", 3, tc)
                    End If

                Next

            Next
            Exit For
        End If
    Next




    If CantFindUnit <> 4 Then
        MsgBox "異常です。" & vbCrLf & "チェック対象のユニット  CantFindUnit : " & CantFindUnit & " しかありませんでした。４つあるべきです。", Buttons:=vbCritical
    End If





    Call Fin("終了しました。" & vbCrLf & "", 1)
    Exit Sub    ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical

End Sub

























Function IsCellErrorType(Target As Variant) As Boolean
    If IsError(Target.Value) Then
        Select Case Target.Value
        Case CVErr(xlErrDiv0)
            IsCellErrorType = False    'IsCellErrorType = "#DIV/0! エラー"
        Case CVErr(xlErrNA)
            IsCellErrorType = False    'IsCellErrorType = "#N/A エラー"
        Case CVErr(xlErrValue)
            IsCellErrorType = False    'IsCellErrorType = "#VALUE! エラー"
        Case Else
            IsCellErrorType = False    'IsCellErrorType = "その他のエラー"
        End Select
        Call CMsg("このセルでエラーが発生しています。@IsCellErrorType", 3, Target)
    Else
        IsCellErrorType = True    'IsCellErrorType = "エラーなし"
    End If
End Function







Function CheckDateTimeFormat(Target As Variant) As Boolean
    Dim compareDate As Date
    CheckDateTimeFormat = False
    If IsDate(Target.Value) Then
        If Format(Target.Value, "yyyy/mm/dd hh:mm") <> Target.Text Then
            Call CMsg("フォーマットが正しくありません。@CheckDateTimeFormat" & vbCrLf & "正しい形式: 2025/01/28 22:00", 3, Target)
        Else
            CheckDateTimeFormat = True
            compareDate = DateSerial(2025, 1, 1) + TimeSerial(12, 30, 0)
            If Target.Value < compareDate Then
                MsgBox Target.Value & " が、 " & compareDate & " より前です。確認した方がいいです。", vbExclamation
            End If
        End If
    Else
        Call CMsg("有効な日付が入力されていません。@CheckDateTimeFormat", 3, Target)
    End If
End Function



Function CheckTimeFormat(Target As Variant) As Boolean
    Debug.Print "CheckTimeFormat         target.Value = " & Target.Value
    CheckTimeFormat = False
    If Not IsNumeric(Target.Value) Or Target.Value < 0 Then
        Debug.Print "有効な時間が入力されていません。@CheckTimeFormat    target.Value = " & Target.Value
        Call CMsg("有効な時間が入力されていません。@CheckTimeFormat", 3, Target)
    Else
        If IsDate(CDate(Target.Value)) Then
            Dim fmt As String
            fmt = Target.NumberFormat
            'Debug.Print "フォーマットは　     target.Value = " & target.Value & "  fmt = " & fmt
            If fmt = "h:mm" Or fmt = "hh:mm" Or fmt = "[h]:mm" Or fmt = "h:mm;@" Or fmt = "hh:mm;@" Then    ' [h]:mmは累積時間
                'Debug.Print "時刻データで正しいフォーマットです。    target.Value = " & target.Value
                CheckTimeFormat = True
            Else
                Debug.Print "時刻データですが、フォーマットが異なります。    target.Value = " & Target.Value
                Call CMsg("時刻データですが、フォーマットが異なります。@CheckTimeFormat", 3, Target)
            End If
        End If
    End If
End Function












'--------------------------------------------------------------------------------------------------------------------------------------------
' セルの値が指定したパターンに一致するかチェックする関数
Function IsValidFormat(cell As Variant, pattern As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    regEx.pattern = pattern
    regEx.IgnoreCase = True
    regEx.Global = False

    ' 正規表現がマッチするかを判定
    IsValidFormat = regEx.Test(cell.Value)

    ' オブジェクト解放
    Set regEx = Nothing
End Function









