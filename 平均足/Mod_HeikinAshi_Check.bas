Attribute VB_Name = "Mod_HeikinAshi_Check"
Option Explicit
'==================================================================
' 平均足　データチェック　v1.0
'
'   実行 : Alt+F8 →「平均足_データチェック」
'   「データチェック」シートに結果を出します。
'
'   ◆一番の目的
'     楽天RSSが切れていて、データが古いまま候補を出してしまう事故を防ぐ。
'     最新日が2営業日以上古い場合は、はっきり警告します。
'
'   ◆調べること
'     1. 最新日が何営業日前か（RSS切断・取込忘れの検出）
'     2. 日付の逆転・重複・土日混入・欠損
'     3. 5枚のシートで日付と銘柄の並びが一致しているか
'     4. 5行目（TOPX）が他の銘柄と同じ値になっていないか
'     5. 銘柄ごとの空白セルの数
'==================================================================

Private Const SH As String = "データチェック"
Private Const R_TOP As Long = 6
Private Const R_END As Long = 520
Private Const C_NEW As Long = 5
Private Const N_MAX As Long = 300

'============ 全マクロが使う「最新営業日」============
Public Function HA_LatestDate() As Date
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("終値")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim c As Long, dv As Double, best As Double, miss As Long
    '日付行を左から走査。E列の日付は D3 に入っている
    dv = 0
    If IsNumeric(ws.Cells(3, 4).Value2) Then dv = CDbl(ws.Cells(3, 4).Value2)
    If dv > 40000 And dv < 80000 Then best = dv
    miss = 0
    For c = C_NEW + 1 To C_NEW + N_MAX
        dv = 0
        If IsNumeric(ws.Cells(3, c).Value2) Then dv = CDbl(ws.Cells(3, c).Value2)
        If dv > 40000 And dv < 80000 Then
            If dv > best Then best = dv
            miss = 0
        Else
            miss = miss + 1
            If miss > 5 Then Exit For
        End If
    Next c
    If best > 0 Then HA_LatestDate = CDate(best)
End Function

'============ データが新しいか ============
Public Function HA_DataFresh(ByRef msg As String) As Boolean
    Dim d As Date, k As Long, cur As Date
    d = HA_LatestDate()
    If d = 0 Then
        msg = "最新日が読み取れません。": HA_DataFresh = False: Exit Function
    End If
    '今日から数えて何営業日前か（祝日は考えないので目安）
    cur = Date
    k = 0
    Do While cur > d And k < 30
        cur = cur - 1
        If Weekday(cur, vbMonday) <= 5 Then k = k + 1
    Loop
    msg = "最新日 " & Format(d, "yyyy/mm/dd(aaa)") & "　＝　およそ " & k & " 営業日前"
    HA_DataFresh = (k <= 1)
End Function

'==================== 本体 ====================
Public Sub 平均足_データチェック()

    Dim wsC As Worksheet
    Set wsC = HA_C_Ws("終値")
    If wsC Is Nothing Then
        MsgBox "「終値」シートがありません。", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH
    End If

    Application.ScreenUpdating = False
    ws.Cells.UnMerge
    ws.Cells.Clear

    Dim r As Long
    r = 1
    HA_C_Title ws, r, "  平均足　データチェック　（" & Format(Now, "yyyy/mm/dd hh:nn") & "）"

    '--- 1. 鮮度 ---
    Dim msg As String, fresh As Boolean
    fresh = HA_DataFresh(msg)
    HA_C_Head ws, r, "1. データの新しさ"
    If fresh Then
        HA_C_Line ws, r, "OK", msg, RGB(0, 97, 0)
    Else
        HA_C_Line ws, r, "★警告★", msg & "　→　楽天RSSが切れているか、取込を忘れています。" & _
                  "この状態で候補を出さないでください。", RGB(192, 0, 0)
    End If

    '--- 日付を集める ---
    Dim dts() As Double, cols() As Long, nD As Long
    ReDim dts(1 To N_MAX): ReDim cols(1 To N_MAX)
    Dim c As Long, dv As Double, miss As Long
    nD = 0: miss = 0
    dv = 0
    If IsNumeric(wsC.Cells(3, 4).Value2) Then dv = CDbl(wsC.Cells(3, 4).Value2)
    If dv > 40000 And dv < 80000 Then
        nD = 1: dts(1) = dv: cols(1) = C_NEW
    End If
    For c = C_NEW + 1 To C_NEW + N_MAX - 1
        dv = 0
        If IsNumeric(wsC.Cells(3, c).Value2) Then dv = CDbl(wsC.Cells(3, c).Value2)
        If dv > 40000 And dv < 80000 Then
            nD = nD + 1: dts(nD) = dv: cols(nD) = c: miss = 0
        Else
            miss = miss + 1
            If miss > 5 Then Exit For
        End If
    Next c

    HA_C_Head ws, r, "2. 日付（" & nD & "日分）"
    If nD < 60 Then
        HA_C_Line ws, r, "★警告★", "日数が " & nD & " しかありません。週足の判定に60日以上必要です。", RGB(192, 0, 0)
    Else
        HA_C_Line ws, r, "OK", Format(CDate(dts(nD)), "yyyy/mm/dd") & " 〜 " & _
                  Format(CDate(dts(1)), "yyyy/mm/dd") & "　" & nD & "日分", RGB(0, 97, 0)
    End If

    '--- 2. 逆転・重複・土日・欠損 ---
    Dim k As Long, bad As Long, dup As Long, wknd As Long, gap As Long
    For k = 1 To nD - 1
        If dts(k) <= dts(k + 1) Then bad = bad + 1
        If dts(k) = dts(k + 1) Then dup = dup + 1
    Next k
    For k = 1 To nD
        If Weekday(CDate(dts(k)), vbMonday) >= 6 Then wknd = wknd + 1
    Next k
    For k = 1 To nD - 1
        If dts(k) - dts(k + 1) >= 5 Then
            gap = gap + 1
            HA_C_Line ws, r, "確認", Format(CDate(dts(k + 1)), "yyyy/mm/dd(aaa)") & " → " & _
                      Format(CDate(dts(k)), "yyyy/mm/dd(aaa)") & "　" & CLng(dts(k) - dts(k + 1)) & "日あいています" & _
                      "（年末年始・GW・連休なら正常）", RGB(128, 96, 0)
        End If
    Next k
    HA_C_Line ws, r, IIf(bad = 0, "OK", "★警告★"), "日付の逆転・並び順の乱れ：" & bad & "件", IIf(bad = 0, RGB(0, 97, 0), RGB(192, 0, 0))
    HA_C_Line ws, r, IIf(dup = 0, "OK", "★警告★"), "日付の重複：" & dup & "件", IIf(dup = 0, RGB(0, 97, 0), RGB(192, 0, 0))
    HA_C_Line ws, r, IIf(wknd = 0, "OK", "★警告★"), "土日が混ざっている日：" & wknd & "件", IIf(wknd = 0, RGB(0, 97, 0), RGB(192, 0, 0))

    '--- 3. 5枚のシートの一致 ---
    HA_C_Head ws, r, "3. 5枚のシートの一致"
    Dim nmArr As Variant, i As Long
    nmArr = Array("始値", "高値", "安値", "終値", "出来高")
    Dim lastRow As Long
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    If lastRow > R_END Then lastRow = R_END
    For i = 0 To UBound(nmArr)
        Dim w2 As Worksheet
        Set w2 = HA_C_Ws(CStr(nmArr(i)))
        If w2 Is Nothing Then
            HA_C_Line ws, r, "★警告★", CStr(nmArr(i)) & "シートがありません", RGB(192, 0, 0)
        Else
            Dim dNG As Long, cNG As Long
            dNG = 0: cNG = 0
            For k = 1 To nD
                If HA_C_Num(w2.Cells(3, IIf(cols(k) = C_NEW, 4, cols(k))).Value2) <> dts(k) Then dNG = dNG + 1
            Next k
            Dim rr As Long
            For rr = 5 To lastRow
                If Trim$(CStr(w2.Cells(rr, 1).Value)) <> Trim$(CStr(wsC.Cells(rr, 1).Value)) Then cNG = cNG + 1
            Next rr
            HA_C_Line ws, r, IIf(dNG = 0 And cNG = 0, "OK", "★警告★"), _
                      CStr(nmArr(i)) & "：日付のずれ " & dNG & "件／銘柄の並びのずれ " & cNG & "件", _
                      IIf(dNG = 0 And cNG = 0, RGB(0, 97, 0), RGB(192, 0, 0))
        End If
    Next i

    '--- 4. TOPX行が他の銘柄と同じでないか ---
    HA_C_Head ws, r, "4. 5行目（TOPX）の中身"
    Dim same As Long, tot As Long, rr2 As Long, hitRow As Long, hitCnt As Long
    For rr2 = R_TOP To lastRow
        same = 0: tot = 0
        For k = 1 To nD
            Dim a As Double, b As Double
            a = HA_C_Num(wsC.Cells(5, cols(k)).Value2)
            b = HA_C_Num(wsC.Cells(rr2, cols(k)).Value2)
            If a > 0 And b > 0 Then
                tot = tot + 1
                If a = b Then same = same + 1
            End If
        Next k
        If tot > 20 Then
            If same > hitCnt Then hitCnt = same: hitRow = rr2
        End If
    Next rr2
    If hitCnt >= nD * 0.5 Then
        HA_C_Line ws, r, "★警告★", "TOPX行（5行目）の " & hitCnt & "/" & nD & "日分が、" & _
                  Trim$(CStr(wsC.Cells(hitRow, 1).Value)) & " " & Trim$(CStr(wsC.Cells(hitRow, 2).Value)) & _
                  "（" & hitRow & "行目）と同じ値です。RSSの銘柄指定を確認してください。", RGB(192, 0, 0)
        HA_C_Line ws, r, "参考", "※平均足の売買判定はTOPXを一切使っていないので、成績への影響はありません。", RGB(90, 90, 90)
    Else
        HA_C_Line ws, r, "OK", "他の銘柄と重複していません。", RGB(0, 97, 0)
    End If

    '--- 5. 空白セル ---
    HA_C_Head ws, r, "5. 空白（欠損）セル"
    Dim blanks As Long, worst As Long, worstRow As Long
    For rr2 = R_TOP To lastRow
        If Trim$(CStr(wsC.Cells(rr2, 1).Value)) <> "" Then
            Dim bl As Long
            bl = 0
            For k = 1 To nD
                If HA_C_Num(wsC.Cells(rr2, cols(k)).Value2) <= 0 Then bl = bl + 1
            Next k
            blanks = blanks + bl
            If bl > worst Then worst = bl: worstRow = rr2
        End If
    Next rr2
    HA_C_Line ws, r, IIf(worst < nD * 0.2, "OK", "確認"), _
              "空白の合計 " & blanks & "個／一番多い銘柄は " & _
              Trim$(CStr(wsC.Cells(worstRow, 1).Value)) & " " & Trim$(CStr(wsC.Cells(worstRow, 2).Value)) & _
              " の " & worst & "個（" & nD & "日中）", _
              IIf(worst < nD * 0.2, RGB(0, 97, 0), RGB(128, 96, 0))

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 110
    Application.ScreenUpdating = True

    On Error Resume Next
    ws.Activate
    On Error GoTo 0

    If fresh Then
        MsgBox "データチェックが終わりました。" & vbCrLf & msg, vbInformation
    Else
        MsgBox "★データが古いです★" & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
               "楽天RSSが切れているか、取込を忘れています。" & vbCrLf & _
               "この状態で候補を出さないでください。", vbExclamation
    End If
End Sub

'==================== 部品 ====================
Private Function HA_C_Ws(ByVal n As String) As Worksheet
    On Error Resume Next
    Set HA_C_Ws = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
End Function

Private Function HA_C_Num(ByVal v As Variant) As Double
    If IsNumeric(v) Then HA_C_Num = CDbl(v) Else HA_C_Num = 0
End Function

Private Sub HA_C_Title(ByVal ws As Worksheet, ByRef r As Long, ByVal s As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Merge
        .Value = s
        .Font.Name = "Meiryo UI": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
    End With
    ws.Rows(r).RowHeight = 28
    r = r + 2
End Sub

Private Sub HA_C_Head(ByVal ws As Worksheet, ByRef r As Long, ByVal s As String)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Merge
        .Value = s
        .Font.Name = "Meiryo UI": .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
    End With
    r = r + 1
End Sub

Private Sub HA_C_Line(ByVal ws As Worksheet, ByRef r As Long, ByVal mark As String, _
                      ByVal s As String, ByVal col As Long)
    ws.Cells(r, 2).Value = mark
    ws.Cells(r, 3).Value = s
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Font.Name = "Meiryo UI": .Font.Size = 11
        .Font.Color = col
    End With
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
End Sub
