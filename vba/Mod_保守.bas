Attribute VB_Name = "Mod_保守"
Option Explicit
'=========================================================
'  T850 保守モジュール  v1.0 (完成版)
'
'   初期設定        … 足りないシートを作り、見出しと設定表を整える
'                     (すでにあるシート・データには手を触れません)
'   銘柄反映        … 銘柄シートの一覧を 取得シートと集積19シートに反映
'                     過去の値は「銘柄コードごと」に正しい行へ引っ越します
'   集積シート整備  … 19シートの見出し・書式・列数を整える
'
'   ※ 銘柄反映の前に MarketSpeed2 を立ち上げ、RSSを使える状態に
'      しておいてください(Mod_RSS起動 の「一発起動」)。
'=========================================================

Private Const 最大銘柄 As Long = 300      '4行目～303行目
Private Const 開始行   As Long = 4
Private Const 履歴列数 As Long = 250      'C列から右に残す列数

'=========================================================
'  1. 初期設定
'=========================================================
Public Sub 初期設定()
    Dim need As Variant, i As Long, n As Long, added As String
    Dim ws As Worksheet

    Application.ScreenUpdating = False
    設定表を整える

    need = 必要シート一覧()
    For i = 0 To UBound(need)
        If Not 有(CStr(need(i))) Then
            Set ws = ThisWorkbook.Worksheets.Add( _
                     After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(need(i))
            n = n + 1
            added = added & CStr(need(i)) & " "
        End If
    Next i

    見出しを整える
    Application.ScreenUpdating = True

    記 "情報", "初期設定を実行しました (シート " & ThisWorkbook.Worksheets.Count & "枚 / 新規 " & n & "枚)"
    MsgBox "初期設定が終わりました。" & vbLf & vbLf & _
           "新しく作ったシート: " & IIf(n = 0, "なし(すべて揃っていました)", added), _
           vbInformation, "初期設定"
End Sub

'設定シートの空いている項目だけ埋める(入っている値は上書きしません)
Private Sub 設定表を整える()
    Dim ws As Worksheet, i As Long
    Dim tm As Variant
    If Not 有("設定") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("設定")

    埋 ws, "A1", "取得時刻(表示)"
    埋 ws, "B1", "実取得時刻"
    埋 ws, "C1", "CSV保存フォルダ"
    埋 ws, "F1", "※B列=実際に取りに行く時刻。空ならA列と同じ"
    埋 ws, "F2", "※11:30は11:31 / 15:30は15:31(引け確定後に取る)"

    tm = 既定の時刻表()
    For i = 0 To UBound(tm)
        If Len(CStr(ws.Cells(i + 2, 1).Value)) = 0 Then
            ws.Cells(i + 2, 1).Value = CDate(tm(i))
            ws.Cells(i + 2, 1).NumberFormat = "hh:mm"
        End If
    Next i

    埋 ws, "C2", "ブザー使用":        埋 ws, "D2", "○"
    埋 ws, "C3", "ブザー継続分":      埋 ws, "D3", 10
    埋 ws, "C4", "ブザー間隔秒":      埋 ws, "D4", 30
    埋 ws, "C5", "取得遅延秒":        埋 ws, "D5", 5
    埋 ws, "C6", "自動保存":          埋 ws, "D6", "○"
    埋 ws, "C7", "起動時に自動監視":  埋 ws, "D7", "○"
    埋 ws, "C8", "自動保存時刻1"
    埋 ws, "C9", "自動保存時刻2"
    埋 ws, "C10", "履歴を残す列数(空=消さない)"

    埋 ws, "AA1", "予定時刻(自動)"
    埋 ws, "AB1", "内容"
    埋 ws, "AC1", "状態"
End Sub

Private Function 既定の時刻表() As Variant
    既定の時刻表 = Array("9:00", "9:10", "9:20", "9:45", "10:00", "10:15", "10:30", _
                         "11:30", "12:30", "13:30", "14:30", "15:00", "15:20", "15:30")
End Function

Private Function 必要シート一覧() As Variant
    Dim base As Variant, tm As Variant, r() As String, i As Long, n As Long
    base = Array("設定", "銘柄", "取得", "記録", "ログ", "始値", "高値", "安値", "終値", "出来高")
    tm = 時間シート名()
    ReDim r(0 To UBound(base) + UBound(tm) + 1)
    For i = 0 To UBound(base)
        r(n) = CStr(base(i)): n = n + 1
    Next i
    For i = 0 To UBound(tm)
        r(n) = CStr(tm(i)): n = n + 1
    Next i
    必要シート一覧 = r
End Function

'時間シートの名前は 設定!A2:A15 から作る(空なら既定値)
Private Function 時間シート名() As Variant
    Dim ws As Worksheet, i As Long, n As Long, t As Double
    Dim r() As String, tm As Variant
    ReDim r(0 To 13)
    If 有("設定") Then
        Set ws = ThisWorkbook.Worksheets("設定")
        For i = 2 To 15
            t = 時刻を分数に(ws.Cells(i, 1).Value)
            If t > 0 Then
                r(n) = Format$(CDate(t), "hhnn")
                n = n + 1
            End If
        Next i
    End If
    If n = 0 Then
        tm = 既定の時刻表()
        For i = 0 To UBound(tm)
            r(n) = Format$(CDate(tm(i)), "hhnn")
            n = n + 1
        Next i
    End If
    ReDim Preserve r(0 To n - 1)
    時間シート名 = r
End Function

'見出しが空のところだけ入れる
Private Sub 見出しを整える()
    Dim ws As Worksheet, v As Variant, i As Long

    If 有("取得") Then
        Set ws = ThisWorkbook.Worksheets("取得")
        v = Array("コード", "銘柄名", "始値", "高値", "安値", "現在値", "出来高")
        For i = 0 To UBound(v)
            If Len(CStr(ws.Cells(1, i + 1).Value)) = 0 Then ws.Cells(1, i + 1).Value = v(i)
        Next i
    End If

    If 有("記録") Then
        Set ws = ThisWorkbook.Worksheets("記録")
        v = Array("日付", "時刻区分", "取得時刻", "コード", "銘柄名", _
                  "始値", "高値", "安値", "現在値", "出来高")
        For i = 0 To UBound(v)
            If Len(CStr(ws.Cells(1, i + 1).Value)) = 0 Then ws.Cells(1, i + 1).Value = v(i)
        Next i
    End If

    If 有("ログ") Then
        Set ws = ThisWorkbook.Worksheets("ログ")
        埋 ws, "A1", "日時": 埋 ws, "B1", "種別": 埋 ws, "C1", "内容"
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("C").ColumnWidth = 60
    End If

    If 有("銘柄") Then
        Set ws = ThisWorkbook.Worksheets("銘柄")
        埋 ws, "A1", "コード": 埋 ws, "B1", "銘柄名(自動)": 埋 ws, "C1", "止める時は×"
        埋 ws, "E1", "◆ A列にコードを打つ → 「銘柄反映」ボタン"
        埋 ws, "E2", "◆ 消す時は その行を丸ごと削除 → 「銘柄反映」ボタン"
        埋 ws, "E3", "◆ 一時的に外すだけなら C列に × を入れる"
    End If

    集積見出し
End Sub

'=========================================================
'  2. 銘柄反映
'=========================================================
Public Sub 銘柄反映()
    Dim codes() As String, names() As String
    Dim n As Long, i As Long, add As Long, del As Long
    Dim shts As Variant, calcOld As Long

    n = 銘柄一覧を読む(codes, names)
    If n = 0 Then
        MsgBox "銘柄シートのA列にコードがありません", vbExclamation, "銘柄反映"
        Exit Sub
    End If

    増減を数える codes, n, add, del
    If MsgBox(n & " 銘柄を 取得シートと集積19シートに反映します。" & vbLf & vbLf & _
              "追加 " & add & " 銘柄 / 一覧から外れる " & del & " 銘柄" & vbLf & _
              "過去の値は銘柄コードごとに正しい行へ引っ越します。" & vbLf & _
              "外れた銘柄の過去データは消えます。よろしいですか?", _
              vbYesNo + vbQuestion, "銘柄反映") <> vbYes Then Exit Sub

    calcOld = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo EH

    銘柄シート式(n)
    取得シート反映 codes, n
    shts = 集積シート一覧()
    For i = 0 To UBound(shts)
        If 有(CStr(shts(i))) Then 集積1シート反映 ThisWorkbook.Worksheets(CStr(shts(i))), codes, names, n
    Next i

    Application.Calculation = calcOld
    Application.ScreenUpdating = True
    Application.CalculateFull

    記 "情報", "銘柄反映 " & n & "銘柄 (追加 " & add & " / 削除 " & del & ")"
    MsgBox n & " 銘柄を反映しました。" & vbLf & "追加 " & add & " / 削除 " & del, vbInformation, "銘柄反映"
    Exit Sub
EH:
    Application.Calculation = calcOld
    Application.ScreenUpdating = True
    記 "エラー", "銘柄反映で失敗: " & Err.Description
    MsgBox "銘柄反映で失敗しました: " & Err.Description, vbExclamation
End Sub

'銘柄シートから コードと名前を読む(C列に × があれば飛ばす)
Private Function 銘柄一覧を読む(ByRef codes() As String, ByRef names() As String) As Long
    Dim ws As Worksheet, i As Long, n As Long, c As String
    ReDim codes(1 To 最大銘柄)
    ReDim names(1 To 最大銘柄)
    If Not 有("銘柄") Then Exit Function
    Set ws = ThisWorkbook.Worksheets("銘柄")
    For i = 2 To 最大銘柄 + 1
        c = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(c) > 0 Then
            If InStr(CStr(ws.Cells(i, 3).Value), "×") = 0 Then
                n = n + 1
                codes(n) = c
                names(n) = Trim$(CStr(ws.Cells(i, 2).Value))
                If names(n) = "0" Then names(n) = ""
            End If
        End If
    Next i
    銘柄一覧を読む = n
End Function

Private Sub 増減を数える(ByRef codes() As String, ByVal n As Long, _
                         ByRef add As Long, ByRef del As Long)
    Dim ws As Worksheet, i As Long, d As Object, c As String
    add = 0: del = 0
    If Not 有("取得") Then add = n: Exit Sub
    Set ws = ThisWorkbook.Worksheets("取得")
    Set d = CreateObject("Scripting.Dictionary")
    For i = 2 To 最大銘柄 + 1
        c = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(c) > 0 Then d(c) = 1
    Next i
    For i = 1 To n
        If d.Exists(codes(i)) Then
            d.Remove codes(i)
        Else
            add = add + 1
        End If
    Next i
    del = d.Count
End Sub

'銘柄シートのB列(銘柄名)にRSS式を入れ直す
Private Sub 銘柄シート式(ByVal n As Long)
    Dim ws As Worksheet, i As Long, f() As Variant
    If Not 有("銘柄") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("銘柄")
    ReDim f(1 To 最大銘柄, 1 To 1)
    For i = 1 To 最大銘柄
        If Len(Trim$(CStr(ws.Cells(i + 1, 1).Value))) > 0 Then
            f(i, 1) = "=IFERROR(RssMarket($A" & (i + 1) & ",""銘柄名称""),"""")"
        Else
            f(i, 1) = ""
        End If
    Next i
    ws.Range("B2:B" & (最大銘柄 + 1)).Formula = f
End Sub

'取得シートに コードとRSS式を並べ直す
Private Sub 取得シート反映(ByRef codes() As String, ByVal n As Long)
    Dim ws As Worksheet, i As Long, r As Long, f() As Variant
    Dim itm As Variant
    itm = Array("銘柄名称", "始値", "高値", "安値", "現在値", "出来高")
    Set ws = ThisWorkbook.Worksheets("取得")
    ReDim f(1 To 最大銘柄, 1 To 7)
    For i = 1 To 最大銘柄
        r = i + 1
        If i <= n Then
            f(i, 1) = codes(i)
            f(i, 2) = "=IFERROR(RssMarket($A" & r & ",""" & itm(0) & """),"""")"
            f(i, 3) = "=IFERROR(RssMarket($A" & r & ",""" & itm(1) & """),"""")"
            f(i, 4) = "=IFERROR(RssMarket($A" & r & ",""" & itm(2) & """),"""")"
            f(i, 5) = "=IFERROR(RssMarket($A" & r & ",""" & itm(3) & """),"""")"
            f(i, 6) = "=IFERROR(RssMarket($A" & r & ",""" & itm(4) & """),"""")"
            f(i, 7) = "=IFERROR(RssMarket($A" & r & ",""" & itm(5) & """),"""")"
        Else
            f(i, 1) = "": f(i, 2) = "": f(i, 3) = "": f(i, 4) = ""
            f(i, 5) = "": f(i, 6) = "": f(i, 7) = ""
        End If
    Next i
    ws.Range("A2:A" & (最大銘柄 + 1)).NumberFormat = "@"
    ws.Range("A2:G" & (最大銘柄 + 1)).Formula = f
End Sub

'集積シート1枚ぶんを並べ直す(過去の値はコードごとに引っ越す)
Private Sub 集積1シート反映(ByVal ws As Worksheet, ByRef codes() As String, _
                            ByRef names() As String, ByVal n As Long)
    Dim old As Variant, dat As Variant
    Dim nAB() As Variant, nD() As Variant
    Dim d As Object, i As Long, j As Long, r As Long, lastC As Long, w As Long
    Dim endRow As Long

    endRow = 開始行 + 最大銘柄 - 1
    lastC = ws.Cells(日付行(ws.Name), ws.Columns.Count).End(xlToLeft).Column
    If lastC < 3 Then lastC = 3
    w = lastC - 2

    old = ws.Range(ws.Cells(開始行, 1), ws.Cells(endRow, 2)).Value
    dat = ws.Range(ws.Cells(開始行, 3), ws.Cells(endRow, lastC)).Value

    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To 最大銘柄
        If Len(Trim$(CStr(old(i, 1)))) > 0 Then
            If Not d.Exists(Trim$(CStr(old(i, 1)))) Then d(Trim$(CStr(old(i, 1)))) = i
        End If
    Next i

    ReDim nAB(1 To 最大銘柄, 1 To 2)
    ReDim nD(1 To 最大銘柄, 1 To w)
    For i = 1 To 最大銘柄
        If i <= n Then
            nAB(i, 1) = codes(i)
            r = 0
            If d.Exists(codes(i)) Then r = CLng(d(codes(i)))
            If Len(names(i)) > 0 Then
                nAB(i, 2) = names(i)
            ElseIf r > 0 Then
                nAB(i, 2) = old(r, 2)
            Else
                nAB(i, 2) = ""
            End If
            If r > 0 Then
                For j = 1 To w
                    nD(i, j) = dat(r, j)
                Next j
            End If
        Else
            nAB(i, 1) = "": nAB(i, 2) = ""
        End If
    Next i

    ws.Range(ws.Cells(開始行, 1), ws.Cells(endRow, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(開始行, 1), ws.Cells(endRow, 2)).Value = nAB
    ws.Range(ws.Cells(開始行, 3), ws.Cells(endRow, lastC)).Value = nD
End Sub

'=========================================================
'  3. 集積シート整備 (19シート)
'=========================================================
Public Sub 集積シート整備()
    Dim shts As Variant, i As Long, ws As Worksheet
    Dim keep As Long, over As Long, lastC As Long, dr As Long

    keep = 履歴列数
    shts = 集積シート一覧()

    '削られる列があるか先に数える
    For i = 0 To UBound(shts)
        If 有(CStr(shts(i))) Then
            Set ws = ThisWorkbook.Worksheets(CStr(shts(i)))
            dr = 日付行(CStr(shts(i)))
            lastC = ws.Cells(dr, ws.Columns.Count).End(xlToLeft).Column
            If lastC > 2 + keep Then over = over + (lastC - 2 - keep)
        End If
    Next i

    If over > 0 Then
        If MsgBox("見出しと書式を整えます。" & vbLf & vbLf & _
                  keep & "列(=" & keep & "回分)より古い列 合計 " & over & "列を削除します。" & vbLf & _
                  "削除された古いデータは戻せません。よろしいですか?" & vbLf & _
                  "(いいえ = 削除せず、見出しと書式だけ整えます)", _
                  vbYesNoCancel + vbExclamation, "集積シート整備") = vbCancel Then Exit Sub
        If MsgBox("古い列を削除しますか?" & vbLf & "「いいえ」なら列はそのまま残します。", _
                  vbYesNo + vbQuestion, "集積シート整備") <> vbYes Then keep = 0
    End If

    Application.ScreenUpdating = False
    集積見出し
    For i = 0 To UBound(shts)
        If 有(CStr(shts(i))) Then
            Set ws = ThisWorkbook.Worksheets(CStr(shts(i)))
            dr = 日付行(CStr(shts(i)))
            lastC = ws.Cells(dr, ws.Columns.Count).End(xlToLeft).Column
            If lastC < 3 Then lastC = 3
            ws.Range(ws.Cells(開始行, 1), ws.Cells(開始行 + 最大銘柄 - 1, 1)).NumberFormat = "@"
            ws.Range(ws.Cells(開始行, 3), ws.Cells(開始行 + 最大銘柄 - 1, lastC)).NumberFormat = "#,##0"
            ws.Range(ws.Cells(dr, 3), ws.Cells(dr, lastC)).NumberFormat = "m/d"
            ws.Columns("A").ColumnWidth = 8
            ws.Columns("B").ColumnWidth = 18
            If keep >= 30 And lastC > 2 + keep Then
                ws.Range(ws.Columns(2 + keep + 1), ws.Columns(lastC)).Delete
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    記 "情報", "集積19シートを整えました(" & keep & "列ぶん / 削除 " & IIf(keep >= 30, over, 0) & "列)"
    MsgBox "集積シートを整えました。", vbInformation, "集積シート整備"
End Sub

Private Sub 集積見出し()
    Dim shts As Variant, i As Long, ws As Worksheet, nm As String
    shts = 集積シート一覧()
    For i = 0 To UBound(shts)
        nm = CStr(shts(i))
        If 有(nm) Then
            Set ws = ThisWorkbook.Worksheets(nm)
            If 日付行(nm) = 1 Then
                埋 ws, "A1", "日付"
                埋 ws, "A2", "曜日/取得"
            Else
                埋 ws, "A2", "日付"
            End If
            埋 ws, "A3", "コード"
            埋 ws, "B3", "銘柄名"
        End If
    Next i
End Sub

'日足5シートは1行目が日付、時間14シートは2行目が日付
Private Function 日付行(ByVal nm As String) As Long
    Select Case nm
        Case "始値", "高値", "安値", "終値", "出来高": 日付行 = 1
        Case Else: 日付行 = 2
    End Select
End Function

Private Function 集積シート一覧() As Variant
    Dim tm As Variant, r() As String, i As Long, n As Long
    Dim hi As Variant
    hi = Array("始値", "高値", "安値", "終値", "出来高")
    tm = 時間シート名()
    ReDim r(0 To UBound(hi) + UBound(tm) + 1)
    For i = 0 To UBound(hi)
        r(n) = CStr(hi(i)): n = n + 1
    Next i
    For i = 0 To UBound(tm)
        r(n) = CStr(tm(i)): n = n + 1
    Next i
    集積シート一覧 = r
End Function

'=========================================================
'  小物
'=========================================================
Private Function 有(ByVal n As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
    有 = Not ws Is Nothing
End Function

'空のセルにだけ値を入れる
Private Sub 埋(ByVal ws As Worksheet, ByVal adr As String, ByVal v As Variant)
    If Len(CStr(ws.Range(adr).Value)) = 0 Then ws.Range(adr).Value = v
End Sub

Private Function 時刻を分数に(ByVal x As Variant) As Double
    Dim s As String, d As Date, n As Double
    時刻を分数に = 0
    If IsEmpty(x) Then Exit Function
    If IsDate(x) Then
        d = CDate(x)
        時刻を分数に = CDbl(d) - Int(CDbl(d))
        Exit Function
    End If
    If IsNumeric(x) Then
        n = CDbl(x)
        If n >= 1 And n <= 2359 Then
            時刻を分数に = TimeSerial(Int(n / 100), CLng(n) Mod 100, 0)
        ElseIf n > 0 And n < 1 Then
            時刻を分数に = n
        End If
        Exit Function
    End If
    s = Trim$(CStr(x))
    If Len(s) = 0 Then Exit Function
    On Error Resume Next
    d = CDate(s)
    If Err.Number = 0 Then 時刻を分数に = CDbl(d) - Int(CDbl(d))
    Err.Clear
End Function

Private Sub 記(ByVal kind As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    If Not 有("ログ") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("ログ")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "m/d hh:mm:ss"
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).Value = msg
End Sub
