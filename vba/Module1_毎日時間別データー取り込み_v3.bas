Attribute VB_Name = "Module1"
Option Explicit
'=========================================================
'  毎日時間別データー取り込み  v3.1 (完成版)
'
'  【今回直したところ】
'  (1) 設定シートの取得時刻(9:00 など)を読めず、予約が1件も
'      入らなかった不具合を修正。
'      → VBAの IsNumeric は「時刻」に False を返すため、
'         全スロットが「時刻なし」と判定されていた。
'  (2) Application.OnTime に引数つきの長い文字列を渡す方式をやめ、
'      予約表(設定シート AA:AC)に書いて、窓口を1本にした。
'      → 取消しが確実にでき、他ブックと名前がぶつからない。
'  (3) 3分ごとの「見張り」を追加。パソコンが忙しくて予約を
'      取りこぼしても、自動で追いつく。
'  (4) 日付が変わったら翌日分を自動で予約し直す。
'  (5) シートが無いスロットは予約せずログに残す。
'  (6) 記録シートの日次整理は、CSVへ退避してから削除する方式に変更。
'  (7) 設定!D10 に残す列数(例 250)を入れると、集積シートの古い列を自動で捨てる。
'      空のままなら何も消さない。
'
'  【シートの決まり】
'  時間シート(0900～1530) 1行目=空 2行目=日付 3行目=見出し 4行目～データ
'  日足シート(始値～出来高) 1行目=日付 2行目=曜日 3行目=見出し 4行目～
'  いずれも C列が最新(左が新しい)
'=========================================================

Private Const RETRY_SEC As Long = 60      'ダメだった時に次を試すまでの秒数
Private Const RETRY_MAX As Long = 30      '60秒×30 = 約30分であきらめ
Private Const HB_SEC    As Long = 180     '見張りの間隔(秒)

Private Const C_TIME As Long = 27         'AA列 予定時刻
Private Const C_JOB  As Long = 28         'AB列 内容
Private Const C_STAT As Long = 29         'AC列 状態

Private mBusy   As Boolean                '取得中フラグ(二重実行防止)
Private mBusyAt As Double                 'いつから取得中か

'================== 起動・停止 ==================
Public Sub 自動起動()
    If Trim$(CStr(SH("設定").Range("D7").Value)) = "○" Then 監視開始
End Sub

Public Sub 監視開始()
    Dim i As Long, t As Double, dt As Date, nm As String, n As Long

    予約全解除

    With SH("設定")
        For i = 2 To 15
            t = 実取得時刻(i)
            nm = スロット名(i)
            If t > 0 And Len(nm) > 0 Then
                If シートあり(nm) Then
                    dt = CDate(Int(Now) + t)
                    If dt > Now + TimeSerial(0, 0, 2) Then
                        予約追加 dt, "SLOT|" & nm & "|0"
                        n = n + 1
                    End If
                Else
                    ログ書込 "警告", nm & " というシートが無いので予約しません"
                End If
            End If
        Next i

        If Trim$(CStr(.Range("D6").Value)) = "○" Then
            For i = 8 To 9
                t = 時刻値(.Cells(i, 4).Value)
                If t > 0 Then
                    dt = CDate(Int(Now) + t)
                    If dt > Now Then 予約追加 dt, "SAVE"
                End If
            Next i
        End If
    End With

    予約追加 翌朝830, "REPLAN"                       '日付が変わったら翌日分を予約
    予約追加 CDate(Now + HB_SEC / 86400#), "HB"      '見張り

    ログ書込 "情報", "監視を開始しました(本日の残り " & n & " スロットを予約)"
    状態表示 "予約済 " & n & " 件  次回 " & 次回予約表示
End Sub

Public Sub 監視停止()
    予約全解除
    ログ書込 "情報", "監視を停止しました"
    状態表示 "(監視していません)"
End Sub

Public Sub 再予約()
    監視開始
End Sub

Public Sub 自動保存()
    On Error Resume Next
    ThisWorkbook.Save
    ログ書込 "情報", "ブックを保存しました"
End Sub

'手動テスト用(シート名を入れて実行 例:1530)
Public Sub 手動取得()
    Dim s As String
    s = InputBox("取得して書き込むシート名(例 0900 / 1530)", "手動取得", Format$(Now, "hhnn"))
    s = Trim$(s)
    If Len(s) = 0 Then Exit Sub
    If Not シートあり(s) Then
        MsgBox s & " というシートがありません", vbExclamation, "手動取得"
        Exit Sub
    End If
    スロット実行 s, RETRY_MAX
End Sub

'今すぐ予約表を点検する(ボタン用・手で押してもよい)
Public Sub 今すぐ点検()
    予約処理
    状態表示 "点検 " & Format$(Now, "hh:nn:ss") & "  次回 " & 次回予約表示
End Sub

'================== 予約の窓口(OnTime はここだけを呼ぶ) ==================
Public Sub 予約時刻到来()
    予約処理
End Sub

Private Sub 予約処理()
    Dim ws As Worksheet, r As Long, lastR As Long, job As String
    Set ws = SH("設定")
    lastR = ws.Cells(ws.Rows.Count, C_TIME).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, C_STAT).Value) = "待機" Then
            If 予定時刻(ws, r) > 0 And 予定時刻(ws, r) <= CDbl(Now) + 2# / 86400# Then
                job = CStr(ws.Cells(r, C_JOB).Value)
                ws.Cells(r, C_STAT).Value = "済"
                If ジョブ実行(job) Then Exit Sub      '予約表を作り直したので抜ける
            End If
        End If
    Next r
End Sub

'戻り値 True = 予約表を作り直したので、これ以上ループしない
Private Function ジョブ実行(ByVal job As String) As Boolean
    Dim p As Variant
    ジョブ実行 = False
    If Len(job) = 0 Then Exit Function
    p = Split(job, "|")
    Select Case CStr(p(0))
        Case "SLOT"
            スロット実行 CStr(p(1)), CLng(p(2))
        Case "SAVE"
            自動保存
        Case "REPLAN"
            監視開始
            ジョブ実行 = True
        Case "HB"
            予約追加 CDate(Now + HB_SEC / 86400#), "HB"
    End Select
End Function

'================== 本体 ==================
Public Sub スロット実行(ByVal slot As String, ByVal tryCnt As Long)
    Dim v As Variant, ok As Long, i As Long

    If mBusy Then
        If (CDbl(Now) - mBusyAt) * 86400# > 300 Then
            mBusy = False                              '5分以上その状態なら取り残しとみなす
        Else
            予約追加 CDate(Now + 20# / 86400#), "SLOT|" & slot & "|" & tryCnt
            Exit Sub
        End If
    End If

    mBusy = True
    mBusyAt = CDbl(Now)
    On Error GoTo EH

    Application.Calculate
    待機 数値(SH("設定").Range("D5").Value, 5)
    Application.Calculate

    v = SH("取得").Range("A2:G301").Value
    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 6)) Then If v(i, 6) > 0 Then ok = ok + 1
    Next i

    If ok = 0 Then
        If tryCnt < RETRY_MAX Then
            予約追加 CDate(Now + RETRY_SEC / 86400#), _
                     "SLOT|" & slot & "|" & (tryCnt + 1)
        Else
            ログ書込 "警告", slot & " は約30分過ぎても取れませんでした(あきらめ)"
            ブザー
        End If
        mBusy = False
        Exit Sub
    End If

    時間シート書込 slot, v
    記録追加 slot, v
    If slot = "1530" Then
        日足書込 v
        記録日次整理
    End If

    ログ書込 "取得", slot & "  " & UBound(v, 1) & "銘柄 (値あり " & ok & ")"
    状態表示 "最終取得 " & slot & "  " & Format$(Now, "hh:nn:ss") & "  次回 " & 次回予約表示
    mBusy = False
    Exit Sub
EH:
    mBusy = False
    ログ書込 "エラー", slot & " 取得中にエラー: " & Err.Description
End Sub

'--- 時間シート(1日1列・現在値) ---
Private Sub 時間シート書込(ByVal slot As String, ByVal v As Variant)
    Dim ws As Worksheet, i As Long, key As String
    Dim codes As Variant, cur As Variant, d As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(slot)
    On Error GoTo 0
    If ws Is Nothing Then
        ログ書込 "警告", slot & " シートが無いので書き込めません"
        Exit Sub
    End If

    If ws.Cells(2, 3).Value2 <> CDbl(Date) Then
        ws.Columns("C").Insert Shift:=xlToRight
        ws.Range("D1:D303").Copy
        ws.Range("C1:C303").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ws.Range("C1:C303").ClearContents
        ws.Cells(2, 3).Value = Date
        With ws.Cells(2, 3)
            .NumberFormat = "m/d"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 235, 235)
        End With
        ws.Range("C4:C303").NumberFormat = "#,##0"
        列トリム ws, 2
    End If

    codes = ws.Range("A4:A303").Value
    cur = ws.Range("C4:C303").Value
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To 300
        If Len(codes(i, 1)) > 0 Then d(CStr(codes(i, 1))) = i
    Next i
    For i = 1 To UBound(v, 1)
        key = CStr(v(i, 1))
        If d.Exists(key) Then
            If IsNumeric(v(i, 6)) Then If v(i, 6) > 0 Then cur(d(key), 1) = v(i, 6)
        End If
    Next i
    ws.Range("C4:C303").Value = cur
End Sub

'--- 日足5シート(1530のときだけ) ---
Private Sub 日足書込(ByVal v As Variant)
    Dim nms As Variant, cols As Variant, k As Long
    nms = Array("始値", "高値", "安値", "終値", "出来高")
    cols = Array(3, 4, 5, 6, 7)          '取得シート C,D,E,F(現在値=終値),G
    For k = 0 To 4
        日足1シート CStr(nms(k)), CLng(cols(k)), v
    Next k
End Sub

Private Sub 日足1シート(ByVal shName As String, ByVal srcCol As Long, ByVal v As Variant)
    Dim ws As Worksheet, i As Long, key As String
    Dim codes As Variant, cur As Variant, d As Object

    If Not シートあり(shName) Then
        ログ書込 "警告", shName & " シートが無いので書き込めません"
        Exit Sub
    End If
    Set ws = SH(shName)

    If ws.Cells(1, 3).Value2 <> CDbl(Date) Then
        ws.Columns("C").Insert Shift:=xlToRight
        ws.Range("D1:D303").Copy
        ws.Range("C1:C303").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ws.Range("C1:C303").ClearContents
        If InStr(CStr(ws.Cells(2, 4).Value), "本日") > 0 Then
            If IsDate(ws.Cells(1, 4).Value) Then ws.Cells(2, 4).Value = 曜日(CDate(ws.Cells(1, 4).Value))
        End If
        ws.Cells(1, 3).Value = Date
        ws.Cells(2, 3).Value = "本日" & Format$(Now, "hh:nn") & "取得"
        列トリム ws, 1
    End If

    codes = ws.Range("A4:A303").Value
    cur = ws.Range("C4:C303").Value
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To 300
        If Len(codes(i, 1)) > 0 Then d(CStr(codes(i, 1))) = i
    Next i
    For i = 1 To UBound(v, 1)
        key = CStr(v(i, 1))
        If d.Exists(key) Then
            If IsNumeric(v(i, srcCol)) Then If v(i, srcCol) > 0 Then cur(d(key), 1) = v(i, srcCol)
        End If
    Next i
    ws.Range("C4:C303").Value = cur
End Sub

'--- 記録シートへ追記(コードが空の行は書かない) ---
Private Sub 記録追加(ByVal slot As String, ByVal v As Variant)
    Dim ws As Worksheet, r As Long, i As Long, n As Long, o() As Variant
    Set ws = SH("記録")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ReDim o(1 To UBound(v, 1), 1 To 10)
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, 1)))) > 0 Then
            n = n + 1
            o(n, 1) = Date
            o(n, 2) = スロット時刻(slot)
            o(n, 3) = Format$(Now, "hh:nn:ss")
            o(n, 4) = v(i, 1): o(n, 5) = v(i, 2)
            o(n, 6) = v(i, 3): o(n, 7) = v(i, 4): o(n, 8) = v(i, 5)
            o(n, 9) = v(i, 6): o(n, 10) = v(i, 7)
        End If
    Next i
    If n = 0 Then Exit Sub
    ws.Cells(r, 1).Resize(n, 10).Value = o
End Sub

'================== 予約の管理 ==================
Private Function 呼び出し先() As String
    呼び出し先 = "'" & ThisWorkbook.Name & "'!予約時刻到来"
End Function

Private Sub 予約追加(ByVal dt As Date, ByVal job As String)
    Dim ws As Worksheet, r As Long

    dt = CDate(Int(CDbl(dt) * 86400# + 0.5) / 86400#)   '1秒単位にそろえる(取消しを確実に)

    On Error Resume Next
    Application.OnTime EarliestTime:=dt, Procedure:=呼び出し先(), Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ログ書込 "エラー", "予約できませんでした " & Format$(dt, "m/d hh:nn:ss") & "  " & job
        Exit Sub
    End If
    On Error GoTo 0

    Set ws = SH("設定")
    If Len(CStr(ws.Range("AA1").Value)) = 0 Then
        ws.Range("AA1").Value = "予定時刻(自動)"
        ws.Range("AB1").Value = "内容"
        ws.Range("AC1").Value = "状態"
    End If
    r = ws.Cells(ws.Rows.Count, C_TIME).End(xlUp).Row + 1
    If r < 2 Then r = 2
    With ws.Cells(r, C_TIME)
        .NumberFormat = "m/d hh:mm:ss"
        .Value = dt
    End With
    ws.Cells(r, C_JOB).Value = job
    ws.Cells(r, C_STAT).Value = "待機"
End Sub

Private Sub 予約全解除()
    Dim ws As Worksheet, r As Long, lastR As Long, t As Double
    Set ws = SH("設定")
    lastR = ws.Cells(ws.Rows.Count, C_TIME).End(xlUp).Row
    For r = 2 To lastR
        t = 予定時刻(ws, r)
        If t > 0 And CStr(ws.Cells(r, C_STAT).Value) = "待機" Then
            On Error Resume Next
            Application.OnTime EarliestTime:=CDate(t), Procedure:=呼び出し先(), Schedule:=False
            Err.Clear
            On Error GoTo 0
        End If
    Next r
    If lastR >= 2 Then ws.Range(ws.Cells(2, C_TIME), ws.Cells(lastR, C_STAT)).ClearContents
End Sub

Private Function 予定時刻(ByVal ws As Worksheet, ByVal r As Long) As Double
    Dim x As Variant
    予定時刻 = 0
    x = ws.Cells(r, C_TIME).Value2
    If IsEmpty(x) Then Exit Function
    If Not IsNumeric(x) Then Exit Function
    予定時刻 = CDbl(x)
End Function

Private Function 次回予約表示() As String
    Dim ws As Worksheet, r As Long, lastR As Long, m As Double, x As Double
    Set ws = SH("設定")
    lastR = ws.Cells(ws.Rows.Count, C_TIME).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, C_STAT).Value) = "待機" Then
            x = 予定時刻(ws, r)
            If x > CDbl(Now) And CStr(ws.Cells(r, C_JOB).Value) <> "HB" Then
                If m = 0 Or x < m Then m = x
            End If
        End If
    Next r
    If m = 0 Then 次回予約表示 = "なし" Else 次回予約表示 = Format$(CDate(m), "m/d hh:mm")
End Function

'================== 予約一覧 ==================
Public Sub 予約一覧表示()
    Dim ws As Worksheet, r As Long, lastR As Long, s As String, n As Long, x As Double
    Set ws = SH("設定")
    lastR = ws.Cells(ws.Rows.Count, C_TIME).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, C_STAT).Value) = "待機" Then
            x = 予定時刻(ws, r)
            If x > 0 Then
                n = n + 1
                If n <= 40 Then
                    s = s & Format$(CDate(x), "m/d hh:mm:ss") & "   " _
                          & CStr(ws.Cells(r, C_JOB).Value) & vbLf
                End If
            End If
        End If
    Next r
    If n = 0 Then s = "予約はありません(監視していません)"
    MsgBox "待機中の予約 " & n & " 件" & vbLf & vbLf & s, vbInformation, "予約一覧"
End Sub

'================== 小物 ==================
'履歴が増えすぎたら古い列(右側)を消す
'  設定!D10 に「残す列数(例 250)」を入れたときだけ働く。空なら何も消さない。
Private Sub 列トリム(ByVal ws As Worksheet, ByVal dateRow As Long)
    Dim keep As Long, lastC As Long
    keep = 数値(SH("設定").Range("D10").Value, 0)
    If keep < 30 Then Exit Sub
    lastC = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column
    If lastC > 2 + keep Then
        ws.Range(ws.Columns(2 + keep + 1), ws.Columns(lastC)).Delete
        ログ書込 "情報", ws.Name & " の古い列 " & (lastC - 2 - keep) & "列を削除(" & keep & "日分に調整)"
    End If
End Sub

Private Function SH(ByVal n As String) As Worksheet
    Set SH = ThisWorkbook.Worksheets(n)
End Function

Private Function シートあり(ByVal n As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
    シートあり = Not ws Is Nothing
End Function

Private Sub 状態表示(ByVal msg As String)
    On Error Resume Next
    SH("取得").Range("I1").Value = msg
End Sub

'設定シートの時刻セルを「1日の中の時間」(0～1)に直す
'  9:00 / 0.375 / "9:00" / 900 のどれでも読める
Private Function 時刻値(ByVal x As Variant) As Double
    Dim s As String, d As Date, n As Double
    時刻値 = 0
    If IsEmpty(x) Then Exit Function
    If IsDate(x) Then
        d = CDate(x)
        時刻値 = CDbl(d) - Int(CDbl(d))
        Exit Function
    End If
    If IsNumeric(x) Then
        n = CDbl(x)
        If n >= 1 And n <= 2359 Then                 'hhmm 形式で書かれた場合
            時刻値 = TimeSerial(Int(n / 100), CLng(n) Mod 100, 0)
        ElseIf n > 0 And n < 1 Then
            時刻値 = n
        End If
        Exit Function
    End If
    s = Trim$(CStr(x))
    If Len(s) = 0 Then Exit Function
    On Error Resume Next
    d = CDate(s)
    If Err.Number = 0 Then 時刻値 = CDbl(d) - Int(CDbl(d))
    Err.Clear
End Function

'B列(実取得時刻)が空ならA列を使う
Private Function 実取得時刻(ByVal rw As Long) As Double
    Dim t As Double
    With SH("設定")
        t = 時刻値(.Cells(rw, 2).Value)
        If t <= 0 Then t = 時刻値(.Cells(rw, 1).Value)
    End With
    実取得時刻 = t
End Function

'シート名はA列(表示用の時刻)から作る 例 9:00 → "0900"
Private Function スロット名(ByVal rw As Long) As String
    Dim t As Double
    t = 時刻値(SH("設定").Cells(rw, 1).Value)
    If t <= 0 Then t = 時刻値(SH("設定").Cells(rw, 2).Value)
    If t <= 0 Then
        スロット名 = ""
    Else
        スロット名 = Format$(CDate(t), "hhnn")
    End If
End Function

Private Function 翌朝830() As Date
    翌朝830 = Int(Now) + TimeSerial(8, 30, 0)
    If 翌朝830 <= Now Then 翌朝830 = 翌朝830 + 1
End Function

Private Function スロット時刻(ByVal slot As String) As Double
    On Error Resume Next
    スロット時刻 = TimeSerial(CLng(Left$(slot, 2)), CLng(Right$(slot, 2)), 0)
    Err.Clear
End Function

Private Function 曜日(ByVal d As Date) As String
    曜日 = Mid$("日月火水木金土", Weekday(d, vbSunday), 1)
End Function

Private Sub 待機(ByVal sec As Long)
    Dim t As Double
    If sec <= 0 Then Exit Sub
    t = Timer + sec
    Do While Timer < t
        DoEvents
    Loop
End Sub

Private Function 数値(ByVal x As Variant, ByVal def As Long) As Long
    If IsNumeric(x) Then 数値 = CLng(x) Else 数値 = def
End Function

Private Sub ブザー()
    Dim i As Long
    If Trim$(CStr(SH("設定").Range("D2").Value)) <> "○" Then Exit Sub
    For i = 1 To 3
        Beep
        待機 1
    Next i
End Sub

Private Sub ログ書込(ByVal kind As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SH("ログ")
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "m/d hh:mm:ss"
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).Value = msg
End Sub

'================== ボタンパネル作成 ==================
'これを1回だけ実行すると 設定シートのボタンを作り直します
Public Sub ボタン再作成()
    Dim ws As Worksheet, b As Object, i As Long
    Dim nm As Variant, mc As Variant
    Set ws = SH("設定")
    For i = ws.Buttons.Count To 1 Step -1
        ws.Buttons(i).Delete
    Next i
    nm = Array("監視 開始", "監視 停止", "手動で取得", "ブック保存", "予約を確認", "今すぐ点検", "記録をCSV出力")
    mc = Array("監視開始", "監視停止", "手動取得", "自動保存", "予約一覧表示", "今すぐ点検", "記録CSV出力")
    For i = 0 To UBound(nm)
        Set b = ws.Buttons.Add(620, 20 + i * 34, 150, 28)
        b.Caption = nm(i)
        b.OnAction = "'" & ThisWorkbook.Name & "'!" & mc(i)
        b.Font.Size = 11
    Next i
    ログ書込 "情報", "ボタンを作り直しました(" & (UBound(nm) + 1) & "個)"
End Sub

'================== 記録CSV出力 ==================
'設定!D1 のフォルダへ 当日分の記録をCSVで書き出す
Public Sub 記録CSV出力()
    Dim ws As Worksheet, fso As Object, ts As Object
    Dim fol As String, fn As String
    Dim v As Variant, i As Long, j As Long, lastR As Long, n As Long
    Dim line As String
    Set ws = SH("記録")
    fol = Trim$(CStr(SH("設定").Range("D1").Value))
    If Len(fol) = 0 Then MsgBox "設定!D1 に保存フォルダを入れてください", vbExclamation: Exit Sub
    If Right$(fol, 1) <> "\" Then fol = fol & "\"
    fn = fol & "記録_" & Format$(Date, "yyyymmdd") & ".csv"
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    v = ws.Range("A1:J" & lastR).Value
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo EH
    Set ts = fso.CreateTextFile(fn, True)
    ts.WriteLine "日付,時刻区分,取得時刻,コード,銘柄名,始値,高値,安値,現在値,出来高"
    For i = 2 To lastR
        If IsDate(v(i, 1)) Or IsNumeric(v(i, 1)) Then
            If CDbl(v(i, 1)) = CDbl(Date) Then
                line = Format$(v(i, 1), "yyyy/mm/dd") & "," & Format$(v(i, 2), "hh:mm")
                For j = 3 To 10
                    line = line & "," & CStr(v(i, j))
                Next j
                ts.WriteLine line
                n = n + 1
            End If
        End If
    Next i
    ts.Close
    ログ書込 "情報", "CSV出力 " & n & "行  " & fn
    MsgBox n & " 行を書き出しました" & vbLf & fn, vbInformation
    Exit Sub
EH:
    ログ書込 "エラー", "CSV出力に失敗: " & Err.Description
    MsgBox "CSV出力に失敗: " & Err.Description, vbExclamation
End Sub

'================== 記録シートの日次整理 ==================
'当日より前の行を 日付ごとのCSVへ退避してから削除する(1530取得後に自動実行)
'  退避先は 設定!D1 のフォルダ。書き出せなかったときは削除しません。
Public Sub 記録日次整理()
    Dim ws As Worksheet, v As Variant, i As Long, lastR As Long, cut As Long
    Set ws = SH("記録")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    v = ws.Range("A2:A" & lastR).Value
    For i = 1 To UBound(v, 1)
        If Not (IsDate(v(i, 1)) Or IsNumeric(v(i, 1))) Then Exit For
        If CDbl(v(i, 1)) >= CDbl(Date) Then Exit For     '当日分は残す
        cut = i + 1                                      'シート上の行番号
    Next i
    If cut < 2 Then Exit Sub

    If 記録退避CSV(2, cut) Then
        ws.Rows("2:" & cut).Delete
        ログ書込 "情報", "記録整理 " & (cut - 1) & "行をCSVへ退避して削除(残り " _
                       & (ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1) & "行)"
    Else
        ログ書込 "警告", "記録をCSVへ退避できなかったので削除しませんでした(設定!D1 を確認)"
    End If
End Sub

'指定した行範囲を 日付ごとの 記録_yyyymmdd.csv へ書き出す(成功=True)
Private Function 記録退避CSV(ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim ws As Worksheet, fso As Object, ts As Object
    Dim fol As String, fn As String, cur As String, d As String, line As String
    Dim v As Variant, i As Long, j As Long

    記録退避CSV = False
    fol = Trim$(CStr(SH("設定").Range("D1").Value))
    If Len(fol) = 0 Then Exit Function
    If Right$(fol, 1) <> "\" Then fol = fol & "\"

    Set ws = SH("記録")
    v = ws.Range("A" & r1 & ":J" & r2).Value
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Left$(fol, Len(fol) - 1)) Then Exit Function

    On Error GoTo EH
    For i = 1 To UBound(v, 1)
        d = Format$(v(i, 1), "yyyymmdd")
        If d <> cur Then
            If Not ts Is Nothing Then ts.Close: Set ts = Nothing
            fn = fol & "記録_" & d & ".csv"
            If fso.FileExists(fn) Then
                Set ts = fso.OpenTextFile(fn, 8)          '追記
            Else
                Set ts = fso.CreateTextFile(fn, True)
                ts.WriteLine "日付,時刻区分,取得時刻,コード,銘柄名,始値,高値,安値,現在値,出来高"
            End If
            cur = d
        End If
        line = Format$(v(i, 1), "yyyy/mm/dd") & "," & Format$(v(i, 2), "hh:mm")
        For j = 3 To 10
            line = line & "," & CStr(v(i, j))
        Next j
        ts.WriteLine line
    Next i
    If Not ts Is Nothing Then ts.Close
    記録退避CSV = True
    Exit Function
EH:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    ログ書込 "エラー", "記録の退避に失敗: " & Err.Description
End Function

'================== 過去分の一括退避(1回だけ実行) ==================
Public Sub 記録過去分エクスポート()
    If MsgBox("記録シートの当日より前の行を、日付ごとのCSVに書き出してから削除します。" _
            & vbLf & "書き出し先: " & CStr(SH("設定").Range("D1").Value) _
            & vbLf & vbLf & "先にブックのバックアップを取りましたか?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    記録日次整理
    Application.ScreenUpdating = True
    MsgBox "完了しました。ログを確認してください。", vbInformation
End Sub

Public Sub CSV削除()
    Dim fso As Object, fol As String
    Dim f As Object, nm() As String, i As Long, n As Long
    fol = Trim$(CStr(SH("設定").Range("D1").Value))
    If Len(fol) = 0 Then MsgBox "設定!D1 に保存フォルダを入れてください", vbExclamation: Exit Sub
    If Right$(fol, 1) <> "\" Then fol = fol & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReDim nm(1 To 20000)
    For Each f In fso.GetFolder(fol).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "csv" Then
            If Left$(f.Name, 3) = "記録_" Then
                n = n + 1
                nm(n) = f.Name
            End If
        End If
    Next f
    If n = 0 Then MsgBox "対象のCSVはありません", vbInformation: Exit Sub
    If MsgBox(fol & vbLf & "の 記録_*.csv を " & n & " 件、完全削除します。よろしいですか?", _
              vbYesNo + vbExclamation, "CSV削除") <> vbYes Then Exit Sub
    For i = 1 To n
        On Error Resume Next
        fso.DeleteFile fol & nm(i), True
        On Error GoTo 0
    Next i
    ログ書込 "情報", "CSV " & n & "件を削除(" & fol & ")"
    MsgBox n & " 件を削除しました", vbInformation
End Sub
