Attribute VB_Name = "Mod_OHLCV取得"
Option Explicit
'==================================================================
' OHLCV 時刻別 取得BOOK   v1.0
'
' 【このBOOKの目的】
'   決められた時刻の OHLCV を「確実に残す」ことだけ。
'   計算・抽出はしない。
'
' 【取得時刻】設定シート A列 で自由に変更できます
'   9:00 9:10 9:20 9:45 10:00 10:15 10:30
'   11:30 12:30 13:30 14:30 15:00 15:20 15:30
'
' 【PCトラブル対策】
'   1回取るたびに   (1)記録シートに追記
'                   (2)CSVファイルに追記   ← ここが本命
'                   (3)ブックを自動保存
'   CSVは1行ずつ書き込むので、途中で落ちてもそこまでは必ず残ります。
'
' 【RSS切断】
'   30秒ごとに監視。切れたらブザーを30秒おきに10分間鳴らします。
'   「ブザー停止」ボタンでいつでも止まります。
'
' 【他BOOKとの競合】
'   ・取得は「時刻ちょうど」ではなく 設定の遅延秒(既定5秒)後に実行
'   ・CSV書き込みはロックファイルで排他
'   ・取れなかったら30秒後に自動で再挑戦(30分以内)
'   ・自動処理中は MsgBox を一切出しません(止まらないため)
'==================================================================

Private Const SH_MEI As String = "銘柄"
Private Const SH_BAN As String = "取得"
Private Const SH_REC As String = "記録"
Private Const SH_LOG As String = "ログ"
Private Const SH_CFG As String = "設定"

Private Const ロック名 As String = "OHLCV_LOCK.tmp"
Private Const 間隔秒 As Long = 30          ' 監視タイマーの間隔
Private Const 猶予分 As Long = 30          ' 取得時刻を何分過ぎるまで挑戦するか

Private g監視中 As Boolean
Private g実行中 As Boolean
Private g次回 As Date

Private gブザー中 As Boolean
Private gブザー終了 As Date
Private g次ブザー As Date

Private g前回出来高 As Double
Private g出来高時刻 As Date


'==================================================================
'  1. 初期設定  ← 最初に1回だけ実行してください
'==================================================================
Public Sub 初期設定()
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    '--- 設定シート ---
    Set ws = シート確保(SH_CFG)
    If CStr(ws.Range("A1").Value) = "" Then 設定初期値 ws
    設定補完 ws
    ws.Columns("H").NumberFormat = "@"

    '--- 銘柄シート ---
    Set ws = シート確保(SH_MEI)
    ws.Range("A1").Value = "コード"
    ws.Range("B1").Value = "銘柄名(自動)"
    ws.Range("C1").Value = "止める時は×"
    見出し ws.Range("A1:C1")
    ws.Columns("A").NumberFormat = "@"
    ws.Columns("A:C").ColumnWidth = 16
    ws.Range("E1").Value = "◆ A列にコードを打つ → 「銘柄反映」ボタン"
    ws.Range("E2").Value = "◆ 消す時は その行を丸ごと削除 → 「銘柄反映」ボタン"
    ws.Range("E3").Value = "◆ 一時的に外すだけなら C列に × を入れる"
    ws.Range("E1:E3").Font.Color = RGB(0, 0, 192)

    '--- 取得シート(RSS実況板) ---
    Set ws = シート確保(SH_BAN)
    ws.Range("A1").Value = "コード"
    ws.Range("B1").Value = "銘柄名"
    ws.Range("C1").Value = "始値"
    ws.Range("D1").Value = "高値"
    ws.Range("E1").Value = "安値"
    ws.Range("F1").Value = "現在値"
    ws.Range("G1").Value = "出来高"
    見出し ws.Range("A1:G1")
    ws.Columns("A").NumberFormat = "@"
    ws.Columns("C:G").NumberFormat = "#,##0.##"
    ws.Columns("A:G").ColumnWidth = 13
    ws.Range("I1").ColumnWidth = 30
    ws.Range("I1").Value = "(監視していません)"
    ws.Range("I1").Font.Bold = True
    ws.Range("I2").Value = "↑ RSSの状態。赤くなったら切断です。"

    '--- 記録シート ---
    Set ws = シート確保(SH_REC)
    記録見出し ws

    '--- ログシート ---
    Set ws = シート確保(SH_LOG)
    ws.Range("A1").Value = "日時"
    ws.Range("B1").Value = "種別"
    ws.Range("C1").Value = "内容"
    見出し ws.Range("A1:C1")
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("A").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 70

    ボタン作成
    Application.ScreenUpdating = True

    ログ書く "情報", "初期設定を実行しました"
    MsgBox "初期設定が終わりました。" & vbCrLf & vbCrLf & _
           "1. 銘柄シートのA列に銘柄コードを入れる" & vbCrLf & _
           "2. 「銘柄反映」ボタンを押す" & vbCrLf & _
           "3. 「監視開始」ボタンを押す" & vbCrLf & vbCrLf & _
           "あとは自動です。", vbInformation, "OHLCV取得BOOK"
End Sub


Private Sub 設定初期値(ByVal ws As Worksheet)
    Dim t As Variant
    t = Array("9:00", "9:10", "9:20", "9:45", "10:00", "10:15", "10:30", _
              "11:30", "12:30", "13:30", "14:30", "15:00", "15:20", "15:30")

    ws.Range("A1").Value = "取得時刻"
    見出し ws.Range("A1")
    Dim i As Long
    For i = LBound(t) To UBound(t)
        ws.Cells(i + 2, 1).Value = TimeValue(CStr(t(i)))
    Next i
    ws.Range("A2:A40").NumberFormat = "hh:mm"
    ws.Columns("A").ColumnWidth = 12

    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 46

    ws.Range("H1").Value = ""
    ws.Range("G1").Value = "取得済(本日)"
    見出し ws.Range("G1")
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
End Sub


Private Sub 設定補完(ByVal ws As Worksheet)
    設定1つ ws, "CSV保存フォルダ", ThisWorkbook.Path
    設定1つ ws, "ブザー使用", "○"
    設定1つ ws, "ブザー継続分", 10
    設定1つ ws, "ブザー間隔秒", 30
    設定1つ ws, "取得遅延秒", 5
    設定1つ ws, "自動保存", "○"
    設定1つ ws, "起動時に自動監視", "○"
    ws.Columns("D").ColumnWidth = 46
End Sub

Private Sub 設定1つ(ByVal ws As Worksheet, ByVal 名前 As String, ByVal 既定 As Variant)
    Dim r As Long
    For r = 1 To 30
        If CStr(ws.Cells(r, 3).Value) = 名前 Then Exit Sub
    Next r
    For r = 1 To 30
        If CStr(ws.Cells(r, 3).Value) = "" Then
            ws.Cells(r, 3).Value = 名前
            ws.Cells(r, 4).Value = 既定
            見出し ws.Cells(r, 3)
            Exit Sub
        End If
    Next r
End Sub


Private Sub 記録見出し(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日付"
    ws.Range("B1").Value = "時刻区分"
    ws.Range("C1").Value = "取得時刻"
    ws.Range("D1").Value = "コード"
    ws.Range("E1").Value = "銘柄名"
    ws.Range("F1").Value = "始値"
    ws.Range("G1").Value = "高値"
    ws.Range("H1").Value = "安値"
    ws.Range("I1").Value = "現在値"
    ws.Range("J1").Value = "出来高"
    見出し ws.Range("A1:J1")
    ws.Columns("A").NumberFormat = "yyyy/mm/dd"
    ws.Columns("D").NumberFormat = "@"
    ws.Columns("F:J").NumberFormat = "#,##0.##"
    ws.Columns("A:J").ColumnWidth = 12
    If Not ws.AutoFilterMode Then ws.Range("A1:J1").AutoFilter
End Sub


'==================================================================
'  2. 銘柄反映  ← 銘柄を足した/消したら押す
'==================================================================
Public Sub 銘柄反映()
    Dim mei As Worksheet, ban As Worksheet
    Set mei = ThisWorkbook.Worksheets(SH_MEI)
    Set ban = ThisWorkbook.Worksheets(SH_BAN)

    Application.ScreenUpdating = False

    ' 銘柄シートの銘柄名式を入れ直す
    Dim 最終 As Long
    最終 = mei.Cells(mei.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To 最終
        If Trim$(CStr(mei.Cells(r, 1).Value)) <> "" Then
            mei.Cells(r, 2).Formula = _
                "=IFERROR(@RssMarket($A" & r & ",""銘柄名称""),"""")"
        End If
    Next r

    ' 取得シートを作り直す
    ban.Range("A2:G100000").ClearContents

    Dim w As Long: w = 2
    Dim code As String
    For r = 2 To 最終
        code = Trim$(CStr(mei.Cells(r, 1).Value))
        If code <> "" And InStr(CStr(mei.Cells(r, 3).Value), "×") = 0 Then
            ban.Cells(w, 1).NumberFormat = "@"
            ban.Cells(w, 1).Value = code
            ban.Cells(w, 2).Formula = "=IFERROR(@RssMarket($A" & w & ",""銘柄名称""),"""")"
            ban.Cells(w, 3).Formula = "=IFERROR(@RssMarket($A" & w & ",""始値""),"""")"
            ban.Cells(w, 4).Formula = "=IFERROR(@RssMarket($A" & w & ",""高値""),"""")"
            ban.Cells(w, 5).Formula = "=IFERROR(@RssMarket($A" & w & ",""安値""),"""")"
            ban.Cells(w, 6).Formula = "=IFERROR(@RssMarket($A" & w & ",""現在値""),"""")"
            ban.Cells(w, 7).Formula = "=IFERROR(@RssMarket($A" & w & ",""出来高""),"""")"
            w = w + 1
        End If
    Next r

    Application.ScreenUpdating = True
    ログ書く "情報", "銘柄反映 " & (w - 2) & "銘柄"
    MsgBox (w - 2) & " 銘柄を取得シートにセットしました。", vbInformation, "銘柄反映"
End Sub


'------------------------------------------------------------------
'  銘柄取込 : 他のBOOK(売買BOOKなど)からコードをもらう
'------------------------------------------------------------------
Public Sub 銘柄取込_他BOOKから()
    Dim fn As Variant
    fn = Application.GetOpenFilename("Excelブック,*.xls*", , _
         "銘柄コードのあるBOOKを選んでください(ボリンジャー送信BOOKなど)")
    If VarType(fn) = vbBoolean Then Exit Sub

    Dim パス As String: パス = CStr(fn)
    Dim ファイル名 As String
    ファイル名 = Mid$(パス, InStrRev(パス, "\") + 1)

    Dim wb As Workbook
    Dim 既に開いていた As Boolean

    ' すでに開いているなら、それを使う(勝手に閉じない)
    On Error Resume Next
    Set wb = Workbooks(ファイル名)
    On Error GoTo 0
    既に開いていた = Not (wb Is Nothing)

    Dim 見つけたWs As Worksheet
    Dim 見つけた列 As Long, 件数 As Long
    Dim シート名 As String, 列名 As String
    Dim 集めた As Object
    見つけた列 = 0: 件数 = 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo L_ERR

    If Not 既に開いていた Then
        Set wb = Workbooks.Open(Filename:=パス, ReadOnly:=True, UpdateLinks:=0)
    End If

    '--- 全シートを見て、銘柄コードが一番多い列を自動で探す ---
    Dim ws As Worksheet, c As Long, n As Long
    For Each ws In wb.Worksheets
        c = 0: n = 0
        シート走査 ws, c, n
        If n > 件数 Then
            件数 = n: 見つけた列 = c: Set 見つけたWs = ws
        End If
    Next ws

    If 件数 < 3 Or 見つけたWs Is Nothing Then
        Application.ScreenUpdating = True
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing: Set 見つけたWs = Nothing
        MsgBox "銘柄コードらしい列が見つかりませんでした。" & vbCrLf & _
               "(4桁のコードが3件以上ある列を探しています)", vbExclamation, "銘柄取込"
        Exit Sub
    End If

    ' ★ BOOKを閉じる前に、必要な情報を全部こちらへ移す
    シート名 = 見つけたWs.Name
    列名 = 列文字(見つけた列)
    Set 集めた = コード収集(見つけたWs, 見つけた列)

    BOOK後始末 wb, 既に開いていた
    Set wb = Nothing
    Set 見つけたWs = Nothing

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    '--- ここから先は自分のBOOKだけ ---
    If 集めた.Count = 0 Then
        MsgBox "コードを1件も読み取れませんでした。", vbExclamation, "銘柄取込"
        Exit Sub
    End If

    If MsgBox("次の場所からコードを読み込みました。" & vbCrLf & vbCrLf & _
              "BOOK  : " & ファイル名 & vbCrLf & _
              "シート: " & シート名 & vbCrLf & _
              "列    : " & 列名 & "列" & vbCrLf & _
              "件数  : " & 集めた.Count & " 件 (重複除き)" & vbCrLf & vbCrLf & _
              "銘柄シートに書き込みますか?" & vbCrLf & _
              "※今の銘柄リストは入れ替わります", _
              vbYesNo + vbQuestion, "銘柄取込") <> vbYes Then Exit Sub

    Dim mei As Worksheet: Set mei = ThisWorkbook.Worksheets(SH_MEI)
    Application.ScreenUpdating = False
    mei.Range("A2:C100000").ClearContents

    Dim k As Variant, w As Long: w = 2
    For Each k In 集めた.Keys
        mei.Cells(w, 1).NumberFormat = "@"
        mei.Cells(w, 1).Value = CStr(k)
        w = w + 1
    Next k
    Application.ScreenUpdating = True

    ログ書く "情報", "銘柄取込 " & (w - 2) & "件 (" & ファイル名 & " / " & シート名 & " / " & 列名 & "列)"

    If MsgBox((w - 2) & " 銘柄を書き込みました。" & vbCrLf & vbCrLf & _
              "続けて「銘柄反映」を実行しますか?", _
              vbYesNo + vbQuestion, "銘柄取込") = vbYes Then
        銘柄反映
    End If
    Exit Sub

L_ERR:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    BOOK後始末 wb, 既に開いていた
    Set wb = Nothing
    Set 見つけたWs = Nothing
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ログ書く "エラー", "銘柄取込: " & eN & " " & eD
    MsgBox "取込に失敗しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & eN & vbCrLf & eD, vbExclamation, "銘柄取込"
End Sub


'------------------------------------------------------------------
'  開いたBOOKの後始末(元から開いていたBOOKは閉じない)
'------------------------------------------------------------------
Private Sub BOOK後始末(ByRef wb As Workbook, ByVal 既に開いていた As Boolean)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If 既に開いていた Then Exit Sub
    wb.Close SaveChanges:=False
End Sub


'------------------------------------------------------------------
'  1シートを走査して、コードが一番多い列と件数を返す
'------------------------------------------------------------------
Private Sub シート走査(ByVal ws As Worksheet, ByRef 最良列 As Long, ByRef 最良件数 As Long)
    最良列 = 0: 最良件数 = 0

    On Error Resume Next
    Dim lastR As Long: lastR = 0
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If lastR < 1 Then Exit Sub
    If lastR > 2000 Then lastR = 2000

    Dim v As Variant
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 12)).Value
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If Not IsArray(v) Then Exit Sub
    On Error GoTo 0

    Dim c As Long, r As Long, n As Long
    For c = 1 To UBound(v, 2)
        n = 0
        For r = 1 To UBound(v, 1)
            If コードか(v(r, c)) Then n = n + 1
        Next r
        If n > 最良件数 Then 最良件数 = n: 最良列 = c
    Next c
End Sub


'------------------------------------------------------------------
'  指定列からコードを集める(重複は捨てる)
'------------------------------------------------------------------
Private Function コード収集(ByVal ws As Worksheet, ByVal col As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set コード収集 = d

    On Error Resume Next
    Dim lastR As Long: lastR = 0
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR < 1 Then Exit Function
    If lastR > 20000 Then lastR = 20000

    Dim v As Variant
    v = ws.Range(ws.Cells(1, col), ws.Cells(lastR, col)).Value
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0

    Dim r As Long, s As String
    If IsArray(v) Then
        For r = 1 To UBound(v, 1)
            If コードか(v(r, 1)) Then
                s = Trim$(CStr(v(r, 1)))
                If Not d.Exists(s) Then d.Add s, 1
            End If
        Next r
    Else
        If コードか(v) Then
            s = Trim$(CStr(v))
            If Not d.Exists(s) Then d.Add s, 1
        End If
    End If
End Function


'------------------------------------------------------------------
'  列番号 → 列文字 (1→A)
'------------------------------------------------------------------
Private Function 列文字(ByVal col As Long) As String
    Dim s As String, n As Long
    n = col
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    列文字 = s
End Function


'------------------------------------------------------------------
'  銘柄コードらしいか (7203 / 130A など4桁)
'------------------------------------------------------------------
Private Function コードか(ByVal v As Variant) As Boolean
    コードか = False
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then Exit Function

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) <> 4 Then Exit Function
    If Left$(s, 1) = "0" Then Exit Function

    Dim i As Long, ch As String
    For i = 1 To 3
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ch = UCase$(Mid$(s, 4, 1))
    If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then コードか = True
End Function


'==================================================================
'  3. 監視開始 / 停止
'==================================================================
Public Sub 監視開始()
    If g監視中 Then
        MsgBox "すでに監視中です。", vbInformation, "監視"
        Exit Sub
    End If
    監視開始_内部
    MsgBox "監視を開始しました。" & vbCrLf & vbCrLf & _
           "・" & 間隔秒 & "秒ごとにRSSを見張ります" & vbCrLf & _
           "・設定シートの時刻になったら自動で記録します" & vbCrLf & _
           "・CSVは " & CSVフォルダ() & " に保存します", _
           vbInformation, "監視開始"
End Sub

Public Sub 監視開始_内部()
    g監視中 = True
    g前回出来高 = -1
    g出来高時刻 = Now
    日付切替チェック
    予約する
    ログ書く "情報", "監視を開始しました"
End Sub

Public Sub 監視停止()
    On Error Resume Next
    If g監視中 Then
        Application.OnTime g次回, "'" & ThisWorkbook.Name & "'!TICK処理", , False
    End If
    g監視中 = False
    ブザー停止
    ThisWorkbook.Worksheets(SH_BAN).Range("I1").Value = "(監視していません)"
    ThisWorkbook.Worksheets(SH_BAN).Range("I1").Interior.ColorIndex = xlNone
    ThisWorkbook.Worksheets(SH_BAN).Range("I1").Font.Color = RGB(0, 0, 0)
    ログ書く "情報", "監視を停止しました"
End Sub

Private Sub 予約する()
    On Error Resume Next
    g次回 = Now + TimeSerial(0, 0, 間隔秒)
    Application.OnTime g次回, "'" & ThisWorkbook.Name & "'!TICK処理"
End Sub


'==================================================================
'  4. タイマー本体 (30秒ごと)
'==================================================================
Public Sub TICK処理()
    If Not g監視中 Then Exit Sub
    予約する                       ' 何があっても次を先に予約する
    If g実行中 Then Exit Sub
    g実行中 = True

    On Error GoTo L_ERR
    日付切替チェック
    接続監視
    時刻チェック
    g実行中 = False
    Exit Sub
L_ERR:
    ログ書く "エラー", "TICK処理: " & Err.Description
    g実行中 = False
End Sub


'------------------------------------------------------------------
'  取得時刻が来ているか調べる
'------------------------------------------------------------------
Private Sub 時刻チェック()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    Dim 遅延 As Long: 遅延 = 数値設定("取得遅延秒", 5)

    Dim 最終 As Long
    最終 = cfg.Cells(cfg.Rows.Count, 1).End(xlUp).Row
    If 最終 < 2 Then Exit Sub

    Dim r As Long, t As Date, key As String, 予定 As Date
    For r = 2 To 最終
        If IsDate(cfg.Cells(r, 1).Value) Then
            t = CDate(cfg.Cells(r, 1).Value)
            key = Format(t, "hh:mm")
            If Not 済か(key) Then
                予定 = Date + TimeValue(key) + TimeSerial(0, 0, 遅延)
                If Now >= 予定 Then
                    If Now <= 予定 + TimeSerial(0, 猶予分, 0) Then
                        If 取得実行(key) Then 済にする key
                    Else
                        ログ書く "警告", key & " は " & 猶予分 & "分過ぎても取れませんでした(あきらめ)"
                        済にする key
                    End If
                End If
            End If
        End If
    Next r
End Sub


'==================================================================
'  5. 取得の実行
'==================================================================
Public Sub 今すぐ取得()
    Dim key As String
    key = "手動" & Format(Now, "hhmm")
    If 取得実行(key) Then
        MsgBox "取得しました。記録シートを見てください。", vbInformation, "今すぐ取得"
    Else
        MsgBox "RSSの値が取れませんでした。接続を確認してください。", vbExclamation, "今すぐ取得"
    End If
End Sub


Private Function 取得実行(ByVal 区分 As String) As Boolean
    取得実行 = False
    On Error GoTo L_ERR

    Dim ban As Worksheet, rec As Worksheet
    Set ban = ThisWorkbook.Worksheets(SH_BAN)
    Set rec = ThisWorkbook.Worksheets(SH_REC)

    ban.Calculate                     ' RSSを最新にする
    DoEvents

    Dim 最終 As Long
    最終 = ban.Cells(ban.Rows.Count, 1).End(xlUp).Row
    If 最終 < 2 Then
        ログ書く "警告", "銘柄が0件です。銘柄反映を実行してください"
        Exit Function
    End If

    Dim n As Long: n = 最終 - 1
    Dim buf() As Variant
    ReDim buf(1 To n, 1 To 10)

    Dim ts As String: ts = Format(Now, "hh:mm:ss")
    Dim i As Long, r As Long, 有効 As Long: 有効 = 0

    For r = 2 To 最終
        i = r - 1
        buf(i, 1) = Date
        buf(i, 2) = 区分
        buf(i, 3) = ts
        buf(i, 4) = CStr(ban.Cells(r, 1).Value)
        buf(i, 5) = CStr(ban.Cells(r, 2).Value)
        buf(i, 6) = 数値(ban.Cells(r, 3).Value)
        buf(i, 7) = 数値(ban.Cells(r, 4).Value)
        buf(i, 8) = 数値(ban.Cells(r, 5).Value)
        buf(i, 9) = 数値(ban.Cells(r, 6).Value)
        buf(i, 10) = 数値(ban.Cells(r, 7).Value)
        If buf(i, 9) > 0 Then 有効 = 有効 + 1
    Next r

    ' 半分も取れていなければ失敗。30秒後にまた挑戦する
    If 有効 * 2 < n Then
        ログ書く "再試行", 区分 & " 有効 " & 有効 & "/" & n & " → 30秒後に再挑戦"
        Exit Function
    End If

    Dim w As Long
    w = rec.Cells(rec.Rows.Count, 1).End(xlUp).Row + 1
    If w < 2 Then w = 2
    If w + n > 1000000 Then
        ログ書く "重大", "記録シートが満杯に近いです。古い分をCSVへ退避してください"
    End If
    rec.Cells(w, 1).Resize(n, 10).Value = buf

    CSV追記 buf                       ' ★ここが本命(落ちても残る)
    ログ書く "取得", 区分 & "  " & n & "銘柄 (値あり " & 有効 & ")"
    自動保存

    取得実行 = True
    Exit Function
L_ERR:
    ログ書く "エラー", "取得実行(" & 区分 & "): " & Err.Description
    取得実行 = False
End Function


'==================================================================
'  6. CSV追記 (排他ロックつき)
'==================================================================
Private Sub CSV追記(ByRef buf As Variant)
    Dim フォルダ As String: フォルダ = CSVフォルダ()
    If フォルダ = "" Then
        ログ書く "警告", "CSV保存フォルダが未設定です"
        Exit Sub
    End If

    Dim パス As String
    パス = フォルダ & "\OHLCV_" & Format(Date, "yyyymmdd") & ".csv"

    If Not ロック取得(フォルダ) Then
        ログ書く "警告", "他BOOKが使用中。CSV保存は次回にまわします"
        Exit Sub
    End If

    Dim ff As Integer, i As Long, 新規 As Boolean
    On Error GoTo L_ERR
    新規 = (Dir(パス) = "")
    ff = FreeFile
    Open パス For Append As #ff
    If 新規 Then
        Print #ff, "日付,時刻区分,取得時刻,コード,銘柄名,始値,高値,安値,現在値,出来高"
    End If
    For i = LBound(buf, 1) To UBound(buf, 1)
        Print #ff, Format(buf(i, 1), "yyyy/mm/dd") & "," & _
                   CSV安全(CStr(buf(i, 2))) & "," & CStr(buf(i, 3)) & "," & _
                   CSV安全(CStr(buf(i, 4))) & "," & CSV安全(CStr(buf(i, 5))) & "," & _
                   CStr(buf(i, 6)) & "," & CStr(buf(i, 7)) & "," & CStr(buf(i, 8)) & "," & _
                   CStr(buf(i, 9)) & "," & CStr(buf(i, 10))
    Next i
    Close #ff
    ロック解放 フォルダ
    Exit Sub
L_ERR:
    On Error Resume Next
    Close #ff
    ロック解放 フォルダ
    ログ書く "エラー", "CSV保存: " & Err.Description
End Sub


Private Function ロック取得(ByVal フォルダ As String) As Boolean
    Dim p As String: p = フォルダ & "\" & ロック名
    Dim i As Long, ff As Integer
    ロック取得 = False
    On Error Resume Next
    For i = 1 To 10                     ' 最大約5秒待つ
        If Dir(p) = "" Then
            ff = FreeFile
            Open p For Output As #ff
            Print #ff, Format(Now, "yyyy/mm/dd hh:nn:ss")
            Close #ff
            ロック取得 = True
            Exit Function
        End If
        ' 60秒より古いロックは壊れ物とみなして消す
        If FileDateTime(p) < Now - TimeSerial(0, 1, 0) Then Kill p
        待つ 0.5
    Next i
End Function

Private Sub ロック解放(ByVal フォルダ As String)
    On Error Resume Next
    Dim p As String: p = フォルダ & "\" & ロック名
    If Dir(p) <> "" Then Kill p
End Sub

Private Sub 待つ(ByVal 秒 As Double)
    Dim t As Double: t = Timer
    Do While Timer - t < 秒
        DoEvents
        If Timer < t Then Exit Do
    Loop
End Sub

Private Function CSV安全(ByVal s As String) As String
    s = Replace(s, ",", " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, """", "")
    CSV安全 = s
End Function


'==================================================================
'  7. RSS接続監視 + ブザー
'==================================================================
Private Sub 接続監視()
    Dim ban As Worksheet: Set ban = ThisWorkbook.Worksheets(SH_BAN)
    Dim c As Range: Set c = ban.Range("I1")

    Dim 最終 As Long
    最終 = ban.Cells(ban.Rows.Count, 1).End(xlUp).Row
    If 最終 < 2 Then
        c.Value = "銘柄が未設定です"
        c.Interior.Color = RGB(255, 242, 204)
        c.Font.Color = RGB(0, 0, 0)
        Exit Sub
    End If

    ban.Calculate
    Dim r As Long, 見る As Long, 有効 As Long, 合計 As Double
    見る = 最終
    If 見る > 21 Then 見る = 21              ' 先頭20銘柄で判定
    有効 = 0: 合計 = 0
    For r = 2 To 見る
        If 数値(ban.Cells(r, 6).Value) > 0 Then 有効 = 有効 + 1
        合計 = 合計 + 数値(ban.Cells(r, 7).Value)
    Next r

    Dim 切断 As Boolean, 理由 As String
    切断 = False: 理由 = ""

    If 有効 = 0 Then
        切断 = True: 理由 = "値がまったく取れていません"
    End If

    ' ザラバ中に出来高が5分間まったく動かない → RSS停止の疑い
    If Not 切断 Then
        If ザラバ中() Then
            If 合計 <> g前回出来高 Then
                g前回出来高 = 合計
                g出来高時刻 = Now
            ElseIf Now - g出来高時刻 > TimeSerial(0, 5, 0) Then
                切断 = True: 理由 = "出来高が5分間まったく更新されていません"
            End If
        Else
            g前回出来高 = 合計
            g出来高時刻 = Now
        End If
    End If

    If Not 切断 Then
        c.Value = "RSS 正常  " & Format(Now, "hh:mm:ss")
        c.Font.Color = RGB(0, 100, 0)
        c.Interior.Color = RGB(226, 242, 226)
        If gブザー中 Then
            ブザー停止
            ログ書く "復旧", "RSSが回復しました"
        End If
    Else
        c.Value = "★RSS切断★  " & Format(Now, "hh:mm:ss")
        c.Font.Color = RGB(255, 255, 255)
        c.Interior.Color = RGB(192, 0, 0)
        If 監視時間帯() Then
            If Not gブザー中 Then
                ログ書く "重大", "RSS切断を検知: " & 理由
                ブザー開始
            End If
        End If
    End If
End Sub


Private Sub ブザー開始()
    If 文字設定("ブザー使用", "○") <> "○" Then Exit Sub
    gブザー中 = True
    gブザー終了 = Now + TimeSerial(0, 数値設定("ブザー継続分", 10), 0)
    ブザー鳴動
End Sub

Public Sub ブザー鳴動()
    If Not gブザー中 Then Exit Sub
    If Now > gブザー終了 Then
        gブザー中 = False
        ログ書く "情報", "ブザーを自動停止しました(時間満了)"
        Exit Sub
    End If

    Beep
    待つ 0.3
    Beep
    待つ 0.3
    Beep

    Dim 間隔 As Long
    間隔 = 数値設定("ブザー間隔秒", 30)
    On Error Resume Next
    g次ブザー = Now + TimeSerial(0, 0, 間隔)
    Application.OnTime g次ブザー, "'" & ThisWorkbook.Name & "'!ブザー鳴動"
End Sub

Public Sub ブザー停止()
    On Error Resume Next
    If gブザー中 Then
        Application.OnTime g次ブザー, "'" & ThisWorkbook.Name & "'!ブザー鳴動", , False
    End If
    gブザー中 = False
End Sub

Public Sub RSS状態を見る()
    接続監視
    MsgBox ThisWorkbook.Worksheets(SH_BAN).Range("I1").Value, vbInformation, "RSS状態"
End Sub


'==================================================================
'  8. 取得済リスト (設定シート G:H列 / PCが落ちても残る)
'==================================================================
Private Sub 日付切替チェック()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    If CStr(cfg.Range("H1").Value) <> Format(Date, "yyyy/mm/dd") Then
        cfg.Range("H1").Value = Format(Date, "yyyy/mm/dd")
        cfg.Range("H2:H200").ClearContents
        ログ書く "情報", "日付が変わりました。取得済リストを初期化しました"
    End If
End Sub

Private Function 済か(ByVal key As String) As Boolean
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    Dim r As Long
    済か = False
    For r = 2 To 200
        If CStr(cfg.Cells(r, 8).Value) = key Then 済か = True: Exit Function
        If CStr(cfg.Cells(r, 8).Value) = "" Then Exit Function
    Next r
End Function

Private Sub 済にする(ByVal key As String)
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    Dim r As Long
    For r = 2 To 200
        If CStr(cfg.Cells(r, 8).Value) = "" Then
            cfg.Cells(r, 8).NumberFormat = "@"
            cfg.Cells(r, 8).Value = key
            Exit Sub
        End If
    Next r
End Sub


'==================================================================
'  9. ログ (シート + テキストファイル の二重書き)
'==================================================================
Private Sub ログ書く(ByVal 種別 As String, ByVal 内容 As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If Not ws Is Nothing Then
        Dim r As Long
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        If r > 500000 Then
            ws.Range("2:100000").Delete
            r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        End If
        ws.Cells(r, 1).Value = Now
        ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
        ws.Cells(r, 2).Value = 種別
        ws.Cells(r, 3).Value = 内容
        If 種別 = "重大" Or 種別 = "エラー" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Color = RGB(192, 0, 0)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
        End If
    End If
    ログファイル 種別, 内容
End Sub

Private Sub ログファイル(ByVal 種別 As String, ByVal 内容 As String)
    On Error Resume Next
    Dim フォルダ As String: フォルダ = CSVフォルダ()
    If フォルダ = "" Then Exit Sub
    Dim p As String
    p = フォルダ & "\OHLCV_LOG_" & Format(Date, "yyyymm") & ".txt"
    Dim ff As Integer: ff = FreeFile
    Open p For Append As #ff
    Print #ff, Format(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & 種別 & vbTab & 内容
    Close #ff
End Sub


'==================================================================
'  10. 小道具
'==================================================================
Private Function シート確保(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set シート確保 = ws
End Function

Private Sub 見出し(ByVal rg As Range)
    With rg
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Function 数値(ByVal v As Variant) As Double
    数値 = 0
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then 数値 = CDbl(v)
End Function

Private Function 設定値(ByVal 名前 As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    On Error GoTo 0
    設定値 = Empty
    If ws Is Nothing Then Exit Function
    Dim r As Long
    For r = 1 To 30
        If CStr(ws.Cells(r, 3).Value) = 名前 Then
            設定値 = ws.Cells(r, 4).Value
            Exit Function
        End If
    Next r
End Function

Private Function 数値設定(ByVal 名前 As String, ByVal 既定 As Long) As Long
    Dim v As Variant: v = 設定値(名前)
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then 数値設定 = CLng(v) Else 数値設定 = 既定
    Else
        数値設定 = 既定
    End If
End Function

Private Function 文字設定(ByVal 名前 As String, ByVal 既定 As String) As String
    Dim v As Variant: v = 設定値(名前)
    If IsEmpty(v) Then 文字設定 = 既定 Else 文字設定 = Trim$(CStr(v))
End Function

Private Function CSVフォルダ() As String
    Dim s As String
    s = 文字設定("CSV保存フォルダ", "")
    If s = "" Then s = ThisWorkbook.Path
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    On Error Resume Next
    If Dir(s, vbDirectory) = "" Then s = ThisWorkbook.Path
    CSVフォルダ = s
End Function

Private Function 監視時間帯() As Boolean
    Dim t As Date: t = TimeValue(Format(Now, "hh:nn:ss"))
    監視時間帯 = (t >= TimeSerial(9, 0, 0) And t <= TimeSerial(15, 35, 0))
    If Weekday(Date, vbMonday) > 5 Then 監視時間帯 = False
End Function

Private Function ザラバ中() As Boolean
    Dim t As Date: t = TimeValue(Format(Now, "hh:nn:ss"))
    ザラバ中 = (t >= TimeSerial(9, 5, 0) And t <= TimeSerial(11, 25, 0)) Or _
               (t >= TimeSerial(12, 35, 0) And t <= TimeSerial(15, 20, 0))
    If Weekday(Date, vbMonday) > 5 Then ザラバ中 = False
End Function

Private Sub 自動保存()
    If 文字設定("自動保存", "○") <> "○" Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub


'==================================================================
'  11. ボタンを設定シートに作る
'==================================================================
Private Sub ボタン作成()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CFG)
    Dim shp As Shape, i As Long
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0

    ボタン1つ ws, 1, "監視開始", "監視開始"
    ボタン1つ ws, 2, "監視停止", "監視停止"
    ボタン1つ ws, 3, "今すぐ取得", "今すぐ取得"
    ボタン1つ ws, 4, "ブザー停止", "ブザー停止"
    ボタン1つ ws, 5, "銘柄反映", "銘柄反映"
    ボタン1つ ws, 6, "RSS状態", "RSS状態を見る"
    ボタン1つ ws, 7, "銘柄取込", "銘柄取込_他BOOKから"
End Sub

Private Sub ボタン1つ(ByVal ws As Worksheet, ByVal n As Long, _
                      ByVal 表示 As String, ByVal マクロ As String)
    Dim b As Object
    Set b = ws.Shapes.AddFormControl(0, 820, 20 + (n - 1) * 34, 110, 28)
    b.TextFrame.Characters.Text = 表示
    b.OnAction = "'" & ThisWorkbook.Name & "'!" & マクロ
End Sub


'==================================================================
'  12. 起動時の自動監視 (ThisWorkbook から呼ばれる)
'==================================================================
Public Sub 自動起動()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    If ws Is Nothing Then Exit Sub
    If 文字設定("起動時に自動監視", "○") <> "○" Then Exit Sub
    If g監視中 Then Exit Sub
    監視開始_内部
End Sub
