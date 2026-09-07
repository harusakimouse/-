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
'   ※11:30 は前引け確定後の 11:31、15:30 は大引け確定後の 15:31 に取ります
'     (設定シートB列「実取得時刻」で調整できます)
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
    Dim 段階 As String

    On Error GoTo L_ERR
    Application.ScreenUpdating = False

    '--- 設定シート ---
    段階 = "設定シート"
    Set ws = シート確保(SH_CFG)
    If CStr(ws.Range("A1").Value) = "" Then 設定初期値 ws
    設定補完 ws
    取得時刻補修 ws
    実取得時刻補完 ws
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

    '--- 集積5シート(始値/高値/安値/終値/出来高) ---
    段階 = "集積5シート"
    Dim nm As Variant: nm = 集積シート名()
    Dim k As Long
    For k = LBound(nm) To UBound(nm)
        集積シート確保 CStr(nm(k))
    Next k

    '--- 時刻シート(9:00～15:30 の14枚) ---
    段階 = "時刻シート"
    Dim 報告 As String
    報告 = 時刻シート一括作成()

    段階 = "枠固定"
    枠固定
    段階 = "ボタン"
    ボタン作成
    ThisWorkbook.Worksheets(SH_CFG).Activate
    Application.ScreenUpdating = True
    On Error GoTo 0

    ログ書く "情報", "初期設定を実行しました (シート " & ThisWorkbook.Worksheets.Count & "枚)"

    MsgBox "初期設定が終わりました。" & vbCrLf & vbCrLf & _
           報告 & vbCrLf & _
           "このBOOKのシート合計 : " & ThisWorkbook.Worksheets.Count & " 枚" & vbCrLf & _
           "(作業5枚 + 値別5枚 + 時刻14枚 = 24枚が正常)" & vbCrLf & vbCrLf & _
           "1. 銘柄シートのA列に銘柄コードを入れる" & vbCrLf & _
           "2. 「銘柄反映」ボタンを押す" & vbCrLf & _
           "3. 「監視開始」ボタンを押す", vbInformation, "OHLCV取得BOOK"
    Exit Sub

L_ERR:
    Application.ScreenUpdating = True
    MsgBox "初期設定が [" & 段階 & "] で止まりました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & Err.Description, _
           vbExclamation, "初期設定"
End Sub


Private Sub 設定初期値(ByVal ws As Worksheet)
    Dim t As Variant
    t = Array("9:00", "9:10", "9:20", "9:45", "10:00", "10:15", "10:30", _
              "11:30", "12:30", "13:30", "14:30", "15:00", "15:20", "15:30")

    ws.Range("A1").Value = "取得時刻"
    見出し ws.Range("A1")
    ws.Range("A2:A60").NumberFormat = "hh:mm"
    Dim i As Long
    For i = LBound(t) To UBound(t)
        ws.Cells(i + 2, 1).Value = TimeValue(CStr(t(i)))
    Next i
    ws.Columns("A").ColumnWidth = 12

    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 46

    ws.Range("H1").Value = ""
    ws.Range("G1").Value = "取得済(本日)"
    見出し ws.Range("G1")
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
End Sub


Private Sub 実取得時刻補完(ByVal ws As Worksheet)
    ' B列 = 実際に取りに行く時刻。空ならA列と同じ。
    ' 15:30 は大引け確定を取りたいので 15:31 を既定にする。
    If CStr(ws.Range("B1").Value) = "" Then
        ws.Range("B1").Value = "実取得時刻"
        見出し ws.Range("B1")
        ws.Columns("B").ColumnWidth = 12
        ws.Range("B2:B60").NumberFormat = "hh:mm"
    End If

    Dim 最終 As Long
    最終 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, hm As String
    For r = 2 To 最終
        hm = 時刻文字(ws.Cells(r, 1).Value)
        If hm = "11:30" Or hm = "15:30" Then
            If 時刻文字(ws.Cells(r, 2).Value) = "" Then
                ' 前引け・大引けは確定してから取る(+1分)
                ws.Cells(r, 2).NumberFormat = "hh:mm"
                ws.Cells(r, 2).Value = TimeSerial(CLng(Left$(hm, 2)), 31, 0)
            End If
        End If
    Next r

    ws.Range("A1").Value = "取得時刻(表示)"
    ws.Range("F1").Value = "※B列=実際に取りに行く時刻。空ならA列と同じ"
    ws.Range("F2").Value = "※11:30は11:31 / 15:30は15:31(引け確定後に取る)"
    ws.Range("F1:F2").Font.Color = RGB(0, 0, 192)
    ws.Columns("F").ColumnWidth = 3
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
    If CStr(ws.Range("A2").Value) = "" Then ws.Columns("B:C").NumberFormat = "@"
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

    高速化開始
    On Error GoTo L_ERR

    If Not 既に開いていた Then
        進捗 "元BOOKを開いています… " & ファイル名
        Set wb = Workbooks.Open(Filename:=パス, ReadOnly:=True, UpdateLinks:=0)
    End If

    進捗 "銘柄コードの列を探しています…"
    '--- 全シートを見て、銘柄コードが多い列を候補として集める ---
    Dim 候補Ws(1 To 20) As Worksheet
    Dim 候補列(1 To 20) As Long
    Dim 候補数(1 To 20) As Long
    Dim 候補件数 As Long: 候補件数 = 0

    Dim ws As Worksheet, c As Long, n As Long
    For Each ws In wb.Worksheets
        c = 0: n = 0
        シート走査 ws, c, n
        If n >= 3 Then
            If 候補件数 < 20 Then
                候補件数 = 候補件数 + 1
                Set 候補Ws(候補件数) = ws
                候補列(候補件数) = c
                候補数(候補件数) = n
            End If
        End If
    Next ws

    If 候補件数 = 0 Then
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing
        画面戻す
        MsgBox "銘柄コードらしい列が見つかりませんでした。" & vbCrLf & _
               "(4桁のコードが3件以上ある列を探しています)", vbExclamation, "銘柄取込"
        Exit Sub
    End If

    ' 件数の多い順に並べ替え
    Dim a As Long, b As Long, tn As Long, tc As Long
    Dim tw As Worksheet
    For a = 1 To 候補件数 - 1
        For b = a + 1 To 候補件数
            If 候補数(b) > 候補数(a) Then
                tn = 候補数(a): 候補数(a) = 候補数(b): 候補数(b) = tn
                tc = 候補列(a): 候補列(a) = 候補列(b): 候補列(b) = tc
                Set tw = 候補Ws(a): Set 候補Ws(a) = 候補Ws(b): Set 候補Ws(b) = tw
            End If
        Next b
    Next a

    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' 候補が1つならそのまま、複数なら選んでもらう
    Dim 選択 As Long: 選択 = 1
    If 候補件数 > 1 Then
        Dim msg As String
        msg = "BOOK: " & ファイル名 & vbCrLf & _
              "銘柄コードの列が " & 候補件数 & " か所見つかりました。" & vbCrLf & vbCrLf
        Dim q As Long
        For q = 1 To 候補件数
            msg = msg & q & " : " & 候補Ws(q).Name & "  /  " & _
                  列文字(候補列(q)) & "列  /  " & 候補数(q) & "件" & vbCrLf
        Next q
        msg = msg & vbCrLf & "使う番号を入れてください (既定 1)"

        Dim ans As String
        ans = InputBox(msg, "銘柄取込 - どこから読みますか", "1")
        If ans = "" Then
            BOOK後始末 wb, 既に開いていた
            Set wb = Nothing
            画面戻す
            Exit Sub
        End If
        選択 = Val(ans)
        If 選択 < 1 Or 選択 > 候補件数 Then 選択 = 1
    End If

    Set 見つけたWs = 候補Ws(選択)
    見つけた列 = 候補列(選択)
    件数 = 候補数(選択)

    ' ★ BOOKを閉じる前に、必要な情報を全部こちらへ移す
    シート名 = 見つけたWs.Name
    列名 = 列文字(見つけた列)
    Set 集めた = コード収集(見つけたWs, 見つけた列)

    BOOK後始末 wb, 既に開いていた
    Set wb = Nothing
    Set 見つけたWs = Nothing

    画面戻す
    On Error GoTo 0

    '--- ここから先は自分のBOOKだけ ---
    If 集めた.Count = 0 Then
        MsgBox "コードを1件も読み取れませんでした。", vbExclamation, "銘柄取込"
        Exit Sub
    End If

    If MsgBox("読み込みました。" & vbCrLf & vbCrLf & _
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
    画面戻す
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

    Dim r As Long, key As String, 実文 As String, 予定 As Date
    For r = 2 To 最終
        key = 時刻文字(cfg.Cells(r, 1).Value)
        If key <> "" Then

            ' B列に実取得時刻があればそちらを使う(15:30 → 15:31 など)
            実文 = 時刻文字(cfg.Cells(r, 2).Value)
            If 実文 = "" Then 実文 = key

            If Not 済か(key) Then
                予定 = Date + TimeValue(実文) + TimeSerial(0, 0, 遅延)
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
    集積へ書く buf, 区分                ' 5シートへ右へずらして集積
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
'------------------------------------------------------------------
'  重い処理の前後で使う (他BOOKを開く時は計算を必ず止める)
'------------------------------------------------------------------
Private Sub 高速化開始()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub 画面戻す()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.AskToUpdateLinks = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Private Sub 進捗(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = "OHLCV取得BOOK : " & msg
    DoEvents
End Sub


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
    監視時間帯 = (t >= TimeSerial(9, 0, 0) And t <= TimeSerial(15, 40, 0))
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
    ボタン1つ ws, 8, "集積作り直し", "集積を作り直す"
    ボタン1つ ws, 9, "時刻シート作成", "時刻シートを作る"
    ボタン1つ ws, 10, "過去データ取込", "過去データ取込"
End Sub

Private Sub ボタン1つ(ByVal ws As Worksheet, ByVal n As Long, _
                      ByVal 表示 As String, ByVal マクロ As String)
    Dim b As Object
    Set b = ws.Shapes.AddFormControl(0, 820, 20 + (n - 1) * 34, 110, 28)
    b.TextFrame.Characters.Text = 表示
    b.OnAction = "'" & ThisWorkbook.Name & "'!" & マクロ
End Sub


'==================================================================
'  11-B. 値別5シート (行=銘柄 / 列=日付+時刻 で右へ伸びる)
'
'   1行目 = 日付
'   2行目 = 時刻
'   3行目 = 見出し (A3=コード  B3=銘柄名)
'   4行目～ = 銘柄
'   C列から右へ、1日14列ずつ増えます
'
'   ※ここは「見るため」の表です。大元は 記録シート と CSV。
'     おかしくなったら「集積作り直し」ボタンで記録シートから作り直せます。
'==================================================================
Private Function 集積シート名() As Variant
    集積シート名 = Array("始値", "高値", "安値", "終値", "出来高")
End Function

Private Function 集積元列() As Variant
    集積元列 = Array(6, 7, 8, 9, 10)      ' buf の 始値/高値/安値/現在値/出来高
End Function


Private Function 集積シート確保(ByVal nm As String) As Worksheet
    Dim ws As Worksheet: Set ws = シート確保(nm)
    If CStr(ws.Range("A3").Value) <> "コード" Then
        ws.Range("A1").Value = "日付"
        ws.Range("A2").Value = "時刻"
        ws.Range("A3").Value = "コード"
        ws.Range("B3").Value = "銘柄名"
        見出し ws.Range("A3:B3")
        ws.Range("A1:B2").Font.Bold = True
        ws.Columns("A").NumberFormat = "@"
        ws.Columns("A").ColumnWidth = 9
        ws.Columns("B").ColumnWidth = 20
    End If
    Set 集積シート確保 = ws
End Function


Private Sub 枠固定()
    On Error Resume Next
    Dim nm As Variant: nm = 集積シート名()
    Dim k As Long
    For k = LBound(nm) To UBound(nm)
        枠固定1枚 CStr(nm(k))
    Next k
    Dim tm As Variant: tm = 取得時刻一覧()
    For k = LBound(tm) To UBound(tm)
        枠固定1枚 時刻シート名(CStr(tm(k)))
    Next k
End Sub

Private Sub 枠固定1枚(ByVal nm As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nm)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C4").Select
    ActiveWindow.FreezePanes = True
End Sub


'------------------------------------------------------------------
'  1回分のデータを5シートへ書く
'------------------------------------------------------------------
Private Sub 集積へ書く(ByRef buf As Variant, ByVal 区分 As String)
    On Error GoTo L_ERR
    If Left$(区分, 2) = "手動" Then Exit Sub      ' 手動取得は集積に入れない

    Dim nm As Variant: nm = 集積シート名()
    Dim cl As Variant: cl = 集積元列()
    Dim k As Long
    For k = LBound(nm) To UBound(nm)
        集積1シート CStr(nm(k)), CLng(cl(k)), buf, 区分
    Next k

    時刻シートへ書く buf, 区分
    Exit Sub
L_ERR:
    ログ書く "エラー", "集積書込(" & 区分 & "): " & Err.Description
End Sub


Private Sub 集積1シート(ByVal shName As String, ByVal 元列 As Long, _
                        ByRef buf As Variant, ByVal 区分 As String)
    Dim ws As Worksheet: Set ws = 集積シート確保(shName)

    Dim d As Date: d = CDate(buf(LBound(buf, 1), 1))
    Dim t As Date: t = TimeValue(区分)

    Dim col As Long: col = 集積列(ws, d, t)
    If col = 0 Then Exit Sub

    Dim map As Object: Set map = 集積行マップ(ws)

    Dim 最終行 As Long
    最終行 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If 最終行 < 3 Then 最終行 = 3

    ' 新しい銘柄は下に足す
    Dim i As Long, code As String
    For i = LBound(buf, 1) To UBound(buf, 1)
        code = Trim$(CStr(buf(i, 4)))
        If code <> "" Then
            If Not map.Exists(code) Then
                最終行 = 最終行 + 1
                ws.Cells(最終行, 1).NumberFormat = "@"
                ws.Cells(最終行, 1).Value = code
                ws.Cells(最終行, 2).Value = CStr(buf(i, 5))
                map.Add code, 最終行
            Else
                ' 銘柄名が空なら埋める
                If CStr(ws.Cells(CLng(map(code)), 2).Value) = "" Then
                    ws.Cells(CLng(map(code)), 2).Value = CStr(buf(i, 5))
                End If
            End If
        End If
    Next i

    If 最終行 < 4 Then Exit Sub

    ' 値を一括で書く
    Dim vals() As Variant
    ReDim vals(1 To 最終行 - 3, 1 To 1)
    Dim r As Long
    For i = LBound(buf, 1) To UBound(buf, 1)
        code = Trim$(CStr(buf(i, 4)))
        If code <> "" Then
            If map.Exists(code) Then
                r = CLng(map(code)) - 3
                If r >= 1 And r <= UBound(vals, 1) Then vals(r, 1) = buf(i, 元列)
            End If
        End If
    Next i
    ws.Cells(4, col).Resize(最終行 - 3, 1).Value = vals
End Sub


'------------------------------------------------------------------
'  その日付+時刻の列を返す (無ければ右端に作る)
'------------------------------------------------------------------
Private Function 集積列(ByVal ws As Worksheet, ByVal d As Date, ByVal t As Date) As Long
    集積列 = 0
    Dim 最終列 As Long
    最終列 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If 最終列 < 2 Then 最終列 = 2

    Dim tt As String: tt = Format(t, "hh:mm")

    If 最終列 >= 3 Then
        Dim h As Variant
        h = ws.Range(ws.Cells(1, 1), ws.Cells(2, 最終列)).Value
        Dim c As Long
        For c = 3 To UBound(h, 2)
            If IsDate(h(1, c)) And IsDate(h(2, c)) Then
                If CDate(h(1, c)) = d And Format(h(2, c), "hh:mm") = tt Then
                    集積列 = c
                    Exit Function
                End If
            End If
        Next c
    End If

    Dim nc As Long: nc = 最終列 + 1
    If nc > 16300 Then
        ログ書く "重大", ws.Name & " が列の上限です。古い分を別ファイルへ退避してください"
        Exit Function
    End If
    If nc > 15000 Then
        ログ書く "警告", ws.Name & " の列が " & nc & " です。そろそろ退避してください"
    End If

    ws.Cells(1, nc).Value = d
    ws.Cells(1, nc).NumberFormat = "mm/dd"
    ws.Cells(2, nc).Value = t
    ws.Cells(2, nc).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(1, nc), ws.Cells(2, nc)).Font.Bold = True
    ws.Range(ws.Cells(1, nc), ws.Cells(2, nc)).HorizontalAlignment = xlCenter
    ws.Columns(nc).ColumnWidth = 9
    ws.Columns(nc).NumberFormat = "#,##0.##"
    集積列 = nc
End Function


'------------------------------------------------------------------
'  コード → 行番号 の対応表
'------------------------------------------------------------------
Private Function 集積行マップ(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set 集積行マップ = d

    Dim 最終行 As Long
    最終行 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If 最終行 < 4 Then Exit Function

    Dim v As Variant
    v = ws.Range(ws.Cells(4, 1), ws.Cells(最終行, 1)).Value
    Dim i As Long, sc As String
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            sc = Trim$(CStr(v(i, 1)))
            If sc <> "" Then
                If Not d.Exists(sc) Then d.Add sc, i + 3
            End If
        Next i
    Else
        sc = Trim$(CStr(v))
        If sc <> "" Then d.Add sc, 4
    End If
End Function


'==================================================================
'  10-B. 過去データ取込 (既存BOOKのOHLCV履歴を丸ごと持ってくる)
'
'   売買BOOKなどの 始値/高値/安値/終値/出来高 シート
'   (行=銘柄 / 列=日付) を読んで、このBOOKの記録シートに
'   「15:30 の記録」として書き写します。
'   そのあと集積19シートを作り直せば、全部そろいます。
'==================================================================
Public Sub 過去データ取込()
    Dim fn As Variant
    fn = Application.GetOpenFilename("Excelブック,*.xls*", , _
         "OHLCV履歴のあるBOOKを選んでください")
    If VarType(fn) = vbBoolean Then Exit Sub

    Dim パス As String: パス = CStr(fn)
    Dim ファイル名 As String
    ファイル名 = Mid$(パス, InStrRev(パス, "\") + 1)

    Dim wb As Workbook
    Dim 既に開いていた As Boolean
    On Error Resume Next
    Set wb = Workbooks(ファイル名)
    On Error GoTo 0
    既に開いていた = Not (wb Is Nothing)

    Dim 出力() As Variant
    Dim 件数 As Long: 件数 = 0

    If Not 既に開いていた Then
        If MsgBox("元BOOKを『計算を止めた状態』で開きます。" & vbCrLf & vbCrLf & _
                  "大きいBOOKだと開くのに1～2分かかることがあります。" & vbCrLf & _
                  "画面が止まって見えても、そのまま待ってください。" & vbCrLf & vbCrLf & _
                  "始めますか?", vbYesNo + vbQuestion, "過去データ取込") <> vbYes Then Exit Sub
    End If

    高速化開始
    On Error GoTo L_ERR

    If Not 既に開いていた Then
        進捗 "元BOOKを開いています… " & ファイル名
        Set wb = Workbooks.Open(Filename:=パス, ReadOnly:=True, UpdateLinks:=0)
    End If

    '--- 5シートがあるか確認 ---
    進捗 "シートを調べています…"
    Dim shs As Variant: shs = 集積シート名()
    Dim i As Long, ws As Worksheet, 無し As String
    無し = ""
    For i = LBound(shs) To UBound(shs)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(shs(i)))
        On Error GoTo L_ERR
        If ws Is Nothing Then 無し = 無し & shs(i) & " "
    Next i
    If 無し <> "" Then
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing
        画面戻す
        MsgBox "このBOOKには次のシートがありません。" & vbCrLf & vbCrLf & 無し & vbCrLf & vbCrLf & _
               "始値/高値/安値/終値/出来高 の5シートが必要です。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    '--- レイアウトを自動判定 (終値シートが基準) ---
    Dim base As Worksheet: Set base = wb.Worksheets("終値")
    Dim ヘッダ As Long: ヘッダ = 日付ヘッダ行(base)
    If ヘッダ = 0 Then
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing
        画面戻す
        MsgBox "日付の見出し行が見つかりませんでした。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    Dim 最終行 As Long, 最終列 As Long
    最終行 = base.Cells(base.Rows.Count, 1).End(xlUp).Row
    最終列 = base.Cells(ヘッダ, base.Columns.Count).End(xlToLeft).Column
    If 最終行 <= ヘッダ Or 最終列 < 3 Then
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing
        画面戻す
        MsgBox "履歴データが見つかりませんでした。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    '--- 大きすぎないか確認 ---
    If (最終行 - ヘッダ) * (最終列 - 2) > 2000000 Then
        BOOK後始末 wb, 既に開いていた
        Set wb = Nothing
        画面戻す
        MsgBox "データが大きすぎます (" & (最終行 - ヘッダ) & "行 × " & (最終列 - 2) & "列)。" & vbCrLf & _
               "元BOOKの範囲を絞ってからやり直してください。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    '--- 5シートを配列で読む ---
    Dim arr(0 To 4) As Variant
    For i = 0 To 4
        進捗 "読み込み中… " & shs(i)
        Set ws = wb.Worksheets(CStr(shs(i)))
        arr(i) = ws.Range(ws.Cells(ヘッダ, 1), ws.Cells(最終行, 最終列)).Value
    Next i

    進捗 "元BOOKを閉じています…"
    BOOK後始末 wb, 既に開いていた
    Set wb = Nothing
    Set base = Nothing
    Set ws = Nothing

    '--- 日付の列をひろう ---
    Dim 日列() As Long, 日付() As Date, 日数 As Long
    ReDim 日列(1 To 最終列): ReDim 日付(1 To 最終列)
    日数 = 0
    Dim c As Long, hv As Variant
    For c = 3 To 最終列
        hv = arr(3)(1, c)
        If Not IsError(hv) Then
            If IsDate(hv) Then
                If CDate(hv) >= DateSerial(1990, 1, 1) Then
                    日数 = 日数 + 1
                    日列(日数) = c
                    日付(日数) = CDate(hv)
                End If
            End If
        End If
    Next c

    If 日数 = 0 Then
        画面戻す
        MsgBox "日付の列が見つかりませんでした。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    '--- 銘柄行をひろって出力を組み立てる ---
    進捗 "組み立て中… " & 日数 & "日分"
    Dim 行数 As Long: 行数 = UBound(arr(3), 1)
    ReDim 出力(1 To 行数 * 日数, 1 To 10)

    Dim r As Long, j As Long, k As Long
    Dim code As String, 銘柄 As String
    For r = 2 To 行数
        code = Trim$(CStr(arr(3)(r, 1)))
        If コードか(code) Then
            銘柄 = Trim$(CStr(arr(3)(r, 2)))
            For j = 1 To 日数
                c = 日列(j)
                Dim v終 As Double: v終 = 数値(arr(3)(r, c))
                If v終 > 0 Then
                    件数 = 件数 + 1
                    出力(件数, 1) = 日付(j)
                    出力(件数, 2) = "15:30"
                    出力(件数, 3) = "(履歴)"
                    出力(件数, 4) = code
                    出力(件数, 5) = 銘柄
                    For k = 0 To 4
                        出力(件数, 6 + k) = 履歴値(arr(k), r, c, code)
                    Next k
                End If
            Next j
        End If
    Next r

    画面戻す

    If 件数 = 0 Then
        MsgBox "書き写せるデータがありませんでした。", vbExclamation, "過去データ取込"
        Exit Sub
    End If

    If MsgBox("読み取りました。" & vbCrLf & vbCrLf & _
              "BOOK  : " & ファイル名 & vbCrLf & _
              "日数  : " & 日数 & " 日分" & vbCrLf & _
              "行数  : " & 件数 & " 行" & vbCrLf & vbCrLf & _
              "記録シートに「15:30 の記録」として書き足しますか?" & vbCrLf & _
              "※今の記録シートは消しません。足すだけです", _
              vbYesNo + vbQuestion, "過去データ取込") <> vbYes Then Exit Sub

    '--- 記録シートへ書き足す ---
    高速化開始
    進捗 "記録シートに書き足しています…"
    Dim rec As Worksheet: Set rec = ThisWorkbook.Worksheets(SH_REC)
    Dim w As Long
    w = rec.Cells(rec.Rows.Count, 1).End(xlUp).Row + 1
    If w < 2 Then w = 2
    rec.Cells(w, 1).Resize(件数, 10).Value = 出力

    ' 日付・時刻の順に並べ替える
    Dim 末 As Long: 末 = w + 件数 - 1
    On Error Resume Next
    rec.Range("A1:J" & 末).Sort _
        Key1:=rec.Range("A2"), Order1:=xlAscending, _
        Key2:=rec.Range("B2"), Order2:=xlAscending, _
        Header:=xlYes
    On Error GoTo 0
    画面戻す

    ログ書く "情報", "過去データ取込 " & 件数 & "行 / " & 日数 & "日分 (" & ファイル名 & ")"

    If MsgBox(件数 & " 行を記録シートに書き足しました。" & vbCrLf & vbCrLf & _
              "続けて「集積作り直し」を実行しますか?" & vbCrLf & _
              "(19シートに全部並べ直します。少し時間がかかります)", _
              vbYesNo + vbQuestion, "過去データ取込") = vbYes Then
        集積を作り直す
    End If
    Exit Sub

L_ERR:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    BOOK後始末 wb, 既に開いていた
    Set wb = Nothing
    画面戻す
    ログ書く "エラー", "過去データ取込: " & eN & " " & eD
    MsgBox "取込に失敗しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & eN & vbCrLf & eD, vbExclamation, "過去データ取込"
End Sub


'------------------------------------------------------------------
'  日付見出しの行を探す (1～10行目で日付が3つ以上ある行)
'------------------------------------------------------------------
Private Function 日付ヘッダ行(ByVal ws As Worksheet) As Long
    日付ヘッダ行 = 0
    Dim r As Long, c As Long, n As Long
    Dim v As Variant
    For r = 1 To 10
        n = 0
        For c = 3 To 40
            v = ws.Cells(r, c).Value
            If Not IsError(v) Then
                If IsDate(v) Then
                    If CDate(v) >= DateSerial(1990, 1, 1) Then n = n + 1
                End If
            End If
        Next c
        If n >= 3 Then
            日付ヘッダ行 = r
            Exit Function
        End If
    Next r
End Function


'------------------------------------------------------------------
'  同じ行のコードが一致する時だけ値を返す (行ズレ対策)
'------------------------------------------------------------------
Private Function 履歴値(ByRef a As Variant, ByVal r As Long, _
                        ByVal c As Long, ByVal code As String) As Double
    履歴値 = 0
    On Error Resume Next
    If r > UBound(a, 1) Or c > UBound(a, 2) Then Exit Function
    If Trim$(CStr(a(r, 1))) <> code Then Exit Function
    履歴値 = 数値(a(r, c))
End Function


'==================================================================
'  11-C. 時刻シート (9:00～15:30 の14枚)
'
'   1行目 = 日付 (5列で1日分)
'   2行目 = 始値 高値 安値 終値 出来高
'   3行目 = 見出し (A3=コード  B3=銘柄名)
'   4行目～ = 銘柄
'   C列から右へ、1日5列ずつ増えます
'==================================================================
Private Function 取得時刻一覧() As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    Dim 最終 As Long
    最終 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim tmp() As String
    ReDim tmp(0 To 100)
    Dim r As Long, n As Long: n = -1
    Dim hm As String
    For r = 2 To 最終
        hm = 時刻文字(ws.Cells(r, 1).Value)
        If hm <> "" Then
            If n >= 99 Then Exit For
            n = n + 1
            tmp(n) = hm
        End If
    Next r
    If n < 0 Then
        取得時刻一覧 = 既定時刻()          ' 設定シートが空でも必ず14本返す
    Else
        ReDim Preserve tmp(0 To n)
        取得時刻一覧 = tmp
    End If
End Function


'------------------------------------------------------------------
'  セルの中身を "hh:mm" にする (時刻でなければ "")
'  ※ 時刻セルは Date で返る時と 数値(0～1)で返る時があるので両対応
'------------------------------------------------------------------
Private Function 時刻文字(ByVal v As Variant) As String
    時刻文字 = ""
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then
        時刻文字 = Format(CDate(v), "hh:mm")
        Exit Function
    End If

    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        If d >= 0 And d < 1 Then 時刻文字 = Format(d, "hh:mm")
        Exit Function
    End If

    Dim t As String: t = Trim$(CStr(v))
    If t = "" Then Exit Function
    If IsDate(t) Then 時刻文字 = Format(CDate(t), "hh:mm")
End Function


Private Function 既定時刻() As Variant
    既定時刻 = Array("09:00", "09:10", "09:20", "09:45", "10:00", "10:15", "10:30", _
                     "11:30", "12:30", "13:30", "14:30", "15:00", "15:20", "15:30")
End Function


'------------------------------------------------------------------
'  設定シートA列に時刻が1つも無ければ入れ直す
'------------------------------------------------------------------
Private Sub 取得時刻補修(ByVal ws As Worksheet)
    Dim 最終 As Long
    最終 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, 有 As Long: 有 = 0
    For r = 2 To 最終
        If 時刻文字(ws.Cells(r, 1).Value) <> "" Then 有 = 有 + 1
    Next r
    If 有 > 0 Then Exit Sub

    ws.Range("A2:A60").NumberFormat = "hh:mm"
    Dim t As Variant: t = 既定時刻()
    Dim i As Long
    For i = LBound(t) To UBound(t)
        ws.Cells(i + 2, 1).Value = TimeValue(CStr(t(i)))
    Next i
    ws.Columns("A").ColumnWidth = 14
End Sub


'------------------------------------------------------------------
'  時刻シートを作る (単独でも実行できます)
'------------------------------------------------------------------
Public Sub 時刻シートを作る()
    Dim 結果 As String
    結果 = 時刻シート一括作成()
    MsgBox 結果, vbInformation, "時刻シート"
End Sub


Private Function 時刻シート一括作成() As String
    Dim cfg As Worksheet
    Set cfg = シート確保(SH_CFG)
    取得時刻補修 cfg

    Dim tm As Variant: tm = 取得時刻一覧()
    If UBound(tm) < LBound(tm) Then
        時刻シート一括作成 = "取得時刻が1つもありません。"
        Exit Function
    End If

    Dim 作成 As Long, 既存 As Long, 失敗 As String
    作成 = 0: 既存 = 0: 失敗 = ""

    Dim k As Long, nm As String, ws As Worksheet
    For k = LBound(tm) To UBound(tm)
        nm = 時刻シート名(CStr(tm(k)))

        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(nm)
        On Error GoTo 0

        If ws Is Nothing Then
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets.Add( _
                     After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            If Err.Number <> 0 Then
                失敗 = 失敗 & nm & "(追加不可) "
                Err.Clear
                On Error GoTo 0
                GoTo 次へ
            End If
            ws.Name = nm
            If Err.Number <> 0 Then
                Err.Clear
                ws.Name = "T" & nm              ' 名前が使えない時の逃げ道
                If Err.Number <> 0 Then
                    失敗 = 失敗 & nm & "(名前不可) "
                    Err.Clear
                End If
            End If
            On Error GoTo 0
            作成 = 作成 + 1
        Else
            既存 = 既存 + 1
        End If

        時刻シート見出し ws, CStr(tm(k))
次へ:
    Next k

    Dim msg As String
    msg = "時刻シート " & (UBound(tm) - LBound(tm) + 1) & " 本ぶん" & vbCrLf & _
          "  新しく作った : " & 作成 & " 枚" & vbCrLf & _
          "  すでにあった : " & 既存 & " 枚" & vbCrLf
    If 失敗 <> "" Then msg = msg & "  作れなかった : " & 失敗 & vbCrLf
    時刻シート一括作成 = msg
End Function


Private Function 時刻シート名(ByVal 区分 As String) As String
    時刻シート名 = Replace(区分, ":", "")
End Function


Private Function 時刻シート確保(ByVal 区分 As String) As Worksheet
    Dim ws As Worksheet
    Set ws = シート確保(時刻シート名(区分))
    時刻シート見出し ws, 区分
    Set 時刻シート確保 = ws
End Function


Private Sub 時刻シート見出し(ByVal ws As Worksheet, ByVal 区分 As String)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If CStr(ws.Range("A3").Value) <> "コード" Then
        ws.Range("A1").Value = 区分 & " の記録"
        ws.Range("A2").Value = "項目"
        ws.Range("A3").Value = "コード"
        ws.Range("B3").Value = "銘柄名"
        見出し ws.Range("A3:B3")
        ws.Range("A1:B2").Font.Bold = True
        ws.Columns("A").NumberFormat = "@"
        ws.Columns("A").ColumnWidth = 9
        ws.Columns("B").ColumnWidth = 20
    End If
End Sub


Private Sub 時刻シートへ書く(ByRef buf As Variant, ByVal 区分 As String)
    On Error GoTo L_ERR
    Dim ws As Worksheet: Set ws = 時刻シート確保(区分)
    Dim d As Date: d = CDate(buf(LBound(buf, 1), 1))

    Dim col As Long: col = 時刻シート列(ws, d)
    If col = 0 Then Exit Sub

    Dim map As Object: Set map = 集積行マップ(ws)
    Dim 最終行 As Long
    最終行 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If 最終行 < 3 Then 最終行 = 3

    Dim i As Long, code As String
    For i = LBound(buf, 1) To UBound(buf, 1)
        code = Trim$(CStr(buf(i, 4)))
        If code <> "" Then
            If Not map.Exists(code) Then
                最終行 = 最終行 + 1
                ws.Cells(最終行, 1).NumberFormat = "@"
                ws.Cells(最終行, 1).Value = code
                ws.Cells(最終行, 2).Value = CStr(buf(i, 5))
                map.Add code, 最終行
            ElseIf CStr(ws.Cells(CLng(map(code)), 2).Value) = "" Then
                ws.Cells(CLng(map(code)), 2).Value = CStr(buf(i, 5))
            End If
        End If
    Next i
    If 最終行 < 4 Then Exit Sub

    Dim vals() As Variant
    ReDim vals(1 To 最終行 - 3, 1 To 5)
    Dim r As Long, j As Long
    For i = LBound(buf, 1) To UBound(buf, 1)
        code = Trim$(CStr(buf(i, 4)))
        If code <> "" Then
            If map.Exists(code) Then
                r = CLng(map(code)) - 3
                If r >= 1 And r <= UBound(vals, 1) Then
                    For j = 1 To 5
                        vals(r, j) = buf(i, 5 + j)
                    Next j
                End If
            End If
        End If
    Next i
    ws.Cells(4, col).Resize(最終行 - 3, 5).Value = vals
    Exit Sub
L_ERR:
    ログ書く "エラー", "時刻シート書込(" & 区分 & "): " & Err.Description
End Sub


'------------------------------------------------------------------
'  その日の5列ブロックの先頭列を返す (無ければ右端に作る)
'------------------------------------------------------------------
Private Function 時刻シート列(ByVal ws As Worksheet, ByVal d As Date) As Long
    時刻シート列 = 0
    Dim 最終列 As Long
    最終列 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If 最終列 < 2 Then 最終列 = 2

    If 最終列 >= 3 Then
        Dim h As Variant
        h = ws.Range(ws.Cells(1, 1), ws.Cells(2, 最終列)).Value
        Dim c As Long
        For c = 3 To UBound(h, 2)
            If IsDate(h(1, c)) Then
                If CDate(h(1, c)) = d And CStr(h(2, c)) = "始値" Then
                    時刻シート列 = c
                    Exit Function
                End If
            End If
        Next c
    End If

    Dim nc As Long: nc = 最終列 + 1
    If nc + 4 > 16300 Then
        ログ書く "重大", ws.Name & " が列の上限です。古い分を別ファイルへ退避してください"
        Exit Function
    End If
    If nc > 15000 Then
        ログ書く "警告", ws.Name & " の列が " & nc & " です。そろそろ退避してください"
    End If

    Dim lbl As Variant
    lbl = Array("始値", "高値", "安値", "終値", "出来高")
    Dim j As Long
    For j = 0 To 4
        ws.Cells(1, nc + j).Value = d
        ws.Cells(1, nc + j).NumberFormat = "mm/dd"
        ws.Cells(2, nc + j).Value = lbl(j)
        ws.Columns(nc + j).ColumnWidth = 9
        ws.Columns(nc + j).NumberFormat = "#,##0.##"
    Next j
    With ws.Range(ws.Cells(1, nc), ws.Cells(2, nc + 4))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Range(ws.Cells(1, nc), ws.Cells(3, nc)).Borders(xlEdgeLeft).Weight = xlMedium
    時刻シート列 = nc
End Function


'------------------------------------------------------------------
'  記録シートから5シートを作り直す (復旧用)
'------------------------------------------------------------------
Public Sub 集積を作り直す()
    Dim rec As Worksheet: Set rec = ThisWorkbook.Worksheets(SH_REC)
    Dim 最終行 As Long
    最終行 = rec.Cells(rec.Rows.Count, 1).End(xlUp).Row
    If 最終行 < 2 Then
        MsgBox "記録シートにデータがありません。", vbExclamation, "集積の作り直し"
        Exit Sub
    End If
    If 最終行 > 500000 Then
        MsgBox "記録シートが " & 最終行 & " 行あります。" & vbCrLf & _
               "多すぎるので、古い分を別ファイルへ退避してから実行してください。", _
               vbExclamation, "集積の作り直し"
        Exit Sub
    End If

    If MsgBox("記録シート " & (最終行 - 1) & " 行から" & vbCrLf & _
              "集積19シート(OHLCV 5枚 + 時刻別 14枚)を作り直します。" & vbCrLf & vbCrLf & _
              "※今の19シートの中身は消えて、記録シートの内容で作り直されます" & vbCrLf & _
              "※記録シートとCSVは触りません" & vbCrLf & vbCrLf & _
              "よろしいですか?", vbYesNo + vbExclamation, "集積の作り直し") <> vbYes Then Exit Sub

    On Error GoTo L_ERR
    高速化開始

    Dim nm As Variant: nm = 集積シート名()
    Dim k As Long, ws As Worksheet
    For k = LBound(nm) To UBound(nm)
        Set ws = シート確保(CStr(nm(k)))
        ws.Cells.Clear
        集積シート確保 CStr(nm(k))
    Next k

    Dim tm As Variant: tm = 取得時刻一覧()
    For k = LBound(tm) To UBound(tm)
        Set ws = シート確保(時刻シート名(CStr(tm(k))))
        ws.Cells.Clear
        時刻シート確保 CStr(tm(k))
    Next k

    Dim v As Variant
    v = rec.Range(rec.Cells(2, 1), rec.Cells(最終行, 10)).Value

    Dim i As Long, st As Long, 件 As Long, 読 As Long
    件 = 0: 読 = 0
    Dim キー As String, 前キー As String
    前キー = ""
    st = 1
    For i = 1 To UBound(v, 1) + 1
        If i <= UBound(v, 1) Then
            キー = CStr(v(i, 1)) & "|" & CStr(v(i, 2))
        Else
            キー = "**終わり**"
        End If
        If キー <> 前キー Then
            If 前キー <> "" And i > st Then
                読 = 読 + 1
                If ブロック書く(v, st, i - 1) Then 件 = 件 + 1
                If 読 Mod 5 = 0 Then 進捗 "並べ直し中… " & 読 & " 回分"
            End If
            st = i
            前キー = キー
        End If
    Next i

    枠固定
    ThisWorkbook.Worksheets(SH_CFG).Activate
    画面戻す

    ログ書く "情報", "集積19シートを作り直しました (" & 件 & "/" & 読 & "回分)"
    If 件 = 0 Then
        MsgBox "並べ直せませんでした。" & vbCrLf & vbCrLf & _
               "読んだ回数: " & 読 & vbCrLf & _
               "書けた回数: 0" & vbCrLf & vbCrLf & _
               "記録シートB列(時刻区分)の中身を確認してください。", _
               vbExclamation, "集積の作り直し"
    Else
        MsgBox "作り直しました。" & vbCrLf & vbCrLf & _
               "読んだ回数: " & 読 & vbCrLf & _
               "並べた回数: " & 件, vbInformation, "集積の作り直し"
    End If
    Exit Sub

L_ERR:
    画面戻す
    ログ書く "エラー", "集積作り直し: " & Err.Description
    MsgBox "作り直しに失敗しました: " & Err.Description, vbExclamation, "集積の作り直し"
End Sub


Private Function ブロック書く(ByRef v As Variant, ByVal 開始 As Long, ByVal 終了 As Long) As Boolean
    ブロック書く = False
    Dim n As Long: n = 終了 - 開始 + 1
    If n < 1 Then Exit Function

    Dim buf() As Variant
    ReDim buf(1 To n, 1 To 10)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To 10
            buf(i, j) = v(開始 + i - 1, j)
        Next j
    Next i

    ' ★時刻区分は Excel が時刻に変換していることがあるので両対応
    Dim 区分 As String
    区分 = 時刻文字(buf(1, 2))
    If 区分 = "" Then 区分 = Trim$(CStr(buf(1, 2)))
    If 区分 = "" Then Exit Function
    If Left$(区分, 2) = "手動" Then Exit Function
    If Len(区分) <> 5 Then Exit Function

    集積へ書く buf, 区分
    ブロック書く = True
End Function


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
