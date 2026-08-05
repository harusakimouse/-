Attribute VB_Name = "Mod_一括再取込"
Option Explicit

' ==================================================================
' OHLCV 一括再取込（全銘柄）  RssChartPast 版
'
' 【何をするか】
'   MarketSpeed II RSS の RssChartPast（過去チャートデータ）を使って、
'   登録全銘柄の日足 OHLCV を取り直し、始値/高値/安値/終値/出来高 の
'   5シートに書き戻します。
'
' 【なぜ RssChartPast か】
'   RssMarket は「今の値」しか返しません。過去列を直す手段になりません。
'   データ取込シートの A3 に既にある
'       =RssChartPast(B2:J2, B1, "D", $M$1, $K$1)
'   と同じ関数を、全銘柄ぶんプログラムから回します。
'
' 【2つのモード】
'   モードA: 日付ヘッダ(3行目)も作り直してから全銘柄を再取込
'            → 列と日付の対応が完全に整合する。汚れを一掃したいとき。
'            → 2026/7/26(日曜)のような不正な列や、抜けている営業日も直る。
'   モードB: 既存の日付ヘッダは触らず、一致する日付にだけ書き込む
'            → 安全側。欠測セルの穴埋めだけしたいとき。
'
' 【安全設計】
'   ・作業用の一時シートを作って処理し、最後に消す（データ取込シートは触らない）
'   ・C列(RSS式)・D列(前日比式) には一切触らない
'   ・ESC で中断できる。中断・エラーのどちらでも
'     計算モード / 画面更新 / シート保護 を必ず元に戻す
'   ・中断しても、次に実行すると続きから再開できる
'
' 【所要時間】
'   1銘柄あたり 1〜3秒 × 287銘柄 ＝ おおむね 8〜15分。
'   実行中は Excel を触らないでください。
' ==================================================================

Private Const PWD           As String = "ne19480314"

Private Const R_TOPIX       As Long = 5
Private Const R_FIRST       As Long = 6
Private Const R_LAST        As Long = 505
Private Const R_DATE        As Long = 3        ' 日付ヘッダ行
Private Const C_CODE        As Long = 1        ' A列
Private Const C_TODAY       As Long = 5        ' E列（最新日）
Private Const C_HISTEND     As Long = 104      ' CZ列（最古日）

Private Const WORK_SHEET    As String = "_RSS作業"
Private Const RESUME_CELL   As String = "AC1"  ' 銘柄管理 に再開位置を記録

Private Const START_BACK_DAYS As Long = 170    ' 何日前から取得するか(暦日)
Private Const MAX_BARS        As Long = 200    ' 最大取得本数
Private Const WAIT_SEC        As Long = 20     ' 1銘柄あたりの応答待ち上限(秒)

Private gAbort As Boolean


' ==================================================================
'  メイン
' ==================================================================
Public Sub OHLCV_一括再取込_全銘柄()

    ' ---------- 事前チェック ----------
    If Not AddinAlive() Then
        MsgBox "MarketSpeed II RSS アドインが読み込まれていません（#NAME?）。" & vbCrLf & vbCrLf & _
               "MarketSpeed II を起動・ログインしてから Excel を開き直して、" & vbCrLf & _
               "もう一度実行してください。", vbCritical, "中止"
        Exit Sub
    End If

    Dim closeWs As Worksheet
    On Error Resume Next
    Set closeWs = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If closeWs Is Nothing Then MsgBox "終値シートがありません。", vbCritical: Exit Sub

    ' ---------- モード選択 ----------
    Dim ans As VbMsgBoxResult
    ans = MsgBox("OHLCV 一括再取込（全銘柄）" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf & _
                 "【はい】 モードA: 日付ヘッダも作り直して全部取り直す" & vbCrLf & _
                 "         列と日付の対応が完全に整います。" & vbCrLf & _
                 "         不正な列(日曜など)・抜けている営業日も直ります。" & vbCrLf & vbCrLf & _
                 "【いいえ】モードB: 既存の日付ヘッダは変えず、" & vbCrLf & _
                 "         一致する日付にだけ書き込む（安全側）" & vbCrLf & vbCrLf & _
                 "【キャンセル】やめる" & vbCrLf & vbCrLf & _
                 "★実行前に必ずブックのコピーを取ってください★" & vbCrLf & _
                 "★所要 8〜15分。実行中は Excel を触らないでください★", _
                 vbYesNoCancel + vbExclamation, "一括再取込")
    If ans = vbCancel Then Exit Sub

    Dim rebuildHeader As Boolean
    rebuildHeader = (ans = vbYes)

    ' ---------- 対象銘柄の収集 ----------
    Dim codes() As String, rowsOf() As Long, nCode As Long
    ReDim codes(1 To R_LAST): ReDim rowsOf(1 To R_LAST)
    Dim r As Long
    For r = R_FIRST To R_LAST
        Dim cd As String
        cd = Trim$(CStr(closeWs.Cells(r, C_CODE).Value))
        If IsStockCode(cd) Then
            nCode = nCode + 1
            codes(nCode) = cd
            rowsOf(nCode) = r
        End If
    Next r
    If nCode = 0 Then MsgBox "対象銘柄がありません。", vbExclamation: Exit Sub

    ' ---------- 再開位置 ----------
    Dim startIdx As Long: startIdx = 1
    Dim savedIdx As Long: savedIdx = GetResumeIndex()
    If savedIdx > 0 And savedIdx < nCode Then
        If MsgBox("前回 " & savedIdx & " 銘柄目まで完了しています。" & vbCrLf & _
                  "続きから再開しますか?" & vbCrLf & vbCrLf & _
                  "【いいえ】＝ 最初からやり直す", _
                  vbYesNo + vbQuestion, "再開") = vbYes Then
            startIdx = savedIdx + 1
        End If
    End If

    ' ---------- 実行 ----------
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim wk As Worksheet
    Dim okCount As Long, ngCount As Long, cellCount As Long, skipDate As Long
    Dim ngList As String
    Dim t0 As Double: t0 = Timer

    gAbort = False
    On Error GoTo Cleanup
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic   ' RTD を回すため自動が必須

    Set wk = MakeWorkSheet()
    UnprotectAll

    ' ---------- モードA: 日付ヘッダの作り直し ----------
    If rebuildHeader Then
        Application.StatusBar = "日付ヘッダを作り直しています（基準銘柄 " & codes(1) & "）..."
        Dim hdrDates() As Date, nHdr As Long
        If Not FetchBars(wk, codes(1), hdrDates, Nothing, Nothing, Nothing, Nothing, Nothing, nHdr, True) Then
            MsgBox "基準銘柄 " & codes(1) & " の過去データが取得できませんでした。" & vbCrLf & _
                   "モードB（既存ヘッダ維持）で実行し直してください。", vbCritical, "中止"
            GoTo Cleanup
        End If
        WriteHeader hdrDates, nHdr
        startIdx = 1                     ' ヘッダを作り直したら全銘柄やり直し
        ClearAllHistory
    End If

    ' ---------- 銘柄ループ ----------
    Dim i As Long
    For i = startIdx To nCode
        If gAbort Then Exit For

        Application.StatusBar = "一括再取込 " & i & " / " & nCode & _
                                "  [" & codes(i) & "]  取込済セル=" & cellCount & _
                                "  経過 " & Format((Timer - t0) / 60, "0.0") & " 分" & _
                                "   （中断は ESC）"

        Dim dts() As Date
        Dim vo() As Double, vh() As Double, vl() As Double, vc() As Double, vv() As Double
        Dim nBar As Long

        If FetchBars(wk, codes(i), dts, vo, vh, vl, vc, vv, nBar, False) Then
            Dim wrote As Long, missed As Long
            ApplyBars rowsOf(i), dts, vo, vh, vl, vc, vv, nBar, wrote, missed
            cellCount = cellCount + wrote
            skipDate = skipDate + missed
            okCount = okCount + 1
            SetResumeIndex i
        Else
            ngCount = ngCount + 1
            If Len(ngList) < 300 Then ngList = ngList & codes(i) & " "
        End If

        DoEvents
    Next i

Cleanup:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description

    On Error Resume Next
    If Not wk Is Nothing Then wk.Delete
    ProtectAll
    Application.Calculation = xlCalculationAutomatic
    If prevCalc = xlCalculationManual Then Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0

    Dim tail As String
    If eN = 18 Or gAbort Then
        tail = "★ESC で中断しました。もう一度実行すると続きから再開できます。"
    ElseIf eN <> 0 Then
        tail = "★エラーで中断: " & eN & " " & eD & vbCrLf & _
               "  計算モード・シート保護は元に戻しました。"
    Else
        tail = "完了しました。"
        SetResumeIndex 0
    End If

    MsgBox "【OHLCV 一括再取込】" & vbCrLf & String(34, "-") & vbCrLf & _
           "モード: " & IIf(rebuildHeader, "A (ヘッダ作り直し)", "B (既存ヘッダ維持)") & vbCrLf & _
           "対象銘柄: " & nCode & vbCrLf & _
           "成功: " & okCount & " / 失敗: " & ngCount & vbCrLf & _
           IIf(ngCount > 0, "失敗コード: " & Trim$(ngList) & vbCrLf, "") & _
           "書込セル: " & cellCount & vbCrLf & _
           "ヘッダに無い日付でスキップ: " & skipDate & vbCrLf & _
           "所要: " & Format((Timer - t0) / 60, "0.0") & " 分" & vbCrLf & _
           String(34, "-") & vbCrLf & tail, _
           IIf(eN = 0 And Not gAbort, vbInformation, vbExclamation), "一括再取込"
End Sub


' ==================================================================
'  RssChartPast で 1 銘柄ぶんの日足を取得
'  戻り値: 取得できたら True
'  onlyDates=True なら日付だけ返す（ヘッダ作成用）
' ==================================================================
Private Function FetchBars(ByVal wk As Worksheet, ByVal code As String, _
                           ByRef dts() As Date, _
                           ByRef vo() As Double, ByRef vh() As Double, _
                           ByRef vl() As Double, ByRef vc() As Double, _
                           ByRef vv() As Double, _
                           ByRef nBar As Long, _
                           ByVal onlyDates As Boolean) As Boolean
    nBar = 0

    ' --- 出力域をクリアして条件をセット ---
    wk.Range("A3:L400").ClearContents
    wk.Range("B1").Value = code
    wk.Range("K1").Value = MAX_BARS
    wk.Range("M1").Value = CLng(Format(Date - START_BACK_DAYS, "yyyymmdd"))

    wk.Range("A3").Formula = "=RssChartPast(B2:J2,B1,""D"",$M$1,$K$1)"

    ' --- 応答待ち（RssChartPast は非同期） ---
    Dim t0 As Double: t0 = Timer
    Dim ready As Boolean
    Do
        DoEvents
        Application.Calculate

        Dim st As String
        st = CStr(wk.Range("A3").Text)
        Dim d1 As Variant: d1 = wk.Range("D3").Value

        If IsDate(d1) Then
            ' 最終行の終端マーカーか、次の空行が現れたら完了とみなす
            If InStr(1, st, "応答待ち") = 0 Then ready = True
            If Not ready Then
                ' 保険: データが 2 行以上そろっていれば取れているとみなす
                If IsDate(wk.Range("D4").Value) Then ready = True
            End If
        End If
        If ready Then Exit Do

        Dim w As Double: w = Timer + 0.3
        Do While Timer < w: DoEvents: Loop
    Loop While (Timer - t0) < WAIT_SEC

    If Not ready Then Exit Function

    ' 書き込み途中を掴まないよう、少しだけ落ち着かせる
    Dim w2 As Double: w2 = Timer + 0.4
    Do While Timer < w2: DoEvents: Loop
    Application.Calculate

    ' --- 読み取り ---
    ReDim dts(1 To MAX_BARS + 10)
    If Not onlyDates Then
        ReDim vo(1 To MAX_BARS + 10): ReDim vh(1 To MAX_BARS + 10)
        ReDim vl(1 To MAX_BARS + 10): ReDim vc(1 To MAX_BARS + 10)
        ReDim vv(1 To MAX_BARS + 10)
    End If

    Dim rr As Long
    For rr = 3 To 3 + MAX_BARS + 5
        Dim dv As Variant: dv = wk.Cells(rr, 4).Value      ' D列 = 日付
        If Not IsDate(dv) Then Exit For

        Dim o As Double, h As Double, l As Double, c As Double, v As Double
        o = SafeD(wk.Cells(rr, 7).Value)                    ' G列 = 始値
        h = SafeD(wk.Cells(rr, 8).Value)                    ' H列 = 高値
        l = SafeD(wk.Cells(rr, 9).Value)                    ' I列 = 安値
        c = SafeD(wk.Cells(rr, 10).Value)                   ' J列 = 終値
        v = SafeD(wk.Cells(rr, 6).Value)                    ' F列 = 出来高

        If onlyDates Or (o > 0 And h > 0 And l > 0 And c > 0) Then
            nBar = nBar + 1
            dts(nBar) = CDate(dv)
            If Not onlyDates Then
                vo(nBar) = o: vh(nBar) = h: vl(nBar) = l: vc(nBar) = c: vv(nBar) = v
            End If
        End If
    Next rr

    FetchBars = (nBar > 0)
End Function


' ==================================================================
'  取得した足を 5 シートの該当列へ書き込む
'  既存の日付ヘッダに一致する列だけに書く（ヘッダは絶対に書き換えない）
' ==================================================================
Private Sub ApplyBars(ByVal targetRow As Long, _
                      ByRef dts() As Date, _
                      ByRef vo() As Double, ByRef vh() As Double, _
                      ByRef vl() As Double, ByRef vc() As Double, _
                      ByRef vv() As Double, _
                      ByVal nBar As Long, _
                      ByRef wrote As Long, ByRef missed As Long)

    Dim map As Object
    Set map = DateColumnMap()

    Dim shNames As Variant, s As Integer
    shNames = Array("始値", "高値", "安値", "終値", "出来高")

    Dim k As Long
    For k = 1 To nBar
        Dim key As String: key = Format(dts(k), "yyyymmdd")
        If Not map.Exists(key) Then
            missed = missed + 1
        Else
            Dim col As Long: col = CLng(map(key))
            For s = 0 To 4
                Dim ws As Worksheet
                Set ws = Nothing
                On Error Resume Next
                Set ws = ThisWorkbook.Sheets(shNames(s))
                On Error GoTo 0
                If ws Is Nothing Then GoTo NextS

                Dim val As Double
                Select Case s
                    Case 0: val = vo(k)
                    Case 1: val = vh(k)
                    Case 2: val = vl(k)
                    Case 3: val = vc(k)
                    Case 4: val = vv(k)
                End Select

                If val > 0 Then
                    ws.Cells(targetRow, col).Value = val
                    wrote = wrote + 1
                End If
NextS:
            Next s
        End If
    Next k
End Sub


' ------------------------------------------------------------------
'  日付ヘッダ → 列番号 の辞書
' ------------------------------------------------------------------
Private Function DateColumnMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("終値")

    Dim c As Long
    For c = C_TODAY To C_HISTEND
        Dim v As Variant: v = ws.Cells(R_DATE, c).Value
        If IsDate(v) Then
            Dim key As String: key = Format(CDate(v), "yyyymmdd")
            If Not d.Exists(key) Then d.Add key, c
        End If
    Next c

    Set DateColumnMap = d
End Function


' ------------------------------------------------------------------
'  日付ヘッダを 5 シートすべてに書き直す（新しい日付ほど左＝E列側）
' ------------------------------------------------------------------
Private Sub WriteHeader(ByRef dts() As Date, ByVal nBar As Long)
    Dim shNames As Variant: shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim nCols As Long: nCols = C_HISTEND - C_TODAY + 1

    Dim s As Integer
    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextS

        ws.Range(ws.Cells(R_DATE, C_TODAY), ws.Cells(R_DATE, C_HISTEND)).ClearContents

        Dim c As Long, k As Long
        For c = C_TODAY To C_HISTEND
            k = nBar - (c - C_TODAY)          ' 末尾(最新)から逆順に
            If k < 1 Then Exit For
            ws.Cells(R_DATE, c).Value = dts(k)
            ws.Cells(R_DATE, c).NumberFormat = "m/d"
        Next c
NextS:
    Next s
End Sub


' ------------------------------------------------------------------
'  履歴領域(E:CZ)をすべてクリア（モードA用）
' ------------------------------------------------------------------
Private Sub ClearAllHistory()
    Dim shNames As Variant: shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer
    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextS
        ws.Range(ws.Cells(R_TOPIX, C_TODAY), ws.Cells(R_LAST, C_HISTEND)).ClearContents
NextS:
    Next s
End Sub


' ==================================================================
'  補助
' ==================================================================
Private Function MakeWorkSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WORK_SHEET)
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = WORK_SHEET
    ws.Visible = xlSheetVisible

    ' データ取込シートと同じ並びにする（この並びで動作実績がある）
    ws.Range("B2").Value = "銘柄名称"
    ws.Range("C2").Value = "市場名称"
    ws.Range("D2").Value = "日付"
    ws.Range("E2").Value = "時刻"
    ws.Range("F2").Value = "出来高"
    ws.Range("G2").Value = "始値"
    ws.Range("H2").Value = "高値"
    ws.Range("I2").Value = "安値"
    ws.Range("J2").Value = "終値"

    Set MakeWorkSheet = ws
End Function

Private Sub UnprotectAll()
    Dim shNames As Variant: shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer
    For s = 0 To 4
        On Error Resume Next
        ThisWorkbook.Sheets(shNames(s)).Unprotect Password:=PWD
        On Error GoTo 0
    Next s
End Sub

Private Sub ProtectAll()
    Dim shNames As Variant: shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer
    For s = 0 To 4
        On Error Resume Next
        ThisWorkbook.Sheets(shNames(s)).Protect Password:=PWD, UserInterfaceOnly:=True, _
            DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    Next s
End Sub

Private Function GetResumeIndex() As Long
    On Error Resume Next
    GetResumeIndex = CLng(val(ThisWorkbook.Sheets("銘柄管理").Range(RESUME_CELL).Value))
    On Error GoTo 0
End Function

Private Sub SetResumeIndex(ByVal i As Long)
    On Error Resume Next
    With ThisWorkbook.Sheets("銘柄管理")
        .Unprotect Password:=PWD
        If i = 0 Then
            .Range(RESUME_CELL).ClearContents
        Else
            .Range(RESUME_CELL).Value = i
        End If
        .Protect Password:=PWD, UserInterfaceOnly:=True, _
                 DrawingObjects:=True, Contents:=True, Scenarios:=True
    End With
    On Error GoTo 0
End Sub

Private Function SafeD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    SafeD = CDbl(Replace(CStr(v), ",", ""))
    On Error GoTo 0
End Function

Private Function IsStockCode(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) <> 4 Then Exit Function
    Dim i As Long
    For i = 1 To 3
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    Dim c As String: c = UCase$(Mid$(s, 4, 1))
    IsStockCode = ((c >= "0" And c <= "9") Or (c >= "A" And c <= "Z"))
End Function

Private Function AddinAlive() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = Application.Evaluate("RssMarket(""7203"",""銘柄名称"")")
    On Error GoTo 0
    If IsError(v) Then
        If CLng(v) = xlErrName Then Exit Function
    End If
    AddinAlive = True
End Function


' ==================================================================
'  おまけ①: 1銘柄だけ再取込（動作確認用。まずこれで試す）
' ==================================================================
Public Sub OHLCV_再取込_1銘柄()
    If Not AddinAlive() Then
        MsgBox "RSS アドインが読み込まれていません。", vbCritical: Exit Sub
    End If

    Dim code As String
    code = Trim$(InputBox("再取込する銘柄コード（4桁）", "1銘柄テスト", "7203"))
    If code = "" Then Exit Sub
    If Not IsStockCode(code) Then MsgBox "4桁のコードを入力してください。", vbExclamation: Exit Sub

    Dim closeWs As Worksheet: Set closeWs = ThisWorkbook.Sheets("終値")
    Dim targetRow As Long, r As Long
    For r = R_FIRST To R_LAST
        If Trim$(CStr(closeWs.Cells(r, C_CODE).Value)) = code Then targetRow = r: Exit For
    Next r
    If targetRow = 0 Then MsgBox "コード " & code & " は OHLCV シートにありません。", vbExclamation: Exit Sub

    Dim wk As Worksheet
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic
    Set wk = MakeWorkSheet()
    UnprotectAll

    Dim dts() As Date, vo() As Double, vh() As Double, vl() As Double, vc() As Double, vv() As Double
    Dim nBar As Long, wrote As Long, missed As Long

    If FetchBars(wk, code, dts, vo, vh, vl, vc, vv, nBar, False) Then
        ApplyBars targetRow, dts, vo, vh, vl, vc, vv, nBar, wrote, missed
        MsgBox "【" & code & "】 " & targetRow & " 行目" & vbCrLf & _
               "取得した足: " & nBar & " 日分" & vbCrLf & _
               "書込セル: " & wrote & vbCrLf & _
               "ヘッダに無い日付: " & missed & vbCrLf & vbCrLf & _
               "期間: " & Format(dts(1), "m/d") & " 〜 " & Format(dts(nBar), "m/d"), _
               vbInformation, "1銘柄テスト完了"
    Else
        MsgBox "取得できませんでした。" & vbCrLf & _
               "MarketSpeed II が起動しているか、" & vbCrLf & _
               "コードが正しいかを確認してください。", vbExclamation
    End If

Cleanup:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wk Is Nothing Then wk.Delete
    ProtectAll
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0
    If eN <> 0 Then MsgBox "エラー: " & eN & " " & eD, vbCritical
End Sub


' ==================================================================
'  おまけ②: 日付ヘッダの点検（土日・重複・欠測を洗い出す。無変更）
' ==================================================================
Public Sub 日付ヘッダ_点検()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("終値")

    Dim msg As String
    msg = "【日付ヘッダ 点検】（変更しません）" & vbCrLf & String(40, "-") & vbCrLf

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim bad As Long, prevD As Date, c As Long, n As Long

    For c = C_TODAY To C_HISTEND
        Dim v As Variant: v = ws.Cells(R_DATE, c).Value
        If Not IsDate(v) Then GoTo NextC
        Dim d As Date: d = CDate(v)
        n = n + 1

        If Weekday(d, vbMonday) >= 6 Then
            bad = bad + 1
            msg = msg & "  ★列" & c & " " & Format(d, "yyyy/m/d(aaa)") & " ← 土日の列（存在しない営業日）" & vbCrLf
        End If
        If seen.Exists(Format(d, "yyyymmdd")) Then
            bad = bad + 1
            msg = msg & "  ★列" & c & " " & Format(d, "yyyy/m/d") & " ← 日付が重複" & vbCrLf
        Else
            seen.Add Format(d, "yyyymmdd"), c
        End If
        If prevD <> 0 Then
            If d >= prevD Then
                bad = bad + 1
                msg = msg & "  ★列" & c & " " & Format(d, "yyyy/m/d") & " ← 並び順が逆" & vbCrLf
            End If
        End If
        prevD = d
NextC:
    Next c

    ' 充填率
    msg = msg & vbCrLf & "列ごとの欠測（終値シート）:" & vbCrLf
    Dim r As Long, codeCnt As Long
    For r = R_FIRST To R_LAST
        If IsStockCode(Trim$(CStr(ws.Cells(r, C_CODE).Value))) Then codeCnt = codeCnt + 1
    Next r

    For c = C_TODAY To C_HISTEND
        If Not IsDate(ws.Cells(R_DATE, c).Value) Then GoTo NextC2
        Dim filled As Long: filled = 0
        For r = R_FIRST To R_LAST
            If IsStockCode(Trim$(CStr(ws.Cells(r, C_CODE).Value))) Then
                If IsNumeric(ws.Cells(r, c).Value) And ws.Cells(r, c).Value <> "" Then filled = filled + 1
            End If
        Next r
        If filled < codeCnt Then
            msg = msg & "  列" & c & " " & Format(CDate(ws.Cells(R_DATE, c).Value), "m/d") & _
                  "  " & filled & "/" & codeCnt & "  ← 欠測 " & (codeCnt - filled) & " 件" & vbCrLf
        End If
NextC2:
    Next c

    msg = msg & String(40, "-") & vbCrLf & _
          "日付列 " & n & " 本 / 問題 " & bad & " 件" & vbCrLf & _
          "銘柄 " & codeCnt & " 件" & vbCrLf & vbCrLf & _
          IIf(bad > 0, "→ OHLCV_一括再取込_全銘柄 のモードA で直せます。", "→ ヘッダの並びは正常です。")

    MsgBox msg, vbInformation, "日付ヘッダ 点検"
End Sub
