Attribute VB_Name = "Mod_MS一括取込"
Option Explicit

' ==============================================================================
'  MS時系列 一括取込
'
'  マーケットスピード過去データ取込（Mod_MS_TimeSeriesImport）で
'  動いている書き込み処理を、そのまま銘柄管理の全銘柄に回すだけのもの。
'  書き込みのロジックは1文字も変えていない。足したのは
'
'    ・銘柄管理 B5:B305 を上から順に データ取込!B1 へ入れる
'    ・RSS の応答を待つ
'    ・確認/完了ダイアログを出さない（最後に1回だけまとめて出す）
'    ・日付の並びが1銘柄目と違う銘柄を見つけて報告する
'
'  の4つだけ。
'
'  使い方
'    1. マーケットスピードを起動してログイン
'    2. データ取込!K1 = 400 / M1 = 開始日(yyyymmdd, 今日の約420日前)
'    3. Alt+F8 → MS一括取込
'    ※ 途中で止めるときは ESC
' ==============================================================================

Private Const PWD          As String = "ne19480314"
Private Const LAST_COL     As Long = 254   ' IT列 = E列 + 250日 - 1
Private Const MEI_ROW1     As Long = 5     ' 銘柄管理の先頭行（5行目=TOPX）
Private Const MEI_ROWN     As Long = 305
Private Const WAIT_LIMIT   As Double = 30  ' RSS応答の待ち上限（秒）
Private Const STABLE_POLLS As Long = 3     ' この回数だけ内容が変わらなければ受信完了
Private Const POLL_WAIT    As Double = 0.2 ' 1回の待ち（秒）
Private Const MIN_SETTLE   As Double = 1   ' 受信完了と認めるまでの最低待ち（秒）

' データ取込シートの列位置（FindHeaderCols が埋める）
Private hRow As Long
Private cName As Long, cDt As Long, cOpen As Long, cHigh As Long
Private cLow As Long, cClose As Long, cVol As Long

' 1銘柄目の日付の並び（ズレ検出の基準）
Private refDates() As Date
Private refCnt As Long


Public Sub MS一括取込()
    Dim prevCalc As XlCalculation, prevScr As Boolean
    prevCalc = Application.Calculation
    prevScr = Application.ScreenUpdating

    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler

    Dim dataWs As Worksheet, mws As Worksheet
    Set dataWs = Nothing: Set mws = Nothing
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("データ取込")
    Set mws = ThisWorkbook.Sheets("銘柄管理")
    On Error GoTo ErrHandler
    If dataWs Is Nothing Then Err.Raise 5, , "「データ取込」シートが見つかりません。"
    If mws Is Nothing Then Err.Raise 5, , "「銘柄管理」シートが見つかりません。"

    ' --- ヘッダ行と列位置（1銘柄目の応答が出ている前提で探す）---
    FindHeaderCols dataWs
    If hRow = 0 Then Err.Raise 5, , _
        "データ取込シートのヘッダ行が見つかりません。" & vbCrLf & _
        "「日付」「始値」「終値」を含む行が必要です。" & vbCrLf & vbCrLf & _
        "先に B1 に1銘柄入れて、RSSが表示されている状態にしてください。"

    ' --- 対象銘柄を数える ---
    Dim total As Long, r As Long
    For r = MEI_ROW1 To MEI_ROWN
        If Trim$(CStr(mws.Cells(r, 2).Value)) <> "" Then total = total + 1
    Next r
    If total = 0 Then Err.Raise 5, , "銘柄管理に銘柄がありません。"

    If MsgBox("銘柄管理の " & total & " 銘柄を、250営業日分まで取り込みます。" & vbCrLf & vbCrLf & _
              "取得条件（データ取込シート）" & vbCrLf & _
              "  K1 本数   = " & dataWs.Range("K1").Value & vbCrLf & _
              "  M1 開始日 = " & dataWs.Range("M1").Value & vbCrLf & vbCrLf & _
              "銘柄数によっては10～30分かかります。" & vbCrLf & _
              "途中で止めるときは ESC キーを押してください。" & vbCrLf & vbCrLf & _
              "始めますか？", vbYesNo + vbQuestion, "MS時系列 一括取込") <> vbYes Then
        Exit Sub
    End If

    refCnt = 0
    Dim okCnt As Long, ngCnt As Long, doneCnt As Long
    Dim ngList As String, shiftList As String
    Dim t0 As Double: t0 = Timer

    For r = MEI_ROW1 To MEI_ROWN
        Dim code As String: code = Trim$(CStr(mws.Cells(r, 2).Value))
        If code <> "" Then
            doneCnt = doneCnt + 1
            Dim meiName As String: meiName = Trim$(CStr(mws.Cells(r, 3).Value))

            ShowProgress doneCnt, total, code, meiName, t0

            ' ---- ① RSS に投げて応答を待つ（画面は止めない）----
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic

            dataWs.Range("B1").Value = code
            Application.Calculate

            Dim gotIt As Boolean: gotIt = WaitForRss(dataWs, meiName)

            ' ---- ② 5シートへ書き込む（ここは速度優先で画面を止める）----
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False

            Dim wrote As Long: wrote = 0
            Dim shifted As Boolean: shifted = False
            If gotIt Then wrote = WriteOneStock(dataWs, r, shifted)

            If wrote > 0 Then
                okCnt = okCnt + 1
                If shifted Then shiftList = shiftList & "  " & code & " " & meiName & vbCrLf
            Else
                ngCnt = ngCnt + 1
                ngList = ngList & "  " & code & " " & meiName & vbCrLf
            End If
        End If
    Next r

    RestoreApp prevCalc, prevScr

    Dim msg As String
    msg = "一括取込 完了" & vbCrLf & String(34, "-") & vbCrLf & _
          "  成功:   " & okCnt & " 銘柄" & vbCrLf & _
          "  失敗:   " & ngCnt & " 銘柄" & vbCrLf & _
          "  所要:   " & Format$((Timer - t0) / 60, "0.0") & " 分" & vbCrLf
    If ngList <> "" Then
        msg = msg & vbCrLf & "取得できなかった銘柄:" & vbCrLf & Left$(ngList, 900)
    End If
    If shiftList <> "" Then
        msg = msg & vbCrLf & "★日付の並びが1銘柄目と違う銘柄:" & vbCrLf & Left$(shiftList, 700) & _
              vbCrLf & "  （売買停止などで日が抜けた銘柄です。" & vbCrLf & _
              "    この銘柄だけ列と日付がズレています）" & vbCrLf
    End If
    msg = msg & vbCrLf & "F9 を押して再計算してください。"

    MsgBox msg, vbInformation, "MS時系列 一括取込"
    Exit Sub

ErrHandler:
    If Err.Number = 18 Then          ' ESC
        RestoreApp prevCalc, prevScr
        MsgBox "ESC で中断しました。" & vbCrLf & _
               "ここまでに取り込んだぶんは残っています。", vbExclamation, "中断"
        Exit Sub
    End If
    Dim e As String: e = "Err " & Err.Number & ": " & Err.Description
    RestoreApp prevCalc, prevScr
    MsgBox "取込を中断しました。" & vbCrLf & vbCrLf & e, vbCritical, "エラー"
End Sub


'------------------------------------------------------------------------------
' 1銘柄を5シートへ書き込む
'   Mod_MS_TimeSeriesImport の書き込み部分をそのまま持ってきたもの。
'   変えたのは、ダイアログを出さないことと、戻り値を返すことだけ。
'------------------------------------------------------------------------------
Private Function WriteOneStock(ByVal dataWs As Worksheet, ByVal targetRow As Long, _
                               ByRef shifted As Boolean) As Long
    ' --- データ行を収集 (ヘッダ行+1 から) ---
    Dim dateRows() As Long
    Dim dateVals() As Date
    ReDim dateRows(1 To 400)
    ReDim dateVals(1 To 400)
    Dim cnt As Long: cnt = 0

    Dim lastDataRow As Long
    lastDataRow = dataWs.Cells(dataWs.Rows.Count, cDt).End(xlUp).Row

    Dim dr As Long
    For dr = hRow + 1 To lastDataRow
        If cnt >= 400 Then Exit For
        Dim dv As Variant: dv = dataWs.Cells(dr, cDt).Value
        If IsDate(dv) Then
            cnt = cnt + 1
            dateRows(cnt) = dr
            dateVals(cnt) = CDate(dv)
        End If
    Next dr
    If cnt = 0 Then Exit Function

    ' --- 日付で降順ソート (バブルソート) ---
    Dim i As Long, j As Long, tmpD As Date, tmpR As Long
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dateVals(j) < dateVals(j + 1) Then
                tmpD = dateVals(j): dateVals(j) = dateVals(j + 1): dateVals(j + 1) = tmpD
                tmpR = dateRows(j): dateRows(j) = dateRows(j + 1): dateRows(j + 1) = tmpR
            End If
        Next j
    Next i

    ' --- 1銘柄目の並びを基準にして、ズレていないか見る ---
    Dim useN As Long: useN = cnt
    If useN > LAST_COL - 4 Then useN = LAST_COL - 4      ' 250日分まで
    If refCnt = 0 Then
        ReDim refDates(1 To useN)
        For i = 1 To useN
            refDates(i) = dateVals(i)
        Next i
        refCnt = useN
    Else
        Dim chk As Long: chk = useN
        If chk > refCnt Then chk = refCnt
        For i = 1 To chk
            If dateVals(i) <> refDates(i) Then shifted = True: Exit For
        Next i
    End If

    ' --- 5シートに書込 ---
    Dim ShNames As Variant
    ShNames = Array("終値", "高値", "安値", "始値", "出来高")
    Dim srcCols As Variant
    srcCols = Array(cClose, cHigh, cLow, cOpen, cVol)

    Dim writeCount As Long: writeCount = 0

    Dim s As Integer
    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(ShNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSh

        Dim wasP As Boolean: wasP = ws.ProtectContents
        On Error Resume Next
        ws.Unprotect Password:=PWD
        On Error GoTo 0

        ' E～IT (col 5～254) をクリア
        ws.Range(ws.Cells(targetRow, 5), ws.Cells(targetRow, LAST_COL)).ClearContents

        ' 書込 (E列=最新, F列=1日前, ... IT列=249日前)
        Dim d As Long
        For d = 1 To cnt
            Dim col As Long: col = 4 + d   ' E=5, F=6, ...
            If col > LAST_COL Then Exit For

            Dim srcR As Long: srcR = dateRows(d)
            Dim srcV As Variant: srcV = dataWs.Cells(srcR, srcCols(s)).Value

            Dim numV As Double: numV = 0
            On Error Resume Next
            numV = CDbl(Replace(CStr(srcV), ",", ""))
            On Error GoTo 0

            If numV > 0 Then
                ws.Cells(targetRow, col).Value = numV
                ws.Cells(targetRow, col).NumberFormat = "#,##0"
                writeCount = writeCount + 1
            End If

            ' 3行目 日付ヘッダも更新
            ws.Cells(3, col).Value = dateVals(d)
            ws.Cells(3, col).NumberFormat = "m/d"
        Next d

        If wasP Then
            On Error Resume Next
            ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                       DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                       AllowFiltering:=True, AllowSorting:=True
            On Error GoTo 0
        End If
NextSh:
    Next s

    WriteOneStock = writeCount
End Function


'------------------------------------------------------------------------------
' ヘッダ行と列位置を探す（元マクロと同じ探し方）
'------------------------------------------------------------------------------
Private Sub FindHeaderCols(ByVal dataWs As Worksheet)
    hRow = 0
    cName = 0: cDt = 0: cOpen = 0: cHigh = 0: cLow = 0: cClose = 0: cVol = 0

    Dim hr As Long, hc As Long
    Dim hasDate As Boolean, hasOpen As Boolean, hasClose As Boolean
    For hr = 1 To 50
        hasDate = False: hasOpen = False: hasClose = False
        For hc = 1 To 30
            Dim hv As String: hv = Trim$(CStr(dataWs.Cells(hr, hc).Value))
            If hv = "日付" Then hasDate = True
            If hv = "始値" Then hasOpen = True
            If hv = "終値" Then hasClose = True
        Next hc
        If hasDate And hasOpen And hasClose Then
            hRow = hr
            Exit For
        End If
    Next hr
    If hRow = 0 Then Exit Sub

    Dim cc As Long
    For cc = 1 To 30
        Dim cv As String: cv = Trim$(CStr(dataWs.Cells(hRow, cc).Value))
        Select Case cv
            Case "日付":   cDt = cc
            Case "始値":   cOpen = cc
            Case "高値":   cHigh = cc
            Case "安値":   cLow = cc
            Case "終値":   cClose = cc
            Case "銘柄名称", "銘柄名": cName = cc
            Case "出来高", "売買高", "出来高(株)", "出来高（株）": cVol = cc
        End Select
    Next cc

    If cDt = 0 Or cOpen = 0 Or cHigh = 0 Or cLow = 0 Or cClose = 0 Or cVol = 0 Then hRow = 0
End Sub


'------------------------------------------------------------------------------
' RSS の応答を待つ
'
'   銘柄名が目的の銘柄になっていて、かつ行数が3回続けて同じなら受信完了。
'   行数だけで見ると、連続する2銘柄がたまたま同じ日数のときに
'   前の銘柄の古いデータをそのまま読んでしまう。
'------------------------------------------------------------------------------
Private Function WaitForRss(ByVal dataWs As Worksheet, ByVal meiName As String) As Boolean
    Dim t0 As Double: t0 = Timer
    Dim prevN As Long: prevN = -1
    Dim same As Long: same = 0

    Do While Timer - t0 < WAIT_LIMIT
        PauseFor POLL_WAIT
        DoEvents

        Dim n As Long
        n = dataWs.Cells(dataWs.Rows.Count, cDt).End(xlUp).Row - hRow
        If n < 0 Then n = 0

        Dim nameOk As Boolean
        If cName = 0 Or meiName = "" Or IsNumeric(meiName) Then
            ' 名前で確認できないときは行数だけで見る。
            ' 銘柄管理 C5（TOPX）には銘柄名ではなく指数値が入っているため、
            ' 数値のときも名前照合をあきらめる（照合しないと永久に一致しない）。
            nameOk = True
        Else
            nameOk = (Trim$(CStr(dataWs.Cells(hRow + 1, cName).Value)) = meiName)
        End If

        If n > 0 And nameOk Then
            If n = prevN Then
                same = same + 1
            Else
                same = 0
            End If
            If same >= STABLE_POLLS And (Timer - t0) >= MIN_SETTLE Then
                WaitForRss = True
                Exit Function
            End If
        Else
            same = 0
        End If
        prevN = n
    Loop
End Function


Private Sub PauseFor(ByVal sec As Double)
    Dim t0 As Double: t0 = Timer
    Do While Timer - t0 < sec
        DoEvents
    Loop
End Sub


Private Sub ShowProgress(ByVal doneCnt As Long, ByVal total As Long, _
                         ByVal code As String, ByVal meiName As String, _
                         ByVal t0 As Double)
    Dim el As Double: el = Timer - t0
    Dim eta As String: eta = "--"
    If doneCnt > 1 Then
        eta = Format$((el / (doneCnt - 1)) * (total - doneCnt + 1) / 60, "0.0") & " 分"
    End If
    Application.StatusBar = "一括取込 " & doneCnt & " / " & total & "   " & _
                            code & " " & meiName & "   経過 " & _
                            Format$(el / 60, "0.0") & " 分 / 残り約 " & eta & _
                            "   （中止は ESC）"
    DoEvents
End Sub


Private Sub RestoreApp(ByVal prevCalc As XlCalculation, ByVal prevScr As Boolean)
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScr
    Application.StatusBar = False
    On Error GoTo 0
End Sub


'==============================================================================
' データ取込シートのヘッダと数式を作り直す
'
'  「データ取込シートをクリアしますか？」で「はい」を押すと
'  Range("A2:Z500").ClearContents が走り、2行目のヘッダと A3 の
'  RssChartPast 数式まで消える。そうなると
'  「ヘッダ行が見つかりません」でどのマクロも動かなくなる。
'  これを1発で戻すためのもの。
'==============================================================================
Public Sub データ取込_ヘッダ修復()
    Dim dws As Worksheet
    Set dws = Nothing
    On Error Resume Next
    Set dws = ThisWorkbook.Sheets("データ取込")
    On Error GoTo 0
    If dws Is Nothing Then
        MsgBox "「データ取込」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    If MsgBox("データ取込シートの 1～2行目と A3 の数式を作り直します。" & vbCrLf & _
              "取得済みの価格シートには触りません。" & vbCrLf & vbCrLf & _
              "続けますか？", vbYesNo + vbQuestion, "ヘッダ修復") <> vbYes Then Exit Sub

    On Error Resume Next
    dws.Unprotect Password:=PWD
    On Error GoTo 0

    ' --- 1行目：ラベルと取得条件 ---
    dws.Range("A1").Value = "コード →"
    dws.Range("D1").Value = "→ 株探(過去データ)"
    dws.Range("E1").Value = "→"
    dws.Range("J1").Value = "本数"
    ' L1 はどの数式からも VBA からも参照されていない空き。
    ' J1 が K1 のラベルなのに M1 だけラベルが無かったので、ここに入れる。
    dws.Range("L1").Value = "開始日"
    If Trim$(CStr(dws.Range("B1").Value)) = "" Then dws.Range("B1").Value = "3401"
    dws.Range("K1").Value = 400
    dws.Range("M1").Value = Format$(Date - 420, "yyyymmdd")

    ' --- 2行目：RssChartPast に渡す項目名（これがヘッダ行）---
    dws.Range("B2").Value = "銘柄名称"
    dws.Range("C2").Value = "市場名称"
    dws.Range("D2").Value = "日付"
    dws.Range("E2").Value = "時刻"
    dws.Range("F2").Value = "出来高"
    dws.Range("G2").Value = "始値"
    dws.Range("H2").Value = "高値"
    dws.Range("I2").Value = "安値"
    dws.Range("J2").Value = "終値"
    dws.Range("K2").Value = "移動平均1"
    dws.Range("L2").Value = "移動平均2"

    ' --- 数式 ---
    Dim errMsg As String
    On Error Resume Next
    dws.Range("C1").Formula = "=IF(B1="""","""",RssMarket(B1,""銘柄名称""))"
    If Err.Number <> 0 Then errMsg = errMsg & "  C1 (RssMarket)" & vbCrLf: Err.Clear
    dws.Range("A3").Formula = "=RssChartPast(B2:J2,B1,""D"",$M$1,$K$1)"
    If Err.Number <> 0 Then errMsg = errMsg & "  A3 (RssChartPast)" & vbCrLf: Err.Clear
    On Error GoTo 0

    Application.Calculate
    DoEvents

    If errMsg <> "" Then
        MsgBox "数式を入れられませんでした:" & vbCrLf & errMsg & vbCrLf & _
               "マーケットスピードが起動してログイン済みか確認し、" & vbCrLf & _
               "もう一度実行してください。", vbExclamation, "ヘッダ修復"
        Exit Sub
    End If

    MsgBox "作り直しました。" & vbCrLf & vbCrLf & _
           "  B1 銘柄コード = " & dws.Range("B1").Value & vbCrLf & _
           "  K1 本数       = " & dws.Range("K1").Value & vbCrLf & _
           "  M1 開始日     = " & dws.Range("M1").Value & _
           "  （" & Format$(Date - 420, "yyyy/mm/dd") & "）" & vbCrLf & vbCrLf & _
           "B3 に銘柄名、D3 以降に日付が並べば復旧完了です。" & vbCrLf & _
           "（出るまで数秒かかります）", vbInformation, "ヘッダ修復"
End Sub


'==============================================================================
' 取り込めていない銘柄だけ、もう一度取りに行く
'
'  一括取込で失敗した銘柄（RSSの応答が間に合わなかった等）を拾い直す。
'  終値シートの E:IT に何日ぶん入っているかで判定するので、
'  何度実行しても、そろっている銘柄には触らない。
'  待ち時間は一括取込より長め（60秒）。
'==============================================================================
Public Sub MS取込_不足分だけ()
    Const NEED As Long = 200          ' この日数未満の銘柄を取り直す
    Const LONG_WAIT As Double = 60

    Dim prevCalc As XlCalculation, prevScr As Boolean
    prevCalc = Application.Calculation
    prevScr = Application.ScreenUpdating

    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler

    Dim dataWs As Worksheet, mws As Worksheet, cws As Worksheet
    Set dataWs = Nothing: Set mws = Nothing: Set cws = Nothing
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("データ取込")
    Set mws = ThisWorkbook.Sheets("銘柄管理")
    Set cws = ThisWorkbook.Sheets("終値")
    On Error GoTo ErrHandler
    If dataWs Is Nothing Or mws Is Nothing Or cws Is Nothing Then
        Err.Raise 5, , "データ取込 / 銘柄管理 / 終値 のいずれかが見つかりません。"
    End If

    FindHeaderCols dataWs
    If hRow = 0 Then Err.Raise 5, , _
        "データ取込シートのヘッダ行が見つかりません。" & vbCrLf & _
        "先に データ取込_ヘッダ修復 を実行してください。"

    ' --- 足りない銘柄を数える ---
    Dim targets As String, total As Long, r As Long
    For r = MEI_ROW1 To MEI_ROWN
        If Trim$(CStr(mws.Cells(r, 2).Value)) <> "" Then
            If FilledDays(cws, r) < NEED Then
                total = total + 1
                targets = targets & "  " & Trim$(CStr(mws.Cells(r, 2).Value)) & _
                          " (" & FilledDays(cws, r) & " 日)" & vbCrLf
            End If
        End If
    Next r

    If total = 0 Then
        MsgBox "足りない銘柄はありません。" & vbCrLf & _
               "全銘柄が " & NEED & " 日以上そろっています。", vbInformation, "不足分の取り直し"
        Exit Sub
    End If

    If MsgBox(total & " 銘柄が " & NEED & " 日分に足りていません。" & vbCrLf & _
              "取り直しますか？" & vbCrLf & vbCrLf & Left$(targets, 900), _
              vbYesNo + vbQuestion, "不足分の取り直し") <> vbYes Then Exit Sub

    refCnt = 0
    Dim okCnt As Long, ngCnt As Long, doneCnt As Long, ngList As String
    Dim t0 As Double: t0 = Timer

    For r = MEI_ROW1 To MEI_ROWN
        Dim code As String: code = Trim$(CStr(mws.Cells(r, 2).Value))
        If code <> "" Then
            If FilledDays(cws, r) < NEED Then
                doneCnt = doneCnt + 1
                Dim meiName As String: meiName = Trim$(CStr(mws.Cells(r, 3).Value))
                ShowProgress doneCnt, total, code, meiName, t0

                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                dataWs.Range("B1").Value = ""       ' いったん空にして確実に取り直させる
                Application.Calculate
                dataWs.Range("B1").Value = code
                Application.Calculate

                Dim gotIt As Boolean: gotIt = WaitLong(dataWs, meiName, LONG_WAIT)

                Application.Calculation = xlCalculationManual
                Application.ScreenUpdating = False

                Dim wrote As Long: wrote = 0
                Dim shifted As Boolean: shifted = False
                If gotIt Then wrote = WriteOneStock(dataWs, r, shifted)

                If wrote > 0 Then
                    okCnt = okCnt + 1
                Else
                    ngCnt = ngCnt + 1
                    ngList = ngList & "  " & code & " " & meiName & vbCrLf
                End If
            End If
        End If
    Next r

    RestoreApp prevCalc, prevScr

    Dim msg As String
    msg = "不足分の取り直し 完了" & vbCrLf & String(30, "-") & vbCrLf & _
          "  成功: " & okCnt & " 銘柄" & vbCrLf & _
          "  失敗: " & ngCnt & " 銘柄" & vbCrLf
    If ngList <> "" Then
        msg = msg & vbCrLf & "まだ取得できない銘柄:" & vbCrLf & Left$(ngList, 700) & vbCrLf & _
              "RSSがこの銘柄の日足を返していない可能性があります。" & vbCrLf
    End If
    msg = msg & vbCrLf & "Ctrl+S で保存し、F9 で再計算してください。"
    MsgBox msg, vbInformation, "不足分の取り直し"
    Exit Sub

ErrHandler:
    If Err.Number = 18 Then
        RestoreApp prevCalc, prevScr
        MsgBox "ESC で中断しました。", vbExclamation, "中断"
        Exit Sub
    End If
    Dim e As String: e = "Err " & Err.Number & ": " & Err.Description
    RestoreApp prevCalc, prevScr
    MsgBox "中断しました。" & vbCrLf & vbCrLf & e, vbCritical, "エラー"
End Sub


' 終値シートの指定行に、値が入っている日数
Private Function FilledDays(ByVal cws As Worksheet, ByVal r As Long) As Long
    Dim n As Long
    On Error Resume Next
    n = Application.WorksheetFunction.Count( _
            cws.Range(cws.Cells(r, 5), cws.Cells(r, LAST_COL)))
    If Err.Number <> 0 Then Err.Clear: n = 0
    On Error GoTo 0
    FilledDays = n
End Function


' WaitForRss と同じだが、待ち時間を指定できる
Private Function WaitLong(ByVal dataWs As Worksheet, ByVal meiName As String, _
                          ByVal limitSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer
    Dim prevN As Long: prevN = -1
    Dim same As Long: same = 0

    Do While Timer - t0 < limitSec
        PauseFor POLL_WAIT
        DoEvents

        Dim n As Long
        n = dataWs.Cells(dataWs.Rows.Count, cDt).End(xlUp).Row - hRow
        If n < 0 Then n = 0

        Dim nameOk As Boolean
        If cName = 0 Or meiName = "" Or IsNumeric(meiName) Then
            nameOk = True
        Else
            nameOk = (Trim$(CStr(dataWs.Cells(hRow + 1, cName).Value)) = meiName)
        End If

        If n > 0 And nameOk Then
            If n = prevN Then same = same + 1 Else same = 0
            If same >= STABLE_POLLS And (Timer - t0) >= MIN_SETTLE Then
                WaitLong = True
                Exit Function
            End If
        Else
            same = 0
        End If
        prevN = n
    Loop
End Function
