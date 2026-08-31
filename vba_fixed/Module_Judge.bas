Attribute VB_Name = "Module_Judge"
Option Explicit

'==================================================================
' Module_Judge : ティック判定ロジック本体
'
'   ① 方向判定（連続ティック）
'   ② 出来高初動判定（大口の本気）
'   ③ 約定速度（Tick Speed）
'   ④ 総合ジャッジ → 買い / 売り / 中立 / 未計測
'
'   ※ ティックが1件も無い銘柄は「中立」ではなく「未計測」を出します。
'      配信が止まっていたのか、本当に方向が出なかったのかを
'      取り違えないための区別です。
'==================================================================

'------------------------------------------------------------------
' 全銘柄を判定して Judge_Results にまとめる（ボタン④）
'------------------------------------------------------------------
Public Sub JudgeAll()
    If BlockedWhileLogging("判定 実行") Then Exit Sub
    JudgeAllCore
End Sub

'------------------------------------------------------------------
' 判定の本体（記録終了時にも内部から呼ばれます）
'------------------------------------------------------------------
Public Sub JudgeAllCore()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim res() As TickJudgeResult
    Dim i As Long
    Dim prevCalc As XlCalculation
    Dim calcChanged As Boolean
    Dim errNum As Long
    Dim errDesc As String

    Set shts = TargetSheets()

    If shts.Count = 0 Then
        MsgBox "判定対象の銘柄シートがありません。" & vbCrLf & _
               "各シートの B3 に証券コードを入れてから「① 準備」を実行してください。", _
               vbExclamation, "ティック判定"
        Exit Sub
    End If

    ReDim res(1 To shts.Count)

    '--- 判定中は再計算を止める --------------------------------------
    '   RSS 関数を含むブックでは、セルを1つ書くたびにブック全体が
    '   再計算されます（実測で1セルあたり約0.1秒）。判定は RSS の
    '   最新値を必要としないので、ここだけ手動計算にします。
    '   ※ 記録中は絶対に手動にしないこと。RTD の値が反映されなくなります。
    On Error Resume Next
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    calcChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = False
    On Error GoTo Fin

    For i = 1 To shts.Count
        Set ws = shts(i)
        res(i) = JudgeSheet(ws)
        WriteSheetResult ws, res(i)
    Next i

    WriteSummary res

Fin:
    '--- 後始末の前にエラー情報を退避する ----------------------------
    '   On Error Resume Next は Err をクリアするので、先に控えないと
    '   エラーが起きても検知できません。
    errNum = Err.Number
    errDesc = Err.Description

    Application.ScreenUpdating = True
    If calcChanged Then
        On Error Resume Next
        Application.Calculation = prevCalc
        Err.Clear
        On Error GoTo 0
    End If

    If errNum <> 0 Then
        MsgBox "判定中にエラーが発生しました。" & vbCrLf & _
               errNum & " : " & errDesc, vbCritical, "ティック判定"
    End If
End Sub

'------------------------------------------------------------------
' 1銘柄シートを判定する
'------------------------------------------------------------------
Public Function JudgeSheet(ws As Worksheet) As TickJudgeResult

    Dim r As TickJudgeResult
    Dim lastRow As Long, n As Long, i As Long
    Dim src As Variant
    Dim outCalc() As Variant          ' G:I（重み / UpScore / DnScore）書き戻し用
    Dim outDir() As Variant           ' L:M（方向 / 秒内連番）書き戻し用

    Dim sec As Long, startSec As Long, endSec As Long, bucket As Long, bkSec As Long
    Dim price As Double, prevPrice As Double
    Dim bid As Double, ask As Double, prevBid As Double, prevAsk As Double
    Dim vol As Double, w As Long
    Dim upRun As Long, dnRun As Long
    Dim prevSec As Long, secCount As Long
    Dim upScore As Double, dnScore As Double
    Dim buyBid As Long, sellAsk As Long
    Dim havePrev As Boolean
    Dim dirTick As Long
    Dim lastBid As Double, lastAsk As Double
    Dim useMark As Boolean

    r.SheetName = ws.Name
    r.Code = CStr(ws.Cells(ROW_CODE, COL_CODE).Value)
    r.StockName = CStr(ws.Cells(ROW_CODE, COL_NAME).Value)
    r.Judge = JUDGE_FLAT
    r.DirText = JUDGE_FLAT
    r.Confidence = "通常"

    startSec = SecOfText(JUDGE_START)
    endSec = SecOfText(JUDGE_END)
    bkSec = BucketSec()                        ' 判定窓を4等分した1区間の長さ
    useMark = USE_TICK_MARK And (TB_OFF_MARK >= 0)

    lastRow = LastTickRow(ws)
    If lastRow < TICK_FIRST_ROW Then
        r.Judge = JUDGE_NODATA
        r.DirText = "データなし"
        r.Confidence = "－"
        JudgeSheet = r
        Exit Function
    End If

    n = lastRow - TICK_FIRST_ROW + 1
    src = ws.Range(ws.Cells(TICK_FIRST_ROW, COL_TIME), ws.Cells(lastRow, COL_MARK)).Value
    ' src の列は IX_* 定数で参照します（COL_* の並べ替えに自動追従）

    ReDim outCalc(1 To n, 1 To 3)
    ReDim outDir(1 To n, 1 To 2)

    prevSec = -1
    secCount = 0

    For i = 1 To n

        sec = TickSec(src(i, IX_TIME))
        If sec < startSec Or sec > endSec Then GoTo NextI     ' 判定窓のみ

        price = NumOrZero(src(i, IX_PRICE))
        If price <= 0 Then GoTo NextI                          ' 約定値が無い行は無視

        vol = NumOrZero(src(i, IX_VOL))
        bid = NumOrZero(src(i, IX_BID))
        ask = NumOrZero(src(i, IX_ASK))

        r.TickCount = r.TickCount + 1

        '--- ② 出来高初動判定：判定窓を4等分した区間に振り分け -----------
        bucket = (sec - startSec) \ bkSec
        If bucket > 3 Then bucket = 3                          ' 終端ちょうどは第4区間
        Select Case bucket
            Case 0: r.Vol1 = r.Vol1 + vol
            Case 1: r.Vol2 = r.Vol2 + vol
            Case 2: r.Vol3 = r.Vol3 + vol
            Case 3: r.Vol4 = r.Vol4 + vol
        End Select

        '--- ③ 約定速度：同一秒内の約定回数 ------------------------------
        If sec = prevSec Then
            secCount = secCount + 1
        Else
            secCount = 1
            prevSec = sec
        End If
        If secCount > r.SpeedMax Then r.SpeedMax = secCount
        outDir(i, 2) = secCount

        '--- 重み（参考スコア用） ----------------------------------------
        w = VolWeight(vol)
        outCalc(i, 1) = w

        '--- ① 方向判定：ティック記号（↑↓）優先、無ければ前ティック比 -----
        dirTick = 0
        If useMark Then dirTick = MarkDir(src(i, IX_MARK))

        If dirTick = 0 And havePrev Then
            If price > prevPrice Then
                dirTick = 1
            ElseIf price < prevPrice Then
                dirTick = -1
            End If
        End If

        If Not havePrev And dirTick = 0 Then
            outDir(i, 1) = "-"
        Else
            Select Case dirTick
                Case 1
                    upRun = upRun + 1
                    dnRun = 0
                    upScore = upScore + w
                    outDir(i, 1) = ChrW(&H2191)                ' ↑
                Case -1
                    dnRun = dnRun + 1
                    upRun = 0
                    dnScore = dnScore + w
                    outDir(i, 1) = ChrW(&H2193)                ' ↓
                Case Else
                    If RESET_ON_FLAT Then
                        upRun = 0
                        dnRun = 0
                    End If
                    outDir(i, 1) = ChrW(&H2192)                ' →
            End Select

            If upRun > r.UpSeqMax Then r.UpSeqMax = upRun
            If dnRun > r.DnSeqMax Then r.DnSeqMax = dnRun

            '--- 参考：気配の切り上げ／切り下げ --------------------------
            If bid > 0 And prevBid > 0 Then
                If bid > prevBid Then buyBid = buyBid + 1
            End If
            If ask > 0 And prevAsk > 0 Then
                If ask < prevAsk Then sellAsk = sellAsk + 1
            End If
        End If

        If bid > 0 Then lastBid = bid
        If ask > 0 Then lastAsk = ask

        outCalc(i, 2) = upScore
        outCalc(i, 3) = dnScore

        prevPrice = price
        prevBid = bid
        prevAsk = ask
        havePrev = True
NextI:
    Next i

    '--- 計算過程をシートに書き戻す ----------------------------------
    ws.Range(ws.Cells(TICK_FIRST_ROW, COL_W), ws.Cells(lastRow, COL_DN)).Value = outCalc
    ws.Range(ws.Cells(TICK_FIRST_ROW, COL_DIR), ws.Cells(lastRow, COL_SPD)).Value = outDir

    '--- 判定窓の中に1本も無かった場合 --------------------------------
    If r.TickCount = 0 Then
        r.Judge = JUDGE_NODATA
        r.DirText = "判定窓（" & JUDGE_START & "～" & JUDGE_END & "）にティックがありません"
        r.Confidence = "－"
        r.BestBid = NumOrZero(ws.Range(LIVE_BID).Value)
        r.BestAsk = NumOrZero(ws.Range(LIVE_ASK).Value)
        If r.BestBid > 0 And r.BestAsk > 0 Then r.Spread = r.BestAsk - r.BestBid
        JudgeSheet = r
        Exit Function
    End If

    '--- ① 判定 ------------------------------------------------------
    If r.UpSeqMax >= SEQ_MIN And r.UpSeqMax > r.DnSeqMax Then
        r.Direction = 1
        r.DirText = "引け上方向（買い優勢）"
    ElseIf r.DnSeqMax >= SEQ_MIN And r.DnSeqMax > r.UpSeqMax Then
        r.Direction = -1
        r.DirText = "引け下方向（売り優勢）"
    Else
        r.Direction = 0
        r.DirText = "中立（板で最終確認）"
    End If

    '--- ② 判定 : 終盤で大口集中 -------------------------------------
    r.VolDominant = (r.Vol4 > r.Vol3 * VOL_RATIO) And (r.Vol4 > r.Vol2) And (r.Vol4 > r.Vol1)

    '--- ③ 判定 : アルゴ／大口が本気 ---------------------------------
    r.SpeedFast = (r.SpeedMax >= SPEED_MIN)
    r.Confidence = IIf(r.SpeedFast, "高（アルゴ/大口）", "通常")

    '--- 参考スコア（中立時に板を見るときの材料） --------------------
    r.BuyTotal = upScore + buyBid
    r.SellTotal = dnScore + sellAsk

    '--- 最良気配（RSS のライブ値優先。取れなければ最後のティック時点） ---
    r.BestBid = NumOrZero(ws.Range(LIVE_BID).Value)
    r.BestAsk = NumOrZero(ws.Range(LIVE_ASK).Value)
    If r.BestBid <= 0 Then r.BestBid = lastBid
    If r.BestAsk <= 0 Then r.BestAsk = lastAsk
    If r.BestBid > 0 And r.BestAsk > 0 Then r.Spread = r.BestAsk - r.BestBid

    '--- ④ 総合ジャッジ ----------------------------------------------
    If r.Direction = 1 And r.VolDominant Then
        r.Judge = JUDGE_BUY
    ElseIf r.Direction = -1 And r.VolDominant Then
        r.Judge = JUDGE_SELL
    Else
        r.Judge = JUDGE_FLAT
    End If

    JudgeSheet = r
End Function

'------------------------------------------------------------------
' 銘柄シートの結果ブロック（O列ラベル／P列値）を更新
'------------------------------------------------------------------
Public Sub WriteSheetResult(ws As Worksheet, r As TickJudgeResult)

    Dim top As Range
    Dim v() As Variant

    ReDim v(1 To RES_ROWS, 1 To 1)
    Set top = ws.Range(RES_TOP)

    v(1, 1) = r.Judge
    v(2, 1) = r.Confidence
    v(3, 1) = r.BestBid
    v(4, 1) = r.BestAsk
    v(5, 1) = r.Spread
    v(6, 1) = r.DirText
    v(7, 1) = r.UpSeqMax
    v(8, 1) = r.DnSeqMax
    v(9, 1) = r.Vol1
    v(10, 1) = r.Vol2
    v(11, 1) = r.Vol3
    v(12, 1) = r.Vol4
    v(13, 1) = IIf(r.VolDominant, "○ 終盤に大口集中", "×")
    v(14, 1) = r.SpeedMax
    v(15, 1) = r.TickCount
    v(16, 1) = r.BuyTotal
    v(17, 1) = r.SellTotal
    v(18, 1) = Format$(Now, "hh:mm:ss")

    top.Offset(0, 1).Resize(RES_ROWS, 1).Value = v

    PaintJudge top.Offset(0, 1), r.Judge
End Sub

'------------------------------------------------------------------
' Judge_Results 一覧を更新
'------------------------------------------------------------------
Public Sub WriteSummary(res() As TickJudgeResult)

    Dim ws As Worksheet
    Dim i As Long, rw As Long
    Dim nBuy As Long, nSell As Long, nNeutral As Long, nNoData As Long
    Dim body() As Variant
    Dim cnt As Long
    Dim stamp As String

    Set ws = ResultSheet()
    cnt = UBound(res) - LBound(res) + 1
    stamp = Format$(Now, "hh:mm:ss")

    ws.Range(ws.Cells(ROW_HEADER + 1, SUM_COL_FIRST), _
             ws.Cells(ROW_HEADER + SUM_MAX_ROWS, SUM_COL_LAST)).ClearContents
    ws.Range(ws.Cells(ROW_HEADER + 1, SUM_COL_FIRST), _
             ws.Cells(ROW_HEADER + SUM_MAX_ROWS, SUM_COL_LAST)).Interior.Pattern = xlNone

    ReDim body(1 To cnt, 1 To SUM_COL_LAST - SUM_COL_FIRST + 1)

    rw = 0
    For i = LBound(res) To UBound(res)
        rw = rw + 1
        body(rw, 1) = res(i).Code
        body(rw, 2) = res(i).StockName
        body(rw, 3) = res(i).Judge
        body(rw, 4) = res(i).Confidence
        body(rw, 5) = res(i).UpSeqMax
        body(rw, 6) = res(i).DnSeqMax
        body(rw, 7) = res(i).DirText
        body(rw, 8) = res(i).Vol1
        body(rw, 9) = res(i).Vol2
        body(rw, 10) = res(i).Vol3
        body(rw, 11) = res(i).Vol4
        body(rw, 12) = IIf(res(i).VolDominant, "○", "×")
        body(rw, 13) = res(i).SpeedMax
        body(rw, 14) = res(i).TickCount
        body(rw, 15) = res(i).BuyTotal
        body(rw, 16) = res(i).SellTotal
        body(rw, 17) = stamp
        body(rw, 18) = res(i).BestBid
        body(rw, 19) = res(i).BestAsk
        body(rw, 20) = res(i).Spread

        Select Case res(i).Judge
            Case JUDGE_BUY:    nBuy = nBuy + 1
            Case JUDGE_SELL:   nSell = nSell + 1
            Case JUDGE_NODATA: nNoData = nNoData + 1
            Case Else:         nNeutral = nNeutral + 1
        End Select
    Next i

    ws.Cells(ROW_HEADER + 1, SUM_COL_FIRST).Resize(cnt, SUM_COL_LAST - SUM_COL_FIRST + 1).Value = body

    rw = 0
    For i = LBound(res) To UBound(res)
        rw = rw + 1
        PaintJudge ws.Cells(ROW_HEADER + rw, SUM_COL_FIRST + 2), res(i).Judge
    Next i

    ws.Range(RES_STATUS).Value = "判定完了 " & stamp
    PaintStatus ws.Range(RES_STATUS), IIf(nNoData > 0, "warn", "done")

    ws.Range(RES_BREAK).Value = "買い " & nBuy & " / 売り " & nSell & " / 中立 " & nNeutral & _
                                IIf(nNoData > 0, " / 未計測 " & nNoData & "（配信を確認してください）", "")
End Sub

'------------------------------------------------------------------
' 最良気配のリアルタイム表示
'   記録中に Judge_Results の S:U 列を1秒ごとに更新します。
'------------------------------------------------------------------
Public Sub RefreshQuotes(shts As Collection)

    Dim ws As Worksheet, res As Worksheet
    Dim i As Long, rw As Long
    Dim bid As Double, ask As Double
    Dim v(1 To 1, 1 To 3) As Variant

    Set res = ResultSheet()

    For i = 1 To shts.Count
        Set ws = shts(i)
        rw = ROW_HEADER + i

        bid = NumOrZero(ws.Range(LIVE_BID).Value)
        ask = NumOrZero(ws.Range(LIVE_ASK).Value)

        If Len(res.Cells(rw, SUM_COL_FIRST).Value & "") = 0 Then
            res.Cells(rw, SUM_COL_FIRST).Value = ws.Cells(ROW_CODE, COL_CODE).Value
            res.Cells(rw, SUM_COL_FIRST + 1).Value = ws.Cells(ROW_CODE, COL_NAME).Value
        End If

        v(1, 1) = bid
        v(1, 2) = ask
        If bid > 0 And ask > 0 Then
            v(1, 3) = ask - bid
        Else
            v(1, 3) = Empty
        End If
        res.Cells(rw, SUM_COL_LAST - 2).Resize(1, 3).Value = v
    Next i
End Sub

'------------------------------------------------------------------
' 分別の出来高プロファイル（判定窓を決めるための診断）
'------------------------------------------------------------------
Public Sub ShowVolumeProfile()

    Dim ws As Worksheet, shts As Collection
    Dim lastRow As Long, n As Long, i As Long
    Dim src As Variant
    Dim sec As Long, mi As Long
    Dim vol(0 To 1439) As Double
    Dim cnt(0 To 1439) As Long
    Dim first As Long, last As Long
    Dim total As Double
    Dim msg As String

    If BlockedWhileLogging("出来高プロファイル") Then Exit Sub

    Set ws = ActiveSheet
    If ws.Name = RESULT_SHEET Then
        Set shts = TargetSheets()
        If shts.Count = 0 Then Exit Sub
        Set ws = shts(1)
    End If

    lastRow = LastTickRow(ws)
    If lastRow < TICK_FIRST_ROW Then
        MsgBox "ティックが記録されていません。", vbExclamation, "出来高プロファイル"
        Exit Sub
    End If

    n = lastRow - TICK_FIRST_ROW + 1
    src = ws.Range(ws.Cells(TICK_FIRST_ROW, COL_TIME), ws.Cells(lastRow, COL_MARK)).Value

    first = 9999
    last = -1

    For i = 1 To n
        sec = TickSec(src(i, IX_TIME))
        If sec >= 0 Then
            mi = sec \ 60
            vol(mi) = vol(mi) + NumOrZero(src(i, IX_VOL))
            cnt(mi) = cnt(mi) + 1
            total = total + NumOrZero(src(i, IX_VOL))
            If mi < first Then first = mi
            If mi > last Then last = mi
        End If
    Next i

    If last < 0 Then
        MsgBox "時刻を読める行がありませんでした。", vbExclamation, "出来高プロファイル"
        Exit Sub
    End If

    msg = ws.Name & " / " & ws.Cells(ROW_CODE, COL_NAME).Value & vbCrLf & _
          "記録範囲 " & SecHM(first * 60) & " ～ " & SecHM(last * 60) & _
          " / 合計 " & Format$(total, "#,##0") & " 株" & vbCrLf & vbCrLf & _
          "時刻    出来高        ティック   構成比" & vbCrLf & _
          String$(42, "-") & vbCrLf

    For mi = first To last
        If cnt(mi) > 0 Or vol(mi) > 0 Then
            msg = msg & SecHM(mi * 60) & "  " & _
                  Right$(Space$(11) & Format$(vol(mi), "#,##0"), 11) & "  " & _
                  Right$(Space$(8) & cnt(mi), 8) & "  " & _
                  Right$(Space$(6) & Format$(IIf(total > 0, vol(mi) / total, 0), "0.0%"), 6) & vbCrLf
        End If
    Next mi

    msg = msg & vbCrLf & _
          "現在の判定窓 : " & JUDGE_START & " ～ " & JUDGE_END & vbCrLf & _
          "1区間の長さ  : " & (BucketSec() \ 60) & "分" & (BucketSec() Mod 60) & "秒"

    MsgBox msg, vbInformation, "出来高プロファイル（判定窓の検討用）"
End Sub

'------------------------------------------------------------------
' ティック記号（↑↓）を方向に変換  1=上 / -1=下 / 0=不明・変わらず
'------------------------------------------------------------------
Public Function MarkDir(ByVal v As Variant) As Long

    Dim s As String

    On Error GoTo Done
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    Select Case True
        Case InStr(s, ChrW(&H2191)) > 0, InStr(s, "+") > 0, _
             UCase$(s) = "U", s = "上", s = "買", InStr(s, ChrW(&HFF0B)) > 0
            MarkDir = 1
        Case InStr(s, ChrW(&H2193)) > 0, InStr(s, "-") > 0, _
             UCase$(s) = "D", s = "下", s = "売", InStr(s, ChrW(&HFF0D)) > 0
            MarkDir = -1
    End Select

Done:
End Function

'------------------------------------------------------------------
' 判定セルの色付け
'------------------------------------------------------------------
Public Sub PaintJudge(cel As Range, ByVal Judge As String)

    With cel
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        Select Case Judge
            Case JUDGE_BUY
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            Case JUDGE_SELL
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
            Case JUDGE_NODATA
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 87, 0)
            Case Else
                .Interior.Color = RGB(235, 235, 235)
                .Font.Color = RGB(80, 80, 80)
        End Select
    End With
End Sub
