Attribute VB_Name = "Module_Judge"
Option Explicit

'==================================================================
' Module_Judge : ティック判定ロジック本体（15:00～15:20専用）
'
'   ① 方向判定（連続ティック）
'   ② 出来高初動判定（大口の本気）
'   ③ 約定速度（Tick Speed）
'   ④ 総合ジャッジ → 買い / 売り / 中立
'==================================================================

'------------------------------------------------------------------
' 全銘柄を判定して Judge_Results にまとめる（ボタン④）
'------------------------------------------------------------------
Public Sub JudgeAll()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim res() As TickJudgeResult
    Dim i As Long

    Set shts = TargetSheets()

    If shts.Count = 0 Then
        MsgBox "判定対象の銘柄シートがありません。" & vbCrLf & _
               "各シートの B3 に証券コードを入れてから「① 準備」を実行してください。", _
               vbExclamation, "ティック判定"
        Exit Sub
    End If

    ReDim res(1 To shts.Count)

    Application.ScreenUpdating = False
    On Error GoTo Fin

    For i = 1 To shts.Count
        Set ws = shts(i)
        res(i) = JudgeSheet(ws)
        WriteSheetResult ws, res(i)
    Next i

    WriteSummary res

Fin:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "判定中にエラーが発生しました。" & vbCrLf & _
               Err.Number & " : " & Err.Description, vbCritical, "ティック判定"
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

    r.SheetName = ws.Name
    r.Code = CStr(ws.Cells(ROW_CODE, COL_CODE).Value)
    r.StockName = CStr(ws.Cells(ROW_CODE, COL_NAME).Value)
    r.Judge = "中立"
    r.DirText = "中立"
    r.Confidence = "通常"

    startSec = SecOfText(JUDGE_START)
    endSec = SecOfText(JUDGE_END)
    bkSec = BucketSec()                        ' 判定窓を4等分した1区間の長さ

    lastRow = LastTickRow(ws)
    If lastRow < TICK_FIRST_ROW Then
        r.DirText = "データなし"
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
        If sec < startSec Or sec > endSec Then GoTo NextI     ' 15:00～15:20 のみ

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
        If USE_TICK_MARK Then dirTick = MarkDir(src(i, IX_MARK))

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
        r.Judge = "買い"
    ElseIf r.Direction = -1 And r.VolDominant Then
        r.Judge = "売り"
    Else
        r.Judge = "中立"
    End If

    '--- 計算過程をシートに書き戻す ----------------------------------
    ws.Range(ws.Cells(TICK_FIRST_ROW, COL_W), ws.Cells(lastRow, COL_DN)).Value = outCalc
    ws.Range(ws.Cells(TICK_FIRST_ROW, COL_DIR), ws.Cells(lastRow, COL_SPD)).Value = outDir

    JudgeSheet = r
End Function

'------------------------------------------------------------------
' 銘柄シートの結果ブロック（O列ラベル／P列値）を更新
'------------------------------------------------------------------
Public Sub WriteSheetResult(ws As Worksheet, r As TickJudgeResult)

    Dim top As Range
    Set top = ws.Range(RES_TOP)

    top.Offset(0, 1).Value = r.Judge
    top.Offset(1, 1).Value = r.Confidence
    top.Offset(2, 1).Value = r.BestBid
    top.Offset(3, 1).Value = r.BestAsk
    top.Offset(4, 1).Value = r.Spread
    top.Offset(5, 1).Value = r.DirText
    top.Offset(6, 1).Value = r.UpSeqMax
    top.Offset(7, 1).Value = r.DnSeqMax
    top.Offset(8, 1).Value = r.Vol1
    top.Offset(9, 1).Value = r.Vol2
    top.Offset(10, 1).Value = r.Vol3
    top.Offset(11, 1).Value = r.Vol4
    top.Offset(12, 1).Value = IIf(r.VolDominant, "○ 終盤に大口集中", "×")
    top.Offset(13, 1).Value = r.SpeedMax
    top.Offset(14, 1).Value = r.TickCount
    top.Offset(15, 1).Value = r.BuyTotal
    top.Offset(16, 1).Value = r.SellTotal
    top.Offset(17, 1).Value = Format$(Now, "hh:mm:ss")

    PaintJudge top.Offset(0, 1), r.Judge
End Sub

'------------------------------------------------------------------
' Judge_Results 一覧を更新
'------------------------------------------------------------------
Public Sub WriteSummary(res() As TickJudgeResult)

    Dim ws As Worksheet
    Dim i As Long, rw As Long
    Dim nBuy As Long, nSell As Long, nNeutral As Long

    Set ws = ResultSheet()

    ws.Range(ws.Cells(ROW_HEADER + 1, 2), ws.Cells(ROW_HEADER + 200, 21)).ClearContents
    ws.Range(ws.Cells(ROW_HEADER + 1, 2), ws.Cells(ROW_HEADER + 200, 21)).Interior.Pattern = xlNone

    rw = ROW_HEADER + 1

    For i = LBound(res) To UBound(res)
        With ws
            .Cells(rw, 2).Value = res(i).Code
            .Cells(rw, 3).Value = res(i).StockName
            .Cells(rw, 4).Value = res(i).Judge
            .Cells(rw, 5).Value = res(i).Confidence
            .Cells(rw, 6).Value = res(i).UpSeqMax
            .Cells(rw, 7).Value = res(i).DnSeqMax
            .Cells(rw, 8).Value = res(i).DirText
            .Cells(rw, 9).Value = res(i).Vol1
            .Cells(rw, 10).Value = res(i).Vol2
            .Cells(rw, 11).Value = res(i).Vol3
            .Cells(rw, 12).Value = res(i).Vol4
            .Cells(rw, 13).Value = IIf(res(i).VolDominant, "○", "×")
            .Cells(rw, 14).Value = res(i).SpeedMax
            .Cells(rw, 15).Value = res(i).TickCount
            .Cells(rw, 16).Value = res(i).BuyTotal
            .Cells(rw, 17).Value = res(i).SellTotal
            .Cells(rw, 18).Value = Format$(Now, "hh:mm:ss")
            .Cells(rw, 19).Value = res(i).BestBid
            .Cells(rw, 20).Value = res(i).BestAsk
            .Cells(rw, 21).Value = res(i).Spread
        End With

        PaintJudge ws.Cells(rw, 4), res(i).Judge

        Select Case res(i).Judge
            Case "買い": nBuy = nBuy + 1
            Case "売り": nSell = nSell + 1
            Case Else:   nNeutral = nNeutral + 1
        End Select

        rw = rw + 1
    Next i

    ws.Range(RES_STATUS).Value = "判定完了 " & Format$(Now, "hh:mm:ss")
    ws.Range(RES_BREAK).Value = "買い " & nBuy & " / 売り " & nSell & " / 中立 " & nNeutral
End Sub

'------------------------------------------------------------------
' 最良気配のリアルタイム表示
'   記録中に Judge_Results の S:U 列を1秒ごとに更新します。
'------------------------------------------------------------------
Public Sub RefreshQuotes(shts As Collection)

    Dim ws As Worksheet, res As Worksheet
    Dim i As Long, rw As Long
    Dim bid As Double, ask As Double

    Set res = ResultSheet()

    For i = 1 To shts.Count
        Set ws = shts(i)
        rw = ROW_HEADER + i

        bid = NumOrZero(ws.Range(LIVE_BID).Value)
        ask = NumOrZero(ws.Range(LIVE_ASK).Value)

        If Len(res.Cells(rw, 2).Value & "") = 0 Then
            res.Cells(rw, 2).Value = ws.Cells(ROW_CODE, COL_CODE).Value
            res.Cells(rw, 3).Value = ws.Cells(ROW_CODE, COL_NAME).Value
        End If

        res.Cells(rw, 19).Value = bid
        res.Cells(rw, 20).Value = ask
        If bid > 0 And ask > 0 Then
            res.Cells(rw, 21).Value = ask - bid
        Else
            res.Cells(rw, 21).Value = ""
        End If
    Next i
End Sub

'------------------------------------------------------------------
' 分別の出来高プロファイル（判定窓を決めるための診断）
'   記録済みティックを1分ごとに集計します。判定窓の外も表示するので、
'   「何分から大口が入るか」を自分のデータで確かめて窓を決められます。
'------------------------------------------------------------------
Public Sub ShowVolumeProfile()

    Dim ws As Worksheet, shts As Collection
    Dim lastRow As Long, n As Long, i As Long
    Dim src As Variant
    Dim sec As Long, mi As Long, base As Long
    Dim vol(0 To 1439) As Double
    Dim cnt(0 To 1439) As Long
    Dim first As Long, last As Long
    Dim total As Double
    Dim msg As String

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
Private Sub PaintJudge(cel As Range, ByVal judge As String)
    With cel
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        Select Case judge
            Case "買い"
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            Case "売り"
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
            Case Else
                .Interior.Color = RGB(235, 235, 235)
                .Font.Color = RGB(80, 80, 80)
        End Select
    End With
End Sub
