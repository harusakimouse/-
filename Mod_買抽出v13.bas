Attribute VB_Name = "Mod_買抽出v13"
' ==================================================================
' ★ 100番ブック / Mod_買抽出v13  （2026/07/29 絞り込み強化版）
'   変更点：①MAX_EXTRACT_COUNT追加 ②必須フィルタ3つON
'           ③高精度しきい値12→16 ④ソート後に上位N件だけ残す
' ==================================================================

Option Explicit
'=============================================================================
' 打ち出の無限こづち　候補抽出モジュール　v13 (2026/06/17) → v13.1 絞り込み強化
'=============================================================================
Public Const MAX_STOCKS       As Long = 500
Public Const ANA_DATA_START   As Long = 3
Public Const STOCK_DATA_START As Long = 6
Public Const MEI_DATA_START   As Long = 5

'★①追加：1日の最大抽出件数（1にすれば1銘柄だけ）
Public Const MAX_EXTRACT_COUNT As Long = 2   '高精度抽出の上限（TOPからN件）
Public Const MAX_EXTRACT_BUY   As Long = 5   '買い抽出の上限（TOPからN件）
Public Const MAX_EXTRACT_EASE  As Long = 10  '緩和抽出の上限（TOPからN件）

Public Const TOPIX_SMALL_DROP_PCT As Double = -0.5
Public Const TOPIX_LARGE_DROP_PCT As Double = -1.5
Public Const BUY_REQUIRE_MOMENTUM_COUNT As Long = 2
Public Const MOMENTUM_FULL_BONUS    As Long = 2
Public Const MOMENTUM_PARTIAL_BONUS As Long = 1

'★②ここを変更：必須フィルタを3つON
Public Const BUY_REQUIRE_GOLDEN_CROSS  As Boolean = False
Public Const BUY_REQUIRE_PERFECT_ORDER As Boolean = True   'False→True
Public Const BUY_REQUIRE_VOL_BURST     As Boolean = True   'False→True
Public Const BUY_REQUIRE_ADX           As Boolean = False
Public Const BUY_REQUIRE_HIST_EXPAND   As Boolean = True   'False→True
Public Const BUY_REQUIRE_RS_VS_TOPIX   As Boolean = False

Public Const GOLDEN_CROSS_LOOKBACK As Long = 3
Public Const VOLUME_BURST_RATIO    As Double = 1.5
Public Const ADX_THRESHOLD         As Double = 25#
Public Const GC_BONUS     As Long = 2
Public Const PO_BONUS     As Long = 2
Public Const VOL_BONUS    As Long = 2
Public Const ADX_BONUS    As Long = 1
Public Const HIST_BONUS   As Long = 1
Public Const RS_BONUS     As Long = 1
Public Const CANDLE_LOWER_TAIL_YOSEN As Long = 2
Public Const CANDLE_BIG_YOSEN        As Long = 2
Public Const CANDLE_LONG_LOWER_TAIL  As Long = 1
Public Const CANDLE_UPPER_TAIL_WARN  As Long = -2
Private Const COL_NAVY  As Long = 6299648
Private Const COL_WHITE As Long = 16777215
Private Const COL_BLACK As Long = 2105376
Private Const EXT_DATA_ROW As Long = 4
Private Const EXT_LAST_COL As Long = 18
Public Const BUY_REQUIRE_OSIME As Boolean = True
Public Const FAKE_BREAKOUT_VOL_RATIO As Double = 1.5

'=============================================================================
' ボタン設置マクロ
Public Sub 買抽出v13ボタン設置()
    Dim ws As Worksheet: Set ws = GetWs("買抽出v13")
    If ws Is Nothing Then
        UI_Msg "「買抽出v13」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp
    Dim btn1 As Shape
    Set btn1 = ws.Shapes.AddFormControl(xlButtonControl, 10, 5, 180, 35)
    With btn1
        .Name = "btn高精度抽出"
        .TextFrame.Characters.Text = "★高精度抽出(16以上/上位" & MAX_EXTRACT_COUNT & "件)"
        .TextFrame.Characters.Font.Size = 14
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "Mod_買抽出v13.株価抽出"
    End With
    Dim btn2 As Shape
    Set btn2 = ws.Shapes.AddFormControl(xlButtonControl, 200, 5, 180, 35)
    With btn2
        .Name = "btn通常抽出"
        .TextFrame.Characters.Text = "買い抽出(11以上)"
        .TextFrame.Characters.Font.Size = 16
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "Mod_買抽出v13.買い抽出_スコア4"
    End With
    Dim btn3 As Shape
    Set btn3 = ws.Shapes.AddFormControl(xlButtonControl, 390, 5, 180, 35)
    With btn3
        .Name = "btn緩和抽出"
        .TextFrame.Characters.Text = "緩和抽出(11以上・緩和)"
        .TextFrame.Characters.Font.Size = 16
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "Mod_買抽出v13.買い抽出_スコア4_緩和"
    End With
    UI_Msg "ボタンを再設置しました！", vbInformation, "完了"
End Sub

'=============================================================================
' 抽出メイン
'=============================================================================
Public Sub 株価抽出()
    Dim P As Boolean: P = gUnattended
    gUnattended = True
    On Error GoTo done
    Dim h As Integer: h = Hour(Now)
    Dim m As Integer: m = Minute(Now)
    If (h >= 9 And h < 15) Or (h = 15 And m <= 30) Then
        '★③しきい値 12→16
        Call 抽出実行_RSS(16, True, "高精度抽出[RSS現在値](スコア16以上・上位" & MAX_EXTRACT_COUNT & "件)")
    Else
        Call 抽出実行(16, True, "高精度抽出(スコア16以上・全フィルター必須・上位" & MAX_EXTRACT_COUNT & "件)")
    End If
done:
    gUnattended = P
End Sub

Public Sub 買い抽出_スコア4()
    抽出実行 11, True, "買い抽出(スコア11以上・中程度・上位5件)", False, MAX_EXTRACT_BUY
End Sub

Public Sub 買い抽出_スコア4_緩和()
    抽出実行 11, False, "買い抽出(スコア11以上・緩和・上位10件)", False, MAX_EXTRACT_EASE
End Sub

Private Sub 抽出実行(ByVal minScore As Long, _
                     ByVal doMark As Boolean, _
                     ByVal titleText As String, _
                     Optional ByVal useRSS As Boolean = False, _
                     Optional ByVal maxCount As Long = 0)
    If maxCount <= 0 Then maxCount = MAX_EXTRACT_COUNT
    Dim closeWs As Worksheet: Set closeWs = GetWs("終値")
    Dim openWs  As Worksheet: Set openWs = GetWs("始値")
    Dim highWs  As Worksheet: Set highWs = GetWs("高値")
    Dim lowWs   As Worksheet: Set lowWs = GetWs("安値")
    Dim volWs   As Worksheet: Set volWs = GetWs("出来高")
    Dim extWs   As Worksheet: Set extWs = GetWs("買抽出v13")
    Dim mws     As Worksheet: Set mws = GetWs("銘柄管理")
    Dim anaWs   As Worksheet: Set anaWs = GetWs("分析")
    If closeWs Is Nothing Or extWs Is Nothing Then
        UI_Msg "「終値」または「買抽出v13」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    If mws Is Nothing Then
        UI_Msg "「銘柄管理」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    If Not anaWs Is Nothing Then anaWs.Calculate
    closeWs.Calculate
    On Error GoTo 0
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error Resume Next
    extWs.Unprotect Password:="ne19480314"
    mws.Unprotect Password:="ne19480314"
    If Not anaWs Is Nothing Then anaWs.Unprotect Password:="ne19480314"
    On Error GoTo 0
    Dim extLast As Long
    extLast = extWs.Cells(extWs.Rows.Count, 2).End(xlUp).Row
    Dim clearLast As Long
    clearLast = WorksheetFunction.Max(extLast, EXT_DATA_ROW + MAX_STOCKS)
    With extWs.Range(extWs.Cells(EXT_DATA_ROW, 1), extWs.Cells(clearLast, EXT_LAST_COL))
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With
    Dim topixTrend As String: topixTrend = CStr(mws.Cells(5, 6).Value)
    Dim topixNow As Double, topixRef As Double, topixPct As Double
    topixNow = SafeNum(mws.Cells(5, 3).Value, 0)
    topixRef = SafeNum(mws.Cells(5, 5).Value, 0)
    If topixRef > 0 And topixNow > 0 Then
        topixPct = (topixNow - topixRef) / topixRef * 100#
    Else
        topixPct = 0
    End If
    Dim topixMode As String
    If topixPct <= TOPIX_LARGE_DROP_PCT Then
        topixMode = "LARGE"
    ElseIf topixPct <= TOPIX_SMALL_DROP_PCT Then
        topixMode = "MID"
    Else
        topixMode = "OK"
    End If
    ' Futures filter: H5 = RssIndexMarket(N225.FUT01.OS, prev day %)
    Dim futuresPct As Double
    futuresPct = SafeNum(mws.Cells(5, 8).Value, 0)
    If futuresPct < -1.5 Then
        topixMode = "LARGE"
    ElseIf futuresPct < -0.5 Then
        If topixMode = "OK" Then topixMode = "MID"
    End If
    Dim nameDic As Object: Set nameDic = CreateObject("Scripting.Dictionary")
    Dim mLast As Long
    mLast = mws.Cells(mws.Rows.Count, 2).End(xlUp).Row
    If mLast > MEI_DATA_START + MAX_STOCKS Then mLast = MEI_DATA_START + MAX_STOCKS
    Dim mr As Long, mCode As String, mName As String
    For mr = MEI_DATA_START To mLast
        mCode = Trim$(CStr(mws.Cells(mr, 2).Value))
        mName = Trim$(CStr(mws.Cells(mr, 3).Value))
        If mCode <> "" And Not nameDic.Exists(mCode) Then nameDic.Add mCode, mName
    Next mr
    Dim anaDic As Object: Set anaDic = CreateObject("Scripting.Dictionary")
    If Not anaWs Is Nothing Then
        Dim anaLast As Long
        anaLast = anaWs.Cells(anaWs.Rows.Count, 3).End(xlUp).Row
        If anaLast < ANA_DATA_START Then anaLast = ANA_DATA_START
        If anaLast > ANA_DATA_START + MAX_STOCKS Then anaLast = ANA_DATA_START + MAX_STOCKS
        Dim k As Long, anaCode As String
        For k = ANA_DATA_START To anaLast
            anaCode = Trim$(CStr(anaWs.Cells(k, 3).Value))
            If anaCode <> "" And Not anaDic.Exists(anaCode) Then anaDic(anaCode) = k
        Next k
    End If
    Dim lastRow As Long
    lastRow = closeWs.Cells(closeWs.Rows.Count, 1).End(xlUp).Row
    If lastRow > STOCK_DATA_START + MAX_STOCKS Then lastRow = STOCK_DATA_START + MAX_STOCKS
    Dim outRow As Long: outRow = EXT_DATA_ROW
    Dim totalScore As Long, signals As String
    Dim scannedCount As Long: scannedCount = 0
    Dim i As Long
    Dim cntGC As Long, cntPO As Long, cntVB As Long
    Dim cntADX As Long, cntHE As Long, cntRS As Long
    For i = STOCK_DATA_START To lastRow
        Dim code As String
        code = Trim$(CStr(closeWs.Cells(i, 1).Value))
        If code = "" Or code = "0" Or code = "TOPX" Then GoTo NextStock
        scannedCount = scannedCount + 1
        Application.StatusBar = titleText & " : " & scannedCount & " 件目 (" & code & ")"
        totalScore = 0: signals = ""
        Dim closePrice As Double
        Dim highVal As Double: highVal = 0
        Dim lowVal  As Double: lowVal = 0
        Dim closes As Variant, opens As Variant, highs As Variant, lows As Variant, vols As Variant
        If useRSS Then
            '★ RSS現在値(C列)を先頭にした配列
            Dim rssClose As Double: rssClose = val(closeWs.Cells(i, 3).Value)
            closePrice = IIf(rssClose > 0, rssClose, val(closeWs.Cells(i, 5).Value))
            If Not highWs Is Nothing Then
                Dim rssH As Double: rssH = val(highWs.Cells(i, 3).Value)
                highVal = IIf(rssH > 0, rssH, val(highWs.Cells(i, 5).Value))
            End If
            If Not lowWs Is Nothing Then
                Dim rssL As Double: rssL = val(lowWs.Cells(i, 3).Value)
                lowVal = IIf(rssL > 0, rssL, val(lowWs.Cells(i, 5).Value))
            End If
            closes = RssMerge(closeWs, i)
            If Not openWs Is Nothing Then opens = RssMerge(openWs, i)
            If Not highWs Is Nothing Then highs = RssMerge(highWs, i)
            If Not lowWs Is Nothing Then lows = RssMerge(lowWs, i)
            If Not volWs Is Nothing Then vols = RssMerge(volWs, i)
        Else
            closePrice = val(closeWs.Cells(i, 5).Value)
            If Not highWs Is Nothing Then highVal = val(highWs.Cells(i, 5).Value)
            If Not lowWs Is Nothing Then lowVal = val(lowWs.Cells(i, 5).Value)
            closes = closeWs.Range(closeWs.Cells(i, 5), closeWs.Cells(i, 104)).Value
            If Not openWs Is Nothing Then opens = openWs.Range(openWs.Cells(i, 5), openWs.Cells(i, 104)).Value
            If Not highWs Is Nothing Then highs = highWs.Range(highWs.Cells(i, 5), highWs.Cells(i, 104)).Value
            If Not lowWs Is Nothing Then lows = lowWs.Range(lowWs.Cells(i, 5), lowWs.Cells(i, 104)).Value
            If Not volWs Is Nothing Then vols = volWs.Range(volWs.Cells(i, 5), volWs.Cells(i, 104)).Value
        End If
        If closePrice = 0 Then GoTo NextStock
        Dim anaRow As Long: anaRow = 0
        If anaDic.Exists(code) Then anaRow = CLng(anaDic(code))
        Dim rsiVal As Double: rsiVal = 50
        If anaRow > 0 Then rsiVal = SafeNum(anaWs.Cells(anaRow, 23).Value, 50)
        If rsiVal <= 30 Then
            totalScore = totalScore + 2: signals = signals & "RSI売られ過ぎ "
        ElseIf rsiVal <= 40 Then
            totalScore = totalScore + 1: signals = signals & "RSI安 "
        ElseIf rsiVal >= 70 Then
            totalScore = totalScore - 1: signals = signals & "RSI過熱 "
        End If
        Dim emaTrend As String: emaTrend = ""
        If anaRow > 0 Then emaTrend = SafeStr(anaWs.Cells(anaRow, 30).Value)
        If emaTrend = "パーフェクト▲" Then
            totalScore = totalScore + 2: signals = signals & "パーフェクト▲ "
        ElseIf emaTrend = "上昇トレンド▲" Then
            totalScore = totalScore + 1: signals = signals & "上昇▲ "
        ElseIf emaTrend = "下降トレンド▼" Then
            totalScore = totalScore - 1
        End If
        Dim osime As String: osime = ""
        If anaRow > 0 Then osime = SafeStr(anaWs.Cells(anaRow, 31).Value)
        If osime = "25日押し目★" Then
            totalScore = totalScore + 2: signals = signals & "25日押し目● "
        ElseIf osime = "5日押し目○" Then
            totalScore = totalScore + 1: signals = signals & "5日押し目○ "
        End If
        Dim volToday As Double: volToday = 0
        Dim vol5avg  As Double: vol5avg = 0
        Dim volRatio As Double: volRatio = 0
        If Not volWs Is Nothing Then
            volToday = val(volWs.Cells(i, 5).Value)
            If anaRow > 0 Then vol5avg = SafeNum(anaWs.Cells(anaRow, 13).Value, 0)
            If vol5avg > 0 And volToday > 0 Then
                volRatio = volToday / vol5avg
                If volRatio >= 2# Then
                    totalScore = totalScore + 3
                    signals = signals & "出来高急増(" & Format(volRatio, "0.0") & "倍) "
                ElseIf volRatio >= 1.5 Then
                    totalScore = totalScore + 2
                    signals = signals & "出来高増(" & Format(volRatio, "0.0") & "倍) "
                End If
            End If
        End If
        Dim macdVal As Double: macdVal = 0
        If anaRow > 0 Then macdVal = SafeNum(anaWs.Cells(anaRow, 24).Value, 0)
        If macdVal > 0 Then
            totalScore = totalScore + 1: signals = signals & "MACD+ "
        ElseIf macdVal < 0 Then
            totalScore = totalScore - 1
        End If
        Dim highUpdate As String: highUpdate = ""
        If anaRow > 0 Then highUpdate = SafeStr(anaWs.Cells(anaRow, 22).Value)
        Dim isFakeBreakout As Boolean: isFakeBreakout = False
        If highUpdate = "年初来高値更新●" Then
            If (volRatio > 0 And volRatio < FAKE_BREAKOUT_VOL_RATIO) Or _
               (highVal > 0 And closePrice < highVal * 0.95) Then
                isFakeBreakout = True
                signals = signals & "高値更新ダマシ警戒 "
            Else
                totalScore = totalScore + 2
                signals = signals & "年初来高値更新● "
            End If
        End If
        Dim ema12V As Double, ema26V As Double, macdNow As Double, signalNow As Double
        Dim histNow As Double, histPrev As Double
        ComputeMACDSeries closes, ema12V, ema26V, macdNow, signalNow, histNow, histPrev
        Dim emaUp As Boolean:  emaUp = (ema12V > ema26V) And (ema26V > 0)
        Dim macdUp As Boolean: macdUp = (macdNow > signalNow)
        Dim rsiUp As Boolean:  rsiUp = (rsiVal > 50)
        Dim momCount As Long: momCount = 0
        If emaUp Then momCount = momCount + 1: signals = signals & "EMA12>26 "
        If macdUp Then momCount = momCount + 1: signals = signals & "MACD>Sig "
        If rsiUp Then momCount = momCount + 1: signals = signals & "RSI>50 "
        If momCount = 3 Then
            totalScore = totalScore + MOMENTUM_FULL_BONUS
            signals = signals & "[◎モメンタム完全] "
        ElseIf momCount = 2 Then
            totalScore = totalScore + MOMENTUM_PARTIAL_BONUS
            signals = signals & "[○モメンタム部分] "
        End If
        Dim ok_GC As Boolean, ok_PO As Boolean, ok_VB As Boolean
        Dim ok_ADX As Boolean, ok_HE As Boolean, ok_RS As Boolean
        Dim adxVal As Double: adxVal = 0
        ok_GC = IsGoldenCrossRecent(closes, GOLDEN_CROSS_LOOKBACK)
        If ok_GC Then
            totalScore = totalScore + GC_BONUS
            signals = signals & "[GC直近" & GOLDEN_CROSS_LOOKBACK & "日] "
            cntGC = cntGC + 1
        End If
        ok_PO = IsPerfectOrder(closes)
        If ok_PO Then
            totalScore = totalScore + PO_BONUS
            signals = signals & "[Perfect5>25>75] "
            cntPO = cntPO + 1
        End If
        If Not IsEmpty(opens) And Not IsEmpty(vols) Then
            ok_VB = IsVolumeBurstBullish(vols, closes, opens, VOLUME_BURST_RATIO)
            If ok_VB Then
                totalScore = totalScore + VOL_BONUS
                signals = signals & "[Vol急増陽線] "
                cntVB = cntVB + 1
            End If
        End If
        If Not IsEmpty(highs) And Not IsEmpty(lows) Then
            adxVal = ComputeADX(highs, lows, closes)
            ok_ADX = (adxVal > ADX_THRESHOLD)
            If ok_ADX Then
                totalScore = totalScore + ADX_BONUS
                signals = signals & "[ADX=" & Format(adxVal, "0") & "] "
                cntADX = cntADX + 1
            End If
        End If
        ok_HE = (histNow > histPrev) And (histNow > 0)
        If ok_HE Then
            totalScore = totalScore + HIST_BONUS
            signals = signals & "[Hist拡大] "
            cntHE = cntHE + 1
        End If
        Dim stockPctChg As Double: stockPctChg = 0
        If IsArrOk(closes) Then
            If UBound(closes, 2) >= 2 Then
                Dim prevClose As Double: prevClose = SafeNum(closes(1, 2), 0)
                If prevClose > 0 Then stockPctChg = (closePrice - prevClose) / prevClose * 100#
            End If
        End If
        ok_RS = (stockPctChg > topixPct)
        If ok_RS Then
            totalScore = totalScore + RS_BONUS
            signals = signals & "[対TOPIX勝 +" & Format(stockPctChg - topixPct, "0.00") & "%] "
            cntRS = cntRS + 1
        End If
        totalScore = totalScore + CandleScoreBuy(opens, highs, lows, closes, signals)
        totalScore = totalScore + CandleScoreBuyPrev(opens, highs, lows, closes, signals)
        Dim canExtract As Boolean: canExtract = False
        If totalScore >= minScore Then
            Select Case topixMode
                Case "LARGE"
                    If totalScore >= minScore + 3 Then canExtract = True
                Case "MID"
                    If totalScore >= minScore + 1 Then canExtract = True
                Case Else
                    canExtract = True
            End Select
        End If
        If canExtract And momCount < 2 Then canExtract = False
        If doMark And canExtract Then
            If BUY_REQUIRE_GOLDEN_CROSS And Not ok_GC Then canExtract = False
            If BUY_REQUIRE_PERFECT_ORDER And Not ok_PO Then canExtract = False
            If BUY_REQUIRE_VOL_BURST And Not ok_VB Then canExtract = False
            If BUY_REQUIRE_ADX And Not ok_ADX Then canExtract = False
            If BUY_REQUIRE_HIST_EXPAND And Not ok_HE Then canExtract = False
            If BUY_REQUIRE_RS_VS_TOPIX And Not ok_RS Then canExtract = False
        End If
        If doMark And canExtract And BUY_REQUIRE_OSIME Then
            If osime <> "25日押し目★" And osime <> "5日押し目○" Then
                canExtract = False
            End If
        End If
        If rsiVal >= 70 Then canExtract = False
        If canExtract And Not ok_RS Then canExtract = False
        If canExtract Then
            Dim meiName As String
            If nameDic.Exists(code) Then
                meiName = CStr(nameDic(code))
            Else
                meiName = SafeStr(closeWs.Cells(i, 2).Value)
            End If
            If meiName = "" Then meiName = "(" & code & ")"
            Dim osimeOut As String:      osimeOut = osime
            If osimeOut = "" Then osimeOut = "-"
            Dim highUpdateOut As String: highUpdateOut = highUpdate
            If highUpdateOut = "" Then highUpdateOut = "-"
            Dim currentPrice As Double: currentPrice = val(closeWs.Cells(i, 3).Value)
            If currentPrice <= 0 Then currentPrice = closePrice
            Dim currentHigh As Double: currentHigh = highVal
            If Not highWs Is Nothing Then
                Dim tmpH As Double: tmpH = val(highWs.Cells(i, 3).Value)
                If tmpH > 0 Then currentHigh = tmpH
            End If
            Dim currentLow As Double: currentLow = lowVal
            If Not lowWs Is Nothing Then
                Dim tmpL As Double: tmpL = val(lowWs.Cells(i, 3).Value)
                If tmpL > 0 Then currentLow = tmpL
            End If
            Dim currentVol As Double: currentVol = volToday
            If Not volWs Is Nothing Then
                Dim tmpV As Double: tmpV = val(volWs.Cells(i, 3).Value)
                If tmpV > 0 Then currentVol = tmpV
            End If
            writeRow extWs, outRow, code, meiName, currentPrice, currentHigh, currentLow, _
                     currentVol, rsiVal, emaTrend, osimeOut, volRatio, macdVal, _
                     highUpdateOut, isFakeBreakout, totalScore, signals
            outRow = outRow + 1
        End If
NextStock:
        DoEvents
    Next i

    '★④ ソート（スコア降順→出来高倍率降順）してから上位N件だけ残す
    If outRow > EXT_DATA_ROW + 1 Then
        extWs.Range(extWs.Cells(EXT_DATA_ROW, 1), _
                    extWs.Cells(outRow - 1, EXT_LAST_COL)).Sort _
            Key1:=extWs.Cells(EXT_DATA_ROW, 17), Order1:=xlDescending, _
            Key2:=extWs.Cells(EXT_DATA_ROW, 12), Order2:=xlDescending, Header:=xlNo
    End If
    If outRow - EXT_DATA_ROW > maxCount Then
        With extWs.Range(extWs.Cells(EXT_DATA_ROW + maxCount, 1), _
                         extWs.Cells(outRow - 1, EXT_LAST_COL))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
        End With
        outRow = EXT_DATA_ROW + maxCount
    End If

    '順位列（1,2,3…）を B列(2) に出力
    extWs.Cells(3, 2).Value = "順位"
    Dim rkRow As Long
    For rkRow = EXT_DATA_ROW To outRow - 1
        extWs.Cells(rkRow, 2).Value = rkRow - EXT_DATA_ROW + 1
        extWs.Cells(rkRow, 2).HorizontalAlignment = xlCenter
        extWs.Cells(rkRow, 2).Font.Bold = True
    Next rkRow

    On Error Resume Next
    extWs.Range("A2:R2").UnMerge
    On Error GoTo 0
    extWs.Range("A2:R2").Merge
    Dim modeDesc As String
    Select Case topixMode
        Case "LARGE": modeDesc = "【厳格】"
        Case "MID":   modeDesc = "【中程度】"
        Case Else:    modeDesc = "【通常】"
    End Select
    extWs.Cells(2, 1).Value = "TOPIX: " & topixTrend & _
        " (" & Format(topixPct, "+0.00;-0.00;0.00") & "%) " & modeDesc & _
        "　　最終更新: " & Format(Now(), "yyyy/mm/dd hh:mm") & "　　" & titleText
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Dim extracted As Long: extracted = 0
    Dim chkRow As Long
    For chkRow = EXT_DATA_ROW To outRow - 1
        If IsNumeric(extWs.Cells(chkRow, 17).Value) Then extracted = extracted + 1
    Next chkRow
    Dim msg As String
    msg = "【打ち出の無限こづち v13.1 絞り込み強化】" & vbCrLf & _
          "TOPIX: " & Format(topixPct, "+0.00;-0.00;0.00") & "% " & modeDesc & vbCrLf & _
          titleText & " 完了！" & vbCrLf & vbCrLf & _
          "スキャン件数: " & scannedCount & " 銘柄" & vbCrLf & _
          "抽出件数:    " & extracted & " 件" & vbCrLf & _
          "(抽出基準: スコア " & minScore & " 点以上 / 最大 " & maxCount & " 件)" & vbCrLf & vbCrLf & _
          "--- v13.1変更点 ---" & vbCrLf & _
          "件数上限:   " & maxCount & " 件（スコア→出来高倍率の順で上位のみ採用）" & vbCrLf & _
          "押し目必須: " & IIf(BUY_REQUIRE_OSIME, "ON（25日or5日押し目のみ）", "OFF") & vbCrLf & _
          "PO必須:     " & IIf(BUY_REQUIRE_PERFECT_ORDER, "ON", "OFF") & vbCrLf & _
          "Vol急増必須:" & IIf(BUY_REQUIRE_VOL_BURST, "ON", "OFF") & vbCrLf & _
          "Hist拡大必須:" & IIf(BUY_REQUIRE_HIST_EXPAND, "ON", "OFF") & vbCrLf & _
          "ダマシ閾値: 出来高" & Format(FAKE_BREAKOUT_VOL_RATIO, "0.0") & "倍未満→ダマシ扱い"
    If doMark Then
        msg = msg & vbCrLf & vbCrLf & "--- 高精度フィルター通過数 ---" & vbCrLf & _
              "ゴールデンクロス直近: " & cntGC & vbCrLf & _
              "パーフェクトオーダー: " & cntPO & vbCrLf & _
              "出来高急増+陽線:    " & cntVB & vbCrLf & _
              "ADX>" & Format(ADX_THRESHOLD, "0") & ":         " & cntADX & vbCrLf & _
              "MACDヒスト拡大:     " & cntHE & vbCrLf & _
              "対TOPIX相対強度:    " & cntRS
    End If
    UI_Msg msg, vbInformation, "抽出完了 (v13.1)"
End Sub

Private Sub 抽出実行_RSS(ByVal minScore As Long, _
                          ByVal doMark As Boolean, _
                          ByVal titleText As String, _
                          Optional ByVal maxCount As Long = 0)
    Call 抽出実行(minScore, doMark, titleText & "[RSS]", True, maxCount)
End Sub

'=============================================================================
' RSS現在値を先頭にした配列作成
'=============================================================================
Private Function RssMerge(ByVal ws As Worksheet, ByVal r As Long) As Variant
    Dim histLen As Long: histLen = 100
    Dim arr() As Double
    ReDim arr(1 To 1, 1 To 1 + histLen)
    Dim rssV As Variant: rssV = ws.Cells(r, 3).Value
    If IsNumeric(rssV) Then
        If CDbl(rssV) > 0 Then arr(1, 1) = CDbl(rssV)
    End If
    If arr(1, 1) = 0 Then
        Dim fb As Variant: fb = ws.Cells(r, 5).Value
        If IsNumeric(fb) Then arr(1, 1) = CDbl(fb)
    End If
    Dim hist As Variant
    hist = ws.Range(ws.Cells(r, 5), ws.Cells(r, 104)).Value
    Dim j As Long
    For j = 1 To histLen
        If IsNumeric(hist(1, j)) Then arr(1, 1 + j) = CDbl(hist(1, j))
    Next j
    RssMerge = arr
End Function

Private Function CandleScoreBuy(ByRef opens As Variant, ByRef highs As Variant, _
                                ByRef lows As Variant, ByRef closes As Variant, _
                                ByRef sig As String) As Long
    CandleScoreBuy = 0
    On Error GoTo ExitFunc
    If IsEmpty(opens) Or IsEmpty(highs) Or IsEmpty(lows) Or IsEmpty(closes) Then Exit Function
    If Not (IsNumeric(opens(1, 1)) And IsNumeric(highs(1, 1)) _
            And IsNumeric(lows(1, 1)) And IsNumeric(closes(1, 1))) Then Exit Function
    Dim o As Double: o = CDbl(opens(1, 1))
    Dim h As Double: h = CDbl(highs(1, 1))
    Dim l As Double: l = CDbl(lows(1, 1))
    Dim c As Double: c = CDbl(closes(1, 1))
    If o <= 0 Or h <= 0 Or l <= 0 Or c <= 0 Then Exit Function
    Dim rng As Double: rng = h - l
    If rng <= 0 Then Exit Function
    Dim body As Double: body = Abs(c - o)
    Dim up As Double: up = h - Application.Max(o, c)
    Dim lo As Double: lo = Application.Min(o, c) - l
    If body > 0 And lo >= body * 2# And up <= body * 0.5 Then
        CandleScoreBuy = CANDLE_LOWER_TAIL_YOSEN: sig = sig & "[下ヒゲ陽線●+2] "
    ElseIf c >= o And body >= rng * 0.7 Then
        CandleScoreBuy = CANDLE_BIG_YOSEN: sig = sig & "[大陽線▲+2] "
    ElseIf lo >= rng * 0.5 And up <= rng * 0.2 Then
        CandleScoreBuy = CANDLE_LONG_LOWER_TAIL: sig = sig & "[長い下ヒゲ+1] "
    ElseIf body > 0 And up >= body * 2# And lo <= body * 0.5 Then
        CandleScoreBuy = CANDLE_UPPER_TAIL_WARN: sig = sig & "[上ヒゲ警戒▼-2] "
    End If
ExitFunc:
End Function

Private Function CandleScoreBuyPrev(ByRef opens As Variant, ByRef highs As Variant, _
                                    ByRef lows As Variant, ByRef closes As Variant, _
                                    ByRef sig As String) As Long
    CandleScoreBuyPrev = 0
    On Error GoTo ExitFunc
    If IsEmpty(opens) Or IsEmpty(highs) Or IsEmpty(lows) Or IsEmpty(closes) Then Exit Function
    If UBound(opens, 2) < 2 Or UBound(highs, 2) < 2 Then Exit Function
    If UBound(lows, 2) < 2 Or UBound(closes, 2) < 2 Then Exit Function
    If Not (IsNumeric(opens(1, 2)) And IsNumeric(highs(1, 2)) _
            And IsNumeric(lows(1, 2)) And IsNumeric(closes(1, 2))) Then Exit Function
    Dim o As Double: o = CDbl(opens(1, 2))
    Dim h As Double: h = CDbl(highs(1, 2))
    Dim l As Double: l = CDbl(lows(1, 2))
    Dim c As Double: c = CDbl(closes(1, 2))
    If o <= 0 Or h <= 0 Or l <= 0 Or c <= 0 Then Exit Function
    Dim rng As Double: rng = h - l
    If rng <= 0 Then Exit Function
    Dim body As Double: body = Abs(c - o)
    Dim up As Double: up = h - Application.Max(o, c)
    Dim lo As Double: lo = Application.Min(o, c) - l
    If body > 0 And lo >= body * 2# And up <= body * 0.5 Then
        CandleScoreBuyPrev = CANDLE_LOWER_TAIL_YOSEN: sig = sig & "[前日:下ヒゲ陽線●+2] "
    ElseIf c >= o And body >= rng * 0.7 Then
        CandleScoreBuyPrev = CANDLE_BIG_YOSEN: sig = sig & "[前日:大陽線▲+2] "
    ElseIf lo >= rng * 0.5 And up <= rng * 0.2 Then
        CandleScoreBuyPrev = CANDLE_LONG_LOWER_TAIL: sig = sig & "[前日:長い下ヒゲ+1] "
    ElseIf body > 0 And up >= body * 2# And lo <= body * 0.5 Then
        CandleScoreBuyPrev = CANDLE_UPPER_TAIL_WARN: sig = sig & "[前日:上ヒゲ警戒▼-2] "
    End If
ExitFunc:
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim$(CStr(v))
    End If
End Function

Private Function SafeNum(ByVal v As Variant, ByVal defaultVal As Double) As Double
    If IsError(v) Then
        SafeNum = defaultVal
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeNum = defaultVal
    ElseIf IsNumeric(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = defaultVal
    End If
End Function

Private Function IsArrOk(ByVal v As Variant) As Boolean
    On Error Resume Next
    Dim ub As Long: ub = UBound(v, 2)
    IsArrOk = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetWs(ByVal n As String) As Worksheet
    On Error Resume Next
    Set GetWs = ThisWorkbook.Sheets(n)
    On Error GoTo 0
End Function

Private Sub ComputeMACDSeries(ByRef prices As Variant, _
                               ByRef ema12Today As Double, ByRef ema26Today As Double, _
                               ByRef macdToday As Double, ByRef signalToday As Double, _
                               ByRef histToday As Double, ByRef histYesterday As Double)
    ema12Today = 0: ema26Today = 0: macdToday = 0: signalToday = 0
    histToday = 0: histYesterday = 0
    On Error GoTo ExitProc
    If IsEmpty(prices) Then Exit Sub
    Dim n As Long: n = UBound(prices, 2)
    If n < 13 Then Exit Sub
    Dim k12 As Double: k12 = 2# / 13#
    Dim k26 As Double: k26 = 2# / 27#
    Dim k9  As Double: k9 = 2# / 10#
    ReDim macdArr(1 To n) As Double
    Dim e12 As Double: e12 = 0
    Dim e26 As Double: e26 = 0
    Dim seeded As Boolean: seeded = False
    Dim mCount As Long: mCount = 0
    Dim i As Long, pv As Double, v As Variant
    For i = n To 1 Step -1
        v = prices(1, i)
        If IsNumeric(v) Then
            pv = CDbl(v)
            If pv > 0 Then
                If Not seeded Then
                    e12 = pv: e26 = pv: seeded = True
                Else
                    e12 = k12 * pv + (1# - k12) * e12
                    e26 = k26 * pv + (1# - k26) * e26
                End If
                mCount = mCount + 1
                macdArr(mCount) = e12 - e26
            End If
        End If
    Next i
    If mCount < 2 Then Exit Sub
    ema12Today = e12: ema26Today = e26: macdToday = macdArr(mCount)
    Dim sig As Double: sig = macdArr(1)
    Dim sigPrev As Double: sigPrev = sig
    Dim j As Long
    For j = 2 To mCount
        sigPrev = sig
        sig = k9 * macdArr(j) + (1# - k9) * sig
    Next j
    signalToday = sig
    histToday = macdArr(mCount) - sig
    histYesterday = macdArr(mCount - 1) - sigPrev
ExitProc:
End Sub

Private Function IsGoldenCrossRecent(ByRef prices As Variant, ByVal lookback As Long) As Boolean
    IsGoldenCrossRecent = False
    On Error GoTo ExitProc
    If IsEmpty(prices) Then Exit Function
    Dim n As Long: n = UBound(prices, 2)
    If n < 26 Then Exit Function
    Dim k5 As Double:  k5 = 2# / 6#
    Dim k25 As Double: k25 = 2# / 26#
    ReDim e5Arr(1 To n) As Double
    ReDim e25Arr(1 To n) As Double
    Dim e5 As Double: e5 = 0
    Dim e25 As Double: e25 = 0
    Dim seeded As Boolean: seeded = False
    Dim cnt As Long: cnt = 0
    Dim i As Long, pv As Double, v As Variant
    For i = n To 1 Step -1
        v = prices(1, i)
        If IsNumeric(v) Then
            pv = CDbl(v)
            If pv > 0 Then
                If Not seeded Then
                    e5 = pv: e25 = pv: seeded = True
                Else
                    e5 = k5 * pv + (1# - k5) * e5
                    e25 = k25 * pv + (1# - k25) * e25
                End If
                cnt = cnt + 1
                e5Arr(cnt) = e5: e25Arr(cnt) = e25
            End If
        End If
    Next i
    If cnt < 2 Then Exit Function
    If e5Arr(cnt) <= e25Arr(cnt) Then Exit Function
    Dim lookStart As Long: lookStart = cnt - lookback
    If lookStart < 1 Then lookStart = 1
    Dim j As Long
    For j = lookStart To cnt - 1
        If e5Arr(j) <= e25Arr(j) Then
            IsGoldenCrossRecent = True: Exit Function
        End If
    Next j
ExitProc:
End Function

Private Function IsPerfectOrder(ByRef prices As Variant) As Boolean
    IsPerfectOrder = False
    On Error GoTo ExitProc
    If IsEmpty(prices) Then Exit Function
    Dim e5  As Double: e5 = ComputeSingleEMA(prices, 5)
    Dim e25 As Double: e25 = ComputeSingleEMA(prices, 25)
    Dim e75 As Double: e75 = ComputeSingleEMA(prices, 75)
    If e5 > 0 And e25 > 0 And e75 > 0 Then
        IsPerfectOrder = (e5 > e25) And (e25 > e75)
    End If
ExitProc:
End Function

Private Function ComputeSingleEMA(ByRef prices As Variant, ByVal period As Long) As Double
    ComputeSingleEMA = 0
    On Error GoTo ExitProc
    If IsEmpty(prices) Then Exit Function
    Dim n As Long: n = UBound(prices, 2)
    If n < 2 Then Exit Function
    Dim k As Double: k = 2# / (period + 1#)
    Dim ema As Double: ema = 0
    Dim seeded As Boolean: seeded = False
    Dim i As Long, pv As Double, v As Variant
    For i = n To 1 Step -1
        v = prices(1, i)
        If IsNumeric(v) Then
            pv = CDbl(v)
            If Not seeded Then
                ema = pv: seeded = True
            Else
                ema = k * pv + (1# - k) * ema
            End If
        End If
    Next i
    ComputeSingleEMA = ema
ExitProc:
End Function

Private Function IsVolumeBurstBullish(ByRef vols As Variant, ByRef closes As Variant, _
                                       ByRef opens As Variant, ByVal ratio As Double) As Boolean
    IsVolumeBurstBullish = False
    On Error GoTo ExitProc
    If IsEmpty(vols) Or IsEmpty(closes) Or IsEmpty(opens) Then Exit Function
    Dim n As Long: n = UBound(vols, 2)
    If n < 2 Then Exit Function
    If Not IsNumeric(vols(1, 1)) Then Exit Function
    If Not IsNumeric(closes(1, 1)) Then Exit Function
    If Not IsNumeric(opens(1, 1)) Then Exit Function
    Dim todayVol   As Double: todayVol = CDbl(vols(1, 1))
    Dim todayClose As Double: todayClose = CDbl(closes(1, 1))
    Dim todayOpen  As Double: todayOpen = CDbl(opens(1, 1))
    If todayVol <= 0 Or todayClose <= 0 Or todayOpen <= 0 Then Exit Function
    If todayClose <= todayOpen Then Exit Function
    Dim endIdx As Long: endIdx = 21
    If endIdx > n Then endIdx = n
    Dim Sum As Double: Sum = 0
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = 2 To endIdx
        If IsNumeric(vols(1, i)) Then
            Dim vv As Double: vv = CDbl(vols(1, i))
            If vv > 0 Then Sum = Sum + vv: cnt = cnt + 1
        End If
    Next i
    If cnt < 5 Then Exit Function
    IsVolumeBurstBullish = (todayVol >= (Sum / cnt) * ratio)
ExitProc:
End Function

Private Function ComputeADX(ByRef highs As Variant, ByRef lows As Variant, ByRef closes As Variant) As Double
    ComputeADX = 0
    On Error GoTo ExitProc
    If IsEmpty(highs) Or IsEmpty(lows) Or IsEmpty(closes) Then Exit Function
    Dim n As Long: n = UBound(highs, 2)
    If n < 28 Then Exit Function
    Const P As Long = 14
    ReDim hArr(1 To n) As Double
    ReDim lArr(1 To n) As Double
    ReDim cArr(1 To n) As Double
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = n To 1 Step -1
        If IsNumeric(highs(1, i)) And IsNumeric(lows(1, i)) And IsNumeric(closes(1, i)) Then
            Dim hv As Double, lv As Double, cv As Double
            hv = CDbl(highs(1, i)): lv = CDbl(lows(1, i)): cv = CDbl(closes(1, i))
            If hv > 0 And lv > 0 And cv > 0 Then
                cnt = cnt + 1
                hArr(cnt) = hv: lArr(cnt) = lv: cArr(cnt) = cv
            End If
        End If
    Next i
    If cnt < 28 Then Exit Function
    ReDim trArr(2 To cnt) As Double
    ReDim pdmArr(2 To cnt) As Double
    ReDim mdmArr(2 To cnt) As Double
    Dim j As Long
    For j = 2 To cnt
        Dim hd As Double, ld As Double
        hd = hArr(j) - hArr(j - 1): ld = lArr(j - 1) - lArr(j)
        pdmArr(j) = IIf(hd > ld And hd > 0, hd, 0)
        mdmArr(j) = IIf(ld > hd And ld > 0, ld, 0)
        Dim t1 As Double, t2 As Double, t3 As Double
        t1 = hArr(j) - lArr(j)
        t2 = Abs(hArr(j) - cArr(j - 1))
        t3 = Abs(lArr(j) - cArr(j - 1))
        trArr(j) = t1
        If t2 > trArr(j) Then trArr(j) = t2
        If t3 > trArr(j) Then trArr(j) = t3
    Next j
    If cnt < 2 + P - 1 Then Exit Function
    Dim atr As Double, pdm As Double, mdm As Double
    atr = 0: pdm = 0: mdm = 0
    For j = 2 To 1 + P
        atr = atr + trArr(j): pdm = pdm + pdmArr(j): mdm = mdm + mdmArr(j)
    Next j
    ReDim dxArr(1 To cnt) As Double
    Dim dxCnt As Long: dxCnt = 0
    Dim pDI As Double, mDI As Double, sDI As Double
    If atr > 0 Then
        pDI = 100# * pdm / atr: mDI = 100# * mdm / atr: sDI = pDI + mDI
        If sDI > 0 Then dxCnt = dxCnt + 1: dxArr(dxCnt) = 100# * Abs(pDI - mDI) / sDI
    End If
    For j = 2 + P To cnt
        atr = atr - (atr / P) + trArr(j)
        pdm = pdm - (pdm / P) + pdmArr(j)
        mdm = mdm - (mdm / P) + mdmArr(j)
        If atr > 0 Then
            pDI = 100# * pdm / atr: mDI = 100# * mdm / atr: sDI = pDI + mDI
            If sDI > 0 Then dxCnt = dxCnt + 1: dxArr(dxCnt) = 100# * Abs(pDI - mDI) / sDI
        End If
    Next j
    If dxCnt < P Then Exit Function
    Dim adx As Double: adx = 0
    For j = 1 To P: adx = adx + dxArr(j): Next j
    adx = adx / P
    For j = P + 1 To dxCnt: adx = (adx * (P - 1) + dxArr(j)) / P: Next j
    ComputeADX = adx
ExitProc:
End Function

Private Sub writeRow(ByVal ws As Worksheet, ByVal r As Long, _
                     ByVal code As String, ByVal meiName As String, _
                     ByVal closePrice As Double, ByVal highVal As Double, _
                     ByVal lowVal As Double, ByVal volToday As Double, _
                     ByVal rsiVal As Double, ByVal emaTrend As String, _
                     ByVal osime As String, ByVal volRatio As Double, _
                     ByVal macdVal As Double, ByVal highUpdate As String, _
                     ByVal isFake As Boolean, ByVal totalScore As Long, _
                     ByVal signals As String)
    With ws
        .Cells(r, 1).Formula = "=HYPERLINK(""https://kabutan.jp/stock/?code=""&TEXT(C" & r & ",""0000""),""株探"")"
        .Cells(r, 1).Interior.Color = COL_NAVY
        .Cells(r, 1).Font.Color = COL_WHITE
        .Cells(r, 1).Font.Bold = True
        .Cells(r, 1).HorizontalAlignment = xlCenter
        .Cells(r, 2).Font.Color = COL_BLACK
        .Cells(r, 3).Value = code
        .Cells(r, 3).NumberFormat = "0000"
        .Cells(r, 3).Font.Color = COL_BLACK
        .Cells(r, 4).Value = meiName
        .Cells(r, 4).Font.Color = COL_BLACK
        .Cells(r, 4).Font.Bold = True
        .Cells(r, 4).HorizontalAlignment = xlLeft
        .Cells(r, 5).Value = closePrice
        .Cells(r, 5).NumberFormat = "#,##0"
        .Cells(r, 5).Font.Color = COL_BLACK
        .Cells(r, 6).Value = highVal
        .Cells(r, 6).NumberFormat = "#,##0"
        .Cells(r, 7).Value = lowVal
        .Cells(r, 7).NumberFormat = "#,##0"
        .Cells(r, 8).Value = volToday
        .Cells(r, 8).NumberFormat = "#,##0"
        .Cells(r, 9).Value = Format(rsiVal, "0.0")
        .Cells(r, 10).Value = emaTrend
        Select Case emaTrend
            Case "パーフェクト▲"
                .Cells(r, 10).Interior.Color = RGB(0, 176, 80)
                .Cells(r, 10).Font.Color = COL_WHITE
                .Cells(r, 10).Font.Bold = True
            Case "上昇トレンド▲"
                .Cells(r, 10).Interior.Color = RGB(146, 208, 80)
                .Cells(r, 10).Font.Color = COL_BLACK
            Case "下降トレンド▼"
                .Cells(r, 10).Interior.Color = RGB(255, 0, 0)
                .Cells(r, 10).Font.Color = COL_WHITE
        End Select
        .Cells(r, 11).Value = osime
        .Cells(r, 12).Value = Format(volRatio, "0.0")
        .Cells(r, 13).Value = Format(macdVal, "0.00")
        .Cells(r, 14).Value = highUpdate
        If isFake Then
            .Cells(r, 14).Interior.Color = RGB(255, 255, 0)
            .Cells(r, 14).Font.Color = RGB(255, 0, 0)
        End If
        .Cells(r, 15).Value = Format(closePrice * 1.08, "#,##0")
        .Cells(r, 16).Value = Format(closePrice * 0.96, "#,##0")
        .Cells(r, 17).Value = totalScore
        If totalScore >= 17 Then
            .Cells(r, 17).Interior.Color = RGB(255, 0, 0)
            .Cells(r, 17).Font.Color = COL_WHITE
            .Cells(r, 17).Font.Bold = True
        ElseIf totalScore >= 14 Then
            .Cells(r, 17).Interior.Color = RGB(255, 102, 0)
            .Cells(r, 17).Font.Color = COL_WHITE
        End If
        .Cells(r, 18).Value = signals
        .Cells(r, 18).Font.Color = RGB(80, 80, 80)
    End With
End Sub

