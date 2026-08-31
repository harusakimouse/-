Attribute VB_Name = "Mod_SellExtrac"
' ==================================================================
' ★ 100番ブック / Mod_SellExtrac  v3 （2026/08/23 検証反映版）
'
' 【！ 先に読んでください ！】
'   このBOOKの実データ（300銘柄×250営業日）で売り側を検証したところ、
'   どの設定でも期待値がマイナスかゼロ近辺でした。
'
'     入口の比較（スコア11・損切1.5ATR・RR2・同時2枠）
'       翌日寄成          約54件 勝率37.0% 平均-0.246R 総-13.3R PF0.61
'       指値 終値+0.25ATR 約49件 勝率28.6% 平均-0.293R 総-14.4R PF0.56
'       指値 終値+0.75ATR 約45件 勝率46.7% 平均+0.046R 総 +2.1R PF1.10
'
'     スコア閾値を変えても改善しません（6以上でも15以上でもマイナス）。
'     損切・RRを振っても、最良で 総+3.1R（PF1.16）＝誤差の範囲です。
'
'   同じ期間・同じ枠で、買い側は 平均+1.0R以上 出ています。
'   資金枠が限られているなら、売りに枠を使うのは明確に不利です。
'
' 【ただし断定はできません】
'   検証期間は +16% の上昇相場1年のみで、空売りに構造的に不利な期間です。
'   「売りロジックがダメ」ではなく「この相場では機能しなかった」が正確です。
'   下降相場のデータが手に入ったら、必ず検証し直してください。
'
' 【v3で直したこと】
'   ① 抽出の冒頭にあった 分析シート／終値シートの強制再計算に歯止めを付けた。
'      従来は無条件に Calculate していたため、RSSが死んでいる状態で抽出すると
'      各シートのRSSキャッシュ値が一斉に消え、表示が飛んでいた（実際に起きた事故）。
'      v3ではRSSが応答しているときだけ再計算する。
'   ② ザラ場の出来高判定を直した。ザラ場中の出来高は途中までの累計なので、
'      時刻に応じた進捗率で1日分に換算してから5日平均と比べる。
'      15:00は1日の約78%しか出来ていないため、補正しないと出来高急増を拾えない。
'   ③ 安全装置モジュールがあれば、実行前にRSSの生死を確認する。
'      無ければ何もしない（従来どおり動く）。
' ==================================================================

Option Explicit
'=============================================================================
' Mod_売り抽出
' 売り候補抽出モジュール v2 (2026/06/10)
' 買いロジック(候補抽出v10)と同じ指標・同じ構造
' シート：終値/始値/高値/安値/出来高/分析/銘柄管理 → 売抽出
'=============================================================================

Dim NextSellRunTime As Date

' ===== スコア定数 =====
Private Const S_DC_BONUS   As Long = 2   ' デッドクロス
Private Const S_PD_BONUS   As Long = 2   ' パーフェクト下降
Private Const S_VOL_BONUS  As Long = 2   ' 出来高急増+陰線
Private Const S_ADX_BONUS  As Long = 1   ' ADX>25
Private Const S_HIST_BONUS As Long = 1   ' MACDヒスト悪化
Private Const S_MIN_SCORE  As Long = 11  ' 抽出最低スコア
Private Const S_MAX_HIGH   As Long = 2   ' 高精度売りの上限（TOPからN件）
Private Const S_MAX_NORMAL As Long = 5   ' 通常売りの上限（TOPからN件）
Private Const S_MAX_EASE   As Long = 10  ' 緩和売りの上限（TOPからN件）

' ===== 必須フィルタ =====
Private Const S_REQUIRE_DOWNTREND As Boolean = True

' ===== v3追加 =====
Private Const S_PROBE_COL As Long = 60   'RSS点検に使う作業列（BH列）

'=============================================================================
' ■ 公開Sub（ボタンに割り当て）
'=============================================================================
Public Sub 売抽出更新()
    Dim p As Boolean: p = gUnattended
    gUnattended = True
    On Error GoTo done
    Dim h As Integer: h = Hour(Now)
    Dim m As Integer: m = Minute(Now)
    ' 営業時間内(9:00～15:30)はRSS現在値版、それ以外は保存済みデータ版
    If (h >= 9 And h < 15) Or (h = 15 And m <= 30) Then
        Call S_抽出実行_RSS(S_MIN_SCORE, True, "売り抽出[RSS現在値](スコア" & S_MIN_SCORE & "以上・上位2件)", S_MAX_HIGH)
    Else
        Call S_抽出実行(S_MIN_SCORE, True, "売り高精度(スコア" & S_MIN_SCORE & "以上・上位2件)", S_MAX_HIGH)
    End If
done:
    gUnattended = p
End Sub

Public Sub 売り候補抽出()
    If Not S_安全確認("売り候補抽出") Then Exit Sub
    Call S_抽出実行(S_MIN_SCORE, True, "売り高精度(スコア" & S_MIN_SCORE & "以上・上位2件)", S_MAX_HIGH)
End Sub

'=============================================================================
' ★v3追加：ザラ場売り抽出（15:00頃の見直し用）
'   時刻に関係なく、必ずRSSの現在値・今日の出来高で抽出する。
'   出来高は時刻に応じた進捗率で1日分に換算してから判定する。
'=============================================================================
Public Sub ザラ場売り抽出_今すぐ()
    If Not S_安全確認("ザラ場売り抽出") Then Exit Sub
    Dim vp As Double: vp = S_VolProgress()
    Call S_抽出実行_RSS(S_MIN_SCORE, True, _
         "ザラ場売り " & Format(Now, "hh:mm") & "(スコア" & S_MIN_SCORE & "以上・上位" & _
         S_MAX_HIGH & "件・出来高進捗" & Format(vp * 100, "0") & "%換算)", S_MAX_HIGH)
End Sub

Public Sub 売り抽出_通常()
    Call S_抽出実行(8, True, "売り通常(スコア8以上・上位5件)", S_MAX_NORMAL)
End Sub

Public Sub 売り抽出_緩和()
    Call S_抽出実行(8, False, "売り緩和(スコア8以上・上位10件)", S_MAX_EASE)
End Sub

Public Sub 売抽出自動更新開始()
    NextSellRunTime = Now + TimeValue("00:05:00")
    Application.OnTime NextSellRunTime, "'" & ThisWorkbook.Name & "'!売抽出タイマー実行"
    UI_Msg "自動更新開始（5分ごと）", vbInformation
End Sub

Public Sub 売抽出タイマー実行()
    Dim p As Boolean: p = gUnattended
    gUnattended = True
    On Error Resume Next
    Call 売り候補抽出
    On Error GoTo 0
    gUnattended = p
    NextSellRunTime = Now + TimeValue("00:05:00")
    Application.OnTime NextSellRunTime, "'" & ThisWorkbook.Name & "'!売抽出タイマー実行"

End Sub

Public Sub 売抽出自動更新停止()
    On Error Resume Next
    Application.OnTime NextSellRunTime, "'" & ThisWorkbook.Name & "'!売抽出タイマー実行", , False

    On Error GoTo 0
    UI_Msg "自動更新停止", vbInformation
End Sub

'=============================================================================
' ■ メイン処理
'=============================================================================
Private Sub S_抽出実行(ByVal minScore As Long, _
                       ByVal doMark As Boolean, _
                       ByVal titleText As String, _
                       Optional ByVal maxCount As Long = 0)

    If maxCount <= 0 Then maxCount = S_MAX_HIGH
    Dim closeWs As Worksheet: Set closeWs = S_GetWs("終値")
    Dim openWs  As Worksheet: Set openWs = S_GetWs("始値")
    Dim highWs  As Worksheet: Set highWs = S_GetWs("高値")
    Dim lowWs   As Worksheet: Set lowWs = S_GetWs("安値")
    Dim volWs   As Worksheet: Set volWs = S_GetWs("出来高")
    Dim extWs   As Worksheet: Set extWs = S_GetWs("売抽出v13")
    Dim mws     As Worksheet: Set mws = S_GetWs("銘柄管理")
    Dim anaWs   As Worksheet: Set anaWs = S_GetWs("分析")

    If closeWs Is Nothing Or extWs Is Nothing Then
        UI_Msg "「終値」または「売抽出」シートが見つかりません。", vbExclamation: Exit Sub
    End If
    If mws Is Nothing Then
        UI_Msg "「銘柄管理」シートが見つかりません。", vbExclamation: Exit Sub
    End If

    '★v3：RSSが死んでいる状態でこの再計算をすると、分析シートや終値シートの
    '  RSSキャッシュ値が一斉に消え、各シートの表示が飛ぶ。
    '  RSSが応答しているときだけ再計算する。
    If S_RSS生存(extWs) Then
        On Error Resume Next
        If Not anaWs Is Nothing Then anaWs.Calculate
        closeWs.Calculate
        On Error GoTo 0
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    extWs.Unprotect Password:="ne19480314"
    mws.Unprotect Password:="ne19480314"
    If Not anaWs Is Nothing Then anaWs.Unprotect Password:="ne19480314"
    On Error GoTo 0

    ' ヘッダー行を生成（3行目：買い抽出シートと同じ構造）
    On Error Resume Next
    extWs.Unprotect Password:="ne19480314"
    On Error GoTo 0
    Dim hdr As Variant
    hdr = Array("株探", "メモ", "銘柄コード", "銘柄名称", "現在値", "高値", "安値", _
                "出来高", "RSI(14)", "EMAトレンド", "戻り売り判定", "出来高倍率", _
                "MACD", "安値更新", "利確ライン", "損切ライン", "スコア", "シグナル内容")
    Dim hc As Long
    For hc = 0 To 17
        With extWs.Cells(3, hc + 1)
            .Value = hdr(hc)
            .Font.Name = "Meiryo UI"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(80, 0, 0)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next hc
    extWs.Rows(3).RowHeight = 40

    ' 抽出シートのデータ行をクリア（5行目以降）
    ' 5行目以降をクリア
    Dim extLast As Long
    extLast = extWs.Cells(extWs.Rows.Count, 2).End(xlUp).Row
    If extLast < 5 Then extLast = 5
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).ClearContents
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Interior.ColorIndex = xlNone
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Font.ColorIndex = xlAutomatic
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Font.Bold = False

    ' TOPIX
    Dim topixTrend As String: topixTrend = S_SafeStr(mws.Cells(5, 6).Value)
    Dim topixNow As Double: topixNow = S_SafeNum(mws.Cells(5, 3).Value, 0)
    Dim topixRef As Double: topixRef = S_SafeNum(mws.Cells(5, 5).Value, 0)
    Dim topixPct As Double: topixPct = 0
    If topixRef > 0 And topixNow > 0 Then topixPct = (topixNow - topixRef) / topixRef * 100#

    Dim topixMode As String
    If topixPct <= -1.5 Then
        topixMode = "LARGE"
    ElseIf topixPct <= -0.5 Then
        topixMode = "MID"
    Else
        topixMode = "OK"
    End If
    ' Futures filter for SELL: block short when futures surge
    ' (修正: 先物急騰時は空売りを難しくする。LARGE=最も緩い判定なので誤り。OK=最も厳しい判定が正しい)
    Dim futuresPct As Double
    futuresPct = S_SafeNum(mws.Cells(5, 8).Value, 0)
    If futuresPct > 1.5 Then
        topixMode = "OK"
    ElseIf futuresPct > 0.5 Then
        If topixMode = "LARGE" Then topixMode = "MID"
    End If

    ' 銘柄名辞書
    Dim nameDic As Object: Set nameDic = CreateObject("Scripting.Dictionary")
    Dim mLast As Long: mLast = mws.Cells(mws.Rows.Count, 2).End(xlUp).Row
    Dim mr As Long
    For mr = 5 To mLast
        Dim mCode As String: mCode = Trim$(CStr(mws.Cells(mr, 2).Value))
        Dim mName As String: mName = Trim$(CStr(mws.Cells(mr, 3).Value))
        If mCode <> "" And Not nameDic.Exists(mCode) Then nameDic.Add mCode, mName
    Next mr

    ' 分析シート辞書
    Dim anaDic As Object: Set anaDic = CreateObject("Scripting.Dictionary")
    If Not anaWs Is Nothing Then
        Dim anaLast As Long: anaLast = anaWs.Cells(anaWs.Rows.Count, 3).End(xlUp).Row
        Dim k As Long
        For k = 3 To anaLast
            Dim anaCode As String: anaCode = Trim$(CStr(anaWs.Cells(k, 3).Value))
            If anaCode <> "" And Not anaDic.Exists(anaCode) Then anaDic(anaCode) = k
        Next k
    End If

    Dim lastRow As Long: lastRow = closeWs.Cells(closeWs.Rows.Count, 1).End(xlUp).Row
    Dim outRow As Long: outRow = 4
    Dim totalScore As Long, signals As String
    Dim scannedCount As Long: scannedCount = 0
    Dim i As Long

    For i = 6 To lastRow
        Dim code As String: code = Trim$(CStr(closeWs.Cells(i, 1).Value))
        If code = "" Or code = "0" Or code = "TOPX" Then GoTo NextS

        scannedCount = scannedCount + 1
        Application.StatusBar = titleText & " : " & code
        totalScore = 0: signals = ""

        Dim closePrice As Double: closePrice = Val(closeWs.Cells(i, 5).Value)
        If closePrice = 0 Then GoTo NextS

        Dim highVal As Double: highVal = 0
        Dim lowVal  As Double: lowVal = 0
        If Not highWs Is Nothing Then highVal = Val(highWs.Cells(i, 5).Value)
        If Not lowWs Is Nothing Then lowVal = Val(lowWs.Cells(i, 5).Value)

        Dim closes As Variant, opens As Variant, highs As Variant, loWs As Variant, vols As Variant
        closes = closeWs.Range(closeWs.Cells(i, 5), closeWs.Cells(i, 64)).Value
        If Not openWs Is Nothing Then opens = openWs.Range(openWs.Cells(i, 5), openWs.Cells(i, 64)).Value
        If Not highWs Is Nothing Then highs = highWs.Range(highWs.Cells(i, 5), highWs.Cells(i, 64)).Value
        If Not lowWs Is Nothing Then loWs = lowWs.Range(lowWs.Cells(i, 5), lowWs.Cells(i, 64)).Value
        If Not volWs Is Nothing Then vols = volWs.Range(volWs.Cells(i, 5), volWs.Cells(i, 64)).Value

        Dim anaRow As Long: anaRow = 0
        If anaDic.Exists(code) Then anaRow = CLng(anaDic(code))
        ' RSI（分析シートから取得、なければVBAで直接計算）
        Dim rsiVal As Double: rsiVal = 50
        If anaRow > 0 Then rsiVal = S_SafeNum(anaWs.Cells(anaRow, 23).Value, 50)
        If rsiVal = 50 Or anaRow = 0 Then rsiVal = S_ComputeRSI(closes, 14)
        If rsiVal >= 70 Then
            totalScore = totalScore + 2: signals = signals & "RSI買われ過ぎ "
        ElseIf rsiVal >= 60 Then
            totalScore = totalScore + 1: signals = signals & "RSI高 "
        ElseIf rsiVal <= 30 Then
            totalScore = totalScore - 1
        End If

        ' EMAトレンド（分析シートから取得、なければVBAで直接計算）
        Dim emaTrend As String: emaTrend = ""
        If anaRow > 0 Then emaTrend = S_SafeStr(anaWs.Cells(anaRow, 30).Value)
        ' 分析シートにない場合はVBAで直接計算
        If emaTrend = "" Then
            Dim e5c As Double: e5c = S_EMA(closes, 5)
            Dim e25c As Double: e25c = S_EMA(closes, 25)
            Dim e75c As Double: e75c = S_EMA(closes, 75)
            If e5c > 0 And e25c > 0 And e75c > 0 Then
                If e5c > e25c And e25c > e75c Then
                    emaTrend = "パーフェクト▲"
                ElseIf e5c < e25c And e25c < e75c Then
                    emaTrend = "下降トレンド▼"
                ElseIf e5c > e25c Then
                    emaTrend = "上昇トレンド▲"
                Else
                    emaTrend = "もみ合い→"
                End If
            End If
        End If
        If emaTrend = "下降トレンド▼" Then
            totalScore = totalScore + 2: signals = signals & "下降トレンド▼ "
        ElseIf emaTrend = "パーフェクト▲" Then
            totalScore = totalScore - 1
        End If

        ' 戻り売り
        Dim osime As String: osime = ""
        If anaRow > 0 Then osime = S_SafeStr(anaWs.Cells(anaRow, 31).Value)
        Dim uwamodori As String: uwamodori = "-"
        If osime = "25日押し目★" Then
            uwamodori = "25日線戻り売り●"
            totalScore = totalScore + 2: signals = signals & "25日線戻り● "
        ElseIf osime = "5日押し目○" Then
            uwamodori = "5日線戻り売り○"
            totalScore = totalScore + 1: signals = signals & "5日線戻り○ "
        End If

        ' 出来高（分析シートから取得、なければVBAで直接計算）
        Dim volToday As Double: volToday = 0
        Dim vol5avg  As Double: vol5avg = 0
        Dim volRatio As Double: volRatio = 0
        If Not volWs Is Nothing Then
            volToday = Val(volWs.Cells(i, 5).Value)
            If anaRow > 0 Then vol5avg = S_SafeNum(anaWs.Cells(anaRow, 13).Value, 0)
            ' 分析シートにない場合は直近5日平均を計算
            If vol5avg = 0 And Not IsEmpty(vols) Then
                Dim vi As Long: Dim vsum As Double: vsum = 0: Dim vcnt As Long: vcnt = 0
                For vi = 2 To 6
                    If UBound(vols, 2) >= vi Then
                        If IsNumeric(vols(1, vi)) And CDbl(vols(1, vi)) > 0 Then
                            vsum = vsum + CDbl(vols(1, vi)): vcnt = vcnt + 1
                        End If
                    End If
                Next vi
                If vcnt > 0 Then vol5avg = vsum / vcnt
            End If
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

        ' MACD（分析シートから取得、なければMACDシリーズ計算後に設定）
        Dim macdVal As Double: macdVal = 0
        If anaRow > 0 Then macdVal = S_SafeNum(anaWs.Cells(anaRow, 24).Value, 0)

        ' 年初来安値更新
        Dim lowUpdate As String: lowUpdate = ""
        If anaRow > 0 Then lowUpdate = S_SafeStr(anaWs.Cells(anaRow, 22).Value)
        If InStr(lowUpdate, "安値更新") > 0 Then
            totalScore = totalScore + 2: signals = signals & "年初来安値更新▼ "
        End If

        ' MACDシリーズ
        Dim ema12V As Double, ema26V As Double, macdNow As Double, signalNow As Double
        Dim histNow As Double, histPrev As Double
        S_ComputeMACD closes, ema12V, ema26V, macdNow, signalNow, histNow, histPrev

        ' 分析シートにない場合はMACDシリーズの値を使用
        If macdVal = 0 And anaRow = 0 Then macdVal = macdNow
        If macdVal < 0 Then
            totalScore = totalScore + 1: signals = signals & "MACD- "
        ElseIf macdVal > 0 Then
            totalScore = totalScore - 1
        End If

        Dim emaDown  As Boolean: emaDown = (ema12V < ema26V) And (ema26V > 0)
        Dim macdDown As Boolean: macdDown = (macdNow < signalNow)
        Dim rsiDown  As Boolean: rsiDown = (rsiVal < 50)

        ' デッドクロス直近5日
        Dim ok_DC As Boolean: ok_DC = S_IsDeadCross(closes, 5)
        If ok_DC Then
            totalScore = totalScore + S_DC_BONUS: signals = signals & "[DC直近5日] "
        End If

        ' パーフェクト下降
        Dim ok_PD As Boolean: ok_PD = S_IsPerfectDown(closes)
        If ok_PD Then
            totalScore = totalScore + S_PD_BONUS: signals = signals & "[完全下降] "
        End If

        ' 出来高急増+陰線
        Dim ok_VBB As Boolean: ok_VBB = False
        If Not IsEmpty(opens) And Not IsEmpty(vols) Then
            ok_VBB = S_IsVolBear(vols, closes, opens)
            If ok_VBB Then
                totalScore = totalScore + S_VOL_BONUS: signals = signals & "[Vol急増陰線] "
            End If
        End If

        ' ADX
        Dim adxVal As Double: adxVal = 0
        Dim ok_ADX As Boolean: ok_ADX = False
        If Not IsEmpty(highs) And Not IsEmpty(loWs) Then
            adxVal = S_ComputeADX(highs, loWs, closes)
            ok_ADX = (adxVal > 25#)
            If ok_ADX Then
                totalScore = totalScore + S_ADX_BONUS: signals = signals & "[ADX=" & Format(adxVal, "0") & "] "
            End If
        End If

        ' MACDヒスト悪化
        Dim ok_HE As Boolean: ok_HE = (histNow < histPrev) And (histNow < 0)
        If ok_HE Then
            totalScore = totalScore + S_HIST_BONUS: signals = signals & "[Hist悪化] "
        End If

        ' 売りモメンタム
        Dim momCount As Long: momCount = 0
        If emaDown Then momCount = momCount + 1: signals = signals & "EMA12<26 "
        If macdDown Then momCount = momCount + 1: signals = signals & "MACD<Sig "
        If rsiDown Then momCount = momCount + 1: signals = signals & "RSI<50 "
        If momCount = 3 Then
            totalScore = totalScore + 2: signals = signals & "[◎売りモメンタム完全] "
        ElseIf momCount = 2 Then
            totalScore = totalScore + 1: signals = signals & "[○売りモメンタム部分] "
        End If

        ' ローソク足（当日・前日）
        totalScore = totalScore + S_CandleBear(opens, highs, loWs, closes, signals)
        totalScore = totalScore + S_CandleBearPrev(opens, highs, loWs, closes, signals)

        ' 対TOPIX相対弱度
        Dim stockPct As Double: stockPct = 0
        If Not IsEmpty(closes) Then
            If UBound(closes, 2) >= 2 Then
                Dim prevClose As Double: prevClose = S_SafeNum(closes(1, 2), 0)
                If prevClose > 0 Then stockPct = (closePrice - prevClose) / prevClose * 100#
            End If
        End If
        If stockPct < topixPct Then
            totalScore = totalScore + 1: signals = signals & "[対TOPIX弱] "
        End If

        ' 抽出判定
        Dim canExtract As Boolean: canExtract = False
        If totalScore >= minScore Then
            Select Case topixMode
                Case "LARGE": canExtract = True
                Case "MID": If emaTrend <> "パーフェクト▲" Then canExtract = True
                Case Else: If emaTrend = "下降トレンド▼" Or ok_PD Then canExtract = True
            End Select
        End If
        If canExtract And rsiVal < 15 Then canExtract = False
        If canExtract And rsiVal >= 80 Then canExtract = False
        If canExtract And stockPct >= topixPct Then canExtract = False
        If doMark And canExtract Then
            If momCount < 2 Then canExtract = False
            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False
        End If

        If canExtract And code <> "" Then
            Dim meiName As String
            If nameDic.Exists(code) Then
                meiName = CStr(nameDic(code))
            Else
                meiName = S_SafeStr(closeWs.Cells(i, 2).Value)
            End If
            If meiName = "" Then meiName = "(" & code & ")"

            Dim currentPrice As Double: currentPrice = Val(closeWs.Cells(i, 3).Value)
            If currentPrice <= 0 Then currentPrice = closePrice
            Dim currentHigh As Double: currentHigh = highVal
            If Not highWs Is Nothing Then
                Dim tmpH As Double: tmpH = Val(highWs.Cells(i, 3).Value)
                If tmpH > 0 Then currentHigh = tmpH
            End If
            Dim currentLow As Double: currentLow = lowVal
            If Not lowWs Is Nothing Then
                Dim tmpL As Double: tmpL = Val(lowWs.Cells(i, 3).Value)
                If tmpL > 0 Then currentLow = tmpL
            End If
            Dim currentVol As Double: currentVol = volToday
            If Not volWs Is Nothing Then
                Dim tmpV As Double: tmpV = Val(volWs.Cells(i, 3).Value)
                If tmpV > 0 Then currentVol = tmpV
            End If

            S_WriteRow extWs, outRow, code, meiName, currentPrice, currentHigh, currentLow, _
                       currentVol, rsiVal, emaTrend, uwamodori, volRatio, macdVal, _
                       lowUpdate, totalScore, signals
            outRow = outRow + 1
        End If
NextS:
        DoEvents
    Next i

    ' 空白行削除してからソート
    Dim delRow As Long
    For delRow = outRow - 1 To 4 Step -1
        If Trim(CStr(extWs.Cells(delRow, 3).Value)) = "" Then
            extWs.Rows(delRow).Delete
            outRow = outRow - 1
        End If
    Next delRow

    ' スコア降順ソート（結合セル回避）
    If outRow > 6 Then
        On Error Resume Next
        extWs.Range(extWs.Cells(4, 1), extWs.Cells(outRow - 1, 18)).Sort _
            Key1:=extWs.Cells(4, 17), Order1:=xlDescending, Header:=xlNo
        On Error GoTo 0
    End If

    '★上位N件だけ残す（スコア降順ソート後）
    If outRow - 4 > maxCount Then
        extWs.Range(extWs.Cells(4 + maxCount, 1), extWs.Cells(outRow - 1, 18)).ClearContents
        extWs.Range(extWs.Cells(4 + maxCount, 1), extWs.Cells(outRow - 1, 18)).Interior.ColorIndex = xlNone
        outRow = 4 + maxCount
    End If

    '★順位列（1,2,3…）を B列(2) に出力
    extWs.Cells(3, 2).Value = "順位"
    Dim rkRow As Long
    For rkRow = 4 To outRow - 1
        extWs.Cells(rkRow, 2).Value = rkRow - 3
        extWs.Cells(rkRow, 2).HorizontalAlignment = xlCenter
        extWs.Cells(rkRow, 2).Font.Bold = True
    Next rkRow

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim extracted As Long: extracted = outRow - 4
    Dim modeDesc As String
    Select Case topixMode
        Case "LARGE": modeDesc = "【相場下落】"
        Case "MID":   modeDesc = "【中程度】"
        Case Else:    modeDesc = "【通常】"
    End Select

    ' 2行目にTOPIX情報を書く（買い抽出シートと同じ構造）
    On Error Resume Next
    extWs.Cells(2, 1).MergeArea.UnMerge
    On Error GoTo 0
    extWs.Range("A2:R2").Merge
    extWs.Cells(2, 1).Value = "TOPIX: " & topixTrend & _
        " (" & Format(topixPct, "+0.00;-0.00;0.00") & "%) " & modeDesc & _
        "　最終更新: " & Format(Now(), "yyyy/mm/dd hh:mm")
    With extWs.Cells(2, 1)
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Color = RGB(255, 215, 0)
        .Interior.Color = RGB(40, 0, 0)
        .HorizontalAlignment = xlRight
    End With

    UI_Msg "【売り候補抽出 v2】" & vbCrLf & _
           "TOPIX: " & Format(topixPct, "+0.00;-0.00;0.00") & "% " & modeDesc & vbCrLf & _
           "スキャン: " & scannedCount & " 銘柄" & vbCrLf & _
           "抽出:    " & extracted & " 件 (スコア" & minScore & "点以上)", _
           vbInformation, "売り抽出完了"
End Sub

'=============================================================================
' RSS現在値版抽出（営業時間内用）
'=============================================================================
Private Sub S_抽出実行_RSS(ByVal minScore As Long, _
                            ByVal doMark As Boolean, _
                            ByVal titleText As String, _
                            Optional ByVal maxCount As Long = 0)

    If maxCount <= 0 Then maxCount = S_MAX_HIGH
    Dim closeWs As Worksheet: Set closeWs = S_GetWs("終値")
    Dim openWs  As Worksheet: Set openWs = S_GetWs("始値")
    Dim highWs  As Worksheet: Set highWs = S_GetWs("高値")
    Dim lowWs   As Worksheet: Set lowWs = S_GetWs("安値")
    Dim volWs   As Worksheet: Set volWs = S_GetWs("出来高")
    Dim extWs   As Worksheet: Set extWs = S_GetWs("売抽出v13")
    Dim mws     As Worksheet: Set mws = S_GetWs("銘柄管理")
    Dim anaWs   As Worksheet: Set anaWs = S_GetWs("分析")

    If closeWs Is Nothing Or extWs Is Nothing Then
        UI_Msg "「終値」または「売抽出」シートが見つかりません。", vbExclamation: Exit Sub
    End If

    '★v3：RSSが死んでいる状態でこの再計算をすると、分析シートや終値シートの
    '  RSSキャッシュ値が一斉に消え、各シートの表示が飛ぶ。
    '  RSSが応答しているときだけ再計算する。
    If S_RSS生存(extWs) Then
        On Error Resume Next
        If Not anaWs Is Nothing Then anaWs.Calculate
        closeWs.Calculate
        On Error GoTo 0
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    extWs.Unprotect Password:="ne19480314"
    mws.Unprotect Password:="ne19480314"
    If Not anaWs Is Nothing Then anaWs.Unprotect Password:="ne19480314"
    On Error GoTo 0

    ' ヘッダー生成
    Dim hdr As Variant
    hdr = Array("株探", "メモ", "コード", "銘柄名", "現在値", "高値", "安値", _
                "出来高", "RSI(14)", "EMAトレンド", "戻り売り判定", "出来高倍率", _
                "MACD", "安値更新", "利確ライン", "損切ライン", "スコア", "シグナル内容")
    Dim hc As Long
    For hc = 0 To 17
        With extWs.Cells(4, hc + 1)
            .Value = hdr(hc)
            .Font.Name = "Meiryo UI"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(80, 0, 0)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next hc
    extWs.Rows(4).RowHeight = 40

    ' クリア
    ' 5行目以降をクリア
    Dim extLast As Long
    extLast = extWs.Cells(extWs.Rows.Count, 2).End(xlUp).Row
    If extLast < 5 Then extLast = 5
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).ClearContents
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Interior.ColorIndex = xlNone
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Font.ColorIndex = xlAutomatic
    extWs.Range(extWs.Cells(4, 1), extWs.Cells(extLast + 500, 18)).Font.Bold = False

    ' TOPIX
    Dim topixTrend As String: topixTrend = S_SafeStr(mws.Cells(5, 6).Value)
    Dim topixNow As Double: topixNow = S_SafeNum(mws.Cells(5, 3).Value, 0)
    Dim topixRef As Double: topixRef = S_SafeNum(mws.Cells(5, 5).Value, 0)
    Dim topixPct As Double: topixPct = 0
    If topixRef > 0 And topixNow > 0 Then topixPct = (topixNow - topixRef) / topixRef * 100#
    Dim topixMode As String
    If topixPct <= -1.5 Then
        topixMode = "LARGE"
    ElseIf topixPct <= -0.5 Then
        topixMode = "MID"
    Else
        topixMode = "OK"
    End If
    ' Futures filter for SELL: block short when futures surge
    ' (修正: 先物急騰時は空売りを難しくする。LARGE=最も緩い判定なので誤り。OK=最も厳しい判定が正しい)
    Dim futuresPct As Double
    futuresPct = S_SafeNum(mws.Cells(5, 8).Value, 0)
    If futuresPct > 1.5 Then
        topixMode = "OK"
    ElseIf futuresPct > 0.5 Then
        If topixMode = "LARGE" Then topixMode = "MID"
    End If

    ' 銘柄名辞書
    Dim nameDic As Object: Set nameDic = CreateObject("Scripting.Dictionary")
    Dim mLast As Long: mLast = mws.Cells(mws.Rows.Count, 2).End(xlUp).Row
    Dim mr As Long
    For mr = 5 To mLast
        Dim mCode As String: mCode = Trim$(CStr(mws.Cells(mr, 2).Value))
        Dim mName As String: mName = Trim$(CStr(mws.Cells(mr, 3).Value))
        If mCode <> "" And Not nameDic.Exists(mCode) Then nameDic.Add mCode, mName
    Next mr

    ' 分析シート辞書
    Dim anaDic As Object: Set anaDic = CreateObject("Scripting.Dictionary")
    If Not anaWs Is Nothing Then
        Dim anaLast As Long: anaLast = anaWs.Cells(anaWs.Rows.Count, 3).End(xlUp).Row
        Dim k As Long
        For k = 3 To anaLast
            Dim anaCode As String: anaCode = Trim$(CStr(anaWs.Cells(k, 3).Value))
            If anaCode <> "" And Not anaDic.Exists(anaCode) Then anaDic(anaCode) = k
        Next k
    End If

    Dim lastRow As Long: lastRow = closeWs.Cells(closeWs.Rows.Count, 1).End(xlUp).Row
    Dim outRow As Long: outRow = 4
    Dim totalScore As Long, signals As String
    Dim scannedCount As Long: scannedCount = 0
    Dim i As Long

    For i = 6 To lastRow
        Dim code As String: code = Trim$(CStr(closeWs.Cells(i, 1).Value))
        If code = "" Or code = "0" Or code = "TOPX" Then GoTo NextRSS

        scannedCount = scannedCount + 1
        Application.StatusBar = titleText & " : " & code
        totalScore = 0: signals = ""

        '★ RSS現在値(C列)を先頭に使用
        Dim rssClose As Double: rssClose = Val(closeWs.Cells(i, 3).Value)
        Dim closePrice As Double
        If rssClose > 0 Then
            closePrice = rssClose
        Else
            closePrice = Val(closeWs.Cells(i, 5).Value)
        End If
        If closePrice = 0 Then GoTo NextRSS

        Dim highVal As Double: highVal = 0
        Dim lowVal  As Double: lowVal = 0
        If Not highWs Is Nothing Then
            Dim rssH As Double: rssH = Val(highWs.Cells(i, 3).Value)
            highVal = IIf(rssH > 0, rssH, Val(highWs.Cells(i, 5).Value))
        End If
        If Not lowWs Is Nothing Then
            Dim rssL As Double: rssL = Val(lowWs.Cells(i, 3).Value)
            lowVal = IIf(rssL > 0, rssL, Val(lowWs.Cells(i, 5).Value))
        End If

        '★ RSS現在値を先頭にした配列
        Dim closes As Variant, opens As Variant, highs As Variant, loWs As Variant, vols As Variant
        closes = S_RssMerge(closeWs, i)
        If Not openWs Is Nothing Then opens = S_RssMerge(openWs, i)
        If Not highWs Is Nothing Then highs = S_RssMerge(highWs, i)
        If Not lowWs Is Nothing Then loWs = S_RssMerge(lowWs, i)
        If Not volWs Is Nothing Then vols = S_RssMerge(volWs, i)

        Dim anaRow As Long: anaRow = 0
        If anaDic.Exists(code) Then anaRow = CLng(anaDic(code))
        ' RSI（分析シートから取得、なければVBAで直接計算）
        Dim rsiVal As Double: rsiVal = 50
        If anaRow > 0 Then rsiVal = S_SafeNum(anaWs.Cells(anaRow, 23).Value, 50)
        If rsiVal = 50 Or anaRow = 0 Then rsiVal = S_ComputeRSI(closes, 14)
        If rsiVal >= 70 Then
            totalScore = totalScore + 2: signals = signals & "RSI買われ過ぎ "
        ElseIf rsiVal >= 60 Then
            totalScore = totalScore + 1: signals = signals & "RSI高 "
        ElseIf rsiVal <= 30 Then
            totalScore = totalScore - 1
        End If

        ' EMAトレンド（分析シートから取得、なければVBAで直接計算）
        Dim emaTrend As String: emaTrend = ""
        If anaRow > 0 Then emaTrend = S_SafeStr(anaWs.Cells(anaRow, 30).Value)
        ' 分析シートにない場合はVBAで直接計算
        If emaTrend = "" Then
            Dim e5c As Double: e5c = S_EMA(closes, 5)
            Dim e25c As Double: e25c = S_EMA(closes, 25)
            Dim e75c As Double: e75c = S_EMA(closes, 75)
            If e5c > 0 And e25c > 0 And e75c > 0 Then
                If e5c > e25c And e25c > e75c Then
                    emaTrend = "パーフェクト▲"
                ElseIf e5c < e25c And e25c < e75c Then
                    emaTrend = "下降トレンド▼"
                ElseIf e5c > e25c Then
                    emaTrend = "上昇トレンド▲"
                Else
                    emaTrend = "もみ合い→"
                End If
            End If
        End If
        If emaTrend = "下降トレンド▼" Then
            totalScore = totalScore + 2: signals = signals & "下降トレンド▼ "
        ElseIf emaTrend = "パーフェクト▲" Then
            totalScore = totalScore - 1
        End If

        ' 戻り売り
        Dim osime As String: osime = ""
        If anaRow > 0 Then osime = S_SafeStr(anaWs.Cells(anaRow, 31).Value)
        Dim uwamodori As String: uwamodori = "-"
        If osime = "25日押し目★" Then
            uwamodori = "25日線戻り売り●"
            totalScore = totalScore + 2: signals = signals & "25日線戻り● "
        ElseIf osime = "5日押し目○" Then
            uwamodori = "5日線戻り売り○"
            totalScore = totalScore + 1: signals = signals & "5日線戻り○ "
        End If

        ' 出来高
        Dim volToday As Double: volToday = 0
        Dim vol5avg  As Double: vol5avg = 0
        Dim volRatio As Double: volRatio = 0
        If Not volWs Is Nothing Then
            Dim rssV As Double: rssV = Val(volWs.Cells(i, 3).Value)
            volToday = IIf(rssV > 0, rssV, Val(volWs.Cells(i, 5).Value))
            '★v3：ザラ場は途中までの出来高なので1日分に換算してから比べる
            Dim volCalc As Double
            volCalc = IIf(rssV > 0, volToday / S_VolProgress(), volToday)
            If anaRow > 0 Then vol5avg = S_SafeNum(anaWs.Cells(anaRow, 13).Value, 0)
            If vol5avg > 0 And volCalc > 0 Then
                volRatio = volCalc / vol5avg
                If volRatio >= 2# Then
                    totalScore = totalScore + 3
                    signals = signals & "出来高急増(" & Format(volRatio, "0.0") & "倍) "
                ElseIf volRatio >= 1.5 Then
                    totalScore = totalScore + 2
                    signals = signals & "出来高増(" & Format(volRatio, "0.0") & "倍) "
                End If
            End If
        End If

        ' MACD（分析シートから取得、なければMACDシリーズ計算後に設定）
        Dim macdVal As Double: macdVal = 0
        If anaRow > 0 Then macdVal = S_SafeNum(anaWs.Cells(anaRow, 24).Value, 0)

        ' 年初来安値更新
        Dim lowUpdate As String: lowUpdate = ""
        If anaRow > 0 Then lowUpdate = S_SafeStr(anaWs.Cells(anaRow, 22).Value)
        If InStr(lowUpdate, "安値更新") > 0 Then
            totalScore = totalScore + 2: signals = signals & "年初来安値更新▼ "
        End If

        ' MACDシリーズ
        Dim ema12V As Double, ema26V As Double, macdNow As Double, signalNow As Double
        Dim histNow As Double, histPrev As Double
        S_ComputeMACD closes, ema12V, ema26V, macdNow, signalNow, histNow, histPrev

        ' 分析シートにない場合はMACDシリーズの値を使用
        If macdVal = 0 And anaRow = 0 Then macdVal = macdNow
        If macdVal < 0 Then
            totalScore = totalScore + 1: signals = signals & "MACD- "
        ElseIf macdVal > 0 Then
            totalScore = totalScore - 1
        End If

        Dim emaDown  As Boolean: emaDown = (ema12V < ema26V) And (ema26V > 0)
        Dim macdDown As Boolean: macdDown = (macdNow < signalNow)
        Dim rsiDown  As Boolean: rsiDown = (rsiVal < 50)

        Dim ok_DC As Boolean: ok_DC = S_IsDeadCross(closes, 5)
        If ok_DC Then totalScore = totalScore + S_DC_BONUS: signals = signals & "[DC直近5日] "

        Dim ok_PD As Boolean: ok_PD = S_IsPerfectDown(closes)
        If ok_PD Then totalScore = totalScore + S_PD_BONUS: signals = signals & "[完全下降] "

        Dim ok_VBB As Boolean: ok_VBB = False
        If Not IsEmpty(opens) And Not IsEmpty(vols) Then
            ok_VBB = S_IsVolBear(vols, closes, opens)
            If ok_VBB Then totalScore = totalScore + S_VOL_BONUS: signals = signals & "[Vol急増陰線] "
        End If

        Dim adxVal As Double: adxVal = 0
        Dim ok_ADX As Boolean: ok_ADX = False
        If Not IsEmpty(highs) And Not IsEmpty(loWs) Then
            adxVal = S_ComputeADX(highs, loWs, closes)
            ok_ADX = (adxVal > 25#)
            If ok_ADX Then totalScore = totalScore + S_ADX_BONUS: signals = signals & "[ADX=" & Format(adxVal, "0") & "] "
        End If

        Dim ok_HE As Boolean: ok_HE = (histNow < histPrev) And (histNow < 0)
        If ok_HE Then totalScore = totalScore + S_HIST_BONUS: signals = signals & "[Hist悪化] "

        Dim momCount As Long: momCount = 0
        If emaDown Then momCount = momCount + 1: signals = signals & "EMA12<26 "
        If macdDown Then momCount = momCount + 1: signals = signals & "MACD<Sig "
        If rsiDown Then momCount = momCount + 1: signals = signals & "RSI<50 "
        If momCount = 3 Then
            totalScore = totalScore + 2: signals = signals & "[◎売りモメンタム完全] "
        ElseIf momCount = 2 Then
            totalScore = totalScore + 1: signals = signals & "[○売りモメンタム部分] "
        End If

        totalScore = totalScore + S_CandleBear(opens, highs, loWs, closes, signals)
        totalScore = totalScore + S_CandleBearPrev(opens, highs, loWs, closes, signals)

        Dim stockPct As Double: stockPct = 0
        If Not IsEmpty(closes) Then
            If UBound(closes, 2) >= 2 Then
                Dim prevClose As Double: prevClose = S_SafeNum(closes(1, 2), 0)
                If prevClose > 0 Then stockPct = (closePrice - prevClose) / prevClose * 100#
            End If
        End If
        If stockPct < topixPct Then
            totalScore = totalScore + 1: signals = signals & "[対TOPIX弱] "
        End If

        Dim canExtract As Boolean: canExtract = False
        If totalScore >= minScore Then
            Select Case topixMode
                Case "LARGE": canExtract = True
                Case "MID": If emaTrend <> "パーフェクト▲" Then canExtract = True
                Case Else: If emaTrend = "下降トレンド▼" Or ok_PD Then canExtract = True
            End Select
        End If
        If canExtract And rsiVal < 15 Then canExtract = False
        If canExtract And rsiVal >= 80 Then canExtract = False
        If canExtract And stockPct >= topixPct Then canExtract = False
        If doMark And canExtract Then
            If momCount < 2 Then canExtract = False
            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False
        End If

        If canExtract And code <> "" Then
            Dim meiName As String
            If nameDic.Exists(code) Then
                meiName = CStr(nameDic(code))
            Else
                meiName = S_SafeStr(closeWs.Cells(i, 2).Value)
            End If
            If meiName = "" Then meiName = "(" & code & ")"

            Dim currentPrice As Double: currentPrice = closePrice
            Dim currentHigh As Double: currentHigh = highVal
            Dim currentLow  As Double: currentLow = lowVal
            Dim currentVol  As Double: currentVol = volToday

            S_WriteRow extWs, outRow, code, meiName, currentPrice, currentHigh, currentLow, _
                       currentVol, rsiVal, emaTrend, uwamodori, volRatio, macdVal, _
                       lowUpdate, totalScore, signals
            outRow = outRow + 1
        End If
NextRSS:
        DoEvents
    Next i

    ' 空白行削除
    Dim delRow As Long
    For delRow = outRow - 1 To 4 Step -1
        If Trim(CStr(extWs.Cells(delRow, 3).Value)) = "" Then
            extWs.Rows(delRow).Delete
            outRow = outRow - 1
        End If
    Next delRow

    ' ソート
    If outRow > 6 Then
        On Error Resume Next
        extWs.Range(extWs.Cells(4, 1), extWs.Cells(outRow - 1, 18)).Sort _
            Key1:=extWs.Cells(4, 17), Order1:=xlDescending, Header:=xlNo
        On Error GoTo 0
    End If

    '★上位N件だけ残す（スコア降順ソート後）
    If outRow - 4 > maxCount Then
        extWs.Range(extWs.Cells(4 + maxCount, 1), extWs.Cells(outRow - 1, 18)).ClearContents
        extWs.Range(extWs.Cells(4 + maxCount, 1), extWs.Cells(outRow - 1, 18)).Interior.ColorIndex = xlNone
        outRow = 4 + maxCount
    End If

    '★順位列（1,2,3…）を B列(2) に出力
    extWs.Cells(3, 2).Value = "順位"
    Dim rkRow As Long
    For rkRow = 4 To outRow - 1
        extWs.Cells(rkRow, 2).Value = rkRow - 3
        extWs.Cells(rkRow, 2).HorizontalAlignment = xlCenter
        extWs.Cells(rkRow, 2).Font.Bold = True
    Next rkRow

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim extracted As Long: extracted = outRow - 4
    Dim modeDesc As String
    Select Case topixMode
        Case "LARGE": modeDesc = "【相場下落】"
        Case "MID":   modeDesc = "【中程度】"
        Case Else:    modeDesc = "【通常】"
    End Select

    ' 2行目にTOPIX情報を書く（買い抽出シートと同じ構造）
    On Error Resume Next
    extWs.Cells(2, 1).MergeArea.UnMerge
    On Error GoTo 0
    extWs.Range("A2:R2").Merge
    extWs.Cells(2, 1).Value = "TOPIX: " & topixTrend & _
        " (" & Format(topixPct, "+0.00;-0.00;0.00") & "%) " & modeDesc & _
        "　最終更新: " & Format(Now(), "yyyy/mm/dd hh:mm")
    With extWs.Cells(2, 1)
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Color = RGB(255, 215, 0)
        .Interior.Color = RGB(40, 0, 0)
        .HorizontalAlignment = xlRight
    End With

    UI_Msg "【売り候補抽出 RSS現在値版】" & vbCrLf & _
           "TOPIX: " & Format(topixPct, "+0.00;-0.00;0.00") & "% " & modeDesc & vbCrLf & _
           "スキャン: " & scannedCount & " 銘柄" & vbCrLf & _
           "抽出:    " & extracted & " 件 (スコア" & minScore & "点以上)" & vbCrLf & _
           "※RSS現在値使用", _
           vbInformation, "売り抽出完了(RSS版)"
End Sub

'=============================================================================
' RSS現在値を先頭にした配列作成
'=============================================================================
Private Function S_RssMerge(ByVal ws As Worksheet, ByVal r As Long) As Variant
    Dim histLen As Long: histLen = 60
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
    S_RssMerge = arr
End Function

'=============================================================================
' 出力行書き込み
'=============================================================================
Private Sub S_WriteRow(ByVal ws As Worksheet, ByVal r As Long, _
                        ByVal code As String, ByVal meiName As String, _
                        ByVal closePrice As Double, ByVal highVal As Double, _
                        ByVal lowVal As Double, ByVal volToday As Double, _
                        ByVal rsiVal As Double, ByVal emaTrend As String, _
                        ByVal modori As String, ByVal volRatio As Double, _
                        ByVal macdVal As Double, ByVal lowUpdate As String, _
                        ByVal totalScore As Long, ByVal signals As String)
    With ws
        .Cells(r, 1).Formula = "=HYPERLINK(""https://kabutan.jp/stock/?code=""&TEXT(C" & r & ",""0000""),""株探"")"
        .Cells(r, 1).Interior.Color = RGB(128, 0, 0)
        .Cells(r, 1).Font.Color = RGB(255, 255, 255)
        .Cells(r, 1).Font.Bold = True
        .Cells(r, 1).HorizontalAlignment = xlCenter
        .Cells(r, 3).Value = code
        .Cells(r, 3).NumberFormat = "0000"
        .Cells(r, 4).Value = meiName
        .Cells(r, 4).Font.Bold = True
        .Cells(r, 5).Value = closePrice
        .Cells(r, 5).NumberFormat = "#,##0"
        .Cells(r, 6).Value = highVal
        .Cells(r, 6).NumberFormat = "#,##0"
        .Cells(r, 7).Value = lowVal
        .Cells(r, 7).NumberFormat = "#,##0"
        .Cells(r, 8).Value = volToday
        .Cells(r, 8).NumberFormat = "#,##0"
        .Cells(r, 9).Value = Format(rsiVal, "0.0")
        .Cells(r, 10).Value = emaTrend
        Select Case emaTrend
            Case "下降トレンド▼"
                .Cells(r, 10).Interior.Color = RGB(192, 0, 0)
                .Cells(r, 10).Font.Color = RGB(255, 255, 255)
                .Cells(r, 10).Font.Bold = True
            Case "パーフェクト▲"
                .Cells(r, 10).Interior.Color = RGB(0, 176, 80)
                .Cells(r, 10).Font.Color = RGB(255, 255, 255)
        End Select
        .Cells(r, 11).Value = modori
        .Cells(r, 11).HorizontalAlignment = xlCenter
        If modori = "25日線戻り売り●" Then
            .Cells(r, 11).Interior.Color = RGB(192, 0, 0)
            .Cells(r, 11).Font.Color = RGB(255, 255, 255)
            .Cells(r, 11).Font.Bold = True
        ElseIf modori = "5日線戻り売り○" Then
            .Cells(r, 11).Interior.Color = RGB(255, 150, 150)
        End If
        .Cells(r, 12).Value = IIf(volRatio > 0, Format(volRatio, "0.00") & "倍", "-")
        .Cells(r, 13).Value = Format(macdVal, "0.00")
        .Cells(r, 14).Value = lowUpdate
        .Cells(r, 14).HorizontalAlignment = xlCenter
        If InStr(lowUpdate, "安値更新") > 0 Then
            .Cells(r, 14).Font.Color = RGB(192, 0, 0)
            .Cells(r, 14).Font.Bold = True
        End If
        .Cells(r, 15).Value = Round(closePrice * 0.92, 0)  'O:利確ライン -8%
        .Cells(r, 15).NumberFormat = "#,##0"
        .Cells(r, 16).Value = Round(closePrice * 1.04, 0)  'P:損切ライン +4%
        .Cells(r, 16).NumberFormat = "#,##0"
        .Cells(r, 17).Value = totalScore
        Dim bg As Long
        Select Case totalScore
            Case Is >= 14: bg = RGB(80, 0, 0)
            Case Is >= 12: bg = RGB(128, 0, 0)
            Case Is >= 10: bg = RGB(192, 0, 0)
            Case Is >= 8:  bg = RGB(237, 85, 49)
            Case Else:     bg = RGB(200, 100, 50)
        End Select
        .Cells(r, 17).Interior.Color = bg
        .Cells(r, 17).Font.Color = RGB(255, 255, 255)
        .Cells(r, 17).Font.Bold = True
        .Cells(r, 17).HorizontalAlignment = xlCenter
        .Cells(r, 18).Value = Trim$(signals)
        .Range(.Cells(r, 1), .Cells(r, 18)).Font.Name = "Meiryo UI"
        .Range(.Cells(r, 1), .Cells(r, 18)).Font.Size = 18
        ' 行の背景色：奇数行=白、偶数行=薄いピンク
        Dim cc As Long
        For cc = 2 To 18
            If cc <> 10 And cc <> 11 And cc <> 14 And cc <> 17 Then
                If r Mod 2 = 0 Then
                    .Cells(r, cc).Interior.Color = RGB(255, 220, 220)
                Else
                    .Cells(r, cc).Interior.Color = RGB(255, 255, 255)
                End If
            End If
        Next cc
    End With
End Sub

'=============================================================================
' ヘルパー関数
'=============================================================================
Private Function S_ComputeRSI(ByRef prices As Variant, ByVal period As Long) As Double
    S_ComputeRSI = 50
    On Error GoTo ExitF
    If IsEmpty(prices) Then Exit Function
    Dim n As Long: n = UBound(prices, 2)
    If n < period + 1 Then Exit Function
    Dim gains As Double: gains = 0
    Dim losses As Double: losses = 0
    Dim i As Long
    ' 最新のperiod日分を逆順で計算
    For i = 1 To period
        If UBound(prices, 2) >= i + 1 Then
            Dim chg As Double
            chg = CDbl(prices(1, i)) - CDbl(prices(1, i + 1))
            If chg > 0 Then gains = gains + chg Else losses = losses - chg
        End If
    Next i
    gains = gains / period
    losses = losses / period
    If losses = 0 Then S_ComputeRSI = 100: Exit Function
    S_ComputeRSI = 100 - (100 / (1 + gains / losses))
ExitF:
End Function

Private Function S_GetWs(ByVal n As String) As Worksheet
    On Error Resume Next
    Set S_GetWs = ThisWorkbook.Sheets(n)
    On Error GoTo 0
End Function

Private Function S_SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        S_SafeStr = ""
    Else
        S_SafeStr = Trim$(CStr(v))
    End If
End Function

Private Function S_SafeNum(ByVal v As Variant, ByVal d As Double) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        S_SafeNum = d
    ElseIf IsNumeric(v) Then
        S_SafeNum = CDbl(v)
    Else
        S_SafeNum = d
    End If
End Function

Private Function S_CandleBear(ByRef opens As Variant, ByRef highs As Variant, _
                               ByRef loWs As Variant, ByRef closes As Variant, _
                               ByRef sig As String) As Long
    S_CandleBear = 0
    On Error GoTo ExitF
    If IsEmpty(opens) Or IsEmpty(highs) Or IsEmpty(loWs) Or IsEmpty(closes) Then Exit Function
    Dim o As Double: o = CDbl(opens(1, 1))
    Dim h As Double: h = CDbl(highs(1, 1))
    Dim L As Double: L = CDbl(loWs(1, 1))
    Dim c As Double: c = CDbl(closes(1, 1))
    If o <= 0 Or h <= 0 Or L <= 0 Or c <= 0 Then Exit Function
    Dim rng As Double: rng = h - L
    If rng <= 0 Then Exit Function
    Dim body As Double: body = Abs(c - o)
    Dim up As Double: up = h - WorksheetFunction.Max(o, c)
    Dim lo As Double: lo = WorksheetFunction.Min(o, c) - L
    If body > 0 And up >= body * 2# And lo <= body * 0.5 And c < o Then
        S_CandleBear = 2: sig = sig & "[上ヒゲ陰線▼+2] "
    ElseIf c < o And body >= rng * 0.7 Then
        S_CandleBear = 2: sig = sig & "[大陰線▼+2] "
    ElseIf up >= rng * 0.5 And lo <= rng * 0.2 Then
        S_CandleBear = 1: sig = sig & "[長い上ヒゲ+1] "
    ElseIf body > 0 And lo >= body * 2# And up <= body * 0.5 Then
        S_CandleBear = -2: sig = sig & "[下ヒゲ陽線-2] "
    End If
ExitF:
End Function

Private Function S_CandleBearPrev(ByRef opens As Variant, ByRef highs As Variant, _
                                   ByRef loWs As Variant, ByRef closes As Variant, _
                                   ByRef sig As String) As Long
    S_CandleBearPrev = 0
    On Error GoTo ExitF
    If IsEmpty(opens) Or IsEmpty(highs) Or IsEmpty(loWs) Or IsEmpty(closes) Then Exit Function
    If UBound(opens, 2) < 2 Then Exit Function
    Dim o As Double: o = CDbl(opens(1, 2))
    Dim h As Double: h = CDbl(highs(1, 2))
    Dim L As Double: L = CDbl(loWs(1, 2))
    Dim c As Double: c = CDbl(closes(1, 2))
    If o <= 0 Or h <= 0 Or L <= 0 Or c <= 0 Then Exit Function
    Dim rng As Double: rng = h - L
    If rng <= 0 Then Exit Function
    Dim body As Double: body = Abs(c - o)
    Dim up As Double: up = h - WorksheetFunction.Max(o, c)
    Dim lo As Double: lo = WorksheetFunction.Min(o, c) - L
    If body > 0 And up >= body * 2# And lo <= body * 0.5 And c < o Then
        S_CandleBearPrev = 2: sig = sig & "[前日:上ヒゲ陰線▼+2] "
    ElseIf c < o And body >= rng * 0.7 Then
        S_CandleBearPrev = 2: sig = sig & "[前日:大陰線▼+2] "
    ElseIf up >= rng * 0.5 And lo <= rng * 0.2 Then
        S_CandleBearPrev = 1: sig = sig & "[前日:長い上ヒゲ+1] "
    ElseIf body > 0 And lo >= body * 2# And up <= body * 0.5 Then
        S_CandleBearPrev = -2: sig = sig & "[前日:下ヒゲ陽線-2] "
    End If
ExitF:
End Function

Private Function S_IsDeadCross(ByRef prices As Variant, ByVal lookback As Long) As Boolean
    S_IsDeadCross = False
    On Error GoTo ExitF
    If IsEmpty(prices) Then Exit Function
    Dim n As Long: n = UBound(prices, 2)
    If n < 26 Then Exit Function
    Dim k5 As Double: k5 = 2# / 6#
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
    If e5Arr(cnt) >= e25Arr(cnt) Then Exit Function
    Dim lookStart As Long: lookStart = cnt - lookback
    If lookStart < 1 Then lookStart = 1
    Dim j As Long
    For j = lookStart To cnt - 1
        If e5Arr(j) >= e25Arr(j) Then
            S_IsDeadCross = True: Exit Function
        End If
    Next j
ExitF:
End Function

Private Function S_IsPerfectDown(ByRef prices As Variant) As Boolean
    S_IsPerfectDown = False
    On Error GoTo ExitF
    If IsEmpty(prices) Then Exit Function
    Dim e5  As Double: e5 = S_EMA(prices, 5)
    Dim e25 As Double: e25 = S_EMA(prices, 25)
    Dim e75 As Double: e75 = S_EMA(prices, 75)
    If e5 > 0 And e25 > 0 And e75 > 0 Then
        S_IsPerfectDown = (e5 < e25) And (e25 < e75)
    End If
ExitF:
End Function

Private Function S_EMA(ByRef prices As Variant, ByVal period As Long) As Double
    S_EMA = 0
    On Error GoTo ExitF
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
    S_EMA = ema
ExitF:
End Function

Private Function S_IsVolBear(ByRef vols As Variant, ByRef closes As Variant, _
                              ByRef opens As Variant) As Boolean
    S_IsVolBear = False
    On Error GoTo ExitF
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
    If todayClose >= todayOpen Then Exit Function
    Dim endIdx As Long: endIdx = IIf(n < 21, n, 21)
    Dim sm As Double: sm = 0
    Dim cnt As Long: cnt = 0
    Dim j As Long
    For j = 2 To endIdx
        If IsNumeric(vols(1, j)) Then
            Dim vV As Double: vV = CDbl(vols(1, j))
            If vV > 0 Then sm = sm + vV: cnt = cnt + 1
        End If
    Next j
    If cnt < 5 Then Exit Function
    S_IsVolBear = (todayVol >= (sm / cnt) * 1.5)
ExitF:
End Function

Private Sub S_ComputeMACD(ByRef prices As Variant, _
                           ByRef ema12 As Double, ByRef ema26 As Double, _
                           ByRef macdV As Double, ByRef sigV As Double, _
                           ByRef histNow As Double, ByRef histPrev As Double)
    ema12 = 0: ema26 = 0: macdV = 0: sigV = 0: histNow = 0: histPrev = 0
    On Error GoTo ExitP
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
    Dim mCnt As Long: mCnt = 0
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
                mCnt = mCnt + 1
                macdArr(mCnt) = e12 - e26
            End If
        End If
    Next i
    If mCnt < 2 Then Exit Sub
    ema12 = e12: ema26 = e26: macdV = macdArr(mCnt)
    Dim sg As Double: sg = macdArr(1)
    Dim sgPrev As Double: sgPrev = sg
    Dim j As Long
    For j = 2 To mCnt
        sgPrev = sg
        sg = k9 * macdArr(j) + (1# - k9) * sg
    Next j
    sigV = sg
    histNow = macdArr(mCnt) - sg
    histPrev = macdArr(mCnt - 1) - sgPrev
ExitP:
End Sub

Private Function S_ComputeADX(ByRef highs As Variant, ByRef loWs As Variant, _
                                ByRef closes As Variant) As Double
    S_ComputeADX = 0
    On Error GoTo ExitF
    If IsEmpty(highs) Or IsEmpty(loWs) Or IsEmpty(closes) Then Exit Function
    Dim n As Long: n = UBound(highs, 2)
    If n < 28 Then Exit Function
    Const p As Long = 14
    ReDim hA(1 To n) As Double
    ReDim lA(1 To n) As Double
    ReDim cA(1 To n) As Double
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = n To 1 Step -1
        If IsNumeric(highs(1, i)) And IsNumeric(loWs(1, i)) And IsNumeric(closes(1, i)) Then
            Dim hv As Double: hv = CDbl(highs(1, i))
            Dim lv As Double: lv = CDbl(loWs(1, i))
            Dim cv As Double: cv = CDbl(closes(1, i))
            If hv > 0 And lv > 0 And cv > 0 Then
                cnt = cnt + 1
                hA(cnt) = hv: lA(cnt) = lv: cA(cnt) = cv
            End If
        End If
    Next i
    If cnt < 28 Then Exit Function
    ReDim trA(2 To cnt) As Double
    ReDim pdA(2 To cnt) As Double
    ReDim mdA(2 To cnt) As Double
    Dim j As Long
    For j = 2 To cnt
        Dim hd As Double: hd = hA(j) - hA(j - 1)
        Dim ld As Double: ld = lA(j - 1) - lA(j)
        pdA(j) = IIf(hd > ld And hd > 0, hd, 0)
        mdA(j) = IIf(ld > hd And ld > 0, ld, 0)
        Dim t1 As Double: t1 = hA(j) - lA(j)
        Dim t2 As Double: t2 = Abs(hA(j) - cA(j - 1))
        Dim t3 As Double: t3 = Abs(lA(j) - cA(j - 1))
        trA(j) = t1
        If t2 > trA(j) Then trA(j) = t2
        If t3 > trA(j) Then trA(j) = t3
    Next j
    If cnt < 2 + p - 1 Then Exit Function
    Dim atr As Double: atr = 0
    Dim pdm As Double: pdm = 0
    Dim mdm As Double: mdm = 0
    For j = 2 To 1 + p
        atr = atr + trA(j): pdm = pdm + pdA(j): mdm = mdm + mdA(j)
    Next j
    ReDim dxA(1 To cnt) As Double
    Dim dxCnt As Long: dxCnt = 0
    Dim pDI As Double, mDI As Double, sDI As Double
    If atr > 0 Then
        pDI = 100# * pdm / atr: mDI = 100# * mdm / atr: sDI = pDI + mDI
        If sDI > 0 Then dxCnt = dxCnt + 1: dxA(dxCnt) = 100# * Abs(pDI - mDI) / sDI
    End If
    For j = 2 + p To cnt
        atr = atr - (atr / p) + trA(j)
        pdm = pdm - (pdm / p) + pdA(j)
        mdm = mdm - (mdm / p) + mdA(j)
        If atr > 0 Then
            pDI = 100# * pdm / atr: mDI = 100# * mdm / atr: sDI = pDI + mDI
            If sDI > 0 Then dxCnt = dxCnt + 1: dxA(dxCnt) = 100# * Abs(pDI - mDI) / sDI
        End If
    Next j
    If dxCnt < p Then Exit Function
    Dim adx As Double: adx = 0
    For j = 1 To p: adx = adx + dxA(j): Next j
    adx = adx / p
    For j = p + 1 To dxCnt: adx = (adx * (p - 1) + dxA(j)) / p: Next j
    S_ComputeADX = adx
ExitF:
End Function

'=============================================================================
' ★v3追加：補助
'=============================================================================
'   RSSが応答しているかを実際に問い合わせて確かめる。
'   作業セル1つに RssMarket の式を入れ、そのシートだけを計算して確認する。
'   他のシート・他のBOOKには一切触らない。
Private Function S_RSS生存(ByVal ws As Worksheet) As Boolean
    S_RSS生存 = False
    If ws Is Nothing Then Exit Function
    On Error GoTo Fail
    Dim sc As Range
    Set sc = ws.Cells(1, S_PROBE_COL)
    sc.Formula = "=IFERROR(RssMarket(""7203"",""現在値""),-1)"
    ws.Calculate
    S_RSS生存 = (Val(sc.Value) > 0)
    sc.ClearContents
    Exit Function
Fail:
    On Error Resume Next
    ws.Cells(1, S_PROBE_COL).ClearContents
    On Error GoTo 0
End Function

'   出来高の進捗率。いまの時刻までに1日の出来高の何割が出来ているかの目安。
'   ※ 実測ではなく一般的な目安。値を変えたいときはこの表を直す。
Private Function S_VolProgress() As Double
    Dim hh As Variant, pp As Variant
    hh = Array("09:00", "09:05", "09:30", "10:00", "11:00", "11:30", _
               "12:30", "13:00", "14:00", "14:30", "15:00", "15:20", "15:25", "15:30")
    pp = Array(0.02, 0.1, 0.25, 0.33, 0.42, 0.47, _
               0.47, 0.52, 0.62, 0.68, 0.78, 0.88, 0.93, 1#)

    Dim t As Double: t = CDbl(TimeValue(Format(Now, "hh:mm:ss")))
    Dim i As Long, t0 As Double, t1 As Double

    If t < CDbl(TimeValue(CStr(hh(0)))) Then S_VolProgress = 1#: Exit Function
    If t >= CDbl(TimeValue(CStr(hh(UBound(hh))))) Then S_VolProgress = 1#: Exit Function

    For i = 0 To UBound(hh) - 1
        t0 = CDbl(TimeValue(CStr(hh(i))))
        t1 = CDbl(TimeValue(CStr(hh(i + 1))))
        If t >= t0 And t < t1 Then
            If t1 > t0 Then
                S_VolProgress = CDbl(pp(i)) + (CDbl(pp(i + 1)) - CDbl(pp(i))) * (t - t0) / (t1 - t0)
            Else
                S_VolProgress = CDbl(pp(i))
            End If
            Exit Function
        End If
    Next i
    S_VolProgress = 1#
End Function

'   安全装置（Mod_安全装置）があれば実行前にRSSの生死を確認する。
'   無ければ何もせず True を返すので、従来どおり動く。
Private Function S_安全確認(ByVal nm As String) As Boolean
    S_安全確認 = True
    If gUnattended Then Exit Function
    Dim r As Variant
    On Error Resume Next
    r = Application.Run("安全装置_RSS必要", nm)
    If Err.Number = 0 Then S_安全確認 = CBool(r)
    Err.Clear
    On Error GoTo 0
End Function
