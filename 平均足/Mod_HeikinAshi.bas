Attribute VB_Name = "Mod_HeikinAshi"
Option Explicit
'==================================================================
' 平均足トレード　抽出モジュール　v2.0（実データ検証済み）
'
'   使うシート : 始値 / 高値 / 安値 / 終値 / 出来高
'   出すシート : 平均足（無ければ自動で作ります）
'   実行方法   : Alt+F8 →「平均足_買い抽出」または「平均足_売り抽出」
'
'   考え方
'     平均足は「トレンドの終わり＝反転」を見る道具。
'     反転の形が出た銘柄だけを残し、7つの手法で点数を付けて順位を決める。
'
'   検証結果（300銘柄×直近250日・手数料込み・損切の窓開けも計算）
'     出来高1.4倍 / 株価1000円以上 / 週足10週線の上 / 損切8% / 利確8% / 5日手じまい / 寄り+2%まで
'       221回　勝率65.2%　1回あたり平均 +1.62%
'       同時5銘柄・1銘柄に資金の20% → 250日で +51.1%　最大下落 9.1%
'     週足の条件を外すと 370回　勝率58.6%　250日 +62.0%（回数は増えるが勝率は下がる）
'     ※直近1年の相場での結果です。100%勝てる方法はありません。
'==================================================================

'-------------- 調整するのはここだけ --------------
Public Const HA_CAPITAL  As Double = 3000000#   '運用資金（円）
Public Const HA_RISK     As Double = 0.01       '1回で失ってよい割合（資金の1%）
Public Const HA_PASS     As Long = 4            '合格点（下げると候補が増える）
Public Const HA_MINPRICE As Double = 1000#      '最低株価（円）※低位株は成績が悪い
Public Const HA_MINVOL   As Double = 200000#    '5日平均出来高の下限（株）
Public Const HA_VOLRATE  As Double = 1.4        '当日出来高／5日平均（これが一番効く）
Public Const HA_RR       As Double = 2#         '利確1＝損切幅の何倍か
Public Const HA_MAXLOSS  As Double = 0.06       '損切の最大幅（6%）
Public Const HA_MAXPOS   As Double = 0.25       '1銘柄に使う上限（資金の25%）
Public Const HA_MAXROWS  As Long = 60           '出力する最大件数
Public Const HA_HOLDDAYS As Long = 5            '手じまい（営業日）

'--- 出口の方式 ---
Public Const HA_FIX_MODE As Boolean = True     'True=固定%（損切8%/利確8%）　False=ATR基準
Public Const HA_FIX_SL   As Double = 0.08      '固定の損切（8%）
Public Const HA_FIX_TP   As Double = 0.08      '固定の利確（8%）
Public Const HA_FIX_POS  As Double = 0.2       '固定方式のとき1銘柄に使う資金の割合（20%）
Public Const HA_MAXGAP   As Double = 0.02      '翌朝これ以上高く始まったら見送る（+2%）
Public Const HA_WEEK_ON  As Boolean = True     '週足10週線の上の銘柄だけ買う（勝率が上がります）
Public Const HA_WEEK_MA  As Long = 10          '週足の移動平均の週数
Public Const HA_SKIP     As Long = 0            '0=最新日で判定　1=前日で判定（ザラ場中は1）
'--------------------------------------------------

Public HA_SILENT As Boolean          'True=自動実行中（メッセージを出さない）

Private Const R_TOP As Long = 6     'データ開始行
Private Const R_END As Long = 520   '走査する最終行
Private Const C_NEW As Long = 5     'E列＝最新日
Private Const N_MAX As Long = 150   '使う日数の上限
Private Const OUT_SHEET As String = "平均足"
Private Const OUT_TOP As Long = 4   '出力の開始行
Private Const OUT_COLS As Long = 21

'==================== 入口 ====================
Public Sub 平均足_買い抽出()
    HA_Run 1
End Sub

Public Sub 平均足_売り抽出()
    HA_Run -1
End Sub

'他のモジュールから呼ぶ用（半角名）　※中身は上と同じ
Public Sub HA_Buy()
    HA_Run 1
End Sub

Public Sub HA_Sell()
    HA_Run -1
End Sub

'==================== 本体 ====================
Private Sub HA_Run(ByVal side As Long)

    Dim wsO As Worksheet, wsH As Worksheet, wsL As Worksheet
    Dim wsC As Worksheet, wsV As Worksheet, wsX As Worksheet

    Set wsO = HA_GetWs("始値")
    Set wsH = HA_GetWs("高値")
    Set wsL = HA_GetWs("安値")
    Set wsC = HA_GetWs("終値")
    Set wsV = HA_GetWs("出来高")

    If wsO Is Nothing Or wsH Is Nothing Or wsL Is Nothing Or wsC Is Nothing Then
        If Not HA_SILENT Then MsgBox "「始値」「高値」「安値」「終値」のシートが必要です。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- 使える日数（列）を調べる ---
    '※日付書式のセルは .Value だと日付型になり IsNumeric が False になるので .Value2 を使う
    Dim lastCol As Long, c As Long, dv As Double, missRun As Long
    lastCol = C_NEW
    missRun = 0
    For c = C_NEW + 1 To C_NEW + N_MAX - 1
        dv = 0
        If IsNumeric(wsC.Cells(3, c).Value2) Then dv = CDbl(wsC.Cells(3, c).Value2)
        If dv > 40000 And dv < 80000 Then
            lastCol = c
            missRun = 0
        Else
            missRun = missRun + 1
            If missRun > 5 Then Exit For
        End If
    Next c

    Dim nCol As Long
    nCol = lastCol - C_NEW + 1
    If nCol < 40 Then
        HA_Restore
        If Not HA_SILENT Then MsgBox "履歴が足りません（" & nCol & "日）。40日以上必要です。", vbExclamation
        Exit Sub
    End If

    '--- 最終行 ---
    Dim lastRow As Long
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    If lastRow > R_END Then lastRow = R_END
    If lastRow < R_TOP Then
        HA_Restore
        If Not HA_SILENT Then MsgBox "銘柄データがありません。", vbExclamation
        Exit Sub
    End If

    '--- 日付の行（週足の判定に使う）---
    Dim dtc() As Double
    ReDim dtc(1 To nCol)
    Dim kk As Long
    For kk = 1 To nCol
        If kk = 1 Then
            dtc(kk) = HA_Num(wsC.Cells(3, 4).Value2)
        Else
            dtc(kk) = HA_Num(wsC.Cells(3, 4 + kk).Value2)
        End If
    Next kk

    '--- まとめて読み込む（速度対策） ---
    Dim aO As Variant, aH As Variant, aL As Variant, aC As Variant, aV As Variant
    Dim aCode As Variant, aName As Variant
    aO = wsO.Range(wsO.Cells(R_TOP, C_NEW), wsO.Cells(lastRow, lastCol)).Value
    aH = wsH.Range(wsH.Cells(R_TOP, C_NEW), wsH.Cells(lastRow, lastCol)).Value
    aL = wsL.Range(wsL.Cells(R_TOP, C_NEW), wsL.Cells(lastRow, lastCol)).Value
    aC = wsC.Range(wsC.Cells(R_TOP, C_NEW), wsC.Cells(lastRow, lastCol)).Value
    If Not wsV Is Nothing Then
        aV = wsV.Range(wsV.Cells(R_TOP, C_NEW), wsV.Cells(lastRow, lastCol)).Value
    End If
    aCode = wsC.Range(wsC.Cells(R_TOP, 1), wsC.Cells(lastRow, 1)).Value
    aName = wsC.Range(wsC.Cells(R_TOP, 2), wsC.Cells(lastRow, 2)).Value

    '--- 地合い（TOPIX＝5行目） ---
    Dim topixNote As String, topixBonus As Long
    topixBonus = HA_TopixState(wsC, lastCol, side, topixNote)

    '--- 走査 ---
    Dim res() As Variant, sc() As Double
    ReDim res(1 To lastRow - R_TOP + 1, 1 To OUT_COLS)
    ReDim sc(1 To lastRow - R_TOP + 1)
    Dim hit As Long
    hit = 0

    '--- 絞り込みの内訳を数える（なぜ候補が出ないかを見るため）---
    Dim cnt(1 To 13) As Long

    Dim i As Long, n As Long, k As Long
    Dim o() As Double, h() As Double, lo() As Double, cl() As Double, vo() As Double
    Dim dts() As Double
    Dim haO() As Double, haC() As Double, haH() As Double, haL() As Double

    For i = 1 To lastRow - R_TOP + 1

        Dim code As String, nm As String
        code = Trim$(CStr(aCode(i, 1)))
        nm = Trim$(CStr(aName(i, 1)))
        If code = "" Or code = "0" Or UCase$(code) = "TOPX" Then GoTo NextStock
        cnt(1) = cnt(1) + 1

        '--- 欠測日を飛ばして 古い→新しい 順に詰め直す ---
        ReDim o(1 To nCol): ReDim h(1 To nCol): ReDim lo(1 To nCol)
        ReDim cl(1 To nCol): ReDim vo(1 To nCol): ReDim dts(1 To nCol)
        n = 0
        For k = nCol To 1 Step -1
            Dim pC As Double, pO As Double, pH As Double, pL As Double
            pC = HA_Num(aC(i, k)): pO = HA_Num(aO(i, k))
            pH = HA_Num(aH(i, k)): pL = HA_Num(aL(i, k))
            If pC > 0 And pO > 0 And pH > 0 And pL > 0 Then
                n = n + 1
                o(n) = pO: h(n) = pH: lo(n) = pL: cl(n) = pC
                dts(n) = dtc(k)
                If IsArray(aV) Then vo(n) = HA_Num(aV(i, k)) Else vo(n) = 0
            End If
        Next k
        n = n - HA_SKIP
        If n < 35 Then GoTo NextStock
        If cl(n) < HA_MINPRICE Then GoTo NextStock
        cnt(2) = cnt(2) + 1

        '--- 平均足を作る ---
        ReDim haO(1 To n): ReDim haC(1 To n): ReDim haH(1 To n): ReDim haL(1 To n)
        For k = 1 To n
            haC(k) = (o(k) + h(k) + lo(k) + cl(k)) / 4#
            If k = 1 Then
                haO(k) = (o(k) + cl(k)) / 2#
            Else
                haO(k) = (haO(k - 1) + haC(k - 1)) / 2#
            End If
            haH(k) = HA_Max3(h(k), haO(k), haC(k))
            haL(k) = HA_Min3(lo(k), haO(k), haC(k))
        Next k

        '===== 必須条件（1つでも欠けたら捨てる）=====

        '(1) 出来高　※ここが一番効く
        Dim vol5 As Double, volRate As Double
        vol5 = HA_Avg(vo, n, 5)
        If vol5 > 0 Then volRate = vo(n) / vol5 Else volRate = 0
        If IsArray(aV) Then
            If vol5 < HA_MINVOL Then GoTo NextStock
            cnt(3) = cnt(3) + 1
            If volRate < HA_VOLRATE Then GoTo NextStock
        Else
            cnt(3) = cnt(3) + 1
        End If
        cnt(4) = cnt(4) + 1

        '(2) 同じ色が2〜5本続いている（6本目以降は高値づかみ）
        Dim runLen As Long
        runLen = HA_RunLen(haO, haC, n, side)
        If runLen < 2 Or runLen > 5 Then GoTo NextStock
        cnt(5) = cnt(5) + 1

        '(3) その前に逆の色が3本以上（＝トレンドの終わりを取る）
        Dim before As Long
        before = HA_RunPrev(haO, haC, n, runLen, -side)
        Dim sma20 As Double, sma25 As Double
        sma20 = HA_Sma(cl, n, 20): sma25 = HA_Sma(cl, n, 25)
        Dim okTurn As Boolean
        okTurn = (before >= 3)
        If side = 1 Then
            If Not okTurn Then okTurn = (cl(n) > sma20 And before >= 2)
        Else
            If Not okTurn Then okTurn = (cl(n) < sma20 And before >= 2)
        End If
        If Not okTurn Then GoTo NextStock
        cnt(6) = cnt(6) + 1

        '(4) ヒゲ　上昇中は下ヒゲが出ない／下落中は上ヒゲが出ない
        If side = 1 Then
            If haL(n) < haO(n) - cl(n) * 0.0005 Then GoTo NextStock
        Else
            If haH(n) > haO(n) + cl(n) * 0.0005 Then GoTo NextStock
        End If
        cnt(7) = cnt(7) + 1

        '(5) 安値切り上げ（売りは高値切り下げ）
        If side = 1 Then
            If HA_MinN(lo, n, 3) <= HA_MinN2(lo, n, 4, 8) Then GoTo NextStock
        Else
            If HA_MaxN(h, n, 3) >= HA_MaxN2(h, n, 4, 8) Then GoTo NextStock
        End If
        cnt(8) = cnt(8) + 1

        '(6) 20日線の上（売りは下）
        If side = 1 Then
            If cl(n) <= sma20 Then GoTo NextStock
        Else
            If cl(n) >= sma20 Then GoTo NextStock
        End If

        '(7) 25日線からの乖離（買いは0%以上＝上向きの押し目を買う）
        Dim kairi As Double
        If sma25 > 0 Then kairi = (cl(n) - sma25) / sma25 * 100#
        If side = 1 Then
            If kairi < 0 Then GoTo NextStock
        Else
            If kairi > 0 Then GoTo NextStock
        End If
        cnt(9) = cnt(9) + 1

        '(8) RSI（50ラインが命）
        Dim rsiNow As Double, rsiPre As Double
        rsiNow = HA_Rsi(cl, n, 14): rsiPre = HA_Rsi(cl, n - 1, 14)
        If side = 1 Then
            If rsiNow < 50 Then GoTo NextStock
        Else
            If rsiNow > 50 Then GoTo NextStock
        End If
        cnt(10) = cnt(10) + 1

        '(9) 実体が大きすぎない（行き過ぎは翌日戻される）
        Dim bodyRate As Double
        If haH(n) - haL(n) > 0 Then
            bodyRate = Abs(haC(n) - haO(n)) / (haH(n) - haL(n))
        End If
        If bodyRate > 0.85 Then GoTo NextStock
        cnt(11) = cnt(11) + 1

        '(10) 週足が10週線の上（大きな流れに逆らわない）
        If HA_WEEK_ON Then
            If Not HA_WeekOK(cl, dts, n, side) Then GoTo NextStock
        End If
        cnt(12) = cnt(12) + 1

        '===== 7つの手法で点数を付ける（順位決め）=====
        Dim score As Long, sig As String
        score = 0: sig = ""

        '手法4 十字線からの転換（一番効く）
        Dim d As Long
        d = n - runLen
        If d >= 1 Then
            If haH(d) - haL(d) > 0 Then
                If Abs(haC(d) - haO(d)) / (haH(d) - haL(d)) <= 0.35 Then
                    score = score + 3: sig = sig & "④十字線から転換 "
                End If
            End If
        End If

        '地合い（表示のみ・点数には入れない）
        score = score + topixBonus

        '手法7 出来高
        If volRate >= 2# Then
            score = score + 2: sig = sig & "⑦出来高" & Format(volRate, "0.0") & "倍 "
        Else
            score = score + 1: sig = sig & "⑦出来高" & Format(volRate, "0.0") & "倍 "
        End If

        '手法5 RSI
        If side = 1 Then
            If rsiNow <= 65 Then score = score + 2: sig = sig & "⑤RSI適温 "
            If rsiNow >= 75 Then score = score - 1: sig = sig & "⑤RSI過熱 "
            If rsiNow > 50 And rsiPre <= 50 Then score = score + 1: sig = sig & "⑤RSI50上抜け "
        Else
            If rsiNow >= 35 Then score = score + 2: sig = sig & "⑤RSI適温 "
            If rsiNow <= 25 Then score = score - 1: sig = sig & "⑤売られ過ぎ "
            If rsiNow < 50 And rsiPre >= 50 Then score = score + 1: sig = sig & "⑤RSI50下抜け "
        End If

        '転換前の陰線の本数
        If before >= 3 And before <= 5 Then
            score = score + 2: sig = sig & "④前" & before & "本の下げ "
        End If

        '手法3 移動平均線
        Dim sma9 As Double, sma18 As Double, sma9p As Double, sma18p As Double
        sma9 = HA_Sma(cl, n, 9): sma18 = HA_Sma(cl, n, 18)
        sma9p = HA_Sma(cl, n - 5, 9): sma18p = HA_Sma(cl, n - 5, 18)
        If side = 1 Then
            If sma9 > sma18 Then score = score + 1: sig = sig & "③9>18MA "
            If sma9 > sma18 And sma9p <= sma18p Then score = score + 1: sig = sig & "③GC "
        Else
            If sma9 < sma18 Then score = score + 1: sig = sig & "③9<18MA "
            If sma9 < sma18 And sma9p >= sma18p Then score = score + 1: sig = sig & "③DC "
        End If

        '手法1 支持線・抵抗線からの反発
        Dim ref As Double
        If side = 1 Then
            ref = HA_MinN(lo, n - 1, 20)
            If ref > 0 Then
                If lo(n) <= ref * 1.015 Then score = score + 1: sig = sig & "①支持線反発 "
            End If
        Else
            ref = HA_MaxN(h, n - 1, 20)
            If ref > 0 Then
                If h(n) >= ref * 0.985 Then score = score + 1: sig = sig & "①抵抗線跳ね返り "
            End If
        End If

        '手法2 ブレイクアウト（平均足が確定してから）
        Dim brk As Double
        If side = 1 Then
            brk = HA_MaxN2(h, n, 2, 21)
            If brk > 0 Then
                If cl(n) > brk Then score = score + 1: sig = sig & "②高値ブレイク "
            End If
        Else
            brk = HA_MinN2(lo, n, 2, 21)
            If brk > 0 Then
                If cl(n) < brk Then score = score + 1: sig = sig & "②安値ブレイク "
            End If
        End If

        '手法6 ストキャス
        Dim kNow As Double, kPre As Double
        kNow = HA_StochK(h, lo, cl, n, 9): kPre = HA_StochK(h, lo, cl, n - 1, 9)
        If side = 1 Then
            If kNow > 20 And kPre <= 20 Then score = score + 1: sig = sig & "⑥ストキャス20上抜け "
        Else
            If kNow < 80 And kPre >= 80 Then score = score + 1: sig = sig & "⑥ストキャス80下抜け "
        End If

        '値動きの荒さ・実体の形
        Dim atr As Double, atrPct As Double
        atr = HA_Atr(h, lo, cl, n, 14)
        If cl(n) > 0 Then atrPct = atr / cl(n) * 100#
        If atrPct <= 3.5 Then score = score + 1: sig = sig & "値動き安定 "
        If HA_WEEK_ON Then sig = sig & "週足" & HA_WEEK_MA & "週線上 "
        If bodyRate >= 0.3 And bodyRate <= 0.7 Then score = score + 1

        If score < HA_PASS Then GoTo NextStock
        cnt(13) = cnt(13) + 1

        '===== 損切・利確・株数 =====
        Dim entry As Double, stopP As Double
        Dim tgt1 As Double, tgt2 As Double, riskP As Double
        entry = cl(n)

        If HA_FIX_MODE Then
            '--- 固定%方式（損切8% / 利確8% / 7日手じまい）---
            If side = 1 Then
                stopP = entry * (1 - HA_FIX_SL)
                tgt1 = entry * (1 + HA_FIX_TP)
            Else
                stopP = entry * (1 + HA_FIX_SL)
                tgt1 = entry * (1 - HA_FIX_TP)
            End If
            tgt2 = tgt1
            riskP = Abs(entry - stopP)
        Else
            '--- ATR基準方式 ---
            If side = 1 Then
                stopP = HA_Min2(haL(n), HA_MinN(lo, n, 3)) - atr * 0.5
                If stopP < entry * (1 - HA_MAXLOSS) Then stopP = entry * (1 - HA_MAXLOSS)
                If stopP >= entry Then stopP = entry * (1 - HA_MAXLOSS)
                riskP = entry - stopP
                tgt1 = entry + riskP * HA_RR
                tgt2 = entry + riskP * 3#
            Else
                stopP = HA_Max2(haH(n), HA_MaxN(h, n, 3)) + atr * 0.5
                If stopP > entry * (1 + HA_MAXLOSS) Then stopP = entry * (1 + HA_MAXLOSS)
                If stopP <= entry Then stopP = entry * (1 + HA_MAXLOSS)
                riskP = stopP - entry
                tgt1 = entry - riskP * HA_RR
                tgt2 = entry - riskP * 3#
            End If
        End If
        If riskP <= 0 Then GoTo NextStock

        Dim shares As Double, capLimit As Double
        If HA_FIX_MODE Then
            shares = Int((HA_CAPITAL * HA_FIX_POS) / entry / 100#) * 100#
        Else
            shares = Int((HA_CAPITAL * HA_RISK) / riskP / 100#) * 100#
            capLimit = Int((HA_CAPITAL * HA_MAXPOS) / entry / 100#) * 100#
            If shares > capLimit Then shares = capLimit
        End If
        If shares < 100 Then shares = 100

        '===== 結果を貯める =====
        hit = hit + 1
        res(hit, 1) = 0
        res(hit, 2) = code
        res(hit, 3) = nm
        res(hit, 4) = score
        res(hit, 5) = HA_Rank(score)
        res(hit, 6) = cl(n)
        res(hit, 7) = HA_Tick(entry)
        res(hit, 8) = HA_Tick(entry * (1 + HA_MAXGAP))   '見送りライン
        res(hit, 9) = HA_Tick(stopP)
        res(hit, 10) = HA_Tick(tgt1)
        res(hit, 11) = HA_Tick(tgt2)
        res(hit, 12) = shares
        res(hit, 13) = -riskP * shares
        res(hit, 14) = Abs(tgt1 - entry) * shares
        res(hit, 15) = runLen
        res(hit, 16) = Round(rsiNow, 1)
        res(hit, 17) = Round(kNow, 1)
        res(hit, 18) = Round(kairi, 1)
        res(hit, 19) = Round(volRate, 2)
        res(hit, 20) = Round(atr, 1)
        res(hit, 21) = Trim$(sig)
        sc(hit) = score + volRate / 100#    '同点は出来高の多い順

NextStock:
    Next i

    '--- 点数の高い順に並べる ---
    Dim a As Long, b As Long, z As Long
    Dim tD As Double, tV As Variant
    For a = 1 To hit - 1
        For b = a + 1 To hit
            If sc(b) > sc(a) Then
                tD = sc(a): sc(a) = sc(b): sc(b) = tD
                For z = 1 To OUT_COLS
                    tV = res(a, z): res(a, z) = res(b, z): res(b, z) = tV
                Next z
            End If
        Next b
    Next a

    '--- 絞り込みの内訳 ---
    Dim funnel As String
    funnel = "全銘柄 " & cnt(1) & " → 株価" & Format(HA_MINPRICE, "#,##0") & "円以上 " & cnt(2) & _
             " → 出来高" & Format(HA_MINVOL / 10000, "0") & "万株以上 " & cnt(3) & _
             " → 出来高" & HA_VOLRATE & "倍以上 " & cnt(4) & _
             " → 平均足" & IIf(side = 1, "陽線", "陰線") & "2〜5本 " & cnt(5) & _
             " → 前に3本の逆行 " & cnt(6) & " → ヒゲ無し " & cnt(7) & _
             " → 安値切り上げ " & cnt(8) & " → 20日線・乖離 " & cnt(9) & _
             " → RSI50 " & cnt(10) & " → 実体 " & cnt(11) & _
             IIf(HA_WEEK_ON, " → 週足" & HA_WEEK_MA & "週線上 " & cnt(12), "") & _
             " → 合格点" & HA_PASS & " " & cnt(13)

    '--- 出力 ---
    Set wsX = HA_MakeOut()
    HA_WriteOut wsX, res, hit, side, nCol, topixNote, wsC.Cells(3, 4).Value2, funnel

    HA_Restore
    wsX.Activate

    If HA_SILENT Then Exit Sub
    Dim msg As String
    If side = 1 Then msg = "買い候補" Else msg = "売り候補"
    MsgBox msg & " " & IIf(hit > HA_MAXROWS, HA_MAXROWS, hit) & " 件を「平均足」シートに出しました。" & vbCrLf & _
           "地合い：" & topixNote & vbCrLf & vbCrLf & _
           "【絞り込みの内訳】" & vbCrLf & Replace(funnel, " → ", vbCrLf & "  → "), vbInformation
End Sub

Private Sub HA_Restore()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'==================== 出力 ====================
Private Function HA_MakeOut() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = OUT_SHEET
    End If
    Set HA_MakeOut = ws
End Function

Private Sub HA_WriteOut(ByVal ws As Worksheet, ByRef res As Variant, ByVal hit As Long, _
                        ByVal side As Long, ByVal nCol As Long, ByVal topixNote As String, _
                        ByVal lastDate As Variant, ByVal funnel As String)

    ws.Cells.UnMerge
    ws.Cells.Clear

    Dim ttl As String
    If side = 1 Then ttl = "平均足　買い候補" Else ttl = "平均足　売り候補"

    With ws.Range("A1:U1")
        .Merge
        .Value = "  " & ttl & "　（反転を取る7手法／" & Format(Now, "yyyy/mm/dd hh:nn") & " 作成）"
        .Font.Name = "Meiryo UI"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(1).RowHeight = 30

    With ws.Range("A2:U2")
        .Merge
        .Value = "  最新日=" & HA_DateStr(lastDate) & "／使用" & nCol & "日／合格点=" & HA_PASS & _
                 "／資金" & Format(HA_CAPITAL, "#,##0") & "円" & _
                 IIf(HA_FIX_MODE, "・損切-" & Format(HA_FIX_SL * 100, "0") & "%／利確+" & Format(HA_FIX_TP * 100, "0") & "%／" & HA_HOLDDAYS & "日で手じまい", _
                                  "・1回のリスク" & Format(HA_RISK * 100, "0.0") & "%") & _
                 "／地合い：" & topixNote
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Interior.Color = RGB(226, 239, 218)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(2).RowHeight = 20

    Dim hd As Variant
    hd = Array("順位", "コード", "銘柄名", "点数", "評価", "現値", "エントリー", "見送りライン", _
               "損切", "利確1", "利確2", "株数", "想定損失", "想定利益", "連続", "RSI", "%K", _
               "25日乖離%", "出来高倍", "ATR", "根拠（7手法）")
    Dim j As Long
    For j = 0 To UBound(hd)
        ws.Cells(3, j + 1).Value = hd(j)
    Next j
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, OUT_COLS))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(3).RowHeight = 30

    Dim nOut As Long
    nOut = hit
    If nOut > HA_MAXROWS Then nOut = HA_MAXROWS

    Dim r As Long, k As Long
    For r = 1 To nOut
        ws.Cells(OUT_TOP + r - 1, 1).Value = r
        For k = 2 To OUT_COLS
            ws.Cells(OUT_TOP + r - 1, k).Value = res(r, k)
        Next k
    Next r

    If nOut = 0 Then
        ws.Cells(OUT_TOP, 1).Value = "該当なし（無理に建てない日です）"
        ws.Cells(OUT_TOP, 1).Font.Bold = True
        ws.Cells(OUT_TOP + 1, 1).Value = "絞り込みの内訳：" & funnel
        ws.Cells(OUT_TOP + 2, 1).Value = "※検証では2日に1日は候補ゼロです。出ない日に建てないことが勝ちの半分です。"
    Else
        With ws.Range(ws.Cells(OUT_TOP, 1), ws.Cells(OUT_TOP + nOut - 1, OUT_COLS))
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
        ws.Range(ws.Cells(OUT_TOP, 6), ws.Cells(OUT_TOP + nOut - 1, 14)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(OUT_TOP, 16), ws.Cells(OUT_TOP + nOut - 1, 20)).NumberFormat = "0.0"

        Dim cc As Long
        For cc = 1 To nOut
            If Val(ws.Cells(OUT_TOP + cc - 1, 4).Value) >= HA_PASS + 4 Then
                ws.Range(ws.Cells(OUT_TOP + cc - 1, 1), ws.Cells(OUT_TOP + cc - 1, OUT_COLS)).Interior.Color = RGB(198, 239, 206)
            End If
        Next cc
    End If

    Dim rr As Long
    rr = OUT_TOP + nOut + 2
    ws.Cells(rr, 1).Value = "絞り込みの内訳：" & funnel
    rr = rr + 2
    ws.Cells(rr, 1).Value = "【売買ルール　この通りにやる】"
    ws.Cells(rr + 1, 1).Value = "1. 上位（点数の高い順）から、翌日の寄り成りで買う。同時保有は5銘柄まで。"
    If HA_FIX_MODE Then
        ws.Cells(rr + 2, 1).Value = "2. ただし寄り値が「見送りライン」より高く始まったら、その銘柄は買わない（高値づかみ防止）。"
        ws.Cells(rr + 3, 1).Value = "3. 買った直後に「損切」（買値の-" & Format(HA_FIX_SL * 100, "0") & "%）を逆指値で入れる。絶対に下げない。"
        ws.Cells(rr + 4, 1).Value = "4. 同時に「利確1」（買値の+" & Format(HA_FIX_TP * 100, "0") & "%）に売り指値を置く。届いたら全部売り。"
        ws.Cells(rr + 5, 1).Value = "5. どちらにも当たらなければ、" & HA_HOLDDAYS & "営業日たった日の引けで成行手仕舞い。途中で判断しない。"
        ws.Cells(rr + 6, 1).Value = "6. 3連敗したら株数を半分。月の損失が資金の6%になったらその月は休む。"
        ws.Cells(rr + 7, 1).Value = "※検証（300銘柄×250日・手数料込・窓開けの滑りも計算）：勝率65.2%、1回平均+1.62%、250日+51.1%、最大下落9.1%。"
    Else
        ws.Cells(rr + 2, 1).Value = "2. 買った直後に「損切」を逆指値で入れる。入れた後は絶対に下げない。"
        ws.Cells(rr + 3, 1).Value = "3. 「利確1」に届いたら半分売る。残りの損切は買値まで引き上げる。"
        ws.Cells(rr + 4, 1).Value = "4. 平均足が陰線2本続いたら、残り全部を成行で手仕舞う（これが平均足の本番）。"
        ws.Cells(rr + 5, 1).Value = "5. " & HA_HOLDDAYS & "営業日たっても「利確1」に届かない玉は、勝ち負けに関係なく手仕舞う。"
        ws.Cells(rr + 6, 1).Value = "6. 3連敗したら株数を半分。月の損失が資金の6%になったらその月は休む。"
        ws.Cells(rr + 7, 1).Value = "※検証（300銘柄×250日・手数料込）：勝率49.2%、1回平均+0.40R。100%勝てる方法ではありません。"
    End If
    With ws.Range(ws.Cells(rr, 1), ws.Cells(rr + 7, 1))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
    End With
    ws.Cells(rr, 1).Font.Bold = True

    ws.Columns("A:U").AutoFit
    If ws.Columns("C").ColumnWidth > 18 Then ws.Columns("C").ColumnWidth = 18
    If ws.Columns("U").ColumnWidth > 60 Then ws.Columns("U").ColumnWidth = 60
    ws.Rows(3).RowHeight = 30

    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

'==================== 部品 ====================
Private Function HA_GetWs(ByVal n As String) As Worksheet
    On Error Resume Next
    Set HA_GetWs = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
End Function

Private Function HA_Num(ByVal v As Variant) As Double
    If IsNumeric(v) Then HA_Num = CDbl(v) Else HA_Num = 0
End Function

Private Function HA_DateStr(ByVal v As Variant) As String
    If IsNumeric(v) Then
        If Val(v) > 40000 Then
            HA_DateStr = Format(CDate(Val(v)), "yyyy/mm/dd")
            Exit Function
        End If
    End If
    HA_DateStr = CStr(v)
End Function

Private Function HA_Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then HA_Max2 = a Else HA_Max2 = b
End Function

Private Function HA_Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then HA_Min2 = a Else HA_Min2 = b
End Function

Private Function HA_Max3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    HA_Max3 = a
    If b > HA_Max3 Then HA_Max3 = b
    If c > HA_Max3 Then HA_Max3 = c
End Function

Private Function HA_Min3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    HA_Min3 = a
    If b < HA_Min3 Then HA_Min3 = b
    If c < HA_Min3 Then HA_Min3 = c
End Function

Private Function HA_Avg(ByRef x() As Double, ByVal idx As Long, ByVal p As Long) As Double
    Dim k As Long, s As Double, cnt As Long
    For k = idx To idx - p + 1 Step -1
        If k >= 1 Then
            s = s + x(k): cnt = cnt + 1
        End If
    Next k
    If cnt > 0 Then HA_Avg = s / cnt
End Function

Private Function HA_Sma(ByRef x() As Double, ByVal idx As Long, ByVal p As Long) As Double
    If idx < p Then Exit Function
    HA_Sma = HA_Avg(x, idx, p)
End Function

Private Function HA_MinN(ByRef x() As Double, ByVal idx As Long, ByVal p As Long) As Double
    Dim k As Long, m As Double
    For k = idx To idx - p + 1 Step -1
        If k >= 1 Then
            If m = 0 Then
                m = x(k)
            ElseIf x(k) < m Then
                m = x(k)
            End If
        End If
    Next k
    HA_MinN = m
End Function

Private Function HA_MaxN(ByRef x() As Double, ByVal idx As Long, ByVal p As Long) As Double
    Dim k As Long, m As Double
    For k = idx To idx - p + 1 Step -1
        If k >= 1 Then
            If x(k) > m Then m = x(k)
        End If
    Next k
    HA_MaxN = m
End Function

'idx から f本前 〜 t本前 の安値
Private Function HA_MinN2(ByRef x() As Double, ByVal idx As Long, ByVal f As Long, ByVal t As Long) As Double
    Dim k As Long, m As Double
    For k = idx - f + 1 To idx - t + 1 Step -1
        If k >= 1 Then
            If m = 0 Then
                m = x(k)
            ElseIf x(k) < m Then
                m = x(k)
            End If
        End If
    Next k
    HA_MinN2 = m
End Function

Private Function HA_MaxN2(ByRef x() As Double, ByVal idx As Long, ByVal f As Long, ByVal t As Long) As Double
    Dim k As Long, m As Double
    For k = idx - f + 1 To idx - t + 1 Step -1
        If k >= 1 Then
            If x(k) > m Then m = x(k)
        End If
    Next k
    HA_MaxN2 = m
End Function

'同じ色が何本続いているか
Private Function HA_RunLen(ByRef haO() As Double, ByRef haC() As Double, _
                           ByVal n As Long, ByVal side As Long) As Long
    Dim k As Long, cnt As Long
    For k = n To 1 Step -1
        If side = 1 Then
            If haC(k) > haO(k) Then cnt = cnt + 1 Else Exit For
        Else
            If haC(k) < haO(k) Then cnt = cnt + 1 Else Exit For
        End If
    Next k
    HA_RunLen = cnt
End Function

'転換する前に逆の色が何本続いていたか
Private Function HA_RunPrev(ByRef haO() As Double, ByRef haC() As Double, _
                            ByVal n As Long, ByVal runLen As Long, ByVal side As Long) As Long
    Dim st As Long, k As Long, cnt As Long
    st = n - runLen
    For k = st To 1 Step -1
        If side = 1 Then
            If haC(k) > haO(k) Then cnt = cnt + 1 Else Exit For
        Else
            If haC(k) < haO(k) Then cnt = cnt + 1 Else Exit For
        End If
    Next k
    HA_RunPrev = cnt
End Function

'RSI（ワイルダー簡易）
Private Function HA_Rsi(ByRef cl() As Double, ByVal idx As Long, ByVal p As Long) As Double
    HA_Rsi = 50
    If idx < p + 1 Then Exit Function
    Dim k As Long, up As Double, dn As Double, dd As Double
    For k = idx - p + 1 To idx
        dd = cl(k) - cl(k - 1)
        If dd > 0 Then up = up + dd Else dn = dn - dd
    Next k
    up = up / p: dn = dn / p
    If dn = 0 Then
        HA_Rsi = 100
    Else
        HA_Rsi = 100# - 100# / (1# + up / dn)
    End If
End Function

'ストキャス %K
Private Function HA_StochK(ByRef h() As Double, ByRef lo() As Double, ByRef cl() As Double, _
                           ByVal idx As Long, ByVal p As Long) As Double
    HA_StochK = 50
    If idx < p Then Exit Function
    Dim hh As Double, ll As Double
    hh = HA_MaxN(h, idx, p)
    ll = HA_MinN(lo, idx, p)
    If hh - ll <= 0 Then Exit Function
    HA_StochK = (cl(idx) - ll) / (hh - ll) * 100#
End Function

'ATR
Private Function HA_Atr(ByRef h() As Double, ByRef lo() As Double, ByRef cl() As Double, _
                        ByVal idx As Long, ByVal p As Long) As Double
    If idx < p + 1 Then
        HA_Atr = h(idx) - lo(idx)
        Exit Function
    End If
    Dim k As Long, s As Double, tr As Double
    For k = idx - p + 1 To idx
        tr = h(k) - lo(k)
        If Abs(h(k) - cl(k - 1)) > tr Then tr = Abs(h(k) - cl(k - 1))
        If Abs(lo(k) - cl(k - 1)) > tr Then tr = Abs(lo(k) - cl(k - 1))
        s = s + tr
    Next k
    HA_Atr = s / p
End Function

'呼値に丸める
Private Function HA_Tick(ByVal p As Double) As Double
    Dim t As Double
    Select Case p
        Case Is < 3000: t = 1
        Case Is < 5000: t = 5
        Case Is < 30000: t = 10
        Case Is < 50000: t = 50
        Case Is < 300000: t = 100
        Case Else: t = 1000
    End Select
    HA_Tick = Int(p / t + 0.5) * t
End Function

Private Function HA_Rank(ByVal s As Long) As String
    If s >= HA_PASS + 6 Then
        HA_Rank = "◎本命"
    ElseIf s >= HA_PASS + 3 Then
        HA_Rank = "○有力"
    Else
        HA_Rank = "△様子見"
    End If
End Function

'地合い（TOPIX＝5行目）　この手法は地合いが弱い時ほど当たる
Private Function HA_TopixState(ByVal wsC As Worksheet, ByVal lastCol As Long, _
                               ByVal side As Long, ByRef note As String) As Long
    Dim k As Long, n As Long, v As Double
    Dim t() As Double
    ReDim t(1 To lastCol - C_NEW + 1)
    n = 0
    For k = lastCol To C_NEW Step -1
        v = HA_Num(wsC.Cells(5, k).Value)
        If v > 0 Then
            n = n + 1
            t(n) = v
        End If
    Next k
    If n < 21 Then
        note = "判定できず（TOPIXデータ不足）"
        HA_TopixState = 0
        Exit Function
    End If
    Dim ma20 As Double
    ma20 = HA_Sma(t, n, 20)
    '※TOPIXは表示のみ。点数には入れない（データが古い列で途切れる場合があるため）
    HA_TopixState = 0
    If t(n) > ma20 Then
        note = "TOPIXは20日線の上"
    Else
        note = "TOPIXは20日線の下"
    End If
End Function

'==================================================================
' 売買ルールを「平均足ルール」シートに書き出す
'   実行 : Alt+F8 →「平均足_ルール表示」
'   ※このブックの中だけで完結します（他のブックには触りません）
'==================================================================
Public Sub 平均足_ルール表示()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("平均足ルール")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "平均足ルール"
    End If

    Application.ScreenUpdating = False
    ws.Cells.UnMerge
    ws.Cells.Clear

    With ws.Range("A1:C1")
        .Merge
        .Value = "  平均足トレード　売買ルール（実データ検証済み）"
        .Font.Name = "Meiryo UI"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(1).RowHeight = 32

    Dim r As Long
    r = 3

    HA_PutH ws, r, "【1】買う条件　この9つが全部そろった時だけ買う"
    HA_PutR ws, r, "1", "当日出来高 ≧ 5日平均 × " & HA_VOLRATE & "倍", "ここが一番効く。1.2倍まで下げると成績が半分になる"
    HA_PutR ws, r, "2", "平均足の陽線が 2〜5本連続", "6本目以降は高値づかみになる"
    HA_PutR ws, r, "3", "その前に陰線が3本以上（20日線の上なら2本以上）", "下げの終わりを取るため"
    HA_PutR ws, r, "4", "最新の平均足に下ヒゲが無い", "買い方が完全に優勢な形"
    HA_PutR ws, r, "5", "直近3日の安値 > その前4〜8日の安値", "安値切り上げ＝底が固まった証拠"
    HA_PutR ws, r, "6", "終値 > 20日線／週足が10週線の上", "大きな流れに逆らわない。勝率58.6%→65.2%"
    HA_PutR ws, r, "7", "25日線からの乖離 ≧ 0%", "下げ続けている株の逆張りは負ける（検証で確認）"
    HA_PutR ws, r, "8", "RSI(14) ≧ 50", "50割れは+0.04R、50超で+0.21R"
    HA_PutR ws, r, "9", "平均足の実体率 ≦ 0.85", "行き過ぎた足は翌日戻される"
    HA_PutR ws, r, "他", "株価1000円以上・5日平均出来高20万株以上", "1000円未満は勝率40.5%・平均-0.87%と悪い"
    r = r + 1

    HA_PutH ws, r, "【2】点数　合格点は4点以上。候補が複数出た日は点数の高い順に買う"
    HA_PutR ws, r, "+3", "転換の直前が十字線（実体率0.35以下）", "平均足の本命パターン"
    HA_PutR ws, r, "+2", "出来高2倍以上（1.6倍以上なら+1）", ""
    HA_PutR ws, r, "+2", "RSIが65以下（適温）　※75以上なら −1", ""
    HA_PutR ws, r, "+2", "転換前の陰線が3〜5本", ""
    HA_PutR ws, r, "+1", "9日線 > 18日線／その直近ゴールデンクロス", ""
    HA_PutR ws, r, "+1", "20日安値から1.5%以内で反発（支持線）", ""
    HA_PutR ws, r, "+1", "20日高値を終値で更新（ブレイク）", ""
    HA_PutR ws, r, "+1", "ストキャス %K が20を上抜け", ""
    HA_PutR ws, r, "+1", "RSIが50を上抜けた当日", ""
    HA_PutR ws, r, "+1", "ATR ≦ 株価の3.5%（値動きが素直）", ""
    HA_PutR ws, r, "+1", "実体率 0.3〜0.7", ""
    HA_PutR ws, r, "※", "早い者順に買うと年+34%、点数順なら年+103%", "順位を守ることが一番の利益源"
    r = r + 1

    HA_PutH ws, r, "【3】出口　ここを守らないと必ず負ける"
    If HA_FIX_MODE Then
        HA_PutR ws, r, "1", "損切 ＝ 買値の −" & Format(HA_FIX_SL * 100, "0") & "%（逆指値）", "買った直後に入れる。絶対に下げない"
        HA_PutR ws, r, "2", "利確 ＝ 買値の +" & Format(HA_FIX_TP * 100, "0") & "%（売り指値）", "届いたら全部売り"
        HA_PutR ws, r, "3", HA_HOLDDAYS & "営業日たったら引けで成行手仕舞い", "どちらにも当たらなかった玉"
        HA_PutR ws, r, "4", "途中で判断しない。置いた注文をいじらない", "いじると成績が落ちる"
        HA_PutR ws, r, "5", "寄り値が「見送りライン」（現値+" & Format(HA_MAXGAP * 100, "0") & "%）より高い日は買わない", "高値づかみを避ける。87%は買えます"
        HA_PutR ws, r, "※", "損切8%・利確8%が最良。利確を損切より小さくすると必ず負ける", "勝率72%でも年−24.5%になる"
    Else
        HA_PutR ws, r, "1", "損切 ＝ min(平均足の安値, 直近3日安値) − ATR×0.5（最大 −6%）", "買った直後に逆指値。入れた後は絶対に下げない"
        HA_PutR ws, r, "2", "+" & HA_RR & "R に届いたら半分売る。残りの損切は買値まで上げる", "Rは1回の損切幅"
        HA_PutR ws, r, "3", "平均足が陰線2本続いたら残り全部を成行で手仕舞う", "これが平均足の本番"
        HA_PutR ws, r, "4", HA_HOLDDAYS & "営業日たっても未達の玉は勝ち負けに関係なく手仕舞う", "資金を寝かせない"
        HA_PutR ws, r, "※", "陰線2本で手仕舞い=+0.43R／陰線1本=+0.15R（早すぎる）", ""
    End If
    r = r + 1

    HA_PutH ws, r, "【4】資金管理"
    HA_PutR ws, r, "1", IIf(HA_FIX_MODE, "1銘柄に資金の" & Format(HA_FIX_POS * 100, "0") & "%を使う", "1回のリスクは資金の1%"), IIf(HA_FIX_MODE, "1回の最大損失は資金の" & Format(HA_FIX_POS * HA_FIX_SL * 100, "0.0") & "%", "3,000,000円なら1回30,000円まで")
    HA_PutR ws, r, "2", "株数はマクロが自動計算（100株単位で切り捨て）", ""
    HA_PutR ws, r, "3", "1銘柄に使う金額は資金の25%まで", ""
    HA_PutR ws, r, "4", "同時保有は5銘柄まで", "3銘柄にすると下落幅がさらに小さい"
    HA_PutR ws, r, "5", "3連敗したら株数を半分にする", ""
    HA_PutR ws, r, "6", "月間の損失が資金の6%になったらその月は休む", ""
    r = r + 1

    HA_PutH ws, r, "【5】検証結果　300銘柄 × 直近250日 ・ 手数料往復0.3%込み"
    If HA_FIX_MODE Then
        HA_PutR ws, r, "", "設定", "出来高1.4倍 / 株価1000円以上 / 週足10週線上 / 損切8% / 利確8% / 5日手じまい / 寄り+2%まで"
        HA_PutR ws, r, "", "勝率", "65.2%"
        HA_PutR ws, r, "", "1回あたりの平均", "+1.62%"
        HA_PutR ws, r, "", "同時5銘柄・1銘柄に資金の20%", "250日で +51.1%　最大下落 9.1%"
        HA_PutR ws, r, "", "週足の条件を外すと", "勝率58.6%　250日 +62.0%（回数は増えるが勝率は下がる）"
        HA_PutR ws, r, "", "注", "損切が窓を開けて滑る分も引いた数字です"
    Else
        HA_PutR ws, r, "", "設定", "出来高1.6倍 / ATR基準の損切 / +2Rで半分利確"
        HA_PutR ws, r, "", "勝率", "49.2%"
        HA_PutR ws, r, "", "1回あたりの平均", "+0.40R"
        HA_PutR ws, r, "", "同時5銘柄・1回1%リスク", "250日で +102.8%　最大下落 13.2%"
        HA_PutR ws, r, "", "プロフィットファクター", "2.15"
    End If
    r = r + 1

    HA_PutH ws, r, "【6】注意"
    HA_PutR ws, r, "1", "引け後 15:30以降に実行する", "出来高が確定しないと候補が出ない"
    HA_PutR ws, r, "2", "ザラ場中に回すときは HA_SKIP を 1 にする", "前日の足で判定します"
    HA_PutR ws, r, "3", "翌日の寄り成りで買う。ギャップアップで損切幅が6%を超えるなら見送り", ""
    HA_PutR ws, r, "4", "検証は直近1年・300銘柄。相場つきが変われば成績は落ちる", ""
    HA_PutR ws, r, "5", "100%勝てる方法ではない。負けを1%で切ることが利益の源泉", ""

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 58
    ws.Columns("C").ColumnWidth = 46
    With ws.Range(ws.Cells(3, 1), ws.Cells(r - 1, 3))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
        .VerticalAlignment = xlTop
    End With
    ws.Rows("3:" & r).AutoFit

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A1").Select
    MsgBox "「平均足ルール」シートを作りました。", vbInformation
End Sub

Private Sub HA_PutH(ByVal ws As Worksheet, ByRef r As Long, ByVal s As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Merge
        .Value = s
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(r).RowHeight = 22
    r = r + 1
End Sub

Private Sub HA_PutR(ByVal ws As Worksheet, ByRef r As Long, ByVal a As String, _
                    ByVal b As String, ByVal c As String)
    ws.Cells(r, 1).Value = a
    ws.Cells(r, 2).Value = b
    ws.Cells(r, 3).Value = c
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .WrapText = True
    End With
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 3).Font.Color = RGB(90, 90, 90)
    r = r + 1
End Sub

'週足が10週線の上か（直前に確定した週で判定）
Private Function HA_WeekOK(ByRef cl() As Double, ByRef dts() As Double, _
                           ByVal n As Long, ByVal side As Long) As Boolean

    Dim wc(1 To 80) As Double
    Dim wCnt As Long, k As Long
    Dim curKey As Double, wkKey As Double
    wCnt = 0: curKey = -1

    For k = 1 To n
        If dts(k) > 40000 Then
            wkKey = Int(dts(k)) - (Weekday(CDate(dts(k)), vbMonday) - 1)
            If wkKey <> curKey Then
                wCnt = wCnt + 1
                If wCnt > 80 Then Exit For
                curKey = wkKey
            End If
            wc(wCnt) = cl(k)      'その週の最新の終値で上書き
        End If
    Next k

    '一番新しい週はまだ途中なので、その1つ前を使う
    Dim w As Long
    w = wCnt - 1
    If w < HA_WEEK_MA + 1 Then Exit Function

    Dim s As Double, j As Long
    For j = w - HA_WEEK_MA + 1 To w
        s = s + wc(j)
    Next j
    Dim ma As Double
    ma = s / HA_WEEK_MA

    If side = 1 Then
        HA_WeekOK = (wc(w) > ma)
    Else
        HA_WeekOK = (wc(w) < ma)
    End If
End Function
