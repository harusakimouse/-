Attribute VB_Name = "Mod_HeikinAshi"
Option Explicit
'==================================================================
' 平均足　買い候補　抽出モジュール　v3.0（ルール一本化版）
'
'   ◆売買ルール（これ1つだけ。旧ATR方式と点数制は廃止）
'     1. 平均足が陽線2〜5本連続（その前に陰線3本以上）
'     2. 最新足に下ヒゲなし／安値切り上げ／終値が20日線の上
'     3. 当日出来高 ≧ 5日平均 × HA_VOLRATE
'     4. 週足の終値が10週線の上
'     5. 株価 HA_MINPRICE 以上／5日平均出来高 HA_MINVOL 以上
'     6. 翌営業日の寄りで買う。ただし寄りが現値+HA_MAXGAP%超なら見送り
'     7. 損切 -HA_SL%（逆指値）／利確 +HA_TP%（売り指値）
'     8. どちらにも当たらなければ、買った日から HA_HOLDDAYS 営業日後の引けで成行
'     ※順位は「出来高倍率の高い順」。点数は使いません（検証で逆効果と判明）。
'
'   ◆検証結果（300銘柄×250日・手数料往復0.3%込み・窓開けの滑りも計算）
'     A 安全重視(1.6倍/2000円/実体)  54件 勝率68.5% 平均+2.34% PF2.21 後半PF1.48 最大4連敗
'     B 標準   (1.5倍/2000円)      115件 勝率66.1% 平均+1.87% PF1.85 後半PF1.30 最大6連敗 ←採用
'     C 利益重視(1.4倍/1000円)      221件 勝率63.8% 平均+1.48% PF1.67 後半PF1.02 最大17連敗
'     ※採用基準（全期間PF1.3以上／後半PF1.2以上／取引100回以上）を満たすのはBだけ。
'     ※ただし前半PF3.59→後半PF1.30と半減、直近3か月はPF1.03。判定は「少額でテスト運用」。
'     ※資金300万・1銘柄15%(45万)で回した実績：65件 勝率66.2% PF2.02 最大DD2.7% 年+13.8%
'       1銘柄10%(30万)だと株価3000円以下しか買えず、候補の75%を取り逃がします。
'     ※点数制度は検証の結果、点数が高いほど成績が悪かったため廃止しました
'       （4〜6点 PF2.77／7〜8点 PF1.57／9点以上 PF0.89）。
'==================================================================

'============ 設定（A/B/Cはここを変えるだけ）============
'  A 安全重視 : VOLRATE=1.6  MINPRICE=2000  BODY_ON=True   （PF2.21だが54件と少ない）
'  B 標準     : VOLRATE=1.5  MINPRICE=2000  BODY_ON=False  ← 採用（唯一 基準を全部満たした）
'  C 利益重視 : VOLRATE=1.4  MINPRICE=1000  BODY_ON=False  （後半PF1.02・直近3か月-1.67%）
Public Const HA_VOLRATE  As Double = 1.5       '当日出来高 ÷ 5日平均
Public Const HA_MINPRICE As Double = 2000#     '最低株価
Public Const HA_BODY_ON  As Boolean = False    '平均足の実体率で絞るか
Public Const HA_MAXGAP   As Double = 0.02      '寄りがこれ以上高かったら見送る
'============================================================

Public Const HA_CAPITAL  As Double = 3000000#  '運用資金（円）
Public Const HA_POS      As Double = 0.15      '1銘柄に使う資金の割合（資金300万なら45万＝株価4500円まで買える）
Public Const HA_SLOTS    As Long = 5           '同時に持てる銘柄数
Public Const HA_SL       As Double = 0.08      '損切（8%）
Public Const HA_TP       As Double = 0.08      '利確（8%）
Public Const HA_HOLDDAYS As Long = 5           '買った日から何営業日後の引けで手じまうか
Public Const HA_MINVOL   As Double = 200000#   '5日平均出来高の下限（株）
Public Const HA_WEEK_ON  As Boolean = True     '週足フィルター
Public Const HA_WEEK_MA  As Long = 10          '週足の移動平均（週）
Public Const HA_BODY_MIN As Double = 0.3       '実体率の下限
Public Const HA_BODY_MAX As Double = 0.7       '実体率の上限
Public Const HA_MAXROWS  As Long = 60          '出力する最大件数
Public Const HA_SKIP     As Long = 0           '0=最新日で判定　1=前日で判定

Public HA_SILENT As Boolean                    'True=自動実行中（メッセージを出さない）

Private Const R_TOP As Long = 6
Private Const R_END As Long = 520
Private Const C_NEW As Long = 5
Private Const N_MAX As Long = 150
Private Const OUT_SHEET As String = "平均足"
Private Const OUT_TOP As Long = 4
Private Const OUT_COLS As Long = 20

'==================== 入口 ====================
Public Sub 平均足_買い抽出()
    HA_Run
End Sub

'==================== 本体 ====================
Private Sub HA_Run()

    Dim wsO As Worksheet, wsH As Worksheet, wsL As Worksheet
    Dim wsC As Worksheet, wsV As Worksheet, wsX As Worksheet

    Set wsO = HA_GetWs("始値"): Set wsH = HA_GetWs("高値")
    Set wsL = HA_GetWs("安値"): Set wsC = HA_GetWs("終値")
    Set wsV = HA_GetWs("出来高")

    If wsO Is Nothing Or wsH Is Nothing Or wsL Is Nothing Or wsC Is Nothing Or wsV Is Nothing Then
        If Not HA_SILENT Then MsgBox "始値・高値・安値・終値・出来高の5シートが必要です。", vbExclamation
        Exit Sub
    End If

    '--- データが古くないか（楽天RSSの切断・取込忘れを検出）---
    Dim fmsg As String
    If Not HA_DataFresh(fmsg) Then
        If Not HA_SILENT Then
            If MsgBox("★データが古いです★" & vbCrLf & vbCrLf & fmsg & vbCrLf & vbCrLf & _
                      "楽天RSSが切れているか、取込を忘れている可能性があります。" & vbCrLf & _
                      "この状態の候補で売買しないでください。" & vbCrLf & vbCrLf & _
                      "それでも候補を出しますか？", vbYesNo + vbExclamation, "データが古い") <> vbYes Then Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- 使える列（日付は .Value2 で読む）---
    Dim lastCol As Long, c As Long, dv As Double, miss As Long
    lastCol = C_NEW: miss = 0
    For c = C_NEW + 1 To C_NEW + N_MAX - 1
        dv = 0
        If IsNumeric(wsC.Cells(3, c).Value2) Then dv = CDbl(wsC.Cells(3, c).Value2)
        If dv > 40000 And dv < 80000 Then
            lastCol = c: miss = 0
        Else
            miss = miss + 1
            If miss > 5 Then Exit For
        End If
    Next c
    Dim nCol As Long
    nCol = lastCol - C_NEW + 1
    If nCol < 60 Then
        HA_Restore
        If Not HA_SILENT Then MsgBox "履歴が足りません（" & nCol & "日）。週足の判定に60日以上必要です。", vbExclamation
        Exit Sub
    End If

    '--- 日付（E列の日付はD3にある）---
    Dim dtc() As Double, kk As Long
    ReDim dtc(1 To nCol)
    For kk = 1 To nCol
        If kk = 1 Then
            dtc(kk) = HA_Num(wsC.Cells(3, 4).Value2)
        Else
            dtc(kk) = HA_Num(wsC.Cells(3, 4 + kk).Value2)
        End If
    Next kk

    Dim lastRow As Long
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    If lastRow > R_END Then lastRow = R_END
    If lastRow < R_TOP Then
        HA_Restore
        If Not HA_SILENT Then MsgBox "銘柄データがありません。", vbExclamation
        Exit Sub
    End If

    Dim aO As Variant, aH As Variant, aL As Variant, aC As Variant, aV As Variant
    Dim aCode As Variant, aName As Variant
    aO = wsO.Range(wsO.Cells(R_TOP, C_NEW), wsO.Cells(lastRow, lastCol)).Value
    aH = wsH.Range(wsH.Cells(R_TOP, C_NEW), wsH.Cells(lastRow, lastCol)).Value
    aL = wsL.Range(wsL.Cells(R_TOP, C_NEW), wsL.Cells(lastRow, lastCol)).Value
    aC = wsC.Range(wsC.Cells(R_TOP, C_NEW), wsC.Cells(lastRow, lastCol)).Value
    aV = wsV.Range(wsV.Cells(R_TOP, C_NEW), wsV.Cells(lastRow, lastCol)).Value
    aCode = wsC.Range(wsC.Cells(R_TOP, 1), wsC.Cells(lastRow, 1)).Value
    aName = wsC.Range(wsC.Cells(R_TOP, 2), wsC.Cells(lastRow, 2)).Value

    Dim res() As Variant, srt() As Double
    ReDim res(1 To lastRow - R_TOP + 1, 1 To OUT_COLS)
    ReDim srt(1 To lastRow - R_TOP + 1)
    Dim hit As Long: hit = 0
    Dim cnt(1 To 10) As Long

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
                vo(n) = HA_Num(aV(i, k)): dts(n) = dtc(k)
            End If
        Next k
        n = n - HA_SKIP
        If n < 60 Then GoTo NextStock
        If cl(n) < HA_MINPRICE Then GoTo NextStock
        cnt(2) = cnt(2) + 1

        '--- 平均足 ---
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

        '--- (1) 出来高 ---
        Dim vol5 As Double, volRate As Double
        vol5 = HA_Avg(vo, n, 5)
        If vol5 < HA_MINVOL Then GoTo NextStock
        cnt(3) = cnt(3) + 1
        volRate = vo(n) / vol5
        If volRate < HA_VOLRATE Then GoTo NextStock
        cnt(4) = cnt(4) + 1

        '--- (2) 平均足の転換 ---
        Dim runLen As Long
        runLen = HA_RunLen(haO, haC, n)
        If runLen < 2 Or runLen > 5 Then GoTo NextStock
        cnt(5) = cnt(5) + 1

        Dim before As Long, sma20 As Double, sma25 As Double
        before = HA_RunPrev(haO, haC, n, runLen)
        sma20 = HA_Sma(cl, n, 20): sma25 = HA_Sma(cl, n, 25)
        If Not (before >= 3 Or (cl(n) > sma20 And before >= 2)) Then GoTo NextStock

        '--- (3) 下ヒゲなし ---
        If haL(n) < haO(n) - cl(n) * 0.0005 Then GoTo NextStock
        '--- (4) 安値切り上げ ---
        If HA_MinN(lo, n, 3) <= HA_MinN2(lo, n, 4, 8) Then GoTo NextStock
        '--- (5) 20日線の上 ---
        If cl(n) <= sma20 Then GoTo NextStock
        cnt(6) = cnt(6) + 1

        '--- (6) 実体率（A設定の時だけ）---
        Dim bodyRate As Double
        If haH(n) - haL(n) > 0 Then bodyRate = Abs(haC(n) - haO(n)) / (haH(n) - haL(n))
        If HA_BODY_ON Then
            If bodyRate < HA_BODY_MIN Or bodyRate > HA_BODY_MAX Then GoTo NextStock
        End If
        cnt(7) = cnt(7) + 1

        '--- (7) 週足10週線の上 ---
        Dim weekDev As Double
        weekDev = HA_WeekDev(cl, dts, n)
        If HA_WEEK_ON Then
            If weekDev < -998 Then GoTo NextStock
            If weekDev <= 0 Then GoTo NextStock
        End If
        cnt(8) = cnt(8) + 1

        '--- 建玉 ---
        Dim entry As Double, stopP As Double, tgt As Double, guard As Double
        Dim shares As Double
        entry = cl(n)
        stopP = entry * (1 - HA_SL)
        tgt = entry * (1 + HA_TP)
        guard = entry * (1 + HA_MAXGAP)
        shares = Int((HA_CAPITAL * HA_POS) / entry / 100#) * 100#
        If shares < 100 Then shares = 100

        hit = hit + 1
        res(hit, 1) = 0
        res(hit, 2) = code
        res(hit, 3) = nm
        res(hit, 4) = cl(n)
        res(hit, 5) = HA_Tick(entry)
        res(hit, 6) = HA_Tick(guard)
        res(hit, 7) = HA_Tick(stopP)
        res(hit, 8) = HA_Tick(tgt)
        res(hit, 9) = shares
        res(hit, 10) = -(entry - stopP) * shares
        res(hit, 11) = (tgt - entry) * shares
        res(hit, 12) = Round(volRate, 2)
        res(hit, 13) = Round(weekDev, 1)
        res(hit, 14) = runLen
        res(hit, 15) = before
        res(hit, 16) = Round(bodyRate, 2)
        res(hit, 17) = Round(HA_Rsi(cl, n, 14), 1)
        res(hit, 18) = IIf(sma25 > 0, Round((cl(n) - sma25) / sma25 * 100#, 1), 0)
        res(hit, 19) = Round(HA_Atr(h, lo, cl, n, 14), 1)
        res(hit, 20) = "出来高" & Format(volRate, "0.0") & "倍／週足+" & Format(weekDev, "0.0") & _
                       "%／陽線" & runLen & "本／前" & before & "本の下げ／実体" & Format(bodyRate, "0.00")
        srt(hit) = volRate

NextStock:
    Next i

    '--- 出来高倍率の高い順に並べる（点数は使わない）---
    Dim a As Long, b As Long, z As Long
    Dim tD As Double, tV As Variant
    For a = 1 To hit - 1
        For b = a + 1 To hit
            If srt(b) > srt(a) Then
                tD = srt(a): srt(a) = srt(b): srt(b) = tD
                For z = 1 To OUT_COLS
                    tV = res(a, z): res(a, z) = res(b, z): res(b, z) = tV
                Next z
            End If
        Next b
    Next a

    Dim funnel As String
    funnel = "全銘柄 " & cnt(1) & " → 株価" & Format(HA_MINPRICE, "#,##0") & "円以上 " & cnt(2) & _
             " → 出来高" & Format(HA_MINVOL / 10000, "0") & "万株以上 " & cnt(3) & _
             " → 出来高" & HA_VOLRATE & "倍以上 " & cnt(4) & _
             " → 平均足陽線2〜5本 " & cnt(5) & _
             " → 転換・ヒゲ・安値切上・20日線 " & cnt(6) & _
             IIf(HA_BODY_ON, " → 実体" & HA_BODY_MIN & "〜" & HA_BODY_MAX & " " & cnt(7), "") & _
             IIf(HA_WEEK_ON, " → 週足" & HA_WEEK_MA & "週線上 " & cnt(8), "")

    Set wsX = HA_MakeOut()
    HA_WriteOut wsX, res, hit, nCol, wsC.Cells(3, 4).Value2, funnel

    HA_Restore
    On Error Resume Next
    wsX.Activate
    On Error GoTo 0

    If HA_SILENT Then Exit Sub
    MsgBox "買い候補 " & IIf(hit > HA_MAXROWS, HA_MAXROWS, hit) & " 件を「平均足」シートに出しました。" & vbCrLf & vbCrLf & _
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
                        ByVal nCol As Long, ByVal lastDate As Variant, ByVal funnel As String)

    ws.Cells.UnMerge
    ws.Cells.Clear

    With ws.Range("A1:T1")
        .Merge
        .Value = "  平均足　買い候補　（" & Format(Now, "yyyy/mm/dd hh:nn") & " 作成）"
        .Font.Name = "Meiryo UI": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(1).RowHeight = 30

    With ws.Range("A2:T2")
        .Merge
        .Value = "  最新日=" & HA_DateStr(lastDate) & "／使用" & nCol & "日" & _
                 "／出来高" & HA_VOLRATE & "倍以上" & _
                 "／週足" & IIf(HA_WEEK_ON, HA_WEEK_MA & "週線の上だけ", "見ない") & _
                 "／実体" & IIf(HA_BODY_ON, HA_BODY_MIN & "〜" & HA_BODY_MAX, "制限なし") & _
                 "／損切-" & Format(HA_SL * 100, "0") & "%／利確+" & Format(HA_TP * 100, "0") & "%" & _
                 "／" & HA_HOLDDAYS & "営業日で手じまい" & _
                 "／1銘柄=資金の" & Format(HA_POS * 100, "0") & "%（" & Format(HA_CAPITAL * HA_POS, "#,##0") & "円）"
        .Font.Name = "Meiryo UI": .Font.Size = 10
        .Interior.Color = RGB(226, 239, 218)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(2).RowHeight = 20

    Dim hd As Variant
    hd = Array("順位", "コード", "銘柄名", "現値", "エントリー", "見送りライン", "損切", "利確", _
               "株数", "想定損失", "想定利益", "出来高倍", "週足乖離%", "陽線", "前の下げ", _
               "実体率", "RSI", "25日乖離%", "ATR", "内訳")
    Dim j As Long
    For j = 0 To UBound(hd)
        ws.Cells(3, j + 1).Value = hd(j)
    Next j
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, OUT_COLS))
        .Font.Name = "Meiryo UI": .Font.Size = 10: .Font.Bold = True
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
    Else
        With ws.Range(ws.Cells(OUT_TOP, 1), ws.Cells(OUT_TOP + nOut - 1, OUT_COLS))
            .Font.Name = "Meiryo UI": .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
        ws.Range(ws.Cells(OUT_TOP, 4), ws.Cells(OUT_TOP + nOut - 1, 11)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(OUT_TOP, 12), ws.Cells(OUT_TOP + nOut - 1, 19)).NumberFormat = "0.0"
        ws.Range(ws.Cells(OUT_TOP, 1), ws.Cells(OUT_TOP + IIf(nOut < HA_SLOTS, nOut, HA_SLOTS) - 1, OUT_COLS)).Interior.Color = RGB(198, 239, 206)
    End If

    Dim rr As Long
    rr = OUT_TOP + nOut + 2
    ws.Cells(rr, 1).Value = "絞り込みの内訳：" & funnel
    rr = rr + 2
    ws.Cells(rr, 1).Value = "【売買ルール　この通りにやる】"
    ws.Cells(rr + 1, 1).Value = "1. 上位（出来高倍率の高い順）から、翌日の寄り成りで買う。同時保有は" & HA_SLOTS & "銘柄まで。"
    ws.Cells(rr + 2, 1).Value = "2. ただし寄り値が「見送りライン」より高く始まったら、その銘柄は買わない。"
    ws.Cells(rr + 3, 1).Value = "3. 買った直後に「損切」を逆指値、「利確」に売り指値。両方すぐ出す。"
    ws.Cells(rr + 4, 1).Value = "4. どちらにも当たらなければ、買った日から" & HA_HOLDDAYS & "営業日後の引けで成行手じまい。"
    ws.Cells(rr + 5, 1).Value = "5. 途中で判断しない。出した注文をいじらない。"
    ws.Cells(rr + 6, 1).Value = "6. 1銘柄は資金の" & Format(HA_POS * 100, "0") & "%（" & Format(HA_CAPITAL * HA_POS, "#,##0") & "円）。3連敗したら株数を半分。月の損失が資金の6%で当月休み。"
    ws.Cells(rr + 7, 1).Value = "※検証：勝率63〜65%／1回平均+1.5〜2.0%／PF1.7〜2.1。ただし直近四半期は赤字。今は少額テストのみ。"
    With ws.Range(ws.Cells(rr, 1), ws.Cells(rr + 7, 1))
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    ws.Cells(rr, 1).Font.Bold = True
    ws.Cells(rr + 7, 1).Font.Color = RGB(192, 0, 0)

    ws.Columns("A:T").AutoFit
    If ws.Columns("C").ColumnWidth > 18 Then ws.Columns("C").ColumnWidth = 18
    If ws.Columns("T").ColumnWidth > 60 Then ws.Columns("T").ColumnWidth = 60
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
        If k >= 1 Then s = s + x(k): cnt = cnt + 1
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

'陽線が何本続いているか
Private Function HA_RunLen(ByRef haO() As Double, ByRef haC() As Double, ByVal n As Long) As Long
    Dim k As Long, cnt As Long
    For k = n To 1 Step -1
        If haC(k) > haO(k) Then cnt = cnt + 1 Else Exit For
    Next k
    HA_RunLen = cnt
End Function

'その前に陰線が何本続いていたか
Private Function HA_RunPrev(ByRef haO() As Double, ByRef haC() As Double, _
                            ByVal n As Long, ByVal runLen As Long) As Long
    Dim k As Long, cnt As Long
    For k = n - runLen To 1 Step -1
        If haC(k) < haO(k) Then cnt = cnt + 1 Else Exit For
    Next k
    HA_RunPrev = cnt
End Function

Private Function HA_Rsi(ByRef cl() As Double, ByVal idx As Long, ByVal p As Long) As Double
    HA_Rsi = 50
    If idx < p + 1 Then Exit Function
    Dim k As Long, up As Double, dn As Double, dd As Double
    For k = idx - p + 1 To idx
        dd = cl(k) - cl(k - 1)
        If dd > 0 Then up = up + dd Else dn = dn - dd
    Next k
    up = up / p: dn = dn / p
    If dn = 0 Then HA_Rsi = 100 Else HA_Rsi = 100# - 100# / (1# + up / dn)
End Function

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

'週足の終値が10週線からどれだけ離れているか（%）　-999＝判定できない
Private Function HA_WeekDev(ByRef cl() As Double, ByRef dts() As Double, ByVal n As Long) As Double

    HA_WeekDev = -999

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
            wc(wCnt) = cl(k)
        End If
    Next k

    Dim w As Long
    w = wCnt - 1                      '一番新しい週は途中なので1つ前を使う
    If w < HA_WEEK_MA + 1 Then Exit Function

    Dim s As Double, j As Long
    For j = w - HA_WEEK_MA + 1 To w
        s = s + wc(j)
    Next j
    If s <= 0 Then Exit Function
    Dim ma As Double
    ma = s / HA_WEEK_MA
    HA_WeekDev = (wc(w) - ma) / ma * 100#
End Function

'==================================================================
' 売買ルールを「平均足ルール」シートに書き出す
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
        .Value = "  平均足　売買ルール（検証済み・条件B）"
        .Font.Name = "Meiryo UI": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
    End With
    ws.Rows(1).RowHeight = 30

    Dim r As Long
    r = 3
    HA_PutH ws, r, "【1】買う条件（全部そろった時だけ）"
    HA_PutR ws, r, "1", "平均足が陽線2〜5本連続（その前に陰線3本以上）", "下げの終わりを取る"
    HA_PutR ws, r, "2", "最新の平均足に下ヒゲが無い", "買い方が優勢な形"
    HA_PutR ws, r, "3", "直近3日の安値 > その前4〜8日の安値", "安値切り上げ"
    HA_PutR ws, r, "4", "終値が20日線の上", "逆行しているものは買わない"
    HA_PutR ws, r, "5", "当日出来高 ≧ 5日平均 × " & HA_VOLRATE & "倍", "ここが一番効く"
    HA_PutR ws, r, "6", "週足の終値が" & HA_WEEK_MA & "週線の上", "大きな流れに逆らわない"
    HA_PutR ws, r, "7", "株価 " & Format(HA_MINPRICE, "#,##0") & "円以上／5日平均出来高 20万株以上", "低位株・薄い株を外す"
    HA_PutR ws, r, "8", "翌朝の寄りが「見送りライン」（現値+" & Format(HA_MAXGAP * 100, "0") & "%）より高ければ買わない", "高値づかみ防止"
    r = r + 1

    HA_PutH ws, r, "【2】出口（この3つだけ。途中で判断しない）"
    HA_PutR ws, r, "1", "損切 ＝ 買値の −" & Format(HA_SL * 100, "0") & "%（逆指値）", "買った直後に入れる。絶対に下げない"
    HA_PutR ws, r, "2", "利確 ＝ 買値の +" & Format(HA_TP * 100, "0") & "%（売り指値）", "届いたら全部売り"
    HA_PutR ws, r, "3", "買った日を0日目として" & HA_HOLDDAYS & "営業日後の引けで成行手じまい", "どちらにも当たらなかった玉"
    HA_PutR ws, r, "※", "損切・利確は「実際の買値」が基準（前日終値ではない）", "GUして買ったらその値段の±8%"
    r = r + 1

    HA_PutH ws, r, "【3】資金管理"
    HA_PutR ws, r, "1", "1銘柄に資金の" & Format(HA_POS * 100, "0") & "%（" & Format(HA_CAPITAL * HA_POS, "#,##0") & "円）", "100株単位で切り捨て"
    HA_PutR ws, r, "2", "同時保有は" & HA_SLOTS & "銘柄まで。出来高倍率の高い順に採用", "点数は使いません"
    HA_PutR ws, r, "3", "同じ銘柄を持っている間は買い増ししない", ""
    HA_PutR ws, r, "4", "3連敗したら株数を半分／月の損失が資金の6%で当月休み", ""
    r = r + 1

    HA_PutH ws, r, "【4】検証結果（300銘柄×250日・手数料往復0.3%込み）"
    HA_PutR ws, r, "", "取引数 / 勝率", "115件 / 66.1%"
    HA_PutR ws, r, "", "1回あたりの平均", "+1.87%（勝ち+6.16% / 負け-6.49%）"
    HA_PutR ws, r, "", "PF（稼いだ÷失った）", "1.85"
    HA_PutR ws, r, "", "最大ドローダウン / 最大連敗", "4.6% / 6連敗"
    HA_PutR ws, r, "", "前半PF / 後半PF", "3.59 / 1.30"
    HA_PutR ws, r, "", "直近3か月", "PF1.03・平均+0.08%（ほぼトントン）"
    HA_PutR ws, r, "", "資金300万・1銘柄15%での実績", "65件 勝率66.2% PF2.02 最大DD2.7% 年+13.8%"
    r = r + 1

    HA_PutH ws, r, "【5】注意（必ず読む）"
    HA_PutR ws, r, "1", "判定は「少額でテスト運用」。実戦投入可ではありません", "記録30件で再判定します"
    HA_PutR ws, r, "2", "検証は1年分だけ。しかも全銘柄平均+64%という異常な上げ相場でした", ""
    HA_PutR ws, r, "3", "点数制度は廃止（点数が高いほど成績が悪かった）", "4〜6点PF2.77／9点以上PF0.89"
    HA_PutR ws, r, "4", "空売りは使わない（勝率41%・1回-3.41%）", ""
    HA_PutR ws, r, "5", "データが2営業日以上古い時は候補を出さない", "⑪データチェックで確認"

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 46
    With ws.Range(ws.Cells(3, 1), ws.Cells(r - 1, 3))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
        .VerticalAlignment = xlTop
    End With
    ws.Rows("3:" & r).AutoFit

    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "「平均足ルール」シートを作りました。", vbInformation
End Sub

Private Sub HA_PutH(ByVal ws As Worksheet, ByRef r As Long, ByVal s As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Merge
        .Value = s
        .Font.Name = "Meiryo UI": .Font.Size = 11: .Font.Bold = True
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
        .Font.Name = "Meiryo UI": .Font.Size = 10: .WrapText = True
    End With
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 3).Font.Color = RGB(90, 90, 90)
    r = r + 1
End Sub
