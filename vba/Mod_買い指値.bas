Attribute VB_Name = "Mod_買い指値"
'=============================================================================
' 打ち出のこづち　買い指値モジュール v1 (2026/08/15)
'
'   目的：買抽出v13 の候補に「いくらで指値を置くか」を追加する。
'
'   根拠：300銘柄×250営業日(2025/07/31〜2026/08/14)の実データ検証
'         ・シグナル翌日の安値は、終値から中央値 0.46ATR 下まで落ちる
'         ・寄付成行だと 95.8% が当日一度は含み損（中央 -1.33%）
'         ・終値 - 0.75ATR の指値／3営業日有効 が最も成績が良い
'             寄付成行  平均+0.179R 勝率46.2%
'             指値      平均+0.413R 勝率53.1%
'
'   使い方：買抽出v13 シートで 指値ラダー計算 を実行。U〜AB列に出力する。
'=============================================================================
Option Explicit

Public Const LIMIT_K_LIGHT As Double = 0.5    '第1指値（浅い・約定しやすい）
Public Const LIMIT_K_MAIN  As Double = 0.75   '★本命指値（検証で最良）
Public Const LIMIT_K_DEEP  As Double = 1.25   '深押し指値（滅多に刺さらない）
Public Const LIMIT_VALID_DAYS As Long = 3     '指値の有効日数（営業日）
Public Const ATR_PERIOD As Long = 14
Public Const HITRATE_LOOKBACK As Long = 120   '到達率を測る過去日数
Public Const STOP_ATR_MULT As Double = 2#     'V811設定に合わせた損切ATR倍率
Public Const STOP_MIN_PCT  As Double = 0.03
Public Const STOP_MAX_PCT  As Double = 0.1
Public Const REWARD_RATIO  As Double = 2#     '利確 = 損切幅 × この倍率

Private Const EXT_SHEET   As String = "買抽出v13"
Private Const EXT_DATA_ROW As Long = 4
Private Const CODE_COL    As Long = 3
Private Const PRICE_COL   As Long = 5
Private Const OUT_FIRST   As Long = 21        'U列から出力
Private Const STOCK_DATA_START As Long = 6
Private Const HIST_FIRST_COL   As Long = 5    'E列＝最新日
Private Const SHEET_PW    As String = "ne19480314"

'=============================================================================
Public Sub 指値ラダー計算()
    Dim extWs As Worksheet: Set extWs = GetWsL(EXT_SHEET)
    If extWs Is Nothing Then
        MsgBox "「" & EXT_SHEET & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    Dim hiWs As Worksheet: Set hiWs = GetWsL("高値")
    Dim loWs As Worksheet: Set loWs = GetWsL("安値")
    Dim clWs As Worksheet: Set clWs = GetWsL("終値")
    If hiWs Is Nothing Or loWs Is Nothing Or clWs Is Nothing Then
        MsgBox "「高値」「安値」「終値」シートが必要です。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error Resume Next
    extWs.Unprotect Password:=SHEET_PW
    On Error GoTo 0

    WriteHeader extWs

    Dim rowDic As Object: Set rowDic = BuildCodeIndex(clWs)
    Dim lastRow As Long
    lastRow = extWs.Cells(extWs.Rows.Count, CODE_COL).End(xlUp).Row
    If lastRow < EXT_DATA_ROW Then lastRow = EXT_DATA_ROW

    extWs.Range(extWs.Cells(EXT_DATA_ROW, OUT_FIRST), _
                extWs.Cells(lastRow + 200, OUT_FIRST + 7)).ClearContents

    Dim r As Long, done As Long: done = 0
    For r = EXT_DATA_ROW To lastRow
        Dim code As String
        code = Trim$(CStr(extWs.Cells(r, CODE_COL).Value))
        If code <> "" And code <> "0" Then
            If Len(code) < 4 Then code = Format(code, "0000")
            If rowDic.Exists(code) Then
                If WriteOneRow(extWs, r, CLng(rowDic(code)), hiWs, loWs, clWs) Then done = done + 1
            End If
        End If
    Next r

    extWs.Columns(OUT_FIRST).Resize(, 8).AutoFit
    Application.ScreenUpdating = True
    MsgBox "指値ラダーを計算しました（" & done & " 銘柄）。" & vbCrLf & vbCrLf & _
           "★本命指値 = 終値 − " & LIMIT_K_MAIN & " × ATR14" & vbCrLf & _
           "　" & LIMIT_VALID_DAYS & "営業日有効で出し、刺さらなければ取り消し。" & vbCrLf & _
           "　到達率は、その銘柄の過去" & HITRATE_LOOKBACK & "日の実績です。", _
           vbInformation, "買い指値 v1"
End Sub

'=============================================================================
Private Sub WriteHeader(ByVal ws As Worksheet)
    Dim hdr As Variant
    hdr = Array("ATR14", "第1指値(-" & LIMIT_K_LIGHT & "ATR)", "★本命指値(-" & LIMIT_K_MAIN & "ATR)", _
                "深押し(-" & LIMIT_K_DEEP & "ATR)", "本命到達率", "本命約定時の損切", _
                "本命約定時の利確", "ひとこと")
    Dim i As Long
    ws.Cells(2, OUT_FIRST).Value = "▼ 買い指値 v1：本命指値を " & LIMIT_VALID_DAYS & "営業日有効で出す（終値−" & LIMIT_K_MAIN & "ATR）"
    ws.Cells(2, OUT_FIRST).Font.Bold = True
    For i = 0 To UBound(hdr)
        With ws.Cells(3, OUT_FIRST + i)
            .Value = hdr(i)
            .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub

Private Function WriteOneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal srcRow As Long, _
                             ByVal hiWs As Worksheet, ByVal loWs As Worksheet, _
                             ByVal clWs As Worksheet) As Boolean
    WriteOneRow = False
    Dim nCol As Long: nCol = HistLastCol(clWs)
    If nCol < HIST_FIRST_COL + ATR_PERIOD + 2 Then Exit Function

    Dim H() As Double, L() As Double, C() As Double, cnt As Long
    If Not LoadSeries(hiWs, loWs, clWs, srcRow, nCol, H, L, C, cnt) Then Exit Function
    If cnt < ATR_PERIOD + 2 Then Exit Function

    Dim atrArr() As Double
    ReDim atrArr(1 To cnt)
    If Not BuildATR(H, L, C, cnt, atrArr) Then Exit Function

    Dim atrNow As Double: atrNow = atrArr(cnt)
    Dim closeNow As Double: closeNow = C(cnt)
    If atrNow <= 0 Or closeNow <= 0 Then Exit Function

    '── 画面の現在値があればそちらを基準にする（場中の値動きを反映）
    Dim basePx As Double: basePx = closeNow
    Dim shown As Double: shown = SafeNumL(ws.Cells(r, PRICE_COL).Value, 0)
    If shown > 0 Then
        If Abs(shown / closeNow - 1) < 0.3 Then basePx = shown
    End If

    Dim pLight As Double, pMain As Double, pDeep As Double
    pLight = RoundTick(basePx - LIMIT_K_LIGHT * atrNow)
    pMain = RoundTick(basePx - LIMIT_K_MAIN * atrNow)
    pDeep = RoundTick(basePx - LIMIT_K_DEEP * atrNow)

    Dim hit As Double: hit = HitRate(H, L, C, atrArr, cnt, LIMIT_K_MAIN)

    Dim sd As Double
    sd = STOP_ATR_MULT * atrNow
    If sd < pMain * STOP_MIN_PCT Then sd = pMain * STOP_MIN_PCT
    If sd > pMain * STOP_MAX_PCT Then sd = pMain * STOP_MAX_PCT

    With ws
        .Cells(r, OUT_FIRST).Value = Round(atrNow, 1)
        .Cells(r, OUT_FIRST + 1).Value = pLight
        .Cells(r, OUT_FIRST + 2).Value = pMain
        .Cells(r, OUT_FIRST + 3).Value = pDeep
        .Cells(r, OUT_FIRST + 4).Value = Format(hit, "0.0") & "%"
        .Cells(r, OUT_FIRST + 5).Value = RoundTick(pMain - sd)
        .Cells(r, OUT_FIRST + 6).Value = RoundTick(pMain + REWARD_RATIO * sd)
        .Cells(r, OUT_FIRST + 7).Value = Comment(basePx, atrNow, C, cnt)
        .Range(.Cells(r, OUT_FIRST), .Cells(r, OUT_FIRST + 3)).NumberFormat = "#,##0"
        .Range(.Cells(r, OUT_FIRST + 5), .Cells(r, OUT_FIRST + 6)).NumberFormat = "#,##0"
        With .Cells(r, OUT_FIRST + 2)
            .Interior.Color = RGB(255, 215, 0)
            .Font.Bold = True
        End With
    End With
    WriteOneRow = True
End Function

'=============================================================================
' 系列読み込み：シートは E列＝最新 なので、配列は 1=最古 … cnt=最新 に詰め直す
'=============================================================================
Private Function LoadSeries(ByVal hiWs As Worksheet, ByVal loWs As Worksheet, ByVal clWs As Worksheet, _
                            ByVal srcRow As Long, ByVal nCol As Long, _
                            ByRef H() As Double, ByRef L() As Double, ByRef C() As Double, _
                            ByRef cnt As Long) As Boolean
    LoadSeries = False
    On Error GoTo Fail
    Dim maxN As Long: maxN = nCol - HIST_FIRST_COL + 1
    If maxN > HITRATE_LOOKBACK + ATR_PERIOD + 10 Then maxN = HITRATE_LOOKBACK + ATR_PERIOD + 10
    Dim lastCol As Long: lastCol = HIST_FIRST_COL + maxN - 1

    Dim hv As Variant, lv As Variant, cv As Variant
    hv = hiWs.Range(hiWs.Cells(srcRow, HIST_FIRST_COL), hiWs.Cells(srcRow, lastCol)).Value
    lv = loWs.Range(loWs.Cells(srcRow, HIST_FIRST_COL), loWs.Cells(srcRow, lastCol)).Value
    cv = clWs.Range(clWs.Cells(srcRow, HIST_FIRST_COL), clWs.Cells(srcRow, lastCol)).Value

    ReDim H(1 To maxN): ReDim L(1 To maxN): ReDim C(1 To maxN)
    cnt = 0
    Dim k As Long
    For k = maxN To 1 Step -1          '古い側から詰める
        Dim h1 As Double, l1 As Double, c1 As Double
        h1 = SafeNumL(hv(1, k), 0): l1 = SafeNumL(lv(1, k), 0): c1 = SafeNumL(cv(1, k), 0)
        If h1 > 0 And l1 > 0 And c1 > 0 And h1 >= l1 Then
            cnt = cnt + 1
            H(cnt) = h1: L(cnt) = l1: C(cnt) = c1
        End If
    Next k
    LoadSeries = (cnt > 0)
    Exit Function
Fail:
End Function

Private Function BuildATR(ByRef H() As Double, ByRef L() As Double, ByRef C() As Double, _
                          ByVal cnt As Long, ByRef outATR() As Double) As Boolean
    BuildATR = False
    If cnt < ATR_PERIOD + 1 Then Exit Function
    Dim i As Long, tr As Double, a As Double, seeded As Boolean
    seeded = False: a = 0
    outATR(1) = 0
    For i = 2 To cnt
        tr = H(i) - L(i)
        If Abs(H(i) - C(i - 1)) > tr Then tr = Abs(H(i) - C(i - 1))
        If Abs(L(i) - C(i - 1)) > tr Then tr = Abs(L(i) - C(i - 1))
        If Not seeded Then
            a = tr: seeded = True
        Else
            a = (a * (ATR_PERIOD - 1) + tr) / ATR_PERIOD
        End If
        outATR(i) = a
    Next i
    BuildATR = (outATR(cnt) > 0)
End Function

'   その銘柄が「終値 − k×ATR」まで翌日に下げた実績の割合
Private Function HitRate(ByRef H() As Double, ByRef L() As Double, ByRef C() As Double, _
                         ByRef atrArr() As Double, ByVal cnt As Long, ByVal k As Double) As Double
    Dim first As Long: first = cnt - HITRATE_LOOKBACK
    If first < ATR_PERIOD + 1 Then first = ATR_PERIOD + 1
    Dim i As Long, n As Long, hitN As Long
    For i = first To cnt - 1
        If atrArr(i) > 0 Then
            n = n + 1
            If L(i + 1) <= C(i) - k * atrArr(i) Then hitN = hitN + 1
        End If
    Next i
    If n = 0 Then
        HitRate = 0
    Else
        HitRate = hitN / n * 100#
    End If
End Function

Private Function Comment(ByVal px As Double, ByVal atrNow As Double, _
                         ByRef C() As Double, ByVal cnt As Long) As String
    Dim s As String
    Dim atrPct As Double: atrPct = atrNow / px * 100#
    If cnt >= 2 Then
        Dim chg As Double: chg = (C(cnt) / C(cnt - 1) - 1) * 100#
        If chg >= 9 Then
            s = s & "前日比" & Format(chg, "+0.0") & "% 急騰直後 "
        ElseIf chg >= 6 Then
            s = s & "前日比" & Format(chg, "+0.0") & "% "
        End If
    End If
    Dim e25 As Double: e25 = EMAOf(C, cnt, 25)
    If e25 > 0 Then
        Dim dev As Double: dev = (px / e25 - 1) * 100#
        If dev >= 15 Then s = s & "25日EMA乖離" & Format(dev, "+0") & "% 高値警戒 "
        If dev < 0 Then s = s & "25日EMA割れ 見送り推奨 "
    End If
    If atrPct >= 5 Then
        s = s & "高ボラ(ATR" & Format(atrPct, "0.0") & "%) 株数を絞る "
    ElseIf atrPct <= 1.5 Then
        s = s & "低ボラ(ATR" & Format(atrPct, "0.0") & "%) 指値は浅めで可 "
    End If
    If s = "" Then s = "標準"
    Comment = Trim$(s)
End Function

Private Function EMAOf(ByRef C() As Double, ByVal cnt As Long, ByVal period As Long) As Double
    If cnt < 2 Then Exit Function
    Dim k As Double: k = 2# / (period + 1#)
    Dim e As Double: e = C(1)
    Dim i As Long
    For i = 2 To cnt
        e = k * C(i) + (1# - k) * e
    Next i
    EMAOf = e
End Function

'=============================================================================
' 補助
'=============================================================================
Private Function BuildCodeIndex(ByVal ws As Worksheet) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, c As String
    For i = STOCK_DATA_START To last
        c = Trim$(CStr(ws.Cells(i, 1).Value))
        If c <> "" And c <> "0" And c <> "TOPX" Then
            If Len(c) < 4 Then c = Format(c, "0000")
            If Not dic.Exists(c) Then dic.Add c, i
        End If
    Next i
    Set BuildCodeIndex = dic
End Function

Private Function HistLastCol(ByVal ws As Worksheet) As Long
    Dim c As Long
    c = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    If c < HIST_FIRST_COL Then c = HIST_FIRST_COL
    HistLastCol = c
End Function

'   呼値に丸める（東証・株式の呼値／TOPIX100構成銘柄以外の一般刻み）
Public Function RoundTick(ByVal p As Double) As Double
    Dim t As Double
    Select Case p
        Case Is < 3000: t = 1
        Case Is < 5000: t = 5
        Case Is < 30000: t = 10
        Case Is < 50000: t = 50
        Case Else: t = 100
    End Select
    RoundTick = Int(p / t) * t          '買い指値なので切り捨て
End Function

Private Function SafeNumL(ByVal v As Variant, ByVal dflt As Double) As Double
    If IsError(v) Then
        SafeNumL = dflt
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeNumL = dflt
    ElseIf IsNumeric(v) Then
        SafeNumL = CDbl(v)
    Else
        SafeNumL = dflt
    End If
End Function

Private Function GetWsL(ByVal n As String) As Worksheet
    On Error Resume Next
    Set GetWsL = ThisWorkbook.Sheets(n)
    On Error GoTo 0
End Function
