Attribute VB_Name = "MS2_Core"
Option Explicit

'============================================================
' MS2 コアマクロ（列順: No / 株探リンク / 銘柄コード / 銘柄名 / 現在値～）
'   DATA_MS2（1行目=帯 / 2行目=見出し / 3行目～=データ）
'     A=No  B=株探リンク  C=銘柄コード  D=銘柄名
'     E=現在値 F=高値 G=安値 H=終値 I=出来高 J=前日高値 K=前日安値 L=前日終値
'     M=寄り方向 N=ATR O=ゾーン P=STOP-BUY Q=STOP-SELL R=売買種別
'     S=Entry T=SL U=TP1 V=TP2 W=TP3 X=2回目戻し Y=時間帯
'     Z=ボラフィルタ AA=ダマシ除去 AB=トレンド方向
'   ・入力列は E～L。数値の空/エラー値は自動で 0 扱い（NzD）。
'   ・ATRは銘柄別の14期間ATR（Wilder平滑）を ATRHIST_MS2 に蓄積（N列）。
'
' 【導入手順】
'   1) VBEで 旧 Module1 / 重複 Module2 / 古い MS2_Core を削除
'   2) ファイル→ファイルのインポート→本ファイル
'   3) マクロ一覧の MS2_Build_Menu を実行→「操作パネル_MS2」にボタン
'============================================================

Private Const DATA_FIRST As Long = 3
Private Const STOCK_MAX As Long = 350
Private Const HIST As String = "ATRHIST_MS2"
Private Const ATR_N As Long = 14

'============================================================
' 数値安全読み取り：数値なら値、そうでなければ 0
'============================================================
Private Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v) Else NzD = 0#
    On Error GoTo 0
End Function

'============================================================
' A=No / B=株探リンク / C=銘柄コード / D=銘柄名 を再設定（350銘柄）
'============================================================
Sub MS2_Update_StockList_To_DATA()

    Dim wsD As Worksheet
    Dim i As Long, s As Long
    Dim q As String, sref As String, nref As String

    Set wsD = Sheets("DATA_MS2")
    q = Chr(34)

    For i = DATA_FIRST To DATA_FIRST + STOCK_MAX - 1
        s = i - 1
        sref = "'StockList_MS2'!A" & s
        nref = "'StockList_MS2'!B" & s
        wsD.Cells(i, "A").Formula = "=IF(C" & i & "=" & q & q & "," & q & q & ",ROW()-2)"
        wsD.Cells(i, "B").Formula = _
            "=IF(" & sref & "=" & q & q & "," & q & q & "," & _
            "HYPERLINK(" & q & "https://kabutan.jp/stock/?code=" & q & "&" & sref & _
            "," & q & "株探" & q & "))"
        wsD.Cells(i, "C").Formula = "=IF(" & sref & "=" & q & q & "," & q & q & "," & sref & ")"
        wsD.Cells(i, "D").Formula = "=IF(" & sref & "=" & q & q & "," & q & q & "," & nref & ")"
    Next i

    MsgBox "A:No / B:株探 / C:銘柄コード / D:銘柄名 を再設定しました（350銘柄）。", vbInformation, "MS2"

End Sub

'============================================================
' RssMarket 式を E～L にセット（3行目開始）
'============================================================
Sub MS2_Set_RssMarket_Formulas()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim code As String

    Set ws = Sheets("DATA_MS2")
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = DATA_FIRST To lastRow
        code = Trim(CStr(ws.Cells(i, "C").Value))
        If code <> "" Then
            ws.Cells(i, "E").Formula = "=RssMarket(""" & code & """,""現在値"")"
            ws.Cells(i, "F").Formula = "=RssMarket(""" & code & """,""高値"")"
            ws.Cells(i, "G").Formula = "=RssMarket(""" & code & """,""安値"")"
            ws.Cells(i, "H").Formula = "=RssMarket(""" & code & """,""終値"")"
            ws.Cells(i, "I").Formula = "=RssMarket(""" & code & """,""出来高"")"
            ws.Cells(i, "J").Formula = "=RssMarket(""" & code & """,""前日高値"")"
            ws.Cells(i, "K").Formula = "=RssMarket(""" & code & """,""前日安値"")"
            ws.Cells(i, "L").Formula = "=RssMarket(""" & code & """,""前日終値"")"
        End If
    Next i

End Sub

'============================================================
' ATR履歴シート（無ければ作成・非表示）
'============================================================
Private Function MS2_Get_History() As Worksheet

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(HIST)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = HIST
        ws.Range("A1").Value = "銘柄コード"
        ws.Range("B1").Value = "サンプル数"
        ws.Range("C1").Value = "TR累計"
        ws.Range("D1").Value = "ATR"
        ws.Visible = xlSheetHidden
    End If
    Set MS2_Get_History = ws

End Function

'============================================================
' ATR(14) 銘柄別（Wilder平滑）→ N列
'============================================================
Sub MS2_Calc_ATR_5min_14()

    Dim ws As Worksheet, wsH As Worksheet
    Dim i As Long, lastRow As Long, hLast As Long, hr As Long
    Dim code As String
    Dim high As Double, low As Double, prevClose As Double, TR As Double
    Dim cnt As Long, sumTR As Double, atr As Double
    Dim dict As Object

    Set ws = Sheets("DATA_MS2")
    Set wsH = MS2_Get_History()

    Set dict = CreateObject("Scripting.Dictionary")
    hLast = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
    For hr = 2 To hLast
        code = Trim(CStr(wsH.Cells(hr, "A").Value))
        If code <> "" Then dict(code) = hr
    Next hr

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = DATA_FIRST To lastRow

        code = Trim(CStr(ws.Cells(i, "C").Value))
        If code = "" Then GoTo NextI

        high = NzD(ws.Cells(i, "F").Value)       ' 高値
        low = NzD(ws.Cells(i, "G").Value)        ' 安値
        prevClose = NzD(ws.Cells(i, "L").Value)  ' 前日終値

        If high = 0 And low = 0 And prevClose = 0 Then
            ws.Cells(i, "N").Value = 0
            GoTo NextI
        End If

        TR = Application.WorksheetFunction.Max( _
                high - low, Abs(high - prevClose), Abs(low - prevClose))

        If dict.Exists(code) Then
            hr = dict(code)
        Else
            hr = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row + 1
            If hr < 2 Then hr = 2
            wsH.Cells(hr, "A").Value = code
            wsH.Cells(hr, "B").Value = 0
            wsH.Cells(hr, "C").Value = 0
            wsH.Cells(hr, "D").Value = 0
            dict(code) = hr
        End If

        cnt = CLng(NzD(wsH.Cells(hr, "B").Value))

        If cnt < ATR_N Then
            cnt = cnt + 1
            sumTR = NzD(wsH.Cells(hr, "C").Value) + TR
            atr = sumTR / cnt
            wsH.Cells(hr, "B").Value = cnt
            wsH.Cells(hr, "C").Value = sumTR
        Else
            atr = (NzD(wsH.Cells(hr, "D").Value) * (ATR_N - 1) + TR) / ATR_N
        End If

        wsH.Cells(hr, "D").Value = atr
        ws.Cells(i, "N").Value = atr        ' ATR列

NextI:
    Next i

End Sub

'============================================================
' ATR履歴のリセット
'============================================================
Sub MS2_Reset_ATR_History()
    Dim wsH As Worksheet
    Set wsH = MS2_Get_History()
    wsH.Range("A2:D1000000").ClearContents
    MsgBox "ATR履歴をリセットしました。次回のATR更新から再蓄積します。", vbInformation, "MS2"
End Sub

'============================================================
' 売買ロジック（新列・堅牢化：空/エラー値は0扱い）
'============================================================
Sub MS2_Stock_Logic_Run()

    Dim ws As Worksheet, wsT As Worksheet
    Dim lastRow As Long, i As Long, tradeRow As Long
    Dim dir As String, stopBuy As String, stopSell As String
    Dim zone As String, kind As String
    Dim volaFlag As String, fakeFlag As String, trendFlag As String, secondPull As String
    Dim vol As Double, prevVol As Double, atr As Double
    Dim entry As Double, sl As Double, tp1 As Double, tp2 As Double, tp3 As Double
    Dim code As String, tNow As Date, tmStr As String
    Dim eCur As Double, fHigh As Double, gLow As Double, hClose As Double
    Dim jPH As Double, kPL As Double, lPC As Double

    Set ws = Sheets("DATA_MS2")
    Set wsT = Sheets("TRADE_MS2")

    wsT.Range("A2:L1000").ClearContents

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    tradeRow = 2
    tNow = Now
    tmStr = Format(tNow, "hh:mm")

    For i = DATA_FIRST To lastRow

        code = Trim(CStr(ws.Cells(i, "C").Value))
        If code = "" Then GoTo NextI

        eCur = NzD(ws.Cells(i, "E").Value)    ' 現在値
        fHigh = NzD(ws.Cells(i, "F").Value)   ' 高値
        gLow = NzD(ws.Cells(i, "G").Value)    ' 安値
        hClose = NzD(ws.Cells(i, "H").Value)  ' 終値
        vol = NzD(ws.Cells(i, "I").Value)     ' 出来高
        jPH = NzD(ws.Cells(i, "J").Value)     ' 前日高値
        kPL = NzD(ws.Cells(i, "K").Value)     ' 前日安値
        lPC = NzD(ws.Cells(i, "L").Value)     ' 前日終値

        If eCur > lPC Then dir = "BUY-DAY" Else dir = "SELL-DAY"
        ws.Cells(i, "M").Value = dir

        If vol > 0 Then
            prevVol = vol / 3
            If vol > prevVol * 2.5 Then zone = "VOL-OK" Else zone = "VOL-NG"
        Else
            zone = "VOL-NG"
        End If
        ws.Cells(i, "O").Value = zone

        If eCur > jPH Then stopBuy = "ON" Else stopBuy = "OFF"
        If eCur < kPL Then stopSell = "ON" Else stopSell = "OFF"
        ws.Cells(i, "P").Value = stopBuy
        ws.Cells(i, "Q").Value = stopSell

        atr = NzD(ws.Cells(i, "N").Value)
        If atr <= 0 Then GoTo NextI

        If atr > eCur * 0.05 Then volaFlag = "VOL-HIGH" Else volaFlag = "VOL-NORMAL"
        ws.Cells(i, "Z").Value = volaFlag

        If fHigh > jPH And gLow > kPL Then
            trendFlag = "UP"
        ElseIf fHigh < jPH And gLow < kPL Then
            trendFlag = "DOWN"
        Else
            trendFlag = "FLAT"
        End If
        ws.Cells(i, "AB").Value = trendFlag

        If tmStr >= "09:00" And tmStr <= "09:30" Then
            ws.Cells(i, "Y").Value = "寄り付き"
        ElseIf tmStr >= "12:30" And tmStr <= "14:30" Then
            ws.Cells(i, "Y").Value = "後場"
        ElseIf tmStr >= "14:30" And tmStr <= "15:00" Then
            ws.Cells(i, "Y").Value = "引け前"
        Else
            ws.Cells(i, "Y").Value = "その他"
        End If

        If eCur > hClose And Abs(eCur - hClose) < atr * 0.5 Then
            secondPull = "2ND"
        Else
            secondPull = "1ST"
        End If
        ws.Cells(i, "X").Value = secondPull

        If zone = "VOL-OK" And volaFlag = "VOL-NORMAL" And secondPull = "2ND" Then
            fakeFlag = "OK"
        Else
            fakeFlag = "NG"
        End If
        ws.Cells(i, "AA").Value = fakeFlag

        kind = ""
        If dir = "BUY-DAY" And zone = "VOL-OK" And stopSell = "ON" And fakeFlag = "OK" And trendFlag <> "DOWN" Then
            kind = "BUY"
        ElseIf dir = "SELL-DAY" And zone = "VOL-OK" And stopBuy = "ON" And fakeFlag = "OK" And trendFlag <> "UP" Then
            kind = "SELL"
        End If
        ws.Cells(i, "R").Value = kind

        If kind = "" Then GoTo NextI

        entry = eCur
        If kind = "BUY" Then
            sl = entry - atr * 1.5
            tp1 = entry + atr * 1.5
            tp2 = entry + atr * 2#
            tp3 = entry + atr * 3#
        Else
            sl = entry + atr * 1.5
            tp1 = entry - atr * 1.5
            tp2 = entry - atr * 2#
            tp3 = entry - atr * 3#
        End If

        ws.Cells(i, "S").Value = entry
        ws.Cells(i, "T").Value = sl
        ws.Cells(i, "U").Value = tp1
        ws.Cells(i, "V").Value = tp2
        ws.Cells(i, "W").Value = tp3

        wsT.Cells(tradeRow, "A").Value = tradeRow - 1
        wsT.Cells(tradeRow, "B").Value = code
        wsT.Cells(tradeRow, "C").Value = kind
        wsT.Cells(tradeRow, "D").Value = entry
        wsT.Cells(tradeRow, "E").Value = sl
        wsT.Cells(tradeRow, "F").Value = tp1
        wsT.Cells(tradeRow, "G").Value = tp2
        wsT.Cells(tradeRow, "H").Value = tp3
        wsT.Cells(tradeRow, "K").Value = tNow
        tradeRow = tradeRow + 1

NextI:
    Next i

End Sub

'============================================================
' TRADE結果判定（簡易：SL/TP1/END）
'============================================================
Sub MS2_Eval_Trades()

    Dim wsT As Worksheet
    Dim lastRow As Long, i As Long
    Dim kind As String
    Dim entry As Double, sl As Double, tp1 As Double
    Dim result As String, pl As Double

    Set wsT = Sheets("TRADE_MS2")
    lastRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        kind = CStr(wsT.Cells(i, "C").Value)
        entry = NzD(wsT.Cells(i, "D").Value)
        sl = NzD(wsT.Cells(i, "E").Value)
        tp1 = NzD(wsT.Cells(i, "F").Value)
        result = "": pl = 0
        If kind = "BUY" Then
            If tp1 > entry Then
                result = "TP1": pl = tp1 - entry
            ElseIf sl < entry Then
                result = "SL": pl = sl - entry
            Else
                result = "END": pl = tp1 - entry
            End If
        ElseIf kind = "SELL" Then
            If tp1 < entry Then
                result = "TP1": pl = entry - tp1
            ElseIf sl > entry Then
                result = "SL": pl = entry - sl
            Else
                result = "END": pl = entry - tp1
            End If
        End If
        wsT.Cells(i, "I").Value = result
        wsT.Cells(i, "J").Value = pl
        wsT.Cells(i, "L").Value = Now
    Next i

End Sub

'============================================================
' RANK集計（A=株探リンク / D=銘柄名 / 指標）
'============================================================
Sub MS2_Update_Ranking()

    Dim wsT As Worksheet, wsR As Worksheet, wsL As Worksheet
    Dim dict As Object, nameDict As Object
    Dim lastRow As Long, lr As Long, k As Long, i As Long, c As Long, r As Long
    Dim v As Variant, arr, pl As Double
    Dim winRate As Double, pf As Double, rr As Double
    Dim maxDD As Double, avgRR As Double
    Dim streakWin As Long, streakLose As Long
    Dim monthPL As Double, tradeCount As Long
    Dim q As String, cc As String, nm As String, key As String
    Dim hdr As Variant

    Set wsT = Sheets("TRADE_MS2")
    Set wsR = Sheets("RANK_MS2")
    Set wsL = Sheets("StockList_MS2")
    q = Chr(34)

    hdr = Array("銘柄コード（株探）", "勝率", "PF", "銘柄名", "RR", "最大DD", _
                "平均RR", "連勝", "連敗", "月次損益", "トレード数")
    For c = 1 To 11
        With wsR.Cells(1, c)
            .Value = hdr(c - 1)
            .Font.Name = "Meiryo"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    wsR.Rows(1).RowHeight = 46

    wsR.Range("A2:K100000").ClearContents

    Set nameDict = CreateObject("Scripting.Dictionary")
    lr = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row
    For k = 2 To lr
        cc = Trim(CStr(wsL.Cells(k, "A").Value))
        If cc <> "" And Not nameDict.Exists(cc) Then nameDict(cc) = CStr(wsL.Cells(k, "B").Value)
    Next k

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        key = Trim(CStr(wsT.Cells(i, "B").Value))
        If key = "" Then GoTo NextI
        If Not dict.Exists(key) Then
            dict.Add key, Array(0, 0, 0#, 0#, 0#, 0#, 0&, 0&, 0#, 0&)
        End If
        arr = dict(key)
        pl = NzD(wsT.Cells(i, "J").Value)
        arr(9) = arr(9) + 1
        If pl > 0 Then
            arr(0) = arr(0) + 1
            arr(2) = arr(2) + pl
            arr(6) = arr(6) + 1
            arr(7) = 0
        ElseIf pl < 0 Then
            arr(1) = arr(1) + 1
            arr(3) = arr(3) + Abs(pl)
            arr(7) = arr(7) + 1
            arr(6) = 0
            If arr(4) < Abs(pl) Then arr(4) = Abs(pl)
        End If
        If IsDate(wsT.Cells(i, "K").Value) Then
            If Month(wsT.Cells(i, "K").Value) = Month(Date) Then arr(8) = arr(8) + pl
        End If
        dict(key) = arr
NextI:
    Next i

    r = 2
    For Each v In dict.Keys
        arr = dict(v)
        key = CStr(v)
        If arr(0) + arr(1) > 0 Then winRate = arr(0) / (arr(0) + arr(1)) Else winRate = 0
        If arr(3) <> 0 Then pf = arr(2) / arr(3) Else pf = 0
        If arr(0) > 0 Then rr = arr(2) / arr(0) Else rr = 0
        maxDD = arr(4)
        avgRR = rr
        streakWin = arr(6)
        streakLose = arr(7)
        monthPL = arr(8)
        tradeCount = arr(9)
        If nameDict.Exists(key) Then nm = nameDict(key) Else nm = ""

        wsR.Cells(r, "A").Formula = "=HYPERLINK(" & q & "https://kabutan.jp/stock/?code=" & q & "&" & q & key & q & "," & q & key & q & ")"
        wsR.Cells(r, "B").Value = winRate
        wsR.Cells(r, "C").Value = pf
        wsR.Cells(r, "D").Value = nm
        wsR.Cells(r, "E").Value = rr
        wsR.Cells(r, "F").Value = maxDD
        wsR.Cells(r, "G").Value = avgRR
        wsR.Cells(r, "H").Value = streakWin
        wsR.Cells(r, "I").Value = streakLose
        wsR.Cells(r, "J").Value = monthPL
        wsR.Cells(r, "K").Value = tradeCount
        r = r + 1
    Next v

End Sub

'============================================================
' ログ表示
'============================================================
Sub MS2_Show_Log()
    Dim wsT As Worksheet, lastRow As Long
    Set wsT = Sheets("TRADE_MS2")
    lastRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
    MsgBox "TRADE_MS2 のトレード数: " & (lastRow - 1), vbInformation, "MS2_LOG"
End Sub

'============================================================
' DATA_MS2 の見出し(2行目)・タイトル帯(1行目)を修復（28列）
'============================================================
Sub MS2_Fix_DATA_Header()

    Dim ws As Worksheet
    Dim headers As Variant
    Dim c As Long, nCol As Long

    Set ws = Sheets("DATA_MS2")
    headers = Array("No", "株探リンク", "銘柄コード", "銘柄名", "現在値", "高値", "安値", "終値", _
        "出来高", "前日高値", "前日安値", "前日終値", "寄り方向", "ATR(14・5分足)", "ゾーン", _
        "STOP-BUY", "STOP-SELL", "売買種別", "Entry", "SL", "TP1", "TP2", "TP3", _
        "2回目戻し", "時間帯", "ボラフィルタ", "ダマシ除去", "トレンド方向")
    nCol = UBound(headers) - LBound(headers) + 1     ' 28

    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCol)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCol)).Merge
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCol)).Interior.Color = RGB(189, 215, 238)
    ws.Rows(1).RowHeight = 44

    For c = 1 To nCol
        With ws.Cells(2, c)
            .Value = headers(c - 1)
            .Font.Name = "Meiryo"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
    Next c
    ws.Rows(2).RowHeight = 46

    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("E3").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    MsgBox "DATA_MS2 の見出し（2行目）を修復しました。", vbInformation, "MS2"

End Sub

'============================================================
' 全自動
'============================================================
Sub MS2_Auto_All()
    MS2_Set_RssMarket_Formulas
    MS2_Calc_ATR_5min_14
    MS2_Stock_Logic_Run
    MS2_Eval_Trades
    MS2_Update_Ranking
End Sub

'============================================================
' 操作ボタンを専用シート「操作パネル_MS2」にまとめる
'============================================================
Sub MS2_Build_Menu()

    Dim ws As Worksheet
    Dim btn As Object
    Dim y As Single
    Dim items As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = Sheets("操作パネル_MS2")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Sheets.Add(Before:=Sheets(1))
        ws.Name = "操作パネル_MS2"
    End If

    ws.Cells.Clear
    ws.Buttons.Delete
    ws.Range("A1").Value = "◆ MS2 操作パネル"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    items = Array( _
        "銘柄反映_MS2", "MS2_Update_StockList_To_DATA", _
        "RSS式セット_MS2", "MS2_Set_RssMarket_Formulas", _
        "ATR更新_MS2", "MS2_Calc_ATR_5min_14", _
        "売買抽出_MS2", "MS2_Stock_Logic_Run", _
        "結果判定_MS2", "MS2_Eval_Trades", _
        "ランキング_MS2", "MS2_Update_Ranking", _
        "ヘッダー修復_MS2", "MS2_Fix_DATA_Header", _
        "ATR履歴リセット_MS2", "MS2_Reset_ATR_History", _
        "ログ_MS2", "MS2_Show_Log", _
        "全自動_MS2", "MS2_Auto_All")

    y = 36
    For i = LBound(items) To UBound(items) Step 2
        Set btn = ws.Buttons.Add(20, y, 180, 28)
        btn.Caption = items(i)
        btn.OnAction = items(i + 1)
        y = y + 34
    Next i

    MsgBox "「操作パネル_MS2」シートにボタンを作成しました。", vbInformation, "MS2"

End Sub
