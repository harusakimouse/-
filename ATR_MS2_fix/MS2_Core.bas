Attribute VB_Name = "MS2_Core"
Option Explicit
'============================================================
' MS2 コアマクロ v7（2026-08-31 修正版）
'
'  v6 からの変更点（不具合ID は診断レポートに対応）
'  ------------------------------------------------------------
'  A-1  ATRHIST が当日更新でなければ実測ATRを使わない（H17で判定）
'  A-2  前営業日以前の建玉は当日値で SL/TP 判定せず EXPIRED で閉じる
'  A-4  現在値が 0/文字列 の行を判定対象から除外（N()ガード）
'  B-2  銘柄反映は StockList の実件数ぶんだけ書き、余りは消す
'  B-4  シグナル追記時に TRADE_MS2 の N/O/P/Q 補助列も同時に張る
'  B-5  RANK_MS2 の損益を売買単位(H13)で円換算し、TRADE と単位を統一
'  C-1  Workbook_Activate で gShuttingDown を戻す前提。
'       予約が入らなかった場合に「開始しました」と言わないよう修正
'  C-2  AD列/銘柄名の列全体参照を有限範囲に。最終更新の書き込みは足の切替時のみ
'  C-4  足の長さを操作パネル H9 から読む（従来は定数5で固定されていた）
'  C-5  同じ日に二重保存しない
'  C-6  RANK_MS2 を月次損益の降順に並べ替え。PF/RR は数値のまま(999=∞)
'  C-8  エラー時に成功メッセージを出さない
'  C-9  実銘柄数だけをループ
'  ---  TRADE_MS2 に U列（同日損益・内部）を追加。MINIFS を使わず最大損失を出す
'
'  DATA_MS2 列: A=No B=株探 C=コード D=銘柄名
'    E=現在値 F=高値 G=安値 H=終値 I=出来高 J=前日高値 K=前日安値 L=前日終値
'    M=寄り方向 N=ATR O=ゾーン P=STOP-BUY Q=STOP-SELL R=売買種別
'    S=Entry T=SL U=TP1 V=TP2 W=TP3 X=2回目戻し Y=時間帯
'    Z=ボラフィルタ AA=ダマシ除去 AB=トレンド方向 AC=売買代金 AD=実測5分ATR
'
'  ATRHIST_MS2 列: A=コード B=バー高値 C=バー安値 D=バー終値
'                  E=前バー終値 F=本数 G=TR累計 H=ATR I=最終更新
'
'  TRADE_MS2 列: A=No B=コード C=銘柄名 D=種別 E=Entry F=SL
'                G=TP1 H=TP2 I=TP3 J=結果 K=損益(円/株) L=開始時刻 M=終了時刻
'                N=約定代金(円) O=損益(円) P=R倍数 Q=同日決済フラグ
'  RANK_MS2 列: A=コード(株探) B=銘柄名 C=勝率 D=PF E=平均損益(円) F=最大DD(円)
'               G=平均RR H=最大連勝 I=最大連敗 J=月次損益(円) K=トレード数
'============================================================
Private Const DATA_FIRST   As Long = 3
Private Const STOCK_MAX    As Long = 350
Private Const HIST         As String = "ATRHIST_MS2"
Private Const PREV         As String = "PREVDAY_MS2"
Private Const PANEL        As String = "操作パネル_MS2"
Private Const ATR_N        As Long = 14
Private Const ATR_BAR_DEF  As Long = 5      ' 足の長さ既定（実際は H9 を読む）
Private Const TICK_SEC     As Long = 30     ' サンプリング間隔（秒）
Private Const SAVE_HHMM    As String = "15:45"
Private Const TRADE_LAST   As Long = 400    ' 補助列とサマリーの対象最終行
Private Const HIST_LAST    As Long = 400    ' ATRHIST の参照上限
Private Const PF_INF       As Double = 999  ' PF/RR の「∞」を表す数値

'★予約時刻の永続化セル（PREVDAY_MS2 の未使用列）
Private Const CELL_TICK    As String = "Z1"
Private Const CELL_SAVE    As String = "Z2"
Private Const CELL_SAVED   As String = "Z3"   '★最後に前日値を保存した日付

Public gPrevDaySaveTime As Date
Public gTickTime        As Date
Public gShuttingDown    As Boolean
Private gBucket         As Long
Private gSampling       As Boolean
Private gBarMin         As Long              '★足の長さ（開始時に H9 から読む）

'============================================================
' 共通ユーティリティ
'============================================================
Private Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v) Else NzD = 0#
    On Error GoTo 0
End Function

' 0除算・異常値でオーバーフローしない除算
Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo Zero
    If b = 0# Then GoTo Zero
    Dim x As Double
    x = a / b
    If Abs(x) > 1E+300 Then GoTo Zero
    SafeDiv = x
    Exit Function
Zero:
    SafeDiv = 0#
End Function

Private Function PanelNum(ByVal addr As String, ByVal dflt As Double) As Double
    Dim v As Double
    On Error Resume Next
    v = NzD(ThisWorkbook.Sheets(PANEL).Range(addr).Value)
    On Error GoTo 0
    If v <= 0 Then v = dflt
    PanelNum = v
End Function

'★足の長さ（分）— 操作パネル H9。1〜60 に丸める
Private Function BarMinutes() As Long
    Dim m As Long
    m = CLng(PanelNum("H9", ATR_BAR_DEF))
    If m < 1 Then m = 1
    If m > 60 Then m = 60
    BarMinutes = m
End Function

'★DATA_MS2 の実データ最終行（C列に銘柄コードが入っている最後の行）
Private Function DataLastRow(ws As Worksheet) As Long
    Dim r As Long, last As Long
    last = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If last < DATA_FIRST Then
        DataLastRow = DATA_FIRST - 1
        Exit Function
    End If
    For r = last To DATA_FIRST Step -1
        If Len(Trim(CStr(ws.Cells(r, "C").Value))) > 0 Then
            DataLastRow = r
            Exit Function
        End If
    Next r
    DataLastRow = DATA_FIRST - 1
End Function

' 場中経過分（前場9:00-11:30=150分 + 後場12:30-15:30=180分）
Private Function SessionMinutes(ByVal t As Date) As Double
    Dim tt As Double, m As Double
    tt = t - Int(t)
    m = 0
    If tt > TimeValue("09:00") Then m = m + WorksheetFunction.Min((tt - TimeValue("09:00")) * 1440#, 150#)
    If tt > TimeValue("12:30") Then m = m + WorksheetFunction.Min((tt - TimeValue("12:30")) * 1440#, 180#)
    If m < 0 Then m = 0
    SessionMinutes = m
End Function

Private Function InSession(ByVal t As Date) As Boolean
    Dim tt As Double
    tt = t - Int(t)
    If Weekday(t, vbMonday) >= 6 Then Exit Function   '土日は場外
    InSession = (tt >= TimeValue("09:00") And tt <= TimeValue("15:30"))
End Function

' TRADE_MS2 C列用：銘柄名を引く数式（★列全体参照をやめた）
Private Function NameFormula(ByVal rowNo As Long) As String
    NameFormula = "=IFERROR(INDEX(StockList_MS2!$B$2:$B$400,MATCH($B" & rowNo & _
                  ",StockList_MS2!$A$2:$A$400,0)),"""")"
End Function

'★TRADE_MS2 の補助列（N/O/P/Q）を1行ぶん張る
Private Sub PutTradeAux(wsT As Worksheet, ByVal r As Long)
    Dim p As String
    p = "'" & PANEL & "'"
    wsT.Cells(r, "N").Formula = "=IF($E" & r & "="""","""",ROUND($E" & r & "*" & p & "!$H$13,0))"
    wsT.Cells(r, "O").Formula = "=IF($K" & r & "="""","""",ROUND($K" & r & "*" & p & "!$H$13,0))"
    wsT.Cells(r, "P").Formula = "=IF(OR($K" & r & "="""",$E" & r & "="""",$F" & r & "="""",$E" & r & "=$F" & r & "),""""," & _
                                "$K" & r & "/ABS($E" & r & "-$F" & r & "))"
    wsT.Cells(r, "Q").Formula = "=IF(OR($B" & r & "="""",$J" & r & "=""""),""""," & _
                                "IF(AND($J" & r & "<>""EXPIRED"",ISNUMBER($L" & r & "),ISNUMBER($M" & r & ")," & _
                                "INT($L" & r & ")=INT($M" & r & ")),1,0))"
    '★U列＝同日決済ぶんの損益だけを取り出す内部ヘルパー。
    '  MIN は空白と文字列を無視するので、MINIFS（Excel2016以降の関数で
    '  XMLには _xlfn. 接頭辞が要る）を使わずに最大損失を出せる。
    wsT.Cells(r, "U").Formula = "=IF($Q" & r & "=1,$O" & r & ","""")"
End Sub

'============================================================
' ★予約時刻の永続化（VBAリセットでも消えない保険）
'============================================================
Private Sub PutTimerCell(ByVal addr As String, ByVal t As Date)
    On Error Resume Next
    MS2_Get_PrevDay.Range(addr).Value = CDbl(t)
    On Error GoTo 0
End Sub

Private Function GetTimerCell(ByVal addr As String) As Date
    On Error Resume Next
    GetTimerCell = CDate(NzD(MS2_Get_PrevDay.Range(addr).Value))
    On Error GoTo 0
End Function

'============================================================
' 操作パネルの設定ブロック（既存値は保持し、空欄のみ既定値を投入）
'============================================================
Private Sub PutSetting(ws As Worksheet, ByVal r As Long, ByVal label As String, _
                       ByVal dflt As Variant, ByVal fmt As String)
    ws.Cells(r, "G").Value = label
    If Len(Trim(CStr(ws.Cells(r, "H").Value))) = 0 Then ws.Cells(r, "H").Value = dflt
    ws.Cells(r, "H").NumberFormat = fmt
    ws.Cells(r, "H").Font.Color = RGB(0, 0, 255)
    ws.Cells(r, "H").HorizontalAlignment = xlRight
End Sub

Private Sub PutAuto(ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal f As String, ByVal fmt As String)
    ws.Cells(r, "G").Value = label
    ws.Cells(r, "H").Formula = f
    ws.Cells(r, "H").NumberFormat = fmt
    ws.Cells(r, "H").Font.Color = RGB(0, 0, 0)
    ws.Cells(r, "H").HorizontalAlignment = xlRight
End Sub

Sub MS2_Install_Settings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PANEL)
    ws.Range("G2").Value = "◆ 判定パラメータ設定"
    ws.Range("H2").Value = "値"
    With ws.Range("G2:H2")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
    End With
    PutSetting ws, 3, "最低売買代金（千円）", 1000000, "#,##0"
    PutSetting ws, 4, "ボラ上限（現在値比）", 0.015, "0.0%"
    PutSetting ws, 5, "SL倍率（×ATR）", 1.5, "0.0"
    PutSetting ws, 6, "TP1倍率（×ATR）", 1.5, "0.0"
    PutSetting ws, 7, "TP2倍率（×ATR）", 2#, "0.0"
    PutSetting ws, 8, "TP3倍率（×ATR）", 3#, "0.0"
    PutSetting ws, 9, "足の長さ（分）", ATR_BAR_DEF, "0"
    PutSetting ws, 10, "ATR調整係数", 1#, "0.00"
    PutAuto ws, 11, "経過場中分数（自動）", _
        "=MEDIAN(0,(MOD(NOW(),1)-TIME(9,0,0))*1440,150)+MEDIAN(0,(MOD(NOW(),1)-TIME(12,30,0))*1440,180)", "#,##0"
    PutAuto ws, 12, "経過足本数（自動）", "=MAX(1,H11/MAX(1,H9))", "#,##0.0"
    PutSetting ws, 13, "売買単位（株）", 100, "#,##0"
    '★v7 追加：Y列の NOW() を1セルに集約 ＆ ATR履歴の鮮度判定
    PutAuto ws, 15, "現在時刻（自動）", "=TEXT(NOW(),""hh:mm"")", "@"
    PutAuto ws, 16, "ATR履歴の最終更新（自動）", _
        "=IFERROR(MAX(" & HIST & "!$I$2:$I$" & HIST_LAST & "),0)", "yyyy/mm/dd hh:mm"
    PutAuto ws, 17, "実測ATR採用 1=当日/0=推定（自動）", "=IF(INT($H$16)=TODAY(),1,0)", "0"
    ws.Range("G14").Value = "※ ATRは ATRHIST_MS2 が「当日更新済み」のときだけ実測5分足(AD列)を使います。" & _
                            "古い/未蓄積なら 当日レンジ÷√経過足本数 に自動フォールバックします（H17で判定）。"
    ws.Range("G14").Font.Color = RGB(192, 0, 0)
    ws.Range("G14").Font.Size = 11
    ws.Columns("G").AutoFit
End Sub

'============================================================
' ★判定数式の一括セット
'============================================================
Sub MS2_Install_Formulas()
    Dim ws As Worksheet, lastRow As Long, p As String, i As Long
    Dim ge As String, gn As String, gj As String, gk As String, gl As String
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    p = "'" & PANEL & "'"
    lastRow = DataLastRow(ws)
    If lastRow < DATA_FIRST Then
        MsgBox "DATA_MS2 に銘柄がありません。先に「銘柄反映_MS2」を実行してください。", vbExclamation, "MS2"
        Exit Sub
    End If
    MS2_Install_Settings

    i = DATA_FIRST
    ge = "N($E" & i & ")<=0"
    gn = "N($N" & i & ")<=0"
    gj = "N($J" & i & ")<=0"
    gk = "N($K" & i & ")<=0"
    gl = "N($L" & i & ")<=0"

    Application.ScreenUpdating = False
    On Error GoTo Fail

    ws.Range("M" & i).Formula = "=IF(OR(" & ge & "," & gl & "),""""," & _
        "IF($E" & i & ">$L" & i & ",""BUY-DAY"",""SELL-DAY""))"

    ws.Range("N" & i).Formula = "=IF(OR(" & ge & ",N($F" & i & ")<=0,N($G" & i & ")<=0),""""," & _
        "IF(AND($AD" & i & ">0," & p & "!$H$17=1),$AD" & i & "," & _
        "MAX(0,$F" & i & "-$G" & i & ")/SQRT(MAX(1," & p & "!$H$12))*" & p & "!$H$10))"

    ws.Range("O" & i).Formula = "=IF(OR(" & ge & ",$AC" & i & "=""""),""""," & _
        "IF(N($AC" & i & ")>=" & p & "!$H$3,""VOL-OK"",""VOL-NG""))"

    ws.Range("P" & i).Formula = "=IF(OR(" & ge & "," & gj & "),""""," & _
        "IF($E" & i & ">$J" & i & ",""ON"",""OFF""))"

    ws.Range("Q" & i).Formula = "=IF(OR(" & ge & "," & gk & "),""""," & _
        "IF($E" & i & "<$K" & i & ",""ON"",""OFF""))"

    ws.Range("R" & i).Formula = "=IF(OR(" & ge & "," & gn & "),""""," & _
        "IF(AND($M" & i & "=""BUY-DAY"",$O" & i & "=""VOL-OK"",$P" & i & "=""ON""," & _
        "$AA" & i & "=""OK"",$AB" & i & "<>""DOWN""),""BUY""," & _
        "IF(AND($M" & i & "=""SELL-DAY"",$O" & i & "=""VOL-OK"",$Q" & i & "=""ON""," & _
        "$AA" & i & "=""OK"",$AB" & i & "<>""UP""),""SELL"","""")))"

    ws.Range("S" & i).Formula = "=IF($R" & i & "="""","""",$E" & i & ")"
    ws.Range("T" & i).Formula = "=IF($R" & i & "="""","""",IF($R" & i & "=""BUY""," & _
        "$E" & i & "-$N" & i & "*" & p & "!$H$5,$E" & i & "+$N" & i & "*" & p & "!$H$5))"
    ws.Range("U" & i).Formula = "=IF($R" & i & "="""","""",IF($R" & i & "=""BUY""," & _
        "$E" & i & "+$N" & i & "*" & p & "!$H$6,$E" & i & "-$N" & i & "*" & p & "!$H$6))"
    ws.Range("V" & i).Formula = "=IF($R" & i & "="""","""",IF($R" & i & "=""BUY""," & _
        "$E" & i & "+$N" & i & "*" & p & "!$H$7,$E" & i & "-$N" & i & "*" & p & "!$H$7))"
    ws.Range("W" & i).Formula = "=IF($R" & i & "="""","""",IF($R" & i & "=""BUY""," & _
        "$E" & i & "+$N" & i & "*" & p & "!$H$8,$E" & i & "-$N" & i & "*" & p & "!$H$8))"

    ws.Range("X" & i).Formula = "=IF(OR(" & ge & "," & gn & "," & gl & "),""""," & _
        "IF(AND($E" & i & ">$L" & i & ",ABS($E" & i & "-$L" & i & ")<$N" & i & "*0.5),""2ND"",""1ST""))"

    '★NOW() を350セルに撒くのをやめ、操作パネル H15 の1セットを参照。前場も追加
    ws.Range("Y" & i).Formula = "=IF(OR(" & ge & "," & gn & "),""""," & _
        "IF(AND(" & p & "!$H$15>=""09:00""," & p & "!$H$15<=""09:30""),""寄り付き""," & _
        "IF(AND(" & p & "!$H$15>""09:30""," & p & "!$H$15<=""11:30""),""前場""," & _
        "IF(AND(" & p & "!$H$15>=""12:30""," & p & "!$H$15<=""14:30""),""後場""," & _
        "IF(AND(" & p & "!$H$15>""14:30""," & p & "!$H$15<=""15:30""),""引け前"",""その他"")))))"

    ws.Range("Z" & i).Formula = "=IF(OR(" & ge & "," & gn & "),""""," & _
        "IF($N" & i & ">$E" & i & "*" & p & "!$H$4,""VOL-HIGH"",""VOL-NORMAL""))"

    ws.Range("AA" & i).Formula = "=IF(OR(" & ge & "," & gn & "),""""," & _
        "IF(AND($O" & i & "=""VOL-OK"",$Z" & i & "=""VOL-NORMAL""),""OK"",""NG""))"

    ws.Range("AB" & i).Formula = "=IF(OR(" & ge & "," & gn & "," & gj & "," & gk & "),""""," & _
        "IF(AND($F" & i & ">$J" & i & ",$G" & i & ">$K" & i & "),""UP""," & _
        "IF(AND($F" & i & "<$J" & i & ",$G" & i & "<$K" & i & "),""DOWN"",""FLAT"")))"

    ws.Range("AC" & i).Formula = "=IF($C" & i & "="""","""",RssMarket($C" & i & "&"""",""売買代金""))"

    '★列全体参照をやめて有限範囲に
    ws.Range("AD" & i).Formula = "=IFERROR(INDEX(" & HIST & "!$H$2:$H$" & HIST_LAST & _
        ",MATCH($C" & i & "," & HIST & "!$A$2:$A$" & HIST_LAST & ",0)),0)"

    If lastRow > DATA_FIRST Then ws.Range("M" & DATA_FIRST & ":AD" & lastRow).FillDown
    '★余った行に古い数式が残っていたら消す
    If lastRow < DATA_FIRST + STOCK_MAX - 1 Then
        ws.Range("M" & (lastRow + 1) & ":AD" & (DATA_FIRST + STOCK_MAX - 1)).ClearContents
    End If

    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "判定数式をセットしました（" & DATA_FIRST & "〜" & lastRow & "行）。", vbInformation, "MS2"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "判定数式のセット中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical, "MS2"
End Sub

'============================================================
' A=No / B=株探リンク / C=銘柄コード / D=銘柄名
'  ★StockList の実件数ぶんだけ書き、余りは消す
'============================================================
Sub MS2_Update_StockList_To_DATA()
    Dim wsD As Worksheet, wsL As Worksheet, i As Long, s As Long, n As Long
    Dim q As String, sref As String, nref As String, lastL As Long
    Set wsD = ThisWorkbook.Sheets("DATA_MS2")
    Set wsL = ThisWorkbook.Sheets("StockList_MS2")
    q = Chr(34)

    n = 0
    lastL = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row
    For s = 2 To lastL
        If Len(Trim(CStr(wsL.Cells(s, "A").Value))) > 0 And _
           Left$(Trim(CStr(wsL.Cells(s, "A").Value)), 1) <> "※" Then n = s - 1
    Next s
    If n <= 0 Then
        MsgBox "StockList_MS2 に銘柄コードがありません。", vbExclamation, "MS2"
        Exit Sub
    End If
    If n > STOCK_MAX Then n = STOCK_MAX

    Application.ScreenUpdating = False
    For i = DATA_FIRST To DATA_FIRST + n - 1
        s = i - 1
        sref = "'StockList_MS2'!A" & s
        nref = "'StockList_MS2'!B" & s
        wsD.Cells(i, "A").Formula = "=IF(C" & i & "=" & q & q & "," & q & q & ",ROW()-2)"
        wsD.Cells(i, "B").Formula = _
            "=IF(" & sref & "=" & q & q & "," & q & q & ",HYPERLINK(" & q & _
            "https://kabutan.jp/stock/?code=" & q & "&" & sref & "," & q & "株探" & q & "))"
        wsD.Cells(i, "C").Formula = "=IF(" & sref & "=" & q & q & "," & q & q & "," & sref & ")"
        wsD.Cells(i, "D").Formula = "=IF(" & sref & "=" & q & q & "," & q & q & "," & nref & ")"
    Next i
    '★余りは A〜AD ごと消す（古い銘柄のRSS式や前日値が残らないように）
    If DATA_FIRST + n <= DATA_FIRST + STOCK_MAX - 1 Then
        wsD.Range("A" & (DATA_FIRST + n) & ":AD" & (DATA_FIRST + STOCK_MAX - 1)).ClearContents
    End If
    Application.ScreenUpdating = True

    MsgBox "A:No / B:株探 / C:コード / D:銘柄名 を再設定しました（" & n & "銘柄）。" & vbCrLf & vbCrLf & _
           "★このあと必ず「RSS式セット_MS2」→「判定式セット_MS2」→「前日値反映_MS2」を" & vbCrLf & _
           "　この順で実行してください（E〜L列は銘柄コード直書きのため、単独では追従しません）。", _
           vbInformation, "MS2"
End Sub

'============================================================
' RssMarket 式（E〜L）
'============================================================
Sub MS2_Set_RssMarket_Formulas()
    Dim ws As Worksheet, lastRow As Long, i As Long, code As String, n As Long
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    lastRow = DataLastRow(ws)
    If lastRow < DATA_FIRST Then
        MsgBox "DATA_MS2 に銘柄がありません。", vbExclamation, "MS2"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    n = 0
    For i = DATA_FIRST To lastRow
        code = Trim(CStr(ws.Cells(i, "C").Value))
        If code <> "" Then
            ws.Cells(i, "E").Formula = "=RssMarket(""" & code & """,""現在値"")"
            ws.Cells(i, "F").Formula = "=RssMarket(""" & code & """,""高値"")"
            ws.Cells(i, "G").Formula = "=RssMarket(""" & code & """,""安値"")"
            ws.Cells(i, "H").Formula = "=RssMarket(""" & code & """,""終値"")"
            ws.Cells(i, "I").Formula = "=RssMarket(""" & code & """,""出来高"")"
            ws.Cells(i, "J").ClearContents
            ws.Cells(i, "K").ClearContents
            ws.Cells(i, "L").Formula = "=RssMarket(""" & code & """,""前日終値"")"
            n = n + 1
        End If
    Next i
    '★余った行の RSS 式を消す（購読数を減らす）
    If lastRow < DATA_FIRST + STOCK_MAX - 1 Then
        ws.Range("E" & (lastRow + 1) & ":L" & (DATA_FIRST + STOCK_MAX - 1)).ClearContents
    End If
    Application.ScreenUpdating = True
    MsgBox "RSS式をセットしました（" & n & "銘柄）。" & vbCrLf & _
           "J・K列（前日高安）は消えているので「前日値反映_MS2」を実行してください。", vbInformation, "MS2"
End Sub

'============================================================
' ATR履歴シート（9列レイアウト）
'============================================================
Private Function MS2_Get_History() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HIST)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = HIST
    End If
    If CStr(ws.Range("B1").Value) <> "バー高値" Then
        ws.Cells.ClearContents
        ws.Range("A1:I1").Value = Array("銘柄コード", "バー高値", "バー安値", "バー終値", _
                                        "前バー終値", "本数", "TR累計", "ATR", "最終更新")
    End If
    ws.Visible = xlSheetHidden
    Set MS2_Get_History = ws
End Function

Sub MS2_Reset_ATR_History()
    Dim wsH As Worksheet
    Set wsH = MS2_Get_History()
    wsH.Range("A2:I" & HIST_LAST).ClearContents
    gBucket = -1
    MsgBox "ATR履歴をリセットしました。次のサンプリングから再蓄積します。" & vbCrLf & _
           "有効値（" & ATR_N & "本）まで約 " & ATR_N * BarMinutes() & " 分かかります。", vbInformation, "MS2"
End Sub

'============================================================
' ★5分足ATR(14) 収集
'============================================================
Public Sub MS2_Ensure_Sampler()
    If gSampling Then Exit Sub
    MS2_Get_History
    gBarMin = BarMinutes()          '★足の長さを操作パネルから読む
    gBucket = -1
    gSampling = True
    MS2_Schedule_Tick
End Sub

Sub MS2_Start_ATR_Sampler()
    If gSampling And gTickTime > 0 Then
        MsgBox "ATR収集はすでに稼働中です（次回 " & Format(gTickTime, "hh:nn:ss") & "）。", vbInformation, "MS2"
        Exit Sub
    End If
    gSampling = False               '★取り残しフラグをいったん落としてから開始
    MS2_Ensure_Sampler

    '★予約が実際に入ったかで判定する（gSampling だけ見ると嘘の成功になる）
    If gTickTime <= 0 Then
        If gShuttingDown Then
            gSampling = False
            MsgBox "終了処理フラグが立ったままのため開始できませんでした。" & vbCrLf & vbCrLf & _
                   "「★タイマー全解除」を一度押してから、もう一度実行してください。" & vbCrLf & _
                   "（ブックを閉じる操作をキャンセルすると、この状態になることがあります）", _
                   vbExclamation, "MS2"
        Else
            MsgBox "現在は場外（平日 9:00-15:30 以外）のため、ATR収集は開始しません。" & vbCrLf & _
                   "取引時間中に再度実行してください。", vbExclamation, "MS2"
        End If
        Exit Sub
    End If

    MsgBox "ATR収集を開始しました（" & TICK_SEC & "秒間隔で" & gBarMin & "分足を生成）。" & vbCrLf & _
           "実測ATRが" & ATR_N & "本たまるまでは推定値が使われます（約" & ATR_N * gBarMin & "分）。", _
           vbInformation, "MS2"
End Sub

'★予約時刻をセルからも復元して確実に解除
Public Sub MS2_Cancel_Tick()
    Dim t As Date
    gSampling = False
    On Error Resume Next
    If gTickTime > 0 Then Application.OnTime gTickTime, "MS2_ATR_Tick", , False
    t = GetTimerCell(CELL_TICK)
    If t > 0 Then Application.OnTime t, "MS2_ATR_Tick", , False
    On Error GoTo 0
    gTickTime = 0
    PutTimerCell CELL_TICK, 0
End Sub

Sub MS2_Stop_ATR_Sampler()
    MS2_Cancel_Tick
    MsgBox "ATR収集を停止しました。", vbInformation, "MS2"
End Sub

'★終了処理中・場外は再予約しない
Private Sub MS2_Schedule_Tick()
    If gShuttingDown Then Exit Sub
    If Not gSampling Then Exit Sub
    If Not InSession(Now) Then
        gSampling = False
        gTickTime = 0
        PutTimerCell CELL_TICK, 0
        Exit Sub
    End If
    gTickTime = Now + TimeSerial(0, 0, TICK_SEC)
    PutTimerCell CELL_TICK, gTickTime
    On Error Resume Next
    Application.OnTime gTickTime, "MS2_ATR_Tick"
    On Error GoTo 0
End Sub

Public Sub MS2_ATR_Tick()
    Dim ws As Worksheet, wsH As Worksheet
    Dim lastRow As Long, hLast As Long, i As Long, r As Long
    Dim codes As Variant, prices As Variant, h As Variant
    Dim dict As Object, code As String, p As Double
    Dim bk As Long, newBar As Boolean
    Dim hi As Double, lo As Double, cl As Double, pc As Double
    Dim TR As Double, cnt As Long, sumTR As Double, atr As Double
    Dim needed As Long, arr() As Variant, nRows As Long, nCode As Long
    Dim sub_() As Variant, k As Long, cc As Long
    Dim c1(1 To 1, 1 To 1) As Variant, p1(1 To 1, 1 To 1) As Variant
    If gShuttingDown Then Exit Sub
    If Not gSampling Then Exit Sub
    On Error GoTo Bye
    If Not InSession(Now) Then GoTo Bye
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    Set wsH = MS2_Get_History()
    lastRow = DataLastRow(ws)
    If lastRow < DATA_FIRST Then GoTo Bye
    If gBarMin < 1 Then gBarMin = BarMinutes()

    bk = Int(SessionMinutes(Now) / gBarMin)
    newBar = (gBucket >= 0 And bk <> gBucket)

    codes = ws.Range("C" & DATA_FIRST & ":C" & lastRow).Value
    prices = ws.Range("E" & DATA_FIRST & ":E" & lastRow).Value
    If Not IsArray(codes) Then                      '1銘柄だけのとき
        c1(1, 1) = codes: p1(1, 1) = prices
        codes = c1: prices = p1
    End If
    nCode = UBound(codes, 1)

    hLast = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    nRows = 0
    If hLast >= 2 Then
        h = wsH.Range("A2:I" & hLast + 1).Value
        nRows = hLast - 1
        For r = 1 To nRows
            code = Trim(CStr(h(r, 1)))
            If code <> "" Then dict(code) = r
        Next r
    End If
    ReDim arr(1 To nRows + nCode + 1, 1 To 9)
    For r = 1 To nRows
        For i = 1 To 9
            arr(r, i) = h(r, i)
        Next i
    Next r
    needed = nRows
    For i = 1 To nCode
        code = Trim(CStr(codes(i, 1)))
        If code <> "" And Not dict.Exists(code) Then
            needed = needed + 1
            arr(needed, 1) = codes(i, 1)
            arr(needed, 6) = 0
            arr(needed, 7) = 0
            arr(needed, 8) = 0
            dict(code) = needed
        End If
    Next i
    For i = 1 To nCode
        code = Trim(CStr(codes(i, 1)))
        If code = "" Then GoTo NextI
        p = NzD(prices(i, 1))
        If p <= 0 Then GoTo NextI
        r = CLng(dict(code))
        If newBar Then
            hi = NzD(arr(r, 2)): lo = NzD(arr(r, 3))
            cl = NzD(arr(r, 4)): pc = NzD(arr(r, 5))
            If hi > 0 And lo > 0 Then
                If pc > 0 Then
                    TR = WorksheetFunction.Max(hi - lo, Abs(hi - pc), Abs(lo - pc))
                Else
                    TR = hi - lo
                End If
                cnt = CLng(NzD(arr(r, 6)))
                If cnt < ATR_N Then
                    cnt = cnt + 1
                    sumTR = NzD(arr(r, 7)) + TR
                    arr(r, 6) = cnt
                    arr(r, 7) = sumTR
                    atr = SafeDiv(sumTR, CDbl(cnt))
                Else
                    atr = SafeDiv(NzD(arr(r, 8)) * (ATR_N - 1) + TR, CDbl(ATR_N))
                End If
                arr(r, 8) = atr
                arr(r, 5) = cl
            End If
            arr(r, 2) = p: arr(r, 3) = p: arr(r, 4) = p
            arr(r, 9) = Now                 '★最終更新は足が切り替わったときだけ更新
        Else
            If NzD(arr(r, 2)) = 0 Then
                arr(r, 2) = p: arr(r, 3) = p
                arr(r, 9) = Now             '★新規銘柄の初回だけ時刻を入れる
            Else
                If p > NzD(arr(r, 2)) Then arr(r, 2) = p
                If p < NzD(arr(r, 3)) Then arr(r, 3) = p
            End If
            arr(r, 4) = p
        End If
NextI:
    Next i
    If needed > 0 Then
        If newBar Then
            wsH.Range("A2").Resize(needed, 9).Value = arr        '最終更新(I列)も含めて書く
        Else
            '★通常ティックは A〜H だけ書き戻す（I列を触らないので再計算の連鎖が減る）
            ReDim sub_(1 To needed, 1 To 8)
            For k = 1 To needed
                For cc = 1 To 8
                    sub_(k, cc) = arr(k, cc)
                Next cc
            Next k
            wsH.Range("A2").Resize(needed, 8).Value = sub_
        End If
    End If
    gBucket = bk
Bye:
    On Error Resume Next
    MS2_Schedule_Tick
End Sub

'============================================================
' ★タイマー強制全解除（緊急用 / 総当たり）
'============================================================
Public Sub MS2_Force_Cancel_All()
    Dim i As Long, base As Date
    gShuttingDown = True
    gSampling = False
    On Error Resume Next
    '直近3分ぶんの毎秒を総当たりで解除（予約が無い時刻はエラー→無視）
    base = Now - TimeSerial(0, 0, 10)
    For i = 0 To 190
        Application.OnTime base + TimeSerial(0, 0, i), "MS2_ATR_Tick", , False
    Next i
    '当日と翌日の 15:45 も解除
    Application.OnTime Int(Now) + TimeValue(SAVE_HHMM), "MS2_AutoSave_PrevDay", , False
    Application.OnTime Int(Now) + 1 + TimeValue(SAVE_HHMM), "MS2_AutoSave_PrevDay", , False
    If gTickTime > 0 Then Application.OnTime gTickTime, "MS2_ATR_Tick", , False
    If gPrevDaySaveTime > 0 Then Application.OnTime gPrevDaySaveTime, "MS2_AutoSave_PrevDay", , False
    Dim t As Date
    t = GetTimerCell(CELL_TICK)
    If t > 0 Then Application.OnTime t, "MS2_ATR_Tick", , False
    t = GetTimerCell(CELL_SAVE)
    If t > 0 Then Application.OnTime t, "MS2_AutoSave_PrevDay", , False
    On Error GoTo 0
    gTickTime = 0
    gPrevDaySaveTime = 0
    PutTimerCell CELL_TICK, 0
    PutTimerCell CELL_SAVE, 0
    gShuttingDown = False
    MsgBox "すべてのタイマー予約を強制解除しました。" & vbCrLf & _
           "この後ブックを保存して閉じれば、再オープンは起きません。", vbInformation, "MS2"
End Sub

'============================================================
' ★シグナル抽出
'============================================================
Sub MS2_Extract_Signals()
    Dim ws As Worksheet, wsT As Worksheet
    Dim lastRow As Long, tLast As Long, i As Long, j As Long
    Dim code As String, kind As String
    Dim openKey As Object, added As Long, entry As Double
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    Set wsT = ThisWorkbook.Sheets("TRADE_MS2")
    Application.CalculateFull
    lastRow = DataLastRow(ws)
    tLast = wsT.Cells(wsT.Rows.Count, "B").End(xlUp).Row
    If tLast < 1 Then tLast = 1
    Set openKey = CreateObject("Scripting.Dictionary")
    For j = 2 To tLast
        code = Trim(CStr(wsT.Cells(j, "B").Value))
        If code <> "" Then
            If Len(Trim(CStr(wsT.Cells(j, "J").Value))) = 0 Then
                openKey(code) = True
            ElseIf IsDate(wsT.Cells(j, "L").Value) Then
                If Int(CDate(wsT.Cells(j, "L").Value)) = Int(Date) Then openKey(code) = True
            End If
        End If
    Next j
    added = 0
    For i = DATA_FIRST To lastRow
        code = Trim(CStr(ws.Cells(i, "C").Value))
        kind = Trim(CStr(ws.Cells(i, "R").Value))
        If code = "" Or (kind <> "BUY" And kind <> "SELL") Then GoTo NextI
        If openKey.Exists(code) Then GoTo NextI
        entry = NzD(ws.Cells(i, "S").Value)
        If entry <= 0 Then GoTo NextI              '★Entry=0 の玉を作らない
        tLast = tLast + 1
        If tLast > TRADE_LAST Then
            MsgBox "TRADE_MS2 が " & TRADE_LAST & " 行に達しました。" & vbCrLf & _
                   "古い記録を別シートへ退避してから続行してください。", vbExclamation, "MS2"
            Exit For
        End If
        wsT.Cells(tLast, "A").Value = tLast - 1
        wsT.Cells(tLast, "B").Value = ws.Cells(i, "C").Value
        wsT.Cells(tLast, "C").Formula = NameFormula(tLast)
        wsT.Cells(tLast, "D").Value = kind
        wsT.Cells(tLast, "E").Value = entry                           ' Entry
        wsT.Cells(tLast, "F").Value = NzD(ws.Cells(i, "T").Value)     ' SL
        wsT.Cells(tLast, "G").Value = NzD(ws.Cells(i, "U").Value)     ' TP1
        wsT.Cells(tLast, "H").Value = NzD(ws.Cells(i, "V").Value)     ' TP2
        wsT.Cells(tLast, "I").Value = NzD(ws.Cells(i, "W").Value)     ' TP3
        wsT.Cells(tLast, "J").ClearContents                           ' 結果
        wsT.Cells(tLast, "K").ClearContents                           ' 損益
        wsT.Cells(tLast, "L").Value = Now                             ' 開始時刻
        wsT.Cells(tLast, "M").ClearContents                           ' 終了時刻
        PutTradeAux wsT, tLast                                        '★補助列 N/O/P/Q
        openKey(code) = True
        added = added + 1
NextI:
    Next i
    MsgBox "新規シグナル " & added & " 件を TRADE_MS2 に追加しました。", vbInformation, "MS2"
End Sub

'============================================================
' ★結果判定
'   前営業日以前の建玉は当日値で SL/TP 判定せず EXPIRED で閉じる
'============================================================
Sub MS2_Eval_Trades()
    Dim ws As Worksheet, wsT As Worksheet
    Dim lastRow As Long, tLast As Long, i As Long, j As Long
    Dim px As Object, code As String, kind As String
    Dim entry As Double, sl As Double, tp1 As Double, cur As Double
    Dim closed As Long, expired As Long, forceEnd As Boolean
    Dim startDt As Date, isOld As Boolean
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    Set wsT = ThisWorkbook.Sheets("TRADE_MS2")
    Set px = CreateObject("Scripting.Dictionary")
    lastRow = DataLastRow(ws)
    For i = DATA_FIRST To lastRow
        code = Trim(CStr(ws.Cells(i, "C").Value))
        If code <> "" Then px(code) = NzD(ws.Cells(i, "E").Value)
    Next i
    forceEnd = Not InSession(Now)
    tLast = wsT.Cells(wsT.Rows.Count, "B").End(xlUp).Row
    closed = 0: expired = 0
    For j = 2 To tLast
        If Len(Trim(CStr(wsT.Cells(j, "J").Value))) > 0 Then GoTo NextJ
        code = Trim(CStr(wsT.Cells(j, "B").Value))
        kind = Trim(CStr(wsT.Cells(j, "D").Value))
        If code = "" Then GoTo NextJ

        '★建玉日が今日より前なら、当日値では判定しない
        isOld = False
        If IsDate(wsT.Cells(j, "L").Value) Then
            startDt = CDate(wsT.Cells(j, "L").Value)
            If Int(startDt) < Int(Date) Then isOld = True
        End If
        If isOld Then
            wsT.Cells(j, "J").Value = "EXPIRED"
            wsT.Cells(j, "K").Value = 0                       '損益は不明なので0（集計から除外される）
            wsT.Cells(j, "M").Value = Int(startDt) + TimeValue("15:30")
            PutTradeAux wsT, j
            expired = expired + 1
            GoTo NextJ
        End If

        If Not px.Exists(code) Then GoTo NextJ
        entry = NzD(wsT.Cells(j, "E").Value)
        sl = NzD(wsT.Cells(j, "F").Value)
        tp1 = NzD(wsT.Cells(j, "G").Value)
        cur = NzD(px(code))
        If cur <= 0 Or entry <= 0 Then GoTo NextJ
        If kind = "BUY" Then
            If sl > 0 And cur <= sl Then
                wsT.Cells(j, "J").Value = "SL": wsT.Cells(j, "K").Value = sl - entry
            ElseIf tp1 > 0 And cur >= tp1 Then
                wsT.Cells(j, "J").Value = "TP1": wsT.Cells(j, "K").Value = tp1 - entry
            ElseIf forceEnd Then
                wsT.Cells(j, "J").Value = "END": wsT.Cells(j, "K").Value = cur - entry
            End If
        ElseIf kind = "SELL" Then
            If sl > 0 And cur >= sl Then
                wsT.Cells(j, "J").Value = "SL": wsT.Cells(j, "K").Value = entry - sl
            ElseIf tp1 > 0 And cur <= tp1 Then
                wsT.Cells(j, "J").Value = "TP1": wsT.Cells(j, "K").Value = entry - tp1
            ElseIf forceEnd Then
                wsT.Cells(j, "J").Value = "END": wsT.Cells(j, "K").Value = entry - cur
            End If
        End If
        If Len(Trim(CStr(wsT.Cells(j, "J").Value))) > 0 Then
            wsT.Cells(j, "M").Value = Now
            PutTradeAux wsT, j
            closed = closed + 1
        End If
NextJ:
    Next j
    MsgBox "決済確定 " & closed & " 件。" & vbCrLf & _
           "期限切れ（前日以前の建玉）" & expired & " 件を EXPIRED で閉じました（成績集計からは除外）。" & vbCrLf & vbCrLf & _
           "※ 引け（15:30〜15:45）に必ず1回実行してください。翌日以降に持ち越すと" & vbCrLf & _
           "　その玉は損益を確定できません。", vbInformation, "MS2"
End Sub

'============================================================
' ★RANK集計（円換算・月次損益の降順ソート付き）
'============================================================
Sub MS2_Update_Ranking()
    Dim wsT As Worksheet, wsR As Worksheet, wsL As Worksheet
    Dim nameDict As Object, rows_ As Object, coll As Object
    Dim lastRow As Long, lr As Long, k As Long, i As Long, c As Long, r As Long
    Dim q As String, cc As String, key As String, hdr As Variant, v As Variant
    Dim pl As Double, cum As Double, peak As Double, dd As Double
    Dim win As Long, lose As Long, sw As Long, sl_ As Long, mw As Long, ml As Long
    Dim gp As Double, gls As Double, mpl As Double, n As Long
    Dim endRow As Long, unit As Double, res As String
    Set wsT = ThisWorkbook.Sheets("TRADE_MS2")
    Set wsR = ThisWorkbook.Sheets("RANK_MS2")
    Set wsL = ThisWorkbook.Sheets("StockList_MS2")
    q = Chr(34)
    unit = PanelNum("H13", 100)                       '★売買単位（株）で円換算

    hdr = Array("銘柄コード（株探）", "銘柄名", "勝率", "PF", "平均損益(円)", "最大DD(円)", _
                "平均RR", "最大連勝", "最大連敗", "月次損益(円)", "トレード数")
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
        If cc <> "" And Left$(cc, 1) <> "※" And Not nameDict.Exists(cc) Then
            nameDict(cc) = CStr(wsL.Cells(k, "B").Value)
        End If
    Next k

    Set rows_ = CreateObject("Scripting.Dictionary")
    lastRow = wsT.Cells(wsT.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        key = Trim(CStr(wsT.Cells(i, "B").Value))
        If key = "" Then GoTo NextI
        res = Trim(CStr(wsT.Cells(i, "J").Value))
        If Len(res) = 0 Then GoTo NextI
        If res = "EXPIRED" Then GoTo NextI            '★損益不明の玉は集計しない
        '★日をまたいだ記録も集計しない（当日値で誤判定された古いデータ対策）
        If IsDate(wsT.Cells(i, "L").Value) And IsDate(wsT.Cells(i, "M").Value) Then
            If Int(CDate(wsT.Cells(i, "L").Value)) <> Int(CDate(wsT.Cells(i, "M").Value)) Then GoTo NextI
        Else
            GoTo NextI
        End If
        If Not rows_.Exists(key) Then
            Set coll = New Collection
            rows_.Add key, coll
        End If
        rows_(key).Add i
NextI:
    Next i

    r = 2
    For Each v In rows_.Keys
        key = CStr(v)
        Set coll = rows_(key)
        win = 0: lose = 0: gp = 0: gls = 0: mpl = 0: n = 0
        cum = 0: peak = 0: dd = 0: sw = 0: sl_ = 0: mw = 0: ml = 0
        For k = 1 To coll.Count
            i = CLng(coll(k))
            pl = NzD(wsT.Cells(i, "K").Value) * unit      '★円に換算
            n = n + 1
            cum = cum + pl
            If cum > peak Then peak = cum
            If peak - cum > dd Then dd = peak - cum
            If pl > 0 Then
                win = win + 1: gp = gp + pl
                sw = sw + 1: sl_ = 0
                If sw > mw Then mw = sw
            ElseIf pl < 0 Then
                lose = lose + 1: gls = gls + Abs(pl)
                sl_ = sl_ + 1: sw = 0
                If sl_ > ml Then ml = sl_
            End If
            If IsDate(wsT.Cells(i, "L").Value) Then
                If Month(wsT.Cells(i, "L").Value) = Month(Date) And _
                   Year(wsT.Cells(i, "L").Value) = Year(Date) Then mpl = mpl + pl
            End If
        Next k
        wsR.Cells(r, "A").Formula = "=HYPERLINK(" & q & "https://kabutan.jp/stock/?code=" & q & "&" & q & key & q & "," & q & key & q & ")"
        If nameDict.Exists(key) Then
            wsR.Cells(r, "B").Value = nameDict(key)
        Else
            wsR.Cells(r, "B").Value = ""
        End If
        If win + lose > 0 Then
            wsR.Cells(r, "C").Value = SafeDiv(CDbl(win), CDbl(win + lose))
        Else
            wsR.Cells(r, "C").Value = 0
        End If
        '★"∞" の文字列をやめ、999 を「∞」と表示する書式にする（数値のまま並べ替えできる）
        If gls > 0 Then
            wsR.Cells(r, "D").Value = SafeDiv(gp, gls)
        ElseIf gp > 0 Then
            wsR.Cells(r, "D").Value = PF_INF
        Else
            wsR.Cells(r, "D").Value = 0
        End If
        If n > 0 Then
            wsR.Cells(r, "E").Value = SafeDiv(gp - gls, CDbl(n))
        Else
            wsR.Cells(r, "E").Value = 0
        End If
        wsR.Cells(r, "F").Value = dd
        If win > 0 And lose > 0 And gls > 0 Then
            wsR.Cells(r, "G").Value = SafeDiv(SafeDiv(gp, CDbl(win)), SafeDiv(gls, CDbl(lose)))
        ElseIf win > 0 And gp > 0 Then
            wsR.Cells(r, "G").Value = PF_INF
        Else
            wsR.Cells(r, "G").Value = 0
        End If
        wsR.Cells(r, "H").Value = mw
        wsR.Cells(r, "I").Value = ml
        wsR.Cells(r, "J").Value = mpl
        wsR.Cells(r, "K").Value = n
        r = r + 1
    Next v

    endRow = WorksheetFunction.Max(2, r - 1)
    wsR.Range("C2:C" & endRow).NumberFormat = "0.0%"
    wsR.Range("D2:D" & endRow).NumberFormat = "[=" & PF_INF & "]""∞"";0.00"
    wsR.Range("E2:F" & endRow).NumberFormat = "#,##0;(#,##0);-"
    wsR.Range("G2:G" & endRow).NumberFormat = "[=" & PF_INF & "]""∞"";0.00"
    wsR.Range("H2:I" & endRow).NumberFormat = "0"
    wsR.Range("J2:J" & endRow).NumberFormat = "#,##0;(#,##0);-"
    wsR.Range("K2:K" & endRow).NumberFormat = "0"
    wsR.Range("B2:B" & endRow).Font.Color = RGB(0, 0, 0)
    wsR.Range("B2:B" & endRow).Font.Underline = xlUnderlineStyleNone

    '★「ランキング」なので月次損益の降順に並べ替える
    If endRow > 2 Then
        On Error Resume Next
        wsR.Sort.SortFields.Clear
        wsR.Sort.SortFields.Add Key:=wsR.Range("J2:J" & endRow), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With wsR.Sort
            .SetRange wsR.Range("A2:K" & endRow)
            .Header = xlNo
            .Apply
        End With
        wsR.Sort.SortFields.Clear
        On Error GoTo 0
    End If

    MsgBox "RANK_MS2 を集計しました（" & (endRow - 1) & "銘柄 / 月次損益の降順）。" & vbCrLf & _
           "※ 損益は売買単位 " & Format(unit, "#,##0") & " 株での円換算です。" & vbCrLf & _
           "※ EXPIRED と日跨ぎの記録は集計から除外しています。", vbInformation, "MS2"
End Sub

'============================================================
' ★TRADE_MS2 の銘柄名・補助列を一括補完
'============================================================
Sub MS2_Fill_Trade_Names()
    Dim wsT As Worksheet, tLast As Long, j As Long
    Set wsT = ThisWorkbook.Sheets("TRADE_MS2")
    wsT.Range("C1").Value = "銘柄名"
    tLast = wsT.Cells(wsT.Rows.Count, "B").End(xlUp).Row
    Application.ScreenUpdating = False
    For j = 2 To tLast
        If Len(Trim(CStr(wsT.Cells(j, "B").Value))) > 0 Then
            wsT.Cells(j, "C").Formula = NameFormula(j)
            PutTradeAux wsT, j
        End If
    Next j
    Application.ScreenUpdating = True
    MsgBox "TRADE_MS2 の銘柄名と補助列(N/O/P/Q)を補完しました（2〜" & tLast & "行）。", vbInformation, "MS2"
End Sub

'============================================================
' 前日高値・安値ストア
'============================================================
Private Function MS2_Get_PrevDay() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PREV)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = PREV
        ws.Range("A1:D1").Value = Array("銘柄コード", "前日高値", "前日安値", "保存日時")
        ws.Visible = xlSheetHidden
    End If
    If Len(Trim(CStr(ws.Range("Y1").Value))) = 0 Then
        ws.Range("Y1").Value = "次回tick予約(内部)"
        ws.Range("Y2").Value = "次回自動保存予約(内部)"
        ws.Range("Y3").Value = "最終保存日(内部)"
    End If
    Set MS2_Get_PrevDay = ws
End Function

Public Function MS2_Save_PrevDay_Core(Optional ByVal force As Boolean = False) As Long
    Dim wsD As Worksheet, wsP As Worksheet
    Dim i As Long, lastRow As Long, pr As Long, hLast As Long, cnt As Long
    Dim code As String, hi As Double, lo As Double
    Dim dict As Object
    If (Not force) And InSession(Now) Then
        MS2_Save_PrevDay_Core = -1
        Exit Function
    End If
    Set wsD = ThisWorkbook.Sheets("DATA_MS2")
    Set wsP = MS2_Get_PrevDay()
    Set dict = CreateObject("Scripting.Dictionary")
    hLast = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    For pr = 2 To hLast
        code = Trim(CStr(wsP.Cells(pr, "A").Value))
        If code <> "" Then dict(code) = pr
    Next pr
    lastRow = DataLastRow(wsD)
    cnt = 0
    For i = DATA_FIRST To lastRow
        code = Trim(CStr(wsD.Cells(i, "C").Value))
        If code = "" Then GoTo NextI
        hi = NzD(wsD.Cells(i, "F").Value)
        lo = NzD(wsD.Cells(i, "G").Value)
        If hi <= 0 Or lo <= 0 Or lo > hi Then GoTo NextI
        If dict.Exists(code) Then
            pr = CLng(dict(code))
        Else
            pr = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row + 1
            If pr < 2 Then pr = 2
            wsP.Cells(pr, "A").Value = wsD.Cells(i, "C").Value
            dict(code) = pr
        End If
        wsP.Cells(pr, "B").Value = hi
        wsP.Cells(pr, "C").Value = lo
        wsP.Cells(pr, "D").Value = Now
        cnt = cnt + 1
NextI:
    Next i
    If cnt > 0 Then wsP.Range(CELL_SAVED).Value = CDbl(Int(Date))
    MS2_Save_PrevDay_Core = cnt
End Function

Sub MS2_Save_PrevDay()
    Dim n As Long, wsD As Worksheet, total As Long
    Set wsD = ThisWorkbook.Sheets("DATA_MS2")
    total = DataLastRow(wsD) - DATA_FIRST + 1
    n = MS2_Save_PrevDay_Core()
    If n < 0 Then
        If MsgBox("現在は場中（9:00-15:30）です。当日高安が未確定のまま保存すると" & vbCrLf & _
                  "翌日の前日高値・安値が誤った値になります。" & vbCrLf & vbCrLf & _
                  "それでも保存しますか？", vbYesNo + vbExclamation, "MS2") = vbYes Then
            n = MS2_Save_PrevDay_Core(True)
        Else
            Exit Sub
        End If
    End If
    If n = 0 Then
        MsgBox "保存できる高値・安値がありませんでした（RSS未接続、または休場日）。" & vbCrLf & _
               "前日値は変更していません。", vbExclamation, "MS2"
        Exit Sub
    End If
    MsgBox "本日の高値・安値を『前日値』として保存しました（" & n & " / " & total & " 銘柄）。" & _
           IIf(n < total, vbCrLf & "※ " & (total - n) & " 銘柄は高安が取得できずスキップしました。", ""), _
           vbInformation, "MS2"
End Sub

Sub MS2_Load_PrevDay()
    Dim wsD As Worksheet, wsP As Worksheet
    Dim i As Long, lastRow As Long, pr As Long, hLast As Long, n As Long
    Dim code As String, dict As Object
    Set wsD = ThisWorkbook.Sheets("DATA_MS2")
    Set wsP = MS2_Get_PrevDay()
    Set dict = CreateObject("Scripting.Dictionary")
    hLast = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    For pr = 2 To hLast
        code = Trim(CStr(wsP.Cells(pr, "A").Value))
        If code <> "" Then dict(code) = pr
    Next pr
    lastRow = DataLastRow(wsD)
    For i = DATA_FIRST To lastRow
        code = Trim(CStr(wsD.Cells(i, "C").Value))
        If code = "" Then GoTo NextI
        If dict.Exists(code) Then
            pr = CLng(dict(code))
            wsD.Cells(i, "J").Value = NzD(wsP.Cells(pr, "B").Value)
            wsD.Cells(i, "K").Value = NzD(wsP.Cells(pr, "C").Value)
            n = n + 1
        Else
            wsD.Cells(i, "J").ClearContents      '★古い他銘柄の値を残さない
            wsD.Cells(i, "K").ClearContents
        End If
NextI:
    Next i
End Sub

'============================================================
' 前日値の自動保存（OnTime）
'============================================================
Public Sub MS2_Enable_AutoPrevDay()
    MS2_Load_PrevDay
    MS2_Schedule_AutoSave
    MsgBox "前日値の自動化をONにしました（毎営業日 " & SAVE_HHMM & " 頃に保存）。", vbInformation, "MS2"
End Sub

Public Sub MS2_Disable_AutoPrevDay()
    MS2_Cancel_AutoSave
    MsgBox "前日値の自動保存を停止しました。", vbInformation, "MS2"
End Sub

Public Sub MS2_Schedule_AutoSave()
    If gShuttingDown Then Exit Sub
    MS2_Cancel_AutoSave
    gPrevDaySaveTime = MS2_NextSaveTime()
    PutTimerCell CELL_SAVE, gPrevDaySaveTime
    On Error Resume Next
    Application.OnTime gPrevDaySaveTime, "MS2_AutoSave_PrevDay"
    On Error GoTo 0
End Sub

Public Sub MS2_Cancel_AutoSave()
    Dim t As Date
    On Error Resume Next
    If gPrevDaySaveTime > 0 Then Application.OnTime gPrevDaySaveTime, "MS2_AutoSave_PrevDay", , False
    t = GetTimerCell(CELL_SAVE)
    If t > 0 Then Application.OnTime t, "MS2_AutoSave_PrevDay", , False
    On Error GoTo 0
    gPrevDaySaveTime = 0
    PutTimerCell CELL_SAVE, 0
End Sub

Private Function MS2_NextSaveTime() As Date
    Dim d As Date
    d = Int(Now) + TimeValue(SAVE_HHMM)
    If Now >= d Then d = d + 1
    Do While Weekday(d, vbMonday) >= 6
        d = d + 1
    Loop
    MS2_NextSaveTime = d
End Function

Public Sub MS2_AutoSave_PrevDay()
    Static runningFlag As Boolean
    Dim lastSaved As Double
    If gShuttingDown Then Exit Sub
    If runningFlag Then Exit Sub
    runningFlag = True
    On Error Resume Next
    '★同じ日に二重保存しない（休場日は高安が取れず cnt=0 で自然にスキップされる）
    lastSaved = NzD(MS2_Get_PrevDay.Range(CELL_SAVED).Value)
    If Int(lastSaved) <> Int(CDbl(Date)) Then
        MS2_Save_PrevDay_Core True
    End If
    On Error GoTo 0
    runningFlag = False
    MS2_Schedule_AutoSave
End Sub

'============================================================
' ログ表示
'============================================================
Sub MS2_Show_Log()
    Dim wsT As Worksheet, wsH As Worksheet
    Dim lastRow As Long, i As Long, opn As Long, bars As Long
    Dim hLast As Long, sState As String, tinfo As String, warn As String
    Dim lastUpd As Double
    Set wsT = ThisWorkbook.Sheets("TRADE_MS2")
    lastRow = wsT.Cells(wsT.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If Len(Trim(CStr(wsT.Cells(i, "J").Value))) = 0 Then opn = opn + 1
    Next i
    Set wsH = MS2_Get_History()
    bars = 0
    hLast = wsH.Cells(wsH.Rows.Count, "F").End(xlUp).Row
    If hLast >= 2 Then
        On Error Resume Next
        bars = CLng(WorksheetFunction.Max(wsH.Range("F2:F" & hLast)))
        lastUpd = NzD(WorksheetFunction.Max(wsH.Range("I2:I" & hLast)))
        On Error GoTo 0
    End If
    If gSampling Then sState = "稼働中" Else sState = "停止中"

    tinfo = "tick予約: "
    If GetTimerCell(CELL_TICK) > 0 Then
        tinfo = tinfo & Format(GetTimerCell(CELL_TICK), "yyyy/mm/dd hh:nn:ss")
    Else
        tinfo = tinfo & "なし"
    End If
    tinfo = tinfo & vbCrLf & "自動保存予約: "
    If GetTimerCell(CELL_SAVE) > 0 Then
        tinfo = tinfo & Format(GetTimerCell(CELL_SAVE), "yyyy/mm/dd hh:nn")
    Else
        tinfo = tinfo & "なし"
    End If

    '★状態の食い違いを明示する
    If gSampling And GetTimerCell(CELL_TICK) <= 0 Then
        warn = vbCrLf & vbCrLf & "⚠ 「稼働中」なのに予約がありません。" & vbCrLf & _
               "　「★タイマー全解除」→「ATR収集_開始」で復旧してください。"
    End If
    If lastUpd > 0 Then
        If Int(lastUpd) <> Int(CDbl(Date)) Then
            warn = warn & vbCrLf & vbCrLf & "⚠ ATR履歴が古い（" & Format(CDate(lastUpd), "yyyy/mm/dd hh:nn") & "）。" & vbCrLf & _
                   "　実測ATRは使われず、当日レンジ推定にフォールバックしています。"
        End If
    End If

    MsgBox "トレード総数: " & WorksheetFunction.Max(0, lastRow - 1) & vbCrLf & _
           "うち建玉中: " & opn & vbCrLf & vbCrLf & _
           "ATR収集: " & sState & vbCrLf & _
           "蓄積バー数(最大): " & bars & " / " & ATR_N & vbCrLf & _
           "ATR履歴の最終更新: " & IIf(lastUpd > 0, Format(CDate(lastUpd), "yyyy/mm/dd hh:nn:ss"), "なし") & vbCrLf & vbCrLf & _
           tinfo & warn, vbInformation, "MS2_LOG"
End Sub

'============================================================
' 見出し修復（DATA_MS2：30列）
'============================================================
Sub MS2_Fix_DATA_Header()
    Dim ws As Worksheet, headers As Variant, c As Long, nCol As Long
    Set ws = ThisWorkbook.Sheets("DATA_MS2")
    headers = Array("No", "株探リンク", "銘柄コード", "銘柄名", "現在値", "高値", "安値", "終値", _
        "出来高", "前日高値", "前日安値", "前日終値", "寄り方向", "ATR(5分足14)", "ゾーン", _
        "STOP-BUY", "STOP-SELL", "売買種別", "Entry", "SL", "TP1", "TP2", "TP3", _
        "2回目戻し", "時間帯", "ボラフィルタ", "ダマシ除去", "トレンド方向", "売買代金(千円)", "実測5分ATR")
    nCol = UBound(headers) - LBound(headers) + 1
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
    MsgBox "DATA_MS2 の見出しを修復しました（30列）。", vbInformation, "MS2"
End Sub

'============================================================
' ★見出し修復（TRADE_MS2：17列）＋ 補助列・サマリーの再生成
'============================================================
Sub MS2_Fix_TRADE_Header()
    Dim ws As Worksheet, headers As Variant, c As Long, j As Long, tLast As Long
    Set ws = ThisWorkbook.Sheets("TRADE_MS2")
    headers = Array("No", "銘柄コード", "銘柄名", "売買種別", "Entry", "SL", _
                    "TP1", "TP2", "TP3", "結果", "損益", "開始時刻", "終了時刻", _
                    "約定代金(100株)", "損益(100株)", "R倍数", "同日決済")
    For c = 1 To 17
        With ws.Cells(1, c)
            .Value = headers(c - 1)
            .Font.Name = "Meiryo"
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c

    '補助列を全行に張り直す
    Application.ScreenUpdating = False
    tLast = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If tLast < 2 Then tLast = 2
    For j = 2 To WorksheetFunction.Max(tLast, TRADE_LAST)
        PutTradeAux ws, j
    Next j

    ws.Range("U1").Value = "同日損益(内部)"

    'サマリー（★同日決済のみを集計）
    ws.Range("R1").Value = "◆ 100株換算サマリー（同日決済のみ）"
    ws.Range("R2").Value = "決済済み件数"
    ws.Range("R3").Value = "合計損益（円）"
    ws.Range("R4").Value = "平均損益（円）"
    ws.Range("R5").Value = "最大損失（円）"
    ws.Range("R6").Value = "勝率"
    ws.Range("R7").Value = "日跨ぎ除外（無効）件数"
    ws.Range("S2").Formula = "=COUNTIFS($Q$2:$Q$" & TRADE_LAST & ",1)"
    ws.Range("S3").Formula = "=SUMIFS($O$2:$O$" & TRADE_LAST & ",$Q$2:$Q$" & TRADE_LAST & ",1)"
    ws.Range("S4").Formula = "=IF(S2=0,0,ROUND(S3/S2,0))"
    ws.Range("S5").Formula = "=IF(S2=0,0,MIN($U$2:$U$" & TRADE_LAST & "))"
    ws.Range("S6").Formula = "=IF(S2=0,0,COUNTIFS($O$2:$O$" & TRADE_LAST & ","">0"",$Q$2:$Q$" & TRADE_LAST & ",1)/S2)"
    ws.Range("S7").Formula = "=COUNTIFS($Q$2:$Q$" & TRADE_LAST & ",0)"
    ws.Range("S6").NumberFormat = "0.0%"
    Application.ScreenUpdating = True

    MsgBox "TRADE_MS2 の見出し（17列）と補助列・サマリーを再生成しました。", vbInformation, "MS2"
End Sub

'============================================================
' 全自動
'============================================================
Sub MS2_Auto_All()
    MS2_Ensure_Sampler
    MS2_Load_PrevDay
    Application.CalculateFull
    MS2_Extract_Signals
    MS2_Eval_Trades
    MS2_Update_Ranking
End Sub

'============================================================
' 操作パネル
'============================================================
Sub MS2_Build_Menu()
    Dim ws As Worksheet, btn As Object, y As Single, items As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PANEL)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = PANEL
    End If
    ws.Buttons.Delete
    ws.Range("A1").Value = "◆ MS2 操作パネル"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    items = Array( _
        "銘柄反映_MS2", "MS2_Update_StockList_To_DATA", _
        "RSS式セット_MS2", "MS2_Set_RssMarket_Formulas", _
        "判定式セット_MS2", "MS2_Install_Formulas", _
        "ATR収集_開始", "MS2_Start_ATR_Sampler", _
        "ATR収集_停止", "MS2_Stop_ATR_Sampler", _
        "シグナル抽出_MS2", "MS2_Extract_Signals", _
        "結果判定_MS2", "MS2_Eval_Trades", _
        "ランキング_MS2", "MS2_Update_Ranking", _
        "銘柄名補完_MS2", "MS2_Fill_Trade_Names", _
        "ヘッダー修復_MS2", "MS2_Fix_DATA_Header", _
        "TRADE見出し修復", "MS2_Fix_TRADE_Header", _
        "ATR履歴リセット_MS2", "MS2_Reset_ATR_History", _
        "前日値保存_MS2", "MS2_Save_PrevDay", _
        "前日値反映_MS2", "MS2_Load_PrevDay", _
        "前日値_自動ON", "MS2_Enable_AutoPrevDay", _
        "前日値_自動OFF", "MS2_Disable_AutoPrevDay", _
        "★タイマー全解除", "MS2_Force_Cancel_All", _
        "ログ_MS2", "MS2_Show_Log", _
        "全自動_MS2", "MS2_Auto_All")
    y = 36
    For i = LBound(items) To UBound(items) Step 2
        Set btn = ws.Buttons.Add(20, y, 180, 28)
        btn.Caption = items(i)
        btn.OnAction = items(i + 1)
        y = y + 34
    Next i
    MS2_Install_Settings
    MsgBox "操作パネルのボタンを再作成しました（設定ブロックは保持）。", vbInformation, "MS2"
End Sub
