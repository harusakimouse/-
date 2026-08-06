Attribute VB_Name = "modFetch250"
'==============================================================================
' modFetch250 ― OHLCV 5シートへ 250営業日分を取得して書き込む
'
'  データ取込シートに元からある仕組みをそのまま回す。
'      A3 = RssChartPast(B2:J2, B1, "D", $M$1, $K$1)
'        B1  銘柄コード
'        K1  本数（日数）
'        M1  開始日（yyyymmdd）
'
'  やること
'    1. K1=250 / M1=今日 をセット
'    2. 日付軸（row 3）を作る … TOPX を基準にする
'    3. 銘柄管理 B6:B305 を上から順に
'         B1 に銘柄コードを入れる → 再計算 → RSSの応答を待つ
'         → 日付を突き合わせて 始値/高値/安値/終値/出来高 の5シートへ書く
'    4. 途中経過をステータスバーに出す。ESC で中断できる
'
'  ★日付で突き合わせて書く理由
'    V805 の取込は「E列＝最新、F列＝1日前…」と順番に詰めていた。
'    この方式だと売買停止などで1日抜けた銘柄だけ列が1つズレ、
'    銘柄どうしの日付が揃わなくなる。指標はすべて銘柄間の比較で
'    計算するので、ズレると結果が壊れる。
'    そこで row 3 の日付軸に対して MATCH して、同じ日付の列に書く。
'    軸に無い日付は書かない（欠測として空欄のまま）。
'
'  使い方
'    OHLCV_全銘柄取得     … 銘柄管理の全銘柄をまとめて取得（時間がかかる）
'    OHLCV_1銘柄取得      … データ取込!B1 の1銘柄だけ
'    OHLCV_取得状況を確認  … 何銘柄そろっているかを表示
'==============================================================================
Option Explicit

'------------------------------------------------------------ 設定（ここを触る）
Private Const HIST_MAX     As Long = 250     ' 取得する営業日数（E列～IT列）
Private Const WAIT_LIMIT   As Double = 25    ' RSS応答の待ち時間の上限（秒）
Private Const STABLE_POLLS As Long = 3       ' この回数だけ内容が変わらなければ受信完了
Private Const POLL_WAIT    As Double = 0.2   ' 1回の待ち（秒）
Private Const SKIP_IF_HAVE As Long = 0       ' 既にこの日数以上ある銘柄を飛ばす（0=飛ばさない）

Private Const PWD          As String = "ne19480314"
Private Const FIRST_COL    As Long = 5       ' E列
Private Const DATE_ROW     As Long = 3       ' 価格シートの日付ヘッダ行
Private Const TOPIX_ROW    As Long = 5       ' 価格シートの TOPX 行
Private Const STOCK_ROW1   As Long = 6       ' 価格シートの1銘柄目
Private Const MEI_ROW1     As Long = 6       ' 銘柄管理の1銘柄目
Private Const MEI_ROWN     As Long = 305
Private Const SH_DATA      As String = "データ取込"
Private Const SH_MEI       As String = "銘柄管理"

Private Const SHEETS_LIST  As String = "始値,高値,安値,終値,出来高"

' データ取込シートの列位置（FindHeader が埋める）
Private hRow As Long, colName As Long, colDate As Long
Private colOpen As Long, colHigh As Long, colLow As Long, colClose As Long, colVol As Long

Private gAxis() As Date          ' 日付軸（新しい順）
Private gAxisN  As Long
Private gCancel As Boolean


'==============================================================================
' ① 全銘柄を取得して5シートへ書き込む
'==============================================================================
Public Sub OHLCV_全銘柄取得()
    Dim prevCalc As XlCalculation, prevScr As Boolean
    prevCalc = Application.Calculation: prevScr = Application.ScreenUpdating
    gCancel = False
    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler      ' ESC を捕まえる
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic  ' RSS は自動計算でないと更新されない

    Dim dws As Worksheet: Set dws = MustSheet(SH_DATA)
    Dim mws As Worksheet: Set mws = MustSheet(SH_MEI)

    If MsgBox("銘柄管理に登録された全銘柄について、" & HIST_MAX & "営業日分を取得します。" & vbCrLf & _
              "銘柄数によっては10～30分かかります。" & vbCrLf & vbCrLf & _
              "途中で止めるときは ESC キーを押してください。" & vbCrLf & _
              "始めますか？", vbYesNo + vbQuestion, "OHLCV 一括取得") <> vbYes Then
        Cleanup prevCalc, prevScr
        Exit Sub
    End If

    PrepareDataSheet dws
    FindHeader dws
    If hRow = 0 Then Err.Raise 5, , "データ取込シートに「日付」「始値」「終値」を含むヘッダ行が見つかりません。"

    '--- 日付軸を作る（TOPX 基準。取れなければ最初に成功した銘柄で作る）---
    Application.StatusBar = "日付軸を作成中…"
    If Not BuildAxisFrom(dws, "TOPX") Then
        Dim r0 As Long
        For r0 = MEI_ROW1 To MEI_ROWN
            If Trim$(CStr(mws.Cells(r0, 2).Value)) <> "" Then
                If BuildAxisFrom(dws, CStr(mws.Cells(r0, 2).Value)) Then Exit For
            End If
        Next r0
    End If
    If gAxisN = 0 Then Err.Raise 5, , "日付軸を作れませんでした。" & vbCrLf & _
                                      "データ取込シートで RSS がデータを返しているか確認してください。"
    WriteAxis
    Application.StatusBar = "日付軸 " & gAxisN & " 営業日分を確定しました。"

    '--- 銘柄をループ ---
    Dim r As Long, nTarget As Long, nDone As Long, nSkip As Long, nFail As Long
    Dim failList As String
    For r = MEI_ROW1 To MEI_ROWN
        If Trim$(CStr(mws.Cells(r, 2).Value)) <> "" Then nTarget = nTarget + 1
    Next r

    Dim idx As Long
    For r = MEI_ROW1 To MEI_ROWN
        Dim code As String: code = Trim$(CStr(mws.Cells(r, 2).Value))
        If code <> "" Then
            idx = idx + 1
            Dim priceRow As Long: priceRow = STOCK_ROW1 + (r - MEI_ROW1)

            If SKIP_IF_HAVE > 0 Then
                If FilledDays(priceRow) >= SKIP_IF_HAVE Then
                    nSkip = nSkip + 1
                    GoTo NextStock
                End If
            End If

            Application.StatusBar = "[" & idx & "/" & nTarget & "] " & code & " " & _
                                    CStr(mws.Cells(r, 3).Value) & " を取得中…" & _
                                    "  (完了 " & nDone & " / 失敗 " & nFail & ")  ESCで中断"
            Dim got As Long
            got = FetchAndWrite(dws, code, priceRow)
            If got > 0 Then
                nDone = nDone + 1
            Else
                nFail = nFail + 1
                If Len(failList) < 400 Then failList = failList & code & " "
            End If
            If gCancel Then Exit For
        End If
NextStock:
    Next r

    Cleanup prevCalc, prevScr
    MsgBox IIf(gCancel, "ESC で中断しました。" & vbCrLf & vbCrLf, "") & _
           "取得完了: " & nDone & " 銘柄" & vbCrLf & _
           "スキップ: " & nSkip & " 銘柄" & vbCrLf & _
           "失敗:     " & nFail & " 銘柄" & vbCrLf & _
           IIf(nFail > 0, vbCrLf & "失敗した銘柄: " & failList & vbCrLf & _
               "（時間をおいて、もう一度実行すると取れることがあります）", "") & vbCrLf & vbCrLf & _
           "日付軸: " & gAxisN & " 営業日分", _
           vbInformation, "OHLCV 一括取得"
    Exit Sub

ErrHandler:
    If Err.Number = 18 Then                 ' ESC が押された
        gCancel = True
        Resume Next
    End If
    Dim msg As String: msg = "Err " & Err.Number & ": " & Err.Description
    Cleanup prevCalc, prevScr
    MsgBox "取得を中断しました。" & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
           "アプリの状態は元に戻しました。", vbCritical, "取得エラー"
End Sub


'==============================================================================
' ② データ取込!B1 の1銘柄だけ取得する
'==============================================================================
Public Sub OHLCV_1銘柄取得()
    Dim prevCalc As XlCalculation, prevScr As Boolean
    prevCalc = Application.Calculation: prevScr = Application.ScreenUpdating
    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic  ' RSS は自動計算でないと更新されない

    Dim dws As Worksheet: Set dws = MustSheet(SH_DATA)
    Dim mws As Worksheet: Set mws = MustSheet(SH_MEI)
    Dim code As String: code = UCase$(Trim$(CStr(dws.Range("B1").Value)))
    If code = "" Then Err.Raise 5, , "データ取込シートの B1 に銘柄コードを入れてください。"

    Dim r As Long, priceRow As Long
    For r = MEI_ROW1 To MEI_ROWN
        If UCase$(Trim$(CStr(mws.Cells(r, 2).Value))) = code Then
            priceRow = STOCK_ROW1 + (r - MEI_ROW1): Exit For
        End If
    Next r
    If priceRow = 0 Then Err.Raise 5, , "銘柄コード " & code & " が銘柄管理にありません。"

    PrepareDataSheet dws
    FindHeader dws
    If hRow = 0 Then Err.Raise 5, , "データ取込シートのヘッダ行が見つかりません。"

    ReadAxis                                  ' 既にある日付軸を使う
    If gAxisN = 0 Then
        If Not BuildAxisFrom(dws, "TOPX") Then
            If Not BuildAxisFrom(dws, code) Then Err.Raise 5, , "日付軸を作れませんでした。"
        End If
        WriteAxis
    End If

    Dim got As Long: got = FetchAndWrite(dws, code, priceRow)
    Cleanup prevCalc, prevScr
    If got > 0 Then
        MsgBox code & " を " & got & " 日分 書き込みました。（日付軸 " & gAxisN & " 日）", _
               vbInformation, "取得完了"
    Else
        MsgBox code & " のデータを取得できませんでした。" & vbCrLf & vbCrLf & _
               "・マーケットスピードが起動してログイン済みか" & vbCrLf & _
               "・データ取込シートの A3 に RssChartPast の数式があるか" & vbCrLf & _
               "を確認してください。", vbExclamation, "取得できず"
    End If
    Exit Sub

ErrHandler:
    If Err.Number = 18 Then Resume Next
    Dim msg As String: msg = "Err " & Err.Number & ": " & Err.Description
    Cleanup prevCalc, prevScr
    MsgBox "取得を中断しました。" & vbCrLf & vbCrLf & msg, vbCritical, "取得エラー"
End Sub


'==============================================================================
' ③ 取得状況の確認
'==============================================================================
Public Sub OHLCV_取得状況を確認()
    Dim mws As Worksheet: Set mws = MustSheet(SH_MEI)
    Dim cws As Worksheet: Set cws = MustSheet("終値")
    Dim r As Long, n As Long, few As Long, none As Long, full As Long
    Dim minD As Long: minD = 99999
    Dim firstShort As String

    For r = MEI_ROW1 To MEI_ROWN
        If Trim$(CStr(mws.Cells(r, 2).Value)) <> "" Then
            n = n + 1
            Dim d As Long: d = FilledDays(STOCK_ROW1 + (r - MEI_ROW1))
            If d = 0 Then
                none = none + 1
            ElseIf d < 100 Then
                few = few + 1
                If firstShort = "" Then firstShort = CStr(mws.Cells(r, 2).Value) & _
                                                     "(" & d & "日)"
            Else
                full = full + 1
            End If
            If d < minD Then minD = d
        End If
    Next r

    Dim axisN As Long
    Dim c As Long
    For c = FIRST_COL To FIRST_COL + HIST_MAX - 1
        If IsDate(cws.Cells(DATE_ROW, c).Value) Then axisN = axisN + 1
    Next c

    MsgBox "登録銘柄:        " & n & vbCrLf & _
           "  100日以上:     " & full & vbCrLf & _
           "  100日未満:     " & few & IIf(firstShort <> "", "  例) " & firstShort, "") & vbCrLf & _
           "  データなし:     " & none & vbCrLf & vbCrLf & _
           "日付軸(3行目):   " & axisN & " 営業日分" & vbCrLf & _
           "最少の銘柄:      " & IIf(minD = 99999, 0, minD) & " 日分" & vbCrLf & vbCrLf & _
           "指標の計算には25日以上、検証には100日以上あると安心です。", _
           vbInformation, "OHLCV 取得状況"
End Sub


'------------------------------------------------------------------ 取得と書込
Private Function FetchAndWrite(ByVal dws As Worksheet, ByVal code As String, _
                               ByVal priceRow As Long) As Long
    Dim d() As Date, o() As Double, h() As Double, l() As Double
    Dim c() As Double, v() As Double, n As Long

    n = FetchSeries(dws, code, d, o, h, l, c, v)
    If n = 0 Then Exit Function

    ' 日付軸に合わせて 1×HIST_MAX の配列を作る（軸に無い日付は書かない）
    Dim ao As Variant, ah As Variant, al As Variant, ac As Variant, av As Variant
    ReDim ao(1 To 1, 1 To HIST_MAX): ReDim ah(1 To 1, 1 To HIST_MAX)
    ReDim al(1 To 1, 1 To HIST_MAX): ReDim ac(1 To 1, 1 To HIST_MAX)
    ReDim av(1 To 1, 1 To HIST_MAX)

    Dim i As Long, k As Long, placed As Long
    For i = 1 To n
        k = AxisIndex(d(i))
        If k > 0 Then
            ao(1, k) = o(i): ah(1, k) = h(i): al(1, k) = l(i)
            ac(1, k) = c(i): av(1, k) = v(i)
            placed = placed + 1
        End If
    Next i
    If placed = 0 Then Exit Function

    PutRow "始値", priceRow, ao
    PutRow "高値", priceRow, ah
    PutRow "安値", priceRow, al
    PutRow "終値", priceRow, ac
    PutRow "出来高", priceRow, av
    FetchAndWrite = placed
End Function


' RSS に銘柄を投げて、返ってきた日足を配列に読み取る
Private Function FetchSeries(ByVal dws As Worksheet, ByVal code As String, _
                             ByRef d() As Date, ByRef o() As Double, ByRef h() As Double, _
                             ByRef l() As Double, ByRef c() As Double, _
                             ByRef v() As Double) As Long
    ReDim d(1 To HIST_MAX): ReDim o(1 To HIST_MAX): ReDim h(1 To HIST_MAX)
    ReDim l(1 To HIST_MAX): ReDim c(1 To HIST_MAX): ReDim v(1 To HIST_MAX)

    dws.Range("B1").Value = code
    Application.Calculate
    If Not WaitForRss(dws, ExpectedName(code)) Then Exit Function

    Dim lastRow As Long, i As Long, n As Long
    lastRow = dws.Cells(dws.Rows.Count, colDate).End(xlUp).Row
    For i = hRow + 1 To lastRow
        If n >= HIST_MAX Then Exit For
        Dim dv As Variant: dv = dws.Cells(i, colDate).Value
        If IsDate(dv) Then
            Dim cv As Variant: cv = dws.Cells(i, colClose).Value
            Dim vv As Variant: vv = IIf(colVol > 0, dws.Cells(i, colVol).Value, 0)
            ' 出来高0の日は休業・欠損として飛ばす
            If IsNumeric(cv) And IsNumeric(vv) Then
                If cv > 0 And vv > 0 Then
                    n = n + 1
                    d(n) = CDate(dv)
                    c(n) = CDbl(cv)
                    v(n) = CDbl(vv)
                    o(n) = PickNum(dws, i, colOpen, c(n))
                    h(n) = PickNum(dws, i, colHigh, c(n))
                    l(n) = PickNum(dws, i, colLow, c(n))
                End If
            End If
        End If
    Next i
    FetchSeries = n
End Function


' RSS は非同期なので、内容が落ち着くまで待つ。
'
' 行数が安定しただけでは足りない。連続する2銘柄がたまたま同じ日数だと、
' 前の銘柄の古いデータをそのまま読んでしまう。そこで
'   ・銘柄名の列が目的の銘柄になっている
'   ・かつ行数が STABLE_POLLS 回続けて同じ
' の両方がそろって初めて受信完了とみなす。
Private Function WaitForRss(ByVal dws As Worksheet, ByVal wantName As String) As Boolean
    Dim t0 As Double: t0 = Timer
    Dim lastN As Long, same As Long
    lastN = -1

    Do
        DoEvents                                  ' RSS に配信の機会を与える
        PauseFor POLL_WAIT

        Dim n As Long
        n = dws.Cells(dws.Rows.Count, colDate).End(xlUp).Row - hRow
        If n < 0 Then n = 0

        Dim nameOK As Boolean
        If wantName = "" Or colName = 0 Then
            nameOK = True                         ' 名前で確認できない場合は行数だけで判断
        Else
            nameOK = (Trim$(CStr(dws.Cells(hRow + 1, colName).Value)) = wantName)
        End If

        If n > 0 And n = lastN And nameOK Then
            same = same + 1
            If same >= STABLE_POLLS Then
                WaitForRss = True
                Exit Function
            End If
        Else
            same = 0
            lastN = n
        End If

        If Timer - t0 > WAIT_LIMIT Then Exit Do   ' 応答なしで打ち切り
        If Timer < t0 Then t0 = Timer             ' 日付をまたいだ場合の保険
    Loop
End Function


' 銘柄管理に登録されている銘柄名を返す（受信完了の判定に使う）
Private Function ExpectedName(ByVal code As String) As String
    Dim mws As Worksheet
    On Error Resume Next
    Set mws = ThisWorkbook.Sheets(SH_MEI)
    On Error GoTo 0
    If mws Is Nothing Then Exit Function
    Dim r As Long
    For r = MEI_ROW1 - 1 To MEI_ROWN            ' TOPX は5行目にあるので1つ上から見る
        If UCase$(Trim$(CStr(mws.Cells(r, 2).Value))) = UCase$(Trim$(code)) Then
            ExpectedName = Trim$(CStr(mws.Cells(r, 3).Value))
            Exit Function
        End If
    Next r
End Function


'------------------------------------------------------------------ 日付軸
' 指定銘柄のデータから日付軸（新しい順）を作る
Private Function BuildAxisFrom(ByVal dws As Worksheet, ByVal code As String) As Boolean
    Dim d() As Date, o() As Double, h() As Double, l() As Double
    Dim c() As Double, v() As Double
    Dim n As Long: n = FetchSeries(dws, code, d, o, h, l, c, v)
    If n = 0 Then Exit Function

    SortDesc d, o, h, l, c, v, n
    ReDim gAxis(1 To n)
    Dim i As Long
    For i = 1 To n
        gAxis(i) = d(i)
    Next i
    gAxisN = n
    BuildAxisFrom = True
End Function


' 既に価格シートにある日付軸を読み込む
Private Sub ReadAxis()
    Dim ws As Worksheet: Set ws = MustSheet("終値")
    ReDim gAxis(1 To HIST_MAX)
    gAxisN = 0
    Dim c As Long
    For c = FIRST_COL To FIRST_COL + HIST_MAX - 1
        If IsDate(ws.Cells(DATE_ROW, c).Value) Then
            gAxisN = gAxisN + 1
            gAxis(gAxisN) = CDate(ws.Cells(DATE_ROW, c).Value)
        Else
            Exit For
        End If
    Next c
End Sub


' 日付軸を5シートの3行目へ書く
Private Sub WriteAxis()
    Dim arr As Variant
    ReDim arr(1 To 1, 1 To HIST_MAX)
    Dim i As Long
    For i = 1 To gAxisN
        arr(1, i) = gAxis(i)
    Next i

    Dim s As Variant
    For Each s In Split(SHEETS_LIST, ",")
        Dim ws As Worksheet: Set ws = MustSheet(CStr(s))
        ws.Unprotect PWD
        With ws.Range(ws.Cells(DATE_ROW, FIRST_COL), _
                      ws.Cells(DATE_ROW, FIRST_COL + HIST_MAX - 1))
            .ClearContents
            .Value = arr
            .NumberFormatLocal = "m/d"
        End With
        ws.Protect PWD, UserInterfaceOnly:=True
    Next s
End Sub


' 日付軸の中での位置（1起点）。無ければ 0
Private Function AxisIndex(ByVal dt As Date) As Long
    Dim lo As Long, hi As Long, midIdx As Long
    lo = 1: hi = gAxisN
    Do While lo <= hi                        ' 軸は新しい順なので降順の二分探索
        midIdx = (lo + hi) \ 2
        If gAxis(midIdx) = dt Then
            AxisIndex = midIdx
            Exit Function
        ElseIf gAxis(midIdx) > dt Then
            lo = midIdx + 1
        Else
            hi = midIdx - 1
        End If
    Loop
End Function


'------------------------------------------------------------------ 書き込み
Private Sub PutRow(ByVal shName As String, ByVal priceRow As Long, ByRef arr As Variant)
    Dim ws As Worksheet: Set ws = MustSheet(shName)
    ws.Unprotect PWD
    With ws.Range(ws.Cells(priceRow, FIRST_COL), _
                  ws.Cells(priceRow, FIRST_COL + HIST_MAX - 1))
        .ClearContents
        .Value = arr                          ' 1行まとめて書く（1セルずつより桁違いに速い）
    End With
    ws.Protect PWD, UserInterfaceOnly:=True
End Sub


Private Function FilledDays(ByVal priceRow As Long) As Long
    Dim ws As Worksheet: Set ws = MustSheet("終値")
    Dim arr As Variant
    arr = ws.Range(ws.Cells(priceRow, FIRST_COL), _
                   ws.Cells(priceRow, FIRST_COL + HIST_MAX - 1)).Value
    Dim i As Long, n As Long
    For i = 1 To HIST_MAX
        If IsNumeric(arr(1, i)) Then
            If arr(1, i) > 0 Then n = n + 1
        End If
    Next i
    FilledDays = n
End Function


'------------------------------------------------------------------ 補助
Private Sub PrepareDataSheet(ByVal dws As Worksheet)
    dws.Range("K1").Value = HIST_MAX                    ' 本数
    dws.Range("M1").Value = Format$(Date, "yyyymmdd")   ' 開始日＝今日
End Sub


Private Sub FindHeader(ByVal dws As Worksheet)
    hRow = 0: colName = 0: colDate = 0
    colOpen = 0: colHigh = 0: colLow = 0: colClose = 0: colVol = 0

    Dim r As Long, c As Long
    For r = 1 To 60
        Dim okD As Boolean, okO As Boolean, okC As Boolean
        okD = False: okO = False: okC = False
        For c = 1 To 30
            Select Case Trim$(CStr(dws.Cells(r, c).Value))
                Case "日付": okD = True
                Case "始値": okO = True
                Case "終値": okC = True
            End Select
        Next c
        If okD And okO And okC Then
            hRow = r
            For c = 1 To 30
                Select Case Trim$(CStr(dws.Cells(r, c).Value))
                    Case "銘柄名称", "銘柄名":                    colName = c
                    Case "日付":                                  colDate = c
                    Case "始値":                                  colOpen = c
                    Case "高値":                                  colHigh = c
                    Case "安値":                                  colLow = c
                    Case "終値":                                  colClose = c
                    Case "出来高", "売買高", "出来高(株)", "出来高（株）": colVol = c
                End Select
            Next c
            Exit For
        End If
    Next r
End Sub


Private Function PickNum(ByVal dws As Worksheet, ByVal r As Long, _
                         ByVal col As Long, ByVal fallback As Double) As Double
    If col = 0 Then PickNum = fallback: Exit Function
    Dim v As Variant: v = dws.Cells(r, col).Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then PickNum = CDbl(v): Exit Function
    End If
    PickNum = fallback                        ' 欠損は終値で代用
End Function


' 日付の新しい順に並べ替える（挿入ソート。250件なので十分速い）
Private Sub SortDesc(ByRef d() As Date, ByRef o() As Double, ByRef h() As Double, _
                     ByRef l() As Double, ByRef c() As Double, ByRef v() As Double, _
                     ByVal n As Long)
    Dim i As Long, j As Long
    For i = 2 To n
        Dim kd As Date, ko As Double, kh As Double
        Dim kl As Double, kc As Double, kv As Double
        kd = d(i): ko = o(i): kh = h(i): kl = l(i): kc = c(i): kv = v(i)
        j = i - 1
        Do While j >= 1
            If d(j) >= kd Then Exit Do
            d(j + 1) = d(j): o(j + 1) = o(j): h(j + 1) = h(j)
            l(j + 1) = l(j): c(j + 1) = c(j): v(j + 1) = v(j)
            j = j - 1
        Loop
        d(j + 1) = kd: o(j + 1) = ko: h(j + 1) = kh
        l(j + 1) = kl: c(j + 1) = kc: v(j + 1) = kv
    Next i
End Sub


' Application.Wait は1秒未満を扱えないので Timer で自前に待つ
Private Sub PauseFor(ByVal sec As Double)
    Dim t As Double: t = Timer
    Do
        DoEvents
        If Timer < t Then Exit Do             ' 日付をまたいだ
    Loop While Timer - t < sec
End Sub


Private Function MustSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set MustSheet = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If MustSheet Is Nothing Then Err.Raise 5, , "「" & nm & "」シートが見つかりません。"
End Function


' 中断しても Calculation が Manual のまま残らないようにする。
' V805 の抽出Sub4本はこれが無く、残ると現在値が更新されずに
' 損切判定が完全に沈黙する状態になっていた。
Private Sub Cleanup(ByVal prevCalc As XlCalculation, ByVal prevScr As Boolean)
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0
End Sub
