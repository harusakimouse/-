Attribute VB_Name = "modImport250"
'==============================================================================
' modImport250 ― 価格履歴の一括取込（最大250営業日）
'
'  V805 の取込は 60日固定（E〜BL列）だった。V808 で価格シートを
'  E列〜IT列（250営業日分）へ拡張したので、その全域に書き込めるようにする。
'
'  データ取込シートの既存の仕組みをそのまま使う
'    A3 = RssChartPast(B2:J2, B1, "D", $M$1, $K$1)
'      B1  銘柄コード     例) 3401
'      K1  本数（日数）   ここを 250 にする
'      M1  開始日        yyyymmdd 形式  例) 20260806
'    → K1 と M1 を入れて再計算すると、3行目以降に日足が降ってくる。
'
'  このモジュールは「降ってきた日足を5枚の価格シートへ転記する」係。
'  ヘッダ行（日付／始値／高値／安値／終値／出来高）は自動検出するので、
'  貼り付け位置や列順が違っても動く。
'
'  書き込み先： 始値／高値／安値／終値／出来高 の5シート
'               3行目＝日付ヘッダ、E列＝最新、F列＝1日前 …（右へ行くほど古い）
'               銘柄の行は 銘柄管理!B6:B305 の並び順（6行目＝1銘柄目）
'==============================================================================
Option Explicit

Private Const PWD          As String = "ne19480314"
Private Const HIST_MAX     As Long = 250      ' E列〜IT列
Private Const FIRST_COL    As Long = 5        ' E列
Private Const DATE_ROW     As Long = 3
Private Const STOCK_ROW1   As Long = 6        ' 価格シートの1銘柄目
Private Const MEI_ROW1     As Long = 6        ' 銘柄管理の1銘柄目
Private Const MEI_ROWN     As Long = 305

Private Const SH_DATA      As String = "データ取込"
Private Const SH_MEI       As String = "銘柄管理"


'==============================================================================
' ① 1銘柄だけ取り込む（データ取込シートに貼り付けたものを反映）
'==============================================================================
Public Sub 価格履歴_取込_1銘柄()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    On Error GoTo ErrHandler
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim dws As Worksheet
    Set dws = SheetOrNothing(SH_DATA)
    If dws Is Nothing Then Err.Raise 5, , "「" & SH_DATA & "」シートがありません。"

    Dim code As String
    code = UCase$(Trim$(CStr(dws.Range("B1").Value)))
    If code = "" Then Err.Raise 5, , "データ取込シートの B1 に銘柄コードを入れてください。"

    Dim n As Long
    n = ImportOne(dws, code, HIST_MAX)

    Cleanup prevCalc
    MsgBox code & " を " & n & " 日分 取り込みました。（上限 " & HIST_MAX & " 日）" & vbCrLf & vbCrLf & _
           "足りない場合は データ取込!K1（本数）を大きくし、" & vbCrLf & _
           "M1（開始日 yyyymmdd）を新しい日付にしてから、" & vbCrLf & _
           "再計算してもう一度実行してください。", _
           vbInformation, "取込完了"
    Exit Sub

ErrHandler:
    Dim msg As String: msg = Err.Description
    Cleanup prevCalc
    MsgBox "取込を中断しました。" & vbCrLf & vbCrLf & msg, vbCritical, "取込エラー"
End Sub


'==============================================================================
' ② どの銘柄がまだ足りないかを確認する
'    RssChartPast は1銘柄ずつなので、B1 を書き換えて①を繰り返す。
'    その進み具合をこのSubで確認する。
'==============================================================================
Public Sub 価格履歴_取込_残り銘柄を確認()
    Dim mws As Worksheet, cws As Worksheet
    Set mws = SheetOrNothing(SH_MEI)
    Set cws = SheetOrNothing("終値")
    If mws Is Nothing Or cws Is Nothing Then Exit Sub

    Dim r As Long, done As Long, todo As Long, firstTodo As String
    For r = MEI_ROW1 To MEI_ROWN
        If Trim$(CStr(mws.Cells(r, 2).Value)) <> "" Then
            Dim pr As Long: pr = STOCK_ROW1 + (r - MEI_ROW1)
            If CountFilled(cws, pr) >= 100 Then
                done = done + 1
            Else
                todo = todo + 1
                If firstTodo = "" Then firstTodo = CStr(mws.Cells(r, 2).Value) & " " & _
                                                  CStr(mws.Cells(r, 3).Value)
            End If
        End If
    Next r

    MsgBox "100日以上そろっている銘柄: " & done & vbCrLf & _
           "まだ足りない銘柄:          " & todo & vbCrLf & vbCrLf & _
           IIf(todo > 0, "次に取り込む銘柄: " & firstTodo, "すべて完了しています。"), _
           vbInformation, "取込の進み具合"
End Sub


'==============================================================================
' ③ 現在の履歴日数を確認する
'==============================================================================
Public Sub 価格履歴_日数を確認()
    Dim s As Variant, msg As String
    For Each s In Array("始値", "高値", "安値", "終値", "出来高")
        Dim ws As Worksheet: Set ws = SheetOrNothing(CStr(s))
        If Not ws Is Nothing Then
            Dim c As Long, n As Long: n = 0
            For c = FIRST_COL To FIRST_COL + HIST_MAX - 1
                If IsDate(ws.Cells(DATE_ROW, c).Value) Then n = n + 1
            Next c
            msg = msg & CStr(s) & ": " & n & " 日分" & vbCrLf
        End If
    Next s
    MsgBox msg & vbCrLf & "上限は " & HIST_MAX & " 日（E列〜IT列）です。", _
           vbInformation, "履歴日数"
End Sub


'------------------------------------------------------------------ 取込本体
Private Function ImportOne(ByVal dws As Worksheet, ByVal code As String, _
                           ByVal nDays As Long) As Long
    ' --- 銘柄管理での位置 → 価格シートの行 ---
    Dim mws As Worksheet: Set mws = SheetOrNothing(SH_MEI)
    If mws Is Nothing Then Err.Raise 5, , "「" & SH_MEI & "」シートがありません。"
    Dim r As Long, meiRow As Long
    For r = MEI_ROW1 To MEI_ROWN
        If UCase$(Trim$(CStr(mws.Cells(r, 2).Value))) = code Then meiRow = r: Exit For
    Next r
    If meiRow = 0 Then Err.Raise 5, , "銘柄コード " & code & " が銘柄管理にありません。"
    Dim priceRow As Long: priceRow = STOCK_ROW1 + (meiRow - MEI_ROW1)

    ' --- ヘッダ行と列を自動検出 ---
    Dim hRow As Long, cDate As Long, cO As Long, cH As Long, cL As Long, cC As Long, cV As Long
    FindHeader dws, hRow, cDate, cO, cH, cL, cC, cV
    If hRow = 0 Then Err.Raise 5, , "「日付」「始値」「終値」を含むヘッダ行が見つかりません。"

    ' --- 明細を読み、日付降順（新しい順）に並べる ---
    Dim d() As Date, o() As Double, h() As Double, lo() As Double, c() As Double, v() As Double
    ReDim d(1 To HIST_MAX): ReDim o(1 To HIST_MAX): ReDim h(1 To HIST_MAX)
    ReDim lo(1 To HIST_MAX): ReDim c(1 To HIST_MAX): ReDim v(1 To HIST_MAX)

    Dim lastRow As Long, i As Long, n As Long
    lastRow = dws.Cells(dws.Rows.Count, cDate).End(xlUp).Row
    For i = hRow + 1 To lastRow
        If n >= nDays Then Exit For
        Dim cell As Variant: cell = dws.Cells(i, cDate).Value
        If IsDate(cell) Then
            Dim cv As Variant: cv = dws.Cells(i, cC).Value
            Dim vv As Variant: vv = IIf(cV > 0, dws.Cells(i, cV).Value, 0)
            ' 出来高0の日は休業・欠損として飛ばす（V805の取込と同じ扱い）
            If IsNumeric(cv) And cv > 0 And IsNumeric(vv) And vv > 0 Then
                n = n + 1
                d(n) = CDate(cell)
                c(n) = CDbl(cv)
                v(n) = CDbl(vv)
                o(n) = PickNum(dws, i, cO, c(n))
                h(n) = PickNum(dws, i, cH, c(n))
                lo(n) = PickNum(dws, i, cL, c(n))
            End If
        End If
    Next i
    If n = 0 Then Err.Raise 5, , "取り込める明細行がありませんでした。" & vbCrLf & _
                                 "データ取込!K1(本数) と M1(開始日) を確認し、" & vbCrLf & _
                                 "再計算してからもう一度実行してください。"

    SortDescByDate d, o, h, lo, c, v, n

    ' --- 5シートへ書き込み ---
    WriteSheet "始値", priceRow, d, o, n
    WriteSheet "高値", priceRow, d, h, n
    WriteSheet "安値", priceRow, d, lo, n
    WriteSheet "終値", priceRow, d, c, n
    WriteSheet "出来高", priceRow, d, v, n

    ImportOne = n
End Function


Private Sub WriteSheet(ByVal name As String, ByVal priceRow As Long, _
                       ByRef d() As Date, ByRef val() As Double, ByVal n As Long)
    Dim ws As Worksheet: Set ws = SheetOrNothing(name)
    If ws Is Nothing Then Exit Sub
    ws.Unprotect PWD

    ' 対象銘柄行の履歴域をいったん消す（E列〜IT列）
    ws.Range(ws.Cells(priceRow, FIRST_COL), _
             ws.Cells(priceRow, FIRST_COL + HIST_MAX - 1)).ClearContents

    Dim i As Long
    For i = 1 To n
        ws.Cells(priceRow, FIRST_COL + i - 1).Value = val(i)
        ' 3行目の日付ヘッダは、まだ空か古い場合だけ更新する
        Dim hc As Range: Set hc = ws.Cells(DATE_ROW, FIRST_COL + i - 1)
        If Not IsDate(hc.Value) Then
            hc.Value = d(i)
        ElseIf CDate(hc.Value) < d(i) Then
            hc.Value = d(i)
        End If
    Next i

    ws.Protect PWD, UserInterfaceOnly:=True
End Sub


'------------------------------------------------------------------ 補助関数
Private Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
End Function

Public Sub 取込設定_250日にする()
    Dim dws As Worksheet: Set dws = SheetOrNothing(SH_DATA)
    If dws Is Nothing Then Exit Sub
    dws.Range("K1").Value = HIST_MAX                       ' 本数
    dws.Range("M1").Value = Format(Date, "yyyymmdd")       ' 開始日＝今日
    Application.CalculateFull
    MsgBox "データ取込シートを設定しました。" & vbCrLf & vbCrLf & _
           "  K1（本数）  = " & HIST_MAX & vbCrLf & _
           "  M1（開始日）= " & dws.Range("M1").Value & vbCrLf & vbCrLf & _
           "RSSがデータを取り終えたら、B1 に銘柄コードを入れて" & vbCrLf & _
           "「価格履歴_取込_1銘柄」を実行してください。", _
           vbInformation, "取込設定"
End Sub

Private Function PickNum(ByVal dws As Worksheet, ByVal r As Long, _
                         ByVal col As Long, ByVal fallback As Double) As Double
    If col = 0 Then PickNum = fallback: Exit Function
    Dim v As Variant: v = dws.Cells(r, col).Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then PickNum = CDbl(v): Exit Function
    End If
    PickNum = fallback          ' 欠損は終値で代用（レンジ0扱いになる）
End Function

Private Sub FindHeader(ByVal dws As Worksheet, ByRef hRow As Long, _
                       ByRef cDate As Long, ByRef cO As Long, ByRef cH As Long, _
                       ByRef cL As Long, ByRef cC As Long, ByRef cV As Long)
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
                    Case "日付":                              cDate = c
                    Case "始値":                              cO = c
                    Case "高値":                              cH = c
                    Case "安値":                              cL = c
                    Case "終値":                              cC = c
                    Case "出来高", "売買高", "出来高(株)", "出来高（株）": cV = c
                End Select
            Next c
            Exit For
        End If
    Next r
End Sub

' 日付の新しい順に並べ替える（単純挿入ソート。250件なので十分速い）
Private Sub SortDescByDate(ByRef d() As Date, ByRef o() As Double, ByRef h() As Double, _
                           ByRef l() As Double, ByRef c() As Double, ByRef v() As Double, _
                           ByVal n As Long)
    Dim i As Long, j As Long
    For i = 2 To n
        Dim kd As Date, ko As Double, kh As Double, kl As Double, kc As Double, kv As Double
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

Private Function CountFilled(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim c As Long, n As Long
    For c = FIRST_COL To FIRST_COL + HIST_MAX - 1
        If IsNumeric(ws.Cells(r, c).Value) Then
            If ws.Cells(r, c).Value > 0 Then n = n + 1
        End If
    Next c
    CountFilled = n
End Function

' 中断しても Calculation が Manual のまま残らないようにする。
' V805 の抽出Sub4本はこれが無く、残留すると現在値が更新されず
' 損切判定が完全に沈黙する状態になっていた。
Private Sub Cleanup(ByVal prevCalc As XlCalculation)
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub
