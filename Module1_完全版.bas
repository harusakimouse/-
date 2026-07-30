Attribute VB_Name = "Module1"
Option Explicit
' =============================================
' Module1: 楽天RSS データ自動取得（完成版 v4）
' =============================================
' Sheet1にRSS数式が常駐（触らない）
' VBAはSheet1の値を読み取り→ターゲットシートに蓄積
'
' v4の修正点（v3からの差分）
'  ★ HasToday を IsDate ベースに修正（日付書式セルでも本日分を正しく判定）
'     - これで「30秒ごとの無限再取得＝重複」と「書込0件」の誤表示が解消
'  ★ ReArmScheduler を追加（OnTimeチェーンが切れても自己復旧）
'
' v3の修正点
'  1. Application.OnTime をブック名で修飾（他ブックへの誤解決を防止）
'  2. 稼働状態・次回時刻を「実行ログ」シートに保存（VBAリセットで消えない）
'  3. 完了判定を対象シートの実データで行う（重複防止＋失敗リトライを両立）
'  4. 自動実行中はMsgBoxを出さない（モーダルはOnTimeを止めるため）
'  5. 実行履歴・エラー内容を「実行ログ」シートに追記
' =============================================

Private Const LOGSH As String = "実行ログ"

' ==============  ログ・状態管理  ==============
Private Function LogSheet() As Worksheet
    Set LogSheet = ThisWorkbook.Worksheets(LOGSH)
End Function

Private Sub LogWrite(ByVal proc As String, ByVal result As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = LogSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = proc
    ws.Cells(r, 3).Value = result
    ws.Range("F3").Value = Now
End Sub

Private Property Get IsRunning() As Boolean
    On Error Resume Next
    IsRunning = (LogSheet.Range("F2").Value = True)
End Property

Private Property Let IsRunning(ByVal v As Boolean)
    On Error Resume Next
    LogSheet.Range("F2").Value = v
End Property

' ★ OnTimeは必ずブック名で修飾する
Private Function MacroRef(ByVal nm As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!" & nm
End Function

' ★ 予約済みポーリングを確実に取り消す（二重予約・他マクロ競合対策）
'   自分の (時刻 + 'ブック'!PollSchedule) だけを取り消すので、
'   他VBAが登録した OnTime には一切干渉しない。
Private Sub CancelPending()
    On Error Resume Next
    Dim nx As Variant: nx = LogSheet.Range("F1").Value
    If IsDate(nx) Then
        Application.OnTime CDate(nx), MacroRef("PollSchedule"), , False
    End If
    LogSheet.Range("F1").ClearContents
End Sub

' ==============  スケジューラ  ==============
Public Sub CheckSchedule()
    On Error Resume Next
    CancelPending                     ' ★ 既存の予約を先に解除（二重起動を防ぐ）
    If Time >= TimeValue("15:45:00") Then
        IsRunning = False
        Application.StatusBar = "15:45以降 スケジューラ未起動"
        LogWrite "起動見送り", "15:45以降のため"
        Exit Sub
    End If
    IsRunning = True
    Application.StatusBar = Format(Now, "hh:nn:ss") & " スケジューラ起動済み"
    LogWrite "スケジューラ起動", "OK"
    ScheduleNext
End Sub

Public Sub StopScheduler()
    On Error Resume Next
    CancelPending
    IsRunning = False
    Application.StatusBar = "スケジューラ停止"
    LogWrite "スケジューラ停止", "OK"
End Sub

Private Sub ScheduleNext()
    On Error Resume Next
    If Not IsRunning Then Exit Sub
    CancelPending                     ' ★ 直前の予約を解除してから入れ直す（チェーン多重化防止）
    Dim nx As Date: nx = Now + TimeValue("00:00:30")
    Err.Clear
    Application.OnTime nx, MacroRef("PollSchedule")
    If Err.Number = 0 Then
        LogSheet.Range("F1").Value = nx           ' 成功した予約だけを記録
    Else
        LogWrite "OnTime予約失敗", "Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Sub

Public Sub PollSchedule()
    On Error Resume Next               ' ★ 何が起きてもチェーンを止めない
    If Not IsRunning Then Exit Sub

    Dim t As Date: t = Time

    If t >= TimeValue("15:45:00") Then
        StopScheduler
        Application.StatusBar = "15:45 自動終了"
        LogWrite "自動終了", "15:45"
        Exit Sub
    End If

    If t >= TimeValue("15:02:00") Then RunOnce "1500"
    If t >= TimeValue("15:17:00") Then RunOnce "1520"
    If t >= TimeValue("15:31:00") Then RunOnce "1530"

    ScheduleNext   ' 途中で何があっても必ず次を予約
End Sub

' ★ v4追加：自動発火チェーンの自己復旧
'   OnTimeの30秒チェーンは、稼働中に一度でもタイミングを取り逃すと
'   そのまま停止し、開き直すまで復活しない。稼働フラグは立っているのに
'   次回予約(F1)が失われている/過ぎている場合だけ静かに立て直す。
'   ※ HasToday修正済みで再発火は無害（冪等）なので、重複は発生しない。
'   ThisWorkbook 側に次の1行を入れると、シート切替のたびに点検される：
'       Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'           On Error Resume Next
'           ReArmScheduler
'       End Sub
Public Sub ReArmScheduler()
    On Error Resume Next
    If Time >= TimeValue("15:45:00") Then Exit Sub   ' 稼働時間外は何もしない
    If Not IsRunning Then Exit Sub                   ' 停止中なら触らない

    Dim nx As Variant: nx = LogSheet.Range("F1").Value
    If IsDate(nx) Then
        If CDate(nx) > Now - TimeValue("00:01:30") Then Exit Sub  ' まだ生きている
    End If

    ScheduleNext                                     ' 切れている → 立て直す
    LogWrite "自己復旧", "ポーリング再武装"
End Sub

Public Sub スケジュール設定()
    MsgBox "スケジューラはブック起動時に自動開始します。" & vbCrLf & _
           "  開始: CheckSchedule / 停止: StopScheduler" & vbCrLf & _
           "15:02→1500 / 15:17→1520 / 15:31→1530 / 15:45終了" & vbCrLf & vbCrLf & _
           "状態と履歴は「実行ログ」シートで確認できます。", _
           vbInformation, "スケジュール設定"
End Sub

' ==============  1日1回だけ実行する制御  ==============
Private Sub RunOnce(ByVal nm As String)
    On Error GoTo Fail
    If HasToday(nm) Then Exit Sub          ' 既に本日分あり → 何もしない
    CaptureStockData nm
    If HasToday(nm) Then
        LogWrite nm & " 自動取得", "成功"
    Else
        LogWrite nm & " 自動取得", "書込0件 → 次回リトライ"
    End If
    Exit Sub
Fail:
    LogWrite nm & " 自動取得", "エラー: " & Err.Description & " → 次回リトライ"
End Sub

' 対象シートに本日日付の行が既にあるか
' ★ v4修正：A列は日付書式セルなので .Value は Date 型で返る。
'   IsNumeric(Date型) は False になるため、IsDate で判定する。
Public Function HasToday(ByVal nm As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(nm)
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    If lr < 2 Then Exit Function

    Dim v As Variant: v = ws.Cells(lr, 1).Value
    If IsDate(v) Then
        HasToday = (Int(CDbl(CDate(v))) = CLng(Date))   ' 日付書式セル
    ElseIf IsNumeric(v) Then
        HasToday = (CLng(v) = CLng(Date))               ' 生のシリアル値の保険
    End If
End Function

' ==============  手動実行用（重複ガード付き）  ==============
Public Sub Auto1500()
    ManualCapture "1500"
End Sub

Public Sub Auto1520()
    ManualCapture "1520"
End Sub

Public Sub Auto1530()
    ManualCapture "1530"
End Sub

' 旧名互換：古いボタン登録が残っていても動くようにする
Public Sub Get1500()
    ManualCapture "1500"
End Sub

Public Sub SaveHistory1500()
    ManualCapture "1500"
End Sub

Public Sub SaveHistory1520()
    ManualCapture "1520"
End Sub

Public Sub SaveHistory1530()
    ManualCapture "1530"
End Sub

Private Sub ManualCapture(ByVal nm As String)
    If HasToday(nm) Then
        If MsgBox(nm & " は本日分がすでに蓄積されています。" & vbCrLf & _
                  "もう一度追記すると重複します。実行しますか？", _
                  vbYesNo + vbExclamation, "重複確認") <> vbYes Then
            LogWrite nm & " 手動取得", "ユーザー中止（本日分あり）"
            Exit Sub
        End If
    End If
    CaptureStockData nm
    LogWrite nm & " 手動取得", "実行"
End Sub

' ==============  共通キャプチャ  ==============
Private Sub CaptureStockData(ByVal sheetName As String)
    On Error GoTo ErrHandler

    Dim wsTarget As Worksheet, wsFunc As Worksheet, wsWork As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets(sheetName)
    Set wsFunc = ThisWorkbook.Worksheets("関数")
    Set wsWork = ThisWorkbook.Worksheets("Sheet1")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' --- 銘柄数（Sheet1のA列） ---
    Dim n As Long
    n = wsWork.Cells(wsWork.Rows.Count, 1).End(xlUp).Row - 1  ' ヘッダー除く
    If n <= 0 Then GoTo Cleanup

    ' --- 市場指標（関数シート） ---
    Dim topixVal As Variant: topixVal = wsFunc.Range("B1").Value
    Dim topixRate As Variant: topixRate = wsFunc.Range("C1").Value
    Dim nkVal As Variant: nkVal = wsFunc.Range("D1").Value
    Dim nkRate As Variant: nkRate = wsFunc.Range("E1").Value
    If IsNumeric(topixRate) Then topixRate = topixRate / 100
    If IsNumeric(nkRate) Then nkRate = nkRate / 100

    ' --- 書き込み開始行 ---
    Dim startRow As Long
    If wsTarget.Cells(2, 6).Value = "" Then
        startRow = 2
    Else
        startRow = wsTarget.Cells(wsTarget.Rows.Count, 6).End(xlUp).Row + 1
    End If

    Dim todayVal As Long: todayVal = CLng(Date)

    ' --- Sheet1から値を読み取って書き込み ---
    ' Sheet1列: A=コード B=銘柄名称 C=現在値 D=前日比 E=前日終値
    '           F=始値 G=高値 H=安値 I=出来高 J=VWAP
    '           K=年高 L=年安 M=売気配数量 N=買気配 O=買気配数量
    '           P=更新時間 Q=後場始値

    Dim i As Long
    Dim r As Long, srcRow As Long
    Dim curP As Variant, chg As Variant, prevC As Variant
    Dim openP As Variant, highP As Variant, lowP As Variant
    Dim vol As Variant, vwap As Variant, yHi As Variant, yLo As Variant
    Dim askQ As Variant, bidP As Variant, bidQ As Variant
    Dim updT As Variant, pmO As Variant

    For i = 1 To n
        r = startRow + i - 1
        srcRow = i + 1  ' Sheet1はヘッダー付き

        With wsTarget
            .Cells(r, 1).Value = todayVal
            .Cells(r, 1).NumberFormat = "m/d/yyyy"
            .Cells(r, 2).Value = topixVal
            .Cells(r, 3).Value = topixRate
            .Cells(r, 4).Value = nkVal
            .Cells(r, 5).Value = nkRate
            .Cells(r, 6).Value = wsWork.Cells(srcRow, 1).Value  ' F: コード
        End With

        curP = wsWork.Cells(srcRow, 3).Value
        chg = wsWork.Cells(srcRow, 4).Value
        prevC = wsWork.Cells(srcRow, 5).Value
        openP = wsWork.Cells(srcRow, 6).Value
        highP = wsWork.Cells(srcRow, 7).Value
        lowP = wsWork.Cells(srcRow, 8).Value
        vol = wsWork.Cells(srcRow, 9).Value
        vwap = wsWork.Cells(srcRow, 10).Value
        yHi = wsWork.Cells(srcRow, 11).Value
        yLo = wsWork.Cells(srcRow, 12).Value
        askQ = wsWork.Cells(srcRow, 13).Value
        bidP = wsWork.Cells(srcRow, 14).Value
        bidQ = wsWork.Cells(srcRow, 15).Value
        updT = wsWork.Cells(srcRow, 16).Value
        pmO = wsWork.Cells(srcRow, 17).Value

        wsTarget.Cells(r, 7).Value = wsWork.Cells(srcRow, 2).Value  ' G: 銘柄名

        If IsNumeric(curP) And curP <> 0 Then
            wsTarget.Cells(r, 8).Value = curP                       ' H: 現在値
            wsTarget.Cells(r, 9).Value = chg                        ' I: 前日比
            If IsNumeric(prevC) And prevC <> 0 Then _
                wsTarget.Cells(r, 10).Value = chg / prevC           ' J: 前日比率
            wsTarget.Cells(r, 11).Value = prevC                     ' K: 前日終値
            wsTarget.Cells(r, 12).Value = openP                     ' L: 始値
            wsTarget.Cells(r, 13).Value = highP                     ' M: 高値
            wsTarget.Cells(r, 14).Value = lowP                      ' N: 安値
            If IsNumeric(openP) And openP <> 0 Then _
                wsTarget.Cells(r, 15).Value = (curP - openP) / openP   ' O: 寄付騰落率
            If IsNumeric(pmO) And pmO <> 0 Then
                wsTarget.Cells(r, 16).Value = pmO                   ' P: 後場始値
                wsTarget.Cells(r, 17).Value = (curP - pmO) / pmO    ' Q: 後場騰落率
            End If
            If IsNumeric(lowP) And lowP <> 0 Then _
                wsTarget.Cells(r, 18).Value = (curP - lowP) / lowP  ' R: 安値からの上昇率
            wsTarget.Cells(r, 19).Value = vol                       ' S: 出来高
            If IsNumeric(vwap) Then _
                wsTarget.Cells(r, 20).Value = vwap                  ' T: VWAP
            wsTarget.Cells(r, 21).Value = ExtractNum(yHi)           ' U: 年高
            wsTarget.Cells(r, 22).Value = ExtractNum(yLo)           ' V: 年安
            If IsNumeric(askQ) Then wsTarget.Cells(r, 23).Value = askQ  ' W
            If IsNumeric(bidP) Then wsTarget.Cells(r, 24).Value = bidP  ' X
            If IsNumeric(bidQ) Then wsTarget.Cells(r, 25).Value = bidQ  ' Y
            If Len(CStr(updT)) > 0 Then
                If IsDate(updT) Then
                    wsTarget.Cells(r, 26).Value = CDbl(TimeValue(CStr(updT)))
                ElseIf IsNumeric(updT) Then
                    wsTarget.Cells(r, 26).Value = updT
                End If
            End If
        End If
    Next i

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = sheetName & " 取得完了 (" & _
        Format(Now, "hh:nn:ss") & ") " & n & "銘柄"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = sheetName & " エラー"
    LogWrite sheetName & " CaptureStockData", "エラー: " & Err.Description
End Sub

Private Function ExtractNum(ByVal v As Variant) As Variant
    If IsNumeric(v) Then ExtractNum = v: Exit Function
    If IsEmpty(v) Or v = "" Then ExtractNum = Empty: Exit Function
    Dim s As String: s = CStr(v)
    Dim pos As Long: pos = InStr(s, " ")
    If pos > 0 Then s = Left(s, pos - 1)
    s = Replace(s, ",", "")
    If IsNumeric(s) Then ExtractNum = CDbl(s) Else ExtractNum = v
End Function
