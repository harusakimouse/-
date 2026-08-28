Attribute VB_Name = "Module_TickLogger"
Option Explicit

'==================================================================
' Module_TickLogger : 15:00〜15:20 のティックを各銘柄シートに記録する
'
'   RSS は「歩み値」を直接くれないので、現在値・出来高（当日累計）を
'   高速ポーリングし、出来高が増えた瞬間を1ティックとして記録します。
'     ティック出来高 = 今回の累計出来高 − 前回の累計出来高
'
'   POLL_MS = 200 なら 1秒あたり最大5ティックまで分離できるので
'   ③ 約定速度（SpeedMax >= 3）の判定が成立します。
'==================================================================

Public gLogging As Boolean
Private mArmedTime As Date

'------------------------------------------------------------------
' ティック記録 開始（ボタン②）
'------------------------------------------------------------------
Public Sub StartTickLogging()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim prevVol() As Double
    Dim nextRow() As Long
    Dim i As Long, n As Long
    Dim tickTotal As Long
    Dim res As Worksheet
    Dim lastPaint As Single

    If gLogging Then Exit Sub

    Set shts = TargetSheets()
    If shts.Count = 0 Then
        MsgBox "記録対象の銘柄シートがありません。" & vbCrLf & _
               "各シートの B3 に証券コードを入れてから「① 準備」を実行してください。", _
               vbExclamation, "ティック記録"
        Exit Sub
    End If

    n = shts.Count
    ReDim prevVol(1 To n)
    ReDim nextRow(1 To n)

    For i = 1 To n
        Set ws = shts(i)
        prevVol(i) = 0                       ' 最初の1回は基準値取りに使う
        nextRow(i) = LastTickRow(ws) + 1
        If nextRow(i) < TICK_FIRST_ROW Then nextRow(i) = TICK_FIRST_ROW
    Next i

    Set res = ResultSheet()
    res.Range("U1").Value = "記録中… 開始 " & Format$(Now, "hh:mm:ss")

    gLogging = True
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Fin

    Do While gLogging

        If NowTime() > TimeValue(JUDGE_END) Then Exit Do          ' 15:20 で自動終了

        If NowTime() >= TimeValue(JUDGE_START) Then
            For i = 1 To n
                Set ws = shts(i)
                If PollOne(ws, prevVol(i), nextRow(i)) Then tickTotal = tickTotal + 1
            Next i
        End If

        If Timer - lastPaint > 1 Then
            res.Range("U2").Value = "ティック " & tickTotal & " 件 / 最終 " & Format$(Now, "hh:mm:ss")
            lastPaint = Timer
        End If

        WaitMs POLL_MS
    Loop

Fin:
    gLogging = False
    Application.EnableCancelKey = xlInterrupt

    If Err.Number <> 0 And Err.Number <> 18 Then
        res.Range("U1").Value = "記録エラー " & Err.Number & " : " & Err.Description
        MsgBox "記録中にエラーが発生しました。" & vbCrLf & _
               Err.Number & " : " & Err.Description, vbCritical, "ティック記録"
    Else
        res.Range("U1").Value = "記録停止 " & Format$(Now, "hh:mm:ss") & _
                                "（ティック " & tickTotal & " 件）"
    End If

    If AUTO_JUDGE_ON_STOP Then JudgeAll
End Sub

'------------------------------------------------------------------
' ティック記録 停止（ボタン③）
'------------------------------------------------------------------
Public Sub StopTickLogging()
    gLogging = False
End Sub

'------------------------------------------------------------------
' 記録を止めてすぐ判定（ボタン③の代替）
'------------------------------------------------------------------
Public Sub StopAndJudge()
    If gLogging Then
        gLogging = False        ' StartTickLogging 側で自動的に JudgeAll まで走ります
    Else
        JudgeAll
    End If
End Sub

'------------------------------------------------------------------
' 1銘柄を1回ポーリングし、約定があれば1行追記する
'   戻り値 : 追記したら True
'------------------------------------------------------------------
Private Function PollOne(ws As Worksheet, ByRef prevVol As Double, ByRef nextRow As Long) As Boolean

    Dim vol As Double, price As Double, bid As Double, ask As Double
    Dim d As Double
    Dim t As Date

    vol = NumOrZero(ws.Range(LIVE_VOL).Value)
    price = NumOrZero(ws.Range(LIVE_PRICE).Value)

    If vol <= 0 Or price <= 0 Then Exit Function      ' RSS 未接続 / 気配のみ

    If prevVol = 0 Then                               ' 記録開始時点の基準値
        prevVol = vol
        Exit Function
    End If

    If vol <= prevVol Then Exit Function              ' 約定なし
    d = vol - prevVol
    prevVol = vol

    bid = NumOrZero(ws.Range(LIVE_BID).Value)
    ask = NumOrZero(ws.Range(LIVE_ASK).Value)
    t = TickTime(ws)

    If nextRow > ws.Rows.Count - 1 Then Exit Function

    With ws
        .Cells(nextRow, COL_TIME).Value = t
        .Cells(nextRow, COL_TIME).NumberFormatLocal = "hh:mm:ss"
        .Cells(nextRow, COL_PRICE).Value = price
        .Cells(nextRow, COL_VOL).Value = d
        .Cells(nextRow, COL_BID).Value = bid
        .Cells(nextRow, COL_ASK).Value = ask
    End With

    nextRow = nextRow + 1
    PollOne = True
End Function

'------------------------------------------------------------------
' ティックの時刻（RSS の現在値時刻優先、取れなければ PC 時計）
'------------------------------------------------------------------
Private Function TickTime(ws As Worksheet) As Date

    Dim v As Variant

    If USE_RSS_TIME Then
        v = ws.Range(LIVE_TIME).Value
        On Error Resume Next
        If Not IsError(v) And Not IsEmpty(v) Then
            If VarType(v) = vbDate Then
                TickTime = CDate(v)
            ElseIf IsNumeric(v) Then
                TickTime = CDate(CDbl(v))
            Else
                TickTime = CDate(CStr(v))
            End If
        End If
        On Error GoTo 0
        If TickTime > 0 Then
            TickTime = TimeSerial(Hour(TickTime), Minute(TickTime), Second(TickTime))
            Exit Function
        End If
    End If

    TickTime = TimeSerial(Hour(Now), Minute(Now), Second(Now))
End Function

'==================================================================
' 自動開始の予約（14:59:30 に記録を開始し、15:20 で自動判定）
'==================================================================

Public Sub ArmAutoRun()

    DisarmAutoRun

    mArmedTime = Date + TimeValue(AUTO_ARM_TIME)

    If mArmedTime <= Now Then
        MsgBox "予約時刻（" & AUTO_ARM_TIME & "）を過ぎています。" & vbCrLf & _
               "「② ティック記録 開始」を手で押してください。", vbInformation, "自動実行の予約"
        mArmedTime = 0
        Exit Sub
    End If

    Application.OnTime mArmedTime, "StartTickLogging"
    ResultSheet().Range("U1").Value = "自動開始を予約しました（" & AUTO_ARM_TIME & "）"
End Sub

Public Sub DisarmAutoRun()
    On Error Resume Next
    If mArmedTime > 0 Then
        Application.OnTime mArmedTime, "StartTickLogging", , False
    End If
    mArmedTime = 0
    On Error GoTo 0
End Sub

'==================================================================
' ティックログのクリア（ボタン⑤）
'==================================================================
Public Sub ClearAllTicks()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    If MsgBox("全銘柄シートのティックログを消去します。よろしいですか？", _
              vbYesNo + vbQuestion, "ティックログ消去") <> vbYes Then Exit Sub

    Set shts = TargetSheets()

    Application.ScreenUpdating = False
    For i = 1 To shts.Count
        Set ws = shts(i)
        lastRow = LastTickRow(ws)
        If lastRow >= TICK_FIRST_ROW Then
            ws.Range(ws.Cells(TICK_FIRST_ROW, COL_TIME), ws.Cells(lastRow, COL_SPD)).ClearContents
        End If
        ws.Range(RES_TOP).Offset(0, 1).Resize(15, 1).ClearContents
        ws.Range(RES_TOP).Offset(0, 1).Interior.Pattern = xlNone
    Next i

    With ResultSheet()
        .Range(.Cells(ROW_HEADER + 1, 2), .Cells(ROW_HEADER + 200, 18)).ClearContents
        .Range(.Cells(ROW_HEADER + 1, 2), .Cells(ROW_HEADER + 200, 18)).Interior.Pattern = xlNone
        .Range("U1:U3").ClearContents
    End With
    Application.ScreenUpdating = True
End Sub
