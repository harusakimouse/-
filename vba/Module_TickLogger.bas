Attribute VB_Name = "Module_TickLogger"
Option Explicit

'==================================================================
' Module_TickLogger : 15:00〜15:20 のティックを各銘柄シートに記録する
'
'  ● モード1（既定 / LOG_MODE = 1）… 歩み値（TICK）ブロック追従
'      RSS の歩み値をシート上のブロック（既定 AB3〜）に出しておき、
'      前回から増えた行だけを検出してティックログに追記します。
'      約定1本ずつのデータなので ③SpeedMax が正確に出ます。
'
'      ・並び順（新しい順 / 古い順）は自動判定します
'      ・記録開始時にブロック内の履歴もまとめて取り込みます
'      ・ブロックを追い越した（＝取りこぼした）場合は警告を出します
'
'  ● モード2（LOG_MODE = 2）… 現在値・出来高ポーリング
'      歩み値が使えない環境向け。出来高（当日累計）の増分を1ティックとみなします。
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
    Dim anchor() As String
    Dim gapWarned() As Boolean
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
    ReDim anchor(1 To n)
    ReDim gapWarned(1 To n)

    For i = 1 To n
        Set ws = shts(i)
        prevVol(i) = 0                       ' モード2：最初の1回は基準値取りに使う
        nextRow(i) = LastTickRow(ws) + 1
        If nextRow(i) < TICK_FIRST_ROW Then nextRow(i) = TICK_FIRST_ROW
        anchor(i) = SeedAnchorFromLog(ws)    ' モード1：途中再開でも二重取込しない
    Next i

    Set res = ResultSheet()
    res.Range("X1").Value = "記録中… 開始 " & Format$(Now, "hh:mm:ss") & _
                            "（" & IIf(LOG_MODE = 1, "歩み値追従", "出来高ポーリング") & "）"

    gLogging = True
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Fin

    Do While gLogging

        If NowTime() > TimeValue(JUDGE_END) Then Exit Do          ' 15:20 で自動終了

        If NowTime() >= TimeValue(JUDGE_START) Then
            For i = 1 To n
                Set ws = shts(i)
                If LOG_MODE = 1 Then
                    tickTotal = tickTotal + FollowTickBlock(ws, anchor(i), nextRow(i), gapWarned(i))
                Else
                    If PollOne(ws, prevVol(i), nextRow(i)) Then tickTotal = tickTotal + 1
                End If
            Next i
        End If

        If Timer - lastPaint > 1 Then
            res.Range("X2").Value = "ティック " & tickTotal & " 件 / 最終 " & Format$(Now, "hh:mm:ss")
            RefreshQuotes shts                    ' 最良気配のリアルタイム表示
            lastPaint = Timer
        End If

        WaitMs POLL_MS
    Loop

Fin:
    gLogging = False
    Application.EnableCancelKey = xlInterrupt

    If Err.Number <> 0 And Err.Number <> 18 Then
        res.Range("X1").Value = "記録エラー " & Err.Number & " : " & Err.Description
        MsgBox "記録中にエラーが発生しました。" & vbCrLf & _
               Err.Number & " : " & Err.Description, vbCritical, "ティック記録"
    Else
        res.Range("X1").Value = "記録停止 " & Format$(Now, "hh:mm:ss") & _
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
' 記録を止めてすぐ判定（ボタン③）
'------------------------------------------------------------------
Public Sub StopAndJudge()
    If gLogging Then
        gLogging = False        ' StartTickLogging 側で自動的に JudgeAll まで走ります
    Else
        JudgeAll
    End If
End Sub

'==================================================================
' モード1 : 歩み値（TICK）ブロック追従
'==================================================================

'------------------------------------------------------------------
' 歩み値ブロックを読み、前回以降に増えた行だけを追記する
'   戻り値 : 追記したティック数
'------------------------------------------------------------------
Private Function FollowTickBlock(ws As Worksheet, ByRef anchorKey As String, _
                                 ByRef nextRow As Long, ByRef gapWarned As Boolean) As Long

    Dim b As Variant
    Dim n As Long, i As Long, idx As Long
    Dim added As Long
    Dim bid As Double, ask As Double

    b = ReadTickBlock(ws)                    ' 古い順に正規化された配列
    If Not IsArray(b) Then Exit Function
    n = UBound(b, 1)
    If n < 1 Then Exit Function

    '--- 歩み値には気配が無いので、そのときの最良気配をスタンプする ---
    bid = NumOrZero(ws.Range(LIVE_BID).Value)
    ask = NumOrZero(ws.Range(LIVE_ASK).Value)

    '--- 前回の最終ティックがブロックのどこにあるかを探す -------------
    idx = 0
    If Len(anchorKey) > 0 Then
        For i = n To 1 Step -1
            If BlockRowKey(b, i) = anchorKey Then
                idx = i
                Exit For
            End If
        Next i

        If idx = 0 Then
            ' ブロックが一周してしまい前回位置を見失った＝取りこぼしの可能性
            If Not gapWarned Then
                gapWarned = True
                ResultSheet().Range("X3").Value = _
                    "警告：" & ws.Name & " で歩み値ブロックを追い越しました。" & _
                    "TICK_BLOCK_ROWS を増やすか POLL_MS を短くしてください。"
            End If
        End If
    End If

    '--- 新しい行だけ追記 --------------------------------------------
    For i = idx + 1 To n
        If AppendTickRow(ws, nextRow, b, i, bid, ask) Then added = added + 1
    Next i

    anchorKey = BlockRowKey(b, n)
    FollowTickBlock = added
End Function

'------------------------------------------------------------------
' 歩み値ブロックを読み、（時刻, 約定値, 出来高, 記号）を古い順で返す
'   有効行が無ければ Empty を返す
'------------------------------------------------------------------
Private Function ReadTickBlock(ws As Worksheet) As Variant

    Dim src As Variant
    Dim cols As Long
    Dim i As Long, k As Long, n As Long
    Dim sec As Long
    Dim price As Double, vol As Double
    Dim tmp() As Variant
    Dim outArr() As Variant
    Dim newestFirst As Boolean

    cols = TB_OFF_TIME
    If TB_OFF_PRICE > cols Then cols = TB_OFF_PRICE
    If TB_OFF_VOL > cols Then cols = TB_OFF_VOL
    If TB_OFF_MARK > cols Then cols = TB_OFF_MARK
    cols = cols + 1

    src = ws.Range(TICK_BLOCK_CELL).Resize(TICK_BLOCK_ROWS, cols).Value

    ReDim tmp(1 To TICK_BLOCK_ROWS, 1 To 4)

    For i = 1 To TICK_BLOCK_ROWS
        sec = TickSec(src(i, TB_OFF_TIME + 1))
        If sec >= 0 Then
            price = NumOrZero(src(i, TB_OFF_PRICE + 1))
            If price > 0 Then
                vol = NumOrZero(src(i, TB_OFF_VOL + 1))
                n = n + 1
                tmp(n, 1) = sec
                tmp(n, 2) = price
                tmp(n, 3) = vol
                If TB_OFF_MARK >= 0 Then
                    tmp(n, 4) = CStr2(src(i, TB_OFF_MARK + 1))
                Else
                    tmp(n, 4) = ""
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function

    '--- 並び順の判定 -------------------------------------------------
    Select Case TICK_BLOCK_ORDER
        Case 1: newestFirst = True
        Case 2: newestFirst = False
        Case Else
            If n >= 2 Then newestFirst = (CLng(tmp(1, 1)) > CLng(tmp(n, 1)))
    End Select

    ReDim outArr(1 To n, 1 To 4)
    For i = 1 To n
        k = IIf(newestFirst, n - i + 1, i)
        outArr(i, 1) = tmp(k, 1)
        outArr(i, 2) = tmp(k, 2)
        outArr(i, 3) = tmp(k, 3)
        outArr(i, 4) = tmp(k, 4)
    Next i

    ReadTickBlock = outArr
End Function

'------------------------------------------------------------------
' ブロック1行の識別キー（時刻＋約定値＋出来高）
'------------------------------------------------------------------
Private Function BlockRowKey(b As Variant, ByVal i As Long) As String
    BlockRowKey = CStr(b(i, 1)) & "|" & CStr(b(i, 2)) & "|" & CStr(b(i, 3))
End Function

'------------------------------------------------------------------
' ティックログ最終行から追従用アンカーを作る（途中再開の二重取込防止）
'------------------------------------------------------------------
Private Function SeedAnchorFromLog(ws As Worksheet) As String

    Dim r As Long, sec As Long

    r = LastTickRow(ws)
    If r < TICK_FIRST_ROW Then Exit Function

    sec = TickSec(ws.Cells(r, COL_TIME).Value)
    If sec < 0 Then Exit Function

    SeedAnchorFromLog = CStr(sec) & "|" & _
                        CStr(NumOrZero(ws.Cells(r, COL_PRICE).Value)) & "|" & _
                        CStr(NumOrZero(ws.Cells(r, COL_VOL).Value))
End Function

'------------------------------------------------------------------
' ブロック1行をティックログに追記（15:00〜15:20 の範囲だけ）
'------------------------------------------------------------------
Private Function AppendTickRow(ws As Worksheet, ByRef nextRow As Long, _
                               b As Variant, ByVal i As Long, _
                               ByVal bid As Double, ByVal ask As Double) As Boolean

    Dim sec As Long

    sec = CLng(b(i, 1))
    If sec < SecOfText(JUDGE_START) Or sec > SecOfText(JUDGE_END) Then Exit Function
    If nextRow > ws.Rows.Count - 1 Then Exit Function

    With ws
        .Cells(nextRow, COL_TIME).Value = TimeSerial(sec \ 3600, (sec \ 60) Mod 60, sec Mod 60)
        .Cells(nextRow, COL_TIME).NumberFormatLocal = "hh:mm:ss"
        .Cells(nextRow, COL_PRICE).Value = b(i, 2)
        .Cells(nextRow, COL_VOL).Value = b(i, 3)
        .Cells(nextRow, COL_MARK).Value = b(i, 4)
        If bid > 0 Then .Cells(nextRow, COL_BID).Value = bid
        If ask > 0 Then .Cells(nextRow, COL_ASK).Value = ask
    End With

    nextRow = nextRow + 1
    AppendTickRow = True
End Function

'------------------------------------------------------------------
' 歩み値ブロックを今すぐ一括取込（ボタン）
'   記録を回していないときの取りこぼし回収・検証用
'------------------------------------------------------------------
Public Sub ImportTickBlockNow()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, nextRow As Long, total As Long
    Dim anchorKey As String
    Dim dummy As Boolean

    Set shts = TargetSheets()
    If shts.Count = 0 Then Exit Sub

    Application.ScreenUpdating = False
    For i = 1 To shts.Count
        Set ws = shts(i)
        nextRow = LastTickRow(ws) + 1
        If nextRow < TICK_FIRST_ROW Then nextRow = TICK_FIRST_ROW
        anchorKey = SeedAnchorFromLog(ws)
        total = total + FollowTickBlock(ws, anchorKey, nextRow, dummy)
    Next i
    RefreshQuotes shts
    Application.ScreenUpdating = True

    MsgBox total & " 件のティックを取り込みました。", vbInformation, "歩み値の取込"
End Sub

'------------------------------------------------------------------
' 歩み値ブロックの読み取り状態を確認する（ボタン）
'   列の割り当てが合っているかを目視確認するための診断です
'------------------------------------------------------------------
Public Sub ShowTickBlockInfo()

    Dim ws As Worksheet
    Dim b As Variant
    Dim n As Long
    Dim msg As String

    Dim shts As Collection

    Set ws = ActiveSheet
    If ws.Name = RESULT_SHEET Then
        Set shts = TargetSheets()
        If shts.Count = 0 Then
            MsgBox "銘柄シートがありません。", vbExclamation, "歩み値ブロックの確認"
            Exit Sub
        End If
        Set ws = shts(1)
    End If

    b = ReadTickBlock(ws)

    msg = "シート : " & ws.Name & vbCrLf & _
          "ブロック : " & TICK_BLOCK_CELL & " から " & TICK_BLOCK_ROWS & " 行" & vbCrLf & vbCrLf

    If Not IsArray(b) Then
        msg = msg & "有効なティックが1件も読めませんでした。" & vbCrLf & vbCrLf & _
              "・" & TICK_BLOCK_CELL & " に歩み値のRSS数式が入っているか" & vbCrLf & _
              "・TB_OFF_TIME / TB_OFF_PRICE / TB_OFF_VOL / TB_OFF_MARK の列オフセット" & vbCrLf & _
              "を確認してください。"
        MsgBox msg, vbExclamation, "歩み値ブロックの確認"
        Exit Sub
    End If

    n = UBound(b, 1)
    msg = msg & "読めたティック数 : " & n & vbCrLf & _
          "最古 : " & SecText(CLng(b(1, 1))) & "  " & b(1, 2) & "  " & b(1, 3) & "  " & b(1, 4) & vbCrLf & _
          "最新 : " & SecText(CLng(b(n, 1))) & "  " & b(n, 2) & "  " & b(n, 3) & "  " & b(n, 4) & vbCrLf & vbCrLf & _
          "この並びで「最古→最新」になっていれば設定は正しいです。"

    MsgBox msg, vbInformation, "歩み値ブロックの確認"
End Sub

Private Function SecText(ByVal sec As Long) As String
    SecText = Format$(TimeSerial(sec \ 3600, (sec \ 60) Mod 60, sec Mod 60), "hh:mm:ss")
End Function

Private Function CStr2(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CStr2 = Trim$(CStr(v))
End Function

'==================================================================
' モード2 : 現在値・出来高ポーリング（フォールバック）
'==================================================================

'------------------------------------------------------------------
' 1銘柄を1回ポーリングし、出来高が増えていれば1行追記する
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
    ResultSheet().Range("X1").Value = "自動開始を予約しました（" & AUTO_ARM_TIME & "）"
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
            ws.Range(ws.Cells(TICK_FIRST_ROW, COL_TIME), ws.Cells(lastRow, COL_MARK)).ClearContents
        End If
        ws.Range(RES_TOP).Offset(0, 1).Resize(15, 1).ClearContents
        ws.Range(RES_TOP).Offset(0, 1).Interior.Pattern = xlNone
    Next i

    With ResultSheet()
        .Range(.Cells(ROW_HEADER + 1, 2), .Cells(ROW_HEADER + 200, 18)).ClearContents
        .Range(.Cells(ROW_HEADER + 1, 2), .Cells(ROW_HEADER + 200, 18)).Interior.Pattern = xlNone
        .Range("X1:X3").ClearContents
    End With
    Application.ScreenUpdating = True
End Sub
