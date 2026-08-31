Attribute VB_Name = "Module_TickLogger"
Option Explicit

'==================================================================
' Module_TickLogger : 判定窓のティックを各銘柄シートに記録する
'
'  ● 記録は Application.OnTime による1秒ごとの再スケジュール方式です。
'    1回ぶんの取込が終わるたびにマクロを完全に終了し、Excel に制御を返します。
'    RSS（RTD）がシートへ値を書けるのは Excel がアイドルのときだけなので、
'    「Do While + DoEvents」で回し続けると、記録したいデータ自体が届きません。
'
'  ● モード1 … 歩み値（TICK）ブロック追従
'      RSS の歩み値をシート上のブロック（既定 AB4～）に出しておき、
'      前回から増えた行だけを検出してティックログに追記します。
'      位置合わせは「時刻＋その秒に取り込んだ本数」で行うため、
'      同一秒・同値・同数量のティックが連続しても取りこぼしません。
'
'  ● モード2 … 現在値・出来高ポーリング
'      歩み値が使えない環境向け。出来高（当日累計）の増分を1ティックとみなします。
'      1秒に1本しか拾えないので ③SpeedMax は不正確になります。
'
'  ● 配信が止まったら黙って 0 件で終わらせず、必ず警告を出します。
'==================================================================

Public gLogging As Boolean

'--- 記録中の状態 ---
Private mSheets     As Collection
Private mPrevVol()  As Double     ' モード2：前回の当日累計出来高
Private mNextRow()  As Long       ' 次に書き込むティックログ行
Private mWmSec()    As Long       ' モード1：取込済みの最終ティックの秒（-1=未取込）
Private mWmCnt()    As Long       ' モード1：その秒のうち何本を取込済みか
Private mGapWarned() As Boolean
Private mMode       As Long
Private mTickTotal  As Long
Private mStartedAt  As Date
Private mLastTickAt As Date       ' 最後に1本でも取り込めた時刻
Private mBusy       As Boolean
Private mErrCount   As Long
Private mStaleWarned As Boolean
Private mTimeWarned As Boolean
Private mUiCount    As Long

' 画面表示を何回に1回更新するか
'   RSS 関数を含むブックでは1セルの書き込みごとに再計算が走ります。
'   毎秒フル更新すると、その再計算だけで1秒を使い切りかねません。
Private Const UI_EVERY As Long = 3

'==================================================================
' 開始 / 停止
'==================================================================

'------------------------------------------------------------------
' ボタン② から呼ぶ版（確認ダイアログあり）
'------------------------------------------------------------------
Public Sub StartTickLoggingButton()
    StartCore True
End Sub

'------------------------------------------------------------------
' Application.OnTime から呼ぶ版（ダイアログを一切出さない）
'   予約時刻に誰も見ていない状態で MsgBox を出すと、そこで全部止まります。
'------------------------------------------------------------------
Public Sub StartTickLogging()
    StartCore False
End Sub

Private Sub StartCore(ByVal interactive As Boolean)

    Dim i As Long, n As Long
    Dim ws As Worksheet
    Dim res As Worksheet
    Dim warn As String
    Dim feedNote As String

    If gLogging Then Exit Sub

    ClearSched SCHED_ARM

    '--- 設定の整合性チェック ---------------------------------------
    warn = LayoutWarning()
    If Len(warn) > 0 Then
        If interactive Then MsgBox "列レイアウトの設定に問題があります。" & vbCrLf & vbCrLf & warn, _
                                   vbCritical, "設定"
        WriteRunLog LOGR_WARN, "設定エラー：" & warn
        Exit Sub
    End If

    warn = WindowWarning()
    If Len(warn) > 0 Then
        If interactive Then
            If MsgBox("判定窓の設定に問題があります。" & vbCrLf & vbCrLf & warn & vbCrLf & vbCrLf & _
                      "このまま記録を開始しますか？", vbYesNo + vbExclamation, "判定窓の設定") <> vbYes Then
                Exit Sub
            End If
        End If
        WriteRunLog LOGR_WARN, "判定窓：" & Replace$(warn, vbCrLf, " ")
    End If

    Set mSheets = TargetSheets()
    n = mSheets.Count

    If n = 0 Then
        If interactive Then
            MsgBox "記録対象の銘柄シートがありません。" & vbCrLf & _
                   "各シートの B3 に証券コードを入れてから「① 準備」を実行してください。", _
                   vbExclamation, "ティック記録"
        End If
        Set res = ResultSheet()
        res.Range(RES_STATUS).Value = "記録対象の銘柄シートがありません"
        Exit Sub
    End If

    warn = DuplicateCodeWarning()
    If Len(warn) > 0 Then
        If interactive Then MsgBox warn, vbExclamation, "銘柄コードの重複"
        WriteRunLog LOGR_WARN, Replace$(warn, vbCrLf, " ")
    End If

    '--- 状態の初期化 ------------------------------------------------
    ReDim mPrevVol(1 To n)
    ReDim mNextRow(1 To n)
    ReDim mWmSec(1 To n)
    ReDim mWmCnt(1 To n)
    ReDim mGapWarned(1 To n)

    For i = 1 To n
        Set ws = mSheets(i)
        mPrevVol(i) = 0                       ' モード2：最初の1回は基準値取りに使う
        mNextRow(i) = LastTickRow(ws) + 1
        If mNextRow(i) < TICK_FIRST_ROW Then mNextRow(i) = TICK_FIRST_ROW
        SeedWatermarkFromLog ws, mWmSec(i), mWmCnt(i)   ' 途中再開でも二重取込しない
        mGapWarned(i) = False
    Next i

    mTickTotal = 0
    mErrCount = 0
    mStaleWarned = False
    mTimeWarned = False
    mUiCount = 0
    mStartedAt = Now
    mLastTickAt = Now

    '--- 取込方式を決める --------------------------------------------
    mMode = DecideMode(mSheets)

    '--- 配信の生死を確認 --------------------------------------------
    feedNote = FeedReadiness(mSheets)
    If Len(feedNote) > 0 Then
        If interactive Then
            If MsgBox("RSS の配信が確認できません。" & vbCrLf & vbCrLf & feedNote & vbCrLf & vbCrLf & _
                      "このまま記録を開始しますか？" & vbCrLf & _
                      "（この状態では 0 件のまま終わる可能性があります）", _
                      vbYesNo + vbExclamation, "配信状態") <> vbYes Then
                Exit Sub
            End If
        End If
    End If

    '--- 開始 --------------------------------------------------------
    Set res = ResultSheet()
    ClearRunLog
    WriteRunLog LOGR_START, mStartedAt
    WriteRunLog LOGR_MODE, ModeText(mMode)
    WriteRunLog LOGR_FEED, FeedSummary(mSheets)
    If Len(feedNote) > 0 Then WriteRunLog LOGR_WARN, "開始時：" & feedNote
    res.Range(LOG_TOP).Offset(LOGR_START, 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"

    res.Range(RES_STATUS).Value = "記録中… 開始 " & Format$(mStartedAt, "hh:mm:ss") & "（" & ModeText(mMode) & "）"
    PaintStatus res.Range(RES_STATUS), IIf(Len(feedNote) > 0, "warn", "run")
    res.Range(RES_COUNT).Value = "ティック 0 件"

    gLogging = True
    ScheduleNextPoll

    If interactive Then
        MsgBox "記録を開始しました（" & ModeText(mMode) & "）。" & vbCrLf & vbCrLf & _
               "1秒ごとに自動で取り込みます。ブックはこのまま置いておいてください。" & vbCrLf & _
               JUDGE_END & " に自動で判定まで走ります。", vbInformation, "ティック記録"
    End If
End Sub

'------------------------------------------------------------------
' ティック記録 停止（判定はしない）
'   ブックを閉じるときにも呼ばれます。
'------------------------------------------------------------------
Public Sub StopTickLogging()
    If Not gLogging Then Exit Sub
    FinishLogging False
End Sub

'------------------------------------------------------------------
' 記録を止めてすぐ判定（ボタン③）
'------------------------------------------------------------------
Public Sub StopAndJudge()
    If gLogging Then
        FinishLogging True
    Else
        JudgeAll
    End If
End Sub

'------------------------------------------------------------------
' 記録の終了処理
'------------------------------------------------------------------
Private Sub FinishLogging(ByVal doJudge As Boolean)

    Dim res As Worksheet

    gLogging = False
    CancelSched SCHED_POLL, "PollTick"

    On Error Resume Next
    Set res = ResultSheet()
    res.Range(RES_STATUS).Value = "記録停止 " & Format$(Now, "hh:mm:ss") & _
                                  "（" & ModeText(mMode) & " / ティック " & mTickTotal & " 件）"
    PaintStatus res.Range(RES_STATUS), IIf(mTickTotal = 0, "warn", "done")
    WriteRunLog LOGR_END, Now
    res.Range(LOG_TOP).Offset(LOGR_END, 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    If mTickTotal = 0 Then
        WriteRunLog LOGR_WARN, "ティックを1件も記録できませんでした。RSS の配信状態を確認してください。"
    End If
    On Error GoTo 0

    If doJudge Then JudgeAllCore
End Sub

'==================================================================
' ポーリング（Application.OnTime の連鎖）
'==================================================================

Private Sub ScheduleNextPoll()

    Dim t As Date

    If Not gLogging Then Exit Sub

    t = Now + TimeSerial(0, 0, POLL_SEC)

    On Error Resume Next
    Application.OnTime t, MacroRef("PollTick")
    If Err.Number <> 0 Then
        Err.Clear
        gLogging = False
        ClearSched SCHED_POLL
        WriteRunLog LOGR_WARN, "次回ポーリングの予約に失敗しました。記録を停止します。"
    Else
        SaveSched SCHED_POLL, t
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' 1回ぶんの取込。OnTime から呼ばれます。
'   ここを抜けるとマクロは完全に終了し、次の1秒間 Excel は自由になります。
'------------------------------------------------------------------
Public Sub PollTick()

    Dim i As Long
    Dim ws As Worksheet
    Dim got As Long

    If Not gLogging Then Exit Sub
    If mBusy Then Exit Sub

    If mSheets Is Nothing Then
        '   VBE のリセットなどで記録中の状態が失われた場合
        gLogging = False
        ClearSched SCHED_POLL
        Exit Sub
    End If

    mBusy = True
    ClearSched SCHED_POLL
    On Error GoTo Fail

    '--- 判定窓の終了 -------------------------------------------------
    If NowTime() > TimeValue(JUDGE_END) Then
        mBusy = False
        FinishLogging AUTO_JUDGE_ON_STOP
        Exit Sub
    End If

    '--- 判定窓に入るまでは待機（配信の様子だけ見る） ------------------
    If NowTime() >= TimeValue(JUDGE_START) Then

        MaybeSwitchMode

        For i = 1 To mSheets.Count
            Set ws = mSheets(i)
            If mMode = 1 Then
                got = FollowTickBlock(ws, mWmSec(i), mWmCnt(i), mNextRow(i), mGapWarned(i))
            Else
                got = IIf(PollOne(ws, mPrevVol(i), mNextRow(i)), 1, 0)
            End If
            If got > 0 Then
                mTickTotal = mTickTotal + got
                mLastTickAt = Now
            End If
        Next i
    End If

    '--- 画面表示（UI_EVERY 回に1回） ---------------------------------
    mUiCount = mUiCount + 1
    If mUiCount >= UI_EVERY Then
        mUiCount = 0
        ResultSheet().Range(RES_COUNT).Value = "ティック " & mTickTotal & " 件 / 最終 " & Format$(Now, "hh:mm:ss")
        RefreshQuotes mSheets
        WriteRunLog LOGR_FEED, FeedSummary(mSheets)
        If mTickTotal > 0 Then
            WriteRunLog LOGR_LAST, Format$(mLastTickAt, "hh:mm:ss") & "（" & mTickTotal & " 件目）"
        End If
    End If

    WarnIfStale

    mErrCount = 0
    mBusy = False
    ScheduleNextPoll
    Exit Sub

Fail:
    mBusy = False
    mErrCount = mErrCount + 1
    WriteRunLog LOGR_WARN, "ポーリングエラー " & Err.Number & " : " & Err.Description & _
                           "（連続 " & mErrCount & " 回目 / " & Format$(Now, "hh:mm:ss") & "）"
    If mErrCount >= 5 Then
        FinishLogging False
        MsgBox "記録中にエラーが続いたため停止しました。" & vbCrLf & _
               Err.Number & " : " & Err.Description, vbCritical, "ティック記録"
    Else
        ScheduleNextPoll
    End If
End Sub

'------------------------------------------------------------------
' まだ1件も取れていないうちは、取込方式の見直しを許す
'   （14:59:30 に開始した時点では歩み値がまだ来ていないことがあるため）
'------------------------------------------------------------------
Private Sub MaybeSwitchMode()

    Dim newMode As Long

    If mTickTotal > 0 Then Exit Sub
    If LOG_MODE <> 0 Then Exit Sub

    newMode = DecideMode(mSheets)
    If newMode = mMode Then Exit Sub

    mMode = newMode
    ReseedAll
    WriteRunLog LOGR_MODE, ModeText(mMode) & "（" & Format$(Now, "hh:mm:ss") & " に切替）"
    ResultSheet().Range(RES_STATUS).Value = "記録中… " & ModeText(mMode) & " に切替 " & Format$(Now, "hh:mm:ss")
End Sub

Private Sub ReseedAll()

    Dim i As Long
    Dim ws As Worksheet

    For i = 1 To mSheets.Count
        Set ws = mSheets(i)
        mNextRow(i) = LastTickRow(ws) + 1
        If mNextRow(i) < TICK_FIRST_ROW Then mNextRow(i) = TICK_FIRST_ROW
        SeedWatermarkFromLog ws, mWmSec(i), mWmCnt(i)
        mPrevVol(i) = 0
    Next i
End Sub

'------------------------------------------------------------------
' 配信が止まっていないか見張る
'------------------------------------------------------------------
Private Sub WarnIfStale()

    Dim quiet As Long
    Dim note As String

    If NowTime() < TimeValue(JUDGE_START) Then Exit Sub

    quiet = DateDiff("s", mLastTickAt, Now)
    If quiet <= STALE_LIMIT_SEC Then
        If mStaleWarned Then
            mStaleWarned = False
            ResultSheet().Range(RES_STATUS).Value = "記録中… 開始 " & Format$(mStartedAt, "hh:mm:ss") & _
                                                    "（" & ModeText(mMode) & "）"
            PaintStatus ResultSheet().Range(RES_STATUS), "run"
            WriteRunLog LOGR_WARN, "配信が回復しました（" & Format$(Now, "hh:mm:ss") & "）"
        End If
        Exit Sub
    End If

    If mStaleWarned Then Exit Sub          ' 毎秒書き直さない（再計算がもったいない）

    note = quiet & " 秒間ティックがありません（" & Format$(Now, "hh:mm:ss") & "）。" & FeedSummary(mSheets)

    WriteRunLog LOGR_WARN, note
    ResultSheet().Range(RES_STATUS).Value = "【警告】配信停止の疑い " & Format$(Now, "hh:mm:ss") & "（" & ModeText(mMode) & "）"
    PaintStatus ResultSheet().Range(RES_STATUS), "warn"
    mStaleWarned = True
End Sub

'==================================================================
' 取込方式の判定と配信状態
'==================================================================

Public Function ModeText(ByVal m As Long) As String
    ModeText = IIf(m = 1, "歩み値追従", "出来高ポーリング")
End Function

'------------------------------------------------------------------
' 歩み値が「生きている」かどうかで方式を決める
'   読めるかどうかではなく、新しいかどうかで判定します。
'------------------------------------------------------------------
Private Function DecideMode(shts As Collection) As Long

    Dim i As Long

    If LOG_MODE <> 0 Then
        DecideMode = LOG_MODE
        Exit Function
    End If

    For i = 1 To shts.Count
        If BlockIsFresh(shts(i)) Then
            DecideMode = 1
            Exit Function
        End If
    Next i

    DecideMode = 2
End Function

' 歩み値ブロックの最新ティックの秒（読めなければ -1）
Public Function BlockNewestSec(ws As Worksheet) As Long

    Dim b As Variant

    BlockNewestSec = -1
    b = ReadTickBlock(ws)
    If Not IsArray(b) Then Exit Function
    BlockNewestSec = CLng(b(UBound(b, 1), 1))
End Function

' 歩み値ブロックが STALE_LIMIT_SEC 秒以内か
Public Function BlockIsFresh(ws As Worksheet) As Boolean

    Dim s As Long

    s = BlockNewestSec(ws)
    If s < 0 Then Exit Function
    BlockIsFresh = (SecondsAgo(s) <= STALE_LIMIT_SEC)
End Function

' 「0時からの経過秒」が今から何秒前か
Public Function SecondsAgo(ByVal sec As Long) As Long

    Dim d As Long

    d = NowSec() - sec
    If d < 0 Then d = d + 86400
    SecondsAgo = d
End Function

' 配信状態セル（RssTickList の戻り値）
Public Function FeedStatusText(ws As Worksheet) As String
    FeedStatusText = TextOrEmpty(ws.Range(TICK_FORMULA_CELL).Value)
End Function

' 配信状態の要約（ログ表示用）
Public Function FeedSummary(shts As Collection) As String

    Dim ws As Worksheet
    Dim s As Long

    If shts.Count = 0 Then Exit Function
    Set ws = shts(1)

    s = BlockNewestSec(ws)
    FeedSummary = ws.Name & " 歩み値:" & IIf(s < 0, "読取不可", SecHMS(s) & "（" & AgoText(SecondsAgo(s)) & "）") & _
                  " / 現在値:" & Format$(NumOrZero(ws.Range(LIVE_PRICE).Value), "#,##0") & _
                  " / " & Left$(FeedStatusText(ws), 40)
End Function

'------------------------------------------------------------------
' 開始前の配信チェック。問題なければ空文字。
'------------------------------------------------------------------
Public Function FeedReadiness(shts As Collection) As String

    Dim i As Long
    Dim ws As Worksheet
    Dim liveOk As Boolean, blockOk As Boolean
    Dim st As String
    Dim msg As String

    If shts.Count = 0 Then
        FeedReadiness = "銘柄シートがありません。"
        Exit Function
    End If

    For i = 1 To shts.Count
        Set ws = shts(i)
        If NumOrZero(ws.Range(LIVE_PRICE).Value) > 0 And NumOrZero(ws.Range(LIVE_VOL).Value) > 0 Then liveOk = True
        If BlockIsFresh(ws) Then blockOk = True
    Next i

    Set ws = shts(1)
    st = FeedStatusText(ws)

    If Not liveOk And Not blockOk Then
        msg = "現在値も歩み値も更新されていません。"
    ElseIf mMode = 1 And Not blockOk Then
        msg = "歩み値ブロックが " & STALE_LIMIT_SEC & " 秒以上更新されていません。"
    ElseIf mMode = 2 And Not liveOk Then
        msg = "現在値・出来高が取得できていません。"
    End If

    If InStr(st, "接続待ち") > 0 Or InStr(st, "エラー") > 0 Then
        msg = msg & IIf(Len(msg) > 0, " ", "") & "配信状態：" & st
    End If

    If Len(msg) > 0 Then
        FeedReadiness = msg & vbCrLf & _
                        "マーケットスピードにログインしているか、" & TICK_FORMULA_CELL & " の数式が入っているかを確認してください。"
    End If
End Function

'==================================================================
' モード1 : 歩み値（TICK）ブロック追従
'==================================================================

'------------------------------------------------------------------
' 歩み値ブロックを読み、前回以降に増えた行だけを追記する
'
'   位置合わせは「最後に取り込んだティックの秒（wmSec）」と
'   「その秒のうち何本を取り込んだか（wmCnt）」で行います。
'   歩み値は時刻順に並ぶので、
'       秒 <  wmSec              → 取込済み
'       秒 =  wmSec の先頭 wmCnt 本 → 取込済み
'       それ以外                  → 新規
'   と一意に決まります。値の一致で探さないので、同一秒・同値・同数量の
'   ティックが連続しても取りこぼしも二重取込も起きません。
'
'   戻り値 : 追記したティック数
'------------------------------------------------------------------
Private Function FollowTickBlock(ws As Worksheet, ByRef wmSec As Long, ByRef wmCnt As Long, _
                                 ByRef nextRow As Long, ByRef gapWarned As Boolean) As Long

    Dim b As Variant
    Dim n As Long, i As Long
    Dim sec As Long, seenAtWm As Long
    Dim bid As Double, ask As Double
    Dim startSec As Long, endSec As Long
    Dim aTVP() As Variant, aBA() As Variant, aMk() As Variant
    Dim cnt As Long
    Dim newSec As Long, newCnt As Long
    Dim useMark As Boolean

    b = ReadTickBlock(ws)
    If Not IsArray(b) Then Exit Function
    n = UBound(b, 1)
    If n < 1 Then Exit Function

    startSec = SecOfText(JUDGE_START)
    endSec = SecOfText(JUDGE_END)
    useMark = (TB_OFF_MARK >= 0)

    '--- 取りこぼし検知 ----------------------------------------------
    '   ブロックの最古行が監視点より新しい＝間のティックが押し出された
    If wmSec >= 0 Then
        If CLng(b(1, 1)) > wmSec Then
            If Not gapWarned Then
                gapWarned = True
                WriteRunLog LOGR_WARN, "警告：" & ws.Name & " で歩み値ブロックを追い越しました（" & _
                                       SecHMS(wmSec) & " → " & SecHMS(CLng(b(1, 1))) & "）。" & _
                                       "TICK_SHOW_ROWS を増やすか POLL_SEC を短くしてください。"
            End If
        End If
    End If

    '--- 歩み値には気配が無いので、そのときの最良気配をスタンプする ---
    bid = NumOrZero(ws.Range(LIVE_BID).Value)
    ask = NumOrZero(ws.Range(LIVE_ASK).Value)

    ReDim aTVP(1 To n, 1 To 3)
    ReDim aBA(1 To n, 1 To 2)
    ReDim aMk(1 To n, 1 To 1)

    For i = 1 To n
        sec = CLng(b(i, 1))

        '--- 取込済みかどうか ----------------------------------------
        If wmSec >= 0 Then
            If sec < wmSec Then GoTo SkipRow
            If sec = wmSec Then
                seenAtWm = seenAtWm + 1
                If seenAtWm <= wmCnt Then GoTo SkipRow
            End If
        End If

        '--- 判定窓の中だけ書き込む（窓の外でも「消費済み」にはする） ---
        If sec < startSec Or sec > endSec Then GoTo SkipRow
        If nextRow + cnt > ws.Rows.Count - 1 Then GoTo SkipRow

        cnt = cnt + 1
        aTVP(cnt, 1) = TimeSerial(sec \ 3600, (sec \ 60) Mod 60, sec Mod 60)
        aTVP(cnt, 2) = b(i, 3)                     ' 出来高
        aTVP(cnt, 3) = b(i, 2)                     ' 約定値
        aBA(cnt, 1) = IIf(bid > 0, bid, Empty)
        aBA(cnt, 2) = IIf(ask > 0, ask, Empty)
        If useMark Then aMk(cnt, 1) = b(i, 4)

SkipRow:
    Next i

    '--- 一括書き込み（1ティックずつ書くと再計算で潰れます） ----------
    If cnt > 0 Then
        ws.Cells(nextRow, COL_TIME).Resize(cnt, 3).Value = aTVP
        ws.Cells(nextRow, COL_BID).Resize(cnt, 2).Value = aBA
        If useMark Then ws.Cells(nextRow, COL_MARK).Resize(cnt, 1).Value = aMk
        nextRow = nextRow + cnt
    End If

    '--- 監視点を進める ----------------------------------------------
    newSec = CLng(b(n, 1))
    If newSec >= wmSec Then
        newCnt = 0
        For i = n To 1 Step -1
            If CLng(b(i, 1)) <> newSec Then Exit For
            newCnt = newCnt + 1
        Next i
        wmSec = newSec
        wmCnt = newCnt
    End If

    FollowTickBlock = cnt
End Function

'------------------------------------------------------------------
' 歩み値ブロックを読み、（時刻秒, 約定値, 出来高, 記号）を古い順で返す
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
                    tmp(n, 4) = TextOrEmpty(src(i, TB_OFF_MARK + 1))
                Else
                    tmp(n, 4) = ""
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function

    '--- 並び順 -------------------------------------------------------
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
' ティックログ最終行から監視点を作る（途中再開の二重取込防止）
'------------------------------------------------------------------
Private Sub SeedWatermarkFromLog(ws As Worksheet, ByRef wmSec As Long, ByRef wmCnt As Long)

    Dim lastRow As Long, r As Long, s As Long

    wmSec = -1
    wmCnt = 0

    lastRow = LastTickRow(ws)
    If lastRow < TICK_FIRST_ROW Then Exit Sub

    s = TickSec(ws.Cells(lastRow, COL_TIME).Value)
    If s < 0 Then Exit Sub          ' 読めない行があれば履歴から作らない

    wmSec = s
    For r = lastRow To TICK_FIRST_ROW Step -1
        If TickSec(ws.Cells(r, COL_TIME).Value) <> s Then Exit For
        wmCnt = wmCnt + 1
    Next r
End Sub

'------------------------------------------------------------------
' 歩み値ブロックを今すぐ一括取込（ボタン）
'   記録を回していないときの取りこぼし回収・検証用
'------------------------------------------------------------------
Public Sub ImportTickBlockNow()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, nextRow As Long, total As Long
    Dim wmSec As Long, wmCnt As Long
    Dim dummy As Boolean

    If BlockedWhileLogging("歩み値を今すぐ取込") Then Exit Sub

    Set shts = TargetSheets()
    If shts.Count = 0 Then
        MsgBox "銘柄シートがありません。", vbExclamation, "歩み値の取込"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Fin

    For i = 1 To shts.Count
        Set ws = shts(i)
        nextRow = LastTickRow(ws) + 1
        If nextRow < TICK_FIRST_ROW Then nextRow = TICK_FIRST_ROW
        SeedWatermarkFromLog ws, wmSec, wmCnt
        total = total + FollowTickBlock(ws, wmSec, wmCnt, nextRow, dummy)
    Next i
    RefreshQuotes shts

Fin:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "取込中にエラーが発生しました。" & vbCrLf & Err.Number & " : " & Err.Description, _
               vbCritical, "歩み値の取込"
        Exit Sub
    End If

    MsgBox total & " 件のティックを取り込みました。" & vbCrLf & vbCrLf & _
           FeedSummary(shts), vbInformation, "歩み値の取込"
End Sub

'------------------------------------------------------------------
' 歩み値ブロックの読み取り状態を確認する（ボタン）
'------------------------------------------------------------------
Public Sub ShowTickBlockInfo()

    Dim ws As Worksheet
    Dim b As Variant
    Dim n As Long, s As Long
    Dim msg As String
    Dim shts As Collection

    Set ws = ActiveSheet
    If IsSystemSheet(ws.Name) Then
        Set shts = TargetSheets()
        If shts.Count = 0 Then
            MsgBox "銘柄シートがありません。", vbExclamation, "歩み値ブロックの確認"
            Exit Sub
        End If
        Set ws = shts(1)
    End If

    b = ReadTickBlock(ws)

    msg = "シート : " & ws.Name & vbCrLf & _
          "ブロック : " & TICK_BLOCK_CELL & " から " & TICK_BLOCK_ROWS & " 行" & vbCrLf & _
          "配信状態 : " & FeedStatusText(ws) & vbCrLf & vbCrLf

    If Not IsArray(b) Then
        msg = msg & "有効なティックが1件も読めませんでした。" & vbCrLf & vbCrLf & _
              "・" & TICK_FORMULA_CELL & " に歩み値のRSS数式が入っているか" & vbCrLf & _
              "・マーケットスピードにログインしているか" & vbCrLf & _
              "・TB_OFF_TIME / TB_OFF_PRICE / TB_OFF_VOL の列オフセット" & vbCrLf & _
              "を確認してください。"
        MsgBox msg, vbExclamation, "歩み値ブロックの確認"
        Exit Sub
    End If

    n = UBound(b, 1)
    s = CLng(b(n, 1))

    msg = msg & "読めたティック数 : " & n & "（配信状態行・見出し行・終端記号は自動で読み飛ばし）" & vbCrLf & _
          "最古 : " & SecHMS(CLng(b(1, 1))) & "  " & b(1, 2) & "  " & b(1, 3) & vbCrLf & _
          "最新 : " & SecHMS(s) & "  " & b(n, 2) & "  " & b(n, 3) & vbCrLf & vbCrLf & _
          "最新ティックは " & AgoText(SecondsAgo(s)) & "です。" & vbCrLf & _
          IIf(SecondsAgo(s) <= STALE_LIMIT_SEC, _
              "→ 配信は生きています（歩み値追従モードで動きます）。", _
              "→ " & STALE_LIMIT_SEC & " 秒以上止まっています。ザラバ中ならRSSの接続を確認してください。")

    MsgBox msg, vbInformation, "歩み値ブロックの確認"
End Sub

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
    Dim sec As Long
    Dim aTVP(1 To 1, 1 To 3) As Variant
    Dim aBA(1 To 1, 1 To 2) As Variant

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

    t = TickTime(ws)
    sec = Hour(t) * 3600& + Minute(t) * 60& + Second(t)
    If sec < SecOfText(JUDGE_START) Or sec > SecOfText(JUDGE_END) Then Exit Function
    If nextRow > ws.Rows.Count - 1 Then Exit Function

    bid = NumOrZero(ws.Range(LIVE_BID).Value)
    ask = NumOrZero(ws.Range(LIVE_ASK).Value)

    aTVP(1, 1) = t
    aTVP(1, 2) = d
    aTVP(1, 3) = price
    aBA(1, 1) = IIf(bid > 0, bid, Empty)
    aBA(1, 2) = IIf(ask > 0, ask, Empty)

    ws.Cells(nextRow, COL_TIME).Resize(1, 3).Value = aTVP
    ws.Cells(nextRow, COL_BID).Resize(1, 2).Value = aBA

    nextRow = nextRow + 1
    PollOne = True
End Function

'------------------------------------------------------------------
' ティックの時刻
'   RSS の「現在値時刻」が hh:mm（秒なし）で返る環境があります。
'   そのまま使うと全ティックの秒が :00 になり、③約定速度が
'   「同一分内の約定回数」になって壊れるので、検出して PC 時計に戻します。
'------------------------------------------------------------------
Private Function TickTime(ws As Worksheet) As Date

    Dim raw As Variant
    Dim s As String
    Dim d As Date
    Dim ok As Boolean

    If USE_RSS_TIME Then
        raw = ws.Range(LIVE_TIME).Value
        s = TextOrEmpty(raw)

        If VarType(raw) = vbDate Then
            d = CDate(raw)
            ok = True
        ElseIf IsNumeric(raw) Then
            d = CDate(CDbl(raw))
            ok = True
        ElseIf Len(s) > 0 Then
            On Error Resume Next
            d = CDate(s)
            ok = (Err.Number = 0)
            Err.Clear
            On Error GoTo 0
            '--- "14:33" のように秒が無い形なら使わない ---
            If ok Then
                If UBound(Split(s, ":")) < 2 Then
                    ok = False
                    If Not mTimeWarned Then
                        mTimeWarned = True
                        WriteRunLog LOGR_WARN, "RSS の現在値時刻に秒がありません（" & s & "）。PC 時計に切り替えました。"
                    End If
                End If
            End If
        End If

        If ok Then
            TickTime = TimeSerial(Hour(d), Minute(d), Second(d))
            Exit Function
        End If
    End If

    TickTime = TimeSerial(Hour(Now), Minute(Now), Second(Now))
End Function

'==================================================================
' 自動開始の予約
'==================================================================

Public Sub ArmAutoRun()
    ArmAutoRunCore True
End Sub

' ブックを開いたときに呼ぶ版（メッセージを出さない）
Public Sub ArmAutoRunSilent()
    If AUTO_ARM_ON_OPEN Then ArmAutoRunCore False
End Sub

Private Sub ArmAutoRunCore(ByVal showMsg As Boolean)

    Dim armAt As Date
    Dim res As Worksheet

    DisarmAutoRun
    If gLogging Then Exit Sub

    Set res = ResultSheet()
    armAt = Date + TimeValue(AUTO_ARM_TIME)

    If Now < armAt Then
        '--- 予約時刻前：その時刻に開始 ------------------------------
        If ScheduleAt(armAt, "StartTickLogging", SCHED_ARM) Then
            res.Range(RES_STATUS).Value = "自動開始を予約しました（" & AUTO_ARM_TIME & "）"
            PaintStatus res.Range(RES_STATUS), "run"
            If showMsg Then
                MsgBox AUTO_ARM_TIME & " に自動で記録を開始します。" & vbCrLf & _
                       "そのまま置いておいてください。", vbInformation, "自動実行の予約"
            End If
        End If

    ElseIf NowTime() < TimeValue(JUDGE_END) Then
        '--- 予約時刻は過ぎたが判定窓の終了前：すぐ開始 ---------------
        If ScheduleAt(Now + TimeSerial(0, 0, 3), "StartTickLogging", SCHED_ARM) Then
            res.Range(RES_STATUS).Value = "予約時刻を過ぎていたため、記録を開始します"
            PaintStatus res.Range(RES_STATUS), "run"
            If showMsg Then
                MsgBox "予約時刻（" & AUTO_ARM_TIME & "）を過ぎていたので、" & vbCrLf & _
                       "このまま記録を開始します。", vbInformation, "自動実行の予約"
            End If
        End If

    Else
        '--- 本日の判定時間は終了 ------------------------------------
        res.Range(RES_STATUS).Value = "本日の判定時間（～" & JUDGE_END & "）は終了しています"
        PaintStatus res.Range(RES_STATUS), "done"
        If showMsg Then
            MsgBox "本日の判定時間（～" & JUDGE_END & "）は終了しています。", _
                   vbInformation, "自動実行の予約"
        End If
    End If
End Sub

Public Sub DisarmAutoRun()
    CancelSched SCHED_ARM, "StartTickLogging"
End Sub

'------------------------------------------------------------------
' ブックを開いた直後に、前回の残骸を掃除する
'   （前セッションの予約はブックを閉じた時点で無効になっています）
'------------------------------------------------------------------
Public Sub RecoverSchedules()

    On Error Resume Next
    '   同じ Excel セッション内でブックを閉じて開き直した場合、前回の予約が
    '   生き残っていることがあります。忘れるだけでなく解除まで行います。
    CancelSched SCHED_ARM, "StartTickLogging"
    CancelSched SCHED_POLL, "PollTick"
    gLogging = False
    On Error GoTo 0
End Sub

'==================================================================
' Application.OnTime の予約管理
'   予約時刻をシートにも残します。VBA の状態がリセットされて
'   モジュール変数が消えても、ここから復元して確実に解除できます。
'==================================================================

' マクロ名をブック名で修飾する
'   ブック名に空白があると、別ブックがアクティブなときに解決に失敗します。
Private Function MacroRef(ByVal procName As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!" & procName
End Function

Private Function ScheduleAt(ByVal t As Date, ByVal procName As String, ByVal addr As String) As Boolean

    On Error GoTo Fail
    Application.OnTime t, MacroRef(procName)
    SaveSched addr, t
    ScheduleAt = True
    Exit Function
Fail:
    MsgBox "予約に失敗しました。" & vbCrLf & Err.Number & " : " & Err.Description, vbExclamation, "自動実行の予約"
End Function

Private Sub SaveSched(ByVal addr As String, ByVal t As Date)

    On Error Resume Next
    With ResultSheet().Range(addr)
        .NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        .Value = t
    End With
End Sub

Private Function LoadSched(ByVal addr As String) As Date

    Dim v As Variant

    On Error Resume Next
    v = ResultSheet().Range(addr).Value
    If IsDate(v) Then LoadSched = CDate(v)
End Function

Private Sub ClearSched(ByVal addr As String)
    On Error Resume Next
    ResultSheet().Range(addr).ClearContents
End Sub

Private Sub CancelSched(ByVal addr As String, ByVal procName As String)

    Dim t As Date

    t = LoadSched(addr)
    If t > 0 Then
        On Error Resume Next
        Application.OnTime t, MacroRef(procName), , False
        Err.Clear
        On Error GoTo 0
    End If
    ClearSched addr
End Sub

'==================================================================
' 状態セルの色付け
'==================================================================
Public Sub PaintStatus(cel As Range, ByVal kind As String)

    On Error Resume Next
    With cel
        .Font.Bold = True
        Select Case kind
            Case "run"
                .Interior.Color = RGB(226, 239, 218)
                .Font.Color = RGB(0, 97, 0)
            Case "warn"
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 87, 0)
            Case "error"
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
            Case Else
                .Interior.Pattern = xlNone
                .Font.Color = RGB(64, 64, 64)
        End Select
    End With
End Sub

'==================================================================
' ティックログのクリア（ボタン⑤）
'==================================================================
Public Sub ClearAllTicks()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    If BlockedWhileLogging("ティックログ消去") Then Exit Sub

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
        ws.Range(RES_TOP).Offset(0, 1).Resize(RES_ROWS, 1).ClearContents
        ws.Range(RES_TOP).Offset(0, 1).Interior.Pattern = xlNone
    Next i

    With ResultSheet()
        .Range(.Cells(ROW_HEADER + 1, SUM_COL_FIRST), .Cells(ROW_HEADER + SUM_MAX_ROWS, SUM_COL_LAST)).ClearContents
        .Range(.Cells(ROW_HEADER + 1, SUM_COL_FIRST), .Cells(ROW_HEADER + SUM_MAX_ROWS, SUM_COL_LAST)).Interior.Pattern = xlNone
        .Range(RES_STATUS & ":" & RES_BREAK).ClearContents
        PaintStatus .Range(RES_STATUS), ""
    End With
    ClearRunLog
    Application.ScreenUpdating = True
End Sub
