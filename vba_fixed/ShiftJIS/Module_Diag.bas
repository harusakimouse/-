Attribute VB_Name = "Module_Diag"
Option Explicit

'==================================================================
' Module_Diag : RSS 疎通テスト（切り分け）
'
'  「記録を始めると配信が止まる」のか「そもそも配信が来ていない」のかを
'  自動で切り分けます。2段階で同じものを観測するだけの単純なテストです。
'
'   第1段階（アイドル観測 10秒）
'       Application.OnTime で1秒ごとに歩み値の先頭と現在値を見る。
'       この間 Excel はアイドルなので、配信が生きていれば必ず変化します。
'
'   第2段階（ループ占有 10秒）
'       旧実装と同じ「Do While + DoEvents + Sleep」で10秒回しながら
'       同じものを見る。ここで変化が止まるなら、ループが RSS を
'       締め出していたことになります。
'
'  ※ ザラバ中（約定が流れている時間帯）に実行してください。
'     昼休みや引け後は、どちらの段階でも変化が0になり判定できません。
'==================================================================

Private Const DIAG_SECONDS As Long = 10

Private mDgWs         As Worksheet
Private mDgStep       As Long
Private mDgPrev       As String
Private mDgIdleChange As Long
Private mDgRunning    As Boolean

'------------------------------------------------------------------
' 疎通テストの開始（ボタン）
'------------------------------------------------------------------
Public Sub RunFeedDiagnosis()

    Dim shts As Collection
    Dim note As String

    If BlockedWhileLogging("RSS 疎通テスト") Then Exit Sub
    If mDgRunning Then
        MsgBox "疎通テストを実行中です。" & DIAG_SECONDS * 2 & " 秒ほどお待ちください。", _
               vbInformation, "RSS 疎通テスト"
        Exit Sub
    End If

    Set shts = TargetSheets()
    If shts.Count = 0 Then
        MsgBox "銘柄シートがありません。B3 に証券コードを入れてください。", vbExclamation, "RSS 疎通テスト"
        Exit Sub
    End If
    Set mDgWs = shts(1)

    note = MarketHoursNote()

    If MsgBox("RSS 疎通テストを実行します（約 " & DIAG_SECONDS * 2 & " 秒）。" & vbCrLf & vbCrLf & _
              "　対象シート : " & mDgWs.Name & "（" & mDgWs.Cells(ROW_CODE, COL_CODE).Value & "）" & vbCrLf & _
              "　第1段階 : Excel をアイドルにして " & DIAG_SECONDS & " 秒観測" & vbCrLf & _
              "　第2段階 : 旧方式のループで占有して " & DIAG_SECONDS & " 秒観測" & vbCrLf & vbCrLf & _
              IIf(Len(note) > 0, note & vbCrLf & vbCrLf, "") & _
              "テスト中はブックを触らないでください。開始しますか？", _
              vbOKCancel + vbQuestion, "RSS 疎通テスト") <> vbOK Then Exit Sub

    mDgRunning = True
    mDgStep = 0
    mDgIdleChange = 0
    mDgPrev = FeedSignature(mDgWs)

    ShowProgress "第1段階（アイドル観測）開始…"
    ScheduleDiagStep
End Sub

Private Sub ScheduleDiagStep()

    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!DiagStep"
    If Err.Number <> 0 Then
        Err.Clear
        mDgRunning = False
        MsgBox "疎通テストの予約に失敗しました。", vbExclamation, "RSS 疎通テスト"
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' 第1段階の1回ぶん（Application.OnTime から呼ばれます）
'------------------------------------------------------------------
Public Sub DiagStep()

    Dim sig As String

    If Not mDgRunning Then Exit Sub

    sig = FeedSignature(mDgWs)
    If sig <> mDgPrev Then
        mDgIdleChange = mDgIdleChange + 1
        mDgPrev = sig
    End If

    mDgStep = mDgStep + 1
    ShowProgress "第1段階（アイドル観測） " & mDgStep & "/" & DIAG_SECONDS & _
                 "　更新 " & mDgIdleChange & " 回"

    If mDgStep < DIAG_SECONDS Then
        ScheduleDiagStep
    Else
        RunBusyPhase
    End If
End Sub

'------------------------------------------------------------------
' 第2段階：旧方式のループで Excel を占有しながら観測する
'------------------------------------------------------------------
Private Sub RunBusyPhase()

    Dim t0 As Single
    Dim prev As String, sig As String
    Dim busyChange As Long
    Dim loops As Long

    ShowProgress "第2段階（ループ占有中）… " & DIAG_SECONDS & " 秒お待ちください"

    prev = FeedSignature(mDgWs)
    t0 = Timer

    Do
        DoEvents
        Sleep 15
        sig = FeedSignature(mDgWs)
        If sig <> prev Then
            busyChange = busyChange + 1
            prev = sig
        End If
        loops = loops + 1
        If Timer < t0 Then Exit Do            ' 日付跨ぎ保険
    Loop While (Timer - t0) < DIAG_SECONDS

    mDgRunning = False
    ReportDiag mDgIdleChange, busyChange, loops
End Sub

'------------------------------------------------------------------
' 結果の判定と表示
'------------------------------------------------------------------
Private Sub ReportDiag(ByVal idleChange As Long, ByVal busyChange As Long, ByVal loops As Long)

    Dim verdict As String, advice As String, kind As String
    Dim msg As String
    Dim s As Long

    If idleChange = 0 And busyChange = 0 Then
        verdict = "配信が来ていません"
        kind = "error"
        advice = "アイドル状態でも歩み値・現在値がまったく変化しませんでした。" & vbCrLf & _
                 "・マーケットスピードにログインしているか" & vbCrLf & _
                 "・" & TICK_FORMULA_CELL & " の配信状態（現在：" & Left$(FeedStatusText(mDgWs), 60) & "）" & vbCrLf & _
                 "・いま板が動いている時間帯か（昼休み・引け後は変化しません）" & vbCrLf & _
                 "を確認してから、ザラバ中にもう一度実行してください。"

    ElseIf idleChange > 0 And busyChange = 0 Then
        verdict = "ループが RSS を締め出しています（F-01 確定）"
        kind = "error"
        advice = "アイドル時は " & idleChange & " 回更新されたのに、ループ占有中は1回も更新されませんでした。" & vbCrLf & _
                 "旧実装（Do While + DoEvents で回し続ける記録ループ）が原因です。" & vbCrLf & vbCrLf & _
                 "この修正版は Application.OnTime による1秒ごとの再スケジュール方式なので、" & vbCrLf & _
                 "1回ぶんの取込が終わるたびに Excel が空きます。そのまま本番で使えます。"

    ElseIf idleChange > 0 And busyChange > 0 Then
        verdict = "ループは配信を止めていません"
        kind = "warn"
        advice = "アイドル時 " & idleChange & " 回、ループ占有中 " & busyChange & " 回の更新がありました。" & vbCrLf & _
                 "この環境では VBA ループが RSS を締め出してはいないようです。" & vbCrLf & vbCrLf & _
                 "8月31日の 14:33:23 の停止は、マーケットスピード側の切断" & vbCrLf & _
                 "（自動ログアウト・回線断・端末のスリープなど）だった可能性が高くなります。" & vbCrLf & _
                 "修正版の配信停止検知（" & STALE_LIMIT_SEC & " 秒）が働くので、次に起きたら気付けます。"

    Else
        verdict = "判定できませんでした"
        kind = "warn"
        advice = "アイドル時に変化が無く、ループ占有中だけ変化しました。" & vbCrLf & _
                 "板が動き始めた直後などのタイミングの可能性があります。" & vbCrLf & _
                 "ザラバの流れている時間帯にもう一度実行してください。"
    End If

    s = BlockNewestSec(mDgWs)

    msg = "■ 結果 : " & verdict & vbCrLf & String$(46, "-") & vbCrLf & _
          "第1段階（アイドル " & DIAG_SECONDS & "秒） 更新 " & idleChange & " 回" & vbCrLf & _
          "第2段階（ループ " & DIAG_SECONDS & "秒）  更新 " & busyChange & " 回（" & loops & " 周）" & vbCrLf & vbCrLf & _
          "対象 : " & mDgWs.Name & "（" & mDgWs.Cells(ROW_CODE, COL_CODE).Value & "）" & vbCrLf & _
          "歩み値の最新 : " & IIf(s < 0, "読取不可", SecHMS(s) & "（" & AgoText(SecondsAgo(s)) & "）") & vbCrLf & _
          "配信状態 : " & Left$(FeedStatusText(mDgWs), 60) & vbCrLf & vbCrLf & _
          advice

    WriteRunLog LOGR_WARN, "疎通テスト " & Format$(Now, "hh:mm:ss") & "：" & verdict & _
                           "（アイドル " & idleChange & " 回 / ループ " & busyChange & " 回）"
    ResultSheet().Range(RES_STATUS).Value = "疎通テスト結果：" & verdict
    PaintStatus ResultSheet().Range(RES_STATUS), kind

    MsgBox msg, vbInformation, "RSS 疎通テスト"
End Sub

'------------------------------------------------------------------
' 観測対象の指紋
'   歩み値ブロックの先頭3行（新しい順なので最新側）と現在値・出来高。
'   ループ内で毎回読むので、ブロック全体は読みません。
'------------------------------------------------------------------
Private Function FeedSignature(ws As Worksheet) As String

    Dim v As Variant

    On Error GoTo Bad
    v = ws.Range(TICK_BLOCK_CELL).Offset(2, 0).Resize(3, 3).Value

    FeedSignature = TextOrEmpty(v(1, 1)) & "/" & TextOrEmpty(v(1, 2)) & "/" & TextOrEmpty(v(1, 3)) & _
                    "/" & TextOrEmpty(v(2, 1)) & "/" & TextOrEmpty(v(3, 1)) & _
                    "|" & CStr(NumOrZero(ws.Range(LIVE_PRICE).Value)) & _
                    "|" & CStr(NumOrZero(ws.Range(LIVE_VOL).Value))
    Exit Function
Bad:
    FeedSignature = "?"
End Function

Private Sub ShowProgress(ByVal s As String)

    On Error Resume Next
    ResultSheet().Range(RES_STATUS).Value = "RSS 疎通テスト：" & s
    PaintStatus ResultSheet().Range(RES_STATUS), "run"
End Sub

'------------------------------------------------------------------
' いま板が動く時間帯かどうかの注意書き
'------------------------------------------------------------------
Private Function MarketHoursNote() As String

    Dim s As Long

    s = NowSec()

    If Weekday(Date, vbMonday) >= 6 Then
        MarketHoursNote = "※ 土日です。板が動かないので判定できません。"
    ElseIf s < SecOfText("09:00:00") Or s > SecOfText("15:25:00") Then
        MarketHoursNote = "※ いまはザラバの時間外です。板が動かないので判定できません。"
    ElseIf s >= SecOfText("11:30:00") And s <= SecOfText("12:30:00") Then
        MarketHoursNote = "※ いまは昼休みです。板が動かないので判定できません。"
    End If
End Function

'==================================================================
' 現在の設定を点検して一覧表示する（ボタン）
'==================================================================
Public Sub ShowConfigCheck()

    Dim msg As String
    Dim warn As String
    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, s As Long
    Dim problems As String

    Set shts = TargetSheets()

    msg = "■ 判定窓" & vbCrLf & _
          "　" & JUDGE_START & " ～ " & JUDGE_END & _
          "（1区間 " & (BucketSec() \ 60) & "分" & (BucketSec() Mod 60) & "秒）" & vbCrLf & vbCrLf & _
          "■ 判定パラメータ" & vbCrLf & _
          "　① 連続ティック SEQ_MIN = " & SEQ_MIN & "　同値でリセット = " & RESET_ON_FLAT & vbCrLf & _
          "　② 出来高倍率 VOL_RATIO = " & VOL_RATIO & vbCrLf & _
          "　③ 約定速度 SPEED_MIN = " & SPEED_MIN & vbCrLf & vbCrLf & _
          "■ 取込" & vbCrLf & _
          "　LOG_MODE = " & LOG_MODE & IIf(LOG_MODE = 0, "（自動）", "") & vbCrLf & _
          "　ポーリング間隔 = " & POLL_SEC & " 秒（OnTime 方式）" & vbCrLf & _
          "　配信停止とみなす秒数 = " & STALE_LIMIT_SEC & vbCrLf & _
          "　歩み値ブロック = " & TICK_BLOCK_CELL & " から " & TICK_BLOCK_ROWS & " 行" & vbCrLf & _
          "　並び順 = " & IIf(TICK_BLOCK_ORDER = 1, "新しい順（固定）", IIf(TICK_BLOCK_ORDER = 2, "古い順（固定）", "自動")) & vbCrLf & _
          "　ティック記号 = " & IIf(TB_OFF_MARK >= 0, "使用（列オフセット " & TB_OFF_MARK & "）", "未設定 → ①は前ティック比で判定") & vbCrLf & _
          "　時刻の取得元 = " & IIf(USE_RSS_TIME, "RSS の現在値時刻", "PC 時計") & vbCrLf & vbCrLf & _
          "■ 銘柄シート（" & shts.Count & " 枚）" & vbCrLf

    For i = 1 To shts.Count
        Set ws = shts(i)
        s = BlockNewestSec(ws)
        msg = msg & "　" & ws.Name & " : " & ws.Cells(ROW_CODE, COL_CODE).Value & _
              " / 歩み値 " & IIf(s < 0, "読取不可", SecHMS(s) & "（" & AgoText(SecondsAgo(s)) & "）") & _
              " / ログ " & Application.Max(0, LastTickRow(ws) - TICK_FIRST_ROW + 1) & " 行" & vbCrLf
    Next i

    '--- 問題点 -------------------------------------------------------
    warn = LayoutWarning()
    If Len(warn) > 0 Then problems = problems & Bullet(warn)
    warn = WindowWarning()
    If Len(warn) > 0 Then problems = problems & Bullet(warn)
    warn = DuplicateCodeWarning()
    If Len(warn) > 0 Then problems = problems & Bullet(warn)
    If shts.Count > 0 Then
        If Len(FeedStatusText(shts(1))) = 0 Then
            problems = problems & Bullet(TICK_FORMULA_CELL & " に歩み値の数式が入っていません。" & vbCrLf & _
                       "「歩み値の数式を全シートに設定」を実行してください。")
        End If
    End If

    If Len(problems) > 0 Then
        msg = msg & vbCrLf & "■ 要確認" & vbCrLf & problems
    Else
        msg = msg & vbCrLf & "■ 設定に問題は見つかりませんでした。"
    End If

    MsgBox msg, vbInformation, "設定の点検"
End Sub
