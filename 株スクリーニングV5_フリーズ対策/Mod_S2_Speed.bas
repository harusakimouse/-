Attribute VB_Name = "Mod_S2_Speed"
Option Explicit
'==============================================================================
' S2 フリーズ対策モジュール  Ver 1.0
'
'  やること
'   1. 重い処理の間だけ Excel を「手動計算」にする（フリーズの一番の原因を止める）
'   2. ダイアログを必ず一番手前に出す（裏に隠れて固まって見えるのを防ぐ）
'   3. 進み具合をステータスバーに出し、DoEvents で「応答なし」を防ぐ
'
'  使い方
'   ・このモジュールを丸ごと取り込む（VBE → ファイル → ファイルのインポート）
'   ・そのうえで「差し替えコード.txt」の4か所を差し替える
'==============================================================================

Private sCalc  As Long
Private sEvt   As Boolean
Private sScr   As Boolean
Private sAlt   As Boolean
Private sDepth As Long


' 重い処理の入口で呼ぶ。入れ子で呼んでも一番外側だけが効く。
Public Sub S2_FastOn()
    sDepth = sDepth + 1
    If sDepth > 1 Then Exit Sub

    On Error Resume Next
    sCalc = Application.Calculation
    sEvt = Application.EnableEvents
    sScr = Application.ScreenUpdating
    sAlt = Application.DisplayAlerts

    Application.Calculation = xlCalculationManual      '★これが最重要
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Application.Cursor = xlWait
    On Error GoTo 0
End Sub


' 重い処理の出口で必ず呼ぶ（エラーで抜けるときも呼ぶこと）
Public Sub S2_FastOff()
    If sDepth > 0 Then sDepth = sDepth - 1
    If sDepth > 0 Then Exit Sub

    On Error Resume Next
    Application.Cursor = xlDefault
    Application.DisplayAlerts = sAlt
    Application.ScreenUpdating = True
    Application.EnableEvents = sEvt
    If sCalc = 0 Then sCalc = xlCalculationAutomatic
    Application.Calculation = sCalc
    Application.StatusBar = False
    Application.Calculate
    On Error GoTo 0
End Sub


' 途中経過の表示。DoEvents を入れて「応答なし」にさせない。
Public Sub S2_Beat(ByVal txt As String)
    On Error Resume Next
    Application.StatusBar = txt
    DoEvents
    On Error GoTo 0
End Sub


' 必ず手前に出る質問ダイアログ（はい／いいえ）
Public Function S2_Ask(ByVal msg As String, ByVal ico As Long, ByVal ttl As String) As Long
    Dim scr As Boolean
    On Error Resume Next
    scr = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    DoEvents
    AppActivate Application.Caption
    Err.Clear
    On Error GoTo 0

    S2_Ask = MsgBox(msg, ico + vbSystemModal, ttl)

    On Error Resume Next
    Application.Cursor = xlWait
    Application.ScreenUpdating = scr
    On Error GoTo 0
End Function


' 予約（OnTime）を確実に解除するための後始末。ブックを閉じる前に呼ぶ。
Public Sub S2_予約全解除()
    Dim d As Date, i As Long
    On Error Resume Next
    d = Int(Now)
    For i = 0 To 1439
        Application.OnTime d + TimeSerial(0, i, 0), "S2_DailyRun", , False
        Application.OnTime d + 1 + TimeSerial(0, i, 0), "S2_DailyRun", , False
    Next i
    Err.Clear
    On Error GoTo 0
    MsgBox "自動更新の予約をすべて解除しました。", vbInformation, "S2"
End Sub
