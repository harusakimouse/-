VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    ' 判定サマリーを開く
    On Error Resume Next
    ThisWorkbook.Worksheets(RESULT_SHEET).Activate
    On Error GoTo 0

    ' 前セッションの予約の残骸を掃除する
    '   ブックを閉じた時点で Application.OnTime の予約は無効になっているので、
    '   内部管理セルに残っている時刻を消すだけです。
    RecoverSchedules

    ' 自動開始を予約（Module_Config の AUTO_ARM_ON_OPEN = True のとき）
    '   予約時刻前     → その時刻に記録開始
    '   予約時刻を過ぎ → 判定窓の終了前ならすぐ開始
    '   判定窓の後     → 何もしない
    ArmAutoRunSilent

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' 記録ループとタイマー予約を確実に止める
    '   ここでは判定を走らせません。閉じかけのブックに書き込むと落ちます。
    StopTickLogging
    DisarmAutoRun

End Sub
