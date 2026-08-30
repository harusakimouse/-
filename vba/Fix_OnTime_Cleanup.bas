Attribute VB_Name = "Fix_OnTime_Cleanup"
Option Explicit
'=============================================================================
' OnTime 予約の一元管理モジュール
'
' 目的: ブックを閉じた後に Excel が勝手にブックを開き直す問題の解消。
'
' 原因: Application.OnTime の予約先は「ブック」ではなく Excel アプリケーション。
'       予約を残したまま閉じると、予約時刻に Excel がマクロを実行するために
'       ブックを自動的に開き直してしまう。
'
' 対策: 予約時刻を「非表示シートに永続保存」し、Workbook_BeforeClose で
'       全件を確実に解除する。モジュールレベル変数に持つと閉じた時点で
'       値が消え、時刻の完全一致が必要な解除に失敗するため使用しない。
'
' 導入手順:
'   1. 本モジュールを標準モジュールとして取り込む
'   2. ThisWorkbook モジュールに Workbook_BeforeClose を追加（下部のコメント参照）
'   3. 既存の Application.OnTime 呼び出しを ScheduleTask / CancelTask に置換
'=============================================================================

Private Const REG_SHEET As String = "_OnTimeReg"   ' 予約台帳（非表示シート）

'-----------------------------------------------------------------------------
' 予約する（Application.OnTime の代わりにこれを呼ぶ）
'-----------------------------------------------------------------------------
Public Sub ScheduleTask(ByVal runAt As Date, ByVal procName As String)
    Application.OnTime runAt, procName
    RegisterTask runAt, procName
End Sub

'-----------------------------------------------------------------------------
' 個別に解除する
'-----------------------------------------------------------------------------
Public Sub CancelTask(ByVal runAt As Date, ByVal procName As String)
    On Error Resume Next
    Application.OnTime runAt, procName, , False
    On Error GoTo 0
    UnregisterTask runAt, procName
End Sub

'-----------------------------------------------------------------------------
' 予約を全件解除する（Workbook_BeforeClose から呼ぶ）
'   台帳を後ろから走査し、解除成否にかかわらず行を削除する。
'   既に発火済みの予約は Application.OnTime の解除でエラーになるが、
'   On Error Resume Next で無視して良い（台帳から消えれば目的は達成）。
'-----------------------------------------------------------------------------
Public Sub CancelAllTasks()
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim runAt As Date
        Dim procName As String

        On Error Resume Next
        runAt = CDate(ws.Cells(i, 1).Value)
        procName = CStr(ws.Cells(i, 2).Value)
        On Error GoTo 0

        If Len(procName) > 0 Then
            On Error Resume Next
            Application.OnTime runAt, procName, , False
            On Error GoTo 0
        End If

        ws.Rows(i).Delete
    Next i
End Sub

'-----------------------------------------------------------------------------
' 予約状況を確認する（デバッグ用・ボタン割当可）
'-----------------------------------------------------------------------------
Public Sub 予約状況を表示()
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)

    If ws Is Nothing Then
        MsgBox "予約はありません。", vbInformation, "OnTime 予約状況"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "予約はありません。", vbInformation, "OnTime 予約状況"
        Exit Sub
    End If

    Dim msg As String
    Dim i As Long
    For i = 2 To lastRow
        msg = msg & Format(CDate(ws.Cells(i, 1).Value), "yyyy/mm/dd hh:nn:ss") & _
              "  →  " & CStr(ws.Cells(i, 2).Value) & vbCrLf
    Next i

    MsgBox "現在の OnTime 予約（" & (lastRow - 1) & " 件）:" & vbCrLf & vbCrLf & msg, _
           vbInformation, "OnTime 予約状況"
End Sub

'=============================================================================
' 内部: 台帳操作
'=============================================================================

Private Sub RegisterTask(ByVal runAt As Date, ByVal procName As String)
    Dim ws As Worksheet
    Set ws = GetRegSheet(True)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Cells(lastRow + 1, 1).Value = CDbl(runAt)   ' 数値で保存し丸め誤差を防ぐ
    ws.Cells(lastRow + 1, 2).Value = procName
End Sub

Private Sub UnregisterTask(ByVal runAt As Date, ByVal procName As String)
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To 2 Step -1
        If CStr(ws.Cells(i, 2).Value) = procName Then
            If CDbl(ws.Cells(i, 1).Value) = CDbl(runAt) Then
                ws.Rows(i).Delete
            End If
        End If
    Next i
End Sub

Private Function GetRegSheet(ByVal createIfMissing As Boolean) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REG_SHEET)
    On Error GoTo 0

    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = REG_SHEET
        ws.Cells(1, 1).Value = "RunAt"
        ws.Cells(1, 2).Value = "Procedure"
        ws.Visible = xlSheetVeryHidden
    End If

    Set GetRegSheet = ws
End Function

'=============================================================================
' 【ThisWorkbook モジュールに追記するコード】
'
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     Call CancelAllTasks
' End Sub
'
'
' 【既存コードの置換例】
'
' ThisWorkbook.Workbook_Open:
'   変更前) Application.OnTime 発火時刻, "日次更新_自動"
'   変更後) ScheduleTask 発火時刻, "日次更新_自動"
'
' Mod_SellExtrac.売抽出タイマー実行:
'   変更前) Application.OnTime NextSellRunTime, "売抽出タイマー実行"
'   変更後) ScheduleTask NextSellRunTime, "売抽出タイマー実行"
'
' Mod_SellExtrac.売抽出自動更新停止:
'   変更前) Application.OnTime NextSellRunTime, "売抽出タイマー実行", , False
'   変更後) CancelAllTasks        ' 台帳から実時刻を引くので確実に解除できる
'
' Mod_RSS接続状態.RSS監視タイマー:
'   変更前) Application.OnTime NextMonitorTime, "RSS監視タイマー"
'   変更後) ScheduleTask NextMonitorTime, "RSS監視タイマー"
'
' Mod_RSS接続状態.RSS監視停止:
'   変更前) Application.OnTime EarliestTime:=NextMonitorTime, _
'                              Procedure:="RSS監視タイマー", Schedule:=False
'   変更後) CancelAllTasks
'=============================================================================
