Attribute VB_Name = "Fix_OnTime_Cleanup"
Option Explicit
'=============================================================================
' Fix_OnTime_Cleanup : OnTime 予約の一元管理モジュール
'
' 目的: ブックを閉じた後に Excel が勝手にブックを開き直す問題の解消。
'
' 原因: Application.OnTime の予約先は「ブック」ではなく Excel アプリケーション。
'       予約を残したまま閉じると、予約時刻に Excel がマクロを実行するために
'       ブックを自動的に開き直してしまう。
'
' 対策: 予約時刻を非表示シート「_OnTimeReg」に永続記録し、
'       Workbook_BeforeClose で全件を確実に解除する。
'       予約時刻をモジュール変数に持つと閉じた時点で値が消え、
'       時刻の完全一致が必要な OnTime の解除に失敗するため使用しない。
'
' 導入手順: docs/恒久対策手順.md を参照
'=============================================================================

Private Const REG_SHEET  As String = "_OnTimeReg"  ' 予約台帳（非表示シート）
Private Const COL_RUNAT  As Long = 1               ' A列: 予約時刻（シリアル値）
Private Const COL_PROC   As Long = 2               ' B列: プロシージャ名
Private Const FIRST_ROW  As Long = 2               ' 1行目はヘッダー

'=============================================================================
' ■ 公開API
'=============================================================================

'-----------------------------------------------------------------------------
' 予約する（Application.OnTime の代わりにこれを呼ぶ）
'   台帳への記録に失敗しても予約自体は成立させる。
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
'   台帳を後ろから走査し、解除の成否にかかわらず行を削除する。
'   既に発火済みの予約は解除でエラーになるが無視して良い
'   （台帳から消えれば目的は達成される）。
'-----------------------------------------------------------------------------
Public Sub CancelAllTasks()
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)
    If ws Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Cleanup

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_RUNAT).End(xlUp).Row

    Dim i As Long
    For i = lastRow To FIRST_ROW Step -1
        Dim serial   As Double
        Dim procName As String

        serial = 0
        procName = vbNullString

        On Error Resume Next
        serial = CDbl(ws.Cells(i, COL_RUNAT).Value)
        procName = CStr(ws.Cells(i, COL_PROC).Value)
        On Error GoTo Cleanup

        If Len(procName) > 0 And serial > 0 Then
            On Error Resume Next
            Application.OnTime CDate(serial), procName, , False
            On Error GoTo Cleanup
        End If

        ws.Rows(i).Delete
    Next i

Cleanup:
    Application.EnableEvents = prevEvents
End Sub

'-----------------------------------------------------------------------------
' 予約状況を確認する（動作確認用・ボタン割当可）
'-----------------------------------------------------------------------------
Public Sub 予約状況を表示()
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)

    Dim lastRow As Long
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, COL_RUNAT).End(xlUp).Row
    End If

    If ws Is Nothing Or lastRow < FIRST_ROW Then
        MsgBox "現在 OnTime 予約はありません。" & vbCrLf & _
               "この状態でブックを閉じれば勝手に起動しません。", _
               vbInformation, "OnTime 予約状況"
        Exit Sub
    End If

    Dim msg As String
    Dim i As Long
    For i = FIRST_ROW To lastRow
        msg = msg & Format(CDate(CDbl(ws.Cells(i, COL_RUNAT).Value)), "yyyy/mm/dd hh:nn:ss") & _
              "  →  " & CStr(ws.Cells(i, COL_PROC).Value) & vbCrLf
    Next i

    MsgBox "現在の OnTime 予約（" & (lastRow - FIRST_ROW + 1) & " 件）:" & vbCrLf & vbCrLf & _
           msg & vbCrLf & _
           "これらはブックを閉じるときに自動解除されます。", _
           vbInformation, "OnTime 予約状況"
End Sub

'-----------------------------------------------------------------------------
' 台帳を強制的に空にする（復旧用）
'   台帳と実際の予約がずれた場合に手動で使う。
'-----------------------------------------------------------------------------
Public Sub 予約台帳をクリア()
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)
    If ws Is Nothing Then
        MsgBox "台帳はまだ作成されていません。", vbInformation, "クリア"
        Exit Sub
    End If

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    ws.Rows(FIRST_ROW & ":" & ws.Rows.Count).ClearContents
    On Error GoTo 0
    Application.EnableEvents = prevEvents

    MsgBox "予約台帳をクリアしました。" & vbCrLf & vbCrLf & _
           "※Excel 本体に残っている予約は消えません。" & vbCrLf & _
           "　完全に消すには Excel を一度終了してください。", _
           vbInformation, "クリア完了"
End Sub

'=============================================================================
' ■ 内部: 台帳操作
'=============================================================================

Private Sub RegisterTask(ByVal runAt As Date, ByVal procName As String)
    Dim ws As Worksheet
    Set ws = GetRegSheet(True)
    If ws Is Nothing Then Exit Sub   ' 台帳を作れなくても予約自体は生きている

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Cleanup

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_RUNAT).End(xlUp).Row
    If lastRow < FIRST_ROW - 1 Then lastRow = FIRST_ROW - 1

    ' シリアル値（Double）で保存する。文字列や書式付き日付にすると
    ' 秒未満が丸められ、解除時の完全一致に失敗する。
    ws.Cells(lastRow + 1, COL_RUNAT).Value = CDbl(runAt)
    ws.Cells(lastRow + 1, COL_PROC).Value = procName

Cleanup:
    Application.EnableEvents = prevEvents
End Sub

Private Sub UnregisterTask(ByVal runAt As Date, ByVal procName As String)
    Dim ws As Worksheet
    Set ws = GetRegSheet(False)
    If ws Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Cleanup

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_RUNAT).End(xlUp).Row

    Dim i As Long
    For i = lastRow To FIRST_ROW Step -1
        If CStr(ws.Cells(i, COL_PROC).Value) = procName Then
            If CDbl(ws.Cells(i, COL_RUNAT).Value) = CDbl(runAt) Then
                ws.Rows(i).Delete
            End If
        End If
    Next i

Cleanup:
    Application.EnableEvents = prevEvents
End Sub

'-----------------------------------------------------------------------------
' 台帳シートを取得（無ければ作成）。
' 作成に失敗した場合（ブック構成の保護など）は Nothing を返し、
' 呼び出し側は処理を継続する。
'-----------------------------------------------------------------------------
Private Function GetRegSheet(ByVal createIfMissing As Boolean) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REG_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set GetRegSheet = ws
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim prevActive As Object

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    Set prevActive = ActiveSheet

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    If Not ws Is Nothing Then
        ws.Name = REG_SHEET
        ws.Cells(1, COL_RUNAT).Value = "RunAt"
        ws.Cells(1, COL_PROC).Value = "Procedure"
        ws.Visible = xlSheetVeryHidden
    End If

    If Not prevActive Is Nothing Then prevActive.Activate
    On Error GoTo 0

    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents

    Set GetRegSheet = ws
End Function
