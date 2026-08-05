Attribute VB_Name = "Mod_RSS接続状態"
Option Explicit

' ================================================================
' RSS接続状態表示  【修正版】
'
' 【旧版のバグ】
'   CheckRSSStatus が closeWs.Cells(r, 5) ＝ E列 を見ていた。
'   E列は「日次更新が書き込んだ保存済みの値」であって RSS のライブ値ではない。
'   その結果:
'     ・RSS が完全に切れていても E列に過去値があれば「接続中」と誤表示
'     ・逆に E列が空になった直後は RSS が生きていても「切断」と誤表示
'
' 【修正版】
'   ・見るのは C列（＝ _xll.RssMarket の結果）。
'   ・#NAME? を「アドイン未ロード」として個別に報告する。
'   ・計算モードが手動なら、それ自体を切断原因として報告する。
'   ・サンプルは TOPIX(5) + 先頭 8 銘柄に拡大（一部銘柄だけ止まる場合の検出）。
' ================================================================

Private Const STATUS_CELL_COL As Long = 18  ' R列
Private Const STATUS_CELL_ROW As Long = 1
Private Const MONITOR_INTERVAL_SEC As Long = 30

Private Monitoring As Boolean
Private NextMonitorTime As Double


' ================================================================
' オンデマンドチェック (ボタン割当推奨)
' ================================================================
Public Sub RSS接続状態チェック()
    Dim status As String, details As String
    Dim isConnected As Boolean

    isConnected = CheckRSSStatus(status, details)
    UpdateStatusCell isConnected, status

    Dim msg As String
    msg = "【RSS接続状態 診断結果】" & vbCrLf & _
          String(30, "-") & vbCrLf & _
          "状態: " & status & vbCrLf & vbCrLf & _
          details & vbCrLf & _
          String(30, "-") & vbCrLf

    If isConnected Then
        msg = msg & "→ RSS式が正常に動作しています。"
    Else
        msg = msg & "→ 次の順で確認してください。" & vbCrLf & _
              "  1. MarketSpeed II を起動・ログイン（Excel より先に）" & vbCrLf & _
              "  2. 数式 → 計算方法の設定 → 自動" & vbCrLf & _
              "  3. RSS環境診断 マクロで詳細を確認"
    End If

    MsgBox msg, IIf(isConnected, vbInformation, vbExclamation), "RSS接続状態"
End Sub


' ================================================================
' 内部: 状態チェック本体  ★C列を見る★
' ================================================================
Private Function CheckRSSStatus(ByRef status As String, _
                                ByRef details As String) As Boolean
    Dim closeWs As Worksheet
    On Error Resume Next
    Set closeWs = ThisWorkbook.Sheets("終値")
    On Error GoTo 0

    If closeWs Is Nothing Then
        status = "[ERROR] 終値シートなし"
        details = "終値シートが見つかりません。"
        CheckRSSStatus = False
        Exit Function
    End If

    ' --- 計算モードは先に見る。手動ならその時点で切断扱い ---
    Dim calcNG As Boolean
    calcNG = (Application.Calculation = xlCalculationManual)

    Dim sampleRows As Variant
    sampleRows = Array(5, 6, 7, 8, 9, 10, 11, 12, 13)

    Dim validCount As Long, nameErrCount As Long
    Dim naCount As Long, zeroCount As Long, noFormulaCount As Long
    Dim sampleDetail As String
    sampleDetail = "サンプル値 (C列 = RSSライブ値):" & vbCrLf

    Dim i As Integer
    For i = 0 To UBound(sampleRows)
        Dim r As Long: r = CLng(sampleRows(i))
        Dim code As String: code = CStr(closeWs.Cells(r, 1).Value)
        Dim state As String

        If Not closeWs.Cells(r, 3).HasFormula Then
            state = "数式なし ★RSS式が消えている"
            noFormulaCount = noFormulaCount + 1
        Else
            Dim v As Variant: v = closeWs.Cells(r, 3).Value
            If IsError(v) Then
                Select Case CLng(v)
                    Case xlErrName
                        state = "#NAME? ★アドイン未ロード"
                        nameErrCount = nameErrCount + 1
                    Case xlErrNA
                        state = "#N/A (接続待ち)"
                        naCount = naCount + 1
                    Case Else
                        state = "エラー(" & CLng(v) & ")"
                        naCount = naCount + 1
                End Select
            ElseIf Not IsNumeric(v) Then
                state = "[" & CStr(v) & "] 非数値"
                zeroCount = zeroCount + 1
            ElseIf CDbl(v) <= 0 Then
                state = "0"
                zeroCount = zeroCount + 1
            Else
                state = Format(CDbl(v), "#,##0.##") & " [OK]"
                validCount = validCount + 1
            End If
        End If

        sampleDetail = sampleDetail & "  C" & r & " (" & code & "): " & state & vbCrLf
    Next i

    details = sampleDetail & vbCrLf & _
              "有効: " & validCount & " / " & (UBound(sampleRows) + 1) & vbCrLf & _
              "#NAME?: " & nameErrCount & " / 待ち: " & naCount & _
              " / 0・非数値: " & zeroCount & " / 数式なし: " & noFormulaCount & vbCrLf & _
              "計算モード: " & IIf(calcNG, "★手動★（RTDが更新されません）", "自動")

    If nameErrCount > 0 Then
        status = "[RSS切断] アドイン未ロード"
        CheckRSSStatus = False
    ElseIf calcNG Then
        status = "[RSS停止] 計算モードが手動"
        CheckRSSStatus = False
    ElseIf validCount >= 5 Then
        status = "[RSS接続中]"
        CheckRSSStatus = True
    ElseIf validCount >= 1 Then
        status = "[RSS部分接続] 一部銘柄のみ取得"
        CheckRSSStatus = True
    Else
        status = "[RSS切断]"
        CheckRSSStatus = False
    End If
End Function


' ================================================================
' 自動監視開始 / 停止 / タイマー （旧版と同じ挙動）
' ================================================================
Public Sub RSS監視開始()
    If Monitoring Then
        MsgBox "既に監視中です。停止するには「RSS監視停止」を実行してください。", _
               vbInformation, "監視中"
        Exit Sub
    End If

    Monitoring = True
    NextMonitorTime = Now + TimeSerial(0, 0, MONITOR_INTERVAL_SEC)

    UpdateStatusSilent
    Application.OnTime NextMonitorTime, "'" & ThisWorkbook.Name & "'!RSS監視タイマー"

    MsgBox "RSS接続状態の自動監視を開始しました。" & vbCrLf & _
           "監視間隔: " & MONITOR_INTERVAL_SEC & " 秒" & vbCrLf & _
           "銘柄管理!R1 にリアルタイム状態が表示されます。", _
           vbInformation, "監視開始"
End Sub

Public Sub RSS監視停止()
    If Not Monitoring Then
        MsgBox "監視は実行されていません。", vbInformation, "停止"
        Exit Sub
    End If

    On Error Resume Next
    Application.OnTime EarliestTime:=NextMonitorTime, _
                   Procedure:="'" & ThisWorkbook.Name & "'!RSS監視タイマー", _
                   Schedule:=False
    On Error GoTo 0

    Monitoring = False
    MsgBox "RSS接続状態の自動監視を停止しました。", vbInformation, "停止完了"
End Sub

Public Sub RSS監視タイマー()
    If Not Monitoring Then Exit Sub
    UpdateStatusSilent
    NextMonitorTime = Now + TimeSerial(0, 0, MONITOR_INTERVAL_SEC)
    Application.OnTime NextMonitorTime, "'" & ThisWorkbook.Name & "'!RSS監視タイマー"
End Sub


' ================================================================
' 内部: ステータスセル更新
' ================================================================
Private Sub UpdateStatusCell(ByVal isConnected As Boolean, ByVal status As String)
    Dim stampText As String
    stampText = status & " " & Format(Now, "hh:mm:ss")

    Dim mws As Worksheet, dws As Worksheet
    On Error Resume Next
    Set mws = ThisWorkbook.Sheets("銘柄管理")
    Set dws = ThisWorkbook.Sheets("ダッシュボード")
    On Error GoTo 0

    If Not mws Is Nothing Then WriteStatus mws, isConnected, stampText
    If Not dws Is Nothing Then WriteStatus dws, isConnected, stampText
End Sub

Private Sub WriteStatus(ByVal ws As Worksheet, _
                        ByVal isConnected As Boolean, _
                        ByVal txt As String)
    On Error Resume Next
    With ws.Cells(STATUS_CELL_ROW, STATUS_CELL_COL)
        .Value = txt
        .Font.Bold = True
        If isConnected Then
            .Interior.Color = RGB(200, 255, 200)
            .Font.Color = RGB(0, 80, 0)
        Else
            .Interior.Color = RGB(255, 200, 200)
            .Font.Color = RGB(150, 0, 0)
        End If
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

Private Sub UpdateStatusSilent()
    Dim dummyStatus As String, dummyDetails As String
    Dim isConn As Boolean
    isConn = CheckRSSStatus(dummyStatus, dummyDetails)
    UpdateStatusCell isConn, dummyStatus
End Sub
