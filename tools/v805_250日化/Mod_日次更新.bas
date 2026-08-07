Attribute VB_Name = "Mod_日次更新"
Option Explicit
' ==================================================================
' ★ 100番ブック / Mod_日次更新  v6 (2026/08/05)
'
' 【変更履歴】
'   v6: ガード0（gStaleGuardBusy）を追加。買い抽出_安全版が終値!A列を
'       退避している間に OnTime が発火してデータが消える事故を防止。
'       予約/予約解除Subは ThisWorkbook 側に移したため本モジュールから削除。
'   v5: 曜日ガード＋日付前進ガードを内蔵（7/26(日)混入事故の再発防止）
'   v4: 行単位スキップを廃止し「シート単位の一括ブロックシフト」に変更
'       → スキップ行による列ズレを根絶
'       「値不変skip」を削除（前日と同値は正常な相場のため）
'       COL_LAST を250営業日分(IT=254)に変更。78のままだと E:BZ しか
'       ずれず、CA以降が固定されて履歴が静かに壊れる
' ==================================================================
Private Const TOPIX_ROW       As Long = 5
Private Const STOCK_START_ROW As Long = 6
Private Const STOCK_END_ROW   As Long = 505
Private Const DATE_HEADER_ROW As Long = 3
Private Const COL_RSS         As Long = 3      ' C列 = 本日(RSS)
Private Const COL_TODAY       As Long = 5      ' E列 = 直近確定日
Private Const COL_LAST        As Long = 254    ' IT列 = 最古列（250営業日分）
Private Const MAX_INVALID_RATE As Double = 0.1 ' データ未着 許容率10%
Private Const SHEET_PW        As String = "ne19480314"

' ---------- 手動実行 ----------
Public Sub 日次更新_本日分()
    Dim ng As String: ng = WindowRejectReason()
    If ng <> "" Then
        UI_Msg "日次更新を実行できません。" & vbCrLf & vbCrLf & ng & vbCrLf & vbCrLf & _
               "現在時刻: " & Format(Now, "yyyy/mm/dd hh:mm:ss"), vbExclamation, "実行不可"
        Exit Sub
    End If
    ExecuteDailyUpdate True
End Sub

' ---------- 自動実行（ThisWorkbook の OnTime から呼ばれる） ----------
Public Sub 日次更新_自動()
    If WindowRejectReason() <> "" Then Exit Sub
    Dim prevFlag As Boolean: prevFlag = gUnattended
    gUnattended = True
    On Error GoTo SafeExit
    ExecuteDailyUpdate False
SafeExit:
    gUnattended = prevFlag
End Sub

' ---------- 業務日 ----------
Private Function GetBusinessDate() As Date
    If Hour(Now) < 8 Then
        GetBusinessDate = Date - 1
    Else
        GetBusinessDate = Date
    End If
End Function

' ---------- 実行可否（理由付き）----------
Private Function WindowRejectReason() As String
    Dim bd As Date: bd = GetBusinessDate()

    ' ★ガード0: 買い抽出_安全版が終値!A列を退避中は絶対に走らせない
    If gStaleGuardBusy Then
        WindowRejectReason = "買い抽出_安全版の実行中です。"
        Exit Function
    End If

    ' ★ガード1: 業務日が土日なら絶対に走らせない（7/26(日)混入事故の再発防止）
    If Weekday(bd, vbMonday) > 5 Then
        WindowRejectReason = "業務日 " & Format(bd, "m/d") & " は土日です。"
        Exit Function
    End If

    ' ★ガード2: E3より新しい日付でなければ走らせない（二重取込・巻き戻し防止）
    Dim e3 As Variant
    On Error Resume Next
    e3 = ThisWorkbook.Sheets("始値").Cells(DATE_HEADER_ROW, COL_TODAY).Value
    On Error GoTo 0
    If IsDate(e3) Then
        If CDate(e3) >= bd Then
            WindowRejectReason = "始値!E3 = " & Format(CDate(e3), "m/d") & _
                                 " のため " & Format(bd, "m/d") & " 分は取込不要です。"
            Exit Function
        End If
    End If

    ' ★ガード3: 時間窓 15:31 ～ 翌08:00
    Dim h As Integer, m As Integer
    h = Hour(Now): m = Minute(Now)
    If h > 15 Then Exit Function
    If h = 15 And m >= 31 Then Exit Function
    If h < 8 Then Exit Function
    WindowRejectReason = "時間窓(15:31～翌08:00)の外です。"
End Function

' ---------- 本体 ----------
Private Sub ExecuteDailyUpdate(ByVal showMsg As Boolean)
    On Error GoTo ErrHandler
    Dim ShNames As Variant
    ShNames = Array("終値", "始値", "高値", "安値", "出来高")
    Dim bizDate As Date:           bizDate = GetBusinessDate()
    Dim totalUpdated As Long:      totalUpdated = 0
    Dim totalSheetSkipped As Long: totalSheetSkipped = 0
    Dim totalRowInvalid As Long:   totalRowInvalid = 0
    Dim reportMsg As String:       reportMsg = ""
    Dim s As Integer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ★ シート保護を解除
    For s = 0 To 4
        On Error Resume Next
        ThisWorkbook.Sheets(ShNames(s)).Unprotect Password:=SHEET_PW
        On Error GoTo ErrHandler
    Next s

    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(ShNames(s))
        On Error GoTo ErrHandler
        If ws Is Nothing Then GoTo NextSheet
        Dim sheetName As String: sheetName = ws.Name

        ' ② 日付ベース冪等性
        Dim e3 As Variant: e3 = ws.Cells(DATE_HEADER_ROW, COL_TODAY).Value
        If IsDate(e3) Then
            If CDate(e3) >= bizDate Then
                reportMsg = reportMsg & sheetName & ": 既に取込済 (E3=" & _
                            Format(CDate(e3), "m/d") & ")" & vbCrLf
                totalSheetSkipped = totalSheetSkipped + 1
                GoTo NextSheet
            End If
        End If

        ' ③ 非取引日判定（祝日はここで捕まる）
        Dim topixC As Variant: topixC = ws.Cells(TOPIX_ROW, COL_RSS).Value
        Dim topixE As Variant: topixE = ws.Cells(TOPIX_ROW, COL_TODAY).Value
        If IsNumeric(topixC) And IsNumeric(topixE) Then
            If CDbl(topixC) > 0 And CDbl(topixE) > 0 Then
                If CDbl(topixC) = CDbl(topixE) Then
                    reportMsg = reportMsg & sheetName & ": 非取引日 (TOPIX不変)" & vbCrLf
                    totalSheetSkipped = totalSheetSkipped + 1
                    GoTo NextSheet
                End If
            End If
        End If

        ' ④ 事前検証: C列を読み込み、未着率が高ければシート丸ごと中止
        Dim rssVals() As Variant
        ReDim rssVals(TOPIX_ROW To STOCK_END_ROW)
        Dim r As Long, codedCnt As Long, invalidCnt As Long
        codedCnt = 0: invalidCnt = 0
        For r = TOPIX_ROW To STOCK_END_ROW
            rssVals(r) = Empty
            If r > TOPIX_ROW Then
                Dim sCode As String: sCode = Trim$(CStr(ws.Cells(r, 1).Text))
                If sCode = "" Or sCode = "0" Then GoTo NextCheck
            End If
            codedCnt = codedCnt + 1
            Dim cv As Variant: cv = ws.Cells(r, COL_RSS).Value
            If IsError(cv) Then
                invalidCnt = invalidCnt + 1
            ElseIf Not IsNumeric(cv) Then
                invalidCnt = invalidCnt + 1
            ElseIf CDbl(cv) <= 0 Then
                invalidCnt = invalidCnt + 1
            Else
                rssVals(r) = CDbl(cv)
            End If
NextCheck:
        Next r

        If codedCnt = 0 Then
            reportMsg = reportMsg & sheetName & ": 銘柄コード無し → 中止" & vbCrLf
            totalSheetSkipped = totalSheetSkipped + 1
            GoTo NextSheet
        End If
        If invalidCnt / codedCnt > MAX_INVALID_RATE Then
            reportMsg = reportMsg & sheetName & ": データ未着 " & invalidCnt & "/" & codedCnt & _
                        " (許容超) → 中止" & vbCrLf
            totalSheetSkipped = totalSheetSkipped + 1
            GoTo NextSheet
        End If

        ' ⑤ 一括ブロックシフト E:IS → F:IT（全行同時なので列ズレしない）
        ws.Range(ws.Cells(DATE_HEADER_ROW, COL_TODAY + 1), _
                 ws.Cells(DATE_HEADER_ROW, COL_LAST)).Value = _
            ws.Range(ws.Cells(DATE_HEADER_ROW, COL_TODAY), _
                     ws.Cells(DATE_HEADER_ROW, COL_LAST - 1)).Value

        ws.Range(ws.Cells(TOPIX_ROW, COL_TODAY + 1), _
                 ws.Cells(STOCK_END_ROW, COL_LAST)).Value = _
            ws.Range(ws.Cells(TOPIX_ROW, COL_TODAY), _
                     ws.Cells(STOCK_END_ROW, COL_LAST - 1)).Value

        ' ⑥ E列に本日値を書込（未着行は空欄のまま＝列ズレを起こさない）
        Dim sheetUpdated As Long: sheetUpdated = 0
        ws.Range(ws.Cells(TOPIX_ROW, COL_TODAY), _
                 ws.Cells(STOCK_END_ROW, COL_TODAY)).ClearContents
        For r = TOPIX_ROW To STOCK_END_ROW
            If Not IsEmpty(rssVals(r)) Then
                ws.Cells(r, COL_TODAY).Value = rssVals(r)
                sheetUpdated = sheetUpdated + 1
            End If
        Next r

        ' ⑦ 日付ヘッダ
        ws.Cells(DATE_HEADER_ROW, COL_TODAY).Value = bizDate
        ws.Cells(DATE_HEADER_ROW, COL_TODAY).NumberFormat = "m/d"

        reportMsg = reportMsg & sheetName & _
                    ": 更新=" & sheetUpdated & " 未着=" & invalidCnt & vbCrLf
        totalUpdated = totalUpdated + sheetUpdated
        totalRowInvalid = totalRowInvalid + invalidCnt
NextSheet:
    Next s
    GoTo Cleanup

ErrHandler:
    UI_Msg "【日次更新エラー】" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
Cleanup:
    ' ★ シート保護を再設定
    For s = 0 To 4
        On Error Resume Next
        ThisWorkbook.Sheets(ShNames(s)).Protect Password:=SHEET_PW, _
            DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    Next s
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    ' ★ 厳選TOP2 を計算し直し、件数を H7 セルに書く。
    '   これで大引け後に再計算ボタンを押さなくても候補が出揃う。
    '   Mod_再計算ボタン が無くてもエラーにしない。
    On Error Resume Next
    Application.Run "厳選TOP2_再計算_静か"
    Err.Clear
    On Error GoTo 0

    If showMsg Then
        UI_Msg "【日次更新 完了】" & vbCrLf & _
               "業務日: " & Format(bizDate, "yyyy/m/d") & vbCrLf & _
               "実行時刻: " & Format(Now, "hh:mm:ss") & vbCrLf & _
               "━━━━━━━━━━━━━━━━━━" & vbCrLf & _
               reportMsg & _
               "━━━━━━━━━━━━━━━━━━" & vbCrLf & _
               "更新: " & totalUpdated & " 行" & vbCrLf & _
               "シートskip: " & totalSheetSkipped & vbCrLf & _
               "データ未着: " & totalRowInvalid, _
               vbInformation, "日次更新"
    End If
End Sub

' ---------- 診断 ----------
Public Sub 日次更新_時間窓確認()
    Dim ng As String: ng = WindowRejectReason()
    UI_Msg "現在時刻: " & Format(Now, "yyyy/mm/dd hh:mm:ss") & vbCrLf & _
           "業務日: " & Format(GetBusinessDate(), "yyyy/m/d") & _
           " (" & Format(GetBusinessDate(), "aaa") & ")" & vbCrLf & _
           "始値!E3: " & ThisWorkbook.Sheets("始値").Range("E3").Text & vbCrLf & _
           "ロック(gStaleGuardBusy): " & IIf(gStaleGuardBusy, "ON", "OFF") & vbCrLf & _
           "判定: " & IIf(ng = "", "★実行可", "実行不可 → " & ng), _
           vbInformation, "時間窓確認"
End Sub