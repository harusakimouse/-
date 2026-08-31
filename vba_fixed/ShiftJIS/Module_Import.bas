Attribute VB_Name = "Module_Import"
Option Explicit

'==================================================================
' Module_Import : 歩み値CSV を銘柄シートに取り込む
'
'   RSS でリアルタイム記録できないとき（検証・バックテスト）に使います。
'   CSV の列並び（1行目が見出しでも可）
'       1列目 : 時刻      15:00:03 / 15:00:03.250 / 2026-08-28 15:00:03
'       2列目 : 出来高    ティック単体 でも 当日累計 でも可（取込時に選択）
'       3列目 : 約定値
'       4列目 : 最良買気配値（任意）
'       5列目 : 最良売気配値（任意）
'       6列目 : ティック記号 ↑↓（任意）
'==================================================================

Public Sub ImportTickCsv()

    Dim ws As Worksheet
    Dim f As Variant
    Dim ff As Integer
    Dim buf As String
    Dim parts() As String
    Dim rw As Long, nRead As Long
    Dim isCumulative As Boolean
    Dim prevVol As Double, v As Double, dv As Double
    Dim tTxt As String
    Dim t As Date
    Dim ans As VbMsgBoxResult
    Dim errNum As Long, errDesc As String
    Dim opened As Boolean

    If BlockedWhileLogging("CSV（歩み値）取込") Then Exit Sub

    Set ws = ActiveSheet
    If ws.Name = RESULT_SHEET Then
        MsgBox "取り込み先の銘柄シートを開いてから実行してください。", vbExclamation, "CSV取込"
        Exit Sub
    End If

    f = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "歩み値CSVを選択")
    If VarType(f) = vbBoolean Then Exit Sub

    ans = MsgBox("CSVの2列目「出来高」は当日累計ですか？" & vbCrLf & vbCrLf & _
                 "  はい   = 累計出来高（差分をティック出来高にします）" & vbCrLf & _
                 "  いいえ = そのティックの出来高", vbYesNoCancel + vbQuestion, "CSV取込")
    If ans = vbCancel Then Exit Sub
    isCumulative = (ans = vbYes)

    rw = LastTickRow(ws)
    If rw >= TICK_FIRST_ROW Then
        If MsgBox("既存のティックログを消してから取り込みますか？", vbYesNo + vbQuestion, "CSV取込") = vbYes Then
            ws.Range(ws.Cells(TICK_FIRST_ROW, COL_TIME), ws.Cells(rw, COL_MARK)).ClearContents
        End If
    End If
    rw = LastTickRow(ws) + 1
    If rw < TICK_FIRST_ROW Then rw = TICK_FIRST_ROW

    Application.ScreenUpdating = False
    ff = FreeFile
    On Error GoTo Fin
    Open CStr(f) For Input As #ff
    opened = True

    Do Until EOF(ff)
        Line Input #ff, buf
        buf = Trim$(buf)
        If Len(buf) = 0 Then GoTo NextLine

        parts = Split(buf, ",")
        If UBound(parts) < 2 Then GoTo NextLine

        tTxt = CleanCell(parts(0))
        t = ParseTimeText(tTxt)
        If t = 0 Then GoTo NextLine                       ' 見出し行など

        If Not IsNumeric(CleanCell(parts(2))) Then GoTo NextLine

        v = Val(CleanCell(parts(1)))
        If isCumulative Then
            If prevVol = 0 Then
                prevVol = v
                GoTo NextLine                             ' 1件目は基準値
            End If
            dv = v - prevVol
            prevVol = v
            If dv <= 0 Then GoTo NextLine
        Else
            dv = v
        End If

        ws.Cells(rw, COL_TIME).Value = t
        ws.Cells(rw, COL_PRICE).Value = Val(CleanCell(parts(2)))
        ws.Cells(rw, COL_VOL).Value = dv
        If UBound(parts) >= 3 Then ws.Cells(rw, COL_BID).Value = Val(CleanCell(parts(3)))
        If UBound(parts) >= 4 Then ws.Cells(rw, COL_ASK).Value = Val(CleanCell(parts(4)))
        If UBound(parts) >= 5 Then ws.Cells(rw, COL_MARK).Value = CleanCell(parts(5))

        rw = rw + 1
        nRead = nRead + 1
NextLine:
    Loop

Fin:
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If opened Then Close #ff
    Err.Clear
    On Error GoTo 0

    ws.Columns(COL_TIME).NumberFormatLocal = "hh:mm:ss"
    Application.ScreenUpdating = True

    If errNum <> 0 Then
        MsgBox "取込中にエラーが発生しました。" & vbCrLf & errNum & " : " & errDesc, _
               vbCritical, "CSV取込"
        Exit Sub
    End If

    MsgBox nRead & " 件のティックを取り込みました。" & vbCrLf & _
           "「④ 判定 実行」で判定できます。", vbInformation, "CSV取込"
End Sub

'------------------------------------------------------------------
' セル文字列の掃除（ダブルクォート・タブ・前後空白を除去）
'------------------------------------------------------------------
Private Function CleanCell(ByVal s As String) As String

    s = Replace$(s, """", "")
    s = Replace$(s, vbTab, "")
    CleanCell = Trim$(s)
End Function

'------------------------------------------------------------------
' 時刻文字列 → Date（時分秒のみ）。解釈できなければ 0
'------------------------------------------------------------------
Private Function ParseTimeText(ByVal s As String) As Date

    Dim p As Long
    Dim d As Date

    If Len(s) = 0 Then Exit Function

    ' "15:00:03.250" のようなミリ秒は切り捨て
    p = InStr(s, ".")
    If p > 0 Then
        If InStr(s, ":") > 0 And p > InStr(s, ":") Then s = Left$(s, p - 1)
    End If

    If InStr(s, ":") = 0 Then Exit Function

    On Error GoTo Bad
    d = CDate(s)
    ParseTimeText = TimeSerial(Hour(d), Minute(d), Second(d))
    Exit Function
Bad:
    ParseTimeText = 0
End Function
