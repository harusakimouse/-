Attribute VB_Name = "日付ヘッダー整合"
Option Explicit
'=============================================================================
' 打ち出の無限こづち  日付ヘッダー整合ツール  v2.1
'
' 【v1からの設計変更】v1は重大な欠陥があったため全面的に作り直しました。
'   × v1: 入力された基準日から営業日カレンダーを「生成」して上書き
'          - 日本の祝日を考慮できず実データと最大11日ズレる（実測62/74列が誤り）
'          - DATA_DAYS=100 で CZ列まで書き込み、CA:CZ に残骸を作る
'          - 土日を基準日に入力できてしまう（7/26(日)混入の原因）
'   ○ v2: 日付を生成しない。5シートを列ごとに突合し「多数派」に合わせる
'          - 実際の営業日（祝日休場を含む）をそのまま保持
'          - 書き込み範囲を E列～IT列（250営業日分）に制限
'          - 土日日付を検出したら自動修正せず警告して中止
'          - シート保護の解除/再保護に対応
'   v2.1: 変数名 fix → bFix（Fix はVBA組み込み関数のためコンパイルエラー）
'
' 【安全性】row 5 以降（実データ）と row 2 は一切変更しません。row 3 のみ。
'=============================================================================
Private Const DATE_ROW       As Long = 3      ' 日付が入る行
Private Const DATA_COL_START As Long = 5      ' E列
Private Const DATA_COL_END   As Long = 254    ' IT列（250営業日分）
Private Const SHEET_PW       As String = "ne19480314"

Private Function ShNames() As Variant
    ShNames = Array("始値", "高値", "安値", "終値", "出来高")
End Function

'=============================================================================
' 診断のみ（シート無変更）? 全列を突合します
'=============================================================================
Public Sub 日付整合性チェック()
    Dim shs As Variant: shs = ShNames()
    Dim msg As String
    msg = "=== 日付整合性チェック (" & Format(Now, "yyyy/mm/dd hh:mm") & ") ===" & vbCrLf & vbCrLf

    Dim i As Long, ws As Worksheet
    msg = msg & "【E列（最新確定日）】" & vbCrLf
    For i = LBound(shs) To UBound(shs)
        Set ws = GetWs(CStr(shs(i)))
        If ws Is Nothing Then
            msg = msg & "  " & shs(i) & ": [シートなし]" & vbCrLf
        Else
            Dim dv As Variant: dv = ws.Cells(DATE_ROW, DATA_COL_START).Value
            If IsDate(dv) Then
                msg = msg & "  " & shs(i) & ": " & Format(CDate(dv), "yyyy/mm/dd (aaa)") & vbCrLf
            Else
                msg = msg & "  " & shs(i) & ": [日付でない: " & CStr(dv) & "]" & vbCrLf
            End If
        End If
    Next i

    Dim badCols As String, badCount As Long
    Dim weekendMsg As String, weekendCount As Long
    Dim c As Long
    For c = DATA_COL_START To DATA_COL_END
        Dim majN As Long
        Dim maj As Variant: maj = 列の多数派日付(c, majN)
        If Not IsEmpty(maj) Then
            ' 土日チェック
            If Weekday(CDate(maj), vbMonday) > 5 Then
                weekendCount = weekendCount + 1
                If Len(weekendMsg) < 400 Then
                    weekendMsg = weekendMsg & "  " & ColLetter(c) & "3 = " & _
                                 Format(CDate(maj), "m/d (aaa)") & vbCrLf
                End If
            End If
            ' 少数派シートを列挙
            For i = LBound(shs) To UBound(shs)
                Set ws = GetWs(CStr(shs(i)))
                If Not ws Is Nothing Then
                    Dim cv As Variant: cv = ws.Cells(DATE_ROW, c).Value
                    If IsDate(cv) Then
                        If CDate(cv) <> CDate(maj) Then
                            badCount = badCount + 1
                            If Len(badCols) < 700 Then
                                badCols = badCols & "  " & shs(i) & "!" & ColLetter(c) & "3 = " & _
                                          Format(CDate(cv), "m/d") & "  (多数派: " & _
                                          Format(CDate(maj), "m/d") & ")" & vbCrLf
                            End If
                        End If
                    ElseIf Not IsEmpty(cv) Then
                        badCount = badCount + 1
                    End If
                End If
            Next i
        End If
    Next c

    msg = msg & vbCrLf & "【全列突合（E列～IT列）】" & vbCrLf
    If badCount = 0 Then
        msg = msg & "  [OK] 5シートの日付ヘッダーは全列一致しています。" & vbCrLf
    Else
        msg = msg & "  [警告] " & badCount & " セルが多数派と不一致:" & vbCrLf & badCols
        msg = msg & "  → 日付ヘッダー一括整合 で多数派に揃えられます。" & vbCrLf
    End If

    msg = msg & vbCrLf & "【土日チェック】" & vbCrLf
    If weekendCount = 0 Then
        msg = msg & "  [OK] 土日の日付は含まれていません。" & vbCrLf
    Else
        msg = msg & "  [重大] 土日の日付が " & weekendCount & " 列あります:" & vbCrLf & weekendMsg
        msg = msg & "  → 自動修正できません。手動で正しい営業日に直してください。" & vbCrLf
    End If

    MsgBox msg, vbInformation, "日付整合性チェック v2.1"
End Sub

'=============================================================================
' 一括整合: 列ごとの多数派に合わせる（日付は生成しない）
'=============================================================================
Public Sub 日付ヘッダー一括整合()
    Dim shs As Variant: shs = ShNames()

    ' --- 事前に土日混入を確認。あれば中止 ---
    Dim c As Long, weekendMsg As String
    For c = DATA_COL_START To DATA_COL_END
        Dim majN0 As Long
        Dim m0 As Variant: m0 = 列の多数派日付(c, majN0)
        If Not IsEmpty(m0) Then
            If Weekday(CDate(m0), vbMonday) > 5 Then
                weekendMsg = weekendMsg & "  " & ColLetter(c) & "3 = " & _
                             Format(CDate(m0), "m/d (aaa)") & vbCrLf
            End If
        End If
    Next c
    If weekendMsg <> "" Then
        MsgBox "土日の日付ヘッダーが見つかりました。" & vbCrLf & vbCrLf & weekendMsg & vbCrLf & _
               "これは多数派修正では直せません（実データの実営業日が不明なため）。" & vbCrLf & _
               "該当セルを手動で正しい営業日に修正してから再実行してください。", _
               vbCritical, "中止: 土日の日付"
        Exit Sub
    End If

    ' --- 修正対象を集計 ---
    Dim plan As String, planCount As Long
    Dim i As Long, ws As Worksheet
    For c = DATA_COL_START To DATA_COL_END
        Dim majN As Long
        Dim maj As Variant: maj = 列の多数派日付(c, majN)
        If Not IsEmpty(maj) Then
            For i = LBound(shs) To UBound(shs)
                Set ws = GetWs(CStr(shs(i)))
                If Not ws Is Nothing Then
                    Dim cv As Variant: cv = ws.Cells(DATE_ROW, c).Value
                    Dim needFix As Boolean: needFix = False
                    If IsDate(cv) Then
                        If CDate(cv) <> CDate(maj) Then needFix = True
                    Else
                        needFix = True
                    End If
                    If needFix Then
                        planCount = planCount + 1
                        If Len(plan) < 700 Then
                            plan = plan & "  " & shs(i) & "!" & ColLetter(c) & "3 → " & _
                                   Format(CDate(maj), "m/d") & vbCrLf
                        End If
                    End If
                End If
            Next i
        End If
    Next c

    If planCount = 0 Then
        MsgBox "日付ヘッダーは既に5シートで一致しています。" & vbCrLf & _
               "修正の必要はありません。", vbInformation, "整合済み"
        Exit Sub
    End If

    If MsgBox("以下の " & planCount & " セルを多数派の日付に揃えます。" & vbCrLf & vbCrLf & _
              plan & vbCrLf & _
              "※ 対象は row 3 の E列～IT列のみです。" & vbCrLf & _
              "※ row 2 と実データ（row 5 以降）は変更しません。" & vbCrLf & _
              "※ 日付の生成は行いません（実営業日を保持します）。" & vbCrLf & vbCrLf & _
              "実行しますか?", vbQuestion + vbYesNo, "実行確認") = vbNo Then
        Exit Sub
    End If

    ' --- 保護解除 ---
    Dim wasProt(0 To 4) As Boolean
    For i = 0 To 4
        Set ws = GetWs(CStr(shs(i)))
        If Not ws Is Nothing Then
            wasProt(i) = ws.ProtectContents
            If wasProt(i) Then
                On Error Resume Next
                ws.Unprotect Password:=SHEET_PW
                On Error GoTo 0
            End If
        End If
    Next i

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nDone As Long
    For c = DATA_COL_START To DATA_COL_END
        Dim majN2 As Long
        Dim maj2 As Variant: maj2 = 列の多数派日付(c, majN2)
        If Not IsEmpty(maj2) Then
            For i = LBound(shs) To UBound(shs)
                Set ws = GetWs(CStr(shs(i)))
                If Not ws Is Nothing Then
                    Dim cv2 As Variant: cv2 = ws.Cells(DATE_ROW, c).Value
                    Dim bFix As Boolean: bFix = False
                    If IsDate(cv2) Then
                        If CDate(cv2) <> CDate(maj2) Then bFix = True
                    Else
                        bFix = True
                    End If
                    If bFix Then
                        With ws.Cells(DATE_ROW, c)
                            .Value = CDate(maj2)
                            .NumberFormat = "m/d"
                        End With
                        nDone = nDone + 1
                    End If
                End If
            Next i
        End If
    Next c

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' --- 再保護 ---
    For i = 0 To 4
        Set ws = GetWs(CStr(shs(i)))
        If Not ws Is Nothing Then
            If wasProt(i) Then
                On Error Resume Next
                ws.Protect Password:=SHEET_PW, DrawingObjects:=True, _
                           Contents:=True, Scenarios:=True
                On Error GoTo 0
            End If
        End If
    Next i

    MsgBox "日付ヘッダー一括整合 完了" & vbCrLf & vbCrLf & _
           "修正: " & nDone & " セル" & vbCrLf & _
           "範囲: row 3 の E列～IT列" & vbCrLf & vbCrLf & _
           "※ 日付は生成せず、5シートの多数派に揃えました。" & vbCrLf & _
           "※ 実データ（row 5 以降）は変更していません。", _
           vbInformation, "整合完了 v2.1"
End Sub

'=============================================================================
' ヘルパー
'=============================================================================
' 指定列について5シートの日付の最頻値を返す（同数なら出現順で先頭）
Private Function 列の多数派日付(ByVal c As Long, ByRef bestN As Long) As Variant
    Dim shs As Variant: shs = ShNames()
    Dim vals(0 To 4) As Variant
    Dim i As Long, j As Long, ws As Worksheet
    For i = 0 To 4
        vals(i) = Empty
        Set ws = GetWs(CStr(shs(i)))
        If Not ws Is Nothing Then
            Dim cv As Variant: cv = ws.Cells(DATE_ROW, c).Value
            If IsDate(cv) Then vals(i) = CDate(cv)
        End If
    Next i

    Dim best As Variant: best = Empty
    bestN = 0
    For i = 0 To 4
        If Not IsEmpty(vals(i)) Then
            Dim n As Long: n = 0
            For j = 0 To 4
                If Not IsEmpty(vals(j)) Then
                    If CDate(vals(j)) = CDate(vals(i)) Then n = n + 1
                End If
            Next j
            If n > bestN Then
                bestN = n
                best = vals(i)
            End If
        End If
    Next i
    列の多数派日付 = best
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim s As String, x As Long: x = c
    Do While x > 0
        Dim r As Long: r = (x - 1) Mod 26
        s = Chr$(65 + r) & s
        x = (x - 1 - r) \ 26
    Loop
    ColLetter = s
End Function

Private Function GetWs(ByVal n As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(n)
    On Error GoTo 0
    Set GetWs = ws
End Function