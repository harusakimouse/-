Attribute VB_Name = "Mod_Launcher"
Option Explicit

'==== 設定 ========================================================
Private Const SHEET_NAME As String = "起動リスト"
Private Const MARK       As String = "〇"
Private Const FIRST_ROW  As Long = 2
' A:起動  B:表示名  C:ファイル名  D:フォルダ  E:状態
'==================================================================
Private g_LastErr As String
'==================================================================

'------------------------------------------------------------------
' ① リスト更新  デスクトップの xls* を一覧化（〇印は保持）
'------------------------------------------------------------------
Public Sub リスト更新()
    Dim ws As Worksheet, fso As Object, fld As Object, f As Object
    Dim marks As Object, deskPath As String
    Dim r As Long, lastRow As Long, ext As String

    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    Set marks = CreateObject("Scripting.Dictionary")
    marks.CompareMode = 1
    lastRow = LastDataRow(ws)
    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "C").Value)) <> "" Then
            marks(Trim$(CStr(ws.Cells(r, "C").Value))) = _
                (Trim$(CStr(ws.Cells(r, "A").Value)) = MARK)
        End If
    Next r

    deskPath = DesktopPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(deskPath) Then
        MsgBox "デスクトップが見つかりません:" & vbCrLf & deskPath, vbExclamation
        Exit Sub
    End If
    Set fld = fso.GetFolder(deskPath)

    Application.ScreenUpdating = False
    If lastRow >= FIRST_ROW Then
        ws.Range(ws.Cells(FIRST_ROW, "A"), ws.Cells(lastRow, "E")).ClearContents
    End If

    r = FIRST_ROW
    For Each f In fld.Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If Left$(ext, 3) = "xls" Then
            If Left$(f.Name, 2) <> "~$" _
               And StrComp(f.Name, ThisWorkbook.Name, vbTextCompare) <> 0 Then
                ws.Cells(r, "A").Value = IIf(marks.Exists(f.Name), _
                                             IIf(marks(f.Name), MARK, ""), "")
                ws.Cells(r, "B").Value = fso.GetBaseName(f.Name)
                ws.Cells(r, "C").Value = f.Name
                ws.Cells(r, "D").Value = deskPath
                ws.Cells(r, "E").Value = ""
                r = r + 1
            End If
        End If
    Next f
    Application.ScreenUpdating = True

    MsgBox (r - FIRST_ROW) & " 件を一覧にしました。", vbInformation
End Sub

'------------------------------------------------------------------
' ② 起動  A列に〇の付いたブックだけ開く
'   ※開いた相手ブックの Workbook_Open でエラーが出ても
'     「開けたかどうか」だけで判定する（誤って 失敗:1004 と出さない）
'------------------------------------------------------------------
Public Sub 起動()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim fullPath As String
    Dim opened As Long, skipped As Long, failCnt As Long
    Dim wb As Workbook, errNo As Long

    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = LastDataRow(ws)

    Application.ScreenUpdating = False
    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = MARK Then
            fullPath = BuildPath(ws, r)

            If fullPath = "" Then
                ws.Cells(r, "E").Value = "行が未入力"

            ElseIf Not FindOpenWorkbook(fullPath) Is Nothing Then
                ws.Cells(r, "E").Value = "既に開いています"
                skipped = skipped + 1

            ElseIf Dir$(fullPath) = "" Then
                ws.Cells(r, "E").Value = "ファイルなし"

            Else
                Set wb = Nothing
                Err.Clear
                On Error Resume Next
                Set wb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0)
                errNo = Err.Number
                Err.Clear
                On Error GoTo 0

                ' 相手ブックの起動マクロでエラーが出ると
                ' 変数に受け取れないことがあるので、開いた一覧から探し直す
                If wb Is Nothing Then Set wb = FindOpenWorkbook(fullPath)

                If Not wb Is Nothing Then
                    opened = opened + 1
                    If errNo = 0 Then
                        ws.Cells(r, "E").Value = "起動OK"
                    Else
                        ws.Cells(r, "E").Value = "起動OK(相手マクロ警告 " & errNo & ")"
                    End If
                Else
                    failCnt = failCnt + 1
                    ws.Cells(r, "E").Value = "失敗:" & errNo
                End If
                Set wb = Nothing
            End If
        End If
    Next r
    ThisWorkbook.Activate
    Application.ScreenUpdating = True

    MsgBox "起動 " & opened & " 件 / 既に開いていた " & skipped & " 件 / 失敗 " & failCnt & " 件", _
           vbInformation
End Sub

'------------------------------------------------------------------
' 保存して閉じる
'   起動リストで〇の付いたブック（＝起動したブック）だけを保存して閉じ、
'   最後にランチャー自身を保存して Excel を終了する
'------------------------------------------------------------------
Public Sub 保存して閉じる_選択ブック()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim wb As Workbook, nm As String, fullPath As String
    Dim cnt As Long, failCnt As Long, restCnt As Long
    Dim closed As String, failed As String, rest As String
    Dim ans As VbMsgBoxResult
    Dim i As Long

    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = LastDataRow(ws)

    ans = MsgBox("起動リストで〇の付いたブックを保存して閉じ、" & vbCrLf & _
                 "そのあと Excel を終了します。よろしいですか？", _
                 vbQuestion + vbOKCancel, "保存して終了")
    If ans <> vbOK Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- ① 〇の付いたブックを保存して閉じる ----
    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = MARK Then
            fullPath = BuildPath(ws, r)
            Set wb = FindOpenWorkbook(fullPath)

            If wb Is Nothing Then
                ws.Cells(r, "E").Value = "開いていません"
            ElseIf StrComp(wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
                ' 自分自身は最後に処理する
            Else
                nm = wb.Name
                If SaveAndClose(wb) Then
                    cnt = cnt + 1
                    closed = closed & "・" & nm & vbCrLf
                    ws.Cells(r, "E").Value = "保存して閉じました"
                Else
                    failCnt = failCnt + 1
                    failed = failed & "・" & nm & "（" & g_LastErr & "）" & vbCrLf
                    ws.Cells(r, "E").Value = "閉じられません"
                End If
            End If
            Set wb = Nothing
        End If
    Next r

    Application.DisplayAlerts = True      ' ← 必ず戻す
    Application.ScreenUpdating = True

    If failCnt > 0 Then
        MsgBox cnt & "ブックを保存して閉じました。" & vbCrLf & vbCrLf & _
               "処理できなかったブック:" & vbCrLf & failed & vbCrLf & _
               "Excel は終了しません。手動で確認してください。", vbExclamation
        Exit Sub
    End If

    ' ---- ② 〇以外でまだ開いているブックの確認 ----
    For i = 1 To Workbooks.Count
        If StrComp(Workbooks(i).Name, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            rest = rest & "・" & Workbooks(i).Name & vbCrLf
            restCnt = restCnt + 1
        End If
    Next i

    If restCnt = 0 Then
        MsgBox cnt & "ブックを保存して閉じました。" & vbCrLf & vbCrLf & closed & vbCrLf & _
               "OK で Excel を終了します。", vbInformation
    Else
        ans = MsgBox(cnt & "ブックを保存して閉じました。" & vbCrLf & vbCrLf & closed & vbCrLf & _
                     "次のブックがまだ開いています:" & vbCrLf & rest & vbCrLf & _
                     "これも保存して Excel を終了しますか？" & vbCrLf & _
                     "（いいえ＝Excel を終了しません）", vbQuestion + vbYesNo, "確認")
        If ans <> vbYes Then Exit Sub

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        For i = Workbooks.Count To 1 Step -1
            Set wb = Workbooks(i)
            If StrComp(wb.Name, ThisWorkbook.Name, vbTextCompare) <> 0 Then
                nm = wb.Name
                If Not SaveAndClose(wb) Then
                    failCnt = failCnt + 1
                    failed = failed & "・" & nm & "（" & g_LastErr & "）" & vbCrLf
                End If
            End If
            Set wb = Nothing
        Next i
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

        If failCnt > 0 Then
            MsgBox "次のブックは処理できませんでした:" & vbCrLf & failed & vbCrLf & _
                   "Excel は終了しません。手動で確認してください。", vbExclamation
            Exit Sub
        End If
    End If

    ' ---- ③ ランチャー自身を保存して Excel 終了 ----
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True             ' ← これが無いと異常終了扱い＝回復ファイル
    Application.Quit
End Sub

' 1ブックを保存して閉じる（成功でTrue／失敗理由は g_LastErr）
Private Function SaveAndClose(wb As Workbook) As Boolean
    g_LastErr = ""
    Err.Clear
    On Error Resume Next

    If wb.Path = "" Then
        wb.Saved = True                   ' 一度も保存していない新規ブックは破棄
    ElseIf Not wb.ReadOnly Then
        wb.Save
    End If
    If Err.Number <> 0 Then
        g_LastErr = "保存失敗 " & Err.Number
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    wb.Close SaveChanges:=False
    If Err.Number <> 0 Then
        g_LastErr = "閉じる失敗 " & Err.Number
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0
    SaveAndClose = True
End Function

'------------------------------------------------------------------
' 全部ON / 全部OFF（押すたび切替）
'------------------------------------------------------------------
Public Sub 全部ONOFF()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim allOn As Boolean, cnt As Long, onCnt As Long

    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = LastDataRow(ws)

    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "C").Value)) <> "" Then
            cnt = cnt + 1
            If Trim$(CStr(ws.Cells(r, "A").Value)) = MARK Then onCnt = onCnt + 1
        End If
    Next r
    allOn = (cnt > 0 And onCnt = cnt)

    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "C").Value)) <> "" Then
            ws.Cells(r, "A").Value = IIf(allOn, "", MARK)
        End If
    Next r
End Sub

'==== 旧ボタン名との互換用（登録名がどれでも動くように） ==========
Public Sub 起動_5ブック一括()
    起動
End Sub

Public Sub 終了_5ブック保存して閉じる()
    保存して閉じる_選択ブック
End Sub

Public Sub 全部ON_OFF()
    全部ONOFF
End Sub

'==== 内部処理 ====================================================
Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        MsgBox "シート「" & SHEET_NAME & "」が見つかりません。", vbCritical
    End If
End Function

Private Function LastDataRow(ws As Worksheet) As Long
    Dim a As Long, c As Long
    a = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    c = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    LastDataRow = IIf(a > c, a, c)
    If LastDataRow < FIRST_ROW Then LastDataRow = FIRST_ROW
End Function

Private Function BuildPath(ws As Worksheet, r As Long) As String
    Dim fld As String, fn As String
    fld = Trim$(CStr(ws.Cells(r, "D").Value))
    fn = Trim$(CStr(ws.Cells(r, "C").Value))
    If fn = "" Then Exit Function
    If fld = "" Then fld = DesktopPath()
    If Right$(fld, 1) <> "\" Then fld = fld & "\"
    BuildPath = fld & fn
End Function

'------------------------------------------------------------------
' 開いているブックを探す
'   ① フルパス一致  ② ファイル名一致（OneDriveでパス表記が変わる保険）
'------------------------------------------------------------------
Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook, fn As String
    If fullPath = "" Then Exit Function
    fn = Mid$(fullPath, InStrRev(fullPath, "\") + 1)

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fn, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function DesktopPath() As String
    Dim p As String
    On Error Resume Next
    p = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If p = "" Then p = Environ$("USERPROFILE") & "\Desktop"
    If Right$(p, 1) <> "\" Then p = p & "\"
    DesktopPath = p
End Function

'------------------------------------------------------------------
' 起動失敗の理由を詳しく調べる（〇の行が対象）
'------------------------------------------------------------------
Public Sub 診断_起動失敗()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim fullPath As String, msg As String
    Dim lockFile As String, wb As Workbook, errNo As Long

    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = LastDataRow(ws)

    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = MARK Then
            fullPath = BuildPath(ws, r)
            msg = msg & "■ " & ws.Cells(r, "C").Value & vbCrLf

            If Dir$(fullPath) = "" Then
                msg = msg & "   ファイルが見つかりません" & vbCrLf & vbCrLf
                GoTo NextRow
            End If

            msg = msg & "   サイズ : " & Format$(FileLen(fullPath), "#,##0") & " バイト" & vbCrLf

            ' 排他ロックファイル（~$名前.xlsm）の有無
            lockFile = Left$(fullPath, InStrRev(fullPath, "\")) & _
                       "~$" & Mid$(fullPath, InStrRev(fullPath, "\") + 1)
            If Dir$(lockFile) <> "" Then
                msg = msg & "   *ロックファイルあり（別プロセスが掴んでいます）" & vbCrLf
            End If

            If Not FindOpenWorkbook(fullPath) Is Nothing Then
                msg = msg & "   既に開いています" & vbCrLf & vbCrLf
                GoTo NextRow
            End If

            ' 実際に開いてみる
            Set wb = Nothing
            Err.Clear
            On Error Resume Next
            Set wb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, Notify:=False)
            errNo = Err.Number
            Err.Clear
            On Error GoTo 0
            If wb Is Nothing Then Set wb = FindOpenWorkbook(fullPath)

            If Not wb Is Nothing Then
                If errNo = 0 Then
                    msg = msg & "   -> 正常に開けました" & vbCrLf
                Else
                    msg = msg & "   -> 開けています。ただし相手ブックの起動マクロで" & vbCrLf & _
                                "      エラー " & errNo & " が出ています（起動そのものは成功）" & vbCrLf
                End If
                If wb.ReadOnly Then msg = msg & "   （読み取り専用で開いています）" & vbCrLf
            Else
                msg = msg & "   オープン失敗 : " & errNo & vbCrLf
            End If
            Set wb = Nothing
            msg = msg & vbCrLf
        End If
NextRow:
    Next r

    MsgBox msg, vbInformation, "起動失敗の診断"
End Sub
