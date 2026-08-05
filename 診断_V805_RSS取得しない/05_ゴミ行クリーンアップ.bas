Attribute VB_Name = "Mod_ゴミ行クリーンアップ"
Option Explicit

' ==================================================================
' OHLCV 5シートの「ゴミ行」除去
'
' 【対象】
'   A列に値があるのに、それが銘柄コードの形（4桁数字 / 3桁数字+英字1）
'   でなく、C列の RSS 数式も無い行。
'
'   V805 で実測された 13 行（終値シート基準）:
'     25("45")  27("47")  31("53")  75("139") 102("191") 134("253")
'     149("280") 177("334") 205("386") 221("415") 228("421")
'     230("422") 245("436")
'
'   ThisWorkbook のコメントにある「汚染13銘柄を除去」の残骸で、
'   銘柄管理側は空なのに OHLCV の A列だけ値が残っている状態。
'
' 【放置するとどうなるか】
'   ・Mod_日次更新 が code <> "" で有効行と判定 → 毎日「欠測」にカウント
'   ・Mod_買抽出v13.抽出実行 も同じ判定でスキャン対象に入れる
'
' 【この処理】
'   まず 点検 で一覧を出し、内容を確認してから 実行 する。
'   A〜D 列だけを消し、E列以降の履歴は消さない（誤操作時に戻せるよう）。
'   履歴も消したい場合は CLEAR_HISTORY を True にする。
' ==================================================================

Private Const PWD           As String = "ne19480314"
Private Const R_FIRST       As Long = 6
Private Const R_LAST        As Long = 505
Private Const C_HISTEND     As Long = 104
Private Const CLEAR_HISTORY As Boolean = False   ' True にすると E:CZ も消す


Private Function IsStockCode(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) <> 4 Then Exit Function
    Dim i As Long
    For i = 1 To 3
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    Dim c As String: c = UCase$(Mid$(s, 4, 1))
    IsStockCode = ((c >= "0" And c <= "9") Or (c >= "A" And c <= "Z"))
End Function


' ==================================================================
' ① まずこれを実行して内容を確認する（無変更）
' ==================================================================
Public Sub ゴミ行_点検()
    Dim msg As String
    msg = "【ゴミ行 点検】（変更しません）" & vbCrLf & String(40, "-") & vbCrLf

    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer, total As Long

    For s = 0 To UBound(shNames)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet

        Dim n As Long, list As String
        Dim r As Long
        For r = R_FIRST To R_LAST
            If IsJunkRow(ws, r) Then
                n = n + 1
                If Len(list) < 400 Then
                    list = list & "    行" & r & " A=[" & _
                           Trim$(CStr(ws.Cells(r, 1).Value)) & "] E=[" & _
                           CStr(ws.Cells(r, 5).Value) & "]" & vbCrLf
                End If
            End If
        Next r

        msg = msg & ws.Name & ": " & n & " 行" & vbCrLf & list
        total = total + n
NextSheet:
    Next s

    msg = msg & String(40, "-") & vbCrLf & "合計 " & total & " 行" & vbCrLf & vbCrLf & _
          "問題なければ ゴミ行_実行 を走らせてください。" & vbCrLf & _
          "（履歴列 E:CZ を消すかどうかは定数 CLEAR_HISTORY = " & CLEAR_HISTORY & "）"

    MsgBox msg, vbInformation, "ゴミ行 点検"
End Sub


' ==================================================================
' ② 実行
' ==================================================================
Public Sub ゴミ行_実行()
    If MsgBox("ゴミ行の A〜D 列を消去します。" & vbCrLf & _
              IIf(CLEAR_HISTORY, "★履歴列 E:CZ も消去します★" & vbCrLf, "履歴列 E:CZ は残します。" & vbCrLf) & vbCrLf & _
              "先に ゴミ行_点検 で内容を確認しましたか?" & vbCrLf & _
              "バックアップは取りましたか?" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbExclamation, "ゴミ行 除去") <> vbYes Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer, report As String, total As Long

    For s = 0 To UBound(shNames)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo ErrHandler
        If ws Is Nothing Then GoTo NextSheet

        On Error Resume Next
        ws.Unprotect Password:=PWD
        On Error GoTo ErrHandler

        Dim n As Long
        Dim r As Long
        For r = R_FIRST To R_LAST
            If IsJunkRow(ws, r) Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).ClearContents
                If CLEAR_HISTORY Then
                    ws.Range(ws.Cells(r, 5), ws.Cells(r, C_HISTEND)).ClearContents
                End If
                n = n + 1
            End If
        Next r

        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo ErrHandler

        report = report & ws.Name & ": " & n & " 行" & vbCrLf
        total = total + n
NextSheet:
    Next s

    Application.ScreenUpdating = True
    MsgBox "ゴミ行を除去しました。" & vbCrLf & String(30, "-") & vbCrLf & _
           report & String(30, "-") & vbCrLf & "合計 " & total & " 行", _
           vbInformation, "完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Number & " " & Err.Description, vbCritical, "エラー"
End Sub


' ------------------------------------------------------------------
' A列に値があり、銘柄コードでなく、C列に RSS 数式も無い行 = ゴミ
' ------------------------------------------------------------------
Private Function IsJunkRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim code As String
    code = Trim$(CStr(ws.Cells(r, 1).Value))
    If code = "" Then Exit Function
    If code = "0" Then IsJunkRow = True: Exit Function
    If IsStockCode(code) Then Exit Function
    If ws.Cells(r, 3).HasFormula Then Exit Function   ' 念のため数式があれば触らない
    IsJunkRow = True
End Function
