Attribute VB_Name = "Mod_RSS式修復_v3"
Option Explicit

' ==================================================================
' OHLCV RSS式修復 v3
'
' 【v2 のバグ（実行すると悪化する）】
'   ① D列に  =IF(OR(C6=0,G6=0),"",C6-F6)  を書いていた。
'      実際のシートは  =IF(OR(C6=0,E6=0),"",C6-E6)  で参照列が違う。
'      → 実行すると前日比列が全行壊れる。
'   ② C列に  =@RssMarket($A6,"始値")  を書いていた。
'      実シートは  =RssMarket("5706","始値")  のコード直書き。書式が混在する。
'   ③ On Error Resume Next の直下で
'          .Formula = "..."
'          fixedC = fixedC + 1
'      としていたため、アドイン未ロードで代入が 1004 で失敗しても
'      カウンタだけ増えて「修復しました」と嘘の報告をしていた。
'
' 【v3】
'   ・実シートと同じ形の式を書く（コード直書き、D列は E 列参照）
'   ・書き込み後に HasFormula で検証し、成功した数だけカウント
'   ・最初にアドインの生死を確認し、死んでいたら何もせず中止
'   ・TOPIX 行(5) は RssIndexMarket を使う
'   ・A列にコードが無くコードでもない「ゴミ行」は触らない（報告のみ）
' ==================================================================

Private Const PWD          As String = "ne19480314"
Private Const R_TOPIX      As Long = 5
Private Const R_FIRST      As Long = 6
Private Const R_LAST       As Long = 505
Private Const C_CODE       As Long = 1
Private Const C_RSS        As Long = 3
Private Const C_DIFF       As Long = 4
Private Const C_TODAY      As Long = 5      ' E列


Private Function SheetItem(ByVal shName As String) As String
    ' 終値シートだけ RSS 項目名が「現在値」
    If shName = "終値" Then
        SheetItem = "現在値"
    Else
        SheetItem = shName
    End If
End Function

Private Function IsStockCode(ByVal s As String) As Boolean
    ' 4桁数字、または 3桁数字+英字1（285A 等）を銘柄コードとみなす
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
' アドインの生死確認。死んでいれば False。
' ==================================================================
Public Function RSSアドイン生存確認() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = Application.Evaluate("RssMarket(""7203"",""銘柄名称"")")
    On Error GoTo 0
    If IsError(v) Then
        If CLng(v) = xlErrName Then Exit Function   ' #NAME? = 未ロード
    End If
    RSSアドイン生存確認 = True
End Function


' ==================================================================
' 本体
' ==================================================================
Public Sub OHLCV_RSS式修復_v3()

    If Not RSSアドイン生存確認() Then
        MsgBox "MarketSpeed II RSS アドインが読み込まれていません（RssMarket が #NAME?）。" & vbCrLf & vbCrLf & _
               "この状態で修復を実行すると、数式の書き込みが失敗したまま" & vbCrLf & _
               "「修復した」と誤報告されます（v2 の不具合）。" & vbCrLf & vbCrLf & _
               "MarketSpeed II を起動・ログインしてから、Excel を開き直して" & vbCrLf & _
               "もう一度実行してください。", _
               vbCritical, "中止: アドイン未ロード"
        Exit Sub
    End If

    If MsgBox("OHLCV 5シートの C列(RSS式) と D列(前日比) を補修します。" & vbCrLf & vbCrLf & _
              "・既に数式がある セルは触りません" & vbCrLf & _
              "・A列が銘柄コードでない「ゴミ行」は触りません（報告のみ）" & vbCrLf & vbCrLf & _
              "実行前にブックのバックアップを取ることを推奨します。" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbQuestion, "RSS式修復 v3") <> vbYes Then Exit Sub

    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")

    Dim report As String
    Dim totC As Long, totD As Long, totFail As Long, totJunk As Long
    Dim s As Integer

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

        Dim shName As String: shName = CStr(shNames(s))
        Dim item As String:   item = SheetItem(shName)

        Dim fixC As Long, fixD As Long, failC As Long, junk As Long
        Dim junkRows As String
        Dim r As Long

        ' ---- TOPIX 行 ----
        If Not ws.Cells(R_TOPIX, C_RSS).HasFormula Then
            If WriteFormula(ws.Cells(R_TOPIX, C_RSS), _
                            "=RssIndexMarket(""TOPX"",""" & item & """)") Then
                fixC = fixC + 1
            Else
                failC = failC + 1
            End If
        End If

        ' ---- 銘柄行 ----
        For r = R_FIRST To R_LAST
            Dim code As String
            code = Trim$(CStr(ws.Cells(r, C_CODE).Value))
            If code = "" Or code = "0" Then GoTo NextRow

            If Not IsStockCode(code) Then
                junk = junk + 1
                If Len(junkRows) < 120 Then junkRows = junkRows & r & "(" & code & ") "
                GoTo NextRow
            End If

            ' C列
            If Not ws.Cells(r, C_RSS).HasFormula Then
                If WriteFormula(ws.Cells(r, C_RSS), _
                                "=RssMarket(""" & code & """,""" & item & """)") Then
                    fixC = fixC + 1
                Else
                    failC = failC + 1
                End If
            End If

            ' D列（実シートと同じ E列参照）
            If Not ws.Cells(r, C_DIFF).HasFormula Then
                If WriteFormula(ws.Cells(r, C_DIFF), _
                                "=IF(OR(C" & r & "=0,E" & r & "=0),"""",C" & r & "-E" & r & ")") Then
                    fixD = fixD + 1
                End If
            End If
NextRow:
        Next r

        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo ErrHandler

        report = report & ws.Name & ": C式=" & fixC & " D式=" & fixD & _
                 " 失敗=" & failC & " ゴミ行=" & junk
        If junk > 0 Then report = report & " [" & Trim$(junkRows) & "]"
        report = report & vbCrLf

        totC = totC + fixC: totD = totD + fixD
        totFail = totFail + failC: totJunk = totJunk + junk
NextSheet:
    Next s

    Application.Calculation = prevCalc
    If Application.Calculation = xlCalculationManual Then _
        Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate

    MsgBox "RSS式修復 v3 完了" & vbCrLf & String(30, "-") & vbCrLf & _
           report & String(30, "-") & vbCrLf & _
           "合計: C式=" & totC & " / D式=" & totD & vbCrLf & _
           "書込失敗: " & totFail & vbCrLf & _
           "ゴミ行(未処理): " & totJunk & vbCrLf & vbCrLf & _
           IIf(totJunk > 0, "ゴミ行は 05_ゴミ行クリーンアップ で除去してください。" & vbCrLf, "") & _
           "RSS値が入るまで数十秒待ってください。", _
           vbInformation, "修復完了 v3"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "修復中にエラー:" & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & _
           "計算モードは自動に戻しました。", vbCritical, "エラー"
End Sub


' ------------------------------------------------------------------
' 数式を書き、本当に入ったかを検証して返す
' ------------------------------------------------------------------
Private Function WriteFormula(ByVal target As Range, ByVal f As String) As Boolean
    On Error Resume Next
    Err.Clear
    target.Formula = f
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function                 ' ← 失敗。カウントしない
    End If
    On Error GoTo 0
    WriteFormula = target.HasFormula  ' ← 検証してから True
End Function


' ==================================================================
' 確認用: C列式の欠損を数えるだけ（無変更）
' ==================================================================
Public Sub OHLCV_C列式チェック_v3()
    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")

    Dim msg As String
    msg = "【C列式 欠損チェック】" & vbCrLf & String(30, "-") & vbCrLf

    Dim s As Integer
    For s = 0 To UBound(shNames)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet

        Dim miss As Long, junk As Long, missRows As String
        Dim r As Long
        For r = R_TOPIX To R_LAST
            Dim code As String
            code = Trim$(CStr(ws.Cells(r, C_CODE).Value))
            If code = "" Or code = "0" Then GoTo NextRow
            If r > R_TOPIX Then
                If Not IsStockCode(code) Then junk = junk + 1: GoTo NextRow
            End If
            If Not ws.Cells(r, C_RSS).HasFormula Then
                miss = miss + 1
                If Len(missRows) < 80 Then missRows = missRows & r & ","
            End If
NextRow:
        Next r

        msg = msg & ws.Name & ": 欠損=" & miss
        If miss > 0 Then msg = msg & " 行=[" & missRows & "]"
        msg = msg & "  ゴミ行=" & junk & vbCrLf
NextSheet:
    Next s

    MsgBox msg, vbInformation, "C列式チェック v3"
End Sub
