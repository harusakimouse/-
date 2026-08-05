Attribute VB_Name = "Mod_アット記号除去"
Option Explicit

' ==================================================================
' RSS式の "@" を剥がして書き直す
'
' 【症状】
'   =@RssMarket("8058","始値")  が入っているのに 数字が出ず空っぽ。
'
' 【原因】
'   正常な行は  =RssMarket("8058","始値")  （保存形式は _xll.RssMarket）で、
'   "@" は付いていない。"@" 付きの式は VBA が後から書いたもので、
'   出どころは次の3つ:
'
'     Mod銘柄管理.銘柄追加
'       ws2.Cells(emptyRow,3).Formula = "=@RssMarket(""" & code & """,""" & items(s) & """)"
'     Mod銘柄管理.C列_銘柄名RSS式_一括適用
'       ws.Cells(i,"C").Formula = "=@RssMarket($B" & i & ",""銘柄名称"")"
'     Mod_RSS式修復_v2.OHLCV_RSS式修復_v2
'       ws.Cells(r,3).Formula = "=@RssMarket($A" & r & ",""" & shName & """)"
'
'   VBA から "@" 付きで書いた式は、書いた瞬間に RSS アドインが
'   読み込まれていないと XLL 関数として解決されず、
'   エラーも出さないまま値が来ない（空のまま）状態になる。
'
' 【この処理】
'   ブック全体を走査し、"@Rss..." を含む数式から "@" だけを外して
'   書き直す。書いた後に本当に直ったかを1セルずつ検証する。
'
' 【重要】
'   アドインが生きている状態で実行してください。
'   死んだ状態で書き直すと同じことの繰り返しになるので、
'   このマクロは最初にアドインの生死を確認し、死んでいたら中止します。
' ==================================================================

Private Const PWD As String = "ne19480314"


' ==================================================================
' ① 点検（無変更）: どこに "@" 付きの式があるか数える
' ==================================================================
Public Sub アット記号_点検()
    Dim msg As String
    msg = "【@付きRSS式 点検】（変更しません）" & vbCrLf & String(42, "=") & vbCrLf

    Dim total As Long, blankCnt As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim n As Long, nb As Long, sample As String
        n = 0: nb = 0: sample = ""

        Dim rng As Range, cell As Range
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWs

        For Each cell In rng
            Dim f As String
            f = cell.Formula
            If InStr(1, f, "@Rss", vbTextCompare) > 0 Or InStr(1, f, "SINGLE", vbTextCompare) > 0 Then
                n = n + 1
                Dim v As Variant: v = cell.Value
                If IsEmpty(v) Or (Not IsError(v) And CStr(v) = "") Then nb = nb + 1
                If Len(sample) < 90 Then sample = sample & cell.Address(False, False) & " "
            End If
        Next cell

        If n > 0 Then
            msg = msg & ws.Name & ": " & n & " セル（うち空 " & nb & "）  例: " & sample & vbCrLf
            total = total + n
            blankCnt = blankCnt + nb
        End If
        Set rng = Nothing
NextWs:
    Next ws

    msg = msg & String(42, "-") & vbCrLf & _
          "合計 " & total & " セル / うち値が空 " & blankCnt & " セル" & vbCrLf & vbCrLf

    If total = 0 Then
        msg = msg & "→ @付きの式はありません。原因は別にあります。" & vbCrLf & _
              "  03_Mod_RSS環境診断 の RSS環境診断 を実行してください。"
    Else
        msg = msg & "→ アット記号_除去 で直せます。" & vbCrLf & _
              "  ※ MarketSpeed II が起動・ログイン済みの状態で実行してください。"
    End If

    CopyToClip msg
    MsgBox msg, vbInformation, "@付きRSS式 点検"
End Sub


' ==================================================================
' ② 除去実行
' ==================================================================
Public Sub アット記号_除去()

    If Not AddinAlive() Then
        MsgBox "MarketSpeed II RSS アドインが読み込まれていません（RssMarket が #NAME?）。" & vbCrLf & vbCrLf & _
               "この状態で式を書き直しても、また同じ「空っぽ」になります。" & vbCrLf & vbCrLf & _
               "MarketSpeed II を起動・ログインしてから Excel を開き直して、" & vbCrLf & _
               "もう一度実行してください。", vbCritical, "中止: アドイン未ロード"
        Exit Sub
    End If

    If MsgBox("RSS式から ""@"" を外して書き直します。" & vbCrLf & vbCrLf & _
              "  =@RssMarket(""8058"",""始値"")" & vbCrLf & _
              "        ↓" & vbCrLf & _
              "  =RssMarket(""8058"",""始値"")" & vbCrLf & vbCrLf & _
              "★実行前にブックのコピーを取ってください★" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbExclamation, "@除去") <> vbYes Then Exit Sub

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic   ' RTD を回すため自動必須

    Dim fixed As Long, failed As Long, report As String, failList As String
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Dim wasProtected As Boolean
        wasProtected = ws.ProtectContents
        If wasProtected Then
            On Error Resume Next
            ws.Unprotect Password:=PWD
            On Error GoTo Cleanup
        End If

        Dim rng As Range
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Cleanup

        Dim nSheet As Long: nSheet = 0

        If Not rng Is Nothing Then
            Dim cell As Range
            For Each cell In rng
                Dim f As String
                f = cell.Formula
                If InStr(1, f, "@Rss", vbTextCompare) > 0 Then
                    Dim newF As String
                    newF = Replace(f, "@Rss", "Rss", 1, -1, vbTextCompare)

                    On Error Resume Next
                    Err.Clear
                    cell.Formula = newF
                    Dim wroteOK As Boolean
                    wroteOK = (Err.Number = 0)
                    Err.Clear
                    On Error GoTo Cleanup

                    ' 本当に直ったか検証
                    If wroteOK Then
                        If InStr(1, cell.Formula, "@Rss", vbTextCompare) = 0 _
                           And cell.HasFormula Then
                            fixed = fixed + 1
                            nSheet = nSheet + 1
                        Else
                            failed = failed + 1
                            If Len(failList) < 120 Then _
                                failList = failList & ws.Name & "!" & cell.Address(False, False) & " "
                        End If
                    Else
                        failed = failed + 1
                        If Len(failList) < 120 Then _
                            failList = failList & ws.Name & "!" & cell.Address(False, False) & " "
                    End If
                End If
            Next cell
        End If

        If wasProtected Then
            On Error Resume Next
            ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                       DrawingObjects:=True, Contents:=True, Scenarios:=True
            On Error GoTo Cleanup
        End If

        If nSheet > 0 Then report = report & "  " & ws.Name & ": " & nSheet & " セル" & vbCrLf
        Set rng = Nothing
    Next ws

Cleanup:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description

    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.StatusBar = False
    On Error GoTo 0

    Application.CalculateFullRebuild

    If eN <> 0 Then
        MsgBox "処理中にエラー: " & eN & " " & eD & vbCrLf & vbCrLf & _
               "計算モード・保護は元に戻しました。", vbCritical, "エラー"
        Exit Sub
    End If

    MsgBox "@ の除去が完了しました。" & vbCrLf & String(34, "-") & vbCrLf & _
           report & String(34, "-") & vbCrLf & _
           "書き直し成功: " & fixed & " セル" & vbCrLf & _
           "失敗: " & failed & " セル" & IIf(failed > 0, vbCrLf & "  " & failList, "") & vbCrLf & vbCrLf & _
           "RSS の値が入るまで 30〜60 秒待ってください。" & vbCrLf & _
           "（このブックは RTD トピックが約1,700件あります）" & vbCrLf & vbCrLf & _
           "※ 今後 銘柄追加 マクロを使うと また @ が付きます。" & vbCrLf & _
           "  Mod銘柄管理 の3か所から ""=@Rss"" を ""=Rss"" に直してください。", _
           IIf(failed = 0, vbInformation, vbExclamation), "@除去 完了"
End Sub


' ==================================================================
' ③ 追加確認: 1セルだけ手動テストする（結果を切り分ける）
' ==================================================================
Public Sub RSS_1セルテスト()
    Dim code As String
    code = Trim$(InputBox("テストする銘柄コード（4桁）", "RSS 1セルテスト", "8058"))
    If code = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "_RSSテスト"

    Application.Calculation = xlCalculationAutomatic

    ws.Range("A1").Value = "@なし"
    ws.Range("A2").Value = "@あり"
    On Error Resume Next
    ws.Range("B1").Formula = "=RssMarket(""" & code & """,""始値"")"
    Dim e1 As Long: e1 = Err.Number: Err.Clear
    ws.Range("B2").Formula = "=@RssMarket(""" & code & """,""始値"")"
    Dim e2 As Long: e2 = Err.Number: Err.Clear
    On Error GoTo 0

    ' RTD が来るのを待つ
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) < 15
        DoEvents
        Application.Calculate
        If IsNumeric(ws.Range("B1").Value) Then
            If ws.Range("B1").Value > 0 Then Exit Do
        End If
    Loop

    Dim msg As String
    msg = "【RSS 1セルテスト】 コード " & code & vbCrLf & String(38, "-") & vbCrLf & _
          "=RssMarket(...)   → 式:[" & ws.Range("B1").Formula & "]" & vbCrLf & _
          "                    値:[" & CStr(ws.Range("B1").Text) & "]  書込err=" & e1 & vbCrLf & vbCrLf & _
          "=@RssMarket(...)  → 式:[" & ws.Range("B2").Formula & "]" & vbCrLf & _
          "                    値:[" & CStr(ws.Range("B2").Text) & "]  書込err=" & e2 & vbCrLf & _
          String(38, "-") & vbCrLf

    If IsNumeric(ws.Range("B1").Value) And ws.Range("B1").Value > 0 Then
        If Not (IsNumeric(ws.Range("B2").Value) And ws.Range("B2").Value > 0) Then
            msg = msg & "→ @なしは取れて @ありは取れない。" & vbCrLf & _
                  "  原因は @ です。アット記号_除去 を実行してください。"
        Else
            msg = msg & "→ どちらも取れています。@ は原因ではありません。" & vbCrLf & _
                  "  03_Mod_RSS環境診断 で他の要因を確認してください。"
        End If
    ElseIf InStr(ws.Range("B1").Text, "NAME") > 0 Then
        msg = msg & "→ #NAME? です。RSS アドインが読み込まれていません。" & vbCrLf & _
              "  MarketSpeed II を先に起動・ログインしてから Excel を開き直してください。"
    Else
        msg = msg & "→ @なしでも取れません。MarketSpeed II 側の問題です。" & vbCrLf & _
              "  ・ログインしているか" & vbCrLf & _
              "  ・その銘柄が取扱対象か" & vbCrLf & _
              "  ・RSS の登録上限に達していないか" & vbCrLf & _
              "  を確認してください。"
    End If

    CopyToClip msg
    MsgBox msg, vbInformation, "RSS 1セルテスト"

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub


' ==================================================================
' 補助
' ==================================================================
Private Function AddinAlive() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = Application.Evaluate("RssMarket(""7203"",""銘柄名称"")")
    On Error GoTo 0
    If IsError(v) Then
        If CLng(v) = xlErrName Then Exit Function
    End If
    AddinAlive = True
End Function

Private Sub CopyToClip(ByVal s As String)
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    o.SetText s
    o.PutInClipboard
    On Error GoTo 0
End Sub
