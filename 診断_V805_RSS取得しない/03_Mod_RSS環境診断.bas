Attribute VB_Name = "Mod_RSS環境診断"
Option Explicit

' ==================================================================
' RSS環境診断  ―― 「RSS取得しない」の原因を一発で切り分ける
'
' 使い方: Alt+F11 → 挿入 → 標準モジュール → 全文貼り付け → RSS環境診断 を実行
'
' 見るもの:
'   A. RssMarket / RssIndexMarket が名前解決できるか（＝アドインが生きているか）
'   B. 計算モード（手動だと RTD は絶対に更新されない）
'   C. RTD スロットル間隔（-1 だと自動更新されない）
'   D. 読み込まれているアドインの一覧
'   E. OHLCV 5 シートの C 列の実状（#NAME? / 数式欠損 / 値ゼロ の件数）
'   F. 日付ヘッダ E3 の状態（当日日付が入っていると日次更新が空振り済み）
'   G. A 列のゴミ行
' ==================================================================

Private Const PWD As String = "ne19480314"


Public Sub RSS環境診断()
    Dim msg As String
    msg = "【RSS環境診断】 " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
          String(46, "=") & vbCrLf

    ' ---------- A. アドインの名前解決 ----------
    msg = msg & "[A] RSS関数の名前解決" & vbCrLf
    Dim rv As Variant, riv As Variant
    On Error Resume Next
    rv = Application.Evaluate("RssMarket(""7203"",""銘柄名称"")")
    riv = Application.Evaluate("RssIndexMarket(""TOPX"",""現在値"")")
    On Error GoTo 0

    Dim addinOK As Boolean: addinOK = True
    If IsError(rv) Then
        msg = msg & "   RssMarket      : " & ErrText(rv) & vbCrLf
        If CLng(rv) = xlErrName Then addinOK = False
    Else
        msg = msg & "   RssMarket      : OK -> [" & CStr(rv) & "]" & vbCrLf
    End If
    If IsError(riv) Then
        msg = msg & "   RssIndexMarket : " & ErrText(riv) & vbCrLf
    Else
        msg = msg & "   RssIndexMarket : OK -> [" & CStr(riv) & "]" & vbCrLf
    End If

    If Not addinOK Then
        msg = msg & "   ★ #NAME? = MarketSpeed II RSS アドインが読み込まれていません。" & vbCrLf & _
                    "      MarketSpeed II を先に起動・ログインしてから Excel を開き直してください。" & vbCrLf
    End If
    msg = msg & vbCrLf

    ' ---------- B. 計算モード ----------
    msg = msg & "[B] 計算モード" & vbCrLf & "   "
    Select Case Application.Calculation
        Case xlCalculationAutomatic:            msg = msg & "自動  [OK]"
        Case xlCalculationManual:               msg = msg & "★手動★  ← RTD が更新されません！"
        Case xlCalculationSemiautomatic:        msg = msg & "データテーブル以外自動"
        Case Else:                              msg = msg & "不明(" & Application.Calculation & ")"
    End Select
    msg = msg & vbCrLf & "   反復計算: " & Application.Iteration & vbCrLf & vbCrLf

    ' ---------- C. RTD スロットル ----------
    msg = msg & "[C] RTD スロットル間隔" & vbCrLf
    Dim thr As Long: thr = -999
    On Error Resume Next
    thr = Application.RTD.ThrottleInterval
    On Error GoTo 0
    If thr = -999 Then
        msg = msg & "   取得できませんでした" & vbCrLf
    ElseIf thr < 0 Then
        msg = msg & "   ★ " & thr & " (手動更新のみ) ← RSS が自動で流れません！" & vbCrLf & _
                    "      修正: Application.RTD.ThrottleInterval = 2000" & vbCrLf
    Else
        msg = msg & "   " & thr & " ms  [OK]" & vbCrLf
    End If
    msg = msg & vbCrLf

    ' ---------- D. アドイン一覧 ----------
    msg = msg & "[D] 読み込み済みアドイン" & vbCrLf
    Dim a As Object, n As Long
    On Error Resume Next
    For Each a In Application.AddIns
        If a.Installed Then
            n = n + 1
            If n <= 15 Then msg = msg & "   (xla/xll) " & a.Name & vbCrLf
        End If
    Next a
    Dim ca As Object
    For Each ca In Application.COMAddIns
        If ca.Connected Then
            n = n + 1
            If n <= 25 Then msg = msg & "   (COM) " & ca.Description & vbCrLf
        End If
    Next ca
    On Error GoTo 0
    If n = 0 Then msg = msg & "   （有効なアドインなし）" & vbCrLf
    msg = msg & vbCrLf

    ' ---------- E. OHLCV 5 シートの C 列 ----------
    msg = msg & "[E] OHLCV C列の実状" & vbCrLf
    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")
    Dim s As Integer
    For s = 0 To UBound(shNames)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSh

        Dim nRows As Long, nNoFormula As Long, nErr As Long, nZero As Long, nOK As Long
        nRows = 0: nNoFormula = 0: nErr = 0: nZero = 0: nOK = 0

        Dim r As Long
        For r = 5 To 505
            Dim code As String
            code = Trim$(CStr(ws.Cells(r, 1).Value))
            If code = "" Or code = "0" Then GoTo NextR
            nRows = nRows + 1

            If Not ws.Cells(r, 3).HasFormula Then
                nNoFormula = nNoFormula + 1
                GoTo NextR
            End If

            Dim v As Variant: v = ws.Cells(r, 3).Value
            If IsError(v) Then
                nErr = nErr + 1
            ElseIf Not IsNumeric(v) Then
                nZero = nZero + 1
            ElseIf CDbl(v) <= 0 Then
                nZero = nZero + 1
            Else
                nOK = nOK + 1
            End If
NextR:
        Next r

        msg = msg & "   " & ws.Name & String(4 - Len(ws.Name) \ 2, " ") & _
              " 対象=" & nRows & " 正常=" & nOK & " エラー=" & nErr & _
              " 0/空=" & nZero & " 数式なし=" & nNoFormula & vbCrLf
NextSh:
    Next s
    msg = msg & vbCrLf

    ' ---------- F. 日付ヘッダ ----------
    msg = msg & "[F] 日付ヘッダ E3" & vbCrLf
    Dim cw As Worksheet
    On Error Resume Next
    Set cw = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If Not cw Is Nothing Then
        Dim e3 As Variant: e3 = cw.Cells(3, 5).Value
        If IsDate(e3) Then
            msg = msg & "   終値!E3 = " & Format(CDate(e3), "yyyy/m/d (aaa)")
            If CDate(e3) = Date Then
                msg = msg & "  ★当日日付が既に入っています" & vbCrLf & _
                      "      → 日次更新は「取込済」と判断して今日はもう走りません。" & vbCrLf & _
                      "      RSS 未接続のまま Workbook_Open で空振りした可能性大。" & vbCrLf
            Else
                msg = msg & "  [OK]" & vbCrLf
            End If
        Else
            msg = msg & "   終値!E3 が日付ではありません: [" & CStr(e3) & "]" & vbCrLf
        End If
    End If
    msg = msg & vbCrLf

    ' ---------- G. A列ゴミ行 ----------
    msg = msg & "[G] A列のゴミ行（コードでない値）" & vbCrLf
    Dim junk As String, jn As Long
    If Not cw Is Nothing Then
        For r = 6 To 505
            Dim c2 As String: c2 = Trim$(CStr(cw.Cells(r, 1).Value))
            If c2 <> "" And c2 <> "0" Then
                If Not cw.Cells(r, 3).HasFormula Then
                    jn = jn + 1
                    If Len(junk) < 160 Then junk = junk & r & "(" & c2 & ") "
                End If
            End If
        Next r
    End If
    If jn = 0 Then
        msg = msg & "   なし [OK]" & vbCrLf
    Else
        msg = msg & "   " & jn & " 行: " & junk & vbCrLf & _
              "   → 05_ゴミ行クリーンアップ で除去してください" & vbCrLf
    End If

    msg = msg & String(46, "=") & vbCrLf & "(クリップボードにコピーしました)"

    On Error Resume Next
    Dim DataObj As Object
    Set DataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText msg
    DataObj.PutInClipboard
    On Error GoTo 0

    MsgBox msg, vbInformation, "RSS環境診断"
End Sub


Private Function ErrText(ByVal v As Variant) As String
    If Not IsError(v) Then ErrText = CStr(v): Exit Function
    Select Case CLng(v)
        Case xlErrName:  ErrText = "#NAME?  ★アドイン未ロード"
        Case xlErrValue: ErrText = "#VALUE!"
        Case xlErrNA:    ErrText = "#N/A （接続待ちの可能性）"
        Case xlErrRef:   ErrText = "#REF!"
        Case xlErrNum:   ErrText = "#NUM!"
        Case xlErrDiv0:  ErrText = "#DIV/0!"
        Case xlErrNull:  ErrText = "#NULL!"
        Case Else:       ErrText = "エラー(" & CLng(v) & ")"
    End Select
End Function


' ==================================================================
' 応急処置: 計算モードと RTD を正常値に戻す
' ==================================================================
Public Sub RSS応急復帰()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.RTD.ThrottleInterval = 2000
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.CutCopyMode = False
    On Error GoTo 0

    Application.CalculateFullRebuild

    MsgBox "計算モード=自動 / RTDスロットル=2000ms に戻し、全再計算しました。" & vbCrLf & _
           "数秒待って C 列に値が入るか確認してください。", _
           vbInformation, "応急復帰"
End Sub
