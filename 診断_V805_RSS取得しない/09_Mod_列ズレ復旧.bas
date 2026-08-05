Attribute VB_Name = "Mod_列ズレ復旧"
Option Explicit

' ==================================================================
' OHLCV 列ズレ復旧（本日分が E列 から右へずれてしまった状態を直す）
'
' 【症状】
'   本日分が E列ではなく V列（や他の列）にある。
'   E列〜U列 が空。抽出も日次更新も RSS 表示も全部おかしい。
'
' 【原因】
'   Mod_日次更新.ExecuteDailyUpdate は
'     ① F:CZ ← E:CY へ右シフト
'     ② E列を ClearContents
'     ③ RSSが取れた行だけ E列へ書込
'     ④ 日付ヘッダをシフト
'     ⑤ E3 に当日日付（＝「取込済」の目印）
'   という順。④⑤ に着く前にエラーが出ると ErrHandler → Cleanup へ飛び、
'   E3 が空のまま終わる。E3 が空だと冒頭の重複ガード
'       If IsDate(e3) Then If CDate(e3) = bizDate Then 取込済
'   が効かないので、次に開くとまた①②が走って もう1列ずれる。
'   ThisWorkbook.Workbook_Open が開いた瞬間に 日次更新_本日分 を呼ぶため、
'   開くたびに 1列ずつ増える。
'
' 【この復旧でやること】
'   空になった左側の列ぶんだけ、データと日付ヘッダを左へ詰め直す。
'   E:U が空なだけでデータは消えていないので、詰めれば元に戻る。
'
' 【使う順番】
'   1) 列ズレ_点検        ← まず必ずこれ（無変更。状態を報告するだけ）
'   2) 列ズレ_左詰め復旧  ← 点検の内容に納得してから
'   3) 01_ThisWorkbook_修正版 を必ず入れる（入れないと再発します）
' ==================================================================

Private Const PWD       As String = "ne19480314"
Private Const R_TOPIX   As Long = 5
Private Const R_FIRST   As Long = 6
Private Const R_LAST    As Long = 505
Private Const R_DATE    As Long = 3
Private Const C_CODE    As Long = 1
Private Const C_TODAY   As Long = 5      ' E列（本来ここが本日分）
Private Const C_HISTEND As Long = 104    ' CZ列

Private Function ShNames() As Variant
    ShNames = Array("始値", "高値", "安値", "終値", "出来高")
End Function


' ==================================================================
' 1) 点検（無変更）
' ==================================================================
Public Sub 列ズレ_点検()
    Dim msg As String
    msg = "【OHLCV 列ズレ 点検】（変更しません）" & vbCrLf & String(44, "=") & vbCrLf

    Dim sh As Variant: sh = ShNames()
    Dim s As Integer
    Dim dataOff(0 To 4) As Long, hdrOff(0 To 4) As Long
    Dim allSame As Boolean: allSame = True

    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sh(s))
        On Error GoTo 0
        If ws Is Nothing Then
            msg = msg & sh(s) & ": シートなし" & vbCrLf
            GoTo NextS
        End If

        dataOff(s) = FirstDataCol(ws) - C_TODAY
        hdrOff(s) = FirstDateCol(ws) - C_TODAY

        msg = msg & ws.Name & ":" & vbCrLf
        msg = msg & "   データ先頭列 = " & ColLetter(FirstDataCol(ws)) & _
              " 列 (ズレ " & dataOff(s) & " 列)" & vbCrLf
        msg = msg & "   日付先頭列   = " & ColLetter(FirstDateCol(ws)) & _
              " 列 (ズレ " & hdrOff(s) & " 列)"
        If FirstDateCol(ws) >= C_TODAY Then
            Dim dv As Variant: dv = ws.Cells(R_DATE, FirstDateCol(ws)).Value
            If IsDate(dv) Then msg = msg & "  最新日付=" & Format(CDate(dv), "yyyy/m/d(aaa)")
        End If
        msg = msg & vbCrLf

        If dataOff(s) <> dataOff(0) Or hdrOff(s) <> hdrOff(0) Then allSame = False
NextS:
    Next s

    msg = msg & String(44, "-") & vbCrLf

    Dim cw As Worksheet
    On Error Resume Next
    Set cw = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If Not cw Is Nothing Then
        msg = msg & "終値!E3 (取込済の目印) = ["
        Dim e3 As Variant: e3 = cw.Cells(R_DATE, C_TODAY).Value
        If IsDate(e3) Then
            msg = msg & Format(CDate(e3), "yyyy/m/d") & "]" & vbCrLf
        Else
            msg = msg & CStr(e3) & "]  ★空です → 開くたびに1列ずれます" & vbCrLf
        End If
    End If

    msg = msg & String(44, "=") & vbCrLf

    If dataOff(0) = 0 And hdrOff(0) = 0 Then
        msg = msg & "→ ズレは検出されませんでした。"
    ElseIf Not allSame Then
        msg = msg & "★ シートごとにズレ量が違います。" & vbCrLf & _
              "  左詰め復旧では直りません。" & vbCrLf & _
              "  08_Mod_一括再取込 の OHLCV_一括再取込_全銘柄（モードA）で" & vbCrLf & _
              "  作り直してください。"
    ElseIf dataOff(0) <> hdrOff(0) Then
        msg = msg & "★ データと日付ヘッダのズレ量が違います" & vbCrLf & _
              "  (データ " & dataOff(0) & " 列 / ヘッダ " & hdrOff(0) & " 列)" & vbCrLf & vbCrLf & _
              "  左詰め復旧は「データもヘッダもそれぞれのズレ量で」詰めます。" & vbCrLf & _
              "  不安なら 08 の一括再取込（モードA）の方が確実です。"
    Else
        msg = msg & "→ データ・ヘッダとも " & dataOff(0) & " 列 右にズレています。" & vbCrLf & _
              "  列ズレ_左詰め復旧 で元に戻せます。" & vbCrLf & _
              "  （E列〜" & ColLetter(C_TODAY + dataOff(0) - 1) & "列 は空なので失われるデータはありません）"
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "※ 復旧後は必ず 01_ThisWorkbook_修正版 を入れてください。" & vbCrLf & _
          "  入れないと、開くたびにまた1列ずつずれます。"

    CopyToClip msg
    MsgBox msg, vbInformation, "列ズレ 点検"
End Sub


' ==================================================================
' 2) 左詰め復旧
' ==================================================================
Public Sub 列ズレ_左詰め復旧()

    ' --- ズレ量の確定 ---
    Dim cw As Worksheet
    On Error Resume Next
    Set cw = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If cw Is Nothing Then MsgBox "終値シートがありません。", vbCritical: Exit Sub

    Dim dOff As Long, hOff As Long
    dOff = FirstDataCol(cw) - C_TODAY
    hOff = FirstDateCol(cw) - C_TODAY

    If dOff <= 0 And hOff <= 0 Then
        MsgBox "ズレは検出されませんでした。何もしません。", vbInformation, "復旧不要"
        Exit Sub
    End If

    If MsgBox("OHLCV 5シートを左へ詰め直します。" & vbCrLf & String(38, "-") & vbCrLf & _
              "データ:     " & dOff & " 列ぶん左へ" & vbCrLf & _
              "日付ヘッダ: " & hOff & " 列ぶん左へ" & vbCrLf & String(38, "-") & vbCrLf & vbCrLf & _
              "★実行前に必ずブックのコピーを取ってください★" & vbCrLf & vbCrLf & _
              "先に 列ズレ_点検 の内容を確認しましたか?" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbExclamation, "左詰め復旧") <> vbYes Then Exit Sub

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim sh As Variant: sh = ShNames()
    Dim s As Integer, report As String

    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sh(s))
        On Error GoTo Cleanup
        If ws Is Nothing Then GoTo NextS

        On Error Resume Next
        ws.Unprotect Password:=PWD
        On Error GoTo Cleanup

        ' ---- データ本体を dOff 列ぶん左へ ----
        If dOff > 0 Then
            ws.Range(ws.Cells(R_TOPIX, C_TODAY), ws.Cells(R_LAST, C_HISTEND - dOff)).Value = _
                ws.Range(ws.Cells(R_TOPIX, C_TODAY + dOff), ws.Cells(R_LAST, C_HISTEND)).Value
            ' 右端の空いた dOff 列をクリア
            ws.Range(ws.Cells(R_TOPIX, C_HISTEND - dOff + 1), ws.Cells(R_LAST, C_HISTEND)).ClearContents
        End If

        ' ---- 日付ヘッダを hOff 列ぶん左へ ----
        If hOff > 0 Then
            ws.Range(ws.Cells(R_DATE, C_TODAY), ws.Cells(R_DATE, C_HISTEND - hOff)).Value = _
                ws.Range(ws.Cells(R_DATE, C_TODAY + hOff), ws.Cells(R_DATE, C_HISTEND)).Value
            ws.Range(ws.Cells(R_DATE, C_HISTEND - hOff + 1), ws.Cells(R_DATE, C_HISTEND)).ClearContents
            ws.Range(ws.Cells(R_DATE, C_TODAY), ws.Cells(R_DATE, C_HISTEND)).NumberFormat = "m/d"
        End If

        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo Cleanup

        report = report & ws.Name & ": OK" & vbCrLf
NextS:
    Next s

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

    If eN <> 0 Then
        MsgBox "復旧中にエラー: " & eN & " " & eD & vbCrLf & vbCrLf & _
               "計算モード・保護は元に戻しました。" & vbCrLf & _
               "バックアップから戻して 08 の一括再取込を使ってください。", _
               vbCritical, "エラー"
        Exit Sub
    End If

    MsgBox "左詰め復旧が完了しました。" & vbCrLf & String(34, "-") & vbCrLf & _
           report & String(34, "-") & vbCrLf & _
           "この後の手順:" & vbCrLf & _
           " ① 四本値_整合性チェック を実行して違反 0 を確認" & vbCrLf & _
           " ② 01_ThisWorkbook_修正版 を必ず入れる" & vbCrLf & _
           "    （入れないと開くたびにまたズレます）" & vbCrLf & _
           " ③ 心配なら 08 の一括再取込（モードA）で取り直す", _
           vbInformation, "復旧完了"
End Sub


' ==================================================================
' 3) 四本値の整合性チェック（復旧が正しかったかの検証）
'    高値 >= 安値 / 高値 >= 始値・終値 / 安値 <= 始値・終値
' ==================================================================
Public Sub 四本値_整合性チェック()
    Dim wo As Worksheet, wh As Worksheet, wl As Worksheet, wc As Worksheet
    On Error Resume Next
    Set wo = ThisWorkbook.Sheets("始値")
    Set wh = ThisWorkbook.Sheets("高値")
    Set wl = ThisWorkbook.Sheets("安値")
    Set wc = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If wo Is Nothing Or wh Is Nothing Or wl Is Nothing Or wc Is Nothing Then
        MsgBox "OHLC シートが揃っていません。", vbCritical: Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim chk As Long, bad As Long, worst As String
    Dim r As Long, c As Long
    For r = R_FIRST To R_LAST
        Dim code As String: code = Trim$(CStr(wc.Cells(r, C_CODE).Value))
        If code = "" Or code = "0" Then GoTo NextR

        For c = C_TODAY To C_HISTEND
            Dim o As Variant, h As Variant, l As Variant, k As Variant
            o = wo.Cells(r, c).Value: h = wh.Cells(r, c).Value
            l = wl.Cells(r, c).Value: k = wc.Cells(r, c).Value
            If Not (IsNumeric(o) And IsNumeric(h) And IsNumeric(l) And IsNumeric(k)) Then GoTo NextC
            If o = "" Or h = "" Or l = "" Or k = "" Then GoTo NextC
            If CDbl(o) <= 0 Or CDbl(h) <= 0 Or CDbl(l) <= 0 Or CDbl(k) <= 0 Then GoTo NextC

            chk = chk + 1
            If CDbl(h) < CDbl(l) Or CDbl(h) < CDbl(k) Or CDbl(h) < CDbl(o) _
               Or CDbl(l) > CDbl(k) Or CDbl(l) > CDbl(o) Then
                bad = bad + 1
                If Len(worst) < 200 Then worst = worst & code & "/" & ColLetter(c) & "列 "
            End If
NextC:
        Next c
NextR:
    Next r

    Application.ScreenUpdating = True

    MsgBox "【四本値 整合性チェック】" & vbCrLf & String(30, "=") & vbCrLf & _
           "検証セル: " & Format(chk, "#,##0") & vbCrLf & _
           "違反セル: " & bad & "  (" & Format(IIf(chk = 0, 0, bad / chk), "0.00%") & ")" & vbCrLf & vbCrLf & _
           IIf(bad = 0, "問題なし。列ズレは発生していません。", _
               "★列ズレが残っています:" & vbCrLf & worst & vbCrLf & vbCrLf & _
               "→ 08 の一括再取込（モードA）で作り直してください。"), _
           IIf(bad = 0, vbInformation, vbExclamation), "整合性チェック"
End Sub


' ==================================================================
' 補助
' ==================================================================

' データが実際に入っている一番左の列（E列以降で探す）
Private Function FirstDataCol(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long, hit As Long
    For c = C_TODAY To C_HISTEND
        hit = 0
        For r = R_TOPIX To 60          ' 先頭 55 行を標本にする
            Dim v As Variant: v = ws.Cells(r, c).Value
            If IsNumeric(v) And v <> "" Then
                If CDbl(v) > 0 Then hit = hit + 1
            End If
        Next r
        If hit >= 5 Then FirstDataCol = c: Exit Function
    Next c
    FirstDataCol = C_TODAY
End Function

' 日付ヘッダが入っている一番左の列
Private Function FirstDateCol(ByVal ws As Worksheet) As Long
    Dim c As Long
    For c = C_TODAY To C_HISTEND
        If IsDate(ws.Cells(R_DATE, c).Value) Then FirstDateCol = c: Exit Function
    Next c
    FirstDateCol = C_TODAY
End Function

Private Function ColLetter(ByVal c As Long) As String
    If c < 1 Then ColLetter = "?": Exit Function
    ColLetter = Split(Cells(1, c).Address(True, False), "$")(0)
End Function

Private Sub CopyToClip(ByVal s As String)
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    o.SetText s
    o.PutInClipboard
    On Error GoTo 0
End Sub
