Attribute VB_Name = "Mod_OHLCV移植"
Option Explicit

' ==================================================================
' OHLCV 履歴の移植（V805 → 今のブック）
'
' 【用途】
'   検証済みの V805 の OHLCV 履歴（8/4 〜 4/10、287銘柄）を、
'   いま開いているブック（バックアップ版など）へ丸ごと移す。
'
' 【突き合わせ方】
'   行 = 銘柄コード（A列）で対応をとる  ← 行番号がずれていても正しく入る
'   列 = 日付ヘッダ（3行目）を V805 のものに置き換える
'
' 【触る範囲】
'   3行目 E:CZ（日付ヘッダ）
'   5〜505行 E:CZ（履歴データ）
'   ★ A列(コード) B列(銘柄名) C列(RSS式) D列(差分式) には一切触りません
'
' 【安全設計】
'   ・V805 は「マクロ無効・読み取り専用」で開く
'     （そのまま開くと V805 の Workbook_Open が走って日次更新が暴れるため）
'   ・処理後 V805 は保存せずに閉じる
'   ・エラー・中断のどちらでも 計算モード / 画面更新 / シート保護 を必ず戻す
'
' 【使う前に】
'   いまのブックのコピーを取ってください。
' ==================================================================

Private Const PWD       As String = "ne19480314"
Private Const R_DATE    As Long = 3
Private Const R_TOPIX   As Long = 5
Private Const R_FIRST   As Long = 6
Private Const R_LAST    As Long = 505
Private Const C_CODE    As Long = 1
Private Const C_TODAY   As Long = 5      ' E列
Private Const C_HISTEND As Long = 104    ' CZ列

Private Function ShNames() As Variant
    ShNames = Array("始値", "高値", "安値", "終値", "出来高")
End Function


' ==================================================================
'  メイン
' ==================================================================
Public Sub OHLCV_V805から移植()

    ' ---------- 移植元ファイルの選択 ----------
    Dim srcPath As String
    srcPath = PickFile()
    If srcPath = "" Then Exit Sub

    If InStr(1, srcPath, ThisWorkbook.Name, vbTextCompare) > 0 Then
        MsgBox "自分自身は選べません。", vbExclamation: Exit Sub
    End If

    If MsgBox("OHLCV 履歴を移植します。" & vbCrLf & String(40, "-") & vbCrLf & _
              "移植元: " & Dir(srcPath) & vbCrLf & _
              "移植先: " & ThisWorkbook.Name & vbCrLf & String(40, "-") & vbCrLf & vbCrLf & _
              "・行は 銘柄コード で突き合わせます" & vbCrLf & _
              "・日付ヘッダ(3行目)は移植元のものに置き換えます" & vbCrLf & _
              "・A〜D列（コード/銘柄名/RSS式/差分式）は触りません" & vbCrLf & vbCrLf & _
              "★このブックの E:CZ は全部上書きされます★" & vbCrLf & _
              "★事前にコピーを取りましたか?★" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbExclamation, "OHLCV 移植") <> vbYes Then Exit Sub

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevSec As MsoAutomationSecurity: prevSec = Application.AutomationSecurity
    Dim src As Workbook
    Dim report As String

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False                       ' Workbook_Open を止める
    Application.AutomationSecurity = msoAutomationSecurityForceDisable  ' マクロ無効で開く
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "移植元を開いています..."
    Set src = Workbooks.Open(Filename:=srcPath, ReadOnly:=True, UpdateLinks:=0)

    ' ---------- 移植元の検査 ----------
    Dim sCl As Worksheet
    On Error Resume Next
    Set sCl = src.Sheets("終値")
    On Error GoTo Cleanup
    If sCl Is Nothing Then
        MsgBox "移植元に 終値シートがありません。", vbCritical
        GoTo Cleanup
    End If

    ' ---------- 日付ヘッダを取得 ----------
    Dim hdr As Variant
    hdr = sCl.Range(sCl.Cells(R_DATE, C_TODAY), sCl.Cells(R_DATE, C_HISTEND)).Value

    ' 日曜/土曜の列を検出して報告（V805 は 7/26(日) が混ざっている）
    Dim weekendMsg As String, wi As Long
    For wi = 1 To UBound(hdr, 2)
        If IsDate(hdr(1, wi)) Then
            If Weekday(CDate(hdr(1, wi)), vbMonday) >= 6 Then
                weekendMsg = weekendMsg & "  " & ColLetter(C_TODAY + wi - 1) & "列 " & _
                             Format(CDate(hdr(1, wi)), "yyyy/m/d(aaa)") & vbCrLf
            End If
        End If
    Next wi

    ' ---------- 移植先の コード→行 辞書 ----------
    Dim dCl As Worksheet
    On Error Resume Next
    Set dCl = ThisWorkbook.Sheets("終値")
    On Error GoTo Cleanup
    If dCl Is Nothing Then MsgBox "このブックに 終値シートがありません。", vbCritical: GoTo Cleanup

    Dim dstRow As Object: Set dstRow = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = R_FIRST To R_LAST
        Dim dc As String: dc = Trim$(CStr(dCl.Cells(r, C_CODE).Value))
        If dc <> "" And dc <> "0" Then
            If Not dstRow.Exists(dc) Then dstRow.Add dc, r
        End If
    Next r

    ' ---------- 保護解除 ----------
    UnprotectAll ThisWorkbook

    ' ---------- シートごとに移植 ----------
    Dim sh As Variant: sh = ShNames()
    Dim s As Integer
    Dim matched As Long, notFound As String, nNotFound As Long
    Dim srcCodes As Object: Set srcCodes = CreateObject("Scripting.Dictionary")

    For s = 0 To 4
        Application.StatusBar = "移植中: " & sh(s) & " ..."

        Dim sw As Worksheet, dw As Worksheet
        Set sw = Nothing: Set dw = Nothing
        On Error Resume Next
        Set sw = src.Sheets(sh(s))
        Set dw = ThisWorkbook.Sheets(sh(s))
        On Error GoTo Cleanup
        If sw Is Nothing Or dw Is Nothing Then
            report = report & sh(s) & ": シートなし(スキップ)" & vbCrLf
            GoTo NextS
        End If

        ' 履歴領域をいったん全消し
        dw.Range(dw.Cells(R_TOPIX, C_TODAY), dw.Cells(R_LAST, C_HISTEND)).ClearContents

        ' 日付ヘッダを置き換え
        dw.Range(dw.Cells(R_DATE, C_TODAY), dw.Cells(R_DATE, C_HISTEND)).Value = hdr
        dw.Range(dw.Cells(R_DATE, C_TODAY), dw.Cells(R_DATE, C_HISTEND)).NumberFormat = "m/d"

        ' TOPIX 行(5) は行番号そのままで移す
        dw.Range(dw.Cells(R_TOPIX, C_TODAY), dw.Cells(R_TOPIX, C_HISTEND)).Value = _
            sw.Range(sw.Cells(R_TOPIX, C_TODAY), sw.Cells(R_TOPIX, C_HISTEND)).Value

        ' 銘柄行はコードで突き合わせ
        Dim n As Long: n = 0
        For r = R_FIRST To R_LAST
            Dim code As String
            code = Trim$(CStr(sw.Cells(r, C_CODE).Value))
            If code = "" Or code = "0" Then GoTo NextR
            If Not IsStockCode(code) Then GoTo NextR      ' ゴミ行は移さない

            If s = 0 Then
                If Not srcCodes.Exists(code) Then srcCodes.Add code, r
            End If

            If dstRow.Exists(code) Then
                Dim tr As Long: tr = CLng(dstRow(code))
                dw.Range(dw.Cells(tr, C_TODAY), dw.Cells(tr, C_HISTEND)).Value = _
                    sw.Range(sw.Cells(r, C_TODAY), sw.Cells(r, C_HISTEND)).Value
                n = n + 1
            Else
                If s = 0 Then
                    nNotFound = nNotFound + 1
                    If Len(notFound) < 200 Then notFound = notFound & code & " "
                End If
            End If
NextR:
        Next r

        report = report & sh(s) & ": " & n & " 銘柄" & vbCrLf
        If s = 0 Then matched = n
NextS:
    Next s

    ' ---------- 移植先にしかないコード ----------
    Dim onlyDst As String, nOnlyDst As Long
    Dim k As Variant
    For Each k In dstRow.Keys
        If Not srcCodes.Exists(CStr(k)) Then
            nOnlyDst = nOnlyDst + 1
            If Len(onlyDst) < 200 Then onlyDst = onlyDst & CStr(k) & " "
        End If
    Next k

Cleanup:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description

    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    ProtectAll ThisWorkbook
    Application.AutomationSecurity = prevSec
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.StatusBar = False
    On Error GoTo 0

    If eN <> 0 Then
        MsgBox "移植中にエラー: " & eN & " " & eD & vbCrLf & vbCrLf & _
               "計算モード・保護は戻しました。" & vbCrLf & _
               "このブックは中途半端な状態の可能性があります。" & vbCrLf & _
               "保存せずに閉じて、コピーからやり直してください。", vbCritical, "エラー"
        Exit Sub
    End If

    Dim msg As String
    msg = "【OHLCV 移植 完了】" & vbCrLf & String(38, "-") & vbCrLf & _
          report & String(38, "-") & vbCrLf & _
          "移植した銘柄: " & matched & vbCrLf

    If nNotFound > 0 Then
        msg = msg & vbCrLf & "★移植元にあってこのブックに無いコード: " & nNotFound & vbCrLf & _
              "  " & Trim$(notFound) & vbCrLf & _
              "  → 銘柄管理に追加してからもう一度実行すると入ります。" & vbCrLf
    End If
    If nOnlyDst > 0 Then
        msg = msg & vbCrLf & "★このブックにあって移植元に無いコード: " & nOnlyDst & vbCrLf & _
              "  " & Trim$(onlyDst) & vbCrLf & _
              "  → この銘柄の履歴は空のままです。" & vbCrLf
    End If
    If weekendMsg <> "" Then
        msg = msg & vbCrLf & "★日付ヘッダに土日の列があります（移植元の既知の不具合）:" & vbCrLf & _
              weekendMsg & _
              "  → 日曜列を修正 で直せます（7/26→7/24）。" & vbCrLf
    End If

    msg = msg & vbCrLf & "この後:" & vbCrLf & _
          " ① 四本値_整合性チェック で違反 0 を確認" & vbCrLf & _
          " ② Workbook_Open の 日次更新_本日分 を止めたか確認"

    CopyToClip msg
    MsgBox msg, vbInformation, "移植完了"
End Sub


' ==================================================================
'  日曜/土曜の日付ヘッダを直前の営業日に直す
'  （V805 の 7/26(日) → 7/24(金)）
' ==================================================================
Public Sub 日曜列を修正()
    Dim cw As Worksheet
    On Error Resume Next
    Set cw = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If cw Is Nothing Then MsgBox "終値シートがありません。", vbCritical: Exit Sub

    ' 対象列を洗い出す
    Dim list As String, cols As String, n As Long
    Dim c As Long
    For c = C_TODAY To C_HISTEND
        Dim v As Variant: v = cw.Cells(R_DATE, c).Value
        If IsDate(v) Then
            Dim d As Date: d = CDate(v)
            If Weekday(d, vbMonday) >= 6 Then
                Dim nd As Date: nd = d
                Do While Weekday(nd, vbMonday) >= 6
                    nd = nd - 1
                Loop
                n = n + 1
                list = list & "  " & ColLetter(c) & "列  " & _
                       Format(d, "m/d(aaa)") & " → " & Format(nd, "m/d(aaa)") & vbCrLf
                cols = cols & c & ","
            End If
        End If
    Next c

    If n = 0 Then MsgBox "土日の日付ヘッダはありません。", vbInformation: Exit Sub

    If MsgBox("土日になっている日付ヘッダを、直前の営業日に直します。" & vbCrLf & vbCrLf & _
              list & vbCrLf & _
              "5シートすべてに適用します。続行しますか?", _
              vbYesNo + vbQuestion, "日曜列を修正") <> vbYes Then Exit Sub

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    UnprotectAll ThisWorkbook

    Dim sh As Variant: sh = ShNames()
    Dim s As Integer
    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sh(s))
        On Error GoTo ErrH
        If ws Is Nothing Then GoTo NextS

        Dim p As Variant
        For Each p In Split(cols, ",")
            If Trim$(CStr(p)) <> "" Then
                Dim cc As Long: cc = CLng(p)
                Dim vv As Variant: vv = ws.Cells(R_DATE, cc).Value
                If IsDate(vv) Then
                    Dim dd As Date: dd = CDate(vv)
                    Do While Weekday(dd, vbMonday) >= 6
                        dd = dd - 1
                    Loop
                    ws.Cells(R_DATE, cc).Value = dd
                    ws.Cells(R_DATE, cc).NumberFormat = "m/d"
                End If
            End If
        Next p
NextS:
    Next s

    ProtectAll ThisWorkbook
    Application.ScreenUpdating = True
    MsgBox n & " 列の日付を直しました。", vbInformation, "完了"
    Exit Sub

ErrH:
    ProtectAll ThisWorkbook
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Number & " " & Err.Description, vbCritical
End Sub


' ==================================================================
'  四本値 整合性チェック（移植後の検証用）
' ==================================================================
Public Sub 四本値_整合性チェック()
    Dim wo As Worksheet, wh As Worksheet, wl As Worksheet, wc As Worksheet
    On Error Resume Next
    Set wo = ThisWorkbook.Sheets("始値"): Set wh = ThisWorkbook.Sheets("高値")
    Set wl = ThisWorkbook.Sheets("安値"): Set wc = ThisWorkbook.Sheets("終値")
    On Error GoTo 0
    If wo Is Nothing Or wh Is Nothing Or wl Is Nothing Or wc Is Nothing Then
        MsgBox "OHLC シートが揃っていません。", vbCritical: Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim chk As Long, bad As Long, worst As String
    Dim r As Long, c As Long
    For r = R_FIRST To R_LAST
        If Trim$(CStr(wc.Cells(r, C_CODE).Value)) = "" Then GoTo NextR
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
                If Len(worst) < 200 Then worst = worst & wc.Cells(r, C_CODE).Value & "/" & ColLetter(c) & " "
            End If
NextC:
        Next c
NextR:
    Next r
    Application.ScreenUpdating = True

    MsgBox "【四本値 整合性チェック】" & vbCrLf & String(30, "=") & vbCrLf & _
           "検証セル: " & Format(chk, "#,##0") & vbCrLf & _
           "違反セル: " & bad & vbCrLf & vbCrLf & _
           IIf(bad = 0, "問題なし。", "★違反あり:" & vbCrLf & worst), _
           IIf(bad = 0, vbInformation, vbExclamation), "整合性チェック"
End Sub


' ==================================================================
'  補助
' ==================================================================
Private Function PickFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "移植元の V805 ファイルを選んでください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel マクロ有効ブック", "*.xlsm;*.xlsx;*.xlsb"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Sub UnprotectAll(ByVal wb As Workbook)
    Dim sh As Variant: sh = ShNames()
    Dim s As Integer
    For s = 0 To 4
        On Error Resume Next
        wb.Sheets(sh(s)).Unprotect Password:=PWD
        On Error GoTo 0
    Next s
End Sub

Private Sub ProtectAll(ByVal wb As Workbook)
    Dim sh As Variant: sh = ShNames()
    Dim s As Integer
    For s = 0 To 4
        On Error Resume Next
        wb.Sheets(sh(s)).Protect Password:=PWD, UserInterfaceOnly:=True, _
            DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    Next s
End Sub

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

Private Function ColLetter(ByVal c As Long) As String
    If c < 1 Then ColLetter = "?": Exit Function
    ColLetter = Split(ThisWorkbook.Sheets(1).Cells(1, c).Address(True, False), "$")(0)
End Function

Private Sub CopyToClip(ByVal s As String)
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    o.SetText s
    o.PutInClipboard
    On Error GoTo 0
End Sub
