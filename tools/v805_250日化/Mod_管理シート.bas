Attribute VB_Name = "Mod_管理シート"
Option Explicit

' ==============================================================================
'  管理シート（仮想売買の記録）
'
'   管理シート_全消去        … 過去の記録を消して、まっさらにする
'                              （消す前に自動でバックアップシートを作ります）
'   本日の候補を管理へ追加    … 厳選TOP2 の候補を管理シートに書き足す
'   管理シート_成績          … いまの記録の成績を集計して表示する
'   管理シートにボタンを置く  … 上の3つをボタンにして管理シートに並べる
'                              （最初に1回だけ実行してください）
'
'  管理シートで手入力するのは次の9列だけです。ほかは全部数式です。
'    B コード / C 売買 / E 購入日 / G 購入数量
'    T 売却日 / U 売却価格
'    Z スコア / AA シグナル内容 / AD 抽出元
'
'  F（購入価格）は購入日の始値を価格シートから自動で拾うので入力しません。
' ==============================================================================

Private Const SH_KAN  As String = "管理"
Private Const SH_TOP  As String = "厳選TOP2"
Private Const PWD     As String = "ne19480314"
Private Const ROW1    As Long = 4      ' 最初のデータ行
Private Const ROWN    As Long = 461    ' 最後のデータ行

' 手入力の列（これ以外は数式なので触らない）
Private Function InputCols() As Variant
    InputCols = Array(2, 3, 5, 7, 20, 21, 26, 27, 30)   ' B C E G T U Z AA AD
End Function


'==============================================================================
' ① 過去の記録を消す
'==============================================================================
Public Sub 管理シート_全消去()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_KAN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SH_KAN & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim n As Long, r As Long
    For r = ROW1 To ROWN
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then n = n + 1
    Next r
    If n = 0 Then
        MsgBox "管理シートには既に記録がありません。", vbInformation
        Exit Sub
    End If

    If MsgBox("管理シートの記録 " & n & " 件を消します。" & vbCrLf & vbCrLf & _
              "消す前にバックアップシートを作ります。" & vbCrLf & _
              "数式（利確ライン・損切ライン・判断など）は残します。" & vbCrLf & vbCrLf & _
              "続けますか？", vbYesNo + vbExclamation, "管理シートの消去") <> vbYes Then Exit Sub

    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Dim prevAlt As Boolean: prevAlt = Application.DisplayAlerts
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- バックアップ（値だけの複製）---
    Dim bkName As String
    bkName = "管理_控" & Format$(Now, "mmdd_hhnn")
    Dim bk As Worksheet
    Set bk = ThisWorkbook.Sheets.Add(After:=ws)
    bk.Name = bkName
    ws.Range(ws.Cells(1, 1), ws.Cells(ROWN, 40)).Copy
    bk.Range("A1").PasteSpecial xlPasteValues
    bk.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    bk.Range("A1").Select

    ' --- 消す ---
    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo ErrHandler

    Dim cols As Variant: cols = InputCols()
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ws.Range(ws.Cells(ROW1, cols(i)), ws.Cells(ROWN, cols(i))).ClearContents
    Next i

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo ErrHandler
    End If

    ws.Activate
    ws.Range("B4").Select
    Application.ScreenUpdating = prevScr
    Application.DisplayAlerts = prevAlt

    MsgBox n & " 件を消しました。" & vbCrLf & vbCrLf & _
           "バックアップ: 「" & bkName & "」シート" & vbCrLf & vbCrLf & _
           "これから、厳選TOP2 の候補を「本日の候補を管理へ追加」で" & vbCrLf & _
           "足していってください。", vbInformation, "消去完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = prevScr
    Application.DisplayAlerts = prevAlt
    MsgBox "消去を中断しました。" & vbCrLf & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbCritical, "エラー"
End Sub


'==============================================================================
' ② 厳選TOP2 の候補を管理シートに書き足す
'
'   購入日は「翌営業日」を入れます。購入価格(F列)はその日の始値を
'   価格シートから自動で拾うので、入力しません。
'==============================================================================
Public Sub 本日の候補を管理へ追加()
    Dim kws As Worksheet, tws As Worksheet
    Set kws = Nothing: Set tws = Nothing
    On Error Resume Next
    Set kws = ThisWorkbook.Sheets(SH_KAN)
    Set tws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo 0
    If kws Is Nothing Or tws Is Nothing Then
        MsgBox "管理シートまたは厳選TOP2シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' --- 候補を読む（買い候補は 18行目から5行）---
    Dim codes(1 To 5) As String, nmz(1 To 5) As String
    Dim scr(1 To 5) As String, sig(1 To 5) As String
    Dim cnt As Long, i As Long
    For i = 1 To 5
        Dim r As Long: r = 17 + i
        Dim cd As String: cd = Trim$(CStr(tws.Cells(r, 2).Value))
        If cd <> "" And cd <> "0" Then
            cnt = cnt + 1
            codes(cnt) = cd
            nmz(cnt) = Trim$(CStr(tws.Cells(r, 3).Value))
            scr(cnt) = "厳選 " & Trim$(CStr(tws.Cells(r, 5).Value)) & "点"
            sig(cnt) = Trim$(CStr(tws.Cells(r, 8).Value)) & _
                       " RSI" & Format$(tws.Cells(r, 7).Value, "0.0") & _
                       " 出来高" & Format$(tws.Cells(r, 6).Value, "0.00") & "倍" & _
                       " 前日比" & Format$(tws.Cells(r, 10).Value, "+0.0;-0.0") & "%"
        End If
    Next i

    If cnt = 0 Then
        MsgBox "厳選TOP2 に買い候補がありません。" & vbCrLf & vbCrLf & _
               "・15:35 より前だと出来高が足りず該当なしになります" & vbCrLf & _
               "・大引け後でも0件なら、今日は見送りです", vbInformation, "候補なし"
        Exit Sub
    End If

    ' --- 購入日 = 翌営業日 ---
    Dim buyDate As Date
    buyDate = Application.WorksheetFunction.WorkDay(Date, 1)

    ' --- 数量 = 許容損失額 ÷ (株価 × 損切幅) を100株単位で ---
    Dim lossAmt As Double, slPct As Double
    lossAmt = 20000: slPct = 0.04
    On Error Resume Next
    lossAmt = CDbl(ThisWorkbook.Sheets("V811設定").Range("B5").Value)
    slPct = CDbl(ThisWorkbook.Sheets("V811設定").Range("B3").Value)
    Err.Clear
    On Error GoTo 0
    If lossAmt <= 0 Then lossAmt = 20000
    If slPct <= 0 Then slPct = 0.04

    Dim msg As String
    msg = "厳選TOP2 の買い候補 " & cnt & " 件を管理シートに追加します。" & vbCrLf & _
          "購入日: " & Format$(buyDate, "yyyy/m/d") & "（翌営業日）" & vbCrLf & _
          String(34, "-") & vbCrLf
    For i = 1 To cnt
        Dim px As Double: px = Val(CStr(tws.Cells(17 + i, 4).Value))
        Dim qty As Long: qty = 100
        If px > 0 Then qty = Application.WorksheetFunction.Max(100, _
                             Int(lossAmt / (px * slPct) / 100) * 100)
        msg = msg & "  " & codes(i) & " " & nmz(i) & "   " & qty & " 株" & vbCrLf
    Next i
    msg = msg & String(34, "-") & vbCrLf & _
          "購入価格は " & Format$(buyDate, "m/d") & " の始値を" & vbCrLf & _
          "価格シートから自動で拾います（翌日以降に入ります）。" & vbCrLf & vbCrLf & _
          "追加しますか？"
    If MsgBox(msg, vbYesNo + vbQuestion, "候補の追加") <> vbYes Then Exit Sub

    ' --- 空き行を探して書き込む ---
    Dim wasP As Boolean: wasP = kws.ProtectContents
    On Error Resume Next
    kws.Unprotect Password:=PWD
    kws.Unprotect
    Err.Clear
    On Error GoTo 0

    Dim wr As Long: wr = ROW1
    Do While wr <= ROWN
        If Trim$(CStr(kws.Cells(wr, 2).Value)) = "" Then Exit Do
        wr = wr + 1
    Loop

    Dim added As Long, skipped As String
    For i = 1 To cnt
        If wr > ROWN Then Exit For
        ' 同じ銘柄を同じ日に二重に入れない
        Dim dup As Boolean: dup = False
        Dim k As Long
        For k = ROW1 To wr - 1
            If Trim$(CStr(kws.Cells(k, 2).Value)) = codes(i) Then
                If kws.Cells(k, 20).Value = "" Then dup = True: Exit For   ' T列が空＝保有中
            End If
        Next k
        If dup Then
            skipped = skipped & "  " & codes(i) & " " & nmz(i) & "（保有中）" & vbCrLf
        Else
            Dim px2 As Double: px2 = Val(CStr(tws.Cells(17 + i, 4).Value))
            Dim qty2 As Long: qty2 = 100
            If px2 > 0 Then qty2 = Application.WorksheetFunction.Max(100, _
                                   Int(lossAmt / (px2 * slPct) / 100) * 100)
            kws.Cells(wr, 2).Value = codes(i)          ' B コード
            kws.Cells(wr, 3).Value = "買"              ' C 売買
            kws.Cells(wr, 5).Value = buyDate           ' E 購入日
            kws.Cells(wr, 7).Value = qty2              ' G 購入数量
            kws.Cells(wr, 26).Value = scr(i)           ' Z スコア
            kws.Cells(wr, 27).Value = sig(i)           ' AA シグナル内容
            kws.Cells(wr, 30).Value = "厳選TOP2"       ' AD 抽出元
            wr = wr + 1: added = added + 1
        End If
    Next i

    If wasP Then
        On Error Resume Next
        kws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                    DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If

    kws.Activate
    Dim done As String
    done = added & " 件を追加しました。"
    If skipped <> "" Then done = done & vbCrLf & vbCrLf & "飛ばした銘柄:" & vbCrLf & skipped
    done = done & vbCrLf & "購入価格(F列)は " & Format$(buyDate, "m/d") & " の始値が" & vbCrLf & _
           "価格シートに入った時点で自動で表示されます。"
    MsgBox done, vbInformation, "追加完了"
End Sub


'==============================================================================
' ③ いまの成績を集計する
'==============================================================================
Public Sub 管理シート_成績()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_KAN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim r As Long, nAll As Long, nDone As Long, nHold As Long
    Dim win As Long, los As Long
    Dim sumWin As Double, sumLos As Double, total As Double
    Dim maxLose As Long, curLose As Long
    Dim eq As Double, peak As Double, dd As Double

    For r = ROW1 To ROWN
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "" Then GoTo NextR
        nAll = nAll + 1
        Dim v As Variant: v = ws.Cells(r, 22).Value      ' V 確定損益
        If Not IsNumeric(v) Or CStr(v) = "" Then
            nHold = nHold + 1
        Else
            nDone = nDone + 1
            Dim p As Double: p = CDbl(v)
            total = total + p
            eq = eq + p
            If eq > peak Then peak = eq
            If eq - peak < dd Then dd = eq - peak
            If p > 0 Then
                win = win + 1: sumWin = sumWin + p: curLose = 0
            Else
                los = los + 1: sumLos = sumLos + p
                curLose = curLose + 1
                If curLose > maxLose Then maxLose = curLose
            End If
        End If
NextR:
    Next r

    If nDone = 0 Then
        MsgBox "決済済みの記録がまだありません。" & vbCrLf & _
               "（登録 " & nAll & " 件 / 保有中 " & nHold & " 件）", vbInformation, "成績"
        Exit Sub
    End If

    Dim pf As String
    If sumLos <> 0 Then pf = Format$(sumWin / Abs(sumLos), "0.00") Else pf = "―"

    MsgBox "管理シートの成績" & vbCrLf & String(32, "-") & vbCrLf & _
           "  登録:      " & nAll & " 件" & vbCrLf & _
           "  決済済:    " & nDone & " 件" & vbCrLf & _
           "  保有中:    " & nHold & " 件" & vbCrLf & vbCrLf & _
           "  勝ち:      " & win & " 件" & vbCrLf & _
           "  負け:      " & los & " 件" & vbCrLf & _
           "  勝率:      " & Format$(100# * win / nDone, "0.0") & " %" & vbCrLf & vbCrLf & _
           "  合計損益:  " & Format$(total, "#,##0") & " 円" & vbCrLf & _
           "  1回平均:   " & Format$(total / nDone, "#,##0") & " 円" & vbCrLf & _
           "  平均利益:  " & Format$(IIf(win > 0, sumWin / win, 0), "#,##0") & " 円" & vbCrLf & _
           "  平均損失:  " & Format$(IIf(los > 0, sumLos / los, 0), "#,##0") & " 円" & vbCrLf & _
           "  PF:        " & pf & vbCrLf & vbCrLf & _
           "  最大DD:    " & Format$(dd, "#,##0") & " 円" & vbCrLf & _
           "  最大連敗:  " & maxLose & " 回", _
           vbInformation, "成績"
End Sub


'==============================================================================
' ④ 管理シートにボタンを置く（最初に1回だけ実行してください）
'==============================================================================
Public Sub 管理シートにボタンを置く()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_KAN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SH_KAN & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo 0

    ' 同名のボタンがあれば消してから作り直す（押すたびに増えないように）
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Select Case ws.Shapes(i).Name
            Case "btn候補追加", "btn成績", "btn全消去": ws.Shapes(i).Delete
        End Select
    Next i

    Dim tl As Range: Set tl = ws.Range("N2")
    Dim x As Double: x = tl.Left
    Dim y As Double: y = tl.Top + 6

    MakeBtn ws, "btn候補追加", "本日の候補を追加", "本日の候補を管理へ追加", _
            x, y, 210, 56, RGB(0, 112, 192), RGB(0, 70, 130), 16
    x = x + 218
    MakeBtn ws, "btn成績", "成績を見る", "管理シート_成績", _
            x, y, 140, 56, RGB(0, 130, 90), RGB(0, 90, 60), 15
    x = x + 148
    MakeBtn ws, "btn全消去", "全消去", "管理シート_全消去", _
            x, y, 110, 56, RGB(150, 150, 150), RGB(100, 100, 100), 14

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If

    ws.Activate
    MsgBox "管理シートにボタンを3つ置きました。（2行目のN列あたり）" & vbCrLf & vbCrLf & _
           "  青  本日の候補を追加 … 毎日15:35以降に押す" & vbCrLf & _
           "  緑  成績を見る       … 勝率・PF・最大DDを表示" & vbCrLf & _
           "  灰  全消去           … 記録を消す（控えは自動で残ります）" & vbCrLf & vbCrLf & _
           "位置は右クリックしてドラッグで動かせます。", _
           vbInformation, "設置完了"
End Sub


Private Sub MakeBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, _
                    ByVal act As String, ByVal x As Double, ByVal y As Double, _
                    ByVal w As Double, ByVal h As Double, _
                    ByVal fillRGB As Long, ByVal lineRGB As Long, ByVal fsize As Single)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With sh
        .Name = nm
        .Fill.ForeColor.RGB = fillRGB
        .Line.ForeColor.RGB = lineRGB
        .Line.Weight = 1.5
        With .TextFrame2
            .TextRange.Text = cap
            .TextRange.Font.Size = fsize
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0: .MarginBottom = 0
            .WordWrap = msoTrue
        End With
        .OnAction = act
    End With
End Sub
