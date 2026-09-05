Attribute VB_Name = "Mod_HeikinAshi_Log"
Option Explicit
'==================================================================
' 平均足　トレード記録シート　v1.0
'
'   実行 : Alt+F8 →「平均足_記録シート作成」
'   このブックの中に「平均足記録」シートを作ります。
'   ※既にある場合は、入力した記録を残したまま計算式だけ直します。
'
'   入力するのは 黄色いセルだけ です。
'     買った時 : 買日 / コード / 銘柄名 / 買値 / 株数
'     売れた時 : 売日 / 売値
'   あとは全部 自動で計算されます。
'==================================================================

Private Const SH As String = "平均足記録"
Private Const TOP_ROW As Long = 5      'データ開始行
Private Const MAX_ROW As Long = 204    'データ最終行（200件）
Private Const HOLD_DAYS As Long = 5    '何営業日で手じまいするか

Public Sub 平均足_記録シート作成()

    Dim ws As Worksheet, isNew As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH
        isNew = True
    End If

    Application.ScreenUpdating = False

    '--- 入力済みの記録を退避（作り直しでも消さない）---
    Dim keep As Variant
    keep = ws.Range(ws.Cells(TOP_ROW, 1), ws.Cells(MAX_ROW, 5)).Value
    Dim keep2 As Variant
    keep2 = ws.Range(ws.Cells(TOP_ROW, 10), ws.Cells(MAX_ROW, 11)).Value
    Dim keepMemo As Variant
    keepMemo = ws.Range(ws.Cells(TOP_ROW, 16), ws.Cells(MAX_ROW, 16)).Value
    Dim keepCap As Double, keepLot As Double
    keepCap = HA_L_Num(ws.Cells(2, 2).Value, 3000000#)
    keepLot = HA_L_Num(ws.Cells(2, 4).Value, 0.1)
    If keepLot <= 0 Or keepLot > 1 Then keepLot = 0.1

    ws.Cells.UnMerge
    ws.Cells.Clear

    '=============== 見出し ===============
    With ws.Range("A1:S1")
        .Merge
        .Value = "  平均足　トレード記録　（黄色いセルだけ入力してください）"
        .Font.Name = "Meiryo UI": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(1).RowHeight = 32

    '=============== 資金の設定 ===============
    ws.Cells(2, 1).Value = "資金"
    ws.Cells(2, 2).Value = keepCap
    ws.Cells(2, 3).Value = "1銘柄の割合"
    ws.Cells(2, 4).Value = keepLot
    ws.Cells(2, 5).Value = "1銘柄の投資額"
    ws.Cells(2, 6).Formula = "=B2*D2"
    ws.Cells(2, 7).Value = "1回の損切額"
    ws.Cells(2, 8).Formula = "=-F2*0.08"
    ws.Cells(2, 9).Value = "月間ストップ"
    ws.Cells(2, 10).Formula = "=-B2*0.06"
    ws.Cells(2, 11).Value = "← ここまで負けたら今月は休む"

    ws.Cells(2, 2).NumberFormat = "#,##0""円"""
    ws.Cells(2, 4).NumberFormat = "0%"
    ws.Cells(2, 6).NumberFormat = "#,##0""円"""
    ws.Cells(2, 8).NumberFormat = "#,##0""円"""
    ws.Cells(2, 10).NumberFormat = "#,##0""円"""
    HA_L_Input ws.Range("B2")
    HA_L_Input ws.Range("D2")

    '=============== 成績のまとめ ===============
    ws.Cells(3, 1).Value = "終わった取引"
    ws.Cells(3, 2).Formula = "=COUNT(K" & TOP_ROW & ":K" & MAX_ROW & ")"
    ws.Cells(3, 3).Value = "勝率"
    ws.Cells(3, 4).Formula = "=IF(B3=0,"""",COUNTIFS(K" & TOP_ROW & ":K" & MAX_ROW & ",""<>"",L" & TOP_ROW & ":L" & MAX_ROW & ","">0"")/B3)"
    ws.Cells(3, 5).Value = "合計損益"
    ws.Cells(3, 6).Formula = "=SUM(L" & TOP_ROW & ":L" & MAX_ROW & ")"
    ws.Cells(3, 7).Value = "今月の損益"
    ws.Cells(3, 8).Formula = "=SUMIFS(L" & TOP_ROW & ":L" & MAX_ROW & ",J" & TOP_ROW & ":J" & MAX_ROW & ","">=""&EOMONTH(TODAY(),-1)+1,J" & TOP_ROW & ":J" & MAX_ROW & ",""<=""&EOMONTH(TODAY(),0))"
    ws.Cells(3, 9).Value = "連敗"
    ws.Cells(3, 10).Formula = "=IFERROR(LOOKUP(9.99E+307,Q" & TOP_ROW & ":Q" & MAX_ROW & "),0)"
    ws.Cells(3, 11).Value = "保有中"
    ws.Cells(3, 12).Formula = "=COUNTIF(N" & TOP_ROW & ":N" & MAX_ROW & ",""保有中"")"
    ws.Cells(3, 13).Value = "直近20回の平均"
    ws.Cells(3, 14).Formula = "=IFERROR(AVERAGEIFS(M" & TOP_ROW & ":M" & MAX_ROW & ",R" & TOP_ROW & ":R" & MAX_ROW & ","">=""&MAX(R" & TOP_ROW & ":R" & MAX_ROW & ")-19),"""")"
    ws.Cells(3, 14).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = "今の指示"
    With ws.Range("P3:S3")
        .Merge
        .Formula = "=IF(H3<=$J$2,""今月はもう建てない（月間ストップ）""," & _
                   "IF(AND($B$3>=20,$N$3<=0),""直近20回がマイナス：プラスに戻るまで休む""," & _
                   "IF($J$3>=3,""3連敗中：株数を半分にする""," & _
                   "IF($L$3>=5,""枠が埋まっています"",""通常どおり。上位から買ってよい""))))"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With

    ws.Cells(3, 4).NumberFormat = "0.0%"
    ws.Cells(3, 6).NumberFormat = "#,##0""円"""
    ws.Cells(3, 8).NumberFormat = "#,##0""円"""

    With ws.Range("A2:S3")
        .Font.Name = "Meiryo UI": .Font.Size = 10
        .Interior.Color = RGB(226, 239, 218)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
    End With
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 22

    '=============== 見出し行 ===============
    Dim hd As Variant
    hd = Array("買日", "コード", "銘柄名", "買値", "株数", "投資額", "損切値", "利確値", _
               "期限日", "売日", "売値", "損益(円)", "損益率", "結果", "資金比", "メモ", "連敗", "通番")
    Dim j As Long
    For j = 0 To UBound(hd)
        ws.Cells(4, j + 1).Value = hd(j)
    Next j
    With ws.Range("A4:R4")
        .Font.Name = "Meiryo UI": .Font.Size = 10: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(4).RowHeight = 30

    '=============== 各行の計算式 ===============
    Dim r As Long
    For r = TOP_ROW To MAX_ROW
        ws.Cells(r, 6).Formula = "=IF(OR(D" & r & "="""",E" & r & "=""""),"""",D" & r & "*E" & r & ")"
        ws.Cells(r, 7).Formula = "=IF(D" & r & "="""","""",ROUND(D" & r & "*0.92,0))"
        ws.Cells(r, 8).Formula = "=IF(D" & r & "="""","""",ROUND(D" & r & "*1.08,0))"
        ws.Cells(r, 9).Formula = "=IF(A" & r & "="""","""",WORKDAY(A" & r & "," & HOLD_DAYS & "))"
        ws.Cells(r, 12).Formula = "=IF(OR(K" & r & "="""",D" & r & "="""",E" & r & "=""""),"""",(K" & r & "-D" & r & ")*E" & r & ")"
        ws.Cells(r, 13).Formula = "=IF(OR(K" & r & "="""",D" & r & "=""""),"""",(K" & r & "-D" & r & ")/D" & r & ")"
        ws.Cells(r, 14).Formula = "=IF(D" & r & "="""","""",IF(K" & r & "="""",""保有中"",IF(M" & r & ">=0.07,""利確"",IF(M" & r & "<=-0.075,""損切"",IF(M" & r & ">0,""時間切れ(勝)"",""時間切れ(負)"")))))"
        ws.Cells(r, 15).Formula = "=IF(L" & r & "="""","""",L" & r & "/$B$2)"
        If r = TOP_ROW Then
            ws.Cells(r, 17).Formula = "=IF(K" & r & "="""","""",IF(L" & r & "<=0,1,0))"
        Else
            ws.Cells(r, 17).Formula = "=IF(K" & r & "="""","""",IF(L" & r & "<=0,N(Q" & (r - 1) & ")+1,0))"
        End If
        ws.Cells(r, 18).Formula = "=IF(K" & r & "="""","""",COUNT($K$" & TOP_ROW & ":K" & r & "))"
    Next r

    ws.Range("A" & TOP_ROW & ":A" & MAX_ROW).NumberFormat = "yyyy/mm/dd"
    ws.Range("I" & TOP_ROW & ":I" & MAX_ROW).NumberFormat = "yyyy/mm/dd"
    ws.Range("J" & TOP_ROW & ":J" & MAX_ROW).NumberFormat = "yyyy/mm/dd"
    ws.Range("D" & TOP_ROW & ":H" & MAX_ROW).NumberFormat = "#,##0"
    ws.Range("K" & TOP_ROW & ":L" & MAX_ROW).NumberFormat = "#,##0"
    ws.Range("M" & TOP_ROW & ":M" & MAX_ROW).NumberFormat = "0.0%"
    ws.Range("O" & TOP_ROW & ":O" & MAX_ROW).NumberFormat = "0.00%"
    ws.Range("N" & TOP_ROW & ":N" & MAX_ROW).HorizontalAlignment = xlCenter

    '--- 入力するセルを黄色に ---
    HA_L_Input ws.Range("A" & TOP_ROW & ":E" & MAX_ROW)
    HA_L_Input ws.Range("J" & TOP_ROW & ":K" & MAX_ROW)
    HA_L_Input ws.Range("P" & TOP_ROW & ":P" & MAX_ROW)

    With ws.Range("A4:R" & MAX_ROW)
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    '--- 色分け（勝ち＝緑、負け＝赤、保有中＝青）---
    Dim rng As Range
    Set rng = ws.Range("A" & TOP_ROW & ":P" & MAX_ROW)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$N" & TOP_ROW & "=""保有中""")
        .Interior.Color = RGB(221, 235, 247)
    End With
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($L" & TOP_ROW & "<>"""",$L" & TOP_ROW & ">0)")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($L" & TOP_ROW & "<>"""",$L" & TOP_ROW & "<=0)")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    '--- 連敗の列は使わないので薄く ---
    ws.Columns("Q:R").Font.Color = RGB(200, 200, 200)

    '=============== 使い方 ===============
    Dim g As Long
    g = MAX_ROW + 2
    ws.Cells(g, 1).Value = "【使い方】"
    ws.Cells(g + 1, 1).Value = "1. 買ったら　買日・コード・銘柄名・買値・株数　の5つだけ入れる。損切値と利確値が自動で出ます。"
    ws.Cells(g + 2, 1).Value = "2. その値段で　逆指値（損切値）と　売り指値（利確値）を証券会社に出す。"
    ws.Cells(g + 3, 1).Value = "3. 売れたら　売日・売値　を入れる。損益と勝率が自動で更新されます。"
    ws.Cells(g + 4, 1).Value = "4. 「期限日」（買った日から5営業日後）の引けで、勝ち負けに関係なく成行手仕舞い。"
    ws.Cells(g + 5, 1).Value = ""
    ws.Cells(g + 6, 1).Value = "【怖くなったら、この数字を見てください（300銘柄×250日の検証）】"
    ws.Cells(g + 7, 1).Value = "・負けの平均は －5.19%。－8%満額で切られるのは17回に1回だけです。"
    ws.Cells(g + 8, 1).Value = "・最大の連敗は 4連敗。5連敗以上は1年で一度もありませんでした。"
    ws.Cells(g + 9, 1).Value = "・10回中4回は負けます。それが設計です。負けは家賃だと思ってください。"
    ws.Cells(g + 10, 1).Value = "・「直近20回の平均」がマイナスになったら、プラスに戻るまで休む（自動で指示が出ます）。"
    ws.Cells(g + 11, 1).Value = "・1銘柄は30万円以上にする。15万円では候補の2割しか買えません（候補の株価は平均4,100円）。"
    ws.Cells(g + 12, 1).Value = "・小さくやりたい時は割合を下げず、運用資金そのものを減らす（例：150万円で1銘柄20%=30万円）。"
    With ws.Range("A" & g & ":A" & (g + 12))
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    ws.Cells(g, 1).Font.Bold = True
    ws.Cells(g + 6, 1).Font.Bold = True

    '--- 幅 ---
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D:H").ColumnWidth = 9
    ws.Columns("I:J").ColumnWidth = 11
    ws.Columns("K:L").ColumnWidth = 10
    ws.Columns("M:O").ColumnWidth = 9
    ws.Columns("P").ColumnWidth = 20
    ws.Columns("Q:R").ColumnWidth = 5

    '--- 退避した記録を書き戻す ---
    If Not isNew Then
        ws.Range(ws.Cells(TOP_ROW, 1), ws.Cells(MAX_ROW, 5)).Value = keep
        ws.Range(ws.Cells(TOP_ROW, 10), ws.Cells(MAX_ROW, 11)).Value = keep2
        ws.Range(ws.Cells(TOP_ROW, 16), ws.Cells(MAX_ROW, 16)).Value = keepMemo
    End If

    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & TOP_ROW).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "「" & SH & "」シートを作りました。" & vbCrLf & vbCrLf & _
           "黄色いセルだけ入力してください。" & vbCrLf & _
           "まずは B2 の資金と、D2 の1銘柄の割合（最初は10%）を確認してください。", vbInformation
End Sub

'入力セルの色付け
Private Sub HA_L_Input(ByVal rg As Range)
    rg.Interior.Color = RGB(255, 255, 204)
    rg.Locked = False
End Sub

Private Function HA_L_Num(ByVal v As Variant, ByVal dflt As Double) As Double
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            HA_L_Num = CDbl(v)
            Exit Function
        End If
    End If
    HA_L_Num = dflt
End Function
