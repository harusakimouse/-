Attribute VB_Name = "Mod_HeikinAshi_Log"
Option Explicit
'==================================================================
' 平均足　トレード記録シート　v2.0
'
'   実行 : Alt+F8 →「平均足_記録シート作成」
'
'   ◆レイアウト
'     1行目とA列は空白（余白）
'     2～5行目 … タイトル／設定／集計／見出し
'                 ここは色やフォントを自由に変えて構いません。
'                 2回目以降の実行では書式を触りません（数式だけ直します）。
'     6行目以降 … 記録
'
'   ◆入力するのは黄色いセルだけ
'     買った時 : 買日 / コード / 銘柄名 / 買値 / 株数
'     売れた時 : 売日 / 売値
'
'   ◆書式を初期状態に戻したい時
'     Alt+F8 →「平均足_記録シート_書式を作り直す」
'==================================================================

Private Const SH As String = "平均足記録"
Private Const TOP_ROW As Long = 6      'データ開始行
Private Const MAX_ROW As Long = 205    'データ最終行（200件）
Private Const HD_ROW As Long = 5       '見出し行
Private Const HD_SIZE As Long = 18     '見出しの文字サイズ
Private Const DT_SIZE As Long = 16     'データの文字サイズ
Private Const HOLD_DAYS As Long = 5    '何営業日で手じまいするか
Private Const SL_RATE As Double = 0.08 '損切の幅
Private Const TP_RATE As Double = 0.08 '利確の幅

'列（A列は余白）
'B買日 C コード D銘柄名 E買値 F株数 G投資額 H損切値 I利確値 J期限日
'K売日 L売値 M損益(円) N損益率 O結果 P資金比 Q メモ R連敗 S通番

'==================== 入口 ====================
Public Sub 平均足_記録シート作成()
    HA_L_Build False
End Sub

Public Sub 平均足_記録シート_書式を作り直す()
    If MsgBox("色やフォントを初期状態に戻します。" & vbCrLf & _
              "入力した記録は消えません。よろしいですか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    HA_L_Build True
End Sub

'==================== 本体 ====================
Private Sub HA_L_Build(ByVal forceStyle As Boolean)

    Dim ws As Worksheet, isNew As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH
        isNew = True
    End If

    Dim doStyle As Boolean
    doStyle = isNew Or forceStyle

    Application.ScreenUpdating = False

    '--- 入力済みの記録を退避 ---
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim capV As Double, lotV As Double
    k1 = ws.Range(ws.Cells(TOP_ROW, 2), ws.Cells(MAX_ROW, 6)).Value      '買日～株数
    k2 = ws.Range(ws.Cells(TOP_ROW, 11), ws.Cells(MAX_ROW, 12)).Value    '売日・売値
    k3 = ws.Range(ws.Cells(TOP_ROW, 17), ws.Cells(MAX_ROW, 17)).Value    'メモ
    capV = HA_L_Num(ws.Cells(3, 3).Value, 3000000#)
    lotV = HA_L_Num(ws.Cells(3, 5).Value, 0.1)
    If lotV <= 0 Or lotV > 1 Then lotV = 0.1

    If doStyle Then
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If

    '=============== 2行目：タイトル ===============
    If doStyle Then
        With ws.Range("B2:S2")
            .Merge
            .Font.Name = "Meiryo UI": .Font.Size = HD_SIZE: .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 70, 127)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(1).RowHeight = 10
        ws.Rows(2).RowHeight = 38
        ws.Columns("A").ColumnWidth = 2
    End If
    ws.Range("B2").Value = "  平均足　トレード記録　（黄色いセルだけ入力してください）"

    '=============== 3行目：資金の設定 ===============
    ws.Cells(3, 2).Value = "資金"
    ws.Cells(3, 3).Value = capV
    ws.Cells(3, 4).Value = "1銘柄の割合"
    ws.Cells(3, 5).Value = lotV
    ws.Cells(3, 6).Value = "1銘柄の投資額"
    ws.Cells(3, 7).Formula = "=C3*E3"
    ws.Cells(3, 8).Value = "1回の損切額"
    ws.Cells(3, 9).Formula = "=-G3*" & SL_RATE
    ws.Cells(3, 10).Value = "月間ストップ"
    ws.Cells(3, 11).Formula = "=-C3*0.06"
    ws.Cells(3, 12).Value = "← ここまで負けたら今月は休む"

    '=============== 4行目：成績のまとめ ===============
    ws.Cells(4, 2).Value = "終わった取引"
    ws.Cells(4, 3).Formula = "=COUNT(L" & TOP_ROW & ":L" & MAX_ROW & ")"
    ws.Cells(4, 4).Value = "勝率"
    ws.Cells(4, 5).Formula = "=IF(C4=0,"""",COUNTIFS(L" & TOP_ROW & ":L" & MAX_ROW & ",""<>"",M" & TOP_ROW & ":M" & MAX_ROW & ","">0"")/C4)"
    ws.Cells(4, 6).Value = "合計損益"
    ws.Cells(4, 7).Formula = "=SUM(M" & TOP_ROW & ":M" & MAX_ROW & ")"
    ws.Cells(4, 8).Value = "今月の損益"
    ws.Cells(4, 9).Formula = "=SUMIFS(M" & TOP_ROW & ":M" & MAX_ROW & ",K" & TOP_ROW & ":K" & MAX_ROW & ","">=""&EOMONTH(TODAY(),-1)+1,K" & TOP_ROW & ":K" & MAX_ROW & ",""<=""&EOMONTH(TODAY(),0))"
    ws.Cells(4, 10).Value = "連敗"
    ws.Cells(4, 11).Formula = "=IFERROR(LOOKUP(9.99E+307,R" & TOP_ROW & ":R" & MAX_ROW & "),0)"
    ws.Cells(4, 12).Value = "保有中"
    ws.Cells(4, 13).Formula = "=COUNTIF(O" & TOP_ROW & ":O" & MAX_ROW & ",""保有中"")"
    ws.Cells(4, 14).Value = "直近20回の平均"
    ws.Cells(4, 15).Formula = "=IFERROR(AVERAGEIFS(N" & TOP_ROW & ":N" & MAX_ROW & ",S" & TOP_ROW & ":S" & MAX_ROW & ","">=""&MAX(S" & TOP_ROW & ":S" & MAX_ROW & ")-19),"""")"
    ws.Cells(4, 16).Value = "今の指示"
    If doStyle Then ws.Range("Q4:S4").Merge
    ws.Range("Q4").Formula = "=IF(I4<=$K$3,""今月はもう建てない（月間ストップ）""," & _
                             "IF(AND($C$4>=20,$O$4<=0),""直近20回がマイナス：プラスに戻るまで休む""," & _
                             "IF($K$4>=3,""3連敗中：株数を半分にする""," & _
                             "IF($M$4>=5,""枠が埋まっています"",""通常どおり。上位から買ってよい""))))"

    '=============== 5行目：見出し ===============
    Dim hd As Variant
    hd = Array("買日", "コード", "銘柄名", "買値", "株数", "投資額", "損切値", "利確値", _
               "期限日", "売日", "売値", "損益(円)", "損益率", "結果", "資金比", "メモ", "連敗", "通番")
    Dim j As Long
    For j = 0 To UBound(hd)
        ws.Cells(HD_ROW, j + 2).Value = hd(j)
    Next j

    '=============== 6行目以降：計算式 ===============
    Dim r As Long
    For r = TOP_ROW To MAX_ROW
        ws.Cells(r, 7).Formula = "=IF(OR(E" & r & "="""",F" & r & "=""""),"""",E" & r & "*F" & r & ")"
        ws.Cells(r, 8).Formula = "=IF(E" & r & "="""","""",ROUND(E" & r & "*" & (1 - SL_RATE) & ",0))"
        ws.Cells(r, 9).Formula = "=IF(E" & r & "="""","""",ROUND(E" & r & "*" & (1 + TP_RATE) & ",0))"
        ws.Cells(r, 10).Formula = "=IF(B" & r & "="""","""",WORKDAY(B" & r & "," & HOLD_DAYS & "))"
        ws.Cells(r, 13).Formula = "=IF(OR(L" & r & "="""",E" & r & "="""",F" & r & "=""""),"""",(L" & r & "-E" & r & ")*F" & r & ")"
        ws.Cells(r, 14).Formula = "=IF(OR(L" & r & "="""",E" & r & "=""""),"""",(L" & r & "-E" & r & ")/E" & r & ")"
        ws.Cells(r, 15).Formula = "=IF(E" & r & "="""","""",IF(L" & r & "="""",""保有中"",IF(N" & r & ">=0.07,""利確"",IF(N" & r & "<=-0.075,""損切"",IF(N" & r & ">0,""時間切れ(勝)"",""時間切れ(負)"")))))"
        ws.Cells(r, 16).Formula = "=IF(M" & r & "="""","""",M" & r & "/$C$3)"
        If r = TOP_ROW Then
            ws.Cells(r, 18).Formula = "=IF(L" & r & "="""","""",IF(M" & r & "<=0,1,0))"
        Else
            ws.Cells(r, 18).Formula = "=IF(L" & r & "="""","""",IF(M" & r & "<=0,N(R" & (r - 1) & ")+1,0))"
        End If
        ws.Cells(r, 19).Formula = "=IF(L" & r & "="""","""",COUNT($L$" & TOP_ROW & ":L" & r & "))"
    Next r

    '=============== 書式（初回・作り直しの時だけ）===============
    If doStyle Then
        With ws.Range("B3:S4")
            .Font.Name = "Meiryo UI": .Font.Size = 12
            .Interior.Color = RGB(226, 239, 218)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
            .VerticalAlignment = xlCenter
        End With
        ws.Range("Q4:S4").Font.Bold = True
        ws.Range("Q4").HorizontalAlignment = xlCenter
        ws.Rows(3).RowHeight = 26
        ws.Rows(4).RowHeight = 26

        With ws.Range("B" & HD_ROW & ":S" & HD_ROW)
            .Font.Name = "Meiryo UI": .Font.Size = HD_SIZE: .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 32, 96)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        ws.Rows(HD_ROW).RowHeight = 46

        With ws.Range("B" & TOP_ROW & ":S" & MAX_ROW)
            .Font.Name = "Meiryo UI"
            .Font.Size = DT_SIZE
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(TOP_ROW & ":" & MAX_ROW).RowHeight = 26

        '数値の書式
        ws.Range("C3").NumberFormat = "#,##0""円"""
        ws.Range("E3").NumberFormat = "0%"
        ws.Range("G3").NumberFormat = "#,##0""円"""
        ws.Range("I3").NumberFormat = "#,##0""円"""
        ws.Range("K3").NumberFormat = "#,##0""円"""
        ws.Range("E4").NumberFormat = "0.0%"
        ws.Range("G4").NumberFormat = "#,##0""円"""
        ws.Range("I4").NumberFormat = "#,##0""円"""
        ws.Range("O4").NumberFormat = "0.00%"
        ws.Range("B" & TOP_ROW & ":B" & MAX_ROW).NumberFormat = "yyyy/mm/dd"
        ws.Range("J" & TOP_ROW & ":K" & MAX_ROW).NumberFormat = "yyyy/mm/dd"
        ws.Range("E" & TOP_ROW & ":I" & MAX_ROW).NumberFormat = "#,##0"
        ws.Range("L" & TOP_ROW & ":M" & MAX_ROW).NumberFormat = "#,##0"
        ws.Range("N" & TOP_ROW & ":N" & MAX_ROW).NumberFormat = "0.0%"
        ws.Range("P" & TOP_ROW & ":P" & MAX_ROW).NumberFormat = "0.00%"
        ws.Range("O" & TOP_ROW & ":O" & MAX_ROW).HorizontalAlignment = xlCenter

        '入力するセルを黄色に
        HA_L_Input ws.Range("C3")
        HA_L_Input ws.Range("E3")
        HA_L_Input ws.Range("B" & TOP_ROW & ":F" & MAX_ROW)
        HA_L_Input ws.Range("K" & TOP_ROW & ":L" & MAX_ROW)
        HA_L_Input ws.Range("Q" & TOP_ROW & ":Q" & MAX_ROW)

        '色分け（保有中＝青、勝ち＝緑、負け＝赤）
        Dim rng As Range
        Set rng = ws.Range("B" & TOP_ROW & ":Q" & MAX_ROW)
        rng.FormatConditions.Delete
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$O" & TOP_ROW & "=""保有中""")
            .Interior.Color = RGB(221, 235, 247)
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M" & TOP_ROW & "<>"""",$M" & TOP_ROW & ">0)")
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M" & TOP_ROW & "<>"""",$M" & TOP_ROW & "<=0)")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With

        '裏方の列は薄く
        ws.Columns("R:S").Font.Color = RGB(210, 210, 210)

        '列幅
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 15
        ws.Columns("C").ColumnWidth = 9
        ws.Columns("D").ColumnWidth = 22
        ws.Columns("E").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 9
        ws.Columns("G").ColumnWidth = 14
        ws.Columns("H").ColumnWidth = 12
        ws.Columns("I").ColumnWidth = 12
        ws.Columns("J").ColumnWidth = 15
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("L").ColumnWidth = 12
        ws.Columns("M").ColumnWidth = 14
        ws.Columns("N").ColumnWidth = 11
        ws.Columns("O").ColumnWidth = 15
        ws.Columns("P").ColumnWidth = 11
        ws.Columns("Q").ColumnWidth = 26
        ws.Columns("R:S").ColumnWidth = 6

        '=============== 使い方 ===============
        Dim g As Long
        g = MAX_ROW + 2
        ws.Cells(g, 2).Value = "【使い方】"
        ws.Cells(g + 1, 2).Value = "1. 買ったら　買日・コード・銘柄名・買値・株数　の5つだけ入れる。損切値と利確値が自動で出ます。"
        ws.Cells(g + 2, 2).Value = "2. その値段で　逆指値（損切値）と　売り指値（利確値）を証券会社に出す。"
        ws.Cells(g + 3, 2).Value = "3. 売れたら　売日・売値　を入れる。損益と勝率が自動で更新されます。"
        ws.Cells(g + 4, 2).Value = "4. 「期限日」（買った日から5営業日後）の引けで、勝ち負けに関係なく成行手仕舞い。"
        ws.Cells(g + 6, 2).Value = "【怖くなったら、この数字を見てください（300銘柄×250日の検証）】"
        ws.Cells(g + 7, 2).Value = "・負けの平均は －5.19%。－8%満額で切られるのは17回に1回だけです。"
        ws.Cells(g + 8, 2).Value = "・最大の連敗は 4連敗。5連敗以上は1年で一度もありませんでした。"
        ws.Cells(g + 9, 2).Value = "・10回中4回は負けます。それが設計です。負けは家賃だと思ってください。"
        ws.Cells(g + 10, 2).Value = "・「直近20回の平均」がマイナスになったら、プラスに戻るまで休む（自動で指示が出ます）。"
        ws.Cells(g + 11, 2).Value = "・1銘柄は30万円以上にする。15万円では候補の2割しか買えません（候補の株価は平均4,100円）。"
        ws.Cells(g + 12, 2).Value = "・小さくやりたい時は割合を下げず、運用資金そのものを減らす（例：150万円で1銘柄20%=30万円）。"
        With ws.Range("B" & g & ":B" & (g + 12))
            .Font.Name = "Meiryo UI": .Font.Size = 12
        End With
        ws.Cells(g, 2).Font.Bold = True
        ws.Cells(g + 6, 2).Font.Bold = True
    End If

    '--- 退避した記録を書き戻す ---
    If Not isNew Then
        ws.Range(ws.Cells(TOP_ROW, 2), ws.Cells(MAX_ROW, 6)).Value = k1
        ws.Range(ws.Cells(TOP_ROW, 11), ws.Cells(MAX_ROW, 12)).Value = k2
        ws.Range(ws.Cells(TOP_ROW, 17), ws.Cells(MAX_ROW, 17)).Value = k3
    End If

    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B" & TOP_ROW).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    Application.ScreenUpdating = True

    Dim msg As String
    If doStyle Then
        msg = "「" & SH & "」シートを作りました。" & vbCrLf & vbCrLf & _
              "・入力するのは黄色いセルだけです。" & vbCrLf & _
              "・C3 の資金と、E3 の1銘柄の割合（最初は10%）を確認してください。" & vbCrLf & _
              "・2～5行目の色やフォントは自由に変えて構いません。" & vbCrLf & _
              "　次回からは書式を触らず、計算式だけ直します。"
    Else
        msg = "計算式を最新にしました。書式（色・フォント）はそのままです。"
    End If
    MsgBox msg, vbInformation
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
