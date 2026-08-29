Attribute VB_Name = "Module_Setup"
Option Explicit

'==================================================================
' Module_Setup : シートの体裁・RSS数式・ボタンを自動生成する
'   最初に1回「Setup_All」を実行すれば表が完成します。
'   （銘柄を足したあと、もう一度実行しても安全です）
'==================================================================

'------------------------------------------------------------------
' すべて準備する（ボタン①）
'------------------------------------------------------------------
Public Sub Setup_All()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long

    Application.ScreenUpdating = False

    Set shts = TargetSheets()

    For i = 1 To shts.Count
        Set ws = shts(i)
        SetupStockSheet ws
    Next i

    SetupResultSheet
    ResultSheet().Activate

    Application.ScreenUpdating = True

    If Len(WindowWarning()) > 0 Then
        MsgBox "判定窓の設定に問題があります。" & vbCrLf & vbCrLf & WindowWarning(), _
               vbExclamation, "判定窓の設定"
    End If

    If shts.Count = 0 Then
        MsgBox "銘柄シートが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "各銘柄シートの B3 に証券コードを入れてから、" & vbCrLf & _
               "もう一度「① 準備」を押してください。", vbExclamation, "準備"
    Else
        MsgBox shts.Count & " 銘柄シートの準備が完了しました。" & vbCrLf & vbCrLf & _
               "判定窓 : " & JUDGE_START & " ～ " & JUDGE_END & vbCrLf & _
               "1区間 : " & (BucketSec() \ 60) & "分" & (BucketSec() Mod 60) & "秒", _
               vbInformation, "準備"
    End If
End Sub

'------------------------------------------------------------------
' 銘柄シート1枚の体裁を整える
'------------------------------------------------------------------
Public Sub SetupStockSheet(ws As Worksheet)

    Dim labels As Variant
    Dim i As Long
    Dim st As Long, bk As Long

    With ws
        '--- 1行目はメモ欄として空けておく（マクロは一切書き込まない） ---

        '--- 見出し行 ------------------------------------------------
        .Cells(ROW_HEADER, COL_CODE).Value = "コード"
        .Cells(ROW_HEADER, COL_NAME).Value = "銘柄名称"
        .Cells(ROW_HEADER, COL_TIME).Value = "時刻"
        .Cells(ROW_HEADER, COL_PRICE).Value = "約定値"
        .Cells(ROW_HEADER, COL_VOL).Value = "出来高"
        .Cells(ROW_HEADER, COL_W).Value = "重み"
        .Cells(ROW_HEADER, COL_UP).Value = "UpScore"
        .Cells(ROW_HEADER, COL_DN).Value = "DnScore"
        .Cells(ROW_HEADER, COL_BID).Value = "最良買気配値"
        .Cells(ROW_HEADER, COL_ASK).Value = "最良売気配値"
        .Cells(ROW_HEADER, COL_DIR).Value = "方向"
        .Cells(ROW_HEADER, COL_SPD).Value = "秒内約定数"
        .Cells(ROW_HEADER, COL_MARK).Value = "ティック"

        With .Range(.Cells(ROW_HEADER, COL_CODE), .Cells(ROW_HEADER, COL_MARK))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With

        '--- 銘柄名称（RSS） ----------------------------------------
        .Cells(ROW_CODE, COL_NAME).Formula = "=RssMarket($B$3,""銘柄名称"")"

        '--- RSS ライブ取得ブロック（S列ラベル / T列値） -------------
        .Range("S2").Value = "現在値"
        .Range("S3").Value = "出来高（当日累計）"
        .Range("S4").Value = "最良買気配値"
        .Range("S5").Value = "最良売気配値"
        .Range("S6").Value = "現在値時刻"
        .Range("S7").Value = "スプレッド"
        .Range("S8").Value = "※このブロックを RSS が更新します"
        .Range(LIVE_SPREAD).Formula = "=IF(COUNT(T4:T5)=2,T5-T4,"""")"

        .Range(LIVE_PRICE).Formula = "=RssMarket($B$3,""現在値"")"
        .Range(LIVE_VOL).Formula = "=RssMarket($B$3,""出来高"")"
        .Range(LIVE_BID).Formula = "=RssMarket($B$3,""最良買気配値"")"
        .Range(LIVE_ASK).Formula = "=RssMarket($B$3,""最良売気配値"")"
        .Range(LIVE_TIME).Formula = "=RssMarket($B$3,""現在値時刻"")"
        .Range(LIVE_TIME).NumberFormatLocal = "hh:mm:ss"

        .Range("S2:S8").Font.Color = RGB(120, 120, 120)
        .Range("S8").Font.Italic = True
        .Range("S4:T5").Font.Bold = True
        .Range("S1:T1").ClearContents

        '--- 判定結果ブロック（O列ラベル / P列値） ------------------
        labels = Array("判定", "信頼度", _
                       "最良買気配値", "最良売気配値", "スプレッド", _
                       "方向判定①", "UpSeqMax", "DnSeqMax", _
                       "Vol1", "Vol2", "Vol3", "Vol4", _
                       "Vol4優勢②", "SpeedMax③", "ティック数", _
                       "BuyTotal(参考)", "SellTotal(参考)", "判定時刻")

        '--- Vol1～Vol4 のラベルは判定窓から自動生成 ------------------
        st = SecOfText(JUDGE_START)
        bk = BucketSec()
        labels(8) = "Vol1 (" & SecHM(st) & "-" & SecHM(st + bk) & ")"
        labels(9) = "Vol2 (" & SecHM(st + bk) & "-" & SecHM(st + bk * 2) & ")"
        labels(10) = "Vol3 (" & SecHM(st + bk * 2) & "-" & SecHM(st + bk * 3) & ")"
        labels(11) = "Vol4 (" & SecHM(st + bk * 3) & "-" & SecHM(SecOfText(JUDGE_END)) & ")"

        For i = LBound(labels) To UBound(labels)
            .Range(RES_TOP).Offset(i, 0).Value = labels(i)
        Next i

        With .Range(RES_TOP).Offset(0, 0).Resize(UBound(labels) + 1, 2)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
        With .Range(RES_TOP).Resize(1, 2)
            .Font.Bold = True
            .Font.Size = 13
        End With
        .Range(RES_TOP).Offset(0, 0).Resize(UBound(labels) + 1, 1).Interior.Color = RGB(242, 242, 242)

        '--- 歩み値（TICK）ブロック ----------------------------------
        '   1行目には書き込まないので、AA2 から並べます
        .Range("AA2").Value = "歩み値（TICK）ブロック"
        .Range("AA2").Font.Bold = True
        .Range("AA3").Value = "RssTickList をここに登録します"
        .Range("AA4").Value = "　表示開始セル = " & TICK_BLOCK_CELL
        .Range("AA5").Value = "　銘柄コード   = B3（セル参照）"
        .Range("AA6").Value = "　表示本数     = 300"
        .Range("AA7").Value = "　取得項目     = 時刻 / 出来高 / 約定値"
        .Range("AA8").Value = "※見出しとデータは RSS が書き出します"
        .Range("AA3:AA8").Font.Color = RGB(120, 120, 120)
        .Range("AA1").ClearContents

        '--- 旧レイアウトの見出しは RSS の見出しと二重になるので撤去 ---
        .Range("AB2:AE2").ClearContents
        .Range("AB2:AE2").Interior.Pattern = xlNone

        .Columns("AA").ColumnWidth = 34
        .Columns("AB:AE").ColumnWidth = 12

        '--- 書式・幅 ------------------------------------------------
        .Columns(COL_TIME).NumberFormatLocal = "hh:mm:ss"
        .Columns(COL_TIME).ColumnWidth = 10
        .Columns(COL_PRICE).ColumnWidth = 10
        .Columns(COL_VOL).ColumnWidth = 10
        .Columns(COL_W).ColumnWidth = 6
        .Columns(COL_UP).ColumnWidth = 10
        .Columns(COL_DN).ColumnWidth = 10
        .Columns(COL_BID).ColumnWidth = 13
        .Columns(COL_ASK).ColumnWidth = 13
        .Columns(COL_DIR).ColumnWidth = 6
        .Columns(COL_SPD).ColumnWidth = 10
        .Columns(COL_MARK).ColumnWidth = 8
        .Columns("O").ColumnWidth = 18
        .Columns("P").ColumnWidth = 20
        .Columns("S").ColumnWidth = 20
        .Columns("T").ColumnWidth = 12

        '--- 見出し固定（1～2行目を常に表示） -------------------------
        '   先に左上までスクロールを戻さないと、1行目が固定ペインの外に
        '   押し出されて二度と表示されなくなります。
        On Error Resume Next
        .Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With
End Sub

'------------------------------------------------------------------
' Judge_Results シートの体裁とボタン
'------------------------------------------------------------------
Public Sub SetupResultSheet()

    Dim ws As Worksheet
    Dim head As Variant
    Dim i As Long

    Set ws = ResultSheet()

    head = Array("コード", "銘柄名", "判定", "信頼度", _
                 "UpSeqMax", "DnSeqMax", "方向判定①", _
                 "Vol1", "Vol2", "Vol3", "Vol4", "Vol4優勢②", _
                 "SpeedMax③", "ティック数", "BuyTotal", "SellTotal", "更新時刻", _
                 "最良買気配", "最良売気配", "スプレッド")

    With ws
        '--- 1行目はメモ欄として空けておく（マクロは一切書き込まない） ---

        For i = LBound(head) To UBound(head)
            .Cells(ROW_HEADER, 2 + i).Value = head(i)
        Next i

        With .Range(.Cells(ROW_HEADER, 2), .Cells(ROW_HEADER, 21))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With

        .Columns("B:C").ColumnWidth = 14
        .Columns("D:E").ColumnWidth = 16
        .Columns("F:R").ColumnWidth = 11
        .Columns("S:U").ColumnWidth = 12
        .Columns("V").ColumnWidth = 3
        .Columns("W").ColumnWidth = 6
        .Columns("X").ColumnWidth = 46

        With .Range(.Cells(ROW_HEADER, 19), .Cells(ROW_HEADER, 21))
            .Interior.Color = RGB(255, 242, 204)
        End With

        '--- 状態表示 ------------------------------------------------
        .Range("W2").Value = "状態"
        .Range("W3").Value = "件数"
        .Range("W4").Value = "内訳"
        .Range("W2:W4").Font.Bold = True
        .Range("W1:X1").ClearContents

        '   ※ 旧版は O1:O3 / T1:U3 に文字を書いていましたが、現在その位置は
        '      サマリー表（O=ティック数 / T=最良売気配 / U=スプレッド）です。
        '      掃除しようとすると見出しとデータを消してしまうため行いません。
        '      1行目に古い文字が残っていたら手で消してください。

        '--- 見出し固定（1～2行目を常に表示） -------------------------
        On Error Resume Next
        .Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With

    SeedResultRows ws

    BuildButtons ws
End Sub

'------------------------------------------------------------------
' 操作ボタンを作り直す
'------------------------------------------------------------------
Private Sub BuildButtons(ws As Worksheet)

    ws.Buttons.Delete

    AddButton ws, 0, "① 準備（表とRSSを作る）", "Setup_All"
    AddButton ws, 1, "② ティック記録 開始", "StartTickLogging"
    AddButton ws, 2, "③ 記録 停止 → 判定", "StopAndJudge"
    AddButton ws, 3, "④ 判定 実行", "JudgeAll"
    AddButton ws, 4, "⑤ ティックログ消去", "ClearAllTicks"
    AddButton ws, 5, "自動開始を予約（14:59:30）", "ArmAutoRun"
    AddButton ws, 6, "歩み値を今すぐ取込", "ImportTickBlockNow"
    AddButton ws, 7, "歩み値ブロックの確認", "ShowTickBlockInfo"
    AddButton ws, 8, "出来高プロファイル", "ShowVolumeProfile"
    AddButton ws, 9, "CSV（歩み値）取込", "ImportTickCsv"
    AddButton ws, 10, "銘柄シートを追加", "AddTickSheet"
End Sub

'------------------------------------------------------------------
' 銘柄行をあらかじめ置いておく
'   記録中（15:00～15:20）から最良気配をここに出すため
'------------------------------------------------------------------
Private Sub SeedResultRows(ws As Worksheet)

    Dim shts As Collection
    Dim src As Worksheet
    Dim i As Long, rw As Long

    Set shts = TargetSheets()

    For i = 1 To shts.Count
        Set src = shts(i)
        rw = ROW_HEADER + i
        ws.Cells(rw, 2).Value = src.Cells(ROW_CODE, COL_CODE).Value
        ws.Cells(rw, 3).Value = src.Cells(ROW_CODE, COL_NAME).Value
    Next i
End Sub

Private Sub AddButton(ws As Worksheet, ByVal idx As Long, ByVal caption As String, ByVal macroName As String)

    Dim b As Object
    Dim anchor As Range

    Set anchor = ws.Range(BTN_ANCHOR)

    Set b = ws.Buttons.Add(anchor.Left, anchor.Top + idx * (BTN_H + 6), BTN_W, BTN_H)
    b.Caption = caption
    b.OnAction = macroName
    b.Name = "btn" & idx
End Sub

'------------------------------------------------------------------
' 銘柄シートを1枚追加する
'------------------------------------------------------------------
Public Sub AddTickSheet()

    Dim s As String
    Dim ws As Worksheet
    Dim nm As String

    s = InputBox("追加する銘柄の証券コードを入力してください。", "銘柄シートの追加")
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Not IsNumeric(s) Then
        MsgBox "証券コードは数値で入力してください。", vbExclamation
        Exit Sub
    End If

    nm = "T" & Trim$(s)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        MsgBox "シート " & nm & " は既に存在します。", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ResultSheet())
    ws.Name = nm
    ws.Cells(ROW_CODE, COL_CODE).Value = CLng(s)

    SetupStockSheet ws
    SetupResultSheet

    MsgBox "シート " & nm & " を作成しました。", vbInformation
End Sub
