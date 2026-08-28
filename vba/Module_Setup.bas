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

    Application.ScreenUpdating = True

    If shts.Count = 0 Then
        MsgBox "銘柄シートが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "各銘柄シートの B3 に証券コードを入れてから、" & vbCrLf & _
               "もう一度「① 準備」を押してください。", vbExclamation, "準備"
    Else
        MsgBox shts.Count & " 銘柄シートの準備が完了しました。", vbInformation, "準備"
    End If
End Sub

'------------------------------------------------------------------
' 銘柄シート1枚の体裁を整える
'------------------------------------------------------------------
Public Sub SetupStockSheet(ws As Worksheet)

    Dim labels As Variant
    Dim i As Long

    With ws
        '--- タイトル ------------------------------------------------
        .Range("A1").Value = "ティック判定ロジック（15:00〜15:20）"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

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

        With .Range(.Cells(ROW_HEADER, COL_CODE), .Cells(ROW_HEADER, COL_SPD))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With

        '--- 銘柄名称（RSS） ----------------------------------------
        .Cells(ROW_CODE, COL_NAME).Formula = "=RssMarket($B$3,""銘柄名称"")"

        '--- RSS ライブ取得ブロック（S列ラベル / T列値） -------------
        .Range("S1").Value = "現在値"
        .Range("S2").Value = "出来高（当日累計）"
        .Range("S3").Value = "最良買気配値"
        .Range("S4").Value = "最良売気配値"
        .Range("S5").Value = "現在値時刻"
        .Range("S6").Value = "※このブロックを RSS が更新します"

        .Range(LIVE_PRICE).Formula = "=RssMarket($B$3,""現在値"")"
        .Range(LIVE_VOL).Formula = "=RssMarket($B$3,""出来高"")"
        .Range(LIVE_BID).Formula = "=RssMarket($B$3,""最良買気配値"")"
        .Range(LIVE_ASK).Formula = "=RssMarket($B$3,""最良売気配値"")"
        .Range(LIVE_TIME).Formula = "=RssMarket($B$3,""現在値時刻"")"
        .Range(LIVE_TIME).NumberFormatLocal = "hh:mm:ss"

        .Range("S1:S6").Font.Color = RGB(120, 120, 120)
        .Range("S6").Font.Italic = True

        '--- 判定結果ブロック（O列ラベル / P列値） ------------------
        labels = Array("判定", "信頼度", "方向判定①", _
                       "UpSeqMax", "DnSeqMax", _
                       "Vol1 (15:00-05)", "Vol2 (15:05-10)", "Vol3 (15:10-15)", "Vol4 (15:15-20)", _
                       "Vol4優勢②", "SpeedMax③", "ティック数", _
                       "BuyTotal(参考)", "SellTotal(参考)", "判定時刻")

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
        .Columns("O").ColumnWidth = 18
        .Columns("P").ColumnWidth = 20
        .Columns("S").ColumnWidth = 20
        .Columns("T").ColumnWidth = 12

        '--- 見出し固定 ---------------------------------------------
        On Error Resume Next
        .Activate
        ActiveWindow.FreezePanes = False
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
                 "SpeedMax③", "ティック数", "BuyTotal", "SellTotal", "更新時刻")

    With ws
        .Range("B1").Value = "引け判定サマリー（15:00〜15:20 のティックで判定）"
        .Range("B1").Font.Bold = True
        .Range("B1").Font.Size = 14

        For i = LBound(head) To UBound(head)
            .Cells(ROW_HEADER, 2 + i).Value = head(i)
        Next i

        With .Range(.Cells(ROW_HEADER, 2), .Cells(ROW_HEADER, 18))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With

        .Columns("B:C").ColumnWidth = 14
        .Columns("D:E").ColumnWidth = 16
        .Columns("F:R").ColumnWidth = 11
        .Columns("T").ColumnWidth = 4
        .Columns("U").ColumnWidth = 40

        '--- 状態表示 ------------------------------------------------
        .Range("T1").Value = "状態"
        .Range("T2").Value = "件数"
        .Range("T3").Value = "内訳"
        .Range("T1:T3").Font.Bold = True

        '--- 旧レイアウトの残骸を掃除 -------------------------------
        .Range("O1:O3").ClearContents
    End With

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
    AddButton ws, 6, "CSV（歩み値）取込", "ImportTickCsv"
    AddButton ws, 7, "銘柄シートを追加", "AddTickSheet"
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
