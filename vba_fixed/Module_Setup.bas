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
    Dim warn As String
    Dim needFormula As Boolean

    If BlockedWhileLogging("① 準備") Then Exit Sub

    Application.ScreenUpdating = False

    Set shts = TargetSheets()

    For i = 1 To shts.Count
        Set ws = shts(i)
        SetupStockSheet ws
        If Len(TextOrEmpty(ws.Range(TICK_FORMULA_CELL).Formula)) = 0 Then needFormula = True
    Next i

    SetupResultSheet
    ResultSheet().Activate

    Application.ScreenUpdating = True

    '--- 設定の点検 ---------------------------------------------------
    warn = LayoutWarning()
    If Len(warn) > 0 Then
        MsgBox "列レイアウトの設定に問題があります。" & vbCrLf & vbCrLf & warn, vbCritical, "設定"
    End If

    warn = WindowWarning()
    If Len(warn) > 0 Then
        MsgBox "判定窓の設定に問題があります。" & vbCrLf & vbCrLf & warn, vbExclamation, "判定窓の設定"
    End If

    warn = DuplicateCodeWarning()
    If Len(warn) > 0 Then
        MsgBox warn, vbExclamation, "銘柄コードの重複"
    End If

    If shts.Count = 0 Then
        MsgBox "銘柄シートが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "各銘柄シートの B3 に証券コードを入れてから、" & vbCrLf & _
               "もう一度「① 準備」を押してください。", vbExclamation, "準備"
        Exit Sub
    End If

    MsgBox shts.Count & " 銘柄シートの準備が完了しました。" & vbCrLf & vbCrLf & _
           "判定窓 : " & JUDGE_START & " ～ " & JUDGE_END & vbCrLf & _
           "1区間 : " & (BucketSec() \ 60) & "分" & (BucketSec() Mod 60) & "秒", _
           vbInformation, "準備"

    '--- 歩み値の数式が無いまま本番に入らせない ------------------------
    If needFormula Then
        If MsgBox("歩み値（RssTickList）の数式が入っていないシートがあります。" & vbCrLf & vbCrLf & _
                  "このままだと出来高ポーリング（③約定速度が不正確）で動きます。" & vbCrLf & _
                  "いま歩み値の数式を設定しますか？", vbYesNo + vbQuestion, "歩み値の数式") = vbYes Then
            ApplyTickFormula
        End If
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
        .Cells(ROW_HEADER, COL_VOL).Value = "出来高"
        .Cells(ROW_HEADER, COL_PRICE).Value = "約定値"
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

        .Range("T2:T7").NumberFormat = "General"
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

        '--- 歩み値（TICK）ブロックの説明 ----------------------------
        .Range("AA2").Value = "歩み値（TICK）ブロック"
        .Range("AA2").Font.Bold = True
        .Range("AA3").Value = "「歩み値の数式を全シートに設定」ボタンで自動設定できます"
        .Range("AA4").Value = "手で登録する場合の RssTickList 設定"
        .Range("AA5").Value = "　表示開始セル = " & TICK_FORMULA_CELL
        .Range("AA6").Value = "　銘柄コード   = B3（セル参照）"
        .Range("AA7").Value = "　表示本数     = " & TICK_SHOW_ROWS
        .Range("AA8").Value = "　取得項目     = " & ItemListText()
        .Range("AA9").Value = "※見出しとデータは RSS が書き出します"
        .Range("AA3:AA9").Font.Color = RGB(120, 120, 120)
        .Range("AA1").ClearContents

        '--- 旧レイアウトの見出しは RSS の見出しと二重になるので撤去 ---
        .Range("AB2:AE2").ClearContents
        .Range("AB2:AE2").Interior.Pattern = xlNone
        .Range("AB3:AE3").ClearContents

        .Columns("AA").ColumnWidth = 34
        .Columns("AB:AE").ColumnWidth = 12

        '--- 書式・幅 ------------------------------------------------
        '   時刻列はここで一度だけ設定します。1ティックごとに書式を
        '   設定すると、そのたびに再計算が走って記録が追いつきません。
        .Columns(COL_TIME).NumberFormatLocal = "hh:mm:ss"
        .Range(.Columns(COL_VOL), .Columns(COL_MARK)).NumberFormat = "General"
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
    End With

    FreezeHeader ws
End Sub

'------------------------------------------------------------------
' Judge_Results シートの体裁とボタン
'------------------------------------------------------------------
Public Sub SetupResultSheet()

    Dim ws As Worksheet
    Dim head As Variant
    Dim logs As Variant
    Dim i As Long

    Set ws = ResultSheet()

    head = Array("コード", "銘柄名", "判定", "信頼度", _
                 "UpSeqMax", "DnSeqMax", "方向判定①", _
                 "Vol1", "Vol2", "Vol3", "Vol4", "Vol4優勢②", _
                 "SpeedMax③", "ティック数", "BuyTotal", "SellTotal", "更新時刻", _
                 "最良買気配", "最良売気配", "スプレッド")

    logs = Array("開始時刻", "取込モード", "配信状態", "最終ティック", "警告", "終了時刻")

    With ws
        '--- 1行目はメモ欄として空けておく（マクロは一切書き込まない） ---

        For i = LBound(head) To UBound(head)
            .Cells(ROW_HEADER, SUM_COL_FIRST + i).Value = head(i)
        Next i

        With .Range(.Cells(ROW_HEADER, SUM_COL_FIRST), .Cells(ROW_HEADER, SUM_COL_LAST))
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
        .Columns("Y").ColumnWidth = 14
        .Columns("Z").ColumnWidth = 60
        .Columns("AA").ColumnWidth = 20

        With .Range(.Cells(ROW_HEADER, SUM_COL_LAST - 2), .Cells(ROW_HEADER, SUM_COL_LAST))
            .Interior.Color = RGB(255, 242, 204)
        End With

        '--- 状態表示 ------------------------------------------------
        .Range("W2").Value = "状態"
        .Range("W3").Value = "件数"
        .Range("W4").Value = "内訳"
        .Range("W2:W4").Font.Bold = True
        .Range("W1:X1").ClearContents

        '--- 実行ログ（状態セルは判定完了で上書きされるため別枠） -----
        For i = LBound(logs) To UBound(logs)
            .Range(LOG_TOP).Offset(i, 0).Value = logs(i)
        Next i
        .Range(LOG_TOP).Offset(0, 0).Resize(LOGR_ROWS, 1).Font.Bold = True
        .Range(LOG_TOP).Offset(0, 0).Resize(LOGR_ROWS, 1).Interior.Color = RGB(242, 242, 242)
        .Range(LOG_TOP).Offset(0, 0).Resize(LOGR_ROWS, 2).Borders.LineStyle = xlContinuous
        .Range(LOG_TOP).Offset(0, 0).Resize(LOGR_ROWS, 2).Borders.Color = RGB(200, 200, 200)

        '--- 内部管理（OnTime の予約時刻。消さないでください） ---------
        .Range(SCHED_NOTE).Value = "※内部管理（消さないでください）"
        .Range(SCHED_NOTE).Font.Color = RGB(150, 150, 150)
        .Range(SCHED_NOTE).Font.Italic = True
        .Range(SCHED_ARM).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        .Range(SCHED_POLL).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        .Range(SCHED_ARM & ":" & SCHED_POLL).Font.Color = RGB(180, 180, 180)
    End With

    SeedResultRows ws
    FreezeHeader ws
    BuildButtons ws
End Sub

'------------------------------------------------------------------
' 見出し固定（1～2行目を常に表示）
'   先に左上までスクロールを戻さないと、1行目が固定ペインの外に
'   押し出されて二度と表示されなくなります。
'------------------------------------------------------------------
Private Sub FreezeHeader(ws As Worksheet)

    Dim prevUpdating As Boolean
    Dim prev As Worksheet

    On Error Resume Next
    prevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set prev = ActiveSheet

    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True

    If Not prev Is Nothing Then prev.Activate
    Application.ScreenUpdating = prevUpdating
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' 操作ボタンを作り直す
'------------------------------------------------------------------
Private Sub BuildButtons(ws As Worksheet)

    ws.Buttons.Delete

    AddButton ws, 0, "① 準備（表とRSSを作る）", "Setup_All"
    AddButton ws, 1, "② ティック記録 開始", "StartTickLoggingButton"
    AddButton ws, 2, "③ 記録 停止 → 判定", "StopAndJudge"
    AddButton ws, 3, "④ 判定 実行", "JudgeAll"
    AddButton ws, 4, "⑤ ティックログ消去", "ClearAllTicks"
    AddButton ws, 5, "自動開始を予約（" & AUTO_ARM_TIME & "）", "ArmAutoRun"
    AddButton ws, 6, "歩み値の数式を全シートに設定", "ApplyTickFormula"
    AddButton ws, 7, "歩み値ブロックの確認", "ShowTickBlockInfo"
    AddButton ws, 8, "歩み値を今すぐ取込", "ImportTickBlockNow"
    AddButton ws, 9, "出来高プロファイル", "ShowVolumeProfile"
    AddButton ws, 10, "CSV（歩み値）取込", "ImportTickCsv"
    AddButton ws, 11, "銘柄シートを追加", "AddTickSheet"
    AddButton ws, 12, "★ RSS 疎通テスト（切り分け）", "RunFeedDiagnosis"
    AddButton ws, 13, "設定の点検", "ShowConfigCheck"
End Sub

'------------------------------------------------------------------
' 銘柄行をあらかじめ置いておく
'   記録中から最良気配をここに出すため
'------------------------------------------------------------------
Private Sub SeedResultRows(ws As Worksheet)

    Dim shts As Collection
    Dim src As Worksheet
    Dim i As Long, rw As Long

    Set shts = TargetSheets()

    For i = 1 To shts.Count
        Set src = shts(i)
        rw = ROW_HEADER + i
        ws.Cells(rw, SUM_COL_FIRST).Value = src.Cells(ROW_CODE, COL_CODE).Value
        ws.Cells(rw, SUM_COL_FIRST + 1).Value = src.Cells(ROW_CODE, COL_NAME).Value
    Next i
End Sub

Private Sub AddButton(ws As Worksheet, ByVal idx As Long, ByVal caption As String, ByVal macroName As String)

    Dim b As Object
    Dim anchor As Range

    Set anchor = ws.Range(BTN_ANCHOR)

    Set b = ws.Buttons.Add(anchor.Left, anchor.top + idx * (BTN_H + 6), BTN_W, BTN_H)
    b.caption = caption
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

    If BlockedWhileLogging("銘柄シートを追加") Then Exit Sub

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

    If MsgBox("シート " & nm & " を作成しました。" & vbCrLf & vbCrLf & _
              "歩み値（RssTickList）の数式もいま設定しますか？", _
              vbYesNo + vbQuestion, "銘柄シートの追加") = vbYes Then
        ApplyTickFormula
    End If
End Sub

'------------------------------------------------------------------
' RssTickList の取得項目（TB_OFF_* の並びどおり）
'------------------------------------------------------------------
Public Function ItemCount() As Long
    ItemCount = IIf(TB_OFF_MARK >= 0, 4, 3)
End Function

Public Function ItemName(ByVal offset As Long) As String

    Select Case offset
        Case TB_OFF_TIME:  ItemName = "時刻"
        Case TB_OFF_VOL:   ItemName = "出来高"
        Case TB_OFF_PRICE: ItemName = "約定値"
        Case TB_OFF_MARK:  ItemName = TICK_MARK_ITEM
    End Select
End Function

Public Function ItemListText() As String

    Dim i As Long
    Dim s As String

    For i = 0 To ItemCount() - 1
        s = s & IIf(Len(s) > 0, " / ", "") & ItemName(i)
    Next i
    ItemListText = s
End Function

'------------------------------------------------------------------
' 歩み値（RssTickList）の数式を全銘柄シートに設定する
'------------------------------------------------------------------
Public Sub ApplyTickFormula()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long, k As Long
    Dim itemsAddr As String
    Dim n As Long

    If BlockedWhileLogging("歩み値の数式を全シートに設定") Then Exit Sub

    Set shts = TargetSheets()
    If shts.Count = 0 Then
        MsgBox "銘柄シートがありません。B3 に証券コードを入れてください。", vbExclamation
        Exit Sub
    End If

    n = ItemCount()

    If MsgBox(shts.Count & " 枚の銘柄シートに歩み値の数式を設定します。" & vbCrLf & vbCrLf & _
              "　" & TICK_ITEMS_CELL & " から右へ : " & ItemListText() & vbCrLf & _
              "　" & TICK_FORMULA_CELL & " : =RssTickList(...)  ※銘柄コードは各シートの B3" & vbCrLf & _
              "　表示本数 : " & TICK_SHOW_ROWS & vbCrLf & vbCrLf & _
              "この位置の既存の内容は上書きされます。よろしいですか？", _
              vbYesNo + vbQuestion, "歩み値の数式を設定") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    For i = 1 To shts.Count
        Set ws = shts(i)

        With ws.Range(TICK_ITEMS_CELL)
            For k = 0 To n - 1
                .Offset(0, k).Value = ItemName(k)
            Next k
            itemsAddr = .Resize(1, n).Address(True, True)
        End With

        ws.Range(TICK_FORMULA_CELL).Formula = _
            "=RssTickList(" & itemsAddr & ",$B$3," & TICK_SHOW_ROWS & ")"
    Next i
    Application.ScreenUpdating = True

    MsgBox shts.Count & " 枚に設定しました。" & vbCrLf & vbCrLf & _
           "RSS が配信を始めるまで数秒待ってから、" & vbCrLf & _
           "「歩み値ブロックの確認」で読めているか確かめてください。", _
           vbInformation, "歩み値の数式を設定"
End Sub

'------------------------------------------------------------------
' 旧バージョンが1行目に書いた残骸を掃除する（1回だけ実行）
'   Alt+F8 から実行してください。ボタンは置いていません。
'------------------------------------------------------------------
Public Sub CleanupOldLayout()

    Dim shts As Collection
    Dim ws As Worksheet
    Dim i As Long

    If BlockedWhileLogging("旧レイアウトの掃除") Then Exit Sub

    If MsgBox("旧バージョンが1行目に書いた文字と書式を消します。" & vbCrLf & _
              "（タイトル、O1の「判定」、S1:T1 のRSS数式など）" & vbCrLf & vbCrLf & _
              "1行目に既にメモを書いている場合は一緒に消えます。" & vbCrLf & _
              "実行しますか？", vbYesNo + vbExclamation, "旧レイアウトの掃除") <> vbYes Then Exit Sub

    Set shts = TargetSheets()

    Application.ScreenUpdating = False
    For i = 1 To shts.Count
        Set ws = shts(i)
        ws.Rows(1).ClearContents
        ws.Rows(1).ClearFormats
        ws.Rows(1).RowHeight = ws.StandardHeight
    Next i

    With ResultSheet()
        .Rows(1).ClearContents
        .Rows(1).ClearFormats
        .Rows(1).RowHeight = .StandardHeight
    End With
    Application.ScreenUpdating = True

    MsgBox "1行目を掃除しました。ここが自由に使えるメモ欄です。", _
           vbInformation, "旧レイアウトの掃除"
End Sub
