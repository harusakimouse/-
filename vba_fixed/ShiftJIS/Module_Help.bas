Attribute VB_Name = "Module_Help"
Option Explicit

'==================================================================
' Module_Help : 「手順書」「表の見方」シートを生成する
'
'   どちらもマクロが書き出す説明シートです。手で編集しても構いませんが、
'   「手順書・表の見方を作り直す」ボタンを押すと作り直されます。
'   （B3 に証券コードが入らないので、判定対象の銘柄シートにはなりません）
'==================================================================

'--- 書き込み中の状態 ---
Private mHws As Worksheet
Private mHrow As Long

Private Const COL_L As Long = 2      ' B 見出し・キー
Private Const COL_R As Long = 3      ' C 説明

'------------------------------------------------------------------
' 無ければ作る（Setup_All から呼ばれる。既にあれば何もしない）
'------------------------------------------------------------------
Public Sub EnsureHelpSheets()

    Dim missing As Boolean

    On Error Resume Next
    missing = (SheetByName(HELP_STEPS_SHEET) Is Nothing) Or (SheetByName(HELP_TABLE_SHEET) Is Nothing)
    On Error GoTo 0

    If missing Then BuildHelpSheets
End Sub

'------------------------------------------------------------------
' 作り直す（ボタン）
'------------------------------------------------------------------
Public Sub RebuildHelpSheets()

    If BlockedWhileLogging("手順書・表の見方を作り直す") Then Exit Sub

    If MsgBox("「" & HELP_STEPS_SHEET & "」と「" & HELP_TABLE_SHEET & "」を作り直します。" & vbCrLf & vbCrLf & _
              "この2枚に書き込んだメモは消えます。よろしいですか？", _
              vbYesNo + vbQuestion, "説明シートの作成") <> vbYes Then Exit Sub

    BuildHelpSheets
    ThisWorkbook.Worksheets(HELP_STEPS_SHEET).Activate

    MsgBox "「" & HELP_STEPS_SHEET & "」と「" & HELP_TABLE_SHEET & "」を作りました。", _
           vbInformation, "説明シートの作成"
End Sub

Public Sub BuildHelpSheets()

    Application.ScreenUpdating = False
    BuildStepsSheet
    BuildTableSheet
    Application.ScreenUpdating = True
End Sub

'==================================================================
' 手順書
'==================================================================
Private Sub BuildStepsSheet()

    Dim t As Variant
    Dim i As Long

    BeginSheet HELP_STEPS_SHEET

    TitleLine "大引ティック判定　手順書"
    NoteLine "このシートはマクロが自動生成します（「手順書・表の見方を作り直す」で再作成）"
    BlankLine

    HeadLine "毎日の運用（ふだんはこれだけ）"
    Step2 "1", "ブックを開く（マクロ有効）", _
          "開いた時点で " & AUTO_ARM_TIME & " に記録開始が自動予約されます。ボタンを押す必要はありません。"
    Step2 "2", "マーケットスピードにログインしておく", _
          "RSS が配信していないと、記録は 0 件のまま終わります。ログインは " & JUDGE_START & " までに。"
    Step2 "3", "そのまま放置する", _
          JUDGE_START & " に記録開始 → " & JUDGE_END & " に記録終了 → そのまま自動で判定まで走ります。" & _
          "記録中はブックを触らないでください。"
    Step2 "4", "Judge_Results を見る", _
          JUDGE_END & " の数秒後に 買い / 売り / 中立 / 未計測 が出ます。読み方は「" & HELP_TABLE_SHEET & "」シートへ。"
    Step2 "5", "15:25 プレクロージング", _
          "「中立」は板を見て最終判断。15:30 の板寄せまで約 6 分あります。"
    BlankLine

    HeadLine "朝いちで見ておくとよいもの"
    Row2 "設定の点検", "判定窓・パラメータ・各銘柄シートの歩み値の鮮度が一覧で出ます。「要確認」が空なら健全です。"
    Row2 "歩み値ブロックの確認", "「配信は生きています」と出れば OK。「" & STALE_LIMIT_SEC & " 秒以上止まっています」なら RSS を疑ってください。"
    BlankLine

    HeadLine "銘柄を変える・増やす"
    Row2 "変える", "銘柄シートの B3 に証券コードを入れ直し、「① 準備」を押す。"
    Row2 "増やす", "「銘柄シートを追加」でコードを入力。歩み値の数式もその場で設定できます。"
    Row2 "注意", "同じコードを複数のシートに入れると、RSS の負荷が増えて結果も重複します。" & _
                 "「① 準備」を押すと重複を警告します。"
    BlankLine

    HeadLine "うまくいかないとき（症状 → 見るところ）"
    Row2 "判定が「未計測」", "判定窓にティックが 1 件も入らなかった状態です。Judge_Results の Z6（警告）を読んでください。"
    Row2 "「配信停止の疑い」と出た", STALE_LIMIT_SEC & " 秒ティックが来ていません。マーケットスピードの接続を確認。回復すれば自動で表示も戻ります。"
    Row2 "配信状態が「接続待ち」", "マーケットスピード未ログイン、または RSS が未接続です。"
    Row2 "記録が始まらない", "Judge_Results の X2（状態）を見る。「本日の判定時間は終了」なら時間外です。"
    Row2 "ボタンが反応しない", "記録中は事故防止のため他のボタンを止めています。「③ 記録 停止 → 判定」で止めてから操作してください。"
    Row2 "原因を切り分けたい", "ザラバ中に「★ RSS 疎通テスト」を 1 回。約 20 秒で原因を名指しします。"
    BlankLine

    HeadLine "ボタン一覧"
    t = ButtonTable()
    For i = LBound(t) To UBound(t)
        Row2 CStr(t(i)(0)), CStr(t(i)(2))
    Next i
    BlankLine

    HeadLine "触らないでほしいところ"
    Row2 "Judge_Results の AA2:AA4", "Application.OnTime の予約時刻を保存しています。消すと予約を解除できなくなります。"
    Row2 "銘柄シートの " & TICK_FORMULA_CELL, "歩み値（RssTickList）の数式です。ここが空だと出来高ポーリングに落ちます。"
    Row2 "銘柄シートの D:N 列", "記録中にマクロが書き込みます。手で編集しないでください。"
    BlankLine

    NoteLine "各シートの 1 行目は自由なメモ欄です。マクロは書き込みません。"

    EndSheet 30, 96
End Sub

'==================================================================
' 表の見方
'==================================================================
Private Sub BuildTableSheet()

    BeginSheet HELP_TABLE_SHEET

    TitleLine "表の見方"
    NoteLine "このシートはマクロが自動生成します（「手順書・表の見方を作り直す」で再作成）"
    BlankLine

    HeadLine "まず結論だけ見るなら"
    Row2 "Judge_Results の D 列", "ここが「買い」「売り」「中立」「未計測」。これだけ見れば判断できます。"
    Row2 "X2 / X3 / X4", "状態 / 件数 / 内訳。判定が終わっているか、何件処理したかが分かります。"
    BlankLine

    HeadLine "判定（D 列）の決まり方"
    Row2 "買い", "① が「引け上方向」かつ ② の Vol4 優勢が ○ のときだけ。"
    Row2 "売り", "① が「引け下方向」かつ ② の Vol4 優勢が ○ のときだけ。"
    Row2 "中立", "①②のどちらかが欠けている状態。板を見て最終判断してください。"
    Row2 "未計測", "判定窓にティックが 1 件もありませんでした。相場の判断ではなく、配信の異常を疑ってください。"
    BlankLine

    HeadLine "①②③の中身"
    Row2 "① 方向判定", "同じ向きの値動きが " & SEQ_MIN & " 回以上続き、かつ反対方向の連続より多いこと。" & _
                       "前ティックと同値のときは連続が" & IIf(RESET_ON_FLAT, "途切れます", "続きます") & "。"
    Row2 "② 出来高初動", "判定窓を 4 等分し、最終区間 Vol4 が Vol3 の " & VOL_RATIO & " 倍を超え、" & _
                         "かつ Vol1・Vol2 も上回ること。終盤に大口が入ったかを見ています。"
    Row2 "③ 約定速度", "同一秒内の約定が " & SPEED_MIN & " 回以上あれば信頼度「高（アルゴ/大口）」。" & _
                       "判定そのものは変えません。確信度の目安です。"
    BlankLine

    HeadLine "Judge_Results の列"
    Col3 "B", "コード", "証券コード"
    Col3 "C", "銘柄名", "RSS から取得した銘柄名称"
    Col3 "D", "判定", "買い / 売り / 中立 / 未計測（緑・赤・灰・黄で色分け）"
    Col3 "E", "信頼度", "③の結果。「高（アルゴ/大口）」か「通常」"
    Col3 "F", "UpSeqMax", "①上昇の最大連続回数"
    Col3 "G", "DnSeqMax", "①下落の最大連続回数"
    Col3 "H", "方向判定①", "①の結論（引け上方向 / 引け下方向 / 中立）"
    Col3 "I:L", "Vol1～Vol4", "判定窓を 4 等分した各区間の出来高（株数）。区間の時刻は銘柄シートの O 列に出ます"
    Col3 "M", "Vol4優勢②", "②の結論。○ なら終盤に大口が集中"
    Col3 "N", "SpeedMax③", "同一秒内の最大約定回数"
    Col3 "O", "ティック数", "判定に使ったティックの本数。0 なら未計測"
    Col3 "P:Q", "BuyTotal / SellTotal", "参考スコア。出来高の重み付き合計＋気配の切り上げ／切り下げ回数。中立のとき板を見る材料に"
    Col3 "R", "更新時刻", "判定した時刻"
    Col3 "S:U", "最良買気配 / 最良売気配 / スプレッド", "記録中は 1 秒おきに更新されます"
    Col3 "W:X", "状態 / 件数 / 内訳", "いま何をしているか。X2 は色でも状態を示します（緑=稼働中、黄=警告）"
    Col3 "Y:Z", "実行ログ", "開始時刻・取込モード・配信状態・最終ティック・警告・終了時刻。あとから検証するための記録"
    Col3 "AA", "内部管理", "OnTime の予約時刻。触らないでください"
    BlankLine

    HeadLine "銘柄シートの列"
    Col3 "B3 / C3", "コード / 銘柄名称", "B3 に証券コードを入れるとこのシートが判定対象になります"
    Col3 "D:F", "時刻 / 出来高 / 約定値", "記録した 1 ティック 1 行。出来高はそのティック単体の株数"
    Col3 "G:I", "重み / UpScore / DnScore", "参考スコアの計算過程。判定そのものには使いません"
    Col3 "J:K", "最良買気配値 / 最良売気配値", "そのティック時点の気配（歩み値には気配が無いので記録時にスタンプ）"
    Col3 "L", "方向", "↑ 上昇 / ↓ 下落 / → 変わらず"
    Col3 "M", "秒内約定数", "その秒で何本目の約定か。最大値が SpeedMax③ になります"
    Col3 "N", "ティック", "取引所のティック記号。取得項目に入れていなければ空欄です"
    Col3 "O:P", "判定結果ブロック", "この銘柄の判定結果。Judge_Results の 1 行ぶんと同じ内容"
    Col3 "S:T", "RSS ライブ取得", "現在値・出来高・気配・現在値時刻。RSS が直接書き込みます"
    Col3 "AA:AD", "歩み値ブロック", "RssTickList の出力。" & TICK_FORMULA_CELL & " が配信状態、その 2 行下から新しい順にティックが並びます"
    BlankLine

    HeadLine "色の意味"
    Row2 "緑", "買い / 記録が正常に動いている"
    Row2 "赤", "売り / エラー"
    Row2 "灰", "中立"
    Row2 "黄", "未計測・警告（配信停止の疑いなど）。数字より先にここを見てください"
    BlankLine

    HeadLine "数字を読むときの注意"
    Row2 "出来高の単位", "株数です。Vol1～Vol4 は区間内の合計。"
    Row2 "SpeedMax が正確なのは", "「歩み値追従」モードのときだけです。「出来高ポーリング」では 1 秒 1 本が上限なので、③は当てになりません。" & _
                                  "いまどちらで動いたかは Judge_Results の Z3 に出ます。"
    Row2 "判定窓の外は数えません", JUDGE_START & "～" & JUDGE_END & " の外にあるティックは、記録されていても判定に入りません。" & _
                                  "15:30 の板寄せ約定を窓に入れないための設計です。"
    Row2 "BuyTotal / SellTotal", "あくまで参考値です。買い／売りの判定はこの数字ではなく ①②で決まります。"

    EndSheet 30, 96
End Sub

'==================================================================
' 書き込みヘルパー
'==================================================================

Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Sub BeginSheet(ByVal nm As String)

    Set mHws = SheetByName(nm)

    If mHws Is Nothing Then
        Set mHws = ThisWorkbook.Worksheets.Add(After:=ResultSheet())
        mHws.Name = nm
    End If

    '   1行目はメモ欄として残し、2行目以降だけ作り直します
    mHws.Range(mHws.Rows(2), mHws.Rows(2000)).Clear
    mHrow = 2
End Sub

Private Sub EndSheet(ByVal wL As Double, ByVal wR As Double)

    With mHws
        .Columns(1).ColumnWidth = 2
        .Columns(COL_L).ColumnWidth = wL
        .Columns(COL_R).ColumnWidth = wR
        .Columns(COL_R).WrapText = True
        .Rows.AutoFit
    End With

    Set mHws = Nothing
End Sub

Private Sub TitleLine(ByVal t As String)

    With mHws.Cells(mHrow, COL_L)
        .Value = t
        .Font.Bold = True
        .Font.Size = 16
    End With
    mHrow = mHrow + 1
End Sub

Private Sub HeadLine(ByVal t As String)

    With mHws.Range(mHws.Cells(mHrow, COL_L), mHws.Cells(mHrow, COL_R))
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(120, 150, 190)
    End With
    mHws.Cells(mHrow, COL_L).Value = t
    mHrow = mHrow + 1
End Sub

' 番号つきの手順
Private Sub Step2(ByVal n As String, ByVal what As String, ByVal detail As String)

    Row2 n & "　" & what, detail
    mHws.Cells(mHrow - 1, COL_L).Font.Bold = True
End Sub

' 左＝キー / 右＝説明
Private Sub Row2(ByVal k As String, ByVal v As String)

    With mHws
        .Cells(mHrow, COL_L).Value = k
        .Cells(mHrow, COL_R).Value = v
        With .Range(.Cells(mHrow, COL_L), .Cells(mHrow, COL_R))
            .VerticalAlignment = xlTop
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
        End With
        .Cells(mHrow, COL_L).Font.Color = RGB(31, 58, 104)
    End With
    mHrow = mHrow + 1
End Sub

' 列番号つき（セル番地 / 見出し / 意味）
Private Sub Col3(ByVal addr As String, ByVal caption As String, ByVal meaning As String)
    Row2 addr & "　" & caption, meaning
End Sub

Private Sub NoteLine(ByVal t As String)

    With mHws.Cells(mHrow, COL_L)
        .Value = t
        .Font.Color = RGB(120, 120, 120)
        .Font.Italic = True
    End With
    mHrow = mHrow + 1
End Sub

Private Sub BlankLine()
    mHrow = mHrow + 1
End Sub
