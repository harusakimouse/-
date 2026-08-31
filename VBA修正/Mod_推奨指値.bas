Attribute VB_Name = "Mod_推奨指値"
'==============================================================================
' 総合売買抽出BOOK　推奨指値モジュール  v1.0  (2026/08/15)
'
'   買抽出v13 / 売抽出v13 / 厳選TOP2 / 個別銘柄 の4シートに
'   「いくらで入るか」を表示する。
'
'------------------------------------------------------------------------------
' 【検証の根拠】300銘柄 × 250営業日（2025/07/31～2026/08/14）の実データ
'
'  ■ 買い（シグナル6,975件）
'     ・寄付で買うと 95.8% が当日一度は含み損（中央 -1.33%）
'     ・翌日安値は当日終値から 中央 0.46ATR 下まで落ちる
'     ・下落幅は%では条件により10倍変わるが、ATRで割ると 0.41～0.60 でほぼ一定
'     ・毎日 上位10候補すべてに 終値-0.75ATR の指値／3営業日有効 が最良
'          翌日寄付で成行 : 平均 +0.179R  勝率 46.2%  年間494取引
'          指値 -0.75ATR : 平均 +0.413R  勝率 53.1%  年間477取引
'       期間前半/後半・建玉1～5銘柄・高精度シグナル限定 いずれでも優劣は不変
'     ・「急騰銘柄はもっと深く指す」「過熱銘柄は見送る」は逆効果だった
'          一律0.75ATR +0.408R → 過熱度連動 +0.357R → 過熱は見送り +0.296R
'
'  ■ 売り（シグナル3,571件）── 買いとは非対称
'     ・翌日高値は当日終値から 中央 0.39ATR 上まで戻る（終値超え率 84.0%）
'     ・しかし「戻りを待つ指値売り」は成績が落ちた
'          翌日寄付で成行  : 平均 +0.062R  勝率 48.1%
'          指値 +0.50ATR  : 平均 +0.010R
'          指値 +0.75ATR  : 平均 -0.044R
'          逆指値 -0.25ATR: 平均 -0.056R  （下抜け確認も改善しない）
'       → 売りは寄付成行のまま。ただし検証期間は +16.2% の上昇相場1年のみで、
'         空売りに構造的に不利な期間。売りの結論は信頼度が低い。
'
'  ■ 決済（V811設定に合わせている）
'     損切 = 建値 -/+ 2ATR（建値の3%～10%に制限）／ 利確 = 損切幅 × 2
'
'  ※ 約定はすべて「安値/高値が指値にタッチすれば約定」として計算している。
'     板の順番で刺さらないことがあるため、実効約定率は表示より低くなる。
'------------------------------------------------------------------------------
' 【使い方】
'   1. VBEで [ファイル]→[ファイルのインポート] からこの .bas を読み込む
'   2. 推奨指値_全シート更新 を実行（またはボタン設置を1回だけ実行）
'   3. 買抽出v13 と 売抽出v13 は U～AC列、厳選TOP2 は N～Q列、
'      個別銘柄 は A24～C29 に出力される（既存の列は一切書き換えない）
'==============================================================================
Option Explicit

'--- 買いの指値（ATR倍数） --------------------------------------------------
Public Const BUY_K_LIGHT As Double = 0.5     '第1指値（浅い・約定しやすい）
Public Const BUY_K_MAIN  As Double = 0.75    '★本命指値（検証で最良）
Public Const BUY_K_DEEP  As Double = 1.25    '深押し指値
Public Const BUY_K_MAIN_HIGHSCORE As Double = 1#  '高精度シグナル(16以上)はここまで引ける
Public Const BUY_VALID_DAYS As Long = 3      '指値の有効日数（営業日）

'--- 売りの参考水準（ATR倍数・エントリーは寄付成行） ------------------------
Public Const SELL_K_MID  As Double = 0.39    '翌日高値の中央到達点
Public Const SELL_K_HIGH As Double = 0.7     '同 P75

'--- 決済（V811設定と同じ） -------------------------------------------------
Public Const ATR_PERIOD   As Long = 14
Public Const STOP_ATR     As Double = 2#
Public Const STOP_MIN_PCT As Double = 0.03
Public Const STOP_MAX_PCT As Double = 0.1
Public Const REWARD_RATIO As Double = 2#
Public Const HITRATE_LOOKBACK As Long = 120  '到達率を測る過去日数

'--- シート・レイアウト -----------------------------------------------------
Private Const SH_BUY  As String = "買抽出v13"
Private Const SH_SELL As String = "売抽出v13"
Private Const SH_TOP2 As String = "厳選TOP2"
Private Const SH_ONE  As String = "個別銘柄"
Private Const SH_HI   As String = "高値"
Private Const SH_LO   As String = "安値"
Private Const SH_CL   As String = "終値"

Private Const EXT_DATA_ROW As Long = 4       '買抽出v13 / 売抽出v13 のデータ開始行
Private Const EXT_CODE_COL As Long = 3       '　同 銘柄コード列(C)
Private Const OUT_COL      As Long = 21      '　同 出力開始列(U)
Private Const OUT_WIDTH    As Long = 9       '　同 出力列数(U～AC)

Private Const TOP2_BUY_HDR  As Long = 17     '厳選TOP2 買い候補ヘッダー行
Private Const TOP2_BUY_ROWS As Long = 5      '　同 データ5行(18～22)
Private Const TOP2_SEL_HDR  As Long = 25     '厳選TOP2 売り候補ヘッダー行
Private Const TOP2_SEL_ROWS As Long = 5      '　同 データ5行(26～30)
Private Const TOP2_CODE_COL As Long = 2      '　同 銘柄コード列(B)
Private Const TOP2_OUT_COL  As Long = 14     '　同 出力開始列(N～Q)

Private Const ONE_CODE_CELL As String = "B3" '個別銘柄 銘柄コード入力セル
Private Const ONE_OUT_ROW   As Long = 24     '　同 出力開始行(24～29)

Private Const STOCK_DATA_START As Long = 6   '高値/安値/終値 のデータ開始行
Private Const HIST_FIRST_COL   As Long = 5   '　同 E列＝最新日
Private Const SHEET_PW As String = "ne19480314"

Private Const COL_HDR   As Long = 8281600    'RGB(0,70,127) 濃紺
Private Const COL_WHITE As Long = 16777215
Private Const COL_GOLD  As Long = 55295      'RGB(255,215,0)
Private Const COL_RED   As Long = 192

'==============================================================================
' メイン
'==============================================================================
Public Sub 推奨指値_全シート更新()
    Dim t As Double: t = Timer
    Dim msg As String
    Dim nBuy As Long, nSell As Long, nTop2 As Long
    Dim okOne As Boolean

    Application.ScreenUpdating = False
    Application.StatusBar = "推奨指値を計算中..."

    If Not DataSheetsReady() Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "「高値」「安値」「終値」シートが必要です。", vbExclamation, "推奨指値"
        Exit Sub
    End If

    nBuy = 更新_抽出シート(SH_BUY, True)
    nSell = 更新_抽出シート(SH_SELL, False)
    nTop2 = 更新_厳選TOP2()
    okOne = 更新_個別銘柄()

    Application.StatusBar = False
    Application.ScreenUpdating = True

    msg = "推奨指値を更新しました。（" & Format(Timer - t, "0.0") & "秒）" & vbCrLf & vbCrLf
    msg = msg & "　買抽出v13 ： " & nBuy & " 銘柄　→ U～AC列" & vbCrLf
    msg = msg & "　売抽出v13 ： " & nSell & " 銘柄　→ U～AC列" & vbCrLf
    msg = msg & "　厳選TOP2 　： " & nTop2 & " 銘柄　→ N～Q列" & vbCrLf
    msg = msg & "　個別銘柄 　： " & IIf(okOne, "表示中", "対象なし") & "　→ A24～C29" & vbCrLf & vbCrLf
    msg = msg & "【買い】指値 = 終値 － " & BUY_K_MAIN & " × ATR14 を " & BUY_VALID_DAYS & "営業日有効で出す。" & vbCrLf
    msg = msg & "　　　　刺さらなければ取り消して翌日の候補へ。" & vbCrLf
    msg = msg & "【売り】検証では寄付成行が最良。戻り待ちの指値売りは成績が落ちる。"
    MsgBox msg, vbInformation, "推奨指値 v1.0"
End Sub

Public Sub 推奨指値_買抽出v13()
    If Not DataSheetsReady() Then Exit Sub
    MsgBox 更新_抽出シート(SH_BUY, True) & " 銘柄を更新しました（U～AC列）。", vbInformation, "推奨指値"
End Sub

Public Sub 推奨指値_売抽出v13()
    If Not DataSheetsReady() Then Exit Sub
    MsgBox 更新_抽出シート(SH_SELL, False) & " 銘柄を更新しました（U～AC列）。", vbInformation, "推奨指値"
End Sub

' ★v1.1追加：メッセージを出さない版。
'   売り抽出／買い抽出の最後から自動で呼ばれる。
'   （抽出のたびに手で押さないと、右側U～AC列に前回の銘柄の指値が残ってしまうため）
Public Sub 推奨指値_売抽出v13_無言()
    If Not DataSheetsReady() Then Exit Sub
    更新_抽出シート SH_SELL, False
End Sub

Public Sub 推奨指値_買抽出v13_無言()
    If Not DataSheetsReady() Then Exit Sub
    更新_抽出シート SH_BUY, True
End Sub

Public Sub 推奨指値_厳選TOP2()
    If Not DataSheetsReady() Then Exit Sub
    MsgBox 更新_厳選TOP2() & " 銘柄を更新しました（N～Q列）。", vbInformation, "推奨指値"
End Sub

Public Sub 推奨指値_個別銘柄()
    If Not DataSheetsReady() Then Exit Sub
    If 更新_個別銘柄() Then
        MsgBox "個別銘柄シートを更新しました（A24～C29）。", vbInformation, "推奨指値"
    Else
        MsgBox "その銘柄コードは 終値シートに見つかりませんでした。", vbExclamation, "推奨指値"
    End If
End Sub

'==============================================================================
' 買抽出v13 / 売抽出v13
'==============================================================================
Private Function 更新_抽出シート(ByVal sheetName As String, ByVal isBuy As Boolean) As Long
    更新_抽出シート = 0
    Dim ws As Worksheet: Set ws = GetWsX(sheetName)
    If ws Is Nothing Then Exit Function

    Dim hiWs As Worksheet: Set hiWs = GetWsX(SH_HI)
    Dim loWs As Worksheet: Set loWs = GetWsX(SH_LO)
    Dim clWs As Worksheet: Set clWs = GetWsX(SH_CL)

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PW
    On Error GoTo 0

    HeaderOut ws, 2, OUT_COL, isBuy

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, EXT_CODE_COL).End(xlUp).Row
    If lastRow < EXT_DATA_ROW Then lastRow = EXT_DATA_ROW
    ws.Range(ws.Cells(EXT_DATA_ROW, OUT_COL), _
             ws.Cells(lastRow + 300, OUT_COL + OUT_WIDTH - 1)).Clear

    Dim dic As Object: Set dic = BuildCodeIndex(clWs)
    Dim r As Long, cnt As Long
    For r = EXT_DATA_ROW To lastRow
        Dim code As String: code = NormCode(ws.Cells(r, EXT_CODE_COL).Value)
        If code <> "" Then
            If dic.Exists(code) Then
                Dim sc As Double: sc = SafeNumX(ws.Cells(r, 17).Value, 0)   'Q列＝スコア
                If OneRow(ws, r, OUT_COL, CLng(dic(code)), hiWs, loWs, clWs, isBuy, sc) Then
                    cnt = cnt + 1
                End If
            End If
        End If
    Next r

    ws.Range(ws.Cells(EXT_DATA_ROW, OUT_COL), ws.Cells(lastRow, OUT_COL + OUT_WIDTH - 1)).Font.Size = 11
    On Error Resume Next
    ws.Columns(OUT_COL).Resize(, OUT_WIDTH).AutoFit
    On Error GoTo 0
    更新_抽出シート = cnt
End Function

Private Sub HeaderOut(ByVal ws As Worksheet, ByVal titleRow As Long, ByVal c0 As Long, ByVal isBuy As Boolean)
    Dim h As Variant, i As Long
    If isBuy Then
        ws.Cells(titleRow, c0).Value = "▼ 推奨指値 v1.0 ： 終値 － " & BUY_K_MAIN & " × ATR14 を " & _
                                       BUY_VALID_DAYS & "営業日有効で出す（平均 +0.413R / 勝率 53.1%）"
        h = Array("基準終値", "ATR14", "第1指値" & vbLf & "-" & BUY_K_LIGHT & "ATR", _
                  "★本命指値" & vbLf & "-" & BUY_K_MAIN & "ATR", "深押し" & vbLf & "-" & BUY_K_DEEP & "ATR", _
                  "本命" & vbLf & "到達率", "約定時" & vbLf & "損切", "約定時" & vbLf & "利確", "ひとこと")
    Else
        ws.Cells(titleRow, c0).Value = "▼ 推奨エントリー v1.0 ： 売りは寄付成行が最良（+0.062R）。" & _
                                       "戻り待ちの指値売りは成績が落ちる（+0.75ATRで -0.044R）"
        h = Array("基準終値", "ATR14", "戻り目安" & vbLf & "+" & SELL_K_MID & "ATR", _
                  "戻り上限" & vbLf & "+" & SELL_K_HIGH & "ATR", "★エントリー" & vbLf & "寄付成行", _
                  "戻り" & vbLf & "到達率", "建値時" & vbLf & "損切", "建値時" & vbLf & "利確", "ひとこと")
    End If
    ws.Cells(titleRow, c0).Font.Bold = True
    For i = 0 To UBound(h)
        With ws.Cells(3, c0 + i)
            .Value = h(i)
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = COL_HDR
            .Font.Color = COL_WHITE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
End Sub

Private Function OneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal c0 As Long, _
                        ByVal srcRow As Long, ByVal hiWs As Worksheet, ByVal loWs As Worksheet, _
                        ByVal clWs As Worksheet, ByVal isBuy As Boolean, ByVal scoreVal As Double) As Boolean
    OneRow = False
    Dim h() As Double, L() As Double, c() As Double, a() As Double, cnt As Long
    If Not LoadBars(hiWs, loWs, clWs, srcRow, h, L, c, a, cnt) Then Exit Function

    Dim px As Double: px = c(cnt)
    Dim atrNow As Double: atrNow = a(cnt)
    If px <= 0 Or atrNow <= 0 Then Exit Function

    Dim kMain As Double: kMain = BUY_K_MAIN
    If isBuy And scoreVal >= 16 Then kMain = BUY_K_MAIN_HIGHSCORE

    Dim p1 As Double, p2 As Double, p3 As Double, entry As Double, hit As Double, sd As Double
    If isBuy Then
        p1 = TickDown(px - BUY_K_LIGHT * atrNow)
        p2 = TickDown(px - kMain * atrNow)
        p3 = TickDown(px - BUY_K_DEEP * atrNow)
        entry = p2
        hit = HitRateDown(h, L, c, a, cnt, kMain)
        sd = StopDist(entry, atrNow)
    Else
        p1 = TickUp(px + SELL_K_MID * atrNow)
        p2 = TickUp(px + SELL_K_HIGH * atrNow)
        p3 = px                                   '★エントリー＝寄付成行（基準は直近終値）
        entry = px
        hit = HitRateUp(h, L, c, a, cnt, SELL_K_MID)
        sd = StopDist(entry, atrNow)
    End If

    With ws
        .Cells(r, c0).Value = px
        .Cells(r, c0 + 1).Value = Round(atrNow, 1)
        .Cells(r, c0 + 2).Value = p1
        .Cells(r, c0 + 3).Value = p2
        .Cells(r, c0 + 4).Value = p3
        .Cells(r, c0 + 5).Value = Format(hit, "0.0") & "%"
        If isBuy Then
            .Cells(r, c0 + 6).Value = TickDown(entry - sd)
            .Cells(r, c0 + 7).Value = TickUp(entry + REWARD_RATIO * sd)
        Else
            .Cells(r, c0 + 6).Value = TickUp(entry + sd)
            .Cells(r, c0 + 7).Value = TickDown(entry - REWARD_RATIO * sd)
        End If
        .Cells(r, c0 + 8).Value = Hitokoto(px, atrNow, c, cnt, isBuy, scoreVal, kMain)

        .Range(.Cells(r, c0), .Cells(r, c0 + 4)).NumberFormat = "#,##0"
        .Range(.Cells(r, c0 + 6), .Cells(r, c0 + 7)).NumberFormat = "#,##0"
        .Range(.Cells(r, c0), .Cells(r, c0 + 8)).HorizontalAlignment = xlRight
        .Cells(r, c0 + 8).HorizontalAlignment = xlLeft

        Dim hi As Range
        If isBuy Then Set hi = .Cells(r, c0 + 3) Else Set hi = .Cells(r, c0 + 4)
        hi.Interior.Color = COL_GOLD
        hi.Font.Bold = True
    End With
    OneRow = True
End Function

'==============================================================================
' 厳選TOP2
'==============================================================================
Private Function 更新_厳選TOP2() As Long
    更新_厳選TOP2 = 0
    Dim ws As Worksheet: Set ws = GetWsX(SH_TOP2)
    If ws Is Nothing Then Exit Function
    Dim hiWs As Worksheet: Set hiWs = GetWsX(SH_HI)
    Dim loWs As Worksheet: Set loWs = GetWsX(SH_LO)
    Dim clWs As Worksheet: Set clWs = GetWsX(SH_CL)

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PW
    On Error GoTo 0

    Top2Header ws, TOP2_BUY_HDR, True
    Top2Header ws, TOP2_SEL_HDR, False

    ws.Range(ws.Cells(TOP2_BUY_HDR + 1, TOP2_OUT_COL), _
             ws.Cells(TOP2_BUY_HDR + TOP2_BUY_ROWS, TOP2_OUT_COL + 3)).Clear
    ws.Range(ws.Cells(TOP2_SEL_HDR + 1, TOP2_OUT_COL), _
             ws.Cells(TOP2_SEL_HDR + TOP2_SEL_ROWS, TOP2_OUT_COL + 3)).Clear

    Dim dic As Object: Set dic = BuildCodeIndex(clWs)
    Dim cnt As Long
    cnt = cnt + Top2Block(ws, TOP2_BUY_HDR + 1, TOP2_BUY_ROWS, True, dic, hiWs, loWs, clWs)
    cnt = cnt + Top2Block(ws, TOP2_SEL_HDR + 1, TOP2_SEL_ROWS, False, dic, hiWs, loWs, clWs)
    On Error Resume Next
    ws.Columns(TOP2_OUT_COL).Resize(, 4).AutoFit
    On Error GoTo 0
    更新_厳選TOP2 = cnt
End Function

Private Sub Top2Header(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal isBuy As Boolean)
    Dim h As Variant, i As Long
    If isBuy Then
        h = Array("ATR14", "★本命指値" & vbLf & "-" & BUY_K_MAIN & "ATR", "到達率", "約定時損切")
    Else
        h = Array("ATR14", "★エントリー" & vbLf & "寄付成行", "戻り目安" & vbLf & "+" & SELL_K_MID & "ATR", "建値時損切")
    End If
    For i = 0 To UBound(h)
        With ws.Cells(hdrRow, TOP2_OUT_COL + i)
            .Value = h(i)
            .Font.Bold = True
            .Interior.Color = COL_HDR
            .Font.Color = COL_WHITE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
End Sub

Private Function Top2Block(ByVal ws As Worksheet, ByVal r0 As Long, ByVal nRows As Long, _
                           ByVal isBuy As Boolean, ByVal dic As Object, _
                           ByVal hiWs As Worksheet, ByVal loWs As Worksheet, _
                           ByVal clWs As Worksheet) As Long
    Dim i As Long, cnt As Long
    For i = 0 To nRows - 1
        Dim r As Long: r = r0 + i
        Dim code As String: code = NormCode(ws.Cells(r, TOP2_CODE_COL).Value)
        If code <> "" Then
            If dic.Exists(code) Then
                Dim h() As Double, L() As Double, c() As Double, a() As Double, cntB As Long
                If LoadBars(hiWs, loWs, clWs, CLng(dic(code)), h, L, c, a, cntB) Then
                    Dim px As Double: px = c(cntB)
                    Dim atrNow As Double: atrNow = a(cntB)
                    If px > 0 And atrNow > 0 Then
                        Dim sd As Double
                        With ws
                            .Cells(r, TOP2_OUT_COL).Value = Round(atrNow, 1)
                            If isBuy Then
                                Dim lim As Double: lim = TickDown(px - BUY_K_MAIN * atrNow)
                                sd = StopDist(lim, atrNow)
                                .Cells(r, TOP2_OUT_COL + 1).Value = lim
                                .Cells(r, TOP2_OUT_COL + 2).Value = _
                                    Format(HitRateDown(h, L, c, a, cntB, BUY_K_MAIN), "0.0") & "%"
                                .Cells(r, TOP2_OUT_COL + 3).Value = TickDown(lim - sd)
                            Else
                                sd = StopDist(px, atrNow)
                                .Cells(r, TOP2_OUT_COL + 1).Value = px
                                .Cells(r, TOP2_OUT_COL + 2).Value = TickUp(px + SELL_K_MID * atrNow)
                                .Cells(r, TOP2_OUT_COL + 3).Value = TickUp(px + sd)
                            End If
                            .Range(.Cells(r, TOP2_OUT_COL), .Cells(r, TOP2_OUT_COL + 3)).NumberFormat = "#,##0"
                            .Cells(r, TOP2_OUT_COL + 2).NumberFormat = "General"
                            With .Cells(r, TOP2_OUT_COL + 1)
                                .Interior.Color = COL_GOLD
                                .Font.Bold = True
                            End With
                        End With
                        cnt = cnt + 1
                    End If
                End If
            End If
        End If
    Next i
    Top2Block = cnt
End Function

'==============================================================================
' 個別銘柄
'==============================================================================
Private Function 更新_個別銘柄() As Boolean
    更新_個別銘柄 = False
    Dim ws As Worksheet: Set ws = GetWsX(SH_ONE)
    If ws Is Nothing Then Exit Function
    Dim hiWs As Worksheet: Set hiWs = GetWsX(SH_HI)
    Dim loWs As Worksheet: Set loWs = GetWsX(SH_LO)
    Dim clWs As Worksheet: Set clWs = GetWsX(SH_CL)

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PW
    On Error GoTo 0

    Dim r As Long
    ws.Range(ws.Cells(ONE_OUT_ROW, 1), ws.Cells(ONE_OUT_ROW + 5, 3)).Clear

    With ws.Cells(ONE_OUT_ROW, 1)
        .Value = "■ 推奨指値（ATR連動 v1.0）"
        .Font.Bold = True
    End With

    Dim code As String: code = NormCode(ws.Range(ONE_CODE_CELL).Value)
    If code = "" Then
        ws.Cells(ONE_OUT_ROW + 1, 1).Value = "銘柄コードを B3 に入力してください"
        Exit Function
    End If

    Dim dic As Object: Set dic = BuildCodeIndex(clWs)
    If Not dic.Exists(code) Then
        ws.Cells(ONE_OUT_ROW + 1, 1).Value = "終値シートにこの銘柄がありません"
        Exit Function
    End If

    Dim h() As Double, L() As Double, c() As Double, a() As Double, cnt As Long
    If Not LoadBars(hiWs, loWs, clWs, CLng(dic(code)), h, L, c, a, cnt) Then
        ws.Cells(ONE_OUT_ROW + 1, 1).Value = "株価データが不足しています（ATR14の計算に15日必要）"
        Exit Function
    End If
    Dim px As Double: px = c(cnt)
    Dim atrNow As Double: atrNow = a(cnt)
    If px <= 0 Or atrNow <= 0 Then Exit Function

    Dim labs As Variant, vals As Variant, memo As Variant
    labs = Array("基準終値", "ATR14", "買 第1 －" & BUY_K_LIGHT & "ATR", _
                 "買 ★本命 －" & BUY_K_MAIN & "ATR", "買 深押し －" & BUY_K_DEEP & "ATR")
    vals = Array(px, Round(atrNow, 1), TickDown(px - BUY_K_LIGHT * atrNow), _
                 TickDown(px - BUY_K_MAIN * atrNow), TickDown(px - BUY_K_DEEP * atrNow))
    memo = Array("売りは寄付成行が最良（検証結果）", _
                 "ボラティリティ " & Format(atrNow / px * 100, "0.0") & "%", _
                 "到達率 " & Format(HitRateDown(h, L, c, a, cnt, BUY_K_LIGHT), "0.0") & "%", _
                 "到達率 " & Format(HitRateDown(h, L, c, a, cnt, BUY_K_MAIN), "0.0") & "%　" & _
                 BUY_VALID_DAYS & "営業日有効で出す", _
                 "到達率 " & Format(HitRateDown(h, L, c, a, cnt, BUY_K_DEEP), "0.0") & "%")

    Dim i As Long
    For i = 0 To 4
        r = ONE_OUT_ROW + 1 + i
        ws.Cells(r, 1).Value = labs(i)
        ws.Cells(r, 2).Value = vals(i)
        ws.Cells(r, 2).NumberFormat = "#,##0"
        ws.Cells(r, 3).Value = memo(i)
        ws.Cells(r, 3).Font.Size = 9
    Next i
    With ws.Cells(ONE_OUT_ROW + 4, 2)
        .Interior.Color = COL_GOLD
        .Font.Bold = True
    End With
    更新_個別銘柄 = True
End Function

'==============================================================================
' 株価系列とATR
'==============================================================================
'   シートは E列＝最新 なので、配列は 1=最古 … cnt=最新 に詰め直す
Private Function LoadBars(ByVal hiWs As Worksheet, ByVal loWs As Worksheet, ByVal clWs As Worksheet, _
                          ByVal srcRow As Long, ByRef h() As Double, ByRef L() As Double, _
                          ByRef c() As Double, ByRef a() As Double, ByRef cnt As Long) As Boolean
    LoadBars = False
    On Error GoTo Fail
    Dim lastCol As Long: lastCol = HistLastCol(clWs)
    Dim maxN As Long: maxN = lastCol - HIST_FIRST_COL + 1
    If maxN < ATR_PERIOD + 2 Then Exit Function
    If maxN > HITRATE_LOOKBACK + ATR_PERIOD + 10 Then maxN = HITRATE_LOOKBACK + ATR_PERIOD + 10
    lastCol = HIST_FIRST_COL + maxN - 1

    Dim hv As Variant, lv As Variant, cv As Variant
    hv = hiWs.Range(hiWs.Cells(srcRow, HIST_FIRST_COL), hiWs.Cells(srcRow, lastCol)).Value
    lv = loWs.Range(loWs.Cells(srcRow, HIST_FIRST_COL), loWs.Cells(srcRow, lastCol)).Value
    cv = clWs.Range(clWs.Cells(srcRow, HIST_FIRST_COL), clWs.Cells(srcRow, lastCol)).Value

    ReDim h(1 To maxN): ReDim L(1 To maxN): ReDim c(1 To maxN)
    cnt = 0
    Dim k As Long, h1 As Double, l1 As Double, c1 As Double
    For k = maxN To 1 Step -1
        h1 = SafeNumX(hv(1, k), 0): l1 = SafeNumX(lv(1, k), 0): c1 = SafeNumX(cv(1, k), 0)
        If h1 > 0 And l1 > 0 And c1 > 0 And h1 >= l1 Then
            cnt = cnt + 1
            h(cnt) = h1: L(cnt) = l1: c(cnt) = c1
        End If
    Next k
    If cnt < ATR_PERIOD + 2 Then Exit Function

    ReDim a(1 To cnt)
    Dim tr As Double, a1 As Double, seeded As Boolean
    a(1) = 0
    For k = 2 To cnt
        tr = h(k) - L(k)
        If Abs(h(k) - c(k - 1)) > tr Then tr = Abs(h(k) - c(k - 1))
        If Abs(L(k) - c(k - 1)) > tr Then tr = Abs(L(k) - c(k - 1))
        If Not seeded Then
            a1 = tr: seeded = True
        Else
            a1 = (a1 * (ATR_PERIOD - 1) + tr) / ATR_PERIOD
        End If
        a(k) = a1
    Next k
    LoadBars = (a(cnt) > 0)
    Exit Function
Fail:
End Function

'   その銘柄で「翌日安値 ≦ 終値 － k×ATR」になった実績の割合
Private Function HitRateDown(ByRef h() As Double, ByRef L() As Double, ByRef c() As Double, _
                             ByRef a() As Double, ByVal cnt As Long, ByVal k As Double) As Double
    Dim i As Long, n As Long, hitN As Long
    Dim first As Long: first = cnt - HITRATE_LOOKBACK
    If first < ATR_PERIOD + 1 Then first = ATR_PERIOD + 1
    For i = first To cnt - 1
        If a(i) > 0 Then
            n = n + 1
            If L(i + 1) <= c(i) - k * a(i) Then hitN = hitN + 1
        End If
    Next i
    If n > 0 Then HitRateDown = hitN / n * 100#
End Function

'   その銘柄で「翌日高値 ≧ 終値 + k×ATR」になった実績の割合
Private Function HitRateUp(ByRef h() As Double, ByRef L() As Double, ByRef c() As Double, _
                           ByRef a() As Double, ByVal cnt As Long, ByVal k As Double) As Double
    Dim i As Long, n As Long, hitN As Long
    Dim first As Long: first = cnt - HITRATE_LOOKBACK
    If first < ATR_PERIOD + 1 Then first = ATR_PERIOD + 1
    For i = first To cnt - 1
        If a(i) > 0 Then
            n = n + 1
            If h(i + 1) >= c(i) + k * a(i) Then hitN = hitN + 1
        End If
    Next i
    If n > 0 Then HitRateUp = hitN / n * 100#
End Function

Private Function StopDist(ByVal entry As Double, ByVal atrNow As Double) As Double
    Dim sd As Double: sd = STOP_ATR * atrNow
    If sd < entry * STOP_MIN_PCT Then sd = entry * STOP_MIN_PCT
    If sd > entry * STOP_MAX_PCT Then sd = entry * STOP_MAX_PCT
    StopDist = sd
End Function

Private Function Hitokoto(ByVal px As Double, ByVal atrNow As Double, ByRef c() As Double, _
                          ByVal cnt As Long, ByVal isBuy As Boolean, ByVal scoreVal As Double, _
                          ByVal kMain As Double) As String
    Dim s As String
    Dim atrPct As Double: atrPct = atrNow / px * 100#
    If isBuy And scoreVal >= 16 Then s = s & "高精度→" & kMain & "ATRまで引ける "
    If cnt >= 2 Then
        Dim chg As Double: chg = (c(cnt) / c(cnt - 1) - 1) * 100#
        If chg >= 9 Then
            s = s & "前日比" & Format(chg, "+0.0") & "% 急騰直後 "
        ElseIf chg <= -9 Then
            s = s & "前日比" & Format(chg, "+0.0") & "% 急落直後 "
        End If
    End If
    Dim e25 As Double: e25 = EMAx(c, cnt, 25)
    If e25 > 0 Then
        Dim dev As Double: dev = (px / e25 - 1) * 100#
        If isBuy Then
            If dev >= 15 Then s = s & "25日EMA乖離" & Format(dev, "+0") & "% 高値警戒 "
            If dev < 0 Then s = s & "25日EMA割れ 見送り推奨 "
        Else
            If dev <= -15 Then s = s & "25日EMA乖離" & Format(dev, "+0") & "% 売られ過ぎ "
            If dev > 0 Then s = s & "25日EMA上 売りには不利 "
        End If
    End If
    If atrPct >= 5 Then
        s = s & "高ボラ(ATR" & Format(atrPct, "0.0") & "%) 株数を絞る "
    ElseIf atrPct <= 1.5 Then
        s = s & "低ボラ(ATR" & Format(atrPct, "0.0") & "%) 指値は浅めで可 "
    End If
    If s = "" Then s = "標準"
    Hitokoto = Trim$(s)
End Function

Private Function EMAx(ByRef c() As Double, ByVal cnt As Long, ByVal period As Long) As Double
    If cnt < 2 Then Exit Function
    Dim k As Double: k = 2# / (period + 1#)
    Dim e As Double: e = c(1)
    Dim i As Long
    For i = 2 To cnt
        e = k * c(i) + (1# - k) * e
    Next i
    EMAx = e
End Function

'==============================================================================
' 補助
'==============================================================================
Private Function DataSheetsReady() As Boolean
    DataSheetsReady = Not (GetWsX(SH_HI) Is Nothing Or GetWsX(SH_LO) Is Nothing Or GetWsX(SH_CL) Is Nothing)
End Function

Private Function BuildCodeIndex(ByVal ws As Worksheet) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set BuildCodeIndex = dic
        Exit Function
    End If
    Dim last As Long: last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, c As String
    For i = STOCK_DATA_START To last
        c = NormCode(ws.Cells(i, 1).Value)
        If c <> "" Then
            If Not dic.Exists(c) Then dic.Add c, i
        End If
    Next i
    Set BuildCodeIndex = dic
End Function

Private Function NormCode(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Trim$(CStr(v))
    If s = "" Or s = "0" Or UCase$(s) = "TOPX" Then Exit Function
    If IsNumeric(s) And Len(s) < 4 Then s = Format(Val(s), "0000")
    NormCode = s
End Function

Private Function HistLastCol(ByVal ws As Worksheet) As Long
    Dim c As Long
    c = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    If c < HIST_FIRST_COL Then c = HIST_FIRST_COL
    HistLastCol = c
End Function

'   東証の呼値（一般刻み）。買い指値は切り捨て、売り指値は切り上げ。
Public Function TickSize(ByVal p As Double) As Double
    Select Case p
        Case Is < 3000: TickSize = 1
        Case Is < 5000: TickSize = 5
        Case Is < 30000: TickSize = 10
        Case Is < 50000: TickSize = 50
        Case Else: TickSize = 100
    End Select
End Function

Public Function TickDown(ByVal p As Double) As Double
    Dim t As Double: t = TickSize(p)
    TickDown = Int(p / t) * t
End Function

Public Function TickUp(ByVal p As Double) As Double
    Dim t As Double: t = TickSize(p)
    TickUp = -Int(-p / t) * t
End Function

Private Function SafeNumX(ByVal v As Variant, ByVal dflt As Double) As Double
    If IsError(v) Then
        SafeNumX = dflt
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeNumX = dflt
    ElseIf IsNumeric(v) Then
        SafeNumX = CDbl(v)
    Else
        SafeNumX = dflt
    End If
End Function

Private Function GetWsX(ByVal n As String) As Worksheet
    On Error Resume Next
    Set GetWsX = ThisWorkbook.Sheets(n)
    On Error GoTo 0
End Function

'==============================================================================
' ボタン設置（1回だけ実行すればよい）
'==============================================================================
Public Sub 推奨指値ボタン設置()
    PutButton SH_BUY, "btn推奨指値", "推奨指値を計算", "Mod_推奨指値.推奨指値_全シート更新", 600, 5
    PutButton SH_SELL, "btn推奨指値", "推奨指値を計算", "Mod_推奨指値.推奨指値_全シート更新", 600, 5
    PutButton SH_ONE, "btn推奨指値", "推奨指値を計算", "Mod_推奨指値.推奨指値_個別銘柄", 380, 5
    MsgBox "ボタンを設置しました。", vbInformation, "推奨指値"
End Sub

Private Sub PutButton(ByVal sheetName As String, ByVal btnName As String, ByVal caption As String, _
                      ByVal action As String, ByVal x As Double, ByVal y As Double)
    Dim ws As Worksheet: Set ws = GetWsX(sheetName)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes(btnName).Delete
    On Error GoTo 0
    Dim b As Shape
    Set b = ws.Shapes.AddFormControl(xlButtonControl, x, y, 170, 34)
    With b
        .Name = btnName
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Size = 13
        .TextFrame.Characters.Font.Bold = True
        .OnAction = action
    End With
End Sub
