Attribute VB_Name = "Module_Config"
Option Explicit

'==================================================================
' Module_Config : 共通設定・共通関数・共通型
'   ティック判定ロジック（大引け前の判定窓専用）
'------------------------------------------------------------------
'   運用タイムライン（既定の設定の場合）
'     14:59:30  自動開始の予約が発火（記録は 15:00:00 まで待機）
'     15:00:00  ティック記録 開始
'     15:24:00  ティック記録 終了 → 自動で判定実行
'     15:24:0x  Judge_Results に 買い / 売り / 中立 が出る
'     15:25:00  ザラバ終了（以降はクロージング・オークション）
'     15:30:00  板寄せ
'------------------------------------------------------------------
'   ※ 記録は Application.OnTime による1秒ごとの再スケジュール方式です。
'      1回のポーリングが終わるたびにマクロを完全に終了させ、Excel に制御を
'      返します。RSS（RTD）がシートへ値を書き込めるのは Excel がアイドルの
'      ときだけなので、ここを塞ぐと記録したいデータ自体が届かなくなります。
'      「Do While + DoEvents で回し続ける」実装に戻さないでください。
'==================================================================

'------------------------------------------------------------------
' 判定対象時間帯
'------------------------------------------------------------------
'   ザラバは 15:25:00 で終了し、以降はクロージング・オークション（気配のみ）。
'   15:30:00 の板寄せ約定は1本で桁違いの出来高になるため、絶対に窓へ入れないこと。
'   → JUDGE_END は 15:25:00 を超えてはいけません。
Public Const JUDGE_START As String = "15:00:00"
Public Const JUDGE_END   As String = "15:24:00"

'------------------------------------------------------------------
' 判定パラメータ（ここだけ触れば感度を変えられます）
'------------------------------------------------------------------
' ① 方向判定 : 連続上昇／連続下落が何回続いたら「方向あり」とみなすか
Public Const SEQ_MIN As Long = 5

' ① 変わらず（前ティックと同値）のティックで連続をリセットするか
Public Const RESET_ON_FLAT As Boolean = True

' ② 出来高初動判定 : Vol4 > Vol3 × この倍率
Public Const VOL_RATIO As Double = 1.5

' ③ 約定速度 : 同一秒内の約定回数がこの値以上なら「アルゴ／大口が本気」
Public Const SPEED_MIN As Long = 3

'------------------------------------------------------------------
' ティックの取込方式
'------------------------------------------------------------------
'   0 = 自動判定（既定）
'       歩み値ブロックが「生きていれば」1、そうでなければ 2 を選びます。
'       ここでいう「生きている」は、読めることではなく
'       最新ティックが STALE_LIMIT_SEC 秒以内であることです。
'   1 = 歩み値（TICK）ブロック追従  … 約定1本ずつ。③SpeedMax が正確に出ます
'   2 = 現在値・出来高ポーリング     … フォールバック。③は不正確になります
Public Const LOG_MODE As Long = 0

' ポーリング間隔（秒）
'   Application.OnTime の下限は1秒です。歩み値ブロックは TICK_SHOW_ROWS 本の
'   履歴を保持しているので、1秒間隔でも分解能は落ちません。
Public Const POLL_SEC As Long = 1

' 「配信が止まった」とみなすまでの秒数
'   最新ティック（またはRSSの現在値）がこの秒数だけ更新されなければ警告します。
Public Const STALE_LIMIT_SEC As Long = 60

'------------------------------------------------------------------
' 歩み値（TICK）ブロックの設定  ※ LOG_MODE = 0（自動）/ 1 のとき使用
'------------------------------------------------------------------
' 表示本数（RssTickList の上限は 300）
Public Const TICK_SHOW_ROWS As Long = 300

' RssTickList を書き込む位置（「歩み値の数式を設定」ボタンで使用）
'   TICK_FORMULA_CELL に数式（ここに配信状態が出ます）
'   TICK_ITEMS_CELL から右へ3セルが取得項目の見出し
'   データはその1行下から流れてきます
Public Const TICK_FORMULA_CELL As String = "AB4"
Public Const TICK_ITEMS_CELL   As String = "AB5"

' 歩み値ブロックの読み取り開始セル
'   配信状態の行（＝数式セル）を先頭にします。見出し行・終端記号は自動で
'   読み飛ばします。TICK_FORMULA_CELL と必ず一致させてください。
Public Const TICK_BLOCK_CELL As String = "AB4"

' 読み取る行数
'   配信状態行 + 見出し行 + データ + 終端記号 + 余裕。
'   TICK_SHOW_ROWS から自動算出するので、表示本数を変えても追従します。
Public Const TICK_BLOCK_ROWS As Long = TICK_SHOW_ROWS + 8

' ブロック内の列オフセット（左上セルから何列目か）
'   RssTickList の「取得項目」に並べた順と同じにしてください。
'   既定は 取得項目 = 時刻 / 出来高 / 約定値 の並び。
Public Const TB_OFF_TIME  As Long = 0     ' 時刻
Public Const TB_OFF_VOL   As Long = 1     ' 出来高
Public Const TB_OFF_PRICE As Long = 2     ' 約定値
Public Const TB_OFF_MARK  As Long = -1    ' ティック記号（取得項目に加えたら 3 に設定）

' ティック記号の RSS 項目名（TB_OFF_MARK を 0 以上にしたときだけ使用）
'   マーケットスピードの「取得項目一覧」で実際の名称を確認してから
'   TB_OFF_MARK を有効にしてください。名前が違うと配信が始まりません。
Public Const TICK_MARK_ITEM As String = "ティック"

' 歩み値の並び順
'   0 = 自動判定 / 1 = 新しい順（上が最新） / 2 = 古い順
'   RssTickList は常に新しい順なので 1 を固定にしています。
'   （自動判定は有効行が1本しかないときに誤るため、既定にはしません）
Public Const TICK_BLOCK_ORDER As Long = 1

' ① の方向をティック記号（↑↓）で判定する
'   TB_OFF_MARK が 0 以上のときだけ効きます。取得項目に記号が無い状態では
'   この設定にかかわらず前ティック比で判定します。
Public Const USE_TICK_MARK As Boolean = True

' 記録終了（JUDGE_END）と同時に自動で判定まで走らせる
Public Const AUTO_JUDGE_ON_STOP As Boolean = True

' 自動開始の予約時刻（ArmAutoRun で使用）
Public Const AUTO_ARM_TIME As String = "14:59:30"

' ブックを開いたときに自動開始を予約するか
Public Const AUTO_ARM_ON_OPEN As Boolean = True

' 時刻の取得元
'   True  = RSS の「現在値時刻」を使う
'   False = PC 時計（Now）を使う ※既定
'
'   注意：RSS の「現在値時刻」は hh:mm（秒なし）で返ってくることがあります。
'   その場合、全ティックの秒が :00 に丸められ、③約定速度が「同一分内の
'   約定回数」になって常に上限に張り付きます。TickTime は秒なしを検出したら
'   自動で PC 時計に切り替えて警告しますが、既定は False にしてあります。
'   秒付きの項目（RSS の項目一覧で秒まで返るもの）に差し替えられる場合は、
'   LIVE_TIME の数式を変更したうえで True にしてください。
Public Const USE_RSS_TIME As Boolean = False

'------------------------------------------------------------------
' シート名・レイアウト定数
'------------------------------------------------------------------
Public Const RESULT_SHEET As String = "Judge_Results"

' 銘柄シートの行
'   1行目はメモ用に空けてあります。マクロは絶対に書き込みません。
Public Const ROW_HEADER     As Long = 2   ' 見出し行
Public Const ROW_CODE       As Long = 3   ' B3=コード / C3=銘柄名称
Public Const TICK_FIRST_ROW As Long = 3   ' ティックログ開始行

' 銘柄シートの列
'   ※ 一括書き込みのため D→E→F（時刻・出来高・約定値）と
'      J→K（買気配・売気配）は隣接している必要があります。
'      並べ替えた場合は LayoutWarning() が起動時に知らせます。
Public Const COL_CODE  As Long = 2   ' B  コード
Public Const COL_NAME  As Long = 3   ' C  銘柄名称
Public Const COL_TIME  As Long = 4   ' D  時刻
Public Const COL_VOL   As Long = 5   ' E  出来高（ティック単体）
Public Const COL_PRICE As Long = 6   ' F  約定値
Public Const COL_W     As Long = 7   ' G  重み
Public Const COL_UP    As Long = 8   ' H  UpScore（累積・参考）
Public Const COL_DN    As Long = 9   ' I  DnScore（累積・参考）
Public Const COL_BID   As Long = 10  ' J  最良買気配値
Public Const COL_ASK   As Long = 11  ' K  最良売気配値
Public Const COL_DIR   As Long = 12  ' L  方向（↑↓→）
Public Const COL_SPD   As Long = 13  ' M  同一秒内の約定連番
Public Const COL_MARK  As Long = 14  ' N  ティック記号（取得項目に無ければ空欄）

' ティックログを配列で読むときの列番号（COL_TIME を 1 とした相対位置）
Public Const IX_TIME  As Long = 1
Public Const IX_VOL   As Long = COL_VOL - COL_TIME + 1
Public Const IX_PRICE As Long = COL_PRICE - COL_TIME + 1
Public Const IX_BID   As Long = COL_BID - COL_TIME + 1
Public Const IX_ASK   As Long = COL_ASK - COL_TIME + 1
Public Const IX_MARK  As Long = COL_MARK - COL_TIME + 1

' 銘柄シートの RSS ライブ取得セル（S列=ラベル / T列=値）
Public Const LIVE_PRICE  As String = "T2"   ' 現在値
Public Const LIVE_VOL    As String = "T3"   ' 出来高（当日累計）
Public Const LIVE_BID    As String = "T4"   ' 最良買気配値
Public Const LIVE_ASK    As String = "T5"   ' 最良売気配値
Public Const LIVE_TIME   As String = "T6"   ' 現在値時刻
Public Const LIVE_SPREAD As String = "T7"   ' スプレッド（数式）

' 銘柄シートの判定結果ブロック（O列=ラベル / P列=値）
Public Const RES_TOP  As String = "O2"
Public Const RES_ROWS As Long = 18          ' 結果ブロックの行数

' Judge_Results の状態表示セル（W列=ラベル / X列=値）
Public Const RES_STATUS As String = "X2"    ' 状態
Public Const RES_COUNT  As String = "X3"    ' 件数
Public Const RES_BREAK  As String = "X4"    ' 内訳

' Judge_Results の実行ログ（Y列=ラベル / Z列=値）
'   状態セルは判定完了で上書きされるため、事後に検証できるよう別枠にします。
Public Const LOG_TOP   As String = "Y2"
Public Const LOGR_START As Long = 0         ' 開始時刻
Public Const LOGR_MODE  As Long = 1         ' 取込モード
Public Const LOGR_FEED  As Long = 2         ' 配信状態
Public Const LOGR_LAST  As Long = 3         ' 最終ティック
Public Const LOGR_WARN  As Long = 4         ' 警告
Public Const LOGR_END   As Long = 5         ' 終了時刻
Public Const LOGR_ROWS  As Long = 6

' Judge_Results の内部管理セル（AA列）
'   Application.OnTime の予約時刻を保存します。VBA の状態がリセットされても
'   ここから復元して確実に解除できるようにするためです。消さないでください。
'   （1行目はメモ欄なので 2行目から使います）
Public Const SCHED_NOTE As String = "AA2"   ' 説明書き
Public Const SCHED_ARM  As String = "AA3"   ' 自動開始の予約時刻
Public Const SCHED_POLL As String = "AA4"   ' 次回ポーリングの予約時刻

' Judge_Results のサマリー列範囲
Public Const SUM_COL_FIRST As Long = 2      ' B
Public Const SUM_COL_LAST  As Long = 21     ' U
Public Const SUM_MAX_ROWS  As Long = 200

' Judge_Results のボタン設置基準セル
Public Const BTN_ANCHOR As String = "W5"
Public Const BTN_W As Double = 190
Public Const BTN_H As Double = 34

' 判定文言
Public Const JUDGE_BUY    As String = "買い"
Public Const JUDGE_SELL   As String = "売り"
Public Const JUDGE_FLAT   As String = "中立"
Public Const JUDGE_NODATA As String = "未計測"

'------------------------------------------------------------------
' 判定結果の型
'------------------------------------------------------------------
Public Type TickJudgeResult
    SheetName    As String
    Code         As String
    StockName    As String
    TickCount    As Long
    UpSeqMax     As Long
    DnSeqMax     As Long
    Direction    As Long      ' 1=引け上方向 / -1=引け下方向 / 0=中立
    DirText      As String
    Vol1         As Double
    Vol2         As Double
    Vol3         As Double
    Vol4         As Double
    VolDominant  As Boolean
    SpeedMax     As Long
    SpeedFast    As Boolean
    BuyTotal     As Double    ' 参考：重み付きスコア＋気配スコア
    SellTotal    As Double
    BestBid      As Double    ' 判定時点の最良買気配値
    BestAsk      As Double    ' 判定時点の最良売気配値
    Spread       As Double    ' 最良売 － 最良買
    Judge        As String    ' 買い / 売り / 中立 / 未計測
    Confidence   As String    ' 高（アルゴ/大口） / 通常 / －
End Type

'------------------------------------------------------------------
' Sleep API（診断用の負荷テストでのみ使用）
'------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'==================================================================
' 共通関数
'==================================================================

' 判定対象の銘柄シートを列挙する
'   Judge_Results 以外で、B3 に証券コードが入っているシートを対象にします。
Public Function TargetSheets() As Collection

    Dim c As New Collection
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> RESULT_SHEET Then
            If IsNumeric(ws.Cells(ROW_CODE, COL_CODE).Value) Then
                If Val(ws.Cells(ROW_CODE, COL_CODE).Value) > 0 Then c.Add ws
            End If
        End If
    Next ws

    Set TargetSheets = c
End Function

' 銘柄コードが重複しているシートを探す。無ければ空文字。
Public Function DuplicateCodeWarning() As String

    Dim shts As Collection
    Dim i As Long, j As Long
    Dim a As String, b As String
    Dim msg As String

    Set shts = TargetSheets()

    For i = 1 To shts.Count
        a = CStr(shts(i).Cells(ROW_CODE, COL_CODE).Value)
        For j = i + 1 To shts.Count
            b = CStr(shts(j).Cells(ROW_CODE, COL_CODE).Value)
            If a = b Then
                msg = msg & "　シート「" & shts(i).Name & "」と「" & shts(j).Name & "」" & _
                      "が同じコード " & a & vbCrLf
            End If
        Next j
    Next i

    If Len(msg) > 0 Then
        DuplicateCodeWarning = "同じ銘柄コードのシートがあります。" & vbCrLf & msg & _
                               "同じ銘柄を重複購読すると RSS の負荷が増え、結果も重複します。"
    End If
End Function

' 結果シートを取得（無ければ作る）
Public Function ResultSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESULT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RESULT_SHEET
    End If

    Set ResultSheet = ws
End Function

' ティックログの最終行
Public Function LastTickRow(ws As Worksheet) As Long
    LastTickRow = ws.Cells(ws.Rows.Count, COL_TIME).End(xlUp).Row
    If LastTickRow < TICK_FIRST_ROW Then LastTickRow = TICK_FIRST_ROW - 1
End Function

' セルの値を「0時からの経過秒」に変換する。変換できなければ -1
Public Function TickSec(ByVal v As Variant) As Long

    Dim d As Date

    On Error GoTo Bad
    If IsError(v) Then GoTo Bad
    If IsEmpty(v) Then GoTo Bad
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then GoTo Bad
        If InStr(CStr(v), ":") = 0 Then GoTo Bad     ' 終端記号 "--------" などを弾く
        d = CDate(CStr(v))
    ElseIf VarType(v) = vbDate Then
        d = CDate(v)
    ElseIf IsNumeric(v) Then
        d = CDate(CDbl(v))
    Else
        GoTo Bad
    End If

    TickSec = Hour(d) * 3600& + Minute(d) * 60& + Second(d)
    Exit Function
Bad:
    TickSec = -1
End Function

' 文字列時刻 → 経過秒
Public Function SecOfText(ByVal s As String) As Long
    SecOfText = TickSec(CDate(s))
End Function

' 経過秒 → "hh:mm"
Public Function SecHM(ByVal sec As Long) As String
    SecHM = Format$(TimeSerial(sec \ 3600, (sec \ 60) Mod 60, 0), "hh:mm")
End Function

' 経過秒 → "hh:mm:ss"
Public Function SecHMS(ByVal sec As Long) As String
    SecHMS = Format$(TimeSerial(sec \ 3600, (sec \ 60) Mod 60, sec Mod 60), "hh:mm:ss")
End Function

' 「何秒前か」を読みやすい日本語にする
'   42秒前 / 3分12秒前 / 5時間51分前
Public Function AgoText(ByVal secAgo As Long) As String

    Dim h As Long, m As Long, s As Long

    If secAgo < 0 Then
        AgoText = "不明"
        Exit Function
    End If

    h = secAgo \ 3600
    m = (secAgo \ 60) Mod 60
    s = secAgo Mod 60

    If h > 0 Then
        AgoText = h & "時間" & m & "分前"
    ElseIf m > 0 Then
        AgoText = m & "分" & s & "秒前"
    Else
        AgoText = s & "秒前"
    End If
End Function

' 複数行の警告文を箇条書きにする（2行目以降は字下げ）
Public Function Bullet(ByVal msg As String) As String

    Dim parts() As String
    Dim i As Long
    Dim out As String

    parts = Split(msg, vbCrLf)
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            out = out & IIf(i = LBound(parts), "・", "　") & parts(i) & vbCrLf
        End If
    Next i

    Bullet = out
End Function

' 現在の 0時からの経過秒
Public Function NowSec() As Long
    NowSec = Hour(Now) * 3600& + Minute(Now) * 60& + Second(Now)
End Function

' ②の1区間の長さ（秒）。判定窓を4等分します。
Public Function BucketSec() As Long
    BucketSec = (SecOfText(JUDGE_END) - SecOfText(JUDGE_START)) \ 4
    If BucketSec < 1 Then BucketSec = 1
End Function

' 判定窓の設定チェック。問題なければ空文字を返します。
Public Function WindowWarning() As String

    Dim st As Long, en As Long

    st = SecOfText(JUDGE_START)
    en = SecOfText(JUDGE_END)

    If en <= st Then
        WindowWarning = "JUDGE_START と JUDGE_END が逆、または同じです。"
    ElseIf en > SecOfText("15:25:00") Then
        WindowWarning = "JUDGE_END が 15:25:00 を超えています。" & vbCrLf & _
                        "ザラバは 15:25:00 で終わり、15:30:00 の板寄せ約定は1本で桁違いの" & vbCrLf & _
                        "出来高になるため、Vol4 と SpeedMax が壊れます。" & vbCrLf & _
                        "JUDGE_END を 15:25:00 以下にしてください。"
    ElseIf (en - st) < 240 Then
        WindowWarning = "判定窓が短すぎます（4分未満）。1区間が1分未満になります。"
    End If
End Function

' 列レイアウトの前提チェック（一括書き込みの隣接条件）
Public Function LayoutWarning() As String

    If COL_VOL <> COL_TIME + 1 Or COL_PRICE <> COL_TIME + 2 Then
        LayoutWarning = "COL_TIME / COL_VOL / COL_PRICE は隣接している必要があります。"
    ElseIf COL_ASK <> COL_BID + 1 Then
        LayoutWarning = "COL_BID と COL_ASK は隣接している必要があります。"
    ElseIf TICK_BLOCK_CELL <> TICK_FORMULA_CELL Then
        LayoutWarning = "TICK_BLOCK_CELL と TICK_FORMULA_CELL が一致していません。"
    End If
End Function

' 現在時刻（0～1 の日内小数）
Public Function NowTime() As Double
    NowTime = Now - Date
End Function

' 数値化（エラー値・空白は 0）
Public Function NumOrZero(ByVal v As Variant) As Double

    On Error GoTo Bad
    If IsError(v) Then GoTo Bad
    If IsEmpty(v) Then GoTo Bad
    If Not IsNumeric(v) Then GoTo Bad
    NumOrZero = CDbl(v)
    Exit Function
Bad:
    NumOrZero = 0
End Function

' 文字列化（エラー値・空白は空文字）
Public Function TextOrEmpty(ByVal v As Variant) As String

    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    TextOrEmpty = Trim$(CStr(v))
End Function

' 出来高から重みを求める（参考スコア用）
Public Function VolWeight(ByVal vol As Double) As Long
    Select Case vol
        Case Is >= 10000#: VolWeight = 5
        Case Is >= 5000#:  VolWeight = 3
        Case Is >= 1000#:  VolWeight = 2
        Case Is > 0#:      VolWeight = 1
        Case Else:         VolWeight = 0
    End Select
End Function

'==================================================================
' 記録中の多重実行ガード
'   記録ループは1秒ごとに完全に終了するため、その隙間で他のボタンが
'   押せてしまいます。ログを壊す操作・実行中のボタン自身を削除する操作を
'   ここで止めます。
'==================================================================
Public Function BlockedWhileLogging(ByVal actionName As String) As Boolean

    If Not gLogging Then Exit Function

    BlockedWhileLogging = True
    MsgBox "ティック記録中は「" & actionName & "」を実行できません。" & vbCrLf & vbCrLf & _
           "「③ 記録 停止 → 判定」で記録を止めてから実行してください。", _
           vbExclamation, "記録中"
End Function

'==================================================================
' Judge_Results の実行ログ
'==================================================================
Public Sub WriteRunLog(ByVal rowOffset As Long, ByVal v As Variant)

    On Error Resume Next
    ResultSheet().Range(LOG_TOP).Offset(rowOffset, 1).Value = v
End Sub

Public Sub ClearRunLog()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ResultSheet()
    ws.Range(LOG_TOP).Offset(0, 1).Resize(LOGR_ROWS, 1).ClearContents
End Sub

'==================================================================
' Excel を固めずに指定ミリ秒待つ（診断用の負荷テスト専用）
'   ※ 本番の記録ループでは使いません。ここで待つと RSS が更新できません。
'==================================================================
Public Sub WaitMs(ByVal ms As Long)

    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        Sleep 15
        If Timer < t0 Then Exit Do   ' 日付跨ぎ保険
    Loop While (Timer - t0) * 1000# < ms
End Sub
