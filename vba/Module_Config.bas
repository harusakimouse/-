Attribute VB_Name = "Module_Config"
Option Explicit

'==================================================================
' Module_Config : 共通設定・共通関数・共通型
'   ティック判定ロジック（15:00～15:20専用）
'------------------------------------------------------------------
'   運用タイムライン
'     15:00:00  ティック記録 開始（自動 or ボタン）
'     15:20:00  ティック記録 終了 → 自動で判定実行
'     15:20:03  Judge_Results に 買い / 売り / 中立 が出る
'     15:25     プレクロージング（中立は板を見て最終判断）
'     15:30     板寄せ終了
'==================================================================

'------------------------------------------------------------------
' 判定対象時間帯
'------------------------------------------------------------------
'   ザラバは 15:25:00 で終了し、以降はクロージング・オークション（気配のみ）。
'   15:30:00 の板寄せ約定は1本で桁違いの出来高になるため、絶対に窓へ入れないこと。
'   → JUDGE_END は 15:25:00 を超えてはいけません。
'
'   既定は 15:00:00～15:24:00。
'   ザラバ終盤（15:22～15:25）の大口集中を Vol4 で捉えつつ、
'   15:24:02 頃に判定が出るので 15:30 の板寄せまで約6分の発注時間が残ります。
Public Const JUDGE_START As String = "15:00:00"
Public Const JUDGE_END   As String = "15:24:00"

'------------------------------------------------------------------
' 判定パラメータ（ここだけ触れば感度を変えられます）
'------------------------------------------------------------------
' ① 方向判定 : 連続上昇／連続下落が何回続いたら「方向あり」とみなすか
Public Const SEQ_MIN As Long = 5

' ① 変わらず（前ティックと同値）のティックで連続をリセットするか
'    True  = 同値で連続が途切れる（仕様どおり／推奨）
'    False = 同値は無視して連続を継続する（緩め）
Public Const RESET_ON_FLAT As Boolean = True

' ② 出来高初動判定 : Vol4 > Vol3 × この倍率
'   区間は JUDGE_START～JUDGE_END を4等分します（窓を変えれば区間幅も自動で追従）。
Public Const VOL_RATIO As Double = 1.5

' ③ 約定速度 : 同一秒内の約定回数がこの値以上なら「アルゴ／大口が本気」
Public Const SPEED_MIN As Long = 3

'------------------------------------------------------------------
' ティックの取込方式
'------------------------------------------------------------------
'   0 = 自動判定（既定）
'       記録開始時に歩み値ブロックを1回読み、ティックが読めれば 1、読めなければ 2 を選びます。
'       歩み値の数式がまだ入っていなくても、そのまま本番が回ります。
'   1 = 歩み値（TICK）ブロック追従
'       RSS の歩み値をシート上のブロックに出しておき、増えた行だけを追記します。
'       約定1本ずつのデータなので SpeedMax（③約定速度）が正確に出ます。
'   2 = 現在値・出来高ポーリング（フォールバック）
'       出来高（当日累計）の増分を1ティックとみなします。
'       この方式では SpeedMax の上限が POLL_MS で決まります（200ms なら最大5）。
Public Const LOG_MODE As Long = 0

' ポーリング間隔（ミリ秒）
'   モード1 : 歩み値ブロックが履歴を保持するので 300～1000 で十分（分解能は落ちません）
'   モード2 : 200 で 1秒あたり最大5ティックまで分離（ここが SpeedMax の上限になります）
'   自動判定でどちらに転んでもよいよう、既定は 200 にしてあります。
Public Const POLL_MS As Long = 200

'------------------------------------------------------------------
' 歩み値（TICK）ブロックの設定  ※ LOG_MODE = 0（自動）/ 1 のとき使用
'------------------------------------------------------------------
' 歩み値ブロックの読み取り開始セル
'   RssTickList は「表示開始セル」に配信状態、その1行下に見出し、2行下からデータを
'   書き出します。ここは配信状態の行を含む一番上を指定してください。
'   （見出し行・配信状態行は自動で読み飛ばします）
Public Const TICK_BLOCK_CELL As String = "AB3"

' 読み取る行数
'   表示本数(300) + 配信状態行 + 見出し行 + 余裕。多い分は空行を読み飛ばすだけです。
Public Const TICK_BLOCK_ROWS As Long = 310

' 追従用アンカーに使う行数
'   「15:24:59 / 100 / 3130」のように同一秒・同値・同数量のティックが連続することが
'   あるため、直前1本だけでは位置を取り違えます。直近この本数の並びで位置合わせします。
Public Const ANCHOR_DEPTH As Long = 3

' ブロック内の列オフセット（左上セルから何列目か）
'   RssTickList の「取得項目」に並べた順と同じにしてください。
'   既定は 取得項目 = 時刻 / 出来高 / 約定値 の並び。
Public Const TB_OFF_TIME  As Long = 0     ' 時刻
Public Const TB_OFF_VOL   As Long = 1     ' 出来高
Public Const TB_OFF_PRICE As Long = 2     ' 約定値
Public Const TB_OFF_MARK  As Long = -1    ' ティック記号（取得項目に加えたら 3 などに設定）

' 歩み値の並び順
'   0 = 自動判定（既定） / 1 = 新しい順（上が最新） / 2 = 古い順
Public Const TICK_BLOCK_ORDER As Long = 0

' ① の方向をティック記号（↑↓）で判定する
'   True  = 記号があればそれを使う（取引所の呼値そのものなので正確）
'           記号が空欄の行は前ティック比で判定します
'   False = 常に前ティック比で判定
Public Const USE_TICK_MARK As Boolean = True

' 記録終了（15:20）と同時に自動で判定まで走らせる
Public Const AUTO_JUDGE_ON_STOP As Boolean = True

' 自動開始の予約時刻（ArmAutoRun で使用）
Public Const AUTO_ARM_TIME As String = "14:59:30"

' ブックを開いたときに自動開始を予約するか
'   True  = 開くだけで予約完了。ボタンを押す必要がありません（推奨）
'           予約時刻を過ぎていても、判定窓の終了前なら自動で記録を始めます。
'   False = 「自動開始を予約」ボタンを自分で押す
Public Const AUTO_ARM_ON_OPEN As Boolean = True

' 時刻の取得元
'   True  = RSS の「現在値時刻」を使う（約定時刻そのもの／推奨）
'   False = PC 時計（Now）を使う
Public Const USE_RSS_TIME As Boolean = True

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
Public Const COL_CODE  As Long = 2   ' B  コード
Public Const COL_NAME  As Long = 3   ' C  銘柄名称
'   D～F の並びは RssTickList の取得項目（時刻 / 出来高 / 約定値）に合わせています
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
'   上の COL_* を並べ替えても自動で追従します。
Public Const IX_TIME  As Long = 1
Public Const IX_VOL   As Long = COL_VOL - COL_TIME + 1
Public Const IX_PRICE As Long = COL_PRICE - COL_TIME + 1
Public Const IX_BID   As Long = COL_BID - COL_TIME + 1
Public Const IX_ASK   As Long = COL_ASK - COL_TIME + 1
Public Const IX_MARK  As Long = COL_MARK - COL_TIME + 1

' 銘柄シートの RSS ライブ取得セル（S列=ラベル / T列=値）
'   1行目はメモ欄なので 2行目から並べます。
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

' Judge_Results のボタン設置基準セル
Public Const BTN_ANCHOR As String = "W5"
Public Const BTN_W As Double = 190
Public Const BTN_H As Double = 34

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
    Judge        As String    ' 買い / 売り / 中立
    Confidence   As String    ' 高（アルゴ/大口） / 通常
End Type

'------------------------------------------------------------------
' Sleep API（ポーリング間隔用）
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
'   ※ シート名が全角「１」でも半角「1」でも取りこぼしません。
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
    If IsEmpty(v) Then GoTo Bad
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then GoTo Bad
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

' Excel を固めずに指定ミリ秒待つ
Public Sub WaitMs(ByVal ms As Long)
    Dim t0 As Single
    t0 = Timer
    Do
        DoEvents
        Sleep 15
        If Timer < t0 Then Exit Do   ' 日付跨ぎ保険
    Loop While (Timer - t0) * 1000# < ms
End Sub
