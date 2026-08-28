Attribute VB_Name = "Module_Config"
Option Explicit

'==================================================================
' Module_Config : 共通設定・共通関数・共通型
'   ティック判定ロジック（15:00〜15:20専用）
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
Public Const JUDGE_START As String = "15:00:00"
Public Const JUDGE_END   As String = "15:20:00"

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
Public Const VOL_RATIO As Double = 1.5

' ③ 約定速度 : 同一秒内の約定回数がこの値以上なら「アルゴ／大口が本気」
Public Const SPEED_MIN As Long = 3

'------------------------------------------------------------------
' ティック記録の設定
'------------------------------------------------------------------
' RSS のポーリング間隔（ミリ秒）
'   Application.OnTime は 1 秒刻みしか扱えないため Sleep + DoEvents で刻みます。
'   200ms = 1秒あたり最大5ティックまで分離できる → SpeedMax >= 3 の判定が成立します。
'   重い場合は 250〜500 に上げてください（そのぶん SpeedMax の分解能は落ちます）。
Public Const POLL_MS As Long = 200

' 記録終了（15:20）と同時に自動で判定まで走らせる
Public Const AUTO_JUDGE_ON_STOP As Boolean = True

' 自動開始の予約時刻（ArmAutoRun で使用）
Public Const AUTO_ARM_TIME As String = "14:59:30"

' 時刻の取得元
'   True  = RSS の「現在値時刻」を使う（約定時刻そのもの／推奨）
'   False = PC 時計（Now）を使う
Public Const USE_RSS_TIME As Boolean = True

'------------------------------------------------------------------
' シート名・レイアウト定数
'------------------------------------------------------------------
Public Const RESULT_SHEET As String = "Judge_Results"

' 銘柄シートの行
Public Const ROW_HEADER     As Long = 2   ' 見出し行
Public Const ROW_CODE       As Long = 3   ' B3=コード / C3=銘柄名称
Public Const TICK_FIRST_ROW As Long = 3   ' ティックログ開始行

' 銘柄シートの列
Public Const COL_CODE  As Long = 2   ' B  コード
Public Const COL_NAME  As Long = 3   ' C  銘柄名称
Public Const COL_TIME  As Long = 4   ' D  時刻
Public Const COL_PRICE As Long = 5   ' E  約定値
Public Const COL_VOL   As Long = 6   ' F  出来高（ティック単体）
Public Const COL_W     As Long = 7   ' G  重み
Public Const COL_UP    As Long = 8   ' H  UpScore（累積・参考）
Public Const COL_DN    As Long = 9   ' I  DnScore（累積・参考）
Public Const COL_BID   As Long = 10  ' J  最良買気配値
Public Const COL_ASK   As Long = 11  ' K  最良売気配値
Public Const COL_DIR   As Long = 12  ' L  方向（↑↓→）
Public Const COL_SPD   As Long = 13  ' M  同一秒内の約定連番

' 銘柄シートの RSS ライブ取得セル（S列=ラベル / T列=値）
Public Const LIVE_PRICE As String = "T1"   ' 現在値
Public Const LIVE_VOL   As String = "T2"   ' 出来高（当日累計）
Public Const LIVE_BID   As String = "T3"   ' 最良買気配値
Public Const LIVE_ASK   As String = "T4"   ' 最良売気配値
Public Const LIVE_TIME  As String = "T5"   ' 現在値時刻

' 銘柄シートの判定結果ブロック（O列=ラベル / P列=値）
Public Const RES_TOP As String = "O1"

' Judge_Results のボタン設置基準セル
Public Const BTN_ANCHOR As String = "T5"
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

' 現在時刻（0〜1 の日内小数）
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
