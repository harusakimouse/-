Attribute VB_Name = "Mod_HeikinAshi_Auto"
Option Explicit
'==================================================================
' OHLCVデータの自動取込　＋　15:35の自動実行　v1.0
'
'   取りに行く先 : このブックと同じフォルダーの
'                  「OHLCV　明日売買　ボリンジャー送信.xlsm」
'   ※元ブックは【読み取り専用】で開き、【保存しません】。
'     元ブックのマクロも動かさない設定で開きます。
'
'   ※本体モジュール(Mod_HeikinAshi)は半角名 HA_Buy で呼びます。
'     日本語名をまたいで呼ばないのでコンパイルエラーになりません。
'
'   実行 : Alt+F8 →
'     平均足_取込して抽出   … 今すぐ取り込んで候補を出す（手動）
'     平均足_自動開始       … 毎日15:35に自動で取込→抽出
'     平均足_自動停止       … 自動をやめる
'==================================================================

'-------------- 調整するのはここだけ --------------
Public Const HA_SRC_NAME As String = "OHLCV　明日売買　ボリンジャー送信.xlsm"
Public Const HA_SRC_FIND As String = "OHLCV*.xlsm"   '名前が少し違っても拾う
Public Const HA_AUTO_TIME As String = "15:35:00"     '自動実行の時刻
Public Const HA_SHEET_PW As String = ""             'シート保護のパスワード（かかっている場合だけ入れる）
'--------------------------------------------------

Private Const SRC_MAXROW As Long = 520
Private Const SRC_MAXCOL As Long = 300

Private nextRun As Date

'OnTimeで呼ぶ手続き名（ブック名付きで確実に呼ぶ）
Private Function HA_ProcName() As String
    HA_ProcName = "'" & ThisWorkbook.Name & "'!平均足_自動実行"
End Function

'==================== 手動で使う ====================
Public Sub 平均足_データ取込()
    If HA_Import(True) Then
        MsgBox "OHLCVデータを取り込みました。", vbInformation
    End If
End Sub

'保護がかかっていて書き込めない時に、5枚のデータシートを作り直します
Public Sub 平均足_データシート作り直し()

    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")

    Dim msg As String
    msg = "始値・高値・安値・終値・出来高　の5枚を、保護のかかっていない新しいシートに作り直します。" & vbCrLf & vbCrLf & _
          "・中のデータは、この直後に元ブックから取り込み直すので消えません。" & vbCrLf & _
          "・シート保護のパスワードは、以後いっさい不要になります。" & vbCrLf & _
          "・「平均足」「平均足記録」「メニュー」などの他のシートは触りません。" & vbCrLf & vbCrLf & _
          "実行してよろしいですか？"
    If MsgBox(msg, vbYesNo + vbQuestion, "データシートの作り直し") <> vbYes Then Exit Sub

    Dim oldAlert As Boolean, oldUpd As Boolean
    oldAlert = Application.DisplayAlerts
    oldUpd = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim i As Long, ws As Worksheet, nm As String
    For i = 0 To UBound(shNames)
        nm = CStr(shNames(i))
        Set ws = HA_FindSheet(ThisWorkbook, nm)
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
        ws.Cells(1, 1).Value = nm & "シート"
        ws.Cells(2, 1).Value = "コード"
        ws.Cells(2, 2).Value = "銘 柄 名"
        ws.Cells(3, 1).Value = "日  付"
        ws.Cells(1, 1).Font.Bold = True
        ws.Range("A2:B2").Font.Bold = True
        ws.Cells(3, 4).NumberFormat = "yyyy/mm/dd"
        ws.Range("D3:GZ3").NumberFormat = "yyyy/mm/dd"
        ws.Columns("A").ColumnWidth = 8
        ws.Columns("B").ColumnWidth = 16
    Next i

    Application.ScreenUpdating = oldUpd
    Application.DisplayAlerts = oldAlert

    If MsgBox("5枚を作り直しました。" & vbCrLf & vbCrLf & _
              "続けてデータを取り込みますか？", vbYesNo + vbQuestion) = vbYes Then
        平均足_取込して抽出
    End If
End Sub

Public Sub 平均足_取込して抽出()
    If HA_Import(True) Then HA_Buy
End Sub

'==================== 自動実行 ====================
Public Sub 平均足_自動開始()
    HA_SetTimer True
End Sub

'ブックを開いた時に静かにセットする用
Public Sub 平均足_自動開始_起動時()
    HA_SetTimer False
End Sub

Public Sub 平均足_自動停止()
    On Error Resume Next
    If nextRun > 0 Then Application.OnTime nextRun, HA_ProcName(), , False
    On Error GoTo 0
    nextRun = 0
    MsgBox "自動実行を止めました。", vbInformation
End Sub

Private Sub HA_SetTimer(ByVal showMsg As Boolean)
    On Error Resume Next
    If nextRun > 0 Then Application.OnTime nextRun, HA_ProcName(), , False
    On Error GoTo 0

    nextRun = HA_NextTime()
    Application.OnTime nextRun, HA_ProcName(), , True

    If showMsg Then
        MsgBox "自動実行をセットしました。" & vbCrLf & vbCrLf & _
               Format(nextRun, "yyyy/mm/dd (aaa) hh:nn") & " に" & vbCrLf & _
               "　OHLCV取込 → 買い候補の抽出　を行います。" & vbCrLf & vbCrLf & _
               "※Excelとこのブックを開いたままにしてください。", vbInformation
    End If
End Sub

Public Sub 平均足_自動実行()
    HA_SILENT = True
    On Error Resume Next
    If HA_Import(False) Then HA_Buy
    On Error GoTo 0
    HA_SILENT = False

    '次の営業日に予約し直す
    nextRun = HA_NextTime()
    On Error Resume Next
    Application.OnTime nextRun, HA_ProcName(), , True
    On Error GoTo 0
End Sub

'次の実行時刻（土日は飛ばす）
Private Function HA_NextTime() As Date
    Dim d As Date
    d = Int(Now) + TimeValue(HA_AUTO_TIME)
    If Now >= d Then d = d + 1
    Do While Weekday(d, vbMonday) >= 6      '6=土 7=日
        d = d + 1
    Loop
    HA_NextTime = d
End Function

'==================== 取込本体 ====================
Private Function HA_Import(ByVal showMsg As Boolean) As Boolean

    HA_Import = False

    '--- 今のExcelの設定を覚えておく ---
    Dim oldEvents As Boolean, oldUpd As Boolean, oldAlert As Boolean
    Dim oldSec As Long, oldCalc As Long, oldAsk As Boolean
    oldEvents = Application.EnableEvents
    oldUpd = Application.ScreenUpdating
    oldAlert = Application.DisplayAlerts
    oldSec = Application.AutomationSecurity
    oldCalc = Application.Calculation
    oldAsk = Application.AskToUpdateLinks

    Dim wbSrc As Workbook, opened As Boolean
    Dim done As Long, detail As String

    On Error GoTo CleanUp

    Dim myPath As String
    myPath = ThisWorkbook.Path
    If myPath = "" Then
        If showMsg Then MsgBox "このブックを一度保存してから実行してください。", vbExclamation
        GoTo CleanUp
    End If

    '--- 元ファイルを探す ---
    Dim fName As String, fPath As String
    fName = HA_SRC_NAME
    If Dir(myPath & "\" & fName) = "" Then
        fName = Dir(myPath & "\" & HA_SRC_FIND)
    End If
    If fName = "" Then
        If showMsg Then
            MsgBox "同じフォルダーに「" & HA_SRC_NAME & "」が見つかりません。" & vbCrLf & _
                   "フォルダー：" & myPath, vbExclamation
        End If
        GoTo CleanUp
    End If
    fPath = myPath & "\" & fName

    '--- ここが大事：重い再計算を止めてから開く ---
    Application.StatusBar = "OHLCVデータを読み込んでいます…"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual          'RSSなどの再計算を止める
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    '--- すでに開いているか調べる ---
    On Error Resume Next
    Set wbSrc = Workbooks(fName)
    On Error GoTo CleanUp

    If wbSrc Is Nothing Then
        On Error Resume Next
        Set wbSrc = Workbooks.Open(FileName:=fPath, ReadOnly:=True, UpdateLinks:=0, _
                                   IgnoreReadOnlyRecommended:=True, Notify:=False)
        On Error GoTo CleanUp
        opened = True
    End If

    If wbSrc Is Nothing Then
        If showMsg Then MsgBox "「" & fName & "」を開けませんでした。", vbExclamation
        GoTo CleanUp
    End If

    '--- 5つのシートを写す ---
    Dim shNames As Variant
    shNames = Array("始値", "高値", "安値", "終値", "出来高")

    Dim i As Long, info As String
    Dim wsS As Worksheet, wsD As Worksheet
    For i = 0 To UBound(shNames)
        Application.StatusBar = "取込中… " & CStr(shNames(i))
        Set wsS = HA_FindSheet(wbSrc, CStr(shNames(i)))
        Set wsD = HA_FindSheet(ThisWorkbook, CStr(shNames(i)))
        info = ""
        If wsS Is Nothing Then
            detail = detail & CStr(shNames(i)) & " : 元ブックに無し" & vbCrLf
        ElseIf wsD Is Nothing Then
            detail = detail & CStr(shNames(i)) & " : このブックに無し" & vbCrLf
        ElseIf HA_CopyOne(wsS, wsD, info) Then
            done = done + 1
            detail = detail & CStr(shNames(i)) & " : OK  " & info & vbCrLf
        Else
            detail = detail & CStr(shNames(i)) & " : 失敗  " & info & vbCrLf
        End If
    Next i

    '--- 元ブックを閉じる（保存しない）---
    If opened Then
        Application.StatusBar = "元ブックを閉じています…"
        wbSrc.Close SaveChanges:=False
    End If
    Set wbSrc = Nothing

CleanUp:
    Dim errNo As Long, errTx As String
    errNo = Err.Number: errTx = Err.Description
    On Error Resume Next
    If opened And Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing
    Application.AutomationSecurity = oldSec
    Application.AskToUpdateLinks = oldAsk
    Application.DisplayAlerts = oldAlert
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldUpd
    Application.StatusBar = False
    On Error GoTo 0

    If errNo <> 0 Then
        If showMsg Then MsgBox "取込中にエラーが出ました。" & vbCrLf & _
                              "番号 " & errNo & "：" & errTx, vbExclamation
        Exit Function
    End If

    If done = 0 Then
        If showMsg And detail <> "" Then
            MsgBox "取り込めませんでした。" & vbCrLf & vbCrLf & detail, vbExclamation
        End If
        Exit Function
    End If

    If done < 5 And showMsg Then
        MsgBox done & "シートだけ取り込みました。" & vbCrLf & vbCrLf & detail, vbExclamation
    End If

    HA_Import = True
End Function

'シートを名前で探す（前後や途中の空白は無視する）
Private Function HA_FindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet, key As String
    key = HA_Norm(nm)
    For Each ws In wb.Worksheets
        If HA_Norm(ws.Name) = key Then
            Set HA_FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function HA_Norm(ByVal s As String) As String
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    HA_Norm = Trim$(s)
End Function

'1シート分を写す（値だけ・C列D列のRSS式は残す）
Private Function HA_CopyOne(ByVal wsS As Worksheet, ByVal wsD As Worksheet, _
                            ByRef info As String) As Boolean

    HA_CopyOne = False

    Dim lastRow As Long, r2 As Long, lastCol As Long, c As Long, dv As Double
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    r2 = wsS.Cells(wsS.Rows.Count, 2).End(xlUp).Row
    If r2 > lastRow Then lastRow = r2
    If lastRow > SRC_MAXROW Then lastRow = SRC_MAXROW

    '日付の行から最終列を探す（途中が空でも最後まで見る）
    lastCol = 0
    For c = 4 To SRC_MAXCOL
        dv = 0
        If IsNumeric(wsS.Cells(3, c).Value2) Then dv = CDbl(wsS.Cells(3, c).Value2)
        If dv > 40000 And dv < 80000 Then lastCol = c
    Next c
    If lastCol = 0 Then
        '日付行が無いときは使用範囲の右端を使う
        lastCol = wsS.UsedRange.Column + wsS.UsedRange.Columns.Count - 1
        If lastCol > SRC_MAXCOL Then lastCol = SRC_MAXCOL
    End If

    info = "行" & lastRow & " 列" & lastCol
    If lastRow < 6 Then
        info = info & "（銘柄行が見つからない）"
        Exit Function
    End If
    If lastCol < 10 Then
        info = info & "（日付の列が見つからない）"
        Exit Function
    End If

    'シート保護を外す（パスワード無しの Unprotect はダイアログが出るので必ず指定する）
    If wsD.ProtectContents Then
        On Error Resume Next
        wsD.Unprotect Password:=HA_SHEET_PW
        On Error GoTo 0
    End If
    If wsD.ProtectContents Then
        info = info & "（このブックの" & wsD.Name & "シートに保護がかかっています。" & _
                      "モジュール先頭の HA_SHEET_PW にパスワードを入れてください）"
        Exit Function
    End If

    On Error GoTo CopyErr
    '日付の行（D列～）
    wsD.Range(wsD.Cells(3, 4), wsD.Cells(3, lastCol)).Value = _
        wsS.Range(wsS.Cells(3, 4), wsS.Cells(3, lastCol)).Value
    'コードと銘柄名
    wsD.Range(wsD.Cells(5, 1), wsD.Cells(lastRow, 2)).Value = _
        wsS.Range(wsS.Cells(5, 1), wsS.Cells(lastRow, 2)).Value
    '株価・出来高（E列～）
    wsD.Range(wsD.Cells(5, 5), wsD.Cells(lastRow, lastCol)).Value = _
        wsS.Range(wsS.Cells(5, 5), wsS.Cells(lastRow, lastCol)).Value

    HA_CopyOne = True
    Exit Function

CopyErr:
    info = info & "（書き込みエラー " & Err.Number & "）"
    HA_CopyOne = False
End Function

'元ブックのシート名を並べて返す（原因調べ用）
Private Function HA_SheetNames(ByVal fPath As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String, opened As Boolean
    On Error Resume Next
    Set wb = Workbooks(Dir(fPath))
    If wb Is Nothing Then
        Application.EnableEvents = False
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set wb = Workbooks.Open(FileName:=fPath, ReadOnly:=True, UpdateLinks:=0)
        opened = True
    End If
    On Error GoTo 0
    If wb Is Nothing Then
        HA_SheetNames = "（開けませんでした）"
        Exit Function
    End If
    For Each ws In wb.Worksheets
        s = s & ws.Name & " / "
    Next ws
    If opened Then
        Application.DisplayAlerts = False
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.AutomationSecurity = msoAutomationSecurityByUI
    End If
    HA_SheetNames = s
End Function
