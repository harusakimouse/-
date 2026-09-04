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
'   実行 : Alt+F8 →
'     平均足_取込して抽出   … 今すぐ取り込んで候補を出す（手動）
'     平均足_自動開始       … 毎日15:35に自動で取込→抽出
'     平均足_自動停止       … 自動をやめる
'==================================================================

'-------------- 調整するのはここだけ --------------
Public Const HA_SRC_NAME As String = "OHLCV　明日売買　ボリンジャー送信.xlsm"
Public Const HA_SRC_FIND As String = "OHLCV*.xlsm"   '名前が少し違っても拾う
Public Const HA_AUTO_TIME As String = "15:35:00"     '自動実行の時刻
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

Public Sub 平均足_取込して抽出()
    If HA_Import(True) Then 平均足_買い抽出
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
    If HA_Import(False) Then 平均足_買い抽出
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

    Dim myPath As String
    myPath = ThisWorkbook.Path
    If myPath = "" Then
        If showMsg Then MsgBox "このブックを一度保存してから実行してください。", vbExclamation
        Exit Function
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
        Exit Function
    End If
    fPath = myPath & "\" & fName

    '--- すでに開いているか調べる ---
    Dim wbSrc As Workbook, opened As Boolean
    On Error Resume Next
    Set wbSrc = Workbooks(fName)
    On Error GoTo 0

    Dim oldEvents As Boolean, oldUpd As Boolean, oldSec As Long
    oldEvents = Application.EnableEvents
    oldUpd = Application.ScreenUpdating
    oldSec = Application.AutomationSecurity

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable   '元ブックのマクロは動かさない

    If wbSrc Is Nothing Then
        On Error Resume Next
        Set wbSrc = Workbooks.Open(FileName:=fPath, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        opened = True
    End If

    If wbSrc Is Nothing Then
        Application.AutomationSecurity = oldSec
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldUpd
        If showMsg Then MsgBox "「" & fName & "」を開けませんでした。", vbExclamation
        Exit Function
    End If

    '--- 5つのシートを写す ---
    Dim names As Variant
    names = Array("始値", "高値", "安値", "終値", "出来高")

    Dim i As Long, done As Long, miss As String
    Dim wsS As Worksheet, wsD As Worksheet
    For i = 0 To UBound(names)
        Set wsS = Nothing: Set wsD = Nothing
        On Error Resume Next
        Set wsS = wbSrc.Worksheets(CStr(names(i)))
        Set wsD = ThisWorkbook.Worksheets(CStr(names(i)))
        On Error GoTo 0
        If wsS Is Nothing Or wsD Is Nothing Then
            miss = miss & CStr(names(i)) & " "
        Else
            If HA_CopyOne(wsS, wsD) Then done = done + 1 Else miss = miss & CStr(names(i)) & " "
        End If
    Next i

    '--- 後始末（元ブックは保存しない）---
    If opened Then
        Application.DisplayAlerts = False
        wbSrc.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    Set wbSrc = Nothing

    Application.AutomationSecurity = oldSec
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldUpd

    If done = 0 Then
        If showMsg Then
            MsgBox "元ブックから同じ名前のシートが見つかりませんでした。" & vbCrLf & _
                   "元ブックのシート名：" & HA_SheetNames(fPath), vbExclamation
        End If
        Exit Function
    End If

    If miss <> "" And showMsg Then
        MsgBox done & "シートを取り込みました。" & vbCrLf & _
               "取り込めなかったシート：" & miss, vbExclamation
    End If

    HA_Import = True
End Function

'1シート分を写す（値だけ・C列D列のRSS式は残す）
Private Function HA_CopyOne(ByVal wsS As Worksheet, ByVal wsD As Worksheet) As Boolean

    HA_CopyOne = False

    Dim lastRow As Long, lastCol As Long, c As Long
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastRow > SRC_MAXROW Then lastRow = SRC_MAXROW
    If lastRow < 6 Then Exit Function

    lastCol = 5
    For c = 6 To SRC_MAXCOL
        If IsNumeric(wsS.Cells(3, c).Value) Then
            If Val(wsS.Cells(3, c).Value) > 40000 Then
                lastCol = c
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next c
    If lastCol < 20 Then Exit Function      '並びが違う＝写さない

    On Error Resume Next
    wsD.Unprotect
    On Error GoTo 0

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
