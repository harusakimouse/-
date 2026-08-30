Attribute VB_Name = "Diag_AutoStart"
Option Explicit
'=============================================================================
' Diag_AutoStart : ブックが勝手に起動する原因の切り分け診断
'
' 使い方:
'   1. 症状の出ているブック（0945_baibai.xlsm）で Alt+F11 → VBE を開く
'   2. ファイル → ファイルのインポート で本モジュールを読み込む
'   3. Alt+F8 → 「自動起動診断」を実行
'   4. デスクトップに 自動起動診断_yyyymmdd_hhnnss.txt が出力される
'
' 何も書き換えないので、そのまま実行して安全です。
'=============================================================================

Public Sub 自動起動診断()
    Dim rpt As String
    rpt = BuildReport()

    Dim path As String
    path = DesktopPath() & "\自動起動診断_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim fnum As Integer
    fnum = FreeFile
    On Error GoTo WriteFailed
    Open path For Output As #fnum
    Print #fnum, rpt
    Close #fnum
    On Error GoTo 0

    MsgBox "診断レポートを出力しました。" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _
           "このテキストファイルの中身を貼り付けてください。", _
           vbInformation, "自動起動診断"
    Exit Sub

WriteFailed:
    On Error Resume Next
    Close #fnum
    On Error GoTo 0
    ' 書き出せない場合はイミディエイトウィンドウへ（Ctrl+G で表示）
    Debug.Print rpt
    MsgBox "ファイル出力に失敗しました。" & vbCrLf & _
           "Ctrl+G でイミディエイトウィンドウを開いて内容をコピーしてください。", _
           vbExclamation, "自動起動診断"
End Sub

'=============================================================================
Private Function BuildReport() As String
    Dim s As String
    Const NL As String = vbCrLf

    s = "==================================================" & NL
    s = s & " 自動起動 診断レポート" & NL
    s = s & " 生成: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & NL
    s = s & "==================================================" & NL & NL

    '--- 1. このブック ---
    s = s & "[1] 診断対象ブック" & NL
    s = s & "  名前      : " & ThisWorkbook.Name & NL
    s = s & "  フルパス  : " & ThisWorkbook.FullName & NL
    On Error Resume Next
    s = s & "  保存日時  : " & FileDateTime(ThisWorkbook.FullName) & NL
    On Error GoTo 0
    s = s & "  Excel版   : " & Application.Version & " (" & Application.Build & ")" & NL & NL

    '--- 2. 起動フォルダ ★最重要 ---
    s = s & "[2] Excel 起動フォルダ  ★勝手に開く原因No.1" & NL
    s = s & "  XLSTART      : " & Application.StartupPath & NL
    s = s & DumpFolder(Application.StartupPath, "    ")
    s = s & "  代替起動folder: " & IIf(Len(Application.AltStartupPath) = 0, _
                                       "(未設定 = 正常)", Application.AltStartupPath) & NL
    If Len(Application.AltStartupPath) > 0 Then
        s = s & DumpFolder(Application.AltStartupPath, "    ")
        s = s & "    ※ここが設定されていると中のファイルが毎回自動で開きます" & NL
    End If
    s = s & NL

    '--- 3. 現在開いているブック ---
    s = s & "[3] 現在開いているブック" & NL
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        s = s & "  - " & wb.Name & NL
        s = s & "      " & wb.path & NL
    Next wb
    s = s & NL

    '--- 4. アドイン ---
    s = s & "[4] 有効なアドイン" & NL
    Dim ai As AddIn
    Dim addinCount As Long
    For Each ai In Application.AddIns
        If ai.Installed Then
            addinCount = addinCount + 1
            s = s & "  - " & ai.Name & "  [" & ai.FullName & "]" & NL
        End If
    Next ai
    If addinCount = 0 Then s = s & "  (なし)" & NL
    s = s & NL

    Dim ai2 As COMAddIn
    s = s & "[5] 有効な COM アドイン" & NL
    Dim comCount As Long
    On Error Resume Next
    For Each ai2 In Application.COMAddIns
        If ai2.Connect Then
            comCount = comCount + 1
            s = s & "  - " & ai2.Description & "  [" & ai2.progID & "]" & NL
        End If
    Next ai2
    On Error GoTo 0
    If comCount = 0 Then s = s & "  (なし / 取得不可)" & NL
    s = s & NL

    '--- 6. VBAプロジェクトの中身 ---
    s = s & "[6] VBA モジュール一覧と OnTime 出現箇所" & NL
    s = s & DumpVBProject("  ")
    s = s & NL

    '--- 7. 外部リンク ---
    s = s & "[7] 外部リンク (他ブックへの参照)" & NL
    Dim lnk As Variant
    On Error Resume Next
    lnk = ThisWorkbook.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If IsEmpty(lnk) Or IsNull(lnk) Then
        s = s & "  (なし)" & NL
    Else
        Dim i As Long
        For i = LBound(lnk) To UBound(lnk)
            s = s & "  - " & lnk(i) & NL
        Next i
    End If
    s = s & NL

    s = s & "==================================================" & NL
    s = s & " レポート終了" & NL
    s = s & "==================================================" & NL

    BuildReport = s
End Function

'-----------------------------------------------------------------------------
' フォルダ内のファイルを列挙
'-----------------------------------------------------------------------------
Private Function DumpFolder(ByVal folderPath As String, ByVal indent As String) As String
    Const NL As String = vbCrLf
    Dim s As String

    If Len(folderPath) = 0 Then
        DumpFolder = indent & "(パスが空)" & NL
        Exit Function
    End If

    Dim f As String
    On Error GoTo NoFolder
    f = Dir(folderPath & "\*.*", vbNormal)
    On Error GoTo 0

    If Len(f) = 0 Then
        DumpFolder = indent & "(空 = 正常)" & NL
        Exit Function
    End If

    Do While Len(f) > 0
        s = s & indent & "→ " & f & "   ★このファイルは Excel 起動時に自動で開かれます" & NL
        f = Dir
    Loop

    DumpFolder = s
    Exit Function

NoFolder:
    DumpFolder = indent & "(フォルダが存在しない = 正常)" & NL
End Function

'-----------------------------------------------------------------------------
' VBAプロジェクトの全モジュールを走査し OnTime / 自動実行を探す
'   ※「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が
'     OFF の場合はここでエラーになるため、その旨を出力する。
'-----------------------------------------------------------------------------
Private Function DumpVBProject(ByVal indent As String) As String
    Const NL As String = vbCrLf
    Dim s As String

    Dim vbc As Object
    Dim cm As Object

    On Error GoTo NoAccess

    Dim total As Long
    total = ThisWorkbook.VBProject.VBComponents.Count

    s = s & indent & "モジュール数: " & total & NL & NL

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Set cm = vbc.CodeModule
        s = s & indent & "- " & vbc.Name & "  (" & cm.CountOfLines & " 行)" & NL

        Dim keywords As Variant
        keywords = Array("Application.OnTime", "OnTime", "Workbook_Open", _
                         "Auto_Open", "Workbook_BeforeClose", "Auto_Close", _
                         "Workbooks.Open", "Shell", "CreateObject")

        Dim k As Long
        For k = LBound(keywords) To UBound(keywords)
            Dim hits As String
            hits = FindInModule(cm, CStr(keywords(k)), indent & "      ")
            If Len(hits) > 0 Then
                s = s & indent & "    ◆ " & keywords(k) & NL & hits
            End If
        Next k
    Next vbc

    DumpVBProject = s
    Exit Function

NoAccess:
    s = "" 
    s = s & indent & "!! VBAプロジェクトを読み取れませんでした。" & NL
    s = s & indent & "   ファイル → オプション → トラスト センター →" & NL
    s = s & indent & "   トラスト センターの設定 → マクロの設定 →" & NL
    s = s & indent & "   「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」" & NL
    s = s & indent & "   にチェックを入れて Excel を再起動し、もう一度実行してください。" & NL
    s = s & indent & "   (エラー: " & Err.Number & " " & Err.Description & ")" & NL
    DumpVBProject = s
End Function

'-----------------------------------------------------------------------------
' 1モジュール内のキーワード出現行を返す
'-----------------------------------------------------------------------------
Private Function FindInModule(ByVal cm As Object, ByVal keyword As String, _
                              ByVal indent As String) As String
    Const NL As String = vbCrLf
    Dim s As String
    Dim i As Long

    For i = 1 To cm.CountOfLines
        Dim ln As String
        ln = cm.Lines(i, 1)

        ' コメント行は除外（先頭が ' の行）
        If Left(LTrim(ln), 1) <> "'" Then
            If InStr(1, ln, keyword, vbTextCompare) > 0 Then
                s = s & indent & i & ": " & Trim(ln) & NL
            End If
        End If
    Next i

    FindInModule = s
End Function

'-----------------------------------------------------------------------------
Private Function DesktopPath() As String
    Dim p As String
    On Error Resume Next
    p = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If Len(p) = 0 Then p = Environ$("USERPROFILE") & "\Desktop"
    If Len(Dir(p, vbDirectory)) = 0 Then p = Environ$("TEMP")
    DesktopPath = p
End Function
