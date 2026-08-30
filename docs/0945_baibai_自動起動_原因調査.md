# `0945_baibai.xlsm` が勝手に立ち上がる — 原因調査レポート

調査対象: `打ち出のこづち_0615MOD整理 改善 .xlsm`（`0945_baibai.xlsm` と同系統のブック）
調査方法: xlsm を展開し `xl/vbaProject.bin` から VBA 全モジュール（36個）を抽出して解析

---

## 結論

**`Application.OnTime` で予約したタイマーが、ブックを閉じるときに解除されていないため。**

`Application.OnTime` の予約先は「ブック」ではなく **Excel アプリケーション本体** です。
予約が残ったままブックを閉じると、予約時刻になった Excel は指定されたマクロを実行するために
**そのブックを自動的に開き直します**。これが「勝手に立ち上がる」の正体です。

ウイルス・マルウェアの類ではありません。

---

## 根拠となるコード

### 1. `ThisWorkbook` — 起動時に3件のタイマーを予約

```vba
Private Sub Workbook_Open()
    Call 日次更新_本日分

    Dim 発火時刻 As Date
    発火時刻 = Int(Now()) + TimeValue("15:31:00")

    If Now() < 発火時刻 Then
        Application.OnTime 発火時刻, "日次更新_自動"                          ' 15:31:00
        Application.OnTime 発火時刻 + TimeValue("00:00:30"), "売抽出更新"     ' 15:31:30
        Application.OnTime 発火時刻 + TimeValue("00:01:00"), "株価抽出"       ' 15:32:00
    End If

    Call 損切りアラート
End Sub
```

→ **`Workbook_BeforeClose` / `Auto_Close` がプロジェクト内に一つも存在しない**（全36モジュールを検索して0件）。
つまり予約を解除する処理がどこにもありません。

**症状の典型例**: 午前中にブックを開いて閉じる → 15:31 になると勝手に起動する。

### 2. `Mod_SellExtrac` — 5分ごとに自己再予約（無限ループ）

```vba
Public Sub 売抽出タイマー実行()
    Call 売り候補抽出
    NextSellRunTime = Now + TimeValue("00:05:00")
    Application.OnTime NextSellRunTime, "売抽出タイマー実行"   ' 自分自身を再予約
End Sub
```

一度「売抽出自動更新開始」を実行すると、明示的に「売抽出自動更新停止」を押さない限り
**5分間隔で永久に再予約され続けます**。閉じても5分後に再起動します。

さらに停止処理には欠陥があります:

```vba
Public Sub 売抽出自動更新停止()
    Application.OnTime NextSellRunTime, "売抽出タイマー実行", , False
End Sub
```

`NextSellRunTime` はモジュールレベル変数なので、**ブックを閉じた時点で値が失われます**（リセットされる）。
再度開いてから「停止」を押しても `NextSellRunTime` は 0 になっており、予約時刻が一致せず
**解除に失敗します**。`Application.OnTime` の解除は「時刻」と「プロシージャ名」の完全一致が必須です。

### 3. `Mod_RSS接続状態` — 30秒ごとに自己再予約（最も影響大）

```vba
Private Const MONITOR_INTERVAL_SEC As Long = 30

Public Sub RSS監視タイマー()
    If Not Monitoring Then Exit Sub
    UpdateStatusSilent
    NextMonitorTime = Now + TimeSerial(0, 0, MONITOR_INTERVAL_SEC)
    Application.OnTime NextMonitorTime, "RSS監視タイマー"   ' 自分自身を再予約
End Sub
```

「RSS監視開始」を実行していた場合、**閉じてから最短30秒で勝手に開きます**。
`Monitoring` もモジュール変数なので、再オープン後は `False` に戻って `Exit Sub` しますが、
**「ブックが開かれてしまう」という事象自体は防げません**（Exit Sub はブックが開いた後に評価されるため）。
`RSS監視停止` も上記2と同じ理由で、閉じた後は解除に失敗します。

---

## OnTime 予約 一覧

| モジュール | 予約先マクロ | 発火 | 自己再予約 | 閉じる時に解除 |
|---|---|---|---|---|
| `ThisWorkbook` | `日次更新_自動` | 15:31:00 | なし | **なし** |
| `ThisWorkbook` | `売抽出更新` | 15:31:30 | なし | **なし** |
| `ThisWorkbook` | `株価抽出` | 15:32:00 | なし | **なし** |
| `Mod_SellExtrac` | `売抽出タイマー実行` | 5分毎 | **あり** | **なし**（停止も失敗する） |
| `Mod_RSS接続状態` | `RSS監視タイマー` | 30秒毎 | **あり** | **なし**（停止も失敗する） |

---

## 今すぐ止める応急処置

1. タスクマネージャーで **`EXCEL.EXE` プロセスを全て終了**（予約は Excel プロセス上にあるため、
   プロセスを落とせば予約も消えます）。
2. その後 `0945_baibai.xlsm` を **マクロを無効にして開く**
   （開く前に Shift キーを押しっぱなしにする、またはセキュリティの警告で「コンテンツの有効化」を押さない）。
   → `Workbook_Open` が走らないので新しい予約が入りません。
3. 下記の恒久対策コードを入れてから、通常どおり開き直す。

---

## 恒久対策

`vba/Fix_OnTime_Cleanup.bas`（本コミットに同梱）を標準モジュールとして取り込み、
`ThisWorkbook` モジュールに `Workbook_BeforeClose` を追加してください。

要点は3つです。

1. **予約した時刻を必ず記録する** — 解除には時刻の完全一致が必要
2. **記録はモジュール変数ではなくシート/名前定義など永続領域に置く** — ブックを閉じても消えない
3. **`Workbook_BeforeClose` で全予約を解除する**

---

## 補足: OnTime 以外に確認すべき箇所

VBA 以外で同じ症状を起こす設定です。上記対策後もまだ起動する場合はこちらを確認してください。

- **XLSTART フォルダ**: `%APPDATA%\Microsoft\Excel\XLSTART\` に `0945_baibai.xlsm` の
  実体やショートカットが置かれていないか
- **Excel の起動時読み込みフォルダ**: ファイル → オプション → 詳細設定 → 全般 →
  「起動時にすべてのファイルを開くフォルダー」が設定されていないか
- **Windows スタートアップ**: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\`
- **タスクスケジューラ**: 「タスク スケジューラ ライブラリ」に該当ファイルを開くタスクがないか
- **楽天 MarketSpeed II の RSS 連携**: RSS 関数を含むブックが外部から呼び出されるケース
  （本ブックは `xl/volatileDependencies.xml` に RSS 参照あり）

なお、本ブックには外部リンク（`xl/externalLinks/`）は存在せず、他ブックを開く
`Workbooks.Open` / `Shell` / `CreateObject("Excel.Application")` も **VBA 内に一切ありません**。
他ファイルから呼ばれているのではなく、**このブック自身の OnTime 予約による自己再オープン**です。
