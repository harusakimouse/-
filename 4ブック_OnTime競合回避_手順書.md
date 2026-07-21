# 4ブックの OnTime 時間競合 回避手順（V0720 が止まる問題）

## 結論（何が起きているか）

`Application.OnTime` は **Excel 全体（アプリ単位）** で動き、発火時刻になると
**その瞬間にアクティブなブック** から指定マクロ名を探して実行します。
マクロ名に **ブック名を付けていない（未修飾）** と、別ブックがアクティブな時に
発火した予約は、そのブックにマクロが無くて

> 実行時エラー1004「マクロ '○○' を実行できません」

になり、**そのブックのタイマー連鎖が途切れて自動更新が止まります**。

### なぜ V0720 が影響を受けるのか

- 調査の結果、**同名マクロの衝突ではありません**
  （V1 と V0720 で共通する Sub 名は `Workbook_Open` だけで、これは OnTime 対象外）。
- 原因は純粋に「未修飾の OnTime」です。
- **V1** は `CheckSchedule` を **5〜10秒ごと** に予約し続けるため、V1 が
  ほぼ常にアクティブ／稼働中になります。
- 一方 **V0720** のタイマー（`RSS監視タイマー`＝数十秒ごと、`売抽出タイマー実行`＝5分ごと、
  `日次更新_自動`／`売抽出更新`／`株価抽出`＝1日1回）は発火間隔が長く、
  発火した瞬間に V1（や400番）がアクティブだと **V0720 側が誤爆して止まる**。
- V1 は毎秒予約し直すので一度ミスしても自己回復しますが、V0720 の長間隔タイマーは
  一度連鎖が切れると復帰しません。だから **V0720 だけが止まって見える** わけです。

### 対策

**すべての `Application.OnTime` のマクロ名を「'ブック名'!マクロ名」に修飾** します。
こうすると予約は必ず **そのブック自身** のマクロを呼び、誰がアクティブでも競合しません。

書き換えは機械的で、`"マクロ名"` を
`"'" & ThisWorkbook.Name & "'!マクロ名"` に置き換えるだけです。

---

## ✅ 400番ブック（bf20f900 / 6af77788）… 対応不要

2つとも前回の修正版が入っており、`MyProc("時刻チェック")`＝
`"'" & ThisWorkbook.Name & "'!時刻チェック"` で **すでに修飾済み**です。
そのまま使えます。

---

## 【要修正①】V0720（9574def8）

VBAエディタ（`Alt`+`F11`）で下記モジュールを開き、該当行を差し替えてください。
**行まるごと** を置き換えるのが安全です（前後の空白・時刻の部分はそのまま）。

### ● ThisWorkbook モジュール（3か所）

```vba
' 変更前
Application.OnTime 発火時刻, "日次更新_自動"
' 変更後
Application.OnTime 発火時刻, "'" & ThisWorkbook.Name & "'!日次更新_自動"
```
```vba
' 変更前
Application.OnTime 発火時刻 + TimeValue("00:00:30"), "売抽出更新"
' 変更後
Application.OnTime 発火時刻 + TimeValue("00:00:30"), "'" & ThisWorkbook.Name & "'!売抽出更新"
```
```vba
' 変更前
Application.OnTime 発火時刻 + TimeValue("00:01:00"), "株価抽出"
' 変更後
Application.OnTime 発火時刻 + TimeValue("00:01:00"), "'" & ThisWorkbook.Name & "'!株価抽出"
```

### ● Mod_RSS接続状態 モジュール（3か所）

同じ行が2か所（開始時と再予約時）あります。**両方**同じように直します。
```vba
' 変更前（2か所とも）
Application.OnTime NextMonitorTime, "RSS監視タイマー"
' 変更後（2か所とも）
Application.OnTime NextMonitorTime, "'" & ThisWorkbook.Name & "'!RSS監視タイマー"
```
取り消し（停止）の3行ブロックも修飾します。
```vba
' 変更前
Application.OnTime EarliestTime:=NextMonitorTime, _
                   Procedure:="RSS監視タイマー", _
                   Schedule:=False
' 変更後
Application.OnTime EarliestTime:=NextMonitorTime, _
                   Procedure:="'" & ThisWorkbook.Name & "'!RSS監視タイマー", _
                   Schedule:=False
```

### ● Mod_SellExtrac モジュール（3か所）

予約が2か所（開始時・再予約時）、取り消しが1か所です。
```vba
' 変更前（予約：2か所とも）
Application.OnTime NextSellRunTime, "売抽出タイマー実行"
' 変更後（予約：2か所とも）
Application.OnTime NextSellRunTime, "'" & ThisWorkbook.Name & "'!売抽出タイマー実行"
```
```vba
' 変更前（取り消し）
Application.OnTime NextSellRunTime, "売抽出タイマー実行", , False
' 変更後（取り消し）
Application.OnTime NextSellRunTime, "'" & ThisWorkbook.Name & "'!売抽出タイマー実行", , False
```

> ※ 取り消し行も予約行と同じ「'ブック名'!…」にそろえてください。
> 予約と取り消しでマクロ名の文字列が一致していないと取り消せなくなります。

---

## 【要修正②】V1（9c79df65）

### ● ThisWorkbook モジュール（2か所）
```vba
' 変更前
Application.OnTime Now + TimeSerial(0, 0, 5), "CheckSchedule"
' 変更後
Application.OnTime Now + TimeSerial(0, 0, 5), "'" & ThisWorkbook.Name & "'!CheckSchedule"
```
```vba
' 変更前（取り消し）
Application.OnTime Now + TimeSerial(0, 0, 10), "CheckSchedule", , False
' 変更後（取り消し）
Application.OnTime Now + TimeSerial(0, 0, 10), "'" & ThisWorkbook.Name & "'!CheckSchedule", , False
```

### ● Module1 モジュール（2か所）
```vba
' 変更前
Application.OnTime Now + TimeSerial(0, 0, 10), "CheckSchedule"
' 変更後
Application.OnTime Now + TimeSerial(0, 0, 10), "'" & ThisWorkbook.Name & "'!CheckSchedule"
```
```vba
' 変更前
Application.OnTime Now + TimeSerial(0, 0, 60), "CheckSchedule"
' 変更後
Application.OnTime Now + TimeSerial(0, 0, 60), "'" & ThisWorkbook.Name & "'!CheckSchedule"
```

---

## 作業のまとめ

| ブック | 対応 | 直す箇所 |
|---|---|---|
| bf20f900 (400番) | 不要 | 修正済み |
| 6af77788 (400番) | 不要 | 修正済み |
| **9574def8 (V0720)** | **要修正** | ThisWorkbook×3・Mod_RSS接続状態×3・Mod_SellExtrac×3 |
| **9c79df65 (V1)** | **要修正** | ThisWorkbook×2・Module1×2 |

1. `Alt`+`F11` でVBAエディタを開く
2. 上記の各モジュールで「変更前」の行を探し、「変更後」に差し替える
3. `Ctrl`+`S` で xlsm のまま上書き保存
4. 4ブックを開き直せば、同時起動しても競合しません

## 補足（任意の追加改善）
V1 の停止処理は取り消し時刻に `Now + TimeSerial(0,0,10)` を使っており、
予約した時刻と一致しないため実際には取り消せていません（従来からの潜在バグ）。
今回の競合回避には影響しませんが、確実に止めたい場合は、予約した時刻を
モジュール変数に保持して同じ時刻で取り消す方式（400番の `NextRunTime` と同じやり方）に
そろえると万全です。ご希望あれば対応します。
