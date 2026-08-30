# vba/ — VBAモジュールの取り込み方

## ★文字コードの注意（重要）

VBE の **ファイル → ファイルのインポート** は、`.bas` を
**Shift-JIS（CP932）** として読み込みます。

このフォルダ直下の `.bas` は GitHub 上で読めるように **UTF-8** で保存しているため、
そのままインポートすると**日本語コメントと文字列が文字化けし、
「コンパイル エラー: 構文エラー」になります**。

### 対処: 次のどちらかを使ってください

#### 方法A: コピー＆ペースト（最も確実・推奨）

クリップボードは Unicode でやり取りされるため、文字化けが起きません。

1. GitHub 上で `.bas` を開き、コード全体をコピー
2. VBE で **挿入 → 標準モジュール**
3. 貼り付ける
4. **先頭の `Attribute VB_Name = "..."` の行だけ削除する**
   （インポート時専用の行で、直接貼るとエラーになります）
5. プロパティウィンドウ（`F4`）でモジュール名をファイル名と同じに変更

#### 方法B: `vba/sjis/` の Shift-JIS 版をインポート

`vba/sjis/` に **CP932 + CRLF** に変換済みの同内容ファイルを置いています。
こちらは **ファイル → ファイルのインポート** でそのまま読み込めます。

> `vba/` 直下（UTF-8）と `vba/sjis/`（CP932）は中身が同一です。
> 編集するときは UTF-8 側を直し、下記コマンドで sjis 側を再生成してください。
>
> ```bash
> for f in Fix_OnTime_Cleanup Diag_AutoStart; do
>   iconv -f UTF-8 -t CP932 "vba/$f.bas" | sed 's/$/\r/' > "vba/sjis/$f.bas"
> done
> ```

---

## モジュール一覧

| ファイル | 用途 | 対象ブック |
|---|---|---|
| `Fix_OnTime_Cleanup.bas` | `Application.OnTime` 予約を台帳管理し、閉じるときに全解除する | **`打ち出のこづち_0615MOD整理 改善 .xlsm` 用** |
| `Diag_AutoStart.bas` | 自動起動の原因を切り分ける診断レポートを出力（読み取りのみ） | どのブックでも可 |

### ⚠ `Fix_OnTime_Cleanup` を `0945_baibai.xlsm` に入れないでください

`0945_baibai.xlsm` には `Application.OnTime` が存在しないため、
このモジュールは**何もしません**（`CancelAllTasks` は常に空振りします）。

既に取り込んでしまった場合は削除してください:

1. VBE左のツリーで `Fix_OnTime_Cleanup` を右クリック
2. **「Fix_OnTime_Cleanup の解放」**
3. 「エクスポートしますか？」→ **いいえ**

あわせて、`ThisWorkbook` に `Call CancelAllTasks` を追記していた場合は
**その行（または `Workbook_BeforeClose` ごと）も削除**してください。
残すと「Sub または Function が定義されていません」エラーになります。

なお `Workbook_BeforeClose` を**二重に定義するとコンパイルエラー**
（名前が適切ではありません）になり、**ブック内の全マクロが動かなくなります**。
1モジュールにつき1つだけです。
