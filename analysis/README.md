# 買いタイミング（価格）分析

打ち出のこづち V1 (V8.17) の OHLCV データ 300銘柄×250営業日を用いて、
「買い候補にいくらの指値を置くべきか」を検証したスクリプト一式。

実行順:

    python3 extract.py   # xlsm -> raw.pkl   （XLSM_PATH を自分の環境に合わせる）
    python3 build.py     # raw.pkl -> data.npz
    python3 signal.py    # data.npz -> feat.npz（Mod_買抽出v13 のスコア再現）
    python3 analyze1.py analyze3.py kmodel.py sim.py portfolio.py robust.py adaptive.py final_stats.py cand_price.py

必要パッケージ: numpy, pandas, openpyxl, oletools

結論と全数値は `../docs/買いタイミング分析_2026-08.md` を参照。
全出力ログは `results_2026-08-14.txt`。
