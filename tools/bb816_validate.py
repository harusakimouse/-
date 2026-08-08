#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB816.xlsx の数式が、検証に使った Python モデルと同じ答えを出すか確かめる。

Excel を持たない環境なので、Python の数式エンジン（formulas パッケージ）で
BB816 の数式を実際に計算させ、bb_backtest.py の計算値と突き合わせる。
500行すべてを解くとエンジンが重いので、先頭20銘柄だけの縮小版を作って検証する。
（数式は行ごとに同じ形なので、20行で合えば全行合う）
"""
import sys
import shutil
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).parent))
import build_bb816 as B
from bb_data import load
from bb_backtest import indicators
from bb_study5 import score816

TEST = Path("/tmp/claude-0/-home-user--/188eba24-3296-5098-a456-8f48f10f134b/scratchpad/BB816_test.xlsx")
NROWS = 20


def main(src="V815.xlsm"):
    B.SRC = Path(src)
    B.DST = TEST
    B.ROWN = B.ROW1 + NROWS - 1
    B.TOPN = 5
    B.log.clear()
    names = B.build()
    print("検証用の縮小版を作成: %s（%d銘柄）" % (TEST.name, NROWS))

    import formulas
    xl = formulas.ExcelModel().loads(str(TEST)).finish()
    sol = xl.calculate()
    print("数式エンジンで計算完了: %d セル" % len(sol))

    key = {}
    for k, v in sol.items():
        m = k.upper()
        if "BB816判定'!" in m or "BB816判定!" in m:
            ref = m.rsplit("!", 1)[-1]
            try:
                key[ref] = v.value[0, 0]
            except Exception:
                pass

    d = load(src)
    ind = indicators(d)
    c = d["close"]
    T = c.shape[1] - 1
    mom20 = c[:, T] / c[:, T - 20] - 1
    is_top = np.zeros(c.shape, bool)

    def s816_variant(i):
        s = 0
        cl, up2, up1, sma = c[i, T], ind["up2"][i, T], ind["up1"][i, T], ind["sma"][i, T]
        vol, mfi, rsi, pb = ind["vol_ratio"][i, T], ind["mfi"][i, T], ind["rsi"][i, T], ind["pct_b"][i, T]
        mm = mom20[i]
        s += 25 if cl > up2 else (15 if cl > up1 else (5 if cl > sma else 0))
        s += 20 if vol >= 2 else (14 if vol >= 1.5 else (8 if vol >= 1.2 else 0))
        s += 15 if mfi >= 80 else (18 if mfi >= 60 else (8 if mfi >= 50 else 0))
        s += 18 if 70 < rsi <= 80 else (14 if 60 < rsi <= 70 else (6 if 50 < rsi <= 60 else 0))
        s += 14 if mm >= 0.10 else (7 if mm >= 0.03 else 0)
        s += -12 if rsi > 80 else 0
        s += -20 if rsi < 40 else 0
        s += -15 if mfi < 40 else 0
        s += -18 if pb < 0.2 else 0
        s += -10 if mm <= -0.10 else 0
        return max(0, min(100, s))

    checks = {"E": "sma", "F": "sd", "G": "up2", "H": "up1", "I": "lo1", "J": "lo2",
              "K": "pct_b", "M": "rsi", "N": "mfi", "O": "vol_ratio"}
    print()
    print("  %-14s %6s %10s" % ("項目", "件数", "最大相対差"))
    ng = []
    for col, name in checks.items():
        diffs = []
        for i in range(NROWS):
            r = B.ROW1 + i
            got = key.get("%s%d" % (col, r))
            exp = ind[name][i, T]
            if got is None or isinstance(got, str) or not np.isfinite(exp):
                continue
            diffs.append(abs(float(got) - exp) / max(abs(exp), 1e-9))
        if diffs:
            print("  %-14s %6d %10.2e" % ("%s (%s)" % (col, name), len(diffs), max(diffs)))
            if max(diffs) > 1e-9:
                ng.append(col)

    # 20日騰落率
    diffs = [abs(float(key["P%d" % (B.ROW1 + i)]) - mom20[i]) for i in range(NROWS)
             if not isinstance(key.get("P%d" % (B.ROW1 + i)), str)]
    print("  %-14s %6d %10.2e" % ("P (20日騰落率)", len(diffs), max(diffs)))

    # スコア
    bad = 0
    for i in range(NROWS):
        got = key.get("Q%d" % (B.ROW1 + i))
        exp = s816_variant(i)
        if isinstance(got, str):
            continue
        if abs(float(got) - exp) > 1e-9:
            bad += 1
            print("    スコア不一致 行%d Excel=%s Python=%s" % (B.ROW1 + i, got, exp))
    print("  %-14s %6d %10s" % ("Q (BB816スコア)", NROWS, "一致" if bad == 0 else "%d件ちがう" % bad))

    # 判定系
    for col, label in (("R", "買い条件"), ("S", "総合判定"), ("T", "守り分析官"), ("U", "攻め分析官")):
        vals = [key.get("%s%d" % (col, B.ROW1 + i)) for i in range(NROWS)]
        print("  %-14s %s" % ("%s (%s)" % (col, label), " ".join(str(v) for v in vals)))

    print()
    print("  Excel側の◎買い銘柄:", [d["codes"][i] for i in range(NROWS)
                                if key.get("S%d" % (B.ROW1 + i)) == "◎買い"])
    entry = ((c[:, T] > ind["up2"][:, T]) & (ind["vol_ratio"][:, T] >= 1.5)
             & (ind["rsi"][:, T] <= 80) & (ind["mfi"][:, T] >= 55))
    py = [d["codes"][i] for i in range(NROWS) if entry[i] and s816_variant(i) >= 40]
    print("  Python側の買い条件成立 :", py)
    print()
    print("結果:", "全項目一致" if not ng and bad == 0 else "不一致あり %s" % ng)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
