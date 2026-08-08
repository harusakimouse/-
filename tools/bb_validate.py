#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Python で組み直した BB 判定が、V815 の BBスクリーニング（Excel 数式の計算結果）と
一致するかを突き合わせる。ここが合っていないと、この先の検証は全部無意味になる。

比較するのは最新営業日（終値シートのE列 = 2026-08-06）の各銘柄。
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
COLS = {"E": "sma", "F": "sd", "G": "up2", "H": "up1", "I": "lo1", "J": "lo2",
        "K": "pct_b", "L": "width", "M": "width_min41", "Q": "rsi", "R": "mfi",
        "V": "vol_ratio"}
TEXT_COLS = {"N": "squeeze", "O": "expand", "P": "brk", "S": "diver", "T": "walk",
             "U": "dbl", "W": "signal", "X": "score"}


def read_screen(path):
    z = zipfile.ZipFile(path)
    shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
    root = ET.fromstring(z.read("xl/worksheets/sheet24.xml"))
    out = {}
    for row in root.iter(NS + "row"):
        r = int(row.get("r"))
        if r < 6:
            continue
        for c in row:
            ref = c.get("r")
            col = re.sub(r"\d", "", ref)
            v = c.find(NS + "v")
            if v is None or v.text is None:
                continue
            t = c.get("t")
            val = shared[int(v.text)] if t == "s" else v.text
            out[(r, col)] = val
    z.close()
    return out


def main(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    ind["signal"], ind["score"] = sig, sc
    xl = read_screen(path)
    T = d["close"].shape[1] - 1          # 最新営業日

    print("突き合わせ: BBスクリーニング(Excelキャッシュ値) vs Python 再現")
    print("  対象日 %s / 銘柄 %d" % (d["dates"][-1], len(d["codes"])))
    print()
    print("  %-12s %6s %8s %10s" % ("項目", "件数", "一致率", "最大差"))

    ng_detail = []
    for col, key in COLS.items():
        got, exp = [], []
        for i in range(len(d["codes"])):
            v = xl.get((6 + i, col))
            if v is None:
                continue
            try:
                v = float(v)
            except ValueError:
                continue
            mine = ind[key][i, T]
            if not np.isfinite(mine):
                continue
            got.append(mine)
            exp.append(v)
        got, exp = np.array(got), np.array(exp)
        if len(got) == 0:
            print("  %-12s     -" % col)
            continue
        scale = np.maximum(np.abs(exp), 1e-9)
        rel = np.abs(got - exp) / scale
        ok = rel < 1e-6
        print("  %-12s %6d %7.1f%% %10.2e" % ("%s (%s)" % (col, key), len(got), 100 * ok.mean(), rel.max()))
        if ok.mean() < 1:
            bad = np.where(~ok)[0][:3]
            ng_detail.append((col, [(int(b), exp[b], got[b]) for b in bad]))

    for col, key in TEXT_COLS.items():
        got, exp = [], []
        for i in range(len(d["codes"])):
            v = xl.get((6 + i, col))
            if v is None:
                continue
            mine = ind[key][i, T]
            if key == "score":
                try:
                    if abs(float(v) - float(mine)) < 1e-6:
                        got.append(1)
                    else:
                        got.append(0)
                        exp.append((i, v, mine))
                except ValueError:
                    pass
                continue
            if key == "expand":
                mine = "拡大↑" if mine else "縮小↓"
            got.append(1 if str(mine) == str(v) else 0)
            if str(mine) != str(v):
                exp.append((i, v, mine))
        if not got:
            continue
        print("  %-12s %6d %7.1f%%" % ("%s (%s)" % (col, key), len(got), 100 * np.mean(got)))
        if np.mean(got) < 1:
            ng_detail.append((col, exp[:3]))

    if ng_detail:
        print("\n  不一致の例")
        for col, items in ng_detail:
            for it in items:
                print("    %s 行%s Excel=%s Python=%s" % (col, it[0] + 6, it[1], it[2]))
    else:
        print("\n  全項目一致")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
