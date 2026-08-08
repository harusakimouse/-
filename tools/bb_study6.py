#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定の深掘り（第6部）― 2人の分析官の対立点を数字で裁定する

対立点
  (a) リスク分析官の指摘: 第4部【16】で「現行スコア上位」が総+143.7%と最良だった。
      大標本（第5部【22】）では現行スコア上位3の超過は -0.29% で負けている。どちらが本当か。
  (b) 損切り: リスク分析官 -8% / トレーダー -10%
  (c) 保有:   リスク分析官 10日で必ず切る / トレーダー 15日・SMA20割れまで引っ張る

裁定の方法
  ・(a) 開始日をずらして何十通りも回し、成績のバラつきを見る（1回の結果は運で動く）
  ・(b)(c) 実際に採用する形（枠8・1日3銘柄・スコア降順）で全組み合わせを回す

出力は tools/out/bb_study6.txt
"""
import sys
from pathlib import Path

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores
from bb_sim import simulate
from bb_study5 import score816

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
lines = []


def p(s=""):
    print(s)
    lines.append(s)


def window(mask, lo, hi):
    m = np.zeros(mask.shape, bool)
    m[:, lo:hi] = mask[:, lo:hi]
    return m


def main(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    c = d["close"]
    N, T = c.shape
    dates = d["dates"]
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])
    live = np.isfinite(ind["sma"]) & np.isfinite(c)
    s816 = score816(ind, c, is_top)

    entry816 = ((c > ind["up2"]) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80)
                & (ind["mfi"] >= 55) & ~is_top & live)
    entry_cur = live & (sc >= 40)

    p("=" * 112)
    p("ボリンジャーバンド判定 深掘りレポート（第6部・分析官の対立点の裁定）")
    p("=" * 112)
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【26】対立点(a) 「現行スコア上位が実運用で最良だった」は本物か")
    p("-" * 112)
    p("  開始日を10日ずつずらして13通り走らせ、成績のバラつきを見る（枠5・1日2銘柄・10日保有・損切なし）")
    p("")
    p("  %-20s %10s %10s %10s %10s %10s" % ("選び方", "中央値%", "最小%", "最大%", "中央DD%", "最悪DD%"))
    for label, m, rank in [("現行スコア", entry_cur, sc), ("BB816スコア", entry816, s816)]:
        tot, dds = [], []
        for start in range(0, 130, 10):
            r = simulate(d, window(m, start, T), rank, hold=10, slots=5, topk=2)
            tot.append(r["total"]); dds.append(r["maxdd"])
        tot, dds = np.array(tot), np.array(dds)
        p("  %-20s %+11.1f %+10.1f %+10.1f %+10.1f %+10.1f" %
          (label, np.median(tot), tot.min(), tot.max(), np.median(dds), dds.min()))
    p("")
    p("  同じことを枠8・1日3銘柄で")
    p("  %-20s %10s %10s %10s %10s %10s" % ("選び方", "中央値%", "最小%", "最大%", "中央DD%", "最悪DD%"))
    for label, m, rank in [("現行スコア", entry_cur, sc), ("BB816スコア", entry816, s816)]:
        tot, dds = [], []
        for start in range(0, 130, 10):
            r = simulate(d, window(m, start, T), rank, hold=10, slots=8, topk=3)
            tot.append(r["total"]); dds.append(r["maxdd"])
        tot, dds = np.array(tot), np.array(dds)
        p("  %-20s %+11.1f %+10.1f %+10.1f %+10.1f %+10.1f" %
          (label, np.median(tot), tot.min(), tot.max(), np.median(dds), dds.min()))
    p("")
    p("  前半125日 / 後半125日 に切って比べる（枠8・1日3銘柄）")
    p("  %-20s %22s %22s" % ("選び方", "前半", "後半"))
    for label, m, rank in [("現行スコア", entry_cur, sc), ("BB816スコア", entry816, s816)]:
        out = []
        for lo, hi in ((0, T // 2), (T // 2, T)):
            r = simulate(d, window(m, lo, hi), rank, hold=10, slots=8, topk=3)
            out.append("%+8.1f%% DD%+7.1f%% n=%3d" % (r["total"], r["maxdd"], r["n"]))
        p("  %-20s %22s %22s" % (label, out[0], out[1]))
    p("")
    p("  ※ 第5部【22】の大標本（毎日上位3を機械的に買う、n=639）では")
    p("     現行スコア: 10日平均+1.52% 超過-0.29%  /  BB816スコア: +4.88% 超過+3.07%")
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【27】対立点(b)(c) 損切りと保有期間の総当たり（枠8・1日3銘柄・BB816スコア降順）")
    p("-" * 112)
    p("  %-30s %10s %9s %8s %6s %7s %8s %8s" %
      ("出口の組み合わせ", "総リターン%", "最大DD%", "Sharpe", "回数", "勝率%", "平均%", "最悪%"))
    grid = []
    for hold in (10, 15):
        for stop in (None, 8, 10):
            for line in (None, "sma"):
                name = "%2d日" % hold
                name += " / 損切%s" % ("なし" if stop is None else "-%d%%" % stop)
                name += " / SMA20割れ%s" % ("あり" if line else "なし")
                r = simulate(d, entry816, s816, hold=hold, slots=8, topk=3,
                             stop_pct=stop, exit_line=line, ind=ind)
                grid.append((name, r))
                p("  %-30s %+11.1f %+9.1f %8.2f %6d %7.1f %+8.2f %+8.1f" %
                  (name, r["total"], r["maxdd"], r["sharpe"], r["n"], r["win"], r["avg"], r["worst"]))
    p("")
    p("  開始日を10日ずつずらした13通りでの安定性（上位4案）")
    top4 = sorted(grid, key=lambda x: -x[1]["sharpe"])[:4]
    p("  %-30s %10s %10s %10s %10s" % ("出口の組み合わせ", "中央値%", "最小%", "最大%", "中央DD%"))
    for name, _ in top4:
        hold = 10 if name.startswith("10") else 15
        stop = None if "損切なし" in name else (8 if "-8%" in name else 10)
        line = "sma" if "SMA20割れあり" in name else None
        tot, dds = [], []
        for start in range(0, 130, 10):
            r = simulate(d, window(entry816, start, T), s816, hold=hold, slots=8, topk=3,
                         stop_pct=stop, exit_line=line, ind=ind)
            tot.append(r["total"]); dds.append(r["maxdd"])
        tot, dds = np.array(tot), np.array(dds)
        p("  %-30s %+11.1f %+10.1f %+10.1f %+10.1f" %
          (name, np.median(tot), tot.min(), tot.max(), np.median(dds)))
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【28】採用案を下落局面だけで測る（市場が崩れた65日に建てた分）")
    p("-" * 112)
    from bb_study2 import fwd
    tradable = live.copy(); tradable[:, T - 12:] = False; tradable[:, :25] = False
    r10 = fwd(d, 10)
    mkt10 = np.nanmean(np.where(tradable, r10, np.nan), axis=0)
    bad = (mkt10 < 0)[None, :]
    from bb_backtest import run_trades, summarize, fmt
    for stop in (None, 8, 10):
        tr = run_trades(d, entry816 & tradable & bad, hold_max=10, stop_pct=stop)
        s = summarize(tr, "下落局面 / 損切%s" % ("なし" if stop is None else "-%d%%" % stop))
        p("  " + fmt(s))
    for stop in (None, 8, 10):
        tr = run_trades(d, entry816 & tradable & ~bad, hold_max=10, stop_pct=stop)
        s = summarize(tr, "上昇局面 / 損切%s" % ("なし" if stop is None else "-%d%%" % stop))
        p("  " + fmt(s))
    p("")
    p("  1トレードの最大損失が総資金に与える影響（枠8 = 1銘柄12.5%）")
    for stop, worst in (("なし", -47.6), ("-8%", -8.0), ("-10%", -10.0)):
        p("    損切%-5s → 1銘柄の最悪 %+6.1f%% → 総資金への影響 %+5.2f%%" % (stop, worst, worst * 0.125))
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【29】最終採用案の通し成績")
    p("-" * 112)
    final = simulate(d, entry816, s816, hold=10, slots=8, topk=3, stop_pct=8, ind=ind)
    p("  入口   : +2σ超え & 出来高1.5倍以上 & MFI>=55 & RSI<=80 & 二番天井でない")
    p("  順位   : BB816スコア降順、1日最大3銘柄")
    p("  枠     : 8（1銘柄あたり資金の12.5%）")
    p("  出口   : 10営業日で手仕舞い / 損切-8%")
    p("")
    p("  総リターン %+.1f%%  最大DD %+.1f%%  Sharpe %.2f  取引 %d回  勝率 %.1f%%  平均 %+.2f%%  最悪 %+.1f%%" %
      (final["total"], final["maxdd"], final["sharpe"], final["n"], final["win"], final["avg"], final["worst"]))
    ref = simulate(d, live, np.zeros((N, T)), hold=10, slots=8, topk=3)
    p("  比較: 何も選ばず毎日3銘柄買う  総リターン %+.1f%%  最大DD %+.1f%%  Sharpe %.2f" %
      (ref["total"], ref["maxdd"], ref["sharpe"]))
    cur = simulate(d, entry_cur, sc, hold=10, slots=8, topk=3, stop_pct=8, ind=ind)
    p("  比較: 現行スコアで同じ運用       総リターン %+.1f%%  最大DD %+.1f%%  Sharpe %.2f" %
      (cur["total"], cur["maxdd"], cur["sharpe"]))
    p("")
    reasons = {}
    for x in final["trades"]:
        reasons.setdefault(x["reason"], []).append(x["ret"])
    p("  手仕舞い理由の内訳")
    for k, v in sorted(reasons.items()):
        v = np.array(v)
        p("    %-8s %4d件 (%4.1f%%) 平均%+6.2f%%" % (k, len(v), 100 * len(v) / final["n"], v.mean()))
    p("")
    eq = final["equity"]
    p("  資産曲線（月末値・開始=1.00）")
    marks = {}
    for t in range(T):
        marks[dates[t].strftime("%Y-%m")] = eq[t]
    p("    " + "  ".join("%s:%.2f" % (k, v) for k, v in list(marks.items())))

    (OUT / "bb_study6.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study6.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
