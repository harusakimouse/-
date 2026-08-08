#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定の深掘り（第4部）― 同じ土俵での実運用比較

条件を全戦略で揃える。
    資金5枠 / 1日の新規は最大2銘柄 / 翌営業日始値で建てる / 同一銘柄の重複なし

比べるもの
    ・V815 の現行BB（スコア順・シグナル別）
    ・順張り型の新ルール（守り型 / 攻め型 / 折衷）
    ・地合いフィルタの有無、損切りの有無

出力は tools/out/bb_study4.txt
"""
import sys
from pathlib import Path

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores
from bb_sim import simulate
from bb_study2 import fwd, daily_returns

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
lines = []


def p(s=""):
    print(s)
    lines.append(s)


def build(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    c = d["close"]
    N, T = c.shape

    live = np.isfinite(ind["sma"]) & np.isfinite(c)
    # breadth と等金額指数は「その日までの情報」だけで作る（先読みなし）
    above = np.where(live, (c > ind["sma"]).astype(float), np.nan)
    breadth = np.nanmean(above, axis=0)
    dr = daily_returns(d)
    idx = np.cumprod(1 + np.nan_to_num(np.nanmean(np.where(live, dr, np.nan), axis=0)))
    idx_sma = np.array([idx[max(0, t - 19):t + 1].mean() for t in range(T)])
    regime = (breadth >= 0.50) & (idx > idx_sma)
    return d, ind, sig, sc, live, breadth, idx, idx_sma, regime


def main(path="V815.xlsm"):
    d, ind, sig, sc, live, breadth, idx, idx_sma, regime = build(path)
    c = d["close"]
    N, T = c.shape
    half = T // 2
    dates = d["dates"]
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])

    p("=" * 116)
    p("ボリンジャーバンド判定 深掘りレポート（第4部・同じ土俵での実運用比較）")
    p("=" * 116)
    p("共通条件: 資金5枠 / 1日の新規は最大2銘柄 / 翌営業日の始値で建てる / 同じ銘柄は重複して持たない")
    p("期間    : %s 〜 %s（%d営業日・%d銘柄）" % (dates[0], dates[-1], T, N))
    p("地合い  : breadth（SMA20超の銘柄割合）>=50%% かつ 等金額指数>20日平均 の日は %d/%d日"
      % (regime.sum(), T))
    p("")

    momentum_rank = np.nan_to_num(ind["rsi"]) + np.nan_to_num(ind["mfi"]) + 10 * np.nan_to_num(ind["vol_ratio"])

    strategies = [
        # (表示名, エントリー条件, ランキング, 地合いフィルタ, 損切り%, 保有日数)
        ("① 買い持ち（全銘柄・比較用）", live, np.zeros((N, T)), None, None, 10),
        ("② 現行BB スコア上位", live & (sc >= 40), sc, None, None, 10),
        ("③ 現行BB ◎スクイーズブレイク買", sig == "◎スクイーズブレイク買", sc, None, None, 10),
        ("④ 現行BB △二番底買(確)", sig == "△二番底買(確)", sc, None, None, 10),
        ("⑤ 現行BB ○反発(ダイバ)買", sig == "○反発(ダイバ)買", sc, None, None, 10),
        ("⑥ 現行BB スコア上位＋損切-6%", live & (sc >= 40), sc, None, 6, 10),
        ("⑦ 守り型（+2σ・出来高2倍・MFI60）", (c > ind["up2"]) & (ind["vol_ratio"] >= 2)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 60) & ~is_top, momentum_rank, None, None, 10),
        ("⑧ 攻め型（+1σ・RSI60・MFI60）", (c > ind["up1"]) & (ind["rsi"] >= 60)
         & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2) & ~is_top, momentum_rank, None, None, 10),
        ("⑨ 折衷（+2σ・出来高1.5倍・MFI55）", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, None, None, 10),
        ("⑩ 折衷＋損切-6%", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, None, 6, 10),
        ("⑪ 折衷＋地合いフィルタ", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, regime, None, 10),
        ("⑫ 折衷＋地合い＋損切-6%", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, regime, 6, 10),
        ("⑬ ⑫＋保有15日", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, regime, 6, 15),
        ("⑭ ⑫＋保有5日", (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5)
         & (ind["rsi"] <= 80) & (ind["mfi"] >= 55) & ~is_top, momentum_rank, regime, 6, 5),
    ]

    p("-" * 116)
    p("【16】戦略別の資金推移（全期間）")
    p("-" * 116)
    p("  %-34s %9s %9s %8s %6s %7s %8s %8s %8s" %
      ("戦略", "総リターン%", "最大DD%", "Sharpe", "回数", "勝率%", "平均%", "最悪%", "PF"))
    results = {}
    for name, m, rank, reg, stop, hold in strategies:
        r = simulate(d, m & live, rank, hold=hold, slots=5, topk=2, stop_pct=stop, regime=reg)
        results[name] = r
        p("  %-34s %+10.1f %+9.1f %8.2f %6d %7.1f %+8.2f %+8.1f %8.2f" %
          (name, r["total"], r["maxdd"], r["sharpe"], r["n"], r["win"], r["avg"], r["worst"], r["pf"]))
    p("")

    p("-" * 116)
    p("【17】前半125日 / 後半125日 に分けた場合")
    p("-" * 116)
    p("  %-34s %20s %20s" % ("戦略", "前半", "後半"))
    for name, m, rank, reg, stop, hold in strategies:
        halves = []
        for lo, hi in ((0, half), (half, T)):
            mm = np.zeros_like(m if m.dtype == bool else m.astype(bool))
            mm[:, lo:hi] = (m & live)[:, lo:hi]
            rr = simulate(d, mm, rank, hold=hold, slots=5, topk=2, stop_pct=stop, regime=reg)
            halves.append("%+7.1f%% DD%+6.1f%%" % (rr["total"], rr["maxdd"]))
        p("  %-34s %20s %20s" % (name, halves[0], halves[1]))
    p("")

    p("-" * 116)
    p("【18】最悪の5トレード（採用候補⑫）と、そのとき何が起きたか")
    p("-" * 116)
    r = results["⑫ 折衷＋地合い＋損切-6%"]
    worst = sorted(r["trades"], key=lambda x: x["ret"])[:5]
    for x in worst:
        p("  %s %-12s  %s 建て → %s 手仕舞い  %+6.2f%%  (%s)" %
          (d["codes"][x["i"]], d["names"][x["i"]][:12], dates[x["in"]], dates[x["out"]],
           x["ret"], x["reason"]))
    p("")
    p("  手仕舞い理由の内訳")
    reasons = {}
    for x in r["trades"]:
        reasons.setdefault(x["reason"], []).append(x["ret"])
    for k, v in sorted(reasons.items()):
        v = np.array(v)
        p("    %-8s %4d件 平均%+6.2f%%" % (k, len(v), v.mean()))
    p("")

    p("-" * 116)
    p("【19】枠数・1日の建玉数を変えた場合（採用候補⑫）")
    p("-" * 116)
    base = ((c > ind["up2"]) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80)
            & (ind["mfi"] >= 55) & ~is_top & live)
    p("  %-24s %10s %9s %8s %6s" % ("枠数 / 1日の新規", "総リターン%", "最大DD%", "Sharpe", "回数"))
    for slots in (3, 5, 8, 10):
        for topk in (1, 2, 3):
            r = simulate(d, base, momentum_rank, hold=10, slots=slots, topk=topk,
                         stop_pct=6, regime=regime)
            p("  枠%-2d / 1日%-2d銘柄        %+10.1f %+9.1f %8.2f %6d" %
              (slots, topk, r["total"], r["maxdd"], r["sharpe"], r["n"]))
    p("")

    p("-" * 116)
    p("【20】最新日 %s の地合いと候補" % dates[-1])
    p("-" * 116)
    t = T - 1
    p("  breadth = %.1f%%（SMA20超の銘柄割合）" % (breadth[t] * 100))
    p("  等金額指数 %.4f  20日平均 %.4f  → %s" %
      (idx[t], idx_sma[t], "上（買い可）" if idx[t] > idx_sma[t] else "下（見送り）"))
    p("  地合い判定: %s" % ("エントリー可" if regime[t] else "見送り（新規は建てない）"))
    p("")
    for label, m in [("折衷（採用案）", base[:, t]),
                     ("守り型", ((c > ind["up2"]) & (ind["vol_ratio"] >= 2) & (ind["rsi"] <= 80)
                                & (ind["mfi"] >= 60) & ~is_top & live)[:, t]),
                     ("攻め型", ((c > ind["up1"]) & (ind["rsi"] >= 60) & (ind["mfi"] >= 60)
                                & (ind["vol_ratio"] >= 1.2) & ~is_top & live)[:, t])]:
        hit = np.where(m)[0]
        p("  %s: %d銘柄" % (label, len(hit)))
        order = sorted(hit, key=lambda i: -momentum_rank[i, t])
        for i in order[:8]:
            p("    %-5s %-14s 終値%9.1f +2σ%9.1f RSI%5.1f MFI%5.1f 出来高%4.1f倍 現行スコア%3d" %
              (d["codes"][i], d["names"][i][:14], c[i, t], ind["up2"][i, t],
               ind["rsi"][i, t], ind["mfi"][i, t], ind["vol_ratio"][i, t], sc[i, t]))
        p("")

    (OUT / "bb_study4.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study4.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
