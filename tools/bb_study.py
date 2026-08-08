#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定のスイング適性を実測する。

前提
  ・データ  300銘柄 × 250営業日（2025-07-29 〜 2026-08-06、V815の価格シート）
  ・売買    引け後にシグナル → 翌営業日の始値で買う（先読みなし）
  ・出口    ルール別に比較
  ・分割    前半125日を train、後半125日を test として、
            「後付けで良く見えるだけ」の条件を弾く

出力は tools/out/bb_study.txt
"""
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores, run_trades, summarize, fmt

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
lines = []


def p(s=""):
    print(s)
    lines.append(s)


def forward_returns(d, mask, horizons=(3, 5, 10)):
    """翌日始値で買って N 日後の終値で売った場合の単純リターン。"""
    o, c = d["open"], d["close"]
    N, T = c.shape
    res = {}
    ii, tt = np.where(mask)
    for hz in horizons:
        r = []
        for i, t in zip(ii, tt):
            e = t + 1
            x = e + hz - 1
            if x >= T:
                continue
            ep, xp = o[i, e], c[i, x]
            if np.isfinite(ep) and np.isfinite(xp) and ep > 0:
                r.append((xp / ep - 1) * 100)
        res[hz] = np.array(r)
    return res


def main(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    N, T = d["close"].shape
    half = T // 2
    dates = d["dates"]

    tradable = np.isfinite(ind["sma"]) & np.isfinite(d["close"])
    tradable[:, T - 2:] = False          # 翌日始値が無い直近2日は検証から除く
    tradable[:, :25] = False             # 指標が立ち上がるまで

    train = np.zeros_like(tradable); train[:, 25:half] = tradable[:, 25:half]
    test = np.zeros_like(tradable); test[:, half:] = tradable[:, half:]

    p("=" * 108)
    p("ボリンジャーバンド判定 スイング実測レポート")
    p("=" * 108)
    p("データ      : %d銘柄 × %d営業日  %s 〜 %s" % (N, T, dates[0], dates[-1]))
    p("train期間   : %s 〜 %s" % (dates[25], dates[half - 1]))
    p("test期間    : %s 〜 %s" % (dates[half], dates[-1]))
    p("売買前提    : 引け後シグナル → 翌営業日の始値で買い（スリッページ・手数料は考慮せず）")
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【0】ベースライン ― 何も選ばず全銘柄・全日を買った場合")
    p("-" * 108)
    base = forward_returns(d, tradable)
    for hz in (3, 5, 10):
        r = base[hz]
        p("  %2d営業日保有  n=%6d  平均%+6.2f%%  中央%+6.2f%%  勝率%5.1f%%" %
          (hz, len(r), r.mean(), np.median(r), 100 * (r > 0).mean()))
    p("  ※ この期間の地合いは強い。以降の数字はこのベースラインを超えて初めて意味がある。")
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【1】V815 の戦略シグナル別 ― 保有日数を変えた素の成績")
    p("-" * 108)
    kinds = ["◎スクイーズブレイク買", "○順張り継続(walk)", "○反発(ダイバ)買", "△二番底買(確)",
             "△二番底形成", "☆スクイーズ監視", "・押し目注目", "▼-2σ割れ危険", "▼二番天井警戒", "様子見"]
    p("  %-22s %7s %26s %26s" % ("シグナル", "件数", "3日保有 平均/勝率", "10日保有 平均/勝率"))
    stats_by_sig = {}
    for k in kinds:
        m = (sig == k) & tradable
        if m.sum() == 0:
            continue
        fr = forward_returns(d, m)
        r3, r10 = fr[3], fr[10]
        if len(r3) == 0:
            continue
        stats_by_sig[k] = (r3, r10)
        p("  %-22s %7d %14.2f%% %8.1f%% %14.2f%% %8.1f%%" %
          (k, m.sum(), r3.mean(), 100 * (r3 > 0).mean(), r10.mean(), 100 * (r10 > 0).mean()))
    p("")
    p("  同じ集計を train / test で分けたもの（10日保有・平均%）")
    p("  %-22s %18s %18s" % ("シグナル", "train", "test"))
    for k in kinds:
        a = forward_returns(d, (sig == k) & train, (10,))[10]
        b = forward_returns(d, (sig == k) & test, (10,))[10]
        if len(a) < 10 and len(b) < 10:
            continue
        p("  %-22s %8d件 %+6.2f%% %8d件 %+6.2f%%" %
          (k, len(a), a.mean() if len(a) else float("nan"), len(b), b.mean() if len(b) else float("nan")))
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【2】スコア（X列）は効いているか ― スコア帯別の10日リターン")
    p("-" * 108)
    edges = [0, 20, 30, 40, 50, 60, 70, 101]
    p("  %-12s %8s %10s %10s %10s" % ("スコア帯", "件数", "平均%", "中央%", "勝率%"))
    for a, b in zip(edges[:-1], edges[1:]):
        m = tradable & (sc >= a) & (sc < b)
        r = forward_returns(d, m, (10,))[10]
        if len(r) == 0:
            continue
        p("  %3d〜%3d      %8d %+9.2f %+9.2f %9.1f" %
          (a, b - 1, len(r), r.mean(), np.median(r), 100 * (r > 0).mean()))
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【3】条件を1つずつ入れた場合の効き ― 10日保有・ベースライン比")
    p("-" * 108)
    c = d["close"]
    conds = {
        "スクイーズ★": ind["squeeze"] == "スクイーズ★",
        "収縮以上（★or収縮）": (ind["squeeze"] == "スクイーズ★") | (ind["squeeze"] == "収縮"),
        "帯拡大↑": ind["expand"],
        "終値 > +2σ": c > ind["up2"],
        "終値 > +1σ": c > ind["up1"],
        "終値 > SMA20": c > ind["sma"],
        "終値 < SMA20": c < ind["sma"],
        "終値 < -2σ": c < ind["lo2"],
        "%B 0.5〜0.8": (ind["pct_b"] >= 0.5) & (ind["pct_b"] <= 0.8),
        "%B < 0.2": ind["pct_b"] < 0.2,
        "MFI >= 80": ind["mfi"] >= 80,
        "MFI 50〜80": (ind["mfi"] >= 50) & (ind["mfi"] < 80),
        "RSI 50〜70": (ind["rsi"] >= 50) & (ind["rsi"] <= 70),
        "RSI > 70": ind["rsi"] > 70,
        "RSI < 40": ind["rsi"] < 40,
        "強気ダイバージェンス": ind["diver"] == "強気div▲",
        "上walk▲": ind["walk"] == "上walk▲",
        "下walk▼": ind["walk"] == "下walk▼",
        "出来高2倍以上": ind["vol_ratio"] >= 2,
        "二番底W▲": ind["dbl"] == "二番底W▲",
        "二番天井（形成含む）": np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]]),
    }
    b10 = base[10]
    p("  %-24s %8s %9s %9s %9s %9s" % ("条件", "件数", "平均%", "対base", "勝率%", "test平均%"))
    for name, m in conds.items():
        mm = m & tradable
        r = forward_returns(d, mm, (10,))[10]
        rt = forward_returns(d, m & test, (10,))[10]
        if len(r) < 20:
            continue
        p("  %-24s %8d %+8.2f %+8.2f %8.1f %+9.2f" %
          (name, len(r), r.mean(), r.mean() - b10.mean(), 100 * (r > 0).mean(),
           rt.mean() if len(rt) else float("nan")))
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【4】出口ルールの比較 ― 入口は「◎スクイーズブレイク買」に固定")
    p("-" * 108)
    entry = (sig == "◎スクイーズブレイク買") & tradable
    rules = [
        ("10日期日のみ", dict(hold_max=10)),
        ("5日期日のみ", dict(hold_max=5)),
        ("+1σ実体割れ（現行の利確）", dict(hold_max=20, exit_below="up1", ind=ind)),
        ("SMA20割れ（現行の損切り）", dict(hold_max=20, exit_below="sma", ind=ind)),
        ("損切-3% / 期日10日", dict(hold_max=10, stop_pct=3)),
        ("損切-5% / 期日10日", dict(hold_max=10, stop_pct=5)),
        ("損切-3% 利確+6% / 10日", dict(hold_max=10, stop_pct=3, target_pct=6)),
        ("損切-4% 利確+8% / 10日", dict(hold_max=10, stop_pct=4, target_pct=8)),
        ("損切-5% 利確+10% / 15日", dict(hold_max=15, stop_pct=5, target_pct=10)),
        ("損切-4% トレール(3日以降)", dict(hold_max=15, stop_pct=4, trail_after=3)),
    ]
    for name, kw in rules:
        tr = run_trades(d, entry, **kw)
        p("  " + fmt(summarize(tr, name)))
    p("")

    # ------------------------------------------------------------------
    p("-" * 108)
    p("【5】入口候補の比較 ― 出口は「損切-4% / 利確+8% / 期日10日」に固定")
    p("-" * 108)
    exitkw = dict(hold_max=10, stop_pct=4, target_pct=8)
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])
    cands = {
        "①◎スクイーズブレイク買（現行）": (sig == "◎スクイーズブレイク買"),
        "②○順張り継続(walk)（現行）": (sig == "○順張り継続(walk)"),
        "③○反発(ダイバ)買（現行）": (sig == "○反発(ダイバ)買"),
        "④△二番底買(確)（現行）": (sig == "△二番底買(確)"),
        "⑤・押し目注目（現行）": (sig == "・押し目注目"),
        "⑥スコア60以上": (sc >= 60),
        "⑦スコア70以上": (sc >= 70),
        "⑧+1σ超え & MFI>=60 & 帯拡大": (c > ind["up1"]) & (ind["mfi"] >= 60) & ind["expand"],
        "⑨SMA20上 & %B0.5-0.8 & MFI>=50": (c > ind["sma"]) & (ind["pct_b"] >= 0.5) & (ind["pct_b"] <= 0.8) & (ind["mfi"] >= 50),
        "⑩収縮 & +1σ超え & 帯拡大": ((ind["squeeze"] == "スクイーズ★") | (ind["squeeze"] == "収縮")) & (c > ind["up1"]) & ind["expand"],
        "⑪⑧に「二番天井でない」を追加": (c > ind["up1"]) & (ind["mfi"] >= 60) & ind["expand"] & ~is_top,
        "⑫⑪に RSI<=75 を追加": (c > ind["up1"]) & (ind["mfi"] >= 60) & ind["expand"] & ~is_top & (ind["rsi"] <= 75),
    }
    for name, m in cands.items():
        tr = run_trades(d, m & tradable, **exitkw)
        p("  " + fmt(summarize(tr, name)))
    p("")
    p("  同じ入口を train / test に分けたもの")
    p("  %-32s %26s %26s" % ("入口", "train", "test"))
    for name, m in cands.items():
        a = summarize(run_trades(d, m & train, **exitkw), name)
        b = summarize(run_trades(d, m & test, **exitkw), name)
        def one(s):
            return "n=%4d 勝率%5.1f%% 平均%+6.2f%%" % (s["n"], s.get("win", 0), s.get("avg", 0)) if s["n"] else "n=   0            "
        p("  %-32s %26s %26s" % (name, one(a), one(b)))
    p("")

    np.save(OUT / "sig.npy", sig)
    np.save(OUT / "score.npy", sc)
    (OUT / "bb_study.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
