#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定の深掘り（第3部）― 地合いフィルタと最終ルールの詰め

第2部までで分かったこと
    ・逆張り側（スクイーズ、強気ダイバージェンス、%B低位、RSI40未満）は超過リターンがマイナス
    ・順張り側（+2σ超え、RSI70超、MFI高位、出来高急増）はプラスで、test でも生き残る
    ・損切りは「下げ相場では効く／上げ相場では利益を削る」。固定の損切りだけでは解けない

第3部
    【11】地合い（市場全体）でスイッチを入れる／切ることの効果
    【12】保有期間と利確ラインの詰め
    【13】BB816 最終ルールの通し検証（train/test、資産曲線、最大DD、月次）
    【14】パラメータをずらしても壊れないか（頑健性）

出力は tools/out/bb_study3.txt
"""
import sys
from pathlib import Path

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores, run_trades, summarize, fmt
from bb_study2 import fwd, portfolio, daily_returns

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
lines = []


def p(s=""):
    print(s)
    lines.append(s)


def market_series(d, ind, tradable):
    """市場全体の状態。Excel 側でも同じものが作れる指標だけを使う。"""
    c = d["close"]
    N, T = c.shape
    # 騰落レシオ的な breadth（SMA20 より上にいる銘柄の割合）
    above = np.where(tradable, (c > ind["sma"]).astype(float), np.nan)
    breadth = np.nanmean(above, axis=0)
    # 等金額指数（全銘柄の日次リターン平均を積み上げたもの）
    dr = daily_returns(d)
    idx = np.cumprod(1 + np.nan_to_num(np.nanmean(np.where(tradable, dr, np.nan), axis=0)))
    idx_sma = np.full(T, np.nan)
    for t in range(T):
        if t >= 19:
            idx_sma[t] = idx[t - 19:t + 1].mean()
    return breadth, idx, idx_sma


def main(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    c = d["close"]
    N, T = c.shape
    half = T // 2
    dates = d["dates"]

    tradable = np.isfinite(ind["sma"]) & np.isfinite(c)
    tradable[:, T - 12:] = False
    tradable[:, :25] = False
    train = np.zeros_like(tradable); train[:, 25:half] = tradable[:, 25:half]
    test = np.zeros_like(tradable); test[:, half:] = tradable[:, half:]

    r10 = fwd(d, 10)
    mkt10 = np.nanmean(np.where(tradable, r10, np.nan), axis=0)
    ex10 = r10 - mkt10[None, :]
    breadth, idx, idx_sma = market_series(d, ind, tradable)
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])

    p("=" * 112)
    p("ボリンジャーバンド判定 深掘りレポート（第3部・地合いフィルタと最終ルール）")
    p("=" * 112)
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【11】地合いフィルタ ― 市場が弱い日に出たシグナルを捨てると何が変わるか")
    p("-" * 112)
    regimes = {
        "フィルタなし": np.ones(T, bool),
        "breadth>=50%（SMA20超の銘柄が半分以上）": breadth >= 0.50,
        "breadth>=40%": breadth >= 0.40,
        "breadth>=60%": breadth >= 0.60,
        "等金額指数 > その20日平均": idx > idx_sma,
        "breadth>=50% かつ 指数>20日平均": (breadth >= 0.50) & (idx > idx_sma),
    }
    entryI = (c > ind["up2"]) & (ind["vol_ratio"] >= 2) & (ind["rsi"] <= 80)
    entryF = (c > ind["up1"]) & (ind["rsi"] >= 60) & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2) & ~is_top
    p("  対象日数と、その日に出たシグナルの成績（入口 I: +2σ超え&出来高2倍&RSI<=80）")
    p("  %-38s %7s %8s %9s %9s %9s" % ("地合い条件", "対象日", "件数", "平均%", "超過%", "勝率%"))
    for name, reg in regimes.items():
        m = entryI & tradable & reg[None, :]
        v = r10[m & np.isfinite(r10)]
        e = ex10[m & np.isfinite(ex10)]
        if len(v) < 10:
            continue
        p("  %-38s %7d %8d %+8.2f %+8.2f %8.1f" %
          (name, reg.sum(), len(v), v.mean(), e.mean(), 100 * (v > 0).mean()))
    p("")
    p("  同じことを入口 F（+1σ上&RSI60&MFI60&出来高1.2倍）で")
    for name, reg in regimes.items():
        m = entryF & tradable & reg[None, :]
        v = r10[m & np.isfinite(r10)]
        e = ex10[m & np.isfinite(ex10)]
        if len(v) < 10:
            continue
        p("  %-38s %7d %8d %+8.2f %+8.2f %8.1f" %
          (name, reg.sum(), len(v), v.mean(), e.mean(), 100 * (v > 0).mean()))
    p("")
    p("  ★ 下げ局面だけを見たときの効果（市場の10日先リターンがマイナスだった日）")
    bad = (mkt10 < 0)[None, :]
    for name, reg in regimes.items():
        m = entryI & tradable & reg[None, :] & bad
        v = r10[m & np.isfinite(r10)]
        if len(v) < 5:
            p("  %-38s 該当 %d件（フィルタで回避）" % (name, len(v)))
            continue
        p("  %-38s %8d件 平均%+7.2f%% 最悪%+7.1f%%" % (name, len(v), v.mean(), v.min()))
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【12】保有期間と出口 ― 入口を I（地合いフィルタ付き）に固定して比較")
    p("-" * 112)
    reg = (breadth >= 0.50) & (idx > idx_sma)
    entry = entryI & tradable & reg[None, :]
    p("  シグナル総数 %d件" % entry.sum())
    for name, kw in [
        ("3日で手仕舞い", dict(hold_max=3)),
        ("5日で手仕舞い", dict(hold_max=5)),
        ("10日で手仕舞い", dict(hold_max=10)),
        ("15日で手仕舞い", dict(hold_max=15)),
        ("10日 + 損切-5%", dict(hold_max=10, stop_pct=5)),
        ("10日 + 損切-6%", dict(hold_max=10, stop_pct=6)),
        ("10日 + 損切-8%", dict(hold_max=10, stop_pct=8)),
        ("10日 + 損切-6% + 利確+10%", dict(hold_max=10, stop_pct=6, target_pct=10)),
        ("10日 + 損切-6% + 利確+15%", dict(hold_max=10, stop_pct=6, target_pct=15)),
        ("15日 + 損切-6% + 3日以降トレール", dict(hold_max=15, stop_pct=6, trail_after=3)),
        ("+1σ実体割れ（現行の利確）", dict(hold_max=20, exit_below="up1", ind=ind)),
        ("SMA20割れ（現行の損切り）", dict(hold_max=20, exit_below="sma", ind=ind)),
    ]:
        p("  " + fmt(summarize(run_trades(d, entry, **kw), name)))
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【13】BB816 最終候補の通し検証")
    p("-" * 112)
    rules = {
        "守り型（保守）": ((c > ind["up2"]) & (ind["vol_ratio"] >= 2) & (ind["rsi"] <= 80)
                      & (ind["mfi"] >= 60) & ~is_top),
        "攻め型（積極）": ((c > ind["up1"]) & (ind["rsi"] >= 60) & (ind["mfi"] >= 60)
                      & (ind["vol_ratio"] >= 1.2) & ~is_top),
        "折衷（採用案）": ((c > ind["up2"]) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80)
                      & (ind["mfi"] >= 55) & ~is_top),
    }
    for rname, rmask in rules.items():
        for regname, rg in [("地合いフィルタなし", np.ones(T, bool)), ("地合いフィルタあり", reg)]:
            m = rmask & tradable & rg[None, :]
            for stop in (None, 6):
                pf = portfolio(d, m, hold=10, stop_pct=stop)
                trs = summarize(run_trades(d, m, hold_max=10, stop_pct=stop), "")
                p("  %-14s %-18s 損切%-4s  件数%5d 勝率%5.1f%% 平均%+6.2f%% | 資産%+8.1f%% 最大DD%+7.1f%% Sharpe%5.2f" %
                  (rname, regname, "なし" if stop is None else "-%d%%" % stop,
                   trs["n"], trs.get("win", 0), trs.get("avg", 0),
                   pf["total"], pf["maxdd"], pf["sharpe"]))
        p("")
    p("  比較用: 全銘柄を毎日買う（地合いそのもの）")
    pf = portfolio(d, tradable, hold=10)
    p("    資産%+8.1f%% 最大DD%+7.1f%% Sharpe%5.2f" % (pf["total"], pf["maxdd"], pf["sharpe"]))
    p("")

    p("  採用案の train / test 別成績（10日保有・損切-6%）")
    adopt = rules["折衷（採用案）"] & reg[None, :]
    for label, msk in [("train", train), ("test", test)]:
        s = summarize(run_trades(d, adopt & msk, hold_max=10, stop_pct=6), label)
        p("    " + fmt(s))
    p("")

    p("  採用案の月別成績（エントリー月ごと・10日保有・損切-6%）")
    tr = run_trades(d, adopt & tradable, hold_max=10, stop_pct=6)
    by_month = {}
    for x in tr:
        key = "%s" % dates[x["entry_day"]].strftime("%Y-%m")
        by_month.setdefault(key, []).append(x["ret"])
    for k in sorted(by_month):
        v = np.array(by_month[k])
        p("    %s  n=%3d  平均%+6.2f%%  勝率%5.1f%%  最悪%+7.1f%%" %
          (k, len(v), v.mean(), 100 * (v > 0).mean(), v.min()))
    p("")

    # ---------------------------------------------------------------
    p("-" * 112)
    p("【14】頑健性 ― しきい値をずらしても成績が崩れないか（10日保有・損切-6%・地合いフィルタあり）")
    p("-" * 112)
    p("  %-34s %8s %9s %9s %9s" % ("条件", "件数", "勝率%", "平均%", "超過%"))
    for volth in (1.2, 1.5, 2.0, 2.5):
        for mfith in (50, 55, 60, 70):
            m = ((c > ind["up2"]) & (ind["vol_ratio"] >= volth) & (ind["rsi"] <= 80)
                 & (ind["mfi"] >= mfith) & ~is_top & tradable & reg[None, :])
            s = summarize(run_trades(d, m, hold_max=10, stop_pct=6), "")
            e = ex10[m & np.isfinite(ex10)]
            if s["n"] < 20:
                continue
            p("  出来高>=%.1f倍 & MFI>=%-3d              %8d %8.1f %+8.2f %+8.2f" %
              (volth, mfith, s["n"], s["win"], s["avg"], e.mean() if len(e) else float("nan")))
    p("")
    p("  σ倍率を変えた場合（出来高1.5倍・MFI55固定）")
    for k, lbl in ((2.0, "+2σ超え"), (1.5, "+1.5σ超え"), (1.0, "+1σ超え")):
        band = ind["sma"] + k * ind["sd"]
        m = ((c > band) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80) & (ind["mfi"] >= 55)
             & ~is_top & tradable & reg[None, :])
        s = summarize(run_trades(d, m, hold_max=10, stop_pct=6), lbl)
        e = ex10[m & np.isfinite(ex10)]
        p("  %-34s %8d %8.1f %+8.2f %+8.2f" % (lbl, s["n"], s["win"], s["avg"], e.mean()))
    p("")

    # 明日の実弾候補
    p("-" * 112)
    p("【15】最新日（%s）の判定 ― 採用案が拾う銘柄" % dates[-1])
    p("-" * 112)
    last = T - 1
    p("  地合い: breadth=%.0f%%  指数%s20日平均  → %s" %
      (breadth[last] * 100, ">" if idx[last] > idx_sma[last] else "<=",
       "エントリー可" if (breadth[last] >= 0.5 and idx[last] > idx_sma[last]) else "見送り"))
    m = rules["折衷（採用案）"][:, last]
    hit = np.where(m)[0]
    if len(hit) == 0:
        p("  採用案の条件に合致する銘柄なし（無理に買わないのも判断）")
    for i in hit:
        p("    %s %-12s 終値%9.1f  +2σ%9.1f  RSI%5.1f MFI%5.1f 出来高%4.1f倍 スコア%3d" %
          (d["codes"][i], d["names"][i][:12], c[i, last], ind["up2"][i, last],
           ind["rsi"][i, last], ind["mfi"][i, last], ind["vol_ratio"][i, last], sc[i, last]))
    p("")
    p("  参考: 攻め型が拾う銘柄数 %d、守り型 %d" %
      (rules["攻め型（積極）"][:, last].sum(), rules["守り型（保守）"][:, last].sum()))

    (OUT / "bb_study3.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study3.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
