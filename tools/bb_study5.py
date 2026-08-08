#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定の深掘り（第5部）― BB816 のスコアと出口を決める

第4部までで分かったこと
    ・現行スコア(X列)は単調でない。50点以上の超過リターンはほぼゼロ
    ・逆張り側の材料（スクイーズ、強気ダイバ、%B低位、RSI低位）は超過リターンがマイナス
    ・順張り側の材料（+2σ超え、出来高急増、MFI/RSI高位、20日騰落率上位）はプラスで test でも残る
    ・地合いフィルタ（breadth・指数>20日平均）は、下げ局面の回避に役立たなかった
    ・資金5枠のシミュレーションは1回の運から受ける影響が大きい。判断は大標本の超過リターンで行う

第5部でやること
    【21】実測の超過リターンから BB816スコアを組み、単調性を train/test で確かめる
    【22】BB816スコア上位の成績（現行スコア上位との直接比較）
    【23】出口（保有日数・損切り・利確）の最終決定に必要な数字を大標本で出す
    【24】守り型と攻め型それぞれの最悪シナリオ

出力は tools/out/bb_study5.txt
"""
import sys
from pathlib import Path

import numpy as np

from bb_data import load
from bb_backtest import indicators, signals, scores, run_trades, summarize, fmt
from bb_study2 import fwd

OUT = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)
lines = []


def p(s=""):
    print(s)
    lines.append(s)


def score816(ind, c, is_top):
    """
    実測の超過リターンに素直に比例させた 0〜100 のスコア。
    Excel の BBスクリーニング にある列だけで計算できる形にしてある。
    """
    mom20 = c / np.roll(c, 20, axis=1) - 1     # 20日騰落率
    s = np.zeros(c.shape)
    # ---- 加点（超過リターンがプラスだった材料） ----
    s += np.where(c > ind["up2"], 25, np.where(c > ind["up1"], 15, np.where(c > ind["sma"], 5, 0)))
    s += np.where(ind["vol_ratio"] >= 2.0, 20, np.where(ind["vol_ratio"] >= 1.5, 14,
                  np.where(ind["vol_ratio"] >= 1.2, 8, 0)))
    s += np.where(ind["mfi"] >= 80, 15, np.where(ind["mfi"] >= 60, 18, np.where(ind["mfi"] >= 50, 8, 0)))
    s += np.where((ind["rsi"] > 70) & (ind["rsi"] <= 80), 18,
         np.where((ind["rsi"] > 60) & (ind["rsi"] <= 70), 14,
         np.where((ind["rsi"] > 50) & (ind["rsi"] <= 60), 6, 0)))
    s += np.where(mom20 >= 0.10, 14, np.where(mom20 >= 0.03, 7, 0))
    # ---- 減点（超過リターンがマイナスだった材料） ----
    s += np.where(ind["rsi"] > 80, -12, 0)          # 過熱しすぎは伸びしろが薄い
    s += np.where(ind["rsi"] < 40, -20, 0)
    s += np.where(ind["mfi"] < 40, -15, 0)
    s += np.where(ind["pct_b"] < 0.2, -18, 0)
    s += np.where(mom20 <= -0.10, -10, 0)
    s += np.where(is_top, -8, 0)
    s += np.where(ind["squeeze"] == "スクイーズ★", -6, 0)
    return np.clip(s, 0, 100)


def main(path="V815.xlsm"):
    d = load(path)
    ind = indicators(d)
    sig = signals(d, ind)
    sc = scores(d, ind)
    c = d["close"]
    N, T = c.shape
    half = T // 2
    dates = d["dates"]
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])

    tradable = np.isfinite(ind["sma"]) & np.isfinite(c)
    tradable[:, T - 12:] = False
    tradable[:, :25] = False
    train = np.zeros_like(tradable); train[:, 25:half] = tradable[:, 25:half]
    test = np.zeros_like(tradable); test[:, half:] = tradable[:, half:]

    r10 = fwd(d, 10)
    mkt10 = np.nanmean(np.where(tradable, r10, np.nan), axis=0)
    ex10 = r10 - mkt10[None, :]
    s816 = score816(ind, c, is_top)

    p("=" * 112)
    p("ボリンジャーバンド判定 深掘りレポート（第5部・BB816スコアと出口の決定）")
    p("=" * 112)
    p("")

    p("-" * 112)
    p("【21】スコアの単調性 ― 点数が高いほど超過リターンが高くなるか")
    p("-" * 112)
    for title, arr in (("現行スコア（V815 X列）", sc), ("BB816スコア（実測配点）", s816)):
        p("  " + title)
        p("    %-10s %8s %10s %10s %11s %11s" % ("点数帯", "件数", "超過平均%", "勝ち越し%", "train超過%", "test超過%"))
        for lo, hi in [(0, 20), (20, 40), (40, 60), (60, 70), (70, 80), (80, 101)]:
            m = (arr >= lo) & (arr < hi) & tradable
            v = ex10[m & np.isfinite(ex10)]
            a = ex10[m & train & np.isfinite(ex10)]
            b = ex10[m & test & np.isfinite(ex10)]
            if len(v) < 30:
                continue
            p("    %3d〜%3d    %8d %+9.2f %10.1f %+10.2f %+10.2f" %
              (lo, hi - 1, len(v), v.mean(), 100 * (v > 0).mean(),
               a.mean() if len(a) else float("nan"), b.mean() if len(b) else float("nan")))
        p("")

    p("-" * 112)
    p("【22】上位N銘柄だけを買った場合（毎日、スコア上位から順に）")
    p("-" * 112)
    p("  %-26s %6s %8s %10s %10s %10s" % ("選び方", "上位", "件数", "10日平均%", "超過%", "勝率%"))
    for label, arr in (("現行スコア", sc), ("BB816スコア", s816)):
        for topn in (3, 5, 10, 20):
            pick = np.zeros_like(tradable)
            for t in range(T):
                col = np.where(tradable[:, t], arr[:, t], -np.inf)
                if np.all(~np.isfinite(col)):
                    continue
                idx = np.argsort(-col)[:topn]
                idx = [i for i in idx if np.isfinite(col[i]) and col[i] > -np.inf]
                pick[idx, t] = True
            v = r10[pick & np.isfinite(r10)]
            e = ex10[pick & np.isfinite(ex10)]
            p("  %-26s %6d %8d %+9.2f %+9.2f %9.1f" %
              (label, topn, len(v), v.mean(), e.mean(), 100 * (v > 0).mean()))
    p("")

    # ------------------------------------------------------------------
    p("-" * 112)
    p("【23】出口の最終決定（入口＝BB816の買い条件、大標本）")
    p("-" * 112)
    buy = ((c > ind["up2"]) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80)
           & (ind["mfi"] >= 55) & ~is_top & tradable)
    p("  入口シグナル数 %d件（250日・300銘柄あたり）平均 %.1f件/日" % (buy.sum(), buy.sum() / T))
    p("")
    for name, kw in [
        ("A 5日で手仕舞い（損切りなし）", dict(hold_max=5)),
        ("B 10日で手仕舞い（損切りなし）", dict(hold_max=10)),
        ("C 15日で手仕舞い（損切りなし）", dict(hold_max=15)),
        ("D 10日 + 損切-4%", dict(hold_max=10, stop_pct=4)),
        ("E 10日 + 損切-6%", dict(hold_max=10, stop_pct=6)),
        ("F 10日 + 損切-8%", dict(hold_max=10, stop_pct=8)),
        ("G 10日 + 損切-10%", dict(hold_max=10, stop_pct=10)),
        ("H 10日 + 損切-6% + 利確+12%", dict(hold_max=10, stop_pct=6, target_pct=12)),
        ("I SMA20割れで手仕舞い（最長20日）", dict(hold_max=20, exit_below="sma", ind=ind)),
        ("J +1σ割れで手仕舞い（最長20日）", dict(hold_max=20, exit_below="up1", ind=ind)),
        ("K 15日 + 損切-8% + 3日以降トレール8%", dict(hold_max=15, stop_pct=8, trail_after=3)),
    ]:
        tr = run_trades(d, buy, **kw)
        s = summarize(tr, name)
        r = np.array([x["ret"] for x in tr])
        p("  " + fmt(s) + "  |  下位5%%点 %+6.2f%%  -10%%以下の割合 %4.1f%%" %
          (np.percentile(r, 5), 100 * (r <= -10).mean()))
    p("")

    p("  train / test 別（B:10日損切りなし と E:10日損切-6%）")
    for name, kw in [("B 10日・損切りなし", dict(hold_max=10)), ("E 10日・損切-6%", dict(hold_max=10, stop_pct=6))]:
        for label, msk in (("train", train), ("test", test)):
            s = summarize(run_trades(d, buy & msk, **kw), "%s / %s" % (name, label))
            p("    " + fmt(s))
    p("")

    # ------------------------------------------------------------------
    p("-" * 112)
    p("【24】守り型・攻め型それぞれの最悪シナリオ（10日保有）")
    p("-" * 112)
    variants = {
        "守り型 +2σ&出来高2倍&MFI60&RSI<=75": ((c > ind["up2"]) & (ind["vol_ratio"] >= 2)
                                          & (ind["rsi"] <= 75) & (ind["mfi"] >= 60) & ~is_top),
        "折衷 +2σ&出来高1.5倍&MFI55&RSI<=80": buy,
        "攻め型 +1σ&RSI60&MFI60&出来高1.2倍": ((c > ind["up1"]) & (ind["rsi"] >= 60)
                                          & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2) & ~is_top),
    }
    for name, m in variants.items():
        for stop in (None, 6):
            tr = run_trades(d, m & tradable, hold_max=10, stop_pct=stop)
            r = np.array([x["ret"] for x in tr])
            mae = np.array([x["mae"] for x in tr])
            p("  %-34s 損切%-5s n=%5d 平均%+6.2f%% 最悪%+7.1f%% 下位5%%%+7.2f%% "
              "10%%超の損失%4.1f%% MAE中央%+6.2f%%" %
              (name, "なし" if stop is None else "-6%", len(r), r.mean(), r.min(),
               np.percentile(r, 5), 100 * (r <= -10).mean(), np.median(mae)))
    p("")

    # ------------------------------------------------------------------
    p("-" * 112)
    p("【25】最新日 %s ― BB816スコア上位20" % dates[-1])
    p("-" * 112)
    t = T - 1
    order = np.argsort(-s816[:, t])
    p("  %-6s %-14s %8s %8s %6s %6s %6s %7s %7s %6s" %
      ("コード", "銘柄名", "終値", "+2σ", "RSI", "MFI", "出来高", "20日%", "BB816", "現行"))
    for i in order[:20]:
        if not np.isfinite(c[i, t]):
            continue
        mom = (c[i, t] / c[i, t - 20] - 1) * 100 if np.isfinite(c[i, t - 20]) else float("nan")
        p("  %-6s %-14s %8.1f %8.1f %6.1f %6.1f %6.1f %+7.1f %7.0f %6.0f" %
          (d["codes"][i], d["names"][i][:14], c[i, t], ind["up2"][i, t], ind["rsi"][i, t],
           ind["mfi"][i, t], ind["vol_ratio"][i, t], mom, s816[i, t], sc[i, t]))
    p("")
    buy_today = buy[:, t] if t < buy.shape[1] else None
    m = ((c > ind["up2"]) & (ind["vol_ratio"] >= 1.5) & (ind["rsi"] <= 80)
         & (ind["mfi"] >= 55) & ~is_top)[:, t]
    p("  買い条件を満たす銘柄: %d件 → %s" %
      (m.sum(), ", ".join("%s %s" % (d["codes"][i], d["names"][i]) for i in np.where(m)[0])))

    np.save(OUT / "s816.npy", s816)
    (OUT / "bb_study5.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study5.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
