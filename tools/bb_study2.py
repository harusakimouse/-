#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB判定の深掘り（第2部）

第1部で分かったこと
    ・この250日は地合いが強く、何も選ばず買っても10日で平均+1.80%
    ・V815のBBシグナルは、その地合い分をほとんど超えていない
そこで第2部では

    【6】超過リターン（同じ日の全銘柄平均を引いた値＝銘柄選択の実力）で測り直す
    【7】効く条件の組み合わせを train で探し、test で確かめる
    【8】実際に毎日上位K銘柄を買い続けたときの資産曲線・最大ドローダウン
    【9】損切り水準の感度（どこに置くと何が起きるか）
    【10】地合いが悪い日だけを取り出したときの挙動（守りの検証）

出力は tools/out/bb_study2.txt
"""
import sys
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


def fwd(d, hz):
    """翌日始値 → hz営業日後の終値 のリターン(%)。(N,T) 配列で返す。"""
    o, c = d["open"], d["close"]
    N, T = c.shape
    out = np.full((N, T), np.nan)
    for t in range(T - hz - 1):
        e = t + 1
        ep, xp = o[:, e], c[:, e + hz - 1]
        with np.errstate(invalid="ignore", divide="ignore"):
            out[:, t] = (xp / ep - 1) * 100
    return out


def daily_returns(d):
    c = d["close"]
    r = np.full_like(c, np.nan)
    r[:, 1:] = c[:, 1:] / c[:, :-1] - 1
    return r


def portfolio(d, mask, hold=10, slots=None, stop_pct=None):
    """
    毎日 mask が立った銘柄を等金額で買い、hold 日保有する運用を再現。
    その日に保有している全建玉の平均日次リターンをポートフォリオの日次リターンとする
    （建玉ゼロの日は現金＝0%）。
    """
    o, c, l = d["open"], d["close"], d["low"]
    N, T = c.shape
    dr = daily_returns(d)
    port = np.zeros(T)
    npos = np.zeros(T)
    ii, tt = np.where(mask)
    order = np.argsort(tt)
    ii, tt = ii[order], tt[order]
    for i, t in zip(ii, tt):
        e = t + 1
        if e >= T - 1 or not np.isfinite(o[i, e]) or o[i, e] <= 0:
            continue
        ep = o[i, e]
        stop = ep * (1 - stop_pct / 100) if stop_pct else None
        for k in range(hold):
            day = e + k
            if day >= T:
                break
            if k == 0:
                ret = c[i, day] / ep - 1
            else:
                ret = dr[i, day]
            if not np.isfinite(ret):
                continue
            if stop is not None and np.isfinite(l[i, day]) and l[i, day] <= stop:
                prev = ep if k == 0 else c[i, day - 1]
                port[day] += stop / prev - 1
                npos[day] += 1
                break
            port[day] += ret
            npos[day] += 1
    with np.errstate(invalid="ignore"):
        daily = np.where(npos > 0, port / np.maximum(npos, 1), 0.0)
    eq = np.cumprod(1 + daily)
    dd = eq / np.maximum.accumulate(eq) - 1
    return {"daily": daily, "eq": eq, "total": (eq[-1] - 1) * 100, "maxdd": dd.min() * 100,
            "sharpe": daily.mean() / daily.std() * np.sqrt(250) if daily.std() > 0 else 0,
            "avg_pos": npos[npos > 0].mean() if (npos > 0).any() else 0,
            "days_in": (npos > 0).mean() * 100}


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
    r5 = fwd(d, 5)
    # 同じ日の全銘柄平均を引いた超過リターン＝「銘柄を選んだ効果」だけを取り出す
    mkt10 = np.nanmean(np.where(tradable, r10, np.nan), axis=0)
    mkt5 = np.nanmean(np.where(tradable, r5, np.nan), axis=0)
    ex10 = r10 - mkt10[None, :]
    ex5 = r5 - mkt5[None, :]

    p("=" * 112)
    p("ボリンジャーバンド判定 深掘りレポート（第2部・超過リターン評価）")
    p("=" * 112)
    p("超過リターン = その銘柄の10日リターン − 同じ日に全銘柄を買った場合の平均10日リターン")
    p("             プラスなら「地合いを差し引いても銘柄選択が効いている」")
    p("")

    def stat(m, arr):
        v = arr[m & tradable & np.isfinite(arr)]
        if len(v) < 20:
            return None
        return len(v), v.mean(), np.median(v), 100 * (v > 0).mean()

    p("-" * 112)
    p("【6】V815 シグナル別の超過リターン（10日保有）")
    p("-" * 112)
    p("  %-22s %8s %10s %10s %10s %12s %12s" % ("シグナル", "件数", "超過平均%", "超過中央%", "勝ち越し%", "train超過%", "test超過%"))
    for k in ["◎スクイーズブレイク買", "○順張り継続(walk)", "○反発(ダイバ)買", "△二番底買(確)",
              "△二番底形成", "☆スクイーズ監視", "・押し目注目", "▼-2σ割れ危険", "▼二番天井警戒", "様子見"]:
        m = sig == k
        s = stat(m, ex10)
        if not s:
            continue
        a = stat(m & train, ex10)
        b = stat(m & test, ex10)
        p("  %-22s %8d %+9.2f %+9.2f %9.1f %+11.2f %+11.2f" %
          (k, s[0], s[1], s[2], s[3], a[1] if a else float("nan"), b[1] if b else float("nan")))
    p("")
    p("  スコア帯別の超過リターン")
    for a_, b_ in [(0, 20), (20, 30), (30, 40), (40, 50), (50, 60), (60, 70), (70, 101)]:
        s = stat((sc >= a_) & (sc < b_), ex10)
        if s:
            p("    スコア%3d〜%3d  n=%6d  超過%+6.2f%%  勝ち越し%5.1f%%" % (a_, b_ - 1, s[0], s[1], s[3]))
    p("")

    p("-" * 112)
    p("【7】単独条件の超過リターン（効く順）")
    p("-" * 112)
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])
    conds = {
        "終値 > +2σ": c > ind["up2"],
        "終値 > +1σ": c > ind["up1"],
        "終値 > SMA20": c > ind["sma"],
        "%B 0.5〜0.8": (ind["pct_b"] >= 0.5) & (ind["pct_b"] <= 0.8),
        "%B 0.8〜1.0": (ind["pct_b"] >= 0.8) & (ind["pct_b"] <= 1.0),
        "%B > 1.0": ind["pct_b"] > 1.0,
        "%B < 0.2": ind["pct_b"] < 0.2,
        "RSI > 70": ind["rsi"] > 70,
        "RSI 60〜70": (ind["rsi"] > 60) & (ind["rsi"] <= 70),
        "RSI 50〜60": (ind["rsi"] > 50) & (ind["rsi"] <= 60),
        "RSI < 40": ind["rsi"] < 40,
        "MFI >= 80": ind["mfi"] >= 80,
        "MFI 60〜80": (ind["mfi"] >= 60) & (ind["mfi"] < 80),
        "MFI < 40": ind["mfi"] < 40,
        "出来高 >= 2倍": ind["vol_ratio"] >= 2,
        "出来高 1.5〜2倍": (ind["vol_ratio"] >= 1.5) & (ind["vol_ratio"] < 2),
        "出来高 < 0.7倍": ind["vol_ratio"] < 0.7,
        "スクイーズ★": ind["squeeze"] == "スクイーズ★",
        "帯拡大↑": ind["expand"],
        "帯拡大↑ かつ +1σ上": ind["expand"] & (c > ind["up1"]),
        "強気ダイバージェンス": ind["diver"] == "強気div▲",
        "弱気ダイバージェンス": ind["diver"] == "弱気div▼",
        "二番底W▲": ind["dbl"] == "二番底W▲",
        "二番天井（形成含む）": is_top,
        "20日騰落率 上位（+10%超）": (c / np.roll(c, 20, axis=1) - 1 > 0.10),
        "20日騰落率 下位（-10%未満）": (c / np.roll(c, 20, axis=1) - 1 < -0.10),
    }
    rows = []
    for name, m in conds.items():
        s = stat(m, ex10)
        b = stat(m & test, ex10)
        if s:
            rows.append((s[1], name, s[0], s[3], b[1] if b else float("nan")))
    rows.sort(reverse=True)
    p("  %-26s %8s %10s %10s %12s" % ("条件", "件数", "超過平均%", "勝ち越し%", "test超過%"))
    for v, name, n, w, tv in rows:
        p("  %-26s %8d %+9.2f %9.1f %+11.2f" % (name, n, v, w, tv))
    p("")

    p("-" * 112)
    p("【8】組み合わせ ― train で作り、test で確かめる（10日保有・超過リターン）")
    p("-" * 112)
    mom = (c > ind["up1"])
    combos = {
        "A": c > ind["up2"],
        "B": (c > ind["up2"]) & (ind["vol_ratio"] >= 1.5),
        "C": (c > ind["up2"]) & (ind["mfi"] >= 70),
        "D": mom & (ind["rsi"] >= 60) & (ind["mfi"] >= 60),
        "E": mom & (ind["rsi"] >= 60) & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2),
        "F": mom & (ind["rsi"] >= 60) & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2) & ~is_top,
        "G": mom & (ind["rsi"] >= 60) & (ind["mfi"] >= 60) & (ind["vol_ratio"] >= 1.2) & ~is_top & ind["expand"],
        "H": (ind["squeeze"] == "スクイーズ★") & mom & ind["expand"],
        "I": (c > ind["up2"]) & (ind["vol_ratio"] >= 2) & (ind["rsi"] <= 80),
        "J": sig == "◎スクイーズブレイク買",
        "K": sc >= 60,
    }
    labels = {"A": "A +2σ超え 単独", "B": "B +2σ超え & 出来高1.5倍以上", "C": "C +2σ超え & MFI>=70",
              "D": "D +1σ上 & RSI>=60 & MFI>=60", "E": "E D & 出来高1.2倍以上", "F": "F E & 二番天井でない",
              "G": "G F & 帯拡大", "H": "H スクイーズ★ & +1σ上 & 帯拡大", "I": "I +2σ超え & 出来高2倍 & RSI<=80",
              "J": "J 現行◎スクイーズブレイク買", "K": "K 現行スコア60以上"}
    p("  %-30s %26s %26s" % ("条件", "train（超過）", "test（超過）"))
    for key, m in combos.items():
        name = labels[key]
        a = stat(m & train, ex10)
        b = stat(m & test, ex10)
        fa = "n=%5d 超過%+6.2f%% 勝越%5.1f%%" % a[0:1] + "" if False else (
            "n=%5d 超過%+6.2f%% 勝越%5.1f%%" % (a[0], a[1], a[3]) if a else "n=    0")
        fb = "n=%5d 超過%+6.2f%% 勝越%5.1f%%" % (b[0], b[1], b[3]) if b else "n=    0"
        p("  %-30s %26s %26s" % (name, fa, fb))
    p("")

    p("-" * 112)
    p("【9】実運用シミュレーション ― 毎日シグナル銘柄を等金額で買い、10日保有")
    p("-" * 112)
    p("  %-30s %9s %9s %9s %9s %9s" % ("戦略", "総リターン%", "最大DD%", "Sharpe", "平均建玉", "稼働日%"))
    bench = portfolio(d, tradable & (np.random.RandomState(0).rand(N, T) < 0.02), hold=10)
    strategies = [
        ("全銘柄ランダム2%（地合い基準）", tradable & (np.random.RandomState(0).rand(N, T) < 0.02), None),
        ("現行◎スクイーズブレイク買", (sig == "◎スクイーズブレイク買") & tradable, None),
        ("現行スコア60以上", (sc >= 60) & tradable, None),
        ("現行△二番底買(確)", (sig == "△二番底買(確)") & tradable, None),
        ("現行○反発(ダイバ)買", (sig == "○反発(ダイバ)買") & tradable, None),
        ("F momentum（損切りなし）", combos["F"] & tradable, None),
        ("F momentum + 損切-5%", combos["F"] & tradable, 5),
        ("F momentum + 損切-8%", combos["F"] & tradable, 8),
        ("I +2σ&出来高2倍", combos["I"] & tradable, None),
    ]
    for name, m, stop in strategies:
        r = portfolio(d, m, hold=10, stop_pct=stop)
        p("  %-30s %+10.1f %+9.1f %9.2f %9.1f %9.1f" %
          (name, r["total"], r["maxdd"], r["sharpe"], r["avg_pos"], r["days_in"]))
    p("")

    p("-" * 112)
    p("【10】守りの検証 ― 下落局面（市場平均10日リターンがマイナスの日に出たシグナル）")
    p("-" * 112)
    bad_days = mkt10 < 0
    bad = tradable & bad_days[None, :]
    p("  下落局面の該当日数 %d日 / %d日" % (bad_days.sum(), T))
    p("  %-30s %8s %10s %10s %10s" % ("シグナル", "件数", "平均%", "勝率%", "最悪%"))
    for name, m in [("全銘柄（地合い）", tradable),
                    ("現行◎スクイーズブレイク買", sig == "◎スクイーズブレイク買"),
                    ("現行スコア60以上", sc >= 60),
                    ("現行△二番底買(確)", sig == "△二番底買(確)"),
                    ("現行○反発(ダイバ)買", sig == "○反発(ダイバ)買"),
                    ("F momentum", combos["F"]),
                    ("I +2σ&出来高2倍", combos["I"])]:
        v = r10[(m & bad) & np.isfinite(r10)]
        if len(v) < 10:
            continue
        p("  %-30s %8d %+9.2f %9.1f %+9.1f" % (name, len(v), v.mean(), 100 * (v > 0).mean(), v.min()))
    p("")
    p("  下落局面での損切り効果（入口=F momentum、10日保有）")
    for stop in [None, 3, 4, 5, 6, 8]:
        tr = run_trades(d, combos["F"] & bad, hold_max=10, stop_pct=stop)
        s = summarize(tr, "損切り %s" % ("なし" if stop is None else "-%d%%" % stop))
        p("    " + fmt(s))
    p("")
    p("  上昇局面での同じ損切りの副作用（入口=F momentum）")
    good = tradable & (~bad_days)[None, :]
    for stop in [None, 3, 4, 5, 6, 8]:
        tr = run_trades(d, combos["F"] & good, hold_max=10, stop_pct=stop)
        s = summarize(tr, "損切り %s" % ("なし" if stop is None else "-%d%%" % stop))
        p("    " + fmt(s))

    (OUT / "bb_study2.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n-> %s" % (OUT / "bb_study2.txt"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
