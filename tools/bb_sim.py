#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
実運用に近い資金シミュレーション

第2部の portfolio() は「シグナルが出た分だけ無制限に建てる」平均リターン計算なので、
建玉数が戦略ごとにバラバラ（数十〜数百）になり、分散効果の差が最大DDに紛れ込む。
ここでは実際の売買と同じ条件に揃える。

    ・資金を slots 等分（既定5枠）。空き枠がなければ、その日のシグナルは見送り
    ・1日に新規で建てるのは最大 topk 銘柄（ランキング上位から）
    ・翌営業日の始値で建て、hold 日で手仕舞い、途中で損切り/利確に触れたらそこで終了
    ・同じ銘柄の重複建ては禁止

これで「枠5・1日2銘柄まで」という同じ土俵で戦略を比べられる。
"""
import numpy as np


def simulate(d, mask, rank, hold=10, slots=5, topk=2, stop_pct=None, target_pct=None,
             trail_pct=None, regime=None, exit_line=None, ind=None):
    o, h, l, c = d["open"], d["high"], d["low"], d["close"]
    N, T = c.shape
    cash = 1.0
    slot_cap = 1.0 / slots
    open_pos = []          # dict(i, entry_px, shares_value, day_open, stop, target, peak)
    equity = np.full(T, np.nan)
    trades = []

    for t in range(T):
        # --- 保有中の建玉を評価・手仕舞い ---
        still = []
        for pos in open_pos:
            i = pos["i"]
            px = c[i, t] if np.isfinite(c[i, t]) else pos["last"]
            lo, hi = l[i, t], h[i, t]
            held = t - pos["day_open"]
            exit_px, reason = None, None
            if pos["stop"] is not None and np.isfinite(lo) and lo <= pos["stop"]:
                exit_px, reason = pos["stop"], "損切り"
            elif pos["target"] is not None and np.isfinite(hi) and hi >= pos["target"]:
                exit_px, reason = pos["target"], "利確"
            elif exit_line is not None and held >= 1 and ind is not None:
                line = ind[exit_line][i, t]
                if np.isfinite(line) and np.isfinite(px) and px < line:
                    exit_px, reason = px, "線割れ"
            if exit_px is None and held >= hold - 1:
                exit_px, reason = px, "期日"
            if np.isfinite(hi):
                pos["peak"] = max(pos["peak"], hi)
                if trail_pct is not None and held >= 2:
                    ts = pos["peak"] * (1 - trail_pct / 100)
                    if np.isfinite(lo) and lo <= ts and exit_px is None:
                        exit_px, reason = ts, "トレール"
            if exit_px is not None and np.isfinite(exit_px):
                cash += pos["value"] * exit_px / pos["entry_px"]
                trades.append({"i": i, "in": pos["day_open"], "out": t,
                               "ret": (exit_px / pos["entry_px"] - 1) * 100, "reason": reason})
            else:
                pos["last"] = px
                still.append(pos)
        open_pos = still

        # --- 時価評価 ---
        val = cash
        for pos in open_pos:
            px = c[pos["i"], t]
            px = px if np.isfinite(px) else pos["last"]
            val += pos["value"] * px / pos["entry_px"]
        equity[t] = val

        # --- 新規建て（前日引けのシグナル → 当日始値） ---
        if t == 0:
            continue
        if regime is not None and not regime[t - 1]:
            continue
        cands = np.where(mask[:, t - 1])[0]
        if len(cands) == 0:
            continue
        held_codes = {p["i"] for p in open_pos}
        cands = [i for i in cands if i not in held_codes and np.isfinite(o[i, t]) and o[i, t] > 0]
        cands.sort(key=lambda i: -rank[i, t - 1])
        for i in cands[:topk]:
            if len(open_pos) >= slots:
                break
            amount = min(slot_cap * equity[t], cash)
            if amount <= 1e-9:
                break
            ep = o[i, t]
            cash -= amount
            open_pos.append({"i": i, "entry_px": ep, "value": amount, "day_open": t,
                             "stop": ep * (1 - stop_pct / 100) if stop_pct else None,
                             "target": ep * (1 + target_pct / 100) if target_pct else None,
                             "peak": ep, "last": ep})

    equity = np.where(np.isfinite(equity), equity, 1.0)
    dd = equity / np.maximum.accumulate(equity) - 1
    r = np.diff(equity) / equity[:-1]
    rets = np.array([x["ret"] for x in trades]) if trades else np.array([0.0])
    return {
        "total": (equity[-1] - 1) * 100,
        "maxdd": dd.min() * 100,
        "sharpe": r.mean() / r.std() * np.sqrt(250) if r.std() > 0 else 0.0,
        "n": len(trades),
        "win": 100 * (rets > 0).mean(),
        "avg": rets.mean(),
        "worst": rets.min(),
        "best": rets.max(),
        "pf": (rets[rets > 0].sum() / -rets[rets < 0].sum()) if (rets < 0).any() else float("inf"),
        "equity": equity,
        "trades": trades,
        "exposure": 100 * np.mean([1 if x else 0 for x in [len(open_pos)]]) if False else None,
    }
