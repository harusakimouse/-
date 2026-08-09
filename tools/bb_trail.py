#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
「損切-4%、含み益+X%に届いたらトレーリングストップに切り替える」の検証

固定利確（+6%で必ず降りる）とは別物なので、独立に測る。

  ・建値から -4% に逆指値（初期ストップ）
  ・高値が +X%（発動水準）に触れたら、以降はトレーリングに切り替える
      ストップ = 直近高値 × (1 - トレール%)
  ・ストップは絶対に下げない（初期ストップ以下には戻さない）
  ・期日（既定10営業日）に達したら引けで手仕舞う

日中の順序は保守側に倒す。
  その日のストップ判定 → 高値でピークと発動を更新、の順に処理する
  （同じ日に「上げてから下げた」場合、利益側を先に取らない）
"""
import numpy as np


def trail_trades(d, entries, init_stop=4.0, arm=6.0, trail=5.0, hold_max=10, cost=0.0):
    """cost: 往復の売買コスト(%)。約定は窓開けを考慮し、寄り付きが逆指値より下なら寄値で約定させる。"""
    o, h, l, c = d["open"], d["high"], d["low"], d["close"]
    N, T = c.shape
    out = []
    ii, tt = np.where(entries)
    for i, t in zip(ii, tt):
        e = t + 1
        if e >= T - 1:
            continue
        ep = o[i, e]
        if not np.isfinite(ep) or ep <= 0:
            continue
        floor_stop = ep * (1 - init_stop / 100)      # 初期ストップ（ここより下げない）
        arm_px = ep * (1 + arm / 100)
        peak = ep
        armed = False
        stop = floor_stop
        exit_px = exit_day = reason = None

        for k in range(hold_max):
            day = e + k
            if day >= T:
                break
            hi, lo, cl = h[i, day], l[i, day], c[i, day]
            if not np.isfinite(cl):
                continue
            # (1) その日の始めに有効なストップで判定
            #     寄り付きが既にストップより下なら、逆指値では止まらず寄値で約定する（窓開け）
            op = o[i, day]
            if np.isfinite(op) and op <= stop:
                exit_px, exit_day = op, day
                reason = "窓開け" if not armed else "トレール(窓開け)"
                break
            if np.isfinite(lo) and lo <= stop:
                exit_px, exit_day = stop, day
                reason = "トレール" if armed else "損切り"
                break
            # (2) 高値でピーク更新・発動判定 → 翌日以降のストップに反映
            if np.isfinite(hi):
                peak = max(peak, hi)
                if not armed and hi >= arm_px:
                    armed = True
                if armed:
                    stop = max(stop, peak * (1 - trail / 100), floor_stop)
        if exit_px is None:
            day = min(e + hold_max - 1, T - 1)
            while day > e and not np.isfinite(c[i, day]):
                day -= 1
            exit_px, exit_day, reason = c[i, day], day, "期日"
        if not np.isfinite(exit_px):
            continue
        out.append({"i": int(i), "t": int(t), "in": int(e), "out": int(exit_day),
                    "ret": (exit_px / ep - 1) * 100 - cost, "days": int(exit_day - e + 1),
                    "reason": reason, "armed": armed})
    return out


def trail_portfolio(d, entries, rank, init_stop=4.0, arm=6.0, trail=5.0,
                    hold=10, slots=8, topk=3, cost=0.0):
    """資金枠つきの運用シミュレーション（bb_sim と同じ土俵）。"""
    o, h, l, c = d["open"], d["high"], d["low"], d["close"]
    N, T = c.shape
    cash, slot_cap = 1.0, 1.0 / slots
    pos, equity, trades = [], np.full(T, np.nan), []

    for t in range(T):
        keep = []
        for p in pos:
            i = p["i"]
            hi, lo, cl, op = h[i, t], l[i, t], c[i, t], o[i, t]
            px = cl if np.isfinite(cl) else p["last"]
            ex, why = None, None
            if np.isfinite(op) and op <= p["stop"]:
                ex, why = op, "窓開け"           # 寄値がストップ以下 → 寄値で約定
            elif np.isfinite(lo) and lo <= p["stop"]:
                ex, why = p["stop"], ("トレール" if p["armed"] else "損切り")
            elif t - p["in"] >= hold - 1:
                ex, why = px, "期日"
            if ex is None and np.isfinite(hi):
                p["peak"] = max(p["peak"], hi)
                if not p["armed"] and hi >= p["arm_px"]:
                    p["armed"] = True
                if p["armed"]:
                    p["stop"] = max(p["stop"], p["peak"] * (1 - trail / 100), p["floor"])
            if ex is not None and np.isfinite(ex):
                cash += p["value"] * (ex / p["ep"]) * (1 - cost / 100)
                trades.append({"i": i, "in": p["in"], "out": t,
                               "ret": (ex / p["ep"] - 1) * 100, "reason": why})
            else:
                p["last"] = px
                keep.append(p)
        pos = keep

        val = cash
        for p in pos:
            px = c[p["i"], t]
            val += p["value"] * (px if np.isfinite(px) else p["last"]) / p["ep"]
        equity[t] = val

        if t == 0:
            continue
        cand = [i for i in np.where(entries[:, t - 1])[0]
                if i not in {p["i"] for p in pos} and np.isfinite(o[i, t]) and o[i, t] > 0]
        cand.sort(key=lambda i: -rank[i, t - 1])
        for i in cand[:topk]:
            if len(pos) >= slots:
                break
            amt = min(slot_cap * equity[t], cash)
            if amt <= 1e-9:
                break
            ep = o[i, t]
            cash -= amt
            pos.append({"i": i, "ep": ep, "value": amt, "in": t, "peak": ep, "last": ep,
                        "floor": ep * (1 - init_stop / 100), "stop": ep * (1 - init_stop / 100),
                        "arm_px": ep * (1 + arm / 100), "armed": False})

    equity = np.where(np.isfinite(equity), equity, 1.0)
    dd = equity / np.maximum.accumulate(equity) - 1
    r = np.diff(equity) / equity[:-1]
    rets = np.array([x["ret"] for x in trades]) if trades else np.array([0.0])
    return {"total": (equity[-1] - 1) * 100, "maxdd": dd.min() * 100,
            "sharpe": r.mean() / r.std() * np.sqrt(250) if r.std() > 0 else 0,
            "n": len(trades), "win": 100 * (rets > 0).mean(), "avg": rets.mean(),
            "trades": trades, "equity": equity}


def summary(tr):
    r = np.array([x["ret"] for x in tr])
    win = r > 0
    up, dn = r[r > 0].sum(), -r[r < 0].sum()
    return {"n": len(r), "win": 100 * win.mean(), "avg": r.mean(), "med": float(np.median(r)),
            "pf": up / dn if dn > 0 else float("inf"), "worst": r.min(), "best": r.max(),
            "days": float(np.mean([x["days"] for x in tr])),
            "armed": 100 * np.mean([x["armed"] for x in tr])}
