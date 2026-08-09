#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ボリンジャーバンド判定のスイングトレード検証エンジン

V815 の BBスクリーニング／BB_帯幅／BB_スイング が Excel 数式でやっている計算を
そのまま Python で再現し、300銘柄 × 250営業日（2025-07-29〜2026-08-06）で
「そのシグナルで翌日から入って、何日持てば、いくらになったのか」を実測する。

数式との対応（BBスクリーニング 6行目基準）
    E  SMA20            AVERAGE(終値!E:X)          → sma20
    F  σ                STDEVP(終値!E:X)           → sd（母標準偏差）
    G/H/I/J  ±2σ/±1σ                                → up2/up1/lo1/lo2
    K  %B               (C-下)/( 上-下 )            → pct_b
    L  帯幅%            (up2-lo2)/sma*100          → width
    M  帯幅41日最小%    MIN(BB_帯幅!B:AP)*100      → width_min41
    N  スクイーズ       L<=M*1.1 / L<=M*1.3        → squeeze
    O  帯拡大           B>C>D                      → expand
    P  ブレイク         +2σ突破 など                → brk
    Q  RSI14            SUMPRODUCT(E:R vs F:S)     → rsi
    R  MFI14            典型価格×出来高            → mfi
    S  ダイバージェンス 10日前の終値/RSIと比較      → diver
    T  バンドウォーク   C>=+1σ / C<=-1σ            → walk
    U  二番底/天井      BB_スイング BT列           → dbl
    V  出来高倍率       出来高E/AVERAGE(E:I)       → vol_ratio
    W  戦略シグナル     入れ子IF                   → signal()
    X  スコア           加点減点の合計             → score()

すべて「その日の終値まで」しか使わない。エントリーは必ず翌営業日の始値。
先読み（lookahead）が入らないようにしている。
"""
import numpy as np

from bb_data import load

SQ_TOL = 1.10      # BB設定!B7  スクイーズ許容係数
MFI_HOT = 80       # BB設定!B10 MFI強気しきい値
DBL_PRICE_TOL = 0.05   # BB設定!B12 二番底/天井の価格許容
DBL_NECK = 0.03        # BB設定!B13 ネックライン戻り


def _roll(a, w):
    """(N,T) → (N,T,w) の後ろ向き窓。先頭 w-1 本は nan 埋め。"""
    n, t = a.shape
    pad = np.full((n, w - 1), np.nan)
    z = np.concatenate([pad, a], axis=1)
    return np.lib.stride_tricks.sliding_window_view(z, w, axis=1)


def _shift(a, k):
    """k 日前の値（k>0 で過去）。"""
    out = np.full_like(a, np.nan)
    if k == 0:
        return a.copy()
    out[:, k:] = a[:, :-k]
    return out


def indicators(d):
    c, h, l, v = d["close"], d["high"], d["low"], d["volume"]
    N, T = c.shape
    ind = {}

    w20 = _roll(c, 20)
    sma = np.nanmean(w20, axis=2)
    sd = np.nanstd(w20, axis=2)                   # STDEVP = 母標準偏差
    valid20 = (~np.isnan(w20)).sum(axis=2) == 20
    sma = np.where(valid20, sma, np.nan)
    sd = np.where(valid20, sd, np.nan)

    ind["sma"], ind["sd"] = sma, sd
    ind["up2"], ind["up1"] = sma + 2 * sd, sma + sd
    ind["lo1"], ind["lo2"] = sma - sd, sma - 2 * sd
    ind["pct_b"] = (c - ind["lo2"]) / (ind["up2"] - ind["lo2"])
    width = 4 * sd / sma * 100
    ind["width"] = width
    ind["width_min41"] = np.nanmin(_roll(width, 41), axis=2)

    # N列 スクイーズ
    sq = np.full(c.shape, "－", dtype=object)
    sq[width <= ind["width_min41"] * 1.3] = "収縮"
    sq[width <= ind["width_min41"] * SQ_TOL] = "スクイーズ★"
    ind["squeeze"] = sq

    # O列 帯拡大（今日 > 昨日 > 一昨日）
    # 帯幅が完全に同値になる日がある（同じ窓の値が並ぶ）。Excel と Python では
    # STDEVP の丸め方が 1ulp ずれるので、相対1e-9 の余裕を持たせて同値は「拡大でない」に倒す。
    def gt(a, b):
        return a > b * (1 + 1e-9)
    ind["expand"] = gt(width, _shift(width, 1)) & gt(_shift(width, 1), _shift(width, 2))

    # Q列 RSI14 / R列 MFI14（数式と同じ単純平均型）
    def rsi_at(offset):
        cc = _shift(c, offset)
        diffs = np.stack([_shift(cc, k) - _shift(cc, k + 1) for k in range(14)], axis=2)
        up = np.nansum(np.where(diffs > 0, diffs, 0), axis=2)
        dn = np.nansum(np.where(diffs < 0, -diffs, 0), axis=2)
        with np.errstate(divide="ignore", invalid="ignore"):
            r = 100 - 100 / (1 + up / dn)
        return np.where(dn == 0, 100.0, r)

    ind["rsi"] = rsi_at(0)
    ind["rsi10"] = rsi_at(10)

    tp = (h + l + c) / 3
    tpv = tp * v
    pos = np.zeros(c.shape)
    neg = np.zeros(c.shape)
    for k in range(14):
        cur, prv = _shift(tp, k), _shift(tp, k + 1)
        flow = _shift(tpv, k)
        pos += np.where(np.nan_to_num(cur > prv), np.nan_to_num(flow), 0)
        neg += np.where(np.nan_to_num(cur < prv), np.nan_to_num(flow), 0)
    with np.errstate(divide="ignore", invalid="ignore"):
        mfi = 100 - 100 / (1 + pos / neg)
    ind["mfi"] = np.where(neg == 0, 50.0, mfi)

    # S列 ダイバージェンス
    c10 = _shift(c, 10)
    diver = np.full(c.shape, "－", dtype=object)
    diver[(c < c10) & (ind["rsi"] > ind["rsi10"])] = "強気div▲"
    diver[(c > c10) & (ind["rsi"] < ind["rsi10"])] = "弱気div▼"
    ind["diver"] = diver

    # T列 バンドウォーク
    walk = np.full(c.shape, "－", dtype=object)
    walk[c >= ind["up1"]] = "上walk▲"
    walk[c <= ind["lo1"]] = "下walk▼"
    ind["walk"] = walk

    # P列 ブレイク
    brk = np.full(c.shape, "バンド内", dtype=object)
    brk[c > ind["up1"]] = "+1σ上"
    brk[c < ind["lo1"]] = "-1σ下"
    brk[c > ind["up2"]] = "↑+2σ突破"
    brk[c < ind["lo2"]] = "↓-2σ割れ"
    ind["brk"] = brk

    # V列 出来高倍率
    ind["vol_ratio"] = v / np.nanmean(_roll(v, 5), axis=2)

    # AD列 直近5日安値
    ind["low5"] = np.nanmin(_roll(c, 5), axis=2)

    ind["dbl"] = double_pattern(d, ind)
    return ind


def double_pattern(d, ind):
    """
    BB_スイング BT列の再現。
    フラクタル（前後2本より安値が突出）で厳密にスイングを取り、
    直近2つのスイング安値が同水準・間にネックラインがあれば W 型。
    """
    c, h, l = d["close"], d["high"], d["low"]
    N, T = c.shape
    out = np.full(c.shape, "－", dtype=object)

    for t in range(T):
        if t < 60:
            continue
        # オフセット 2..28 のピボット（0=当日）
        lows, highs = [], []
        for k in range(2, 29):
            i = t - k
            if i - 2 < 0:
                break
            lw = l[:, i]
            if np.all(np.isnan(lw)):
                continue
            lows.append((k, (lw < l[:, i + 1]) & (lw < l[:, i + 2]) & (lw < l[:, i - 1]) & (lw < l[:, i - 2])))
            hh = h[:, i]
            highs.append((k, (hh > h[:, i + 1]) & (hh > h[:, i + 2]) & (hh > h[:, i - 1]) & (hh > h[:, i - 2])))

        for kind, piv in (("bottom", lows), ("top", highs)):
            first = np.full(N, -1)
            second = np.full(N, -1)
            for k, mask in piv:
                m = np.nan_to_num(mask).astype(bool)
                take1 = m & (first < 0)
                first[take1] = k
                take2 = m & (first >= 0) & (first != k) & (second < 0)
                second[take2] = k
            ok = (first >= 0) & (second >= 0)
            if not ok.any():
                continue
            idx = np.where(ok)[0]
            f, s = t - first[idx], t - second[idx]
            if kind == "bottom":
                p1, p2 = l[idx, f], l[idx, s]          # 新しい谷 / 古い谷
                neck = np.array([np.nanmax(h[i, min(a, b):max(a, b) + 1]) for i, a, b in zip(idx, f, s)])
                cond = (np.abs(p1 - p2) / p2 <= DBL_PRICE_TOL) & (neck >= np.maximum(p1, p2) * (1 + DBL_NECK)) & (c[idx, t] > p1)
                lower2 = ind["lo2"]
                confirmed = (c[idx, t] >= neck) | ((p2 < lower2[idx, s]) & (p1 >= lower2[idx, f]))
                out[idx[cond & confirmed], t] = "二番底W▲"
                out[idx[cond & ~confirmed], t] = "二番底W形成"
            else:
                p1, p2 = h[idx, f], h[idx, s]
                neck = np.array([np.nanmin(l[i, min(a, b):max(a, b) + 1]) for i, a, b in zip(idx, f, s)])
                cond = (np.abs(p1 - p2) / p2 <= DBL_PRICE_TOL) & (neck <= np.minimum(p1, p2) * (1 - DBL_NECK)) & (c[idx, t] < p1)
                upper2 = ind["up2"]
                confirmed = (c[idx, t] <= neck) | ((p2 > upper2[idx, s]) & (p1 <= upper2[idx, f]))
                # 底判定が既に立っている行は上書きしない（数式も底を先に評価する）
                free = out[idx, t] == "－"
                out[idx[cond & confirmed & free], t] = "二番天井M▼"
                out[idx[cond & ~confirmed & free], t] = "二番天井M形成"
    return out


def signals(d, ind):
    """W列 戦略シグナルの再現。"""
    c = d["close"]
    sig = np.full(c.shape, "様子見", dtype=object)
    dbl, sq = ind["dbl"], ind["squeeze"]
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in dbl])

    sig[(sq == "スクイーズ★")] = "☆スクイーズ監視"
    m = (c > ind["sma"]) & (ind["pct_b"] < 0.6) & (ind["mfi"] >= 50)
    sig[m] = "・押し目注目"
    sig[c < ind["lo2"]] = "▼-2σ割れ危険"
    sig[is_top] = "▼二番天井警戒"
    sig[dbl == "二番底W形成"] = "△二番底形成"
    sig[dbl == "二番底W▲"] = "△二番底買(確)"
    sig[(ind["diver"] == "強気div▲") & (c > ind["lo1"]) & ~is_top] = "○反発(ダイバ)買"
    sig[(c >= ind["up1"]) & (ind["mfi"] >= MFI_HOT)] = "○順張り継続(walk)"
    sig[(c > ind["up2"]) & ind["expand"] & ((sq == "スクイーズ★") | (sq == "収縮"))] = "◎スクイーズブレイク買"
    sig[np.isnan(ind["sma"])] = ""
    return sig


def scores(d, ind):
    """X列 スコアの再現。"""
    c = d["close"]
    s = np.zeros(c.shape)
    L, M = ind["width"], ind["width_min41"]
    s += np.where(L <= M * 1.15, 20, np.where(L <= M * 1.4, 10, 0))
    s += np.where(c > ind["up2"], 20, np.where(c > ind["up1"], 12, np.where(c > ind["sma"], 5, 0)))
    s += np.where(ind["expand"], 10, 0)
    r, q = ind["mfi"], ind["rsi"]
    s += np.where(r >= 80, 15, np.where(r >= 60, 10, np.where(r >= 50, 5, 0)))
    s += np.where((q >= 50) & (q <= 70), 10, np.where(q > 70, 3, np.where(q >= 40, 5, 0)))
    bull = ind["diver"] == "強気div▲"
    down = ind["walk"] == "下walk▼"
    s += np.where(bull & ~down, 15, np.where(bull & down, 5, 0))
    s += np.where(ind["walk"] == "上walk▲", 5, 0)
    s += np.where(ind["vol_ratio"] >= 2, 5, np.where(ind["vol_ratio"] >= 1.5, 3, 0))
    is_top = np.array([[str(x).startswith("二番天井") for x in row] for row in ind["dbl"]])
    s += np.where(is_top, -15, 0)
    s += np.where(c < ind["sma"], -10, 0)
    s += np.where(down, -10, 0)
    s += np.where(ind["pct_b"] < 0.2, -5, 0)
    return np.clip(s, 0, 100)


# ----------------------------------------------------------------------------
# トレード実行
# ----------------------------------------------------------------------------
def run_trades(d, entries, hold_max=10, stop_pct=None, target_pct=None,
               exit_below=None, ind=None, trail_after=None, cost=0.0):
    """
    entries: (N,T) bool  … その日の引け後にシグナルが出た
    翌営業日の始値で買い、以下のいずれかで手仕舞う。
        stop_pct     : 安値が -stop_pct% に触れたらその値段で損切り（同日中）
        target_pct   : 高値が +target_pct% に触れたらその値段で利確
        exit_below   : 'up1' / 'sma' … 終値がその線を割ったら翌日始値で退出
        hold_max     : 最長保有日数（到達したら終値で退出）
    戻り値: list of dict
    """
    o, h, l, c = d["open"], d["high"], d["low"], d["close"]
    N, T = c.shape
    trades = []
    ii, tt = np.where(entries)
    for i, t in zip(ii, tt):
        e = t + 1
        if e >= T - 1:
            continue
        ep = o[i, e]
        if not np.isfinite(ep) or ep <= 0:
            continue
        stop = ep * (1 - stop_pct / 100) if stop_pct else None
        targ = ep * (1 + target_pct / 100) if target_pct else None
        exit_px, exit_day, reason = None, None, None
        peak = ep
        for k in range(0, hold_max):
            day = e + k
            if day >= T:
                break
            hi, lo, cl = h[i, day], l[i, day], c[i, day]
            if not np.isfinite(cl):
                continue
            peak = max(peak, hi if np.isfinite(hi) else cl)
            # 寄り付きが逆指値より下なら、その値段では約定しない。寄値で成立する（窓開け）
            op = o[i, day]
            if stop is not None and np.isfinite(op) and op <= stop:
                exit_px, exit_day, reason = op, day, "窓開け"
                break
            if stop is not None and np.isfinite(lo) and lo <= stop:
                exit_px, exit_day, reason = stop, day, "損切り"
                break
            if targ is not None and np.isfinite(hi) and hi >= targ:
                exit_px, exit_day, reason = targ, day, "利確"
                break
            if trail_after is not None and k >= trail_after:
                tstop = peak * (1 - stop_pct / 100) if stop_pct else None
                if tstop is not None and np.isfinite(lo) and lo <= tstop and k > 0:
                    exit_px, exit_day, reason = tstop, day, "トレール"
                    break
            if exit_below is not None and k >= 1 and ind is not None:
                line = ind[exit_below][i, day]
                if np.isfinite(line) and cl < line and day + 1 < T and np.isfinite(o[i, day + 1]):
                    exit_px, exit_day, reason = o[i, day + 1], day + 1, "線割れ"
                    break
        if exit_px is None:
            day = min(e + hold_max - 1, T - 1)
            while day > e and not np.isfinite(c[i, day]):
                day -= 1
            exit_px, exit_day, reason = c[i, day], day, "期日"
        if not np.isfinite(exit_px):
            continue
        mae = (np.nanmin(l[i, e:exit_day + 1]) / ep - 1) * 100
        mfe = (np.nanmax(h[i, e:exit_day + 1]) / ep - 1) * 100
        trades.append({
            "i": int(i), "t": int(t), "entry_day": int(e), "exit_day": int(exit_day),
            "ret": (exit_px / ep - 1) * 100 - cost, "days": int(exit_day - e + 1),
            "reason": reason, "mae": mae, "mfe": mfe,
        })
    return trades


def summarize(trades, label=""):
    if not trades:
        return {"label": label, "n": 0}
    r = np.array([x["ret"] for x in trades])
    mae = np.array([x["mae"] for x in trades])
    win = r > 0
    avg_win = r[win].mean() if win.any() else 0.0
    avg_loss = r[~win].mean() if (~win).any() else 0.0
    pf_up = r[r > 0].sum()
    pf_dn = -r[r < 0].sum()
    return {
        "label": label,
        "n": len(r),
        "win": 100 * win.mean(),
        "avg": r.mean(),
        "med": float(np.median(r)),
        "avg_win": avg_win,
        "avg_loss": avg_loss,
        "pf": (pf_up / pf_dn) if pf_dn > 0 else float("inf"),
        "worst": r.min(),
        "best": r.max(),
        "mae_med": float(np.median(mae)),
        "mae_p10": float(np.percentile(mae, 10)),
        "days": float(np.mean([x["days"] for x in trades])),
        "sum": r.sum(),
    }


def fmt(s):
    if s["n"] == 0:
        return "%-34s  該当なし" % s["label"]
    return ("%-34s n=%5d 勝率%5.1f%% 平均%+6.2f%% 中央%+6.2f%% PF%5.2f "
            "勝%+5.2f/負%+6.2f 最悪%+7.1f%% MAE中央%+6.2f%% 平均%4.1f日") % (
        s["label"], s["n"], s["win"], s["avg"], s["med"], s["pf"],
        s["avg_win"], s["avg_loss"], s["worst"], s["mae_med"], s["days"])
