#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
指標の有効性検定。

管理シートの48件ではサンプルが足りないので、価格シートの
301銘柄 × 74営業日（約2万観測）のパネル全体で各指標を評価する。

評価方法は「日次クロスセクションIC」。
  各日 t について、全銘柄を指標 X で順位づけし、
  翌日始値で入って k 日後の終値で出た場合のリターンとの
  スピアマン順位相関（＝その日のIC）を出す。
  IC の日次系列を t 検定すれば「この指標に方向性の予測力があるか」が分かる。

クロスセクションなので、地合い（全銘柄が一斉に上げ下げする成分）は
自動的に除去される。「相場が良かったから勝った」を優位性と誤認しない。
"""
import numpy as np, openpyxl, json
from collections import OrderedDict

SRC = "v805.xlsm"
SHEETS = {"o": "始値", "h": "高値", "l": "安値", "c": "終値", "v": "出来高"}


# ---------------------------------------------------------------- データ読み込み
def load_panel():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["終値"]
    # 3行目 E列以降が日付。左が新しいので反転して時系列昇順にする
    cols, dates = [], []
    for c in range(5, ws.max_column + 1):
        d = ws.cell(3, c).value
        if hasattr(d, "year"):
            cols.append(c)
            dates.append(d.date())
    order = np.argsort(dates)
    cols = [cols[i] for i in order]
    dates = [dates[i] for i in order]

    codes, names = [], []
    for r in range(6, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        codes.append(str(a).strip())
        names.append(ws.cell(r, 2).value)
    rows = list(range(6, 6 + len(codes)))

    panel = {}
    for key, sheet in SHEETS.items():
        s = wb[sheet]
        m = np.full((len(rows), len(cols)), np.nan)
        for i, r in enumerate(rows):
            for j, c in enumerate(cols):
                v = s.cell(r, c).value
                if isinstance(v, (int, float)):
                    m[i, j] = float(v)
        panel[key] = m

    # TOPIX（5行目）
    s = wb["終値"]
    topix = np.array([s.cell(5, c).value if isinstance(s.cell(5, c).value, (int, float))
                      else np.nan for c in cols], dtype=float)
    return codes, names, dates, panel, topix


# ---------------------------------------------------------------- 指標
def ema(a, n):
    """行方向（時系列）のEMA。NaNは直前値で埋めて計算する。"""
    out = np.full_like(a, np.nan)
    k = 2.0 / (n + 1)
    for i in range(a.shape[0]):
        prev = np.nan
        for j in range(a.shape[1]):
            x = a[i, j]
            if np.isnan(x):
                out[i, j] = prev
                continue
            prev = x if np.isnan(prev) else prev + k * (x - prev)
            out[i, j] = prev
    return out


def rsi(c, n=14):
    out = np.full_like(c, np.nan)
    d = np.diff(c, axis=1)
    for i in range(c.shape[0]):
        au = ad = np.nan
        for j in range(d.shape[1]):
            x = d[i, j]
            if np.isnan(x):
                continue
            u, dn = max(x, 0.0), max(-x, 0.0)
            au = u if np.isnan(au) else (au * (n - 1) + u) / n
            ad = dn if np.isnan(ad) else (ad * (n - 1) + dn) / n
            if au + ad > 0:
                out[i, j + 1] = 100.0 * au / (au + ad)
    return out


def atr(h, l, c, n=14):
    out = np.full_like(c, np.nan)
    for i in range(c.shape[0]):
        prev = np.nan
        for j in range(c.shape[1]):
            if np.isnan(h[i, j]) or np.isnan(l[i, j]):
                out[i, j] = prev
                continue
            tr = h[i, j] - l[i, j]
            if j > 0 and not np.isnan(c[i, j - 1]):
                tr = max(tr, abs(h[i, j] - c[i, j - 1]), abs(l[i, j] - c[i, j - 1]))
            prev = tr if np.isnan(prev) else (prev * (n - 1) + tr) / n
            out[i, j] = prev
    return out


def roll(a, n, fn):
    out = np.full_like(a, np.nan)
    for j in range(a.shape[1]):
        s = max(0, j - n + 1)
        w = a[:, s:j + 1]
        with np.errstate(all="ignore"):
            out[:, j] = fn(w, axis=1)
    return out


def build_indicators(panel):
    o, h, l, c, v = panel["o"], panel["h"], panel["l"], panel["c"], panel["v"]
    ind = OrderedDict()

    e5, e25, e75 = ema(c, 5), ema(c, 25), ema(c, 75)
    r14 = rsi(c, 14)
    a14 = atr(h, l, c, 14)
    m12, m26 = ema(c, 12), ema(c, 26)
    macd = m12 - m26
    sig = ema(macd, 9)

    v5 = roll(v, 5, np.nanmean)
    v25 = roll(v, 25, np.nanmean)
    hi25 = roll(h, 25, np.nanmax)
    lo25 = roll(l, 25, np.nanmin)
    hi60 = roll(h, 60, np.nanmax)
    lo60 = roll(l, 60, np.nanmin)

    prev_c = np.roll(c, 1, axis=1); prev_c[:, 0] = np.nan

    with np.errstate(all="ignore"):
        ind["RSI14"]          = r14
        ind["EMA乖離5"]        = (c - e5) / e5 * 100
        ind["EMA乖離25"]       = (c - e25) / e25 * 100
        ind["EMAパーフェクト度"]  = ((c > e5).astype(float) + (e5 > e25).astype(float)
                                  + (e25 > e75).astype(float))
        ind["MACDヒスト"]      = (macd - sig) / c * 100
        ind["MACD符号"]        = np.sign(macd)
        ind["出来高倍率"]       = v5 / v25
        ind["出来高Zスコア"]     = (v - v25) / (roll(v, 25, np.nanstd) + 1e-9)
        ind["ATR率"]          = a14 / c * 100
        ind["前日比"]          = (c - prev_c) / prev_c * 100
        ind["25日高値距離"]     = (c - hi25) / hi25 * 100
        ind["25日安値距離"]     = (c - lo25) / lo25 * 100
        ind["60日高値距離"]     = (c - hi60) / hi60 * 100
        ind["レンジ内位置25"]    = (c - lo25) / (hi25 - lo25 + 1e-9)
        ind["5日騰落率"]        = (c - np.roll(c, 5, axis=1)) / np.roll(c, 5, axis=1) * 100
        ind["20日騰落率"]       = (c - np.roll(c, 20, axis=1)) / np.roll(c, 20, axis=1) * 100
        ind["日中終値位置"]      = (c - l) / (h - l + 1e-9)
        ind["ギャップ率"]        = (o - prev_c) / prev_c * 100
        ind["日中レンジ率"]      = (h - l) / o * 100
        ind["終値/始値"]        = (c - o) / o * 100
    for k in ("5日騰落率", "20日騰落率"):
        n = 5 if "5日" in k else 20
        ind[k][:, :n] = np.nan
    return ind


# ---------------------------------------------------------------- 前向きリターン
def forward_returns(panel, kdays):
    """シグナル日 t → 翌営業日 t+1 の始値で約定 → t+k の終値で決済。
    管理シートの実約定規約（購入価格＝翌営業日始値、48/48件一致）に合わせる。"""
    o, c = panel["o"], panel["c"]
    n = c.shape[1]
    out = np.full_like(c, np.nan)
    for t in range(n):
        e, x = t + 1, t + kdays
        if x >= n:
            break
        with np.errstate(all="ignore"):
            out[:, t] = (c[:, x] - o[:, e]) / o[:, e] * 100
    return out


# ---------------------------------------------------------------- IC 検定
def spearman(a, b):
    m = np.isfinite(a) & np.isfinite(b)
    if m.sum() < 20:
        return np.nan
    ra = np.argsort(np.argsort(a[m])).astype(float)
    rb = np.argsort(np.argsort(b[m])).astype(float)
    ra -= ra.mean(); rb -= rb.mean()
    d = np.sqrt((ra ** 2).sum() * (rb ** 2).sum())
    return float((ra * rb).sum() / d) if d > 0 else np.nan


def ic_series(x, fwd):
    return np.array([spearman(x[:, t], fwd[:, t]) for t in range(x.shape[1])])


def tstat(v):
    v = v[np.isfinite(v)]
    if len(v) < 5:
        return np.nan, np.nan, 0
    se = v.std(ddof=1) / np.sqrt(len(v))
    return float(v.mean()), float(v.mean() / se) if se > 0 else np.nan, len(v)


def quintile_spread(x, fwd, q=5):
    """各日に指標で5分位に分け、最上位分位 − 最下位分位のリターン差を取る。"""
    hi, lo = [], []
    for t in range(x.shape[1]):
        a, b = x[:, t], fwd[:, t]
        m = np.isfinite(a) & np.isfinite(b)
        if m.sum() < 50:
            continue
        a2, b2 = a[m], b[m]
        idx = np.argsort(a2)
        n = len(idx) // q
        lo.append(b2[idx[:n]].mean())
        hi.append(b2[idx[-n:]].mean())
    return np.array(hi), np.array(lo)
