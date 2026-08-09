# -*- coding: utf-8 -*-
"""V817 で書き換えた数式の計算結果を、V816 のキャッシュ値と株価シートから再現する。
これを <v> として埋め込むことで、楽天RSS 未接続でも管理シートの表示が壊れない。"""
import math, datetime, openpyxl

SRC = '/tmp/claude-0/-home-user--/8b86e642-77b1-5b85-a6ca-fce6f257f782/scratchpad/v816/V816.xlsm'

# ---- V811設定のパラメータ（build_v817.py と同じ値にすること） ----
B3, B4, B5, B6 = 0.06, 2.0, 40000.0, 10
H3, H4, H5, H6, H7, H8, H9, H10 = 1, 2.0, 3.0, 10.0, 1.07, 2.0, 1.5, 14


def _serial(d):
    if d is None:
        return None
    if isinstance(d, (int, float)):
        return float(d)
    if isinstance(d, datetime.datetime):
        d = d.date()
    if isinstance(d, datetime.date):
        return (d - datetime.date(1899, 12, 30)).days
    return None


def workday(start_serial, days):
    """Excel WORKDAY（祝日引数なし＝土日のみ除外）"""
    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(start_serial))
    n = 0
    while n < days:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5:
            n += 1
    return (d - datetime.date(1899, 12, 30)).days


def tick(p):
    return 1 if p <= 3000 else 5 if p <= 5000 else 10 if p <= 30000 else \
           50 if p <= 50000 else 100 if p <= 300000 else 500


def FLOOR(x, t):
    return math.floor(round(x / t, 9)) * t


def CEILING(x, t):
    return math.ceil(round(x / t, 9)) * t


def load():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    S = {}
    for nm in ('始値', '高値', '安値', '終値'):
        ws = wb[nm]
        S[nm] = {'dates': [_serial(ws.cell(3, c).value) for c in range(5, 255)],
                 'rows': {}}
        for r in range(6, 306):
            S[nm]['rows'][r] = [ws.cell(r, c).value for c in range(5, 255)]
    mk = wb['銘柄管理']
    codes = {}
    for r in range(6, 306):
        v = mk.cell(r, 2).value
        if v not in (None, ''):
            codes[str(v).strip()] = r          # 銘柄管理の実行番号
    return wb, S, codes


def code_row(codes, code):
    if code in (None, ''):
        return None
    return codes.get(str(code).strip())


def atr_pct(S, codes, code, anchor_serial=None):
    """(Σ高値 − Σ安値) ÷ Σ終値 × 100 を H10 日分。
       anchor_serial を渡すと、その日を最新日とする14日窓で計算（＝建玉時点で固定）。"""
    r = code_row(codes, code)
    if r is None:
        return ''
    off = 0
    if anchor_serial is not None:
        try:
            off = S['高値']['dates'].index(anchor_serial)
        except ValueError:
            off = 0
    hi, lo, cl = S['高値']['rows'][r], S['安値']['rows'][r], S['終値']['rows'][r]
    sh = sl = sc = 0.0
    for i in range(off, off + H10):
        if i >= len(hi):
            return ''
        for src in (hi, lo, cl):
            if not isinstance(src[i], (int, float)):
                return ''
        sh += hi[i]; sl += lo[i]; sc += cl[i]
    if sc <= 0:
        return ''
    return (sh - sl) / sc * 100


def stop_width(atr):
    if H3 == 1 and isinstance(atr, (int, float)) and atr > 0:
        return min(H6, max(H5, atr * H4))
    return B3 * 100


def qty(entry, width):
    if not isinstance(entry, (int, float)) or not width:
        return ''
    return int(B5 / (entry * (width / 100.0) * H7) / 100) * 100


def post_entry_extreme(S, codes, code, buy_serial, high=True):
    r = code_row(codes, code)
    if r is None or buy_serial is None:
        return ''
    dates = S['高値']['dates']
    try:
        pos = dates.index(buy_serial) + 1          # MATCH は 1 起点
    except ValueError:
        pos = 10
    win = max(1, pos - 1)
    src = S['高値' if high else '安値']['rows'][r]
    vals = [v for v in src[:win] if isinstance(v, (int, float))]
    if not vals:
        return ''
    return max(vals) if high else min(vals)
