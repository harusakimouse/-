import numpy as np

def ema(X, period):
    """X: (n_stocks, n_days) chronological. Recursive EMA seeded at first valid value."""
    k = 2.0/(period+1.0)
    n, m = X.shape
    E = np.full((n, m), np.nan)
    for i in range(n):
        e = np.nan
        for j in range(m):
            v = X[i, j]
            if not np.isfinite(v) or v <= 0:
                E[i, j] = e
                continue
            e = v if not np.isfinite(e) else k*v + (1-k)*e
            E[i, j] = e
    return E

def rsi(C, period=14):
    n, m = C.shape
    R = np.full((n, m), np.nan)
    d = np.diff(C, axis=1)
    up = np.where(d > 0, d, 0.0)
    dn = np.where(d < 0, -d, 0.0)
    for i in range(n):
        au = ad = np.nan
        for j in range(m-1):
            u, v = up[i, j], dn[i, j]
            if not (np.isfinite(u) and np.isfinite(v)):
                continue
            if not np.isfinite(au):
                au, ad = u, v
            else:
                au = (au*(period-1)+u)/period
                ad = (ad*(period-1)+v)/period
            R[i, j+1] = 100.0 if ad == 0 else 100 - 100/(1+au/ad)
    return R

def macd(C, f=12, s=26, sig=9):
    ef, es = ema(C, f), ema(C, s)
    line = ef - es
    signal = ema(np.where(np.isfinite(line), line, np.nan), sig)
    return ef, es, line, signal, line - signal

def atr(H, L, C, period=14):
    n, m = H.shape
    pc = np.full((n, m), np.nan); pc[:, 1:] = C[:, :-1]
    tr = np.nanmax(np.stack([H-L, np.abs(H-pc), np.abs(L-pc)]), axis=0)
    A = np.full((n, m), np.nan)
    for i in range(n):
        a = np.nan
        for j in range(m):
            t = tr[i, j]
            if not np.isfinite(t):
                A[i, j] = a; continue
            a = t if not np.isfinite(a) else (a*(period-1)+t)/period
            A[i, j] = a
    return A

def adx(H, L, C, period=14):
    n, m = H.shape
    A = np.full((n, m), np.nan)
    for i in range(n):
        atr_ = pdm = ndm = np.nan
        for j in range(1, m):
            h, l, ph, pl, pc = H[i, j], L[i, j], H[i, j-1], L[i, j-1], C[i, j-1]
            if not all(np.isfinite(x) for x in (h, l, ph, pl, pc)):
                A[i, j] = A[i, j-1]; continue
            up, dw = h-ph, pl-l
            p = up if (up > dw and up > 0) else 0.0
            q = dw if (dw > up and dw > 0) else 0.0
            tr = max(h-l, abs(h-pc), abs(l-pc))
            if not np.isfinite(atr_):
                atr_, pdm, ndm = tr, p, q
            else:
                atr_ = (atr_*(period-1)+tr)/period
                pdm = (pdm*(period-1)+p)/period
                ndm = (ndm*(period-1)+q)/period
            if atr_ > 0:
                pdi, ndi = 100*pdm/atr_, 100*ndm/atr_
                dx = 0.0 if (pdi+ndi) == 0 else 100*abs(pdi-ndi)/(pdi+ndi)
                prev = A[i, j-1]
                A[i, j] = dx if not np.isfinite(prev) else (prev*(period-1)+dx)/period
            else:
                A[i, j] = A[i, j-1]
    return A

def roll_mean(X, w, include_today=True):
    n, m = X.shape
    R = np.full((n, m), np.nan)
    for j in range(m):
        a = j-w+1 if include_today else j-w
        b = j+1 if include_today else j
        if a < 0: continue
        R[:, j] = np.nanmean(X[:, a:b], axis=1)
    return R

def roll_max(X, w):
    n, m = X.shape
    R = np.full((n, m), np.nan)
    for j in range(m):
        a = max(0, j-w+1)
        R[:, j] = np.nanmax(X[:, a:j+1], axis=1)
    return R
