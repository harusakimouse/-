import numpy as np, pickle
from indicators import *

d = np.load('data.npz', allow_pickle=True)
O,H,L,C,V = d['O'],d['H'],d['L'],d['C'],d['V']
codes, names, dates = d['codes'], d['names'], d['dates']
n, m = C.shape

E5, E25, E75 = ema(C,5), ema(C,25), ema(C,75)
RSI = rsi(C,14)
EF, ES, MACD_L, MACD_S, HIST = macd(C)
ATR = atr(H,L,C,14)
ADX = adx(H,L,C,14)
VOL5 = roll_mean(V,5,True)           # 分析!M列と同じ（当日含む5日平均）
VOL20P = roll_mean(V,20,False)       # IsVolumeBurstBullish と同じ（前日まで20日）
YTDH = roll_max(H,250)
PC = np.full((n,m), np.nan); PC[:,1:] = C[:,:-1]
RET = C/PC - 1.0

# 市場プロキシ = 全銘柄の日次リターン中央値（TOPXは行データに断層があるため）
MKT = np.nanmedian(RET, axis=0)
MKT_CUM = np.nancumprod(1+np.nan_to_num(MKT))

def candle_score(o,h,l,c):
    s = np.zeros_like(c); lab = np.zeros_like(c)
    rng = h-l; body = np.abs(c-o)
    up = h-np.maximum(o,c); lo = np.minimum(o,c)-l
    ok = np.isfinite(o)&np.isfinite(h)&np.isfinite(l)&np.isfinite(c)&(rng>0)
    c1 = ok&(body>0)&(lo>=body*2)&(up<=body*0.5)                 # 下ヒゲ陽線
    c2 = ok&~c1&(c>=o)&(body>=rng*0.7)                           # 大陽線
    c3 = ok&~c1&~c2&(lo>=rng*0.5)&(up<=rng*0.2)                  # 長い下ヒゲ
    c4 = ok&~c1&~c2&~c3&(body>0)&(up>=body*2)&(lo<=body*0.5)     # 上ヒゲ警戒
    s[c1]=2; s[c2]=2; s[c3]=1; s[c4]=-2
    lab[c1]=1; lab[c2]=2; lab[c3]=3; lab[c4]=4
    return s, lab

CS_T, CL_T = candle_score(O,H,L,C)
CS_P = np.full((n,m), 0.0); CS_P[:,1:] = CS_T[:,:-1]
CL_P = np.full((n,m), 0.0); CL_P[:,1:] = CL_T[:,:-1]

# ---- v13 スコア再現 ----
score = np.zeros((n,m))
score += np.where(RSI<=30, 2, np.where(RSI<=40, 1, np.where(RSI>=70, -1, 0)))
perfect_tr = (C>E5)&(E5>E25)
up_tr      = (~perfect_tr)&(C>E25)
down_tr    = (C<E5)&(E5<E25)
score += np.where(perfect_tr, 2, np.where(up_tr, 1, np.where(down_tr, -1, 0)))
osime25 = (np.abs(C-E25)/E25<=0.02)&(C>E25)
osime5  = (~osime25)&(np.abs(C-E5)/E5<=0.01)&(C>E5)
score += np.where(osime25, 2, np.where(osime5, 1, 0))
volratio = V/VOL5
score += np.where(volratio>=2.0, 3, np.where(volratio>=1.5, 2, 0))
score += np.where(MACD_L>0, 1, np.where(MACD_L<0, -1, 0))
ytd_new = H>=YTDH-1e-9
fake = ytd_new&(((volratio>0)&(volratio<1.2))|(C<H*0.95))
score += np.where(ytd_new&~fake, 2, 0)
emaUp, macdUp, rsiUp = (EF>ES)&(ES>0), MACD_L>MACD_S, RSI>50
mom = emaUp.astype(int)+macdUp.astype(int)+rsiUp.astype(int)
score += np.where(mom==3, 2, np.where(mom==2, 1, 0))
cross = (E5 > E25)                              # 直近3日以内にEMA5がEMA25を上抜け
gc = np.zeros((n, m), bool)
for k in range(1, 4):
    prev_below = np.zeros((n, m), bool)
    prev_below[:, k:] = ~cross[:, :-k]
    still_above = np.ones((n, m), bool)
    for q in range(0, k):
        sh = np.zeros((n, m), bool)
        sh[:, q:] = cross[:, :m-q] if q > 0 else cross
        still_above &= sh
    gc |= (prev_below & still_above)
gc[:, :26] = False
po = (E5>E25)&(E25>E75)
vb = (V>=VOL20P*1.5)&(C>O)
HISTP = np.full((n,m), np.nan); HISTP[:,1:] = HIST[:,:-1]
he = (HIST>HISTP)&(HIST>0)
adxok = ADX>25
rs = RET > MKT[None,:]
score += gc*2 + po*2 + vb*2 + adxok*1 + he*1 + rs*1
score += CS_T + CS_P

VALID = np.isfinite(C)&np.isfinite(O)&np.isfinite(H)&np.isfinite(L)&np.isfinite(V)&np.isfinite(E75)&np.isfinite(ATR)
VALID[:, :80] = False   # 指標のウォームアップ期間を除外

np.savez('feat.npz', score=score, E5=E5,E25=E25,E75=E75,RSI=RSI,MACD_L=MACD_L,MACD_S=MACD_S,HIST=HIST,
         ATR=ATR,ADX=ADX,VOL5=VOL5,VOL20P=VOL20P,volratio=volratio,YTDH=YTDH,RET=RET,MKT=MKT,
         perfect_tr=perfect_tr,up_tr=up_tr,down_tr=down_tr,osime25=osime25,osime5=osime5,
         mom=mom,gc=gc,po=po,vb=vb,he=he,adxok=adxok,rs=rs,ytd_new=ytd_new,fake=fake,
         CS_T=CS_T,CS_P=CS_P,CL_T=CL_T,VALID=VALID)
print('score stats', np.nanpercentile(score[VALID],[50,75,90,95,99]))
print('signals score>=11:', (VALID&(score>=11)).sum(), ' >=16:', (VALID&(score>=16)).sum())
print('v13厳格(PO&VB&HE&mom>=2&押し目&>=16):', (VALID&(score>=16)&po&vb&he&(mom>=2)&(osime25|osime5)).sum())
print('market proxy cum', MKT_CUM[-1], 'ann vol', np.nanstd(MKT)*np.sqrt(250))
