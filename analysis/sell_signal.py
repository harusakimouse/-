import numpy as np
from indicators import *
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C,V=d['O'],d['H'],d['L'],d['C'],d['V']; n,m=C.shape
E5,E25,E75,RSI,ATR,ADX=f['E5'],f['E25'],f['E75'],f['RSI'],f['ATR'],f['ADX']
MACD_L,MACD_S,HIST,volratio,RET,MKT=f['MACD_L'],f['MACD_S'],f['HIST'],f['volratio'],f['RET'],f['MKT']
EF,ES=f['E5'],f['E25']  # placeholder; recompute EMA12/26
from indicators import ema as _ema
EF12,ES26=_ema(C,12),_ema(C,26)
YTDL=np.full((n,m),np.nan)
for j in range(m):
    a=max(0,j-249); YTDL[:,j]=np.nanmin(L[:,a:j+1],axis=1)

def bear_candle(o,h,l,c):
    s=np.zeros_like(c); rng=h-l; body=np.abs(c-o)
    up=h-np.maximum(o,c); lo=np.minimum(o,c)-l
    ok=np.isfinite(o)&np.isfinite(h)&np.isfinite(l)&np.isfinite(c)&(rng>0)
    c1=ok&(body>0)&(up>=body*2)&(lo<=body*0.5)&(c<o)
    c2=ok&~c1&(c<o)&(body>=rng*0.7)
    c3=ok&~c1&~c2&(up>=rng*0.5)&(lo<=rng*0.2)
    c4=ok&~c1&~c2&~c3&(body>0)&(lo>=body*2)&(up<=body*0.5)
    s[c1]=2;s[c2]=2;s[c3]=1;s[c4]=-2
    return s
BC=bear_candle(O,H,L,C); BCP=np.full((n,m),0.0); BCP[:,1:]=BC[:,:-1]

ss=np.zeros((n,m))
ss+=np.where(RSI>=70,2,np.where(RSI>=60,1,np.where(RSI<=30,-1,0)))
down_tr=(C<E5)&(E5<E25); perf_up=(C>E5)&(E5>E25)
ss+=np.where(down_tr,2,np.where(perf_up,-1,0))
osime25=(np.abs(C-E25)/E25<=0.02)&(C>E25)
osime5=(~osime25)&(np.abs(C-E5)/E5<=0.01)&(C>E5)
ss+=np.where(osime25,2,np.where(osime5,1,0))
ss+=np.where(volratio>=2.0,3,np.where(volratio>=1.5,2,0))
ss+=np.where(L<=YTDL+1e-9,2,0)
ss+=np.where(MACD_L<0,1,np.where(MACD_L>0,-1,0))
cross=(E5<E25); dc=np.zeros((n,m),bool)
for k in range(1,6):
    pb=np.zeros((n,m),bool); pb[:,k:]=~cross[:,:-k]
    sa=np.ones((n,m),bool)
    for q in range(k):
        sh=np.zeros((n,m),bool); sh[:,q:]=cross[:,:m-q] if q>0 else cross
        sa&=sh
    dc|=(pb&sa)
dc[:,:26]=False
pd_=(E5<E25)&(E25<E75)
VOL20P=roll_mean(V,20,False)
vbb=(V>=VOL20P*1.5)&(C<O)
HISTP=np.full((n,m),np.nan); HISTP[:,1:]=HIST[:,:-1]
hebad=(HIST<HISTP)&(HIST<0)
adxok=ADX>25
ss+=dc*2+pd_*2+vbb*2+adxok*1+hebad*1
mom=((EF12<ES26)&(ES26>0)).astype(int)+(MACD_L<MACD_S).astype(int)+(RSI<50).astype(int)
ss+=np.where(mom==3,2,np.where(mom==2,1,0))
ss+=BC+BCP
weak=RET<MKT[None,:]
ss+=weak*1
SELL_OK=(down_tr|pd_)&(RSI>=15)&(RSI<80)&weak&(mom>=2)
np.savez('sell.npz',ss=ss,SELL_OK=SELL_OK,down_tr=down_tr,pd_=pd_,mom=mom)
V2=f['VALID']
print('売りシグナル(スコア11以上・条件通過):',(V2&SELL_OK&(ss>=11)).sum())
print('売り高精度(19以上):',(V2&SELL_OK&(ss>=19)).sum())
print('スコア分位',np.nanpercentile(ss[V2],[50,75,90,95,99]))
