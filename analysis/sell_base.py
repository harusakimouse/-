import numpy as np
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz'); s=np.load('sell.npz')
O,H,L,C=d['O'],d['H'],d['L'],d['C']; n,m=C.shape
ATR,VALID=f['ATR'],f['VALID']; ss,SELL_OK=s['ss'],s['SELL_OK']
HOLD=10;RR=2.0;TOPN=10
SIG=VALID&SELL_OK&(ss>=11)&np.isfinite(ATR)&(ATR>0)
H1=np.full((n,m),np.nan); H1[:,:m-1]=H[:,1:]
O1=np.full((n,m),np.nan); O1[:,:m-1]=O[:,1:]
K=(H1-C)/ATR
k=K[SIG&np.isfinite(K)]
up=(H1/O1-1)[SIG]*100; up=up[np.isfinite(up)]

def exits_s(i,start,fill):
    a=ATR[i,start-1] if np.isfinite(ATR[i,start-1]) else ATR[i,start]
    sd=min(max(2.0*a,fill*0.03),fill*0.10); stop=fill+sd; tgt=fill-RR*sd
    for k2 in range(start,min(start+HOLD,m-1)+1):
        if not np.isfinite(H[i,k2]): continue
        if H[i,k2]>=stop: return -1.0
        if L[i,k2]<=tgt: return RR
    j=min(start+HOLD,m-1); return (fill-C[i,j])/sd

def run(k_atr,vd=3,slots=2,market=False,t0=80,t1=None,minscore=11):
    t1=t1 or m-12; R=[]
    for t in range(t0,t1):
        cand=np.where(VALID[:,t]&SELL_OK[:,t]&(ss[:,t]>=minscore)&np.isfinite(ATR[:,t]))[0]
        if cand.size==0: continue
        cand=cand[np.argsort(-ss[cand,t])][:TOPN]; used=0
        for i in cand:
            if used>=slots: break
            if market:
                o=O[i,t+1]
                if not np.isfinite(o): continue
                R.append(exits_s(i,t+1,o)); used+=1
            else:
                lim=C[i,t]+k_atr*ATR[i,t]
                for q in range(1,vd+1):
                    o,h=O[i,t+q],H[i,t+q]
                    if not(np.isfinite(o) and np.isfinite(h)): continue
                    if o>=lim or h>=lim:
                        fill=max(o,lim) if o>=lim else lim
                        R.append(exits_s(i,t+q,fill)); used+=1; break
    R=np.array(R); yrs=(t1-t0)/250
    return dict(n=len(R),per_year=len(R)/yrs,avgR=R.mean() if len(R) else 0,
                win=(R>0).mean()*100 if len(R) else 0, tot=R.sum()/yrs)
