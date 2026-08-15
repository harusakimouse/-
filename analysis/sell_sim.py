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
print(f'売りシグナル翌日：高値は終値から何ATR上か  n={k.size:,}')
for p in [10,25,50,75,90]: print(f'   P{p:<3d} {np.percentile(k,p):+5.2f} ATR')
print(f'   終値超え率 {(k>0).mean()*100:.1f}%')
up=(H1/O1-1)[SIG]*100; up=up[np.isfinite(up)]
print(f'\n寄付で空売りした場合の当日含み損（翌日高値/翌日始値）: 中央 {np.median(up):+.2f}% / 平均 {up.mean():+.2f}% / 一度でも含み損 {(up>0).mean()*100:.1f}%')

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
print('\n■ 売り：毎日スコア上位10候補から最大2銘柄')
print(f'{"戦略":<30s}{"年間取引":>8s}{"平均R":>9s}{"勝率":>8s}{"年間合計R":>10s}')
r=run(0,market=True)
print(f'{"翌日寄付で成行売り（現行）":<24s}{r["per_year"]:8.0f}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["tot"]:+10.1f}')
for kk in [0.25,0.5,0.75,1.0,1.25,1.5]:
    for vd in (1,3):
        r=run(kk,vd)
        print(f'{f"指値売り 終値+{kk:.2f}ATR ({vd}日有効)":<28s}{r["per_year"]:8.0f}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["tot"]:+10.1f}')
mid=(80+m-12)//2
print('\n■ 期間安定性（売り）')
for lbl,kk in [('寄付成行',None),('+0.50ATR 3日',0.5),('+0.75ATR 3日',0.75),('+1.00ATR 3日',1.0)]:
    if kk is None:
        a=run(0,market=True,t0=80,t1=mid); b=run(0,market=True,t0=mid,t1=m-12)
    else:
        a=run(kk,3,t0=80,t1=mid); b=run(kk,3,t0=mid,t1=m-12)
    print(f'  {lbl:<18s} 前半 {a["avgR"]:+.3f}({a["win"]:.1f}%) / 後半 {b["avgR"]:+.3f}({b["win"]:.1f}%)')
