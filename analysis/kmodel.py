import numpy as np
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C,V=d['O'],d['H'],d['L'],d['C'],d['V']; n,m=C.shape
score,VALID,ATR,MKT,volratio,RET=f['score'],f['VALID'],f['ATR'],f['MKT'],f['volratio'],f['RET']
E5,E25=f['E5'],f['E25']
L1=np.full((n,m),np.nan); L1[:,:m-1]=L[:,1:]
SIG=VALID&np.isfinite(L1)&np.isfinite(ATR)&(score>=11)&(ATR>0)
K=(C-L1)/ATR          # 翌日安値は終値から何ATR下か
k=K[SIG]; k=k[np.isfinite(k)]
print(f'翌日安値までの下落幅（ATR倍数） n={k.size:,}')
for p in [10,25,40,50,60,75,90]:
    print(f'  P{p:<3d} = {np.percentile(k,p):+5.2f} ATR')
print(f'  平均 {k.mean():+.2f} / 0ATR以上(=終値割れ)率 {(k>0).mean()*100:.1f}%\n')

atr_pct=ATR/C*100; dev25=(C/E25-1)*100; chg=RET*100
print('■ k(=下落ATR倍数)の中央値・P75 が各要因でどう変わるか')
def tab(name,X,edges,unit):
    print(f'\n  ◆ {name}')
    for lo,hi in zip(edges[:-1],edges[1:]):
        s=SIG&(X>=lo)&(X<hi)&np.isfinite(K)
        kk=K[s]; kk=kk[np.isfinite(kk)]
        if kk.size<40: continue
        print(f'    {lo:>6.1f}〜{hi:<6.1f}{unit} n={kk.size:5d}  中央 k={np.median(kk):+5.2f}  P75 k={np.percentile(kk,75):+5.2f}  P90 k={np.percentile(kk,90):+5.2f}')
tab('対前日比',chg,[-99,0,2,4,6,9,99],'%')
tab('出来高倍率',volratio,[0,1.2,1.6,2.2,3.5,99],'倍')
tab('25日EMA乖離',dev25,[-99,3,6,10,15,99],'%')
tab('ATR率',atr_pct,[0,2,3,4,6,99],'%')
tab('市場全体',np.tile(MKT,(n,1))*100,[-99,-0.5,0.5,99],'%')

# --- 単純加算モデル ---
base=0.35
adj = (np.where(chg>=9,0.55,np.where(chg>=6,0.35,np.where(chg>=4,0.20,np.where(chg>=2,0.05,0.0))))
     + np.where(volratio>=3.5,0.35,np.where(volratio>=2.2,0.20,np.where(volratio>=1.6,0.10,0.0)))
     + np.where(dev25>=15,0.30,np.where(dev25>=10,0.20,np.where(dev25>=6,0.10,0.0))))
KPRED=base+adj
print('\n\n■ 推奨指値モデル: 指値 = 当日終値 − k×ATR,  k = 0.35 + 過熱度調整')
for lo,hi in [(0.35,0.36),(0.36,0.55),(0.55,0.75),(0.75,0.95),(0.95,9)]:
    s=SIG&(KPRED>=lo)&(KPRED<hi)
    kk=K[s]; kk=kk[np.isfinite(kk)]
    if kk.size<40: continue
    hit=(kk>=np.median(KPRED[s]))
    print(f'  予測k {lo:.2f}〜{hi:.2f}: n={kk.size:5d}  実際の中央k={np.median(kk):+5.2f}  約定率(実k≧予測k)={hit.mean()*100:5.1f}%')
np.save('KPRED.npy',KPRED); np.save('K.npy',K)
