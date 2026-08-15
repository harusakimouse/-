import numpy as np
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C,V=d['O'],d['H'],d['L'],d['C'],d['V']; n,m=C.shape
score,VALID,ATR=f['score'],f['VALID'],f['ATR']

HOLD=10; STOP_ATR=2.0; STOP_MIN=0.03; STOP_MAX=0.10; RR=2.0

def simulate(mask, depth, mode='pct', valid_days=1, hold=HOLD):
    """指値を entry_ref*(1-depth) に置いた場合の約定率と期待R。
       mode='pct' → depth は当日終値に対する%、'atr' → depth は ATR倍率。"""
    I,J=np.where(mask)
    keep=J<=m-(valid_days+hold+1)
    I,J=I[keep],J[keep]
    res=dict(n=len(I),filled=0,R=[],ret5=[],miss_ret5=[],fillpx=[],wait=[])
    for i,t in zip(I,J):
        c=C[i,t]; a=ATR[i,t]
        if not (np.isfinite(c) and np.isfinite(a) and c>0): continue
        lim = c*(1-depth) if mode=='pct' else c-depth*a
        fill=None; fday=None
        for k in range(1,valid_days+1):
            o,l=O[i,t+k],L[i,t+k]
            if not np.isfinite(o) or not np.isfinite(l): continue
            if o<=lim: fill,fday=o,t+k; break        # 寄付が指値より下 → 寄付で約定
            if l<=lim: fill,fday=lim,t+k; break
        if fill is None:
            j=t+valid_days
            if np.isfinite(C[i,j+5]) and np.isfinite(C[i,t]):
                res['miss_ret5'].append(C[i,j+5]/C[i,t]-1)
            continue
        res['filled']+=1; res['fillpx'].append(fill/c-1); res['wait'].append(fday-t)
        sd=min(max(STOP_ATR*a, fill*STOP_MIN), fill*STOP_MAX)
        stop=fill-sd; tgt=fill+RR*sd
        R=None
        for k in range(fday, min(fday+hold, m-1)+1):
            if not np.isfinite(L[i,k]): continue
            if L[i,k]<=stop: R=-1.0; break
            if H[i,k]>=tgt:  R=RR; break
        if R is None:
            j=min(fday+hold, m-1); R=(C[i,j]-fill)/sd
        res['R'].append(R)
        j5=min(fday+5,m-1)
        if np.isfinite(C[i,j5]): res['ret5'].append(C[i,j5]/fill-1)
    return res

def market(mask, hold=HOLD):
    """翌日寄付で成行買い（現行の運用）"""
    I,J=np.where(mask); keep=J<=m-(hold+2); I,J=I[keep],J[keep]
    R=[];ret5=[]
    for i,t in zip(I,J):
        o=O[i,t+1]; a=ATR[i,t]
        if not (np.isfinite(o) and np.isfinite(a) and o>0): continue
        sd=min(max(STOP_ATR*a,o*STOP_MIN),o*STOP_MAX); stop=o-sd; tgt=o+RR*sd
        r=None
        for k in range(t+1,min(t+1+hold,m-1)+1):
            if not np.isfinite(L[i,k]): continue
            if L[i,k]<=stop: r=-1.0;break
            if H[i,k]>=tgt: r=RR;break
        if r is None:
            j=min(t+1+hold,m-1); r=(C[i,j]-o)/sd
        R.append(r); 
        j5=min(t+6,m-1)
        if np.isfinite(C[i,j5]): ret5.append(C[i,j5]/o-1)
    return dict(n=len(R),R=R,ret5=ret5)

def report(mask,title):
    print(f'\n{"="*104}\n{title}   (対象 {int(mask.sum()):,} 銘柄日 / 決済ルール: 損切2ATR[3〜10%] 利確2R 10日時間決済)\n{"="*104}')
    mk=market(mask)
    Rm=np.array(mk['R'])
    print(f'{"エントリー方法":<26s}{"約定率":>7s}{"平均約定価":>10s}{"期待R(約定時)":>13s}{"勝率":>7s}{"総期待R":>9s}{"5日後":>8s}')
    print(f'{"翌日寄付で成行(現行)":<24s}{"100.0%":>9s}{"±0.00%":>11s}{np.mean(Rm):+12.3f}{(Rm>0).mean()*100:6.1f}%{np.mean(Rm):+9.3f}{np.mean(mk["ret5"])*100:+7.2f}%')
    out=[('翌日寄付成行',1.0,0.0,float(np.mean(Rm)),float((Rm>0).mean()),float(np.mean(Rm)))]
    for mode,grid,fmt in [('pct',[0.005,0.01,0.015,0.02,0.025,0.03,0.04,0.05],'{:.1f}%下'),
                          ('atr',[0.25,0.5,0.75,1.0,1.25,1.5,2.0],'ATR×{:.2f}下')]:
        print('  ' + '-'*100)
        for dep in grid:
            r=simulate(mask,dep,mode,valid_days=1)
            if r['filled']<30: continue
            R=np.array(r['R']); fr=r['filled']/r['n']
            lbl=fmt.format(dep*100 if mode=='pct' else dep)
            print(f'{"指値 "+lbl:<26s}{fr*100:6.1f}%{np.mean(r["fillpx"])*100:+10.2f}%{np.mean(R):+12.3f}{(R>0).mean()*100:6.1f}%{np.mean(R)*fr:+9.3f}{np.mean(r["ret5"])*100:+7.2f}%')
            out.append((lbl,fr,float(np.mean(r['fillpx'])),float(np.mean(R)),float((R>0).mean()),float(np.mean(R)*fr)))
    return out

BASE=VALID&np.isfinite(O)&np.isfinite(ATR)
o11=report(BASE&(score>=11),'【A】v13シグナル(スコア11以上)  ─ 指値を何%下に置くか')
o16=report(BASE&(score>=16),'【B】v13高精度シグナル(スコア16以上)')
np.save('sim11.npy',np.array(o11,dtype=object)); np.save('sim16.npy',np.array(o16,dtype=object))
