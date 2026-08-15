import numpy as np
exec(open('portfolio.py').read().split("print('毎日スコア")[0])

def run2(k_atr, valid_days, slots, t0, t1, minscore=11):
    R=[];MAE=[];imm=[]
    for t in range(t0, t1):
        cand=np.where(VALID[:,t]&(score[:,t]>=minscore)&np.isfinite(ATR[:,t])&np.isfinite(C[:,t]))[0]
        if cand.size==0: continue
        cand=cand[np.argsort(-score[cand,t])][:TOPN]; used=0
        for i in cand:
            if used>=slots: break
            if k_atr is None:
                o=O[i,t+1]
                if not np.isfinite(o): continue
                r,mae,_=exits(i,t+1,o); R.append(r);MAE.append(mae);used+=1
                if np.isfinite(C[i,t+1]): imm.append(C[i,t+1]>=o)
            else:
                lim=C[i,t]-k_atr*ATR[i,t]
                for q in range(1,valid_days+1):
                    o,l=O[i,t+q],L[i,t+q]
                    if not (np.isfinite(o) and np.isfinite(l)): continue
                    if o<=lim or l<=lim:
                        fill=min(o,lim) if o<=lim else lim
                        r,mae,_=exits(i,t+q,fill); R.append(r);MAE.append(mae);used+=1
                        if np.isfinite(C[i,t+q]): imm.append(C[i,t+q]>=fill)
                        break
    R=np.array(R)
    return len(R), (R.mean() if len(R) else 0), ((R>0).mean()*100 if len(R) else 0), (np.mean(imm)*100 if imm else 0)

mid=(80+m-12)//2
print('■ 期間安定性（前半／後半で分割）')
print(f'{"戦略":<26s}{"前半 平均R":>12s}{"後半 平均R":>12s}{"前半 勝率":>11s}{"後半 勝率":>11s}')
for lbl,k,vd in [('翌日寄付成行(現行)',None,1),('指値-0.50ATR 3日',0.5,3),('指値-0.75ATR 3日',0.75,3),('指値-1.00ATR 3日',1.0,3)]:
    a=run2(k,vd,2,80,mid); b=run2(k,vd,2,mid,m-12)
    print(f'{lbl:<24s}{a[1]:+12.3f}{b[1]:+12.3f}{a[2]:10.1f}%{b[2]:10.1f}%')

print('\n■ 建玉数（スロット）を変えた場合の平均R')
print(f'{"戦略":<26s}{"1銘柄":>9s}{"2銘柄":>9s}{"3銘柄":>9s}{"5銘柄":>9s}')
for lbl,k,vd in [('翌日寄付成行(現行)',None,1),('指値-0.50ATR 3日',0.5,3),('指値-0.75ATR 3日',0.75,3),('指値-1.00ATR 3日',1.0,3)]:
    vals=[run2(k,vd,s,80,m-12)[1] for s in (1,2,3,5)]
    print(f'{lbl:<24s}'+''.join(f'{v:+9.3f}' for v in vals))

print('\n■ 「買ったその日に含み益で終わる」率（心理的負担の指標）')
for lbl,k,vd in [('翌日寄付成行(現行)',None,1),('指値-0.50ATR 3日',0.5,3),('指値-0.75ATR 3日',0.75,3),('指値-1.00ATR 3日',1.0,3),('指値-1.50ATR 3日',1.5,3)]:
    r=run2(k,vd,2,80,m-12)
    print(f'  {lbl:<24s} n={r[0]:4d}  当日プラス引け率 {r[3]:5.1f}%   平均R {r[1]:+.3f}')

print('\n■ 高精度シグナル(スコア16以上)に限定')
print(f'{"戦略":<26s}{"取引数":>8s}{"平均R":>9s}{"勝率":>8s}')
for lbl,k,vd in [('翌日寄付成行(現行)',None,1),('指値-0.50ATR 3日',0.5,3),('指値-0.75ATR 3日',0.75,3),('指値-1.00ATR 3日',1.0,3),('指値-1.25ATR 3日',1.25,3)]:
    r=run2(k,vd,2,80,m-12,minscore=16)
    print(f'{lbl:<24s}{r[0]:8d}{r[1]:+9.3f}{r[2]:7.1f}%')
