import numpy as np
exec(open('portfolio.py').read().split("print('毎日スコア")[0])
RET=f['RET']
K=0.75; VD=3
miss=[];fill_adv=[];fill_days=[]
for t in range(80,m-12):
    cand=np.where(VALID[:,t]&(score[:,t]>=11)&np.isfinite(ATR[:,t])&np.isfinite(C[:,t]))[0]
    if cand.size==0: continue
    cand=cand[np.argsort(-score[cand,t])][:TOPN]
    for i in cand:
        lim=C[i,t]-K*ATR[i,t]; got=False
        for q in range(1,VD+1):
            o,l=O[i,t+q],L[i,t+q]
            if not(np.isfinite(o) and np.isfinite(l)): continue
            if o<=lim or l<=lim:
                fill=min(o,lim) if o<=lim else lim
                fill_adv.append(fill/lim-1); fill_days.append(q); got=True; break
        if not got:
            j=min(t+VD+5,m-1)
            if np.isfinite(C[i,j]): miss.append(C[i,j]/C[i,t]-1)
fa=np.array(fill_adv); fd=np.array(fill_days); ms=np.array(miss)
print(f'指値 −0.75ATR / 3営業日有効  （候補ベース n={len(fa)+len(ms):,}）')
print(f'  約定率            : {len(fa)/(len(fa)+len(ms))*100:.1f}%')
print(f'  約定日の内訳      : 翌日 {(fd==1).mean()*100:.0f}% / 2日目 {(fd==2).mean()*100:.0f}% / 3日目 {(fd==3).mean()*100:.0f}%')
print(f'  ギャップダウンで指値より安く約定した割合: {(fa<-1e-9).mean()*100:.1f}% (平均 {fa[fa<0].mean()*100:+.2f}%お得)')
print(f'  未約定銘柄のその後(8日後)の株価: 平均 {ms.mean()*100:+.2f}% / 中央 {np.median(ms)*100:+.2f}% / 上昇した割合 {(ms>0).mean()*100:.1f}%')
print(f'  → 取り逃がした銘柄も平均{ms.mean()*100:+.1f}%で、「大相場を逃す」損失は限定的')
print()
r_now=run(0,market=True); r_new=run(0.75,3)
yrs=(m-80-HOLD)/250
RISK=40000
print('■ 最終比較（1回の許容損失 40,000円 = 1R として金額換算）')
for lbl,r in [('現行: 翌日寄付で成行',r_now),('推奨: 終値−0.75ATR 指値 3日有効',r_new)]:
    print(f'  {lbl:<32s} 年間{r["per_year"]:.0f}取引  平均{r["avgR"]:+.3f}R  勝率{r["winrate"]:.1f}%  年間{r["totalR_yr"]:+.0f}R = {r["totalR_yr"]*RISK/10000:+,.0f}万円')
print(f'  差分: 平均Rで {r_new["avgR"]/r_now["avgR"]:.2f}倍、年間損益で {(r_new["totalR_yr"]-r_now["totalR_yr"])*RISK/10000:+,.0f}万円')
