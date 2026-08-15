import numpy as np
exec(open('portfolio.py').read().split("print('毎日スコア")[0])
RET,volratio,E25=f['RET'],f['volratio'],f['E25']
dev25=(C/E25-1)*100; chg=RET*100

def run3(kfun, valid_days=3, slots=2, skip_hot=False, t0=80, t1=None, minscore=11):
    t1 = t1 or m-12
    R=[];fills=0;offered=0;imm=[]
    for t in range(t0,t1):
        cand=np.where(VALID[:,t]&(score[:,t]>=minscore)&np.isfinite(ATR[:,t])&np.isfinite(C[:,t]))[0]
        if cand.size==0: continue
        cand=cand[np.argsort(-score[cand,t])][:TOPN]; used=0
        for i in cand:
            if used>=slots: break
            if skip_hot and (chg[i,t]>=9 or dev25[i,t]>=25): continue
            offered+=1
            lim=C[i,t]-kfun(i,t)*ATR[i,t]
            for q in range(1,valid_days+1):
                o,l=O[i,t+q],L[i,t+q]
                if not(np.isfinite(o) and np.isfinite(l)): continue
                if o<=lim or l<=lim:
                    fill=min(o,lim) if o<=lim else lim
                    r,mae,_=exits(i,t+q,fill); R.append(r);used+=1;fills+=1
                    if np.isfinite(C[i,t+q]): imm.append(C[i,t+q]>=fill)
                    break
    R=np.array(R)
    return dict(n=len(R),avgR=R.mean() if len(R) else 0,win=(R>0).mean()*100 if len(R) else 0,
                total=R.sum(), imm=np.mean(imm)*100 if imm else 0)

flat=lambda k:(lambda i,t:k)
def hot(i,t):
    k=0.75
    if chg[i,t]>=6 or volratio[i,t]>=2.2 or dev25[i,t]>=15: k+=0.25
    if chg[i,t]>=9 or volratio[i,t]>=3.5 or dev25[i,t]>=25: k+=0.25
    return k
print('■ 過熱度に応じて指値を深くする効果（3日有効・上位10候補から2銘柄）')
print(f'{"ルール":<38s}{"取引数":>7s}{"平均R":>9s}{"勝率":>8s}{"年間合計R":>10s}{"当日プラス率":>11s}')
for lbl,kf,sk in [('一律 −0.50ATR',flat(0.5),False),
                  ('一律 −0.75ATR',flat(0.75),False),
                  ('一律 −1.00ATR',flat(1.0),False),
                  ('過熱度連動 −0.75〜−1.25ATR',hot,False),
                  ('過熱度連動 ＋ 過熱銘柄は見送り',hot,True)]:
    r=run3(kf,skip_hot=sk); yrs=(m-92)/250
    print(f'{lbl:<36s}{r["n"]:7d}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["total"]/yrs:+10.1f}{r["imm"]:10.1f}%')

print('\n■ 指値の有効日数')
print(f'{"ルール":<38s}{"取引数":>7s}{"平均R":>9s}{"勝率":>8s}{"年間合計R":>10s}')
for vd in [1,2,3,5]:
    r=run3(hot,valid_days=vd); yrs=(m-92)/250
    print(f'{f"過熱度連動 / {vd}営業日有効":<36s}{r["n"]:7d}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["total"]/yrs:+10.1f}')

print('\n■ 前半／後半での安定性（過熱度連動 3日有効）')
mid=(80+m-12)//2
for lbl,kf in [('翌日寄付成行(現行)',None),('一律 −0.75ATR',flat(0.75)),('過熱度連動',hot)]:
    if kf is None:
        a=run(0,market=True); print(f'  {lbl:<24s} 全期間 平均R {a["avgR"]:+.3f} / 勝率 {a["winrate"]:.1f}%'); continue
    x=run3(kf,t0=80,t1=mid); y=run3(kf,t0=mid,t1=m-12)
    print(f'  {lbl:<24s} 前半 {x["avgR"]:+.3f}({x["win"]:.1f}%) / 後半 {y["avgR"]:+.3f}({y["win"]:.1f}%)')
