exec(open('sell_base.py').read())
def run_b(k_atr,vd=3,slots=2,t0=80,t1=None,minscore=11):
    t1=t1 or m-12; R=[]
    for t in range(t0,t1):
        cand=np.where(VALID[:,t]&SELL_OK[:,t]&(ss[:,t]>=minscore)&np.isfinite(ATR[:,t]))[0]
        if cand.size==0: continue
        cand=cand[np.argsort(-ss[cand,t])][:TOPN]; used=0
        for i in cand:
            if used>=slots: break
            trg=C[i,t]-k_atr*ATR[i,t]          # 逆指値売り（下抜けで発動）
            for q in range(1,vd+1):
                o,l=O[i,t+q],L[i,t+q]
                if not(np.isfinite(o) and np.isfinite(l)): continue
                if o<=trg or l<=trg:
                    fill=min(o,trg) if o<=trg else trg
                    R.append(exits_s(i,t+q,fill)); used+=1; break
    R=np.array(R); yrs=(t1-t0)/250
    return dict(per_year=len(R)/yrs,avgR=R.mean() if len(R) else 0,
                win=(R>0).mean()*100 if len(R) else 0,tot=R.sum()/yrs,n=len(R))
print('■ 売り：逆指値（下抜け確認）で入る')
print(f'{"戦略":<32s}{"年間取引":>8s}{"平均R":>9s}{"勝率":>8s}{"年間合計R":>10s}')
r=run(0,market=True)
print(f'{"翌日寄付で成行売り（現行）":<26s}{r["per_year"]:8.0f}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["tot"]:+10.1f}')
for kk in [0.15,0.25,0.4,0.5,0.75,1.0]:
    for vd in (1,3):
        r=run_b(kk,vd)
        print(f'{f"逆指値売り 終値-{kk:.2f}ATR ({vd}日)":<32s}{r["per_year"]:8.0f}{r["avgR"]:+9.3f}{r["win"]:7.1f}%{r["tot"]:+10.1f}')
mid=(80+m-12)//2
print('\n■ 期間安定性')
for lbl,fn,kk in [('寄付成行',None,None),('逆指値 -0.25ATR 3日',run_b,0.25),('逆指値 -0.50ATR 3日',run_b,0.5)]:
    if fn is None:
        a=run(0,market=True,t0=80,t1=mid); b=run(0,market=True,t0=mid,t1=m-12)
    else:
        a=fn(kk,3,t0=80,t1=mid); b=fn(kk,3,t0=mid,t1=m-12)
    print(f'  {lbl:<22s} 前半 {a["avgR"]:+.3f}({a["win"]:.1f}%) / 後半 {b["avgR"]:+.3f}({b["win"]:.1f}%)')
print('\n■ 参考：この期間の市場（300銘柄の日次リターン中央値の累積）')
MK=f['MKT']; print(f'  年間 {(np.nancumprod(1+np.nan_to_num(MK))[-1]-1)*100:+.1f}% の上昇相場 → 空売りには構造的に不利な期間')
