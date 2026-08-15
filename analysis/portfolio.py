import numpy as np
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C=d['O'],d['H'],d['L'],d['C']; n,m=C.shape
score,VALID,ATR=f['score'],f['VALID'],f['ATR']
HOLD=10; RR=2.0; TOPN=10; SLOTS=2

def exits(i,start,fill):
    a=ATR[i,start-1] if np.isfinite(ATR[i,start-1]) else ATR[i,start]
    sd=min(max(2.0*a, fill*0.03), fill*0.10)
    stop, tgt = fill-sd, fill+RR*sd
    mae=0.0
    for k in range(start, min(start+HOLD, m-1)+1):
        if not np.isfinite(L[i,k]): continue
        mae=min(mae,(L[i,k]-fill)/fill)
        if L[i,k]<=stop: return -1.0, mae, k-start
        if H[i,k]>=tgt:  return RR, mae, k-start
    j=min(start+HOLD,m-1)
    return (C[i,j]-fill)/sd, mae, j-start

def run(k_atr, valid_days=1, slots=SLOTS, market=False):
    R=[];MAE=[];HOLDD=[];days_traded=0
    for t in range(80, m-HOLD-valid_days-1):
        cand=np.where(VALID[:,t]&(score[:,t]>=11)&np.isfinite(ATR[:,t])&np.isfinite(C[:,t]))[0]
        if cand.size==0: continue
        cand=cand[np.argsort(-score[cand,t])][:TOPN]
        used=0; got=False
        for i in cand:
            if used>=slots: break
            if market:
                o=O[i,t+1]
                if not np.isfinite(o): continue
                r,mae,hd=exits(i,t+1,o); R.append(r);MAE.append(mae);HOLDD.append(hd);used+=1;got=True
            else:
                lim=C[i,t]-k_atr*ATR[i,t]
                for q in range(1,valid_days+1):
                    o,l=O[i,t+q],L[i,t+q]
                    if not (np.isfinite(o) and np.isfinite(l)): continue
                    if o<=lim or l<=lim:
                        fill=min(o,lim) if o<=lim else lim
                        r,mae,hd=exits(i,t+q,fill); R.append(r);MAE.append(mae);HOLDD.append(hd);used+=1;got=True
                        break
        if got: days_traded+=1
    R=np.array(R);MAE=np.array(MAE)
    yrs=(m-80-HOLD)/250
    return dict(n=len(R),per_year=len(R)/yrs,avgR=R.mean() if len(R) else 0,
                winrate=(R>0).mean()*100 if len(R) else 0, totalR=R.sum(),
                totalR_yr=R.sum()/yrs, mae=MAE.mean()*100 if len(R) else 0,
                mae_p90=np.percentile(MAE,10)*100 if len(R) else 0, hold=np.mean(HOLDD) if len(R) else 0)

print('毎日スコア上位10候補 → 最大2銘柄建てる運用（250営業日 ≒ 1年）')
print(f'{"戦略":<30s}{"年間取引":>8s}{"平均R":>8s}{"勝率":>8s}{"年間合計R":>10s}{"平均含み損":>10s}{"含み損P90":>10s}{"平均保有":>8s}')
r=run(0,market=True)
print(f'{"翌日寄付で成行（現行）":<26s}{r["per_year"]:8.0f}{r["avgR"]:+8.3f}{r["winrate"]:7.1f}%{r["totalR_yr"]:+10.1f}{r["mae"]:9.2f}%{r["mae_p90"]:9.2f}%{r["hold"]:7.1f}日')
print('  '+'-'*100)
for k in [0.25,0.5,0.75,1.0,1.25,1.5]:
    r=run(k,1)
    print(f'{f"指値 終値-{k:.2f}ATR (当日限り)":<28s}{r["per_year"]:8.0f}{r["avgR"]:+8.3f}{r["winrate"]:7.1f}%{r["totalR_yr"]:+10.1f}{r["mae"]:9.2f}%{r["mae_p90"]:9.2f}%{r["hold"]:7.1f}日')
print('  '+'-'*100)
for k in [0.5,0.75,1.0,1.25,1.5]:
    r=run(k,3)
    print(f'{f"指値 終値-{k:.2f}ATR (3日有効)":<28s}{r["per_year"]:8.0f}{r["avgR"]:+8.3f}{r["winrate"]:7.1f}%{r["totalR_yr"]:+10.1f}{r["mae"]:9.2f}%{r["mae_p90"]:9.2f}%{r["hold"]:7.1f}日')
