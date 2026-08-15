import numpy as np, json
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C,V=d['O'],d['H'],d['L'],d['C'],d['V']; codes=list(d['codes']); dates=d['dates']; names=d['names']
n,m=C.shape
ATR,E5,E25,E75,RSI,volratio,score=f['ATR'],f['E5'],f['E25'],f['E75'],f['RSI'],f['volratio'],f['score']
TP=(H+L+C)/3.0
def vwap(w,j):
    num=np.nansum(TP[:,j-w+1:j+1]*V[:,j-w+1:j+1],axis=1); den=np.nansum(V[:,j-w+1:j+1],axis=1)
    return np.where(den>0,num/den,np.nan)
T=m-1
VW5,VW25=vwap(5,T),vwap(25,T)
L1=np.full((n,m),np.nan); L1[:,:m-1]=L[:,1:]
K=(C-L1)/ATR
cands=json.load(open('cands.json'))
def tick(p):
    if p<3000: return 1
    if p<5000: return 5
    if p<30000: return 10
    if p<50000: return 50
    return 100
def rnd(p,down=True):
    t=tick(p); return (np.floor(p/t) if down else np.ceil(p/t))*t
print(f'基準日 {dates[T]}   ラダー: k=ATR倍数、約定率は当該銘柄の過去実績\n')
rows=[]
for c in cands:
    if c['code'] not in codes:
        print(c['code'],c['name'],'→ データなし'); continue
    i=codes.index(c['code'])
    cl,a,e5,e25,e75=C[i,T],ATR[i,T],E5[i,T],E25[i,T],E75[i,T]
    lo,hi,op=L[i,T],H[i,T],O[i,T]
    kk=K[i,:]; kk=kk[np.isfinite(kk)]
    hitrate=lambda k: float((kk>=k).mean()*100)
    atrp=a/cl*100
    rows.append(dict(code=c['code'],name=c['name'],sheet_price=c['price'],close=cl,atr=a,atrp=atrp,
        e5=e5,e25=e25,e75=e75,vw5=VW5[i],vw25=VW25[i],low=lo,high=hi,
        rsi=RSI[i,T],vr=volratio[i,T],score=score[i,T],
        p_light=rnd(cl-0.5*a),p_main=rnd(cl-0.75*a),p_deep=rnd(cl-1.25*a),
        h_light=hitrate(0.5),h_main=hitrate(0.75),h_deep=hitrate(1.25),
        chg=(cl/C[i,T-1]-1)*100, dev25=(cl/e25-1)*100))
    r=rows[-1]
    print(f"【{c['code']}】{c['name']}   8/14終値 {cl:,.0f}円  ATR14 {a:,.0f}円({atrp:.1f}%)  対前日{r['chg']:+.1f}%  25日EMA乖離{r['dev25']:+.1f}%")
    print(f"    第1指値 −0.50ATR = {r['p_light']:>8,.0f}円 ({(r['p_light']/cl-1)*100:+5.2f}%)  この銘柄の到達率 {r['h_light']:4.1f}%")
    print(f"    本命指値 −0.75ATR = {r['p_main']:>8,.0f}円 ({(r['p_main']/cl-1)*100:+5.2f}%)  到達率 {r['h_main']:4.1f}%  ★推奨")
    print(f"    押し目深 −1.25ATR = {r['p_deep']:>8,.0f}円 ({(r['p_deep']/cl-1)*100:+5.2f}%)  到達率 {r['h_deep']:4.1f}%")
    print(f"    支持線: 5日EMA {e5:,.0f} / 5日VWAP {VW5[i]:,.0f} / 25日EMA {e25:,.0f} / 25日VWAP {VW25[i]:,.0f} / 当日安値 {lo:,.0f}")
    print(f"    撤退ライン(25日EMA割れ) {e25:,.0f}円未満は見送り\n")
json.dump(rows,open('cand_rows.json','w'),ensure_ascii=False,indent=1,default=float)
