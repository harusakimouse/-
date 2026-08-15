import numpy as np
d=np.load('data.npz',allow_pickle=True); f=np.load('feat.npz')
O,H,L,C,V=d['O'],d['H'],d['L'],d['C'],d['V']; n,m=C.shape
score,VALID,ATR,MKT,volratio,RET=f['score'],f['VALID'],f['ATR'],f['MKT'],f['volratio'],f['RET']
E5,E25,CS_T,CL_T=f['E5'],f['E25'],f['CS_T'],f['CL_T']

TP=(H+L+C)/3.0
def vwap(w):
    R=np.full((n,m),np.nan)
    for j in range(w-1,m):
        num=np.nansum(TP[:,j-w+1:j+1]*V[:,j-w+1:j+1],axis=1)
        den=np.nansum(V[:,j-w+1:j+1],axis=1)
        R[:,j]=np.where(den>0,num/den,np.nan)
    return R
VW5,VW25=vwap(5),vwap(25)
L1=np.full((n,m),np.nan); L1[:,:m-1]=L[:,1:]
O1=np.full((n,m),np.nan); O1[:,:m-1]=O[:,1:]
Lprev=np.full((n,m),np.nan); Lprev[:,1:]=L[:,:-1]

BASE=VALID&np.isfinite(L1)&np.isfinite(ATR)&np.isfinite(VW25)&np.isfinite(E25)
SIG=BASE&(score>=11)
print(f'翌日安値の着地点分析  n={int(SIG.sum()):,}\n')
anchors={
 '当日終値':C, '当日安値':L, '前日安値':Lprev,
 '5日EMA':E5, '25日EMA':E25, '5日VWAP':VW5, '25日VWAP':VW25,
 '終値-0.5ATR':C-0.5*ATR, '終値-1.0ATR':C-1.0*ATR, '終値-1.5ATR':C-1.5*ATR, '終値-2.0ATR':C-2.0*ATR,
 '(終値+安値)/2':(C+L)/2,
}
print(f'{"支持線候補":<16s}{"終値比":>8s}{"翌日安値がそこまで到達する率":>16s}{"翌日安値/その水準 中央値":>18s}')
for k,A in anchors.items():
    a=A[SIG]; l1=L1[SIG]; c=C[SIG]
    ok=np.isfinite(a)&np.isfinite(l1)&(a>0)
    hit=(l1[ok]<=a[ok])
    print(f'{k:<16s}{np.nanmedian(a[ok]/c[ok]-1)*100:+7.2f}%{hit.mean()*100:15.1f}%{np.nanmedian(l1[ok]/a[ok]-1)*100:+17.2f}%')

print('\n\n■ 翌日の下げ幅を決める要因（シグナル日=スコア11以上、翌日安値/翌日始値 の中央値）')
dd=(L1/O1-1)*100
def buckets(name, X, edges, unit=''):
    print(f'\n  ◆ {name}')
    x=X[SIG]; y=dd[SIG]; ok=np.isfinite(x)&np.isfinite(y)
    x,y=x[ok],y[ok]
    for lo,hi in zip(edges[:-1],edges[1:]):
        s=(x>=lo)&(x<hi)
        if s.sum()<40: continue
        print(f'    {lo:>6.1f}〜{hi:<6.1f}{unit}  n={s.sum():5d}  当日含み損 中央 {np.median(y[s]):+6.2f}%  P25 {np.percentile(y[s],25):+6.2f}%  -2%超え率 {(y[s]<=-2).mean()*100:5.1f}%')
buckets('シグナル日の対前日比（飛びつき度合い）', RET*100, [-99,-2,0,2,4,6,9,99], '%')
buckets('出来高倍率(5日平均比)', volratio, [0,0.8,1.2,1.6,2.2,3.5,99], '倍')
buckets('ATR率(ATR/終値=ボラティリティ)', ATR/C*100, [0,1.5,2.5,3.5,5,8,99], '%')
buckets('25日EMAからの乖離率', (C/E25-1)*100, [-99,-2,0,3,6,10,99], '%')
buckets('25日VWAPからの乖離率', (C/VW25-1)*100, [-99,-2,0,3,6,10,99], '%')
buckets('当日の市場全体(300銘柄中央値)', np.tile(MKT,(n,1))*100, [-99,-1.5,-0.5,0.5,1.5,99], '%')
print('\n  ◆ シグナル日のローソク足の形')
for lab,v in [('下ヒゲ陽線',1),('大陽線',2),('長い下ヒゲ',3),('上ヒゲ警戒',4),('その他',0)]:
    s=SIG&(CL_T==v); y=dd[s]; y=y[np.isfinite(y)]
    if y.size<40: continue
    print(f'    {lab:<10s} n={y.size:5d}  当日含み損 中央 {np.median(y):+6.2f}%  P25 {np.percentile(y,25):+6.2f}%  -2%超え率 {(y<=-2).mean()*100:5.1f}%')
print('\n  ◆ 翌日のギャップ別（寄り付き後の値動き）')
gap=(O1/C-1)*100
for lo,hi in [(-99,-2),(-2,-0.5),(-0.5,0.5),(0.5,2),(2,5),(5,99)]:
    s=SIG&(gap>=lo)&(gap<hi); y=dd[s]; y=y[np.isfinite(y)]
    if y.size<40: continue
    print(f'    ギャップ{lo:+5.1f}〜{hi:+5.1f}%  n={y.size:5d}  寄付から更に 中央 {np.median(y):+6.2f}%  P25 {np.percentile(y,25):+6.2f}%')
