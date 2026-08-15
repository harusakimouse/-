import numpy as np, json
d = np.load('data.npz', allow_pickle=True); f = np.load('feat.npz')
O,H,L,C,V = d['O'],d['H'],d['L'],d['C'],d['V']; codes,names,dates = d['codes'],d['names'],d['dates']
n,m = C.shape
score,VALID,ATR,MKT = f['score'],f['VALID'],f['ATR'],f['MKT']
volratio,RET = f['volratio'],f['RET']
E5,E25 = f['E5'],f['E25']

def shift_fwd(X,k):
    Y=np.full_like(X,np.nan); 
    if k<X.shape[1]: Y[:,:m-k]=X[:,k:]
    return Y
O1,H1,L1,C1 = [shift_fwd(x,1) for x in (O,H,L,C)]

def pct(a,ps=(5,10,25,50,75,90,95)):
    a=a[np.isfinite(a)]
    return {f'p{p}': float(np.percentile(a,p)) for p in ps}

def desc(a,label):
    a=a[np.isfinite(a)]
    q=np.percentile(a,[10,25,50,75,90])
    return dict(label=label,n=int(a.size),mean=float(a.mean()),
                p10=float(q[0]),p25=float(q[1]),p50=float(q[2]),p75=float(q[3]),p90=float(q[4]))

BASE = VALID & np.isfinite(O1)&np.isfinite(H1)&np.isfinite(L1)&np.isfinite(C1)
SIG11 = BASE & (score>=11)
SIG16 = BASE & (score>=16)
print(f'母集団: 全 {BASE.sum():,} 銘柄日 / シグナル(>=11) {SIG11.sum():,} / 高精度(>=16) {SIG16.sum():,}')
print(f'期間 {dates[0]} 〜 {dates[-1]}  {m}営業日 {n}銘柄\n')

def block(mask, title):
    gap   = (O1/C-1)[mask]*100                # 翌日ギャップ
    lowc  = (L1/C-1)[mask]*100                # 翌日安値 vs 当日終値
    lowo  = (L1/O1-1)[mask]*100               # 翌日安値 vs 翌日始値（買った日の含み損）
    cloo  = (C1/O1-1)[mask]*100               # 翌日終値 vs 翌日始値
    cloc  = (C1/C-1)[mask]*100
    hio   = (H1/O1-1)[mask]*100
    rows=[desc(gap,'翌日ギャップ (翌日始値/当日終値)'),
          desc(lowc,'翌日安値 / 当日終値'),
          desc(lowo,'翌日安値 / 翌日始値 ★寄付買いの当日含み損'),
          desc(hio, '翌日高値 / 翌日始値'),
          desc(cloo,'翌日終値 / 翌日始値'),
          desc(cloc,'翌日終値 / 当日終値')]
    print(f'--- {title} (n={int(mask.sum()):,}) ---')
    print(f'{"":44s} {"平均":>7s} {"P10":>7s} {"P25":>7s} {"中央":>7s} {"P75":>7s} {"P90":>7s}')
    for r in rows:
        print(f'{r["label"]:44s} {r["mean"]:+7.2f} {r["p10"]:+7.2f} {r["p25"]:+7.2f} {r["p50"]:+7.2f} {r["p75"]:+7.2f} {r["p90"]:+7.2f}')
    print(f'  寄付買いが当日一度でもマイナスになる率: {(lowo<0).mean()*100:5.1f}%')
    print(f'  寄付買いが当日 -1%以上下げる率        : {(lowo<=-1).mean()*100:5.1f}%')
    print(f'  寄付買いが当日 -2%以上下げる率        : {(lowo<=-2).mean()*100:5.1f}%')
    print(f'  翌日陰線率(終値<始値)                 : {(cloo<0).mean()*100:5.1f}%')
    print(f'  ギャップダウン率                      : {(gap<0).mean()*100:5.1f}%')
    print()
    return rows

r_all = block(BASE,  '母集団すべての銘柄日')
r_11  = block(SIG11, 'v13シグナル日 (スコア11以上)')
r_16  = block(SIG16, 'v13高精度シグナル日 (スコア16以上)')
