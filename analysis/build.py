import pickle, datetime as dt, numpy as np
s=pickle.load(open('raw.pkl','rb'))
dates_row = s['終値'][2]
dates = [dates_row[i] for i in range(4,254)]
assert all(isinstance(d,dt.datetime) for d in dates)
# chronological ascending
order = list(range(len(dates)))[::-1]
dates_asc = [dates[i] for i in order]

codes=[]; names=[]; rowidx=[]
c=s['終値']
for r in range(5,505):
    code = c[r][0]; nm = c[r][1]
    if code in (None,'',0): continue
    codes.append(str(code).strip()); names.append(nm); rowidx.append(r)
print('stocks', len(codes), 'days', len(dates_asc), dates_asc[0], dates_asc[-1])

def mat(sheet):
    rows=s[sheet]
    M=np.full((len(rowidx), len(dates_asc)), np.nan)
    for a,r in enumerate(rowidx):
        row=rows[r]
        for b,j in enumerate(order):
            v=row[4+j]
            if isinstance(v,(int,float)) and not isinstance(v,bool):
                M[a,b]=float(v)
    return M
O=mat('始値'); H=mat('高値'); L=mat('安値'); C=mat('終値'); V=mat('出来高')
np.savez('data.npz', O=O,H=H,L=L,C=C,V=V, codes=np.array(codes), names=np.array([str(x) for x in names]),
         dates=np.array([d.strftime('%Y-%m-%d') for d in dates_asc]))
print('nan rate C', np.isnan(C).mean(), 'O',np.isnan(O).mean(),'H',np.isnan(H).mean(),'L',np.isnan(L).mean(),'V',np.isnan(V).mean())
ok = (L<=np.minimum(O,C)+1e-9)&(H>=np.maximum(O,C)-1e-9)&(H>=L)
valid=~np.isnan(O)&~np.isnan(H)&~np.isnan(L)&~np.isnan(C)
print('OHLC consistency', ok[valid].mean(), 'valid cells', valid.sum())
# where inconsistent
bad = valid&~ok
print('bad count', bad.sum())
ii,jj=np.where(bad)
for k in range(min(5,len(ii))):
    a,b=ii[k],jj[k]; print(codes[a], dates_asc[b], O[a,b],H[a,b],L[a,b],C[a,b])
