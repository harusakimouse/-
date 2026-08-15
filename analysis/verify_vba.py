# Mod_推奨指値.bas と同一ロジックを Python で再現し、各シートの出力を確認する
import numpy as np, openpyxl
F='/root/.claude/uploads/c8c951a1-5f75-54ea-97b2-454c76055775/49421f7c-______V1.xlsm'
wb=openpyxl.load_workbook(F,data_only=True)
HI,LO,CL=wb['高値'],wb['安値'],wb['終値']
LOOK=120; P=14
def hist_last_col(ws):
    c=ws.max_column
    while c>=5 and ws.cell(3,c).value is None: c-=1
    return c
LAST=hist_last_col(CL); print('終値シート 最終列 =',LAST,'（E列から',LAST-4,'日分）')
def norm(v):
    if v is None: return ''
    s=str(v).strip()
    if s in ('','0') or s.upper()=='TOPX': return ''
    try:
        if len(s)<4: s='%04d'%int(float(s))
    except: pass
    return s
IDX={}
r=6
while r<=CL.max_row:
    c=norm(CL.cell(r,1).value)
    if c and c not in IDX: IDX[c]=r
    r+=1
print('終値シートの銘柄数 =',len(IDX))

def load_bars(src):
    maxN=min(LAST-5+1, LOOK+P+10); last=5+maxN-1
    H=[];L=[];C=[]
    for k in range(last,4,-1):
        h,l,c=HI.cell(src,k).value,LO.cell(src,k).value,CL.cell(src,k).value
        try: h,l,c=float(h),float(l),float(c)
        except: continue
        if h>0 and l>0 and c>0 and h>=l: H.append(h);L.append(l);C.append(c)
    n=len(H)
    if n<P+2: return None
    A=[0.0]*n; a=None
    for k in range(1,n):
        tr=max(H[k]-L[k],abs(H[k]-C[k-1]),abs(L[k]-C[k-1]))
        a=tr if a is None else (a*(P-1)+tr)/P
        A[k]=a
    return H,L,C,A,n
def tick(p): return 1 if p<3000 else 5 if p<5000 else 10 if p<30000 else 50 if p<50000 else 100
def tdown(p): t=tick(p); return int(p/t)*t
def tup(p):   t=tick(p); return -int(-p/t)*t
def hit_down(H,L,C,A,n,k):
    f=max(n-LOOK,P+1); tot=hits=0
    for i in range(f,n-1):
        if A[i]>0:
            tot+=1
            if L[i+1]<=C[i]-k*A[i]: hits+=1
    return hits/tot*100 if tot else 0
def hit_up(H,L,C,A,n,k):
    f=max(n-LOOK,P+1); tot=hits=0
    for i in range(f,n-1):
        if A[i]>0:
            tot+=1
            if H[i+1]>=C[i]+k*A[i]: hits+=1
    return hits/tot*100 if tot else 0
def stopd(e,a): return min(max(2*a,e*0.03),e*0.10)

for sheet,isbuy in [('買抽出v13',True),('売抽出v13',False)]:
    ws=wb[sheet]
    print(f'\n{"="*100}\n■ {sheet}  → U〜AC列に出力\n{"="*100}')
    if isbuy:
        print(f'{"コード":<6s}{"銘柄":<16s}{"基準終値":>9s}{"ATR14":>7s}{"第1 -0.5":>9s}{"★本命":>9s}{"深押し":>9s}{"到達率":>7s}{"損切":>8s}{"利確":>8s}  ひとこと')
    else:
        print(f'{"コード":<6s}{"銘柄":<16s}{"基準終値":>9s}{"ATR14":>7s}{"戻り+0.39":>10s}{"戻り+0.70":>10s}{"★寄付":>9s}{"到達率":>7s}{"損切":>8s}{"利確":>8s}  ひとこと')
    r=4
    while r<=ws.max_row:
        code=norm(ws.cell(r,3).value)
        if not code: r+=1; continue
        nm=str(ws.cell(r,4).value or '')[:14]
        sc=ws.cell(r,17).value or 0
        if code not in IDX: print(f'{code:<6s}{nm:<16s}  終値シートに無し'); r+=1; continue
        b=load_bars(IDX[code])
        if not b: print(f'{code:<6s}{nm:<16s}  データ不足'); r+=1; continue
        H,L,C,A,n=b; px,a=C[-1],A[-1]
        note=[]
        atrp=a/px*100
        chg=(C[-1]/C[-2]-1)*100
        e=C[0]
        for x in C[1:]: e=(2/26)*x+(1-2/26)*e
        dev=(px/e-1)*100
        if isbuy:
            km=1.0 if sc>=16 else 0.75
            if sc>=16: note.append(f'高精度→{km}ATRまで引ける')
            if chg>=9: note.append(f'前日比+{chg:.1f}% 急騰直後')
            if dev>=15: note.append(f'25日EMA乖離+{dev:.0f}% 高値警戒')
            if dev<0: note.append('25日EMA割れ 見送り推奨')
            if atrp>=5: note.append(f'高ボラ(ATR{atrp:.1f}%) 株数を絞る')
            elif atrp<=1.5: note.append(f'低ボラ(ATR{atrp:.1f}%)')
            p1,p2,p3=tdown(px-0.5*a),tdown(px-km*a),tdown(px-1.25*a)
            sd=stopd(p2,a)
            print(f'{code:<6s}{nm:<16s}{px:9,.0f}{a:7.0f}{p1:9,.0f}{p2:9,.0f}{p3:9,.0f}{hit_down(H,L,C,A,n,km):6.1f}%{tdown(p2-sd):8,.0f}{tup(p2+2*sd):8,.0f}  {" ".join(note) or "標準"}')
        else:
            if chg<=-9: note.append(f'前日比{chg:.1f}% 急落直後')
            if dev<=-15: note.append(f'25日EMA乖離{dev:.0f}% 売られ過ぎ')
            if dev>0: note.append('25日EMA上 売りには不利')
            if atrp>=5: note.append(f'高ボラ(ATR{atrp:.1f}%) 株数を絞る')
            elif atrp<=1.5: note.append(f'低ボラ(ATR{atrp:.1f}%)')
            sd=stopd(px,a)
            print(f'{code:<6s}{nm:<16s}{px:9,.0f}{a:7.0f}{tup(px+0.39*a):10,.0f}{tup(px+0.70*a):10,.0f}{px:9,.0f}{hit_up(H,L,C,A,n,0.39):6.1f}%{tup(px+sd):8,.0f}{tdown(px-2*sd):8,.0f}  {" ".join(note) or "標準"}')
        r+=1

print(f'\n{"="*100}\n■ 厳選TOP2 → N〜Q列\n{"="*100}')
ws=wb['厳選TOP2']
for hdr,lab,isbuy in [(17,'買い候補',True),(25,'売り候補',False)]:
    print(f'  【{lab}】(行{hdr+1}〜{hdr+5})')
    any_=False
    for i in range(1,6):
        code=norm(ws.cell(hdr+i,2).value)
        if not code: continue
        any_=True
        b=load_bars(IDX[code]) if code in IDX else None
        if not b: print(f'    {code} データなし'); continue
        H,L,C,A,n=b; px,a=C[-1],A[-1]
        if isbuy:
            lim=tdown(px-0.75*a); sd=stopd(lim,a)
            print(f'    {code} ATR{a:.0f} ★指値{lim:,.0f} 到達率{hit_down(H,L,C,A,n,0.75):.1f}% 損切{tdown(lim-sd):,.0f}')
        else:
            sd=stopd(px,a)
            print(f'    {code} ATR{a:.0f} ★寄付{px:,.0f} 戻り目安{tup(px+0.39*a):,.0f} 損切{tup(px+sd):,.0f}')
    if not any_: print('    （本日は条件通過ゼロ＝空欄のまま。候補が出た日に自動で埋まる）')

print(f'\n{"="*100}\n■ 個別銘柄 → A24〜C29\n{"="*100}')
ws=wb['個別銘柄']
code=norm(ws['B3'].value)
print('  B3 の銘柄コード =',code, '→', '終値シートに有り' if code in IDX else '終値シートに無し')
if code in IDX:
    b=load_bars(IDX[code])
    if b:
        H,L,C,A,n=b; px,a=C[-1],A[-1]
        rows=[('基準終値',px,'売りは寄付成行が最良（検証結果）'),
              ('ATR14',round(a,1),f'ボラティリティ {a/px*100:.1f}%'),
              ('買 第1 −0.5ATR',tdown(px-0.5*a),f'到達率 {hit_down(H,L,C,A,n,0.5):.1f}%'),
              ('買 ★本命 −0.75ATR',tdown(px-0.75*a),f'到達率 {hit_down(H,L,C,A,n,0.75):.1f}%　3営業日有効で出す'),
              ('買 深押し −1.25ATR',tdown(px-1.25*a),f'到達率 {hit_down(H,L,C,A,n,1.25):.1f}%')]
        for a1,b1,c1 in rows: print(f'    {a1:<22s}{b1:>10,.0f}   {c1}')
    else: print('    データ不足')
