# -*- coding: utf-8 -*-
"""V816.xlsm → V817.xlsm : 第2版レポートの訂正・修正を OOXML 直編集で反映する。
VBA(vbaProject.bin)・図形・フォームコントロール等はバイト単位でそのまま引き継ぐ。"""
import re, math, shutil, zipfile, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xmlcell import set_formula, set_number, set_text, set_cell, widen_dimension, esc
import values as VAL

_wbv, _S, _CODES = VAL.load()
_K = _wbv['管理']
def _kv(col, row):
    return _K[col + str(row)].value
CACHE = {}          # 計算済みキャッシュ値（"シート!参照" -> 値）

SRC = '/tmp/claude-0/-home-user--/8b86e642-77b1-5b85-a6ca-fce6f257f782/scratchpad/v816/V816.xlsm'
DST = '/tmp/claude-0/-home-user--/8b86e642-77b1-5b85-a6ca-fce6f257f782/scratchpad/v817/V817.xlsm'

SH = {'管理':'xl/worksheets/sheet1.xml', '分析':'xl/worksheets/sheet7.xml',
      '売分析':'xl/worksheets/sheet8.xml', '自動注文決済':'xl/worksheets/sheet17.xml',
      '取扱説明書':'xl/worksheets/sheet18.xml', '変更履歴':'xl/worksheets/sheet22.xml',
      'V811設定':'xl/worksheets/sheet23.xml'}

zin = zipfile.ZipFile(SRC)
parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
order = [i.filename for i in zin.infolist()]
info  = {i.filename: i for i in zin.infolist()}
zin.close()
X = {k: parts[v].decode('utf-8') for k, v in SH.items()}
X['workbook'] = parts['xl/workbook.xml'].decode('utf-8')
log = []

B4v, B6v, H8v, H9v = VAL.B4, VAL.B6, VAL.H8, VAL.H9
TICK = ('IF($F{r}<=3000,1,IF($F{r}<=5000,5,IF($F{r}<=30000,10,'
        'IF($F{r}<=50000,50,IF($F{r}<=300000,100,500)))))')

# ============================================================
# 1. V811設定
# ============================================================
v = X['V811設定']
v = set_text(v, 'A3', '損切幅（％固定モード用）');              v = set_number(v, 'B3', '0.06')
v = set_text(v, 'A4', '利確倍率（損切幅の何倍）');              v = set_number(v, 'B4', '2')
v = set_text(v, 'C3', '← H3=0 のときだけ使用')
v = set_text(v, 'C4', '← 利確幅 = 損切幅 × この倍率')
v = set_number(v, 'B5', '40000')
v = set_text(v, 'C5', '← V817で 20,000→40,000（推奨株数を厳格運用すると2万円では約6割が見送りになるため）')
v = set_number(v, 'E3', '3.5')                       # 14日平均レンジ率 上限 2.2 → 3.5
v = set_text(v, 'D13', '※ E12=0 のあいだ売りは建てない')

# ---- 新ブロック：出口ルール詳細 (G2:I14) ----
v = set_text(v, 'G2', '■ 出口ルール詳細（V817で追加・青字＝変更可）')
P = [('G3','損切の決め方  0:％固定 / 1:ATR連動','H3','1','I3','1 を推奨（分析: +0.151R / ％固定は +0.125R）'),
     ('G4','損切ATR倍率','H4','2','I4','全ATR帯で最良'),
     ('G5','損切幅の下限(%)','H5','3','I5','低ボラ銘柄で呼値ノイズに刺さるのを防ぐ'),
     ('G6','損切幅の上限(%)','H6','10','I6','高ボラ銘柄で損切が広がりすぎるのを防ぐ'),
     ('G7','窓抜け安全率（株数計算に掛ける）','H7','1.07','I7','逆指値の18%は寄付で飛ぶ。実効損切幅は名目の約1.07倍'),
     ('G8','トレーリング開始（ATR倍率）','H8','2','I8','建玉後高値がこの水準に届いたら作動'),
     ('G9','トレーリング幅（ATR倍率）','H9','1.5','I9','建玉後高値からの引き下げ幅'),
     ('G10','ATR参照期間（日）','H10','14','I10','V811指標C列と同じ定義')]
for gc, glab, hc, hval, ic, note in P:
    v = set_text(v, gc, glab); v = set_number(v, hc, hval); v = set_text(v, ic, note)
v = set_text(v, 'G12', '※ 損切幅は銘柄ごとに算出され、V811設定 F列（管理シート各行に対応）に出ます。')
v = set_text(v, 'G13', '※ 利確幅 = 損切幅 × B4。抽出・分析・管理・自動注文決済の全シートがこの1組を参照します。')
v = set_text(v, 'G14', '※ 数値の根拠は「V817変更履歴」シートを参照。')
v = set_text(v, 'G15', '※ ATRは購入日時点で固定。建玉後にATRが変わっても損切・利確ラインは動きません（証券会社に置いた逆指値を出し直す必要なし）。')

# ---- 補助テーブルに ATR / 損切幅 列を追加、推奨株数を修正 ----
v = set_text(v, 'E16', 'ATR14(%)【購入日時点で固定】')
v = set_text(v, 'F16', '損切幅(%)')
v = set_text(v, 'G16', '（推奨株数=0 は 100株でも B5 を超える＝見送り推奨）')
v = set_text(v, 'I11', '← 1単元でも許容損失額を超え「見送り」になる行数')
v = set_text(v, 'H11', '←')
MROW = ('IFERROR(MATCH(管理!$B{k},銘柄管理!$B$6:$B$305,0),'
        'MATCH(TEXT(管理!$B{k},"0"),銘柄管理!$B$6:$B$305,0))')
# ATR は「購入日を最終日とする14日窓」で計算する。
# 直近14日にすると建玉後もATRが毎日変わり、損切・利確ラインが日々動いてしまうため。
ACOL = 'IFERROR(MATCH(管理!$E{k},高値!$E$3:$IT$3,0),1)-1'
ATR = ('IF(管理!$B{k}="","",IFERROR((SUM(OFFSET(高値!$E$6,{mr}-1,{ac},1,$H$10))'
       '-SUM(OFFSET(安値!$E$6,{mr}-1,{ac},1,$H$10)))'
       '/SUM(OFFSET(終値!$E$6,{mr}-1,{ac},1,$H$10))*100,""))')
WID = ('IF(管理!$B{k}="","",IF(AND($H$3=1,N(E{n})>0),MIN($H$6,MAX($H$5,E{n}*$H$4)),$B$3*100))')
QTY = ('IF(OR(管理!$F{k}="",N(F{n})=0),"",ROUNDDOWN($B$5/(管理!$F{k}*(F{n}/100)*$H$7)/100,0)*100)')
for n in range(17, 475):          # V811設定 行 n ↔ 管理 行 n-13
    k = n - 13
    code = _kv('B', k)
    atr  = (VAL.atr_pct(_S, _CODES, code, VAL._serial(_kv('E', k)))
            if code not in (None, '') else '')
    wid  = VAL.stop_width(atr) if code not in (None, '') else ''
    ent  = _kv('F', k)
    qt   = VAL.qty(ent, wid) if (code not in (None, '') and wid != '') else ''
    CACHE['ATR%d' % k] = atr; CACHE['W%d' % k] = wid; CACHE['Q%d' % k] = qt
    v = set_formula(v, 'E%d' % n,
                    ATR.format(k=k, mr=MROW.format(k=k), ac=ACOL.format(k=k)), cached=atr)
    v = set_formula(v, 'F%d' % n, WID.format(k=k, n=n), cached=wid)
    v = set_formula(v, 'C%d' % n, QTY.format(k=k, n=n), cached=qt)
v = set_formula(v, 'G11', 'COUNTIF($C$17:$C$474,0)&"件"',
                cached='%d件' % sum(1 for kk in range(4, 462) if CACHE.get('Q%d' % kk) == 0))
v = widen_dimension(v, 'I474')
X['V811設定'] = v
log.append('V811設定: B3/B4 再定義、E3=3.5、出口ルール詳細ブロック(G2:I14)追加、'
           'E/F列(ATR・損切幅)を458行追加、C列 推奨株数を窓抜け安全率込みに修正')

# ============================================================
# 2. 管理
# ============================================================
m = X['管理']
m = set_text(m, 'M3', '利確ライン（V811設定連動）')
m = set_text(m, 'N3', '損切ライン（V811設定連動）')
m = set_text(m, 'P3', '建玉後 高値')
m = set_text(m, 'Q3', '建玉後 安値')
m = set_text(m, 'Y3', '推奨株数')
m = set_text(m, 'K1', '🔴損切推奨')

# サマリー：ステータス文字列に依存しない集計へ
_open = [r for r in range(4, 462)
         if _kv('B', r) not in (None, '') and _kv('T', r) in (None, '')]
_unreal = sum(_kv('K', r) for r in _open if isinstance(_kv('K', r), (int, float)))
m = set_formula(m, 'B2', 'COUNTIFS($B$4:$B$461,"<>",$T$4:$T$461,"")&"件"',
                cached='%d件' % len(_open))
# E2 は V816 から書式が m/d;@（日付）で、含み損益が日付として表示されていた。
# D2（確定損益）と同じ #,##0 の書式に付け替える。
m = re.sub(r'<c r="E2" s="\d+"', '<c r="E2" s="549"', m, count=1)
m = set_formula(m, 'E2', 'SUMIFS($K$4:$K$461,$B$4:$B$461,"<>",$T$4:$T$461,"")', cached=_unreal)
_K2 = 'PLACEHOLDER_K2'

ROWOFF = ('IFERROR(MATCH($B{r},銘柄管理!$B$6:$B$305,0),'
          'MATCH(TEXT($B{r},"0"),銘柄管理!$B$6:$B$305,0))-1')
WIN    = 'MAX(1,IFERROR(MATCH($E{r},高値!$E$3:$IT$3,0),10)-1)'
W_     = 'V811設定!$F{n}'      # その行の損切幅(%)
A_     = 'V811設定!$E{n}'      # その行のATR(%)

for r in range(4, 462):
    n = r + 13
    t = TICK.format(r=r); w = W_.format(n=n); a = A_.format(n=n)
    # ---- キャッシュ値の再現 ----
    _b, _c, _e = _kv('B', r), _kv('C', r), _kv('E', r)
    _f, _h, _l = _kv('F', r), _kv('H', r), _kv('L', r)
    _t_, _w_ = _kv('T', r), _kv('W', r)
    _wid, _atr, _q = CACHE.get('W%d' % r, ''), CACHE.get('ATR%d' % r, ''), CACHE.get('Q%d' % r, '')
    _has = isinstance(_f, (int, float)) and _wid != ''
    _sell = (_c == '売')
    if _has:
        _tk = VAL.tick(_f)
        _M = (VAL.CEILING(_f * (1 - _wid / 100 * B4v), _tk) if _sell
              else VAL.FLOOR(_f * (1 + _wid / 100 * B4v), _tk))
        _N = (VAL.CEILING(_f * (1 + _wid / 100), _tk) if _sell
              else VAL.FLOOR(_f * (1 - _wid / 100), _tk))
    else:
        _M = _N = ''
    _P = VAL.post_entry_extreme(_S, _CODES, _b, VAL._serial(_e), True) if _b not in (None, '') and _e else ''
    _Q = VAL.post_entry_extreme(_S, _CODES, _b, VAL._serial(_e), False) if _b not in (None, '') and _e else ''
    # R（トレーリング）
    if not _has or not isinstance(_h, (int, float)) or _N == '':
        _R = ''
    elif _t_ not in (None, ''):
        _R = ''
    elif _sell:
        _R = _N if (_Q == '' or _Q > _f * (1 - _atr / 100 * H8v)) else min(_N, _Q * (1 + _atr / 100 * H9v))
    else:
        _R = _N if (_P == '' or _P < _f * (1 + _atr / 100 * H8v)) else max(_N, _P * (1 - _atr / 100 * H9v))
    # O（判断）
    if _b in (None, '') or not isinstance(_f, (int, float)) or not isinstance(_h, (int, float)):
        _O = ''
    elif _t_ not in (None, ''):
        _O = '―決済済―'
    else:
        _wd = _w_ if isinstance(_w_, (int, float)) else 0
        if _sell:
            _O = ('🟢利確' if _h <= _M else '🟠TRAIL' if (_R != '' and _h >= _R)
                  else '🔴損切' if _h >= _N else '⏱時間決済' if _wd >= B6v
                  else '⚠️損切接近' if _h >= _f * (1 + _wid / 100 * 0.7) else '保持継続')
        else:
            _O = ('🟢利確' if _h >= _M else '🟠TRAIL' if (_R != '' and _h <= _R)
                  else '🔴損切' if _h <= _N else '⏱時間決済' if _wd >= B6v
                  else '⚠️損切接近' if _h <= _f * (1 - _wid / 100 * 0.7) else '保持継続')
    # S（ステータス）
    if _b in (None, ''):
        _Sv = ''
    elif _t_ not in (None, ''):
        _Sv = '✅売却済'
    elif not isinstance(_h, (int, float)) or not isinstance(_l, (int, float)):
        _Sv = '📌保有中'
    else:
        _Sv = {'🟢利確': '🟢利確推奨', '🟠TRAIL': '🟠TRAIL STOP推奨', '🔴損切': '🔴損切推奨',
               '⏱時間決済': '⏱時間STOP推奨', '⚠️損切接近': '⚠️損切注意'}.get(
            _O, '🟡利確直近' if (_wid != '' and _l >= _wid / 100 * B4v * 0.8) else '📌保有中')
    CACHE['S%d' % r] = _Sv
    _AE = VAL.workday(VAL._serial(_e), B6v) if _e else ''
    _Y = '' if _q == '' else ('⛔1単元でも超過' if _q == 0 else _q)
    m = set_formula(m, 'M%d' % r,
        'IF($F{r}="","",IF($C{r}="売",CEILING($F{r}*(1-{w}/100*V811設定!$B$4),{t}),'
        'FLOOR($F{r}*(1+{w}/100*V811設定!$B$4),{t})))'.format(r=r, w=w, t=t), cached=_M)
    m = set_formula(m, 'N%d' % r,
        'IF($F{r}="","",IF($C{r}="売",CEILING($F{r}*(1+{w}/100),{t}),'
        'FLOOR($F{r}*(1-{w}/100),{t})))'.format(r=r, w=w, t=t), cached=_N)
    m = set_formula(m, 'P%d' % r,
        'IF(OR($B{r}="",$E{r}=""),"",IFERROR(MAX(OFFSET(高値!$E$6,{o},0,1,{win})),""))'
        .format(r=r, o=ROWOFF.format(r=r), win=WIN.format(r=r)), cached=_P)
    m = set_formula(m, 'Q%d' % r,
        'IF(OR($B{r}="",$E{r}=""),"",IFERROR(MIN(OFFSET(安値!$E$6,{o},0,1,{win})),""))'
        .format(r=r, o=ROWOFF.format(r=r), win=WIN.format(r=r)), cached=_Q)
    # 判断（S列への参照をやめて循環参照を断つ）
    m = set_formula(m, 'O%d' % r,
        'IF(OR($B{r}="",$F{r}="",$H{r}=""),"",IF($T{r}<>"","―決済済―",'
        'IF($C{r}="売",'
        'IF($H{r}<=$M{r},"🟢利確",IF(AND($R{r}<>"",$H{r}>=$R{r}),"🟠TRAIL",'
        'IF($H{r}>=$N{r},"🔴損切",IF($W{r}>=V811設定!$B$6,"⏱時間決済",'
        'IF($H{r}>=$F{r}*(1+{w}/100*0.7),"⚠️損切接近","保持継続"))))),'
        'IF($H{r}>=$M{r},"🟢利確",IF(AND($R{r}<>"",$H{r}<=$R{r}),"🟠TRAIL",'
        'IF($H{r}<=$N{r},"🔴損切",IF($W{r}>=V811設定!$B$6,"⏱時間決済",'
        'IF($H{r}<=$F{r}*(1-{w}/100*0.7),"⚠️損切接近","保持継続")))))'
        ')))'.format(r=r, w=w), cached=_O)
    # トレーリング：建玉日起点の高安値／発動・幅は V811設定 H8・H9
    m = set_formula(m, 'R%d' % r,
        'IF(OR($F{r}="",$H{r}="",$N{r}=""),"",IF($T{r}<>"","",'
        'IF($C{r}="売",'
        'IF(OR($Q{r}="",$Q{r}>$F{r}*(1-{a}/100*V811設定!$H$8)),$N{r},'
        'MIN($N{r},$Q{r}*(1+{a}/100*V811設定!$H$9))),'
        'IF(OR($P{r}="",$P{r}<$F{r}*(1+{a}/100*V811設定!$H$8)),$N{r},'
        'MAX($N{r},$P{r}*(1-{a}/100*V811設定!$H$9))))))'.format(r=r, a=a), cached=_R)
    # ステータス：判断列から導出し、既存VBA(損切りアラート)が拾える文字列を出す
    m = set_formula(m, 'S%d' % r,
        'IF($B{r}="","",IF($T{r}<>"","✅売却済",IF(OR($H{r}="",$L{r}=""),"📌保有中",'
        'IF($O{r}="🟢利確","🟢利確推奨",IF($O{r}="🟠TRAIL","🟠TRAIL STOP推奨",'
        'IF($O{r}="🔴損切","🔴損切推奨",IF($O{r}="⏱時間決済","⏱時間STOP推奨",'
        'IF($O{r}="⚠️損切接近","⚠️損切注意",'
        'IF($L{r}>={w}/100*V811設定!$B$4*0.8,"🟡利確直近","📌保有中")))))))))'.format(r=r, w=w), cached=_Sv)
    m = set_formula(m, 'AE%d' % r, 'IF($E{r}="","",WORKDAY($E{r},V811設定!$B$6))'.format(r=r), cached=_AE)
    m = set_formula(m, 'Y%d' % r,
        'IF(V811設定!$C{n}="","",IF(V811設定!$C{n}=0,"⛔1単元でも超過",V811設定!$C{n}))'.format(n=n), cached=_Y)
m = set_formula(m, 'K2', 'COUNTIF($S:$S,"🔴損切推奨")&"件"',
                cached='%d件' % sum(1 for r in range(4, 462) if CACHE.get('S%d' % r) == '🔴損切推奨'))
X['管理'] = m
log.append('管理: E2の書式を日付(m/d;@)から#,##0に修正（V816からの既存バグ）')
log.append('管理: M/N(V811設定連動＋呼値丸め)、P/Q(建玉日起点)、O(循環参照解消)、'
           'R(建玉後高安値＋H8/H9)、S(O列から導出＋"損切推奨"/"STOP推奨"を出力)、'
           'AE(B6参照)、Y列 推奨株数を458行に適用。1〜2行目の集計をステータス非依存に変更')

# ============================================================
# 3. 分析 / 売分析
# ============================================================
def width_expr(atr_ref):
    return ('IF(AND(V811設定!$H$3=1,N(%s)>0),MIN(V811設定!$H$6,MAX(V811設定!$H$5,%s*V811設定!$H$4)),'
            'V811設定!$B$3*100)' % (atr_ref, atr_ref))
_IND = _wbv['V811指標']
def _w_from(atr):
    return VAL.stop_width(atr if isinstance(atr, (int, float)) else '')
def _xlround(x):
    return math.floor(x + 0.5) if x >= 0 else -math.floor(-x + 0.5)

a = X['分析']; _AN = _wbv['分析']
for r in range(3, 207):
    w = width_expr('V811指標!$C%d' % r)
    _e = _AN['E%d' % r].value
    _ww = _w_from(_IND['C%d' % r].value)
    _R = _xlround(_e * (1 + _ww / 100 * VAL.B4)) if isinstance(_e, (int, float)) else ''
    _Sx = _xlround(_e * (1 - _ww / 100)) if isinstance(_e, (int, float)) else ''
    a = set_formula(a, 'R%d' % r, 'IF($E{r}="","",ROUND($E{r}*(1+{w}/100*V811設定!$B$4),0))'.format(r=r, w=w), cached=_R)
    a = set_formula(a, 'S%d' % r, 'IF($E{r}="","",ROUND($E{r}*(1-{w}/100),0))'.format(r=r, w=w), cached=_Sx)
X['分析'] = a
s = X['売分析']; _SN = _wbv['売分析']
for r in range(3, 207):
    w = width_expr('V811指標!$M%d' % r)
    _e = _SN['E%d' % r].value
    _ww = _w_from(_IND['M%d' % r].value)
    _R = _xlround(_e * (1 - _ww / 100 * VAL.B4)) if isinstance(_e, (int, float)) else ''
    _Sx = _xlround(_e * (1 + _ww / 100)) if isinstance(_e, (int, float)) else ''
    s = set_formula(s, 'R%d' % r, 'IF($E{r}="","",ROUND($E{r}*(1-{w}/100*V811設定!$B$4),0))'.format(r=r, w=w), cached=_R)
    s = set_formula(s, 'S%d' % r, 'IF($E{r}="","",ROUND($E{r}*(1+{w}/100),0))'.format(r=r, w=w), cached=_Sx)
X['売分析'] = s
log.append('分析/売分析: R・S列 204行をV811設定連動に置換（3行目に残っていた×1.05/×0.95も解消）')

# ============================================================
# 4. 自動注文決済
# ============================================================
o = X['自動注文決済']
_ord = _wbv['自動注文決済']
_c3 = _ord['C3'].value
_c24 = (VAL.atr_pct(_S, _CODES, _c3, VAL._serial(_ord['C11'].value))
        if _c3 not in (None, '') else '')
_c17 = VAL.stop_width(_c24); _c18 = _c17 * VAL.B4
_c19 = round(_c24 * VAL.H8, 2) if isinstance(_c24, (int, float)) and _c24 > 0 else _c17
_c20 = round(_c24 * VAL.H9, 2) if isinstance(_c24, (int, float)) and _c24 > 0 else round(_c17 * 0.75, 2)
_c11 = _ord['C11'].value
_c40 = VAL.workday(VAL._serial(_c11), VAL.B6) if _c11 else ''
o = set_text(o, 'B24', 'ATR14(%)（自動）')
o = set_formula(o, 'C24',
    'IFERROR((SUM(OFFSET(高値!$E$6,{mr}-1,{ac},1,V811設定!$H$10))'
    '-SUM(OFFSET(安値!$E$6,{mr}-1,{ac},1,V811設定!$H$10)))'
    '/SUM(OFFSET(終値!$E$6,{mr}-1,{ac},1,V811設定!$H$10))*100,"")'.format(
        mr='IFERROR(MATCH($C$3,銘柄管理!$B$6:$B$305,0),MATCH(TEXT($C$3,"0"),銘柄管理!$B$6:$B$305,0))',
        ac='IFERROR(MATCH($C$11,高値!$E$3:$IT$3,0),1)-1'), cached=_c24)
o = set_text(o, 'D24', '建日(C11)を最終日とする H10 日間の平均レンジ率。建玉後は動きません')
o = set_formula(o, 'C17', cached=_c17, formula='IF(AND(V811設定!$H$3=1,N(C24)>0),MIN(V811設定!$H$6,'
                          'MAX(V811設定!$H$5,C24*V811設定!$H$4)),V811設定!$B$3*100)')
o = set_formula(o, 'C18', cached=_c18, formula='C17*V811設定!$B$4')
o = set_formula(o, 'C19', cached=_c19, formula='IF(AND(V811設定!$H$3=1,N(C24)>0),ROUND(C24*V811設定!$H$8,2),C17)')
o = set_formula(o, 'C20', cached=_c20, formula='IF(AND(V811設定!$H$3=1,N(C24)>0),ROUND(C24*V811設定!$H$9,2),ROUND(C17*0.75,2))')
o = set_formula(o, 'C21', 'V811設定!$B$6', cached=VAL.B6)
o = set_formula(o, 'C40', 'IFERROR(IF(C11="","",WORKDAY(C11,C21)),"")', cached=_c40)
for ref, txt in [('D17','V811設定から自動算出（H3=1でATR連動）'),
                 ('D18','損切幅 × V811設定!B4（利確倍率）'),
                 ('D19','ATR × V811設定!H8'),
                 ('D20','ATR × V811設定!H9'),
                 ('D21','V811設定!B6（営業日）。C40はWORKDAYで計算'),
                 ('D40','建日＋N営業日')]:
    o = set_text(o, ref, txt)
# エントリートリガーを停止（決済側が動かないため）
o = set_number(o, 'C54', '0')
o = set_text(o, 'D54', '⚠ V817で 1→0 に変更。決済トリガーC44を書くマクロが未実装のため'
                       '自動エントリーを停止。再開する場合は =IF(C34="✓エントリー可",1,0) に置換')
o = set_text(o, 'D44', '⚠ このセルを1にするマクロは未実装。C49の自動決済は発火しません（手動決済してください）')
X['自動注文決済'] = o
log.append('自動注文決済: C17〜C21をV811設定連動、C24にATR、C40をWORKDAY化、'
           '★C54 エントリートリガーを 1→0（自動発注を停止）')

# ============================================================
# 5. 取扱説明書 の訂正
# ============================================================
t = X['取扱説明書']
fix = {
 'B18':'C18','C18':'利確%','D18':'損切幅 × V811設定!B4（利確倍率）。V811設定で一元管理',
 'B19':'C17','C19':'ストップロス%','D19':'V811設定!H3=1 のとき ATR×H4（下限H5・上限H6でクランプ）',
 'B20':'C19','C20':'トレーリング開始%','D20':'ATR × V811設定!H8',
 'B21':'C20','C21':'トレーリング幅%','D21':'ATR × V811設定!H9',
 'B22':'C21','C22':'最大保有日数','D22':'V811設定!B6（営業日）。C40はWORKDAYで計算',
 'B23':'C22','C23':'GU率上限%','D23':'この%を超えるGUの日はエントリーしない',
 'B24':'C23','C24':'VWAP比較','D24':'1：始値＞VWAPの日のみ　0：フィルターなし',
}
for ref, val in fix.items():
    t = set_text(t, ref, val)
for ref, val in {
 'A33':'▶ 保有中の監視（※ V817時点で未実装）',
 'A34':'　・ CheckSwingTrailing / ResetSwingTracking はこのブックに存在しません。',
 'A35':'　・ そのためC29高値記録・C30フラグ・C44決済トリガーは自動更新されません。手動決済してください。',
 'A39':'①利確：建値 × (1＋損切幅×V811設定!B4)　※固定%ではありません',
 'A40':'②ストップロス：建値 × (1−損切幅)　損切幅はV811設定で算出',
 'A41':'③トレーリングストップ：建玉後高値 × (1−ATR×H9)　※監視マクロ未実装のため現在は動きません',
 'A42':'④日数期限：建日＋V811設定!B6 営業日',
 'A46':'　※ V817時点では ResetSwingTracking が存在しないため、C29・C30・C44を手動で0に戻してください。',
 'A52':'③ V817では C54（エントリートリガー）を0にしてあります。自動発注を使う場合のみ1に戻してください。',
}.items():
    t = set_text(t, ref, val)
X['取扱説明書'] = t
log.append('取扱説明書: セル番地のずれ（C16/C17…→C17/C18…）と数値を訂正、'
           '未実装マクロについての注記を追加')

# ============================================================
# 6. Sheet3 → V817変更履歴
# ============================================================
CH = [
 ('V817 変更履歴 ── V816 第2版分析（32,576件の約定シミュレーション）にもとづく訂正・修正',''),
 ('',''),
 ('■ 設定の一元化','損切・利確を決めていた9か所を V811設定 の1組に統合'),
 ('V811設定!B3','損切幅（％固定モード用）＝0.06。H3=0 のときだけ使用'),
 ('V811設定!B4','利確倍率＝2.0。利確幅 = 損切幅 × B4'),
 ('V811設定!H3','損切の決め方 0:％固定 / 1:ATR連動（既定 1）'),
 ('V811設定!H4','損切ATR倍率＝2.0（全ATR帯で最良）'),
 ('V811設定!H5・H6','損切幅の下限3% / 上限10%'),
 ('V811設定!H7','窓抜け安全率1.07（逆指値の18%は寄付で飛ぶ。実効損切幅は名目の約1.07倍）'),
 ('V811設定!H8・H9','トレーリング開始 ATR×2.0 / 幅 ATR×1.5'),
 ('V811設定!H10','ATR参照期間 14日'),
 ('V811設定!E3','抽出ゲートの14日平均レンジ率 上限 2.2% → 3.5%（効率をほぼ保ったまま候補が26%→59%に）'),
 ('V811設定!E12','売りは 0（停止）のまま。売りは全設定で −0.16〜−0.18R'),
 ('V811設定!E17:F474','各行のATR14(%)と損切幅(%)を算出（管理シート4〜461行に対応）'),
 ('★ ATRは購入日で固定','ATRの14日窓を「購入日を最終日」に固定。直近14日にすると建玉後もATRが毎日変わり、'),
 ('','損切・利確ラインが日々動いてしまう（実測で5営業日に最大1.92%移動）。'),
 ('','証券会社に置いた逆指値・指値は建玉時に1回出せばよく、出し直す必要はありません。'),
 ('','動くのはトレーリング(管理R列)だけで、しかも有利方向へのラチェットのみです。'),
 ('V811設定!C17:C474','推奨株数を「許容損失額 ÷（建値×損切幅×窓抜け安全率）」に修正。'),
 ('','下限100株を撤廃：1単元でも許容損失を超える銘柄は0（＝見送り推奨）を返す'),
 ('',''),
 ('■ 管理シート',''),
 ('M・N列','V811設定の損切幅から算出。呼値丸めは買い=切り捨て / 売り=切り上げを維持'),
 ('P・Q列','「直近10日」固定窓 → 「建玉日起点」の高値・安値に変更'),
 ('O列','S列への参照をやめ売却日T列を見るよう変更（S列との循環参照を解消）'),
 ('R列','建玉後高安値をもとにトレーリング。発動・幅をV811設定!H8・H9から取得'),
 ('S列','O列から導出。閾値のベタ書き（0.06/0.05/0.045/-0.03/-0.04）を廃止'),
 ('S列（重要）','出力文字列に「損切推奨」「STOP推奨」を含めた。これにより既存VBA Module1.損切りアラート が'),
 ('','VBAを書き換えずに機能するようになる（従来は該当0件で毎回「全ポジション正常です」と表示されていた）'),
 ('AE列','WORKDAY(購入日,12) → WORKDAY(購入日,V811設定!B6)。期間定義を10営業日に統一'),
 ('Y列','推奨株数を新設。0のときは「⛔1単元でも超過」と表示（G列の実績値は履歴保持のため未変更）'),
 ('1〜2行目','保有件数・含み損益の集計をステータス文字列に依存しない SUMIFS/COUNTIFS に変更'),
 ('',''),
 ('■ 分析 / 売分析',''),
 ('R・S列','×1.08/×0.96・×0.92/×1.04 のベタ書きを廃止しV811設定連動に。'),
 ('','3行目だけ残っていた ×1.05/×0.95（TOPIX行）も解消。厳選TOP2!K・L列はここを参照するため自動で追随'),
 ('',''),
 ('■ 自動注文決済',''),
 ('C54','★ 1 → 0。決済トリガーC44を書くマクロが存在せず「買いだけ自動・決済は出ない」状態だったため停止'),
 ('C17〜C21','V811設定連動。C24に銘柄ATRを追加'),
 ('C40','C11+C21（暦日）→ WORKDAY(C11,C21)（営業日）'),
 ('',''),
 ('■ このブックでは直せなかったもの（VBAの再インポートが必要）',''),
 ('買抽出v13!O・P列','Mod_買抽出v13.bas:964-965 の ×1.08 / ×0.96 がVBAで毎回書き込まれます'),
 ('売抽出v13!O・P列','Mod_SellExtrac.bas:1022,1024 の ×0.92 / ×1.04 も同様'),
 ('','→ 同梱の patch/Mod_ExitParams.bas を標準モジュールとしてインポートし、各抽出モジュールの'),
 ('','　 該当2行を ExitTP(closePrice, True) / ExitSL(closePrice, True) に差し替えてください'),
 ('Mod_KanriSheet','現行レイアウトと別物の管理シートを作り直すマクロ。実行しないでください'),
 ('管理!AF・AG列','未到達時に1900/1/1と表示される件は配列数式のため未修正（判定AH列は正しく動作）'),
 ('呼値テーブル','TOPIX100銘柄は実際の刻みが異なります（例：57,000円は実際50円刻み、数式は100円刻み）'),
 ('',''),
 ('■ 許容損失額（B5）と「見送り率」の対照表',''),
 ('','推奨株数を厳格に守ると、1単元でも予算を超える銘柄は建てられません（=見送り）。'),
 ('','銘柄管理300銘柄の株価は中央値2,969円・第3四分位5,240円。単元は100株。'),
 ('B5=20,000円','ATR連動 見送り59% ／ ％固定-6% 59% ／ ％固定-4%(旧) 33%   ← 実建玉49件で計算'),
 ('B5=30,000円','ATR連動 見送り41% ／ ％固定-6% 33% ／ ％固定-4% 10%'),
 ('B5=40,000円 ★採用','ATR連動 見送り29% ／ ％固定-6% 16% ／ ％固定-4%  6%   実リスク中央値 約32,000円'),
 ('B5=50,000円','ATR連動 見送り20% ／ ％固定-6% 10% ／ ％固定-4%  4%'),
 ('B5=100,000円','ATR連動 見送り 6% ／ ％固定-6%  2% ／ ％固定-4%  2%'),
 ('','※ 見送りが出るのはATR連動のせいではなく「リスク上限を実際に守らせた」ためです。'),
 ('','　 従来の-4%固定でもB5=20,000円なら33%が見送りになります（V816では上限が無視されていた）。'),
 ('',''),
 ('■ 設定プリセット（H3セル1つで切り替え）',''),
 ('H3=1（既定・ATR連動）','損切=ATR×2.0（3〜10%でクランプ）／利確=その2.0倍　→ 検証 +0.151R'),
 ('H3=0（％固定）','損切=B3×100=6%／利確=12%　→ 検証 +0.125R。全銘柄同じ幅で運用は単純、見送りも少なめ'),
 ('（参考）V816の設定','損切4%／利確6%　→ 検証 +0.094R。B3=0.04・B4=1.5・H3=0 に戻せば再現できます'),
 ('',''),
 ('■ 根拠（V816 第2版分析より）',''),
 ('リスク1単位あたり期待値','固定-4%/+6% = +0.094R ／ ATR×2.0・利確2.0倍（クランプ込み）= +0.151R'),
 ('窓抜け','-4%逆指値の18.0%が寄付で飛び、平均-1.58%の滑り。実効損切幅は-4.28%'),
 ('到達タイミング','1〜3日目は損切25.2%対利確14.5%。4日目以降はほぼ互角、7日目以降は利確優位'),
 ('売り','全設定で -0.16〜-0.18R。窓抜け率も買いより高い'),
 ('重要な前提','検証母集団（銘柄管理300銘柄・250営業日）は10日で平均+1.82%の上方ドリフトあり。'),
 ('','広い損切・長い保有が有利に出るのは大部分がこのドリフトの取り込み。急落局面では逆になります。'),
 ('','だからこそ V811設定!B5（1回の許容損失額）を守ることが最優先です。'),
]
rows_xml = []
for i, (c1, c2) in enumerate(CH, start=1):
    cells = ''
    if c1: cells += '<c r="A%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (i, esc(c1))
    if c2: cells += '<c r="B%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (i, esc(c2))
    if cells: rows_xml.append('<row r="%d">%s</row>' % (i, cells))
X['変更履歴'] = (
 '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 '<dimension ref="A1:B%d"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
 '<sheetFormatPr defaultRowHeight="18.6"/>'
 '<cols><col min="1" max="1" width="34" customWidth="1"/>'
 '<col min="2" max="2" width="108" customWidth="1"/></cols>'
 '<sheetData>%s</sheetData><pageMargins left="0.7" right="0.7" top="0.75" '
 'bottom="0.75" header="0.3" footer="0.3"/></worksheet>' % (len(CH), ''.join(rows_xml)))
log.append('Sheet3 → 「V817変更履歴」にリネームし、変更内容と根拠を記載')

# ============================================================
# 7. workbook.xml
# ============================================================
w = X['workbook']
w = w.replace('<sheet name="Sheet3" sheetId="28"', '<sheet name="V817変更履歴" sheetId="28"')
# fullCalcOnLoad は付けない：RSS未接続で開いたときにキャッシュ表示が消えるため
X['workbook'] = w
log.append('workbook.xml: シート名変更。fullCalcOnLoad は付けない（キャッシュ表示を保つ）')

# ============================================================
# 書き出し
# ============================================================
parts['xl/workbook.xml'] = X['workbook'].encode('utf-8')
for k, path in SH.items():
    parts[path] = X[k].encode('utf-8')
app = parts['docProps/app.xml'].decode('utf-8').replace('<vt:lpstr>Sheet3</vt:lpstr>',
                                                        '<vt:lpstr>V817変更履歴</vt:lpstr>')
parts['docProps/app.xml'] = app.encode('utf-8')

# calcChain.xml は新規セルを含まないため削除。Excel が開いたときに作り直す。
CHAIN = 'xl/calcChain.xml'
if CHAIN in parts:
    del parts[CHAIN]
    order = [n for n in order if n != CHAIN]
    ct = parts['[Content_Types].xml'].decode('utf-8')
    ct = re.sub(r'<Override[^>]*calcChain[^>]*/>', '', ct)
    parts['[Content_Types].xml'] = ct.encode('utf-8')
    wr = parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
    wr = re.sub(r'<Relationship[^>]*calcChain\.xml"[^>]*/>', '', wr)
    parts['xl/_rels/workbook.xml.rels'] = wr.encode('utf-8')
    log.append('calcChain.xml を削除（新規数式セルを含まないため。Excel が再構築します）')

zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for name in order:
    zi = zipfile.ZipInfo(name, date_time=info[name].date_time)
    zi.compress_type = info[name].compress_type
    zi.external_attr = info[name].external_attr
    zo.writestr(zi, parts[name])
zo.close()
print('作成:', DST, os.path.getsize(DST), 'bytes  （元: %d bytes）' % os.path.getsize(SRC))
print()
for l in log:
    print(' •', l)
