#!/usr/bin/env python3
"""ATR_MS2 修正版の全数検証。Excel が「修復」する原因になりうる項目を機械的に確認する。"""
import re, glob, os, sys, zipfile, hashlib, io, warnings
import xml.etree.ElementTree as ET

ROOT = sys.argv[1] if len(sys.argv) > 1 else 'x'
errs = []
def E(m): errs.append(m)

BASE = set("""ABS ACCRINT ACCRINTM ACOS ACOSH ADDRESS AMORDEGRC AMORLINC AND AREAS ASC ASIN ASINH ATAN ATAN2 ATANH
AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT BESSELI BESSELJ BESSELK BESSELY BETADIST BETAINV BIN2DEC
BIN2HEX BIN2OCT BINOMDIST CEILING CELL CHAR CHIDIST CHIINV CHITEST CHOOSE CLEAN CODE COLUMN COLUMNS COMBIN COMPLEX
CONCATENATE CONFIDENCE CONVERT CORREL COS COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS
COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR CRITBINOM CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS360 DB
DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR
DPRODUCT DSTDEV DSTDEVP DSUM DURATION DVAR DVARP EDATE EFFECT EOMONTH ERF ERFC ERROR.TYPE EVEN EXACT EXP EXPONDIST
FACT FACTDOUBLE FALSE FDIST FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FORECAST FREQUENCY FTEST FV FVSCHEDULE
GAMMADIST GAMMAINV GAMMALN GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR
HYPERLINK HYPGEOMDIST IF IFERROR INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN
ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISPMT ISREF ISTEXT JIS KURT LARGE LCM LEFT LEFTB LEN LENB LINEST LN LOG
LOG10 LOGEST LOGINV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINUTE
MINVERSE MIRR MMULT MOD MODE MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST NETWORKDAYS NOMINAL NORMDIST NORMINV
NORMSDIST NORMSINV NOT NOW NPER NPV OCT2BIN OCT2DEC OCT2HEX ODD OFFSET OR PEARSON PERCENTILE PERCENTRANK PERMUT
PHONETIC PI PMT POISSON POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUOTIENT RADIANS RAND
RANDBETWEEN RANK RATE RECEIVED REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS RSQ RTD
SEARCH SEARCHB SECOND SERIESSUM SIGN SIN SINH SKEW SLN SLOPE SMALL SQRT SQRTPI STANDARDIZE STDEV STDEVA STDEVP
STDEVPA STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SYD T TAN TANH
TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TIME TIMEVALUE TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUE TRUNC TTEST
TYPE UPPER VALUE VAR VARA VARP VARPA VDB VLOOKUP WEEKDAY WEEKNUM WEIBULL WORKDAY XIRR XNPV YEAR YEARFRAC YIELD
YIELDDISC YIELDMAT ZTEST""".split())

def cidx(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n

# --- パッケージ整合
ct = open(f'{ROOT}/[Content_Types].xml', encoding='utf-8').read()
for p in re.findall(r'PartName="([^"]+)"', ct):
    if not os.path.exists(ROOT + p): E(f"Content_Types: {p} が存在しない")
for rf in glob.glob(f'{ROOT}/**/_rels/*.rels', recursive=True):
    base = os.path.dirname(os.path.dirname(rf))
    for m in re.finditer(r'<Relationship\b[^>]*>', open(rf, encoding='utf-8').read()):
        tag = m.group(0)
        if 'TargetMode="External"' in tag: continue
        t = re.search(r'Target="([^"]+)"', tag).group(1)
        if t.startswith(('http', 'file:')): continue
        if not os.path.exists(os.path.normpath(os.path.join(base, t))):
            E(f"{rf}: Target {t} が存在しない")

ss = open(f'{ROOT}/xl/sharedStrings.xml', encoding='utf-8').read()
n_ss = len(re.findall(r'<si>', ss))
st = open(f'{ROOT}/xl/styles.xml', encoding='utf-8').read()
n_xf = len(re.findall(r'<xf\b', re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', st, re.S).group(1)))
if int(re.search(r'<cellXfs count="(\d+)">', st).group(1)) != n_xf: E("cellXfs の count 不一致")

for f in sorted(glob.glob(f'{ROOT}/xl/worksheets/sheet*.xml')):
    nm = os.path.basename(f)
    try: ET.parse(f)
    except Exception as e: E(f"{nm}: XML不正 {e}"); continue
    d = open(f, encoding='utf-8').read()
    prev = 0; sdef = set(); sref = set(); seen_rows = set()
    for m in re.finditer(r'<row([^>]*)>(.*?)</row>', d, re.S):
        rn = int(re.search(r'r="(\d+)"', m.group(1)).group(1))
        if rn in seen_rows: E(f"{nm}: 行 {rn} が重複")
        seen_rows.add(rn)
        if rn <= prev: E(f"{nm}: 行 {rn} が昇順でない")
        prev = rn; pc = 0; seen = set()
        for c in re.findall(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', m.group(2), re.S):
            r2 = re.search(r'r="([A-Z]+)(\d+)"', c)
            if not r2: E(f"{nm} r{rn}: r属性のないセル"); continue
            col, crn = r2.group(1), int(r2.group(2))
            if crn != rn: E(f"{nm}: セル {col}{crn} が row r={rn} の中にある")
            if col in seen: E(f"{nm} r{rn}: セル {col} が重複")
            seen.add(col)
            if cidx(col) <= pc: E(f"{nm} r{rn}: セル {col} が列順でない")
            pc = cidx(col)
            s = re.search(r'\ss="(\d+)"', c)
            if s and int(s.group(1)) >= n_xf: E(f"{nm}: {col}{rn} の style index が範囲外")
            t = re.search(r'\st="([a-zA-Z]+)"', c); ty = t.group(1) if t else None
            v = re.search(r'<v>(.*?)</v>', c, re.S)
            if ty == 's':
                if not v or not v.group(1).isdigit() or int(v.group(1)) >= n_ss:
                    E(f"{nm}: {col}{rn} の共有文字列 index が不正")
            elif ty == 'inlineStr' and '<is>' not in c:
                E(f"{nm}: {col}{rn} t=inlineStr なのに <is> が無い")
            elif ty is None and v and '<f' not in c:
                try: float(v.group(1))
                except ValueError: E(f"{nm}: {col}{rn} 型指定なしなのに数値でない値 {v.group(1)!r}")
            for fm in re.finditer(r'<f\b([^>]*)(/>|>)', c):
                si = re.search(r'si="(\d+)"', fm.group(1))
                if si: (sref if fm.group(2) == '/>' else sdef).add(si.group(1))
    if sref - sdef: E(f"{nm}: 定義の無い共有数式 si={sorted(sref - sdef)[:3]}")
    # ★Excel 2007 以降の追加関数は _xlfn. が必要（今回の修復原因）
    for fm in re.findall(r'<f[^>]*>(.*?)</f>', d, re.S):
        for name in re.findall(r'(?<![A-Za-z0-9_.])([A-Z][A-Z0-9_.]*)\(', fm):
            if name not in BASE and not name.startswith(('_xll', '_xlfn')):
                E(f"{nm}: 関数 {name}() は _xlfn. 接頭辞が必要（またはタイプミス）")
    mc = re.search(r'<mergeCells count="(\d+)">(.*?)</mergeCells>', d, re.S)
    if mc and int(mc.group(1)) != len(re.findall(r'<mergeCell ', mc.group(2))):
        E(f"{nm}: mergeCells の count 不一致")

print("=== 検証結果 ===")
if errs:
    for e in sorted(set(errs))[:30]: print("  NG:", e)
    print(f"  合計 {len(set(errs))} 件")
    sys.exit(1)
print("  問題なし")
