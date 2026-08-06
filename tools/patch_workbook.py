#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打ち出の小づち V805 -> V806 修正パッチ

xlsm を ZIP レベルで書き換えるため、VBA・グラフ・図形・webextension など
openpyxl が失う要素をすべて保持したまま数式だけを差し替える。
"""
import re, shutil, sys, zipfile
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "v805.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "V806_修正版.xlsm")

# 管理=sheet1 / 厳選TOP2=sheet3 / 分析=sheet7 / 売分析=sheet8
KANRI, GENSEN, BUNSEKI, URIBUNSEKI = "xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml", \
                                     "xl/worksheets/sheet7.xml", "xl/worksheets/sheet8.xml"

log = []


def colnum(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def set_formula(xml: str, cell: str, formula: str) -> str:
    """<c r="R4"><f>…</f><v>…</v></c> の <f> を差し替えキャッシュ値を削除する。
    セルが存在しない場合は行内の正しい列順で新規挿入する。"""
    col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
    body = f'<c r="{cell}"><f>{formula}</f></c>'
    m = re.search(r'(<c r="%s"[^>]*?)(/>|>.*?</c>)' % cell, xml, re.S)
    if m:
        attrs = re.sub(r'\s+t="[a-z]+"', "", m.group(1))  # 文字列型指定を外す
        return xml[: m.start()] + attrs + ">" + f"<f>{formula}</f>" + "</c>" + xml[m.end():]

    rm = re.search(r'<row r="%s"[^>]*?(/>|>.*?</row>)' % row, xml, re.S)
    if not rm:
        log.append(f"  [skip] 行{row} が存在しません ({cell})")
        return xml
    chunk = rm.group(0)
    if chunk.endswith("/>"):                      # 空行 → 子要素つきに開き直す
        new = chunk[:-2] + ">" + body + "</row>"
        return xml[: rm.start()] + new + xml[rm.end():]
    target = colnum(col)
    ins = None
    for cm in re.finditer(r'<c r="([A-Z]+)%s"' % row, chunk):
        if colnum(cm.group(1)) > target:
            ins = cm.start()
            break
    new = chunk[:ins] + body + chunk[ins:] if ins is not None else \
          chunk[: chunk.rindex("</row>")] + body + "</row>"
    return xml[: rm.start()] + new + xml[rm.end():]


def set_inline_text(xml: str, cell: str, text: str) -> str:
    """見出しセルをインライン文字列として書き込む（sharedStrings を触らずに済む）。"""
    col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
    body = f'<c r="{cell}" t="inlineStr"><is><t>{esc(text)}</t></is></c>'
    m = re.search(r'<c r="%s"[^>]*?(/>|>.*?</c>)' % cell, xml, re.S)
    if m:
        return xml[: m.start()] + body + xml[m.end():]
    rm = re.search(r'<row r="%s"[^>]*?(/>|>.*?</row>)' % row, xml, re.S)
    if not rm:
        log.append(f"  [skip] 見出し行{row} が存在しません ({cell})")
        return xml
    chunk = rm.group(0)
    if chunk.endswith("/>"):
        new = chunk[:-2] + ">" + body + "</row>"
    else:
        target = colnum(col)
        ins = next((cm.start() for cm in re.finditer(r'<c r="([A-Z]+)%s"' % row, chunk)
                    if colnum(cm.group(1)) > target), None)
        new = chunk[:ins] + body + chunk[ins:] if ins is not None else \
              chunk[: chunk.rindex("</row>")] + body + "</row>"
    return xml[: rm.start()] + new + xml[rm.end():]


def set_number(xml: str, cell: str, value) -> str:
    """設定値セルを数値として書き込む。"""
    col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
    body = f'<c r="{cell}"><v>{value}</v></c>'
    m = re.search(r'<c r="%s"[^>]*?(/>|>.*?</c>)' % cell, xml, re.S)
    if m:
        return xml[: m.start()] + body + xml[m.end():]
    rm = re.search(r'<row r="%s"[^>]*?(/>|>.*?</row>)' % row, xml, re.S)
    if not rm:
        log.append(f"  [skip] 設定行{row} が存在しません ({cell})")
        return xml
    chunk = rm.group(0)
    if chunk.endswith("/>"):
        new = chunk[:-2] + ">" + body + "</row>"
    else:
        target = colnum(col)
        ins = next((cm.start() for cm in re.finditer(r'<c r="([A-Z]+)%s"' % row, chunk)
                    if colnum(cm.group(1)) > target), None)
        new = chunk[:ins] + body + chunk[ins:] if ins is not None else \
              chunk[: chunk.rindex("</row>")] + body + "</row>"
    return xml[: rm.start()] + new + xml[rm.end():]


def rows_present(xml: str, col: str):
    return sorted(int(r) for r in re.findall(r'<c r="%s(\d+)"' % col, xml))


def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- 設定ブロック
# 458行のセルを直接書き換えるのではなく、1か所の設定セルを参照させる。
# 以後のパラメータ変更は AG2:AG6 を触るだけで全行に反映される。
SETTINGS = [
    ("AF1", "⚙ リスク設定（AG列を変更）", None),
    ("AF2", "損切幅", 0.03),      # MAE分析: 勝ちトレードのMAE中央値-0.89%、負けは-8.71%
    ("AF3", "利確幅", 0.08),      # グリッドサーチ: 損切3%なら利確8%がPF最良
    ("AF4", "1回の許容損失額(円)", 20000),
    ("AF5", "時間決済(営業日)", 3),   # 保有日数とリターンは単調悪化(-0.42%→-3.41%→-6.93%)
    ("AF6", "片道手数料率", 0.0005),
    ("AF7", "貸株料(年率)", 0.0115),
]
P_SL, P_TP, P_RISK, P_DAYS, P_FEE, P_KASHI = (
    "$AG$2", "$AG$3", "$AG$4", "$AG$5", "$AG$6", "$AG$7")


# ---------------------------------------------------------------- 数式定義
def _tick(raw: str) -> str:
    """東証の呼値（TOPIX100以外の一般株）。raw の価格帯から刻み幅を返す式。"""
    return (f'IF({raw}<=3000,1,IF({raw}<=5000,5,IF({raw}<=30000,10,'
            f'IF({raw}<=50000,50,IF({raw}<=300000,100,500)))))')


def _snap(raw: str, r: int) -> str:
    """呼値単位に丸める。買いは切り下げ／売りは切り上げ＝どちらも
    「利確は約定しやすい側・損切は余裕のある側」に倒れる。
    V805 は ROUND(...,0) で1円丸めしていたため、145件中71件(49%)が
    呼値に乗らず逆指値注文が発注できない価格になっていた。"""
    t = _tick(raw)
    return f'IF($C{r}="売",CEILING({raw},{t}),FLOOR({raw},{t}))'


def f_tp(r: int) -> str:
    """【修正0a】利確ライン：±6%のハードコードを設定セル参照に変更し、呼値単位へスナップ。"""
    raw = f'$F{r}*IF($C{r}="売",1-{P_TP},1+{P_TP})'
    return esc(f'IF($F{r}="","",{_snap(raw, r)})')


def f_sl(r: int) -> str:
    """【修正0b】損切ライン：∓4%のハードコードを設定セル参照に変更し、呼値単位へスナップ。"""
    raw = f'$F{r}*IF($C{r}="売",1+{P_SL},1-{P_SL})'
    return esc(f'IF($F{r}="","",{_snap(raw, r)})')


def f_trailing(r: int) -> str:
    """【修正1】トレーリングSTOP：売買方向を判定し、必ず固定損切より不利にならない位置に置く。
    含み益が +3% を超えるまでは固定損切 N 列をそのまま使い、発動後のみ建値方向にラチェットする。"""
    return esc(
        f'IF(OR($F{r}="",$H{r}="",$N{r}=""),"",'
        f'IF($S{r}="✅売却済","",'
        f'IF($C{r}="売",'
        f'  IF($H{r}>$F{r}*(1-{P_SL}),$N{r},CEILING(MIN($N{r},IF($Q{r}="",$N{r},$Q{r}*1.03)),{_tick("$N"+str(r))})),'
        f'  IF($H{r}<$F{r}*(1+{P_SL}),$N{r},FLOOR(MAX($N{r},IF($P{r}="",$N{r},$P{r}*0.97)),{_tick("$N"+str(r))}))'
        f')))'
    )


def f_judge(r: int) -> str:
    """【修正2】判断：決済済みの行を本日値で再計算しない（過去記録の破壊を防ぐ）。
    さらにトレーリングSTOP到達を判定に組み込む。"""
    return esc(
        f'IF(OR($B{r}="",$F{r}="",$H{r}=""),"",'
        f'IF($S{r}="✅売却済","―決済済―",'
        f'IF($C{r}="売",'
        f'  IF($H{r}<=$M{r},"🟢利確",'
        f'  IF(AND($R{r}<>"",$H{r}>=$R{r}),"🟠TRAIL決済",'
        f'  IF($H{r}>=$N{r},"🔴損切",'
        f'  IF($W{r}>={P_DAYS},"⏱時間決済",'
        f'  IF($H{r}>=$F{r}*(1+{P_SL}*0.7),"⚠️損切接近","保持継続"))))),'
        f'  IF($H{r}>=$M{r},"🟢利確",'
        f'  IF(AND($R{r}<>"",$H{r}<=$R{r}),"🟠TRAIL決済",'
        f'  IF($H{r}<=$N{r},"🔴損切",'
        f'  IF($W{r}>={P_DAYS},"⏱時間決済",'
        f'  IF($H{r}<=$F{r}*(1-{P_SL}*0.7),"⚠️損切接近","保持継続")))))))'
    )


def f_dup(r: int) -> str:
    """【修正3】Y列：同一銘柄・同一方向で建玉が既に存在する場合に警告（ナンピン防止）。"""
    return esc(
        f'IF($B{r}="","",'
        f'IF(COUNTIFS($B$4:$B{r},$B{r},$C$4:$C{r},$C{r})>1,'
        f'"⛔重複"&COUNTIFS($B$4:$B{r},$B{r},$C$4:$C{r},$C{r})&"回目",""))'
    )


def f_qty(r: int) -> str:
    """【修正4】AB列：リスク均等化した推奨株数。1トレードの損失額を一定にする。
    株数 = 許容損失額 ÷ (株価 × 損切幅) を 100株単位に丸める。"""
    return esc(
        f'IF($F{r}="","",'
        f'MAX(100,ROUNDDOWN({P_RISK}/($F{r}*{P_SL})/100,0)*100))'
    )


def f_netpnl(r: int) -> str:
    """【修正5】AC列：手数料・貸株料を控除した実質損益。"""
    return esc(
        f'IF(OR($F{r}="",$G{r}="",$U{r}=""),"",'
        f'ROUND($V{r}-($F{r}+$U{r})*$G{r}*{P_FEE}'
        f'-IF($C{r}="売",$F{r}*$G{r}*{P_KASHI}*$W{r}/365,0),0))'
    )


def f_buy_score(r: int) -> str:
    """【修正6】分析!T列 買いスコア：RSI項を「押し目狙い」から「トレンド健全性＋過熱ペナルティ」に変更。
    旧式は RSI<=30 に +3 点を与えていたが、モメンタム項(高値更新+パーフェクト▲)と同時成立しないため
    実質デッドコードだった。新式は RSI75超を減点し、高値掴みを構造的に排除する。"""
    return esc(
        f'IF(C{r}="","",IFERROR('
        f'IF(O{r}>=2,3,IF(O{r}>=1.5,2,IF(O{r}>=1,1,0)))'          # 出来高倍率
        f'+IF(E{r}>G{r},1,0)'                                       # 前日比プラス
        f'+IF(L{r}>=-5,1,0)'                                        # 年初来高値近接
        f'+IF(Q{r}="買い▲",2,0)'
        f'+IF(W{r}>=78,-4,IF(W{r}>=72,-2,IF(W{r}>=55,2,IF(W{r}>=45,1,0))))'   # ★RSI項を反転
        f'+IF(X{r}>0,1,0)'                                          # MACD
        f'+IF(V{r}="年初来高値更新●",3,IF(V{r}="高値圏",2,0))'
        f'+IF(AD{r}="パーフェクト▲",3,IF(AD{r}="上昇トレンド▲",2,0))'
        f'+IF(AE{r}="25日線押し目●",2,IF(AE{r}="5日線押し目○",1,0))'
        f'+IF(P{r}="急増↑↑",2,IF(P{r}="増加↑",1,0))'
        f'+N(AI{r}),0))'
    )


def f_sell_score(r: int) -> str:
    """【修正7】売分析!T列 売りスコア：RSI項を「戻り売り狙い」に変更。
    旧式は RSI>=70 に +3 点を与えていたが、年初来安値項と同時成立しないため実質デッドコードだった。
    新式は RSI25未満（売られ過ぎ＝踏み上げリスク）を大きく減点する。"""
    return esc(
        f'IF(C{r}="","",IFERROR('
        f'IF(W{r}<=22,-4,IF(W{r}<=28,-2,IF(W{r}<=45,2,IF(W{r}<=55,1,0))))'    # ★RSI項を反転
        f'+IF(X{r}<0,2,0)'
        f'+IF(AD{r}="下降トレンド▼",3,IF(AD{r}="もみ合い→",1,0))'
        f'+IF(V{r}="年初来安値圏●",3,IF(V{r}="安値圏",2,0))'
        f'+IF(Q{r}="売り▼",2,0)'
        f'+IF(P{r}="急増↑↑",2,IF(P{r}="増加↑",1,0))'
        f'+IF(Z{r}<0,1,0)'
        f'+IF(L{r}<=-30,2,IF(L{r}<=-15,1,0))'
        f'+IF(AE{r}="25日戻り売り★",2,IF(AE{r}="5日戻り売り○",1,0))'
        f'+N(AL{r})+N(AP{r})+N(AU{r}),0))'
    )


# ---------------------------------------------------------------- 実行
def main():
    shutil.copy(SRC, DST)
    zin = zipfile.ZipFile(SRC, "r")
    items = {n: zin.read(n) for n in zin.namelist()}
    order = zin.namelist()
    zin.close()

    # --- 管理シート ---
    x = items[KANRI].decode("utf-8")
    trade_rows = [r for r in rows_present(x, "N") if r >= 4]
    log.append(f"管理シート: 対象 {len(trade_rows)} 行 (行{min(trade_rows)}〜{max(trade_rows)})")
    for r in trade_rows:
        x = set_formula(x, f"M{r}", f_tp(r))
        x = set_formula(x, f"N{r}", f_sl(r))
        x = set_formula(x, f"R{r}", f_trailing(r))
        x = set_formula(x, f"O{r}", f_judge(r))
        x = set_formula(x, f"Y{r}", f_dup(r))
        x = set_formula(x, f"AB{r}", f_qty(r))
        x = set_formula(x, f"AC{r}", f_netpnl(r))
    for cell, text in (("Y3", "重複警告"), ("AB3", "推奨株数"), ("AC3", "実質損益(手数料後)")):
        x = set_inline_text(x, cell, text)
    for cell, label, val in SETTINGS:
        x = set_inline_text(x, cell, label)
        if val is not None:
            x = set_number(x, "AG" + cell[2:], val)
    log.append("  修正0 M/N列 損切・利確を設定セル参照化＋東証呼値へスナップ（損切3%・利確8%）")
    items[KANRI] = x.encode("utf-8")
    log.append("  修正1 R列 トレーリングSTOP 方向判定 … 適用")
    log.append("  修正2 O列 判断 決済済み除外 … 適用")
    log.append("  修正3 Y列 重複エントリー警告 … 追加")
    log.append("  修正4 AB列 リスク均等株数 … 追加")
    log.append("  修正5 AC列 コスト控除後損益 … 追加")
    log.append("  修正9 判断に「⏱時間決済」(AG5=3営業日) を追加")

    # --- 分析シート（買いスコア） ---
    x = items[BUNSEKI].decode("utf-8")
    rs = [r for r in rows_present(x, "T") if r >= 3]
    for r in rs:
        x = set_formula(x, f"T{r}", f_buy_score(r))
    items[BUNSEKI] = x.encode("utf-8")
    log.append(f"分析シート: 修正6 T列 買いスコア RSI項反転 … {len(rs)} 行に適用")

    # --- 売分析シート（売りスコア） ---
    x = items[URIBUNSEKI].decode("utf-8")
    rs = [r for r in rows_present(x, "T") if r >= 3]
    for r in rs:
        x = set_formula(x, f"T{r}", f_sell_score(r))
    items[URIBUNSEKI] = x.encode("utf-8")
    log.append(f"売分析シート: 修正7 T列 売りスコア RSI項反転 … {len(rs)} 行に適用")

    # --- 厳選TOP2：出来高フィルタを 1.0(無効) から 1.5 へ ---
    x = items[GENSEN].decode("utf-8")
    for cell in ("B9", "B13"):
        m = re.search(r'(<c r="%s"[^>]*>)(.*?)</c>' % cell, x, re.S)
        if m:
            x = x[: m.start()] + m.group(1) + "<v>1.5</v></c>" + x[m.end():]
    items[GENSEN] = x.encode("utf-8")
    log.append("厳選TOP2: 修正8 出来高倍率フィルタ 1.0 → 1.5 (該当0件の日は見送り)")

    # --- calcChain は数式追加でキャッシュ不整合になるため参照ごと除去 ---
    if "xl/calcChain.xml" in items:
        ct = items["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        items["[Content_Types].xml"] = ct.encode("utf-8")
        wr = items["xl/_rels/workbook.xml.rels"].decode("utf-8")
        wr = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", wr)
        items["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")
        log.append("calcChain.xml を除去（開いた時点でExcelが全再計算）")

    # --- 開いた瞬間にフル再計算させる ---
    wbx = items["xl/workbook.xml"].decode("utf-8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="0" fullCalcOnLoad="1"/>', wbx)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    items["xl/workbook.xml"] = wbx.encode("utf-8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            if n == "xl/calcChain.xml":
                continue
            zout.writestr(n, items[n])
    log.append(f"\n出力: {DST}  ({DST.stat().st_size:,} bytes)")
    print("\n".join(log))


if __name__ == "__main__":
    main()
