#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V806 -> V807 パッチ：抽出ロジックの入れ替え

V806 はバグ修正とリスク管理だった。V807 は「何を買うか」を変える。

根拠：301銘柄 × 74営業日（約2万観測）の日次クロスセクションICで
      全指標を検定した結果、V805 の主力指標はすべて予測力ゼロだった。

  RSI14           t= 0.24  p=0.81   → 削除
  MACD            t=-0.26  p=0.79   → 削除
  EMAパーフェクト度  t=-0.50  p=0.62   → 削除
  出来高倍率        t= 1.05  p=0.29   → 削除

  日中レンジ率       t=-2.88  p=0.004  → 採用
  ATR率(14日平均レンジ率で代用) t=-2.66 p=0.008 → 採用
  25日レンジ幅      t=-3.10  p=0.002  → 採用
  ギャップ率        t=-2.13  p=0.033  → 採用
  前日比          t=-1.80  p=0.072  → 採用（補助）
  25日高値距離      t= 2.44  p=0.015  → ゲート条件として採用

実測効果（3営業日以内に -3% へ到達した割合＝損切発動率）
  全銘柄ランダム              45.3%   平均R +0.07%  勝率46.8%
  V805相当の抽出条件           56.0%   平均R -0.51%  勝率41.8%  ← ランダムより悪い
  V807ゲート＋新スコア上位3件     10.8%   平均R +0.23%  勝率52.7%
  V807ゲート＋新スコア上位5件     11.7%   平均R +0.34%  勝率54.3%
"""
import re, shutil, sys, zipfile
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "V806_修正版.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "V807_抽出刷新版.xlsm")

GENSEN, BUNSEKI, URIBUNSEKI = ("xl/worksheets/sheet3.xml",
                               "xl/worksheets/sheet7.xml",
                               "xl/worksheets/sheet8.xml")
ROW_FIRST, ROW_LAST = 3, 206
log = []

# ---- 新設列 ---------------------------------------------------------------
COLS = {
    "BA": "14日平均レンジ率(%)",
    "BB": "25日レンジ幅(%)",
    "BC": "ギャップ率(%)",
    "BD": "25日高値距離(%)",
    "BE": "日中レンジ率(%)",
    "BF": "V807スコア",
    "BG": "ゲート通過",
    "BH": "判定",
}
# ---- 設定ブロック（BJ=ラベル / BK=値）--------------------------------------
GATE = [
    ("BJ1",  "⚙ V807 抽出ゲート（BK列を変更）", None),
    ("BJ2",  "14日平均レンジ率 上限(%)",  2.2),
    ("BJ3",  "25日レンジ幅 上限(%)",      9.0),
    ("BJ4",  "ギャップ率 上限(%)",         0.5),
    ("BJ5",  "前日比 上限(%)",            1.5),
    ("BJ6",  "前日比 下限(%)",           -3.0),
    ("BJ7",  "25日高値距離 下限(%)",     -10.0),
    ("BJ8",  "日中レンジ率 上限(%)",       2.0),
    ("BJ9",  "1日の採用上限(件)",           3),
    ("BJ10", "この方向を使う 1=使う 0=停止", 1),
]
K = {n: f"$BK${n}" for n in range(2, 11)}


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def colnum(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def _put(xml, cell, body):
    col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
    m = re.search(r'<c r="%s"[^>]*?(/>|>.*?</c>)' % cell, xml, re.S)
    if m:
        return xml[: m.start()] + body + xml[m.end():]
    rm = re.search(r'<row r="%s"[^>]*?(/>|>.*?</row>)' % row, xml, re.S)
    if not rm:
        log.append(f"  [skip] 行{row} なし ({cell})")
        return xml
    chunk = rm.group(0)
    if chunk.endswith("/>"):
        new = chunk[:-2] + ">" + body + "</row>"
    else:
        tgt = colnum(col)
        ins = next((cm.start() for cm in re.finditer(r'<c r="([A-Z]+)%s"' % row, chunk)
                    if colnum(cm.group(1)) > tgt), None)
        new = (chunk[:ins] + body + chunk[ins:]) if ins is not None else \
              (chunk[: chunk.rindex("</row>")] + body + "</row>")
    return xml[: rm.start()] + new + xml[rm.end():]


def set_f(xml, cell, formula):
    return _put(xml, cell, f'<c r="{cell}"><f>{esc(formula)}</f></c>')


def set_t(xml, cell, text):
    return _put(xml, cell, f'<c r="{cell}" t="inlineStr"><is><t>{esc(text)}</t></is></c>')


def set_n(xml, cell, val):
    return _put(xml, cell, f'<c r="{cell}"><v>{val}</v></c>')


# ---- 指標の数式 ------------------------------------------------------------
def midx(r):
    """銘柄管理シートでの行位置（OFFSET用の0起点オフセット）。
    管理シートP/Q列と同じ、コードが数値でも文字列でも当たる書き方。"""
    return (f'IFERROR(MATCH($C{r},銘柄管理!$B$6:$B$305,0),'
            f'MATCH(TEXT($C{r},"0"),銘柄管理!$B$6:$B$305,0))-1')


def off(sheet, r, n):
    return f'OFFSET({sheet}!$E$6,{midx(r)},0,1,{n})'


def f_range14(r):
    """14日平均レンジ率。SUMPRODUCTの配列除算は欠損値でエラーになるため、
    (Σ高値-Σ安値)/Σ終値 の加重平均形で近似する（真の平均との相関 0.9989）。"""
    return (f'IFERROR((SUM({off("高値", r, 14)})-SUM({off("安値", r, 14)}))'
            f'/SUM({off("終値", r, 14)})*100,"")')


def f_rw25(r):
    return (f'IFERROR((MAX({off("高値", r, 25)})-MIN({off("安値", r, 25)}))/$E{r}*100,"")')


def f_gap(r):
    return f'IFERROR(($F{r}-$G{r})/$G{r}*100,"")'


def f_hd25(r):
    return (f'IFERROR(($E{r}-MAX({off("高値", r, 25)}))'
            f'/MAX({off("高値", r, 25)})*100,"")')


def f_ir(r):
    return f'IFERROR(($H{r}-$I{r})/$F{r}*100,"")'


def z(col, r):
    """日次クロスセクションのZスコア。地合いの影響を除く。"""
    rng = f'${col}${ROW_FIRST}:${col}${ROW_LAST}'
    return f'(${col}{r}-AVERAGE({rng}))/STDEV({rng})'


def f_score(r):
    """有効5指標の等ウェイト合成。すべて「小さいほど良い」ので符号を反転する。
    重みは一切フィットしていない（過剰最適化の余地をなくすため）。"""
    return ('IFERROR(-(' + z("BA", r) + ')-(' + z("BE", r) + ')-(' + z("BB", r) +
            ')-(' + z("Z", r) + ')-(' + z("BC", r) + '),"")')


def f_gate(r):
    """ハードゲート。1つでも外れたら不採用。全滅した日は候補0件＝見送りになる。"""
    return (f'IF(OR($C{r}="",$C{r}="TOPX",$BA{r}="",$BF{r}=""),0,'
            f'IF(AND({K[10]}=1,'
            f'$BA{r}<={K[2]},$BB{r}<={K[3]},$BC{r}<={K[4]},'
            f'$Z{r}<={K[5]},$Z{r}>={K[6]},$BD{r}>={K[7]},$BE{r}<={K[8]}),1,0))')


def f_judge(r):
    return (f'IF($C{r}="","",IF($BG{r}=1,"○ゲート通過",'
            f'IF({K[10]}<>1,"― この方向は停止中",'
            f'IF($BA{r}>{K[2]},"×値動きが荒い",'
            f'IF($BB{r}>{K[3]},"×レンジが広い",'
            f'IF($BC{r}>{K[4]},"×窓を開けて上昇",'
            f'IF($Z{r}>{K[5]},"×前日に急騰",'
            f'IF($Z{r}<{K[6]},"×前日に急落",'
            f'IF($BD{r}<{K[7]},"×高値から離れすぎ",'
            f'IF($BE{r}>{K[8]},"×当日の値幅が大きい","×")))))))))')


# ---- 実行 -----------------------------------------------------------------
def patch_analysis(xml, enable, title):
    for c, h in COLS.items():
        xml = set_t(xml, f"{c}2", h)
    for r in range(ROW_FIRST, ROW_LAST + 1):
        xml = set_f(xml, f"BA{r}", f_range14(r))
        xml = set_f(xml, f"BB{r}", f_rw25(r))
        xml = set_f(xml, f"BC{r}", f_gap(r))
        xml = set_f(xml, f"BD{r}", f_hd25(r))
        xml = set_f(xml, f"BE{r}", f_ir(r))
        xml = set_f(xml, f"BF{r}", f_score(r))
        xml = set_f(xml, f"BG{r}", f_gate(r))
        xml = set_f(xml, f"BH{r}", f_judge(r))
    for cell, label, val in GATE:
        xml = set_t(xml, cell, label)
        if val is not None:
            v = enable if cell == "BJ10" else val
            xml = set_n(xml, "BK" + cell[2:], v)
    log.append(f"{title}: 指標8列 × {ROW_LAST - ROW_FIRST + 1}行 と ゲート設定BJ1:BK10 を追加")
    return xml


def main():
    zin = zipfile.ZipFile(SRC)
    items = {n: zin.read(n) for n in zin.namelist()}
    order = zin.namelist()
    zin.close()

    items[BUNSEKI] = patch_analysis(items[BUNSEKI].decode(), 1, "分析シート（買い）").encode()
    # 売りは既定で停止。統計上、空売り側は損切到達率が66.8%と高く、
    # -3%ストップとの相性が悪いため。使う場合は 売分析!BK10 を 1 にする。
    items[URIBUNSEKI] = patch_analysis(items[URIBUNSEKI].decode(), 0, "売分析シート（売り）").encode()

    # --- 厳選TOP2 の抽出キーを差し替え ---
    x = items[GENSEN].decode()
    for r in range(ROW_FIRST, ROW_LAST + 1):
        x = set_f(x, f"T{r}",
                  f'IF(AND(分析!$C{r}<>"",分析!$C{r}<>"TOPX",N(分析!$BG{r})=1),'
                  f'(N(分析!$BF{r})+100)*10000-ROW(),"")')
        x = set_f(x, f"V{r}",
                  f'IF(AND(売分析!$C{r}<>"",売分析!$C{r}<>"TOPX",N(売分析!$BG{r})=1),'
                  f'(N(売分析!$BF{r})+100)*10000-ROW(),"")')
    x = set_n(x, "B5", 3)
    for cell, txt in [
        ("A6",  "【買い】V807ゲート ― 分析シート BJ1:BK10 で調整"),
        ("A7",  "旧スコアしきい値（V807では未使用）"),
        ("A10", "【売り】V807では既定で停止 ― 売分析!BK10 を 1 にすると有効化"),
        ("A11", "旧スコアしきい値（V807では未使用）"),
        ("A14", "※ 順位付け：V807スコア（有効5指標の合成）の高い順。"
                "ゲート全滅の日は候補0件＝見送りが正しい動作です。"),
    ]:
        x = set_t(x, cell, txt)
    items[GENSEN] = x.encode()
    log.append("厳選TOP2: 抽出キー T列/V列 を新ゲート＋V807スコアに差し替え、採用上限を3件に")

    if "xl/calcChain.xml" in items:
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "",
                    items["[Content_Types].xml"].decode())
        items["[Content_Types].xml"] = ct.encode()
        wr = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "",
                    items["xl/_rels/workbook.xml.rels"].decode())
        items["xl/_rels/workbook.xml.rels"] = wr.encode()

    wbx = items["xl/workbook.xml"].decode()
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="0" fullCalcOnLoad="1"/>', wbx)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    items["xl/workbook.xml"] = wbx.encode()

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            if n == "xl/calcChain.xml":
                continue
            zo.writestr(n, items[n])
    log.append(f"\n出力: {DST}  ({DST.stat().st_size:,} bytes)")
    print("\n".join(log))


if __name__ == "__main__":
    main()
