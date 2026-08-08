#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BB816 ビルダー ― ボリンジャーバンド判定の独立ブック（スイング用に組み直した版）

■ BB816 とは

  V815 から切り離したボリンジャーバンド判定を、300銘柄×250営業日の実測に基づいて
  スイングトレード用に組み直した単独ブック。V815/V816 が無くても単独で動く。

      BB816使い方   ルール・根拠・2人の分析官の意見・最終判断
      BB816候補     明日の買い候補 TOP10（発注に必要な数字だけ）
      BB816判定     全銘柄の判定本体
      BB816設定     しきい値（青字セルで変更可）
      終値/高値/安値/出来高   判定に使う価格データ（250営業日・値のみ）

■ V815 の BB判定から変えたところ（すべて実測が根拠）

  1. 逆張り系のシグナルを捨てた
     スクイーズ★（超過-0.48%）、強気ダイバージェンス（-0.41%）、%B<0.2（-1.16%）、
     RSI<40（-1.21%）はいずれも超過リターンがマイナスだった。
     二番底・二番天井の判定（BB_スイング）は、入れても外しても成績が変わらなかった
     （資産+116.5% → +114.4%、超過+2.62% → +2.61%）ので、5MBの計算ごと落とした。

  2. 順張り系に寄せた
     +2σ超え（超過+1.11%）、出来高2倍（+1.14%）、RSI>70（+1.34%）、
     MFI60-80（+0.96%）、20日騰落率上位（+1.52%）。いずれも前半/後半の両方で残った。

  3. スコアを作り直した
     V815のスコアは点数と成績が対応していない（40点以上でむしろ超過ゼロ）。
     BB816スコアは実測の超過リターンに比例させた配点にしたので単調になる。
        0-19点 -0.79% / 20-39 +0.06% / 40-59 +0.76% / 60-69 +1.27% / 70-79 +1.65% / 80-100 +3.37%

  4. 出口を決めた
     10営業日で手仕舞い、損切り-8%。利確目標とトレーリングは置かない
     （+12%利確で期待値が +3.23% → +1.24% に落ちるため）。

  5. 地合いフィルタは入れない
     breadth>=50%・指数>20日平均のどちらも、下落局面の回避に役立たなかった
     （下落局面の平均が -1.21% → -3.31% とかえって悪化）。

  6. 枠は8（1銘柄12.5%）、1日の新規は最大3銘柄
     枠3/1日1銘柄では最大DD -37.0%・Sharpe 0.74。枠8/1日3銘柄で -13.5%・2.26。
     分散はリターンとDDを同時に改善した。

  検証結果（枠8・1日3銘柄・10日保有・損切-8%、250営業日）
      BB816     総リターン +116.5%  最大DD -22.9%  Sharpe 2.29  1トレード最悪 -8.0%
      現行スコア  総リターン  +48.8%  最大DD -19.6%  Sharpe 1.54
      無選別     総リターン  +57.5%  最大DD -28.9%  Sharpe 1.51

■ 作り方の方針

  ・価格シートは V815 からコピーし、数式を全部キャッシュ値に固定する
    （楽天RSSアドインが無くても開ける／銘柄管理シートに依存しない）
  ・判定シートは新規に生成する。使う関数は V815 に実在するものだけ
    （AVERAGE / STDEVP / SUMPRODUCT / IF / IFERROR / AND / OR / MAX / MIN /
      ROUND / INDEX / MATCH / LARGE / ROW / HYPERLINK / COUNT / COUNTIF）
  ・styles.xml は V815 のものを引き継ぎ、必要な書式は末尾に追記する
    （既存のインデックスがずれないので、価格シートの見た目が保たれる）

使い方:  python3 tools/build_bb816.py V815.xlsm BB816.xlsx
"""
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "BB816.xlsx")

ROW1, ROWN = 6, 505          # 判定シートの銘柄行
TOPN = 10                    # 候補シートに出す件数
PRICE = [("終値", "sheet14"), ("高値", "sheet12"), ("安値", "sheet13"), ("出来高", "sheet15")]

log = []


# ---------------------------------------------------------------------------
# XML の細かい部品
# ---------------------------------------------------------------------------
def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def cell(ref, *, f=None, v=None, s=None, text=None):
    a = ' r="%s"' % ref
    if s is not None:
        a += ' s="%d"' % s
    if f is not None:
        return "<c%s><f>%s</f></c>" % (a, esc(f))
    if text is not None:
        return '<c%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (a, esc(text))
    if v is not None:
        return "<c%s><v>%s</v></c>" % (a, v)
    return "<c%s/>" % a


def row(n, cells, ht=None):
    a = ' r="%d"' % n
    if ht:
        a += ' ht="%s" customHeight="1"' % ht
    return "<row%s>%s</row>" % (a, "".join(cells))


def sheet_xml(dimension, rows, cols=None, freeze=None, tab_selected=False, grid=False):
    view = '<sheetView%s showGridLines="%d" workbookViewId="0">' % (
        ' tabSelected="1"' if tab_selected else "", 1 if grid else 0)
    if freeze:
        col, rw = freeze
        view += ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s%d" activePane="bottomRight" state="frozen"/>'
                 % (col, rw, chr(65 + col), rw + 1))
    view += "</sheetView>"
    colxml = "<cols>%s</cols>" % "".join(cols) if cols else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="%s"/><sheetViews>%s</sheetViews>'
            '<sheetFormatPr defaultRowHeight="18"/>%s<sheetData>%s</sheetData>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            '</worksheet>' % (dimension, view, colxml, "".join(rows))).encode("utf-8")


def col_def(a, b, w):
    return '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (a, b, w)


# ---------------------------------------------------------------------------
# styles.xml への追記
# ---------------------------------------------------------------------------
def extend_styles(xml):
    """既存インデックスを壊さずに、末尾へ書式を足す。"""
    n_font = int(re.search(r'<fonts count="(\d+)"', xml).group(1))
    n_fill = int(re.search(r'<fills count="(\d+)"', xml).group(1))
    n_bord = int(re.search(r'<borders count="(\d+)"', xml).group(1))
    n_xf = int(re.search(r'<cellXfs count="(\d+)"', xml).group(1))
    n_fmt = int(re.search(r'<numFmts count="(\d+)"', xml).group(1))

    FONT = "ＭＳ Ｐゴシック"
    new_fonts = [
        '<font><b/><sz val="14"/><color rgb="FF1F3864"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
        '<font><b/><sz val="11"/><color theme="1"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
        '<font><b/><sz val="11"/><color rgb="FF0070C0"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
        '<font><sz val="10"/><color rgb="FF595959"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
        '<font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % FONT,
    ]
    F_TITLE, F_WHITE, F_BOLD, F_BLUE, F_GRAY, F_SEC = range(n_font, n_font + 6)

    new_fills = [
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>',
    ]
    FL_HDR, FL_BLUE, FL_GRAY, FL_EDIT, FL_BUY = range(n_fill, n_fill + 5)

    new_borders = ['<border><left style="thin"><color rgb="FFBFBFBF"/></left>'
                   '<right style="thin"><color rgb="FFBFBFBF"/></right>'
                   '<top style="thin"><color rgb="FFBFBFBF"/></top>'
                   '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>']
    B_THIN = n_bord

    new_fmts = ['<numFmt numFmtId="200" formatCode="0.0"/>',
                '<numFmt numFmtId="201" formatCode="0.0%"/>',
                '<numFmt numFmtId="202" formatCode="0.00&quot;倍&quot;"/>']

    def xf(fmt=0, font=0, fill=0, bord=0, align=None):
        a = ('numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"'
             ' applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"' % (fmt, font, fill, bord))
        if align:
            return "<xf %s applyAlignment=\"1\"><alignment %s/></xf>" % (a, align)
        return "<xf %s/>" % a

    ctr = 'horizontal="center" vertical="center"'
    wrap = 'horizontal="center" vertical="center" wrapText="1"'
    new_xfs = [
        xf(font=F_TITLE),                                             # 0 タイトル
        xf(font=F_GRAY),                                              # 1 注記
        xf(font=F_WHITE, fill=FL_HDR, bord=B_THIN, align=wrap),       # 2 見出し
        xf(bord=B_THIN),                                              # 3 文字
        xf(fmt=200, bord=B_THIN),                                     # 4 小数1桁
        xf(fmt=3, bord=B_THIN),                                       # 5 整数カンマ
        xf(fmt=201, bord=B_THIN),                                     # 6 パーセント
        xf(font=F_BOLD, bord=B_THIN, align=ctr),                      # 7 スコア
        xf(font=F_BOLD, fill=FL_BUY, bord=B_THIN, align=ctr),         # 8 判定（買い）
        xf(font=F_BLUE, fill=FL_EDIT, bord=B_THIN, align=ctr, fmt=200),  # 9 設定値
        xf(font=F_SEC, fill=FL_BLUE),                                 # 10 セクション
        xf(bord=B_THIN, align=ctr),                                   # 11 中央
        xf(font=F_BOLD),                                              # 12 太字
        xf(fmt=202, bord=B_THIN),                                     # 13 出来高倍率
        xf(fmt=14, bord=B_THIN, align=ctr),                           # 14 日付
        xf(font=F_SEC),                                               # 15 小見出し
    ]
    S = {k: n_xf + i for i, k in enumerate(
        ["TITLE", "NOTE", "HDR", "TXT", "NUM1", "NUM0", "PCT", "SCORE", "BUY",
         "SET", "SECTION", "CTR", "BOLD", "VOL", "DATE", "SUB"])}

    xml = xml.replace('<numFmts count="%d">' % n_fmt,
                      '<numFmts count="%d">' % (n_fmt + len(new_fmts)), 1)
    xml = xml.replace("</numFmts>", "".join(new_fmts) + "</numFmts>", 1)
    xml = xml.replace('<fonts count="%d"' % n_font, '<fonts count="%d"' % (n_font + len(new_fonts)), 1)
    xml = xml.replace("</fonts>", "".join(new_fonts) + "</fonts>", 1)
    xml = xml.replace('<fills count="%d"' % n_fill, '<fills count="%d"' % (n_fill + len(new_fills)), 1)
    xml = xml.replace("</fills>", "".join(new_fills) + "</fills>", 1)
    xml = xml.replace('<borders count="%d"' % n_bord, '<borders count="%d"' % (n_bord + 1), 1)
    xml = xml.replace("</borders>", "".join(new_borders) + "</borders>", 1)
    xml = xml.replace('<cellXfs count="%d"' % n_xf, '<cellXfs count="%d"' % (n_xf + len(new_xfs)), 1)
    xml = xml.replace("</cellXfs>", "".join(new_xfs) + "</cellXfs>", 1)
    return xml, S


# ---------------------------------------------------------------------------
# 価格シート（V815 からコピーして値に固定）
# ---------------------------------------------------------------------------
def flatten(xml):
    xml = re.sub(r'<drawing\s+r:id="[^"]*"\s*/>', "", xml)
    xml = re.sub(r'<legacyDrawing\s+r:id="[^"]*"\s*/>', "", xml)
    xml = re.sub(r"<controls>.*?</controls>", "", xml, flags=re.S)
    xml = re.sub(r'(<pageSetup\b[^>]*?)\s+r:id="[^"]*"', r"\1", xml)
    xml = re.sub(r"<f[^>]*/>", "", xml)
    xml = re.sub(r"<f[^>]*>[^<]*</f>", "", xml)

    def to_inline(m):
        attrs, text = m.group(1) + m.group(2), m.group(3)
        if not text:
            return "<c%s/>" % attrs
        return '<c%s t="inlineStr"><is><t>%s</t></is></c>' % (attrs, text)

    xml = re.sub(r'<c([^>]*?)\st="str"([^>]*?)>(?:<v>(.*?)</v>)?</c>', to_inline, xml)
    xml = re.sub(r'<c([^>]*?)\st="e"([^>]*?)>(?:<v>.*?</v>)?</c>', r"<c\1\2/>", xml)
    xml = re.sub(r'\s+codeName="[^"]*"', "", xml, count=1)
    xml = re.sub(r'\s+tabSelected="1"', "", xml)
    return xml


# ---------------------------------------------------------------------------
# BB816設定
# ---------------------------------------------------------------------------
SETTINGS = [
    ("外側バンド σ倍率", 2, "±2σ。買い条件のブレイク判定に使う"),
    ("内側バンド σ倍率", 1, "±1σ。攻め型の判定と参考表示に使う"),
    ("買い 出来高倍率の下限", 1.5, "当日出来高 ÷ 直近5日平均。実測: 2倍以上が超過+1.14%、1.5倍で件数を確保"),
    ("買い MFI下限", 55, "資金流入の確認。MFI60〜80の超過は+0.96%"),
    ("買い RSI上限", 80, "過熱しすぎを除外。RSI>80は超過が伸びない"),
    ("買い スコア下限", 40, "BB816スコア。40未満の帯は超過がマイナス〜ゼロ"),
    ("監視 スコア下限", 60, "買い条件は未成立だが、翌日以降の候補として見ておく水準"),
    ("損切り %", 8, "建値からの下落率。-8%で1トレード最悪-8%、-10%以下の損失をゼロにできる"),
    ("保有 営業日数", 10, "10営業日で手仕舞い。15日に伸ばすと枠が詰まり資金効率が落ちる"),
    ("資金の枠数", 8, "1銘柄あたり 1÷枠数。8枠なら12.5%。1日の新規は最大3銘柄"),
    ("守り 出来高倍率", 2, "リスク分析官の条件（厳しめ）"),
    ("守り MFI下限", 60, "同上"),
    ("守り RSI上限", 75, "同上"),
    ("攻め RSI下限", 60, "トレーダー分析官の条件（広め）"),
    ("攻め MFI下限", 60, "同上"),
    ("攻め 出来高倍率", 1.2, "同上"),
]
SET_ROW = {name: 4 + i for i, name in enumerate(n for n, _v, _d in SETTINGS)}


def R(name):
    return "BB816設定!$B$%d" % SET_ROW[name]


def build_settings(S):
    rows = [row(1, [cell("A1", text="⚙ BB816 設定（青字のセルは変更できます）", s=S["TITLE"])], ht="24"),
            row(2, [cell("A2", text="ここを変えると BB816判定 と BB816候補 の判定が変わります。既定値は300銘柄×250営業日の実測から決めた値です。"
                                   "／ BB816スコアの配点そのものは数式に固定してあります（実測の超過リターンから決めた重みのため）。"
                                   "下の設定で動くのは「買い条件・監視・守り/攻めの判定・損切り・枠数」です。", s=S["NOTE"])])]
    rows.append(row(3, [cell("A3", text="項目", s=S["HDR"]), cell("B3", text="値", s=S["HDR"]),
                        cell("C3", text="説明", s=S["HDR"])]))
    for i, (name, val, desc) in enumerate(SETTINGS):
        r = 4 + i
        rows.append(row(r, [cell("A%d" % r, text=name, s=S["TXT"]),
                            cell("B%d" % r, v=val, s=S["SET"]),
                            cell("C%d" % r, text=desc, s=S["TXT"])]))
    r = 4 + len(SETTINGS) + 1
    rows.append(row(r, [cell("A%d" % r, text="■ 1トレードで許容する損失", s=S["SECTION"])]))
    rows.append(row(r + 1, [cell("A%d" % (r + 1), text="損切り% ÷ 枠数 = 総資金に対する1トレードの最大損失", s=S["TXT"]),
                            cell("B%d" % (r + 1), f="%s/%s" % (R("損切り %"), R("資金の枠数")), s=S["NUM1"]),
                            cell("C%d" % (r + 1), text="既定値なら 8 ÷ 8 = 1.0%。ここが1%を超える設定にはしないこと", s=S["TXT"])]))
    cols = [col_def(1, 1, 26), col_def(2, 2, 10), col_def(3, 3, 78)]
    return sheet_xml("A1:C%d" % (r + 1), rows, cols)


# ---------------------------------------------------------------------------
# BB816判定
# ---------------------------------------------------------------------------
HEAD = [
    ("A", "🔗", 5, "TXT"), ("B", "コード", 8, "CTR"), ("C", "銘柄名", 20, "TXT"),
    ("D", "現在値", 10, "NUM1"), ("E", "SMA20\nミドル", 10, "NUM1"), ("F", "σ", 9, "NUM1"),
    ("G", "+2σ", 10, "NUM1"), ("H", "+1σ", 10, "NUM1"), ("I", "-1σ", 10, "NUM1"),
    ("J", "-2σ", 10, "NUM1"), ("K", "%B", 8, "NUM1"), ("L", "帯幅%", 8, "NUM1"),
    ("M", "RSI14", 8, "NUM1"), ("N", "MFI14", 8, "NUM1"), ("O", "出来高\n倍率", 9, "VOL"),
    ("P", "20日\n騰落率", 9, "PCT"), ("Q", "BB816\nスコア", 9, "SCORE"),
    ("R", "買い条件", 9, "CTR"), ("S", "総合判定", 12, "BUY"),
    ("T", "守り分析官", 11, "CTR"), ("U", "攻め分析官", 11, "CTR"),
    ("V", "エントリー\n目安", 11, "NUM1"), ("W", "損切り\n価格", 11, "NUM1"),
    ("X", "総資金への\n最大損失%", 11, "NUM1"), ("Y", "sortkey", 10, "NUM1"),
]

RSI = ('IFERROR(100-100/(1+SUMPRODUCT((終値!E{r}:R{r}>終値!F{r}:S{r})*(終値!E{r}:R{r}-終値!F{r}:S{r}))'
       '/SUMPRODUCT((終値!E{r}:R{r}<終値!F{r}:S{r})*(終値!F{r}:S{r}-終値!E{r}:R{r}))),100)')
MFI = ('IFERROR(100-100/(1+SUMPRODUCT((((高値!E{r}:R{r}+安値!E{r}:R{r}+終値!E{r}:R{r})/3)>'
       '((高値!F{r}:S{r}+安値!F{r}:S{r}+終値!F{r}:S{r})/3))*((高値!E{r}:R{r}+安値!E{r}:R{r}+終値!E{r}:R{r})/3)'
       '*出来高!E{r}:R{r})/SUMPRODUCT((((高値!E{r}:R{r}+安値!E{r}:R{r}+終値!E{r}:R{r})/3)<'
       '((高値!F{r}:S{r}+安値!F{r}:S{r}+終値!F{r}:S{r})/3))*((高値!E{r}:R{r}+安値!E{r}:R{r}+終値!E{r}:R{r})/3)'
       '*出来高!E{r}:R{r})),50)')

SCORE = (
    'MAX(0,MIN(100,'
    'IF($D{r}>$G{r},25,IF($D{r}>$H{r},15,IF($D{r}>$E{r},5,0)))'          # バンド位置
    '+IF($O{r}>=2,20,IF($O{r}>=1.5,14,IF($O{r}>=1.2,8,0)))'              # 出来高
    '+IF($N{r}>=80,15,IF($N{r}>=60,18,IF($N{r}>=50,8,0)))'               # MFI
    '+IF(AND($M{r}>70,$M{r}<=80),18,IF(AND($M{r}>60,$M{r}<=70),14,IF(AND($M{r}>50,$M{r}<=60),6,0)))'  # RSI
    '+IF($P{r}>=0.1,14,IF($P{r}>=0.03,7,0))'                             # 20日騰落率
    '+IF($M{r}>80,-12,0)+IF($M{r}<40,-20,0)+IF($N{r}<40,-15,0)'          # 減点
    '+IF($K{r}<0.2,-18,0)+IF($P{r}<=-0.1,-10,0)))')


def build_judge(S):
    rows = []
    rows.append(row(1, [cell("A1", text="📊 BB816 判定 ― ボリンジャーバンド＋出来高による順張りスイング", s=S["TITLE"])], ht="24"))
    rows.append(row(2, [cell("A2", text="終値シートの最新終値で自動計算。発注に使うのは「BB816候補」シート。"
                                       "ここは全銘柄の内訳を確認する場所です。", s=S["NOTE"])]))
    rows.append(row(3, [cell("A3", text="総合判定 ◎買い＝翌営業日の寄付で買う / ○監視＝条件待ち / ▲弱い＝買わない。"
                                       "守り分析官・攻め分析官の欄は、2人の分析官それぞれの基準での可否です。", s=S["NOTE"])]))
    hdr = [cell("%s5" % col, text=label, s=S["HDR"]) for col, label, _w, _s in HEAD]
    rows.append(row(5, hdr, ht="34"))

    for r in range(ROW1, ROWN + 1):
        g = 'IF($D{r}="","",{body})'
        c = []
        c.append(cell("A%d" % r, f='IF(終値!A%d="","",HYPERLINK("https://kabutan.jp/stock/?code="&終値!A%d,"📈"))' % (r, r), s=S["TXT"]))
        c.append(cell("B%d" % r, f='IF(終値!A{r}="","",終値!A{r})'.format(r=r), s=S["CTR"]))
        c.append(cell("C%d" % r, f='IF(終値!A{r}="","",終値!B{r})'.format(r=r), s=S["TXT"]))
        c.append(cell("D%d" % r, f='IF(終値!A{r}="","",終値!E{r})'.format(r=r), s=S["NUM1"]))
        c.append(cell("E%d" % r, f=g.format(r=r, body="AVERAGE(終値!E{r}:X{r})".format(r=r)), s=S["NUM1"]))
        c.append(cell("F%d" % r, f=g.format(r=r, body="STDEVP(終値!E{r}:X{r})".format(r=r)), s=S["NUM1"]))
        c.append(cell("G%d" % r, f=g.format(r=r, body="$E{r}+{p}*$F{r}".format(r=r, p=R("外側バンド σ倍率"))), s=S["NUM1"]))
        c.append(cell("H%d" % r, f=g.format(r=r, body="$E{r}+{p}*$F{r}".format(r=r, p=R("内側バンド σ倍率"))), s=S["NUM1"]))
        c.append(cell("I%d" % r, f=g.format(r=r, body="$E{r}-{p}*$F{r}".format(r=r, p=R("内側バンド σ倍率"))), s=S["NUM1"]))
        c.append(cell("J%d" % r, f=g.format(r=r, body="$E{r}-{p}*$F{r}".format(r=r, p=R("外側バンド σ倍率"))), s=S["NUM1"]))
        c.append(cell("K%d" % r, f=g.format(r=r, body='IFERROR(($D{r}-$J{r})/($G{r}-$J{r}),"")'.format(r=r)), s=S["NUM1"]))
        c.append(cell("L%d" % r, f=g.format(r=r, body='IFERROR(($G{r}-$J{r})/$E{r}*100,"")'.format(r=r)), s=S["NUM1"]))
        c.append(cell("M%d" % r, f=g.format(r=r, body=RSI.format(r=r)), s=S["NUM1"]))
        c.append(cell("N%d" % r, f=g.format(r=r, body=MFI.format(r=r)), s=S["NUM1"]))
        c.append(cell("O%d" % r, f=g.format(r=r, body='IFERROR(出来高!E{r}/AVERAGE(出来高!E{r}:I{r}),"")'.format(r=r)), s=S["VOL"]))
        c.append(cell("P%d" % r, f=g.format(r=r, body='IFERROR($D{r}/終値!Y{r}-1,"")'.format(r=r)), s=S["PCT"]))
        c.append(cell("Q%d" % r, f=g.format(r=r, body=SCORE.format(r=r)), s=S["SCORE"]))
        c.append(cell("R%d" % r, f=g.format(r=r, body=(
            'IF(AND($D{r}>$G{r},$O{r}>={v},$N{r}>={m},$M{r}<={s}),"○","－")'
            .format(r=r, v=R("買い 出来高倍率の下限"), m=R("買い MFI下限"), s=R("買い RSI上限")))), s=S["CTR"]))
        c.append(cell("S%d" % r, f=g.format(r=r, body=(
            'IF(AND($R{r}="○",$Q{r}>={b}),"◎買い",IF($Q{r}>={w},"○監視",IF(OR($M{r}<40,$K{r}<0.2),"▲弱い","－")))'
            .format(r=r, b=R("買い スコア下限"), w=R("監視 スコア下限")))), s=S["BUY"]))
        c.append(cell("T%d" % r, f=g.format(r=r, body=(
            'IF(AND($D{r}>$G{r},$O{r}>={v},$N{r}>={m},$M{r}<={s}),"買い可","見送り")'
            .format(r=r, v=R("守り 出来高倍率"), m=R("守り MFI下限"), s=R("守り RSI上限")))), s=S["CTR"]))
        c.append(cell("U%d" % r, f=g.format(r=r, body=(
            'IF(AND($D{r}>$H{r},$M{r}>={s},$N{r}>={m},$O{r}>={v}),"買い可","見送り")'
            .format(r=r, s=R("攻め RSI下限"), m=R("攻め MFI下限"), v=R("攻め 出来高倍率")))), s=S["CTR"]))
        c.append(cell("V%d" % r, f='IF($S{r}="◎買い",$D{r},"")'.format(r=r), s=S["NUM1"]))
        c.append(cell("W%d" % r, f='IF($S{r}="◎買い",ROUND($D{r}*(1-{p}/100),1),"")'.format(r=r, p=R("損切り %")), s=S["NUM1"]))
        c.append(cell("X%d" % r, f='IF($S{r}="◎買い",{p}/{k},"")'.format(r=r, p=R("損切り %"), k=R("資金の枠数")), s=S["NUM1"]))
        c.append(cell("Y%d" % r, f=g.format(r=r, body="$Q{r}+(10000-ROW())/100000".format(r=r)), s=S["NUM1"]))
        rows.append(row(r, c))

    cols = [col_def(i + 1, i + 1, w) for i, (_c, _l, w, _s) in enumerate(HEAD)]
    return sheet_xml("A1:Y%d" % ROWN, rows, cols, freeze=(3, 5))


# ---------------------------------------------------------------------------
# BB816候補
# ---------------------------------------------------------------------------
TOP_COLS = [
    ("A", "順位", 6, "CTR", None),
    ("B", "コード", 8, "CTR", "B"),
    ("C", "銘柄名", 20, "TXT", "C"),
    ("D", "現在値", 10, "NUM1", "D"),
    ("E", "BB816\nスコア", 9, "SCORE", "Q"),
    ("F", "総合判定", 12, "BUY", "S"),
    ("G", "守り\n分析官", 10, "CTR", "T"),
    ("H", "攻め\n分析官", 10, "CTR", "U"),
    ("I", "RSI14", 8, "NUM1", "M"),
    ("J", "MFI14", 8, "NUM1", "N"),
    ("K", "出来高\n倍率", 9, "VOL", "O"),
    ("L", "20日\n騰落率", 9, "PCT", "P"),
    ("M", "エントリー\n(翌日寄付)", 12, "NUM1", "V"),
    ("N", "損切り\n価格", 11, "NUM1", "W"),
    ("O", "株探", 7, "CTR", None),
]


def build_top(S):
    J = "BB816判定"
    rows = [row(1, [cell("A1", text="🏆 BB816 明日の買い候補", s=S["TITLE"])], ht="26")]
    rows.append(row(2, [cell("A2", text="総合判定が「◎買い」の銘柄を、翌営業日の寄付で買う。"
                                       "1日に建てるのは上から最大3銘柄まで。1銘柄は資金の1÷枠数（既定8枠＝12.5%）。", s=S["NOTE"])]))
    rows.append(row(3, [
        cell("A3", text="◎買いの件数", s=S["SUB"]),
        cell("C3", f='COUNTIF(%s!$S$%d:$S$%d,"◎買い")' % (J, ROW1, ROWN), s=S["CTR"]),
        cell("D3", text="SMA20より上の銘柄の割合（地合いの参考。売買条件には使いません）", s=S["NOTE"]),
        cell("H3", f='IFERROR(SUMPRODUCT((%s!$D$%d:$D$%d<>"")*(%s!$D$%d:$D$%d>%s!$E$%d:$E$%d))'
                     '/COUNT(%s!$D$%d:$D$%d),"")' % (J, ROW1, ROWN, J, ROW1, ROWN, J, ROW1, ROWN, J, ROW1, ROWN),
             s=S["PCT"]),
        cell("I3", text="データ最終日", s=S["NOTE"]),
        cell("K3", f="終値!E3", s=S["DATE"]),
    ]))
    rows.append(row(4, [cell("%s4" % c, text=lbl, s=S["HDR"]) for c, lbl, _w, _s, _src in TOP_COLS], ht="34"))

    for i in range(TOPN):
        r = 5 + i
        c = [cell("A%d" % r, v=i + 1, s=S["CTR"])]
        for col, _lbl, _w, st, src in TOP_COLS[1:]:
            if src is None:
                continue
            f = ('IFERROR(INDEX({j}!${s}${a}:${s}${b},MATCH(LARGE({j}!$Y${a}:$Y${b},$A{r}),'
                 '{j}!$Y${a}:$Y${b},0)),"")').format(j=J, s=src, a=ROW1, b=ROWN, r=r)
            c.append(cell("%s%d" % (col, r), f=f, s=S[st]))
        c.append(cell("O%d" % r, f='IF($B{r}="","",HYPERLINK("https://kabutan.jp/stock/?code="&$B{r},"📈"))'.format(r=r), s=S["CTR"]))
        rows.append(row(r, c))

    r = 5 + TOPN + 1
    notes = [
        "■ 発注の手順",
        "1. 「総合判定」が ◎買い の銘柄を上から最大3件選ぶ（空き枠がある分だけ）。",
        "2. 翌営業日の寄付（始値）で買う。指値にせず寄成でよい。",
        "3. 買えたら、その場で損切り価格（N列）に逆指値を置く。",
        "4. 10営業日たったら、損益に関わらず引けで手仕舞う。",
        "5. 利確目標は置かない。含み益が伸びても途中で降りない（検証では利確を置くと期待値が6割落ちた）。",
        "",
        "■ 守り分析官・攻め分析官の欄の読み方",
        "守り＝出来高2倍以上・MFI60以上・RSI75以下。厳しい代わりに1日1.5件しか出ない。",
        "攻め＝+1σ超えまで許容。1日14件出るが、実運用のSharpeは守りに負けた（1.36 vs 2.08）。",
        "採用しているのは両者の中間で、+2σ超え・出来高1.5倍・MFI55・RSI80以下。",
        "2人とも「買い可」なら、条件としては最も堅い候補になる。",
    ]
    for i, t in enumerate(notes):
        if not t:
            continue
        rows.append(row(r + i, [cell("A%d" % (r + i), text=t,
                                     s=S["SUB"] if t.startswith("■") else S["NOTE"])]))
    cols = [col_def(i + 1, i + 1, w) for i, (_c, _l, w, _s, _src) in enumerate(TOP_COLS)]
    return sheet_xml("A1:O%d" % (r + len(notes)), rows, cols, tab_selected=True)


# ---------------------------------------------------------------------------
# BB816使い方
# ---------------------------------------------------------------------------
GUIDE = [
    ("T", "BB816 ― ボリンジャーバンド判定を切り離して、スイング用に組み直したブック"),
    ("N", "300銘柄 × 250営業日（2025-07-29〜2026-08-06）の実データで検証した結果に基づいています。"),
    ("N", "指標の計算は V815 の BBスクリーニング と突き合わせ、300銘柄すべてで一致することを確認済みです。"),
    ("", ""),
    ("S", "■ 売買ルール（これだけ守れば動く）"),
    ("L", "入口      終値が +2σ を上抜け、かつ 出来高が直近5日平均の1.5倍以上、"),
    ("L", "          かつ MFI14 が 55以上、かつ RSI14 が 80以下"),
    ("L", "順位      BB816スコアの高い順。1日に建てるのは最大3銘柄"),
    ("L", "建て方    翌営業日の寄付（始値）で買う"),
    ("L", "枠        資金を8等分。1銘柄あたり12.5%。同じ銘柄を重ねて持たない"),
    ("L", "損切り    建値から -8%。買えたその日に逆指値を置く"),
    ("L", "手仕舞い  10営業日たったら損益に関わらず引けで売る"),
    ("L", "置かない  利確目標・トレーリングストップ・地合いフィルタ（いずれも検証で成績を下げた）"),
    ("", ""),
    ("S", "■ 検証結果（枠8・1日3銘柄・10日保有・損切-8%・250営業日）"),
    ("L", "BB816            総リターン +116.5%   最大ドローダウン -22.9%   Sharpe 2.29   取引248回   勝率46.0%"),
    ("L", "V815の現行スコア  総リターン  +48.8%   最大ドローダウン -19.6%   Sharpe 1.54"),
    ("L", "何も選ばず買う     総リターン  +57.5%   最大ドローダウン -28.9%   Sharpe 1.51"),
    ("L", "1トレードの最悪   -8.0%（損切りが効くため）。総資金への影響は -1.0%"),
    ("N", "開始日を10日ずつずらした13通りでも、中央値+95.7%・最小+64.0% と崩れませんでした。"),
    ("", ""),
    ("S", "■ V815 の BB判定から何を変えたか"),
    ("L", "1. 逆張り系を全部やめた"),
    ("N", "   スクイーズ★は超過リターン -0.48%、強気ダイバージェンスは -0.41%、%B<0.2 は -1.16%、"),
    ("N", "   RSI<40 は -1.21%。「安いところを買う」条件は、実測ではすべて市場平均に負けていました。"),
    ("L", "2. 順張り系に寄せた"),
    ("N", "   +2σ超え +1.11%、出来高2倍 +1.14%、RSI>70 +1.34%、MFI60-80 +0.96%、20日騰落率上位 +1.52%。"),
    ("N", "   いずれも前半125日・後半125日の両方でプラスでした。"),
    ("L", "3. スコアを作り直した"),
    ("N", "   V815のスコアは点数と成績が対応していません（40点以上でも超過はほぼゼロ）。"),
    ("N", "   BB816スコアは実測の超過リターンに比例した配点なので、点数が高いほど成績も上がります。"),
    ("N", "   0-19点 -0.79% / 20-39 +0.06% / 40-59 +0.76% / 60-69 +1.27% / 70-79 +1.65% / 80-100 +3.37%"),
    ("L", "4. 二番底・二番天井とスクイーズの計算を落とした"),
    ("N", "   入れても外しても成績が変わらなかった（+116.5% → +114.4%）ので、"),
    ("N", "   7MB分の重い計算をやめました。BB816が軽いのはこのためです。"),
    ("", ""),
    ("S", "■ 2人の分析官の意見と、最終判断"),
    ("L", "リスク分析官（損を絶対に出さない）"),
    ("N", "   ・裸の順張りは認めない。損切りなしは最悪 -47.6%、-10%以下の損失が12.2%発生している。"),
    ("N", "   ・損切りは -8%。MAE中央値が -5.02% なので、-4%や-6%は正常な値動きの内側で刈られる。"),
    ("N", "   ・10日で必ず切る。15日保有は実運用でSharpe 0.39まで落ちた。"),
    ("N", "   ・枠は8以上。枠3ではドローダウン -37.0%。集中はリスク管理ではなく賭けの増量。"),
    ("N", "   ・自分の直感が否定された点として「安いところを買うほうが安全」「地合いフィルタは守りになる」"),
    ("N", "     「厳選すれば安全」がすべて数字で否定されたことを認めている。"),
    ("L", "トレーダー分析官（積極的に利益を取る）"),
    ("N", "   ・入口は折衷案（+2σ・出来高1.5倍・MFI55）。守り型は質は高いが1日1.5件しか出ず枠が埋まらない。"),
    ("N", "   ・利確目標は絶対に置くな。+12%利確で期待値が +3.23% → +1.24% に落ちる。"),
    ("N", "   ・損切りは -10%まで緩めたい。-4%は期待値の49%を失う。"),
    ("N", "   ・保有は15日、SMA20割れまで引っ張りたい（1トレードの期待値は +8.73% と最大）。"),
    ("N", "   ・枠8-10へ増やすのは守りではなく攻めの手段。リターンとDDが同時に良くなる。"),
    ("L", "最終判断（採用）"),
    ("N", "   ・入口と順位付け、利確目標を置かないこと、枠8・1日3銘柄は両者が一致したので採用。"),
    ("N", "   ・損切りは -8%を採用。枠8で回した実測では -8% が総リターン+116.5%・最小値+64.0%と"),
    ("N", "     最も底が高く、-10%（+42.5%）にも損切りなし（+141.2%だが最小+12.7%）にも勝った。"),
    ("N", "     トレーダー案の-10%は、1回の運に左右されやすく採らない。"),
    ("N", "   ・保有は10日を採用。1トレードだけ見れば15日のほうが有利だが、枠が塞がって次を買えなくなる。"),
    ("N", "     枠8での実測は10日 +116.5% に対し15日 +86.2%。資金効率でリスク分析官の主張が正しかった。"),
    ("N", "   ・地合いフィルタは両者とも不要で一致。下落局面でむしろ成績が悪化したため入れない。"),
    ("", ""),
    ("S", "■ 弱点（承知した上で使うこと）"),
    ("N", "・検証期間の250日は地合いが強い期間でした（無選別でも10日で+1.80%）。弱気相場での検証はできていません。"),
    ("N", "・下落局面（65日）だけを取り出すと、この入口でも平均 -1.94% と負けます。損切り-8%とポジション量が唯一の防御です。"),
    ("N", "・2026年3月は15戦15敗、平均-6.00%。連敗する月は実際に存在します。"),
    ("N", "・8枠すべてが「+2σ超え・出来高増」という同じ性格の銘柄で埋まります。銘柄は分散しても値動きの理由は分散していません。"),
    ("N", "・投資助言ではありません。最終判断はご自身で行ってください。"),
    ("", ""),
    ("S", "■ 毎日の使い方"),
    ("L", "1. V816（本体ブック）で、いつも通り日次更新を実行して価格を取り込む"),
    ("L", "2. V816 で「BB816へ株価を送る」を実行する（Mod_BB切離し.bas の中にあります）"),
    ("L", "3. BB816.xlsx を開く。開いた瞬間に全再計算されます"),
    ("L", "4. 「BB816候補」シートの ◎買い を上から最大3件、翌営業日の寄付で買う"),
    ("N", "   ※ BB816 は価格シートを自分で持っているので、V816 が開いていなくても計算できます。"),
    ("N", "   ※ 楽天RSSアドインが入っていないパソコンでも開けます。"),
]


def build_guide(S):
    rows = []
    n = 1
    for kind, text in GUIDE:
        if kind == "":
            n += 1
            continue
        st = {"T": "TITLE", "S": "SECTION", "L": "BOLD", "N": "NOTE"}[kind]
        rows.append(row(n, [cell("A%d" % n, text=text, s=S[st])], ht="26" if kind == "T" else None))
        n += 1
    return sheet_xml("A1:A%d" % n, rows, [col_def(1, 1, 118)])


# ---------------------------------------------------------------------------
def build():
    zin = zipfile.ZipFile(SRC)
    styles, S = extend_styles(zin.read("xl/styles.xml").decode("utf-8"))

    sheets = [("BB816使い方", build_guide(S), False),
              ("BB816候補", build_top(S), False),
              ("BB816判定", build_judge(S), False),
              ("BB816設定", build_settings(S), False)]
    for name, src in PRICE:
        xml = flatten(zin.read("xl/worksheets/%s.xml" % src).decode("utf-8"))
        sheets.append((name, xml.encode("utf-8"), False))
        log.append("  価格シート %-6s <- %s（数式を値に固定）" % (name, src))

    parts = {"xl/styles.xml": styles.encode("utf-8"),
             "xl/sharedStrings.xml": zin.read("xl/sharedStrings.xml"),
             "xl/theme/theme1.xml": zin.read("xl/theme/theme1.xml")}
    for i, (name, data, _h) in enumerate(sheets, start=1):
        parts["xl/worksheets/sheet%d.xml" % i] = data

    n = len(sheets)
    parts["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<workbookPr/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800"'
        ' windowHeight="16000" tabRatio="820" activeTab="1"/></bookViews><sheets>%s</sheets>'
        '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
        % "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (nm, i, i)
                  for i, (nm, _d, _h) in enumerate(sheets, start=1))).encode("utf-8")

    rels = "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i) for i in range(1, n + 1))
    rels += ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
             '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
             % (n + 1, n + 2, n + 3))
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
        % rels).encode("utf-8")

    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '%s'
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>' % "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType='
                             '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i
                             for i in range(1, n + 1))).encode("utf-8")

    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>').encode("utf-8")

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    parts["docProps/core.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>BB816 ボリンジャーバンド判定（独立版）</dc:title>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>' % (now, now)).encode("utf-8")
    parts["docProps/app.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
        '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>'
        '<AppVersion>16.0300</AppVersion></Properties>').encode("utf-8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        zout.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for k in sorted(parts):
            zout.writestr(k, parts[k])
    zin.close()
    return [nm for nm, _d, _h in sheets]


def verify(names):
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(DST)
    assert z.testzip() is None
    for n in z.namelist():
        if n.endswith(".xml") or n.endswith(".rels"):
            ET.fromstring(z.read(n))

    ALLOWED = {"AVERAGE", "STDEVP", "SUMPRODUCT", "IF", "IFERROR", "AND", "OR", "MAX", "MIN",
               "ROUND", "INDEX", "MATCH", "LARGE", "ROW", "HYPERLINK", "COUNT", "COUNTIF"}
    ref = re.compile(r"(?<![A-Za-z0-9_.!])'?([^'!()\[\],:;+\-*/<>=&\"]+?)'?!\$?[A-Z]{1,3}\$?\d+")
    used, outside, total = set(), set(), 0
    n_sheets = len(names)
    for i in range(1, n_sheets + 1):
        xml = z.read("xl/worksheets/sheet%d.xml" % i).decode("utf-8")
        assert "r:id=" not in xml
        for f in re.findall(r"<f[^>]*>([^<]*)</f>", xml):
            total += 1
            plain = f.replace("&gt;", ">").replace("&lt;", "<").replace("&quot;", '"')
            used |= set(re.findall(r"([A-Z][A-Z0-9\.]{1,20})\(", plain))
            for m in ref.finditer(plain):
                outside.add(m.group(1))
    bad = used - ALLOWED
    assert not bad, "V815 に無い関数を使っている: %s" % sorted(bad)
    unknown = outside - set(names)
    assert not unknown, "無いシートを参照している: %s" % sorted(unknown)
    price_f = 0
    for i, nm in enumerate(names, start=1):
        if nm in [p[0] for p in PRICE]:
            price_f += z.read("xl/worksheets/sheet%d.xml" % i).decode("utf-8").count("<f")
    assert price_f == 0, "価格シートに数式が残っている"

    out = ["", "  自己点検",
           "    XML パース      : %d パート OK" % len(z.namelist()),
           "    シート          : %s" % " / ".join(names),
           "    数式の本数      : %s 本" % f"{total:,}",
           "    使っている関数  : %s" % " ".join(sorted(used)),
           "    参照シート      : %s" % " ".join(sorted(outside)),
           "    価格シートの数式: 0（値に固定済み・RSS不要）",
           "    ファイルサイズ  : %.2f MB" % (DST.stat().st_size / 1048576)]
    z.close()
    return out


if __name__ == "__main__":
    print("BB816 ビルド")
    print("  入力: %s" % SRC)
    print("  出力: %s" % DST)
    names = build()
    log.extend(verify(names))
    print("\n".join(log))
    print("\n完了")
