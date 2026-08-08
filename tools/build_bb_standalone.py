#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V815 → ボリンジャーバンド判定 独立ブック ビルダー

■ 何をするスクリプトか

  V815.xlsm に同居している「ボリンジャーバンド判定」を丸ごと外へ出して、
  本体ブック（V815）が無くても単独で動く計算ブックを作る。

      出力: V815_BB独立版.xlsx

  独立版に入るシート（10枚）

      BB使い方          使い方・戦略サマリー
      BB設定            パラメータ（σ倍率・スクイーズ許容係数など）
      BB買い候補TOP     スコア上位30銘柄
      BBスクリーニング   全銘柄400行の判定本体
      終値 / 高値 / 安値 / 出来高   判定に必要な価格データ（250営業日）
      BB_帯幅           スクイーズ判定用の帯幅履歴（非表示）
      BB_スイング       二番底/天井の厳密スイング検出（非表示）

■ 「独立」にするために外した依存

  1. 楽天RSS
     価格シートのC列は _xll.RssMarket / _xll.RssIndexMarket。
     RSSアドインが無い環境で開くと #NAME? になるので、価格シートの数式は
     すべてキャッシュ値に固定（フラット化）する。
     ※ BB判定が見ているのは E列以降の確定履歴なので、判定結果は変わらない。

  2. 銘柄管理シート
     価格シートの A5/B5（TOPX行）などが 銘柄管理! を参照している。
     これも上のフラット化で値に落ちるため、銘柄管理シートを持たずに済む。

  3. VBA
     独立版はマクロを持たない（BB判定は全て数式）。
     vbaProject.bin を移植すると、存在しないシートに紐づく Document モジュールが
     残って Excel の修復対象になるため、あえて入れていない。
     データ更新は本体側の Mod_BB切離し.bas（BB独立版へ株価を送る）から行う。

  4. ボタン（BB買い候補TOP の隠しボタン "BB更新"）
     マクロ [0]!BB更新 を呼ぶ図形なので、独立版からは drawing/vml/control ごと外す。
     workbook.xml に fullCalcOnLoad="1" を入れてあるので、開いた時点で全再計算される。

■ 壊さないための方針（V807/V808 で開けなくなった件の教訓）

  ・既存シートの XML に「セルを挿入」しない。やるのは
      (a) 数式の削除（価格シートのフラット化）
      (b) rel を持つ要素の削除（drawing/legacyDrawing/controls/pageSetup r:id）
    だけ。
  ・共有数式(shared formula)のマスターは触らない。フラット化する価格シートは
    <f> を全部消すのでマスター/追従の関係ごと消える（追従セルだけ残らない）。
  ・styles.xml と sharedStrings.xml は丸ごとコピー。s= と t="s" のインデックスが
    ずれないので、書式も文字列もそのまま生きる。
  ・calcChain.xml は入れない（Excel が開いたときに作り直す）。

使い方:  python3 tools/build_bb_standalone.py V815.xlsm V815_BB独立版.xlsx
"""
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "V815_BB独立版.xlsx")

# (シート名, V815内のワークシートXML, 価格シートか, 非表示か)
SHEETS = [
    ("BB使い方",       "sheet22", False, False),
    ("BB設定",         "sheet23", False, False),
    ("BB買い候補TOP",  "sheet25", False, False),
    ("BBスクリーニング", "sheet24", False, False),
    ("終値",           "sheet14", True,  False),
    ("高値",           "sheet12", True,  False),
    ("安値",           "sheet13", True,  False),
    ("出来高",         "sheet15", True,  False),
    ("BB_帯幅",        "sheet27", False, True),
    ("BB_スイング",     "sheet28", False, True),
]

COPY_AS_IS = ["xl/styles.xml", "xl/sharedStrings.xml", "xl/theme/theme1.xml"]

log = []


def strip_rel_parts(xml: str) -> str:
    """rel（別パーツ）を指す要素を落とす。独立版には図形もプリンタ設定も持ち込まない。"""
    # パスワード付きのシート保護も外す。独立版に株価を書き込めなくなるため。
    xml = re.sub(r"<sheetProtection[^>]*/>", "", xml)
    xml = re.sub(r'<drawing\s+r:id="[^"]*"\s*/>', "", xml)
    xml = re.sub(r'<legacyDrawing\s+r:id="[^"]*"\s*/>', "", xml)
    xml = re.sub(r'<picture\s+r:id="[^"]*"\s*/>', "", xml)
    xml = re.sub(r"<tableParts.*?</tableParts>", "", xml, flags=re.S)
    xml = re.sub(r"<oleObjects.*?</oleObjects>", "", xml, flags=re.S)
    xml = re.sub(r"<hyperlinks>.*?</hyperlinks>", "", xml, flags=re.S)
    # <controls> は mc:AlternateContent に包まれている。包みごと外す。
    xml = re.sub(
        r"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*?<controls>.*?</mc:AlternateContent>\s*</mc:Choice>\s*</mc:AlternateContent>",
        "", xml, flags=re.S)
    xml = re.sub(r"<controls>.*?</controls>", "", xml, flags=re.S)
    # pageSetup の r:id（プリンタ設定）だけ属性を外す。要素自体は残す。
    xml = re.sub(r'(<pageSetup\b[^>]*?)\s+r:id="[^"]*"', r"\1", xml)
    return xml


def flatten_formulas(xml: str) -> str:
    """価格シートの数式を全部消して、キャッシュ値だけのデータシートにする。"""
    xml = re.sub(r"<f[^>]*/>", "", xml)          # 共有数式の追従セル
    xml = re.sub(r"<f[^>]*>[^<]*</f>", "", xml)  # 通常/共有マスター/配列数式
    # 数式の文字列結果 t="str" は、数式が無くなると不正になるのでインライン文字列に直す
    def to_inline(m):
        attrs = m.group(1) + m.group(2)
        text = m.group(3)
        if not text:
            return "<c%s/>" % attrs
        return '<c%s t="inlineStr"><is><t>%s</t></is></c>' % (attrs, text)
    xml = re.sub(r'<c([^>]*?)\st="str"([^>]*?)>(?:<v>(.*?)</v>)?</c>', to_inline, xml)
    # 数式のエラー結果 t="e" は空セルにする（RSS未接続時の #N/A などを持ち込まない）
    xml = re.sub(r'<c([^>]*?)\st="e"([^>]*?)>(?:<v>.*?</v>)?</c>', r"<c\1\2/>", xml)
    return xml


def normalize_view(xml: str, first: bool) -> str:
    """タブ選択状態を正規化（tabSelected が複数あると Excel が嫌がる）。"""
    xml = re.sub(r'\s+tabSelected="1"', "", xml)
    if first:
        xml = xml.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
    return xml


def clean(xml: str, is_price: bool, first: bool) -> str:
    xml = strip_rel_parts(xml)
    if is_price:
        before = xml.count("<f")
        xml = flatten_formulas(xml)
        log.append("    数式フラット化: %d 個" % before)
    xml = re.sub(r'\s+codeName="[^"]*"', "", xml, count=1)  # sheetPr の codeName（VBA無しなので不要）
    xml = normalize_view(xml, first)
    return xml


def build():
    zin = zipfile.ZipFile(SRC)
    have = set(zin.namelist())

    parts = {}

    # ---- ワークシート ----
    for i, (name, src, is_price, hidden) in enumerate(SHEETS, start=1):
        path = "xl/worksheets/%s.xml" % src
        assert path in have, "%s が V815 に無い" % path
        log.append("  %-14s <- %s%s" % (name, src, "  (非表示)" if hidden else ""))
        xml = zin.read(path).decode("utf-8")
        parts["xl/worksheets/sheet%d.xml" % i] = clean(xml, is_price, first=(i == 1)).encode("utf-8")

    # ---- そのままコピーするパーツ ----
    for p in COPY_AS_IS:
        assert p in have, "%s が無い" % p
        parts[p] = zin.read(p)

    # ---- workbook.xml ----
    sheets_xml = "".join(
        '<sheet name="%s" sheetId="%d"%s r:id="rId%d"/>'
        % (name, i, ' state="hidden"' if hidden else "", i)
        for i, (name, _s, _p, hidden) in enumerate(SHEETS, start=1))
    parts["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<workbookPr/>'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16000"'
        ' tabRatio="820" activeTab="0"/></bookViews>'
        '<sheets>%s</sheets>'
        # fullCalcOnLoad: 開いた瞬間に全再計算 → 価格を差し替えれば判定も必ず更新される
        '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
        '</workbook>' % sheets_xml).encode("utf-8")

    rels = "".join(
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
        % (i, i) for i in range(1, len(SHEETS) + 1))
    n = len(SHEETS)
    rels += (
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        % (n + 1, n + 2, n + 3))
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
        % rels).encode("utf-8")

    # ---- [Content_Types] ----
    overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        % i for i in range(1, len(SHEETS) + 1))
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
        '</Types>' % overrides).encode("utf-8")

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
        '<dc:title>V815 ボリンジャーバンド判定 独立版</dc:title>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>' % (now, now)).encode("utf-8")

    parts["docProps/app.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
        '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>'
        '<AppVersion>16.0300</AppVersion></Properties>').encode("utf-8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        # [Content_Types] を先頭に置く（Excel はこれを最初に読む）
        zout.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name in sorted(parts):
            zout.writestr(name, parts[name])

    zin.close()


def verify():
    """出来上がりを自己点検する。開けないブックを渡さないための最低限のガード。"""
    import xml.etree.ElementTree as ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    z = zipfile.ZipFile(DST)

    bad = z.testzip()
    assert bad is None, "壊れたエントリ: %s" % bad

    names = set(z.namelist())
    for n in names:
        if n.endswith(".xml") or n.endswith(".rels"):
            ET.fromstring(z.read(n))  # XML として読めるか

    sheet_names = {s[0] for s in SHEETS}
    ref = re.compile(r"(?<![A-Za-z0-9_.!])'?([^'!()\[\],:;+\-*/<>=&\"]+?)'?!\$?[A-Z]{1,3}\$?\d+")
    outside, rss, total_f = set(), 0, 0
    for i, (name, _s, is_price, _h) in enumerate(SHEETS, start=1):
        xml = z.read("xl/worksheets/sheet%d.xml" % i).decode("utf-8")
        assert "r:id=" not in xml, "%s に rel 参照が残っている" % name
        fs = re.findall(r"<f[^>]*>([^<]*)</f>", xml)
        total_f += len(fs)
        if is_price:
            assert not fs, "%s に数式が残っている" % name
            assert 't="str"' not in xml, '%s に数式なしの t="str" が残っている' % name
        for f in fs:
            rss += f.count("_xll.")
            for m in ref.finditer(f.replace("&gt;", ">").replace("&lt;", "<")):
                outside.add(m.group(1))

    unknown = {s for s in outside if s not in sheet_names}
    assert not unknown, "独立版に無いシートを参照している: %s" % sorted(unknown)
    assert rss == 0, "RSS(_xll) 数式が %d 個残っている" % rss

    log.append("")
    log.append("  自己点検")
    log.append("    XML パース          : %d パート OK" % len(names))
    log.append("    残った数式          : %d 本（すべて独立版内のシート参照）" % total_f)
    log.append("    参照シート          : %s" % " ".join(sorted(outside)))
    log.append("    _xll.Rss 数式       : 0")
    log.append("    出力サイズ          : %.1f MB" % (DST.stat().st_size / 1024 / 1024))
    z.close()


if __name__ == "__main__":
    print("V815 → BB独立版 ビルド")
    print("  入力: %s" % SRC)
    print("  出力: %s" % DST)
    build()
    verify()
    print("\n".join(log))
    print("\n完了")
