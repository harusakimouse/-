#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V815 → V816 ビルダー（ボリンジャーバンド判定の切り離し）

■ V816 とは

  V815 から「ボリンジャーバンド判定」を降ろした本体ブック。
  V815 の売買判定（買抽出v13 / 売抽出v13 / 厳選TOP2 / 分析 / 管理）は
  BBシートを一切参照していないことを確認済みなので、外しても判定は変わらない。

      確認方法: 全シートのXMLを走査し「BB設定!」「BBスクリーニング!」等の
                参照を数えたところ、BB系シートを参照しているのはBB系シートだけだった。

  外して得られるもの
      ・ブックから約51,000本の数式が消える（F9・再計算が軽くなる）
      ・ファイルサイズが小さくなる
      ・BB判定は BB816.xlsx（独立ブック）に移り、単独で動く

■ なぜ「シートを物理削除」しないのか

  このブックの VBA プロジェクト(vbaProject.bin)には、BB系6シートに対応する
  Document モジュール（Sheet20〜Sheet23, Sheet25, Sheet26）が入っている。
  ZIP/XML でシートだけ消すと、対応するシートが無い Document モジュールが残り、
  Excel が「修復」に入って VBA プロジェクトごと落とす危険がある。
  V807/V808 で2回ブックが開けなくなった経緯を踏まえ、その賭けはしない。

  代わりに
      ・BB系6シートは「中身を空にして非表示」にする（数式ゼロ・データゼロ）
      ・シート自体は残るので codeName の対応が崩れず、VBA は無傷
      ・タブからは見えない
      ・完全に消したい場合は Mod_BB切離し.bas の BB_シートを完全削除 を
        Excel 上で1回実行する（Excel自身が削除するので安全）

■ 触るもの / 触らないもの

  触る    BB系6シートのXML（最小の説明シートに置き換え）
          workbook.xml の該当 <sheet> に state="hidden" を追加
          BB買い候補TOP にぶら下がっていた図形・ボタン（drawing/vml/ctrlProp）を削除
          calcChain.xml を削除（Excelが開いたときに作り直す。数式構成が変わるため）
  触らない 他の24シート、vbaProject.bin、styles.xml、sharedStrings.xml、
          volatileDependencies.xml、metadata.xml、定義済み名（すべてバイト単位でそのまま）

使い方:  python3 tools/build_v816.py V815.xlsm V816.xlsm
"""
import re
import sys
import zipfile
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "V816.xlsm")

# シート名 → (ワークシートXML, codeName)
BB_SHEETS = {
    "BB使い方":       ("xl/worksheets/sheet22.xml", "Sheet21"),
    "BB設定":         ("xl/worksheets/sheet23.xml", "Sheet22"),
    "BBスクリーニング": ("xl/worksheets/sheet24.xml", "Sheet23"),
    "BB買い候補TOP":  ("xl/worksheets/sheet25.xml", "Sheet20"),
    "BB_帯幅":        ("xl/worksheets/sheet27.xml", "Sheet25"),
    "BB_スイング":     ("xl/worksheets/sheet28.xml", "Sheet26"),
}
# BB買い候補TOP にぶら下がっていた図形類（隠しボタン "BB更新"）
DROP_PARTS = {
    "xl/worksheets/_rels/sheet25.xml.rels",
    "xl/drawings/drawing7.xml",
    "xl/drawings/vmlDrawing10.vml",
    "xl/ctrlProps/ctrlProp17.xml",
    "xl/calcChain.xml",
}

NOTE = [
    "このシートの中身は V816 で切り離しました。",
    "ボリンジャーバンド判定は独立ブック BB816.xlsx に移してあります。",
    "V816 本体の売買判定（買抽出v13 / 売抽出v13 / 厳選TOP2 / 分析 / 管理）は",
    "もともと BB を参照していないため、外しても判定結果は変わりません。",
    "",
    "このシートを完全に消したい場合は、Mod_BB切離し.bas を読み込んで",
    "「BB_シートを完全削除」を1回実行してください（Excel自身が安全に削除します）。",
]


def shell_sheet(code_name, title):
    rows = []
    rows.append('<row r="1" ht="24" customHeight="1"><c r="A1" t="inlineStr"><is><t>%s</t></is></c></row>'
                % esc("【%s】切り離し済み" % title))
    for i, text in enumerate(NOTE, start=3):
        if not text:
            continue
        rows.append('<row r="%d"><c r="A%d" t="inlineStr"><is><t>%s</t></is></c></row>'
                    % (i, i, esc(text)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetPr codeName="%s"/>'
            '<dimension ref="A1:A%d"/>'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="18"/>'
            '<cols><col min="1" max="1" width="90" customWidth="1"/></cols>'
            '<sheetData>%s</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '</worksheet>' % (code_name, len(NOTE) + 2, "".join(rows))).encode("utf-8")


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def hide_sheets(workbook_xml, names):
    """workbook.xml の該当シートに state="hidden" を付ける。"""
    out = workbook_xml
    for name in names:
        m = re.search(r'<sheet name="%s"[^>]*/>' % re.escape(name), out)
        if not m:
            raise SystemExit("workbook.xml に %s が見つからない" % name)
        tag = m.group(0)
        if 'state="' in tag:
            continue
        new = tag.replace(' r:id=', ' state="hidden" r:id=')
        out = out.replace(tag, new)
    return out


def strip_content_types(ct, dropped):
    for part in dropped:
        ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(part), "", ct)
    return ct


def strip_workbook_rels(rels):
    """calcChain への参照を外す。"""
    return re.sub(r'<Relationship[^>]*calcChain\.xml"[^>]*/>', "", rels)


def build():
    zin = zipfile.ZipFile(SRC)
    names = zin.namelist()
    xml_by_sheet = {v[0]: (k, v[1]) for k, v in BB_SHEETS.items()}

    log = []
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        for item in zin.infolist():
            n = item.filename
            if n in DROP_PARTS:
                log.append("  削除    %s" % n)
                continue
            data = zin.read(n)
            if n in xml_by_sheet:
                title, code = xml_by_sheet[n]
                before = len(data)
                data = shell_sheet(code, title)
                log.append("  空にする %-28s %8.1fKB → %.1fKB" % (title, before / 1024, len(data) / 1024))
            elif n == "xl/workbook.xml":
                s = data.decode("utf-8")
                s = hide_sheets(s, BB_SHEETS.keys())
                data = s.encode("utf-8")
                log.append("  非表示化 workbook.xml の BB系6シート")
            elif n == "[Content_Types].xml":
                data = strip_content_types(data.decode("utf-8"), DROP_PARTS).encode("utf-8")
            elif n == "xl/_rels/workbook.xml.rels":
                data = strip_workbook_rels(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    zin.close()
    return log


def verify():
    """V816 の自己点検。"""
    import xml.etree.ElementTree as ET
    a, b = zipfile.ZipFile(SRC), zipfile.ZipFile(DST)
    assert b.testzip() is None

    for n in b.namelist():
        if n.endswith(".xml") or n.endswith(".rels"):
            ET.fromstring(b.read(n))

    # VBA は1バイトも変わっていないこと
    assert a.read("xl/vbaProject.bin") == b.read("xl/vbaProject.bin"), "vbaProject.bin が変わっている"
    # 価格シート・判定シートは完全に同一であること
    same = 0
    for n in a.namelist():
        if n in DROP_PARTS or n in {v[0] for v in BB_SHEETS.values()}:
            continue
        if n in ("xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
            continue
        assert n in b.namelist(), "%s が欠けている" % n
        assert a.read(n) == b.read(n), "%s が書き換わっている" % n
        same += 1

    wb = b.read("xl/workbook.xml").decode()
    for name in BB_SHEETS:
        m = re.search(r'<sheet name="%s"[^>]*/>' % re.escape(name), wb)
        assert m and 'state="hidden"' in m.group(0), "%s が非表示になっていない" % name

    # 残った数式の数を数える（軽くなったことの確認）
    def count_formulas(z):
        n = 0
        for name in z.namelist():
            if name.startswith("xl/worksheets/sheet"):
                n += z.read(name).decode("utf-8", "ignore").count("<f")
        return n

    fa, fb = count_formulas(a), count_formulas(b)
    out = [
        "",
        "  自己点検",
        "    XML パース            : 全パート OK",
        "    vbaProject.bin       : V815 と完全一致（マクロは無傷）",
        "    そのまま引き継いだパート : %d 個（価格シート・判定シートを含む）" % same,
        "    BB系6シート           : 空・非表示",
        "    数式の本数            : %s → %s（%s本 減）" % (f"{fa:,}", f"{fb:,}", f"{fa - fb:,}"),
        "    ファイルサイズ         : %.1fMB → %.1fMB" % (SRC.stat().st_size / 1048576, DST.stat().st_size / 1048576),
    ]
    a.close(); b.close()
    return out


if __name__ == "__main__":
    print("V815 → V816 ビルド（BB判定の切り離し）")
    print("  入力: %s" % SRC)
    print("  出力: %s" % DST)
    log = build()
    log += verify()
    print("\n".join(log))
    print("\n完了")
