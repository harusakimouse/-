#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
切り分け用ビルド A：共有数式の展開と幾何情報の整合だけを行う。
数式の書き換え・列の追加は一切しない。

これが Excel で正常に開けば「展開処理は正しく、原因は私の数式編集側」と分かる。
開けなければ「展開処理そのものが原因」と分かる。
"""
import re, sys, zipfile
from pathlib import Path
from shared_expand import expand_file_sheets
from build_v808 import fix_geometry, log

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "v805.xlsm")
DST = Path(sys.argv[2] if len(sys.argv) > 2 else "V808_検証A_展開のみ.xlsm")
SHEETS = {"xl/worksheets/sheet1.xml": "管理", "xl/worksheets/sheet3.xml": "厳選TOP2",
          "xl/worksheets/sheet7.xml": "分析", "xl/worksheets/sheet8.xml": "売分析"}

zin = zipfile.ZipFile(SRC)
items = {n: zin.read(n) for n in zin.namelist()}
order = zin.namelist()
zin.close()

print("共有数式の展開のみ:")
for line in expand_file_sheets(SRC, items, SHEETS):
    print(line)
for path, name in SHEETS.items():
    items[path] = fix_geometry(items[path].decode(), name).encode()
print("\n".join(log))

if "xl/calcChain.xml" in items:
    items["[Content_Types].xml"] = re.sub(
        r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "",
        items["[Content_Types].xml"].decode()).encode()
    items["xl/_rels/workbook.xml.rels"] = re.sub(
        r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "",
        items["xl/_rels/workbook.xml.rels"].decode()).encode()
w = items["xl/workbook.xml"].decode()
w = (re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="0" fullCalcOnLoad="1"/>', w)
     if "<calcPr" in w else
     w.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>'))
items["xl/workbook.xml"] = w.encode()

with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zo:
    for n in order:
        if n == "xl/calcChain.xml":
            continue
        zo.writestr(n, items[n])
print(f"\n出力: {DST} ({DST.stat().st_size:,} bytes)")
