# -*- coding: utf-8 -*-
"""生成したワークシートXMLを、ECMA-376 の CT_Worksheet 要素順と
   Excel の実ファイルの慣行に照らして機械検証する。"""
import re, sys, zipfile
from xml.etree import ElementTree as ET

SEQ = ['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr',
       'sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate',
       'customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations',
       'hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks',
       'customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing',
       'legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'


def col_idx(ref):
    m = re.match(r'([A-Z]+)', ref); n = 0
    for ch in m.group(1): n = n * 26 + ord(ch) - 64
    return n


def check(xml, label):
    errs = []
    try:
        root = ET.fromstring(xml)
    except Exception as e:
        return ['整形式エラー: %s' % e]

    # 1) 宣言されていない名前空間接頭辞を Requires が参照していないか
    st = xml.index('<worksheet')
    decl = set(re.findall(r'xmlns:([A-Za-z0-9_]+)=', xml[st:xml.index('>', st)]))
    for req in re.findall(r'Requires="([^"]+)"', xml):
        for pfx in req.split():
            if pfx not in decl:
                errs.append('★ Requires="%s" の接頭辞が未宣言（Excelがシートを破棄します）' % pfx)

    # 2) worksheet 直下の要素順
    seen = []
    for ch in root:
        tag = ch.tag.replace(NS, '')
        if ch.tag.startswith(MC):
            tag = 'controls'          # mc:AlternateContent で包まれた controls
        if tag not in SEQ:
            errs.append('未知の子要素: %s' % tag); continue
        seen.append((SEQ.index(tag), tag))
    for i in range(1, len(seen)):
        if seen[i][0] < seen[i - 1][0]:
            errs.append('要素順エラー: <%s> が <%s> の後にある' % (seen[i][1], seen[i - 1][1]))

    # 3) 行番号の昇順、行内のセル列の昇順、r属性と行番号の一致
    sd = root.find(NS + 'sheetData')
    prev_r = 0
    for row in (sd if sd is not None else []):
        r = int(row.get('r'))
        if r <= prev_r: errs.append('行順エラー: r=%d' % r)
        prev_r = r
        prev_c = 0
        for c in row:
            ref = c.get('r')
            if not ref.endswith(str(r)):
                errs.append('セル参照と行番号が不一致: %s (row %d)' % (ref, r))
            ci = col_idx(ref)
            if ci <= prev_c: errs.append('列順エラー: %s' % ref)
            prev_c = ci
            # <c> の子要素順は f → v （または is 単独）
            kids = [k.tag.replace(NS, '') for k in c]
            if kids and kids != sorted(kids, key=lambda t: {'f':0,'v':1,'is':2}.get(t, 9)):
                errs.append('セル内要素順エラー: %s %s' % (ref, kids))
            if 'is' in kids and c.get('t') != 'inlineStr':
                errs.append('<is> があるのに t="inlineStr" でない: %s' % ref)

    # 4) control の r:id が rels に存在するか（呼び出し側で渡す）
    return errs


if __name__ == '__main__':
    z = zipfile.ZipFile(sys.argv[1])
    bad = 0
    for n in z.namelist():
        if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
            e = check(z.read(n).decode('utf-8'), n)
            if e:
                bad += 1
                print('■', n)
                for x in e[:10]: print('   ', x)
    print('検証したシート: %d 枚 / 問題あり: %d 枚'
          % (sum(1 for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)), bad))
