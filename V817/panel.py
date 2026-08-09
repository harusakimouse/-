# -*- coding: utf-8 -*-
"""「発注リスト」シート（旧 Sheet1）を生成する。
   ・保有中のポジションだけを詰めて並べ、MS2に出す注文をそのまま読める形にする
   ・既存マクロを呼ぶボタンを5つ設置する（VBAの追加は不要）"""

from xmlcell import esc

BUTTONS = [
    ('損切りアラート',        '[0]!損切りアラート'),
    ('地合いチェック',        '[0]!地合いチェック'),
    ('勝率レポート',          '[0]!トレード勝率レポート'),
    ('データ更新（本日分）',  '[0]!日次更新_本日分'),
    ('RSS接続チェック',       '[0]!RSS接続状態チェック'),
]
FIRST_ROW = 7
LAST_ROW = 106
COLW = [('A', 5), ('B', 9), ('C', 22), ('D', 6), ('E', 9), ('F', 10),
        ('G', 13), ('H', 13), ('I', 12), ('J', 13), ('K', 34), ('L', 13)]
HEAD = ['No', 'コード', '銘柄名', '売買', '株数', '建値',
        '損切（逆指値）', '利確（指値）', '期日', 'トレーリング', '今日やること', '状況']


def _t(ref, text):
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, esc(text))


_MISS = object()


STYLE = {'B': 517, 'E': 481, 'F': 482, 'G': 482, 'H': 482, 'I': 584, 'J': 482}


def _f(ref, formula, cached=_MISS, style=None):
    st = ' s="%d"' % style if style else ''
    if cached is _MISS:
        return '<c r="%s"%s><f>%s</f></c>' % (ref, st, esc(formula))
    if cached is None or cached == '':
        return '<c r="%s"%s t="str"><f>%s</f><v/></c>' % (ref, st, esc(formula))
    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
        v = int(cached) if float(cached).is_integer() else cached
        return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, st, esc(formula), v)
    return '<c r="%s"%s t="str"><f>%s</f><v>%s</v></c>' % (ref, st, esc(formula), esc(str(cached)))


def build_sheet_xml(V=None):
    V = V or {}
    _todo = sum(1 for d in V.values() if str(d.get('todo', '')).startswith('▶'))
    _up = sum(1 for d in V.values() if str(d.get('todo', '')).startswith('⬆'))
    _risk = sum(abs((d.get('F') or 0) - (d.get('N') or 0)) * (d.get('qty') or 0)
                for d in V.values() if isinstance(d.get('F'), (int, float))
                and isinstance(d.get('N'), (int, float)))
    IDX = 'INDEX(管理!${c}:${c},$N{r})'
    rows = []

    rows.append('<row r="1" ht="24" customHeight="1">' +
                _t('A1', '発注リスト ── 保有中のポジションと、いま証券会社に出しておく注文') + '</row>')
    rows.append('<row r="2" ht="34" customHeight="1"/>')
    rows.append('<row r="4" ht="20" customHeight="1">' +
                _t('A4', '要対応') +
                _f('C4', 'COUNTIF($K$%d:$K$%d,"▶*")&"件"' % (FIRST_ROW, LAST_ROW),
                   cached='%d件' % _todo) +
                _t('D4', '逆指値の引き上げ') +
                _f('F4', 'COUNTIF($K$%d:$K$%d,"⬆*")&"件"' % (FIRST_ROW, LAST_ROW),
                   cached='%d件' % _up) +
                _t('G4', '保有') +
                _f('H4', 'COUNT($B$%d:$B$%d)&"件"' % (FIRST_ROW, LAST_ROW),
                   cached='%d件' % len(V)) +
                _t('I4', '想定最大損失') +
                _f('J4', 'SUMPRODUCT(($B$%d:$B$%d<>"")*ABS($F$%d:$F$%d-$G$%d:$G$%d)*$E$%d:$E$%d)'
                   % (FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW),
                   cached=round(_risk)) +
                '</row>')
    rows.append('<row r="5" ht="18" customHeight="1">' +
                _t('A5', '※ 損切・利確ラインは購入日のATRで固定されているため、建玉後に動きません。'
                         '建玉時に1回ずつ出せば、あとは「今日やること」列だけ見てください。') + '</row>')
    rows.append('<row r="6" ht="30" customHeight="1">' +
                ''.join(_t('%s6' % c, h) for (c, _), h in zip(COLW, HEAD)) + '</row>')

    for r in range(FIRST_ROW, LAST_ROW + 1):
        n = r - (FIRST_ROW - 1)
        cells = []
        d = V.get(n, {})
        cells.append(_f('A%d' % r, 'IF($N{r}="","",{n})'.format(r=r, n=n),
                        cached=(n if d else '')))
        for col, src in (('B', 'B'), ('C', 'D'), ('D', 'C'), ('F', 'F'),
                         ('G', 'N'), ('H', 'M'), ('I', 'AE'), ('J', 'R'), ('L', 'O')):
            cells.append(_f('%s%d' % (col, r),
                            'IF($N{r}="","",{ix})'.format(r=r, ix=IDX.format(c=src, r=r)),
                            cached=d.get(src, ''), style=STYLE.get(col)))
        # 株数：実績があればそれ、無ければ推奨株数
        cells.append(_f('E%d' % r,
                        'IF($N{r}="","",IF(N({g})>0,{g},{y}))'.format(
                            r=r, g=IDX.format(c='G', r=r), y=IDX.format(c='Y', r=r)),
                        cached=d.get('qty', ''), style=STYLE['E']))
        # 今日やること
        o = IDX.format(c='O', r=r)
        rr = IDX.format(c='R', r=r)
        nn = IDX.format(c='N', r=r)
        cc = IDX.format(c='C', r=r)
        cells.append(_f('K%d' % r,
            'IF($N{r}="","",'
            'IF({o}="🟢利確","▶ 利確：成行で決済",'
            'IF({o}="🔴損切","▶ 損切：成行で決済",'
            'IF({o}="🟠TRAIL","▶ トレーリング到達：成行で決済",'
            'IF({o}="⏱時間決済","▶ 期日到来：成行で決済",'
            'IF({fv}="","（建値の確定待ち：まだ注文を出さない）",'
            'IF(AND({rr}<>"",IF({cc}="売",{rr}<{nn},{rr}>{nn})),'
            '"⬆ 逆指値を "&TEXT({rr},"#,##0")&" に変更","そのまま")))))))'
            .format(r=r, o=o, rr=rr, nn=nn, cc=cc, fv=IDX.format(c='F', r=r)),
            cached=d.get('todo', '')))
        # 参照行（補助・非表示列）
        cells.append(_f('N%d' % r,
                        'IFERROR(MATCH({n},管理!$AB$4:$AB$461,0)+3,"")'.format(n=n),
                        cached=d.get('row', '')))
        rows.append('<row r="%d">%s</row>' % (r, ''.join(sorted(
            cells, key=lambda c: (len(c.split('"')[1]) - len(str(r)), c.split('"')[1])))))

    cols = ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>'
                   % (i + 1, i + 1, w) for i, (_, w) in enumerate(COLW))
    cols += '<col min="14" max="14" width="8" hidden="1" customWidth="1"/>'

    ctrls = []
    for i, (label, macro) in enumerate(BUTTONS):
        c0 = i * 2
        ctrls.append(
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<mc:Choice Requires="x14">'
            '<control shapeId="%d" r:id="rId%d" name="Button %d">'
            '<controlPr defaultSize="0" print="0" autoFill="0" autoPict="0" macro="%s">'
            '<anchor moveWithCells="1" sizeWithCells="1">'
            '<from><xdr:col>%d</xdr:col><xdr:colOff>19050</xdr:colOff>'
            '<xdr:row>1</xdr:row><xdr:rowOff>19050</xdr:rowOff></from>'
            '<to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>1</xdr:row><xdr:rowOff>400050</xdr:rowOff></to>'
            '</anchor></controlPr></control></mc:Choice></mc:AlternateContent>'
            % (34817 + i, 3 + i, i + 1, esc(macro), c0, c0 + 2))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'mc:Ignorable="x14ac" '
        'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        '<dimension ref="A1:N%d"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="6" topLeftCell="A7" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18.6"/>'
        '<cols>%s</cols><sheetData>%s</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '<legacyDrawing r:id="rId2"/>'
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14"><controls>%s</controls></mc:Choice></mc:AlternateContent>'
        '</worksheet>' % (LAST_ROW, cols, ''.join(rows), ''.join(ctrls)))


def build_rels():
    r = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
         'relationships/printerSettings" Target="../printerSettings/printerSettings8.bin"/>',
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
         'relationships/vmlDrawing" Target="../drawings/vmlDrawing10.vml"/>']
    for i in range(len(BUTTONS)):
        r.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp%d.xml"/>'
                 % (3 + i, 17 + i))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '%s</Relationships>' % ''.join(r))


def build_vml():
    # 列幅(文字)→おおよそのpt。ボタンは2列ぶんを占める。
    widths = [w * 5.25 for _, w in COLW]
    shapes = []
    x = 0.0
    for i, (label, macro) in enumerate(BUTTONS):
        c0 = i * 2
        left = sum(widths[:c0])
        wpt = widths[c0] + widths[c0 + 1]
        shapes.append(
            "<v:shape id=\"_x0000_s%d\" type=\"#_x0000_t201\" style='position:absolute;"
            "margin-left:%.1fpt;margin-top:26pt;width:%.1fpt;height:30pt;z-index:%d;"
            "mso-wrap-style:tight' o:button=\"t\" fillcolor=\"buttonFace [67]\" "
            "strokecolor=\"windowText [64]\" o:insetmode=\"auto\">"
            "<v:fill color2=\"buttonFace [67]\" o:detectmouseclick=\"t\"/>"
            "<o:lock v:ext=\"edit\" rotation=\"t\"/>"
            "<v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">"
            "<div style='text-align:center'><font face=\"ＭＳ Ｐゴシック\" size=\"200\" "
            "color=\"#000000\">%s</font></div></v:textbox>"
            "<x:ClientData ObjectType=\"Button\">"
            "<x:Anchor>%d, 3, 1, 3, %d, 0, 1, 30</x:Anchor>"
            "<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill>"
            "<x:FmlaMacro>%s</x:FmlaMacro>"
            "<x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>"
            "</x:ClientData></v:shape>"
            % (34817 + i, left, wpt, i + 1, esc(label), c0, c0 + 2, esc(macro)))
    return ('<xml xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel">'
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="34"/></o:shapelayout>'
            '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" '
            'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
            '<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>'
            '<o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
            '%s</xml>' % ''.join(shapes))


CTRLPROP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
            'objectType="Button" lockText="1"/>')
