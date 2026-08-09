# -*- coding: utf-8 -*-
"""SpreadsheetML のセルを、他の要素を壊さずに書き換える最小ユーティリティ。"""
import re

def col2idx(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n

def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return m.group(1), int(m.group(2))

def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def f_body(formula):
    """数式セルの中身。キャッシュ値は載せない（fullCalcOnLoad で再計算させる）。"""
    return '<f>%s</f>' % esc(formula)

def num_body(v):
    return '<v>%s</v>' % v

def _cell_re(ref):
    return re.compile(r'<c r="%s"(?![0-9])((?:[^>"]|"[^"]*")*?)(/>|>(?:.*?)</c>)' % ref, re.S)

def set_cell(xml, ref, body, keep_style=True, extra_attr='', drop_attrs=('t',)):
    """既存セルを置換。無ければ行内の正しい位置に挿入。行が無ければ行ごと挿入。"""
    col, row = split_ref(ref)
    m = _cell_re(ref).search(xml)
    if m:
        attrs = m.group(1)
        if keep_style:
            sm = re.search(r'\ss="\d+"', attrs)
            style = sm.group(0) if sm else ''
        else:
            style = ''
        new = '<c r="%s"%s%s>%s</c>' % (ref, style, extra_attr, body) if body \
              else '<c r="%s"%s%s/>' % (ref, style, extra_attr)
        return xml[:m.start()] + new + xml[m.end():], True
    # セルが無い → 行を探す
    rm = re.search(r'<row r="%d"((?:[^>"]|"[^"]*")*?)(/>|>(.*?)</row>)' % row, xml, re.S)
    newcell = '<c r="%s"%s>%s</c>' % (ref, extra_attr, body)
    if rm:
        if rm.group(2) == '/>':
            new_row = '<row r="%d"%s>%s</row>' % (row, rm.group(1), newcell)
            return xml[:rm.start()] + new_row + xml[rm.end():], True
        inner = rm.group(3)
        tgt = col2idx(col)
        pos = None
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"', inner):
            if col2idx(cm.group(1)) > tgt:
                pos = cm.start()
                break
        inner = (inner[:pos] + newcell + inner[pos:]) if pos is not None else (inner + newcell)
        new_row = '<row r="%d"%s>%s</row>' % (row, rm.group(1), inner)
        return xml[:rm.start()] + new_row + xml[rm.end():], True
    # 行も無い → 行を挿入
    newrow = '<row r="%d">%s</row>' % (row, newcell)
    pos = None
    for rmm in re.finditer(r'<row r="(\d+)"', xml):
        if int(rmm.group(1)) > row:
            pos = rmm.start()
            break
    if pos is None:
        em = re.search(r'</sheetData>', xml)
        pos = em.start()
    return xml[:pos] + newrow + xml[pos:], True

def set_formula(xml, ref, formula, extra_attr=''):
    return set_cell(xml, ref, f_body(formula), extra_attr=extra_attr)[0]

def set_number(xml, ref, v, extra_attr=''):
    return set_cell(xml, ref, num_body(v), extra_attr=extra_attr)[0]

def set_text(xml, ref, text, extra_attr=''):
    body = '<is><t xml:space="preserve">%s</t></is>' % esc(text)
    col, row = split_ref(ref)
    m = _cell_re(ref).search(xml)
    style = ''
    if m:
        sm = re.search(r'\ss="\d+"', m.group(1))
        style = sm.group(0) if sm else ''
        new = '<c r="%s"%s t="inlineStr"%s>%s</c>' % (ref, style, extra_attr, body)
        return xml[:m.start()] + new + xml[m.end():]
    return set_cell(xml, ref, body, extra_attr=' t="inlineStr"' + extra_attr)[0]

def widen_dimension(xml, last_ref):
    return re.sub(r'<dimension ref="([A-Z]+\d+):[A-Z]+\d+"/>',
                  lambda m: '<dimension ref="%s:%s"/>' % (m.group(1), last_ref), xml, count=1)
