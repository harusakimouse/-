#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
共有数式(shared formula)の展開。

Excelは同じ形の数式を「マスター1個＋追従セル」で持つ。
   マスター : <c r="T3"><f t="shared" ref="T3:T66" si="4">IF(...)</f><v>8</v></c>
   追従     : <c r="T4"><f t="shared" si="4"/><v>5</v></c>

マスターのセルを書き換えると si="4" の定義が消え、追従セルが全部孤児になる。
Excel はこれを検出してシート内の数式を丸ごと削除する
（「削除されたレコード: /xl/worksheets/sheetN.xml パーツ内の数式」）。

そこでパッチを当てる前に、共有数式をすべて「そのセル専用の普通の数式」に
展開しておく。展開後の数式文字列は openpyxl が正しく相対参照を
ずらして生成してくれるので、それを使う。

配列数式(<f t="array">, RSS関数など)は触らない。
"""
import re
import openpyxl


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def expand_shared(xml: str, ws) -> tuple:
    """xml 内の shared 数式セルを、openpyxl が展開した数式で置き換える。

    注意：<c> 要素の範囲を厳密に区切ること。`.*?</f>` のような緩い書き方だと
    追従セル（<f t="shared" si="0"/> の自己終了タグ）を処理するときに
    次のマスターセルの </f> まで食ってしまい、間の <row> 構造ごと消える。
    """
    n_master = n_follow = n_fail = 0

    # 1つの <c> 要素だけに確実にマッチさせる（内部に </c> を含ませない）
    cell_pat = re.compile(
        r'<c\s[^>]*?r="(?P<ref>[A-Z]+\d+)"[^>]*?'
        r'(?:/>|>(?P<inner>(?:(?!</c>).)*)</c>)', re.S)

    def repl(m):
        nonlocal n_master, n_follow, n_fail
        inner = m.group("inner")
        if inner is None or 't="shared"' not in inner:
            return m.group(0)                     # 共有数式でないセルは触らない
        ref = m.group("ref")
        try:
            v = ws[ref].value
        except Exception:
            v = None
        if not isinstance(v, str) or not v.startswith("="):
            n_fail += 1
            return m.group(0)
        if "ref=" in inner.split(">", 1)[0]:
            n_master += 1
        else:
            n_follow += 1
        head = re.match(r'<c\s[^>]*?r="[^"]*"[^>]*?>', m.group(0)).group(0)
        head = re.sub(r'\s+t="[a-z]+"', "", head)      # 型指定を外す
        return head + f"<f>{esc(v[1:])}</f></c>"

    out = cell_pat.sub(repl, xml)
    return out, (n_master, n_follow, n_fail)


def expand_file_sheets(src_xlsm, items, sheet_map):
    """sheet_map = {xmlパス: シート名} を openpyxl の展開結果で書き換える。"""
    wb = openpyxl.load_workbook(src_xlsm, data_only=False)
    report = []
    for path, name in sheet_map.items():
        x = items[path].decode("utf-8")
        before = len(re.findall(r't="shared"', x))
        x, (nm, nf, bad) = expand_shared(x, wb[name])
        after = len(re.findall(r't="shared"', x))
        items[path] = x.encode("utf-8")
        report.append(f"  {name:<8} 共有数式 {before}個 → {after}個 "
                      f"(マスター{nm} / 追従{nf} を展開、未展開{bad})")
    return report
