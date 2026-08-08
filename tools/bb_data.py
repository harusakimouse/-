#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V815 の価格シート（始値/高値/安値/終値/出来高）を numpy 配列に取り出す共通モジュール。

シートの並び:
    A列  銘柄コード
    B列  銘柄名
    C列  RSS現値（当日ザラ場・検証には使わない）
    D列  差分
    E列  最新営業日の確定値   ← ここから右へ「古く」なる
    F列  1日前 … IT列(254) まで最大250営業日

    行5   TOPX（指数行・銘柄ではない）
    行6〜 個別銘柄（最大400行）
    行3   各列の日付（Excelシリアル値）

backtest 側は「新しい→古い」の並びだと扱いづらいので、
取り出した時点で時系列昇順（古い→新しい）に反転して返す。
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta

import numpy as np

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# V815 内でのワークシート実体
SHEET_XML = {
    "open":   "sheet11",   # 始値
    "high":   "sheet12",   # 高値
    "low":    "sheet13",   # 安値
    "close":  "sheet14",   # 終値
    "volume": "sheet15",   # 出来高
}
FIRST_DATA_COL = 5      # E列
LAST_DATA_COL = 254     # IT列
FIRST_STOCK_ROW = 6
LAST_STOCK_ROW = 405
DATE_ROW = 3


def col_index(ref: str) -> int:
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
        else:
            break
    return n


def row_index(ref: str) -> int:
    return int(re.sub(r"[A-Z]", "", ref))


def _read_sheet(z, name, shared):
    """1シートを (rows, cols) の dict にして返す。値だけ拾う。"""
    root = ET.fromstring(z.read("xl/worksheets/%s.xml" % SHEET_XML[name]))
    cells = {}
    for row in root.iter(NS + "row"):
        r = int(row.get("r"))
        for c in row:
            ref = c.get("r")
            v = c.find(NS + "v")
            if v is None or v.text is None:
                continue
            t = c.get("t")
            if t == "s":
                cells[(r, col_index(ref))] = shared[int(v.text)]
            elif t in (None, "n"):
                try:
                    cells[(r, col_index(ref))] = float(v.text)
                except ValueError:
                    pass
            elif t == "str":
                cells[(r, col_index(ref))] = v.text
    return cells


def excel_serial_to_date(serial: float) -> date:
    return date(1899, 12, 30) + timedelta(days=int(serial))


def load(path="V815.xlsm"):
    """
    戻り値 dict:
        codes  (N,)      銘柄コード（文字列）
        names  (N,)      銘柄名
        dates  (T,)      datetime.date  古い→新しい
        open/high/low/close/volume  (N, T) float  古い→新しい、欠損は nan
    """
    z = zipfile.ZipFile(path)
    shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]

    close_cells = _read_sheet(z, "close", shared)

    # 日付行（E..IT）。値が入っている列だけを実データ列とみなす
    date_cols, dates = [], []
    for c in range(FIRST_DATA_COL, LAST_DATA_COL + 1):
        v = close_cells.get((DATE_ROW, c))
        if isinstance(v, float) and v > 40000:
            date_cols.append(c)
            dates.append(excel_serial_to_date(v))

    # 銘柄行（コードが数値で入っている行）
    rows, codes, names = [], [], []
    for r in range(FIRST_STOCK_ROW, LAST_STOCK_ROW + 1):
        code = close_cells.get((r, 1))
        if code is None or code == "":
            continue
        rows.append(r)
        codes.append(str(int(code)) if isinstance(code, float) else str(code))
        names.append(close_cells.get((r, 2), ""))

    out = {"codes": np.array(codes), "names": np.array(names)}

    for key in ("open", "high", "low", "close", "volume"):
        cells = close_cells if key == "close" else _read_sheet(z, key, shared)
        arr = np.full((len(rows), len(date_cols)), np.nan)
        for i, r in enumerate(rows):
            for j, c in enumerate(date_cols):
                v = cells.get((r, c))
                if isinstance(v, float):
                    arr[i, j] = v
        out[key] = arr[:, ::-1]          # 新しい→古い を 古い→新しい に反転

    out["dates"] = np.array(dates[::-1])
    z.close()
    return out


if __name__ == "__main__":
    import sys
    d = load(sys.argv[1] if len(sys.argv) > 1 else "V815.xlsm")
    N, T = d["close"].shape
    print("銘柄数 %d / 営業日数 %d" % (N, T))
    print("期間 %s 〜 %s" % (d["dates"][0], d["dates"][-1]))
    ok = ~np.isnan(d["close"])
    print("終値の充足率 %.1f%%" % (100 * ok.mean()))
    per_day = ok.sum(axis=0)
    print("日別の有効銘柄数: 最古%d 中央%d 最新%d" % (per_day[0], int(np.median(per_day)), per_day[-1]))
    full = (ok.all(axis=1)).sum()
    print("全期間そろっている銘柄: %d" % full)
    print("先頭5銘柄:", ", ".join("%s %s" % (c, n) for c, n in zip(d["codes"][:5], d["names"][:5])))
