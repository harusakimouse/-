#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""フォルダ内の PDF だけを、中身を読んで分かりやすい名前にリネームします。

PDF 以外のファイル（.xlsx, .docx, 画像など）には一切触れません。
既定はドライラン（実際には変えず、変更案だけ表示）。--apply で実行します。
スキャンした画像PDFは、Tesseract OCR があれば自動で文字を読み取ります。

使い方:
    python rename_pdfs.py "C:\\Users\\Abe Hideaki\\...\\スキャン　マンション"          # 変更案だけ表示
    python rename_pdfs.py "C:\\Users\\Abe Hideaki\\...\\スキャン　マンション" --apply  # 実際にリネーム
    python rename_pdfs.py フォルダ --apply --recursive                                 # サブフォルダも対象
    python rename_pdfs.py フォルダ --undo                                              # 直前のリネームを元に戻す

必要なライブラリ（日本語PDFは PyMuPDF が最も確実）:
    pip install pymupdf
スキャンPDF（画像）を読む場合は追加で:
    pip install pytesseract pymupdf
    + Tesseract 本体と日本語データ（Windows: https://github.com/UB-Mannheim/tesseract/wiki
      インストール時に「Japanese」にチェック）
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import re
import shutil
import sys
import unicodedata
from pathlib import Path

LOG_NAME = ".pdf_rename_log.csv"
MIN_TEXT = 30  # これ未満の文字数ならスキャンPDFとみなして OCR に回す

# ---------------------------------------------------------------- 本文の抽出


def _text_pymupdf(path: Path, max_pages: int):
    import fitz  # PyMuPDF

    with fitz.open(path) as doc:
        title = (doc.metadata or {}).get("title") or ""
        text = "\n".join(doc[i].get_text() for i in range(min(max_pages, doc.page_count)))
    return text, title


def _text_pdfminer(path: Path, max_pages: int):
    from pdfminer.high_level import extract_text as _extract

    return _extract(str(path), maxpages=max_pages) or "", ""


def _text_pypdf(path: Path, max_pages: int):
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader  # 古い環境向けフォールバック

    reader = PdfReader(str(path))
    title = ""
    try:
        title = (reader.metadata or {}).get("/Title") or ""
    except Exception:
        pass
    return "\n".join((p.extract_text() or "") for p in reader.pages[:max_pages]), str(title)


def ocr_available() -> bool:
    try:
        import fitz  # noqa: F401
        import pytesseract  # noqa: F401
    except ImportError:
        return False
    return shutil.which("tesseract") is not None or bool(
        getattr(__import__("pytesseract").pytesseract, "tesseract_cmd", "") and
        Path(__import__("pytesseract").pytesseract.tesseract_cmd).exists()
    )


def _text_ocr(path: Path, max_pages: int, lang: str = "jpn+eng"):
    """スキャンPDFをページ画像に変換して OCR する。"""
    import fitz
    import pytesseract
    from PIL import Image

    out = []
    with fitz.open(path) as doc:
        for i in range(min(max_pages, doc.page_count)):
            pix = doc[i].get_pixmap(dpi=300)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            out.append(pytesseract.image_to_string(img, lang=lang))
    return "\n".join(out)


def extract_text(path: Path, max_pages: int = 2, use_ocr: bool = True, lang: str = "jpn+eng"):
    """(本文, タイトル, 手段) を返す。取り出せなければ (None, '', 手段)。"""
    best_text, best_title = "", ""
    for fn in (_text_pymupdf, _text_pdfminer, _text_pypdf):
        try:
            text, title = fn(path, max_pages)
        except ImportError:   # そのライブラリが未インストール
            continue
        except Exception:     # 壊れたPDF・暗号化PDFなど
            continue
        text = text or ""
        if len(text.strip()) > len(best_text.strip()):
            best_text, best_title = text, (title or "").strip()
        if len(best_text.strip()) >= MIN_TEXT:
            return best_text, best_title, "text"

    if use_ocr:
        try:
            text = _text_ocr(path, max_pages, lang)
            if len(text.strip()) >= 5:
                return text, best_title, "ocr"
        except ImportError:
            pass
        except Exception:
            pass

    if best_text.strip():
        return best_text, best_title, "text"
    return None, best_title, "none"


# ---------------------------------------------------------------- 情報の推定

DOC_TYPES = [
    # 一般
    "御請求書", "請求書", "御見積書", "見積書", "領収書", "納品書", "受領書",
    "注文書", "発注書", "注文請書", "契約書", "覚書", "誓約書", "同意書",
    "報告書", "議事録", "通知書", "案内書", "明細書", "支払明細", "給与明細",
    "源泉徴収票", "決算報告書", "決算書", "履歴書", "提案書", "企画書",
    "仕様書", "取扱説明書", "マニュアル", "申込書", "届出書", "証明書", "確認書",
    # 不動産・マンション関連
    "重要事項説明書", "重要事項調査報告書", "管理規約", "使用細則",
    "総会議事録", "理事会議事録", "定期総会", "臨時総会",
    "長期修繕計画", "修繕積立金", "管理費", "管理委託契約書",
    "賃貸借契約書", "売買契約書", "媒介契約書", "重要事項",
    "登記事項証明書", "登記簿謄本", "固定資産税", "納税通知書",
    "火災保険", "地震保険", "保険証券", "検査済証", "確認済証",
    "設計図書", "平面図", "間取図", "図面", "点検報告書", "工事報告書",
]

_JP = r"一-龥ぁ-んァ-ヶーA-Za-z0-9＆・．\-"

COMPANY_PATTERNS = [
    re.compile(r"(?:株式会社|有限会社|合同会社|合資会社|一般社団法人|公益社団法人|一般財団法人|医療法人|学校法人)[" + _JP + r"]{1,20}"),
    re.compile(r"[" + _JP + r"]{1,20}(?:株式会社|有限会社|合同会社|合資会社)"),
]

# マンション名・建物名（スキャンした物件書類向け）
BUILDING_PATTERN = re.compile(
    r"[一-龥ぁ-んァ-ヶーA-Za-z0-9・]{2,18}"
    r"(?:マンション|ハイツ|コーポ|レジデンス|パレス|ハイム|ヴィラ|ヴィレッジ|タワー|ホームズ|ガーデン|テラス|荘)"
)

DATE_PATTERNS = [
    re.compile(r"(?P<y>(?:19|20)\d{2})\s*[年/\-\.]\s*(?P<m>\d{1,2})\s*[月/\-\.]\s*(?P<d>\d{1,2})"),
    re.compile(r"(?P<era>令和|平成|昭和)\s*(?P<ey>\d{1,2}|元)\s*年\s*(?P<m>\d{1,2})\s*月\s*(?P<d>\d{1,2})"),
]
ERA_BASE = {"令和": 2018, "平成": 1988, "昭和": 1925}

DATE_LABEL = re.compile(r"(発行日|作成日|請求日|見積日|日付|作成年月日|発行年月日|締結日)")


def normalize(text: str) -> str:
    return unicodedata.normalize("NFKC", text).replace("\u3000", " ")


def find_doc_type(text: str):
    best = None
    for word in DOC_TYPES:
        pos = text.find(word)
        if pos >= 0 and (best is None or pos < best[1] or (pos == best[1] and len(word) > len(best[0]))):
            best = (word, pos)
    return best[0] if best else None


def _pick(counts: dict, order: dict):
    if not counts:
        return None
    return sorted(counts, key=lambda n: (-counts[n], order[n]))[0]


def find_company(text: str):
    counts, order = {}, {}
    for pat in COMPANY_PATTERNS:
        for m in pat.finditer(text):
            name = m.group(0).strip("・．- ")
            if len(name) < 5:
                continue
            counts[name] = counts.get(name, 0) + 1
            order.setdefault(name, m.start())
    return _pick(counts, order)


def find_building(text: str):
    counts, order = {}, {}
    for m in BUILDING_PATTERN.finditer(text):
        name = m.group(0).strip("・- ")
        if len(name) < 4 or name.startswith(("マンション", "この", "本")):
            continue
        counts[name] = counts.get(name, 0) + 1
        order.setdefault(name, m.start())
    return _pick(counts, order)


def _to_date(m: re.Match):
    try:
        if m.groupdict().get("era"):
            ey = m.group("ey")
            year = ERA_BASE[m.group("era")] + (1 if ey == "元" else int(ey))
        else:
            year = int(m.group("y"))
        return _dt.date(year, int(m.group("m")), int(m.group("d")))
    except (ValueError, KeyError):
        return None


def find_date(text: str):
    hits = []
    for pat in DATE_PATTERNS:
        for m in pat.finditer(text):
            d = _to_date(m)
            if d and 1970 <= d.year <= _dt.date.today().year + 2:
                labeled = bool(DATE_LABEL.search(text[max(0, m.start() - 50):m.start()]))
                hits.append((0 if labeled else 1, m.start(), d))
    if not hits:
        return None
    hits.sort()
    return hits[0][2]


def first_meaningful_line(text: str, limit: int = 40):
    for raw in text.splitlines():
        line = raw.strip(" \t　-–—=*_|")
        if len(line) < 3 or len(line) > 60:
            continue
        if re.fullmatch(r"[\d\s\-/\.:,、。]+", line):
            continue
        return line[:limit]
    return None


# ---------------------------------------------------------------- 名前の組み立て

INVALID = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
RESERVED = {"CON", "PRN", "AUX", "NUL",
            *(f"COM{i}" for i in range(1, 10)),
            *(f"LPT{i}" for i in range(1, 10))}


def sanitize(name: str, max_len: int = 80) -> str:
    name = INVALID.sub("", normalize(name))
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"_{2,}", "_", name).strip("_. ")
    if name.upper() in RESERVED:
        name = f"_{name}"
    return name[:max_len].strip("_. ")


def build_name(text: str, meta_title: str = ""):
    """本文から新しいファイル名（拡張子なし）を作る。作れなければ None。"""
    text = normalize(text)
    doc_type = find_doc_type(text)
    subject = find_building(text) or find_company(text)
    date = find_date(text)

    parts = [p for p in (doc_type, subject) if p]
    if not parts:
        title = normalize(meta_title).strip()
        if len(title) < 3 or title.lower().endswith(".pdf") or "untitled" in title.lower():
            title = ""
        fallback = title or first_meaningful_line(text)
        if not fallback:
            return None
        parts = [fallback]
    if date:
        parts.append(date.isoformat())

    return sanitize("_".join(parts)) or None


def unique_path(directory: Path, stem: str, suffix: str, taken: set) -> Path:
    candidate = directory / f"{stem}{suffix}"
    i = 2
    while candidate.exists() or candidate in taken:
        candidate = directory / f"{stem}_{i}{suffix}"
        i += 1
    return candidate


# ---------------------------------------------------------------- 実行部


def collect_pdfs(folder: Path, recursive: bool):
    it = folder.rglob("*") if recursive else folder.glob("*")
    # PDF 以外は一切対象にしない
    return sorted(p for p in it if p.is_file() and p.suffix.lower() == ".pdf")


def run(folder: Path, apply: bool, recursive: bool, use_ocr: bool, lang: str):
    pdfs = collect_pdfs(folder, recursive)
    if not pdfs:
        print(f"PDF が見つかりませんでした: {folder}")
        return 0

    if use_ocr and not ocr_available():
        print("※ OCR は使えません（pytesseract / Tesseract 未導入）。文字データのないスキャンPDFはスキップします。\n")
        use_ocr = False

    print(f"対象: {len(pdfs)} 件の PDF（PDF 以外は対象外）\n")
    renamed = skipped = 0
    taken, log = set(), []

    for pdf in pdfs:
        text, meta_title, how = extract_text(pdf, use_ocr=use_ocr, lang=lang)
        if text is None:
            print(f"[スキップ] {pdf.name} … 文字を取り出せません（スキャン画像PDF。OCR の導入が必要）")
            skipped += 1
            continue

        stem = build_name(text, meta_title)
        if not stem:
            print(f"[スキップ] {pdf.name} … 手掛かりになる情報が見つかりません")
            skipped += 1
            continue
        if stem == pdf.stem:
            print(f"[そのまま] {pdf.name}")
            continue

        target = unique_path(pdf.parent, stem, pdf.suffix, taken)
        taken.add(target)
        mark = "OCR" if how == "ocr" else "文字"
        print(f"[{'リネーム' if apply else '変更案'}/{mark}] {pdf.name}\n        → {target.name}")

        if apply:
            try:
                pdf.rename(target)
            except OSError as e:
                print(f"        !! 失敗: {e}")
                skipped += 1
                continue
            log.append((str(pdf), str(target)))
        renamed += 1

    print(f"\n完了: {renamed} 件{'をリネーム' if apply else 'が変更対象'} / {skipped} 件スキップ")
    if log:
        log_path = folder / LOG_NAME
        new_file = not log_path.exists()
        with log_path.open("a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(["old", "new"])
            w.writerows(log)
        print(f"元に戻す用のログ: {log_path}")
    elif not apply and renamed:
        print("実際に変更するには --apply を付けて再実行してください。")
    return 0


def undo(folder: Path):
    log_path = folder / LOG_NAME
    if not log_path.exists():
        print(f"ログがありません: {log_path}")
        return 1
    with log_path.open(encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))[1:]
    restored = 0
    for old, new in reversed(rows):
        new_p, old_p = Path(new), Path(old)
        if new_p.exists() and not old_p.exists():
            new_p.rename(old_p)
            print(f"[戻す] {new_p.name} → {old_p.name}")
            restored += 1
    log_path.unlink()
    print(f"\n{restored} 件を元に戻しました。")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="フォルダ内の PDF だけを中身から自動命名します")
    ap.add_argument("folder", type=Path, help="対象フォルダ")
    ap.add_argument("--apply", action="store_true", help="実際にリネームする（既定はドライラン）")
    ap.add_argument("--recursive", action="store_true", help="サブフォルダも対象にする")
    ap.add_argument("--undo", action="store_true", help="直前のリネームを元に戻す")
    ap.add_argument("--no-ocr", action="store_true", help="スキャンPDFの OCR を使わない")
    ap.add_argument("--lang", default="jpn+eng", help="OCR の言語（既定: jpn+eng）")
    args = ap.parse_args(argv)

    folder = args.folder.expanduser()
    if not folder.is_dir():
        print(f"フォルダが見つかりません: {folder}", file=sys.stderr)
        return 1
    if args.undo:
        return undo(folder)
    return run(folder, args.apply, args.recursive, not args.no_ocr, args.lang)


if __name__ == "__main__":
    raise SystemExit(main())
