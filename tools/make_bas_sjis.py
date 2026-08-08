#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBAモジュール(.bas)を Shift-JIS(CP932) + CRLF で書き出す

■ なぜ必要か

  日本語Windows の VBE（Alt+F11 → ファイルのインポート）は、.bas を
  システムのANSIコードページ＝**Shift-JIS(CP932)** として読む。
  UTF-8 のまま渡すと、コメント・MsgBoxの文言・日本語のプロシージャ名が
  すべて文字化けする（このリポジトリでは 35d0cd5 でも同じ事故が起きている）。

  このブックの既存モジュール（Mod_日次更新.bas / Module7.bas 等）も
  すべて CP932 なので、それに合わせる。

■ 運用

  編集するのは  tools/bas/*.utf8.bas （UTF-8・改行LF。差分が読める）
  配布するのは  リポジトリ直下の *.bas （CP932・改行CRLF。VBEがそのまま読める）

  編集したら、このスクリプトを実行して配布版を作り直す。

      python3 tools/make_bas_sjis.py

  CP932 で表現できない文字（絵文字・①などの機種依存文字・全角チルダ等）が
  混ざっていたら、その場でエラーにして知らせる。VBAのコメントに絵文字は使わないこと。
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "tools" / "bas"


def convert(src: Path) -> Path:
    dst = ROOT / (src.name[: -len(".utf8.bas")] + ".bas")
    text = src.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
    try:
        data = text.encode("cp932")
    except UnicodeEncodeError as e:
        bad = text[e.start:e.end]
        line = text[: e.start].count("\n") + 1
        raise SystemExit(
            "%s の %d行目に Shift-JIS で表せない文字があります: %r\n"
            "  VBE が読めないので、別の文字に置き換えてください。" % (src.name, line, bad))
    dst.write_bytes(data)

    # 往復確認（読み直して一致するか）
    back = dst.read_bytes().decode("cp932")
    assert back == text, "%s の往復変換で内容が変わった" % dst.name
    return dst


def main():
    if not SRC_DIR.is_dir():
        raise SystemExit("%s がありません" % SRC_DIR)
    srcs = sorted(SRC_DIR.glob("*.utf8.bas"))
    if not srcs:
        raise SystemExit("%s に *.utf8.bas がありません" % SRC_DIR)
    for src in srcs:
        dst = convert(src)
        print("  %-34s -> %-24s %6d bytes  CP932 / CRLF" % (src.name, dst.name, dst.stat().st_size))
    print("完了（%d ファイル）" % len(srcs))


if __name__ == "__main__":
    sys.exit(main())
