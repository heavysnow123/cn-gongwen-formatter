"""命令行入口：word-formatter-cli。

用法示例：
  python -m word_formatter.cli -c config.json -o out *.docx
  python -m word_formatter.cli --text "一、标题\n正文..." -o result.docx
"""

from __future__ import annotations

import argparse
import os
import sys

from .config import FormatterConfig
from .core import WordFormatter


def main(argv=None):
    ap = argparse.ArgumentParser(description="Word Formatter 命令行")
    ap.add_argument("files", nargs="*", help="待处理文件(.docx/.doc/.wps/.txt/.md)")
    ap.add_argument("-c", "--config", help="配置 JSON 路径")
    ap.add_argument("-o", "--out-dir", help="输出目录")
    ap.add_argument("--text", help="直接排版文本（默认强制 A4）")
    ap.add_argument("--markdown", action="store_true", help="--text 为 Markdown")
    ap.add_argument("--default-config", action="store_true",
                    help="生成默认配置文件并退出")
    args = ap.parse_args(argv)

    if args.default_config:
        path = FormatterConfig.default_config_path()
        FormatterConfig().save(path)
        print(f"默认配置已写入：{path}")
        return 0

    cfg = FormatterConfig.load(args.config) if args.config else FormatterConfig()
    fmt = WordFormatter(cfg, log_cb=lambda m, l: print(f"[{l}] {m}"))

    if args.text:
        out = args.out_dir or "."
        os.makedirs(out, exist_ok=True)
        op = os.path.join(out, "formatted_document.docx")
        r = fmt.format_text(args.text, op, is_md=args.markdown)
        if r["error"]:
            print("失败：", r["error"]); return 1
        print("已保存：", r["output"]); return 0

    if not args.files:
        ap.print_help()
        return 1

    ok = skip = err = 0
    for f in args.files:
        r = fmt.format_file(f, args.out_dir)
        if r["skipped"]:
            skip += 1
        elif r["error"]:
            err += 1
        else:
            ok += 1
    print(f"\n完成：成功 {ok} / 跳过 {skip} / 失败 {err}")
    return 0 if err == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
