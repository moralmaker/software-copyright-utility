#!/usr/bin/env python3
"""对已抽取的软著代码材料做重品牌处理（不改动仓库真实源码）。
仅处理 草稿/代码-*.md：包名、@author、@since、@version。
用法：python3 rebrand_code_material.py --workdir 软件著作权申请资料
"""
import argparse
import re
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workdir", required=True)
    ap.add_argument("--pkg-from", default="com.xxx.www")
    ap.add_argument("--pkg-to", default="com.moral.www")
    ap.add_argument("--author", default="道義")
    ap.add_argument("--since", default="2025/09/30")
    ap.add_argument("--version", default="1.0.0")
    a = ap.parse_args()

    wd = Path(a.workdir)
    files = sorted(wd.glob("草稿/代码-*.md"))
    if not files:
        print("未找到 草稿/代码-*.md")
        return

    totals = {"pkg_dot": 0, "pkg_slash": 0, "author": 0, "since": 0, "version": 0}
    for md in files:
        s = md.read_text(encoding="utf-8")
        pkg_from = a.pkg_from
        pkg_to = a.pkg_to
        s, n1 = re.subn(re.escape(pkg_from), pkg_to, s)
        s, n2 = re.subn(re.escape(pkg_from.replace(".", "/")), pkg_to.replace(".", "/"), s)
        s, n3 = re.subn(r"(@author)\s+[^\n]*", lambda m: m.group(1) + " " + a.author, s)
        s, n4 = re.subn(r"(@since)\s+[\d\-./]+", lambda m: m.group(1) + " " + a.since, s)
        s, n5 = re.subn(r"(@version)\s+[\d.]+", lambda m: m.group(1) + " " + a.version, s)
        md.write_text(s, encoding="utf-8")
        totals["pkg_dot"] += n1
        totals["pkg_slash"] += n2
        totals["author"] += n3
        totals["since"] += n4
        totals["version"] += n5
        print(f"{md.name}: pkg_dot={n1} pkg_slash={n2} author={n3} since={n4} version={n5}")

    print("合计:", totals)
    # 校验残留
    for md in files:
        s = md.read_text(encoding="utf-8")
        for bad, label in [
            (a.pkg_from, "pkg_dot"),
            (a.pkg_from.replace(".", "/"), "pkg_slash"),
        ]:
            if bad in s:
                print(f"[残留] {md.name} 仍有 {label}: {bad}")
        if re.search(r"@(author|version)\s+", s) and (
            a.author not in s or a.version not in s
        ):
            pass
    print("OK rebrand done")


if __name__ == "__main__":
    main()
