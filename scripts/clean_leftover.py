#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模板残留清理：按替换表把残留示例文字替换为真实内容（就地保存）。

用法:
    python3 clean_leftover.py <pptx> <repl.json>

repl.json 格式: {"replace": [["旧词","新词"], ...]}
命中即整段替换；清理后必须再跑 scan_leftover.py 验证。
"""
from pptx import Presentation
import json, sys

def walk(sh):
    for s in sh:
        if s.shape_type == 6:
            yield from walk(s.shapes)
        else:
            yield s

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    path, repl_path = sys.argv[1], sys.argv[2]
    repl = json.load(open(repl_path))["replace"]
    prs = Presentation(path)
    n = 0
    for slide in prs.slides:
        for shape in walk(slide.shapes):
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                rs = para.runs
                if not rs:
                    continue
                full = "".join(r.text for r in rs)
                new = full
                for a, b in repl:
                    if a in new:
                        new = new.replace(a, b)
                if new != full:
                    rs[0].text = new
                    for r in rs[1:]:
                        r.text = ""
                    n += 1
    prs.save(path)
    print(f"替换 {n} 段 -> {path}")
