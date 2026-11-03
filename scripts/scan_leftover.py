#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模板残留扫描：逐页找出与主题无关的模板示例词/占位词，交付前必跑。

用法:
    python3 scan_leftover.py <pptx> [keywords.json]

keywords.json 可选，格式: {"words": ["三爱", "议题", "Question", ...]}
不传则用内置默认清单（思政/法治/会计常见示例词 + 通用占位词）。
"""
from pptx import Presentation
import json, sys

DEFAULT_WORDS = [
    "道德与法治", "议题", "三爱", "阶级", "国家统一", "爱国", "爱党", "爱社会",
    "马克思", "思政", "政治", "法治", "责任担当", "公共参与", "政治认同", "价值观",
    "德育", "社会存在", "一带一路", "民族", "法律", "道德", "原理课", "研究目标",
    "课题", "会计要素", "企业会计准则", "财务报告", "消费者权益", "课程思政",
    "Question", "Vivamus", "Lorem", "Key Words", "项目名称", "此处插入", "TODO",
]

def walk(sh):
    for s in sh:
        if s.shape_type == 6:
            yield from walk(s.shapes)
        else:
            yield s

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    path = sys.argv[1]
    words = DEFAULT_WORDS
    if len(sys.argv) >= 3:
        words = json.load(open(sys.argv[2])).get("words", DEFAULT_WORDS)
    prs = Presentation(path)
    hits = 0
    for i, slide in enumerate(prs.slides, 1):
        out = []
        for shape in walk(slide.shapes):
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs).strip()
                if t:
                    out.append(t)
        found = [t for t in out if any(w in t for w in words)]
        if found:
            hits += 1
            print(f"P{i}: {found[:5]}")
    if hits == 0:
        print("OK  无残留命中")
    else:
        print(f"警告 {hits} 页有残留，先清理再交付")
