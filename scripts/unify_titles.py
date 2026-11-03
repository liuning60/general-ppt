#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一标题：删掉所有旧标题文字，按"章名 / 节名"给每个内容页重加统一标题框。

用法:
    python3 unify_titles.py <pptx> <chapters.json>

chapters.json 格式:
{
  "content_pages": [1,2,3,5,6,8,9,11,12,13,15,16,18,20,22,23,25,26,27],
  "chapters": {
    "1": ["课程定位与目标", "建设背景"],
    "2": ["课程定位与目标", "课程定位"]
  },
  "title_color": "1F3A6E",
  "sub_color": "999999"
}
"""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
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
    path, cfg_path = sys.argv[1], sys.argv[2]
    cfg = json.load(open(cfg_path))
    content_pages = cfg["content_pages"]
    chapters = {int(k): v for k, v in cfg["chapters"].items()}
    tc = cfg.get("title_color", "1F3A6E")
    sc = cfg.get("sub_color", "999999")
    TITLE_C = RGBColor(int(tc[0:2],16), int(tc[2:4],16), int(tc[4:6],16))
    SUB_C = RGBColor(int(sc[0:2],16), int(sc[2:4],16), int(sc[4:6],16))

    prs = Presentation(path)

    # 1) 清空所有内容页顶部旧标题文字（顶部 1.4cm 内、字号>=16 的文本框）
    cleared = 0
    for idx in content_pages:
        slide = prs.slides[idx - 1]
        for shape in walk(slide.shapes):
            if not shape.has_text_frame:
                continue
            try:
                top = shape.top or 0
            except Exception:
                continue
            if top > 1300000:
                continue
            maxsz = 0
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        maxsz = max(maxsz, r.font.size.pt)
            if maxsz >= 16:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = ""
                cleared += 1

    # 2) 每页加统一标题框：章名(粗/大/主题色) + 斜杠 + 节名(细/灰)
    for idx in content_pages:
        if idx not in chapters:
            continue
        chap, sec = chapters[idx]
        slide = prs.slides[idx - 1]
        tb = slide.shapes.add_textbox(Emu(457200), Emu(180000), Emu(9000000), Emu(550000))
        p = tb.text_frame.paragraphs[0]
        r1 = p.add_run(); r1.text = chap
        r1.font.size = Pt(28); r1.font.bold = True; r1.font.color.rgb = TITLE_C; r1.font.name = "微软雅黑"
        r2 = p.add_run(); r2.text = "  /  "
        r2.font.size = Pt(24); r2.font.color.rgb = SUB_C
        r3 = p.add_run(); r3.text = sec
        r3.font.size = Pt(22); r3.font.color.rgb = SUB_C; r3.font.name = "微软雅黑"

    prs.save(path)
    print(f"清空旧标题 {cleared} 处，统一标题 {len([i for i in content_pages if i in chapters])} 页 -> {path}")
