#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在内容页后插入配图页（深色标题条 + 2 张图左右并排 + 图注），配图页放在对应内容页之后。

用法:
    python3 add_photo_pages.py <src.pptx> <dst.pptx> <photos.json>

photos.json 格式:
{
  "bar_color": "1F3A6E",           # 标题条颜色，可选，默认深蓝
  "pages": [
    {"after": 3, "title": "教学工具栈 · Premiere Pro + Audition",
     "images": [{"path": "images/pr.jpg", "caption": "图1｜Premiere Pro 时间线剪辑", "source": "https://..."}]}
  ]
}
source 有值时自动拼进图注（"图N｜标题｜来源"）。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import json, sys

DEEP_BLUE = RGBColor(0x1F, 0x3A, 0x6E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__); sys.exit(1)
    src, dst, cfg_path = sys.argv[1], sys.argv[2], sys.argv[3]
    cfg = json.load(open(cfg_path))
    bar_hex = cfg.get("bar_color", "1F3A6E")
    BAR = RGBColor(int(bar_hex[0:2], 16), int(bar_hex[2:4], 16), int(bar_hex[4:6], 16))
    pages = cfg["pages"]

    prs = Presentation(src)
    SW, SH = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    def add_photo_page(title, imgs):
        slide = prs.slides.add_slide(blank)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.9))
        bar.fill.solid(); bar.fill.fore_color.rgb = BAR
        bar.line.fill.background()
        tf = bar.text_frame; tf.margin_left = Inches(0.5); tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = title
        r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = WHITE
        r.font.name = "微软雅黑"
        top = Inches(1.15)
        if len(imgs) >= 2:
            gap = Inches(0.3)
            total_w = SW - Inches(0.6)
            img_w = int((total_w - gap) / 2)
            img_h = Inches(4.6)
            xs = [Inches(0.3), Inches(0.3) + img_w + gap]
            for i, (item) in enumerate(imgs[:2]):
                path = item["path"]
                cap_text = item["caption"]
                if item.get("source"):
                    cap_text = f"{cap_text}｜{item['source']}"
                slide.shapes.add_picture(path, xs[i], top, width=img_w, height=img_h)
                cap = slide.shapes.add_textbox(xs[i], top + img_h + Inches(0.05), img_w, Inches(0.4))
                cp = cap.text_frame.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
                cr = cp.add_run(); cr.text = cap_text
                cr.font.size = Pt(12); cr.font.color.rgb = DARK_GRAY; cr.font.name = "微软雅黑"
        return slide

    orig_count = len(prs.slides)
    new_pages = []
    for item in pages:
        s = add_photo_page(item["title"], item["images"])
        new_pages.append((item["after"], s))

    sldIdLst = prs.slides._sldIdLst
    orig_ids = list(sldIdLst)
    new_ids = list(sldIdLst)[orig_count:]
    after_map = {}
    for (after, s), sid in zip(new_pages, new_ids):
        after_map.setdefault(after, []).append(sid)
    for child in list(sldIdLst):
        sldIdLst.remove(child)
    for i, orig_sid in enumerate(orig_ids, start=1):
        sldIdLst.append(orig_sid)
        if i in after_map:
            for sid in after_map[i]:
                sldIdLst.append(sid)
    prs.save(dst)
    print(f"已生成 {dst}，共 {len(prs.slides)} 页（{orig_count} 内容 + {len(new_pages)} 配图）")
