#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用改色：遍历全部 theme*.xml 的 clrScheme 按位置替换 + slides/charts/layouts/masters 残留 srgbClr 替换。

用法:
    python3 recolor_theme.py <src.pptx> <dst.pptx> [palette.json]

palette.json 可选，格式:
    {"theme": {"accent1":"1F3A6E", ...}, "slide_map": {"E0572D":"2E5C9E", ...}}
不传则使用默认蓝色稳重系映射。
"""
import zipfile, re, json, sys

DEFAULT_TARGET = {
    "dk2": "1F3A6E", "lt2": "E8EFF8",
    "accent1": "1F3A6E", "accent2": "2E5C9E", "accent3": "4A7BBF",
    "accent4": "5B9BD5", "accent5": "7AADD6", "accent6": "90AEE0",
    "hlink": "2E5C9E",
}
DEFAULT_SLIDE_MAP = {
    "C00000": "1F3A6E", "D02300": "1F3A6E", "E0572D": "2E5C9E", "D18654": "3A6EA5",
    "F2B044": "4A7BBF", "F77365": "3A6EA5", "FE9987": "7AADD6",
    "FBD4AD": "D6E4F5", "FFD1AB": "D6E4F5", "EFB7B9": "C9DCEF",
    "D7A4A6": "B8CDE8", "F1C8A8": "DCE8F5",
    "FFC000": "2E5C9E", "FFFF00": "1F3A6E", "FFDC6D": "2E5C9E",
    "614FB3": "2E5C9E", "3E327A": "1F3A6E", "A796FF": "5B9BD5",
    "7A74B7": "3A6EA5", "5B3CAF": "2A4D8F", "463AA0": "243F7A",
    "AAA7D1": "8FB3D9", "C2C1DB": "C5D6EC", "F8F4FA": "EEF3FA",
    "AD84C6": "1F3A6E", "8784C7": "2E5C9E", "5D739A": "4A7BBF",
    "6997AF": "5B9BD5", "84ACB6": "7AADD6", "6F8183": "90AEE0",
}

def fix_theme(text, target):
    def repl(m):
        tag = m.group(1)
        if tag in target:
            return f'<a:{tag}><a:srgbClr val="{target[tag]}"/>'
        return m.group(0)
    return re.sub(r'<a:(dk2|lt2|accent\d|hlink)><a:srgbClr val="[0-9A-Fa-f]{6}"', repl, text)

def map_slide(text, slide_map):
    for old, new in slide_map.items():
        text = re.sub(
            r'(srgbClr\s+val=")(‘ + old + r')(")',
            lambda m: m.group(1) + new + m.group(3),
            text, flags=re.IGNORECASE
        )
    return text

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    target, slide_map = DEFAULT_TARGET, DEFAULT_SLIDE_MAP
    if len(sys.argv) >= 4:
        pal = json.load(open(sys.argv[3]))
        target = pal.get("theme", DEFAULT_TARGET)
        slide_map = pal.get("slide_map", DEFAULT_SLIDE_MAP)
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name.startswith("ppt/theme/") and name.endswith(".xml"):
                    data = fix_theme(data.decode("utf-8"), target).encode("utf-8")
                elif (name.startswith("ppt/slides/") or name.startswith("ppt/charts/")
                      or name.startswith("ppt/slideLayouts/") or name.startswith("ppt/slideMasters/")) and name.endswith(".xml"):
                    data = map_slide(data.decode("utf-8"), slide_map).encode("utf-8")
                zout.writestr(name, data)
    # 校验所有 theme 的 accent1
    z = zipfile.ZipFile(dst)
    for t in sorted([n for n in z.namelist() if n.startswith("ppt/theme/theme")]):
        x = z.read(t).decode("utf-8", errors="ignore")
        a1 = re.search(r'<a:accent1><a:srgbClr val="([0-9A-Fa-f]{6})"', x)
        print(f"  {t.split('/')[-1]} accent1 = {a1.group(1) if a1 else '?'}")
    print(f"已生成 {dst}")
