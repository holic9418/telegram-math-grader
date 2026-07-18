# -*- coding: utf-8 -*-
"""진도표 렌더링. 단원 × (유형서/심화유형서/단원평가) × (수정완료/밴드완료) 칸에
학생별 O/X를 채워 표 이미지(PIL)로 그린다. 사진처럼 단원 몇 개씩 '밴드'로 접는다.
"""
import os
import glob
from PIL import Image, ImageDraw, ImageFont

WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
RED = (220, 0, 0)
GRID = (150, 150, 150)
HEAD_BG = (238, 238, 238)
NAME_BG = (247, 247, 247)

# 기본 항목/단계 (반별로 바꿀 수 있게 저장하되, 기본은 사진과 동일)
DEFAULT_ITEMS = ["유형서", "심화유형서", "단원평가"]
DEFAULT_STEPS = ["수정완료", "밴드완료"]
UNITS_PER_BAND = 3   # 사진처럼 단원 3개씩 한 줄(밴드)로

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
    r"C:\Windows\Fonts\malgun.ttf",
    r"C:\Windows\Fonts\malgunbd.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
]


def _font(size):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    for pat in ("/usr/share/fonts/**/Nanum*.ttf", "/usr/share/fonts/**/*.ttf"):
        for p in glob.glob(pat, recursive=True):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _text_w(md, s, font):
    return md.textbbox((0, 0), s, font=font)[2]


def render_progress(cls_name, units, students, cells,
                    items=None, steps=None, scale=2):
    """cls_name: 반 이름 / units: 단원명 리스트 / students: 학생명 리스트
    cells: {학생: {'<ui>|<item>|<step>': 'O'|'X'}}  (ui = 단원 인덱스)
    """
    items = items or DEFAULT_ITEMS
    steps = steps or DEFAULT_STEPS
    if not units:
        units = ["(단원 미설정)"]
    ncol = len(items) * len(steps)   # 한 단원의 칸 수

    f_title = _font(15 * scale)
    f_unit = _font(12 * scale)
    f_item = _font(11 * scale)
    f_step = _font(9 * scale)
    f_name = _font(12 * scale)
    f_mark = _font(15 * scale)

    tmp = Image.new("RGB", (10, 10)); md = ImageDraw.Draw(tmp)

    # 열 너비: 이름칸 + 각 단계칸(‘수정완료’ 들어갈 만큼)
    name_w = max([_text_w(md, s, f_name) for s in students] + [_text_w(md, "학생", f_name)]) + 16 * scale
    step_w = max([_text_w(md, s, f_step) for s in steps]) + 10 * scale
    unit_w = step_w * ncol           # 단원 한 칸 전체 너비

    row_h = 20 * scale               # 학생 행 높이
    h_unit = 20 * scale              # 헤더: 단원
    h_item = 17 * scale              # 헤더: 항목
    h_step = 16 * scale              # 헤더: 단계
    head_h = h_unit + h_item + h_step
    title_h = 26 * scale
    band_gap = 14 * scale

    # 밴드로 단원 나누기
    bands = [units[i:i + UNITS_PER_BAND] for i in range(0, len(units), UNITS_PER_BAND)]
    band_w = name_w + unit_w * min(UNITS_PER_BAND, len(units)) if units else name_w
    # 각 밴드의 실제 폭(마지막 밴드는 단원이 적을 수 있음)
    def band_width(nunits):
        return name_w + unit_w * nunits
    max_w = max(band_width(len(b)) for b in bands)
    band_h = head_h + row_h * len(students)
    total_h = title_h + (band_h + band_gap) * len(bands)

    img = Image.new("RGB", (int(max_w) + 1, int(total_h) + 1), WHITE)
    d = ImageDraw.Draw(img)

    # 제목
    d.text((6 * scale, 4 * scale), f"{cls_name} 진도표", font=f_title, fill=BLACK)

    def ctext(cx, cy, s, font, fill=BLACK):
        w = _text_w(md, s, font)
        d.text((cx - w / 2, cy), s, font=font, fill=fill)

    y0 = title_h
    unit_base = 0
    for b, bunits in enumerate(bands):
        top = y0 + b * (band_h + band_gap)
        nB = len(bunits)
        # 헤더 배경
        d.rectangle([0, top, name_w + unit_w * nB, top + head_h], fill=HEAD_BG, outline=GRID)
        # 이름칸(헤더 영역 병합 느낌)
        d.rectangle([0, top, name_w, top + head_h], fill=NAME_BG, outline=GRID)
        ctext(name_w / 2, top + head_h / 2 - 8 * scale, "학생", f_item)
        for u in range(nB):
            ux = name_w + unit_w * u
            # 단원 헤더
            d.rectangle([ux, top, ux + unit_w, top + h_unit], outline=GRID)
            ctext(ux + unit_w / 2, top + 4 * scale, f"{unit_base+u+1}. {bunits[u]}", f_unit)
            # 항목 헤더
            for ii, item in enumerate(items):
                ix = ux + step_w * len(steps) * ii
                d.rectangle([ix, top + h_unit, ix + step_w * len(steps), top + h_unit + h_item], outline=GRID)
                ctext(ix + step_w * len(steps) / 2, top + h_unit + 3 * scale, item, f_item)
                # 단계 헤더
                for si, stp in enumerate(steps):
                    sx = ix + step_w * si
                    d.rectangle([sx, top + h_unit + h_item, sx + step_w, top + head_h], outline=GRID)
                    ctext(sx + step_w / 2, top + h_unit + h_item + 2 * scale, stp, f_step)
        # 학생 행
        for ri, name in enumerate(students):
            ry = top + head_h + row_h * ri
            d.rectangle([0, ry, name_w, ry + row_h], fill=NAME_BG, outline=GRID)
            d.text((8 * scale, ry + (row_h - 12 * scale) / 2), name, font=f_name, fill=BLACK)
            for u in range(nB):
                ui = unit_base + u
                for ii, item in enumerate(items):
                    for si, stp in enumerate(steps):
                        cx = name_w + unit_w * u + step_w * (len(steps) * ii + si)
                        d.rectangle([cx, ry, cx + step_w, ry + row_h], outline=GRID)
                        v = (cells.get(name) or {}).get(f"{ui}|{item}|{stp}")
                        if v:
                            col = RED if v.upper() == "X" else BLACK
                            ctext(cx + step_w / 2, ry + (row_h - 15 * scale) / 2, v.upper(), f_mark, fill=col)
        unit_base += nB

    return img
