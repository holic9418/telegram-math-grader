# -*- coding: utf-8 -*-
"""주간 출결 보고서 렌더링.
 - 한 반의 특정 주 날짜블록을 표 이미지(PIL)로 그린다.
 - 여러 반 이미지를 모아 PDF로 저장한다.
"""
import os
import datetime
from PIL import Image, ImageDraw, ImageFont
import attendance_core as ac

# 색
YELLOW = (255, 242, 0)
CREAM = (253, 245, 210)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
RED = (220, 0, 0)
BLUE = (0, 0, 210)
GRID = (120, 120, 120)
LABELS = ['출석', '수업내용', '과제수행', '다음과제', '비고']

# 폰트 후보 (배포: fonts-nanum / 로컬 윈도우: malgun)
FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
    r"C:\Windows\Fonts\malgun.ttf",
    r"C:\Windows\Fonts\malgunbd.ttf",
]


def _font(size):
    import glob
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


def _cell_color(cell):
    try:
        rgb = cell.font.color.rgb
        if isinstance(rgb, str):
            h = rgb[-6:].upper()
            if h == 'FF0000':
                return RED
            if h == '0000FF':
                return BLUE
    except Exception:
        pass
    return BLACK


def render_class_table(cls_name, ws, dates, scale=2, keep=None):
    """cls_name 반의 dates(예: ['7/7','7/8']) 블록을 표 이미지로 그린다.
    keep: 포함할 학생 이름 집합(None이면 전원). 재적 아닌 학생 제외용."""
    roster = [(n, c) for n, c in ac.get_roster(ws).items()
              if keep is None or n in keep]  # [(name, col), ...] 열 순서
    fnt = _font(15 * scale)
    fnt_b = _font(15 * scale)
    tmp = Image.new("RGB", (10, 10)); md = ImageDraw.Draw(tmp)

    def tw(s):
        return md.textlength(str(s), font=fnt)

    pad = 8 * scale
    date_w = int(max(tw("00/00"), 40 * scale)) + pad * 2

    def block_label(top, k):
        """블록의 실제 B열 라벨(비고→일일test 등 반영). 없으면 기본 LABELS."""
        v = ws.cell(top + k, 2).value if top else None
        return str(v) if v not in (None, "") else LABELS[k]

    # 라벨 열 너비: 기본 라벨 + 이번 주 블록들의 실제 라벨 중 최대
    all_labels = list(LABELS)
    for _d in dates:
        _t = ac.find_date_block(ws, _d)
        if _t:
            all_labels += [block_label(_t, k) for k in range(5)]
    label_w = int(max(tw(l) for l in all_labels)) + pad * 2
    # 학생 열 너비: 이름/값 중 최대 (수업내용·다음과제 제외 = 병합)
    stu_w = []
    for name, col in roster:
        w = tw(name)
        for d in dates:
            top = ac.find_date_block(ws, d)
            if top is None:
                continue
            for lab in ('출석', '과제수행', '비고'):
                v = ws.cell(top + ac.ROW_OFFSET[lab], col).value
                if v is not None:
                    w = max(w, tw(v))
        stu_w.append(min(max(int(w) + pad * 2, 55 * scale), 240 * scale))

    row_h = 26 * scale
    W = date_w + label_w + sum(stu_w)
    H = row_h * (1 + 5 * len(dates))
    img = Image.new("RGB", (W + 1, H + 1), WHITE)
    d = ImageDraw.Draw(img)

    def cell(x, y, w, h, text, bg=WHITE, color=BLACK, align="center"):
        d.rectangle([x, y, x + w, y + h], fill=bg, outline=GRID, width=max(1, scale // 2))
        if text is None or str(text) == "":
            return
        s = str(text)
        bbox = d.textbbox((0, 0), s, font=fnt)
        th = bbox[3] - bbox[1]
        ty = y + (h - th) // 2 - bbox[1]
        if align == "center":
            tx = x + (w - d.textlength(s, font=fnt)) // 2
        else:
            tx = x + pad
        d.text((tx, ty), s, font=fnt, fill=color)

    # 헤더
    x = 0
    cell(x, 0, date_w, row_h, "", YELLOW); x += date_w
    cell(x, 0, label_w, row_h, "학생이름", YELLOW); x += label_w
    for (name, _), w in zip(roster, stu_w):
        cell(x, 0, w, row_h, name, YELLOW); x += w

    # 날짜 블록들
    y = row_h
    for di, dt in enumerate(dates):
        top = ac.find_date_block(ws, dt)
        # 날짜 셀 (5행 병합)
        d.rectangle([0, y, date_w, y + row_h * 5], fill=WHITE, outline=GRID, width=scale)
        db = d.textbbox((0, 0), dt, font=fnt)
        d.text((date_w // 2 - d.textlength(dt, font=fnt) // 2,
                y + (row_h * 5) // 2 - (db[3] - db[1]) // 2 - db[1]), dt, font=fnt, fill=BLACK)
        for k, lab in enumerate(LABELS):
            ry = y + k * row_h
            rowbg = CREAM if lab == '출석' else WHITE
            cell(date_w, ry, label_w, row_h, block_label(top, k), rowbg)
            if lab in ('수업내용', '다음과제'):
                # 학생 영역 전체에 하나로
                val = ws.cell(top + ac.ROW_OFFSET[lab], ac.STUDENT_FIRST_COL).value if top else None
                x0 = date_w + label_w
                cell(x0, ry, sum(stu_w), row_h, val, WHITE)
            else:
                x = date_w + label_w
                for (name, col), w in zip(roster, stu_w):
                    v = ws.cell(top + ac.ROW_OFFSET[lab], col).value if top else None
                    col_c = _cell_color(ws.cell(top + ac.ROW_OFFSET[lab], col)) if top else BLACK
                    cell(x, ry, w, row_h, v, rowbg, color=col_c)
                    x += w
        y += row_h * 5
    return img


def render_class_page(cls_name, table_img, title=None, scale=2):
    """반 이름(+선택적 제목) 헤딩 아래 표를 얹은 한 페이지 이미지."""
    margin = 24 * scale
    title_h = 50 * scale if title else 0
    head_h = 40 * scale
    W = max(table_img.width, 500 * scale) + margin * 2
    H = title_h + head_h + table_img.height + margin * 2
    page = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(page)
    y = margin
    if title:
        d.text((margin, y), title, font=_font(28 * scale), fill=BLACK)
        y += title_h
    d.text((margin, y), cls_name, font=_font(23 * scale), fill=BLACK)
    y += head_h
    page.paste(table_img, (margin, y))
    return page


def week_bounds(today):
    """today가 속한 주의 월요일~일요일 (date, date)."""
    monday = today - datetime.timedelta(days=today.weekday())
    return monday, monday + datetime.timedelta(days=6)


def week_dates(ws, monday, sunday, year):
    """ws의 날짜블록 중 [monday, sunday] 안에 드는 'M/D' 목록."""
    out = []
    start = ac._find_start_row(ws)
    if start is None:
        return out
    import re
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and re.match(r'^\d{1,2}/\d{1,2}$', str(v).strip()):
            ds = str(v).strip()
            m, dd = map(int, ds.split('/'))
            try:
                d = datetime.date(year, m, dd)
            except ValueError:
                continue
            if monday <= d <= sunday:
                out.append(ds)
    return out


def build_report_pdf(wb, out_path, monday, sunday, year,
                     title="<수학과 주간 출결사항>", enroll=None):
    """이번 주 각 반 표를 모아 PDF로 저장. 반환: 포함된 반 수(0이면 미생성).
    enroll: {반: {학생: {'from','to'}}} — 그 주에 재적 아닌 학생은 표에서 제외."""
    enroll = enroll or {}
    pages = []
    for cls in wb.sheetnames:
        dates = week_dates(wb[cls], monday, sunday, year)
        if not dates:
            continue
        cls_enroll = enroll.get(cls, {})
        keep = {n for n in ac.get_roster(wb[cls])
                if ac.is_enrolled_during(cls_enroll.get(n), monday, sunday, year)}
        tbl = render_class_table(cls, wb[cls], dates, keep=keep)
        head_title = f"{title}   ({monday.month}/{monday.day}~{sunday.month}/{sunday.day})" if not pages else None
        pages.append(render_class_page(cls, tbl, head_title))
    if not pages:
        return 0
    pages[0].save(out_path, "PDF", save_all=True, append_images=pages[1:], resolution=150.0)
    return len(pages)
