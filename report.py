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

# 폰트 후보 (배포: fonts-nanum / 로컬 윈도우: malgun / 로컬 맥: AppleSDGothicNeo)
FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
    r"C:\Windows\Fonts\malgun.ttf",
    r"C:\Windows\Fonts\malgunbd.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
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


def _wrap(md, s, font, max_w):
    """max_w(px) 안에 들어가도록 줄을 나눈다.
    공백에서 먼저 끊고, 한 낱말이 그보다 길면 글자 단위로 끊는다."""
    lines = []
    for para in str(s).split("\n"):
        cur = ""
        for word in para.split(" "):
            cand = cur + " " + word if cur else word
            if md.textlength(cand, font=font) <= max_w:
                cur = cand
                continue
            if cur:
                lines.append(cur)
                cur = ""
            while md.textlength(word, font=font) > max_w and len(word) > 1:
                cut = 1
                while cut < len(word) and md.textlength(word[:cut + 1], font=font) <= max_w:
                    cut += 1
                lines.append(word[:cut])
                word = word[cut:]
            cur = word
        lines.append(cur)
    return lines or [""]


STUDENT_COLS = ['출석', '과제수행', '수업내용', '다음과제', '비고']


def render_student_table(name, ws, dates, scale=2):
    """한 학생(name)의 dates별 출석/과제/수업내용/다음과제/비고를 세로 표로 그린다.
    (날짜가 행, 항목이 열). 결석=빨강, 지각·조퇴=파랑 색은 파일 그대로 반영."""
    col = ac.get_roster(ws).get(name)
    if col is None:
        return None
    last_col = ac._last_col(ws, ac._find_start_row(ws))
    fnt = _font(15 * scale)
    tmp = Image.new("RGB", (10, 10)); md = ImageDraw.Draw(tmp)

    def tw(s):
        return md.textlength(str(s), font=fnt)

    pad = 8 * scale
    caps = {'수업내용': 260 * scale, '다음과제': 200 * scale}

    rows = []  # (날짜, {항목: (값, 색)})
    for d in dates:
        top = ac.find_date_block(ws, d)
        if top is None:
            continue
        vals = {}
        if ac._is_holiday_block(ws, top, last_col):
            hol = ws.cell(top, ac.STUDENT_FIRST_COL).value
            for h in STUDENT_COLS:
                vals[h] = ('' if h != '출석' else hol, RED)
        else:
            for h in ('출석', '과제수행', '비고'):
                c = ws.cell(top + ac.ROW_OFFSET[h], col)
                vals[h] = (c.value, _cell_color(c))
            for h in ('수업내용', '다음과제'):
                c = ws.cell(top + ac.ROW_OFFSET[h], ac.STUDENT_FIRST_COL)
                vals[h] = (c.value, BLACK)
        rows.append((d, vals))

    headers = ['날짜'] + STUDENT_COLS
    col_w = {'날짜': int(max(tw('00/00(월)'), tw('날짜'))) + pad * 2}
    for h in STUDENT_COLS:
        w = tw(h)
        for _, vals in rows:
            v = vals[h][0]
            if v not in (None, ''):
                w = max(w, tw(v))
        w = int(w) + pad * 2
        if h in caps:
            w = min(w, caps[h])
        col_w[h] = max(w, 52 * scale)

    row_h = 26 * scale
    asc, desc = fnt.getmetrics()
    line_h = asc + desc + 2 * scale

    # 1차 계산: 각 칸의 줄 나눔과 행 높이. caps에 걸려 잘리는 대신 줄을 접는다.
    laid = []  # [(날짜, {항목: (줄들, 색)}, 행높이)]
    for dd, vals in rows:
        cells, n = {}, 1
        for h in STUDENT_COLS:
            v, c = vals[h]
            lines = _wrap(md, v, fnt, col_w[h] - pad * 2) if v not in (None, '') else []
            n = max(n, len(lines))
            cells[h] = (lines, c)
        laid.append((dd, cells, max(row_h, n * line_h + pad)))

    W = sum(col_w[h] for h in headers)
    H = row_h + sum(r[2] for r in laid)
    img = Image.new("RGB", (W + 1, H + 1), WHITE)
    dr = ImageDraw.Draw(img)

    def put(x, y, w, h, lines, color=BLACK):
        """세로 가운데 정렬로 여러 줄을 그린다."""
        ty = y + (h - len(lines) * line_h) // 2
        for i, ln in enumerate(lines):
            tx = x + (w - dr.textlength(ln, font=fnt)) // 2
            dr.text((tx, ty + i * line_h), ln, font=fnt, fill=color)

    def cell(x, y, w, h, text, bg=WHITE, color=BLACK):
        dr.rectangle([x, y, x + w, y + h], fill=bg, outline=GRID, width=max(1, scale // 2))
        if text not in (None, ''):
            put(x, y, w, h, _wrap(md, text, fnt, w - pad * 2), color)

    x = 0
    for h in headers:
        cell(x, 0, col_w[h], row_h, h, YELLOW); x += col_w[h]
    y = row_h
    for dd, cells, rh in laid:
        x = 0
        cell(x, y, col_w['날짜'], rh, ac.date_label(dd), CREAM); x += col_w['날짜']
        for h in STUDENT_COLS:
            lines, c = cells[h]
            dr.rectangle([x, y, x + col_w[h], y + rh], fill=WHITE, outline=GRID,
                         width=max(1, scale // 2))
            put(x, y, col_w[h], rh, lines, c)
            x += col_w[h]
        y += rh
    return img


def render_class_table(cls_name, ws, dates, scale=2, keep=None):
    """cls_name 반의 dates(예: ['7/7','7/8']) 블록을 표 이미지로 그린다.
    keep: 포함할 학생 이름 집합(None이면 전원). 재적 아닌 학생 제외용.
    내용이 칸보다 길면 칸 안에서 줄바꿈하고, 그만큼 행 높이를 늘린다."""
    roster = [(n, c) for n, c in ac.get_roster(ws).items()
              if keep is None or n in keep]  # [(name, col), ...] 열 순서
    fnt = _font(15 * scale)
    tmp = Image.new("RGB", (10, 10)); md = ImageDraw.Draw(tmp)

    def tw(s):
        return md.textlength(str(s), font=fnt)

    pad = 8 * scale
    asc, desc = fnt.getmetrics()
    line_h = asc + desc + 2 * scale
    row_h = 26 * scale          # 한 줄짜리 행의 기본 높이
    merged_min_w = 360 * scale  # 학생 수가 적어도 수업내용이 과하게 접히지 않도록
    date_w = int(max(tw("00/00(월)"), 40 * scale)) + pad * 2

    def block_label(top, k):
        """블록의 실제 B열 라벨(비고→일일test 등 반영). 없으면 기본 LABELS."""
        v = ws.cell(top + k, 2).value if top else None
        return str(v) if v not in (None, "") else LABELS[k]

    tops = {d: ac.find_date_block(ws, d) for d in dates}

    # 라벨 열 너비: 기본 라벨 + 이번 주 블록들의 실제 라벨 중 최대
    all_labels = list(LABELS)
    for _d, _t in tops.items():
        if _t:
            all_labels += [block_label(_t, k) for k in range(5)]
    label_w = int(max(tw(l) for l in all_labels)) + pad * 2
    ws_last_col = ac._last_col(ws, ac._find_start_row(ws))
    hol_tops = {t for t in tops.values() if t and ac._is_holiday_block(ws, t, ws_last_col)}

    # 학생 열 너비: 모든 학생을 같은 폭으로 둔다. 기준은 이름과 O/X 같은 짧은 값뿐
    # — 비고처럼 긴 값까지 반영하면 그 학생 열만 넓어져 표가 뒤틀린다(긴 값은 줄바꿈).
    # 휴강·공휴일 블록은 제외한다. 그 문구가 출석행 첫 칸에 들어 있어서 같이 재면
    # ('휴강(폭우)' 등) 학생 열이 통째로 그만큼 넓어진다.
    w = max([tw(n) for n, _ in roster] or [0])
    for _d, top in tops.items():
        if top is None or top in hol_tops:
            continue
        for lab in ('출석', '과제수행'):
            for _n, col in roster:
                v = ws.cell(top + ac.ROW_OFFSET[lab], col).value
                if v is not None:
                    w = max(w, tw(v))
    stu_w = [max(int(w) + pad * 2, 55 * scale)] * len(roster)

    # 수업내용·다음과제는 학생 열 전체를 병합해 쓴다. 이 폭이 너무 좁으면 줄이
    # 과하게 접히므로, 모자란 만큼 학생 열에 고루 나눠 넓힌다.
    if stu_w and sum(stu_w) < merged_min_w:
        q, r = divmod(merged_min_w - sum(stu_w), len(stu_w))
        stu_w = [w + q + (1 if i < r else 0) for i, w in enumerate(stu_w)]
    merged_w = sum(stu_w)

    # 1차 계산: 각 칸의 줄 나눔과 행 높이. 전체 이미지 크기를 알아야 그릴 수 있다.
    blocks = []  # [(날짜, [(라벨텍스트, 라벨, 높이, 병합여부, [(폭, 줄들, 색)])], 휴강텍스트)]
    for dt in dates:
        top = tops[dt]
        # 공휴일·휴강 블록은 학생 영역 전체가 한 칸으로 병합돼 있다. 값은 왼쪽 위
        # 칸에만 있으므로 학생별로 읽으면 첫 학생 칸에만 찍힌다 — 따로 처리한다.
        hol = None
        if top in hol_tops:
            hc = ws.cell(top, ac.STUDENT_FIRST_COL)
            hol = (hc.value, _cell_color(hc))
        rows = []
        for k, lab in enumerate(LABELS):
            if hol:
                cells, n, merged = [], 1, False
            elif lab in ('수업내용', '다음과제'):
                v = ws.cell(top + ac.ROW_OFFSET[lab], ac.STUDENT_FIRST_COL).value if top else None
                lines = _wrap(md, v, fnt, merged_w - pad * 2) if v not in (None, '') else []
                cells, n, merged = [(merged_w, lines, BLACK)], len(lines), True
            else:
                cells, n, merged = [], 1, False
                for (name, col), w in zip(roster, stu_w):
                    c = ws.cell(top + ac.ROW_OFFSET[lab], col) if top else None
                    v = c.value if c else None
                    lines = _wrap(md, v, fnt, w - pad * 2) if v not in (None, '') else []
                    n = max(n, len(lines))
                    cells.append((w, lines, _cell_color(c) if c else BLACK))
            rows.append((block_label(top, k), lab, max(row_h, n * line_h + pad), merged, cells))
        blocks.append((dt, rows, hol))

    W = date_w + label_w + merged_w
    H = row_h + sum(sum(r[2] for r in rows) for _, rows, _h in blocks)
    img = Image.new("RGB", (W + 1, H + 1), WHITE)
    d = ImageDraw.Draw(img)

    def box(x, y, w, h, bg, width=None):
        d.rectangle([x, y, x + w, y + h], fill=bg, outline=GRID,
                    width=width or max(1, scale // 2))

    def put(x, y, w, h, lines, color=BLACK, align="center"):
        """세로 가운데 정렬로 여러 줄을 그린다."""
        ty = y + (h - len(lines) * line_h) // 2
        for i, ln in enumerate(lines):
            tx = x + (w - d.textlength(ln, font=fnt)) // 2 if align == "center" else x + pad
            d.text((tx, ty + i * line_h), ln, font=fnt, fill=color)

    def cell(x, y, w, h, text, bg=WHITE, color=BLACK, align="center"):
        box(x, y, w, h, bg)
        if text not in (None, ''):
            put(x, y, w, h, _wrap(md, text, fnt, w - pad * 2), color, align)

    # 헤더
    x = 0
    cell(x, 0, date_w, row_h, "", YELLOW); x += date_w
    cell(x, 0, label_w, row_h, "학생이름", YELLOW); x += label_w
    for (name, _), w in zip(roster, stu_w):
        cell(x, 0, w, row_h, name, YELLOW); x += w

    # 날짜 블록들
    y = row_h
    for dt, rows, hol in blocks:
        bh = sum(r[2] for r in rows)
        box(0, y, date_w, bh, WHITE, width=scale)  # 날짜 셀 (5행 병합)
        put(0, y, date_w, bh, _wrap(md, ac.date_label(dt), fnt, date_w - pad * 2))
        ry = y
        for lab_text, lab, h, merged, cells in rows:
            rowbg = CREAM if lab == '출석' else WHITE
            cell(date_w, ry, label_w, h, lab_text, rowbg)
            x = date_w + label_w
            for w, lines, c in cells:
                box(x, ry, w, h, WHITE if merged else rowbg)
                put(x, ry, w, h, lines, c)
                x += w
            ry += h
        if hol:  # 학생 영역 전체를 한 칸으로 (파일의 병합과 같은 모양)
            hv, hc = hol
            x0 = date_w + label_w
            box(x0, y, merged_w, bh, WHITE)
            put(x0, y, merged_w, bh, _wrap(md, hv, fnt, merged_w - pad * 2), hc)
        y += bh
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
    for r in range(start, ws.max_row + 1):
        k = ac.date_key(ws.cell(r, 1).value)
        if not k:
            continue
        m, dd = map(int, k.split('/'))
        try:
            d = datetime.date(year, m, dd)
        except ValueError:
            continue
        if monday <= d <= sunday:
            out.append(k)
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
