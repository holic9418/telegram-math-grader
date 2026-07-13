# -*- coding: utf-8 -*-
"""
월간 출석부 핵심 모듈.
 - 한컴(한셀) xlsx 호환 로드/저장
 - 반별 요일패턴으로 새 달 출석부 생성 (학생/서식 유지, 공휴일 병합+빨강)
 - 시트 → HTML 미리보기 렌더
"""
import io, re, zipfile, datetime, calendar
from copy import copy
import openpyxl
from openpyxl.styles import Font, Alignment, Border
import holidays as _holidays

WEEKDAY_KR = ['월', '화', '수', '목', '금', '토', '일']
KR_WD_TO_IDX = {k: i for i, k in enumerate(WEEKDAY_KR)}

# 국경일이지만 실제 휴무가 아닌 날 (수업 정상 진행) → 공휴일에서 제외
NON_OFF_HOLIDAYS = {'제헌절'}


# ── 1. 한컴 호환 로드/저장 ──────────────────────────────────────
def _normalize_hancom_bytes(data: bytes) -> bytes:
    """styles.xml 의 mc:AlternateContent 를 Fallback 내용으로 치환해 openpyxl 이 읽게 함."""
    zin = zipfile.ZipFile(io.BytesIO(data))
    styles = zin.read('xl/styles.xml').decode('utf-8')

    def repl(m):
        fb = re.search(r'<mc:Fallback>(.*?)</mc:Fallback>', m.group(0), re.S)
        return fb.group(1) if fb else ''

    styles2 = re.sub(r'<mc:AlternateContent\b.*?</mc:AlternateContent>', repl,
                     styles, flags=re.S)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            d = zin.read(item.filename)
            if item.filename == 'xl/styles.xml':
                d = styles2.encode('utf-8')
            zout.writestr(item, d)
    zin.close()
    return out.getvalue()


def load_workbook(path_or_bytes):
    """한컴 xlsx 를 openpyxl Workbook 으로 로드 (경로 또는 bytes)."""
    if isinstance(path_or_bytes, (bytes, bytearray)):
        data = bytes(path_or_bytes)
    else:
        with open(path_or_bytes, 'rb') as f:
            data = f.read()
    try:
        return openpyxl.load_workbook(io.BytesIO(data))
    except Exception:
        return openpyxl.load_workbook(io.BytesIO(_normalize_hancom_bytes(data)))


# ── 2. 시트 구조 분석 ──────────────────────────────────────────
BLOCK_LABELS = ['출석', '수업내용', '과제수행', '다음과제', '비고']
BLOCK_SIZE = len(BLOCK_LABELS)


def _find_start_row(ws):
    """col A 에 날짜가 처음 등장하는 행 = 첫 블록 시작행."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and re.match(r'^\d{1,2}/\d{1,2}$', str(v).strip()):
            return r
    return None


def _last_col(ws, start):
    """헤더+첫 블록에서 실제 사용된 마지막 열."""
    last = 2
    for r in range(1, start + BLOCK_SIZE):
        for c in range(1, 60):
            if ws.cell(r, c).value not in (None, ''):
                last = max(last, c)
    for rng in ws.merged_cells.ranges:
        if rng.min_row < start + BLOCK_SIZE:
            last = max(last, rng.max_col)
    return last


def infer_schedule(ws):
    """기존 시트의 날짜들로부터 수업 요일 집합 추론."""
    start = _find_start_row(ws)
    days = set()
    year = datetime.date.today().year
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and re.match(r'^\d{1,2}/\d{1,2}$', str(v).strip()):
            mm, dd = str(v).split('/')
            try:
                days.add(datetime.date(year, int(mm), int(dd)).weekday())
            except ValueError:
                pass
    return sorted(days)


# ── 3. 공휴일 ──────────────────────────────────────────────────
def holidays_for(year):
    kr = _holidays.SouthKorea(years=[year])
    return {d: name for d, name in kr.items() if name not in NON_OFF_HOLIDAYS}


def class_days(year, month, weekday_idxs):
    n = calendar.monthrange(year, month)[1]
    out = []
    for day in range(1, n + 1):
        d = datetime.date(year, month, day)
        if d.weekday() in weekday_idxs:
            out.append(d)
    return out


# ── 4. 새 달 생성 ──────────────────────────────────────────────
STUDENT_FIRST_COL = 3   # A=날짜, B=라벨, C부터 학생/CLASS
# 반별 한 칸으로 병합할 라벨(상대행) — 수업내용, 다음과제
MERGE_LABEL_ROWS = [1, 3]


def _detect_groups(ws, start, last_col):
    """헤더의 CLASS 병합으로 학생 열 그룹 감지. 없으면 전체를 한 그룹."""
    groups = []
    for rng in ws.merged_cells.ranges:
        if rng.max_row < start and rng.min_col >= STUDENT_FIRST_COL and rng.max_col > rng.min_col:
            groups.append((rng.min_col, min(rng.max_col, last_col)))
    if groups:
        groups.sort()
        return groups
    return [(STUDENT_FIRST_COL, last_col)]


def _capture_template(ws, start, last_col):
    """첫 블록의 셀 스타일/행높이만 캡처 (병합은 재구성)."""
    styles, heights = [], []
    for i in range(BLOCK_SIZE):
        r = start + i
        styles.append({c: copy(ws.cell(r, c)._style) for c in range(1, last_col + 1)})
        heights.append(ws.row_dimensions[r].height)
    label_font = copy(ws.cell(start, 2).font)
    groups = _detect_groups(ws, start, last_col)
    return {'styles': styles, 'heights': heights, 'groups': groups,
            'last_col': last_col, 'label_font': label_font}


def generate_month(src_workbook, year, month, schedules=None):
    """src_workbook(Workbook) 을 바탕으로 새 달 Workbook 을 만들어 반환.
       schedules: {sheet_name: [weekday_idx,...]} 없으면 기존 파일에서 추론."""
    wb = src_workbook
    hol = holidays_for(year)
    schedules = schedules or {}

    for name in wb.sheetnames:
        ws = wb[name]
        start = _find_start_row(ws)
        if start is None:
            continue
        last_col = _last_col(ws, start)
        tpl = _capture_template(ws, start, last_col)

        wds = schedules.get(name) or infer_schedule(ws)
        days = class_days(year, month, wds)

        # 1) 블록 영역(>= start)의 기존 병합 모두 해제
        for rng in list(ws.merged_cells.ranges):
            if rng.max_row >= start:
                ws.unmerge_cells(str(rng))

        # 2) 블록 영역 값 비우기 (넉넉히)
        max_clear = start + BLOCK_SIZE * (len(days) + 4)
        for r in range(start, max(max_clear, ws.max_row) + 1):
            for c in range(1, last_col + 1):
                ws.cell(r, c).value = None

        # 3) 새 블록 기록
        for i, d in enumerate(days):
            top = start + BLOCK_SIZE * i
            # 스타일/행높이 적용 (지난달 결석 표시인 대각선은 제거)
            for k in range(BLOCK_SIZE):
                r = top + k
                for c in range(1, last_col + 1):
                    ws.cell(r, c)._style = copy(tpl['styles'][k][c])
                    _strip_diagonal(ws.cell(r, c))
                if tpl['heights'][k] is not None:
                    ws.row_dimensions[r].height = tpl['heights'][k]
                ws.cell(r, 2).value = BLOCK_LABELS[k]
            # 날짜 (A열, 세로 병합)
            ws.cell(top, 1).value = f'{d.month}/{d.day}'
            ws.merge_cells(start_row=top, start_column=1, end_row=top + BLOCK_SIZE - 1, end_column=1)

            if d in hol:
                # 공휴일: 입력칸(C..last) 5행 병합 후 이름 빨강
                ws.merge_cells(start_row=top, start_column=3,
                               end_row=top + BLOCK_SIZE - 1, end_column=last_col)
                cell = ws.cell(top, 3)
                cell.value = hol[d]
                base = tpl['label_font']
                cell.font = Font(name=base.name, size=base.sz, bold=True, color='FFFF0000')
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                # 평일: 수업내용·다음과제 행을 반(CLASS)별 한 칸으로 병합
                for rel in MERGE_LABEL_ROWS:
                    for (g0, g1) in tpl['groups']:
                        if g1 > g0:
                            ws.merge_cells(start_row=top + rel, start_column=g0,
                                           end_row=top + rel, end_column=g1)

        # 4) 마지막 블록 아래 잔여 영역 비우기 (지난달이 더 길었던 경우)
        last_end = start + BLOCK_SIZE * len(days) - 1
        for r in range(last_end + 1, ws.max_row + 1):
            ws.row_dimensions[r].height = None
            for c in range(1, last_col + 1):
                cell = ws.cell(r, c)
                cell.value = None
                cell.border = Border()
    return wb


# ── 4b. 출석 기록/조회 ─────────────────────────────────────────
def get_roster(ws):
    """{학생이름: 열번호} 반환 (헤더의 학생이름 행 기준)."""
    start = _find_start_row(ws)
    if start is None:
        return {}
    last_col = _last_col(ws, start)
    header_row = start - 1
    roster = {}
    for c in range(STUDENT_FIRST_COL, last_col + 1):
        v = ws.cell(header_row, c).value
        if v and str(v).strip() and str(v).strip() != '학생이름':
            roster[str(v).strip()] = c
    return roster


def find_date_block(ws, date_str):
    """'8/4' 같은 날짜의 블록 시작행 반환 (없으면 None)."""
    start = _find_start_row(ws)
    target = str(date_str).strip()
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip() == target:
            return r
    return None


def rosters_summary(wb):
    """{sheet: [학생이름,...]} — 봇 파싱용."""
    out = {}
    for name in wb.sheetnames:
        out[name] = list(get_roster(wb[name]).keys())
    return out


ROW_OFFSET = {'출석': 0, '수업내용': 1, '과제수행': 2, '다음과제': 3, '비고': 4}

# 반별 수업 요일 기본값 (요일 인덱스: 월0~일6)
DEFAULT_SCHEDULES = {
    '초5': [1, 2, 3], '초6': [0, 2, 4], '중1': [0, 2, 4], '중2': [1, 2, 3],
    '중3': [0, 2, 4], '고1': [1, 3, 5], '고2(미적분)': [1, 3, 5], '고3': [0, 2, 4],
}


def sheet_dates(wb):
    """{sheet: ['8/4', ...]} — 각 시트의 수업일."""
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        start = _find_start_row(ws)
        ds = []
        if start:
            for r in range(start, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if v and re.match(r'^\d{1,2}/\d{1,2}$', str(v).strip()):
                    ds.append(str(v).strip())
        out[name] = ds
    return out


def next_year_month(year, month):
    return (year + 1, 1) if month == 12 else (year, month + 1)


def _strip_diagonal(cell):
    """셀의 대각선(사선) 테두리만 제거하고 나머지 테두리는 유지."""
    b = cell.border
    if (b.diagonal and getattr(b.diagonal, 'style', None)) or b.diagonalUp or b.diagonalDown:
        cell.border = Border(left=b.left, right=b.right, top=b.top, bottom=b.bottom)


def _is_holiday_block(ws, top, last_col):
    """블록에 사각형(여러행×여러열) 병합이 있으면 공휴일/휴강 블록으로 간주."""
    for rng in ws.merged_cells.ranges:
        if (rng.min_row >= top and rng.max_row <= top + BLOCK_SIZE - 1
                and rng.min_col >= STUDENT_FIRST_COL
                and rng.max_row > rng.min_row and rng.max_col > rng.min_col):
            return True
    return False


def normalize_block(ws, top, last_col, groups):
    """블록을 '전원 정상출석' 기준 깨끗한 서식으로 정리한다.
       - 학생 세로병합(결석 표시) 해제, 병합됐던 칸 테두리를 정상 칸 기준으로 통일
       - 대각선(사선) 제거
       - 수업내용/다음과제 행을 반(그룹)별 가로 한 칸으로 재병합
       공휴일/휴강 블록(사각형 병합)은 건드리지 않는다."""
    if _is_holiday_block(ws, top, last_col):
        return
    vcols = set()
    for rng in list(ws.merged_cells.ranges):
        if (rng.min_row >= top and rng.max_row <= top + BLOCK_SIZE - 1
                and rng.min_col >= STUDENT_FIRST_COL):
            if rng.max_row > rng.min_row and rng.min_col == rng.max_col:
                vcols.add(rng.min_col)
            ws.unmerge_cells(str(rng))
    # 병합 해제된 학생열을 같은 행의 정상 참조열 스타일로 통일
    for r in range(top, top + BLOCK_SIZE):
        ref = next((c for c in range(STUDENT_FIRST_COL, last_col + 1) if c not in vcols), None)
        if ref is not None:
            for c in vcols:
                ws.cell(r, c)._style = copy(ws.cell(r, ref)._style)
    # 남은 대각선 제거
    for r in range(top, top + BLOCK_SIZE):
        for c in range(STUDENT_FIRST_COL, last_col + 1):
            _strip_diagonal(ws.cell(r, c))
    # 수업내용/다음과제 가로 병합 복원
    for rel in MERGE_LABEL_ROWS:
        for (g0, g1) in groups:
            if g1 > g0:
                ws.merge_cells(start_row=top + rel, start_column=g0,
                               end_row=top + rel, end_column=g1)


COLOR_RED = 'FFFF0000'
COLOR_BLUE = 'FF0000FF'
COLOR_BLACK = 'FF000000'


def _attendance_color(value):
    """출석 값 → 글자색. 결석=빨강, 지각·조퇴=파랑, 정상=검정."""
    s = str(value)
    if '지각' in s or '조퇴' in s:
        return COLOR_BLUE
    if '결석' in s or '결' == s.strip() or 'X' in s.upper():
        return COLOR_RED
    return COLOR_BLACK


def _homework_color(value):
    """과제수행 값 → 글자색. 안함=빨강, 50%/절반=파랑, 완료=검정."""
    s = str(value).strip()
    if any(k in s for k in ('50', '％', '%', '△', '절반')) or s == '반':
        return COLOR_BLUE
    if any(k in s for k in ('X', 'x', '안', '미', '못', '노')):
        return COLOR_RED
    return COLOR_BLACK


def _apply_font_color(cell, argb):
    f = cell.font
    cell.font = Font(name=f.name, size=f.sz, bold=f.bold, italic=f.italic,
                     underline=f.underline, color=argb)


def write_attendance(wb, sheet, date_str, data):
    """data 예:
       {'출석': {'김규림':'O','남우현':'X(결석)'},
        '수업내용': '분수의 나눗셈',
        '과제수행': {'김규림':'O'},
        '다음과제': '42쪽',
        '비고': {'남우현':'보강필요'} 또는 '단체공지'}
       반환: (기록 목록, 경고 목록)"""
    if sheet not in wb.sheetnames:
        raise KeyError(f'시트 없음: {sheet}')
    ws = wb[sheet]
    top = find_date_block(ws, date_str)
    if top is None:
        raise ValueError(f"{sheet}에 '{date_str}' 날짜가 없습니다")
    # 기록 전에 이 블록의 서식을 깨끗하게 정리(결석 세로병합·대각선 제거 등)
    start = _find_start_row(ws)
    last_col = _last_col(ws, start)
    groups = _detect_groups(ws, start, last_col)
    normalize_block(ws, top, last_col, groups)
    roster = get_roster(ws)
    written, warnings = [], []

    def put(label, student, value):
        r = top + ROW_OFFSET[label]
        if student not in roster:
            warnings.append(f"{label}: '{student}' 학생을 못 찾음")
            return
        cell = ws.cell(r, roster[student])
        cell.value = value
        if label == '출석':
            _apply_font_color(cell, _attendance_color(value))
        elif label == '과제수행':
            _apply_font_color(cell, _homework_color(value))
        written.append(f"{label} · {student} = {value}")

    for label in ('출석', '과제수행', '비고'):
        block = data.get(label)
        if isinstance(block, dict):
            for st, val in block.items():
                put(label, st, val)
        elif isinstance(block, str) and block.strip():
            # 단체 문자열 → 라벨행 첫 학생칸(병합 앵커 아님)에 기록
            r = top + ROW_OFFSET[label]
            ws.cell(r, STUDENT_FIRST_COL).value = block
            written.append(f"{label} = {block}")

    for label in ('수업내용', '다음과제'):
        val = data.get(label)
        if isinstance(val, str) and val.strip():
            r = top + ROW_OFFSET[label]
            ws.cell(r, STUDENT_FIRST_COL).value = val  # 병합 앵커
            written.append(f"{label} = {val}")

    return written, warnings


# ── 5. HTML 미리보기 ───────────────────────────────────────────
def _argb_to_css(argb):
    if not argb or not isinstance(argb, str):
        return None
    s = argb[-6:] if len(argb) >= 6 else argb
    if s.upper() in ('000000',):
        return '#000'
    return '#' + s


def render_html(ws, max_rows=46):
    start = _find_start_row(ws)
    last_col = _last_col(ws, start)
    span = {}
    skip = set()
    for rng in ws.merged_cells.ranges:
        span[(rng.min_row, rng.min_col)] = (rng.max_row - rng.min_row + 1,
                                            rng.max_col - rng.min_col + 1)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    skip.add((r, c))
    rows_html = []
    for r in range(1, min(ws.max_row, max_rows) + 1):
        cells = []
        for c in range(1, last_col + 1):
            if (r, c) in skip:
                continue
            cell = ws.cell(r, c)
            v = cell.value if cell.value is not None else ''
            rs, cs = span.get((r, c), (1, 1))
            cls = []
            try:
                col = cell.font.color
                if col and col.rgb and isinstance(col.rgb, str) and col.rgb.lower().endswith('ff0000'):
                    cls.append('h')          # holiday red
            except Exception:
                pass
            if c == 2:
                cls.append('lbl')            # label column
            attr = f' rowspan="{rs}"' if rs > 1 else ''
            attr += f' colspan="{cs}"' if cs > 1 else ''
            attr += f' class="{" ".join(cls)}"' if cls else ''
            cells.append(f'<td{attr}>{v}</td>')
        rows_html.append('<tr>' + ''.join(cells) + '</tr>')
    return '<table class="att">' + ''.join(rows_html) + '</table>'
