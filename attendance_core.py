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


def attendance_recorded(ws, date_str):
    """해당 날짜의 출석 행에 기록이 있는지 확인한다.
    반환: True  = 이미 출석 입력됨
          False = 그 날짜 블록은 있으나 출석이 비어 있음(입력 필요)
          None  = 해당 날짜 블록 없음 / 공휴일·휴강 블록(확인 불가·불필요)."""
    top = find_date_block(ws, date_str)
    if top is None:
        return None
    start = _find_start_row(ws)
    last_col = _last_col(ws, start)
    if _is_holiday_block(ws, top, last_col):
        return None
    att_row = top + ROW_OFFSET['출석']
    for c in get_roster(ws).values():
        if ws.cell(att_row, c).value not in (None, ''):
            return True
    return False


def rosters_summary(wb):
    """{sheet: [학생이름,...]} — 봇 파싱용."""
    out = {}
    for name in wb.sheetnames:
        out[name] = list(get_roster(wb[name]).keys())
    return out


ROW_OFFSET = {'출석': 0, '수업내용': 1, '과제수행': 2, '다음과제': 3, '비고': 4}

# 반별 수업 요일 기본값 (요일 인덱스: 월0~일6)
DEFAULT_SCHEDULES = {
    '초3': [1, 2, 3], '초4': [0, 2, 4], '초5': [1, 2, 3], '초5A': [1, 2, 3],
    '초6': [0, 2, 4], '초6A': [0, 2, 4], '중1AB': [0, 2, 4], '중1C': [0, 2, 4],
    '중1보충': [0, 4], '중2': [1, 2, 3], '중3': [0, 2, 4],
    '고1': [1, 3, 5], '고2(미적분)': [1, 3, 5], '고3': [0, 2, 4],
}

# 반별 '출석 입력 확인' 시각 = 수업 종료 15분 뒤. {반: {요일idx(str): 'HH:MM'}}
# 요일 idx: 월0 화1 수2 목3 금4 토5 일6. 이 표에 있는 요일·시각에만 알림이 울린다.
DEFAULT_CLASS_TIMES = {
    '초3': {'1': '17:15', '2': '15:15', '3': '15:15'},          # 화 4~5시 / 수목 2~3시
    '초4': {'0': '15:15', '2': '15:15', '4': '15:15'},          # 월수금 2~3시
    '초5': {'1': '16:15', '2': '16:15', '3': '16:15'},          # 화수목 3~4시
    '초5A': {'1': '16:15', '2': '16:15', '3': '16:15'},         # 화수목 3~4시
    '초6': {'0': '16:15', '2': '16:15', '4': '16:15'},          # 월수금 3~4시
    '초6A': {'0': '16:15', '2': '16:15', '4': '16:15'},         # 월수금 3~4시
    '중1AB': {'0': '20:15', '4': '20:15', '2': '19:15'},        # 월금 6~8시 / 수 6~7시
    '중1C': {'0': '20:15', '4': '20:15', '2': '19:15'},         # 월금 6~8시 / 수 6~7시
    '중1보충': {'0': '18:15', '4': '18:15'},                    # 월금 5~6시
    '중2': {'1': '18:15', '3': '18:15', '2': '17:15'},          # 화목 4~6시 / 수 4~5시
    '중3': {'0': '18:15', '4': '18:15', '2': '18:15'},          # 월금 4~6시 / 수 5~6시
    '고1': {'1': '22:15', '3': '22:15', '5': '12:15'},          # 화목 8~10시 / 토 10~12시
    '고2(미적분)': {'1': '20:15', '3': '20:15', '5': '14:15'},  # 화목 6~8시 / 토 12~2시
    '고3': {'0': '22:15', '2': '22:15', '4': '22:15'},          # 월수금 8~10시
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


def _is_late(value):
    """지각 값인지: '지각' 이 있거나 'N분'(지각시간) 형식이면 지각."""
    s = str(value)
    return '지각' in s or bool(re.search(r'\d+\s*분', s))


def _attendance_color(value):
    """출석 값 → 글자색. 결석=빨강, 지각·조퇴=파랑, 정상=검정."""
    s = str(value)
    if _is_late(s) or '조퇴' in s:
        return COLOR_BLUE
    if '결석' in s or '결' == s.strip() or 'X' in s.upper():
        return COLOR_RED
    return COLOR_BLACK


def is_absent(value):
    """출석 값이 '결석'인지 판단한다 (지각·조퇴는 수업 참석으로 보아 제외).
    결석이면 그 학생의 과제수행·비고 칸을 비운다."""
    s = str(value)
    if _is_late(s) or '조퇴' in s:
        return False
    return 'X' in s.upper()


def _parse_md(s, year):
    """'M/D' → date. 실패 시 None."""
    try:
        m, d = map(int, str(s).split('/'))
        return datetime.date(year, m, d)
    except Exception:
        return None


def is_enrolled_during(info, monday, sunday, year):
    """info={'from':'M/D'|None,'to':'M/D'|None} 학생이 [monday,sunday] 주에 재적 중인지.
    - 그 주 시작 전에 이미 나간(to < monday) 학생은 제외 (→ 다음 주부터 안 나옴)
    - 그 주 이후에 들어오는(from > sunday) 학생도 제외"""
    if not info:
        return True
    to = _parse_md(info.get('to'), year)
    if to and to < monday:
        return False
    frm = _parse_md(info.get('from'), year)
    if frm and frm > sunday:
        return False
    return True


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


# ── 학적 변동 (신규등록/퇴원/담당변경) ─────────────────────────
# 출석칸에 찍는 표시
LIFECYCLE_MARK = {'신규등록': '신규등록', '전입': '담당변경', '퇴원': '퇴원', '전출': '담당변경'}
LIFECYCLE_ADD = {'신규등록', '전입'}   # 명단 추가 + 그 날짜부터 재적
LIFECYCLE_END = {'퇴원', '전출'}       # 그 날짜까지만 재적(이후 빈칸)


def _date_key(date_str):
    try:
        m, d = str(date_str).split('/')
        return (int(m), int(d))
    except Exception:
        return None


def student_active(enroll, student, date_str):
    """enroll: {student: {'from':'M/D'|None,'to':'M/D'|None}}.
    해당 날짜에 이 학생이 재적(입력 대상)인지."""
    info = (enroll or {}).get(student)
    if not info:
        return True
    dk = _date_key(date_str)
    if dk is None:
        return True
    fk = _date_key(info.get('from')) if info.get('from') else None
    tk = _date_key(info.get('to')) if info.get('to') else None
    if fk and dk < fk:
        return False
    if tk and dk > tk:
        return False
    return True


def add_student(ws, name):
    """시트 명단 끝에 학생을 새 열로 추가(이웃 학생열 서식 복사). 반환: 새 열번호."""
    start = _find_start_row(ws)
    last_col = _last_col(ws, start)
    new_col = last_col + 1
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, last_col)
        if src.has_style:
            ws.cell(r, new_col)._style = copy(src._style)
    from openpyxl.utils import get_column_letter
    sl, dl = get_column_letter(last_col), get_column_letter(new_col)
    if sl in ws.column_dimensions and ws.column_dimensions[sl].width:
        ws.column_dimensions[dl].width = ws.column_dimensions[sl].width
    for r in range(start, ws.max_row + 1):      # 기존 블록 데이터는 비움
        ws.cell(r, new_col).value = None
    ws.cell(start - 1, new_col).value = name    # 헤더에 이름
    return new_col


def rebuild_without_students(wb, sheet, drop_names):
    """sheet 에서 drop_names 학생 열을 제거한 새 시트로 교체(값·서식·병합 유지).
    delete_cols가 병합을 깨뜨리므로 필요한 열만 복사해 재구성한다."""
    from openpyxl.utils import get_column_letter
    if not drop_names:
        return
    ws = wb[sheet]
    roster = get_roster(ws)
    drop_cols = {roster[n] for n in drop_names if n in roster}
    if not drop_cols:
        return
    start = _find_start_row(ws)
    last_col = _last_col(ws, start)
    keep = [c for c in range(STUDENT_FIRST_COL, last_col + 1) if c not in drop_cols]
    colmap = {1: 1, 2: 2}
    for i, sc in enumerate(keep):
        colmap[sc] = STUDENT_FIRST_COL + i
    idx = wb.sheetnames.index(sheet)
    dest = wb.create_sheet(sheet + "__tmp")
    max_row = ws.max_row
    for sc, dc in colmap.items():
        for r in range(1, max_row + 1):
            s = ws.cell(r, sc)
            dest.cell(r, dc).value = s.value
            if s.has_style:
                dest.cell(r, dc)._style = copy(s._style)
        sl, dl = get_column_letter(sc), get_column_letter(dc)
        if sl in ws.column_dimensions and ws.column_dimensions[sl].width:
            dest.column_dimensions[dl].width = ws.column_dimensions[sl].width
    for r in range(1, max_row + 1):
        if r in ws.row_dimensions and ws.row_dimensions[r].height is not None:
            dest.row_dimensions[r].height = ws.row_dimensions[r].height
    for rng in list(ws.merged_cells.ranges):
        kept = [colmap[c] for c in range(rng.min_col, rng.max_col + 1) if c in colmap]
        if not kept:
            continue  # 병합이 전부 삭제열 안이면 버림
        r0, r1, c0, c1 = rng.min_row, rng.max_row, min(kept), max(kept)
        if r1 > r0 or c1 > c0:  # 단일 셀로 줄면 병합 불필요
            dest.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
    del wb[sheet]
    dest.title = sheet
    wb.move_sheet(sheet, offset=idx - wb.sheetnames.index(sheet))


def write_attendance(wb, sheet, date_str, data, enroll=None):
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

    # 학적 변동 처리: 명단 추가(신규·전입) + 출석칸에 표시
    life = data.get('학적') or {}
    for st, event in life.items():
        if event in LIFECYCLE_ADD and st not in roster:
            roster[st] = add_student(ws, st)
            last_col = _last_col(ws, start)
            groups = _detect_groups(ws, start, last_col)
        mark = LIFECYCLE_MARK.get(event)
        if not mark:
            continue
        if st not in roster:
            warnings.append(f"학적: '{st}' 학생을 못 찾음")
            continue
        cell = ws.cell(top + ROW_OFFSET['출석'], roster[st])
        cell.value = mark
        _apply_font_color(cell, _attendance_color(mark))
        written.append(f"학적 · {st} → {mark}")

    # 결석(지각·조퇴 제외) 학생: 과제수행·비고 칸을 비운다
    absent = {
        st for st, val in (data.get('출석') or {}).items()
        if st in roster and st not in life and is_absent(val)
    }
    for st in absent:
        for label in ('과제수행', '비고'):
            cell = ws.cell(top + ROW_OFFSET[label], roster[st])
            cell.value = None
            _apply_font_color(cell, COLOR_BLACK)

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
                if st in life:
                    continue  # 학적 표시로 이미 처리한 학생
                if label in ('과제수행', '비고') and st in absent:
                    continue  # 결석 학생의 과제·비고는 위에서 비웠으므로 건너뜀
                if st in roster and not student_active(enroll, st, date_str):
                    if label == '출석':
                        warnings.append(f"{st}: 재적 기간이 아니라 건너뜀")
                    continue  # 퇴원/전출 이후 또는 등록 전 → 빈칸 유지
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

    # 비고 행 라벨 변경 (일일테스트 등) — B열 라벨만 이 블록에서 교체
    new_label = data.get('비고라벨') or data.get('비고제목')
    if isinstance(new_label, str) and new_label.strip():
        ws.cell(top + ROW_OFFSET['비고'], 2).value = new_label.strip()
        written.append(f"비고 라벨 → {new_label.strip()}")

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
