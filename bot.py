# -*- coding: utf-8 -*-
"""
텔레그램 출석부 봇.

기능:
  1) 자유 문장으로 출석 입력 → Claude가 해석 → 미리보기 → '확인' 하면 xlsx 기록
  2) /출석부 [YY.MM]  → 해당(기본: 이번 달) 출석부 파일 전송
  3) xlsx 파일을 봇에게 보내면 → 저장(시드)
  4) /일정             → 반별 수업 요일 보기/변경
  5) /생성             → 이번 달 파일을 지난달에서 새로 생성
  6) 매월 1일 자동으로 새 달 파일 생성 + 알림

토큰/키는 환경변수(.env)에서 읽습니다.
데이터(출석부 xlsx, 설정)는 DATA_DIR(기본 ./data, Railway는 영구 볼륨)에 저장됩니다.
"""

import os
import io
import re
import json
import logging
import datetime
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
import anthropic
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

import attendance_core as ac
import report as rpt
import progress as pg
import subjects

# ── 설정 ───────────────────────────────────────────────────────
load_dotenv()
logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s | %(message)s", level=logging.INFO
)
log = logging.getLogger("attendance-bot")

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5")
# 과목 선택 (math/english/korean). 없으면 수학.
SUBJECT = os.environ.get("SUBJECT", "math")
SUBJ = subjects.get(SUBJECT)
SUBJ_NAME = SUBJ["display"]         # 예: 'Zest 수학과'
log.info("과목: %s (%s)", SUBJECT, SUBJ_NAME)
# 저장 폴더 우선순위: DATA_DIR > Railway 볼륨 자동경로 > 로컬 ./data
DATA_DIR = (
    os.environ.get("DATA_DIR")
    or os.environ.get("RAILWAY_VOLUME_MOUNT_PATH")
    or "data"
)
os.makedirs(DATA_DIR, exist_ok=True)
log.info("데이터 저장 폴더: %s", DATA_DIR)

claude = anthropic.AsyncAnthropic()

WD = ["월", "화", "수", "목", "금", "토", "일"]
FNAME_RE = re.compile(r"^\d{2}\.\d{2} 출석부\.xlsx$")
KST = ZoneInfo("Asia/Seoul")  # 알림은 한국시간 기준
# 밀린 미입력 알림은 이 날짜부터만 (그 이전 옛 날짜는 알림 제외)
BACKLOG_SINCE = datetime.date(2026, 7, 13)

# 확인/취소 대기 중인 입력: {chat_id: {"sheet","date","data","warnings"}}
pending: dict[int, dict] = {}
# 신규개강 단계 진행 중: {chat_id: {"sheet","step","weekdays","time","students"}}
opening_flow: dict[int, dict] = {}
# 종강 확인 대기: {chat_id: sheet}
closing_flow: dict[int, str] = {}
# 이상 날짜 삭제 확인 대기: {chat_id: [(sheet, top, date), ...]}
stray_flow: dict[int, list] = {}
# 휴강 지정 확인 대기: {chat_id: {"date", "targets": [(sheet, top), ...]}}
holiday_flow: dict[int, dict] = {}
# 잡담용 짧은 기억: {chat_id: [messages]}
chat_history: dict[int, list] = {}


# ── 백업 · 저장 · 실행취소 ──────────────────────────────────────
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
os.makedirs(BACKUP_DIR, exist_ok=True)
BACKUPS_KEEP = 20            # 파일별 백업 보관 개수
_last_change_path = os.path.join(DATA_DIR, "last_change.json")


def backup_file(path):
    """저장 직전 기존 파일을 타임스탬프로 백업. 파일별 최근 BACKUPS_KEEP개만 유지."""
    if not path or not os.path.exists(path):
        return None
    base = os.path.basename(path)
    ts = datetime.datetime.now(KST).strftime("%Y%m%d-%H%M%S")
    dst = os.path.join(BACKUP_DIR, f"{base}.{ts}.bak")
    try:
        import shutil
        shutil.copy2(path, dst)
    except Exception as e:
        log.warning("백업 실패 %s: %s", path, e)
        return None
    # 오래된 백업 정리
    try:
        olds = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(base + "."))
        for f in olds[:-BACKUPS_KEEP]:
            os.remove(os.path.join(BACKUP_DIR, f))
    except Exception:
        pass
    return dst


def save_wb(wb, path, undoable=False, desc=""):
    """워크북을 저장하되, 덮어쓰기 전 기존 파일을 백업한다.
    undoable=True면 '실행취소'로 되돌릴 수 있게 마지막 변경을 기록."""
    backup = backup_file(path)
    wb.save(path)
    if undoable:
        try:
            with open(_last_change_path, "w", encoding="utf-8") as f:
                json.dump(
                    {"path": path, "backup": backup, "desc": desc,
                     "ts": datetime.datetime.now(KST).isoformat()},
                    f, ensure_ascii=False,
                )
        except Exception as e:
            log.warning("실행취소 기록 실패: %s", e)
    return backup


# 음성·오타로 명령어 사이에 낀 띄어쓰기를 복원 (예: '시간 표' → '시간표')
_GLUE_WORDS = [
    "수업시간표", "주간시간표", "시간표", "출석부", "미리보기", "실행취소", "되돌리기",
    "방금취소", "요일추가", "주간보고서", "관리자등록", "도움말", "안내문",
    "신규개강", "신규등록", "공휴일", "휴강",
    "이번주", "저번주", "지난주", "다음주", "이번달", "지난달",
]
_GLUE_RE = [(re.compile(r"\s*".join(map(re.escape, w))), w) for w in _GLUE_WORDS]


def deglue(text):
    """알려진 명령어 안에 낀 공백을 붙여 인식률을 높인다."""
    for pat, w in _GLUE_RE:
        text = pat.sub(w, text)
    return text


# ── 저장소 헬퍼 ────────────────────────────────────────────────
def month_filename(year, month):
    return f"{str(year)[2:]}.{month:02d} 출석부.xlsx"


def month_path(year, month):
    return os.path.join(DATA_DIR, month_filename(year, month))


def current_ym():
    t = datetime.date.today()
    return t.year, t.month


def schedules_path():
    return os.path.join(DATA_DIR, "schedules.json")


def load_schedules():
    p = schedules_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    save_schedules(SUBJ["schedules"])
    return dict(SUBJ["schedules"])


def save_schedules(s):
    with open(schedules_path(), "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)


def times_path():
    return os.path.join(DATA_DIR, "class_times.json")


def load_times():
    """{반이름: {요일idx(str): 'HH:MM'}} — 반별·요일별 출석 입력 확인 시각.
    파일이 없으면 기본 시간표(수업 종료 15분 뒤)를 심고 반환한다."""
    p = times_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    default = {k: dict(v) for k, v in SUBJ["times"].items()}
    save_times(default)
    return default


def save_times(t):
    with open(times_path(), "w", encoding="utf-8") as f:
        json.dump(t, f, ensure_ascii=False, indent=2)


def hours_path():
    return os.path.join(DATA_DIR, "class_hours.json")


def load_hours():
    """{반이름: {요일idx(str): 'HH:MM~HH:MM'}} — 시간표에 띄우는 실제 수업 시간.
    파일이 없으면 과목 기본값(subjects.py)을 심고 반환한다."""
    p = hours_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    default = {k: dict(v) for k, v in SUBJ.get("hours", {}).items()}
    save_hours(default)
    return default


def save_hours(h):
    with open(hours_path(), "w", encoding="utf-8") as f:
        json.dump(h, f, ensure_ascii=False, indent=2)


def teachers_path():
    return os.path.join(DATA_DIR, "teachers.json")


def load_teachers():
    """{반이름: [chat_id, ...]} — 반별 담당 선생님(알림 수신자)."""
    p = teachers_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_teachers(t):
    with open(teachers_path(), "w", encoding="utf-8") as f:
        json.dump(t, f, ensure_ascii=False, indent=2)


def report_to_path():
    return os.path.join(DATA_DIR, "report_to.json")


def get_report_to():
    """주간 보고서 수신자 chat_id (없으면 None)."""
    p = report_to_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f).get("chat_id")
    return None


def set_report_to(chat_id):
    with open(report_to_path(), "w", encoding="utf-8") as f:
        json.dump({"chat_id": chat_id}, f)


# ── 관리자(주인) ───────────────────────────────────────────────
def admin_path():
    return os.path.join(DATA_DIR, "admin.json")


def get_admin():
    """관리자 chat_id. 환경변수 ADMIN_CHAT_ID 우선, 없으면 admin.json."""
    env = (os.environ.get("ADMIN_CHAT_ID") or "").strip()
    if env.lstrip("-").isdigit():
        return int(env)
    p = admin_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f).get("chat_id")
    return None


def set_admin(chat_id):
    with open(admin_path(), "w", encoding="utf-8") as f:
        json.dump({"chat_id": chat_id}, f)


def is_admin(chat_id):
    a = get_admin()
    return a is not None and a == chat_id


# ── 승인 멤버(사용 권한) ───────────────────────────────────────
def members_path():
    return os.path.join(DATA_DIR, "members.json")


def load_members():
    """{str(chat_id): {'name','status'}} — status: approved|pending."""
    p = members_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_members(m):
    with open(members_path(), "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=2)


def is_approved(chat_id):
    if is_admin(chat_id):
        return True
    m = load_members().get(str(chat_id))
    return bool(m) and m.get("status") == "approved"


def register_pending(chat_id, name):
    """미등록 사용자를 대기 상태로 기록. 새로 추가되면 True."""
    m = load_members()
    key = str(chat_id)
    if key in m:
        return False
    m[key] = {"name": name or "이름미상", "status": "pending"}
    save_members(m)
    return True


def _find_member_key(m, arg):
    """arg(이름 또는 chat_id)로 멤버 키 찾기."""
    arg = str(arg).strip()
    if arg in m:
        return arg
    hits = [k for k, v in m.items() if v.get("name") == arg]
    return hits[0] if len(hits) == 1 else None


def enroll_path():
    return os.path.join(DATA_DIR, "enrollments.json")


def load_enroll():
    """{반: {학생: {'from':'M/D'|None,'to':'M/D'|None}}} — 재적 기간."""
    p = enroll_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_enroll(e):
    with open(enroll_path(), "w", encoding="utf-8") as f:
        json.dump(e, f, ensure_ascii=False, indent=2)


def progress_path():
    return os.path.join(DATA_DIR, "progress.json")


def load_progress():
    """{반: {'units':[..], 'items':[..], 'steps':[..], 'cells':{학생:{'<ui>|<item>|<step>':'O'}}}}"""
    p = progress_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_progress(d):
    with open(progress_path(), "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


def prog_cfg(sheet):
    """반의 진도표 설정(단원/항목/단계) 반환. 없으면 기본값."""
    d = load_progress().get(sheet, {})
    return (d.get("units", []),
            d.get("items") or pg.DEFAULT_ITEMS,
            d.get("steps") or pg.DEFAULT_STEPS)


def closed_path():
    return os.path.join(DATA_DIR, "closed_classes.json")


def load_closed():
    """{반: 'M/D'} — 종강한 반과 종강일. 그 주 이후 입력 차단·다음달 제외."""
    p = closed_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_closed(c):
    with open(closed_path(), "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False, indent=2)


def parse_weekdays(text):
    """'월수금'/'월 수 금'/'화목토' → 정렬된 요일 인덱스 목록."""
    base = str(text).replace("요일", "")
    return sorted({WD.index(ch) for ch in base if ch in WD})


def class_input_blocked(sheet, date_str):
    """종강한 반의 '종강 주 이후' 날짜면 True(입력 차단)."""
    closed = load_closed().get(sheet)
    if not closed:
        return False
    y = datetime.date.today().year
    try:
        cm, cd = map(int, str(closed).split("/"))
        dm, dd = map(int, str(date_str).split("/"))
    except ValueError:
        return False
    _, csun = rpt.week_bounds(datetime.date(y, cm, cd))
    return datetime.date(y, dm, dd) > csun


def class_alarm_blocked(sheet, date_str):
    """알림 전용: 종강한 반은 '종강일 그날부터'(그 주 포함) 알림 제외.
    (입력 차단보다 더 이른 시점부터 끊는다.)"""
    closed = load_closed().get(sheet)
    if not closed:
        return False
    y = datetime.date.today().year
    try:
        cm, cd = map(int, str(closed).split("/"))
        dm, dd = map(int, str(date_str).split("/"))
    except ValueError:
        return False
    return datetime.date(y, dm, dd) >= datetime.date(y, cm, cd)


def class_off_timetable(sheet, closed):
    """종강한 반을 시간표에서 뺄 때인지. 입력 차단과 같은 기준으로,
    종강 주까지는 남기고 다음 주부터 뺀다. closed = load_closed() 결과."""
    day = closed.get(sheet)
    if not day:
        return False
    today = datetime.datetime.now(KST).date()
    try:
        cm, cd = map(int, str(day).split("/"))
    except ValueError:
        return True
    _, csun = rpt.week_bounds(datetime.date(today.year, cm, cd))
    return today > csun


def apply_enroll_events(sheet, date_str, life):
    """학적 이벤트를 enrollments.json 에 반영. life={학생: 유형}."""
    e = load_enroll()
    cls = e.setdefault(sheet, {})
    for st, event in life.items():
        info = cls.setdefault(ac.nfc(st), {"from": None, "to": None})
        if event in ac.LIFECYCLE_ADD:
            info["from"] = date_str
            info["to"] = None
        elif event in ac.LIFECYCLE_END:
            info["to"] = date_str
    save_enroll(e)


def chats_path():
    return os.path.join(DATA_DIR, "chats.json")


def remember_chat(chat_id):
    ids = set()
    if os.path.exists(chats_path()):
        with open(chats_path(), encoding="utf-8") as f:
            ids = set(json.load(f))
    if chat_id not in ids:
        ids.add(chat_id)
        with open(chats_path(), "w", encoding="utf-8") as f:
            json.dump(sorted(ids), f)


def known_chats():
    if os.path.exists(chats_path()):
        with open(chats_path(), encoding="utf-8") as f:
            return json.load(f)
    return []


def list_month_files():
    return sorted(f for f in os.listdir(DATA_DIR) if FNAME_RE.match(f))


def latest_month_file(before=None):
    """가장 최근 YY.MM 파일명. before(파일명) 미만으로 제한 가능."""
    files = list_month_files()
    if before:
        files = [f for f in files if f < before]
    return files[-1] if files else None


def load_current_wb():
    y, m = current_ym()
    p = month_path(y, m)
    if os.path.exists(p):
        return ac.load_workbook(p), p
    return None, p


def load_latest_wb():
    """가장 최근 달 파일을 로드 (파싱 시 명단/맥락용). 없으면 (None, None)."""
    f = latest_month_file()
    if not f:
        return None, None
    p = os.path.join(DATA_DIR, f)
    return ac.load_workbook(p), p


def load_wb_for_date(date_str):
    """'8/4' → 8월 파일 로드. 반환 (wb|None, path, month)."""
    month = int(str(date_str).split("/")[0])
    y = datetime.date.today().year
    p = month_path(y, month)
    if os.path.exists(p):
        return ac.load_workbook(p), p, month
    return None, p, month


# ── 새 달 생성 ────────────────────────────────────────────────
def generate_month_file(year, month):
    """지난달(가장 최근) 파일을 바탕으로 (year,month) 파일 생성. 반환: 저장경로."""
    target_name = month_filename(year, month)
    src_name = latest_month_file(before=target_name)
    if not src_name:
        raise FileNotFoundError("바탕이 될 이전 달 출석부가 없습니다. 먼저 파일을 보내주세요.")
    src = ac.load_workbook(os.path.join(DATA_DIR, src_name))
    scheds = load_schedules()
    schedules = {k: v for k, v in scheds.items() if k in src.sheetnames}
    wb = ac.generate_month(src, year, month, schedules)
    # 종강한 반은 새 달 파일에서 제외
    for sheet in list(load_closed()):
        if sheet in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb[sheet]
    # 퇴원·전출(재적 to 설정) 학생은 새 달 명단에서 제거
    enroll = load_enroll()
    for sheet in list(wb.sheetnames):
        drops = [st for st, info in enroll.get(sheet, {}).items() if info.get("to")]
        if drops:
            ac.rebuild_without_students(wb, sheet, drops)
    out = month_path(year, month)
    wb.save(out)
    log.info("생성 완료: %s (바탕: %s)", out, src_name)
    return out


# ── Claude 파싱 ───────────────────────────────────────────────
PARSE_SYSTEM = """너는 학원 출석부 입력 도우미다. 선생님의 한국어 메시지를 출석부 기록용 JSON으로 변환한다.
반드시 아래 형식의 JSON만 출력한다(설명·코드블록 금지).

{
  "type": "attendance" 또는 "other",
  "sheet": "<반 시트명 정확히>",
  "date": "M/D",
  "출석":   {"학생명": "O" 또는 "X(사유)"},
  "수업내용": "<문자열>",
  "과제수행": {"학생명": "O" 또는 "X"},
  "다음과제": "<문자열>",
  "비고":   {"학생명": "<문자열 또는 점수>"},
  "비고라벨": "<비고 행의 이름을 바꿀 때만. 예: 일일test, 주간test, 단원평가>",
  "학적": {"학생명": "신규등록" 또는 "퇴원" 또는 "전입" 또는 "전출"}
}

규칙:
- 출석/과제/학적(신규등록·퇴원·담당변경) 관련 입력이면 type="attendance", 그 외 잡담·질문이면 {"type":"other"} 만 출력.
- sheet 는 제공된 시트 목록 중 하나와 정확히 일치해야 한다.
- 학생명은 제공된 명단의 이름과 정확히 일치시킨다(부분/별칭이면 가장 맞는 이름으로).
- date 는 'M/D' 형식(예: 8/4). '오늘/어제/지난 화요일' 등은 오늘 날짜 기준으로 계산.
- 언급되지 않은 항목(키)은 넣지 않는다. 값이 없으면 그 키를 생략.
- '전원/다 출석/나머지 다 출석' 같은 표현은 명단 전체에 O 로 채우되, 개별 언급(결석 등)이 우선.
- 출석 값: 정상 출석="O"(대문자 오).
    · 결석은 "X (사유)" 형식. 사용자가 적은 사유 표현을 그대로 넣는다(임의로 바꾸지 말 것).
      예) '결석(사유미상)'="X (사유미상)", 여행으로 결석="X (여행)", 아파서 결석="X (몸살)".
      사유를 전혀 밝히지 않고 '그냥 결석'만 말하면 "X (무단)"으로 적는다.
    · 조퇴는 X 를 쓰지 않고 "조퇴시간 조퇴(사유)" 형식으로 적는다('조퇴' 글자 포함).
      예) 휴가로 20분 조퇴="20분 조퇴(휴가)", 그냥 20분 조퇴="20분 조퇴",
      시간 모르고 사유만 있으면="조퇴(사유)", 시간·사유 둘 다 모르면="조퇴".
    · 지각은 X 를 쓰지 않고 "지각시간(사유)" 형식으로 적는다.
      예) 병원가느라 30분 늦음="30분(병원)", 그냥 30분 지각="30분",
      사유만 있고 시간을 모르면="지각(사유)", 시간·사유 둘 다 모르면="지각".
- 과제수행 값: 완료="O". 절반/50%="50%".
    · 그냥 안 했으면(별다른 사유 없음) 사유 없이 "X" 만 적는다("X (미완성)" 처럼 임의로 사유를 붙이지 말 것).
    · 미지참/미수령 등 사용자가 사유를 말한 경우에만 "X (사유)" 형식으로 그 사유를 넣는다. 예) 미지참="X (미지참)", 미수령="X (미수령)".
- 시험/평가 관련: 시험 이름과 학생·점수가 나오면 "비고라벨"에 시험 이름, "비고"에 학생별 점수를 넣는다.
  · '비고를 ~로 바꿔줘'라고 명시하지 않아도, 시험 이름(일일테스트, 일일test, 주간test, 단원평가 등)이 등장하고
    점수를 말하면 자동으로 "비고라벨"=그 시험 이름, "비고"={학생:점수}로 채운다.
    예) '일일테스트 우현 규림 둘다 75점' → "비고라벨":"일일테스트", "비고":{"남우현":"75","김규림":"75"}.
  · 점수만 있고 시험 이름이 없으면 "비고라벨"은 생략하고 "비고"에 점수만 넣는다.
  · '모두/전원 N점'이면 명단 전체에 N을 채운다(개별 언급 우선). '둘다/셋다 N점'은 언급된 학생들에게 N.
- 학적 변동은 "학적"에 {학생명: 유형}으로 넣는다. 유형은 정확히 다음 중 하나:
    "신규등록"(새 등록), "퇴원"(그만둠), "전입"('담당변경(전입)'), "전출"('담당변경(전출)').
  · 신규등록·전입은 새 학생일 수 있으니 명단에 없어도 말한 이름을 그대로 쓴다.
  · sheet(반)·date 는 평소처럼 채운다. 그 학생은 "출석" 등 다른 키에는 넣지 않는다.
"""


async def parse_message(text, wb):
    today = datetime.date.today()
    rosters = ac.rosters_summary(wb)
    dates = ac.sheet_dates(wb)
    ctx = {
        "오늘": f"{today.year}-{today.month:02d}-{today.day:02d} ({WD[today.weekday()]}요일)",
        "시트별_명단": rosters,
        "시트별_수업일": dates,
    }
    user = (
        f"[참고정보]\n{json.dumps(ctx, ensure_ascii=False)}\n\n"
        f"[선생님 메시지]\n{text}"
    )
    resp = await claude.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1500,
        thinking={"type": "disabled"},  # 파싱은 생각 불필요 — 빠르고 저렴하게
        system=PARSE_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    raw = "".join(b.text for b in resp.content if b.type == "text").strip()
    m = re.search(r"\{.*\}", raw, re.S)
    if not m:
        return {"type": "other"}
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return {"type": "other"}


CHAT_SYSTEM = """너는 학원 선생님을 돕는 '출석부 텔레그램 봇'이다.
- 출석부 데이터는 서버 파일(xlsx)에 실제로 저장·조회된다. "나는 AI라 접근 못 한다"는 식으로 말하지 마라.
- 파일을 받고 싶어 하면 '출석부' 또는 '/download' 라고 보내면 된다고 안내해라. 특정 달은 '출석부 7월' 처럼.
- 출석 입력 예시: '초5 오늘 남우현 결석, 나머지 출석. 수업 분수나눗셈'.
- 매우 중요: 파일을 실제로 바꾸는 작업(날짜 변경/삭제/수정 등)을 네가 방금 처리했다고 말하지 마라.
  하지도 않은 일을 '완료했습니다'라고 절대 말하지 마라. 그런 기능이 명령으로 있으면 사용법을,
  없으면 '아직 그 기능은 없어요'라고 솔직히 안내하라. (예: 날짜 변경은 '고1 7/21 → 7/22')
- 수학 개념 질문 등에는 평소처럼 친절히 답해도 된다. 답은 간결하게."""


async def chat_reply(chat_id, text):
    """출석 입력이 아닌 메시지는 봇 성격을 아는 상태로 답한다."""
    history = chat_history.get(chat_id, [])
    history.append({"role": "user", "content": text})
    resp = await claude.messages.create(
        model=CLAUDE_MODEL, max_tokens=800, thinking={"type": "disabled"},
        system=CHAT_SYSTEM, messages=history
    )
    reply = "".join(b.text for b in resp.content if b.type == "text")
    history.append({"role": "assistant", "content": reply})
    chat_history[chat_id] = history[-12:]
    return reply


# ── 미리보기 텍스트 ────────────────────────────────────────────
def build_preview(wb, parsed):
    sheet = parsed.get("sheet")
    date = parsed.get("date")
    if sheet not in wb.sheetnames:
        return None, f"'{sheet}' 반을 찾을 수 없어요. (반: {', '.join(wb.sheetnames)})"
    ws = wb[sheet]
    if ac.find_date_block(ws, date) is None:
        ds = ", ".join(ac.sheet_dates(wb)[sheet])
        return None, f"{sheet}에 '{date}' 수업일이 없어요.\n가능한 날짜: {ds}"

    roster = set(ac.get_roster(ws).keys())
    lines = [f"📋 <b>{sheet}</b> · <b>{date}</b> 에 이렇게 기록할게요:"]
    warnings = []

    att = parsed.get("출석") or {}
    absent = {n for n, v in att.items() if ac.is_absent(v)}
    if att:
        parts = []
        for name, val in att.items():
            parts.append(f"{name} {val}")
            if ac.nfc(name) not in roster:
                warnings.append(name)
        lines.append("• 출석: " + ", ".join(parts))
    if parsed.get("수업내용"):
        lines.append(f"• 수업내용: {parsed['수업내용']}")
    hw = parsed.get("과제수행") or {}
    if hw:
        shown = {n: v for n, v in hw.items() if n not in absent}
        if shown:
            lines.append("• 과제수행: " + ", ".join(f"{n} {v}" for n, v in shown.items()))
        warnings += [n for n in hw if ac.nfc(n) not in roster]
    if parsed.get("다음과제"):
        lines.append(f"• 다음과제: {parsed['다음과제']}")
    bg_label = (parsed.get("비고라벨") or "비고").strip()
    if parsed.get("비고라벨"):
        lines.append(f"• 비고 칸 이름 → <b>{bg_label}</b> 으로 변경")
    bg = parsed.get("비고") or {}
    if isinstance(bg, dict) and bg:
        shown_bg = {n: v for n, v in bg.items() if n not in absent}
        if shown_bg:
            lines.append(f"• {bg_label}: " + ", ".join(f"{n}: {v}" for n, v in shown_bg.items()))
        warnings += [n for n in bg if ac.nfc(n) not in roster]
    if absent:
        lines.append(f"• 결석: {', '.join(sorted(absent))} → 과제수행·비고는 비워둡니다.")

    life = parsed.get("학적") or {}
    if isinstance(life, dict) and life:
        desc = {"신규등록": "신규등록·명단 추가", "전입": "담당변경(전입)·명단 추가",
                "퇴원": "퇴원·이후 빈칸", "전출": "담당변경(전출)·이후 빈칸"}
        lines.append("• 학적: " + ", ".join(f"{n} → {desc.get(v, v)}" for n, v in life.items()))

    missing = sorted(w for w in set(warnings) if w not in life)
    if missing:
        lines.append("\n⚠️ 명단에 없는 이름: " + ", ".join(missing) + " (그대로 두면 무시됩니다)")
    lines.append("\n맞으면 <b>확인</b>, 아니면 <b>취소</b> 라고 보내주세요.")
    return "\n".join(lines), None


# ── 사용 안내 (처음 시작 시 자동 전송) ─────────────────────────
GUIDE = f"""📋 <b>{SUBJ_NAME} 출석부 봇 사용법</b>

{SUBJ_NAME} 선생님들이 함께 쓰는 공용 출석부예요.
<b>채팅으로 말씀하면 → 봇이 파일에 자동 기록</b>합니다. (엑셀 직접 안 여셔도 돼요!)

━━━━━━━━━━━━━━━
<b>1️⃣ 출석 알려주기</b>

그냥 편하게 문장으로 말씀하세요 👇
<blockquote>초5 오늘 남우현 결석, 나머지 출석. 수업 분수나눗셈, 다음과제 42쪽</blockquote>
→ 봇이 미리보기를 보여드려요. <u><b>확인</b></u> 하면 기록하고 <b>기록 내용을 사진으로</b> 보내드려요. (아니면 <u><b>취소</b></u>)

이미 기록한 것도 <b>특정 칸만</b> 고칠 수 있어요 👇
<blockquote>초5 7/15 남우현만 출석으로 바꿔줘</blockquote>

━━━━━━━━━━━━━━━
<b>2️⃣ 이렇게 말하면 → 이렇게 기록돼요</b>

• 결석 → <code>X (사유)</code>  (사유 없으면 <code>무단</code>)
• 지각 → <code>30분(병원)</code>
• 조퇴 → <code>20분 조퇴(휴가)</code>  (사유 같이 기록)
• 숙제: 안 함 <code>X</code> · 절반 <code>50%</code> · 안 가져옴 <code>미지참</code> · 안 받아감 <code>미수령</code>
• 시험: 시험명·학생·점수만 말하면 자동 입력 👇
<blockquote>초5 오늘 일일테스트 우현 규림 둘다 75점</blockquote>

━━━━━━━━━━━━━━━
<b>3️⃣ 조회 (파일·이미지)</b>

• <b>출석부</b> → 이번 달 파일  (<b>출석부 7월</b> = 특정 달)
• 반+날짜 미리보기 (이미지) 👇
<blockquote>초5 7월 15일 미리보기</blockquote>
• 학생 기간별 출결 (이미지) 👇
<blockquote>남우현 7월</blockquote>
• 학생 과제 조회 👇
<blockquote>7월 15일 원서진 과제</blockquote>

━━━━━━━━━━━━━━━
<b>4️⃣ 내 담당 반 지정</b> ⭐

내 반 이름으로 보내면 그 반 알림만 받아요 👇
<blockquote>담당 초5 중2</blockquote>
반 목록은 <b>담당</b> 이라고만 보내면 떠요.

━━━━━━━━━━━━━━━
<b>5️⃣ 알림 · 시간표</b>

• 수업 끝나고 <u>15분 뒤까지 미입력</u>이면 담당 쌤에게 알림 (매일 1시에 밀린 것도 한 번 더)
• <b>시간표</b> → 수업 시간표 (<code>시간표 초5A</code> · <code>시간표 월</code> · <code>시간표 담당</code>)
• <b>일정</b> → 반별 수업 요일  ·  <code>알림 초5 21:00</code> → 알림 시각 변경

━━━━━━━━━━━━━━━
<b>6️⃣ 그 밖에</b>

• <b>휴강</b>: <code>7/20 휴강</code> (사유는 뒤에 <code>7/20 휴강 폭우</code>) · 취소 <code>7/20 휴강취소</code>
   → 그날 수업 있는 반 전부 빨간 '휴강'으로 (알림도 안 가요)
• <b>개강 / 종강</b>: <code>고3B 신규개강</code> / <code>고3B 종강</code> (반 자체를 만들거나 종료)
   → 개강 때 받은 수업 시간이 <b>시간표에 자동 추가</b>, 종강하면 그 주까지만 뜨고 <b>다음 주부터 빠져요</b>
   ※ 학생 한 명은 <b>신규등록</b> · <b>퇴원</b> 으로
• <b>점검</b>: 이상한 날짜(중복·다른 달) 확인·정리
• <b>날짜 변경</b>: <code>고1 7/21 → 7/22</code> (그 칸의 기록은 그대로, 날짜만 바뀜)
• <b>실행취소</b>: <u>방금 입력을 잘못했을 때</u> 직전 상태로 되돌리기 (매 입력 전 자동 백업)
• <b>통계</b>: 출석률·결석·과제 미제출 집계
   → <code>통계</code>(전체) · <code>통계 초5A</code> · <code>통계 이번주</code> · <code>통계 지난달</code>
• <b>진도표</b>: 단원별 항목 O/X 관리 (기본: 유형서·심화유형서·단원평가 × 수정완료·밴드완료)
   → 단원 설정: <code>초5A 진도단원 분수의 나눗셈, 각기둥과 각뿔, …</code>
   → 항목/단계 바꾸기(반별): <code>초5A 진도항목 개념서, 문제집</code> · <code>초5A 진도단계 1차, 2차</code>
   → 입력: <code>원서진 1단원 유형서 수정완료</code> (X·취소도 됨)
   → 보기: <code>초5A 진도</code>(반 전체) · <code>원서진 진도</code>(개인) · 설정확인 <code>초5A 진도설정</code>

━━━━━━━━━━━━━━━
궁금하면 아무 때나 <b>도움말</b> 보내면 이 안내가 다시 떠요. 🙂
음성입력으로 <u>띄어쓰기가 조금 껴도</u>(예: '시간 표') 알아들어요.

※ 이 채팅은 <b>1:1 개인 채팅</b>이라 다른 쌤은 못 봐요. 편하게 쓰세요.
※ 오류나 보완할 점은 <b>태일쌤</b>에게 알려주세요. 열심히 고쳐보겠습니다! 🙇"""


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    remember_chat(update.effective_chat.id)
    await update.message.reply_text(GUIDE, parse_mode="HTML")


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start(update, context)


async def cmd_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    scheds = load_schedules()
    args = context.args
    if not args:
        lines = ["📅 반별 수업 요일:"]
        for name, idxs in scheds.items():
            lines.append(f"• {name}: {'·'.join(WD[i] for i in idxs)}")
        lines.append("\n변경: /일정 고1 화목토")
        await update.message.reply_text("\n".join(lines))
        return
    if len(args) < 2:
        await update.message.reply_text("형식: /일정 <반> <요일들>  예) /일정 고1 화목토")
        return
    name = args[0]
    if name not in scheds:
        await update.message.reply_text(f"'{name}' 반이 없어요. ({', '.join(scheds)})")
        return
    days = "".join(args[1:])
    idxs = sorted({WD.index(ch) for ch in days if ch in WD})
    if not idxs:
        await update.message.reply_text("요일을 인식 못했어요. 예) 화목토")
        return
    scheds[name] = idxs
    save_schedules(scheds)
    await update.message.reply_text(
        f"✅ {name} 수업 요일을 {'·'.join(WD[i] for i in idxs)} 로 바꿨어요. (다음 달 생성부터 적용)"
    )


def parse_time_token(text):
    """'17:00','17시','5시','오후 5시','오후5시30분','5:30','17' → 'HH:MM'(24시간). 실패 시 None."""
    t = str(text).strip()
    pm = any(k in t for k in ("오후", "저녁", "밤"))
    am = "오전" in t
    t2 = re.sub(r"(오전|오후|저녁|밤|정각)", "", t).strip()
    m = re.search(r"(\d{1,2})\s*[:시]\s*(\d{1,2})?", t2)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2)) if m.group(2) else 0
    else:
        m2 = re.search(r"\b(\d{1,2})\b", t2)
        if not m2:
            return None
        hh, mm = int(m2.group(1)), 0
    if pm and hh < 12:
        hh += 12
    if am and hh == 12:
        hh = 0
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        return None
    return f"{hh:02d}:{mm:02d}"


def parse_hours_token(text):
    """'15:00~16:00', '오후 3시~4시', '4시부터 6시' → 'HH:MM~HH:MM'. 실패 시 None."""
    parts = re.split(r"\s*(?:~|—|–|-|부터|에서)\s*", str(text).strip(), maxsplit=1)
    if len(parts) != 2:
        return None
    head, tail = parts[0], parts[1]
    # '오후 3시~4시' — 앞에만 붙은 오전/오후를 끝 시각에도 적용한다.
    mark = re.search(r"(오전|오후|저녁|밤)", head)
    if mark and not re.search(r"(오전|오후|저녁|밤)", tail):
        tail = f"{mark.group(1)} {tail}"
    start, end = parse_time_token(head), parse_time_token(tail)
    if not start or not end:
        return None
    if end <= start:  # '11시~1시' 처럼 끝이 이르면 오후로 본다
        eh, em = map(int, end.split(":"))
        if eh + 12 <= 23:
            end = f"{eh + 12:02d}:{em:02d}"
    if end <= start:
        return None
    return f"{start}~{end}"


def hours_end_plus(hr, minutes=15):
    """'15:00~16:00' → 수업 종료 +minutes 의 'HH:MM'. 실패 시 None."""
    try:
        eh, em = map(int, hr.split("~")[1].split(":"))
    except (ValueError, IndexError):
        return None
    t = (eh * 60 + em + minutes) % (24 * 60)
    return f"{t // 60:02d}:{t % 60:02d}"


async def cmd_remind(update: Update, context: ContextTypes.DEFAULT_TYPE):
    remember_chat(update.effective_chat.id)
    args = context.args
    times = load_times()
    scheds = load_schedules()

    if not args:  # 현재 알림표 보기
        if not times:
            await update.message.reply_text(
                "설정된 출석 입력 알림이 없어요.\n"
                "예) 알림 초5 화수목 16:15  → 그 요일 16:15에 출석 미입력 시 알림\n"
                "끄기) 알림 초5 끄기"
            )
            return
        lines = ["📝 출석 입력 알림 (한국시간, 수업일에 미입력 시 알림):"]
        for cls in sorted(times):
            table = times[cls]
            if not isinstance(table, dict):
                continue
            parts = " ".join(f"{WD[int(d)]}{t}" for d, t in sorted(table.items()))
            lines.append(f"• {cls}: {parts}")
        lines.append("\n변경) 알림 초5 화수목 16:15    끄기) 알림 초5 끄기 (또는 알림 초5 수 끄기)")
        await update.message.reply_text("\n".join(lines))
        return

    cls = args[0]
    # 남은 토큰을 요일 / 시간 / 끄기 로 분류
    day_idxs, time_tokens, off = [], [], False
    for tok in args[1:]:
        t = tok.strip()
        if t in ("끄기", "해제", "삭제", "제거", "off", "취소"):
            off = True
            continue
        base = t.replace("요일", "")
        if base and all(ch in WD for ch in base):   # 화수목 / 월 / 토
            day_idxs += [WD.index(ch) for ch in base]
        elif t:
            time_tokens.append(t)
    day_idxs = sorted(set(day_idxs))
    table = times[cls] if isinstance(times.get(cls), dict) else {}

    if off:
        if day_idxs:  # 특정 요일만 끄기
            for d in day_idxs:
                table.pop(str(d), None)
            if table:
                times[cls] = table
            else:
                times.pop(cls, None)
            msg = f"✅ {cls} {'·'.join(WD[d] for d in day_idxs)}요일 알림을 껐어요."
        else:  # 그 반 전체 끄기
            times.pop(cls, None)
            msg = f"✅ {cls} 알림을 모두 껐어요."
        save_times(times)
        await update.message.reply_text(msg)
        return

    hhmm = parse_time_token(" ".join(time_tokens))
    if not hhmm:
        await update.message.reply_text(
            "시간을 못 알아들었어요.\n예) 알림 초5 화수목 16:15 · 알림 중1 수 19:15"
        )
        return
    if not day_idxs:
        day_idxs = list(scheds.get(cls, []))
    if not day_idxs:
        await update.message.reply_text(
            f"'{cls}' 수업 요일을 함께 알려주세요.\n예) 알림 {cls} 화목 18:15"
        )
        return
    for d in day_idxs:
        table[str(d)] = hhmm
    times[cls] = table
    save_times(times)
    await update.message.reply_text(
        f"✅ {cls} — {'·'.join(WD[d] for d in day_idxs)}요일 {hhmm}에 출석 미입력 시 알려드릴게요."
    )


def _timetable_byday(classes, active, hours, only_days=None, ending=()):
    """요일별 시간표 줄 목록. only_days(요일idx 집합)가 있으면 그 요일만.
    ending 에 든 반은 이번 주가 마지막이라 '(이번 주 종강)'을 붙인다."""
    byday = {i: [] for i in range(7)}
    for cls in classes:
        for i in active.get(cls, []):
            if only_days is not None and i not in only_days:
                continue
            hr = (hours.get(cls) or {}).get(str(i), "")
            byday[i].append((hr or "99:99", cls, hr))
    lines = []
    for i in range(7):
        items = sorted(byday[i], key=lambda x: (x[0], x[1]))
        if not items:
            continue
        parts = []
        for _, cls, hr in items:
            label = f"{cls} {hr}" if hr else f"{cls} (수업 시간 미설정)"
            if cls in ending:
                label += " ⏹ 이번 주 종강"
            parts.append(label)
        lines.append(f"<b>[{WD[i]}]</b>\n  " + "\n  ".join(parts))
    return lines


async def cmd_timetable(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """수업 시간표. 인자로 요일/반/담당 필터 가능."""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    scheds = load_schedules()
    hours = load_hours()
    closed = load_closed()
    active = {c: idxs for c, idxs in scheds.items() if not class_off_timetable(c, closed)}
    ending = {c for c in active if c in closed}   # 종강했지만 이번 주까지는 수업
    arg = " ".join(context.args).strip() if context.args else ""

    # 1) 담당(내 반)
    if arg in ("담당", "내반", "내", "나", "내시간표", "내 시간표"):
        teachers = load_teachers()
        mine = [c for c in active if chat_id in teachers.get(c, [])]
        if not mine:
            await update.message.reply_text("담당으로 지정한 반이 없어요. '담당 초5A' 처럼 먼저 지정하세요.")
            return
        lines = [f"📅 <b>{SUBJ_NAME} · 내 담당 시간표</b>"] + _timetable_byday(mine, active, hours, ending=ending)
        await update.message.reply_text("\n".join(lines), parse_mode="HTML")
        return

    # 2) 요일 필터 (예: '월', '월요일', '월수금')
    days = parse_weekdays(arg) if arg else []
    if arg and days and all(ch in "월화수목금토일요일 " for ch in arg):
        lines = [f"📅 <b>{SUBJ_NAME} · {'·'.join(WD[i] for i in days)} 시간표</b>"]
        lines += _timetable_byday(list(active), active, hours, only_days=set(days), ending=ending)
        if len(lines) == 1:
            lines.append("그 요일에 수업이 없어요.")
        await update.message.reply_text("\n".join(lines), parse_mode="HTML")
        return

    # 3) 반 필터 (예: '초5A')
    if arg:
        cls = arg if arg in active else next((c for c in active if c.replace(" ", "") == arg.replace(" ", "")), None)
        if cls:
            parts = [f"{WD[i]} {(hours.get(cls) or {}).get(str(i), '')}".strip() for i in sorted(active[cls])]
            note = "\n⏹ 종강한 반이라 이번 주까지만 수업해요." if cls in ending else ""
            await update.message.reply_text(
                f"📅 <b>{cls}</b> 수업 시간표\n• " + "\n• ".join(parts) + note, parse_mode="HTML"
            )
            return
        await update.message.reply_text(
            f"'{arg}'를 못 알아들었어요. 예) 시간표 / 시간표 월 / 시간표 초5A / 시간표 담당"
        )
        return

    # 4) 전체
    lines = [f"📅 <b>{SUBJ_NAME} 주간 수업 시간표</b>"] + _timetable_byday(list(active), active, hours, ending=ending)
    if len(lines) == 1:
        lines.append("아직 등록된 반이 없어요. '일정'으로 요일을 설정하면 여기 모여요.")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def cmd_add_weekday(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """이번 달 출석부 파일의 날짜 셀에 요일을 붙인다(데이터 유지). 관리자 전용."""
    if not await require_admin(update):
        return
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부 파일이 없어요.")
        return
    y, _ = current_ym()
    n = ac.add_weekday_labels(wb, y)
    save_wb(wb, path)
    await update.message.reply_text(
        f"✅ 날짜 {n}개에 요일을 붙였어요 (예: 7/15 → 7/15(수)).\n"
        "받아보려면 '출석부' 보내세요."
    )


async def cmd_undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """마지막 출석 입력을 직전 상태로 되돌린다."""
    if not os.path.exists(_last_change_path):
        await update.message.reply_text("↩️ 되돌릴 최근 입력이 없어요.")
        return
    try:
        with open(_last_change_path, encoding="utf-8") as f:
            ch = json.load(f)
    except Exception:
        await update.message.reply_text("↩️ 되돌릴 기록을 읽지 못했어요.")
        return
    path, backup, desc = ch.get("path"), ch.get("backup"), ch.get("desc", "")
    if not backup or not os.path.exists(backup):
        await update.message.reply_text(
            "↩️ 이 입력은 백업본이 없어 되돌릴 수 없어요. (새 파일 생성 등)"
        )
        return
    try:
        import shutil
        # 되돌리기도 되돌릴 수 있게, 지금 상태를 한 번 더 백업
        backup_file(path)
        shutil.copy2(backup, path)
    except Exception as e:
        await update.message.reply_text(f"↩️ 되돌리기 실패: {e}")
        return
    try:
        os.remove(_last_change_path)  # 한 번만 되돌리도록
    except Exception:
        pass
    ts = (ch.get("ts") or "")[:16].replace("T", " ")
    await update.message.reply_text(
        f"↩️ 되돌렸어요 — <b>{desc}</b> 직전 상태로 복원했어요.\n"
        f"(입력 시각: {ts})\n확인하려면 '출석부' 또는 미리보기 해보세요.",
        parse_mode="HTML",
    )


def _collect_stats(ws, dates):
    """한 시트에서 dates 기간의 학생별 출결·과제 집계."""
    roster = ac.get_roster(ws)
    per = {}
    for nm in roster:
        per[nm] = {"total": 0, "absent": 0, "late": 0, "hw_miss": 0, "hw_half": 0}
    for ds in dates:
        top = ac.find_date_block(ws, ds)
        if top is None:
            continue
        for nm, col in roster.items():
            av = ws.cell(top + ac.ROW_OFFSET["출석"], col).value
            if av not in (None, ""):
                per[nm]["total"] += 1
                if ac.is_absent(av):
                    per[nm]["absent"] += 1
                elif ac._is_late(str(av)) or "조퇴" in str(av):
                    per[nm]["late"] += 1
            hv = ws.cell(top + ac.ROW_OFFSET["과제수행"], col).value
            if hv not in (None, ""):
                c = ac._homework_color(hv)
                if c == ac.COLOR_RED:
                    per[nm]["hw_miss"] += 1
                elif c == ac.COLOR_BLUE:
                    per[nm]["hw_half"] += 1
    return per


async def cmd_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """기간·반별 출결/과제 집계. 예) 통계 / 통계 초5A / 통계 이번주 / 통계 지난달"""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    text = " ".join(context.args) if context.args else ""
    today = datetime.datetime.now(KST).date()

    # 기간 결정
    if any(k in text for k in ("이번주", "금주")):
        mon, sun = rpt.week_bounds(today)
        wb, _, _ = load_wb_for_date(f"{mon.month}/{mon.day}")
        period, mode = f"{mon.month}/{mon.day}~{sun.month}/{sun.day}", ("week", mon, sun)
    elif any(k in text for k in ("저번주", "지난주")):
        mon, sun = rpt.week_bounds(today - datetime.timedelta(days=7))
        wb, _, _ = load_wb_for_date(f"{mon.month}/{mon.day}")
        period, mode = f"{mon.month}/{mon.day}~{sun.month}/{sun.day}", ("week", mon, sun)
    else:
        mm = re.search(r"(\d{1,2})\s*월", text)
        month = (12 if today.month == 1 else today.month - 1) if "지난달" in text \
            else (int(mm.group(1)) if mm else today.month)
        wb, _, _ = load_wb_for_date(f"{month}/1")
        period, mode = f"{month}월", ("month", None, None)
    if wb is None:
        wb, _ = load_latest_wb()
    if wb is None:
        await update.message.reply_text("출석부 파일이 없어요.")
        return

    # 반 필터
    tokens = [t for t in text.split() if not re.search(r"주|달|월|통계|집계|현황|리포트", t)]
    want = None
    for t in tokens:
        m = next((s for s in wb.sheetnames if s.replace(" ", "") == t.replace(" ", "")), None)
        if m:
            want = m
            break
    sheets = [want] if want else [s for s in wb.sheetnames if ac.get_roster(wb[s])]

    def dates_for(sheet):
        if mode[0] == "week":
            return rpt.week_dates(wb[sheet], mode[1], mode[2], mode[2].year)
        return ac.sheet_dates(wb).get(sheet, [])

    # 집계
    agg = {}          # sheet -> per-student
    overall_abs, overall_hw = {}, {}
    cls_line = []
    for sheet in sheets:
        per = _collect_stats(wb[sheet], dates_for(sheet))
        agg[sheet] = per
        tot = sum(v["total"] for v in per.values())
        ab = sum(v["absent"] for v in per.values())
        rate = round((tot - ab) / tot * 100) if tot else None
        cls_line.append((sheet, rate, tot, ab))
        for nm, v in per.items():
            if v["absent"]:
                overall_abs[f"{nm}({sheet})"] = v["absent"]
            if v["hw_miss"]:
                overall_hw[f"{nm}({sheet})"] = v["hw_miss"]

    def _rank(d, n=6):
        items = sorted(d.items(), key=lambda x: -x[1])[:n]
        return ", ".join(f"{k} {v}회" for k, v in items) if items else "없음"

    if want:  # 특정 반: 자세히
        per = agg[want]
        tot = sum(v["total"] for v in per.values())
        ab = sum(v["absent"] for v in per.values())
        rate = round((tot - ab) / tot * 100) if tot else 0
        abs_r = _rank({nm: v["absent"] for nm, v in per.items() if v["absent"]})
        hw_r = _rank({nm: v["hw_miss"] for nm, v in per.items() if v["hw_miss"]})
        half = sum(v["hw_half"] for v in per.values())
        lines = [
            f"📊 <b>{want}</b> 통계 · {period}",
            f"• 출석률 <b>{rate}%</b> (기록 {tot}칸 · 결석 {ab})",
            f"• 결석: {abs_r}",
            f"• 과제 미제출: {hw_r}" + (f"  |  부분제출 {half}회" if half else ""),
        ]
    else:  # 전체 요약
        gtot = sum(l[2] for l in cls_line)
        gab = sum(l[3] for l in cls_line)
        grate = round((gtot - gab) / gtot * 100) if gtot else 0
        lines = [f"📊 <b>{SUBJ_NAME} 전체 통계</b> · {period}",
                 f"• 전체 출석률 <b>{grate}%</b> (기록 {gtot}칸 · 결석 {gab})",
                 "",
                 "<b>반별 출석률</b>"]
        for sheet, rate, tot, ab in sorted(cls_line, key=lambda x: (x[1] is None, x[1] or 0)):
            lines.append(f"• {sheet}: {rate}% ({ab} 결석)" if rate is not None else f"• {sheet}: 기록 없음")
        lines += ["", f"<b>결석 많은 학생</b>\n{_rank(overall_abs)}",
                  "", f"<b>과제 미제출</b>\n{_rank(overall_hw)}"]
    lines.append("\n특정 반: '통계 초5A' · 기간: '통계 이번주' / '통계 지난달'")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def cmd_change_date(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    """한 반의 날짜 하나를 다른 날짜로 바꾼다(칸 안의 기록은 그대로 유지).
    예) '고1 7/21 -> 7/22' / '고1 7/21일 날짜 7/22로 변경'"""
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    cands = _match_preview_sheet(text, latest.sheetnames)
    if not cands:
        await update.message.reply_text("어느 반인지 알려주세요. 예) 고1 7/21 → 7/22")
        return
    if len(cands) > 1:
        await update.message.reply_text("어느 반인지 콕 집어주세요: " + ", ".join(cands))
        return
    sheet = cands[0]
    ds = re.findall(r'(\d{1,2})\s*[/.월]\s*(\d{1,2})', text)
    if len(ds) < 2:
        await update.message.reply_text("바꿀 날짜 두 개를 알려주세요. 예) 고1 7/21 → 7/22")
        return
    old = f"{int(ds[0][0])}/{int(ds[0][1])}"
    new = f"{int(ds[1][0])}/{int(ds[1][1])}"
    wb, path, month = load_wb_for_date(old)
    if wb is None:
        await update.message.reply_text(f"{month}월 출석부 파일이 없어요.")
        return
    if sheet not in wb.sheetnames:
        await update.message.reply_text(f"'{sheet}' 반을 그 달 파일에서 못 찾았어요.")
        return
    ws = wb[sheet]
    top = ac.find_date_block(ws, old)
    if top is None:
        await update.message.reply_text(f"'{sheet}'에 {old} 날짜가 없어요. (미리보기로 확인해 보세요)")
        return
    if ac.find_date_block(ws, new) is not None:
        await update.message.reply_text(
            f"'{sheet}'에 이미 {new} 가 있어서 바꿀 수 없어요. 먼저 정리해 주세요.")
        return
    year = current_ym()[0]
    new_label = ac.date_label(new, year)
    try:
        ws.cell(top, 1).value = new_label
    except Exception as e:
        await update.message.reply_text(f"날짜 칸을 바꾸지 못했어요: {e}")
        return
    save_wb(wb, path, undoable=True, desc=f"{sheet} 날짜 {old}→{new}")
    await update.message.reply_text(
        f"✅ <b>{sheet}</b> {old} → <b>{new_label}</b> 로 바꿨어요.\n"
        "그 날짜 칸의 출석·숙제 기록은 그대로 유지돼요.\n"
        f"확인: <code>{sheet} {new} 미리보기</code> · 되돌리기: <code>실행취소</code>",
        parse_mode="HTML")


# ── 진도표 ──────────────────────────────────────────────────────
async def cmd_progress_units(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    """반 진도표의 단원/항목/단계 설정·조회.
    예) '초5A 진도단원 분수의 나눗셈, 각기둥과 각뿔' / '초5A 진도항목 개념서, 문제집'
        '초5A 진도단계 1차, 2차' / '초5A 진도설정'(전체 조회)"""
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    cands = _match_preview_sheet(text, latest.sheetnames)
    if not cands:
        await update.message.reply_text("어느 반인지 알려주세요. 예) 초5A 진도단원 분수의 나눗셈, 각기둥과 각뿔")
        return
    if len(cands) > 1:
        await update.message.reply_text("어느 반인지 콕 집어주세요: " + ", ".join(cands))
        return
    sheet = cands[0]
    # 어떤 항목을 설정하나
    if re.search(r'진도\s*항목', text):
        field, label = "items", "항목"
    elif re.search(r'진도\s*단계', text):
        field, label = "steps", "단계"
    elif re.search(r'진도\s*설정', text):
        field, label = None, None   # 전체 조회
    else:
        field, label = "units", "단원"

    def show_all():
        u, it, st = prog_cfg(sheet)
        ul = "\n".join(f"  {i+1}. {x}" for i, x in enumerate(u)) or "  (없음)"
        return (f"📚 <b>{sheet}</b> 진도표 설정\n"
                f"• 단원:\n{ul}\n"
                f"• 항목: {', '.join(it)}\n"
                f"• 단계: {', '.join(st)}")

    prog = load_progress()
    cur = prog.get(sheet, {})
    m = re.search(r'진도\s*(?:단원|항목|단계|설정)\s*(.*)$', text, re.S)
    body = (m.group(1) if m else "").strip()

    if field is None or not body:   # 조회
        if field in ("units",) and not prog_cfg(sheet)[0]:
            await update.message.reply_text(
                f"'{sheet}'는 아직 진도 단원이 없어요.\n"
                "예) 초5A 진도단원 분수의 나눗셈, 각기둥과 각뿔, 소수의 나눗셈")
            return
        await update.message.reply_text(show_all(), parse_mode="HTML")
        return

    vals = [re.sub(r'^\s*\d+\s*[.)]\s*', '', u).strip()
            for u in re.split(r'[,、\n]+', body) if u.strip()]
    if not vals:
        await update.message.reply_text(f"{label}을(를) 못 읽었어요.")
        return
    cur[field] = vals
    prog[sheet] = cur
    save_progress(prog)
    if field == "units":
        lines = "\n".join(f"{i+1}. {u}" for i, u in enumerate(vals))
        await update.message.reply_text(
            f"✅ <b>{sheet}</b> 진도 단원 {len(vals)}개 설정했어요.\n{lines}\n\n"
            "이제 <code>원서진 1단원 유형서 수정완료</code> 처럼 입력하고, "
            "<code>초5A 진도</code>로 표를 볼 수 있어요.", parse_mode="HTML")
    else:
        await update.message.reply_text(
            f"✅ <b>{sheet}</b> 진도 {label} 설정: {', '.join(vals)}\n"
            f"입력 예) <code>원서진 1단원 {vals[0]} "
            f"{prog_cfg(sheet)[2][0]}</code>", parse_mode="HTML")


def _prog_students(text, wb):
    """text에 언급된 학생들을 (이름, 반) 목록으로. 동명이인은 애매 목록으로 분리."""
    hits = _student_hits(text, wb, include_given=False) or _student_hits(text, wb, include_given=True)
    by_name = {}
    for nm, sh in hits:
        by_name.setdefault(nm, set()).add(sh)
    ok = [(nm, next(iter(shs))) for nm, shs in by_name.items() if len(shs) == 1]
    amb = [nm for nm, shs in by_name.items() if len(shs) > 1]
    return ok, amb


async def cmd_progress_mark(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    """진도 O/X 입력. 예) '원서진 1단원 유형서 수정완료' / '1단원 유형서 밴드완료 지훈, 규림 X'"""
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    students, amb = _prog_students(text, latest)
    if amb:
        await update.message.reply_text("동명이인이 있어요: " + ", ".join(amb) + " — 성까지 붙여 불러주세요.")
        return
    if not students:
        await update.message.reply_text("누구 진도인지 이름을 알려주세요. 예) 원서진 1단원 유형서 수정완료")
        return
    sheet = students[0][1]
    units, items, steps = prog_cfg(sheet)
    if not units:
        await update.message.reply_text(
            f"'{sheet}'는 진도 단원이 없어요. 먼저 설정해 주세요.\n"
            "예) 초5A 진도단원 분수의 나눗셈, 각기둥과 각뿔")
        return
    # 단원
    mu = re.search(r'(\d+)\s*단원', text) or re.search(r'(\d+)\s*과', text)
    ui = None
    if mu:
        ui = int(mu.group(1)) - 1
    else:  # 단원명으로도 찾기
        for i, u in enumerate(units):
            if u.replace(" ", "") in text.replace(" ", ""):
                ui = i
                break
    if ui is None or not (0 <= ui < len(units)):
        await update.message.reply_text(f"몇 단원인지 알려주세요 (1~{len(units)}). 예) 1단원")
        return
    # 항목 (긴 것 먼저, 그다음 줄임말 심화/평가/유형)
    item = next((it for it in sorted(items, key=len, reverse=True) if it in text), None)
    if item is None:
        if "심화" in text:
            item = next((it for it in items if "심화" in it), None)
        elif "평가" in text:
            item = next((it for it in items if "평가" in it), None)
        elif "유형" in text:
            item = next((it for it in items if "유형" in it and "심화" not in it), None)
    if item is None:
        await update.message.reply_text("항목을 알려주세요: " + " / ".join(items))
        return
    # 단계 (수정/밴드 토큰; 없으면 항목 전체)
    tgt_steps = [st for st in steps if any(tok in text for tok in (st, st[:2]))]
    if not tgt_steps:
        tgt_steps = steps[:]   # 예: '유형서 완료' → 수정·밴드 둘 다
    # 값
    if any(k in text for k in ("취소", "삭제", "지워", "없애", "빼줘")):
        val = None
    elif re.search(r'(?<![A-Za-z])[Xx](?![A-Za-z])', text) or "엑스" in text:
        val = "X"
    else:
        val = "O"
    prog = load_progress()
    cur = prog.setdefault(sheet, {})
    cells = cur.setdefault("cells", {})
    done = []
    for nm, sh in students:
        cell = cells.setdefault(nm, {})
        for st in tgt_steps:
            key = f"{ui}|{item}|{st}"
            if val is None:
                cell.pop(key, None)
            else:
                cell[key] = val
        done.append(nm)
    save_progress(prog)
    mark = "지움" if val is None else val
    await update.message.reply_text(
        f"✅ {units[ui]} · {item} · {'/'.join(tgt_steps)} → <b>{mark}</b>\n"
        f"({', '.join(done)}) — {len(done)}명\n"
        f"표 보기: <code>{sheet} 진도</code>", parse_mode="HTML")


async def cmd_progress_view(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    """진도표 이미지 전송. '초5A 진도'(반 전체) / '원서진 진도'(그 학생만)."""
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    name, sheets, amb = _find_student(text, latest)
    if amb:
        await update.message.reply_text("동명이인이 있어요: " + ", ".join(amb) + " — 성까지 붙여 불러주세요.")
        return
    prog = load_progress()

    async def send(sheet, only=None):
        units, items, steps = prog_cfg(sheet)
        if not units:
            await update.message.reply_text(
                f"'{sheet}'는 아직 진도 단원이 없어요.\n"
                "예) " + sheet + " 진도단원 분수의 나눗셈, 각기둥과 각뿔")
            return False
        roster = list(ac.get_roster(latest[sheet]).keys()) if sheet in latest.sheetnames else []
        cells = (prog.get(sheet, {}) or {}).get("cells", {})
        stu = [only] if only else roster
        if not stu:
            await update.message.reply_text(f"'{sheet}'에 학생이 없어요.")
            return False
        img = pg.render_progress(sheet, units, stu, cells, items, steps)
        bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        cap = f"📚 {sheet} 진도표" + (f" · {only}" if only else "")
        await update.message.reply_photo(photo=bio, caption=cap)
        return True

    if name:  # 학생 개별 (여러 반이면 반마다)
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="upload_photo")
        for sh in sheets:
            await send(sh, only=name)
        return
    cands = _match_preview_sheet(text, latest.sheetnames)
    if not cands:
        await update.message.reply_text(
            "누구/어느 반 진도인지 알려주세요. 예) 초5A 진도 / 원서진 진도")
        return
    if len(cands) > 1:
        await update.message.reply_text("어느 반인지 콕 집어주세요: " + ", ".join(cands))
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="upload_photo")
    await send(cands[0])


async def cmd_teacher(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """반별 담당 지정. 담당으로 지정된 반의 출석 알림은 그 선생님에게만 간다."""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    teachers = load_teachers()
    known = set(load_times()) | set(load_schedules())

    # 인자 분류: 반이름들 + 해제 여부
    remove, classes = False, []
    for tok in context.args:
        if tok in ("해제", "빼기", "제거", "삭제", "off", "취소"):
            remove = True
        else:
            classes.append(tok)

    if not context.args:  # 현재 담당 현황
        mine = sorted(c for c, ids in teachers.items() if chat_id in ids)
        lines = ["👤 내가 담당하는 반: " + (", ".join(mine) if mine else "없음")]
        if teachers:
            lines.append("\n전체 담당 현황:")
            for c in sorted(teachers):
                lines.append(f"• {c}: {len(teachers[c])}명 담당")
            unassigned = sorted(known - set(teachers))
            if unassigned:
                lines.append("\n담당 미지정(알림 안 감): " + ", ".join(unassigned))
        lines.append("\n지정) 담당 초5 중2    해제) 담당 초5 빼기    전체해제) 담당 해제")
        await update.message.reply_text("\n".join(lines))
        return

    if remove and not classes:  # 모든 담당에서 빠지기
        for c in list(teachers):
            if chat_id in teachers[c]:
                teachers[c].remove(chat_id)
                if not teachers[c]:
                    teachers.pop(c)
        save_teachers(teachers)
        await update.message.reply_text("✅ 모든 담당 반에서 빠졌어요. 이제 알림을 받지 않습니다.")
        return

    unknown = [c for c in classes if c not in known]
    if unknown:
        await update.message.reply_text(
            f"'{', '.join(unknown)}' 반을 못 찾겠어요.\n등록된 반: {', '.join(sorted(known))}"
        )
        return

    if remove:  # 특정 반 담당에서 빠지기
        for c in classes:
            if c in teachers and chat_id in teachers[c]:
                teachers[c].remove(chat_id)
                if not teachers[c]:
                    teachers.pop(c)
        save_teachers(teachers)
        await update.message.reply_text(f"✅ {', '.join(classes)} 담당에서 빠졌어요.")
        return

    for c in classes:  # 담당 지정
        ids = teachers.get(c, [])
        if chat_id not in ids:
            ids.append(chat_id)
        teachers[c] = ids
    save_teachers(teachers)
    await update.message.reply_text(
        f"✅ 이제 {', '.join(classes)} 반의 출석 알림을 받으실 거예요.\n"
        f"(다른 반 알림은 그 반 담당 선생님에게만 갑니다.)"
    )


async def require_admin(update: Update) -> bool:
    """관리자면 True. 아니면 안내 메시지 보내고 False."""
    if is_admin(update.effective_chat.id):
        return True
    if get_admin() is None:
        await update.message.reply_text(
            "아직 관리자가 없어요. 주인(선생님)이 '관리자등록' 이라고 보내 등록해 주세요."
        )
    else:
        await update.message.reply_text("🔒 이 기능은 관리자만 쓸 수 있어요.")
    return False


async def cmd_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """관리자 등록. 아직 관리자가 없을 때 처음 보낸 사람이 관리자가 된다."""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    a = get_admin()
    if a is None:
        set_admin(chat_id)
        await update.message.reply_text(
            "✅ 관리자로 등록됐어요.\n\n"
            "관리자 전용 기능:\n"
            "• '설정초기화', '주간보고서' — 관리자만\n"
            "• 승인 관리: 다른 쌤이 처음 쓰려 하면 승인 요청이 와요.\n"
            "   - <b>승인</b> : 대기 목록 보기 / <b>승인 이름</b> 또는 <b>승인 번호</b> : 승인\n"
            "   - <b>박탈 이름/번호</b> : 사용 권한 회수 / <b>멤버</b> : 승인된 명단",
            parse_mode="HTML",
        )
    elif a == chat_id:
        await update.message.reply_text("이미 관리자로 등록돼 있어요. 👍")
    else:
        await update.message.reply_text("이미 다른 분이 관리자로 등록돼 있어요.")


async def gate_member(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """승인된 사용자만 통과. 미승인이면 대기 등록·관리자 알림 후 False."""
    chat_id = update.effective_chat.id
    if is_approved(chat_id):
        return True
    name = update.effective_user.full_name if update.effective_user else None
    newly = register_pending(chat_id, name)
    if newly:
        admin = get_admin()
        if admin:
            try:
                await context.bot.send_message(
                    admin,
                    f"🔔 사용 승인 요청: <b>{name or '이름미상'}</b> (id {chat_id})\n"
                    f"허용: <code>승인 {chat_id}</code>  또는  <code>승인 {name or ''}</code>",
                    parse_mode="HTML",
                )
            except Exception as e:
                log.warning("승인요청 알림 실패: %s", e)
        await update.message.reply_text(
            "이 봇은 <b>승인된 선생님만</b> 사용할 수 있어요. 🔒\n"
            "관리자에게 승인을 요청했으니 잠시만 기다려 주세요.",
            parse_mode="HTML",
        )
    else:
        await update.message.reply_text("관리자 승인 대기 중이에요. 조금만 기다려 주세요. 🙏")
    return False


async def cmd_approve(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """관리자: 대기 중인 사용자 승인. 인자 없으면 대기목록."""
    if not await require_admin(update):
        return
    m = load_members()
    if not context.args:
        pend = [(k, v["name"]) for k, v in m.items() if v.get("status") == "pending"]
        if not pend:
            await update.message.reply_text("승인 대기 중인 사람이 없어요.")
            return
        lines = ["⏳ 승인 대기 목록:"]
        for k, nm in pend:
            lines.append(f"• {nm} (id {k}) → 승인하려면: 승인 {k}")
        await update.message.reply_text("\n".join(lines))
        return
    key = _find_member_key(m, " ".join(context.args))
    if not key or m.get(key, {}).get("status") is None:
        await update.message.reply_text("그런 대기자를 못 찾았어요. '승인' 만 보내 목록을 확인하세요.")
        return
    m[key]["status"] = "approved"
    save_members(m)
    await update.message.reply_text(f"✅ {m[key]['name']} 님을 승인했어요. 이제 사용할 수 있어요.")
    try:
        await context.bot.send_message(int(key), "✅ 관리자 승인이 완료됐어요! 이제 출석부 봇을 쓸 수 있어요. 😊")
    except Exception:
        pass


async def cmd_revoke(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """관리자: 사용 권한 박탈."""
    if not await require_admin(update):
        return
    m = load_members()
    if not context.args:
        appr = [(k, v["name"]) for k, v in m.items() if v.get("status") == "approved"]
        lines = ["👥 승인된 멤버:"] + [f"• {nm} (id {k})" for k, nm in appr] if appr else ["승인된 멤버가 없어요."]
        lines.append("\n박탈하려면: 박탈 <이름 또는 번호>")
        await update.message.reply_text("\n".join(lines))
        return
    key = _find_member_key(m, " ".join(context.args))
    if not key:
        await update.message.reply_text("그런 멤버를 못 찾았어요.")
        return
    nm = m[key]["name"]
    m.pop(key)
    save_members(m)
    await update.message.reply_text(f"🚫 {nm} 님의 사용 권한을 박탈했어요.")
    try:
        await context.bot.send_message(int(key), "안내: 출석부 봇 사용 권한이 해제되었어요.")
    except Exception:
        pass


async def cmd_members(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """관리자: 승인/대기 멤버 목록."""
    if not await require_admin(update):
        return
    m = load_members()
    appr = [v["name"] for v in m.values() if v.get("status") == "approved"]
    pend = [(k, v["name"]) for k, v in m.items() if v.get("status") == "pending"]
    lines = [f"👥 승인된 멤버 ({len(appr)}): " + (", ".join(appr) if appr else "없음")]
    if pend:
        lines.append("\n⏳ 승인 대기:")
        for k, nm in pend:
            lines.append(f"• {nm} → 승인 {k}")
    await update.message.reply_text("\n".join(lines))


async def cmd_reset_config(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """반 목록·수업요일·알림 시간표를 코드 기본값으로 다시 맞춘다.
    (반 구성이 바뀐 뒤 한 번 실행. 담당 지정은 그대로 둔다.)"""
    if not await require_admin(update):
        return
    remember_chat(update.effective_chat.id)
    for p in (schedules_path(), times_path()):
        if os.path.exists(p):
            os.remove(p)
    load_schedules()
    load_times()
    classes = ", ".join(sorted(load_times()))
    await update.message.reply_text(
        "✅ 반 목록·수업요일·알림 시간표를 기본값으로 새로 맞췄어요.\n"
        f"등록된 반: {classes}\n\n"
        "이제 담당 반을 지정하세요. 예) 담당 초5A 초6A 중1AB 중2 중3 고1 고2(미적분) 고3"
    )


def parse_month_token(text):
    """'26.08', '26-8', '8월', '2026-08' 등에서 (year, month) 추출. 없으면 None."""
    t = str(text)
    m = re.search(r"(\d{2,4})[.\-/](\d{1,2})", t)
    if m:
        y = int(m.group(1))
        y = y if y > 100 else 2000 + y
        return y, int(m.group(2))
    m = re.search(r"(\d{1,2})\s*월", t)
    if m:
        return datetime.date.today().year, int(m.group(1))
    return None


async def cmd_download(update: Update, context: ContextTypes.DEFAULT_TYPE):
    remember_chat(update.effective_chat.id)
    token = " ".join(context.args) if context.args else ""
    ym = parse_month_token(token)
    if ym:
        y, mth = ym
    else:
        y, mth = current_ym()  # 월 지정 없으면 이번 달
    path, fname = month_path(y, mth), month_filename(y, mth)
    if not os.path.exists(path):
        avail = ", ".join(f.replace(" 출석부.xlsx", "") for f in list_month_files()) or "없음"
        await update.message.reply_text(f"'{fname}' 파일이 없어요.\n보유: {avail}")
        return
    with open(path, "rb") as f:
        await update.message.reply_document(document=f, filename=fname)


async def cmd_generate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    y, m = current_ym()
    if os.path.exists(month_path(y, m)):
        await update.message.reply_text(
            f"{month_filename(y, m)} 는 이미 있어요. 덮어쓰지 않았습니다. (받으려면 /출석부)"
        )
        return
    try:
        out = generate_month_file(y, m)
    except FileNotFoundError as e:
        await update.message.reply_text(str(e))
        return
    await update.message.reply_text(f"✅ {os.path.basename(out)} 를 만들었어요.")
    with open(out, "rb") as f:
        await update.message.reply_document(document=f, filename=os.path.basename(out))


# ── 파일 업로드(시드) ──────────────────────────────────────────
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await gate_member(update, context):
        return
    doc = update.message.document
    if not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("xlsx 파일만 저장할 수 있어요.")
        return
    tgfile = await doc.get_file()
    data = bytes(await tgfile.download_as_bytearray())
    try:
        wb = ac.load_workbook(data)
        dates = ac.sheet_dates(wb)
        # 첫 날짜로 연·월 추정
        first = next((d[0] for d in dates.values() if d), None)
        y = datetime.date.today().year
        mth = int(first.split("/")[0]) if first else datetime.date.today().month
    except Exception as e:
        await update.message.reply_text(f"파일을 읽지 못했어요: {e}")
        return
    fname = month_filename(y, mth)
    with open(os.path.join(DATA_DIR, fname), "wb") as f:
        f.write(data)
    remember_chat(update.effective_chat.id)
    await update.message.reply_text(
        f"✅ 저장했어요: {fname}\n이제 출석을 자유롭게 입력하시면 됩니다."
    )


# ── 신규개강 / 종강 (반 단위) ──────────────────────────────────
async def start_opening(update: Update, context: ContextTypes.DEFAULT_TYPE, sheet):
    chat_id = update.effective_chat.id
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부 파일이 없어요. 먼저 파일을 보내거나 /생성 해주세요.")
        return
    if sheet in wb.sheetnames:
        # 이름을 띄어쓰기로 넣어 한 칸에 뭉쳐버린 경우 → 각 학생 열로 쪼개 복구
        roster = list(ac.get_roster(wb[sheet]).keys())
        crammed = [n for n in roster if any(c.isspace() for c in n)]
        if crammed:
            names = []
            for n in roster:
                for one in n.split():
                    if one not in names:
                        names.append(one)
            ac.set_roster(wb, sheet, names)
            save_wb(wb, path)
            await update.message.reply_text(
                f"🔧 '{sheet}' 명단이 한 칸에 뭉쳐 있어 {len(names)}명으로 나눠 고쳤어요.\n"
                f"• {', '.join(names)}\n이제 출석 입력이 될 거예요.",
            )
            return
        await update.message.reply_text(f"'{sheet}' 반은 이미 있어요. 새로 만들 필요 없어요.")
        return
    opening_flow[chat_id] = {"sheet": sheet, "step": "weekdays"}
    await update.message.reply_text(
        f"🆕 <b>{sheet}</b> 반을 개강할게요. (취소하려면 '취소')\n\n"
        "1️⃣ 수업 요일을 알려주세요. 예) <code>월수금</code>",
        parse_mode="HTML",
    )


async def handle_opening_step(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    chat_id = update.effective_chat.id
    st = opening_flow[chat_id]
    if text.strip() in ("취소", "그만", "cancel"):
        opening_flow.pop(chat_id)
        await update.message.reply_text("개강을 취소했어요.")
        return
    step = st["step"]
    if step == "weekdays":
        wds = parse_weekdays(text)
        if not wds:
            await update.message.reply_text("요일을 못 알아들었어요. 예) 월수금")
            return
        st["weekdays"] = wds
        st["step"] = "hours"
        await update.message.reply_text(
            f"요일: {'·'.join(WD[i] for i in wds)}\n\n"
            "2️⃣ 수업 시간을 알려주세요. 시간표에 이대로 올라가요.\n"
            "예) <code>15:00~16:00</code>  (나중에 정하려면 '없음')",
            parse_mode="HTML",
        )
    elif step == "hours":
        if text.strip() in ("없음", "없어", "나중에", "스킵", "건너뛰기", "skip"):
            st["hours"] = None
        else:
            hr = parse_hours_token(text)
            if not hr:
                await update.message.reply_text(
                    "수업 시간을 못 알아들었어요. 예) 15:00~16:00 · 오후 3시~4시  (또는 '없음')"
                )
                return
            st["hours"] = hr
        st["step"] = "time"
        suggest = hours_end_plus(st["hours"]) if st["hours"] else None
        tip = f"\n그냥 <b>확인</b>만 보내면 수업 종료 +15분({suggest})으로 할게요." if suggest else ""
        await update.message.reply_text(
            f"3️⃣ 출석 미입력 알림을 언제 보낼까요? 예) <code>저녁 9시</code>  (필요 없으면 '없음'){tip}",
            parse_mode="HTML",
        )
    elif step == "time":
        suggest = hours_end_plus(st["hours"]) if st.get("hours") else None
        if text.strip() in ("없음", "없어", "스킵", "건너뛰기", "skip"):
            st["time"] = None
        elif suggest and text.strip().startswith("확인"):
            st["time"] = suggest
        else:
            hhmm = parse_time_token(text)
            if not hhmm:
                await update.message.reply_text("시간을 못 알아들었어요. 예) 저녁 9시  (또는 '없음')")
                return
            st["time"] = hhmm
        st["step"] = "students"
        await update.message.reply_text(
            "4️⃣ 학생 명단을 알려주세요. 예) <code>김철수, 이영희, 박민수</code>  (없으면 '없음')",
            parse_mode="HTML",
        )
    elif step == "students":
        if text.strip() in ("없음", "없어", "나중에"):
            st["students"] = []
        else:
            # 쉼표·줄바꿈뿐 아니라 띄어쓰기로 나열해도 각각 한 명으로 인식
            # (한글 이름엔 내부 공백이 없으므로 공백 분리가 안전)
            st["students"] = [n.strip() for n in re.split(r"[,\n\s]+", text) if n.strip()]
        st["step"] = "confirm"
        wd = "·".join(WD[i] for i in st["weekdays"])
        hr = st.get("hours") or "없음 (나중에)"
        tm = st["time"] or "없음"
        stu = ", ".join(st["students"]) if st["students"] else "(없음 — 나중에 신규등록)"
        await update.message.reply_text(
            f"이렇게 개강할게요:\n"
            f"• 반: <b>{st['sheet']}</b>\n• 요일: {wd}\n• 수업 시간: {hr}\n• 알림: {tm}\n• 학생: {stu}\n\n"
            "맞으면 <b>확인</b>, 아니면 <b>취소</b>",
            parse_mode="HTML",
        )
    elif step == "confirm":
        if not text.strip().startswith("확인"):
            await update.message.reply_text("'확인' 또는 '취소' 라고 보내주세요.")
            return
        opening_flow.pop(chat_id)
        wb, path = load_current_wb()
        if wb is None:
            await update.message.reply_text("이번 달 파일이 없어요.")
            return
        today = datetime.datetime.now(KST).date()
        try:
            ac.create_class_sheet(wb, st["sheet"], st["students"], st["weekdays"], today.year, today.month)
        except ValueError as e:
            await update.message.reply_text(f"⚠️ {e}")
            return
        save_wb(wb, path)
        sch = load_schedules(); sch[st["sheet"]] = st["weekdays"]; save_schedules(sch)
        if st["time"]:
            tms = load_times(); tms[st["sheet"]] = {str(d): st["time"] for d in st["weekdays"]}; save_times(tms)
        if st.get("hours"):
            hrs = load_hours(); hrs[st["sheet"]] = {str(d): st["hours"] for d in st["weekdays"]}; save_hours(hrs)
        cl = load_closed()
        if cl.pop(st["sheet"], None) is not None:
            save_closed(cl)
        await update.message.reply_text(
            f"✅ <b>{st['sheet']}</b> 개강 완료! 이제 출석 입력이 가능해요.\n"
            f"⚠️ 아직 <b>담당 선생님이 없어요.</b> 담당쌤이 <code>담당 {st['sheet']}</code> 라고 보내 지정해 주세요.",
            parse_mode="HTML",
        )
        admin = get_admin()
        if admin and admin != chat_id:
            try:
                await context.bot.send_message(
                    admin,
                    f"🆕 '{st['sheet']}' 반이 개강됐어요. 담당 지정이 필요해요 (담당쌤이 '담당 {st['sheet']}').",
                )
            except Exception:
                pass


async def start_closing(update: Update, context: ContextTypes.DEFAULT_TYPE, sheet):
    chat_id = update.effective_chat.id
    wb, _ = load_current_wb()
    if wb is None or sheet not in wb.sheetnames:
        await update.message.reply_text(f"'{sheet}' 반을 못 찾았어요.")
        return
    closing_flow[chat_id] = sheet
    await update.message.reply_text(
        f"🛑 <b>{sheet}</b> 반을 종강할까요?\n"
        "이번 주까지는 입력되고, <b>다음 주부터 입력 중단</b> + <b>다음 달 파일엔 제외</b>돼요.\n"
        "(이번 달 기록은 그대로 남아요)\n\n맞으면 <b>확인</b>, 아니면 <b>취소</b>",
        parse_mode="HTML",
    )


async def handle_closing_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    chat_id = update.effective_chat.id
    sheet = closing_flow[chat_id]
    if not text.strip().startswith("확인"):
        closing_flow.pop(chat_id)
        await update.message.reply_text("종강을 취소했어요.")
        return
    closing_flow.pop(chat_id)
    today = datetime.datetime.now(KST).date()
    cl = load_closed()
    cl[sheet] = f"{today.month}/{today.day}"
    save_closed(cl)
    await update.message.reply_text(
        f"✅ <b>{sheet}</b> 종강 처리했어요.\n다음 주부터 입력이 멈추고, 다음 달 파일엔 빠집니다.",
        parse_mode="HTML",
    )


async def start_stray_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """이번 달 파일에서 명백히 잘못된 날짜 블록을 찾아 보여주고 삭제 확인을 받는다."""
    chat_id = update.effective_chat.id
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부가 아직 없어요.")
        return
    _, month = current_ym()
    empty, kept = [], []
    for sheet in wb.sheetnames:
        for s in ac.find_stray_blocks(wb[sheet], month=month):
            (kept if s['has_data'] else empty).append((sheet, s))
    if not empty and not kept:
        await update.message.reply_text("✅ 이상한 날짜 없어요. 전부 깔끔합니다.")
        return

    lines = []
    if empty:
        lines.append("🧹 <b>지울 수 있는 날짜</b> (기록이 비어 있어요)")
        lines += [f"• {sh} — <b>{s['date']}</b> ({s['reason']})" for sh, s in empty]
    if kept:
        lines.append("\n⚠️ <b>기록이 있어서 손대지 않아요</b> — 직접 확인해 주세요")
        lines += [f"• {sh} — <b>{s['date']}</b> ({s['reason']})" for sh, s in kept]
    if not empty:
        await update.message.reply_text("\n".join(lines), parse_mode="HTML")
        return

    stray_flow[chat_id] = [(sh, s['top'], s['date']) for sh, s in empty]
    lines.append(f"\n비어 있는 {len(empty)}건을 지울까요?\n"
                 "맞으면 <b>확인</b>, 아니면 <b>취소</b>")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def handle_stray_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    chat_id = update.effective_chat.id
    targets = stray_flow.pop(chat_id)
    if not text.strip().startswith("확인"):
        await update.message.reply_text("정리를 취소했어요.")
        return
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부를 못 찾았어요.")
        return
    # 파일이 그새 바뀌었을 수 있으니 지우기 직전에 다시 확인한다.
    done, skipped = [], []
    for sheet, top, date in targets:
        ws = wb[sheet]
        last_col = ac._last_col(ws, ac._find_start_row(ws))
        cur = ws.cell(top, 1).value
        if str(cur).strip() != date or ac.block_has_data(ws, top, last_col):
            skipped.append(f"{sheet} {date}")
            continue
        ac.remove_block(ws, top)
        done.append(f"{sheet} {date}")
    if done:
        save_wb(wb, path)
    msg = f"✅ {len(done)}건 정리했어요: {', '.join(done)}" if done else "지운 게 없어요."
    if skipped:
        msg += f"\n⚠️ 그새 내용이 생겨서 건너뛴 것: {', '.join(skipped)}"
    await update.message.reply_text(msg)


# ── 휴강 지정 ─────────────────────────────────────────────────
def parse_day_token(tok):
    """'7/20' '7.20' '오늘' '내일' '어제' → 'M/D'. 못 읽으면 None."""
    tok = (tok or "").strip()
    today = datetime.datetime.now(KST).date()
    delta = {"오늘": 0, "내일": 1, "모레": 2, "어제": -1}.get(tok)
    if delta is not None:
        d = today + datetime.timedelta(days=delta)
        return f"{d.month}/{d.day}"
    m = re.match(r"^(\d{1,2})[/.](\d{1,2})$", tok)
    return f"{int(m.group(1))}/{int(m.group(2))}" if m else None


def holiday_text(reason):
    """출석부에 적을 문구. 사유가 있으면 '휴강(폭우)' 처럼."""
    reason = (reason or "").strip()
    return f"휴강({reason})" if reason else "휴강"


async def start_holiday(update: Update, context: ContextTypes.DEFAULT_TYPE, date_str, reason=None):
    """그 날짜에 수업이 있는 모든 반을 휴강으로 지정 — 확인을 받고 적용한다."""
    chat_id = update.effective_chat.id
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부가 아직 없어요.")
        return
    targets, already, has_data = [], [], []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        top = ac.find_date_block(ws, date_str)
        if top is None:
            continue  # 그 반은 그 날 수업이 없음
        last_col = ac._last_col(ws, ac._find_start_row(ws))
        if ac._is_holiday_block(ws, top, last_col):
            already.append(sheet)
            continue
        targets.append((sheet, top))
        if ac.block_has_data(ws, top, last_col):
            has_data.append(sheet)
    if not targets:
        msg = f"<b>{date_str}</b> 에 수업 있는 반이 없어요."
        if already:
            msg = f"<b>{date_str}</b> 은 이미 휴강이에요. ({', '.join(already)})"
        await update.message.reply_text(msg, parse_mode="HTML")
        return

    text = holiday_text(reason)
    lines = [f"🚫 <b>{date_str}</b> 을 <b>{text}</b> 으로 지정할까요?",
             f"대상 {len(targets)}개 반: {', '.join(sh for sh, _ in targets)}"]
    if already:
        lines.append(f"(이미 휴강인 반은 그대로 둬요: {', '.join(already)})")
    if has_data:
        lines.append(f"\n⚠️ <b>이미 입력된 내용이 지워져요</b>: {', '.join(has_data)}")
    lines.append("\n맞으면 <b>확인</b>, 아니면 <b>취소</b>")
    holiday_flow[chat_id] = {"date": date_str, "targets": targets, "text": text}
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def handle_holiday_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE, text):
    chat_id = update.effective_chat.id
    st = holiday_flow.pop(chat_id)
    if not text.strip().startswith("확인"):
        await update.message.reply_text("휴강 지정을 취소했어요.")
        return
    date_str = st["date"]
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부를 못 찾았어요.")
        return
    done = []
    for sheet, _top in st["targets"]:
        ws = wb[sheet]
        top = ac.find_date_block(ws, date_str)  # 그새 파일이 바뀌었을 수 있어 다시 찾는다
        if top is None:
            continue
        last_col = ac._last_col(ws, ac._find_start_row(ws))
        ac.set_holiday_block(ws, top, last_col, st["text"])
        done.append(sheet)
    if done:
        save_wb(wb, path)
    await update.message.reply_text(
        f"✅ <b>{date_str}</b> <b>{st['text']}</b> 으로 지정했어요. ({', '.join(done)})\n"
        "그 날은 출석 알림도 안 가요.", parse_mode="HTML")


async def undo_holiday(update: Update, context: ContextTypes.DEFAULT_TYPE, date_str):
    """휴강 지정을 되돌린다. 휴강 블록에는 기록이 없으므로 확인 없이 바로 처리."""
    wb, path = load_current_wb()
    if wb is None:
        await update.message.reply_text("이번 달 출석부가 아직 없어요.")
        return
    done, failed, was = [], [], {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        top = ac.find_date_block(ws, date_str)
        if top is None:
            continue
        last_col = ac._last_col(ws, ac._find_start_row(ws))
        if not ac._is_holiday_block(ws, top, last_col):
            continue
        was[sheet] = ws.cell(top, ac.STUDENT_FIRST_COL).value
        (done if ac.clear_holiday_block(ws, top, last_col) else failed).append(sheet)
    if not done and not failed:
        await update.message.reply_text(
            f"<b>{date_str}</b> 은 휴강이 아니에요.", parse_mode="HTML")
        return
    if done:
        save_wb(wb, path)
    names = {v for v in was.values() if v}
    msg = f"✅ <b>{date_str}</b> 휴강을 취소했어요. ({', '.join(done)})"
    if names - {"휴강"}:
        msg += f"\n(원래 표시: {', '.join(names)})"
    msg += "\n칸은 비어 있으니 출석을 새로 입력해 주세요."
    if failed:
        msg += f"\n⚠️ 되돌릴 기준 블록이 없어 실패: {', '.join(failed)}"
    await update.message.reply_text(msg, parse_mode="HTML")


# ── 일반 텍스트 ────────────────────────────────────────────────
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    text = update.message.text.strip()
    # 처음 말 거는 쌤에게는 사용 안내를 공지처럼 먼저 보낸다
    first_contact = chat_id not in known_chats()
    remember_chat(chat_id)
    if first_contact:
        await update.message.reply_text(GUIDE, parse_mode="HTML")

    low = deglue(text.lstrip("/").strip())  # 음성·오타 띄어쓰기 보정 (명령 인식용)
    parts = low.split()
    kw = parts[0] if parts else ""

    # 승인된 멤버만 사용 가능 (관리자 등록만 예외로 통과)
    if low not in ("관리자등록", "관리자", "관리자설정") and not await gate_member(update, context):
        return

    # 개강/종강 진행 중이면 그 단계 처리
    if chat_id in opening_flow:
        return await handle_opening_step(update, context, text)
    if chat_id in closing_flow:
        return await handle_closing_confirm(update, context, text)
    if chat_id in stray_flow:
        return await handle_stray_confirm(update, context, text)
    if chat_id in holiday_flow:
        return await handle_holiday_confirm(update, context, text)

    # 문장 전체가 딱 이 낱말일 때만 — '정리 안 한 학생…' 같은 평범한 말이
    # 점검을 띄우고 다음 메시지까지 확인 응답으로 먹는 걸 막는다.
    if low in ("점검", "정리", "날짜점검"):
        return await start_stray_check(update, context)

    # 휴강 취소 — 반드시 휴강 지정보다 먼저 본다
    mo = re.match(r"^(?:휴강\s*취소\s+(\S+)|(\S+)\s+휴강\s*취소)$", low)
    if mo:
        day = parse_day_token(mo.group(1) or mo.group(2))
        if day is None:
            await update.message.reply_text(
                "날짜를 못 읽었어요. 예) <code>7/20 휴강취소</code>", parse_mode="HTML")
            return
        return await undo_holiday(update, context, day)

    # 휴강 지정 (예: '7/20 휴강', '휴강 오늘') — 학원 전체 휴강이라 반을 안 받는다
    if low == "휴강":
        await update.message.reply_text(
            "어느 날짜를 휴강으로 할까요? 예) <code>7/20 휴강</code>, <code>오늘 휴강</code>",
            parse_mode="HTML")
        return
    # 사유는 선택 — '7/20 휴강 폭우' 처럼 뒤에 붙인다
    mo = re.match(r"^(?:휴강\s+(\S+)(?:\s+(.+))?|(\S+)\s+휴강(?:\s+(.+))?)$", low)
    if mo:
        day = parse_day_token(mo.group(1) or mo.group(3))
        reason = (mo.group(2) or mo.group(4) or "").strip()
        if day is None:
            await update.message.reply_text(
                "날짜를 못 읽었어요. 예) <code>7/20 휴강</code>, <code>오늘 휴강 폭우</code>",
                parse_mode="HTML")
            return
        if reason == "취소":  # '휴강 7/20 취소' — 위 취소 정규식이 못 잡는 어순
            return await undo_holiday(update, context, day)
        return await start_holiday(update, context, day, reason)

    # 신규개강 / 종강 트리거 (어순 무관: '고3B 신규개강' = '신규개강 고3B')
    mo = re.match(r"^(?:(?:신규개강|개강)\s+(.+)|(.+?)\s*(?:신규개강|개강))$", low)
    if mo:
        return await start_opening(update, context, (mo.group(1) or mo.group(2)).strip())
    mc = re.match(r"^(?:종강\s+(.+)|(.+?)\s*종강)$", low)
    if mc:
        return await start_closing(update, context, (mc.group(1) or mc.group(2)).strip())

    # 확인/취소 대기 처리 (확인 뒤에 추가 정보가 붙어도 인식하고 병합)
    if chat_id in pending:
        confirm_words = ("확인", "네", "응", "ㅇㅇ", "ㅇㅋ", "ok", "예", "오케이", "저장")
        head = re.sub(r"[.,!~\s]+$", "", kw.lower())
        if head in confirm_words or low.startswith("확인"):
            info = pending.pop(chat_id)
            extra = re.sub(r"^(확인|네|응|예|오케이|저장|ok)\s*[.,!~]*\s*", "", low, flags=re.I).strip()
            if extra:  # '확인. 수업내용은 …' → 추가 정보 병합
                pwb, _ = load_latest_wb()
                if pwb is not None:
                    merged = await parse_message(f"{info['sheet']} {info['date']} {extra}", pwb)
                    if merged.get("type") == "attendance":
                        for key in ("출석", "수업내용", "과제수행", "다음과제", "비고"):
                            v = merged.get(key)
                            if not v:
                                continue
                            if isinstance(v, dict) and isinstance(info["data"].get(key), dict):
                                info["data"][key].update(v)
                            else:
                                info["data"][key] = v
            wb, path, _ = load_wb_for_date(info["date"])
            if wb is None:
                await update.message.reply_text("해당 달 출석부 파일이 없어요. 파일을 보내거나 /생성 해주세요.")
                return
            enroll = load_enroll().get(info["sheet"], {})
            try:
                written, warnings = ac.write_attendance(
                    wb, info["sheet"], info["date"], info["data"], enroll=enroll
                )
                save_wb(wb, path, undoable=True, desc=f"{info['sheet']} {info['date']} 입력")
                life = info["data"].get("학적") or {}
                if life:  # 학적 변동을 재적 기록에 반영
                    apply_enroll_events(info["sheet"], info["date"], life)
            except Exception as e:
                log.error("기록 저장 실패: %s: %s | info=%s", type(e).__name__, e, info)
                await update.message.reply_text(
                    "⚠️ 확인은 됐는데 <b>기록 저장에 실패</b>했어요.\n"
                    f"• 반/날짜: {info['sheet']} / {info['date']}\n"
                    f"• 원인: {type(e).__name__}: {str(e)[:200]}\n"
                    "다시 시도해 주세요. 계속되면 이 메시지를 저(관리자)에게 알려주세요.",
                    parse_mode="HTML",
                )
                return
            msg = f"✅ 기록 완료 ({info['sheet']} {info['date']}) — {len(written)}건\n" + "\n".join(
                "  • " + w for w in written
            )
            if not written:
                msg += ("\n(기록된 항목이 없어요 — 이름이 명단과 다르거나, "
                        "그 반·날짜에 해당하는 내용이 없을 수 있어요.)")
            if warnings:
                msg += "\n⚠️ " + " / ".join(warnings)
            await update.message.reply_text(msg)
            # 기록된 내용을 사진으로 한 장 보내 정확히 확인하게 함
            try:
                img = rpt.render_class_table(info["sheet"], wb[info["sheet"]], [info["date"]])
                bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
                await update.message.reply_photo(
                    photo=bio, caption=f"📷 {info['sheet']} · {info['date']} 기록 내용"
                )
            except Exception as e:
                log.warning("기록 사진 전송 실패: %s", e)
            return
        if head in ("취소", "아니", "아니오", "no", "cancel"):
            pending.pop(chat_id)
            await update.message.reply_text("취소했어요.")
            return
        # 그 외 입력은 아래에서 새로 처리 (기존 대기는 덮어씀)

    # 한글 키워드를 명령처럼 처리 (슬래시 있어도/없어도)
    if kw in ("시간표", "수업시간표", "주간시간표"):   # '시간표 초5A'
        context.args = parts[1:]
        return await cmd_timetable(update, context)
    _mtt = re.match(r"^(.*?)\s*(?:수업\s*)?시간표$", low)   # '초5A 시간표', '내시간표'
    if _mtt:
        pre = _mtt.group(1).strip()
        context.args = pre.split() if pre else []
        return await cmd_timetable(update, context)
    if kw in ("일정", "스케줄"):
        context.args = parts[1:]
        return await cmd_schedule(update, context)
    if kw in ("알림", "알람", "리마인더"):
        context.args = parts[1:]
        return await cmd_remind(update, context)
    if kw in ("담당", "내반", "담당반"):
        context.args = parts[1:]
        return await cmd_teacher(update, context)
    if low in ("관리자등록", "관리자", "관리자설정"):
        return await cmd_admin(update, context)
    if kw in ("승인",):
        context.args = parts[1:]
        return await cmd_approve(update, context)
    if kw in ("박탈", "권한박탈", "차단"):
        context.args = parts[1:]
        return await cmd_revoke(update, context)
    if low in ("멤버", "명단", "멤버목록", "대기", "대기목록"):
        return await cmd_members(update, context)
    if low in ("요일추가", "날짜요일", "요일넣기"):
        return await cmd_add_weekday(update, context)
    if low in ("실행취소", "되돌리기", "방금취소", "입력취소", "취소하기", "복원"):
        return await cmd_undo(update, context)
    if low in ("통계", "집계", "현황", "리포트"):
        context.args = parts[1:]
        return await cmd_stats(update, context)
    if low in ("설정초기화", "기본설정복원", "반목록갱신"):
        return await cmd_reset_config(update, context)
    if low in ("주간보고서", "보고서", "주간출결"):
        return await cmd_report(update, context)
    # ── 날짜 변경: '고1 7/21 → 7/22' / '고1 7/21일 날짜 7/22로 변경' ──
    _dtoks = re.findall(r'\d{1,2}\s*[/.월]\s*\d{1,2}', text)
    if len(_dtoks) >= 2 and (
        re.search(r'\d\s*(?:->|→|~>|=>)\s*\d', text)
        or any(k in low for k in ("날짜", "변경", "바꿔", "바꾸"))):
        return await cmd_change_date(update, context, text)
    # ── 진도표 ──
    if any(k in low for k in ("진도단원", "진도 단원", "진도항목", "진도단계", "진도설정")):
        return await cmd_progress_units(update, context, text)
    # 진도 입력: 단계어(수정/밴드/완료)가 있고 + (N단원 or 항목)이 있으면 O/X 마킹
    # (그냥 '단원평가 90점' 같은 시험 입력과 헷갈리지 않게 단계어를 필수로)
    # 단, '유형서 수정'을 수업내용으로 적은 출석 입력이 진도로 새지 않게 출결 신호가
    # 있으면 진도 라우팅을 건너뛴다(진도 입력엔 이런 단어가 안 나온다).
    _attend_signal = any(k in text for k in (
        "결석", "지각", "조퇴", "외출", "무단", "출석", "미지참", "미수령", "미제출",
        "다음과제", "다음 숙제", "다음숙제"))
    if not _attend_signal and any(k in text for k in ("수정", "밴드", "완료")):
        _has_unit = re.search(r"\d+\s*단원", text)
        _p_item = any(k in text for k in ("유형", "심화", "단원평가"))
        if not _p_item:   # 반별 커스텀 항목도 인식
            for v in load_progress().values():
                if any(it and it in text for it in (v.get("items") or [])):
                    _p_item = True
                    break
        if _has_unit or _p_item:
            return await cmd_progress_mark(update, context, text)
    # 진도 조회: '초5A 진도' / '원서진 진도' (딱 그 형태일 때만)
    if re.fullmatch(r"\s*.+?\s*진도표?\s*", text) and "진도" in low:
        return await cmd_progress_view(update, context, text)
    # '원서진 7월 15일 과제' — 어순 무관하게 과제 조회 ('과제'/'숙제' 단독 단어 + 날짜)
    # 단, '다음과제 51쪽' 같은 출석 입력과 구분: 짧고 상태어(출석/결석/O/X/점/쪽 등)가 없을 때만
    _hw_word = any(re.fullmatch(r"(?:과제|숙제)[?？!요]*", p) for p in parts)
    _looks_query = len(parts) <= 5 and not re.search(
        r"결석|지각|조퇴|출석|수업|무단|미지참|미수령|다음과제|과제수행|\d+쪽|\d+점|[OoXx]", low)
    if _hw_word and _looks_query and (
        _resolve_preview_date(low) is not None
        or any(k in low for k in ("이번주", "저번주", "지난주", "이번달"))
    ):
        return await cmd_homework(update, context)
    if any(k in low for k in ("미리보기", "캡처", "보여줘", "보여주", "이미지로")):
        return await cmd_preview(update, context)
    # '남우현 7월' / '중1AB 이번주' / '초5 7/15' 처럼 (학생|반)+기간, 순서 무관하게 미리보기로
    _period = (r'(?:\d{1,2}\s*월|이번달|지난달|이번주|저번주|지난주|오늘|어제|\d{1,2}[/.]\d{1,2})'
               r'\s*(?:미리보기|보여줘|보여주세요|캡처|이미지로?)?')
    _cls = r'(?:초|중|고)\d[a-zA-Z]*(?:\([^)]*\))?'
    _stu = r'[가-힣]{2,4}'
    if any(re.fullmatch(a + r'\s+' + b, low)
           for a, b in ((_stu, _period), (_period, _stu), (_cls, _period), (_period, _cls))):
        return await cmd_preview(update, context)
    if low in ("생성", "새달", "새달생성"):
        return await cmd_generate(update, context)
    if low in ("시작", "도움말", "도움"):
        return await start(update, context)
    # '출석부' 다운로드 요청 (예: '출석부', '출석부 7월', '7월 출석부 확인')
    if "출석부" in low and not re.match(r"^(초|중|고)\d", low):
        context.args = parts
        return await cmd_download(update, context)

    parse_wb, _ = load_latest_wb()
    if parse_wb is None:
        await update.message.reply_text(
            "출석부 파일이 아직 없어요.\n출석부 xlsx 파일을 봇에게 보내거나 /생성 을 눌러주세요."
        )
        return

    try:
        await context.bot.send_chat_action(chat_id=chat_id, action="typing")
    except Exception:
        pass
    try:
        parsed = await parse_message(text, parse_wb)
    except Exception as e:
        log.error("parse_message 실패: %s: %s", type(e).__name__, e)
        detail = str(e)[:300]
        hint = ""
        if "credit balance" in detail.lower() or "too low" in detail.lower():
            hint = ("\n💳 Anthropic API 크레딧이 부족한 것 같아요. "
                    "console.anthropic.com → Billing 에서 충전하면 바로 됩니다.")
        await update.message.reply_text(
            "⚠️ 지금 입력을 처리하지 못했어요 (AI 응답 실패).\n"
            f"• 원인: {type(e).__name__}\n"
            f"• 상세: {detail}{hint}"
        )
        return

    if parsed.get("type") != "attendance":
        try:
            reply = await chat_reply(chat_id, text)
        except Exception as e:
            log.error("chat_reply 실패: %s: %s", type(e).__name__, e)
            reply = f"지금은 답변을 못 드리겠어요. 잠시 후 다시 시도해 주세요. (원인: {type(e).__name__})"
        await update.message.reply_text(reply)
        return

    # 날짜의 달에 해당하는 파일을 대상으로 (예외가 나도 조용히 죽지 않게 감쌈)
    try:
        target_wb, _, month = load_wb_for_date(parsed.get("date", "0/0"))
        if target_wb is None:
            await update.message.reply_text(
                f"{month}월 출석부 파일이 없어요. 먼저 파일을 보내거나 /생성 해주세요."
            )
            return
        if class_input_blocked(parsed.get("sheet"), parsed.get("date", "")):
            await update.message.reply_text(
                f"🛑 '{parsed.get('sheet')}' 반은 종강해서 그 날짜는 입력할 수 없어요."
            )
            return
        preview, err = build_preview(target_wb, parsed)
    except Exception as e:
        log.error("미리보기 생성 실패: %s: %s | parsed=%s", type(e).__name__, e, parsed)
        await update.message.reply_text(
            "⚠️ 입력을 처리하다 문제가 생겼어요.\n"
            f"• 반/날짜: {parsed.get('sheet')} / {parsed.get('date')}\n"
            f"• 원인: {type(e).__name__}\n"
            "• '초5 오늘 다 왔어'처럼 간단히 시작해 보시거나, 문장을 조금 바꿔 다시 보내주세요."
        )
        return
    if err:
        await update.message.reply_text("⚠️ " + err)
        return
    pending[chat_id] = {"sheet": parsed["sheet"], "date": parsed["date"], "data": parsed}
    await update.message.reply_text(preview, parse_mode="HTML")


# ── 출석 입력 알림 ─────────────────────────────────────────────
# 오늘 이미 확인/발송한 반 기록: {(YYYY-MM-DD, 반이름)}  (중복 방지)
_sent_reminders: set = set()


async def reminder_job(context: ContextTypes.DEFAULT_TYPE):
    """1분마다 확인 → 각 반의 수업일·설정시각(KST)이 되면 출석부를 점검해서,
    아직 출석 입력이 안 됐을 때만 알림을 보낸다. 하루 한 번, 반별로.
    시각을 1~2분 놓쳐도 보내도록 120초 창을 둔다."""
    now = datetime.datetime.now(KST)
    today = now.date().isoformat()
    # 지난 날짜 기록 정리
    for k in [k for k in _sent_reminders if k[0] != today]:
        _sent_reminders.discard(k)

    wd = str(now.weekday())  # '0'=월 … '6'=일
    wb, _ = load_current_wb()  # 이번 달 출석부 (없으면 None)
    date_str = f"{now.month}/{now.day}"
    # 지금 확인 시각이 된 반들 추리기 (요일별 표에서 오늘 요일의 시각을 찾음)
    due = []
    for cls, table in load_times().items():
        if not isinstance(table, dict):
            continue
        if class_alarm_blocked(cls, date_str):
            continue  # 종강한 반은 종강일부터 알림 보내지 않음
        tm = table.get(wd)
        if not tm:
            # 요일엔 없지만 오늘 날짜 블록이 시트에 있으면(보강·날짜변경 등) 대체 시각으로 확인
            if wb is None or cls not in wb.sheetnames:
                continue
            if ac.find_date_block(wb[cls], date_str) is None:
                continue
            tm = next((v for v in table.values() if v), None)  # 이 반의 아무 알림 시각
            if not tm:
                continue
        try:
            hh, mm = (int(x) for x in str(tm).split(":"))
        except (ValueError, AttributeError):
            continue
        sched = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        delta = (now - sched).total_seconds()
        key = (today, cls)
        if 0 <= delta < 120 and key not in _sent_reminders:
            due.append((cls, key))
    if not due:
        return

    teachers = load_teachers()
    for cls, key in due:
        _sent_reminders.add(key)  # 오늘 이 반은 확인 완료 (재확인·중복발송 방지)
        recorded = None
        if wb is not None and cls in wb.sheetnames:
            recorded = ac.attendance_recorded(wb[cls], date_str)
        if recorded is False:  # 블록은 있는데 출석이 비어 있음 → 입력 안 함
            # 담당 지정이 하나라도 있으면 그 반 담당에게만(미지정=발송 안 함),
            # 아직 담당을 아무도 안 정했으면 전체에게(초기 편의)
            if teachers:
                recipients = teachers.get(cls, [])
            else:
                recipients = known_chats()
            for chat_id in recipients:
                try:
                    await context.bot.send_message(
                        chat_id,
                        f"📝 {cls} 오늘({date_str}) 출석부를 아직 입력 안 하셨어요. 잊지 마세요!",
                    )
                except Exception as e:
                    log.warning("알림 전송 실패 %s: %s", chat_id, e)


# ── 매월 1일 자동 생성 ─────────────────────────────────────────
async def monthly_job(context: ContextTypes.DEFAULT_TYPE):
    """이번 달 파일이 없으면(=달이 바뀌면) 지난달 바탕으로 생성하고 알림.
    시간대에 무관하게 '이번 달 파일 유무'로 판단하므로 매일 돌려도 안전하다."""
    y, m = current_ym()
    if os.path.exists(month_path(y, m)):
        return
    if not latest_month_file():  # 바탕이 될 파일이 하나도 없으면 조용히 대기
        return
    try:
        out = generate_month_file(y, m)
    except FileNotFoundError as e:
        log.warning("자동 생성 보류: %s", e)
        return
    for chat_id in known_chats():
        try:
            await context.bot.send_message(chat_id, f"📄 {os.path.basename(out)} 를 새로 만들었어요!")
            with open(out, "rb") as f:
                await context.bot.send_document(chat_id, document=f, filename=os.path.basename(out))
        except Exception as e:
            log.warning("알림 실패 %s: %s", chat_id, e)


# ── 밀린(지난 날) 미입력 출석 알림 ────────────────────────────
def scan_backlog(wb, today):
    """각 반에서 '오늘 이전' 수업일 중 출석이 아직 안 채워진 날짜. {반: ['M/D',...]}"""
    out = {}
    dates_by = ac.sheet_dates(wb)
    for cls in wb.sheetnames:
        ws = wb[cls]
        miss = []
        for ds in dates_by.get(cls, []):
            try:
                m, d = map(int, ds.split('/'))
                dd = datetime.date(today.year, m, d)
            except ValueError:
                continue
            if dd < BACKLOG_SINCE:
                continue  # 기준일 이전 옛 날짜는 알림 제외
            if dd >= today:
                continue  # 오늘·미래는 당일 알림(reminder_job) 담당
            if class_alarm_blocked(cls, ds):
                continue  # 종강한 반은 종강일부터 밀린 알림도 제외
            if ac.attendance_recorded(ws, ds) is False:
                miss.append(ds)
        if miss:
            out[cls] = miss
    return out


_last_backlog = {"day": None}


async def backlog_job(context: ContextTypes.DEFAULT_TYPE):
    """매일 13:00(KST) — 지난 수업일 중 미입력이 있으면 입력할 때까지 하루 한 번 알림."""
    now = datetime.datetime.now(KST)
    if not (now.hour == 13 and now.minute < 3):
        return
    key = now.date().isoformat()
    if _last_backlog["day"] == key:
        return
    _last_backlog["day"] = key
    wb, _ = load_current_wb()
    if wb is None:
        return
    backlog = scan_backlog(wb, now.date())
    if not backlog:
        return
    teachers = load_teachers()
    for cls, dates in backlog.items():
        recipients = (teachers.get(cls, []) if teachers else known_chats())
        ktxt = ", ".join(f"{int(x.split('/')[0])}월 {int(x.split('/')[1])}일" for x in dates)
        msg = f"📌 {cls} · 아직 출석 입력이 안 된 날이 있어요: {ktxt}\n출석부를 입력해주세요!"
        for chat_id in recipients:
            try:
                await context.bot.send_message(chat_id, msg)
            except Exception as e:
                log.warning("미입력 알림 실패 %s: %s", chat_id, e)


# ── 주간 출결 보고서 ───────────────────────────────────────────
def generate_weekly_report(target=None):
    """target(기본 오늘)이 속한 주(월~일) 보고서 PDF 생성. 반환: (경로, 반수)|(None,0)."""
    today = target or datetime.datetime.now(KST).date()
    monday, sunday = rpt.week_bounds(today)
    p = month_path(sunday.year, sunday.month)  # 일요일이 속한 달 파일
    if not os.path.exists(p):
        f = latest_month_file()
        if not f:
            return None, 0
        p = os.path.join(DATA_DIR, f)
    wb = ac.load_workbook(p)
    fname = f"{SUBJ_NAME} 주간 출결사항 ({monday.month}.{monday.day}~{sunday.month}.{sunday.day}).pdf"
    out = os.path.join(DATA_DIR, fname)
    n = rpt.build_report_pdf(wb, out, monday, sunday, sunday.year,
                             title=f"<{SUBJ_NAME} 주간 출결사항>", enroll=load_enroll())
    return (out, n) if n else (None, 0)


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    if not await require_admin(update):
        return
    await context.bot.send_chat_action(chat_id=chat_id, action="upload_document")
    out, n = generate_weekly_report()
    if not out:
        await update.message.reply_text("이번 주 출결 데이터가 있는 출석부 파일을 찾지 못했어요.")
        return
    with open(out, "rb") as f:
        await update.message.reply_document(
            document=f, filename=os.path.basename(out),
            caption=f"📄 주간 출결 보고서 · {n}개 반"
        )


# ── 특정 날짜·반 미리보기(이미지) ──────────────────────────────
def _match_preview_sheet(text, sheets):
    """텍스트에서 반 시트명 찾기. 반환: 후보 리스트(정확히 1개면 확정)."""
    t = text.replace(" ", "")
    for s in sorted(sheets, key=len, reverse=True):  # '초5A'가 '초5'보다 먼저
        if s.replace(" ", "") in t:
            return [s]
    m = re.search(r'(초|중|고)\d[A-Za-z]*', t)  # '고2' → '고2(미적분)', '중1' → 여러 개
    if m:
        tok = m.group(0)
        return [s for s in sheets if s.replace(" ", "").startswith(tok)]
    return []


def _resolve_preview_date(text):
    """텍스트에서 날짜 추출. 오늘/어제/내일/'M월 D일'/'M/D' 지원."""
    today = datetime.datetime.now(KST).date()
    if '오늘' in text:
        return today
    if '어제' in text:
        return today - datetime.timedelta(days=1)
    if '내일' in text:
        return today + datetime.timedelta(days=1)
    # '이번주 월요일' / '저번주 수요일' / '월요일' 등 (주 + 요일)
    wm = re.search(r'([월화수목금토일])\s*요일', text)
    if wm:
        wi = WD.index(wm.group(1))
        base = today
        if any(k in text for k in ('저번주', '지난주', '전주')):
            base = today - datetime.timedelta(days=7)
        elif '다음주' in text or '담주' in text:
            base = today + datetime.timedelta(days=7)
        monday = base - datetime.timedelta(days=base.weekday())
        return monday + datetime.timedelta(days=wi)
    m = re.search(r'(\d{1,2})\s*월\s*(\d{1,2})', text) or re.search(r'(\d{1,2})[/.](\d{1,2})', text)
    if m:
        try:
            return datetime.date(today.year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def _name_givens(nm):
    """성 뗀 이름 후보. '원서진'→{'서진'}, '남궁민수'→{'궁민수','민수'}."""
    g = set()
    if len(nm) >= 3:
        g.add(nm[1:])
    if len(nm) >= 4:
        g.add(nm[2:])
    return g


def _student_hits(text, wb, include_given):
    """text가 가리키는 학생 (nm, sheet) 목록.
    풀네임/풀네임+'이'를 먼저 보고, include_given이면 성 뗀 이름(+'이')도 본다."""
    hits = []
    for sheet in wb.sheetnames:
        for nm in ac.get_roster(wb[sheet]):
            if not nm:
                continue
            ok = (nm in text) or ((nm + '이') in text)
            if not ok and include_given:
                ok = any(g in text or (g + '이') in text for g in _name_givens(nm))
            if ok:
                hits.append((nm, sheet))
    return hits


def _find_student(text, wb):
    """텍스트의 학생 이름과 그 학생이 속한 모든 반. (name, [sheets], ambiguous).
    성 떼거나 '~이'로 불러도 인식. 동명이인이면 (None, [], [이름들])."""
    hits = _student_hits(text, wb, include_given=False)   # 풀네임 우선
    if not hits:
        hits = _student_hits(text, wb, include_given=True)  # 성 뗀 이름
    names = {nm for nm, _ in hits}
    if not names:
        return None, [], []
    if len(names) > 1:                                     # 동명이인 → 애매
        return None, [], sorted(names)
    name = next(iter(names))
    return name, [s for nm, s in hits if nm == name], []


async def _preview_student(update, context, name, sheets, text):
    """한 학생의 기간(월/주/날짜) 출결을 이미지로 보낸다. 여러 반이면 반마다 한 장씩."""
    chat_id = update.effective_chat.id
    today = datetime.datetime.now(KST).date()

    # 기간 결정 + 파일 로드 ('이번주 월요일'처럼 요일이 콕 집히면 그날 하루)
    _has_weekday = re.search(r'[월화수목금토일]\s*요일', text)
    if not _has_weekday and any(k in text for k in ("이번주", "금주")):
        mon, sun = rpt.week_bounds(today); mode = "week"
    elif not _has_weekday and any(k in text for k in ("저번주", "지난주")):
        mon, sun = rpt.week_bounds(today - datetime.timedelta(days=7)); mode = "week"
    else:
        mon = sun = None; mode = None

    if mode == "week":
        wb, _, _ = load_wb_for_date(f"{mon.month}/{mon.day}")
        period = f"{mon.month}/{mon.day}~{sun.month}/{sun.day}"
    else:
        d = _resolve_preview_date(text)
        if d is not None:
            ds = f"{d.month}/{d.day}"
            wb, _, mo = load_wb_for_date(ds); mode = "date"
            if wb is None:
                await update.message.reply_text(f"{mo}월 출석부 파일이 없어요.")
                return
            period = ds
        else:
            mm = re.search(r'(\d{1,2})\s*월', text)
            month = (12 if today.month == 1 else today.month - 1) if '지난달' in text \
                else (int(mm.group(1)) if mm else today.month)
            wb, _, _ = load_wb_for_date(f"{month}/1"); mode = "month"
            period = f"{month}월"
    if wb is None:
        wb, _ = load_latest_wb()
    if wb is None:
        await update.message.reply_text("출석부 파일이 없어요.")
        return

    def dates_for(sheet):
        if mode == "week":
            return rpt.week_dates(wb[sheet], mon, sun, sun.year)
        if mode == "date":
            ds = period
            return [ds] if ac.find_date_block(wb[sheet], ds) else []
        return ac.sheet_dates(wb).get(sheet, [])

    await context.bot.send_chat_action(chat_id=chat_id, action="upload_photo")
    sent = 0
    for sheet in sheets:
        if sheet not in wb.sheetnames:
            continue
        dates = dates_for(sheet)
        if not dates:
            continue
        img = rpt.render_student_table(name, wb[sheet], dates)
        if img is None:
            continue
        bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        await update.message.reply_photo(photo=bio, caption=f"📷 {name} · {sheet} · {period}")
        sent += 1
    if sent == 0:
        await update.message.reply_text(f"{name} 학생의 해당 기간 수업일이 없어요.")


async def cmd_homework(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """'7월 15일 원서진 과제' → 그 학생의 그날 과제수행·다음과제를 챗으로 알려준다."""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    text = update.message.text
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    name, sheets, amb = _find_student(text, latest)
    if amb:
        await update.message.reply_text("여러 명이에요: " + ", ".join(amb) + " — 성까지 붙여 불러주세요.")
        return
    if not name:
        await update.message.reply_text("누구 과제인지 이름을 알려주세요. 예) 7월 15일 원서진 과제")
        return
    d = _resolve_preview_date(text) or datetime.datetime.now(KST).date()
    ds = f"{d.month}/{d.day}"
    wb, _, _ = load_wb_for_date(ds)
    if wb is None:
        wb = latest
    lines = [f"📚 <b>{name}</b> · {ds} 과제"]
    for sheet in sheets:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        top = ac.find_date_block(ws, ds)
        col = ac.get_roster(ws).get(name)
        if top is None or col is None:
            lines.append(f"[{sheet}] 그 날 수업이 없어요.")
            continue
        did = ws.cell(top + ac.ROW_OFFSET["과제수행"], col).value
        nxt = ws.cell(top + ac.ROW_OFFSET["다음과제"], ac.STUDENT_FIRST_COL).value
        lines.append(f"[<b>{sheet}</b>] 과제수행: {did or '-'} / 다음과제: {nxt or '-'}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def cmd_preview(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """'초5 7월 15일 미리보기'(반+날짜) 또는 '남우현 7월'(학생+기간)을 이미지로 보낸다."""
    chat_id = update.effective_chat.id
    remember_chat(chat_id)
    text = update.message.text
    latest, _ = load_latest_wb()
    if latest is None:
        await update.message.reply_text("아직 출석부 파일이 없어요.")
        return
    # 학생 이름이 있으면 학생별 미리보기 우선 (여러 반이면 반마다)
    sname, ssheets, samb = _find_student(text, latest)
    if samb:
        await update.message.reply_text("여러 명이에요: " + ", ".join(samb) + " — 성까지 붙여 불러주세요.")
        return
    if sname:
        return await _preview_student(update, context, sname, ssheets, text)
    cands = _match_preview_sheet(text, latest.sheetnames)
    if not cands:
        await update.message.reply_text(
            "그 반을 못 찾았어요. 출석부에 있는 반: " + ", ".join(latest.sheetnames)
            + "\n예) 초5 7월 15일 미리보기"
        )
        return
    if len(cands) > 1:
        await update.message.reply_text("어느 반인지 콕 집어주세요: " + ", ".join(cands))
        return
    sheet = cands[0]

    # 이번주/저번주면 그 주 전체, 아니면 특정 날짜 (요일이 콕 집히면 그날 하루)
    base = None
    _has_weekday = re.search(r'[월화수목금토일]\s*요일', text)
    if not _has_weekday and any(k in text for k in ("이번주", "금주")):
        base = datetime.datetime.now(KST).date()
    elif not _has_weekday and any(k in text for k in ("저번주", "지난주")):
        base = datetime.datetime.now(KST).date() - datetime.timedelta(days=7)

    if base is not None:
        mon, sun = rpt.week_bounds(base)
        wb, _, _ = load_wb_for_date(f"{mon.month}/{mon.day}")
        if wb is None:
            wb, _ = load_latest_wb()
        ws = wb[sheet]
        dates = rpt.week_dates(ws, mon, sun, sun.year)
        cap = f"📷 {sheet} · {mon.month}/{mon.day}~{sun.month}/{sun.day}"
    else:
        d = _resolve_preview_date(text)
        if d is None:
            await update.message.reply_text(
                "날짜를 알려주세요. 예) 초5 7월 15일 미리보기 / 초5 오늘 미리보기"
            )
            return
        ds = f"{d.month}/{d.day}"
        wb, _, month = load_wb_for_date(ds)
        if wb is None:
            await update.message.reply_text(f"{month}월 출석부 파일이 없어요.")
            return
        ws = wb[sheet]
        dates = [ds] if ac.find_date_block(ws, ds) else []
        cap = f"📷 {sheet} · {ds}"

    if not dates:
        avail = ", ".join(ac.sheet_dates(wb).get(sheet, [])) or "없음"
        await update.message.reply_text(f"{sheet}에 그 날짜 수업이 없어요.\n가능한 날짜: {avail}")
        return

    await context.bot.send_chat_action(chat_id=chat_id, action="upload_photo")
    img = rpt.render_class_table(sheet, ws, dates)
    bio = io.BytesIO()
    img.save(bio, "PNG")
    bio.seek(0)
    await update.message.reply_photo(photo=bio, caption=cap)


# 이번 주 보고서를 이미 보냈는지 (ISO 주 기준)
_last_report_week = {"key": None}


async def report_job(context: ContextTypes.DEFAULT_TYPE):
    """매주 일요일 18:00(KST)에 주간 출결 보고서를 만들어 전송."""
    now = datetime.datetime.now(KST)
    if now.weekday() != 6 or not (now.hour == 18 and now.minute < 3):
        return
    iso = now.isocalendar()
    key = (iso[0], iso[1])
    if _last_report_week["key"] == key:
        return
    _last_report_week["key"] = key
    out, n = generate_weekly_report(now.date())
    if not out:
        log.info("주간 보고서: 이번 주 데이터 없음")
        return
    to = get_admin() or get_report_to()  # 관리자에게만(없으면 기존 수신자)
    recipients = [to] if to else known_chats()
    for chat_id in recipients:
        try:
            with open(out, "rb") as f:
                await context.bot.send_document(
                    chat_id, document=f, filename=os.path.basename(out),
                    caption=f"📄 이번 주 출결 보고서 · {n}개 반"
                )
        except Exception as e:
            log.warning("주간 보고서 전송 실패 %s: %s", chat_id, e)


# ── 실행 ──────────────────────────────────────────────────────
def main():
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    # 텔레그램은 한글 슬래시명령을 지원하지 않으므로 ASCII 명령만 등록하고,
    # 한글 키워드(출석부/일정/생성)는 아래 handle_text 안에서 처리한다.
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("download", cmd_download))
    app.add_handler(CommandHandler("schedule", cmd_schedule))
    app.add_handler(CommandHandler("remind", cmd_remind))
    app.add_handler(CommandHandler("teacher", cmd_teacher))
    app.add_handler(CommandHandler("resetconfig", cmd_reset_config))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("generate", cmd_generate))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    if app.job_queue is not None:
        # 매일 확인 + 시작 직후 1회 확인 (달이 바뀌면 새 파일 생성)
        app.job_queue.run_daily(monthly_job, time=datetime.time(0, 5))
        app.job_queue.run_once(monthly_job, when=10)
        # 1분마다 출석 입력 알림 확인 (한국시간 기준)
        app.job_queue.run_repeating(reminder_job, interval=60, first=15)
        # 매주 일요일 18:00(KST) 주간 출결 보고서
        app.job_queue.run_repeating(report_job, interval=60, first=25)
        # 매일 13:00(KST) 밀린 미입력 출석 알림
        app.job_queue.run_repeating(backlog_job, interval=60, first=35)
        log.info("월간 생성 + 출석/미입력 알림 + 주간 보고서 스케줄 등록됨")
    else:
        log.warning("job_queue 미설치 — 자동 생성/알림 비활성 (requirements 확인)")

    log.info("출석부 봇 시작")
    app.run_polling()


if __name__ == "__main__":
    main()
