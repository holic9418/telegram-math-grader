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

# ── 설정 ───────────────────────────────────────────────────────
load_dotenv()
logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s | %(message)s", level=logging.INFO
)
log = logging.getLogger("attendance-bot")

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-4-8")
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
# 잡담용 짧은 기억: {chat_id: [messages]}
chat_history: dict[int, list] = {}


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
    save_schedules(ac.DEFAULT_SCHEDULES)
    return dict(ac.DEFAULT_SCHEDULES)


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
    default = {k: dict(v) for k, v in ac.DEFAULT_CLASS_TIMES.items()}
    save_times(default)
    return default


def save_times(t):
    with open(times_path(), "w", encoding="utf-8") as f:
        json.dump(t, f, ensure_ascii=False, indent=2)


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


def apply_enroll_events(sheet, date_str, life):
    """학적 이벤트를 enrollments.json 에 반영. life={학생: 유형}."""
    e = load_enroll()
    cls = e.setdefault(sheet, {})
    for st, event in life.items():
        info = cls.setdefault(st, {"from": None, "to": None})
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
- 수학 개념 질문 등에는 평소처럼 친절히 답해도 된다. 답은 간결하게."""


async def chat_reply(chat_id, text):
    """출석 입력이 아닌 메시지는 봇 성격을 아는 상태로 답한다."""
    history = chat_history.get(chat_id, [])
    history.append({"role": "user", "content": text})
    resp = await claude.messages.create(
        model=CLAUDE_MODEL, max_tokens=800, system=CHAT_SYSTEM, messages=history
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
            if name not in roster:
                warnings.append(name)
        lines.append("• 출석: " + ", ".join(parts))
    if parsed.get("수업내용"):
        lines.append(f"• 수업내용: {parsed['수업내용']}")
    hw = parsed.get("과제수행") or {}
    if hw:
        shown = {n: v for n, v in hw.items() if n not in absent}
        if shown:
            lines.append("• 과제수행: " + ", ".join(f"{n} {v}" for n, v in shown.items()))
        warnings += [n for n in hw if n not in roster]
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
        warnings += [n for n in bg if n not in roster]
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
GUIDE = """📋 <b>Zest 수학과 출석부 봇 사용법</b>

Zest 수학과 선생님들이 함께 쓰는 출석부예요.
<b>선생님이 채팅으로 말씀해 주시면, 봇이 공용 출석부 파일에 자동으로 기록</b>합니다.
(선생님은 엑셀을 직접 안 여셔도 되고, 봇에게 말로 알려주기만 하면 돼요.)

<b>1) 출석 알려주기</b> — 그냥 편하게 문장으로 말씀해 주세요.
예) <code>초5 오늘 남우현 결석, 나머지 출석. 수업 분수나눗셈, 다음과제 42쪽</code>
→ 봇이 "이렇게 기록할게요" 미리보기를 보여드려요. <b>확인</b> 이라고 답하시면 파일에 기록됩니다. (아니면 <b>취소</b>)
→ 이미 기록한 뒤에도 <b>특정 칸만 고쳐달라</b>고 하시면 그 칸만 다시 수정할 수 있어요. 예) <code>초5 7/15 남우현만 출석으로 바꿔줘</code>

<b>2) 이렇게 말씀하시면, 파일엔 이렇게 기록돼요</b>
• 결석: <code>OO 결석</code> → 파일엔 <code>X (사유)</code>. 사유 말하시면 그대로, 없으면 <code>무단</code>으로 입력됩니다.
• 지각: <code>OO 병원가서 30분 늦음</code> → 파일엔 <code>30분(병원)</code>
• 조퇴: <code>OO 휴가로 20분 조퇴</code> → 파일엔 <code>20분 조퇴(휴가)</code> (조퇴도 사유가 같이 기록돼요)
• 숙제: 안 했으면 <code>X</code>, 절반은 <code>50%</code>로 기록돼요. 사유 입력이 필요할 때, 안 가져왔으면 <code>미지참</code>, 안 받아갔으면 <code>미수령</code> 처럼 간략히 말씀해 주세요!
• 시험: 시험 이름과 학생·점수를 말씀해 주시면, 파일에 자동으로 기록됩니다.
   예) <code>초5 오늘 일일테스트 우현 규림 둘다 75점</code>

<b>3) 파일 받기</b>
• <b>출석부</b> 라고 보내면 이번 달 파일을, <b>출석부 7월</b> 이면 특정 달 파일을 보내드려요.

<b>4) 내 담당 반 지정</b> (중요!)
• <b>담당 초5 중2</b> 처럼 보내주시면, 그 반 알림만 받게 됩니다.
• 반 목록: 초3 · 초4 · 초5 · 초5A · 초6 · 초6A · 중1AB · 중1C · 중1보충 · 중2 · 중3 · 고1 · 고2(미적분) · 고3

<b>5) 봇이 보내는 알림</b>
• 수업이 끝나고 <b>15분 뒤까지 출석부 미기입시</b> 담당 쌤에게 알림을 보내드려요.
• 알림 시각을 바꾸려면 <code>알림 초5 21:00</code> 처럼 보내주세요.
• <b>일정</b> : 반별 수업 요일 보기·변경 (예: <code>일정 고1 화목토</code>)
• 매일 1시에 밀린 미입력 출석도 한 번 더 알려드려요.

궁금하면 아무 때나 <b>도움말</b> 이라고 보내주시면 이 안내가 다시 떠요. 🙂

※ 이 채팅은 <b>출석부 기록용 1:1 개인 채팅</b>이라, 여기에 쓰신 내용은 <b>다른 쌤들은 볼 수 없어요.</b> 편하게 작성하시면 됩니다."""


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
            "✅ 관리자로 등록됐어요.\n이제 '설정초기화'와 '주간보고서'는 관리자(당신)만 사용/수신합니다."
        )
    elif a == chat_id:
        await update.message.reply_text("이미 관리자로 등록돼 있어요. 👍")
    else:
        await update.message.reply_text("이미 다른 분이 관리자로 등록돼 있어요.")


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


# ── 일반 텍스트 ────────────────────────────────────────────────
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    text = update.message.text.strip()
    # 처음 말 거는 쌤에게는 사용 안내를 공지처럼 먼저 보낸다
    first_contact = chat_id not in known_chats()
    remember_chat(chat_id)
    if first_contact:
        await update.message.reply_text(GUIDE, parse_mode="HTML")

    low = text.lstrip("/").strip()
    parts = low.split()
    kw = parts[0] if parts else ""

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
            written, warnings = ac.write_attendance(
                wb, info["sheet"], info["date"], info["data"], enroll=enroll
            )
            wb.save(path)
            life = info["data"].get("학적") or {}
            if life:  # 학적 변동을 재적 기록에 반영
                apply_enroll_events(info["sheet"], info["date"], life)
            msg = f"✅ 기록 완료 ({info['sheet']} {info['date']}) — {len(written)}건\n" + "\n".join(
                "  • " + w for w in written
            )
            if warnings:
                msg += "\n⚠️ " + " / ".join(warnings)
            await update.message.reply_text(msg)
            return
        if head in ("취소", "아니", "아니오", "no", "cancel"):
            pending.pop(chat_id)
            await update.message.reply_text("취소했어요.")
            return
        # 그 외 입력은 아래에서 새로 처리 (기존 대기는 덮어씀)

    # 한글 키워드를 명령처럼 처리 (슬래시 있어도/없어도)
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
    if low in ("설정초기화", "기본설정복원", "반목록갱신"):
        return await cmd_reset_config(update, context)
    if low in ("주간보고서", "보고서", "주간출결"):
        return await cmd_report(update, context)
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

    await context.bot.send_chat_action(chat_id=chat_id, action="typing")
    parsed = await parse_message(text, parse_wb)

    if parsed.get("type") != "attendance":
        reply = await chat_reply(chat_id, text)
        await update.message.reply_text(reply)
        return

    # 날짜의 달에 해당하는 파일을 대상으로
    target_wb, _, month = load_wb_for_date(parsed.get("date", "0/0"))
    if target_wb is None:
        await update.message.reply_text(
            f"{month}월 출석부 파일이 없어요. 먼저 파일을 보내거나 /생성 해주세요."
        )
        return

    preview, err = build_preview(target_wb, parsed)
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
    # 지금 확인 시각이 된 반들 추리기 (요일별 표에서 오늘 요일의 시각을 찾음)
    due = []
    for cls, table in load_times().items():
        tm = table.get(wd) if isinstance(table, dict) else None
        if not tm:
            continue  # 오늘 수업(알림) 없는 반은 건너뜀
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

    wb, _ = load_current_wb()  # 이번 달 출석부 (없으면 None)
    teachers = load_teachers()
    date_str = f"{now.month}/{now.day}"
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
    fname = f"수학과 주간 출결사항 ({monday.month}.{monday.day}~{sunday.month}.{sunday.day}).pdf"
    out = os.path.join(DATA_DIR, fname)
    n = rpt.build_report_pdf(wb, out, monday, sunday, sunday.year, enroll=load_enroll())
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
