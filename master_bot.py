# -*- coding: utf-8 -*-
"""원장용 마스터봇 (읽기 전용).

같은 Mac 안의 세 과목 데이터 폴더(data-math/korean/english)를 읽어서
원장이 한 곳에서 전체를 조회한다. 입력·수정 기능은 없다.

기능:
  • 현황 / 미입력      — 오늘·밀린 미입력 반을 세 과목 한 번에
  • <과목> <반> 미리보기 — 과목 지정 출석표 이미지
  • <과목> 통계 [기간]  — 출석률·결석·과제 미제출 집계
  • <과목> 주간보고서    — 그 과목 주간 PDF / '주간보고서'면 세 과목 전부
  • 매일 저녁 미입력 요약 + 일요일 저녁 주간보고서 자동 전송

환경변수:
  TELEGRAM_BOT_TOKEN  : 마스터 봇 토큰(필수)
  MASTER_DATA_BASE    : 과목 폴더들이 있는 상위 경로(기본: 이 파일 위치)
  MASTER_CHAT_ID      : 원장 chat_id(자동전송 대상·사용권한). 없으면 '관리자 등록'한 첫 사람.
  MASTER_DAILY_HHMM   : 매일 미입력 요약 시각(기본 21:30)
"""
import os
import re
import io
import json
import logging
import datetime
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application, ApplicationBuilder, ContextTypes,
    CommandHandler, MessageHandler, filters,
)

import attendance_core as ac
import report as rpt
import subjects

load_dotenv()
logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s | %(message)s", level=logging.INFO)
log = logging.getLogger("master-bot")

TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
BASE = os.environ.get("MASTER_DATA_BASE") or os.path.dirname(os.path.abspath(__file__))
KST = ZoneInfo("Asia/Seoul")
BACKLOG_SINCE = datetime.date(2026, 7, 13)   # 이 날짜 이전 옛 미입력은 무시
WD = ["월", "화", "수", "목", "금", "토", "일"]
# 자동 전송(매일 미입력 요약·주간보고서). 기본 꺼짐 — 켜려면 MASTER_AUTO_SEND=1
AUTO_SEND = (os.environ.get("MASTER_AUTO_SEND", "0").strip().lower()
             in ("1", "on", "true", "yes", "y"))

# 표시이름 → (subjects.py 키, 데이터 폴더)
SUBJECTS = {
    "수학": ("math", "data-math"),
    "국어": ("korean", "data-korean"),
    "영어": ("english", "data-english"),
}
ALIASES = {"수학": "수학", "국어": "국어", "영어": "영어",
           "math": "수학", "korean": "국어", "english": "영어"}


# ── 과목별 데이터 접근 (DATA_DIR 의존 로직을 폴더 인자로 일반화) ──
def sdir(subject_label):
    return os.path.join(BASE, SUBJECTS[subject_label][1])


def _load_json(dd, name, default):
    p = os.path.join(dd, name)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def latest_file(dd):
    """폴더에서 가장 최근 'YY.MM 출석부.xlsx' 경로. 없으면 None."""
    if not os.path.isdir(dd):
        return None
    rx = re.compile(r"^(\d{2})\.(\d{2}) 출석부\.xlsx$")
    best, bestkey = None, None
    for f in os.listdir(dd):
        m = rx.match(f)
        if not m:
            continue
        key = (int(m.group(1)), int(m.group(2)))
        if bestkey is None or key > bestkey:
            bestkey, best = key, f
    return os.path.join(dd, best) if best else None


def month_path(dd, y, m):
    return os.path.join(dd, f"{y % 100:02d}.{m:02d} 출석부.xlsx")


def closed_of(dd):
    return _load_json(dd, "closed_classes.json", {})


def enroll_of(dd):
    return _load_json(dd, "enrollments.json", {})


def input_blocked(dd, sheet, date_str):
    """종강한 반의 '종강 주 이후' 날짜면 True."""
    closed = closed_of(dd).get(sheet)
    if not closed:
        return False
    y = datetime.date.today().year
    try:
        cm, cd = map(int, str(closed).split("/"))
        dm, dd2 = map(int, str(date_str).split("/"))
    except ValueError:
        return False
    _, csun = rpt.week_bounds(datetime.date(y, cm, cd))
    return datetime.date(y, dm, dd2) > csun


def alarm_blocked(dd, sheet, date_str):
    """종강한 반은 종강일부터(그 주 포함) 미입력 집계 제외."""
    closed = closed_of(dd).get(sheet)
    if not closed:
        return False
    y = datetime.date.today().year
    try:
        cm, cd = map(int, str(closed).split("/"))
        dm, dd2 = map(int, str(date_str).split("/"))
    except ValueError:
        return False
    return datetime.date(y, dm, dd2) >= datetime.date(y, cm, cd)


# ── 미입력 현황 ────────────────────────────────────────────────
def subject_missing(label, today):
    """한 과목에서 미입력(오늘까지, 밀린 것 포함) 반·날짜. {반: ['M/D',...]}."""
    dd = sdir(label)
    path = latest_file(dd)
    if not path:
        return None, {}    # 파일 없음
    wb = ac.load_workbook(path)
    dates_by = ac.sheet_dates(wb)
    out = {}
    for cls in wb.sheetnames:
        ws = wb[cls]
        if not ac.get_roster(ws):
            continue
        miss = []
        for ds in dates_by.get(cls, []):
            try:
                m, d = map(int, ds.split("/"))
                dd2 = datetime.date(today.year, m, d)
            except ValueError:
                continue
            if dd2 < BACKLOG_SINCE or dd2 > today:
                continue
            if alarm_blocked(dd, cls, ds):
                continue
            if ac.attendance_recorded(ws, ds) is False:
                miss.append(ds)
        if miss:
            out[cls] = miss
    return wb, out


def build_status(today):
    lines = [f"📋 <b>미입력 현황</b> · {today.month}월 {today.day}일({WD[today.weekday()]}) 기준"]
    any_file = False
    for label in SUBJECTS:
        disp = subjects.get(SUBJECTS[label][0])["display"]
        wb, miss = subject_missing(label, today)
        if wb is None:
            lines.append(f"\n<b>[{label}]</b> 출석부 파일 없음")
            continue
        any_file = True
        if not miss:
            lines.append(f"\n<b>[{label}]</b> ✅ 미입력 없음")
            continue
        lines.append(f"\n<b>[{label}]</b> ⚠️ 미입력 {len(miss)}개 반")
        for cls, dates in sorted(miss.items()):
            ktxt = ", ".join(f"{int(x.split('/')[0])}/{int(x.split('/')[1])}" for x in dates)
            lines.append(f"  • {cls}: {ktxt}")
    if not any_file:
        return "세 과목 모두 출석부 파일이 없어요. 각 과목 봇에 파일을 올려주세요."
    return "\n".join(lines)


# ── 기간·반 파싱 ───────────────────────────────────────────────
def parse_subject(text):
    """맨 앞(또는 안)에서 과목 라벨을 찾아 반환. 없으면 None."""
    low = text.strip()
    for k, label in ALIASES.items():
        if low.startswith(k):
            return label, low[len(k):].strip()
    for k, label in ALIASES.items():
        if k in low:
            return label, low.replace(k, "", 1).strip()
    return None, low


def match_sheet(rest, sheets):
    t = rest.replace(" ", "")
    for s in sorted(sheets, key=len, reverse=True):
        if s.replace(" ", "") in t:
            return s
    return None


def find_student(rest, wb):
    """rest에 들어있는 학생 이름과 그 학생이 속한 반들. (name, [sheets]) / (None, [])."""
    t = rest.replace(" ", "")
    found = {}
    for s in wb.sheetnames:
        for nm in ac.get_roster(wb[s]):
            if nm and nm in t:
                found.setdefault(nm, []).append(s)
    if not found:
        return None, []
    name = max(found, key=len)          # 가장 구체적인(긴) 이름 우선
    return name, found[name]


def period_dates(wb, sheet, text, today, default="week"):
    """text의 기간 표현으로 그 반의 날짜 리스트 반환. default=week|month."""
    if any(k in text for k in ("이번주", "금주")):
        mon, sun = rpt.week_bounds(today)
        return rpt.week_dates(wb[sheet], mon, sun, mon.year)
    if any(k in text for k in ("저번주", "지난주")):
        mon, sun = rpt.week_bounds(today - datetime.timedelta(days=7))
        return rpt.week_dates(wb[sheet], mon, sun, mon.year)
    md = re.search(r"(\d{1,2})\s*[/.월]\s*(\d{1,2})", text)
    if md:
        return [f"{int(md.group(1))}/{int(md.group(2))}"]
    if "오늘" in text:
        return [f"{today.month}/{today.day}"]
    if default == "month":
        return ac.sheet_dates(wb).get(sheet, [])
    mon, sun = rpt.week_bounds(today)
    return rpt.week_dates(wb[sheet], mon, sun, mon.year)


# ── 통계 ───────────────────────────────────────────────────────
def subject_stats(label, text, today):
    dd = sdir(label)
    # 기간 파일 선택
    md = re.search(r"(\d{1,2})\s*월", text)
    if any(k in text for k in ("이번주", "금주", "저번주", "지난주")):
        base = today if any(k in text for k in ("이번주", "금주")) else today - datetime.timedelta(days=7)
        mon, sun = rpt.week_bounds(base)
        path = month_path(dd, mon.year, mon.month)
        mode = ("week", mon, sun)
        period = f"{mon.month}/{mon.day}~{sun.month}/{sun.day}"
    else:
        month = (12 if today.month == 1 else today.month - 1) if "지난달" in text \
            else (int(md.group(1)) if md else today.month)
        path = month_path(dd, today.year, month)
        mode, period = ("month", None, None), f"{month}월"
    if not os.path.exists(path):
        path = latest_file(dd)
    if not path:
        return f"[{label}] 출석부 파일이 없어요."
    wb = ac.load_workbook(path)
    sheets = [s for s in wb.sheetnames if ac.get_roster(wb[s])]

    def dates_for(sheet):
        if mode[0] == "week":
            return rpt.week_dates(wb[sheet], mode[1], mode[2], mode[1].year)
        return ac.sheet_dates(wb).get(sheet, [])

    present = absent = hw_x = days = 0
    for sheet in sheets:
        ws = wb[sheet]
        roster = ac.get_roster(ws)
        for ds in dates_for(sheet):
            top = ac.find_date_block(ws, ds)
            if top is None or ac.attendance_recorded(ws, ds) is False:
                continue
            if input_blocked(dd, sheet, ds):
                continue
            days += 1
            for name, col in roster.items():
                av = ws.cell(top + ac.ROW_OFFSET["출석"], col).value
                if av in (None, ""):
                    continue
                if ac.is_absent(av):
                    absent += 1
                else:
                    present += 1
                hv = ws.cell(top + ac.ROW_OFFSET["과제수행"], col).value
                if hv is not None and str(hv).strip().upper().startswith("X"):
                    hw_x += 1
    total = present + absent
    rate = f"{present*100//total}%" if total else "-"
    return (f"📊 <b>[{label}] 통계</b> · {period}\n"
            f"• 수업 기록: {days}회\n"
            f"• 출석 {present} / 결석 {absent} → 출석률 <b>{rate}</b>\n"
            f"• 과제 미제출: {hw_x}건")


# ── 주간보고서 ─────────────────────────────────────────────────
def subject_weekly(label, today):
    dd = sdir(label)
    disp = subjects.get(SUBJECTS[label][0])["display"]
    mon, sun = rpt.week_bounds(today)
    segments, seen = [], set()
    for d in (mon, sun):
        ym = (d.year, d.month)
        if ym in seen:
            continue
        seen.add(ym)
        p = month_path(dd, d.year, d.month)
        if os.path.exists(p):
            segments.append((ac.load_workbook(p), d.year))
    if not segments:
        p = latest_file(dd)
        if not p:
            return None
        segments = [(ac.load_workbook(p), sun.year)]
    out = os.path.join(dd, f"{disp} 주간 출결 ({mon.month}.{mon.day}~{sun.month}.{sun.day}).pdf")
    skip = {c for c in closed_of(dd) if input_blocked(dd, c, f"{mon.month}/{mon.day}")}
    n = rpt.build_report_pdf(segments, out, mon, sun,
                             title=f"<{disp} 주간 출결사항>", enroll=enroll_of(dd), skip=skip)
    return (out, n) if n else None


# ── 권한 ───────────────────────────────────────────────────────
def master_admin():
    env = (os.environ.get("MASTER_CHAT_ID") or "").strip()
    if env.lstrip("-").isdigit():
        return int(env)
    p = os.path.join(BASE, "master_admin.json")
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8")).get("chat_id")
        except Exception:
            return None
    return None


def set_master_admin(chat_id):
    with open(os.path.join(BASE, "master_admin.json"), "w", encoding="utf-8") as f:
        json.dump({"chat_id": chat_id}, f)


GUIDE = (
    "👑 <b>원장용 마스터봇</b> (조회 전용)\n\n"
    "• <b>현황</b> — 세 과목 미입력 반 한눈에\n"
    "• <b>수학 초5 이번주 미리보기</b> — 과목·반·기간 지정 출석표\n"
    "• <b>원서진</b> / <b>수학 원서진</b> — 특정 학생 출결(과목 생략 시 세 과목 전체에서 검색)\n"
    "• <b>국어 통계 이번주</b> — 출석률·결석·과제 미제출\n"
    "• <b>영어 주간보고서</b> / <b>주간보고서</b>(세 과목 전부)\n\n"
    "필요할 때 직접 물어보시면 돼요. (자동 전송은 꺼져 있어요)"
)


# ── 핸들러 ─────────────────────────────────────────────────────
async def send_long(bot, chat_id, text, **kw):
    """텔레그램 4096자 제한 대비 — 줄 단위로 잘라 나눠 보낸다."""
    while text:
        chunk = text[:4000]
        if len(text) > 4000:
            cut = chunk.rfind("\n")
            if cut > 0:
                chunk = text[:cut]
        await bot.send_message(chat_id, chunk, **kw)
        text = text[len(chunk):].lstrip("\n")


async def send_student_views(update, context, label, dd, wb, name, sheets, rest, today):
    """한 학생의 (반별) 출결표를 이미지로 보낸다. 보낸 장수 반환."""
    sent = 0
    for s in sheets:
        enroll = enroll_of(dd).get(s, {})
        dates = [d for d in period_dates(wb, s, rest, today, default="month")
                 if ac.find_date_block(wb[s], d) is not None
                 and ac.student_active(enroll, name, d)]
        if not dates:
            continue
        img = rpt.render_student_table(name, wb[s], dates)
        if img is None:
            continue
        bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        await update.message.reply_photo(
            photo=bio, caption=f"📷 {label} · {name} · {s} ({dates[0]}~{dates[-1]})")
        sent += 1
    return sent


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(GUIDE, parse_mode="HTML")


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    text = (update.message.text or "").strip()
    low = text.lstrip("/").strip()

    admin = master_admin()
    if low in ("관리자등록", "관리자 등록", "관리자"):
        if admin is None:
            set_master_admin(chat_id)
            await update.message.reply_text("✅ 원장(관리자)으로 등록됐어요.\n" + GUIDE,
                                            parse_mode="HTML")
        elif admin == chat_id:
            await update.message.reply_text("이미 등록돼 있어요. 👍")
        else:
            await update.message.reply_text("이미 다른 분이 등록돼 있어요.")
        return
    if admin is not None and chat_id != admin:
        await update.message.reply_text("이 봇은 원장 전용이에요. 🔒")
        return
    if admin is None:
        await update.message.reply_text("먼저 <b>관리자 등록</b> 을 보내 원장으로 등록해 주세요.",
                                        parse_mode="HTML")
        return

    today = datetime.datetime.now(KST).date()
    if low in ("도움말", "도움", "시작", "start"):
        return await update.message.reply_text(GUIDE, parse_mode="HTML")

    if low in ("현황", "미입력", "미입력현황", "전체현황"):
        return await send_long(context.bot, chat_id, build_status(today), parse_mode="HTML")

    if "주간보고서" in low or low in ("보고서", "주간출결"):
        label, _ = parse_subject(low)
        labels = [label] if label else list(SUBJECTS)
        await context.bot.send_chat_action(chat_id=chat_id, action="upload_document")
        sent = 0
        for lb in labels:
            try:
                res = subject_weekly(lb, today)
            except Exception as e:
                log.warning("주간보고서 실패 %s: %s", lb, e)
                res = None
            if not res:
                await update.message.reply_text(f"[{lb}] 이번 주 보고서 데이터가 없어요.")
                continue
            out, n = res
            with open(out, "rb") as f:
                await update.message.reply_document(document=f, filename=os.path.basename(out),
                                                    caption=f"📄 {lb} 주간 출결 · {n}개 반")
            sent += 1
        if not sent:
            await update.message.reply_text("보낼 주간보고서가 없어요.")
        return

    # 과목 지정 명령 (미리보기 / 통계)
    label, rest = parse_subject(low)
    if label:
        if "통계" in rest or "집계" in rest:
            return await update.message.reply_text(subject_stats(label, rest, today),
                                                   parse_mode="HTML")
        dd = sdir(label)
        path = latest_file(dd)
        if not path:
            return await update.message.reply_text(f"[{label}] 출석부 파일이 없어요.")
        wb = ac.load_workbook(path)
        # 반 지정이면 반 출석표, 아니면 학생 조회
        sheet = match_sheet(rest, wb.sheetnames)
        if sheet:
            dates = [d for d in period_dates(wb, sheet, rest, today)
                     if ac.find_date_block(wb[sheet], d) is not None]
            if not dates:
                return await update.message.reply_text(f"[{label}] {sheet}: 그 기간에 수업일이 없어요.")
            img = rpt.render_class_table(sheet, wb[sheet], dates)
            bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
            return await update.message.reply_photo(
                photo=bio, caption=f"📷 {label} · {sheet} ({dates[0]}~{dates[-1]})")
        # 학생 조회 (여러 반이면 반마다 한 장)
        name, sheets = find_student(rest, wb)
        if not name:
            return await update.message.reply_text(
                f"[{label}] 반이나 학생을 못 찾았어요. 반 목록: " + ", ".join(wb.sheetnames))
        await context.bot.send_chat_action(chat_id=chat_id, action="upload_photo")
        sent = await send_student_views(update, context, label, dd, wb, name, sheets, rest, today)
        if not sent:
            return await update.message.reply_text(f"[{label}] {name} 학생의 해당 기간 수업일이 없어요.")
        return

    # 과목 미지정 — 이름만/‘현황 이름’이면 세 과목 전체에서 그 학생을 찾아 보여준다
    q = re.sub(r"현황|미입력현황|미입력|전체현황|미리보기|조회|보여줘|보여주세요", " ", low).strip()
    if q:
        await context.bot.send_chat_action(chat_id=chat_id, action="upload_photo")
        total = 0
        for lb in SUBJECTS:
            p = latest_file(sdir(lb))
            if not p:
                continue
            wb = ac.load_workbook(p)
            name, sheets = find_student(q, wb)
            if name:
                total += await send_student_views(update, context, lb, sdir(lb), wb, name, sheets, q, today)
        if total:
            return

    await update.message.reply_text("무슨 말인지 잘 모르겠어요.\n" + GUIDE, parse_mode="HTML")


# ── 자동 전송 (매일 미입력 요약 · 일요일 주간보고서) ────────────
_last = {"daily": None, "weekly": None}


async def tick(context: ContextTypes.DEFAULT_TYPE):
    if not AUTO_SEND:
        return
    admin = master_admin()
    if admin is None:
        return
    now = datetime.datetime.now(KST)
    hhmm = os.environ.get("MASTER_DAILY_HHMM", "21:30")
    try:
        dh, dm = map(int, hhmm.split(":"))
    except ValueError:
        dh, dm = 21, 30
    # 매일 미입력 요약
    dkey = now.date().isoformat()
    if now.hour == dh and dm <= now.minute < dm + 3 and _last["daily"] != dkey:
        _last["daily"] = dkey
        try:
            await send_long(context.bot, admin, build_status(now.date()), parse_mode="HTML")
        except Exception as e:
            log.warning("일일 요약 실패: %s", e)
    # 일요일 20:00 주간보고서
    if now.weekday() == 6 and now.hour == 20 and now.minute < 3 and _last["weekly"] != dkey:
        _last["weekly"] = dkey
        for lb in SUBJECTS:
            try:
                res = subject_weekly(lb, now.date())
                if res:
                    out, n = res
                    with open(out, "rb") as f:
                        await context.bot.send_document(admin, document=f,
                                                        filename=os.path.basename(out),
                                                        caption=f"📄 {lb} 주간 출결 · {n}개 반")
            except Exception as e:
                log.warning("주간 자동전송 실패 %s: %s", lb, e)


def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle))
    if AUTO_SEND and app.job_queue:
        app.job_queue.run_repeating(tick, interval=60, first=10)
    log.info("마스터봇 시작 · 데이터 상위 폴더: %s · 자동전송: %s",
             BASE, "ON" if AUTO_SEND else "OFF")
    app.run_polling()


if __name__ == "__main__":
    main()
