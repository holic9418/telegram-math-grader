"""
텔레그램 → Claude 봇.

할 수 있는 일:
  1) 텍스트 메시지를 보내면  → Claude가 답을 해줍니다.
  2) 수학 문제 사진을 보내면 → Claude가 채점해서, 정답에는 ⭕ 오답에는 빗금(／)을
                              직접 그려 넣은 사진을 다시 보내줍니다.

봇 토큰과 API 키는 코드에 직접 넣지 않고 환경변수에서 읽어옵니다.
"""

import base64
import io
import json
import os

from dotenv import load_dotenv
from PIL import Image, ImageDraw

import anthropic
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

# ── 1. 환경변수에서 비밀 값 읽기 ────────────────────────────────
# 같은 폴더의 .env 파일을 읽어 환경변수로 불러옵니다.
load_dotenv()

# 코드에 토큰/키를 직접 쓰지 않고, 실행 환경에서 가져옵니다.
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
# ANTHROPIC_API_KEY 는 anthropic 라이브러리가 환경변수에서 자동으로 읽습니다.

# ── 2. Claude 클라이언트 준비 ──────────────────────────────────
# 인자를 비워두면 환경변수 ANTHROPIC_API_KEY 를 자동으로 사용합니다.
claude = anthropic.Anthropic()

# 사진을 채점할 때 Claude 에게 주는 "역할 설명(시스템 프롬프트)".
# 여기서는 설명 대신, 각 답의 "위치"와 "정답 여부"만 JSON 으로 받아옵니다.
# 그 좌표에 우리가 직접 O / 빗금을 그립니다.
GRADING_PROMPT = """당신은 수학 문제 채점기입니다.
학생이 손으로 푼 수학 문제 사진을 보고, 각 문제(또는 각 답)마다 채점하세요.

반드시 아래 형식의 JSON "만" 출력하세요. 다른 설명, 인사말, 코드블록 표시(```)는 절대 넣지 마세요.

{
  "problems": [
    {"number": 1, "correct": true,  "point": [x, y]},
    {"number": 2, "correct": false, "point": [x, y]}
  ]
}

규칙:
- point 는 그 문제의 "좌측 상단"(보통 문제 번호가 있는 위치)의 좌표입니다.
  여기에 채점 표시(O 또는 빗금)를 그립니다.
- 좌표는 사진의 왼쪽 위를 (0,0), 오른쪽 아래를 (1000,1000) 으로 하는 0~1000 사이 정수입니다.
  (실제 사진 크기와 상관없이 항상 0~1000 비율로 환산해서 쓰세요.)
- correct 는 답이 맞으면 true, 틀리면 false 입니다.
- 답을 알아볼 수 없거나 문제가 하나도 없으면 {"problems": []} 를 출력하세요."""


def _extract_json(text: str) -> dict:
    """Claude 응답에서 JSON 부분만 안전하게 뽑아냅니다.

    가끔 앞뒤에 설명이나 ```json 같은 게 섞여 와도, 첫 '{' 부터
    마지막 '}' 까지를 잘라서 파싱합니다. 실패하면 빈 결과를 돌려줍니다.
    """
    try:
        start = text.index("{")
        end = text.rindex("}") + 1
        return json.loads(text[start:end])
    except (ValueError, json.JSONDecodeError):
        return {"problems": []}


def _draw_marks(image_bytes: bytes, problems: list) -> bytes:
    """각 문제의 좌측 상단에 O(정답) / 빗금(오답)을 그려서 새 사진(바이트)으로 돌려줍니다."""
    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    width, height = image.size
    draw = ImageDraw.Draw(image)

    # 표시 크기/두께를 사진 크기에 비례하게 정합니다. (고정 크기 마크)
    radius = max(12, round(min(width, height) / 22))
    line_width = max(3, round(min(width, height) / 150))
    red = (220, 30, 30)

    for item in problems:
        point = item.get("point")
        if not point or len(point) != 2:
            continue

        # 0~1000 비율 좌표를 실제 픽셀 좌표로 환산.
        cx = point[0] / 1000 * width
        cy = point[1] / 1000 * height
        # 마크가 사진 밖으로 나가지 않게 살짝 안쪽으로 당겨줍니다.
        cx = min(max(cx, radius), width - radius)
        cy = min(max(cy, radius), height - radius)

        if item.get("correct"):
            # 정답 → 빨간 동그라미
            draw.ellipse(
                [cx - radius, cy - radius, cx + radius, cy + radius],
                outline=red,
                width=line_width,
            )
        else:
            # 오답 → 빗금(／). 마크 영역을 가로지르는 대각선.
            draw.line(
                [cx - radius, cy + radius, cx + radius, cy - radius],
                fill=red,
                width=line_width,
            )

    out = io.BytesIO()
    image.save(out, format="JPEG")
    return out.getvalue()


# ── 3-a. 텍스트 메시지가 올 때 실행되는 함수 ────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_text = update.message.text  # 사용자가 보낸 글자

    # Claude 에게 물어보기
    response = claude.messages.create(
        model="claude-opus-4-8",
        max_tokens=1024,
        messages=[{"role": "user", "content": user_text}],
    )

    # 응답에서 텍스트만 뽑아내기
    reply = "".join(
        block.text for block in response.content if block.type == "text"
    )

    # 텔레그램으로 답장 보내기
    await update.message.reply_text(reply)


# ── 3-b. 사진이 올 때 실행되는 함수 (수학 문제 채점) ────────────
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # "채점 중" 이라고 먼저 알려줍니다. (Claude 응답에 몇 초 걸립니다)
    await update.message.reply_text("사진을 받았어요. 채점 중입니다... ✏️")

    # 텔레그램은 같은 사진을 여러 해상도로 보내줍니다.
    # 마지막([-1]) 것이 가장 큰(선명한) 사진이라 채점에 유리합니다.
    photo = update.message.photo[-1]
    tg_file = await context.bot.get_file(photo.file_id)

    # 사진을 메모리로 내려받습니다.
    image_bytes = bytes(await tg_file.download_as_bytearray())
    image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    # Claude 에게 사진을 보내고, 각 답의 위치 + 정답 여부(JSON)를 받습니다.
    response = claude.messages.create(
        model="claude-opus-4-8",
        max_tokens=1024,
        system=GRADING_PROMPT,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            # 텔레그램 사진은 보통 JPEG 로 옵니다.
                            "media_type": "image/jpeg",
                            "data": image_b64,
                        },
                    },
                    {"type": "text", "text": "이 사진을 채점해서 JSON 으로 알려주세요."},
                ],
            }
        ],
    )

    raw = "".join(block.text for block in response.content if block.type == "text")
    result = _extract_json(raw)
    problems = result.get("problems", [])

    if not problems:
        await update.message.reply_text(
            "채점할 답을 찾지 못했어요. 😢 문제와 답이 잘 보이게, 밝은 곳에서 다시 찍어 보내주세요."
        )
        return

    # 원본 사진 위에 O / 빗금을 그려서 다시 보냅니다.
    marked = _draw_marks(image_bytes, problems)

    correct_count = sum(1 for p in problems if p.get("correct"))
    total = len(problems)
    await update.message.reply_photo(
        photo=io.BytesIO(marked),
        caption=f"채점 완료! ⭕ {correct_count} / 전체 {total}",
    )


# ── 4. 봇 실행 ────────────────────────────────────────────────
def main() -> None:
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    # 사진이 오면 handle_photo (수학 문제 채점) 를 호출
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    # 텍스트 메시지(명령어 제외)가 오면 handle_message 를 호출
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)
    )

    print("봇이 실행되었습니다. 텔레그램에서 메시지나 사진을 보내보세요. (종료: Ctrl+C)")
    app.run_polling()


if __name__ == "__main__":
    main()
